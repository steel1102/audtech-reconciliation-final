import { Router, type IRouter } from "express";
import multer from "multer";
import * as XLSX from "xlsx";
import { logger } from "../lib/logger";

const router: IRouter = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

interface LedgerRow {
  ledgerCode: string;
  ledgerName: string;
  balance: number;
}

type ReconciliationStatus = "matched" | "mismatched" | "missing_current" | "missing_prior" | "possible_regroup";
type MatchStrategy = "exact_code" | "exact_name" | "ratio" | "partial" | "token_sort" | "token_set";

interface ReconciliationRow {
  status: ReconciliationStatus;
  ledgerCode: string;
  ledgerName: string;
  priorBalance: number;
  currentBalance: number;
  variance: number;
  matchScore: number | null;
  matchedWith: string | null;
  matchStrategy: MatchStrategy | null;
}

interface ReconciliationSummary {
  totalPrior: number;
  totalCurrent: number;
  matchedCount: number;
  mismatchedCount: number;
  missingCurrentCount: number;
  missingPriorCount: number;
  possibleRegroupCount: number;
  totalVariance: number;
}

let lastResult: { summary: ReconciliationSummary; rows: ReconciliationRow[] } | null = null;

function isNumericValue(v: unknown): boolean {
  const s = String(v ?? "").trim().replace(/,/g, "");
  return s !== "" && !isNaN(parseFloat(s)) && isFinite(Number(s));
}

function detectColumns(headers: string[], sample: Record<string, unknown>[]): {
  codeKey: string;
  nameKey: string;
  debitKey: string | undefined;
  creditKey: string | undefined;
  balanceKey: string;
} {
  const lh = (h: string) => h.toLowerCase().trim();

  // Header keyword matching — more specific patterns first, exclude hybrid like "Account Name"
  const debitKey  = headers.find(h => /\b(debit|dr)\b/i.test(h));
  const creditKey = headers.find(h => /\b(credit|cr)\b/i.test(h));

  // Code: explicit code/account-number headers, but NOT "account name"-style
  const codeKey = headers.find(h => {
    const l = lh(h);
    return /\bcode\b/i.test(l) || /account[\s._-]?no/i.test(l) || /\bacct[\s._-]?no\b/i.test(l) || /\ba\/c\b/i.test(l);
  });

  // Name: explicit name/description/ledger headers
  const nameByHeader = headers.find(h => {
    const l = lh(h);
    return /\bname\b/.test(l) || /\bdescription\b/.test(l) || /\bparticulars\b/.test(l)
        || /\bledger\b/.test(l) || /\btitle\b/.test(l);
  });

  // Balance: explicit balance/amount headers
  const balanceByHeader = headers.find(h => /\b(balance|amount|net|total)\b/i.test(h));

  // ── Value-based fallback ──
  // Compute per-column numeric ratio from sample rows
  const colStats = headers
    .filter(h => h !== debitKey && h !== creditKey)
    .map(h => {
      const vals = sample.map(r => r[h]).filter(v => String(v ?? "").trim() !== "");
      const numCount = vals.filter(isNumericValue).length;
      return { h, numericRatio: vals.length > 0 ? numCount / vals.length : 0, nonEmpty: vals.length };
    })
    .sort((a, b) => b.nonEmpty - a.nonEmpty); // prefer columns with more data

  // Most numeric column (excluding already identified) → balance
  const balanceFallback = colStats.filter(s => s.h !== codeKey && s.h !== nameByHeader)
    .sort((a, b) => b.numericRatio - a.numericRatio)[0]?.h ?? headers[headers.length - 1];

  // Most text-like column → name
  const usedBalance = balanceByHeader ?? balanceFallback;
  const nameFallback = colStats
    .filter(s => s.h !== codeKey && s.h !== usedBalance)
    .sort((a, b) => a.numericRatio - b.numericRatio)[0]?.h ?? headers[0];

  return {
    codeKey:    codeKey ?? "",
    nameKey:    nameByHeader ?? nameFallback,
    debitKey,
    creditKey,
    balanceKey: balanceByHeader ?? balanceFallback,
  };
}

function parseWorkbook(buffer: Buffer): LedgerRow[] {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const raw: Record<string, unknown>[] = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (raw.length === 0) return [];

  const headers = Object.keys(raw[0]);
  const sample = raw.slice(0, Math.min(10, raw.length));
  const { codeKey, nameKey, debitKey, creditKey, balanceKey } = detectColumns(headers, sample);

  logger.info({ headers, codeKey, nameKey, debitKey, creditKey, balanceKey }, "COLUMN_DETECTION");

  const rows: LedgerRow[] = [];
  for (const row of raw) {
    const code = codeKey ? String(row[codeKey] ?? "").trim() : "";
    const name = String(row[nameKey] ?? "").trim();
    if (!name || isNumericValue(name)) continue; // skip rows where name looks like a number

    let balance: number;
    if (debitKey && creditKey) {
      const debit = parseFloat(String(row[debitKey]).replace(/,/g, "")) || 0;
      const credit = parseFloat(String(row[creditKey]).replace(/,/g, "")) || 0;
      balance = debit - credit;
    } else {
      balance = parseFloat(String(row[balanceKey]).replace(/,/g, "")) || 0;
    }

    rows.push({ ledgerCode: code, ledgerName: name, balance });
  }

  return rows;
}

function normalize(s: string): string {
  return s
    .toLowerCase()
    .replace(/-+/g, " ")
    .replace(/[^a-z0-9]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function stem(word: string): string {
  if (word.length > 5 && word.endsWith("ies")) return word.slice(0, -3) + "y";
  if (word.length > 5 && word.endsWith("es") && !word.endsWith("ss")) return word.slice(0, -2);
  if (word.length > 4 && word.endsWith("s") && !word.endsWith("ss")) return word.slice(0, -1);
  return word;
}

function normalizedTokens(s: string): string[] {
  return normalize(s).split(" ").filter(Boolean).map(stem);
}

function normalizedStemmed(s: string): string {
  return normalizedTokens(s).join(" ");
}

function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  const prev: number[] = Array.from({ length: n + 1 }, (_, j) => j);
  const curr: number[] = new Array(n + 1).fill(0);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      curr[j] =
        a[i - 1] === b[j - 1]
          ? prev[j - 1]
          : 1 + Math.min(prev[j], curr[j - 1], prev[j - 1]);
    }
    prev.splice(0, n + 1, ...curr);
  }
  return prev[n];
}

function ratio(a: string, b: string): number {
  if (a === b) return 100;
  if (!a || !b) return 0;
  const dist = levenshtein(a, b);
  const total = a.length + b.length;
  return Math.round(((total - dist) / total) * 100);
}

function partialRatio(a: string, b: string): number {
  if (a === b) return 100;
  if (!a || !b) return 0;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length > b.length ? a : b;
  if (shorter.length === 0) return 0;
  let best = 0;
  for (let i = 0; i <= longer.length - shorter.length; i++) {
    const window = longer.slice(i, i + shorter.length);
    const s = ratio(shorter, window);
    if (s > best) best = s;
    if (best === 100) break;
  }
  return best;
}

function tokenSortRatio(a: string, b: string): number {
  const sortedA = normalize(a).split(" ").filter(Boolean).sort().join(" ");
  const sortedB = normalize(b).split(" ").filter(Boolean).sort().join(" ");
  return ratio(sortedA, sortedB);
}

function tokenSetRatio(a: string, b: string): number {
  const tokA = normalizedTokens(a);
  const tokB = normalizedTokens(b);
  const setA = new Set(tokA);
  const setB = new Set(tokB);
  const intersection = [...setA].filter((t) => setB.has(t)).sort().join(" ");
  const onlyA = [...setA].filter((t) => !setB.has(t)).sort().join(" ");
  const onlyB = [...setB].filter((t) => !setA.has(t)).sort().join(" ");
  const t1 = intersection;
  const t2 = [intersection, onlyA].filter(Boolean).join(" ");
  const t3 = [intersection, onlyB].filter(Boolean).join(" ");
  return Math.max(ratio(t1, t2), ratio(t1, t3), ratio(t2, t3));
}

/**
 * Composite fuzzy scorer. Runs six similarity passes over normalised/stemmed
 * strings plus a domain-specific alias lookup, then returns the highest score.
 *
 * Passes (in order):
 *  1. ratio           — character-level edit distance on normalised strings
 *  2. partial         — sliding-window ratio (good for substrings / abbreviations)
 *  3. ratio (stemmed) — same as 1 but with plural/suffix stripping first
 *  4. partial stemmed — same as 2 but stemmed
 *  5. token_sort      — sort tokens alphabetically then compare (word-order agnostic)
 *  6. token_set       — intersection/difference set comparison (handles extra words)
 *  7. alias           — hardcoded accounting synonym table (domain knowledge boost)
 */
function fuzzyScore(a: string, b: string): { score: number; strategy: MatchStrategy } {
  const na = normalize(a);
  const nb = normalize(b);
  const sa = normalizedStemmed(a);
  const sb = normalizedStemmed(b);
  if (na === nb) return { score: 100, strategy: "ratio" };
  if (!na || !nb) return { score: 0, strategy: "ratio" };

  const candidates: { score: number; strategy: MatchStrategy }[] = [
    { score: ratio(na, nb),         strategy: "ratio" },
    { score: partialRatio(na, nb),  strategy: "partial" },
    { score: ratio(sa, sb),         strategy: "ratio" },
    { score: partialRatio(sa, sb),  strategy: "partial" },
    { score: tokenSortRatio(a, b),  strategy: "token_sort" },
    { score: tokenSetRatio(a, b),   strategy: "token_set" },
    // Domain alias boost — fires for known accounting synonym pairs.
    // Score is fixed at 88 so it lands in "possible_regroup" (not "matched").
    { score: aliasScore(a, b),      strategy: "token_set" },
  ];
  return candidates.reduce((best, c) => (c.score > best.score ? c : best));
}

// ─── Accounting alias pairs ──────────────────────────────────────────────────
// Firms often record the same account under different but semantically equivalent
// names (e.g. "Outstanding Salaries" vs "Salary Payable"). Pure string similarity
// can score these below the regroup threshold because the words are different even
// though the accounting meaning is identical. The table below lists known alias
// pairs; if both names normalise to a matching pair we inject a synthetic score
// of 88 so the reconciliation engine always flags them as "possible_regroup".
const ACCOUNTING_ALIASES: Array<[string, string]> = [
  // Payroll liabilities
  ["outstanding salaries",    "salary payable"],
  ["outstanding wages",       "wage payable"],
  ["accrued salaries",        "salary payable"],
  ["accrued wages",           "wage payable"],
  // Receivables
  ["debtors",                 "trade receivable"],
  ["sundry debtors",          "trade receivable"],
  ["debtors",                 "trade debtor"],
  ["accounts receivable",     "trade receivable"],
  // Payables
  ["creditors",               "trade payable"],
  ["sundry creditors",        "trade payable"],
  ["creditors",               "trade creditor"],
  ["accounts payable",        "trade payable"],
  // Inventory
  ["stock",                   "inventory"],
  ["closing stock",           "inventory"],
  ["opening stock",           "inventory"],
  ["goods in hand",           "inventory"],
  // Fixed assets
  ["fixed assets",            "property plant and equipment"],
  ["fixed assets",            "ppe"],
  ["tangible assets",         "property plant and equipment"],
  // Cash
  ["cash and bank",           "bank balance"],
  ["cash and bank balances",  "bank balance"],
  ["cash and cash equivalents", "bank balance"],
  ["bank overdraft",          "overdraft"],
  // Loans
  ["bank loan",               "term loan"],
  ["bank term loan",          "term loan"],
  ["bank borrowings",         "term loan"],
  // Taxes
  ["gst receivable",          "input tax credit"],
  ["gst payable",             "output tax liability"],
  ["tds receivable",          "tax deducted at source"],
];

/**
 * Returns 88 if the two ledger names match a known accounting alias pair,
 * otherwise 0. Matching is substring-based on normalised strings so minor
 * surrounding words (e.g. "net" or "total") do not prevent a match.
 */
function aliasScore(a: string, b: string): number {
  const na = normalize(a);
  const nb = normalize(b);
  for (const [x, y] of ACCOUNTING_ALIASES) {
    if ((na.includes(x) && nb.includes(y)) || (na.includes(y) && nb.includes(x))) {
      return 88;
    }
  }
  return 0;
}

// ─── Scoring thresholds ───────────────────────────────────────────────────────
// SCORE_HIGH  — treat a fuzzy hit as a full match (matched/mismatched) rather
//               than a regroup suggestion. Set at 90 to allow for minor spelling
//               differences (plurals, hyphens) while still being confident.
// SCORE_REGROUP — minimum score to flag as "possible_regroup" instead of
//               "missing_current". Lowered to 62 to catch semantically close
//               names (e.g. "Outstanding Salaries" ↔ "Salary Payable" = ~69).
const SCORE_HIGH = 90;
const SCORE_REGROUP = 62;

// Self-test at module load — confirms key pairs score correctly
{
  const selfTests: Array<[string, string]> = [
    ["Trade Receivables",    "Trade Receivable - Local"],
    ["Outstanding Salaries", "Salary Payable"],
    ["Debtors",              "Trade Debtors"],
    ["Inventory Raw Material", "Raw Material Inventory"],
  ];
  for (const [a, b] of selfTests) {
    logger.info({ a, b, best: fuzzyScore(a, b) }, "FUZZY_SELF_TEST");
  }
}

function reconcile(prior: LedgerRow[], current: LedgerRow[]): ReconciliationRow[] {
  const results: ReconciliationRow[] = [];
  // Track by index — ledger codes are often empty or duplicated, so code-based
  // tracking would silently block valid candidates sharing the same code.
  const matchedCurrentIdx = new Set<number>();

  for (const p of prior) {
    const exactIdx = p.ledgerCode !== ""
      ? current.findIndex((c, i) => c.ledgerCode === p.ledgerCode && !matchedCurrentIdx.has(i))
      : -1;

    if (exactIdx !== -1) {
      const exact = current[exactIdx];
      matchedCurrentIdx.add(exactIdx);
      const variance = exact.balance - p.balance;
      results.push({
        status: Math.abs(variance) < 0.005 ? "matched" : "mismatched",
        ledgerCode: p.ledgerCode,
        ledgerName: p.ledgerName,
        priorBalance: p.balance,
        currentBalance: exact.balance,
        variance,
        matchScore: 100,
        matchedWith: null,
        matchStrategy: "exact_code",
      });
      continue;
    }

    const exactNameIdx = current.findIndex(
      (c, i) => normalize(c.ledgerName) === normalize(p.ledgerName) && !matchedCurrentIdx.has(i)
    );
    if (exactNameIdx !== -1) {
      const exactName = current[exactNameIdx];
      matchedCurrentIdx.add(exactNameIdx);
      const variance = exactName.balance - p.balance;
      results.push({
        status: Math.abs(variance) < 0.005 ? "matched" : "mismatched",
        ledgerCode: p.ledgerCode,
        ledgerName: p.ledgerName,
        priorBalance: p.balance,
        currentBalance: exactName.balance,
        variance,
        matchScore: 100,
        matchedWith: exactName.ledgerCode !== p.ledgerCode ? exactName.ledgerCode : null,
        matchStrategy: "exact_name",
      });
      continue;
    }

    const candidates = current
      .map((c, i) => ({ c, i, ...fuzzyScore(p.ledgerName, c.ledgerName) }))
      .filter((x) => !matchedCurrentIdx.has(x.i) && x.score >= SCORE_REGROUP)
      .sort((a, b) => b.score - a.score);

    if (candidates.length > 0) {
      const best = candidates[0];
      matchedCurrentIdx.add(best.i);
      const variance = best.c.balance - p.balance;

      if (best.score >= SCORE_HIGH) {
        results.push({
          status: Math.abs(variance) < 0.005 ? "matched" : "mismatched",
          ledgerCode: p.ledgerCode,
          ledgerName: p.ledgerName,
          priorBalance: p.balance,
          currentBalance: best.c.balance,
          variance,
          matchScore: best.score,
          matchedWith: best.c.ledgerCode !== p.ledgerCode ? best.c.ledgerCode : null,
          matchStrategy: best.strategy,
        });
      } else {
        results.push({
          status: "possible_regroup",
          ledgerCode: p.ledgerCode,
          ledgerName: p.ledgerName,
          priorBalance: p.balance,
          currentBalance: best.c.balance,
          variance,
          matchScore: best.score,
          matchedWith: best.c.ledgerCode || best.c.ledgerName,
          matchStrategy: best.strategy,
        });
      }
    } else {
      // Log top-3 near-miss scores so we can diagnose low-score misses
      const topMisses = current
        .map((c, i) => ({ name: c.ledgerName, code: c.ledgerCode, i, ...fuzzyScore(p.ledgerName, c.ledgerName) }))
        .filter((x) => !matchedCurrentIdx.has(x.i))
        .sort((a, b) => b.score - a.score)
        .slice(0, 3)
        .map((x) => ({ name: x.name, score: x.score, strategy: x.strategy }));
      logger.info({ priorName: p.ledgerName, priorCode: p.ledgerCode, topMisses }, "UNMATCHED_ENTRY");
      results.push({
        status: "missing_current",
        ledgerCode: p.ledgerCode,
        ledgerName: p.ledgerName,
        priorBalance: p.balance,
        currentBalance: 0,
        variance: -p.balance,
        matchScore: null,
        matchedWith: null,
        matchStrategy: null,
      });
    }
  }

  for (const [i, c] of current.entries()) {
    if (!matchedCurrentIdx.has(i)) {
      results.push({
        status: "missing_prior",
        ledgerCode: c.ledgerCode,
        ledgerName: c.ledgerName,
        priorBalance: 0,
        currentBalance: c.balance,
        variance: c.balance,
        matchScore: null,
        matchedWith: null,
        matchStrategy: null,
      });
    }
  }

  return results;
}

// ─── PDF extraction placeholder ───────────────────────────────────────────────
// Architecture:
//   PDF Upload → parsePdf() → LedgerRow[]
//               → (same) reconcile() → Dashboard & Export
//
// TODO(pdf-ocr): Install a PDF-to-text library (e.g. pdf-parse or pdfjs-dist)
//   and replace the stub below with real text extraction.
// TODO(pdf-ocr): After text extraction, detect the table structure (likely
//   two-column: name | balance) and parse rows into LedgerRow[].
// TODO(pdf-ocr): If the PDF is scanned (no embedded text), integrate an OCR
//   engine (e.g. Tesseract via tesseract.js) to extract the image layer first.
function parsePdf(_buffer: Buffer): LedgerRow[] {
  // TODO(pdf-ocr): replace this stub with actual PDF → LedgerRow[] extraction.
  throw new Error("PDF_NOT_SUPPORTED");
}

router.post("/reconciliation/upload", upload.fields([
  { name: "priorYear", maxCount: 1 },
  { name: "currentYear", maxCount: 1 },
]), async (req, res): Promise<void> => {
  const files = req.files as Record<string, Express.Multer.File[]> | undefined;
  if (!files?.priorYear?.[0] || !files?.currentYear?.[0]) {
    res.status(400).json({ error: "Both priorYear and currentYear files are required." });
    return;
  }

  // Detect PDF uploads and return a clear, actionable error.
  // The multer limit is 20 MB so large scanned PDFs are blocked before this point.
  const priorMime  = files.priorYear[0].mimetype;
  const currentMime = files.currentYear[0].mimetype;
  if (priorMime === "application/pdf" || files.priorYear[0].originalname.toLowerCase().endsWith(".pdf")) {
    // TODO(pdf-ocr): swap parsePdf() in here once OCR is implemented.
    res.status(422).json({
      error: "PDF_NOT_SUPPORTED",
      message: "PDF upload received for Prior Year FS. Full OCR/text extraction is not yet implemented. Please export your signed FS to Excel (.xlsx) and re-upload.",
    });
    return;
  }
  if (currentMime === "application/pdf" || files.currentYear[0].originalname.toLowerCase().endsWith(".pdf")) {
    res.status(422).json({
      error: "PDF_NOT_SUPPORTED",
      message: "PDF upload received for Current Year TB. Please export to Excel (.xlsx) and re-upload.",
    });
    return;
  }

  try {
    const prior = parseWorkbook(files.priorYear[0].buffer);
    const current = parseWorkbook(files.currentYear[0].buffer);

    if (prior.length === 0) {
      res.status(400).json({ error: "Could not parse any rows from the Prior Year file. Check column headers." });
      return;
    }
    if (current.length === 0) {
      res.status(400).json({ error: "Could not parse any rows from the Current Year file. Check column headers." });
      return;
    }

    const rows = reconcile(prior, current);

    const summary: ReconciliationSummary = {
      totalPrior: prior.reduce((s, r) => s + r.balance, 0),
      totalCurrent: current.reduce((s, r) => s + r.balance, 0),
      matchedCount: rows.filter((r) => r.status === "matched").length,
      mismatchedCount: rows.filter((r) => r.status === "mismatched").length,
      missingCurrentCount: rows.filter((r) => r.status === "missing_current").length,
      missingPriorCount: rows.filter((r) => r.status === "missing_prior").length,
      possibleRegroupCount: rows.filter((r) => r.status === "possible_regroup").length,
      totalVariance: rows.reduce((s, r) => s + r.variance, 0),
    };

    lastResult = { summary, rows };
    req.log.info({ rowCount: rows.length }, "Reconciliation complete");
    res.json({ summary, rows });
  } catch (err) {
    req.log.error({ err }, "Error processing files");
    res.status(400).json({ error: "Failed to parse one or both files. Ensure they are valid Excel or CSV files." });
  }
});

router.get("/reconciliation/result", async (_req, res): Promise<void> => {
  if (!lastResult) {
    res.json({ summary: null, rows: [] });
    return;
  }
  res.json(lastResult);
});

router.get("/reconciliation/summary", async (_req, res): Promise<void> => {
  if (!lastResult) {
    res.json({
      totalPrior: 0,
      totalCurrent: 0,
      matchedCount: 0,
      mismatchedCount: 0,
      missingCurrentCount: 0,
      missingPriorCount: 0,
      possibleRegroupCount: 0,
      totalVariance: 0,
    });
    return;
  }
  res.json(lastResult.summary);
});

router.post("/reconciliation/export", async (req, res): Promise<void> => {
  const { rows, summary } = req.body as {
    rows: ReconciliationRow[];
    summary: ReconciliationSummary;
  };

  if (!rows || !summary) {
    res.status(400).json({ error: "rows and summary are required" });
    return;
  }

  try {
    const wb = XLSX.utils.book_new();

    const summaryData = [
      ["Audtech Reconciliation Report", ""],
      ["Generated", new Date().toLocaleString()],
      [""],
      ["SUMMARY", ""],
      ["Total Prior Year Balance", summary.totalPrior],
      ["Total Current Year Balance", summary.totalCurrent],
      ["Total Variance", summary.totalVariance],
      [""],
      ["Matched Ledgers", summary.matchedCount],
      ["Mismatched Balances", summary.mismatchedCount],
      ["Missing in Current Year", summary.missingCurrentCount],
      ["New in Current Year (Missing in Prior)", summary.missingPriorCount],
      ["Possible Regroupings / Reclassifications", summary.possibleRegroupCount],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary["!cols"] = [{ wch: 40 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

    const headers = [
      "Status",
      "Ledger Code",
      "Ledger Name",
      "Prior Year Balance",
      "Current Year Opening Balance",
      "Variance",
      "Match Score",
      "Matched With Code",
    ];
    const detailData = [
      headers,
      ...rows.map((r) => [
        r.status.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()),
        r.ledgerCode,
        r.ledgerName,
        r.priorBalance,
        r.currentBalance,
        r.variance,
        r.matchScore ?? "",
        r.matchedWith ?? "",
      ]),
    ];
    const wsDetail = XLSX.utils.aoa_to_sheet(detailData);
    wsDetail["!cols"] = [
      { wch: 22 }, { wch: 15 }, { wch: 40 },
      { wch: 22 }, { wch: 28 }, { wch: 16 },
      { wch: 14 }, { wch: 20 },
    ];
    XLSX.utils.book_append_sheet(wb, wsDetail, "Reconciliation Detail");

    for (const status of ["matched", "mismatched", "missing_current", "missing_prior", "possible_regroup"] as ReconciliationStatus[]) {
      const filtered = rows.filter((r) => r.status === status);
      if (filtered.length === 0) continue;
      const sheetData = [headers, ...filtered.map((r) => [
        r.status.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()),
        r.ledgerCode, r.ledgerName, r.priorBalance,
        r.currentBalance, r.variance,
        r.matchScore ?? "", r.matchedWith ?? "",
      ])];
      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      ws["!cols"] = wsDetail["!cols"];
      const label = status.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
      XLSX.utils.book_append_sheet(wb, ws, label.slice(0, 31));
    }

    const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="audtech-reconciliation-${Date.now()}.xlsx"`);
    res.send(buffer);
  } catch (err) {
    logger.error({ err }, "Export failed");
    res.status(500).json({ error: "Export failed" });
  }
});

export default router;
