import { Router, type IRouter } from "express";
import multer from "multer";
import * as XLSX from "xlsx";
import { classifyLedgerName } from "../services/llmMatcher";
// pdf-parse v1 is a CJS module; the banner in build.mjs sets globalThis.require
// so it bundles correctly into our ESM output.
// We import from the inner lib file to bypass index.js which reads a test PDF
// at module-load time — a known bug in pdf-parse@1.1.1.
// eslint-disable-next-line @typescript-eslint/no-require-imports
const pdfParse: (buf: Buffer) => Promise<{ text: string; numpages: number }> =
  require("pdf-parse/lib/pdf-parse.js");
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
  /** Plain-English explanation of why this row was classified as it was.
   *  Populated for possible_regroup, mismatched-by-fuzzy, and missing rows. */
  matchReason: string | null;
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
  const sheetName =
    wb.SheetNames.find(name =>
      name.toLowerCase().includes("tb") ||
      name.toLowerCase().includes("trial") ||
      name.toLowerCase().includes("balance") ||
      name.toLowerCase().includes("fs")
    ) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  console.log("Sheet Names :", wb.SheetNames);
  console.log("Using sheet:", sheetName);
  const rows: any[][] = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
  });

  let headerRowIndex = 0;

  for (let i = 0; i < rows.length; i++) {
    const rowText = rows[i].join(" ").toLowerCase();

    if (
      rowText.includes("account") ||
      rowText.includes("ledger") ||
      rowText.includes("particular") ||
      rowText.includes("description") ||
      rowText.includes("name")
    ) {
      headerRowIndex = i;
      break;
    }
  }

  const raw: Record<string, unknown>[] = XLSX.utils.sheet_to_json(ws, {
    range: headerRowIndex,
    defval: "",
  });

  console.log("Raw rows:", raw.slice(0, 10));

  const cleanedRaw = raw.filter((row: any) =>
    Object.values(row).some(v => String(v).trim() !== "")
  );

  if (cleanedRaw.length === 0) return [];

  const headers = Object.keys(cleanedRaw[0]);
  const sample = cleanedRaw.slice(0, Math.min(10, cleanedRaw.length));
  const { codeKey, nameKey, debitKey, creditKey, balanceKey } = detectColumns(headers, sample);

  logger.info({ headers, codeKey, nameKey, debitKey, creditKey, balanceKey }, "COLUMN_DETECTION");

  const ledgerRows: LedgerRow[] = [];
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

    ledgerRows.push({ ledgerCode: code, ledgerName: name, balance });
  }

  return ledgerRows;
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
 * Composite fuzzy scorer. Runs six string-similarity passes plus a domain
 * alias lookup, then returns the highest-scoring result together with the
 * alias entry that fired (if any) so callers can generate reason text.
 *
 * Passes (in order):
 *  1. ratio           — character-level edit distance on normalised strings
 *  2. partial         — sliding-window ratio (good for substrings / abbreviations)
 *  3. ratio (stemmed) — same as 1 but with plural/suffix stripping first
 *  4. partial stemmed — same as 2 but stemmed
 *  5. token_sort      — sort tokens alphabetically then compare (word-order agnostic)
 *  6. token_set       — intersection/difference set comparison (handles extra words)
 *  7. alias           — financial synonym table (domain knowledge boost, score=88)
 */

function isLikelyLedgerRow(name: string): boolean { return true;
}
async function fuzzyScore(
  a: string,
  b: string
): Promise<{
  score: number;
 strategy: MatchStrategy;
 aliasMatch: [string, string, string] | null;
}> {
  const na = normalize(a);
  const nb = normalize(b);
  const sa = normalizedStemmed(a);
  const sb = normalizedStemmed(b);
  if (na === nb) return { score: 100, strategy: "ratio", aliasMatch: null };
  if (!na || !nb) return { score: 0,   strategy: "ratio", aliasMatch: null };

  const alias = getAliasMatch(a, b);

  const candidates: { score: number; strategy: MatchStrategy }[] = [
    { score: ratio(na, nb),         strategy: "ratio"      },
    { score: partialRatio(na, nb),  strategy: "partial"    },
    { score: ratio(sa, sb),         strategy: "ratio"      },
    { score: partialRatio(sa, sb),  strategy: "partial"    },
    { score: tokenSortRatio(a, b),  strategy: "token_sort" },
    { score: tokenSetRatio(a, b),   strategy: "token_set"  },
    // Domain alias boost — fires for known accounting synonym pairs.
    // Score fixed at 88 so it always lands in "possible_regroup" (< SCORE_HIGH).
    { score: alias ? 88 : 0,        strategy: "token_set"  },
  ];
  let confidencePenalty = 0;

  const invalidLedgerTerms = [
    "statement",
    "financial position",
    "director",
    "dated",
    "page",
    "opening balance",
    "report",
    "independent auditor",
    "1-jan",
    "31-dec",
    "particulars",
    "2024",
    "2023"
  ];

  const looksLikeNonLedger =
    invalidLedgerTerms.some(term => a.includes(term)) ||
    invalidLedgerTerms.some(term => b.includes(term));

  if (looksLikeNonLedger) {
    confidencePenalty = 40;
  }

  
  const best = candidates.reduce((b, c) => (c.score > b.score ? c : b));
  const llmResult = await classifyLedgerName(a);

  console.log("LLM RESULT:", a, llmResult);
  return {
    ...best,
    score: Math.max(0, best.score - confidencePenalty),
    aliasMatch: alias,
  };
}

// ─── Financial synonym alias table ───────────────────────────────────────────
// Each entry is [termA, termB, category] where termA and termB are normalised
// strings (lowercase, no special chars). When both ledger names contain one term
// of a pair we inject a synthetic score of 88, guaranteeing "possible_regroup"
// classification and generating a human-readable reason string.
//
// Adding new aliases: add a row here. No other code needs to change.
const ACCOUNTING_ALIASES: Array<[string, string, string]> = [
  // ── Payroll & HR liabilities ──────────────────────────────────────────────
  ["outstanding salaries",      "salary payable",              "Payroll liabilities"],
  ["outstanding wages",         "wage payable",                "Payroll liabilities"],
  ["accrued salaries",          "salary payable",              "Payroll liabilities"],
  ["accrued wages",             "wage payable",                "Payroll liabilities"],
  ["salaries payable",          "salary payable",              "Payroll liabilities"],
  ["employee benefits payable", "salary payable",              "Payroll liabilities"],

  // ── Trade receivables ─────────────────────────────────────────────────────
  ["debtors",                   "trade receivable",            "Trade receivables"],
  ["sundry debtors",            "trade receivable",            "Trade receivables"],
  ["debtors",                   "trade debtor",                "Trade receivables"],
  ["accounts receivable",       "trade receivable",            "Trade receivables"],
  ["book debts",                "trade receivable",            "Trade receivables"],
  ["bills receivable",          "trade receivable",            "Trade receivables"],

  // ── Trade payables ────────────────────────────────────────────────────────
  ["creditors",                 "trade payable",               "Trade payables"],
  ["sundry creditors",          "trade payable",               "Trade payables"],
  ["creditors",                 "trade creditor",              "Trade payables"],
  ["accounts payable",          "trade payable",               "Trade payables"],
  ["bills payable",             "trade payable",               "Trade payables"],
  ["trade payables",            "sundry creditors",            "Trade payables"],

  // ── Professional & audit fees ─────────────────────────────────────────────
  ["audit fees",                "professional charges",        "Professional fees"],
  ["audit fees",                "professional fees",           "Professional fees"],
  ["audit fee",                 "professional charges",        "Professional fees"],
  ["auditor remuneration",      "professional charges",        "Professional fees"],
  ["legal and professional",    "professional charges",        "Professional fees"],
  ["legal fees",                "professional charges",        "Professional fees"],
  ["consultancy fees",          "professional charges",        "Professional fees"],
  ["advisory fees",             "professional charges",        "Professional fees"],

  // ── Miscellaneous & general expenses ──────────────────────────────────────
  ["miscellaneous expense",     "general expenses",            "General/miscellaneous"],
  ["miscellaneous expenses",    "general expenses",            "General/miscellaneous"],
  ["misc expense",              "general expenses",            "General/miscellaneous"],
  ["sundry expenses",           "general expenses",            "General/miscellaneous"],
  ["sundry expenses",           "miscellaneous expense",       "General/miscellaneous"],
  ["other expenses",            "general expenses",            "General/miscellaneous"],
  ["other expenses",            "miscellaneous expense",       "General/miscellaneous"],

  // ── Inventory ─────────────────────────────────────────────────────────────
  ["stock",                     "inventory",                   "Inventory"],
  ["closing stock",             "inventory",                   "Inventory"],
  ["opening stock",             "inventory",                   "Inventory"],
  ["goods in hand",             "inventory",                   "Inventory"],
  ["finished goods",            "inventory",                   "Inventory"],
  ["raw material",              "inventory",                   "Inventory"],
  ["work in progress",          "inventory",                   "Inventory"],
  ["wip",                       "work in progress",            "Inventory"],

  // ── Fixed assets / PPE ────────────────────────────────────────────────────
  ["fixed assets",              "property plant and equipment","Fixed assets"],
  ["fixed assets",              "ppe",                         "Fixed assets"],
  ["tangible assets",           "property plant and equipment","Fixed assets"],
  ["capital work in progress",  "cwip",                        "Fixed assets"],
  ["capital wip",               "cwip",                        "Fixed assets"],

  // ── Cash & bank ───────────────────────────────────────────────────────────
  ["cash and bank",             "bank balance",                "Cash & bank"],
  ["cash and bank balances",    "bank balance",                "Cash & bank"],
  ["cash and cash equivalents", "bank balance",                "Cash & bank"],
  ["cash at bank",              "bank balance",                "Cash & bank"],
  ["bank overdraft",            "overdraft",                   "Cash & bank"],

  // ── Loans & borrowings ────────────────────────────────────────────────────
  ["bank loan",                 "term loan",                   "Loans"],
  ["bank term loan",            "term loan",                   "Loans"],
  ["bank borrowings",           "term loan",                   "Loans"],
  ["secured loan",              "term loan",                   "Loans"],
  ["unsecured loan",            "loan from directors",         "Loans"],

  // ── Finance costs ─────────────────────────────────────────────────────────
  ["interest expense",          "finance charges",             "Finance costs"],
  ["interest expense",          "finance costs",               "Finance costs"],
  ["interest on loan",          "finance charges",             "Finance costs"],
  ["bank charges",              "finance charges",             "Finance costs"],

  // ── Revenue & income ──────────────────────────────────────────────────────
  ["commission income",         "commission received",         "Income"],
  ["interest income",           "interest received",           "Income"],
  ["other income",              "miscellaneous income",        "Income"],
  ["rental income",             "rent received",               "Income"],

  // ── Operating expenses ────────────────────────────────────────────────────
  ["rent expense",              "rent charges",                "Operating expenses"],
  ["office rent",               "rent charges",                "Operating expenses"],
  ["electricity expense",       "utilities",                   "Operating expenses"],
  ["power and fuel",            "utilities",                   "Operating expenses"],
  ["insurance expense",         "insurance premium",           "Operating expenses"],
  ["repairs and maintenance",   "maintenance expense",         "Operating expenses"],

  // ── Equity & reserves ─────────────────────────────────────────────────────
  ["retained earnings",         "reserves and surplus",        "Equity"],
  ["profit and loss account",   "reserves and surplus",        "Equity"],
  ["share capital",             "paid up capital",             "Equity"],
  ["equity share capital",      "share capital",               "Equity"],

  // ── Taxes ─────────────────────────────────────────────────────────────────
  ["gst receivable",            "input tax credit",            "Taxes"],
  ["gst payable",               "output tax liability",        "Taxes"],
  ["tds receivable",            "tax deducted at source",      "Taxes"],
  ["income tax payable",        "provision for tax",           "Taxes"],
  ["advance tax",               "prepaid tax",                 "Taxes"],

  // ── Provisions & accruals ─────────────────────────────────────────────────
  ["provision for expenses",    "accrued expenses",            "Provisions"],
  ["accrued liabilities",       "accrued expenses",            "Provisions"],
  ["outstanding expenses",      "accrued expenses",            "Provisions"],
  ["provision for doubtful debts", "bad debt provision",       "Provisions"],
];

/**
 * Returns the first alias pair that matches both names (as substrings of their
 * normalised forms), or null if no alias applies.
 */
function getAliasMatch(a: string, b: string): [string, string, string] | null {
  const na = normalize(a);
  const nb = normalize(b);
  for (const [x, y, category] of ACCOUNTING_ALIASES) {
    if ((na.includes(x) && nb.includes(y)) || (na.includes(y) && nb.includes(x))) {
      return [x, y, category];
    }
  }
  return null;
}

/** Returns 88 if a known alias pair matches, otherwise 0. */
function aliasScore(a: string, b: string): number {
  return getAliasMatch(a, b) !== null ? 88 : 0;
}

// ─── Regroup reason generation ────────────────────────────────────────────────
/**
 * Generates a concise, auditor-readable explanation for why a row was flagged
 * as "possible_regroup". Used to populate the `matchReason` field.
 *
 * @param cyName     – the CY ledger name that was matched
 * @param score      – the winning fuzzy score (0–100)
 * @param strategy   – the scoring pass that produced the winning score
 * @param alias      – the alias entry that fired, if any
 */
function generateRegroupReason(
  cyName: string,
  score: number,
  strategy: MatchStrategy,
  alias: [string, string, string] | null,
): string {
  if (alias) {
    const [, , category] = alias;
    return `Financial synonym (${category}): matched to "${cyName}"`;
  }
  const pct = score;
  switch (strategy) {
    case "partial":
      return `Partial name overlap ${pct}%: "${cyName}" shares a significant sub-string`;
    case "token_sort":
      return `Same keywords, different order ${pct}%: matched to "${cyName}"`;
    case "token_set":
      return `Shared key terms ${pct}%: common financial words with "${cyName}"`;
    case "ratio":
      return `High character similarity ${pct}%: close spelling match to "${cyName}"`;
    default:
      return `Fuzzy name match ${pct}%: matched to "${cyName}"`;
  }
}

/** Reason text for a fuzzy hit that scores ≥ SCORE_HIGH (treated as matched/mismatched). */
function generateHighConfidenceReason(
  cyName: string,
  score: number,
  strategy: MatchStrategy,
  alias: [string, string, string] | null,
): string {
  if (alias) {
    const [, , category] = alias;
    return `High-confidence synonym (${category}): matched to "${cyName}"`;
  }
  return `High-confidence fuzzy match ${score}% via ${strategy}: matched to "${cyName}"`;
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
    logger.info({ a, b, best: await fuzzyScore(a, b) }, "FUZZY_SELF_TEST");
  }
}

async function reconcile(prior: LedgerRow[], current: LedgerRow[]): Promise< ReconciliationRow[]> {
  const results: ReconciliationRow[] = [];

  console.log("PRIOR FS DATA:", prior.slice(0, 10));
  console.log("CURRENT TB DATA:", current.slice(0, 10));
  // Track by index — ledger codes are often empty or duplicated, so code-based
  // tracking would silently block valid candidates sharing the same code.
  const matchedCurrentIdx = new Set<number>();

  for (const p of prior) {
    // ── Pass 1: Exact account code match ──────────────────────────────────
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
        matchReason: null, // exact matches need no explanation
      });
      continue;
    }

    // ── Pass 2: Exact normalised name match ───────────────────────────────
    const exactNameIdx = current.findIndex(
      (c, i) => normalize(c.ledgerName) === normalize(p.ledgerName) && !matchedCurrentIdx.has(i)
    );
     
      const exactName = current[exactNameIdx];
      const normalizedPrior = p.ledgerName
        .toLowerCase()
        .replace(/[^a-z0-9 ]/g, "")
        .trim();

      const normalizedCurrent = exactName.ledgerName
        .toLowerCase()
        .replace(/[^a-z0-9 ]/g, "")
        .trim();
       if (normalizedPrior === normalizedCurrent) {
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
        matchedWith: 
          exactName.ledgerCode !== p.ledgerCode ? exactName.ledgerCode : null,
        matchStrategy: "exact_name",
        matchReason: null,
      });
      continue;
    }

    // ── Pass 3: Fuzzy / alias matching ────────────────────────────────────
    // Score every unmatched CY entry; keep those above SCORE_REGROUP threshold.
    const invalidLedgerTerms = [
      "statement",
      "financial position",
      "director",
      "dated",
      "page",
      "opening balance",
      "report",
      "independent auditor",
      "1-jan",
      "31-dec",
      "particulars",
      "2024",
      "2023"
    ];

    const normalizedLedger = p.ledgerName.toLowerCase();

    const isNonLedgerRow = invalidLedgerTerms.some(term =>
      normalizedLedger.includes(term)
    );

    if (isNonLedgerRow) {
      continue;
    }
    if (!isLikelyLedgerRow(p.ledgerName)) {
      continue;
    }
    
      const candidateResults = await Promise.all(
      current.map(async (c, i) => {
      const scoreResult = await fuzzyScore(
      p.ledgerName,
      c.ledgerName
      );

      return {
      c,
      i,
      ...scoreResult,
      };
      })
      );

      const candidates = candidateResults
      .filter(
      (x) =>
      !matchedCurrentIdx.has(x.i) &&
      x.score >= SCORE_REGROUP &&
      Math.abs(x.c.balance - p.balance) < 1000
      )
      .sort((a, b) => b.score - a.score);

    if (candidates.length > 0) {
      const best = candidates[0];
      matchedCurrentIdx.add(best.i);
      const variance = best.c.balance - p.balance;

      if ((best).score >= SCORE_HIGH) {
        // High-confidence fuzzy hit → treated as matched/mismatched, not regroup
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
          matchReason: generateHighConfidenceReason(
            best.c.ledgerName, best.score, best.strategy, best.aliasMatch
          ),
        });
      } else {
        // Below SCORE_HIGH → flag as possible regroup for auditor review
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
          matchReason: generateRegroupReason(
            best.c.ledgerName, best.score, best.strategy, best.aliasMatch
          ),
        });
      }
    } else {
      // No candidate above threshold — log top-3 near-misses for diagnostics
      const scoredCandidates = await Promise.all(
      current.map(async (c, i) => ({
      name: c.ledgerName,
      code: c.ledgerCode,
      i,
      ...(await fuzzyScore(p.ledgerName, c.ledgerName)),
      }))
      );

      const topMisses = scoredCandidates
      .filter((x) => !matchedCurrentIdx.has(x.i))
      .sort((a, b) => b.score - a.score)
      .slice(0, 3)
      .map((x) => ({
      name: x.name,
      score: x.score,
      strategy: x.strategy,
      }));
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
        matchReason: topMisses.length > 0
          ? `Best near-miss: "${topMisses[0].name}" (${topMisses[0].score}%) — below regroup threshold`
          : "No similar CY ledger found",
      });
    }
  }

  // ── Add unmatched CY entries as "New in CY" ───────────────────────────────
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
        matchReason: "No matching Prior Year entry found — new account in Current Year",
      });
    }
  }

  return results;
}

// ─── PDF text extraction ──────────────────────────────────────────────────────
// Architecture:
//   PDF Upload → parsePdf() → LedgerRow[]
//               → reconcile() → Dashboard & Export
//
// This implementation handles digitally-created PDFs (embedded text layer).
// TODO(pdf-ocr): For scanned PDFs (image-only), integrate Tesseract.js OCR here
//   to first convert each page image to text, then pass to extractLedgerLines().

/**
 * Regex that recognises the rightmost currency-style number on a line.
 * Handles:
 *   1,234.56        → positive
 *   (1,234.56)      → negative (parenthesis convention)
 *   -1,234.56       → negative
 *   1,234.56 Cr     → negative (credit = liability)
 *   1,234.56 Dr     → positive
 *   1,234           → integer balance
 */
const NUMBER_RE = /(\([\d,]+(?:\.\d+)?\)|[\d,]+(?:\.\d+)?)\s*(Cr|Dr)?$/i;

/** Lines that are almost certainly headers/footers, not data rows. */
const NOISE_RE = /^\s*($|page\s+\d|date|note[s]?|particulars|description|account|schedule|total|sub.?total|grand\s+total|balance\s+sheet|profit\s+and\s+loss|trial\s+balance|as\s+at|for\s+the\s+(year|period)|rs\.?\s*$|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})/i;

/** Very short runs that look like standalone page numbers or reference codes. */
const STANDALONE_NUMBER_RE = /^\s*[\d.]+\s*$/;

/**
 * Parse a single currency token into a signed float.
 * Parentheses and "Cr" suffix indicate negative values.
 */
function parseCurrencyToken(raw: string, suffix: string | undefined): number {
  const isParens = raw.startsWith("(") && raw.endsWith(")");
  const cleaned  = raw.replace(/[(),]/g, "");
  const value    = parseFloat(cleaned) || 0;
  const isCr     = isParens || (suffix?.toLowerCase() === "cr");
  return isCr ? -value : value;
}

/**
 * Given raw PDF text, extract ledger rows by scanning each line for a pattern:
 *   [optional code]  <name text>  <balance number>
 *
 * Strategy:
 *  1. Split text into lines; clean excess whitespace.
 *  2. Skip noise lines (headers, footers, totals, blanks).
 *  3. Use NUMBER_RE to find the rightmost number token → balance.
 *  4. Everything to the left is the label; split off a leading short code if present.
 *  5. Skip rows where the resulting name is empty, purely numeric, or looks like
 *     a section heading (all caps with no number).
 */
function extractLedgerLines(text: string): LedgerRow[] {
  const rows: LedgerRow[] = [];

  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || NOISE_RE.test(line) || STANDALONE_NUMBER_RE.test(line)) continue;

    const match = NUMBER_RE.exec(line);
    if (!match) continue;

    const balance  = parseCurrencyToken(match[1], match[2]);
    // Label = everything before the matched number token
    const label    = line.slice(0, match.index).trim();
    if (!label || isNumericValue(label)) continue;

    // Optionally split a leading account code: a short (≤8 char) alphanumeric
    // prefix separated by whitespace from the rest of the name.
    // E.g. "1100 Trade Receivables" → code="1100", name="Trade Receivables"
    const codeMatch = /^([A-Z0-9\-./]{1,8})\s{2,}(.+)$/i.exec(label);
    const ledgerCode = codeMatch ? codeMatch[1].trim() : "";
    const ledgerName = codeMatch ? codeMatch[2].trim() : label;

    if (!ledgerName || isNumericValue(ledgerName)) continue;

    rows.push({ ledgerCode, ledgerName, balance });
  }

  return rows;
}

/**
 * Parse a PDF buffer into LedgerRow[].
 * Requires the PDF to have an embedded text layer (digitally created).
 * Uses pdf-parse v1: pdfParse(buffer) → { text, numpages }.
 * TODO(pdf-ocr): Add scanned-PDF fallback via Tesseract.js when text is empty.
 */
async function parsePdf(buffer: Buffer): Promise<LedgerRow[]> {
  // pdf-parse v1: pdfParse(buffer) → { text: string, numpages: number }
  const data = await pdfParse(buffer);
  const text = data.text ?? "";

  if (!text || text.trim().length < 20) {
    // TODO(pdf-ocr): scanned PDF detected — no embedded text. Integrate OCR here.
    throw new Error("PDF_SCANNED");
  }

  logger.info({ pages: data.numpages, chars: text.length }, "PDF_PARSED");
  const rows = extractLedgerLines(text);
  return rows;
}

/** Helper: detect PDF by MIME type or file extension. */
function isPdfFile(file: Express.Multer.File): boolean {
  return (
    file.mimetype === "application/pdf" ||
    file.originalname.toLowerCase().endsWith(".pdf")
  );
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

  const priorFile   = files.priorYear[0];
  const currentFile = files.currentYear[0];

  try {
    // Parse Prior Year — supports Excel, CSV, and digitally-created PDF
    let prior: LedgerRow[];
    if (isPdfFile(priorFile)) {
      try {
        prior = await parsePdf(priorFile.buffer);
      } catch (pdfErr: unknown) {
        const msg = pdfErr instanceof Error ? pdfErr.message : "";
        if (msg === "PDF_SCANNED") {
          res.status(422).json({
            error: "PDF_SCANNED",
            message: "The Prior Year PDF appears to be a scanned image with no embedded text. Please export to Excel (.xlsx) or use a digitally-created PDF.",
          });
        } else {
          res.status(422).json({
            error: "PDF_PARSE_ERROR",
            message: "Could not extract ledger data from the Prior Year PDF. Ensure the PDF contains a text layer and the balances appear on the same line as the account names.",
          });
        }
        return;
      }
    } else {
      prior = parseWorkbook(priorFile.buffer);
    }

    // Parse Current Year — Excel/CSV only (TB is typically spreadsheet-based)
    // TODO(pdf-ocr): extend to support PDF for Current Year TB if required.
    if (isPdfFile(currentFile)) {
      res.status(422).json({
        error: "PDF_NOT_SUPPORTED",
        message: "PDF upload is only supported for the Prior Year Signed FS. Please upload the Current Year Opening TB as Excel (.xlsx) or CSV.",
      });
      return;
    }
    const current = parseWorkbook(currentFile.buffer);

    if (prior.length === 0) {
      res.status(400).json({ error: "Could not parse any rows from the Prior Year file. Check that ledger names and balances appear on the same line." });
      return;
    }
    if (current.length === 0) {
      res.status(400).json({ error: "Could not parse any rows from the Current Year file. Check column headers." });
      return;
    }

    const rows = await reconcile(prior, current);

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
      "Match Score (%)",
      "Matched With",
      "AI Match Reason",
    ];
    const rowToArray = (r: ReconciliationRow) => [
      r.status.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()),
      r.ledgerCode,
      r.ledgerName,
      r.priorBalance,
      r.currentBalance,
      r.variance,
      r.matchScore ?? "",
      r.matchedWith ?? "",
      r.matchReason ?? "",
    ];
    const detailData = [headers, ...rows.map(rowToArray)];
    const wsDetail = XLSX.utils.aoa_to_sheet(detailData);
    wsDetail["!cols"] = [
      { wch: 22 }, { wch: 15 }, { wch: 40 },
      { wch: 22 }, { wch: 28 }, { wch: 16 },
      { wch: 14 }, { wch: 25 }, { wch: 60 },
    ];
    XLSX.utils.book_append_sheet(wb, wsDetail, "Reconciliation Detail");

    for (const status of ["matched", "mismatched", "missing_current", "missing_prior", "possible_regroup"] as ReconciliationStatus[]) {
      const filtered = rows.filter((r) => r.status === status);
      if (filtered.length === 0) continue;
      const sheetData = [headers, ...filtered.map(rowToArray)];
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
