import { useState, useCallback, useRef, useMemo } from "react";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Legend } from "recharts";
import { Upload, FileSpreadsheet, CheckCircle, AlertCircle, XCircle, HelpCircle, ArrowRightLeft, Download, RefreshCw, ChevronUp, ChevronDown } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { cn } from "@/lib/utils";

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

interface ReconciliationResult {
  summary: ReconciliationSummary;
  rows: ReconciliationRow[];
}

interface ManualOverride {
  cyLedgerName: string;
  cyLedgerCode: string;
  cyBalance: number;
  sourceRowIndex: number;
}

const STATUS_CONFIG: Record<ReconciliationStatus, { label: string; color: string; icon: React.ReactNode; bg: string }> = {
  matched: { label: "Matched", color: "#16a34a", bg: "bg-green-50 text-green-700 border-green-200", icon: <CheckCircle className="w-3.5 h-3.5" /> },
  mismatched: { label: "Mismatched", color: "#d97706", bg: "bg-amber-50 text-amber-700 border-amber-200", icon: <AlertCircle className="w-3.5 h-3.5" /> },
  missing_current: { label: "Missing in CY", color: "#dc2626", bg: "bg-red-50 text-red-700 border-red-200", icon: <XCircle className="w-3.5 h-3.5" /> },
  missing_prior: { label: "New in CY", color: "#2563eb", bg: "bg-blue-50 text-blue-700 border-blue-200", icon: <ArrowRightLeft className="w-3.5 h-3.5" /> },
  possible_regroup: { label: "Possible Regroup", color: "#7c3aed", bg: "bg-violet-50 text-violet-700 border-violet-200", icon: <HelpCircle className="w-3.5 h-3.5" /> },
};

const CHART_COLORS = ["#16a34a", "#d97706", "#dc2626", "#2563eb", "#7c3aed"];

function fmt(n: number) {
  return new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
}

function StatusBadge({ status }: { status: ReconciliationStatus }) {
  const cfg = STATUS_CONFIG[status];
  return (
    <span className={cn("inline-flex items-center gap-1 px-2 py-0.5 rounded-full border text-xs font-medium", cfg.bg)}>
      {cfg.icon}
      {cfg.label}
    </span>
  );
}

const STRATEGY_LABELS: Record<MatchStrategy, string> = {
  exact_code:  "Exact Code",
  exact_name:  "Exact Name",
  ratio:       "Char Similarity",
  partial:     "Partial Match",
  token_sort:  "Word Order",
  token_set:   "Word Set",
};

function ConfidenceBand({ score, strategy }: { score: number | null; strategy: MatchStrategy | null }) {
  if (score === null) return <span className="text-muted-foreground text-xs">—</span>;

  const pct = score;
  const barColor =
    pct === 100 ? "bg-green-500" :
    pct >= 95   ? "bg-green-400" :
    pct >= 80   ? "bg-amber-400" :
                  "bg-violet-400";

  const textColor =
    pct === 100 ? "text-green-700" :
    pct >= 95   ? "text-green-600" :
    pct >= 80   ? "text-amber-600" :
                  "text-violet-600";

  return (
    <div className="flex flex-col items-end gap-1 min-w-[96px]">
      <div className="flex items-center gap-1.5 w-full justify-end">
        <div className="flex-1 h-1.5 rounded-full bg-muted overflow-hidden">
          <div className={cn("h-full rounded-full transition-all", barColor)} style={{ width: `${pct}%` }} />
        </div>
        <span className={cn("text-xs font-semibold tabular-nums w-8 text-right", textColor)}>
          {pct}%
        </span>
      </div>
      {strategy && strategy !== "exact_code" && strategy !== "exact_name" && (
        <span className="text-[10px] text-muted-foreground font-medium leading-none">
          {STRATEGY_LABELS[strategy]}
        </span>
      )}
    </div>
  );
}

function FileDropZone({
  label, file, onChange, acceptPdf = false,
}: {
  label: string;
  file: File | null;
  onChange: (f: File) => void;
  acceptPdf?: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) onChange(f);
  }, [onChange]);

  const isPdf = file?.name.toLowerCase().endsWith(".pdf");
  const accept = acceptPdf ? ".xlsx,.xls,.csv,.pdf" : ".xlsx,.xls,.csv";
  const hint   = acceptPdf ? "Excel, CSV, or PDF — click to browse" : "Drop Excel or CSV here, or click to browse";

  return (
    <div
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      className={cn(
        "border-2 border-dashed rounded-xl p-6 cursor-pointer transition-all flex flex-col items-center gap-3 text-center",
        dragging ? "border-primary bg-primary/5" : "border-border hover:border-primary/50 hover:bg-accent/30",
        file && !isPdf && "border-green-400 bg-green-50/30",
        isPdf && "border-amber-400 bg-amber-50/30"
      )}
    >
      <input ref={inputRef} type="file" className="hidden" accept={accept}
        onChange={(e) => e.target.files?.[0] && onChange(e.target.files[0])} />
      {file ? (
        <>
          <FileSpreadsheet className={cn("w-8 h-8", isPdf ? "text-amber-500" : "text-green-600")} />
          <div>
            <p className={cn("text-sm font-medium", isPdf ? "text-amber-700" : "text-green-700")}>{file.name}</p>
            <p className="text-xs text-muted-foreground mt-0.5">{(file.size / 1024).toFixed(1)} KB</p>
            {isPdf && (
              <p className="text-xs text-amber-600 mt-1 font-medium">
                PDF detected — text layer will be extracted automatically
              </p>
            )}
          </div>
        </>
      ) : (
        <>
          <Upload className="w-8 h-8 text-muted-foreground" />
          <div>
            <p className="text-sm font-semibold">{label}</p>
            <p className="text-xs text-muted-foreground mt-0.5">{hint}</p>
          </div>
        </>
      )}
    </div>
  );
}

type SortField = "ledgerCode" | "ledgerName" | "priorBalance" | "currentBalance" | "variance" | "matchScore" | "status";
type SortDir = "asc" | "desc";

function SortIcon({ field, sortField, sortDir }: { field: SortField; sortField: SortField; sortDir: SortDir }) {
  if (field !== sortField) return <ChevronUp className="w-3 h-3 opacity-20" />;
  return sortDir === "asc" ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />;
}

export default function Reconciliation() {
  const [priorFile, setPriorFile] = useState<File | null>(null);
  const [currentFile, setCurrentFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<ReconciliationResult | null>(null);
  const [filterStatus, setFilterStatus] = useState<ReconciliationStatus | "all">("all");
  const [search, setSearch] = useState("");
  const [sortField, setSortField] = useState<SortField>("status");
  const [sortDir, setSortDir] = useState<SortDir>("asc");
  const [exporting, setExporting] = useState(false);
  const [overrides, setOverrides] = useState<Record<number, ManualOverride>>({});
  const [pendingSelection, setPendingSelection] = useState<Record<number, string>>({});

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortDir(d => d === "asc" ? "desc" : "asc");
    } else {
      setSortField(field);
      setSortDir("asc");
    }
  };

  const handleUpload = async () => {
    if (!priorFile || !currentFile) return;
    setLoading(true);
    setError(null);
    try {
      const fd = new FormData();
      fd.append("priorYear", priorFile);
      fd.append("currentYear", currentFile);
      const res = await fetch("/api/reconciliation/upload", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) {
        // PDF_NOT_SUPPORTED returns a human-readable `message` field
        throw new Error(data.message || data.error || "Upload failed");
      }
      setResult(data);
      setFilterStatus("all");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Unknown error");
    } finally {
      setLoading(false);
    }
  };

  const handleExport = async () => {
    if (!result) return;
    setExporting(true);
    try {
      const res = await fetch("/api/reconciliation/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rows: effectiveRows, summary: effectiveSummary }),
      });
      if (!res.ok) throw new Error("Export failed");
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `audtech-reconciliation-${Date.now()}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      setError("Export failed. Please try again.");
    } finally {
      setExporting(false);
    }
  };

  const handleReset = () => {
    setPriorFile(null);
    setCurrentFile(null);
    setResult(null);
    setError(null);
    setFilterStatus("all");
    setSearch("");
    setOverrides({});
    setPendingSelection({});
  };

  // Indices of missing_prior rows consumed by a confirmed override
  const consumedPriorIndices = useMemo(
    () => new Set(Object.values(overrides).map(o => o.sourceRowIndex)),
    [overrides]
  );

  // Available CY rows for manual mapping (missing_prior rows not yet used)
  const availableCyRows = useMemo(() => {
    if (!result) return [];
    return result.rows
      .map((r, i) => ({ ...r, _idx: i }))
      .filter(r => r.status === "missing_prior" && !consumedPriorIndices.has(r._idx));
  }, [result, consumedPriorIndices]);

  // Effective rows: apply confirmed overrides, hide consumed missing_prior rows
  const effectiveRows = useMemo<(ReconciliationRow & { _originalIdx: number })[]>(() => {
    if (!result) return [];
    return result.rows
      .map((row, i) => {
        const ov = overrides[i];
        if (ov) {
          const variance = ov.cyBalance - row.priorBalance;
          return {
            ...row,
            _originalIdx: i,
            status: "possible_regroup" as ReconciliationStatus,
            currentBalance: ov.cyBalance,
            variance,
            matchedWith: ov.cyLedgerCode ? `${ov.cyLedgerCode} – ${ov.cyLedgerName}` : ov.cyLedgerName,
            matchScore: null,
            matchStrategy: null,
          };
        }
        return { ...row, _originalIdx: i };
      })
      .filter((_, i) => !consumedPriorIndices.has(i));
  }, [result, overrides, consumedPriorIndices]);

  const effectiveSummary = useMemo(() => {
    if (!result) return null;
    return {
      ...result.summary,
      matchedCount: effectiveRows.filter(r => r.status === "matched").length,
      mismatchedCount: effectiveRows.filter(r => r.status === "mismatched").length,
      missingCurrentCount: effectiveRows.filter(r => r.status === "missing_current").length,
      missingPriorCount: effectiveRows.filter(r => r.status === "missing_prior").length,
      possibleRegroupCount: effectiveRows.filter(r => r.status === "possible_regroup").length,
      totalVariance: effectiveRows.reduce((s, r) => s + r.variance, 0),
    };
  }, [effectiveRows, result]);

  const applyOverride = (rowIdx: number) => {
    const selected = pendingSelection[rowIdx];
    if (!selected) return;
    const cyRow = result?.rows.find((r, i) => String(i) === selected);
    if (!cyRow) return;
    const sourceRowIndex = result!.rows.indexOf(cyRow);
    setOverrides(prev => ({
      ...prev,
      [rowIdx]: {
        cyLedgerName: cyRow.ledgerName,
        cyLedgerCode: cyRow.ledgerCode,
        cyBalance: cyRow.currentBalance,
        sourceRowIndex,
      },
    }));
    setPendingSelection(prev => { const n = { ...prev }; delete n[rowIdx]; return n; });
  };

  const removeOverride = (rowIdx: number) => {
    setOverrides(prev => { const n = { ...prev }; delete n[rowIdx]; return n; });
  };

  const filteredRows = effectiveRows
    .filter(r => filterStatus === "all" || r.status === filterStatus)
    .filter(r => {
      if (!search) return true;
      const q = search.toLowerCase();
      return r.ledgerCode.toLowerCase().includes(q) || r.ledgerName.toLowerCase().includes(q);
    })
    .sort((a, b) => {
      let av: string | number = a[sortField as keyof ReconciliationRow] ?? "";
      let bv: string | number = b[sortField as keyof ReconciliationRow] ?? "";
      if (typeof av === "number" && typeof bv === "number") {
        return sortDir === "asc" ? av - bv : bv - av;
      }
      av = String(av); bv = String(bv);
      return sortDir === "asc" ? av.localeCompare(bv) : bv.localeCompare(av);
    });

  const pieData = effectiveSummary ? [
    { name: "Matched", value: effectiveSummary.matchedCount, color: CHART_COLORS[0] },
    { name: "Mismatched", value: effectiveSummary.mismatchedCount, color: CHART_COLORS[1] },
    { name: "Missing in CY", value: effectiveSummary.missingCurrentCount, color: CHART_COLORS[2] },
    { name: "New in CY", value: effectiveSummary.missingPriorCount, color: CHART_COLORS[3] },
    { name: "Possible Regroup", value: effectiveSummary.possibleRegroupCount, color: CHART_COLORS[4] },
  ].filter(d => d.value > 0) : [];

  const barData = effectiveSummary ? [
    { name: "Prior Year", value: effectiveSummary.totalPrior },
    { name: "Current Year", value: effectiveSummary.totalCurrent },
    { name: "Variance", value: effectiveSummary.totalVariance },
  ] : [];

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b bg-card px-6 py-4">
        <div className="max-w-[1400px] mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-lg bg-primary flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-primary-foreground" />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight">Audtech Reconciliation</h1>
              <p className="text-xs text-muted-foreground">Prior Year FS vs Current Year Opening TB</p>
            </div>
          </div>
          {result && (
            <div className="flex items-center gap-2">
              <Button variant="outline" size="sm" onClick={handleReset}>
                <RefreshCw className="w-4 h-4 mr-1.5" />
                New Reconciliation
              </Button>
              <Button size="sm" onClick={handleExport} disabled={exporting}>
                <Download className="w-4 h-4 mr-1.5" />
                {exporting ? "Exporting..." : "Export to Excel"}
              </Button>
            </div>
          )}
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto px-6 py-8 space-y-8">
        {!result ? (
          <div className="max-w-2xl mx-auto space-y-6">
            <div className="text-center space-y-2">
              <h2 className="text-2xl font-bold">Upload Financial Files</h2>
              <p className="text-muted-foreground text-sm">
                Upload your Prior Year Signed Financial Statements (Excel, CSV, or PDF) and Current Year Opening Trial Balance (Excel or CSV).
                The tool will automatically reconcile and identify matches, mismatches, and possible regroupings.
              </p>
            </div>

            <Card>
              <CardContent className="p-6 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-sm font-medium flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-full bg-muted text-xs flex items-center justify-center font-bold">1</span>
                      Prior Year Signed FS
                    </label>
                    <FileDropZone label="Prior Year Signed FS" file={priorFile} onChange={setPriorFile} acceptPdf />
                  </div>
                  <div className="space-y-2">
                    <label className="text-sm font-medium flex items-center gap-1.5">
                      <span className="w-5 h-5 rounded-full bg-muted text-xs flex items-center justify-center font-bold">2</span>
                      Current Year Opening TB
                    </label>
                    <FileDropZone label="Current Year Opening TB" file={currentFile} onChange={setCurrentFile} />
                  </div>
                </div>

                {error && (
                  <div className="p-3 rounded-lg bg-destructive/10 border border-destructive/20 text-destructive text-sm flex items-start gap-2">
                    <XCircle className="w-4 h-4 mt-0.5 shrink-0" />
                    {error}
                  </div>
                )}

                <Button className="w-full" size="lg" disabled={!priorFile || !currentFile || loading} onClick={handleUpload}>
                  {loading ? (
                    <>
                      <RefreshCw className="w-4 h-4 mr-2 animate-spin" />
                      Processing files...
                    </>
                  ) : (
                    <>
                      <ArrowRightLeft className="w-4 h-4 mr-2" />
                      Run Reconciliation
                    </>
                  )}
                </Button>

                <div className="text-xs text-muted-foreground space-y-1 pt-2 border-t">
                  <p className="font-medium">Expected column headers (any order, auto-detected):</p>
                  <p>Code/Account No, Name/Description, Debit + Credit (or Balance/Amount). Net balance = Debit − Credit.</p>
                </div>
              </CardContent>
            </Card>
          </div>
        ) : (
          <div className="space-y-6">
            {/* Summary KPI Cards */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <Card className="border-l-4 border-l-green-500">
                <CardContent className="p-4">
                  <p className="text-xs text-muted-foreground uppercase tracking-wide font-medium">Matched</p>
                  <p className="text-3xl font-bold text-green-600 mt-1">{effectiveSummary!.matchedCount}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">Exact ledger matches</p>
                </CardContent>
              </Card>
              <Card className="border-l-4 border-l-amber-500">
                <CardContent className="p-4">
                  <p className="text-xs text-muted-foreground uppercase tracking-wide font-medium">Mismatched</p>
                  <p className="text-3xl font-bold text-amber-600 mt-1">{effectiveSummary!.mismatchedCount}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">Balance differences</p>
                </CardContent>
              </Card>
              <Card className="border-l-4 border-l-red-500">
                <CardContent className="p-4">
                  <p className="text-xs text-muted-foreground uppercase tracking-wide font-medium">Missing in CY</p>
                  <p className="text-3xl font-bold text-red-600 mt-1">{effectiveSummary!.missingCurrentCount}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">In PY, not in CY</p>
                </CardContent>
              </Card>
              <Card className="border-l-4 border-l-violet-500">
                <CardContent className="p-4">
                  <p className="text-xs text-muted-foreground uppercase tracking-wide font-medium">Possible Regroup</p>
                  <p className="text-3xl font-bold text-violet-600 mt-1">{effectiveSummary!.possibleRegroupCount}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">Fuzzy-matched ledgers</p>
                </CardContent>
              </Card>
            </div>

            {/* Balance Summary */}
            <div className="grid grid-cols-3 gap-4">
              <Card>
                <CardContent className="p-4">
                  <p className="text-xs text-muted-foreground uppercase tracking-wide font-medium">Total PY Balance</p>
                  <p className="text-xl font-bold mt-1">{fmt(effectiveSummary!.totalPrior)}</p>
                </CardContent>
              </Card>
              <Card>
                <CardContent className="p-4">
                  <p className="text-xs text-muted-foreground uppercase tracking-wide font-medium">Total CY Balance</p>
                  <p className="text-xl font-bold mt-1">{fmt(effectiveSummary!.totalCurrent)}</p>
                </CardContent>
              </Card>
              <Card className={cn("border-l-4", Math.abs(effectiveSummary!.totalVariance) < 0.01 ? "border-l-green-500" : "border-l-amber-500")}>
                <CardContent className="p-4">
                  <p className="text-xs text-muted-foreground uppercase tracking-wide font-medium">Total Variance</p>
                  <p className={cn("text-xl font-bold mt-1", Math.abs(effectiveSummary!.totalVariance) < 0.01 ? "text-green-600" : "text-amber-600")}>
                    {fmt(effectiveSummary!.totalVariance)}
                  </p>
                </CardContent>
              </Card>
            </div>

            {/* Charts */}
            <div className="grid grid-cols-2 gap-6">
              <Card>
                <CardHeader className="pb-2">
                  <CardTitle className="text-sm font-semibold">Reconciliation Breakdown</CardTitle>
                </CardHeader>
                <CardContent>
                  <ResponsiveContainer width="100%" height={220}>
                    <PieChart>
                      <Pie data={pieData} cx="50%" cy="50%" innerRadius={55} outerRadius={85} paddingAngle={2} dataKey="value">
                        {pieData.map((entry, i) => (
                          <Cell key={i} fill={entry.color} />
                        ))}
                      </Pie>
                      <Tooltip formatter={(v: number) => [v, "Count"]} />
                      <Legend formatter={(val) => <span className="text-xs">{val}</span>} />
                    </PieChart>
                  </ResponsiveContainer>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="pb-2">
                  <CardTitle className="text-sm font-semibold">Balance Comparison</CardTitle>
                </CardHeader>
                <CardContent>
                  <ResponsiveContainer width="100%" height={220}>
                    <BarChart data={barData} margin={{ top: 5, right: 10, bottom: 5, left: 10 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                      <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                      <YAxis tick={{ fontSize: 11 }} tickFormatter={(v) => {
                        if (Math.abs(v) >= 1000000) return `${(v / 1000000).toFixed(1)}M`;
                        if (Math.abs(v) >= 1000) return `${(v / 1000).toFixed(0)}K`;
                        return String(v);
                      }} />
                      <Tooltip formatter={(v: number) => [fmt(v), "Amount"]} />
                      <Bar dataKey="value" fill="hsl(var(--chart-1))" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </CardContent>
              </Card>
            </div>

            {/* Table Section */}
            <Card>
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between gap-4 flex-wrap">
                  <CardTitle className="text-sm font-semibold">Reconciliation Detail ({filteredRows.length} of {effectiveRows.length})</CardTitle>
                  <div className="flex items-center gap-2 flex-wrap">
                    <input
                      type="search"
                      placeholder="Search ledger..."
                      value={search}
                      onChange={e => setSearch(e.target.value)}
                      className="h-8 px-3 text-sm border rounded-md bg-background w-48 focus:outline-none focus:ring-2 focus:ring-ring"
                    />
                    <div className="flex gap-1 flex-wrap">
                      {(["all", "matched", "mismatched", "missing_current", "missing_prior", "possible_regroup"] as const).map((s) => (
                        <button
                          key={s}
                          onClick={() => setFilterStatus(s)}
                          className={cn(
                            "px-2.5 py-1 rounded-full text-xs font-medium border transition-all",
                            filterStatus === s
                              ? "bg-primary text-primary-foreground border-primary"
                              : "bg-background text-muted-foreground border-border hover:border-primary/40"
                          )}
                        >
                          {s === "all" ? "All" : STATUS_CONFIG[s].label}
                          {s !== "all" && ` (${effectiveSummary![`${s === "missing_current" ? "missingCurrentCount" : s === "missing_prior" ? "missingPriorCount" : s === "possible_regroup" ? "possibleRegroupCount" : `${s}Count`}` as keyof ReconciliationSummary]})`}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="p-0">
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b bg-muted/40">
                        {([
                          ["status", "Status"],
                          ["ledgerCode", "Code"],
                          ["ledgerName", "Ledger Name"],
                          ["priorBalance", "Prior Year Balance"],
                          ["currentBalance", "CY Opening Balance"],
                          ["variance", "Variance"],
                          ["matchScore", "Confidence"],
                        ] as [SortField, string][]).map(([field, label]) => (
                          <th
                            key={field}
                            onClick={() => handleSort(field)}
                            className="px-4 py-2.5 text-left text-xs font-medium text-muted-foreground uppercase tracking-wide cursor-pointer hover:text-foreground select-none"
                          >
                            <span className="flex items-center gap-1">
                              {label}
                              <SortIcon field={field} sortField={sortField} sortDir={sortDir} />
                            </span>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="px-4 py-10 text-center text-muted-foreground text-sm">
                            No results found
                          </td>
                        </tr>
                      ) : filteredRows.map((row, i) => {
                        const origIdx = row._originalIdx;
                        const isOverridden = !!overrides[origIdx];
                        const isMissingCurrent = row.status === "missing_current";
                        const hasCyOptions = availableCyRows.length > 0;
                        return (
                        <tr key={i} className={cn("border-b transition-colors hover:bg-muted/30", i % 2 === 0 ? "" : "bg-muted/10")}>
                          <td className="px-4 py-2.5"><StatusBadge status={row.status} /></td>
                          <td className="px-4 py-2.5 font-mono text-xs text-muted-foreground">{row.ledgerCode || "—"}</td>
                          <td className="px-4 py-2.5 font-medium max-w-xs">
                            <span className="truncate block" title={row.ledgerName}>{row.ledgerName}</span>
                            {row.matchedWith && !isOverridden && (
                              <span className="text-xs text-muted-foreground font-normal">→ {row.matchedWith}</span>
                            )}
                            {isOverridden && (
                              <span className="text-xs text-violet-600 font-normal flex items-center gap-1">
                                → {overrides[origIdx].cyLedgerName}
                                <button
                                  onClick={() => removeOverride(origIdx)}
                                  className="text-muted-foreground hover:text-destructive ml-1"
                                  title="Remove manual mapping"
                                >✕</button>
                              </span>
                            )}
                            {isMissingCurrent && !isOverridden && (
                              <div className="flex items-center gap-1 mt-1.5">
                                <select
                                  value={pendingSelection[origIdx] ?? ""}
                                  onChange={e => setPendingSelection(prev => ({ ...prev, [origIdx]: e.target.value }))}
                                  className="text-xs border rounded px-1.5 py-0.5 bg-background text-foreground focus:outline-none focus:ring-1 focus:ring-violet-400 max-w-[180px]"
                                  disabled={!hasCyOptions}
                                >
                                  <option value="">{hasCyOptions ? "Map to CY ledger…" : "No unmatched CY entries"}</option>
                                  {availableCyRows.map((cy) => (
                                    <option key={cy._idx} value={String(cy._idx)}>
                                      {cy.ledgerCode ? `${cy.ledgerCode} – ` : ""}{cy.ledgerName}
                                    </option>
                                  ))}
                                </select>
                                {pendingSelection[origIdx] && (
                                  <button
                                    onClick={() => applyOverride(origIdx)}
                                    className="text-xs px-2 py-0.5 rounded bg-violet-600 text-white hover:bg-violet-700 transition-colors font-medium"
                                  >
                                    Apply
                                  </button>
                                )}
                              </div>
                            )}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono text-sm">{fmt(row.priorBalance)}</td>
                          <td className="px-4 py-2.5 text-right font-mono text-sm">{fmt(row.currentBalance)}</td>
                          <td className={cn("px-4 py-2.5 text-right font-mono text-sm font-medium",
                            Math.abs(row.variance) < 0.005 ? "text-green-600" :
                              row.variance > 0 ? "text-blue-600" : "text-red-600"
                          )}>
                            {row.variance > 0 ? "+" : ""}{fmt(row.variance)}
                          </td>
                          <td className="px-4 py-2.5">
                            <ConfidenceBand score={row.matchScore} strategy={row.matchStrategy} />
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </CardContent>
            </Card>

            {effectiveSummary!.missingPriorCount > 0 && (
              <Card className="border-blue-200 bg-blue-50/30">
                <CardContent className="p-4 flex items-start gap-3">
                  <ArrowRightLeft className="w-4 h-4 text-blue-600 mt-0.5 shrink-0" />
                  <div className="text-sm text-blue-800">
                    <span className="font-semibold">{effectiveSummary!.missingPriorCount} new ledger{effectiveSummary!.missingPriorCount > 1 ? "s" : ""}</span> found in the Current Year Opening TB that have no corresponding entry in the Prior Year FS.
                    These may represent new accounts opened during the year.
                  </div>
                </CardContent>
              </Card>
            )}
            {Object.keys(overrides).length > 0 && (
              <Card className="border-violet-200 bg-violet-50/30">
                <CardContent className="p-4 flex items-start gap-3">
                  <HelpCircle className="w-4 h-4 text-violet-600 mt-0.5 shrink-0" />
                  <div className="text-sm text-violet-800">
                    <span className="font-semibold">{Object.keys(overrides).length} manual mapping{Object.keys(overrides).length > 1 ? "s" : ""}</span> applied by auditor. These overrides will be included in the exported report.
                  </div>
                </CardContent>
              </Card>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
