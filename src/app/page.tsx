"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AlertTriangle,
  ArrowRight,
  Check,
  Download,
  FileDown,
  FileSpreadsheet,
  RefreshCw,
  ShieldCheck,
  Trash2,
  UploadCloud,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";
import { reconcileCsv, type Reconciliation } from "@/lib/reconcile";

const inr = (v: number) =>
  `₹${Number(v || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const signed = (v: number) => `${v > 0 ? "+" : v < 0 ? "−" : ""}${inr(Math.abs(v))}`;
const COLLAPSED_EXCEPTIONS = 12;

function FileBox({
  kind,
  files,
  onChange,
}: {
  kind: "vendor" | "client";
  files: File[];
  onChange: (f: File[]) => void;
}) {
  const [dragging, setDragging] = useState(false);
  const isVendor = kind === "vendor";
  const acceptsFile = (f: File | undefined | null) => !!f && /\.(csv|xls|xlsx)$/i.test(f.name);
  const acceptFiles = (list: FileList | null) => {
    if (!list) return;
    const accepted = [...list].filter(acceptsFile);
    if (accepted.length) onChange(accepted);
  };
  return (
    <label
      className={`file-box ${files.length ? "is-ready" : ""} ${dragging ? "is-dragging" : ""}`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        acceptFiles(e.dataTransfer.files);
      }}
    >
      <input type="file" multiple accept=".csv,.xls,.xlsx,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={(e) => acceptFiles(e.target.files)} />
      <div className="file-icon">
        <FileSpreadsheet size={22} />
      </div>
      <div className="file-copy">
        <strong>{isVendor ? "Vendor payment advice (CSV / XLS)" : "Client transaction report (CSV / XLS)"}</strong>
        <span>
          {files.length === 0
            ? isVendor
              ? "Drop the vendor's payment advice .csv or .xls here — multiple files = batch"
              : "Drop your own booking ledger .csv or .xls here — same order pairs them"
            : files.length === 1
              ? files[0].name
              : `${files.length} files — paired in order`}
        </span>
      </div>
      <div className="file-action">{files.length ? <Check size={18} /> : <UploadCloud size={18} />}</div>
    </label>
  );
}

async function readAsCsv(file: File) {
  if (/\.csv$/i.test(file.name)) return file.text();
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The client spreadsheet does not contain a worksheet.");
  return XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName], { blankrows: true });
}

function CheckRow({ item }: { item: Reconciliation["checks"][number] }) {
  const matched = item.status === "matched";
  const info = item.status === "info";
  return (
    <div className="check-row">
      <div className={`status-dot ${matched ? "ok" : info ? "info" : "review"}`}>
        {matched ? <Check size={13} /> : info ? <ShieldCheck size={13} /> : <AlertTriangle size={13} />}
      </div>
      <div>
        <strong>{item.label}</strong>
        <p>{item.detail}</p>
      </div>
      <span className={`status-label ${matched ? "ok-text" : info ? "info-text" : "review-text"}`}>{item.status}</span>
    </div>
  );
}

function ResultSkeleton() {
  return (
    <div className="skeleton-state" aria-label="Reading files">
      <div className="skeleton-bar w-40" />
      <div className="skeleton-hero" />
      <div className="skeleton-bar" />
      <div className="skeleton-bar w-70" />
      <div className="skeleton-bar w-55" />
      <div className="skeleton-bar" />
      <p className="skeleton-note">Matching both files by PNR and building the settlement bridge…</p>
    </div>
  );
}

export default function Home() {
  const [vendor, setVendor] = useState<File[]>([]);
  const [client, setClient] = useState<File[]>([]);
  const [results, setResults] = useState<Reconciliation[]>([]);
  const [activeResult, setActiveResult] = useState(0);
  const [receivedInput, setReceivedInput] = useState("");
  const [featuredInput, setFeaturedInput] = useState("");
  const [sponsorInput, setSponsorInput] = useState("");
  const [shareCharges, setShareCharges] = useState(false);
  const [auditOpen, setAuditOpen] = useState(false);
  const [auditQuery, setAuditQuery] = useState("");
  const result = results[activeResult] ?? null;
  const [loading, setLoading] = useState(false);
  const [downloaded, setDownloaded] = useState(false);
  const [showAll, setShowAll] = useState(false);
  const [error, setError] = useState("");
  const [savedReportCount, setSavedReportCount] = useState(0);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [history, setHistory] = useState<Array<{ savedAt: string; result: Reconciliation }>>([]);
  const resultRef = useRef<HTMLDivElement | null>(null);
  const inFlight = useRef(0);
  const restoringHistory = useRef(false);
  const num = (s: string) => {
    const v = Number(s.replace(/[,\s]/g, ""));
    return Number.isFinite(v) && v > 0 ? v : undefined;
  };
  const reconcileOptions = useMemo(
    () => ({ received: num(receivedInput), featuredCard: num(featuredInput), sponsorship: num(sponsorInput), shareCharges }),
    [receivedInput, featuredInput, sponsorInput, shareCharges],
  );

  const canRun = Boolean(vendor.length && client.length);

  const reconcile = useCallback(async () => {
    if (!vendor.length || !client.length) return;
    const generation = ++inFlight.current;
    setLoading(true);
    setError("");
    setResults([]);
    setActiveResult(0);
    setShowAll(false);
    try {
      // Batch mode: files are paired in order — vendor[0] with client[0], etc.
      const pairs = Math.min(vendor.length, client.length);
      if (vendor.length !== client.length) setError(`${Math.max(vendor.length, client.length) - pairs} unpaired file(s) skipped — upload one client file per vendor advice.`);
      const out: Reconciliation[] = [];
      for (let i = 0; i < pairs; i++) {
        const [vendorText, clientText] = await Promise.all([readAsCsv(vendor[i]), readAsCsv(client[i])]);
        out.push(reconcileCsv(vendorText, clientText, vendor[i].name, client[i].name, reconcileOptions));
      }
      // A newer file selection supersedes this run; only the latest pair wins.
      if (generation !== inFlight.current) return;
      setResults(out);
    } catch (e) {
      if (generation === inFlight.current) setError(e instanceof Error ? e.message : "Unable to reconcile these files.");
    } finally {
      if (generation === inFlight.current) setLoading(false);
    }
  }, [vendor, client, reconcileOptions]);

  // Auto-run once both sides have files; reconcile legitimately sets loading/result state.
  useEffect(() => {
    if (vendor.length && client.length) {
      // eslint-disable-next-line react-hooks/set-state-in-effect
      void reconcile();
    }
  }, [vendor, client, reconcile]);

  useEffect(() => {
    if (!results.length) return;
    if (restoringHistory.current) {
      restoringHistory.current = false;
      return;
    }
    const key = "humsafarai:reconciliation-reports";
    let saved: Array<{ savedAt: string; result: Reconciliation }> = [];
    try {
      saved = JSON.parse(window.localStorage.getItem(key) || "[]") as Array<{ savedAt: string; result: Reconciliation }>;
    } catch {
      saved = [];
    }
    const stamp = new Date().toISOString();
    // The joined audit grid (thousands of rows) is regenerable from the
    // source files and always complete in the XLS/CSV export — keep it out
    // of localStorage so history fits the ~5MB quota.
    const slim = (r: Reconciliation): Reconciliation => ({ ...r, audit: [] });
    saved.unshift(...results.map((result) => ({ savedAt: stamp, result: slim(result) })));
    const kept = saved.slice(0, 10);
    try {
      window.localStorage.setItem(key, JSON.stringify(kept));
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setSavedReportCount(kept.length);
    } catch {
      // Still over quota: drop the oldest entries until it fits.
      let room = kept.length - 1;
      while (room > 0) {
        try {
          window.localStorage.setItem(key, JSON.stringify(kept.slice(0, room)));
          setSavedReportCount(room);
          return;
        } catch {
          room--;
        }
      }
      window.localStorage.removeItem(key);
      setSavedReportCount(0);
    }
  }, [results]);

  useEffect(() => {
    try {
      const saved = JSON.parse(window.localStorage.getItem("humsafarai:reconciliation-reports") || "[]") as Array<{ savedAt: string; result: Reconciliation }>;
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setSavedReportCount(saved.length);
    } catch {
      setSavedReportCount(0);
    }
  }, []);

  const openHistory = () => {
    try {
      const saved = JSON.parse(window.localStorage.getItem("humsafarai:reconciliation-reports") || "[]") as Array<{ savedAt: string; result: Reconciliation }>;
      setHistory(saved);
    } catch {
      setHistory([]);
    }
    setHistoryOpen(true);
  };

  const restoreHistoryItem = (saved: { savedAt: string; result: Reconciliation }) => {
    restoringHistory.current = true;
    setResults([saved.result]);
    setActiveResult(0);
    setError("");
    setDownloaded(false);
    setHistoryOpen(false);
  };

  useEffect(() => {
    if (result && window.matchMedia("(max-width: 900px)").matches) {
      resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }, [result]);

  const resetOutput = () => {
    setResults([]);
    setActiveResult(0);
    setError("");
    setDownloaded(false);
  };
  const chooseVendor = (files: File[]) => {
    setVendor(files);
    resetOutput();
  };
  const chooseClient = (files: File[]) => {
    setClient(files);
    resetOutput();
  };

  const reviewCount = useMemo(
    () =>
      result
        ? result.checks.filter((c) => c.status === "review").length +
          result.differences.filter((d) => d.status === "review").length
        : 0,
    [result],
  );
  const exceptionGroups = useMemo(() => {
    if (!result) return [] as { category: string; items: Reconciliation["exceptions"]; variance: number }[];
    const order = result.source.vendor === "RedBus"
      ? ["Booking", "Cancellation", "Bus cancelled", "Exceptional refund", "Assured"]
      : ["Booking", "Cancellation", "Exceptional refund", "Assured", "Bus cancelled"];
    const groups = new Map<string, Reconciliation["exceptions"]>();
    for (const item of result.exceptions) {
      const list = groups.get(item.category) ?? [];
      list.push(item);
      groups.set(item.category, list);
    }
    const categories = [...order.filter((c) => groups.has(c)), ...[...groups.keys()].filter((c) => !order.includes(c))];
    return categories.map((category) => {
      const items = groups.get(category)!;
      return { category, items, variance: Math.round(items.reduce((a, e) => a + e.variance, 0) * 100) / 100 };
    });
  }, [result]);

  // Repeat-offender intelligence: PNRs flagged in the most recent saved
  // report before this run.
  const previousReport = useMemo(
    () =>
      result
        ? history.find(
            (h) =>
              h.result !== result &&
              (h.result.source.invoiceNo !== result.source.invoiceNo || h.result.source.vendorName !== result.source.vendorName),
          )
        : undefined,
    [result, history],
  );
  const repeatPnrs = useMemo(() => {
    const set = new Set<string>();
    previousReport?.result.exceptions.forEach((e) => set.add(e.pnr));
    return set;
  }, [previousReport]);

  // Carried-forward DIFF from the previous saved report.
  const carriedForward = useMemo(() => {
    if (!previousReport) return null;
    return Math.abs(previousReport.result.receivable.difference) > 0.005
      ? { difference: previousReport.result.receivable.difference, invoice: previousReport.result.source.invoiceNo }
      : null;
  }, [previousReport]);

  const askVendor = (item: Reconciliation["exceptions"][number]) => {
    const v = result?.source.vendor ?? "the vendor";
    switch (item.issue) {
      case "In vendor advice, missing from client ledger":
        return `Ask ${v}: this PNR is settled on the advice but never reached ITS — share the ticket proof.`;
      case "In client ledger, missing from vendor advice":
        return `Ask ${v}: this PNR was booked in ITS but is absent from the advice — request inclusion in the next settlement.`;
      case "Cancelled in vendor advice, no cancellation in client ledger":
        return `Ask ${v}: a cancellation was deducted that ITS never recorded — confirm the cancel date and charge.`;
      case "Cancelled in client ledger, not in vendor advice":
        return `Ask ${v}: ITS cancelled this PNR but no reversal appears on the advice — escalate for refund.`;
      case "Deducted by vendor, no client equivalent":
        return `Ask ${v}: a ${item.category} deduction with no ITS record — request the reason and supporting document.`;
      case "Vendor reversed in full (bus cancelled); client kept a cancellation charge":
        return `The bus was cancelled, so ${v} reversed in full — reverse the kept charge on your ledger.`;
      default:
        return `Ask ${v} to justify the ${item.variance !== 0 ? "variance" : "entry"} on this PNR.`;
    }
  };

  const exceptionSectionTitle = (category: string) => {
    if (result?.source.vendor === "RedBus") {
      const map: Record<string, string> = {
        Booking: "Payment",
        Cancellation: "Cancellation",
        "Bus cancelled": "Bus Cancelled",
        "Exceptional refund": "Exceptional Refund",
        Assured: "Assured",
      };
      return map[category] ?? category;
    }
    const map: Record<string, string> = {
      Booking: "Booked tickets",
      Cancellation: "Cancellation",
      "Exceptional refund": "Exceptional refunds",
      Assured: "Assured",
      "Bus cancelled": "Bus cancelled",
    };
    return map[category] ?? category;
  };

  async function downloadReport() {
    if (!result) return;
    try {
      const response = await fetch("/api/reconcile/report", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(result),
      });
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.error || "The PDF report could not be generated.");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `humsafarAI-${result.source.invoiceNo.replace(/[^\w-]+/g, "-") || "report"}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
      setDownloaded(true);
    } catch (e) {
      setError(e instanceof Error ? e.message : "The PDF report could not be generated.");
    }
  }

  const fileStem = () => `humsafarAI-${result?.source.invoiceNo.replace(/[^\w-]+/g, "-") || "report"}`;


  function exportSheetRows() {
    if (!result) return { summary: [] as string[][], weekly: [] as string[][], differences: [] as string[][], checks: [] as string[][], exceptions: [] as string[][], audit: [] as string[][], periods: [] as string[][] };
    const s = result.source;
    const summary: string[][] = [
      ["HumsafarAI reconciliation"],
      ["Vendor", s.vendor],
      ["Vendor file", s.vendorName],
      ["Client file", s.clientName],
      ["Operator", s.operator],
      ["Invoice", s.invoiceNo],
      ["Vendor period", s.period],
      ["Client period", s.clientPeriod],
      [],
      ["Vendor final settlement", result.settlement.vendorFinal],
      ["Client ledger net", result.settlement.clientNet],
      ["Gap", result.settlement.gap],
      [],
      ["NET RECEIVABLE (manual formula)"],
      ...result.receivable.rows.map((r) => [r.label, r.amount] as (string | number)[]),
      ["Net receivable", result.receivable.receivable],
      ["Actually received", result.receivable.received],
      ["DIFF", result.receivable.difference],
      [],
      ["SETTLEMENT BRIDGE"],
      ...result.bridge.map((b) => [b.label, b.amount] as (string | number)[]),
    ] as string[][];
    const differences: string[][] = [
      ["Comparison", "Vendor", "Client", "Variance", "Status", "Explanation"],
      ...result.differences.map((d) => [d.label, d.vendorAmount, d.clientAmount, d.variance, d.status, d.explanation] as (string | number)[]),
    ] as string[][];
    const checks: string[][] = [
      ["Check", "Status", "Variance", "Detail"],
      ...result.checks.map((c) => [c.label, c.status, c.variance ?? "", c.detail] as (string | number)[]),
    ] as string[][];
    const carriedByPnr = new Map(result.audit.map((a) => [a.pnr, a.carried ? a.carriedWeek : ""]));
    const exceptions: string[][] = [
      ["PNR", "Category", "Issue", "Vendor", "Client", "Variance", "Prior week", "Detail", "Follow-up"],
      ...result.exceptions.map((e) => [e.pnr, e.category, e.issue, e.vendorAmount, e.clientAmount, e.variance, carriedByPnr.get(e.pnr) ?? "", e.detail, askVendor(e)] as (string | number)[]),
    ] as string[][];
    // Their FINAL tab: one weekly settlement row per advice/ITS pair.
    const w = result.weekly;
    const vname = result.source.vendor;
    const datedPeriod = /\d{2}-\d{2}-\d{4}/.test(result.source.period) ? result.source.period : result.source.clientPeriod;
    const periodParts = datedPeriod.split(/\s*(?:\u2192|\u2013|\u2014|\bto\b)\s*/i).map((s) => s.trim());
    const weekly: (string | number)[][] = [
      ["INV NO", "SETTAL DATE", "FROM", "TO", `${vname} BOOKING`, "ITS BOOKING", `${vname} COMM`, "ITS COMM", `${vname} Cancellation`, "Bus Cancelled", "Exceptional Refund", "ITS Cancellation", "ITS Cancellation Charges", `${vname} Cancellation Charges`, "Operator Offer", "GDS Commission", "SS partner comm w/ GST", "Reschedule", "SRP", "Sponsored Listing", "Dynamic charg", "TDS", "YourBus Charge", "SeatSeller Adjustments", "Recon Adjustments", "TOTAL", "Payment Received", "DIFF", "AS PER ITS", "PAYMNET DIFF"],
      [result.source.invoiceNo, result.source.paymentDate, periodParts[0] ?? "", periodParts[periodParts.length - 1] ?? "", w.vendorBooking, w.itsBooking, w.vendorComm, w.itsComm, w.vendorCancel, w.busCancelled, w.exceptional, w.itsCancel, w.itsCharges, w.vendorCharges, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, w.total, w.received, w.diff, w.asPerIts, w.paymentDiff],
      [],
      ["Note", "Adjustment columns (Operator Offer, GDS, SS comm, Reschedule, SRP, Sponsored, Dynamic, TDS, YourBus, SeatSeller/Recon adjustments) are hand-keyed on the manual sheet and have no source columns in the advice — they are 0 here; fill them from the RedBus partner portal if needed."],
    ];
    const audit: (string | number)[][] = [
      ["PNR", "Ticket", "Route", "DOJ", "Settle date", "Client row", "Vendor section", "Cancel type", "Rescheduled", "Channel",
        "Vendor fare", "ITS fare", "Fare diff", "Vendor comm", "ITS comm", "Comm diff", "Vendor offer", "ITS offer", "Offer diff",
        "Vendor cancel charges", "ITS cancel charges", "Charge diff", "Vendor net", "ITS net", "Variance", "Prior week", "Vendor note", "Status"],
      ...result.audit.map((a) => [
        a.pnr, a.ticket, a.route, a.doj, a.settleDate, a.clientType, a.vendorSection, a.cancelType, a.rescheduled ? "Yes" : "", a.channel,
        a.vendorFare, a.clientFare, a.fareDiff, a.vendorComm, a.clientComm, a.commDiff, a.vendorOffer, a.clientOffer, a.offerDiff,
        a.vendorCharge, a.clientCharge, a.chargeDiff, a.vendorAmount, a.clientAmount, a.variance, a.carried ? a.carriedWeek : "", a.note, a.status,
      ] as (string | number)[]),
    ];
    const periods: (string | number)[][] = [
      ["Period", "Booked", "Cancelled", "Service cancel", "Charges", "Exceptional", "Edge", "Receivable", "Tickets"],
      ...result.periods.map((p) => [p.label, p.booked, p.cancel, p.serviceCancel, p.charges, p.exceptional, p.edge, p.receivable, p.tickets] as (string | number)[]),
    ];
    return { summary, weekly, differences, checks, exceptions, audit, periods };
  }

  function downloadXls() {
    if (!result) return;
    const { summary, weekly, differences, checks, exceptions, audit, periods } = exportSheetRows();
    const wb = XLSX.utils.book_new();
    const sheet = (rows: (string | number)[][]) => {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = (rows[0] ?? []).map((_, i) => ({
        wch: Math.min(Math.max(10, ...rows.slice(0, 80).map((r) => String(r[i] ?? "").length)) + 2, 44),
      }));
      return ws;
    };
    // Three plain tabs: the weekly settlement row (their FINAL tab), the
    // per-PNR report answering every clarification column, and only the
    // PNRs that still need attention.
    XLSX.utils.book_append_sheet(wb, sheet((weekly as (string | number)[][]).concat([[""]], summary, [[""]], [["CONTROL CHECKS"]], checks)), "Weekly settlement");
    XLSX.utils.book_append_sheet(wb, sheet(audit), "PNR wise report");
    XLSX.utils.book_append_sheet(wb, sheet((exceptions as (string | number)[][]).concat([[""]], [["WHERE DIFFERENCES OCCUR"]], differences, [[""]], [["PERIODS BY JOURNEY DATE"]], periods)), "Tickets to review");
    XLSX.writeFile(wb, `${fileStem()}.xlsx`);
    setDownloaded(true);
  }

  function deleteHistoryItem(savedAt: string) {
    const key = "humsafarai:reconciliation-reports";
    const next = history.filter((h) => h.savedAt !== savedAt);
    setHistory(next);
    setSavedReportCount(next.length);
    window.localStorage.setItem(key, JSON.stringify(next));
  }

  function clearHistory() {
    setHistory([]);
    setSavedReportCount(0);
    window.localStorage.setItem("humsafarai:reconciliation-reports", "[]");
  }

  const step1Done = Boolean(vendor && client);
  const step2Done = Boolean(result);
  const steps = [
    { n: "01", label: "Upload files", state: step1Done ? "done" : "active" },
    { n: "02", label: "See the result", state: step2Done ? "done" : loading ? "active" : "idle" },
    { n: "03", label: "Download report", state: downloaded ? "done" : step2Done ? "active" : "idle" },
  ];

  return (
    <div className="app-shell">
      <header className="topnav">
        <div className="brand">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img className="brand-logo" src="/humsafarai-logo.png" alt="HumsafarAI logo" />
          <div>
            <strong>HumsafarAI</strong>
            <span>settlement control</span>
          </div>
        </div>
        <nav className="step-track" aria-label="Workflow">
          {steps.map((step, i) => (
            <div className={`track-step ${step.state}`} key={step.n}>
              <span className="track-dot">{step.state === "done" ? <Check size={11} /> : step.n}</span>
              <span className="track-label">{step.label}</span>
              {i < steps.length - 1 && <span className="track-line" />}
            </div>
          ))}
        </nav>
        <button className="history-btn" onClick={openHistory}>
          History ({savedReportCount})
        </button>
        <div className="topnav-badge">
          <span className="live-dot" /> local · {savedReportCount} report{savedReportCount === 1 ? "" : "s"} saved
        </div>
      </header>

      <section className="hero">
        <span className="eyebrow">WEEKLY SETTLEMENT RECONCILIATION</span>
        <h1>What should you receive — and what did you actually get?</h1>
        <p>Upload the vendor&apos;s weekly payment advice and your ITS report. We match every PNR and show where each rupee was deducted, and why.</p>
      </section>

      <section className="workspace">
        <div className="panel upload-panel">
          <div className="panel-head">
            <span className="kicker">01 · Upload files</span>
            <h2>Two files, one answer.</h2>
            <p className="lead">
              Upload the vendor&apos;s payment advice (RedBus / AbhiBus) and your ITS transaction report. The app matches every ticket by PNR
              and explains each deduction the vendor applied.
            </p>
          </div>
          <div className="file-stack">
            <FileBox kind="vendor" files={vendor} onChange={chooseVendor} />
            <FileBox kind="client" files={client} onChange={chooseClient} />
          </div>
          <details className="options-details">
            <summary>Optional inputs — bank amount received &amp; promo deductions</summary>
            <div className="options-grid">
              <label className="option-field">
                <span>Amount received (bank)</span>
                <input type="text" inputMode="decimal" placeholder="e.g. 786171" value={receivedInput} onChange={(e) => setReceivedInput(e.target.value)} onBlur={reconcile} />
              </label>
              <label className="option-field">
                <span>Featured card deduction</span>
                <input type="text" inputMode="decimal" placeholder="0" value={featuredInput} onChange={(e) => setFeaturedInput(e.target.value)} onBlur={reconcile} />
              </label>
              <label className="option-field">
                <span>Sponsorship deduction</span>
                <input type="text" inputMode="decimal" placeholder="0" value={sponsorInput} onChange={(e) => setSponsorInput(e.target.value)} onBlur={reconcile} />
              </label>
              <label className="option-field option-check">
                <input type="checkbox" checked={shareCharges} onChange={(e) => setShareCharges(e.target.checked)} />
                <span>Share cancellation charges 50% (AbhiBus)</span>
              </label>
            </div>
          </details>
          {error && <div className="error-box">{error}</div>}
          <div className="upload-actions">
            <button onClick={reconcile} disabled={!canRun || loading} className="run-btn">
              {loading ? (
                <>
                  <RefreshCw className="spin" size={17} /> Reading files…
                </>
              ) : (
                <>
                  Run reconciliation <ArrowRight size={17} />
                </>
              )}
            </button>
            {!canRun && <span className="hint">Upload both files — the reconciliation runs automatically.</span>}
            {canRun && loading && <span className="hint">Both files received — reconciling automatically.</span>}
          </div>
          <div className="panel-note">
            <ShieldCheck size={15} />
            <p>
              Files never leave your browser. Every total is tied back to a source row in the files.
            </p>
          </div>
        </div>

        <div className="panel result-panel" ref={resultRef}>
          <div className="result-header">
            <span className="kicker">02 · Result</span>
            <div className="result-actions">
              <span className="confidence">
                <span className="live-dot" />{" "}
                {result ? (reviewCount ? `${reviewCount} to review` : "ready to review") : loading ? "working" : "waiting"}
              </span>
              <div className="export-group">
                <button className="outline-btn" onClick={downloadReport} disabled={!result}>
                  <Download size={15} /> PDF
                </button>
                <button className="outline-btn" onClick={downloadXls} disabled={!result}>
                  <FileDown size={15} /> XLS
                </button>

              </div>
            </div>
          </div>
          {result ? (
            <div className="result-body">
              {results.length > 1 && (
                <div className="result-tabs">
                  {results.map((r, i) => (
                    <button key={i} className={`result-tab ${i === activeResult ? "active" : ""}`} onClick={() => setActiveResult(i)}>
                      {r.source.vendorName.replace(/\.\w+$/, "")}
                    </button>
                  ))}
                </div>
              )}
              <div className="identity">
                <div>
                  <span className="label">VENDOR</span>
                  <strong><span className={`vendor-chip ${result.source.vendor === "RedBus" ? "red" : "abhi"}`}>{result.source.vendor}</span></strong>
                </div>
                <div>
                  <span className="label">OPERATOR</span>
                  <strong>{result.source.operator}</strong>
                </div>
                <div>
                  <span className="label">INVOICE</span>
                  <strong>{result.source.invoiceNo}</strong>
                </div>
                <div>
                  <span className="label">PERIOD</span>
                  <strong>{result.source.period}</strong>
                </div>
              </div>
              <div className="hero-total">
                <div>
                  <span className="label">WHAT ITS SAYS YOU SHOULD RECEIVE</span>
                  <div className="amount expect">{inr(result.settlement.clientNet)}</div>
                  <span className="label" style={{ marginTop: 10 }}>WHAT {result.source.vendor.toUpperCase()} ACTUALLY PAYS</span>
                  <div className="amount">{inr(result.settlement.vendorFinal)}</div>
                  <p>
                    deduction gap <b>{signed(result.settlement.gap)}</b>
                  </p>
                </div>
                <div className={`verified ${reviewCount ? "needs-review" : ""}`}>
                  {reviewCount ? (
                    <>
                      <AlertTriangle size={15} /> {reviewCount} PNRs to review
                    </>
                  ) : (
                    <>
                      <Check size={17} /> fully reconciled
                    </>
                  )}
                </div>
              </div>
              <div className="bridge">
                <div className="bridge-title">
                  <span>WHERE THE MONEY WAS DEDUCTED</span>
                  <span className="mono">from the vendor&apos;s own advice</span>
                </div>
                {result.receivable.rows.map((item, i) => (
                  <div className={`bridge-row ${item.amount < 0 ? "deduction" : "addition"}`} key={i}>
                    <span>{item.label}</span>
                    <strong>{signed(item.amount)}</strong>
                  </div>
                ))}
                <div className="bridge-row base">
                  <span>Should have been received</span>
                  <strong>{inr(result.receivable.receivable)}</strong>
                </div>
                <div className="bridge-row deduction">
                  <span>Actually received</span>
                  <strong>{inr(result.receivable.received)}</strong>
                </div>
                <div className={`bridge-row ${Math.abs(result.receivable.difference) > 2 ? "deduction" : "final"}`}>
                  <span>DIFF</span>
                  <strong>{signed(result.receivable.difference)}</strong>
                </div>
                {carriedForward && (
                  <div className="bridge-row carry">
                    <span>Carried forward from {carriedForward.invoice || "previous report"}</span>
                    <strong>{signed(carriedForward.difference)}</strong>
                  </div>
                )}
                <details className="bridge-details">
                  <summary>Ledger walkthrough (ITS → vendor final)</summary>
                  {result.bridge.map((item, i) => (
                    <div className={`bridge-row ${item.kind}`} key={i}>
                      <span>{item.label}</span>
                      <strong>
                        {item.kind === "base" || item.kind === "final" ? inr(item.amount) : signed(item.amount)}
                      </strong>
                    </div>
                  ))}
                </details>
              </div>
              {result.periods.length > 0 && (
                <div className="bridge">
                  <div className="bridge-title">
                    <span>WEEKS COVERED</span>
                    <span className="mono">by journey date — incl. earlier weeks settled now</span>
                  </div>
                  <div className="period-grid">
                    {result.periods.map((p) => (
                      <div className="period-card" key={p.label}>
                        <strong>{p.label}</strong>
                        <div className="period-row"><span>Booked</span><b>{inr(p.booked)}</b></div>
                        <div className="period-row"><span>Cancelled</span><b>{inr(p.cancel)}</b></div>
                        {p.serviceCancel !== 0 && <div className="period-row"><span>Bus cancelled</span><b>{inr(p.serviceCancel)}</b></div>}
                        {p.charges !== 0 && <div className="period-row"><span>Charges</span><b>{inr(p.charges)}</b></div>}
                        {p.exceptional !== 0 && <div className="period-row"><span>Exceptional</span><b>{inr(p.exceptional)}</b></div>}
                        {p.edge !== 0 && <div className="period-row"><span>Edge</span><b>{inr(p.edge)}</b></div>}
                        <div className="period-total"><span>Receivable</span><b>{inr(p.receivable)}</b></div>
                        <small>{p.tickets.toLocaleString("en-IN")} tickets</small>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              <div className="summary-grid">
                <div className="metric-card">
                  <span className="label">BOOKINGS MATCHED</span>
                  <strong>
                    {result.coverage.matchedBookings.toLocaleString("en-IN")} /{" "}
                    {result.coverage.vendorBookings.toLocaleString("en-IN")}
                  </strong>
                  <small>
                    {result.coverage.vendorOnlyBookings} vendor-only · {result.coverage.clientOnlyBookings} client-only
                  </small>
                </div>
                <div className="metric-card">
                  <span className="label">CANCELLATIONS MATCHED</span>
                  <strong>
                    {result.coverage.matchedCancellations} / {result.coverage.vendorCancellations}
                  </strong>
                  <small>{result.coverage.clientCancellations} in client ledger</small>
                </div>
                <div className="metric-card">
                  <span className="label">PNRs TO REVIEW</span>
                  <strong>{result.exceptionTotal.toLocaleString("en-IN")}</strong>
                  <small>per-PNR differences</small>
                </div>
              </div>
              <div className="differences">
                <div className="subhead">
                  <span>WHAT DOESN&apos;T MATCH</span>
                  <span className="mono">vendor ↔ ITS, by category</span>
                </div>
                {result.differences.map((item, i) => (
                  <div className="difference-row" key={i}>
                    <div className={`difference-badge ${item.status === "matched" ? "match" : "variance"}`}>
                      {item.status === "matched" ? <Check size={13} /> : <AlertTriangle size={13} />}
                    </div>
                    <div className="difference-main">
                      <strong>{item.label}</strong>
                      <p>{item.explanation}</p>
                    </div>
                    <div className="difference-values">
                      <span>Vendor {inr(item.vendorAmount)}</span>
                      <span>Client {inr(item.clientAmount)}</span>
                      <b className={item.variance === 0 ? "zero" : "nonzero"}>{signed(item.variance)}</b>
                    </div>
                  </div>
                ))}
              </div>
              <div className="exceptions">
                <div className="subhead">
                  <span>PNRs NEEDING ATTENTION</span>
                  <span className="mono">
                    {result.exceptionTotal.toLocaleString("en-IN")} total
                  </span>
                </div>
                {result.exceptions.length === 0 ? (
                  <p className="exceptions-empty">Every ticket matches on PNR, amount, commission and edge offer.</p>
                ) : (
                  <>
                    {exceptionGroups.map((group) => {
                      const items = showAll ? group.items : group.items.slice(0, COLLAPSED_EXCEPTIONS);
                      return (
                        <div className="exception-group" key={group.category}>
                          <div className="exception-group-head">
                            <span>{exceptionSectionTitle(group.category)}</span>
                            <span className="mono">
                              {group.items.length.toLocaleString("en-IN")} PNRs · {signed(group.variance)}
                            </span>
                          </div>
                          {items.map((item, i) => (
                            <div className="exception-row" key={`${item.pnr}-${item.issue}-${i}`}>
                              <div className="exception-main">
                                <strong>
                                  {item.pnr} <em>{item.category}</em>
                                </strong>
                                <p>{item.issue}</p>
                                <small>{item.detail}</small>
                                <small className="ask-vendor">
                                  {repeatPnrs.has(item.pnr) && <em className="repeat-badge">repeat</em>} {askVendor(item)}
                                </small>
                              </div>
                              <div className="difference-values">
                                <span>Vendor {inr(item.vendorAmount)}</span>
                                <span>Client {inr(item.clientAmount)}</span>
                                <b className="nonzero">{signed(item.variance)}</b>
                              </div>
                            </div>
                          ))}
                        </div>
                      );
                    })}
                    {result.exceptions.length > COLLAPSED_EXCEPTIONS && (
                      <button className="outline-btn exceptions-toggle" onClick={() => setShowAll((v) => !v)}>
                        {showAll ? "Show fewer" : `Show all ${result.exceptions.length}`}
                      </button>
                    )}
                  </>
                )}
              </div>
              <div className="audit">
                <div className="subhead">
                  <span>EVERY PNR — WHAT HAPPENED</span>
                  <span className="mono">{result.audit.length.toLocaleString("en-IN")} tickets matched</span>
                  <button className="outline-btn audit-toggle" onClick={() => setAuditOpen((v) => !v)} disabled={!result.audit.length}>
                    {auditOpen ? "Hide" : "Show"}
                  </button>
                </div>
                {!result.audit.length && (
                  <small className="audit-note">Restored from history — the joined audit is only kept in memory; re-upload the files or use the XLS export for the full grid.</small>
                )}
                {auditOpen && (
                  <>
                    <input
                      className="audit-search"
                      placeholder="Filter by PNR, route or status…"
                      value={auditQuery}
                      onChange={(e) => setAuditQuery(e.target.value)}
                    />
                    <div className="audit-table">
                      <div className="audit-head">
                        <span>PNR</span><span>Client</span><span>Vendor section</span><span>Cancel type</span><span>Fare diff</span><span>Comm diff</span><span>Offer diff</span><span>Charge diff</span><span>Variance</span><span>Note</span>
                      </div>
                      {result.audit
                        .filter((a) => !auditQuery || `${a.pnr} ${a.route} ${a.status} ${a.vendorSection}`.toLowerCase().includes(auditQuery.toLowerCase()))
                        .slice(0, 200)
                        .map((a, i) => (
                          <div className={`audit-row ${a.status}`} key={`${a.pnr}-${i}`}>
                            <span>{a.pnr}{repeatPnrs.has(a.pnr) && <em className="repeat-badge">repeat</em>}{a.carried && <em className="repeat-badge carried">prev week</em>}</span>
                            <span>{a.clientType}</span>
                            <span>{a.vendorSection}</span>
                            <span>{a.cancelType}{a.rescheduled ? " · resched" : ""}</span>
                            <span>{a.fareDiff === 0 ? "—" : signed(a.fareDiff)}</span>
                            <span>{a.commDiff === 0 ? "—" : signed(a.commDiff)}</span>
                            <span>{a.offerDiff === 0 ? "—" : signed(a.offerDiff)}</span>
                            <span>{a.chargeDiff === 0 ? "—" : signed(a.chargeDiff)}</span>
                            <span className={a.variance === 0 ? "zero" : "nonzero"}>{signed(a.variance)}</span>
                            <span>{a.note || a.route || "—"}</span>
                          </div>
                        ))}
                    </div>
                    <small className="audit-note">First 200 rows shown — the full audit is in the XLS export.</small>
                  </>
                )}
              </div>
              <details className="checks checks-details">
                <summary className="subhead">
                  <span>SANITY CHECKS</span>
                  <span className="mono">totals cross-verified against the files</span>
                </summary>
                {result.checks.map((item, i) => (
                  <CheckRow item={item} key={i} />
                ))}
              </details>
            </div>
          ) : loading ? (
            <ResultSkeleton />
          ) : (
            <div className="empty-state">
              <UploadCloud size={30} />
              <strong>Upload both files to begin</strong>
              <span>The result canvas will appear here.</span>
            </div>
          )}
        </div>
      </section>

      {historyOpen && (
        <section className="history-panel" aria-label="Reconciliation history">
          <div className="history-head">
            <div>
              <span className="kicker">Saved locally</span>
              <h2>Reconciliation history</h2>
              <p>Previous reports are stored only in this browser and can be reopened without uploading the source files again.</p>
            </div>
            <div className="history-head-actions">
              {history.length > 0 && (
                <button className="outline-btn" onClick={clearHistory}>
                  <Trash2 size={14} /> Clear all
                </button>
              )}
              <button className="outline-btn" onClick={() => setHistoryOpen(false)}>
                <X size={14} /> Close
              </button>
            </div>
          </div>
          {history.length === 0 ? (
            <p className="history-empty">No reconciliation reports have been saved yet.</p>
          ) : (
            <div className="history-list">
              {history.map((saved, index) => {
                const review = saved.result.checks.filter((c) => c.status === "review").length + saved.result.differences.filter((d) => d.status === "review").length;
                return (
                  <div className="history-row" key={`${saved.savedAt}-${index}`}>
                    <button className="history-main" onClick={() => restoreHistoryItem(saved)}>
                      <span className="history-title">
                        <span className={`vendor-chip sm ${saved.result.source.vendor === "RedBus" ? "red" : "abhi"}`}>{saved.result.source.vendor}</span>
                        <strong>{saved.result.source.invoiceNo || "Reconciliation report"}</strong>
                      </span>
                      <small>{saved.result.source.operator} · {saved.result.source.period} · {new Date(saved.savedAt).toLocaleString()}</small>
                    </button>
                    <span className="history-meta">
                      <b>{saved.result.coverage.matchedBookings.toLocaleString("en-IN")} matched</b>
                      <small className={review ? "review-text" : "ok-text"}>{review ? `${review} to review` : "clean"}</small>
                    </span>
                    <span className="history-actions">
                      <button className="icon-btn" title="Delete" onClick={() => deleteHistoryItem(saved.savedAt)}>
                        <Trash2 size={15} />
                      </button>
                    </span>
                  </div>
                );
              })}
            </div>
          )}
        </section>
      )}

      <div className="mobile-bar">
        <button onClick={reconcile} disabled={!canRun || loading} className="run-btn">
          {loading ? (
            <>
              <RefreshCw className="spin" size={15} /> Reading…
            </>
          ) : (
            <>
              Reconcile <ArrowRight size={15} />
            </>
          )}
        </button>
        <button className="outline-btn" onClick={downloadReport} disabled={!result}>
          <Download size={15} /> PDF
        </button>
        <button className="outline-btn" onClick={downloadXls} disabled={!result}>
          <FileDown size={15} /> XLS
        </button>
      </div>

      <footer>
        <span>HumsafarAI v1.0</span>
        <span>Built for shareable, audit-ready answers</span>
      </footer>
    </div>
  );
}
