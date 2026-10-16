"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AlertTriangle,
  ArrowRight,
  Check,
  Download,
  FileSpreadsheet,
  RefreshCw,
  ShieldCheck,
  UploadCloud,
} from "lucide-react";
import * as XLSX from "xlsx";
import { reconcileCsv, type Reconciliation } from "@/lib/reconcile";

const inr = (v: number) =>
  `₹${Number(v || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const signed = (v: number) => `${v > 0 ? "+" : v < 0 ? "−" : ""}${inr(Math.abs(v))}`;
const COLLAPSED_EXCEPTIONS = 12;

function FileBox({
  kind,
  file,
  onChange,
}: {
  kind: "vendor" | "client";
  file: File | null;
  onChange: (f: File | null) => void;
}) {
  const [dragging, setDragging] = useState(false);
  const isVendor = kind === "vendor";
  const acceptsFile = (f: File | undefined | null) => !!f && /\.(csv|xls|xlsx)$/i.test(f.name);
  return (
    <label
      className={`file-box ${file ? "is-ready" : ""} ${dragging ? "is-dragging" : ""}`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        const dropped = e.dataTransfer.files?.[0];
        if (acceptsFile(dropped)) onChange(dropped);
      }}
    >
      <input type="file" accept=".csv,.xls,.xlsx,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={(e) => onChange(e.target.files?.[0] ?? null)} />
      <div className="file-icon">
        <FileSpreadsheet size={22} />
      </div>
      <div className="file-copy">
        <strong>{isVendor ? "Vendor payment advice (CSV / XLS)" : "Client transaction report (CSV / XLS)"}</strong>
        <span>
          {file ? file.name : isVendor ? "Drop the vendor's payment advice .csv or .xls here" : "Drop your own booking ledger .csv or .xls here"}
        </span>
      </div>
      <div className="file-action">{file ? <Check size={18} /> : <UploadCloud size={18} />}</div>
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
  const [vendor, setVendor] = useState<File | null>(null);
  const [client, setClient] = useState<File | null>(null);
  const [result, setResult] = useState<Reconciliation | null>(null);
  const [loading, setLoading] = useState(false);
  const [downloaded, setDownloaded] = useState(false);
  const [showAll, setShowAll] = useState(false);
  const [error, setError] = useState("");
  const [savedReportCount, setSavedReportCount] = useState(0);
  const resultRef = useRef<HTMLDivElement | null>(null);
  const inFlight = useRef(false);
  const canRun = Boolean(vendor && client);

  const reconcile = useCallback(async () => {
    if (!vendor || !client || inFlight.current) return;
    inFlight.current = true;
    setLoading(true);
    setError("");
    setResult(null);
    setShowAll(false);
    try {
      const [vendorText, clientText] = await Promise.all([readAsCsv(vendor), readAsCsv(client)]);
      setResult(reconcileCsv(vendorText, clientText, vendor.name, client.name));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Unable to reconcile these files.");
    } finally {
      inFlight.current = false;
      setLoading(false);
    }
  }, [vendor, client]);

  // Auto-run once both files are selected; reconcile legitimately sets loading/result state.
  useEffect(() => {
    if (vendor && client) {
      // eslint-disable-next-line react-hooks/set-state-in-effect
      void reconcile();
    }
  }, [vendor, client, reconcile]);

  useEffect(() => {
    if (!result) return;
    const key = "humsafarai:reconciliation-reports";
    let saved: Array<{ savedAt: string; result: Reconciliation }> = [];
    try {
      saved = JSON.parse(window.localStorage.getItem(key) || "[]") as Array<{ savedAt: string; result: Reconciliation }>;
    } catch {
      saved = [];
    }
    saved.unshift({ savedAt: new Date().toISOString(), result });
    window.localStorage.setItem(key, JSON.stringify(saved.slice(0, 25)));
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setSavedReportCount(saved.slice(0, 25).length);
  }, [result]);

  useEffect(() => {
    try {
      const saved = JSON.parse(window.localStorage.getItem("humsafarai:reconciliation-reports") || "[]") as unknown[];
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setSavedReportCount(saved.length);
    } catch {
      setSavedReportCount(0);
    }
  }, []);

  useEffect(() => {
    if (result && window.matchMedia("(max-width: 900px)").matches) {
      resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }, [result]);

  const resetOutput = () => {
    setResult(null);
    setError("");
    setDownloaded(false);
  };
  const chooseVendor = (file: File | null) => {
    setVendor(file);
    resetOutput();
  };
  const chooseClient = (file: File | null) => {
    setClient(file);
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
  const visibleExceptions = result ? (showAll ? result.exceptions : result.exceptions.slice(0, COLLAPSED_EXCEPTIONS)) : [];

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

  const step1Done = Boolean(vendor && client);
  const step2Done = Boolean(result);
  const steps = [
    { n: "01", label: "Load evidence", state: step1Done ? "done" : "active" },
    { n: "02", label: "Reconcile", state: step2Done ? "done" : loading ? "active" : "idle" },
    { n: "03", label: "Share result", state: downloaded ? "done" : step2Done ? "active" : "idle" },
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
        <div className="topnav-badge">
          <span className="live-dot" /> local · {savedReportCount} report{savedReportCount === 1 ? "" : "s"} saved
        </div>
      </header>

      <section className="hero">
        <span className="eyebrow">PAYMENT ADVICE / CONTROL ROOM</span>
        <h1>Follow the money.</h1>
        <p>Your ledger against the vendor&apos;s invoice, ticket by ticket.</p>
      </section>

      <section className="workspace">
        <div className="panel upload-panel">
          <div className="panel-head">
            <span className="kicker">01 · Load evidence</span>
            <h2>Bring the two sides together.</h2>
            <p className="lead">
              Upload the vendor&apos;s payment advice as CSV or XLS and your own client transaction report as CSV or XLS. HumsafarAI matches
              every ticket by PNR and builds the bridge from your ledger to the vendor&apos;s settlement.
            </p>
          </div>
          <div className="file-stack">
            <FileBox kind="vendor" file={vendor} onChange={chooseVendor} />
            <FileBox kind="client" file={client} onChange={chooseClient} />
          </div>
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
            {!canRun && <span className="hint">Upload the vendor CSV and client CSV/XLS file to generate a reconciliation result.</span>}
            {canRun && loading && <span className="hint">Both files received — reconciling automatically.</span>}
          </div>
          <div className="panel-note">
            <ShieldCheck size={15} />
            <p>
              <strong>Evidence first.</strong> Files are read in your browser; every total is tied back to a source row.
            </p>
          </div>
        </div>

        <div className="panel result-panel" ref={resultRef}>
          <div className="result-header">
            <span className="kicker">02 · Reconciliation result</span>
            <div className="result-actions">
              <span className="confidence">
                <span className="live-dot" />{" "}
                {result ? (reviewCount ? `${reviewCount} to review` : "ready to review") : loading ? "working" : "waiting"}
              </span>
              <button className="outline-btn" onClick={downloadReport} disabled={!result}>
                <Download size={15} /> Download PDF
              </button>
            </div>
          </div>
          {result ? (
            <div className="result-body">
              <div className="identity">
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
                  <span className="label">VENDOR SETTLES</span>
                  <div className="amount">{inr(result.settlement.vendorFinal)}</div>
                  <p>
                    client ledger expects <b>{inr(result.settlement.clientNet)}</b> · gap{" "}
                    <b>{signed(result.settlement.gap)}</b>
                  </p>
                </div>
                <div className={`verified ${reviewCount ? "needs-review" : ""}`}>
                  {reviewCount ? (
                    <>
                      <AlertTriangle size={15} /> {reviewCount} to review
                    </>
                  ) : (
                    <>
                      <Check size={17} /> reconciled
                    </>
                  )}
                </div>
              </div>
              <div className="bridge">
                <div className="bridge-title">
                  <span>SETTLEMENT BRIDGE</span>
                  <span className="mono">client ledger → vendor final</span>
                </div>
                {result.bridge.map((item, i) => (
                  <div className={`bridge-row ${item.kind}`} key={i}>
                    <span>{item.label}</span>
                    <strong>
                      {item.kind === "base" || item.kind === "final" ? inr(item.amount) : signed(item.amount)}
                    </strong>
                  </div>
                ))}
              </div>
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
                  <span className="label">TICKETS TO REVIEW</span>
                  <strong>{result.exceptionTotal.toLocaleString("en-IN")}</strong>
                  <small>per-PNR differences</small>
                </div>
              </div>
              <div className="differences">
                <div className="subhead">
                  <span>WHERE DIFFERENCES OCCUR</span>
                  <span className="mono">vendor ↔ client</span>
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
                  <span>TICKETS TO REVIEW</span>
                  <span className="mono">
                    {result.exceptions.length < result.exceptionTotal
                      ? `largest ${result.exceptions.length} of ${result.exceptionTotal}`
                      : `${result.exceptionTotal} total`}
                  </span>
                </div>
                {result.exceptions.length === 0 ? (
                  <p className="exceptions-empty">Every ticket matches on PNR, amount, commission and edge offer.</p>
                ) : (
                  <>
                    {visibleExceptions.map((item, i) => (
                      <div className="exception-row" key={`${item.pnr}-${item.issue}-${i}`}>
                        <div className="exception-main">
                          <strong>
                            {item.pnr} <em>{item.category}</em>
                          </strong>
                          <p>{item.issue}</p>
                          <small>{item.detail}</small>
                        </div>
                        <div className="difference-values">
                          <span>Vendor {inr(item.vendorAmount)}</span>
                          <span>Client {inr(item.clientAmount)}</span>
                          <b className="nonzero">{signed(item.variance)}</b>
                        </div>
                      </div>
                    ))}
                    {result.exceptions.length > COLLAPSED_EXCEPTIONS && (
                      <button className="outline-btn exceptions-toggle" onClick={() => setShowAll((v) => !v)}>
                        {showAll ? "Show fewer" : `Show all ${result.exceptions.length}`}
                      </button>
                    )}
                  </>
                )}
              </div>
              <div className="checks">
                <div className="subhead">
                  <span>CONTROL CHECKS</span>
                  <span className="mono">source / calculated / variance</span>
                </div>
                {result.checks.map((item, i) => (
                  <CheckRow item={item} key={i} />
                ))}
              </div>
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
      </div>

      <footer>
        <span>HumsafarAI v1.0</span>
        <span>Built for shareable, audit-ready answers</span>
      </footer>
    </div>
  );
}
