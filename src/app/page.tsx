"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  AlertTriangle,
  ArrowRight,
  Check,
  Download,
  FileSpreadsheet,
  FileText,
  RefreshCw,
  ShieldCheck,
  UploadCloud,
} from "lucide-react";
import { normalizeExcelInBrowser } from "@/lib/excel";
import type { Reconciliation } from "@/lib/reconcile";

const inr = (v: number) =>
  `₹${Number(v || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

function FileBox({
  kind,
  file,
  onChange,
}: {
  kind: "pdf" | "excel";
  file: File | null;
  onChange: (f: File | null) => void;
}) {
  const isPdf = kind === "pdf";
  return (
    <label className={`file-box ${file ? "is-ready" : ""}`}>
      <input
        type="file"
        accept={
          isPdf
            ? ".pdf,application/pdf"
            : ".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        onChange={(e) => onChange(e.target.files?.[0] ?? null)}
      />
      <div className="file-icon">{isPdf ? <FileText size={20} /> : <FileSpreadsheet size={20} />}</div>
      <div className="file-copy">
        <strong>{isPdf ? "Payment advice PDF" : "Detailed Excel export"}</strong>
        <span>{file ? file.name : isPdf ? "Drop the settlement advice here" : "Drop the .xls or .xlsx here"}</span>
      </div>
      <div className="file-action">{file ? <Check size={18} /> : <UploadCloud size={18} />}</div>
    </label>
  );
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

export default function Home() {
  const [pdf, setPdf] = useState<File | null>(null);
  const [excel, setExcel] = useState<File | null>(null);
  const [result, setResult] = useState<Reconciliation | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const canRun = Boolean(pdf && excel);

  const reconcile = useCallback(async () => {
    if (!pdf || !excel || loading) return;
    setLoading(true);
    setError("");
    try {
      const excelSummary = await normalizeExcelInBrowser(excel);
      const body = new FormData();
      body.append("pdf", pdf);
      body.append("excelName", excel.name);
      body.append("excelSummary", JSON.stringify(excelSummary));
      const response = await fetch("/api/reconcile", { method: "POST", body });
      const responseText = await response.text();
      let data: Reconciliation | { error?: string } = {} as Reconciliation;
      try {
        data = responseText ? JSON.parse(responseText) : ({} as Reconciliation);
      } catch {
        throw new Error(
          response.ok
            ? "The reconciliation response was not valid JSON."
            : `Reconciliation service error (${response.status}). Please try again.`,
        );
      }
      if (!response.ok)
        throw new Error(
          (data as { error?: string }).error || `Unable to reconcile these files (HTTP ${response.status}).`,
        );
      setResult(data as Reconciliation);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Unable to reconcile these files.");
    } finally {
      setLoading(false);
    }
  }, [pdf, excel, loading]);

  // Auto-run once both files are selected; the async fetch in reconcile legitimately
  // sets loading/result state from this effect.
  useEffect(() => {
    if (pdf && excel) {
      // eslint-disable-next-line react-hooks/set-state-in-effect
      void reconcile();
    }
  }, [pdf, excel, reconcile]);

  const choosePdf = (file: File | null) => {
    setPdf(file);
    setResult(null);
    setError("");
  };
  const chooseExcel = (file: File | null) => {
    setExcel(file);
    setResult(null);
    setError("");
  };
  const subscriptionSum = useMemo(
    () =>
      result
        ? [result.pdf.dynamicPricing, result.pdf.srp, result.pdf.sponsored, result.pdf.nova].reduce(
            (a, b) => a + Number(b || 0),
            0,
          )
        : 0,
    [result],
  );

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
      a.download = `humsafarAI-${result.source.invoiceNo || "report"}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      setError(e instanceof Error ? e.message : "The PDF report could not be generated.");
    }
  }

  return (
    <div className="app-shell">
      <aside className="rail">
        <div className="brand">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img className="brand-mark brand-logo" src="/humsafarai-logo.png" alt="HumsafarAI logo" />
          <div>
            <strong>HumsafarAI</strong>
            <span>settlement control</span>
          </div>
        </div>
        <div className="rail-rule" />
        <div className="steps">
          <div className="step active">
            <span>01</span>
            <div>
              <strong>Load evidence</strong>
              <small>PDF + Excel</small>
            </div>
          </div>
          <div className={`step ${result ? "complete" : ""}`}>
            <span>02</span>
            <div>
              <strong>Reconcile</strong>
              <small>Detect overlap</small>
            </div>
          </div>
          <div className="step">
            <span>03</span>
            <div>
              <strong>Share result</strong>
              <small>Download PDF</small>
            </div>
          </div>
        </div>
        <div className="rail-note">
          <ShieldCheck size={17} />
          <p>
            <strong>Evidence first</strong>
            <br />
            Every total is tied back to a source file or formula.
          </p>
        </div>
      </aside>
      <main className="main">
        <header className="topbar">
          <div>
            <span className="eyebrow">PAYMENT ADVICE / CONTROL ROOM</span>
            <h1>Follow the money.</h1>
            <p>Turn two messy files into one defensible answer.</p>
          </div>
          <div className="topbar-badge">
            <span className="live-dot" /> local analysis · no files stored
          </div>
        </header>
        <section className="workspace">
          <div className="upload-panel">
            <div className="section-kicker">
              <span>01</span>
              <span>Load evidence</span>
            </div>
            <h2>Bring the two sides together.</h2>
            <p className="lead">
              Upload the payment advice and its detailed export. HumsafarAI reads both, separates embedded
              cancellations, and builds the settlement bridge.
            </p>
            <div className="file-stack">
              <FileBox kind="pdf" file={pdf} onChange={choosePdf} />
              <FileBox kind="excel" file={excel} onChange={chooseExcel} />
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
              {!canRun && <span className="hint">Upload both files to generate a reconciliation result.</span>}
              {canRun && loading && <span className="hint">Both files received — reconciling automatically.</span>}
            </div>
          </div>
          <div className="result-panel">
            <div className="result-header">
              <div className="section-kicker">
                <span>02</span>
                <span>Reconciliation result</span>
              </div>
              <div className="result-actions">
                <span className="confidence">
                  <span className="live-dot" /> {result ? "ready to review" : "waiting"}
                </span>
                <button className="outline-btn" onClick={downloadReport} disabled={!result}>
                  <Download size={15} /> Download PDF
                </button>
              </div>
            </div>
            {result ? (
              <>
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
                    <span className="label">FINAL SETTLEMENT</span>
                    <div className="amount">{inr(result.pdf.finalSettlement)}</div>
                    <p>
                      after adjustment <b>{inr(result.pdf.adjustment)}</b>
                    </p>
                  </div>
                  <div className="verified">
                    <Check size={17} /> reconciled
                  </div>
                </div>
                <div className="bridge">
                  <div className="bridge-title">
                    <span>SETTLEMENT BRIDGE</span>
                    <span className="mono">invoice → final</span>
                  </div>
                  {result.bridge.map((item, i) => (
                    <div className={`bridge-row ${item.kind}`} key={i}>
                      <span>{item.label}</span>
                      <strong>
                        {item.amount < 0 ? "−" : ""}
                        {inr(Math.abs(item.amount))}
                      </strong>
                    </div>
                  ))}
                </div>
                <div className="summary-grid">
                  <div className="metric-card">
                    <span className="label">EXCEL BOOKING ROWS</span>
                    <strong>{Number(result.excel.bookingRows).toLocaleString("en-IN")}</strong>
                    <small>raw booking detail</small>
                  </div>
                  <div className="metric-card">
                    <span className="label">CANCELLATION OVERLAP</span>
                    <strong>{inr(result.excel.embeddedCancellationFare)}</strong>
                    <small>excluded from active total</small>
                  </div>
                  <div className="metric-card">
                    <span className="label">SUBSCRIPTION MATCH</span>
                    <strong>{inr(subscriptionSum)}</strong>
                    <small>4 PDF deductions matched</small>
                  </div>
                </div>
                <div className="match-strip">
                  <div className="subhead">
                    <span>SUBSCRIPTION DETAIL MATCH</span>
                    <span className="mono">PDF ↔ Excel</span>
                  </div>
                  {(result.excel.subscriptionItems || []).map((item, i) => (
                    <div className="match-row" key={i}>
                      <span>{item.label}</span>
                      <strong>{inr(item.amount)}</strong>
                      <em>{item.matched || "Review"}</em>
                    </div>
                  ))}
                </div>
                <div className="differences">
                  <div className="subhead">
                    <span>WHERE DIFFERENCES OCCUR</span>
                    <span className="mono">PDF ↔ Excel</span>
                  </div>
                  {(result.differences || []).map((item, i) => (
                    <div className="difference-row" key={i}>
                      <div className={`difference-badge ${item.status === "matched" ? "match" : "variance"}`}>
                        {item.status === "matched" ? <Check size={13} /> : <AlertTriangle size={13} />}
                      </div>
                      <div className="difference-main">
                        <strong>{item.label}</strong>
                        <p>{item.explanation}</p>
                      </div>
                      <div className="difference-values">
                        <span>PDF {inr(item.pdfAmount)}</span>
                        <span>Excel {inr(item.excelAmount)}</span>
                        <b className={item.variance === 0 ? "zero" : "nonzero"}>
                          {item.variance > 0 ? "+" : ""}
                          {inr(item.variance)}
                        </b>
                      </div>
                    </div>
                  ))}
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
              </>
            ) : (
              <div className="empty-state">
                <UploadCloud size={30} />
                <strong>Upload both files to begin</strong>
                <span>The result canvas will appear here.</span>
              </div>
            )}
          </div>
        </section>
        <footer>
          <span>HumsafarAI v1.0</span>
          <span>Built for shareable, audit-ready answers</span>
        </footer>
      </main>
    </div>
  );
}
