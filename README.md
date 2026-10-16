# HumsafarAI

Settlement reconciliation workbench: upload a payment-advice PDF and a detailed XLS/XLSX export, and get an auditable settlement bridge with a downloadable PDF report. Files are processed in memory — nothing is stored.

## Stack

- **Next.js (App Router) + TypeScript** — UI and API in one framework
- **Tailwind CSS** — "Ledger Modern" finance-workbench design system
- **xlsx** — browser-side workbook normalization (TIN/PNR sections, cancellation-overlap removal, subscription values)
- **pdf-parse** — server-side payment-advice extraction (14 category amounts, adjustments, final settlement)
- **pdfkit** — server-side PDF report generation

## Routes

- `POST /api/reconcile` — multipart `pdf` file + `excelSummary` JSON → normalized `Reconciliation`
- `POST /api/reconcile/report` — `Reconciliation` JSON → branded PDF report

## Develop

```bash
npm install
npm run dev    # http://localhost:3000
npm run build  # production build
```

Deploy: push to GitHub and import in Vercel — no extra configuration needed.
