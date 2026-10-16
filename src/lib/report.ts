import path from "node:path";
import PDFDocument from "pdfkit";
import type { Reconciliation } from "./reconcile";

const fontDir = path.resolve(process.cwd(), "public/assets");
const logoPath = path.resolve(process.cwd(), "public/humsafarai-logo.png");

function renderReport(doc: PDFKit.PDFDocument, result: Reconciliation) {
  doc.registerFont("Humsafar Sans", path.join(fontDir, "DejaVuSans.ttf"));
  doc.registerFont("Humsafar Sans Bold", path.join(fontDir, "DejaVuSans-Bold.ttf"));
  const regular = "Humsafar Sans";
  const bold = "Humsafar Sans Bold";
  const navy = "#123331", teal = "#0f766e", mint = "#e8f6f2", ink = "#20312f", muted = "#6c7d79", amber = "#a15c15", red = "#a13f2b", line = "#dfe8e5", paper = "#f7faf9";
  const inr = (v: number) => `₹${Number(v || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const moneyWithSign = (v: number) => `${v < 0 ? "-" : v > 0 ? "+" : ""}${inr(Math.abs(v))}`;
  const rule = () => { doc.moveTo(44, doc.y).lineTo(551, doc.y).strokeColor(line).lineWidth(0.7).stroke(); };
  const newPage = () => { doc.addPage(); doc.y = 76; };
  const section = (number: string, title: string) => { if (doc.y > 700) newPage(); doc.moveDown(0.7); doc.fillColor(teal).font(bold).fontSize(9).text(`${number}  /  ${title.toUpperCase()}`, 44, doc.y, { width: 507 }); doc.moveDown(0.28); rule(); doc.moveDown(0.55); };
  const label = (text: string, x: number, y: number, width = 150) => doc.fillColor(muted).font(bold).fontSize(7.5).text(text.toUpperCase(), x, y, { width, characterSpacing: 0.8 });
  const tableRow = (left: string, right: string, emphasis = false) => { if (doc.y > 740) newPage(); const y = doc.y; doc.fillColor(emphasis ? ink : muted).font(emphasis ? bold : regular).fontSize(emphasis ? 10.5 : 9.5).text(left, 58, y, { width: 330 }); doc.fillColor(emphasis ? teal : ink).font(emphasis ? bold : regular).text(right, 400, y, { width: 135, align: "right" }); doc.moveDown(0.38); rule(); doc.moveDown(0.38); };
  const drawHeader = () => { doc.save(); doc.image(logoPath, 44, 18, { fit: [30, 30] }); doc.fillColor(navy).font(bold).fontSize(8).text("HUMSAFARAI", 82, 24); doc.fillColor(muted).font(regular).fontSize(7.5).text("SETTLEMENT INTELLIGENCE", 82, 36); doc.restore(); };
  doc.on("pageAdded", drawHeader);
  drawHeader();
  doc.y = 76;
  doc.fillColor(navy).font(bold).fontSize(27).text("Settlement reconciliation");
  doc.moveDown(0.25).fillColor(muted).font(regular).fontSize(10).text("Client transaction report compared with the vendor payment advice, matched by PNR.", { width: 470 });
  doc.moveDown(1.05);

  const metaY = doc.y;
  doc.roundedRect(44, metaY, 507, 68, 7).fillAndStroke(paper, line);
  label("Operator", 60, metaY + 14); doc.fillColor(ink).font(bold).fontSize(10).text(result.source.operator, 60, metaY + 29, { width: 145 });
  label("Invoice", 235, metaY + 14); doc.fillColor(ink).font(bold).fontSize(10).text(result.source.invoiceNo, 235, metaY + 29, { width: 145 });
  label("Period", 410, metaY + 14); doc.fillColor(ink).font(bold).fontSize(10).text(result.source.period, 410, metaY + 29, { width: 120 });
  doc.y = metaY + 87;

  const heroY = doc.y;
  const reviewCount = result.differences.filter(item => item.status === "review").length + result.checks.filter(item => item.status === "review").length;
  doc.roundedRect(44, heroY, 507, 108, 9).fill(navy);
  label("Vendor settles", 62, heroY + 18); doc.fillColor("#ffffff").font(bold).fontSize(27).text(inr(result.settlement.vendorFinal), 62, heroY + 35);
  doc.fillColor("#b7dfd6").font(regular).fontSize(9).text(`Client ledger expects ${inr(result.settlement.clientNet)}  ·  gap ${moneyWithSign(result.settlement.gap)}`, 63, heroY + 78, { width: 300 });
  doc.roundedRect(398, heroY + 20, 125, 25, 12).fill(mint); doc.fillColor(reviewCount ? red : teal).font(bold).fontSize(7.2).text(reviewCount ? `${reviewCount} TO REVIEW` : "RECONCILED", 407, heroY + 29, { width: 107, align: "center" });
  doc.fillColor("#b7dfd6").font(regular).fontSize(8).text("Review any item marked REVIEW before external submission.", 325, heroY + 78, { width: 195, align: "right" });
  doc.y = heroY + 130;

  section("01", "Settlement bridge");
  result.bridge.forEach(item => tableRow(item.label, item.kind === "base" || item.kind === "final" ? inr(item.amount) : moneyWithSign(item.amount), item.kind === "final"));

  section("02", "Ticket coverage");
  const c = result.coverage;
  [["Booked tickets matched by PNR", `${c.matchedBookings.toLocaleString("en-IN")} of ${c.vendorBookings.toLocaleString("en-IN")}`], ["Booked tickets only in vendor advice", c.vendorOnlyBookings.toLocaleString("en-IN")], ["Booked tickets only in client ledger", c.clientOnlyBookings.toLocaleString("en-IN")], ["Cancellations matched by PNR", `${c.matchedCancellations} of ${c.vendorCancellations}`], ["Cancellations in client ledger", c.clientCancellations.toLocaleString("en-IN")]].forEach(([left, right]) => tableRow(left, right));

  section("03", "Where differences occur");
  const headY = doc.y;
  doc.roundedRect(50, headY - 3, 495, 18, 4).fill(paper);
  doc.fillColor(muted).font(bold).fontSize(7).text("COMPARISON", 58, headY + 2, { width: 180, characterSpacing: 0.8 });
  doc.text("VENDOR", 245, headY + 2, { width: 75, align: "right", characterSpacing: 0.8 });
  doc.text("CLIENT", 330, headY + 2, { width: 75, align: "right", characterSpacing: 0.8 });
  doc.text("VARIANCE", 400, headY + 2, { width: 74, align: "right", characterSpacing: 0.8 });
  doc.text("STATUS", 476, headY + 2, { width: 62, align: "center", characterSpacing: 0.8 });
  doc.y = headY + 19; doc.moveDown(0.35);
  result.differences.forEach(item => {
    const matched = item.status === "matched";
    const color = matched ? teal : red;
    const explanationLines = Math.max(1, Math.ceil(item.explanation.length / 100));
    const labelLines = Math.max(1, Math.ceil(item.label.length / 30));
    const rowHeight = 16 + labelLines * 10 + explanationLines * 10;
    if (doc.y + rowHeight > 748) newPage();
    const y = doc.y;
    doc.fillColor(ink).font(bold).fontSize(9.2).text(item.label, 58, y + 2, { width: 180 });
    doc.fillColor(ink).font(regular).fontSize(8.6).text(inr(item.vendorAmount), 245, y + 2, { width: 75, align: "right" });
    doc.text(inr(item.clientAmount), 330, y + 2, { width: 75, align: "right" });
    doc.fillColor(color).font(bold).fontSize(7.8).text(moneyWithSign(item.variance), 400, y + 2, { width: 74, align: "right" });
    doc.roundedRect(476, y, 62, 14, 7).fill(matched ? mint : "#fdeee8");
    doc.fillColor(color).font(bold).fontSize(6.6).text(matched ? "MATCHED" : "REVIEW", 476, y + 4.5, { width: 62, align: "center", characterSpacing: 0.5 });
    doc.fillColor(muted).font(regular).fontSize(7.8).text(item.explanation, 58, y + 6 + labelLines * 10, { width: 480, lineGap: 1 });
    doc.y = y + rowHeight + 2;
    rule();
    doc.moveDown(0.4);
  });

  section("04", "Tickets to review");
  const shown = result.exceptions;
  if (!shown.length) {
    doc.fillColor(muted).font(regular).fontSize(9).text("Every ticket matches on PNR, amount, commission and edge offer.");
  }
  shown.forEach(item => {
    if (doc.y > 715) newPage();
    const y = doc.y;
    doc.fillColor(ink).font(bold).fontSize(9).text(`${item.pnr}  ·  ${item.category}`, 58, y, { width: 300 });
    doc.fillColor(red).font(bold).fontSize(8.4).text(moneyWithSign(item.variance), 400, y, { width: 135, align: "right" });
    doc.fillColor(ink).font(regular).fontSize(8.4).text(item.issue, 58, y + 12, { width: 340 });
    doc.fillColor(muted).font(regular).fontSize(7.6).text(`Vendor ${inr(item.vendorAmount)}  ·  Client ${inr(item.clientAmount)}  ·  ${item.detail}`, 58, doc.y + 1, { width: 477 });
    doc.moveDown(0.35); rule(); doc.moveDown(0.35);
  });
  if (result.exceptionTotal > shown.length) {
    doc.fillColor(muted).font(regular).fontSize(8.5).text(`Showing the largest ${shown.length} of ${result.exceptionTotal} tickets; the full list is available in the app.`);
  }

  section("05", "Control checks");
  result.checks.forEach(check => { if (doc.y > 720) newPage(); const color = check.status === "matched" ? teal : check.status === "info" ? amber : red; doc.fillColor(color).font(bold).fontSize(9.5).text(`${check.status.toUpperCase()}  ·  ${check.label}`); doc.fillColor(muted).font(regular).fontSize(8.5).text(check.detail, { width: 470 }); doc.moveDown(0.6); });

  doc.moveDown(0.5); if (doc.y > 720) newPage(); const shareY = doc.y; doc.roundedRect(44, shareY, 507, 50, 6).fillAndStroke(mint, "#c7e8df"); doc.fillColor(teal).font(bold).fontSize(8.5).text("SHARE NOTE", 58, shareY + 11); doc.fillColor(ink).font(regular).fontSize(8.5).text("This report preserves the source-to-calculation trail so the settlement answer can be reviewed without reopening the source files.", 140, shareY + 10, { width: 390, height: 30 }); doc.y = shareY + 58;

  const pages = doc.bufferedPageRange();
  for (let page = pages.start; page < pages.start + pages.count; page += 1) {
    doc.switchToPage(page);
    // drawing below the bottom margin would make pdfkit auto-add blank pages
    const bottomMargin = doc.page.margins.bottom;
    doc.page.margins.bottom = 0;
    doc.moveTo(44, 797).lineTo(551, 797).strokeColor(line).lineWidth(0.7).stroke();
    doc.fillColor(muted).font(regular).fontSize(7.5).text("HumsafarAI  ·  Settlement intelligence", 44, 807);
    doc.text(`Page ${page - pages.start + 1} of ${pages.count}`, 470, 807, { width: 81, align: "right" });
    doc.page.margins.bottom = bottomMargin;
  }
}

export function buildReport(result: Reconciliation): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margin: 44, bufferPages: true, info: { Title: "HumsafarAI Settlement Reconciliation", Author: "HumsafarAI" } });
    const chunks: Buffer[] = [];
    doc.on("data", (chunk: Buffer) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
    renderReport(doc, result);
    doc.end();
  });
}
