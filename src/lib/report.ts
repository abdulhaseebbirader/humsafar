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
  const tableRow = (left: string, right: string, emphasis = false) => { const y = doc.y; doc.fillColor(emphasis ? ink : muted).font(emphasis ? bold : regular).fontSize(emphasis ? 10.5 : 9.5).text(left, 58, y, { width: 330 }); doc.fillColor(emphasis ? teal : ink).font(emphasis ? bold : regular).text(right, 400, y, { width: 135, align: "right" }); doc.moveDown(0.38); rule(); doc.moveDown(0.38); };
  const drawHeader = () => { doc.save(); doc.image(logoPath, 44, 18, { fit: [30, 30] }); doc.fillColor(navy).font(bold).fontSize(8).text("HUMSAFARAI", 82, 24); doc.fillColor(muted).font(regular).fontSize(7.5).text("SETTLEMENT INTELLIGENCE", 82, 36); doc.restore(); };
  doc.on("pageAdded", drawHeader);
  drawHeader();
  doc.y = 76;
  doc.fillColor(navy).font(bold).fontSize(27).text("Settlement reconciliation");
  doc.moveDown(0.25).fillColor(muted).font(regular).fontSize(10).text("An evidence-led comparison of the payment advice and detailed Excel export.", { width: 470 });
  doc.moveDown(1.05);

  const metaY = doc.y;
  doc.roundedRect(44, metaY, 507, 68, 7).fillAndStroke(paper, line);
  label("Operator", 60, metaY + 14); doc.fillColor(ink).font(bold).fontSize(10).text(result.source.operator, 60, metaY + 29, { width: 145 });
  label("Invoice", 235, metaY + 14); doc.fillColor(ink).font(bold).fontSize(10).text(result.source.invoiceNo, 235, metaY + 29, { width: 145 });
  label("Period", 410, metaY + 14); doc.fillColor(ink).font(bold).fontSize(10).text(result.source.period, 410, metaY + 29, { width: 120 });
  doc.y = metaY + 87;

  const heroY = doc.y;
  doc.roundedRect(44, heroY, 507, 108, 9).fill(navy);
  label("Final settlement", 62, heroY + 18); doc.fillColor("#ffffff").font(bold).fontSize(27).text(inr(result.pdf.finalSettlement), 62, heroY + 35);
  doc.fillColor("#b7dfd6").font(regular).fontSize(9).text(`After adjustment  ${moneyWithSign(result.pdf.adjustment)}`, 63, heroY + 78);
  doc.roundedRect(398, heroY + 20, 125, 25, 12).fill(mint); doc.fillColor(teal).font(bold).fontSize(7.2).text("READY TO SHARE", 407, heroY + 29, { width: 107, align: "center" });
  doc.fillColor("#b7dfd6").font(regular).fontSize(8).text("Review any item marked REVIEW before external submission.", 325, heroY + 78, { width: 195, align: "right" });
  doc.y = heroY + 130;

  section("01", "Settlement bridge");
  result.bridge.forEach(item => tableRow(item.label, item.kind === "base" || item.kind === "final" ? inr(item.amount) : moneyWithSign(item.amount), item.kind === "final"));

  section("02", "Source evidence");
  const sourceRows = [["Excel booking rows", result.excel.bookingRows.toLocaleString("en-IN")], ["Excel booking net (raw)", inr(result.excel.rawBookingNet ?? result.excel.bookingNet)], ["Active booking net after overlap", inr(result.excel.activeBookingNet ?? result.excel.bookingNet)], ["Embedded cancellation fare", inr(result.excel.embeddedCancellationFare)], ["Separate cancellation rows", result.excel.cancellationRows.toLocaleString("en-IN")], ["Subscription/detail total", inr(result.excel.subscriptionTotal)]];
  sourceRows.forEach(([left, right]) => tableRow(left, right));

  section("03", "Where differences occur");
  label("Comparison", 58, doc.y, 170); label("PDF", 235, doc.y, 100); label("Excel", 340, doc.y, 100); label("Variance", 445, doc.y, 105); doc.moveDown(0.5); rule(); doc.moveDown(0.45);
  (result.differences ?? []).forEach(item => {
    const color = item.status === "matched" ? teal : red;
    const explanationLines = Math.max(1, Math.ceil(item.explanation.length / 92));
    const rowHeight = 48 + explanationLines * 9;
    if (doc.y + rowHeight > 748) newPage();
    const y = doc.y;
    doc.roundedRect(50, y - 4, 495, rowHeight, 6).fill(item.status === "matched" ? "#f1faf7" : "#fff7f4");
    doc.fillColor(color).font(bold).fontSize(9.2).text(item.label, 62, y + 5, { width: 165 });
    doc.fillColor(ink).font(regular).fontSize(7.6).text(inr(item.pdfAmount), 235, y + 5, { width: 100, align: "right" });
    doc.text(inr(item.excelAmount), 340, y + 5, { width: 100, align: "right" });
    doc.fillColor(color).font(bold).fontSize(7.6).text(moneyWithSign(item.variance), 445, y + 5, { width: 90, align: "right" });
    doc.fillColor(color).font(bold).fontSize(6.8).text(item.status === "matched" ? "MATCHED" : "REVIEW", 62, y + 22, { width: 62, characterSpacing: 0.7 });
    doc.fillColor(muted).font(regular).fontSize(7.8).text(item.explanation, 128, y + 21, { width: 402, lineGap: 1 });
    doc.y = y + rowHeight + 7;
  });
  doc.moveDown(0.2); rule(); doc.moveDown(0.35);

  section("04", "Subscription detail match");
  [["Dynamic Pricing", result.pdf.dynamicPricing], ["SRP", result.pdf.srp], ["Sponsored Listing", result.pdf.sponsored], ["Nova", result.pdf.nova]].forEach(([itemLabel, value]) => { const matched = result.excel.subscriptionItems?.find(item => item.label === itemLabel)?.matched ?? "Review"; tableRow(`${String(itemLabel)}  ·  ${matched}`, inr(Number(value))); });

  section("05", "Control checks");
  result.checks.forEach(check => { if (doc.y > 720) newPage(); const color = check.status === "matched" ? teal : check.status === "info" ? amber : red; doc.fillColor(color).font(bold).fontSize(9.5).text(`${check.status.toUpperCase()}  ·  ${check.label}`); doc.fillColor(muted).font(regular).fontSize(8.5).text(check.detail, { width: 470 }); doc.moveDown(0.6); });

  doc.moveDown(0.5); if (doc.y > 720) newPage(); const shareY = doc.y; doc.roundedRect(44, shareY, 507, 50, 6).fillAndStroke(mint, "#c7e8df"); doc.fillColor(teal).font(bold).fontSize(8.5).text("SHARE NOTE", 58, shareY + 11); doc.fillColor(ink).font(regular).fontSize(8.5).text("This report preserves the source-to-calculation trail so the settlement answer can be reviewed without reopening the workbook.", 140, shareY + 10, { width: 390, height: 30 }); doc.y = shareY + 58;

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
