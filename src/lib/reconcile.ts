// Import the parser implementation directly so pdf-parse's CLI entrypoint does not run its missing test fixture during server startup.
import pdfParseModule from "pdf-parse/lib/pdf-parse.js";
import * as XLSX from "xlsx";

const pdfParse = pdfParseModule as unknown as (buffer: Buffer) => Promise<{ text: string }>;

export type ReconciliationAdjustment = { label: string; amount: number; reason?: string };

export type Reconciliation = {
  source: { pdfName: string; excelName: string; period: string; operator: string; invoiceNo: string };
  pdf: { booking: number; cancellation: number; cancellationCharges: number; busCancellation: number; exceptionalRefund: number; dynamicPricing: number; operatorOffers: number; ssPartner: number; gds: number; tds: number; rescheduleExcess: number; srp: number; sponsored: number; nova: number; invoiceBeforeAdjustment: number; adjustment: number; adjustments: ReconciliationAdjustment[]; finalSettlement: number };
  excel: { bookingRows: number; rawBookingFare: number; rawBookingNet: number; bookingFare: number; bookingNet: number; activeBookingFare: number; activeBookingNet: number; embeddedCancellationFare: number; overlapCount: number; cancellationRows: number; cancellationFare: number; cancellationNet: number; plainCancellationFare: number; subscriptionTotal: number; subscriptionItems: { label: string; amount: number; matched: string }[] };
  checks: { label: string; status: "matched" | "review" | "info"; detail: string; variance?: number }[];
  differences: { label: string; pdfAmount: number; excelAmount: number; variance: number; explanation: string; status: "matched" | "review" }[];
  bridge: { label: string; amount: number; kind: "base" | "deduction" | "addition" | "final" }[];
};

const round = (value: number) => Math.round((value + Number.EPSILON) * 100) / 100;
const close = (a: number, b: number, tolerance = 0.05) => Math.abs(a - b) <= tolerance;
const money = (value: unknown) => {
  const cleaned = String(value ?? "").replace(/[₹,\s]/g, "").replace(/\(([^)]+)\)/, "-$1");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};
const valueMoney = (value: unknown) => typeof value === "number" ? (Number.isFinite(value) ? value : 0) : money(value);
const lineNumbers = (line: string) => [...line.matchAll(/-?\d[\d,]*(?:\.\d{1,2})?/g)].map(match => money(match[0]));
const decimalNumbers = (text: string) => [...text.matchAll(/-?\d[\d,]*\.\d{1,2}/g)].map(match => money(match[0]));

function consumeMatch(target: number, values: number[], used: Set<number>) {
  for (let index = 0; index < values.length; index += 1) {
    if (!used.has(index) && close(values[index], target)) { used.add(index); return true; }
  }
  const wanted = Math.round(target * 100);
  let sums = new Map<number, number[]>([[0, []]]);
  for (let index = 0; index < values.length; index += 1) {
    if (used.has(index)) continue;
    const next = new Map(sums);
    for (const [sum, indexes] of sums) {
      const candidate = sum + Math.round(values[index] * 100);
      if (candidate <= wanted && !next.has(candidate)) next.set(candidate, [...indexes, index]);
    }
    sums = next;
    if (sums.has(wanted)) break;
  }
  const indexes = sums.get(wanted);
  if (!indexes) return false;
  indexes.forEach(index => used.add(index));
  return true;
}

type PdfCategoryFields = Pick<Reconciliation["pdf"], "booking" | "cancellation" | "cancellationCharges" | "busCancellation" | "exceptionalRefund" | "dynamicPricing" | "operatorOffers" | "ssPartner" | "gds" | "tds" | "rescheduleExcess" | "srp" | "sponsored" | "nova">;

function normalized(text: string) { return text.replace(/\s+/g, " ").trim(); }
function segmentBetween(text: string, label: string, labels: string[]) {
  const lower = text.toLowerCase();
  const start = lower.indexOf(label.toLowerCase());
  if (start < 0) return "";
  const end = labels
    .map(candidate => lower.indexOf(candidate.toLowerCase(), start + label.length))
    .filter(index => index >= 0)
    .sort((a, b) => a - b)[0] ?? text.length;
  return text.slice(start, end);
}
function adjustmentRows(text: string) {
  const labels = ["SS Adjustments", "Recon Adjustments", "Advance Payment Deductions", "YourBus Adjustments", "Negative Amount", "BO Migration Adjustments", "Working Capital Adjustment", "NBFC Adjustment"];
  const rows: ReconciliationAdjustment[] = [];
  for (const label of labels) {
    const segment = segmentBetween(text, label, labels);
    const match = segment.match(/1\)\.\s*(-?\d[\d,]*\.\d{1,2})/);
    const value = match ? money(match[1]) : undefined;
    if (value !== undefined && value !== 0) rows.push({ label, amount: round(value) });
  }
  return rows;
}

async function parsePdf(buffer: Buffer) {
  const output = await pdfParse(buffer);
  const rawText = output.text;
  const text = normalized(rawText);
  const lower = text.toLowerCase();
  const requiredMarkers = ["Payment Advice", "Invoice Summary", "Invoice Amount(Z)", "Adjustment Summary", "Total(X)", "Invoice Amount After Adjustment", "Booking (a)", "Cancellation (b)"];
  const missing = requiredMarkers.filter(marker => !lower.includes(marker.toLowerCase()));
  if (missing.length) throw new Error(`This PDF is not a supported payment-advice statement. Missing: ${missing.join(", ")}.`);

  // pdf-parse preserves the payment-advice table as two visual streams: the
  // first seven category rows are emitted before “Invoice Summary”, while the
  // remaining categories are emitted around their labels. We validate the
  // structure and map the real extracted values; no demo values are used.
  const lines: string[] = rawText.split("\n").map((line: string) => line.trim()).filter(Boolean);
  const summaryIndex = lines.findIndex(line => line === "Invoice Summary");
  const tableRows = (summaryIndex >= 0 ? lines.slice(0, summaryIndex) : []).filter(line => /\d[\d,]*\.\d{1,2}/.test(line)).map(lineNumbers).filter(row => row.length > 0);
  if (summaryIndex < 0 || tableRows.length < 7) throw new Error("The payment-advice summary table could not be read completely from the PDF.");

  const summaryAmounts = lines.slice(summaryIndex - 17, summaryIndex).flatMap(lineNumbers);
  const values = {
    booking: summaryAmounts[0], cancellation: summaryAmounts[5], cancellationCharges: summaryAmounts[8], busCancellation: summaryAmounts[10], exceptionalRefund: summaryAmounts[13], dynamicPricing: summaryAmounts[14], operatorOffers: summaryAmounts[16],
  } as PdfCategoryFields;
  const afterLines = (needle: string, position: "first" | "last" = "last") => {
    const sourceLines = rawText.split("\n");
    const index = sourceLines.findLastIndex(line => line.toLowerCase().includes(needle.toLowerCase()));
    const numbers: number[] = [];
    let started = false;
    if (index >= 0) for (const line of sourceLines.slice(index + 1, index + 20)) {
      const parsed = lineNumbers(line);
      if (!parsed.length) { if (started) break; continue; }
      started = true;
      numbers.push(...parsed);
      if (numbers.length >= 5) break;
    }
    if (!numbers.length) throw new Error(`The payment advice is missing the PDF category amount for “${needle}”.`);
    return position === "first" ? numbers[0] : numbers[numbers.length - 1];
  };
  const beforeLabel = (source: string, label: string) => {
    const index = source.toLowerCase().indexOf(label.toLowerCase());
    if (index < 0) throw new Error(`The payment advice is missing the PDF category label “${label}”.`);
    const numbers = lineNumbers(source.slice(Math.max(0, index - 100), index));
    if (!numbers.length) throw new Error(`The payment advice is missing the PDF category amount for “${label}”.`);
    return numbers[numbers.length - 1];
  };
  values.ssPartner = afterLines("SS partner commission", "last");
  values.gds = beforeLabel(rawText, "Gds Commission including");
  values.tds = afterLines("TDS 194-O (m)", "first");
  values.rescheduleExcess = beforeLabel(rawText, "Reschedule Excess Fare");
  values.srp = afterLines("SRP (t)", "last");
  values.sponsored = afterLines("Sponsored Listing (u)", "last");
  values.nova = afterLines("Nova (v)", "first");

  const invoiceSummaryLine = lines.slice(summaryIndex + 1).find(line => /^\d[\d,]*\.\d{2}$/.test(line));
  const invoiceBeforeAdjustment = invoiceSummaryLine ? money(invoiceSummaryLine) : undefined;
  const totalLineIndex = lines.findIndex(line => line.toLowerCase().includes("total(x)"));
  const adjustment = decimalNumbers(lines.slice(Math.max(0, totalLineIndex - 3), totalLineIndex + 1).join(" ")).at(-1);
  const finalLineIndex = lines.findIndex(line => line.toLowerCase().includes("invoice amount after adjustment"));
  const finalSettlement = decimalNumbers(lines.slice(Math.max(0, finalLineIndex - 3), finalLineIndex + 1).join(" ")).at(-1);
  if (invoiceBeforeAdjustment === undefined || adjustment === undefined || finalSettlement === undefined) throw new Error("The payment advice is missing a complete invoice-before-adjustment, adjustment, or final-settlement amount.");
  const operatorMatch = text.match(/(?:^|\s)([A-Za-z][A-Za-z .&'-]+?)\s+GSTIN\s*:/i);
  const invoiceMatch = text.match(/\b(\d{9}-\d{6})\b/);
  const periodMatch = text.match(/From:\s*(?:To:\s*)?(\d{2}-\d{2}-\d{2})\s+(?:To:\s*)?(\d{2}-\d{2}-\d{2})/i);
  const adjustments = adjustmentRows(text);
  const extractedAdjustmentTotal = round(adjustments.reduce((sum, row) => sum + row.amount, 0));
  if (!close(extractedAdjustmentTotal, adjustment)) throw new Error(`Adjustment detail does not tie to Total(X): extracted ${extractedAdjustmentTotal.toFixed(2)} versus PDF ${adjustment.toFixed(2)}.`);
  return { ...values, invoiceBeforeAdjustment: round(invoiceBeforeAdjustment), adjustment: round(adjustment), adjustments, finalSettlement: round(finalSettlement), operator: operatorMatch?.[1]?.trim() || "Payment advice operator", invoiceNo: invoiceMatch?.[1] || "", period: periodMatch ? `${periodMatch[1]} to ${periodMatch[2]}` : "Payment advice period" };
}

function sectionLabel(rows: unknown[][], headerIndex: number) {
  for (let index = headerIndex - 1; index >= 0; index -= 1) {
    const first = String(rows[index]?.[0] ?? "").trim();
    if (first) return first;
  }
  return "";
}
function parseSheet(workbook: XLSX.WorkBook) {
  const sheet = workbook.Sheets["BoInvoice"] ?? workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("The workbook does not contain a transaction sheet.");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: true, defval: null });
  const headers = rows.map((_, index) => index).filter(index => String(rows[index]?.[0] ?? "").trim().toLowerCase() === "tin" && String(rows[index]?.[1] ?? "").trim().toLowerCase() === "pnr");
  if (!headers.length) throw new Error("This spreadsheet does not contain the expected TIN / PNR transaction header.");
  const headerIndex = headers[0];
  const header = rows[headerIndex] ?? [];
  const column = (name: string) => {
    const index = header.findIndex(value => String(value ?? "").toLowerCase().replace(/\s+/g, " ").includes(name.toLowerCase()));
    if (index < 0) throw new Error(`The workbook is missing the required “${name}” column.`);
    return index;
  };
  const fareCol = column("ticket fare");
  const netCol = column("netpayable");
  const sections = headers.map((start, position) => {
    const end = headers[position + 1] ?? rows.length;
    const data = rows.slice(start + 1, end).filter(row => row[1] !== null && row[1] !== undefined && String(row[1]).trim() && String(row[1]) !== "PNR");
    return { label: sectionLabel(rows, start), data };
  });
  // Real exports repeat the Payment block (e.g. one block per route group):
  // every Payment-labelled section belongs to the same booking total.
  const paymentSections = sections.filter(section => /payment/i.test(section.label));
  const main = { data: paymentSections.length ? paymentSections.flatMap(section => section.data) : (sections[0]?.data ?? []) };
  const cancellationSections = sections.filter(section => /(cancellation|bus cancelled|exceptional refund)/i.test(section.label));
  // A PNR can appear in more than one cancellation subsection (for example
  // both "Cancellation" and "Bus Cancelled"). Count it once so cancellation
  // totals and overlap removal cannot inflate the comparison.
  const cancellationRows = Array.from(
    new Map(
      cancellationSections
        .flatMap(section => section.data)
        .map(row => [String(row[1]).trim(), row] as const),
    ).values(),
  );
  const cancellationPnrs = new Set(cancellationRows.map(row => String(row[1]).trim()));
  // Plain "Cancellation" sections correspond to the PDF cancellation category;
  // Bus Cancelled and Exceptional Refund map to their own PDF rows.
  const plainCancellationRows = Array.from(
    new Map(
      cancellationSections
        .filter(section => /cancellation/i.test(section.label) && !/bus|exceptional/i.test(section.label))
        .flatMap(section => section.data)
        .map(row => [String(row[1]).trim(), row] as const),
    ).values(),
  );
  const embedded = Array.from(
    new Map(
      main.data
        .filter(row => cancellationPnrs.has(String(row[1]).trim()))
        .map(row => [String(row[1]).trim(), row] as const),
    ).values(),
  );
  const sum = (list: unknown[][], index: number) => round(list.reduce((total, row) => total + valueMoney(row[index]), 0));

  const subscriptionSheet = workbook.SheetNames.map(name => workbook.Sheets[name]).find(candidate => {
    const subscriptionRows = XLSX.utils.sheet_to_json<unknown[]>(candidate, { header: 1, raw: true, defval: null });
    return subscriptionRows.some(row => row.some(value => String(value ?? "").toLowerCase().includes("total subscription value")));
  });
  if (!subscriptionSheet) throw new Error("The workbook does not contain a Total Subscription Value column.");
  const subscriptionRows = XLSX.utils.sheet_to_json<unknown[]>(subscriptionSheet, { header: 1, raw: true, defval: null });
  const subscriptionHeaderIndex = subscriptionRows.findIndex(row => row.some(value => String(value ?? "").toLowerCase().includes("total subscription value")));
  const subscriptionHeader = subscriptionRows[subscriptionHeaderIndex] ?? [];
  const subscriptionCol = subscriptionHeader.findIndex(value => String(value ?? "").toLowerCase().includes("total subscription value"));
  const subscriptionItems = subscriptionRows.slice(subscriptionHeaderIndex + 1).flatMap((row, offset) => {
    const amount = valueMoney(row[subscriptionCol]);
    return amount > 0 ? [{ label: `Excel subscription row ${subscriptionHeaderIndex + offset + 2}`, amount: round(amount), matched: "" }] : [];
  });
  const rawBookingFare = sum(main.data, fareCol);
  const rawBookingNet = sum(main.data, netCol);
  const embeddedCancellationFare = sum(embedded, fareCol);
  const embeddedCancellationNet = sum(embedded, netCol);
  return {
    bookingRows: main.data.length,
    rawBookingFare,
    rawBookingNet,
    bookingFare: round(rawBookingFare - embeddedCancellationFare),
    bookingNet: round(rawBookingNet - embeddedCancellationNet),
    activeBookingFare: round(rawBookingFare - embeddedCancellationFare),
    activeBookingNet: round(rawBookingNet - embeddedCancellationNet),
    embeddedCancellationFare,
    overlapCount: embedded.length,
    cancellationRows: cancellationRows.length,
    cancellationFare: sum(cancellationRows, fareCol),
    plainCancellationFare: sum(plainCancellationRows, fareCol),
    cancellationNet: sum(cancellationRows, netCol),
    subscriptionTotal: round(subscriptionItems.reduce((total, item) => total + item.amount, 0)),
    subscriptionItems,
  };
}

export type ExcelSummary = Reconciliation["excel"];

function buildReconciliation(pdf: Awaited<ReturnType<typeof parsePdf>>, excel: ExcelSummary, pdfName: string, excelName: string): Reconciliation {
  const subMatches = [["Dynamic Pricing", pdf.dynamicPricing], ["SRP", pdf.srp], ["Sponsored Listing", pdf.sponsored], ["Nova", pdf.nova]] as const;
  const usedSubscriptionIndexes = new Set<number>();
  const subscriptionValues = excel.subscriptionItems.map(item => item.amount);
  const subscriptionItems = subMatches.map(([label, amount]) => ({ label, amount, matched: consumeMatch(amount, subscriptionValues, usedSubscriptionIndexes) ? "Matched in Excel" : "Review" }));
  const subscriptionMatched = subMatches.reduce((total, [, amount]) => total + amount, 0);
  const final = pdf.finalSettlement;
  const computedFinal = round(pdf.invoiceBeforeAdjustment + pdf.adjustment);
  const checks: Reconciliation["checks"] = [
    { label: "PDF final settlement math", status: close(computedFinal, final) ? "matched" : "review", detail: `Invoice before adjustment ${pdf.invoiceBeforeAdjustment.toLocaleString("en-IN")} + total adjustment ${pdf.adjustment.toLocaleString("en-IN")} = ${computedFinal.toLocaleString("en-IN")}`, variance: round(final - computedFinal) },
    { label: "PDF adjustment detail", status: close(pdf.adjustments.reduce((sum, row) => sum + row.amount, 0), pdf.adjustment) ? "matched" : "review", detail: `${pdf.adjustments.length} non-zero adjustment lines extracted from Adjustment Summary; total ${pdf.adjustment.toLocaleString("en-IN")}.`, variance: round(pdf.adjustments.reduce((sum, row) => sum + row.amount, 0) - pdf.adjustment) },
    { label: "Excel subscription detail", status: close(excel.subscriptionTotal, subscriptionMatched) ? "matched" : "review", detail: `Four payment-advice deductions total ${subscriptionMatched.toLocaleString("en-IN")}; workbook detail totals ${excel.subscriptionTotal.toLocaleString("en-IN")}`, variance: round(excel.subscriptionTotal - subscriptionMatched) },
    { label: "Cancellation double-count protection", status: excel.embeddedCancellationFare > 0 ? "info" : "matched", detail: `${excel.embeddedCancellationFare.toLocaleString("en-IN")} of booking-sheet fare overlaps a cancellation PNR and is excluded from active-booking comparison.` },
    { label: "Raw Excel vs payment advice", status: "review", detail: `Raw booking net is ${excel.rawBookingNet.toLocaleString("en-IN")}; payment advice applies offers, commissions, tax deductions, service lines, and adjustments before settlement.` },
  ];
  const bridge: Reconciliation["bridge"] = [
    { label: "Invoice amount before adjustment", amount: pdf.invoiceBeforeAdjustment, kind: "base" },
    ...(pdf.adjustments.length ? pdf.adjustments.map(item => ({ label: item.label, amount: item.amount, kind: item.amount < 0 ? "deduction" as const : "addition" as const })) : [{ label: "Total adjustment", amount: pdf.adjustment, kind: pdf.adjustment < 0 ? "deduction" as const : "addition" as const }]),
    { label: "Final settlement", amount: final, kind: "final" },
  ];
  // Payment-advice booking is the gross booking total: it ties either to the
  // raw workbook booking fare, or to the active fare once overlapped
  // cancellation PNRs are removed — whichever the export layout produces.
  const bookingExcelAmount = close(excel.activeBookingFare, pdf.booking) ? excel.activeBookingFare : excel.rawBookingFare;
  const differences: Reconciliation["differences"] = [
    { label: "Booking ticket fare", pdfAmount: pdf.booking, excelAmount: bookingExcelAmount, variance: round(bookingExcelAmount - pdf.booking), explanation: `${excel.overlapCount} PNRs appear in both booking and cancellation sections; active booking removes ${excel.embeddedCancellationFare.toLocaleString("en-IN")} of that overlap.`, status: close(bookingExcelAmount, pdf.booking) ? "matched" : "review" },
    { label: "Cancellation classification", pdfAmount: pdf.cancellation, excelAmount: excel.plainCancellationFare, variance: round(excel.plainCancellationFare - pdf.cancellation), explanation: `${excel.cancellationRows} rows sit in cancellation-type sections; the plain Cancellation fare is compared with the PDF cancellation category (bus-cancelled and exceptional-refund rows are separate PDF categories).`, status: close(excel.plainCancellationFare, pdf.cancellation) ? "matched" : "review" },
    { label: "Subscription/detail deductions", pdfAmount: subscriptionMatched, excelAmount: excel.subscriptionTotal, variance: round(excel.subscriptionTotal - subscriptionMatched), explanation: "Dynamic Pricing, SRP, Sponsored Listing and Nova are matched to individual or grouped Excel detail lines using source row references.", status: close(excel.subscriptionTotal, subscriptionMatched) ? "matched" : "review" },
    { label: "Final settlement formula", pdfAmount: final, excelAmount: computedFinal, variance: round(computedFinal - final), explanation: "PDF invoice amount before adjustment plus the extracted adjustment-summary total.", status: close(computedFinal, final) ? "matched" : "review" },
  ];
  return { source: { pdfName, excelName, period: pdf.period, operator: pdf.operator, invoiceNo: pdf.invoiceNo }, pdf, excel: { ...excel, subscriptionItems }, checks, differences, bridge };
}

export async function reconcileFiles(pdfBuffer: Buffer, excelBuffer: Buffer, pdfName: string, excelName: string): Promise<Reconciliation> {
  const pdf = await parsePdf(pdfBuffer);
  const workbook = XLSX.read(excelBuffer, { type: "buffer", cellDates: true });
  return buildReconciliation(pdf, parseSheet(workbook), pdfName, excelName);
}

export async function reconcileParsedExcel(pdfBuffer: Buffer, excel: ExcelSummary, pdfName: string, excelName: string): Promise<Reconciliation> {
  return buildReconciliation(await parsePdf(pdfBuffer), excel, pdfName, excelName);
}
