import { NextRequest, NextResponse } from "next/server";
import { reconcileParsedExcel, type ExcelSummary } from "@/lib/reconcile";

export const maxDuration = 60;

const MAX_PDF_BYTES = 20 * 1024 * 1024;

export async function POST(request: NextRequest) {
  let form: FormData;
  try {
    form = await request.formData();
  } catch {
    return NextResponse.json({ error: "Expected a multipart form upload." }, { status: 400 });
  }

  const pdf = form.get("pdf");
  if (!(pdf instanceof File) || !pdf.name.toLowerCase().endsWith(".pdf")) {
    return NextResponse.json({ error: "The payment advice must be a PDF file." }, { status: 400 });
  }
  if (pdf.size > MAX_PDF_BYTES) {
    return NextResponse.json({ error: "The payment advice PDF is larger than 20 MB." }, { status: 400 });
  }

  let summary: ExcelSummary;
  try {
    summary = JSON.parse(String(form.get("excelSummary") ?? ""));
  } catch {
    return NextResponse.json({ error: "The Excel summary is not valid JSON. Please select the workbook again." }, { status: 400 });
  }
  if (!summary || !Array.isArray(summary.subscriptionItems) || !Number.isFinite(summary.rawBookingNet)) {
    return NextResponse.json({ error: "The Excel summary is incomplete. Please select the workbook again." }, { status: 400 });
  }

  try {
    const buffer = Buffer.from(await pdf.arrayBuffer());
    const excelName = typeof form.get("excelName") === "string" ? String(form.get("excelName")) : "browser-normalized Excel export";
    const result = await reconcileParsedExcel(buffer, summary, pdf.name, excelName);
    return NextResponse.json(result);
  } catch (error) {
    console.error("[reconcile] failed", error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Unable to reconcile these files." },
      { status: 400 },
    );
  }
}
