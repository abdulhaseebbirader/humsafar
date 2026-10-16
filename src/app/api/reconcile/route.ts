import { NextRequest, NextResponse } from "next/server";
import { reconcileParsedExcel, type ExcelSummary } from "@/lib/reconcile";

export const maxDuration = 60;

const MAX_CSV_BYTES = 20 * 1024 * 1024;

export async function POST(request: NextRequest) {
  let form: FormData;
  try {
    form = await request.formData();
  } catch {
    return NextResponse.json({ error: "Expected a multipart form upload." }, { status: 400 });
  }

  const csv = form.get("csv");
  if (!(csv instanceof File) || !csv.name.toLowerCase().endsWith(".csv")) {
    return NextResponse.json({ error: "The vendor invoice must be a CSV file." }, { status: 400 });
  }
  if (csv.size > MAX_CSV_BYTES) {
    return NextResponse.json({ error: "The vendor invoice CSV is larger than 20 MB." }, { status: 400 });
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
    const buffer = Buffer.from(await csv.arrayBuffer());
    const excelName = typeof form.get("excelName") === "string" ? String(form.get("excelName")) : "browser-normalized Excel export";
    const result = await reconcileParsedExcel(buffer, summary, csv.name, excelName);
    return NextResponse.json(result);
  } catch (error) {
    console.error("[reconcile] failed", error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Unable to reconcile these files." },
      { status: 400 },
    );
  }
}
