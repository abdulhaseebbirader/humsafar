import { NextRequest, NextResponse } from "next/server";
import { buildReport } from "@/lib/report";
import type { Reconciliation } from "@/lib/reconcile";

export const maxDuration = 60;

export async function POST(request: NextRequest) {
  let result: Reconciliation;
  try {
    result = (await request.json()) as Reconciliation;
  } catch {
    return NextResponse.json({ error: "A reconciliation result is required." }, { status: 400 });
  }
  if (!result?.settlement || !Array.isArray(result?.differences) || !Array.isArray(result?.bridge)) {
    return NextResponse.json({ error: "A reconciliation result is required." }, { status: 400 });
  }
  // Cap rendered rows so a crafted payload cannot exhaust the report worker.
  const MAX_REPORT_EXCEPTIONS = 2000;
  if (Array.isArray(result.exceptions) && result.exceptions.length > MAX_REPORT_EXCEPTIONS) {
    result.exceptions = result.exceptions.slice(0, MAX_REPORT_EXCEPTIONS);
  }

  try {
    const pdf = await buildReport(result);
    return new NextResponse(new Uint8Array(pdf), {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename="humsafarAI-${result.source.invoiceNo || "report"}.pdf"`,
      },
    });
  } catch (error) {
    console.error("[report] failed", error);
    return NextResponse.json({ error: "The PDF report could not be generated." }, { status: 500 });
  }
}
