import { NextResponse } from "next/server";
import pdfParse from "pdf-parse";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");
    if (!(file instanceof File)) {
      return NextResponse.json({ error: "Please provide a vendor payment advice PDF." }, { status: 400 });
    }
    if (file.size > 20 * 1024 * 1024) {
      return NextResponse.json({ error: "The vendor PDF is larger than the 20 MB limit." }, { status: 413 });
    }
    const parsed = await pdfParse(Buffer.from(await file.arrayBuffer()));
    if (!parsed.text.trim()) {
      return NextResponse.json({ error: "The vendor PDF does not contain readable text." }, { status: 422 });
    }
    return NextResponse.json({ text: parsed.text });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Unable to read the vendor payment advice PDF." },
      { status: 422 },
    );
  }
}
