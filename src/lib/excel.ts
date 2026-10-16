"use client";

import * as XLSX from "xlsx";
import type { ExcelSummary } from "@/lib/reconcile";

const excelNumber = (value: unknown) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value ?? "")
    .replace(/[₹,\s]/g, "")
    .replace(/^\((.*)\)$/, "-$1");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
};
const excelRound = (value: number) => Math.round((value + Number.EPSILON) * 100) / 100;

export async function normalizeExcelInBrowser(file: File): Promise<ExcelSummary> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const sheet = workbook.Sheets["BoInvoice"] || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("The workbook does not contain a transaction sheet.");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: true, defval: null });
  const headers = rows
    .map((row, index) => ({ row, index }))
    .filter(
      ({ row }) =>
        String(row?.[0] ?? "").trim().toLowerCase() === "tin" &&
        String(row?.[1] ?? "").trim().toLowerCase() === "pnr",
    );
  if (!headers.length) throw new Error("This spreadsheet does not contain the expected TIN / PNR transaction header.");
  const column = (name: string) => {
    const index = headers[0].row.findIndex((value: unknown) =>
      String(value ?? "").toLowerCase().replace(/\s+/g, " ").includes(name.toLowerCase()),
    );
    if (index < 0) throw new Error(`The workbook is missing the required “${name}” column.`);
    return index;
  };
  const fareCol = column("ticket fare");
  const netCol = column("netpayable");
  const sections = headers.map(({ index }) => {
    const end = headers.find((item) => item.index > index)?.index ?? rows.length;
    let label = "";
    for (let i = index - 1; i >= 0; i -= 1) {
      const value = String(rows[i]?.[0] ?? "").trim();
      if (value) {
        label = value;
        break;
      }
    }
    return {
      label,
      data: rows
        .slice(index + 1, end)
        .filter((row) => row[1] !== null && row[1] !== undefined && String(row[1]).trim() && String(row[1]) !== "PNR"),
    };
  });
  // Real exports repeat the Payment block: every Payment-labelled section
  // belongs to the same booking total.
  const paymentSections = sections.filter((section) => /payment/i.test(section.label));
  const main = {
    data: paymentSections.length ? paymentSections.flatMap((section) => section.data) : (sections[0]?.data ?? []),
  };
  const cancellationSections = sections.filter((section) =>
    /(cancellation|bus cancelled|exceptional refund)/i.test(section.label),
  );
  const cancellationRows = Array.from(
    new Map(
      cancellationSections.flatMap((section) => section.data).map((row) => [String(row[1]).trim(), row] as const),
    ).values(),
  );
  const cancellationPnrs = new Set(cancellationRows.map((row) => String(row[1]).trim()));
  // Plain "Cancellation" sections correspond to the PDF cancellation category.
  const plainCancellationRows = Array.from(
    new Map(
      cancellationSections
        .filter((section) => /cancellation/i.test(section.label) && !/bus|exceptional/i.test(section.label))
        .flatMap((section) => section.data)
        .map((row) => [String(row[1]).trim(), row] as const),
    ).values(),
  );
  const embedded = Array.from(
    new Map(
      main.data
        .filter((row) => cancellationPnrs.has(String(row[1]).trim()))
        .map((row) => [String(row[1]).trim(), row] as const),
    ).values(),
  );
  const sum = (items: unknown[][], index: number) =>
    excelRound(items.reduce((total, row) => total + excelNumber(row[index]), 0));
  const subscriptionSheet = workbook.SheetNames.map((name) => workbook.Sheets[name]).find((candidate) => {
    const values = XLSX.utils.sheet_to_json<unknown[]>(candidate, { header: 1, raw: true, defval: null });
    return values.some((row) =>
      row.some((value: unknown) => String(value ?? "").toLowerCase().includes("total subscription value")),
    );
  });
  if (!subscriptionSheet) throw new Error("The workbook does not contain a Total Subscription Value column.");
  const subscriptionRows = XLSX.utils.sheet_to_json<unknown[]>(subscriptionSheet, {
    header: 1,
    raw: true,
    defval: null,
  });
  const subscriptionHeaderIndex = subscriptionRows.findIndex((row) =>
    row.some((value) => String(value ?? "").toLowerCase().includes("total subscription value")),
  );
  const subscriptionCol = subscriptionRows[subscriptionHeaderIndex].findIndex((value) =>
    String(value ?? "").toLowerCase().includes("total subscription value"),
  );
  const subscriptionItems = subscriptionRows.slice(subscriptionHeaderIndex + 1).flatMap((row, offset) => {
    const amount = excelNumber(row[subscriptionCol]);
    return amount > 0
      ? [{ label: `Excel subscription row ${subscriptionHeaderIndex + offset + 2}`, amount: excelRound(amount), matched: "" }]
      : [];
  });
  const rawBookingFare = sum(main.data, fareCol);
  const rawBookingNet = sum(main.data, netCol);
  const embeddedCancellationFare = sum(embedded, fareCol);
  const embeddedCancellationNet = sum(embedded, netCol);
  return {
    bookingRows: main.data.length,
    rawBookingFare,
    rawBookingNet,
    bookingFare: excelRound(rawBookingFare - embeddedCancellationFare),
    bookingNet: excelRound(rawBookingNet - embeddedCancellationNet),
    activeBookingFare: excelRound(rawBookingFare - embeddedCancellationFare),
    activeBookingNet: excelRound(rawBookingNet - embeddedCancellationNet),
    embeddedCancellationFare,
    overlapCount: embedded.length,
    cancellationRows: cancellationRows.length,
    cancellationFare: sum(cancellationRows, fareCol),
    plainCancellationFare: sum(plainCancellationRows, fareCol),
    cancellationNet: sum(cancellationRows, netCol),
    subscriptionTotal: excelRound(subscriptionItems.reduce((total, item) => total + item.amount, 0)),
    subscriptionItems,
  };
}
