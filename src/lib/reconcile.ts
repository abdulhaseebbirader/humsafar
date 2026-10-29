export type CheckStatus = "matched" | "review" | "info";

export type Reconciliation = {
  source: {
    vendor: string;
    vendorName: string;
    clientName: string;
    operator: string;
    agent: string;
    invoiceNo: string;
    period: string;
    clientPeriod: string;
    paymentDate: string;
  };
  settlement: { vendorFinal: number; clientNet: number; gap: number };
  coverage: {
    vendorBookings: number;
    clientBookings: number;
    matchedBookings: number;
    vendorOnlyBookings: number;
    clientOnlyBookings: number;
    vendorCancellations: number;
    clientCancellations: number;
    matchedCancellations: number;
  };
  differences: {
    label: string;
    vendorAmount: number;
    clientAmount: number;
    variance: number;
    explanation: string;
    status: "matched" | "review";
  }[];
  bridge: { label: string; amount: number; kind: "base" | "deduction" | "addition" | "final" }[];
  // Humsafar's manual SUMMARY-sheet formula: NET RECEIVABLE built from the
  // vendor components, compared to what the vendor actually settled.
  receivable: {
    rows: { label: string; amount: number }[];
    receivable: number;
    received: number;
    difference: number;
  };
  // Their manual FINAL tab: one weekly settlement row per advice/ITS pair.
  // Adjustment columns (Operator Offer, SRP, Sponsored, TDS, YourBus,
  // SeatSeller/Recon adjustments) are hand-keyed in their sheet and have no
  // source columns in the advice, so they stay 0 here.
  weekly: {
    vendorBooking: number;
    itsBooking: number;
    vendorComm: number;
    itsComm: number;
    vendorCancel: number;
    busCancelled: number;
    exceptional: number;
    itsCancel: number;
    itsCharges: number;
    vendorCharges: number;
    yourbus: number;
    total: number;
    received: number;
    diff: number;
    asPerIts: number;
    paymentDiff: number;
  };
  // Joined per-PNR audit — the same view their manual VLOOKUP columns produce.
  audit: {
    pnr: string;
    clientType: string;
    vendorSection: string;
    clientAmount: number;
    vendorAmount: number;
    variance: number;
    // Per-PNR clarification detail: fare / commission / offer / charge
    // compared side by side so every difference is explained on the row.
    vendorFare: number;
    clientFare: number;
    fareDiff: number;
    vendorComm: number;
    clientComm: number;
    commDiff: number;
    vendorOffer: number;
    clientOffer: number;
    offerDiff: number;
    vendorCharge: number;
    clientCharge: number;
    chargeDiff: number;
    cancelType: string;
    rescheduled: boolean;
    channel: string;
    // Journey date falls in a week before this settlement week — a PNR the
    // vendor carried over and only settled now.
    carried: boolean;
    carriedWeek: string;
    ticket: string;
    route: string;
    doj: string;
    settleDate: string;
    note: string;
    status: "matched" | "client-only" | "vendor-only" | "differs";
  }[];
  // Weekly settlement periods (bucketed by journey date), matching how the
  // manual SUMMARY sheet groups each period before comparing receipts.
  periods: {
    label: string;
    start: string;
    end: string;
    booked: number;
    cancel: number;
    serviceCancel: number;
    charges: number;
    exceptional: number;
    edge: number;
    receivable: number;
    tickets: number;
  }[];
  checks: { label: string; status: CheckStatus; detail: string; variance?: number }[];
  exceptions: {
    pnr: string;
    category: string;
    issue: string;
    vendorAmount: number;
    clientAmount: number;
    variance: number;
    detail: string;
  }[];
  exceptionTotal: number;
};

// Keep every exception so the on-screen audit trail and downloaded PDF are complete.
const MAX_EXCEPTIONS = Number.POSITIVE_INFINITY;
const round = (value: number) => Math.round((value + Number.EPSILON) * 100) / 100;
const close = (a: number, b: number, tolerance = 1) => Math.abs(a - b) <= tolerance;
const fmt = (value: number) =>
  `₹${value.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const money = (value: unknown) => {
  const cleaned = String(value ?? "")
    .replace(/[₹,\s]/g, "")
    .replace(/^Rs\.?/i, "")
    .replace(/^\((.*)\)$/, "-$1");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

export function parseCsv(text: string): string[][] {
  const source = text.replace(/^\uFEFF/, "");
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < source.length; i += 1) {
    const ch = source[i];
    if (quoted) {
      if (ch === '"') {
        if (source[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else quoted = false;
      } else cell += ch;
    } else if (ch === '"' && cell === "") quoted = true;
    else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && source[i + 1] === "\n") i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else cell += ch;
  }
  if (cell !== "" || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

const cellAt = (row: string[] | undefined, index: number) => (row?.[index] ?? "").trim();
const onlyFirstCell = (row: string[]) => row.slice(1).every((value) => !value.trim());

/* ----------------------------- vendor advice ----------------------------- */

type VendorKind = "booked" | "cancel" | "charges" | "exceptional" | "assured" | "busCancelled";
type VendorTicket = {
  ticket: string;
  pnr: string;
  doj: string;
  seats: string;
  amount: number;
  commission: number;
  commissionGst?: number;
  edge: number;
  net: number;
  settle?: string;
  charge?: number;
  note?: string;
  route?: string;
  ticketType?: string;
  channel?: string;
  rescheduled?: boolean;
};

function parseVendor(rows: string[][]) {
  const label = (row: string[]) => cellAt(row, 0);
  const rowText = (row: string[]) => row.map((value) => value.trim()).join(" ");
  const totalInvoiceIndex = rows.findIndex((row) =>
    /^(?:Total\s+Invoice|Total\s+Invoice\s*\()/i.test(label(row)) ||
    /\bTotal\s+Invoice\b/i.test(rowText(row)),
  );
  if (totalInvoiceIndex < 0) {
    // RedBus payment advice is a flat TIN/PNR table rather than the
    // sectioned AbhiBus advice with a Total Invoice row.
    const adviceHeaderIndex = rows.findIndex((row) => {
      const normalized = row.map((value) => value.trim());
      const hasPnr = normalized.some((value) => /^PNR$/i.test(value));
      const hasFare = normalized.some((value) => /^(Ticket\s*Fare|Gross\s*(Amount|Fare)|Total\s*Fare|Amount)$/i.test(value));
      const hasNet = normalized.some((value) => /^(Net\s*Payable|NetPayable|Net\s*Amount|Settlement\s*Amount|Payable)$/i.test(value));
      return hasPnr && (hasFare || hasNet);
    });
    if (adviceHeaderIndex >= 0) {
      // RedBus payment reports are multi-section CSVs: a title row such as
      // "<GSTIN>-Payment" / "-Cancellation" / "-Bus Cancelled" / "-Exceptional
      // Refund" (the sections can repeat), each followed by a TIN,PNR header
      // and ticket rows, then a totals row. Every section keeps the same
      // TIN/PNR/Fare columns but settles a different money column.
      type RedSection = "booked" | "cancel" | "busCancelled" | "exceptional";
      const sectionOf = (name: string): RedSection | null => {
        if (/payment\s*$/i.test(name)) return "booked";
        if (/bus\s*cancell?ed/i.test(name)) return "busCancelled";
        if (/exceptional\s*refund/i.test(name)) return "exceptional";
        if (/cancell?ation/i.test(name)) return "cancel";
        return null;
      };
      const isRedHeader = (row: string[]) =>
        /^TIN$/i.test(cellAt(row, 0)) && /^PNR$/i.test(cellAt(row, 1)) && row.some((value) => /ticket\s*fare/i.test(value));
      const redTickets: Record<RedSection, VendorTicket[]> = { booked: [], cancel: [], busCancelled: [], exceptional: [] };
      let section: RedSection | null = null;
      let cols = { pnr: 1, fare: -1, commission: -1, commissionGst: -1, others: -1, netPayable: -1, netReversed: -1, prevNetPaid: -1, refund: -1, cancelCharges: -1, reversal: -1, settle: -1, doj: -1, seats: -1, route: -1, typeof: -1, ticketType: -1, excReason: -1, resched: -1, channel: -1 };
      const applyRedHeader = (header: string[]) => {
        const find = (pattern: RegExp) => header.findIndex((value) => pattern.test(value.trim()));
        const settle = find(/^RED\s*BUS\s*SETTAL\s*DATE$/i);
        cols = {
          pnr: find(/^PNR$/i),
          fare: find(/^Ticket\s*Fare$/i),
          // Prefer pre-GST commission so it compares cleanly to the ITS
          // "Comm." column; GST-inclusive commission is a fallback.
          commission:
            find(/^Commission\s*Without\s*GST$/i) >= 0 ? find(/^Commission\s*Without\s*GST$/i) : find(/^Total\s*Commission\s*With\s*GST$/i),
          commissionGst: find(/^Total\s*Commission\s*With\s*GST$/i),
          others: find(/^Others$/i),
          netPayable: find(/^Net\s*Payable$/i),
          netReversed: find(/^Net\s*Reversed$/i),
          prevNetPaid: find(/^Previous\s*Net\s*Paid$/i),
          refund: find(/^Refund\s*Amount$/i),
          cancelCharges: find(/^Cancellation\s*Charges$/i),
          reversal: find(/^Reversal$/i),
          settle: settle >= 0 ? settle : find(/^Settal\s*date$/i),
          doj: find(/^DOM$/i) >= 0 ? find(/^DOM$/i) : find(/^DOE$/i),
          seats: find(/^Seats$/i),
          route: find(/^Route$/i),
          typeof: find(/^Type\s*of/i),
          ticketType: find(/^Ticket\s*Type$/i),
          excReason: find(/^Exceptional\s*$/i),
          resched: find(/^IsRescheduled$/i),
          channel: find(/^Sales\s*Channel$/i),
        };
      };
      applyRedHeader(rows[adviceHeaderIndex]);
      for (const row of rows.slice(adviceHeaderIndex + 1)) {
        const title = cellAt(row, 0);
        const named = title.match(/-(Payment|Cancellation|Bus\s*Cancelled|Exceptional\s*Refund)\s*$/i);
        if (named && onlyFirstCell(row)) {
          section = sectionOf(named[1]);
          continue;
        }
        if (isRedHeader(row)) {
          applyRedHeader(row);
          continue;
        }
        if (!section) section = "booked";
        const pnr = cellAt(row, cols.pnr);
        const fareValue = cols.fare >= 0 ? cellAt(row, cols.fare) : "";
        if (!title || !pnr || !/\d/.test(pnr) || !/\d/.test(fareValue)) continue;
        const fare = money(fareValue);
        const reversal = cols.reversal >= 0 ? money(cellAt(row, cols.reversal)) : 0;
        const netReversed = cols.netReversed >= 0 ? money(cellAt(row, cols.netReversed)) : 0;
        // Vendor's own reason text: cancellation type ("Normal cancellation"),
        // bus-cancel timing (PreDOJ/PostDOJ via TicketType), exceptional-refund
        // reason, reschedule flag and sales channel.
        const note = [
          cols.typeof >= 0 ? cellAt(row, cols.typeof) : "",
          section === "busCancelled" && cols.ticketType >= 0 ? cellAt(row, cols.ticketType) : "",
          section === "exceptional" && cols.excReason >= 0 ? cellAt(row, cols.excReason) : "",
          cols.resched >= 0 && /^true$/i.test(cellAt(row, cols.resched)) ? "Rescheduled" : "",
          cols.channel >= 0 ? cellAt(row, cols.channel) : "",
        ].filter(Boolean).join(" · ");
        const ticket: VendorTicket = {
          ticket: title,
          pnr,
          doj: cols.doj >= 0 ? cellAt(row, cols.doj) : "",
          seats: cols.seats >= 0 ? cellAt(row, cols.seats) : "",
          amount: fare,
          commission: cols.commission >= 0 ? money(cellAt(row, cols.commission)) : 0,
          commissionGst: cols.commissionGst >= 0 ? money(cellAt(row, cols.commissionGst)) : 0,
          edge: cols.others >= 0 ? money(cellAt(row, cols.others)) : 0,
          net:
            section === "booked"
              ? cols.netPayable >= 0
                ? money(cellAt(row, cols.netPayable))
                : fare
              : section === "cancel"
                ? netReversed || reversal
                : section === "busCancelled"
                  ? cols.prevNetPaid >= 0
                    ? money(cellAt(row, cols.prevNetPaid))
                    : 0
                  : cols.refund >= 0
                    ? money(cellAt(row, cols.refund))
                    : 0,
          settle: cols.settle >= 0 ? cellAt(row, cols.settle) : "",
          note: note || undefined,
          route: cols.route >= 0 ? cellAt(row, cols.route) : "",
          charge:
            section === "cancel"
              ? cols.cancelCharges >= 0
                ? money(cellAt(row, cols.cancelCharges))
                : Math.max(0, round(reversal - netReversed))
              : 0,
          ticketType: section === "busCancelled" && cols.ticketType >= 0 ? cellAt(row, cols.ticketType) : undefined,
          channel: cols.channel >= 0 ? cellAt(row, cols.channel) : undefined,
          rescheduled: cols.resched >= 0 && /^true$/i.test(cellAt(row, cols.resched)) ? true : undefined,
        };
        redTickets[section].push(ticket);
      }
      if (!redTickets.booked.length) throw new Error("The RedBus vendor payment advice has no ticket rows.");
      const total = (kind: RedSection, pick: (ticket: VendorTicket) => number) => round(redTickets[kind].reduce((acc, ticket) => acc + pick(ticket), 0));
      const booked = total("booked", (t) => t.amount);
      const commission = total("booked", (t) => t.commission);
      const commissionGst = total("booked", (t) => t.commissionGst ?? t.commission);
      const edge = total("booked", (t) => t.edge);
      const netPayable = total("booked", (t) => t.net);
      const cancelReversed = total("cancel", (t) => t.net);
      const cancelCharges = total("cancel", (t) => t.charge ?? 0);
      const busReversed = total("busCancelled", (t) => t.net);
      const refunded = total("exceptional", (t) => t.net);
      const final = round(netPayable - cancelReversed - busReversed - refunded - edge);
      return {
        operator: cellAt(rows[0], 0) || "RedBus payment advice",
        invoiceNo: cellAt(rows[0], 0),
        paymentDate: "",
        period: "RedBus payment advice period",
        vendorStyle: "redbus" as const,
        // bookingPayable comes from the fare/commission columns, not the net
        // column, so the "Vendor invoice math" check still catches a wrong NetPayable.
        // bookedCommission is the pre-GST figure for ITS comparison;
        // bookingPayable uses the GST-inclusive commission NetPayable is
        // built from.
        summary: { booked, bookedCommission: commission, cancel: cancelReversed, serviceCancel: busReversed, charges: cancelCharges, exceptional: refunded, edgeGross: edge, edgeCommission: 0, edgePayable: edge, bookingPayable: round(booked - commissionGst), tds: 0, final },
        amountSettled: final,
        tickets: { booked: redTickets.booked, cancel: redTickets.cancel, charges: [], exceptional: redTickets.exceptional, assured: [], busCancelled: redTickets.busCancelled },
      };
    }
    const looksLikeClient = rows.some((row) => /particulares/i.test(row.join(",")));
    throw new Error(
      looksLikeClient
        ? "The vendor slot holds a client transaction report. Please swap the two files."
        : "This CSV is not a supported vendor payment advice (expected AbhiBus sections or a RedBus TIN / PNR table).",
    );
  }

  const headerValue = (pattern: RegExp) => {
    const row = rows.find((candidate) => pattern.test(label(candidate)));
    return row ? label(row).replace(/^[^:]*:\s*/, "").trim() : "";
  };
  const operator = headerValue(/^Bus Operator Payment Advice Operator/i) || "Vendor payment advice";
  const invoiceNo = headerValue(/^Invoice Number/i);
  const paymentDate = headerValue(/^Payment Date/i);
  const periodRow = rows.find((row) => /^Summary from/i.test(label(row)));
  const period = periodRow
    ? `${cellAt(periodRow, 0).replace(/^.*:\s*/, "")} to ${cellAt(periodRow, 1).replace(/^.*:\s*/, "")}`
    : "Payment advice period";

  const sections: Record<string, Record<string, string[]>> = {};
  let section = "";
  for (const row of rows.slice(0, totalInvoiceIndex)) {
    const name = label(row);
    if (!name) continue;
    const title = name.match(/\(([A-G])\)\s*$/);
    if (title && onlyFirstCell(row)) {
      section = title[1];
      sections[section] = {};
      continue;
    }
    if (section) sections[section][/^Total\s*\(/i.test(name) ? "Total" : name] = row;
  }
  const pick = (sectionKey: string, pattern: RegExp) => {
    const entries = Object.entries(sections[sectionKey] ?? {});
    return entries.find(([key]) => pattern.test(key))?.[1];
  };
  const col = (row: string[] | undefined, index: number) => money(cellAt(row, index));

  const summary = {
    booked: col(pick("A", /^Booked Amount/i), 1),
    bookedCommission: col(pick("A", /^Booked Amount/i), 2),
    cancel: col(pick("A", /^Cancel Amount/i), 1),
    serviceCancel: col(pick("A", /^Service Cancel/i), 1),
    charges: col(pick("A", /^Cancellation Charges/i), 1),
    exceptional: col(pick("E", /^Total$/), 6),
    edgeGross: col(pick("G", /^Total$/), 1),
    edgeCommission: col(pick("G", /^Total$/), 2),
    edgePayable: col(pick("G", /^Total$/), 6),
    bookingPayable: col(pick("A", /^Total$/), 6),
    tds: col(rows[totalInvoiceIndex], 4),
    final: col(rows[totalInvoiceIndex], 6),
  };
  const settledRow = rows.find((row) => /^Amount Settled/i.test(label(row)));
  const amountSettled = settledRow ? money(label(settledRow).replace(/^.*Rs\.?/i, "")) : 0;

  const tickets: Record<VendorKind, VendorTicket[]> = { booked: [], cancel: [], charges: [], exceptional: [], assured: [], busCancelled: [] };
  const kindOf = (title: string): VendorKind | null => {
    if (/Cancellation Charges/i.test(title)) return "charges";
    if (/Cancellation Amount/i.test(title)) return "cancel";
    if (/Booked Amount/i.test(title)) return "booked";
    if (/Exceptional/i.test(title)) return "exceptional";
    if (/Assured/i.test(title)) return "assured";
    return null;
  };
  let kind: VendorKind | null = null;
  let columns = { amount: 7, commission: -1, edge: -1, net: -1 };
  for (const row of rows.slice(totalInvoiceIndex + 1)) {
    const name = label(row);
    if (/Invoice Details\s*$/i.test(name) && onlyFirstCell(row)) {
      kind = kindOf(name);
      continue;
    }
    if (/^TDS 194\(O\) Calculation/i.test(name)) {
      kind = null;
      continue;
    }
    if (!kind) continue;
    if (/^S\.?No/i.test(name)) {
      const find = (pattern: RegExp) => row.findIndex((value) => pattern.test(value.trim()));
      columns = {
        amount: Math.max(find(/^(ticket amount|cancel charges)/i), 0),
        commission: find(/commission/i),
        edge: find(/edge offer/i),
        net: find(/^net amount/i),
      };
      continue;
    }
    if (!/^\d+$/.test(name)) continue;
    const pnr = cellAt(row, 2).split("/")[0].trim();
    if (!pnr) continue;
    tickets[kind].push({
      ticket: cellAt(row, 1),
      pnr,
      doj: cellAt(row, 4),
      seats: cellAt(row, 6),
      route: cellAt(row, 5),
      amount: col(row, columns.amount),
      commission: columns.commission >= 0 ? col(row, columns.commission) : 0,
      edge: columns.edge >= 0 ? col(row, columns.edge) : 0,
      net: columns.net >= 0 ? col(row, columns.net) : 0,
    });
  }
  if (!tickets.booked.length) throw new Error("The vendor payment advice has no booked-ticket detail section.");
  return { operator, invoiceNo, paymentDate, period, vendorStyle: "abhibus" as const, summary, amountSettled, tickets };
}

/* ----------------------------- client ledger ----------------------------- */

type ClientRow = {
  type: "B" | "C";
  pnr: string;
  route: string;
  journey: string;
  seats: number;
  discount: number;
  amount: number;
  commission: number;
  debit: number;
  credit: number;
};

function parseClient(rows: string[][]) {
  const legacyHeaderIndex = rows.findIndex((row) => cellAt(row, 0) === "No." && /particular/i.test(cellAt(row, 2)));
  const adviceHeaderIndex = rows.findIndex(
    (row) => /^(TIN|Ticket\s*(No|Number))$/i.test(cellAt(row, 0)) && /^PNR$/i.test(cellAt(row, 1)) && row.some((value) => /ticket\s*fare/i.test(value)),
  );
  const headerIndex = legacyHeaderIndex >= 0 ? legacyHeaderIndex : adviceHeaderIndex;
  if (headerIndex < 0) {
    const looksLikeVendor = rows.some((row) => /Payment Advice/i.test(cellAt(row, 0)));
    throw new Error(
      looksLikeVendor
        ? "The client slot holds a vendor payment advice. Please swap the two files."
        : "This CSV is not a supported client transaction report (expected a RedBus ledger or TIN / PNR transaction table).",
    );
  }

  const top = rows[0] ?? [];
  const agentAt = top.findIndex((value) => /^Agent\s*:/i.test(value.trim()));
  const agent = agentAt >= 0 ? cellAt(top, agentAt + 1) : "";
  const periodMatch = top.join(" ").match(/From\s*:\s*(\S+)\s*To\s*:\s*(\S+)/i);
  const clientPeriod = periodMatch ? `${periodMatch[1]} to ${periodMatch[2]}` : "";
  const company = cellAt(top, 0);
  const data: ClientRow[] = [];
  let printedDebit: number | undefined;
  let printedCredit: number | undefined;

  if (legacyHeaderIndex >= 0) {
    for (const row of rows.slice(legacyHeaderIndex + 1)) {
      const first = cellAt(row, 0);
      if (/^Total\s*:/i.test(first)) {
        // The ITS "Total :" row drops the date/particulares/route columns, so
        // its values start at the Seat column — debit sits at index 7 and
        // credit at index 8, not under the Debit/Credit header positions.
        printedDebit = money(cellAt(row, 7));
        printedCredit = money(cellAt(row, 8));
        continue;
      }
      if (!/^\d+$/.test(first)) continue;
      const match = cellAt(row, 2).match(/^([BC])\s*:\s*(\S+)/i);
      if (!match) continue;
      data.push({
        type: match[1].toUpperCase() as "B" | "C",
        pnr: match[2],
        route: cellAt(row, 3),
        journey: cellAt(row, 4),
        seats: Math.abs(money(cellAt(row, 5))),
        discount: money(cellAt(row, 6)),
        amount: money(cellAt(row, 7)),
        commission: money(cellAt(row, 9)),
        debit: money(cellAt(row, 11)),
        credit: money(cellAt(row, 12)),
      });
    }
  } else {
    // TIN/PNR exports can repeat the header for several sections (Payment,
    // Cancellation, Exceptional Refund, ...). Track each header plus the
    // section title above it so cancellation rows are typed C, not B.
    const isHeaderRow = (row: string[]) =>
      /^(TIN|Ticket\s*(No|Number))$/i.test(cellAt(row, 0)) && /^PNR$/i.test(cellAt(row, 1)) && row.some((value) => /ticket\s*fare/i.test(value));
    const sectionLabel = (headerIndex: number) => {
      for (let index = headerIndex - 1; index >= 0; index -= 1) {
        const first = cellAt(rows[index], 0);
        if (first && !isHeaderRow(rows[index])) return first;
      }
      return "";
    };
    let cols = { pnr: -1, date: -1, route: -1, seat: -1, fare: -1, commission: -1, net: -1, offer: -1 };
    let sectionIsCancel = false;
    const applyHeader = (header: string[], headerIndex: number) => {
      const find = (pattern: RegExp) => header.findIndex((value) => pattern.test(value.trim()));
      const date = find(/^DOM$/i) >= 0 ? find(/^DOM$/i) : find(/^DOE$/i);
      cols = { pnr: find(/^PNR$/i), date, route: find(/^Route$/i), seat: find(/^Seats$/i), fare: find(/^Ticket\s*Fare$/i), commission: find(/^Total\s*Commission\s*With\s*GST$/i), net: find(/^NetPayable$/i), offer: find(/^Others$/i) };
      sectionIsCancel = /(cancell|refund)/i.test(sectionLabel(headerIndex));
    };
    applyHeader(rows[adviceHeaderIndex], adviceHeaderIndex);
    for (const [rowIndex, row] of rows.slice(adviceHeaderIndex + 1).entries()) {
      if (isHeaderRow(row)) {
        applyHeader(row, adviceHeaderIndex + 1 + rowIndex);
        continue;
      }
      const pnr = cellAt(row, cols.pnr);
      if (!pnr || /^total/i.test(pnr) || !cellAt(row, 0)) continue;
      const fare = money(cellAt(row, cols.fare));
      const net = money(cellAt(row, cols.net));
      const commission = money(cellAt(row, cols.commission));
      data.push({
        type: sectionIsCancel ? "C" : "B",
        pnr,
        route: cellAt(row, cols.route),
        journey: cellAt(row, cols.date),
        seats: Math.max(1, cellAt(row, cols.seat).split(",").filter(Boolean).length),
        discount: money(cellAt(row, cols.offer)),
        amount: fare,
        commission,
        // For cancellation sections the NetPayable is the charge the client
        // keeps (debit); the refunded remainder of the fare is the credit.
        debit: net,
        credit: sectionIsCancel ? round(Math.max(0, fare - net)) : 0,
      });
    }
  }
  if (!data.length) throw new Error("The client transaction report has no booking or cancellation rows.");
  return { agent, company, clientPeriod, data, printedDebit, printedCredit };
}

/* ------------------------------ reconciliation ------------------------------ */

const sum = <T,>(items: T[], pick: (item: T) => number) => round(items.reduce((total, item) => total + pick(item), 0));

function groupBy<T extends { pnr: string }>(items: T[]) {
  const map = new Map<string, T[]>();
  for (const item of items) map.set(item.pnr, [...(map.get(item.pnr) ?? []), item]);
  return map;
}

export type ReconcileOptions = {
  // Actual bank receipt for the period — when supplied, DIFF compares the
  // computed receivable against it instead of the invoice's settled figure.
  received?: number;
  // Hand-keyed deductions from their SUMMARY sheet that do not exist in the
  // vendor advice.
  featuredCard?: number;
  sponsorship?: number;
  // YourBus device-subscription charge, read from the RedBus advice's own
  // "YourBus Subscription Detail" sheet when the vendor file is an .xls.
  yourbus?: number;
  // AbhiBus shares cancellation charges 50/50 (their SHERING 50% column).
  shareCharges?: boolean;
};

// Parse dd-mm-yyyy / dd/mm/yyyy (optionally followed by a time) into a Date.
const parseDoj = (raw: string) => {
  const m = raw.match(/(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (!m) return null;
  const year = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
  const date = new Date(year, Number(m[2]) - 1, Number(m[1]));
  return Number.isNaN(date.getTime()) ? null : date;
};
const weekLabel = (date: Date) => {
  const start = new Date(date);
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7)); // Monday
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  const f = (d: Date) => d.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
  return `${f(start)} → ${f(end)}`;
};

export function reconcileCsv(vendorText: string, clientText: string, vendorName: string, clientName: string, options: ReconcileOptions = {}): Reconciliation {
  const vendor = parseVendor(parseCsv(vendorText));
  const client = parseClient(parseCsv(clientText));
  const v = vendor.summary;

  const vBooked = groupBy(vendor.tickets.booked);
  const vCancel = groupBy(vendor.tickets.cancel);
  const vCharges = groupBy(vendor.tickets.charges);
  const vExceptional = groupBy(vendor.tickets.exceptional);
  const vAssured = groupBy(vendor.tickets.assured);
  const vBusCancelled = groupBy(vendor.tickets.busCancelled);
  const cBooked = groupBy(client.data.filter((row) => row.type === "B"));
  const cCancel = groupBy(client.data.filter((row) => row.type === "C"));

  const clientBookings = client.data.filter((row) => row.type === "B");
  const clientCancels = client.data.filter((row) => row.type === "C");
  const clientDebit = sum(client.data, (row) => row.debit);
  const clientCredit = sum(client.data, (row) => row.credit);
  const clientNet = round(clientDebit - clientCredit);
  const clientBookingDebit = sum(clientBookings, (row) => row.debit);
  const clientCancelDebit = sum(clientCancels, (row) => row.debit);

  const vendorBookedAmount = sum(vendor.tickets.booked, (t) => t.amount);
  const vendorBookedCommission = sum(vendor.tickets.booked, (t) => t.commission);
  const vendorBookedEdge = sum(vendor.tickets.booked, (t) => t.edge);
  const vendorCancelAmount = sum(vendor.tickets.cancel, (t) => t.amount);
  // AbhiBus lists cancellation charges in their own detail section; RedBus
  // carries the charge on each cancellation row.
  const vendorChargeAmount = round(
    sum(vendor.tickets.charges, (t) => t.amount) + sum(vendor.tickets.cancel, (t) => t.charge ?? 0),
  );
  const vendorExceptionalTotal = sum([...vendor.tickets.exceptional, ...vendor.tickets.assured], (t) => t.net || t.amount);
  const clientBookedAmount = sum(clientBookings, (row) => row.amount);
  const clientBookedCommission = sum(clientBookings, (row) => row.commission);
  const clientBookedDiscount = sum(clientBookings, (row) => row.discount);
  const clientCancelAmount = round(Math.abs(sum(clientCancels, (row) => row.amount)));

  /* per-PNR exceptions */
  const exceptions: Reconciliation["exceptions"] = [];
  const flag = (pnr: string, category: string, issue: string, vendorAmount: number, clientAmount: number, detail: string) =>
    exceptions.push({ pnr, category, issue, vendorAmount: round(vendorAmount), clientAmount: round(clientAmount), variance: round(clientAmount - vendorAmount), detail });

  const bookingPnrs = new Set([...vBooked.keys(), ...cBooked.keys()]);
  let matchedBookings = 0;
  for (const pnr of bookingPnrs) {
    const vRows = vBooked.get(pnr);
    const cRows = cBooked.get(pnr);
    if (vRows && !cRows) {
      flag(pnr, "Booking", "In vendor advice, missing from client ledger", sum(vRows, (t) => t.amount), 0, `Ticket ${vRows[0].ticket} · DOJ ${vRows[0].doj} · seats ${vRows[0].seats || "n/a"}`);
      continue;
    }
    if (cRows && !vRows) {
      flag(pnr, "Booking", "In client ledger, missing from vendor advice", 0, sum(cRows, (r) => r.amount), `${cRows[0].route} · journey ${cRows[0].journey}`);
      continue;
    }
    if (!vRows || !cRows) continue;
    matchedBookings += 1;
    const where = `${cRows[0].route} · journey ${cRows[0].journey}`;
    const pairs: [string, number, number][] = [
      ["Ticket amount differs", sum(vRows, (t) => t.amount), sum(cRows, (r) => r.amount)],
    ];
    // RedBus settles its own commission rate and does not deduct the agent's
    // customer discount, so only AbhiBus commission/edge figures are
    // comparable to the ITS Comm./Discount columns.
    if (vendor.vendorStyle !== "redbus") {
      pairs.push(
        ["Commission differs", sum(vRows, (t) => t.commission), sum(cRows, (r) => r.commission)],
        ["Edge offer / discount differs", sum(vRows, (t) => t.edge), sum(cRows, (r) => r.discount)],
      );
    }
    for (const [issue, vendorAmount, clientAmount] of pairs) {
      if (!close(vendorAmount, clientAmount, 0.5)) flag(pnr, "Booking", issue, vendorAmount, clientAmount, where);
    }
  }

  const cancelPnrs = new Set([...vCancel.keys(), ...cCancel.keys()]);
  let matchedCancellations = 0;
  for (const pnr of cancelPnrs) {
    const vRows = vCancel.get(pnr);
    const cRows = cCancel.get(pnr);
    const vendorCharge = sum(vCharges.get(pnr) ?? [], (t) => t.amount) + sum(vRows ?? [], (t) => t.charge ?? 0);
    if (vRows && !cRows) {
      const booked = cBooked.get(pnr)?.[0];
      flag(pnr, "Cancellation", "Cancelled in vendor advice, no cancellation in client ledger", sum(vRows, (t) => t.amount), 0, `${booked ? `${booked.route} · ` : ""}vendor cancellation charge ${fmt(vendorCharge)}${vRows[0].note ? ` · ${vRows[0].note}` : ""}`);
      continue;
    }
    if (cRows && !vRows) {
      // RedBus settles cancelled-vehicle PNRs in its Bus Cancelled section as
      // a full reversal rather than a normal cancellation.
      if (vBusCancelled.has(pnr)) continue;
      flag(pnr, "Cancellation", "Cancelled in client ledger, not in vendor advice", 0, Math.abs(sum(cRows, (r) => r.amount)), `${cRows[0].route} · client credit ${fmt(sum(cRows, (r) => r.credit))}`);
      continue;
    }
    if (!vRows || !cRows) continue;
    matchedCancellations += 1;
    const vendorAmount = sum(vRows, (t) => t.amount);
    const clientAmount = Math.abs(sum(cRows, (r) => r.amount));
    if (!close(vendorAmount, clientAmount, 0.5)) flag(pnr, "Cancellation", "Cancelled amount differs", vendorAmount, clientAmount, cRows[0].route);
    const clientCharge = sum(cRows, (r) => r.debit);
    if (!close(vendorCharge, clientCharge, 0.5)) flag(pnr, "Cancellation", "Cancellation charge differs", vendorCharge, clientCharge, cRows[0].route);
  }
  for (const [pnr, rows] of vCharges) {
    if (!vCancel.has(pnr)) flag(pnr, "Cancellation", "Vendor cancellation charge without a cancelled ticket", sum(rows, (t) => t.amount), 0, `Ticket ${rows[0].ticket}`);
  }

  const deductions: [string, Map<string, VendorTicket[]>][] = [
    ["Exceptional refund", vExceptional],
    ["Assured", vAssured],
    ["Bus cancelled", vBusCancelled],
  ];
  for (const [category, map] of deductions) {
    for (const [pnr, rows] of map) {
      const clientRows = cCancel.get(pnr);
      const vendorAmount = sum(rows, (t) => t.net || t.amount);
      if (clientRows && category === "Bus cancelled") {
        // The vendor reversed the ticket in full (cancelled vehicle) while
        // ITS recorded a normal cancellation keeping a charge — a real
        // difference, not a missing record.
        flag(pnr, category, "Vendor reversed in full (bus cancelled); client kept a cancellation charge", vendorAmount, sum(clientRows, (r) => r.credit), `Ticket ${rows[0].ticket}${rows[0].note ? ` · ${rows[0].note}` : ""} · client kept debit ${fmt(sum(clientRows, (r) => r.debit))}`);
        continue;
      }
      const clientCharge = sum(clientRows ?? [], (r) => r.debit);
      const note = clientRows ? `client ledger shows a cancellation charge of ${fmt(clientCharge)} for this PNR` : "no client cancellation for this PNR";
      flag(pnr, category, "Deducted by vendor, no client equivalent", vendorAmount, 0, `Ticket ${rows[0].ticket}${rows[0].note ? ` · ${rows[0].note}` : ""} · ${note}`);
    }
  }
  exceptions.sort((a, b) => Math.abs(b.variance) - Math.abs(a.variance));

  /* category comparison (client − vendor) */
  const row = (label: string, vendorAmount: number, clientAmount: number, explanation: string) => ({
    label,
    vendorAmount: round(vendorAmount),
    clientAmount: round(clientAmount),
    variance: round(clientAmount - vendorAmount),
    explanation,
    status: close(vendorAmount, clientAmount) ? ("matched" as const) : ("review" as const),
  });
  const vendorOnlyBookings = [...vBooked.keys()].filter((pnr) => !cBooked.has(pnr)).length;
  const clientOnlyBookings = [...cBooked.keys()].filter((pnr) => !vBooked.has(pnr)).length;
  const vendorOnlyCancels = [...vCancel.keys()].filter((pnr) => !cCancel.has(pnr));
  const clientOnlyCancels = [...cCancel.keys()].filter((pnr) => !vCancel.has(pnr) && !vBusCancelled.has(pnr));
  const isRed = vendor.vendorStyle === "redbus";
  const differences: Reconciliation["differences"] = [
    row("Booked ticket amount", vendorBookedAmount, clientBookedAmount, `${matchedBookings.toLocaleString("en-IN")} of ${vBooked.size.toLocaleString("en-IN")} vendor tickets matched to the client ledger by PNR; ${vendorOnlyBookings} vendor-only, ${clientOnlyBookings} client-only.`),
    ...(isRed
      ? []
      : [
          row("Booking commission", vendorBookedCommission, clientBookedCommission, "Commission on booked tickets, compared PNR by PNR."),
          row("Edge offer / discount on bookings", vendorBookedEdge, clientBookedDiscount, "Vendor Edge Offer (G1) against the client ledger Discount column. The vendor deducts this from the payable; the client ledger does not."),
        ]),
    row("Cancelled ticket amount", vendorCancelAmount, clientCancelAmount, `Vendor cancellation and service-cancel tickets (${vCancel.size}) against client cancellations (${cCancel.size}). ${vendorOnlyCancels.length} vendor-only, ${clientOnlyCancels.length} client-only.`),
    row("Cancellation charges", vendorChargeAmount, clientCancelDebit, "Charges the operator keeps on cancelled tickets: vendor Cancellation Charges section against the client ledger debit on cancellation rows."),
    row("Exceptional refunds + assured", vendorExceptionalTotal, 0, "Vendor-only deductions (post-journey service cancels and assured payouts). The client ledger has no equivalent line."),
    ...(isRed
      ? [row("Bus cancelled reversals", sum(vendor.tickets.busCancelled, (t) => t.net), sum(clientCancels.filter((r) => vBusCancelled.has(r.pnr)), (r) => r.credit), "PNRs the vendor reversed in full because the bus was cancelled, against the credit the client ledger recorded for the same PNRs.")]
      : []),
    row("Net settlement", v.final, clientNet, "Vendor final invoice against the client ledger net (debit − credit)."),
  ];

  /* settlement bridge: client net → vendor final */
  const bridge: Reconciliation["bridge"] = [{ label: "Client ledger net (debit − credit)", amount: clientNet, kind: "base" }];
  let running = clientNet;
  const step = (label: string, amount: number) => {
    const value = round(amount);
    if (Math.abs(value) < 0.005) return;
    running = round(running + value);
    bridge.push({ label, amount: value, kind: value < 0 ? "deduction" : "addition" });
  };
  // For RedBus the payable base is fare − GST-inclusive commission
  // (bookingPayable); for AbhiBus the booked amount minus its commission.
  step("Booking amount differences", (vendor.vendorStyle === "redbus" ? v.bookingPayable : round(v.booked - v.bookedCommission)) - clientBookingDebit);
  step("Cancellation refunds deducted by vendor beyond client credits", -(v.cancel + v.serviceCancel) + clientCredit);
  // AbhiBus adds cancellation charges back on top of the refund reversal;
  // RedBus's NetReversed already nets the kept charge out, so adding it
  // again would double count.
  if (vendor.vendorStyle !== "redbus") step("Cancellation charges kept by vendor beyond client records", v.charges - clientCancelDebit);
  step("Edge offer deducted by vendor (net of commission)", -round(v.edgeGross - v.edgeCommission));
  step("Exceptional refunds and assured deducted by vendor", -v.exceptional);
  step("TDS 194(O) deducted by vendor", -v.tds);
  const residual = round(v.final - running);
  step(Math.abs(residual) <= 2 ? "Rounding in vendor invoice" : "Unexplained difference", residual);
  bridge.push({ label: "Vendor final settlement", amount: v.final, kind: "final" });

  /* Humsafar SUMMARY-sheet formula (verbatim from their workbook):
     R = Booking − Comm − TDS − Cancel + CancelTDS + CancelCharges
         − Assured − Exceptional − PostJD ServiceCancel − Edge
         + EdgeComm(10%) − FeaturedCard − Sponsorship.
     Cancel TDS, featured-card and sponsorship deductions are keyed in
     manually on their sheet and do not exist in the vendor advice, so
     they are shown as zero lines. DIFF = RECEIVABLE − RECEIVED. */
  const receivableRows: { label: string; amount: number }[] = [];
  const recv = (label: string, amount: number) => receivableRows.push({ label, amount: round(amount) });
  if (vendor.vendorStyle === "redbus") {
    recv("Bookings payable (fare − commission inc. GST)", v.bookingPayable);
    recv("Cancellation reversals", -v.cancel);
    recv("Bus cancelled reversals", -v.serviceCancel);
    recv("Exceptional refunds", -v.exceptional);
    recv("Other deductions", -v.edgePayable);
  } else {
    recv("Booking", v.booked);
    recv("Booking commission", -v.bookedCommission);
    recv("TDS", -v.tds);
    recv("Cancel", -v.cancel);
    recv("Cancellation charges", v.charges);
    recv("Assured + exceptional refunds", -v.exceptional);
    recv("Post-journey service cancel", -v.serviceCancel);
    recv("Abhi Edge", -v.edgeGross);
    recv("Edge comm (10%)", v.edgeGross * 0.1);
    recv("Featured card deduction", 0);
    recv("Sponsorship deduction", 0);
  }
  // Manual overrides: 50% charge sharing, hand-keyed featured-card and
  // sponsorship deductions, and the actual bank receipt when supplied.
  if (options.shareCharges && vendor.vendorStyle === "abhibus") {
    const idx = receivableRows.findIndex((r) => r.label === "Cancellation charges");
    if (idx >= 0) {
      receivableRows[idx] = { label: "Cancellation charges (shared 50%)", amount: round(v.charges / 2) };
    }
  }
  if (options.featuredCard) {
    const idx = receivableRows.findIndex((r) => r.label === "Featured card deduction");
    if (idx >= 0) receivableRows[idx].amount = -Math.abs(options.featuredCard);
    else receivableRows.push({ label: "Featured card deduction", amount: -Math.abs(options.featuredCard) });
  }
  if (options.sponsorship) {
    const idx = receivableRows.findIndex((r) => r.label === "Sponsorship deduction");
    if (idx >= 0) receivableRows[idx].amount = -Math.abs(options.sponsorship);
    else receivableRows.push({ label: "Sponsorship deduction", amount: -Math.abs(options.sponsorship) });
  }
  if (options.yourbus) {
    receivableRows.push({ label: "YourBus subscription charge", amount: -Math.abs(options.yourbus) });
  }
  const receivableTotal = round(sum(receivableRows, (r) => r.amount));
  const amountReceived = options.received !== undefined && options.received > 0 ? options.received : vendor.amountSettled || v.final;
  const receivable = {
    rows: receivableRows,
    receivable: receivableTotal,
    received: round(amountReceived),
    difference: round(receivableTotal - amountReceived),
  };

  /* per-PNR joined audit — the manual sheet's VLOOKUP columns, for every PNR */
  const sectionName = (key: string) =>
    vendor.vendorStyle === "redbus"
      ? ({ booked: "Payment", cancel: "Cancellation", busCancelled: "Bus Cancelled", exceptional: "Exceptional Refund", assured: "Assured" } as Record<string, string>)[key] ?? key
      : ({ booked: "Booked", cancel: "Cancellation", busCancelled: "Bus Cancelled", exceptional: "Exceptional refunds", assured: "Assured" } as Record<string, string>)[key] ?? key;
  const vendorSectionOf = (pnr: string) => {
    if (vBooked.has(pnr)) return "booked";
    if (vCancel.has(pnr)) return "cancel";
    if (vBusCancelled.has(pnr)) return "busCancelled";
    if (vExceptional.has(pnr)) return "exceptional";
    if (vAssured.has(pnr)) return "assured";
    return "";
  };
  const vendorAmountFor = (key: string, pnr: string) => {
    const list =
      key === "booked" ? vBooked.get(pnr) : key === "cancel" ? vCancel.get(pnr) : key === "busCancelled" ? vBusCancelled.get(pnr) : key === "exceptional" ? vExceptional.get(pnr) : vAssured.get(pnr);
    return sum(list ?? [], (t) => (key === "booked" ? t.amount : t.net || t.amount));
  };
  const vendorTicketFor = (key: string, pnr: string) =>
    (key === "booked" ? vBooked.get(pnr) : key === "cancel" ? vCancel.get(pnr) : key === "busCancelled" ? vBusCancelled.get(pnr) : key === "exceptional" ? vExceptional.get(pnr) : vAssured.get(pnr))?.[0];
  const audit: Reconciliation["audit"] = [];
  // Settlement week = the advice/ITS period start. PNRs whose journey date
  // sits in an earlier week are arrears the vendor carried into this run.
  const datedPeriod = /\d{2}-\d{2}-\d{4}/.test(vendor.period) ? vendor.period : client.clientPeriod;
  const periodStart = parseDoj(datedPeriod);
  const weekStart = (d: Date) => { const s = new Date(d); s.setDate(s.getDate() - ((s.getDay() + 6) % 7)); s.setHours(0, 0, 0, 0); return s; };
  const settleWeekStart = periodStart ? weekStart(periodStart).getTime() : null;
  const auditPnrs = new Set<string>([...vBooked.keys(), ...vCancel.keys(), ...vBusCancelled.keys(), ...vExceptional.keys(), ...vAssured.keys(), ...cBooked.keys(), ...cCancel.keys()]);
  for (const pnr of auditPnrs) {
    const cB = cBooked.get(pnr);
    const cC = cCancel.get(pnr);
    const clientType = cB ? "B" : cC ? "C" : "—";
    const vKey = vendorSectionOf(pnr);
    const clientAmount = round(cB ? sum(cB, (r) => r.amount) : cC ? sum(cC, (r) => r.debit - r.credit) : 0);
    const vendorAmount = vKey ? round(vendorAmountFor(vKey, pnr)) : 0;
    const ref = vKey ? vendorTicketFor(vKey, pnr) : undefined;
    const cRef = cB?.[0] ?? cC?.[0];
    // Clarification fields: what each side booked, charged and deducted for
    // this PNR, so every column of their manual check is answered on one row.
    const vB = vBooked.get(pnr) ?? [];
    const vC = vCancel.get(pnr) ?? [];
    const vendorFare = round(sum(vB, (t) => t.amount) + sum(vC, (t) => t.amount) + sum(vBusCancelled.get(pnr) ?? [], (t) => t.amount));
    const clientFare = round((cB ? sum(cB, (r) => r.amount) : 0) + (cC ? Math.abs(sum(cC, (r) => r.amount)) : 0));
    const vendorComm = round(sum(vB, (t) => t.commission));
    const clientComm = round(cB ? sum(cB, (r) => r.commission) : 0);
    const vendorOffer = round(sum(vB, (t) => t.edge));
    const clientOffer = round(cB ? sum(cB, (r) => r.discount) : 0);
    const vendorCharge = round(sum(vCharges.get(pnr) ?? [], (t) => t.amount) + sum(vC, (t) => t.charge ?? 0));
    const clientCharge = round(cC ? sum(cC, (r) => r.debit) : 0);
    const cancelType = vBusCancelled.has(pnr)
      ? "Bus cancelled"
      : vExceptional.has(pnr) || vAssured.has(pnr)
        ? "Exceptional refund"
        : vCancel.has(pnr)
          ? vC[0]?.note?.match(/reschedul|normal|partial|full/i)?.[0] ?? "Normal cancellation"
          : cC
            ? "Client cancellation only"
            : "—";
    audit.push({
      pnr,
      clientType,
      vendorSection: vKey ? sectionName(vKey) : "—",
      clientAmount,
      vendorAmount,
      variance: round(clientAmount - vendorAmount),
      vendorFare,
      clientFare,
      fareDiff: round(clientFare - vendorFare),
      vendorComm,
      clientComm,
      commDiff: round(clientComm - vendorComm),
      vendorOffer,
      clientOffer,
      offerDiff: round(clientOffer - vendorOffer),
      vendorCharge,
      clientCharge,
      chargeDiff: round(clientCharge - vendorCharge),
      cancelType,
      rescheduled: Boolean(ref?.rescheduled) || /reschedul/i.test(ref?.note ?? ""),
      channel: ref?.channel ?? "",
      carried: (() => { const d = parseDoj(ref?.doj ?? cRef?.journey ?? ""); return Boolean(d && settleWeekStart !== null && weekStart(d).getTime() < settleWeekStart); })(),
      carriedWeek: (() => { const d = parseDoj(ref?.doj ?? cRef?.journey ?? ""); return d ? weekLabel(d) : ""; })(),
      ticket: ref?.ticket ?? "",
      route: cRef?.route ?? ref?.route ?? "",
      doj: ref?.doj ?? cRef?.journey ?? "",
      settleDate: ref?.settle ?? "",
      note: ref?.note ?? "",
      status: !vKey ? "client-only" : !cB && !cC ? "vendor-only" : close(clientAmount, vendorAmount) ? "matched" : "differs",
    });
  }
  audit.sort((a, b) => Math.abs(b.variance) - Math.abs(a.variance));

  /* weekly settlement periods, bucketed by journey date */
  const periodBuckets = new Map<string, { start: Date; booked: number; cancel: number; serviceCancel: number; charges: number; exceptional: number; edge: number; tickets: number }>();
  const addToPeriod = (doj: string, field: "booked" | "cancel" | "serviceCancel" | "charges" | "exceptional" | "edge", amount: number) => {
    const date = parseDoj(doj);
    if (!date) return;
    const key = weekLabel(date);
    const bucket = periodBuckets.get(key) ?? { start: new Date(date.setDate(date.getDate() - ((date.getDay() + 6) % 7))), booked: 0, cancel: 0, serviceCancel: 0, charges: 0, exceptional: 0, edge: 0, tickets: 0 };
    bucket[field] += amount;
    if (field === "booked") bucket.tickets += 1;
    periodBuckets.set(key, bucket);
  };
  for (const t of vendor.tickets.booked) addToPeriod(t.doj, "booked", t.net);
  for (const t of vendor.tickets.cancel) addToPeriod(t.doj, "cancel", t.net || t.amount);
  for (const t of vendor.tickets.busCancelled) addToPeriod(t.doj, "serviceCancel", t.net || t.amount);
  for (const t of vendor.tickets.charges) addToPeriod(t.doj, "charges", t.amount);
  for (const t of [...vendor.tickets.exceptional, ...vendor.tickets.assured]) addToPeriod(t.doj, "exceptional", t.net || t.amount);
  for (const t of vendor.tickets.booked) addToPeriod(t.doj, "edge", t.edge);
  const periodEnd = (start: Date) => { const e = new Date(start); e.setDate(e.getDate() + 6); return e.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }); };
  const periods: Reconciliation["periods"] = [...periodBuckets.entries()]
    .sort((a, b) => a[1].start.getTime() - b[1].start.getTime())
    .map(([label, b]) => {
      const chargeShare = options.shareCharges ? b.charges / 2 : b.charges;
      // AbhiBus Net Amount already nets commission and edge inside each row,
      // so edge is not deducted again; RedBus NetPayable nets commission but
      // not the Others/edge column.
      const receivable =
        vendor.vendorStyle === "redbus"
          ? round(b.booked - b.cancel - b.serviceCancel - b.exceptional - b.edge)
          : round(b.booked - b.cancel - b.serviceCancel + chargeShare - b.exceptional);
      return {
        label,
        start: b.start.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }),
        end: periodEnd(b.start),
        booked: round(b.booked),
        cancel: round(b.cancel),
        serviceCancel: round(b.serviceCancel),
        charges: round(b.charges),
        exceptional: round(b.exceptional),
        edge: round(b.edge),
        receivable,
        tickets: b.tickets,
      };
    });

  /* control checks */
  // AbhiBus nets cancellations inside the bookings payable line; RedBus
  // deducts each section (cancellation reversal, bus-cancel reversal,
  // exceptional refund) on top of the payment section.
  const vendorMath =
    vendor.vendorStyle === "redbus"
      ? round(v.bookingPayable - v.cancel - v.serviceCancel - v.exceptional - v.edgePayable)
      : round(v.bookingPayable - v.exceptional - v.edgePayable);
  const vendorMathDetail =
    vendor.vendorStyle === "redbus"
      ? `Bookings payable ${fmt(v.bookingPayable)} − cancellation reversals ${fmt(v.cancel)} − bus-cancel reversals ${fmt(v.serviceCancel)} − exceptional refunds ${fmt(v.exceptional)} − others ${fmt(v.edgePayable)} = ${fmt(vendorMath)} against computed settlement ${fmt(v.final)}.`
      : `Bookings payable ${fmt(v.bookingPayable)} − exceptional ${fmt(v.exceptional)} − edge offer ${fmt(v.edgePayable)} = ${fmt(vendorMath)} against invoice ${fmt(v.final)}.`;
  const vendorCancelNet = round(sum(vendor.tickets.cancel, (t) => t.net) + sum(vendor.tickets.busCancelled, (t) => t.net));
  const checks: Reconciliation["checks"] = [
    { label: "Vendor invoice math", status: close(vendorMath, v.final) ? "matched" : "review", detail: vendorMathDetail, variance: round(vendorMath - v.final) },
    { label: "Vendor detail ties to summary", status: close(vendorBookedAmount, v.booked) && close(vendor.vendorStyle === "redbus" ? vendorCancelNet : vendorCancelAmount, v.cancel + v.serviceCancel) && close(vendorChargeAmount, v.charges) ? "matched" : "review", detail: `Booked ${fmt(vendorBookedAmount)} / ${fmt(v.booked)}, cancelled ${fmt(vendor.vendorStyle === "redbus" ? vendorCancelNet : vendorCancelAmount)} / ${fmt(v.cancel + v.serviceCancel)}, charges ${fmt(vendorChargeAmount)} / ${fmt(v.charges)} (ticket detail / summary; rupee rounding allowed).` },
    { label: "Client ledger totals", status: (client.printedDebit === undefined || close(clientDebit, client.printedDebit)) && (client.printedCredit === undefined || close(clientCredit, client.printedCredit)) ? "matched" : "review", detail: `Rows add up to debit ${fmt(clientDebit)} and credit ${fmt(clientCredit)}${client.printedDebit !== undefined ? `; the report prints ${fmt(client.printedDebit)} and ${fmt(client.printedCredit ?? 0)}` : ""}.` },
    { label: "Amount settled header", status: !vendor.amountSettled || close(vendor.amountSettled, v.final) ? "matched" : "review", detail: `Header shows ${fmt(vendor.amountSettled)} and the Total Invoice row shows ${fmt(v.final)}.`, variance: round(vendor.amountSettled - v.final) },
    { label: "Net receivable vs received", status: close(receivable.receivable, receivable.received, 2) ? "matched" : "review", detail: `SUMMARY-sheet formula gives ${fmt(receivable.receivable)} against ${fmt(receivable.received)} actually settled — DIFF ${fmt(receivable.difference)}.${vendor.vendorStyle === "abhibus" ? " Featured-card and sponsorship deductions are not in the advice and were treated as zero." : ""}`, variance: receivable.difference },
    { label: "Booking PNR coverage", status: vendorOnlyBookings + clientOnlyBookings === 0 ? "matched" : "review", detail: `${matchedBookings} PNRs on both sides, ${vendorOnlyBookings} only in the vendor advice, ${clientOnlyBookings} only in the client ledger.` },
    { label: "Cancellation PNR coverage", status: vendorOnlyCancels.length + clientOnlyCancels.length === 0 ? "matched" : "review", detail: `${matchedCancellations} cancellations on both sides, ${vendorOnlyCancels.length} only in the vendor advice, ${clientOnlyCancels.length} only in the client ledger.` },
    { label: "Reporting period", status: client.clientPeriod && client.clientPeriod === vendor.period ? "matched" : "info", detail: `Vendor ${vendor.period}; client ${client.clientPeriod || "not stated"}.` },
  ];

  const asPerIts = round(
    clientBookedAmount -
      clientBookedCommission -
      clientCancelAmount +
      clientCancelDebit -
      vendorExceptionalTotal,
  );
  const weekly: Reconciliation["weekly"] = {
    vendorBooking: round(vendorBookedAmount),
    itsBooking: round(clientBookedAmount),
    vendorComm: round(vendorBookedCommission),
    itsComm: round(clientBookedCommission),
    vendorCancel: round(vendorCancelAmount),
    busCancelled: round(v.serviceCancel),
    exceptional: round(vendorExceptionalTotal),
    itsCancel: round(clientCancelAmount),
    itsCharges: round(clientCancelDebit),
    vendorCharges: round(vendorChargeAmount),
    yourbus: round(options.yourbus ?? 0),
    total: round(v.final),
    received: receivable.received,
    diff: receivable.difference,
    asPerIts,
    paymentDiff: round(asPerIts - receivable.received),
  };

  return {
    source: { vendor: vendor.vendorStyle === "redbus" ? "RedBus" : vendor.vendorStyle === "abhibus" ? "AbhiBus" : "Vendor", vendorName, clientName, operator: vendor.operator, agent: client.agent || client.company, invoiceNo: vendor.invoiceNo, period: vendor.period, clientPeriod: client.clientPeriod, paymentDate: vendor.paymentDate },
    settlement: { vendorFinal: v.final, clientNet, gap: round(clientNet - v.final) },
    coverage: { vendorBookings: vBooked.size, clientBookings: cBooked.size, matchedBookings, vendorOnlyBookings, clientOnlyBookings, vendorCancellations: vCancel.size + vBusCancelled.size, clientCancellations: cCancel.size, matchedCancellations },
    differences,
    bridge,
    receivable,
    weekly,
    audit,
    periods,
    checks,
    exceptions: exceptions.slice(0, MAX_EXCEPTIONS),
    exceptionTotal: exceptions.length,
  };
}
