export type CheckStatus = "matched" | "review" | "info";

export type Reconciliation = {
  source: {
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

const MAX_EXCEPTIONS = 500;
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

type VendorKind = "booked" | "cancel" | "charges" | "exceptional" | "assured";
type VendorTicket = {
  ticket: string;
  pnr: string;
  doj: string;
  seats: string;
  amount: number;
  commission: number;
  edge: number;
  net: number;
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
    const adviceHeaderIndex = rows.findIndex(
      (row) => /^TIN$/i.test(cellAt(row, 0)) && /^PNR$/i.test(cellAt(row, 1)) && row.some((value) => /ticket\s*fare/i.test(value)),
    );
    if (adviceHeaderIndex >= 0) {
      const header = rows[adviceHeaderIndex];
      const column = (pattern: RegExp) => header.findIndex((value) => pattern.test(value.trim()));
      const pnrColumn = column(/^PNR$/i);
      const fareColumn = column(/^Ticket\s*Fare$/i);
      const commissionColumn = column(/^Total\s*Commission\s*With\s*GST$/i);
      const netColumn = column(/^NetPayable$/i);
      const offerColumn = column(/^Others$/i);
      const flatTickets: VendorTicket[] = [];
      for (const row of rows.slice(adviceHeaderIndex + 1)) {
        const ticket = cellAt(row, 0);
        const pnr = cellAt(row, pnrColumn);
        if (!ticket || !pnr || !/\d/.test(pnr) || fareColumn < 0) continue;
        flatTickets.push({
          ticket,
          pnr,
          doj: cellAt(row, column(/^DOE$/i) >= 0 ? column(/^DOE$/i) : column(/^DOM$/i)),
          seats: cellAt(row, column(/^Seats$/i)),
          amount: money(cellAt(row, fareColumn)),
          commission: commissionColumn >= 0 ? money(cellAt(row, commissionColumn)) : 0,
          edge: offerColumn >= 0 ? money(cellAt(row, offerColumn)) : 0,
          net: netColumn >= 0 ? money(cellAt(row, netColumn)) : 0,
        });
      }
      if (!flatTickets.length) throw new Error("The RedBus vendor payment advice has no ticket rows.");
      const total = (pick: (ticket: VendorTicket) => number) => round(flatTickets.reduce((sum, ticket) => sum + pick(ticket), 0));
      const booked = total((ticket) => ticket.amount);
      const commission = total((ticket) => ticket.commission);
      const edge = total((ticket) => ticket.edge);
      const final = total((ticket) => ticket.net || ticket.amount - ticket.commission - ticket.edge);
      return {
        operator: cellAt(rows[0], 0) || "RedBus payment advice",
        invoiceNo: cellAt(rows[0], 0),
        paymentDate: "",
        period: "RedBus payment advice period",
        summary: { booked, bookedCommission: commission, cancel: 0, serviceCancel: 0, charges: 0, exceptional: 0, edgeGross: edge, edgeCommission: 0, edgePayable: edge, bookingPayable: final, tds: 0, final },
        amountSettled: final,
        tickets: { booked: flatTickets, cancel: [], charges: [], exceptional: [], assured: [] },
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

  const tickets: Record<VendorKind, VendorTicket[]> = { booked: [], cancel: [], charges: [], exceptional: [], assured: [] };
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
      amount: col(row, columns.amount),
      commission: columns.commission >= 0 ? col(row, columns.commission) : 0,
      edge: columns.edge >= 0 ? col(row, columns.edge) : 0,
      net: columns.net >= 0 ? col(row, columns.net) : 0,
    });
  }
  if (!tickets.booked.length) throw new Error("The vendor payment advice has no booked-ticket detail section.");
  return { operator, invoiceNo, paymentDate, period, summary, amountSettled, tickets };
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
    const header = rows[adviceHeaderIndex];
    const column = (pattern: RegExp) => header.findIndex((value) => pattern.test(value.trim()));
    const pnrColumn = column(/^PNR$/i);
    const dateColumn = column(/^DOM$/i) >= 0 ? column(/^DOM$/i) : column(/^DOE$/i);
    const routeColumn = column(/^Route$/i);
    const seatColumn = column(/^Seats$/i);
    const fareColumn = column(/^Ticket\s*Fare$/i);
    const commissionColumn = column(/^Total\s*Commission\s*With\s*GST$/i);
    const netColumn = column(/^NetPayable$/i);
    const offerColumn = column(/^Others$/i);
    for (const row of rows.slice(adviceHeaderIndex + 1)) {
      const pnr = cellAt(row, pnrColumn);
      if (!pnr || /^total/i.test(pnr) || !/^\d+$/.test(cellAt(row, 1))) continue;
      const fare = money(cellAt(row, fareColumn));
      const net = money(cellAt(row, netColumn));
      const commission = money(cellAt(row, commissionColumn));
      data.push({
        type: "B",
        pnr,
        route: cellAt(row, routeColumn),
        journey: cellAt(row, dateColumn),
        seats: Math.max(1, cellAt(row, seatColumn).split(",").filter(Boolean).length),
        discount: money(cellAt(row, offerColumn)),
        amount: fare,
        commission,
        debit: net,
        credit: 0,
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

export function reconcileCsv(vendorText: string, clientText: string, vendorName: string, clientName: string): Reconciliation {
  const vendor = parseVendor(parseCsv(vendorText));
  const client = parseClient(parseCsv(clientText));
  const v = vendor.summary;

  const vBooked = groupBy(vendor.tickets.booked);
  const vCancel = groupBy(vendor.tickets.cancel);
  const vCharges = groupBy(vendor.tickets.charges);
  const vExceptional = groupBy(vendor.tickets.exceptional);
  const vAssured = groupBy(vendor.tickets.assured);
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
  const vendorChargeAmount = sum(vendor.tickets.charges, (t) => t.amount);
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
      ["Commission differs", sum(vRows, (t) => t.commission), sum(cRows, (r) => r.commission)],
      ["Edge offer / discount differs", sum(vRows, (t) => t.edge), sum(cRows, (r) => r.discount)],
    ];
    for (const [issue, vendorAmount, clientAmount] of pairs) {
      if (!close(vendorAmount, clientAmount, 0.5)) flag(pnr, "Booking", issue, vendorAmount, clientAmount, where);
    }
  }

  const cancelPnrs = new Set([...vCancel.keys(), ...cCancel.keys()]);
  let matchedCancellations = 0;
  for (const pnr of cancelPnrs) {
    const vRows = vCancel.get(pnr);
    const cRows = cCancel.get(pnr);
    const vendorCharge = sum(vCharges.get(pnr) ?? [], (t) => t.amount);
    if (vRows && !cRows) {
      const booked = cBooked.get(pnr)?.[0];
      flag(pnr, "Cancellation", "Cancelled in vendor advice, no cancellation in client ledger", sum(vRows, (t) => t.amount), 0, `${booked ? `${booked.route} · ` : ""}vendor cancellation charge ${fmt(vendorCharge)}`);
      continue;
    }
    if (cRows && !vRows) {
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

  const deductions: [string, Map<string, VendorTicket[]>][] = [["Exceptional refund", vExceptional], ["Assured", vAssured]];
  for (const [category, map] of deductions) {
    for (const [pnr, rows] of map) {
      const clientCharge = sum(cCancel.get(pnr) ?? [], (r) => r.debit);
      const note = cCancel.has(pnr) ? `client ledger shows a cancellation charge of ${fmt(clientCharge)} for this PNR` : "no client cancellation for this PNR";
      flag(pnr, category, "Deducted by vendor, no client equivalent", sum(rows, (t) => t.net || t.amount), 0, `Ticket ${rows[0].ticket} · ${note}`);
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
  const clientOnlyCancels = [...cCancel.keys()].filter((pnr) => !vCancel.has(pnr));
  const differences: Reconciliation["differences"] = [
    row("Booked ticket amount", vendorBookedAmount, clientBookedAmount, `${matchedBookings.toLocaleString("en-IN")} of ${vBooked.size.toLocaleString("en-IN")} vendor tickets matched to the client ledger by PNR; ${vendorOnlyBookings} vendor-only, ${clientOnlyBookings} client-only.`),
    row("Booking commission", vendorBookedCommission, clientBookedCommission, "Commission on booked tickets, compared PNR by PNR."),
    row("Edge offer / discount on bookings", vendorBookedEdge, clientBookedDiscount, "Vendor Edge Offer (G1) against the client ledger Discount column. The vendor deducts this from the payable; the client ledger does not."),
    row("Cancelled ticket amount", vendorCancelAmount, clientCancelAmount, `Vendor cancellation and service-cancel tickets (${vCancel.size}) against client cancellations (${cCancel.size}). ${vendorOnlyCancels.length} vendor-only, ${clientOnlyCancels.length} client-only.`),
    row("Cancellation charges", vendorChargeAmount, clientCancelDebit, "Charges the operator keeps on cancelled tickets: vendor Cancellation Charges section against the client ledger debit on cancellation rows."),
    row("Exceptional refunds + assured", vendorExceptionalTotal, 0, "Vendor-only deductions (post-journey service cancels and assured payouts). The client ledger has no equivalent line."),
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
  step("Booking amount differences", round(v.booked - v.bookedCommission) - clientBookingDebit);
  step("Cancellation refunds deducted by vendor beyond client credits", -(v.cancel + v.serviceCancel) + clientCredit);
  step("Cancellation charges kept by vendor beyond client records", v.charges - clientCancelDebit);
  step("Edge offer deducted by vendor (net of commission)", -round(v.edgeGross - v.edgeCommission));
  step("Exceptional refunds and assured deducted by vendor", -v.exceptional);
  step("TDS 194(O) deducted by vendor", -v.tds);
  const residual = round(v.final - running);
  step(Math.abs(residual) <= 2 ? "Rounding in vendor invoice" : "Unexplained difference", residual);
  bridge.push({ label: "Vendor final settlement", amount: v.final, kind: "final" });

  /* control checks */
  const vendorMath = round(v.bookingPayable - v.exceptional - v.edgePayable);
  const checks: Reconciliation["checks"] = [
    { label: "Vendor invoice math", status: close(vendorMath, v.final) ? "matched" : "review", detail: `Bookings payable ${fmt(v.bookingPayable)} − exceptional ${fmt(v.exceptional)} − edge offer ${fmt(v.edgePayable)} = ${fmt(vendorMath)} against invoice ${fmt(v.final)}.`, variance: round(vendorMath - v.final) },
    { label: "Vendor detail ties to summary", status: close(vendorBookedAmount, v.booked) && close(vendorCancelAmount, v.cancel + v.serviceCancel) && close(vendorChargeAmount, v.charges) ? "matched" : "review", detail: `Booked ${fmt(vendorBookedAmount)} / ${fmt(v.booked)}, cancelled ${fmt(vendorCancelAmount)} / ${fmt(v.cancel + v.serviceCancel)}, charges ${fmt(vendorChargeAmount)} / ${fmt(v.charges)} (ticket detail / summary; rupee rounding allowed).` },
    { label: "Client ledger totals", status: (client.printedDebit === undefined || close(clientDebit, client.printedDebit)) && (client.printedCredit === undefined || close(clientCredit, client.printedCredit)) ? "matched" : "review", detail: `Rows add up to debit ${fmt(clientDebit)} and credit ${fmt(clientCredit)}${client.printedDebit !== undefined ? `; the report prints ${fmt(client.printedDebit)} and ${fmt(client.printedCredit ?? 0)}` : ""}.` },
    { label: "Amount settled header", status: !vendor.amountSettled || close(vendor.amountSettled, v.final) ? "matched" : "review", detail: `Header shows ${fmt(vendor.amountSettled)} and the Total Invoice row shows ${fmt(v.final)}.`, variance: round(vendor.amountSettled - v.final) },
    { label: "Booking PNR coverage", status: vendorOnlyBookings + clientOnlyBookings === 0 ? "matched" : "review", detail: `${matchedBookings} PNRs on both sides, ${vendorOnlyBookings} only in the vendor advice, ${clientOnlyBookings} only in the client ledger.` },
    { label: "Cancellation PNR coverage", status: vendorOnlyCancels.length + clientOnlyCancels.length === 0 ? "matched" : "review", detail: `${matchedCancellations} cancellations on both sides, ${vendorOnlyCancels.length} only in the vendor advice, ${clientOnlyCancels.length} only in the client ledger.` },
    { label: "Reporting period", status: client.clientPeriod && client.clientPeriod === vendor.period ? "matched" : "info", detail: `Vendor ${vendor.period}; client ${client.clientPeriod || "not stated"}.` },
  ];

  return {
    source: { vendorName, clientName, operator: vendor.operator, agent: client.agent || client.company, invoiceNo: vendor.invoiceNo, period: vendor.period, clientPeriod: client.clientPeriod, paymentDate: vendor.paymentDate },
    settlement: { vendorFinal: v.final, clientNet, gap: round(clientNet - v.final) },
    coverage: { vendorBookings: vBooked.size, clientBookings: cBooked.size, matchedBookings, vendorOnlyBookings, clientOnlyBookings, vendorCancellations: vCancel.size, clientCancellations: cCancel.size, matchedCancellations },
    differences,
    bridge,
    checks,
    exceptions: exceptions.slice(0, MAX_EXCEPTIONS),
    exceptionTotal: exceptions.length,
  };
}
