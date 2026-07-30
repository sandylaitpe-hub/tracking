const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";               // 13.333 x 7.5
pres.author = "Investment Committee Memorandum";
pres.title = "Chateau in the Woods — Investment Committee Recommendation";

// ---- palette -------------------------------------------------------------
const DARK   = "16302A";   // deep forest, dark slides
const DARK2  = "1F423A";   // card on dark
const FOREST = "2C5F2D";
const MOSS   = "97BC62";
const AMBER  = "C8811F";
const RED    = "A32E22";
const INK    = "18231E";
const MUTED  = "6E7B73";
const LINE   = "D8DED9";
const TINT   = "F0F3EF";
const WHITE  = "FFFFFF";

const HEAD = "Cambria";
const BODY = "Calibri";

const sh = () => ({ type: "outer", color: "0B1512", blur: 8, offset: 2, angle: 90, opacity: 0.16 });

// =========================================================================
// SLIDE 1 — Executive summary, recommendation, returns
// =========================================================================
const s1 = pres.addSlide();
s1.background = { color: DARK };

s1.addText("Chateau in the Woods", {
  x: 0.5, y: 0.28, w: 8.6, h: 0.52, fontFace: HEAD, fontSize: 34, bold: true,
  color: WHITE, margin: 0,
});
s1.addText("118 Units  ·  4020 Monaco Drive, Indianapolis, IN 46220  ·  Built 1974  ·  134,371 SF  ·  96.6% occupied",
  { x: 0.5, y: 0.80, w: 9.2, h: 0.28, fontFace: BODY, fontSize: 11.5, color: MOSS, margin: 0 });

s1.addText([
  { text: "INVESTMENT COMMITTEE", options: { fontSize: 11, bold: true, color: WHITE, breakLine: true } },
  { text: "Recommendation  ·  Value-Add Acquisition", options: { fontSize: 9.5, color: MOSS } },
], { x: 9.9, y: 0.30, w: 2.93, h: 0.6, align: "right", fontFace: BODY, margin: 0 });

// ---- verdict -------------------------------------------------------------
s1.addShape(pres.ShapeType.roundRect, {
  x: 0.5, y: 1.20, w: 12.33, h: 0.80, fill: { color: AMBER }, rectRadius: 0.06, shadow: sh(),
});
s1.addText([
  { text: "CONDITIONAL:  ", options: { bold: true } },
  { text: "approve the plan and the capital — not the price." },
], { x: 0.80, y: 1.20, w: 8.0, h: 0.80, fontFace: HEAD, fontSize: 16, color: "2A1B04",
     valign: "middle", margin: 0 });
s1.addText([
  { text: "Bid $11.25mm ($95,339/unit).", options: { bold: true, breakLine: true } },
  { text: "Hard walk above $12.0mm." },
], { x: 8.85, y: 1.20, w: 3.70, h: 0.80, fontFace: BODY, fontSize: 11.5, color: "2A1B04",
     align: "right", valign: "middle", margin: 0 });

// ---- KPI row -------------------------------------------------------------
const kpis = [
  ["$11.25mm", "Recommended bid", "$95,339 per unit  ·  7.57% going-in cap", MOSS],
  ["14.2% / 1.83x", "At our bid", "Levered IRR / equity multiple", MOSS],
  ["3.1% / 1.15x", "At the $14.2mm ask", "Levered IRR / equity multiple", RED],
  ["1.1%", "Metro rent growth, actual", "Yardi Matrix 1Q-26  ·  T-3 rents flat", AMBER],
];
kpis.forEach(([big, lbl, sub, col], i) => {
  const x = 0.5 + i * 3.11;
  s1.addShape(pres.ShapeType.roundRect, {
    x, y: 2.22, w: 2.91, h: 1.06, fill: { color: DARK2 }, rectRadius: 0.05,
  });
  s1.addText(big, { x: x + 0.18, y: 2.30, w: 2.55, h: 0.42, fontFace: HEAD, fontSize: 21,
    bold: true, color: col, margin: 0 });
  s1.addText(lbl, { x: x + 0.18, y: 2.72, w: 2.55, h: 0.22, fontFace: BODY, fontSize: 10.5,
    bold: true, color: WHITE, margin: 0 });
  s1.addText(sub, { x: x + 0.18, y: 2.94, w: 2.58, h: 0.28, fontFace: BODY, fontSize: 8.5,
    color: "A9B6AE", margin: 0 });
});

// ---- executive summary ---------------------------------------------------
s1.addText("Executive Summary", { x: 0.5, y: 3.48, w: 6.1, h: 0.3, fontFace: HEAD, fontSize: 16,
  bold: true, color: WHITE, margin: 0 });

const summary = [
  "118 units, built 1974, Northside Indianapolis. 96.6% occupied, all units classic. Renovate all 118, add in-unit laundry to the 65 without it, modernise the common areas — $3.07mm, $26,030/unit.",
  "The renovation works. $15,754/unit earns a $166/month premium, a 12.7% return on cost. Renovated rents average $1.29/SF against a $1,310 metro average asking rent — Chateau's units are far larger than market, so we underwrite below the metro average per foot.",
  "But the market is not growing. Yardi Matrix 1Q-26: metro rents FLAT on a trailing-three-month basis, up just 1.1% year over year, and 2025 starts MORE THAN DOUBLED — delivering 2027-28, inside our hold, with Carmel taking over half of 2026 deliveries just north of us. We underwrite 1.5/2.0/2.5/3.0/3.0%, not a flat 3%.",
  "At the ask it is value-destructive. $14.2mm buys a 6.00% going-in cap and a Year-5 yield on cost of 6.58% — 17 bps BELOW our exit cap. The whole $3.07mm plan nets minus $6,880: the interior scope creates $1.06mm, deferred maintenance takes it back.",
  "Bid $11.25mm, best and final $11.5mm, hard walk $12.0mm. We will probably lose — that is the correct outcome at these fundamentals.",
];
s1.addText(summary.map((t, i) => ({
  text: t,
  options: { bullet: { characterCode: "25AA" }, breakLine: i !== summary.length - 1, paraSpaceAfter: 6 },
})), { x: 0.55, y: 3.80, w: 6.05, h: 3.20, fontFace: BODY, fontSize: 9, color: "DFE6E0",
       lineSpacing: 11.2, valign: "top", margin: 0 });

// ---- returns table -------------------------------------------------------
s1.addText("Returns", { x: 6.95, y: 3.48, w: 5.9, h: 0.3, fontFace: HEAD, fontSize: 16,
  bold: true, color: WHITE, margin: 0 });

const tHead = (t) => ({ text: t, options: { bold: true, color: WHITE, fontSize: 10,
  fill: { color: FOREST }, align: "center", valign: "middle" } });
const tRow = (a, b, c, hi) => ([
  { text: a, options: { color: hi ? WHITE : "DFE6E0", bold: !!hi, fontSize: 10 } },
  { text: b, options: { color: hi ? MOSS : "DFE6E0", bold: !!hi, align: "center", fontSize: 10 } },
  { text: c, options: { color: hi ? AMBER : "B9C4BC", bold: !!hi, align: "center", fontSize: 10 } },
]);

s1.addTable([
  [tHead("Five-year hold, Aug-26 to Jul-31"), tHead("At our bid\n$11.25mm"), tHead("At the ask\n$14.2mm")],
  tRow("Price per unit", "$95,339", "$120,339"),
  tRow("Going-in cap rate (our underwriting)", "7.57%", "6.00%"),
  tRow("All-in basis per unit", "$121,369", "$146,369"),
  tRow("Sponsor equity", "$5.37mm", "$7.04mm"),
  tRow("Year-1 / Year-5 NOI", "$852K / $1,137K", "$852K / $1,137K"),
  tRow("Year-5 yield on cost vs 6.75% exit", "+119 bps", "(17) bps"),
  tRow("Average cash-on-cash", "8.2%", "5.2%"),
  tRow("Unlevered IRR / multiple", "9.5% / 1.51x", "5.1% / 1.25x"),
  tRow("Levered IRR", "14.2%", "3.1%", true),
  tRow("Equity multiple", "1.83x", "1.15x", true),
], {
  x: 6.95, y: 3.84, w: 5.88, colW: [3.02, 1.46, 1.40],
  rowH: 0.235, fontFace: BODY, border: { type: "solid", color: "2C4A42", pt: 0.5 },
  fill: { color: DARK }, valign: "middle", margin: [2, 5, 2, 5],
});

s1.addText("Sources: rent roll 07/01/2026 · T-12 Feb-25 to Jan-26 · CBRE OM 09/10/2025 · Yardi Matrix Indianapolis 1Q-2026. No price was provided; $14.2mm is implied by the OM's 6.9% Year-1 cap rate on CBRE's $980,838 forecast NOI. Debt: 65% LTC / 8.0% min debt yield / 5.85% interest-only.",
  { x: 0.5, y: 7.12, w: 12.33, h: 0.24, fontFace: BODY, fontSize: 7.5, color: "8C9A92", margin: 0 });

s1.addNotes("The base case is deliberately conservative on what we do not control — rent growth is set off actual 1Q-26 metro performance rather than a forecast, and the exit cap is held at 6.75%. It is confident on what we do control: the renovation premium, the ancillary income and expense management. To justify $14.2mm you would need roughly 4.0% annual rent growth AND a 6.25% exit cap simultaneously. Metro actual is 1.1%.");

// =========================================================================
// SLIDE 2 — Opportunities, risks, assumptions, assessment
// =========================================================================
const s2 = pres.addSlide();
s2.background = { color: WHITE };

s2.addText("Opportunities, Risks & Assumptions", {
  x: 0.5, y: 0.26, w: 8.4, h: 0.44, fontFace: HEAD, fontSize: 26, bold: true, color: INK, margin: 0,
});
s2.addText("Chateau in the Woods  ·  118 Units  ·  Indianapolis, IN  ·  Investment Committee", {
  x: 8.9, y: 0.36, w: 3.93, h: 0.26, fontFace: BODY, fontSize: 10, color: MUTED,
  align: "right", margin: 0,
});

// ---- column helper -------------------------------------------------------
function column(x, w, heading, accent, items, y0) {
  s2.addShape(pres.ShapeType.ellipse, { x, y: y0, w: 0.26, h: 0.26, fill: { color: accent } });
  s2.addText(heading, { x: x + 0.36, y: y0 - 0.03, w: w - 0.36, h: 0.3, fontFace: HEAD,
    fontSize: 14.5, bold: true, color: INK, margin: 0 });
  s2.addText(items.map((it, i) => ([
    { text: it[0] + "  ", options: { bold: true, color: accent } },
    { text: it[1], options: { color: "3D4A43", breakLine: i !== items.length - 1 } },
  ])).flat(), {
    x, y: y0 + 0.34, w, h: 4.10, fontFace: BODY, fontSize: 9, lineSpacing: 11,
    paraSpaceAfter: 6, margin: 0, valign: "top",
  });
}

const OPPS = [
  ["In-unit laundry, all 118 units.", "65 units lack it. At the case's $5,000/unit it earns an incremental $75/month — an 18% return on cost, the best item in the plan. The rent roll shows only 53 units with machines, not the OM's 78."],
  ["Comprehensive interior renovation.", "$13,000/unit for cabinets, counters, LVP, appliances, fixtures and paint at a $125/month premium. Units were last touched in 2006."],
  ["Rent headroom, priced conservatively.", "Renovated rents average $1.29/SF against a $1,310 metro average asking rent — Chateau's units are far larger than market, so we underwrite below the metro average per foot. Avalon Lake (1974, nearby) gets $1,655 on 1,090 SF; our Verdun is $1,450."],
  ["Under-collected ancillary income.", "The $49 technology fee is only ~60% penetrated, the $25 valet trash charge is not yet billed although the property starts paying for it in Jan-26, and there is no pet rent — ~$110K a year of run-rate."],
  ["Operational slack.", "2.2% loss to lease and 6.8% trailing physical vacancy against a 5.5% stabilised target."],
  ["Upside not underwritten.", "Two units from the Bldg 4042 storage conversion and 9 carport-to-garage conversions — real, but 5-10% returns."],
];

const RISKS = [
  ["Basis, not the plan.", "At the ask we pay a 6.00% going-in cap, fund $26,030/unit of capital, and reach a Year-5 yield on cost of 6.58% — below the 6.75% exit. The capital plan nets minus $6,880."],
  ["Rent growth and new supply.", "Metro rents flat on a trailing-three-month basis, +1.1% year over year in 1Q-26, and 2025 starts more than doubled — delivering 2027-28, inside our hold, with Carmel taking over half of 2026 deliveries just north of us. Each 100 bps of rent growth is worth ~3 points of IRR."],
  ["Tax reassessment.", "Assessed value $5.55mm against a $14.2mm price; the OM's comp study shows AV averaging 69% of sale price post-trade. We step taxes $140K to $210K — a move to 69% adds ~$110K a year, ~$1.6mm of value."],
  ["Unproven rent premium.", "Nothing has been renovated to a current standard. Every $25/month of premium is worth ~1.3 points of IRR. A leased mock-up before we go hard would settle it."],
  ["1974 physical risk and exit.", "3.5 of 5 original rubber roofs, 37 below-grade units needing drains and sump pumps, original sliders, central gas water heaters. Plus 4 evictions and 8 notices on the rent roll, and 75 bps of exit cap expansion costs ~5 points of IRR."],
];

column(0.5, 4.05, "Key Opportunities", FOREST, OPPS, 0.92);
column(4.85, 4.05, "Key Risks", RED, RISKS, 0.92);

// ---- assumptions ---------------------------------------------------------
s2.addShape(pres.ShapeType.ellipse, { x: 9.2, y: 0.92, w: 0.26, h: 0.26, fill: { color: AMBER } });
s2.addText("Major Assumptions", { x: 9.56, y: 0.89, w: 3.3, h: 0.3, fontFace: HEAD,
  fontSize: 14.5, bold: true, color: INK, margin: 0 });

const aHead = (t) => ({ text: t, options: { bold: true, fontSize: 8.5, color: WHITE,
  fill: { color: INK }, valign: "middle" } });
const aRow = (a, b) => ([
  { text: a, options: { fontSize: 8.5, color: "3D4A43" } },
  { text: b, options: { fontSize: 8.5, bold: true, color: INK, align: "right" } },
]);
s2.addTable([
  [aHead("Driver"), aHead("Underwritten")],
  aRow("Purchase price (not provided; from OM cap rate)", "$14.2mm"),
  aRow("Market rent growth Y1 to Y5", "1.5% to 3.0%"),
  aRow("Reno start / pace / completion", "mo 4 / 5 / mo 27"),
  aRow("Interior cost / rent premium", "$13,000 / $125mo"),
  aRow("W/D addition, 65 units (case-set)", "$5,000 / $75mo"),
  aRow("Non-unit capital + 8% contingency", "$1.21mm"),
  aRow("Stabilised vacancy / reno downtime", "5.5% / 14 days"),
  aRow("Bad debt, Year 1 to Year 5", "1.50% to 1.00%"),
  aRow("Real estate taxes, Year 1 to Year 5", "$140K to $210K"),
  aRow("Operating expenses, Year 1", "$8,948/unit"),
  aRow("Debt: LTC / rate / structure", "65% / 5.85% / IO"),
  aRow("Exit cap / cost of sale", "6.75% / 1.5%"),
], {
  x: 9.2, y: 1.26, w: 3.63, colW: [2.22, 1.41], rowH: 0.205, fontFace: BODY,
  border: { type: "solid", color: LINE, pt: 0.5 }, valign: "middle", margin: [2, 4, 2, 4],
});
s2.addText("Rent growth is set off Yardi Matrix 1Q-26 ACTUALS, not a forecast. Where the rent roll and T-12 conflict with the OM, the Excel files govern — including the washer/dryer count (53, not the OM's 78).",
  { x: 9.2, y: 4.60, w: 3.63, h: 0.58, fontFace: BODY, fontSize: 8, italic: true, color: MUTED,
    valign: "top", margin: 0 });

// ---- chart ---------------------------------------------------------------
s2.addText("Levered IRR by purchase price (6.75% exit cap)", {
  x: 0.5, y: 5.30, w: 5.1, h: 0.26, fontFace: HEAD, fontSize: 11.5, bold: true, color: INK, margin: 0,
});
s2.addChart(pres.ChartType.bar, [{
  name: "Levered IRR",
  labels: ["$10.5mm", "$11.0mm", "$11.25mm", "$11.75mm", "$12.5mm", "$13.25mm", "$14.2mm"],
  values: [0.172, 0.152, 0.142, 0.123, 0.094, 0.065, 0.031],
}], {
  x: 0.42, y: 5.56, w: 5.25, h: 1.42,
  barDir: "col", chartColors: [FOREST, FOREST, AMBER, MOSS, MOSS, RED, RED],
  showLegend: false, showTitle: false,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%",
  dataLabelFontSize: 7.5, dataLabelColor: INK, dataLabelFontFace: BODY,
  catAxisLabelFontSize: 7.5, catAxisLabelColor: MUTED, catAxisLabelFontFace: BODY,
  catGridLine: { style: "none" },
  valAxisHidden: true, valGridLine: { style: "none" },
  valAxisMaxVal: 0.20, valAxisMinVal: 0,
  barGapWidthPct: 45, plotArea: { fill: { color: WHITE } },
});
s2.addText("14% hurdle is cleared only at or below $11.25mm — a 21% discount to the ask.", {
  x: 0.5, y: 6.99, w: 5.1, h: 0.22, fontFace: BODY, fontSize: 8, italic: true, color: MUTED, margin: 0,
});

// ---- overall assessment --------------------------------------------------
s2.addShape(pres.ShapeType.roundRect, {
  x: 6.0, y: 5.22, w: 6.83, h: 1.90, fill: { color: DARK }, rectRadius: 0.05, shadow: sh(),
});
s2.addText("Overall Assessment", { x: 6.3, y: 5.34, w: 6.2, h: 0.26, fontFace: HEAD,
  fontSize: 12.5, bold: true, color: MOSS, margin: 0 });
s2.addText("A well-located asset with a genuinely accretive renovation — 12.7% on cost, 18% on the laundry component. Approve the plan and the $3.07mm of capital; do not approve the price. To make $14.2mm work you would need roughly 4.0% annual rent growth AND a 6.25% exit cap simultaneously. Metro rents actually grew 1.1% last year and 2025 starts doubled. Underwriting a market recovery we do not control, to justify a price we do control, is how an acquisitions team loses money it cannot renovate its way out of. Bid $11.25mm and be willing to lose.",
  { x: 6.3, y: 5.64, w: 6.25, h: 1.40, fontFace: BODY, fontSize: 9.5, color: "DFE6E0",
    lineSpacing: 12.5, valign: "top", margin: 0 });

s2.addNotes("Expect to lose this deal at $11.25mm; that is the right outcome. If the seller holds at $14.2mm we pass. Two structures worth testing before walking: a tax escrow or price adjustment tied to the first post-closing Form 11A, and a seller credit for the identified immediate needs (sump pumps, patio drains, asphalt — about $190K).");

pres.writeFile({ fileName: "Chateau_in_the_Woods_IC_Presentation.pptx" })
  .then(f => console.log("wrote", f));
