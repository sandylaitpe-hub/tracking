const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";               // 13.333 x 7.5
pres.author = "Investment Committee Memorandum";
pres.title = "Chateau in the Woods — Investment Committee Recommendation";

// ---- palette -------------------------------------------------------------
// Green house palette. fairlawn.com could not be reached from this environment
// (network policy returned 403 on CONNECT), so the exact brand hexes are not
// sampled from the site — this is a documented stand-in in the same family.
const DEEP   = "0D3B2E";   // deep forest — dark slide background
const CARD   = "164C3A";   // card on dark
const GREEN  = "1F7A5C";   // primary green
const GREEN2 = "2E9E76";   // bright green
const SAGE   = "8FC9A9";   // light green accent (on dark)
const SAGE2  = "BFE0CD";   // pale green (chart fills)
const PALE   = "EAF3EE";   // tint on light slides
const GOLD   = "C9A227";   // accent / hurdle
const RED    = "9E2B25";   // downside
const INK    = "13201A";
const MUTED  = "5F7269";
const LINE   = "D4E0D8";
const WHITE  = "FFFFFF";

const HEAD = "Cambria";
const BODY = "Calibri";

const sh = () => ({ type: "outer", color: "081E16", blur: 8, offset: 2, angle: 90, opacity: 0.18 });

// chart defaults shared by every plot
const CH = {
  showLegend: false, showTitle: false,
  catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
  dataLabelFontFace: BODY,
  catGridLine: { style: "none" },
  catAxisMajorTickMark: "none", catAxisMinorTickMark: "none",
  valAxisMajorTickMark: "none", valAxisMinorTickMark: "none",
  valAxisLineShow: false,
};

// small helper: section heading on a light slide
function h2(slide, x, y, w, text, sub) {
  slide.addShape(pres.ShapeType.rect, { x, y: y + 0.03, w: 0.055, h: 0.20, fill: { color: GREEN2 } });
  slide.addText(text, { x: x + 0.14, y: y - 0.03, w: w - 0.14, h: 0.26, fontFace: HEAD,
    fontSize: 11.5, bold: true, color: INK, margin: 0 });
  if (sub) slide.addText(sub, { x: x + 0.14, y: y + 0.20, w: w - 0.14, h: 0.20, fontFace: BODY,
    fontSize: 8, color: MUTED, margin: 0 });
}

// =========================================================================
// SLIDE 1 — THE RECOMMENDATION
// =========================================================================
const s1 = pres.addSlide();
s1.background = { color: DEEP };
s1.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.10, fill: { color: GREEN2 } });

s1.addText("Chateau in the Woods", {
  x: 0.45, y: 0.28, w: 8.6, h: 0.50, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, margin: 0,
});
s1.addText("118 Units  ·  4020 Monaco Drive, Indianapolis, IN 46220  ·  Built 1974  ·  134,371 SF  ·  96.6% occupied",
  { x: 0.47, y: 0.78, w: 9.2, h: 0.26, fontFace: BODY, fontSize: 11, color: SAGE, margin: 0 });
s1.addText([
  { text: "INVESTMENT COMMITTEE", options: { fontSize: 11, bold: true, color: WHITE, breakLine: true } },
  { text: "Recommendation  ·  Value-Add Acquisition", options: { fontSize: 9.5, color: SAGE } },
], { x: 9.9, y: 0.32, w: 2.98, h: 0.6, align: "right", fontFace: BODY, margin: 0 });

// ---- verdict bar ---------------------------------------------------------
s1.addShape(pres.ShapeType.roundRect, {
  x: 0.45, y: 1.15, w: 12.43, h: 0.74, fill: { color: GOLD }, rectRadius: 0.06, shadow: sh(),
});
s1.addText([
  { text: "BID $12,000,000.  ", options: { bold: true } },
  { text: "Price is the only variable we control." },
], { x: 0.75, y: 1.15, w: 7.2, h: 0.74, fontFace: HEAD, fontSize: 15.5, color: "2A2004",
     valign: "middle", margin: 0 });
s1.addText("It traded at $13,625,000 — 12% above us.", {
  x: 7.95, y: 1.15, w: 4.70, h: 0.74, fontFace: BODY, fontSize: 12, bold: true, color: "2A2004",
  align: "right", valign: "middle", margin: 0,
});

// ---- KPI cards -----------------------------------------------------------
const kpis = [
  ["14.0% / 1.81x", "Levered IRR / multiple at our bid", SAGE],
  ["+111 bps", "Year-5 yield on cost over the 6.75% exit", SAGE],
  ["12.9%", "Return on renovation cost, $14,407/unit", SAGE],
  ["16 months", "All 118 units renovated, nobody displaced", SAGE],
  ["7.9% / 1.42x", "Same plan at the $13.625mm trade price", "F3A6A0"],
];
kpis.forEach(([big, lbl, col], i) => {
  const x = 0.45 + i * 2.49;
  s1.addShape(pres.ShapeType.roundRect, {
    x, y: 2.05, w: 2.33, h: 0.98, fill: { color: CARD }, rectRadius: 0.05,
  });
  s1.addText(big, { x: x + 0.16, y: 2.13, w: 2.05, h: 0.40, fontFace: HEAD, fontSize: 19,
    bold: true, color: col, margin: 0 });
  s1.addText(lbl, { x: x + 0.16, y: 2.53, w: 2.05, h: 0.42, fontFace: BODY, fontSize: 8.5,
    color: "C9D7CE", lineSpacing: 10, valign: "top", margin: 0 });
});

// ---- hero chart: price ladder -------------------------------------------
s1.addText("Levered IRR by purchase price", { x: 0.45, y: 3.24, w: 7.4, h: 0.26,
  fontFace: HEAD, fontSize: 13, bold: true, color: WHITE, margin: 0 });
s1.addText("Five-year hold · 6.75% exit cap held constant · every bar is a live cash flow, not an interpolation",
  { x: 0.45, y: 3.50, w: 7.4, h: 0.20, fontFace: BODY, fontSize: 8, color: SAGE, margin: 0 });

const PL = ["$10.5mm", "$11.0mm", "$11.5mm", "$12.0mm", "$12.5mm", "$13.0mm", "$13.625mm", "$14.2mm"];
s1.addChart([
  { type: pres.ChartType.bar,
    data: [{ name: "Levered IRR", labels: PL,
             values: [0.199, 0.179, 0.159, 0.140, 0.121, 0.102, 0.079, 0.058] }],
    options: { chartColors: [GREEN2, GREEN2, GREEN2, GOLD, "5A7F70", "5A7F70", RED, RED],
               barGapWidthPct: 40, showValue: true } },
  { type: pres.ChartType.line,
    data: [{ name: "14% hurdle", labels: PL,
             values: [0.14, 0.14, 0.14, 0.14, 0.14, 0.14, 0.14, 0.14] }],
    options: { chartColors: [WHITE], lineSize: 1.25, lineDash: "dash", lineDataSymbol: "none",
               showValue: false } },
], Object.assign({}, CH, {
  x: 0.30, y: 3.72, w: 7.55, h: 3.20,
  dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%",
  dataLabelFontSize: 9, dataLabelColor: WHITE,
  catAxisLabelFontSize: 9, catAxisLabelColor: SAGE,
  valAxisHidden: true, valGridLine: { style: "none" },
  valAxisMaxVal: 0.225, valAxisMinVal: 0,
  plotArea: { fill: { color: DEEP } }, chartArea: { fill: { color: DEEP } },
}));
s1.addText("Dashed line = 14% hurdle.  Cleared at or below $12.0mm.  Red = the actual trade and above.", {
  x: 0.45, y: 6.90, w: 7.4, h: 0.20, fontFace: BODY, fontSize: 8, italic: true, color: SAGE, margin: 0,
});

// ---- returns table -------------------------------------------------------
s1.addText("Our bid vs. the trade", { x: 8.15, y: 3.24, w: 4.73, h: 0.26,
  fontFace: HEAD, fontSize: 13, bold: true, color: WHITE, margin: 0 });
s1.addText("Identical business plan, identical NOI — only the basis changes",
  { x: 8.15, y: 3.50, w: 4.73, h: 0.20, fontFace: BODY, fontSize: 8, color: SAGE, margin: 0 });

const tH = (t) => ({ text: t, options: { bold: true, color: WHITE, fontSize: 9.5,
  fill: { color: GREEN }, align: "center", valign: "middle" } });
const tR = (a, b, c, hi) => ([
  { text: a, options: { color: hi ? WHITE : "D5E2DA", bold: !!hi, fontSize: 9.5 } },
  { text: b, options: { color: hi ? SAGE : "D5E2DA", bold: !!hi, align: "center", fontSize: 9.5 } },
  { text: c, options: { color: hi ? "F3A6A0" : "AFC0B7", bold: !!hi, align: "center", fontSize: 9.5 } },
]);
s1.addTable([
  [tH("Aug-26 to Jul-31"), tH("Our bid\n$12.0mm"), tH("Trade\n$13.625mm")],
  tR("Price per unit", "$101,695", "$115,466"),
  tR("Going-in cap rate", "7.28%", "6.41%"),
  tR("All-in basis per unit", "$126,269", "$140,041"),
  tR("Sponsor equity", "$5.59mm", "$6.20mm"),
  tR("Year-1 NOI", "$873K", "$873K"),
  tR("Year-5 NOI", "$1,171K", "$1,171K"),
  tR("Year-5 yield on cost", "7.86%", "7.09%"),
  tR("Spread over 6.75% exit", "+111 bps", "+34 bps"),
  tR("Average cash-on-cash", "8.4%", "6.5%"),
  tR("Unlevered IRR", "9.4%", "6.9%"),
  tR("Levered IRR", "14.0%", "7.9%", true),
  tR("Equity multiple", "1.81x", "1.42x", true),
], {
  x: 8.15, y: 3.76, w: 4.73, colW: [2.31, 1.24, 1.18],
  rowH: 0.238, fontFace: BODY, border: { type: "solid", color: "2A5A48", pt: 0.5 },
  fill: { color: DEEP }, valign: "middle", margin: [2, 5, 2, 5],
});

s1.addText("Rent roll 07/01/2026 · T-12 Feb-25 to Jan-26 · CBRE OM 09/10/2025 · Yardi Matrix Indianapolis 1Q-2026 · Marion County parcel 8000002. Rent growth set off 1Q-26 ACTUALS, not a forecast. Debt 65% LTC / 8.0% min debt yield / 1.25x DSCR / 5.85% interest-only.",
  { x: 0.45, y: 7.16, w: 12.43, h: 0.22, fontFace: BODY, fontSize: 7.5, color: "7C8F85", margin: 0 });

s1.addNotes("The model was not handed a price — it solves for one. $12.0mm is the highest basis at which this plan clears 14% and 1.80x on rent growth taken from 1Q-26 actuals. It traded 12% above that on 12 June 2026. Conservative on what we do not control (rent growth, exit cap, taxes evidenced against 25 years of assessment record); confident on what we do (renovation premium, pace, ancillary income).");
// =========================================================================
// SLIDE 2 — WHERE THE VALUE COMES FROM, AND WHAT MOVES IT
// =========================================================================
const s2 = pres.addSlide();
s2.background = { color: WHITE };
s2.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.10, fill: { color: GREEN2 } });

s2.addText("Where the value comes from — and what moves it", {
  x: 0.45, y: 0.20, w: 9.6, h: 0.40, fontFace: HEAD, fontSize: 21, bold: true, color: INK, margin: 0,
});
s2.addText("Chateau in the Woods  ·  118 Units  ·  Indianapolis, IN", {
  x: 9.4, y: 0.30, w: 3.48, h: 0.26, fontFace: BODY, fontSize: 10, color: MUTED, align: "right", margin: 0,
});

const CW = 4.02, CX = [0.45, 4.66, 8.87];

// ---- (a) NOI by year -----------------------------------------------------
h2(s2, CX[0], 0.76, CW, "NOI grows 34% in five years", "$873K in Year 1 → $1,171K in Year 5");
s2.addChart(pres.ChartType.bar, [{
  name: "NOI", labels: ["Y1", "Y2", "Y3", "Y4", "Y5"],
  values: [873, 1060, 1108, 1142, 1171],
}], Object.assign({}, CH, {
  x: CX[0] - 0.18, y: 1.20, w: CW + 0.28, h: 1.62,
  barDir: "col", chartColors: [GREEN, GREEN2, GREEN2, GREEN2, GREEN2], barGapWidthPct: 45,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '$#,##0"K"',
  dataLabelFontSize: 8, dataLabelColor: INK,
  catAxisLabelFontSize: 8.5, catAxisLabelColor: MUTED,
  valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 1350, valAxisMinVal: 0,
  plotArea: { fill: { color: WHITE } }, chartArea: { fill: { color: WHITE } },
}));

// ---- (b) yield on cost vs exit cap ---------------------------------------
h2(s2, CX[1], 0.76, CW, "Yield on cost clears the exit cap", "The value-add test: 7.86% vs 6.75% = +111 bps");
s2.addChart([
  { type: pres.ChartType.line,
    data: [{ name: "Yield on cost", labels: ["Y1", "Y2", "Y3", "Y4", "Y5"],
             values: [0.0586, 0.0711, 0.0744, 0.0767, 0.0786] }],
    options: { chartColors: [GREEN], lineSize: 2.75, lineDataSymbol: "circle", lineDataSymbolSize: 6 } },
  { type: pres.ChartType.line,
    data: [{ name: "Exit cap 6.75%", labels: ["Y1", "Y2", "Y3", "Y4", "Y5"],
             values: [0.0675, 0.0675, 0.0675, 0.0675, 0.0675] }],
    options: { chartColors: [GOLD], lineSize: 1.5, lineDash: "dash", lineDataSymbol: "none" } },
], Object.assign({}, CH, {
  x: CX[1] - 0.18, y: 1.20, w: CW + 0.28, h: 1.62,
  showValue: false,
  catAxisLabelFontSize: 8.5, catAxisLabelColor: MUTED,
  valAxisLabelFontSize: 8, valAxisLabelColor: MUTED, valAxisLabelFormatCode: "0.0%",
  valAxisMaxVal: 0.085, valAxisMinVal: 0.05, valAxisMajorUnit: 0.01,
  valGridLine: { style: "solid", color: "EAEFEB", size: 0.5 },
  plotArea: { fill: { color: WHITE } }, chartArea: { fill: { color: WHITE } },
}));
s2.addText("Solid = yield on cost   ·   dashed = 6.75% exit cap", {
  x: CX[1], y: 2.84, w: CW, h: 0.18, fontFace: BODY, fontSize: 7.5, italic: true, color: MUTED, margin: 0 });

// ---- (c) renovation schedule --------------------------------------------
h2(s2, CX[2], 0.76, CW, "Two tracks, done in month 16", "Cumulative units renovated, by month");
s2.addChart(pres.ChartType.bar, [
  { name: "Track A — in place", labels: Array.from({ length: 18 }, (_, i) => String(i + 1)),
    values: [0, 10, 20, 30, 40, 50, 53, 53, 53, 53, 53, 53, 53, 53, 53, 53, 53, 53] },
  { name: "Track B — vacant + laundry", labels: Array.from({ length: 18 }, (_, i) => String(i + 1)),
    values: [0, 0, 0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 65, 65] },
], Object.assign({}, CH, {
  x: CX[2] - 0.18, y: 1.20, w: CW + 0.28, h: 1.62,
  barDir: "col", barGrouping: "stacked", chartColors: [GREEN, SAGE2], barGapWidthPct: 20,
  showValue: false,
  catAxisLabelFontSize: 7, catAxisLabelColor: MUTED,
  valAxisLabelFontSize: 7.5, valAxisLabelColor: MUTED,
  valAxisMaxVal: 118, valAxisMinVal: 0, valAxisMajorUnit: 59,
  valGridLine: { style: "solid", color: "EAEFEB", size: 0.5 },
  plotArea: { fill: { color: WHITE } }, chartArea: { fill: { color: WHITE } },
}));
s2.addText("Dark = 53 units renovated in place (no permit, no vacancy loss)  ·  pale = 65 units taken vacant to add laundry", {
  x: CX[2], y: 2.84, w: CW, h: 0.30, fontFace: BODY, fontSize: 7.5, italic: true, color: MUTED,
  lineSpacing: 9, margin: 0 });

// ---- tornado -------------------------------------------------------------
h2(s2, 0.45, 3.24, 7.4, "What moves the 14.0% — breakeven ladders",
   "Levered IRR at our $12.0mm bid. Each bar spans the full range tested; left number is the downside, right the upside.");

// Floating bars: the upside series is drawn first, then the downside series is
// overlaid 100% in white so only the span between them stays coloured.
const TOR = [
  ["Renovation cost $20.0K → $11.0K",     0.1135, 0.1566],
  ["Renovation premium $100 → $190 /mo",  0.1020, 0.1623],
  ["Year-5 tax bill $294K → $137K",       0.0738, 0.1504],
  ["Exit cap 7.50% → 6.25%",              0.0950, 0.1715],
  ["Market rent growth 1.0% → 3.5% flat", 0.0714, 0.1873],
];
s2.addChart(pres.ChartType.bar, [
  { name: "upside", labels: TOR.map(t => t[0]), values: TOR.map(t => t[2]) },
  { name: "downside", labels: TOR.map(t => t[0]), values: TOR.map(t => t[1]) },
], Object.assign({}, CH, {
  x: 0.20, y: 3.76, w: 7.75, h: 2.20,
  barDir: "bar", barGrouping: "clustered", barOverlapPct: 100,
  chartColors: [SAGE2, WHITE], barGapWidthPct: 55,
  showValue: true, dataLabelPosition: "inEnd", dataLabelFormatCode: "0.0%",
  dataLabelFontSize: 8, dataLabelColor: INK,
  catAxisLabelFontSize: 8.5, catAxisLabelColor: INK, catAxisLineShow: false,
  valAxisHidden: true, valGridLine: { style: "none" },
  valAxisMinVal: 0.04, valAxisMaxVal: 0.20,
  plotArea: { fill: { color: WHITE } }, chartArea: { fill: { color: WHITE } },
}));

// ---- major assumptions ---------------------------------------------------
h2(s2, 8.15, 3.24, 4.73, "Major assumptions",
   "Set off Yardi 1Q-26 actuals and the 25-year county tax record — not forecasts");
const aH = (t) => ({ text: t, options: { bold: true, fontSize: 8.5, color: WHITE,
  fill: { color: INK }, valign: "middle" } });
const aR = (a, b) => ([
  { text: a, options: { fontSize: 8.5, color: "3A4A42" } },
  { text: b, options: { fontSize: 8.5, bold: true, color: INK, align: "right" } },
]);
s2.addTable([
  [aH("Driver"), aH("Underwritten")],
  aR("Purchase price — SOLVED, not given", "$12,000,000"),
  aR("Market rent growth, Y1 to Y5", "1.5 / 2.0 / 2.5 / 3.0 / 3.0%"),
  aR("Stabilised vacancy / Year-1 bad debt", "5.5% / 1.50%"),
  aR("Track A — start / pace / cost / prem", "mo 2 / 10 / $10,000 / $100"),
  aR("Track B — start / pace / cost / prem", "mo 4 / 5 / $18,000 / $200"),
  aR("Blended cost / premium / return on cost", "$14,407 / $155 / 12.9%"),
  aR("Non-unit capital + 8% contingency", "$1,199,800"),
  aR("Operating expenses, Year 1", "$8,929 per unit"),
  aR("Real estate taxes, Y1 to Y5", "$137,000 to $160,300"),
  aR("Debt — LTC / rate / structure", "65% / 5.85% / interest-only"),
  aR("Exit cap / cost of sale", "6.75% / 1.5%"),
  aR("Data hierarchy — the Excel files govern", "53 with laundry, not 78"),
], {
  x: 8.15, y: 3.76, w: 4.73, colW: [2.73, 2.00], rowH: 0.183, fontFace: BODY,
  border: { type: "solid", color: LINE, pt: 0.5 }, valign: "middle", margin: [1, 5, 1, 5],
});

// ---- opportunities / risks / assessment ---------------------------------
function mini(x, w, title, accent, items) {
  s2.addShape(pres.ShapeType.rect, { x, y: 6.26, w, h: 0.035, fill: { color: accent } });
  s2.addText(title, { x, y: 6.34, w, h: 0.20, fontFace: BODY, fontSize: 9, bold: true,
    color: accent, margin: 0 });
  s2.addText(items.map((it, i) => ([
    { text: it[0] + " ", options: { bold: true, color: INK } },
    { text: it[1], options: { color: "4A5A52", breakLine: i !== items.length - 1 } },
  ])).flat(), { x, y: 6.56, w, h: 0.86, fontFace: BODY, fontSize: 8, lineSpacing: 9.8,
    paraSpaceAfter: 2, valign: "top", margin: 0 });
}
mini(0.45, 3.90, "KEY OPPORTUNITIES", GREEN, [
  ["12.9% return on renovation cost.", "$14,407 buys $155 a month."],
  ["All 118 units in 16 months.", "Two tracks, nobody displaced, no permit on 53."],
  ["~$110K a year uncollected.", "Tech fee 60% penetrated, valet trash unbilled, no pet rent."],
]);
mini(4.60, 3.90, "KEY RISKS", RED, [
  ["Rent growth — the dominant variable.", "At a flat 1.0%, the IRR is 7.1%."],
  ["Basis.", "It traded 12% above our bid; there the plan returns 7.9%."],
  ["Unproven premium, 1974 plant.", "Nothing renovated to standard; 3.5 of 5 roofs original."],
]);
s2.addShape(pres.ShapeType.roundRect, {
  x: 8.75, y: 6.26, w: 4.13, h: 1.06, fill: { color: DEEP }, rectRadius: 0.05, shadow: sh(),
});
s2.addText("OVERALL ASSESSMENT", { x: 8.95, y: 6.36, w: 3.8, h: 0.20, fontFace: BODY, fontSize: 8.5,
  bold: true, color: SAGE, margin: 0 });
s2.addText("A genuinely accretive plan on a well-located asset — but only at a price. We bid $12.0mm and were outbid by 12%. On 1.1% actual metro rent growth, that is the right outcome.",
  { x: 8.95, y: 6.58, w: 3.78, h: 0.66, fontFace: BODY, fontSize: 8.5, color: "DDE8E1",
    lineSpacing: 10.5, valign: "top", margin: 0 });

s2.addNotes("Read the tornado from the bottom: the only bar that breaks the deal is the one we do not control. Rent growth at a flat 1.0% takes it to 7.1%; at 3.5% it is 18.7%. Everything we do control — premium, cost, pace — moves it roughly 5 points either way, which is why the bid, not the business plan, is the decision in front of the committee.");

// =========================================================================
// APPENDIX — SUPPORTING DETAIL
// =========================================================================
const a1 = pres.addSlide();
a1.background = { color: WHITE };
a1.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.10, fill: { color: GREEN2 } });
a1.addText("Appendix  ·  Business plan, price sensitivity and the evidence", { x: 0.45, y: 0.22, w: 9.6, h: 0.40,
  fontFace: HEAD, fontSize: 20, bold: true, color: INK, margin: 0 });
a1.addText("Full detail in the 13-tab workbook", { x: 9.4, y: 0.32, w: 3.48, h: 0.26,
  fontFace: BODY, fontSize: 10, color: MUTED, align: "right", margin: 0 });

// ---- two-track table -----------------------------------------------------
h2(a1, 0.45, 0.86, 6.0, "Two tracks, chosen by what each unit needs");
const trH = (t) => ({ text: t, options: { bold: true, fontSize: 9, color: WHITE,
  fill: { color: GREEN }, valign: "middle" } });
const trR = (a, b, c) => ([
  { text: a, options: { fontSize: 8.5, color: MUTED } },
  { text: b, options: { fontSize: 8.5, color: INK, bold: true } },
  { text: c, options: { fontSize: 8.5, color: INK, bold: true } },
]);
a1.addTable([
  [trH(""), trH("Track A — in place"), trH("Track B — vacant unit")],
  trR("Units", "53 — laundry already in unit", "65 — laundry to be added"),
  trR("Resident impact", "None — resident stays", "Natural turnover only"),
  trR("Permit", "None — cosmetic scope exempt", "State CDR, then county DBNS"),
  trR("Start / pace", "Month 2 / 10 per month", "Month 4 / 5 per month"),
  trR("Cost per unit", "$10,000", "$13,000 + $5,000 laundry"),
  trR("Rent premium", "$100 / month", "$200 / month"),
  trR("Premium captured", "At renewal — 6-month lag", "Immediately, on the new lease"),
  trR("Downtime", "None", "14 incremental days"),
  trR("Complete", "Month 7", "Month 16"),
], {
  x: 0.45, y: 1.26, w: 6.0, colW: [1.55, 2.25, 2.20], rowH: 0.215, fontFace: BODY,
  border: { type: "solid", color: LINE, pt: 0.5 }, valign: "middle", margin: [2, 5, 2, 5],
});
a1.addText([
  { text: "Blended $14,407 per unit buys a $155 per month premium — a 12.9% return on cost", options: { bold: true, color: INK } },
  { text: ", against a 6.75% exit. A single-track plan at 5 units a month would take 27 months; running the 53 in place in parallel finishes in 16 without forcing a single resident out.", options: { color: MUTED } },
], { x: 0.45, y: 3.66, w: 6.0, h: 0.62, fontFace: BODY, fontSize: 9, lineSpacing: 11.5, margin: 0 });

// ---- capital budget ------------------------------------------------------
h2(a1, 0.45, 4.36, 6.0, "Where the $2.90mm of capital goes", "$24,575 per unit, all funded at closing in Sources & Uses");
a1.addChart(pres.ChartType.bar, [{
  name: "Capital", labels: ["Unit renovation\n118 units", "Non-unit capital\nroof, drains, amenity", "Contingency\n8%"],
  values: [1700, 985, 215],
}], Object.assign({}, CH, {
  x: 0.30, y: 4.90, w: 6.25, h: 2.10,
  barDir: "col", chartColors: [GREEN, GREEN2, SAGE2], barGapWidthPct: 55,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '$#,##0"K"',
  dataLabelFontSize: 9, dataLabelColor: INK,
  catAxisLabelFontSize: 8, catAxisLabelColor: MUTED,
  valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 2100, valAxisMinVal: 0,
  plotArea: { fill: { color: WHITE } }, chartArea: { fill: { color: WHITE } },
}));
a1.addText("Average levered cash-on-cash 8.4% — 4.9% in Year 1 while the renovation runs, 10.1% by Year 5.",
  { x: 0.45, y: 7.02, w: 6.0, h: 0.24, fontFace: BODY, fontSize: 8.5, italic: true, color: MUTED, margin: 0 });

// ---- price x exit cap grid ----------------------------------------------
h2(a1, 6.85, 0.86, 6.03, "Levered IRR — purchase price × exit cap",
   "The hurdle moves with the exit: at a 7.25% exit we can pay $11.25mm; at 6.25%, $12.87mm.");
const gH = (t) => ({ text: t, options: { bold: true, fontSize: 9, color: WHITE,
  fill: { color: GREEN }, align: "center", valign: "middle" } });
const gCell = (v) => ({
  text: (v * 100).toFixed(1) + "%",
  options: { fontSize: 9, align: "center", bold: v >= 0.14,
    color: v >= 0.14 ? INK : MUTED,
    fill: { color: v >= 0.16 ? SAGE2 : v >= 0.14 ? PALE : WHITE } },
});
const GRID = [
  ["$11.0mm", [0.209, 0.194, 0.179, 0.165, 0.151]],
  ["$11.5mm", [0.190, 0.175, 0.159, 0.145, 0.130]],
  ["$12.0mm", [0.171, 0.156, 0.140, 0.125, 0.110]],
  ["$12.5mm", [0.153, 0.137, 0.121, 0.105, 0.090]],
  ["$13.0mm", [0.135, 0.119, 0.102, 0.086, 0.070]],
  ["$13.625mm", [0.113, 0.096, 0.079, 0.062, 0.045]],
];
a1.addTable([
  [gH("Purchase price"), gH("6.25%"), gH("6.50%"), gH("6.75%"), gH("7.00%"), gH("7.25%")],
  ...GRID.map(([p, vals]) => ([
    { text: p, options: { fontSize: 9, bold: p === "$12.0mm", color: p === "$13.625mm" ? RED : INK } },
    ...vals.map(v => gCell(v)),
  ])),
], {
  x: 6.85, y: 1.40, w: 6.03, colW: [1.43, 0.92, 0.92, 0.92, 0.92, 0.92], rowH: 0.235,
  fontFace: BODY, border: { type: "solid", color: LINE, pt: 0.5 }, valign: "middle",
  margin: [2, 5, 2, 5],
});
a1.addText("Shaded ≥ 14%.  Base case $12.0mm at a 6.75% exit = 14.0%.  $13.625mm is what it traded for.", {
  x: 6.85, y: 3.12, w: 6.03, h: 0.20, fontFace: BODY, fontSize: 8, italic: true, color: MUTED, margin: 0 });

// ---- evidence ------------------------------------------------------------
h2(a1, 6.85, 3.48, 6.03, "Market and public-record evidence", "Actuals, not forecasts — this is what the assumptions are set off");
const eH = (t) => ({ text: t, options: { bold: true, fontSize: 9, color: WHITE,
  fill: { color: INK }, valign: "middle" } });
const eR = (a, b, c) => ([
  { text: a, options: { fontSize: 8.5, color: "3A4A42" } },
  { text: b, options: { fontSize: 8.5, bold: true, color: INK, align: "right" } },
  { text: c, options: { fontSize: 8, color: MUTED } },
]);
a1.addTable([
  [eH("Evidence"), eH("Actual"), eH("Source")],
  eR("Indianapolis metro asking rent, 1Q-26", "$1,310", "Yardi Matrix"),
  eR("Trailing-three-month rent change", "FLAT", "Yardi Matrix"),
  eR("Year-over-year rent change", "+1.1%", "Yardi Matrix"),
  eR("2025 construction starts", "more than doubled", "Yardi Matrix"),
  eR("2025 net annual tax bill, parcel 8000002", "$129,648", "Marion County"),
  eR("Assessed value range, 2012 to 2025", "$5.1 – 6.6mm", "Marion County"),
  eR("Last recorded assessment change", "15 Mar 2022", "Marion County"),
  eR("2025 AV as a share of the 2026 trade price", "41%", "Marion County"),
  eR("Trade price, deed 12 Jun 2026", "$13,625,000", "Marion County"),
  eR("Prior sale, Mar 2006", "$4,400,000", "Marion County"),
], {
  x: 6.85, y: 4.02, w: 6.03, colW: [3.12, 1.55, 1.36], rowH: 0.208, fontFace: BODY,
  border: { type: "solid", color: LINE, pt: 0.5 }, valign: "middle", margin: [2, 5, 2, 5],
});
a1.addText("Indiana values apartments at the lowest of the income, market and cost approaches — which is why assessed value here has not tracked trade prices. An earlier draft stepped taxes toward a share of the purchase price; the record does not support it, and correcting it is worth about 1.5 points of levered IRR. The risk is still priced, in Ladder C.",
  { x: 6.85, y: 6.36, w: 6.03, h: 0.76, fontFace: BODY, fontSize: 8.5, italic: true, color: MUTED,
    lineSpacing: 10.5, valign: "top", margin: 0 });

pres.writeFile({ fileName: "Chateau_in_the_Woods_IC_Presentation.pptx" })
  .then(f => console.log("wrote", f));
