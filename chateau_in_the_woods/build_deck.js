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
  { text: "BID $12.0mm.  ", options: { bold: true } },
  { text: "The plan works; the price is what has to be disciplined." },
], { x: 0.80, y: 1.20, w: 8.0, h: 0.80, fontFace: HEAD, fontSize: 16, color: "2A1B04",
     valign: "middle", margin: 0 });
s1.addText([
  { text: "Bid $12.0mm ($101,695/unit).", options: { bold: true, breakLine: true } },
  { text: "It traded at $13.625mm — 12% above us." },
], { x: 8.85, y: 1.20, w: 3.70, h: 0.80, fontFace: BODY, fontSize: 11.5, color: "2A1B04",
     align: "right", valign: "middle", margin: 0 });

// ---- KPI row -------------------------------------------------------------
const kpis = [
  ["$12.0mm", "Our maximum bid", "$101,695 per unit  ·  7.28% going-in cap", MOSS],
  ["14.0% / 1.81x", "At our bid", "Levered IRR / equity multiple", MOSS],
  ["7.9% / 1.42x", "At the $13.625mm trade", "What it actually sold for, 12 Jun 26", RED],
  ["16 months", "Renovation complete", "Two tracks  ·  no resident forced out", MOSS],
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
  "118 units, built 1974, Northside Indianapolis. 96.6% occupied, all units classic. Renovate all 118, add in-unit laundry to the 65 without it, modernise the common areas — $2.90mm, $24,575/unit.",
  "Two-track renovation, complete in 16 months. The 53 units that already have laundry are renovated IN PLACE at 10/month from month 2 — no permit, no vacancy loss, premium captured at renewal. The 65 units needing laundry require a state Construction Design Release plus a local permit and a vacant unit, so they run at natural turnover, 5/month from month 4. Nobody is forced out.",
  "The renovation works. $14,407/unit blended earns a $155/month premium — a 12.9% return on cost, and the Year-5 yield on cost of 7.86% sits 111 bps above our 6.75% exit.",
  "The market does not help. Yardi Matrix 1Q-26: metro rents FLAT on a trailing-three-month basis at $1,310 and up just 1.1% year over year, with 2025 starts MORE THAN DOUBLED — delivering 2027-28, inside our hold. We underwrite 1.5/2.0/2.5/3.0/3.0%.",
  "Bid $12.0mm. It traded at $13,625,000 on 12 June 2026 (Marion County deed, $115,466/unit) — 12% above us, where the same plan returns 7.9% and 1.42x. We would have been outbid, and that is the right outcome.",
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
  [tHead("Five-year hold, Aug-26 to Jul-31"), tHead("At our bid\n$12.0mm"), tHead("At the trade\n$13.625mm")],
  tRow("Price per unit", "$101,695", "$115,466"),
  tRow("Going-in cap rate (our underwriting)", "7.28%", "6.41%"),
  tRow("All-in basis per unit", "$126,269", "$140,041"),
  tRow("Sponsor equity", "$5.59mm", "$6.20mm"),
  tRow("Year-1 / Year-5 NOI", "$873K / $1,171K", "$873K / $1,171K"),
  tRow("Year-5 yield on cost vs 6.75% exit", "+111 bps", "+34 bps"),
  tRow("Average cash-on-cash", "8.4%", "6.5%"),
  tRow("Unlevered IRR / multiple", "9.4% / 1.50x", "6.9% / 1.35x"),
  tRow("Levered IRR", "14.0%", "7.9%", true),
  tRow("Equity multiple", "1.81x", "1.42x", true),
], {
  x: 6.95, y: 3.84, w: 5.88, colW: [3.02, 1.46, 1.40],
  rowH: 0.235, fontFace: BODY, border: { type: "solid", color: "2C4A42", pt: 0.5 },
  fill: { color: DARK }, valign: "middle", margin: [2, 5, 2, 5],
});

s1.addText("Sources: rent roll 07/01/2026 · T-12 Feb-25 to Jan-26 · CBRE OM 09/10/2025 · Yardi Matrix Indianapolis 1Q-2026 · Marion County parcel 8000002 assessment and transfer record. Rent growth is set off 1Q-26 ACTUALS, not a forecast. Debt: 65% LTC / 8.0% min debt yield / 5.85% interest-only.",
  { x: 0.5, y: 7.12, w: 12.33, h: 0.24, fontFace: BODY, fontSize: 7.5, color: "8C9A92", margin: 0 });

s1.addNotes("Conservative on what we do not control (rent growth off 1Q-26 actuals, 6.75% exit cap held); confident on what we do control (renovation premium, ancillary income, expense management, two-track pace). The single biggest correction from an earlier draft: the 25-year Marion County assessment record shows assessed value range-bound at $5.1-6.6mm since 2012 and the 2006 sale did NOT reset it, so stepping taxes to a share of the purchase price was wrong. That correction is worth about 1.5 points of IRR and is still stress-tested in Ladder C.");

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
  ["Two-track renovation, 16 months.", "The 53 units with existing laundry go in place at 10/month from month 2 — no permit, no vacancy loss. The 65 needing laundry follow natural turnover at 5/month. All 118 done in month 16 without forcing a single resident out."],
  ["In-unit laundry, all 118 units.", "65 units lack it. At the case's $5,000/unit the laundry component is the highest-return item in the plan. The rent roll shows 53 units with machines, not the OM's 78."],
  ["12.9% return on renovation cost.", "$14,407/unit blended for a $155/month premium, against a 6.75% exit. Year-5 yield on cost of 7.86% sits 111 bps above the exit cap."],
  ["Taxes are lower than they look.", "Marion County parcel 8000002: 2025 actual bill $129,648, assessed value range-bound $5.1-6.6mm since 2012, last change March 2022. The 2006 sale did not reset it. Today's AV is 41% of the 2026 trade price."],
  ["Under-collected ancillary income.", "The $49 technology fee is ~60% penetrated, the $25 valet trash charge is not yet billed although the property starts paying for it in Jan-26, and there is no pet rent — ~$110K a year of run-rate."],
  ["Operational slack.", "2.2% loss to lease and 6.8% trailing physical vacancy against a 5.5% stabilised target."],
];

const RISKS = [
  ["Rent growth. The dominant variable.", "Metro rents flat on a trailing-three-month basis, +1.1% YoY in 1Q-26, and 2025 starts more than doubled — delivering 2027-28, inside our hold, with Carmel taking over half of 2026 deliveries just north of us. At a flat 1.0% the deal returns 7.1%; at 3.0%, 16.7%."],
  ["Basis. We were outbid by 12%.", "It traded at $13,625,000. At that price the same plan returns 7.9% and 1.42x, and the yield-on-cost spread over the exit collapses from 111 bps to 34 bps."],
  ["Tax reassessment.", "Base case grows the actual bill 4% a year. If assessed value reset to 69% of the trade price the Year-5 bill goes to $238,700 and the IRR falls from 14.0% to 10.3%. History says it will not, but Ladder C prices it."],
  ["Unproven rent premium.", "Nothing has been renovated to a current standard. At $125 rather than $155 the IRR is 12.0%; at $190 it is 16.2%. A leased mock-up before we go hard would settle it."],
  ["1974 physical risk.", "3.5 of 5 original rubber roofs, 37 below-grade units needing drains and sump pumps, original sliders, central gas water heaters. Plus 4 evictions and 8 notices on the rent roll."],
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
  aRow("Purchase price — SOLVED, not given", "$12.0mm"),
  aRow("Market rent growth Y1 to Y5", "1.5% to 3.0%"),
  aRow("Track A: start / pace / cost / prem", "mo 2 / 10 / $10k / $100"),
  aRow("Track B: start / pace / cost / prem", "mo 4 / 5 / $18k / $200"),
  aRow("Blended cost / premium / ROC", "$14,407 / $155 / 12.9%"),
  aRow("Non-unit capital + 8% contingency", "$1.20mm"),
  aRow("Stabilised vacancy / reno downtime", "5.5% / 14 days"),
  aRow("Bad debt, Year 1 to Year 5", "1.50% to 1.00%"),
  aRow("Taxes Y1-Y5 (actual bill +4%/yr)", "$137K to $160K"),
  aRow("Operating expenses, Year 1", "$8,929/unit"),
  aRow("Debt: LTC / rate / structure", "65% / 5.85% / IO"),
  aRow("Exit cap / cost of sale", "6.75% / 1.5%"),
], {
  x: 9.2, y: 1.26, w: 3.63, colW: [2.22, 1.41], rowH: 0.205, fontFace: BODY,
  border: { type: "solid", color: LINE, pt: 0.5 }, valign: "middle", margin: [2, 4, 2, 4],
});
s2.addText("Rent growth off Yardi 1Q-26 ACTUALS and taxes off the 25-year Marion County assessment record — not forecasts. Where the rent roll and T-12 conflict with the OM, the Excel files govern (W/D count 53, not 78).",
  { x: 9.2, y: 4.60, w: 3.63, h: 0.58, fontFace: BODY, fontSize: 8, italic: true, color: MUTED,
    valign: "top", margin: 0 });

// ---- chart ---------------------------------------------------------------
s2.addText("Levered IRR by purchase price (6.75% exit cap)", {
  x: 0.5, y: 5.30, w: 5.1, h: 0.26, fontFace: HEAD, fontSize: 11.5, bold: true, color: INK, margin: 0,
});
s2.addChart(pres.ChartType.bar, [{
  name: "Levered IRR",
  labels: ["$10.5mm", "$11.0mm", "$11.5mm", "$12.0mm", "$12.5mm", "$13.625mm", "$14.2mm"],
  values: [0.199, 0.179, 0.159, 0.140, 0.121, 0.079, 0.058],
}], {
  x: 0.42, y: 5.56, w: 5.25, h: 1.42,
  barDir: "col", chartColors: [FOREST, FOREST, FOREST, AMBER, MOSS, RED, RED],
  showLegend: false, showTitle: false,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%",
  dataLabelFontSize: 7.5, dataLabelColor: INK, dataLabelFontFace: BODY,
  catAxisLabelFontSize: 7.5, catAxisLabelColor: MUTED, catAxisLabelFontFace: BODY,
  catGridLine: { style: "none" },
  valAxisHidden: true, valGridLine: { style: "none" },
  valAxisMaxVal: 0.20, valAxisMinVal: 0,
  barGapWidthPct: 45, plotArea: { fill: { color: WHITE } },
});
s2.addText("14% hurdle is cleared at or below $12.0mm. It traded at $13.625mm.", {
  x: 0.5, y: 6.99, w: 5.1, h: 0.22, fontFace: BODY, fontSize: 8, italic: true, color: MUTED, margin: 0,
});

// ---- overall assessment --------------------------------------------------
s2.addShape(pres.ShapeType.roundRect, {
  x: 6.0, y: 5.22, w: 6.83, h: 1.90, fill: { color: DARK }, rectRadius: 0.05, shadow: sh(),
});
s2.addText("Overall Assessment", { x: 6.3, y: 5.34, w: 6.2, h: 0.26, fontFace: HEAD,
  fontSize: 12.5, bold: true, color: MOSS, margin: 0 });
s2.addText("A well-located asset with a genuinely accretive plan: 12.9% return on renovation cost, all 118 units done in 16 months without displacing a resident, and a Year-5 yield on cost 111 bps above our exit. Our maximum bid is $12.0mm. It traded at $13,625,000 — 12% above us — where the same plan returns 7.9%. We were outbid by someone with a cheaper cost of capital or a braver rent view, and on 1.1% actual metro rent growth we would make the same call again. Being disciplined on the one variable you fully control is the job.",
  { x: 6.3, y: 5.64, w: 6.25, h: 1.40, fontFace: BODY, fontSize: 9.5, color: "DFE6E0",
    lineSpacing: 12.5, valign: "top", margin: 0 });

s2.addNotes("We would have lost this deal, and should have. If we wanted to win it, the two honest routes are a proven rent premium (a leased mock-up moves the bid to ~$12.6mm at $170) or a cheaper scope that still holds the premium (Ladder B). Two structures worth testing before walking: a tax escrow or price adjustment tied to the first post-closing Form 11A, and a seller credit for the identified immediate needs (sump pumps, patio drains, asphalt — about $190K).");

pres.writeFile({ fileName: "Chateau_in_the_Woods_IC_Presentation.pptx" })
  .then(f => console.log("wrote", f));
