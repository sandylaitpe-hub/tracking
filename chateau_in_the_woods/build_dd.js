const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageOrientation, Footer, PageNumber, PageBreak,
} = require("docx");
const fs = require("fs");

const INK = "18231E";
const FOREST = "2C5F2D";
const MUTED = "6E7B73";
const HDRFILL = "16302A";
const ALT = "F0F3EF";
const LINE = "D8DED9";

// docx-js swaps width/height for landscape, so pass PORTRAIT dimensions here.
const PAGE_W = 12240;        // portrait letter width; becomes the page HEIGHT in landscape
const PAGE_H = 15840;        // portrait letter height; becomes the page WIDTH in landscape
const MARGIN = 720;          // 0.5"
const TABLE_W = PAGE_H - 2 * MARGIN;    // 14400 DXA = 10" of usable landscape width
const COLS = [560, 5400, 6640, 1800];   // #, question, assumed answer, priority — sums to 14400

const cell = (children, opts = {}) => new TableCell({
  children,
  width: { size: opts.w, type: WidthType.DXA },
  shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: "auto" } : undefined,
  margins: { top: 60, bottom: 60, left: 90, right: 90 },
  verticalAlign: opts.valign || "top",
});

const p = (runs, opts = {}) => new Paragraph({
  children: runs,
  spacing: { after: opts.after === undefined ? 0 : opts.after, line: 240 },
  alignment: opts.align,
});

const t = (text, opts = {}) => new TextRun({
  text, font: opts.font || "Calibri", size: opts.size || 17,   // half-points
  bold: opts.bold, italics: opts.italics, color: opts.color || INK,
});

// ---------------------------------------------------------------------------
const SECTIONS = [
  ["1. Rent Roll, Leases & Leasing Performance", [
    ["Please provide the complete lease file for all 118 units: executed leases, every addendum, renewal notices, and the current concession log by unit.",
     "No unusual concessions or side agreements. Concessions underwritten at 0.75% of gross potential rent in Years 1-2, falling to 0.50%.", "High"],
    ["The 07/01/2026 rent roll shows 4 units in Evict and 8 in Notice-Unrented. For each, what is the status of the action, what is owed, and when is possession expected?",
     "All four evictions resolve within 90 days of closing and the units are re-leased at renovated rents. The cost sits inside the 1.50% Year-1 bad debt allowance.", "High"],
    ["Unit 4017-C carries a lease rent of $821.03 against a $1,250 market rent and is on notice. Is that a concession, an employee unit, a subsidy, or a partial-month proration?",
     "A proration or one-off concession rather than a genuine below-market lease. Modelled at market rent on turn.", "Medium"],
    ["Please provide 24 months of leasing traffic, tour-to-lease conversion, renewal rate, average renewal increase achieved, and the ILS source report.",
     "Renewal rate of 55-60%, producing roughly 55-60 natural turns a year. This is what makes a 5-unit-per-month renovation pace achievable without forcing residents out.", "High"],
    ["Which units are on month-to-month tenancies and what is being charged? The OM cites a $100/month month-to-month fee that the T-12 does not separately disclose.",
     "A small number of month-to-month units; the premium income sits inside miscellaneous other income.", "Low"],
    ["Units 4005-D, 4029-F, 4013-A, 4017-E and 4034-E show a lease-from date but no lease-to date. Are these expired leases holding over?",
     "Holdover on month-to-month, re-leased at market rent when renovated.", "Medium"],
    ["Please confirm there are no affordable, Section 8, HAP, LIHTC or other regulatory restrictions on any unit.",
     "The property is entirely market rate with no restrictions.", "High"],
  ]],
  ["2. Washer / Dryer Inventory — the largest data conflict in the package", [
    ["Please confirm, unit by unit, which units contain a washer and dryer, which have hookups only, and which have neither. The OM states that 78 of 118 units have a full-size or stackable washer/dryer, then lists only 53 side-by-side plus 1 stackable, and its own floor-plan and building-mix schedules foot to 54.",
     "The rent roll governs per the case data hierarchy: 53 units have machines (Rhone 15, Roman 4, Cabernet 4, Brunswick 30) and 65 require both hookups and machines at $5,000 each. Every 10 units of error moves the capital budget by $50,000.", "High"],
    ["For the 65 units without laundry, please provide plumbing and electrical capacity confirmation: panel amperage, riser and stack capacity, and whether a dedicated 240V circuit and a drain can be run inside the $5,000 allowance.",
     "The $5,000 per unit prescribed by the case covers hookups plus machines with no panel or riser upgrade required. A riser or panel constraint would be a material change.", "High"],
    ["The OM states the Verdun 2-bed / 2-bath layout supports adding machines in the bedroom closet. Has an architect or general contractor confirmed this for all 25 Verdun units, and does it reduce the closet below any code or marketing minimum?",
     "All 25 Verdun units accommodate stacked machines without losing a required closet. CBRE's own forecast assumes the same.", "High"],
    ["The OM identifies unit 4017-D as a Bordeaux with stackable machines at a $1,307 market rent, but the rent roll shows all 25 Bordeaux units at $1,260 in a single unit type. Which is correct?",
     "The rent roll governs, so 4017-D is underwritten as needing the $5,000 addition. Conservative by $5,000.", "Low"],
    ["What was the previous washer/dryer rental programme, why was it discontinued, and are there surviving vendor contracts or leased equipment on site?",
     "No surviving equipment lease. The 15 laundry centres are owned and can be repurposed or removed once in-unit laundry is universal.", "Medium"],
    ["Once every unit has in-unit laundry, is there a use or conversion opportunity for the 15 laundry centres (3 per building)?",
     "No value underwritten. Any storage or unit conversion would be incremental upside.", "Low"],
  ]],
  ["3. Renovation Scope, Cost & Rent Premium", [
    ["Please provide the scope, unit-level cost detail and before/after rents for the 2006 rehabilitation of 116 units, and identify the two units never rehabbed.",
     "The 2006 scope was cabinets, counters and flooring. Those components are now 20 years old, which is why every unit reads as classic against today's market. The two un-rehabbed units require full scope.", "High"],
    ["Have any units been renovated to a current standard in the last three years? If so, what premium was achieved, over what lease-up period, and at what cost?",
     "No recent renovations, so there is no in-place proof of premium. This is the single largest unproven assumption in the model and the reason premium is sensitised from $125 to $190 per month.", "High"],
    ["Please provide two or three general contractor bids on our scope for a representative Abbey, Verdun and Brunswick unit, inclusive of the washer/dryer addition.",
     "$13,000 per unit interior plus $5,000 for the laundry addition, a $15,754 blended cost. A $2,000 per unit miss is $236,000 of capital.", "High"],
    ["What is realistic turn time from vacate to rent-ready under our scope, and can the site sustain five units a month with the current team?",
     "14 days of incremental downtime beyond a normal turn. We fund a second maintenance technician in payroll.", "Medium"],
    ["Are there asbestos, lead paint or mould surveys? A 1974 building rehabbed in 2006 may have regulated materials behind cabinets and in flooring mastic.",
     "No abatement beyond what the 8% contingency absorbs. A material finding would change the programme cost and pace.", "High"],
    ["Do the units have individual electrical panels sized for modern appliance loads, and is there capacity for in-unit laundry plus upgraded lighting?",
     "Cutler-Hammer breaker boxes with adequate capacity, per the OM systems schedule.", "Medium"],
  ]],
  ["4. Physical Condition & Capital Requirements", [
    ["Please provide the most recent property condition assessment, all roof inspection reports, and the roof warranty documentation for the 2020 and 2024 replacements.",
     "3.5 of the 5 rubber roofs are original and require replacement at roughly $103,000 each. Underwritten at $360,000 phased over Years 2-4. Building 4 cost $90,000 in 2020 and half of Building 5 cost $54,507 in 2024.", "High"],
    ["The OM identifies 37 below-grade units with patio drains and 4 sump pumps needed. Is there any history of water intrusion, insurance claims, mould remediation or resident complaints in those units?",
     "No systemic intrusion; $40,000 resolves the drains and pumps. A systemic finding would be a material adverse item and would change the recommendation.", "High"],
    ["What is the remaining useful life of the central gas water heaters and the 8 individual electric units? The T-12 shows $29,053 of water heater replacement in just two months.",
     "Ongoing replacement runs through the $300 per unit replacement reserve. A full central plant replacement is not underwritten.", "High"],
    ["Please provide a unit-level inventory with the age and condition of the individual gas furnaces and condensing units.",
     "Staggered ages with replacement funded through the reserve; $3,946 of HVAC replacement appears in the T-12.", "Medium"],
    ["Windows and sliding patio doors are original aluminium sliders. Please provide any quote for full replacement against the current practice of replacing fogged glass on turns.",
     "Glass replacement on turns at $150,000 across three years, not full window replacement.", "Medium"],
    ["Please provide the paving assessment and any quote for the parking and carport areas, plus the warranties on the 2022 pool deck and the 2020 Hardie siding and gutter work.",
     "$150,000 for patch, sealcoat and stripe. The envelope and pool deck require no capital during the hold.", "Medium"],
    ["Is the property fully separately metered for electricity, and are there master-metered obligations beyond gas and water/sewer?",
     "Electricity is individually metered and resident-paid, per the OM utility schedule.", "Low"],
    ["Please provide the elevator, boiler, fire suppression and life-safety inspection certificates, and the status of the building entry systems the OM flags for upgrade.",
     "All current. $60,000 underwritten for entry system upgrades across the 20 building entries.", "Medium"],
  ]],
  ["5. Financial Statements & Operations", [
    ["Please provide 24 months of monthly general ledger detail, monthly bank statements, and an accounts receivable ageing by unit.",
     "The T-12 provided is complete and accrual-accurate. The AR ageing supports the 1.50% Year-1 bad debt allowance.", "High"],
    ["Bad debt: $28,229 was written off against $19,050 collected over the T-12, and January 2026 alone was $5,670. Please explain the trend and describe the current collections process.",
     "Deterioration is driven by the four evictions in process and normalises as the resident profile improves post-renovation. Underwritten at 1.50% falling to 1.00% of gross potential rent.", "High"],
    ["General and administrative expense carries $17,528 of legal and $14,730 of tax and audit fees concentrated in December 2025 and January 2026. Are these transaction costs that should be excluded from a run-rate?",
     "Sale-related and non-recurring. Underwritten at $50,000 against $62,887 in the T-12 and CBRE's $42,209 forecast.", "Medium"],
    ["Please provide the Comcast bulk internet contract: the 7-year term, the escalator, the stabilised per-unit cost, and confirmation of the $49 resident charge and its billing mechanics.",
     "$27.95 per unit per month stabilised, escalating 3%, billed to residents at $49. The T-12 captures only $11,068 of a ramping contract against $39,577 stabilised — understating this line is the easiest way to overstate NOI on this deal.", "High"],
    ["Please provide the valet trash contract, its start date, per-unit cost and term, and confirm what share of residents are currently being charged the $25.",
     "The property begins paying the contractor in January 2026 and the charge is recoverable from all residents on renewal. Ramped at 75% in Year 1 to 100% by Year 3.", "High"],
    ["What portion of the T-12 'Utility Fees' line of $97,812 is water and sewer recovery versus valet trash, and what is the actual recovery ratio on water and sewer usage?",
     "85% water and sewer recovery with valet trash charged separately, consistent with CBRE's own restatement. Historical recovery has ranged from 78% to 128%.", "High"],
    ["Please provide the payroll register and benefits detail by position, and confirm which costs are reimbursed by the manager versus paid directly.",
     "$235,000 in Year 1 for a manager, part-time leasing consultant, maintenance supervisor and one technician, growing 3.5%.", "Medium"],
    ["Is the current property management agreement assumable, and what is the fee and any termination cost?",
     "Terminable at closing. 4.0% of effective gross revenue going forward against an effective 3.5% in the T-12.", "Low"],
    ["Please provide the ResMan chart of accounts and confirm the accrual policy for real estate taxes; the October 2025 entry of $17,958 appears to be a catch-up.",
     "An accrual true-up rather than a rate change. Year-1 taxes underwritten at $140,000.", "Medium"],
    ["Storage: only 5 of 23 units are rented at $15-20. Is there a reason demand is so low, and is $25 achievable?",
     "No structural constraint. Underwritten at 12 units rented at $25, or $3,600 a year against $141 actual.", "Low"],
    ["The carport income of $16,197 implies roughly 54 of 74 carports rented at $25. What drives the vacancy, and is $35-40 achievable post-renovation?",
     "Held at the T-12 actual with no rate increase underwritten. Any increase is upside.", "Low"],
  ]],
  ["6. Real Estate Taxes — the largest expense judgement", [
    ["Please provide the last three Form 11A reassessment notices, the current tax bills, and any pending appeal or settlement.",
     "The 2025 assessed value of $5,547,800 and the 2.539% rate are correct and no appeal is pending.", "High"],
    ["Has the seller obtained guidance from the Marion County Assessor on how a sale at this price will be treated, given that Indiana values apartments at the lowest of the income, market and cost approaches?",
     "The income approach governs, so assessed value tracks NOI rather than the sale price. Taxes are stepped from $140,000 in Year 1 to $210,000 in Year 5, roughly 10% a year. If assessed value instead moved to the 69% of sale price the OM's own comp study shows, the tax bill would rise about $110,000 a year and cost roughly $1.6mm of value.", "High"],
    ["Please confirm the SB-1 deduction schedule as it applies to this parcel, and the referendum rate applicable in taxing district 800.",
     "CBRE's SB-1 forecast is correct: a 6% assessed value deduction for taxes payable 2026 rising to 33.4% by 2031, taxed at the lower of the actual rate or the 2% cap plus referendum.", "High"],
    ["What is the storm water fee and is it assessed separately from the real estate tax bill?",
     "$2,966 in 2026 growing 5% a year, included inside the tax line.", "Low"],
    ["Are there personal property taxes, and does the SB-1 personal property deduction eliminate the $4,110 assessed in 2025?",
     "Eliminated under the $1mm deduction for taxes payable 2026. Nothing underwritten.", "Low"],
  ]],
  ["7. Insurance & Risk", [
    ["Please provide five years of loss runs, the current policy with limits and deductibles, and the most recent insurance appraisal or replacement-cost estimate.",
     "No material loss history. $95,000 in Year 1 ($805 per unit) growing 5%, against $94,611 in the T-12 and CBRE's $88,326 forecast.", "High"],
    ["Are there open or threatened claims, habitability complaints, fair housing matters, code violations or litigation involving the property or its management?",
     "None. No litigation reserve is underwritten.", "High"],
    ["Please confirm the flood determination. The OM cites Zone X on panel 18097C0151G, but 37 units are below grade.",
     "Zone X confirmed and no flood insurance required. The below-grade units are a drainage matter, not a flood zone matter.", "Medium"],
    ["Please provide any environmental site assessment (Phase I, and Phase II if one exists), including any underground storage tank history.",
     "Clean Phase I with no recognised environmental conditions.", "High"],
  ]],
  ["8. Market, Comparables & the Rent Thesis", [
    ["Please provide the underlying shop reports for the nine rent comparables, including concessions offered, renovation status, and in-unit laundry availability at each.",
     "Comp rents are net of concessions and the renovated comps include in-unit laundry, which is what makes renovated rents of $1,250-$1,725 supportable.", "High"],
    ["Avalon Lake is the same 1974 vintage nearby and achieves $1,178 / $1,655 / $1,834 for one, two and three bedrooms. What explains a two-bedroom gap of roughly $400 to Chateau, and how much of it is interior condition versus location or amenity?",
     "The gap is predominantly interior condition and in-unit laundry, which is exactly what our capital addresses. This single assumption carries the business plan.", "High"],
    ["What new supply is under construction or permitted within three miles, and on what delivery schedule?",
     "No material Class B/C competitive supply. Northside Indianapolis deliveries are Class A at rents well above our renovated levels.", "High"],
    ["What are current market concessions in the submarket for renovated 1970s product?",
     "Two to four weeks free on new leases, inside the 0.50-0.75% concession allowance.", "Medium"],
    ["Please provide any third-party market study, and current submarket occupancy, absorption and rent growth data.",
     "Stable submarket occupancy in the low-90s supporting a 5.0% stabilised vacancy assumption and 3.0% annual market rent growth.", "Medium"],
  ]],
  ["9. Title, Legal, Zoning & Contracts", [
    ["Please provide the title commitment, ALTA survey, a zoning verification letter for the D-1 designation with confirmation of legal non-conforming status if applicable, and all certificates of occupancy.",
     "Clean title with no adverse easements, and the existing density of 12.55 units per acre on 9.40 acres is either permitted or a legal non-conforming use.", "High"],
    ["Please confirm the 8 units converted from storage are properly permitted with certificates of occupancy, and that the proposed Building 4042 storage conversion would be permittable.",
     "The 8 converted units are legal and permitted. The 4042 conversion is upside and is not underwritten.", "Medium"],
    ["Please provide all service contracts with terms, notice periods and termination rights: landscaping, pest control, valet trash, trash hauling, laundry, security and fire, pool, and internet.",
     "All are terminable at or shortly after closing except the 7-year bulk internet contract, which is assumed and underwritten.", "Medium"],
    ["Please confirm no existing debt is being assumed and that the property is delivered free and clear, per the OM offering terms.",
     "Free and clear. New financing at 65% loan-to-cost, constrained by an 8.0% minimum debt yield.", "Low"],
    ["Are there any tenant-in-common, ground lease, easement, reciprocal access or shared-facility arrangements?",
     "None. Fee simple ownership of the entire 9.40-acre parcel.", "Medium"],
  ]],
  ["10. Items we would price rather than assume", [
    ["What will the seller accept, and is guidance still the 6.9% Year-1 cap rate on CBRE's $980,838 Acquisition Forecast — roughly $14.2mm?",
     "$14.2mm is the ask. On our underwriting that produces a 6.7% levered IRR and a 1.35x multiple over five years, which is why the recommendation is conditional on basis rather than on the business plan. Our bid is $12.25mm with a hard walk at $13.0mm.", "High"],
    ["Would the seller entertain a structure that shares the reassessment risk — a tax escrow, or a purchase price adjustment tied to the first post-closing Form 11A?",
     "No structure available, so the buyer takes reassessment risk. This is why taxes are underwritten at roughly 10% annual growth.", "Medium"],
    ["Would the seller fund or credit the identified immediate needs — the 4 sump pumps, the patio drains and the asphalt work?",
     "No credit given; the full $190,000 sits in our Year-1 capital budget.", "Medium"],
    ["Is the seller willing to give an earn-out or holdback against the 4 evictions and 8 notice units being delivered vacant or current?",
     "No holdback; the cost sits inside the Year-1 bad debt and vacancy allowance.", "Low"],
    ["What is the required diligence period and closing timeline, and what earnest money structure is expected?",
     "A 60-day diligence period and 30-day close, with earnest money going hard at the end of diligence.", "Medium"],
  ]],
];

// ---------------------------------------------------------------------------
const children = [];

children.push(new Paragraph({
  children: [t("Chateau in the Woods", { font: "Cambria", size: 34, bold: true, color: INK })],
  spacing: { after: 40 },
}));
children.push(new Paragraph({
  children: [t("Due Diligence & Seller / Broker Question List", { font: "Cambria", size: 24, color: FOREST })],
  spacing: { after: 60 },
}));
children.push(new Paragraph({
  children: [t("118 Units  ·  4020 Monaco Drive, Indianapolis, IN 46220  ·  Built 1974  ·  134,371 SF", { size: 18, color: MUTED })],
  spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 8 } },
}));

children.push(new Paragraph({
  children: [
    t("Purpose. ", { bold: true }),
    t("This list sets out the information we would request from the seller and broker to firm up the underwriting. " +
      "Because the case requires a fully functioning model, the third column records the reasonable answer we have " +
      "assumed in the five-year model in the absence of a response. Items marked ", {}),
    t("High", { bold: true, color: FOREST }),
    t(" are those where a different answer changes the underwriting output or the recommendation itself; the rest " +
      "refine precision without moving the decision.", {}),
  ],
  spacing: { after: 120, line: 260 },
}));
children.push(new Paragraph({
  children: [
    t("Data hierarchy. ", { bold: true }),
    t("Per the case instructions, the Excel rent roll (07/01/2026) and the T-12 profit and loss (February 2025 " +
      "through January 2026) govern wherever they conflict with the CBRE Offering Memorandum dated 09/10/2025. " +
      "The most consequential such conflict is the in-unit washer/dryer count, addressed in Section 2.", {}),
  ],
  spacing: { after: 240, line: 260 },
}));

const HDR = (label, w) => cell([p([t(label, { bold: true, size: 17, color: "FFFFFF" })])],
  { w, fill: HDRFILL, valign: "center" });

let n = 0;
SECTIONS.forEach(([title, items]) => {
  children.push(new Paragraph({
    children: [t(title, { font: "Cambria", size: 22, bold: true, color: FOREST })],
    spacing: { before: 260, after: 100 },
    keepNext: true,
  }));

  const rows = [new TableRow({
    tableHeader: true,
    children: [HDR("#", COLS[0]), HDR("Question for the seller / broker", COLS[1]),
               HDR("Reasonable answer assumed in the model", COLS[2]), HDR("Priority", COLS[3])],
  })];

  items.forEach(([q, a, pri], i) => {
    n += 1;
    const fill = i % 2 === 1 ? ALT : undefined;
    rows.push(new TableRow({
      cantSplit: true,
      children: [
        cell([p([t(String(n), { bold: true, color: MUTED })], { align: AlignmentType.CENTER })], { w: COLS[0], fill }),
        cell([p([t(q)])], { w: COLS[1], fill }),
        cell([p([t(a, { color: "3D4A43" })])], { w: COLS[2], fill }),
        cell([p([t(pri, { bold: pri === "High", color: pri === "High" ? FOREST : MUTED })],
          { align: AlignmentType.CENTER })], { w: COLS[3], fill }),
      ],
    }));
  });

  children.push(new Table({
    rows,
    columnWidths: COLS,
    width: { size: TABLE_W, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE },
    },
  }));
});

children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(new Paragraph({
  children: [t("Where the Answers Matter Most", { font: "Cambria", size: 22, bold: true, color: FOREST })],
  spacing: { after: 120 },
}));
children.push(new Paragraph({
  children: [
    t(`${n} questions across 10 categories. `, { bold: true }),
    t("Most refine precision. Three move the recommendation on their own, and we would not go hard on earnest " +
      "money without answers to all three.", {}),
  ],
  spacing: { after: 160, line: 260 },
}));

const CRITICAL = [
  ["Washer / dryer inventory (Section 2).",
   "The OM's 78-unit figure against the rent roll's 53 is a $60,000 swing in the capital budget and changes " +
   "which units earn the $75 laundry premium — the highest-return component of the whole plan. We underwrite " +
   "the rent roll, but this is the first thing we would walk the property to confirm."],
  ["Post-sale tax reassessment (Section 6).",
   "Assessed value of $5.55mm against a $14.2mm price. If assessed value moved to the 69% of sale price the " +
   "OM's own comp study shows, the bill would rise roughly $110,000 a year — about $1.6mm of value at our exit " +
   "cap, or more than two points of levered IRR."],
  ["Evidence of an achievable rent premium (Section 3).",
   "Nothing has been renovated to a current standard, so the $166/month blended premium is inferred from the " +
   "comp set rather than proven on site. Every $25 per month is worth roughly 1.3 points of levered IRR, which " +
   "is why the model sensitises premium from $125 to $190."],
];
CRITICAL.forEach(([h, b], i) => {
  children.push(new Paragraph({
    children: [t(`${i + 1}.  `, { bold: true, color: FOREST }), t(h, { bold: true }), t("  " + b, {})],
    spacing: { after: 140, line: 260 },
    indent: { left: 260, hanging: 260 },
  }));
});

children.push(new Paragraph({
  children: [
    t("Position going in. ", { bold: true }),
    t("The five-year model returns a 6.7% levered IRR and a 1.35x multiple at the $14.2mm implied by the OM's " +
      "6.9% cap rate, against a 13.8% IRR and 1.81x at our $12.25mm bid. The business plan is sound at either " +
      "price; only the basis is in question. Diligence should therefore be aimed at the three items above and " +
      "at the physical scope in Section 4 — those are what could move our bid, or take us off the deal.", {}),
  ],
  spacing: { before: 60, after: 220, line: 260 },
}));

children.push(new Paragraph({
  children: [t("Sources: rent roll export 07/01/2026 · Trailing Profit and Loss Detail, January 2026 (accrual, " +
    "ResMan), covering February 2025 through January 2026 · Chateau in the Woods Offering Memorandum, CBRE / " +
    "Midwest Multifamily Advisory Group, prepared 09/10/2025.", { size: 15, italics: true, color: MUTED })],
  spacing: { line: 240 },
}));

const doc = new Document({
  creator: "Investment Committee",
  title: "Chateau in the Woods — Due Diligence Question List",
  styles: { default: { document: { run: { font: "Calibri", size: 17, color: INK } } } },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.LANDSCAPE },
        margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            t("Chateau in the Woods — Due Diligence Question List    ", { size: 15, color: MUTED }),
            new TextRun({ children: [PageNumber.CURRENT], size: 15, color: MUTED, font: "Calibri" }),
            t(" of ", { size: 15, color: MUTED }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15, color: MUTED, font: "Calibri" }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("Chateau_in_the_Woods_Due_Diligence_Questions.docx", buf);
  console.log("wrote Chateau_in_the_Woods_Due_Diligence_Questions.docx —", n, "questions");
});
