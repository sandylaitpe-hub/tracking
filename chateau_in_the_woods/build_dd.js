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
// Single source of truth, shared with build_model.py so the Word list and the
// workbook tab can never drift apart.
const SECTIONS = JSON.parse(fs.readFileSync("dd_questions.json", "utf8"));

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
   "The OM's 78-unit figure against the rent roll's 53 is a $125,000 swing in the capital budget — $50,000 for " +
   "every 10 units of error — and it decides which units sit in which renovation track. The 65 units that need " +
   "laundry carry the case's $5,000 addition and earn $200 a month rather than $100. We underwrite the rent " +
   "roll per the case data hierarchy, but this is the first thing we would walk the property to confirm."],
  ["Post-sale tax reassessment (Section 6).",
   "Net assessed value of $5,214,932 against a $13,625,000 trade price — 41%, before any reassessment. The " +
   "base case grows the actual 2025 bill of $129,648 at 4% a year, which the 25-year record supports: assessed " +
   "value has been range-bound at $5.1-6.6mm since 2012 and the 2006 sale did not reset it. If it did reset to " +
   "the 69% of sale price the OM's own comp study assumes, the Year-5 bill would be $238,700 rather than " +
   "$160,300 — about $1.2mm of value at our exit cap, and 3.7 points of levered IRR. Priced as Ladder C."],
  ["Evidence of an achievable rent premium (Section 3).",
   "Nothing has been renovated to a current standard, so the $155/month blended premium is inferred from the " +
   "comp set rather than proven on site. Roughly every $15 a month is worth a point of levered IRR, which is " +
   "why the model sensitises the premium from $100 to $190. A single leased mock-up unit before we go hard " +
   "would settle it."],
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
    t("No purchase price was given, so the model solves for one: $12,000,000 is the highest basis at which " +
      "this plan clears a 14% levered IRR and a 1.80x multiple. The asset traded on 12 June 2026 for " +
      "$13,625,000, where the same plan returns 7.9% and 1.42x. The business plan is sound at either price; " +
      "only the basis is in question. Diligence should therefore be aimed at the three items above and at the " +
      "physical scope in Section 4 — those are what could move our bid, or take us off the deal.", {}),
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
