# Chateau in the Woods — Value-Add Case Study

118 units · 4020 Monaco Drive, Indianapolis, IN 46220 · built 1974 · 134,371 SF · 96.6% occupied

## Deliverables

| File | What it is |
|---|---|
| `Chateau_in_the_Woods_5yr_Cash_Flow_Model.xlsx` | Dynamic five-year cash flow model, 12 tabs, 2,780 formulas, zero formula errors. Every projection cell is a formula off the `Assumptions` tab. |
| `Chateau_in_the_Woods_IC_Presentation.pptx` | Two-slide investment committee memorandum: recommendation and returns; opportunities, risks, assumptions and overall assessment. |
| `Chateau_in_the_Woods_Due_Diligence_Questions.docx` | 62 diligence questions across 10 categories, each paired with the reasonable answer assumed in the model. |

## Recommendation

**Conditional — approve the business plan and the capital, not the price.**

Bid **$12.25mm ($103,814/unit)**. Best and final $12.5mm. Hard walk at $13.0mm.

| | At our bid $12.25mm | At the ask $14.2mm |
|---|---|---|
| Going-in cap rate (our Year-1 NOI) | 7.15% | 6.17% |
| All-in basis per unit | $129,843 | $146,369 |
| Sponsor equity | $5.75mm | $6.74mm |
| Year-5 yield on cost vs 6.75% exit | +114 bps | +25 bps |
| Average cash-on-cash | 8.0% | 6.0% |
| Unlevered IRR / multiple | 9.3% / 1.50x | 6.5% / 1.33x |
| **Levered IRR / multiple** | **13.8% / 1.81x** | **6.7% / 1.35x** |

The renovation is a good project: $15,754/unit blended earns a $166/month premium, a 12.7% return on
cost against a 6.75% exit. The constraint is basis. Capitalised at exit, the entire $3.07mm capital
plan nets **$9,867** of value — the interior scope creates $1.07mm and $1.06mm of deferred maintenance
gives all of it back. And the asking price is struck on a forecast NOI that already embeds the
loss-to-lease burn-off, the recovery to 95% occupancy and the ancillary income a buyer has to create.

No purchase price was given in the case materials. $14,200,000 is derived from the OM cover: a 6.9%
Year-1 cap rate on CBRE's $980,838 Acquisition Forecast NOI. It is an input on the `Assumptions` tab
and is sensitised live on the `Sensitivities` tab.

## Model tabs

`README` · `Assumptions` · `Rent Roll` · `Unit Mix` · `T-12 Recast` · `Other Income` ·
`Capex Budget` · `Reno Schedule` · `Annual Model` · `Returns` · `Sensitivities` · `DD Questions`

Month 1 is August 2026 (the rent roll is dated 07/01/2026). The annual operating model is driven by a
60-month renovation schedule that converts units from classic to renovated rent, spends capital and
charges renovation downtime vacancy. All capital is funded at closing in Sources & Uses, so annual
cash flow deducts interest and replacement reserves only and never double-counts capex — conservative,
since it puts the full equity out on day one.

Both sensitivity grids are live: each cell drives its own cash-flow row and `IRR()`, so they
recalculate with any change to the assumptions rather than being pasted values.

## Data hierarchy

Per the case instructions the Excel rent roll and T-12 govern over the Offering Memorandum. The most
consequential conflict is the in-unit washer/dryer count: the OM says 78 of 118 units have machines but
then lists 53 side-by-side plus 1 stackable, and its own floor-plan and building-mix schedules foot to
54. The rent roll shows 53 units in a "w/ WD" unit type, so 53 units have laundry and **65 carry the
$5,000 addition** prescribed by the case. Every 10 units of error moves the capital budget by $50,000.

Status rules applied as instructed: `Notice-Unrented` and `Evict` count as occupied, any variation of
`Vacant` counts as vacant — 114 occupied, 4 vacant, 96.6%.

## Reproducing

```bash
pip install openpyxl && npm install pptxgenjs docx

python3 build_model.py      # writes the .xlsx
node    build_deck.js       # writes the .pptx
node    build_dd.js         # writes the .docx  (reads dd_questions.json)
python3 verify.py           # independent replica — ties to the workbook to the dollar
```

`build_model.py` and `build_dd.js` both read `dd_questions.json`, so the workbook's `DD Questions` tab
and the Word document never drift apart.

`verify.py` re-implements the whole model from the source data in plain Python. It is the tie-out that
proves the workbook's formulas compute what they are supposed to — NOI by year, the loan sizing tests,
the exit, both IRRs and the price ladder all agree exactly.

## Source files

1. `rent_roll20260701_1.xlsx` — rent roll as of 07/01/2026
2. `Chateau_Trailing_Profit_And_Loss_Detail_01312026.xlsx` — T-12 accrual P&L, Feb 2025 – Jan 2026
3. `Chateau_in_the_Woods_Offering_Memorandum.pdf` — CBRE / Midwest Multifamily Advisory Group, 09/10/2025

Source files are not committed; the build scripts read them from the upload directory.
