# Chateau in the Woods — Value-Add Case Study

118 units · 4020 Monaco Drive, Indianapolis, IN 46220 · built 1974 · 134,371 SF · 96.6% occupied

## Deliverables

| File | What it is |
|---|---|
| `Chateau_in_the_Woods_5yr_Cash_Flow_Model.xlsx` | Dynamic five-year cash flow model, 13 tabs, 3,332 formulas, zero formula errors. Every projection cell is a formula off the `Assumptions` tab. |
| `Chateau_in_the_Woods_IC_Presentation.pptx` | Chart-led investment committee deck: two main pages (recommendation and returns; where the value comes from and what moves it) plus three appendix pages (business plan, opportunities and risks, assumptions and the sensitivity grid). Eight native charts — the price ladder, the yield-on-cost spread, the two-track schedule and a breakeven tornado. |
| `Chateau_in_the_Woods_Due_Diligence_Questions.docx` | 62 diligence questions across 10 categories, each paired with the reasonable answer assumed in the model. |

## The question this model answers

No purchase price was given. So the model **solves for one**: the highest basis at which the five-year
business plan clears a 14% levered IRR and a 1.80x equity multiple on a conservative case.

**Answer: $12,000,000 ($101,695/unit, 7.28% going-in cap on our Year-1 NOI).**

The asset actually traded on 12 June 2026 for **$13,625,000** ($115,466/unit) — Marion County deed,
special warranty, assessor validity flag "Y". That is 12% above our bid, and at that price the same
plan returns 7.9% and 1.42x. We would have been outbid, and on 1.1% actual metro rent growth that is
the right outcome.

| Five-year hold, Aug-26 to Jul-31 | At our bid $12.0mm | At the trade $13.625mm |
|---|---|---|
| Price per unit | $101,695 | $115,466 |
| Going-in cap rate (our Year-1 NOI) | 7.28% | 6.41% |
| All-in basis per unit | $126,269 | $140,041 |
| Sponsor equity | $5.59mm | $6.20mm |
| Year-5 yield on cost vs 6.75% exit | **+111 bps** | +34 bps |
| Average cash-on-cash | 8.4% | 6.5% |
| Unlevered IRR / multiple | 9.4% / 1.50x | 6.9% / 1.35x |
| **Levered IRR / multiple** | **14.0% / 1.81x** | **7.9% / 1.42x** |

## Two-track renovation — 16 months, nobody displaced

| | Track A | Track B |
|---|---|---|
| Units | 53 (already have laundry) | 65 (need laundry added) |
| Method | **In place** — resident stays | Vacant unit |
| Permit | None — cosmetic work is exempt | State CDR **then** local DBNS permit |
| Start / pace | Month 2 / 10 per month | Month 4 / 5 per month (natural turnover) |
| Cost per unit | $10,000 | $13,000 + $5,000 laundry (case-set) |
| Premium | $100/month, captured at renewal (6-month lag) | $200/month, captured immediately |
| Downtime | None | 14 days incremental |

Blended **$14,407/unit for a $155/month premium = 12.9% return on cost**. All 118 units complete in
month 16; a single-track plan at 5 a month would take 27.

## Assumptions set off actuals, not forecasts

- **Rent growth 1.5 / 2.0 / 2.5 / 3.0 / 3.0%.** Yardi Matrix Indianapolis 1Q-2026: metro asking rent
  FLAT on a trailing-three-month basis at $1,310 and up just **1.1% year over year**, with 2025
  construction starts **more than doubled** — that wave delivers 2027-28, inside the hold, and Carmel
  alone takes over half of 2026 metro deliveries immediately north of the asset.
- **Real estate taxes $137K rising to $160K.** Marion County parcel 8000002 paid **$129,648** in 2025.
  The 25-year record on the `Tax Record` tab shows assessed value range-bound at $5.1–6.6mm every year
  since 2012 and — decisively — assessed value *fell* the year after the March 2006 sale. Indiana values
  apartments at the lowest of the income, market and cost approaches, so assessed value does not track
  trade prices here. An earlier draft stepped taxes toward a share of the purchase price; correcting
  that is worth about 1.5 points of levered IRR. The risk is still priced in Ladder C.
- **Exit cap 6.75%**, held constant, against a 7.86% Year-5 yield on cost.

## Sensitivities — all live, all anchored at 14.0%

| Block | What it flexes |
|---|---|
| Grid 1A / 1B | Purchase price × exit cap → levered IRR and equity multiple |
| Grid 2 | Flat market rent growth (1.0%–3.5%) × exit cap |
| Ladder A | Blended renovation premium $100–$190 → IRR and return on cost |
| Ladder B | Blended renovation cost $11,000–$20,000 → IRR and return on cost |
| Ladder C | Year-5 tax bill $137K–$294K, i.e. assessed value at 41%–85% of the trade price |

Every cell drives its own cash-flow row and `IRR()` — nothing is pasted, so the grids move when you
change an assumption.

## Underwriting posture

Conservative on what we do not control (rent growth off actuals, exit cap held, taxes evidenced).
Confident on what we do control (renovation premium, pace, ancillary income, expense management). The
model solves for **the highest price that clears the hurdle**, not for assumptions that justify a price.

## Model tabs

`README` · `Assumptions` · `Rent Roll` · `Unit Mix` · `T-12 Recast` · `Other Income` ·
`Capex Budget` · `Reno Schedule` · `Annual Model` · `Returns` · `Sensitivities` · `Tax Record` ·
`DD Questions`

Month 1 is August 2026 (rent roll dated 07/01/2026). All capital is funded at closing in Sources &
Uses, so annual cash flow deducts interest and replacement reserves only and never double-counts capex.

## Data hierarchy

Per the case instructions the Excel rent roll and T-12 govern over the Offering Memorandum. The most
consequential conflict is the in-unit washer/dryer count: the OM says 78 of 118 units have machines but
then lists 53 side-by-side plus 1 stackable, and its own floor-plan and building-mix schedules foot to
54. The rent roll shows 53, so 53 units are Track A and **65 carry the $5,000 addition** prescribed by
the case. Every 10 units of error moves the capital budget by $50,000.

Status rules applied as instructed: `Notice-Unrented` and `Evict` count as occupied, any variation of
`Vacant` counts as vacant — 114 occupied, 4 vacant, 96.6%.

## Reproducing

```bash
pip install openpyxl && npm install pptxgenjs docx

python3 build_model.py      # writes the .xlsx
node    build_deck.js       # writes the .pptx
node    build_dd.js         # writes the .docx  (reads dd_questions.json)
python3 verify.py           # independent replica — ties to the workbook
```

`verify.py` re-implements the whole model in plain Python and is the tie-out that proves the workbook's
formulas compute what they should: NOI by year, the two-track schedule, loan sizing, the exit, both
IRRs and the price ladder all agree.

## Source files

1. `rent_roll20260701_1.xlsx` — rent roll as of 07/01/2026
2. `Chateau_Trailing_Profit_And_Loss_Detail_01312026.xlsx` — T-12 accrual P&L, Feb 2025 – Jan 2026
3. `Chateau_in_the_Woods_Offering_Memorandum.pdf` — CBRE / Midwest Multifamily Advisory Group, 09/10/2025
4. Yardi Matrix Indianapolis Multifamily Market Report, 1Q 2026 — rent growth, supply, occupancy
5. Marion County parcel 8000002 — 25-year assessment and tax record, transfer/deed record

Source files are not committed; the build scripts read the first three from the upload directory.
