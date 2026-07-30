"""Independent replica of the Excel model, used to verify it and to source the deck."""
UNITS = 118
MKT_MO = 153435.0
N_WD, N_NOWD = 53, 65
COST_INT, COST_WD = 13000, 5000
PREM_INT, PREM_WD = 125, 75
RENO_START, PACE, DOWNTIME = 3, 5, 14

cost_blend = (N_WD * COST_INT + N_NOWD * (COST_INT + COST_WD)) / UNITS
prem_blend = (N_WD * PREM_INT + N_NOWD * (PREM_INT + PREM_WD)) / UNITS
reno_cap = cost_blend * UNITS
mkt_avg = MKT_MO / UNITS
reno_avg = mkt_avg + prem_blend

# monthly renovation schedule
cum, rows = 0, []
for m in range(1, 61):
    bom = cum
    done = 0 if m < RENO_START else max(0, min(PACE, UNITS - cum))
    cum += done
    rows.append(dict(m=m, y=(m - 1) // 12 + 1, done=done, eom=cum, bom=bom,
                     capex=done * COST_INT + done * N_NOWD / UNITS * COST_WD,
                     dvac=done * DOWNTIME / 30 / UNITS))
Y = {y: [x for x in rows if x["y"] == y] for y in range(1, 6)}
avg_reno = {y: sum(x["bom"] for x in Y[y]) / 12 for y in Y}
avg_cls = {y: UNITS - avg_reno[y] for y in Y}
capex_reno = {y: sum(x["capex"] for x in Y[y]) for y in Y}
dvac = {y: sum(x["dvac"] for x in Y[y]) / 12 for y in Y}
cum_y = {y: Y[y][-1]["eom"] for y in Y}

G = {y: 0.03 for y in range(1, 7)}
LL = {1: .022, 2: .020, 3: .0175, 4: .015, 5: .015, 6: .015}
VAC = {y: .05 for y in range(1, 7)}
CONC = {1: .0075, 2: .0075, 3: .005, 4: .005, 5: .005, 6: .005}
BD = {1: .015, 2: .0125, 3: .0125, 4: .010, 5: .010, 6: .010}
idx, fac = {0: 1.0}, {}
for y in range(1, 7):
    idx[y] = idx[y - 1] * (1 + G[y])
    fac[y] = (idx[y - 1] * idx[y]) ** 0.5

# ---- other income -------------------------------------------------------
E_WS, E_WS_G = 90000, 0.04
G_OI = 0.03
ramp_tech = {1: .85, 2: 1., 3: 1., 4: 1., 5: 1.}
ramp_valet = {1: .75, 2: .95, 3: 1., 4: 1., 5: 1.}
ramp_pet = {1: .50, 2: 1., 3: 1., 4: 1., 5: 1.}
OCCF = 0.95
FLAT = {  # T-12 based lines held flat in Y1 then grown
    "carport": 16197.0,
    "app_admin": 7050 + 5250 + 2346.11 + 200,
    "late_nsf": 9675 + 500,
    "shortterm": 6440.71,
    "leasterm": 4000.0,
    "damage": 15000.0,
    "misc": 5000.0,
    "storage": 3600.0,
    "wd_rental": 0.0,
}
oi = {}
for y in range(1, 6):
    ws_exp = E_WS * (1 + E_WS_G) ** (y - 1)
    t = 0.85 * ws_exp
    rt = ramp_tech.get(y, 1.0); rv = ramp_valet.get(y, 1.0); rp = ramp_pet.get(y, 1.0)
    t += 25 * 12 * UNITS * OCCF * rv * (1 + G_OI) ** (y - 1)
    t += 49 * 12 * UNITS * OCCF * rt * (1 + G_OI) ** (y - 1)
    t += 30 * 12 * UNITS * 0.25 * rp * (1 + G_OI) ** (y - 1)
    t += sum(FLAT.values()) * (1 + G_OI) ** (y - 1)
    oi[y] = t
# Year 6 follows the workbook convention: grow the Year-5 total by other-income growth
oi[6] = oi[5] * (1 + G_OI)

# ---- expenses -----------------------------------------------------------
EXP = [("pay", 235000, .035), ("mkt", 28000, .03), ("ga", 50000, .03), ("int", 39577, .03),
       ("elec", 31000, .03), ("gas", 16000, .03), ("ws", 90000, .04), ("rm", 255000, .03),
       ("ins", 95000, .05)]
TAX = {1: 140000, 2: 170000, 3: 190000, 4: 200000, 5: 210000, 6: 210000 * 1.05}
MGMT = 0.04

model = {}
for y in range(1, 7):
    ar = UNITS if y == 6 else avg_reno[y]
    ac = 0 if y == 6 else avg_cls[y]
    gpr = (ac * mkt_avg + ar * reno_avg) * 12 * fac[y]
    dv = 0 if y == 6 else dvac[y]
    nri = gpr * (1 - LL[y] - VAC[y] - dv - CONC[y] - BD[y])
    egr = nri + oi[y]
    opex = sum(v * (1 + g) ** (y - 1) for _, v, g in EXP) + TAX[y] + MGMT * egr
    noi = egr - opex
    model[y] = dict(gpr=gpr, nri=nri, oi=oi[y], egr=egr, opex=opex, noi=noi)

# ---- capitalisation -----------------------------------------------------
PRICE = 14_200_000
ACQ, CONTING, LTC, RATE, FEE, FINCOST = .015, .08, .65, .0585, .01, 100_000
MINDY, MINDSCR = .080, 1.25
PROP_CAPEX = 360000 + 150000 + 40000 + 150000 + 60000 + 75000 + 50000 + 100000
cont = (reno_cap + PROP_CAPEX) * CONTING
tcost = PRICE + reno_cap + PROP_CAPEX + cont
def size_loan(cost, noi1):
    return min(cost * LTC, noi1 / MINDY, noi1 / (MINDSCR * RATE))


loan = size_loan(tcost, 0)  # placeholder, reset below
loan = size_loan(tcost, model[1]["noi"])
uses = PRICE + PRICE * ACQ + reno_cap + PROP_CAPEX + cont + loan * FEE + FINCOST
equity = uses - loan
interest = loan * RATE
res = {y: 300 * UNITS * (1 + G_OI) ** (y - 1) for y in range(1, 6)}

EXITCAP, COS = .0675, .015
gross = model[6]["noi"] / EXITCAP
net = gross * (1 - COS)
toeq = net - loan

ucf = [-(tcost + PRICE * ACQ)] + [model[y]["noi"] - res[y] for y in range(1, 6)]
ucf[-1] += net
lcf = [-equity] + [model[y]["noi"] - res[y] - interest for y in range(1, 6)]
lcf[-1] += toeq


def irr(cf, lo=-0.5, hi=1.0):
    for _ in range(300):
        mid = (lo + hi) / 2
        v = sum(c / (1 + mid) ** i for i, c in enumerate(cf))
        if v > 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


print(f"{'':34}{'Y1':>12}{'Y2':>12}{'Y3':>12}{'Y4':>12}{'Y5':>12}{'Y6':>12}")
for k, lbl in [("gpr", "Gross Potential Rent"), ("nri", "Net Rental Income"), ("oi", "Other Income"),
               ("egr", "Effective Gross Revenue"), ("opex", "Operating Expenses"), ("noi", "NOI")]:
    print(f"{lbl:34}" + "".join(f"{model[y][k]:>12,.0f}" for y in range(1, 7)))
print(f"{'NOI / unit':34}" + "".join(f"{model[y]['noi']/UNITS:>12,.0f}" for y in range(1, 7)))
print(f"{'Expense ratio':34}" + "".join(f"{model[y]['opex']/model[y]['egr']:>11.1%} " for y in range(1, 7)))
print(f"{'Avg renovated units':34}" + "".join(f"{(UNITS if y==6 else avg_reno[y]):>12,.1f}" for y in range(1, 7)))
print(f"{'Cum renovated (YE)':34}" + "".join(f"{(UNITS if y==6 else cum_y[y]):>12,.0f}" for y in range(1, 6)))
print()
print(f"Blended reno cost/unit      ${cost_blend:,.0f}   premium ${prem_blend:,.2f}/mo   "
      f"ROC {prem_blend*12/cost_blend:.1%}")
print(f"Renovation capital          ${reno_cap:,.0f}")
print(f"Non-unit capital            ${PROP_CAPEX:,.0f}   contingency ${cont:,.0f}")
print(f"Total capital               ${reno_cap+PROP_CAPEX+cont:,.0f}  (${(reno_cap+PROP_CAPEX+cont)/UNITS:,.0f}/unit)")
print(f"Total capitalised cost      ${tcost:,.0f}  (${tcost/UNITS:,.0f}/unit)")
print(f"Total uses                  ${uses:,.0f}")
print(f"Loan sizing tests: LTC ${tcost*LTC:,.0f} | DY ${model[1]['noi']/MINDY:,.0f} | "
      f"DSCR ${model[1]['noi']/(MINDSCR*RATE):,.0f}")
print(f"Loan                        ${loan:,.0f} ({loan/tcost:.1%} LTC)  interest ${interest:,.0f}   "
      f"Y1 DY {model[1]['noi']/loan:.1%}  Y1 DSCR {model[1]['noi']/interest:.2f}x")
print(f"Sponsor equity              ${equity:,.0f}")
print(f"Going-in cap (our UW)       {model[1]['noi']/PRICE:.2%}   price/unit ${PRICE/UNITS:,.0f}")
print(f"Y5 yield on total cost      {model[5]['noi']/tcost:.2%}")
print(f"Exit gross @ {EXITCAP:.2%}         ${gross:,.0f}  (${gross/UNITS:,.0f}/unit)  net ${net:,.0f}")
print(f"Proceeds to equity          ${toeq:,.0f}")
print()
print("Unlevered CF ", [f"{c:,.0f}" for c in ucf])
print("Levered CF   ", [f"{c:,.0f}" for c in lcf])
print(f"Unlevered IRR {irr(ucf):.2%}   EM {sum(ucf[1:])/-ucf[0]:.2f}x")
print(f"Levered   IRR {irr(lcf):.2%}   EM {sum(lcf[1:])/-lcf[0]:.2f}x")
print("CoC       ", [f"{(model[y]['noi']-res[y]-interest)/equity:.1%}" for y in range(1, 6)])
print(f"Avg CoC   {sum((model[y]['noi']-res[y]-interest) for y in range(1,6))/5/equity:.1%}")
print()
print("PRICE SENSITIVITY at 6.75% exit cap")
for p in [11_500_000, 12_000_000, 12_500_000, 13_000_000, 13_500_000, 14_200_000, 14_750_000]:
    t = p + reno_cap + PROP_CAPEX + cont
    ln = size_loan(t, model[1]["noi"])
    eq = p + p * ACQ + reno_cap + PROP_CAPEX + cont + ln * FEE + FINCOST - ln
    cf = [-eq] + [model[y]["noi"] - res[y] - ln * RATE for y in range(1, 6)]
    cf[-1] += net - ln
    print(f"  ${p:>11,}  (${p/UNITS:>7,.0f}/unit)  equity ${eq:>10,.0f}  "
          f"IRR {irr(cf):>6.1%}  EM {sum(cf[1:])/-cf[0]:.2f}x  "
          f"going-in {model[1]['noi']/p:.2%}")
print()
print("EXIT CAP SENSITIVITY at $14.2mm")
for c in [.060, .065, .0675, .070, .075]:
    g_ = model[6]["noi"] / c
    n_ = g_ * (1 - COS)
    cf = [-equity] + [model[y]["noi"] - res[y] - interest for y in range(1, 6)]
    cf[-1] += n_ - loan
    print(f"  {c:.2%}  value ${g_:>11,.0f}  IRR {irr(cf):>6.1%}  EM {sum(cf[1:])/-cf[0]:.2f}x")
