"""
Independent replica of the Excel model. Used to verify the workbook and to source the deck.

Two-track renovation programme:
  Track A — 53 units that already have in-unit laundry. Light IN-PLACE scope (residents stay),
            so no vacancy loss, but the premium is only captured at the next lease renewal.
  Track B — 65 units that need laundry added. Requires plumbing and electrical permits and a
            vacant unit, so it runs at natural turnover with real downtime, but the premium is
            captured immediately on the new lease.
"""
UNITS = 118
MKT_MO = 153435.0
N_WD, N_NOWD = 53, 65

# ---- Track A: in-place, units that already have laundry ------------------
A_START, A_PACE, A_COST, A_PREM, A_DOWN, A_LAG = 2, 10, 10000, 100, 0, 6
# ---- Track B: vacant-unit renovation plus laundry addition ---------------
B_START, B_PACE, B_COST, B_PREM, B_DOWN, B_LAG = 4, 5, 13000 + 5000, 200, 14, 1

reno_cap = N_WD * A_COST + N_NOWD * B_COST
cost_blend = reno_cap / UNITS
prem_blend = (N_WD * A_PREM + N_NOWD * B_PREM) / UNITS
mkt_avg = MKT_MO / UNITS

# ---- monthly two-track schedule -----------------------------------------
rows = []
cumA = cumB = 0
for m in range(1, 61):
    dA = 0 if m < A_START else max(0, min(A_PACE, N_WD - cumA))
    dB = 0 if m < B_START else max(0, min(B_PACE, N_NOWD - cumB))
    cumA += dA
    cumB += dB
    rows.append(dict(m=m, y=(m - 1) // 12 + 1, dA=dA, dB=dB, eA=cumA, eB=cumB,
                     capex=dA * A_COST + dB * B_COST,
                     dvac=(dA * A_DOWN + dB * B_DOWN) / 30 / UNITS))
# premium-earning units: Track A lags A_LAG months, Track B lags B_LAG months
for i, r in enumerate(rows):
    jA, jB = i - A_LAG, i - B_LAG
    r["premA"] = rows[jA]["eA"] if jA >= 0 else 0
    r["premB"] = rows[jB]["eB"] if jB >= 0 else 0
    r["prem_units"] = r["premA"] + r["premB"]
    r["prem_dollars"] = r["premA"] * A_PREM + r["premB"] * B_PREM

Y = {y: [x for x in rows if x["y"] == y] for y in range(1, 6)}
avg_premD = {y: sum(x["prem_dollars"] for x in Y[y]) / 12 for y in Y}   # $/month of premium in force
avg_reno = {y: sum(x["prem_units"] for x in Y[y]) / 12 for y in Y}
capex_reno = {y: sum(x["capex"] for x in Y[y]) for y in Y}
dvac = {y: sum(x["dvac"] for x in Y[y]) / 12 for y in Y}
cum_y = {y: Y[y][-1]["eA"] + Y[y][-1]["eB"] for y in Y}
done_month = next(x["m"] for x in rows if x["eA"] + x["eB"] == UNITS)

# ---- market drivers ------------------------------------------------------
# Yardi Matrix Indianapolis 1Q-2026: T-3 asking rent FLAT at $1,310, +1.1% YoY,
# 2025 construction starts more than doubled (delivering 2027-28, inside the hold).
G = {1: .015, 2: .020, 3: .025, 4: .030, 5: .030, 6: .030}
LL = {1: .022, 2: .020, 3: .0175, 4: .015, 5: .015, 6: .015}
VAC = {y: .055 for y in range(1, 7)}
CONC = {1: .0075, 2: .0075, 3: .005, 4: .005, 5: .005, 6: .005}
BD = {1: .015, 2: .0125, 3: .0125, 4: .010, 5: .010, 6: .010}
idx, fac = {0: 1.0}, {}
for y in range(1, 7):
    idx[y] = idx[y - 1] * (1 + G[y])
    fac[y] = (idx[y - 1] * idx[y]) ** 0.5

# ---- other income --------------------------------------------------------
E_WS, E_WS_G = 90000, 0.04
G_OI = 0.03
ramp_tech = {1: .85, 2: 1., 3: 1., 4: 1., 5: 1.}
ramp_valet = {1: .75, 2: .95, 3: 1., 4: 1., 5: 1.}
ramp_pet = {1: .50, 2: 1., 3: 1., 4: 1., 5: 1.}
OCCF = 0.95
FLAT = {"carport": 16197.0, "app_admin": 7050 + 5250 + 2346.11 + 200,
        "late_nsf": 9675 + 500, "shortterm": 6440.71, "leasterm": 4000.0,
        "damage": 15000.0, "misc": 5000.0, "storage": 3600.0, "wd_rental": 0.0}
oi = {}
for y in range(1, 6):
    ws_exp = E_WS * (1 + E_WS_G) ** (y - 1)
    t = 0.85 * ws_exp
    t += 25 * 12 * UNITS * OCCF * ramp_valet[y] * (1 + G_OI) ** (y - 1)
    t += 49 * 12 * UNITS * OCCF * ramp_tech[y] * (1 + G_OI) ** (y - 1)
    t += 30 * 12 * UNITS * 0.25 * ramp_pet[y] * (1 + G_OI) ** (y - 1)
    t += sum(FLAT.values()) * (1 + G_OI) ** (y - 1)
    oi[y] = t
oi[6] = oi[5] * (1 + G_OI)

# ---- expenses ------------------------------------------------------------
EXP = [("pay", 235000, .035), ("mkt", 28000, .03), ("ga", 50000, .03), ("int", 39577, .03),
       ("elec", 31000, .03), ("gas", 16000, .03), ("ws", 90000, .04), ("rm", 255000, .03),
       ("ins", 95000, .05)]
# Marion County parcel 8000002: 2025 net annual tax $129,648 on a $5,214,932 net assessed
# value at 2.4861%. Assessed value has been range-bound $5.1-6.6mm since 2012 (0.35%/yr) and
# the last assessment change was 03/15/2022, so the 2006 sale did NOT reset it. Base case
# grows the actual bill at 4%; a reassessment shock is a named sensitivity, not the base.
TAX = {1: 137000, 2: 142500, 3: 148200, 4: 154100, 5: 160300, 6: 160300 * 1.05}
MGMT = 0.04

model = {}
for y in range(1, 7):
    if y == 6:
        prem_d = UNITS * prem_blend
        dv = 0
    else:
        prem_d = avg_premD[y]
        dv = dvac[y]
    gpr = (UNITS * mkt_avg + prem_d) * 12 * fac[y]
    nri = gpr * (1 - LL[y] - VAC[y] - dv - CONC[y] - BD[y])
    egr = nri + oi[y]
    opex = sum(v * (1 + g) ** (y - 1) for _, v, g in EXP) + TAX[y] + MGMT * egr
    model[y] = dict(gpr=gpr, nri=nri, oi=oi[y], egr=egr, opex=opex, noi=egr - opex)

# ---- capitalisation ------------------------------------------------------
PRICE = 11_500_000            # solved bid — see the price ladder below
ACQ, CONTING, LTC, RATE, FEE, FINCOST = .015, .08, .65, .0585, .01, 100_000
MINDY, MINDSCR = .080, 1.25
PROP_CAPEX = 360000 + 150000 + 40000 + 150000 + 60000 + 75000 + 50000 + 100000
cont = (reno_cap + PROP_CAPEX) * CONTING
EXITCAP, COS = .0675, .015
ACTUAL_TRADE = 13_625_000     # deed 06/12/2026, recorded 06/26/2026, assessor flag "Y"


def size_loan(cost, noi1):
    return min(cost * LTC, noi1 / MINDY, noi1 / (MINDSCR * RATE))


res = {y: 300 * UNITS * (1 + G_OI) ** (y - 1) for y in range(1, 6)}


def irr(cf, lo=-0.9, hi=2.0):
    for _ in range(500):
        mid = (lo + hi) / 2
        if sum(c / (1 + mid) ** i for i, c in enumerate(cf)) > 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def price_case(p, exitcap=EXITCAP):
    t = p + reno_cap + PROP_CAPEX + cont
    ln = size_loan(t, model[1]["noi"])
    eq = p + p * ACQ + reno_cap + PROP_CAPEX + cont + ln * FEE + FINCOST - ln
    net = model[6]["noi"] / exitcap * (1 - COS)
    ucf = [-(t + p * ACQ)] + [model[y]["noi"] - res[y] for y in range(1, 6)]
    ucf[-1] += net
    lcf = [-eq] + [model[y]["noi"] - res[y] - ln * RATE for y in range(1, 6)]
    lcf[-1] += net - ln
    return dict(price=p, tcost=t, loan=ln, equity=eq, net=net, ucf=ucf, lcf=lcf,
                irr=irr(lcf), em=sum(lcf[1:]) / -lcf[0], uirr=irr(ucf),
                uem=sum(ucf[1:]) / -ucf[0], cap=model[1]["noi"] / p,
                yoc=model[5]["noi"] / t,
                coc=[(model[y]["noi"] - res[y] - ln * RATE) / eq for y in range(1, 6)])


def max_bid(target_irr=.14, target_em=1.80, exitcap=EXITCAP):
    lo, hi = 6_000_000, 20_000_000
    for _ in range(50):
        m = (lo + hi) / 2
        r = price_case(m, exitcap)
        if r["irr"] >= target_irr and r["em"] >= target_em:
            lo = m
        else:
            hi = m
    return lo


if __name__ == "__main__":
    b = price_case(PRICE)
    print(f"{'':34}" + "".join(f"{'Y'+str(y):>12}" for y in range(1, 7)))
    for k, lbl in [("gpr", "Gross Potential Rent"), ("nri", "Net Rental Income"),
                   ("oi", "Other Income"), ("egr", "Effective Gross Revenue"),
                   ("opex", "Operating Expenses"), ("noi", "NOI")]:
        print(f"{lbl:34}" + "".join(f"{model[y][k]:>12,.0f}" for y in range(1, 7)))
    print(f"{'NOI / unit':34}" + "".join(f"{model[y]['noi']/UNITS:>12,.0f}" for y in range(1, 7)))
    print(f"{'Expense ratio':34}" + "".join(f"{model[y]['opex']/model[y]['egr']:>11.1%} " for y in range(1, 7)))
    print(f"{'Avg units earning premium':34}" + "".join(f"{avg_reno.get(y, UNITS):>12,.1f}" for y in range(1, 7)))
    print(f"{'Reno downtime vacancy':34}" + "".join(f"{dvac.get(y, 0):>11.2%} " for y in range(1, 7)))
    print()
    print(f"Two-track programme completes in month {done_month}")
    print(f"  Track A  {N_WD} units @ {A_PACE}/mo from month {A_START}, ${A_COST:,}/unit, "
          f"${A_PREM}/mo premium, in-place, {A_LAG}-month lag to capture")
    print(f"  Track B  {N_NOWD} units @ {B_PACE}/mo from month {B_START}, ${B_COST:,}/unit, "
          f"${B_PREM}/mo premium, vacant, {B_DOWN}-day downtime")
    print(f"  Blended  ${cost_blend:,.0f}/unit  ${prem_blend:.2f}/mo  "
          f"ROC {prem_blend*12/cost_blend:.1%}   total ${reno_cap:,.0f}")
    print(f"Non-unit capital ${PROP_CAPEX:,.0f} + contingency ${cont:,.0f}")
    print()
    print(f"{'Price':>12} {'/unit':>9} {'going-in':>9} {'YoC':>7} {'spread':>8} "
          f"{'equity':>11} {'IRR':>7} {'EM':>7} {'uIRR':>7}")
    for p in [10_500_000, 11_000_000, 11_500_000, 12_000_000, 12_500_000,
              13_000_000, ACTUAL_TRADE, 14_200_000]:
        r = price_case(p)
        tag = "  <- actual trade" if p == ACTUAL_TRADE else ""
        print(f"${p:>11,} ${p/UNITS:>8,.0f} {r['cap']:>9.2%} {r['yoc']:>7.2%} "
              f"{r['yoc']-EXITCAP:>+8.2%} ${r['equity']:>10,.0f} {r['irr']:>7.1%} "
              f"{r['em']:>6.2f}x {r['uirr']:>7.1%}{tag}")
    print()
    print(f"Max bid @ 14% IRR / 1.80x : ${max_bid():,.0f}  (${max_bid()/UNITS:,.0f}/unit)")
    print(f"Max bid @ 12% IRR / 1.60x : ${max_bid(.12,1.60):,.0f}")
    print(f"Actual trade ${ACTUAL_TRADE:,}  -> our max bid is "
          f"{max_bid()/ACTUAL_TRADE-1:+.1%} vs what it traded for")
