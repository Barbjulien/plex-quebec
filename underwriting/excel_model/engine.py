"""Python port of the Plex Québec underwriting engine (index.html JS, 2026-11-22):
multi-term DCF (renouvellement de terme + retrait d'équité) and turnover-phased
renovations (reno spend follows the unit-turnover curve; private bridge drawn
pro-rata, interest on the drawn balance).
Used as the reference oracle to verify the Excel workbook's formulas."""
import math

def monthly_rate(annual_pct):
    return (1 + annual_pct / 2 / 100) ** (1 / 6) - 1

def mortgage_constant(annual_pct, amort_years):
    r = monthly_rate(annual_pct); n = amort_years * 12
    if r == 0: return 1 / n * 12
    return (r * (1 + r) ** n / ((1 + r) ** n - 1)) * 12

def loan_balance(principal, annual_pct, amort_years, months_paid):
    r = monthly_rate(annual_pct); n = amort_years * 12
    if r == 0: return max(0, principal - (principal / n) * months_paid)
    pmt = principal * r * (1 + r) ** n / ((1 + r) ** n - 1)
    return max(0, principal * (1 + r) ** months_paid - pmt * (((1 + r) ** months_paid - 1) / r))

def break_penalty(balance, contract_pct, current_pct, years_into_term, term_years):
    remaining = max(0, term_years - years_into_term)
    if remaining <= 0: return 0
    r = monthly_rate(contract_pct)
    three_month = balance * r * 3
    ird = balance * max(0, (contract_pct - current_pct) / 100) * remaining
    return max(three_month, ird)

def irr(cashflows):
    def npv(r): return sum(cf / (1 + r) ** t for t, cf in enumerate(cashflows))
    lo, hi = -0.9, 2.0
    n_lo, n_hi = npv(lo), npv(hi)
    if n_lo * n_hi > 0: return None
    for _ in range(64):
        mid = (lo + hi) / 2
        if npv(mid) * n_lo < 0: hi = mid
        else: lo = mid
    return (lo + hi) / 2

# CMHC tables
def get_max_amort(program, pts):
    if program == 'mli-market': return 40
    if pts >= 100: return 50
    if pts >= 70: return 45
    if pts >= 50: return 40
    return 25

def get_program_max_ltv(program, pts):
    if program == 'mli-market': return 85
    if pts >= 100: return 95
    if pts >= 70: return 90
    return 85

def get_min_dscr(program, building_type, units, term, is_refi):
    if building_type == 'other': return 1.2
    if program == 'mli-select': return 1.1
    if units <= 6: return 1.2 if is_refi else 1.1
    return 1.2 if term >= 10 else 1.3

def get_mli_discount(program, pts):
    if program != 'mli-select': return 0
    if pts >= 100: return 0.3
    if pts >= 70: return 0.2
    if pts >= 50: return 0.1
    return 0

PREMIUM_TABLES = {
    'standard': {
        'purchase':     [(65,.026),(70,.0285),(75,.0335),(80,.0435),(85,.0535),(90,.059),(100,.0615)],
        'construction': [(65,.0325),(70,.0375),(75,.0425),(80,.050),(85,.060),(90,.0675),(100,.070)],
    },
    'other': {
        'purchase':     [(65,.063),(70,.066),(75,.069),(80,.0705),(85,.0775),(90,.080),(100,.0875)],
        'construction': [(65,.0655),(70,.0685),(75,.0715),(80,.073),(85,.080),(90,.0825),(100,.090)],
    },
}
def get_base_premium_rate(ltv_pct, building_type, purpose):
    sched = PREMIUM_TABLES.get(building_type, PREMIUM_TABLES['standard'])['construction' if purpose == 'construction' else 'purchase']
    for max_ltv, rate in sched:
        if ltv_pct <= max_ltv: return rate
    return sched[-1][1]

def get_amort_surcharge(amort):
    return max(0, math.floor((amort - 25) / 5)) * 0.0025

def get_prov_tax(province):
    return {'qc': 0.09975, 'on': 0.08, 'mb': 0.07, 'sk': 0.06}.get(province, 0)

def compute_cmhc(program, building_type, purpose, mli_pts, target_ltv_pct, rate_pct, term, amort,
                 province, property_value, noi, units, lending_value, is_refi):
    lv = lending_value if (lending_value is not None and lending_value > 0 and lending_value < property_value) else property_value
    eff_amort = min(amort, get_max_amort(program, mli_pts))
    prog_max_ltv = get_program_max_ltv(program, mli_pts)
    eff_ltv_cap = min(target_ltv_pct, prog_max_ltv) / 100
    min_dscr = get_min_dscr(program, building_type, units, term, is_refi)
    mc = mortgage_constant(rate_pct, eff_amort)
    ltv_max_loan = lv * eff_ltv_cap
    surcharge = get_amort_surcharge(eff_amort)
    discount = get_mli_discount(program, mli_pts)
    dscr_max_loan = noi / (min_dscr * mc) if mc > 0 else 0
    approved = min(ltv_max_loan, max(0, dscr_max_loan))
    binding = 'DSCR' if dscr_max_loan <= ltv_max_loan else 'LTV'
    eff_ltv = approved / property_value if property_value > 0 else 0
    base_rate = get_base_premium_rate(eff_ltv * 100, building_type, purpose)
    net_rate = max(0, (base_rate + surcharge) * (1 - discount))
    premium = approved * net_rate
    app_fee = units * 150
    insured = approved + premium + app_fee
    r = monthly_rate(rate_pct); n = eff_amort * 12
    mp = insured / n if r == 0 else insured * r * (1 + r) ** n / ((1 + r) ** n - 1)
    ads = mp * 12
    dscr_insured = noi / ads if ads > 0 else 0
    prov_tax = premium * get_prov_tax(province)
    down = property_value - approved
    dscr_base = noi / (approved * mc) if (mc > 0 and approved > 0) else 0
    ltv_implied_dscr = noi / (ltv_max_loan * mc) if (mc > 0 and ltv_max_loan > 0) else 0
    dscr_implied_ltv = dscr_max_loan / lv if lv > 0 else 0
    return dict(mc=mc, ltvMaxLoan=ltv_max_loan, dscrMaxLoan=dscr_max_loan, approvedLoan=approved,
                binding=binding, effLtv=eff_ltv, dscrBase=dscr_base, minDscr=min_dscr,
                ltvImpliedDscr=ltv_implied_dscr, dscrImpliedLtv=dscr_implied_ltv, baseRate=base_rate,
                surcharge=surcharge, discount=discount, netRate=net_rate, premium=premium,
                insured=insured, mp=mp, ads=ads, dscrInsured=dscr_insured, provTax=prov_tax,
                downPayment=down, appFee=app_fee, effAmort=eff_amort, lendingValue=lv)

def compute_conventional(target_ltv_pct, rate_pct, amort, min_dscr, property_value, noi, lending_value):
    lv = lending_value if (lending_value is not None and lending_value > 0 and lending_value < property_value) else property_value
    mc = mortgage_constant(rate_pct, amort)
    ltv_max_loan = lv * (target_ltv_pct / 100)
    dscr_max_loan = noi / (min_dscr * mc) if mc > 0 else 0
    approved = min(ltv_max_loan, max(0, dscr_max_loan))
    binding = 'DSCR' if dscr_max_loan <= ltv_max_loan else 'LTV'
    eff_ltv = approved / property_value if property_value > 0 else 0
    r = monthly_rate(rate_pct); n = amort * 12
    mp = approved / n if r == 0 else approved * r * (1 + r) ** n / ((1 + r) ** n - 1)
    ads = mp * 12
    return dict(mc=mc, ltvMaxLoan=ltv_max_loan, dscrMaxLoan=dscr_max_loan, approvedLoan=approved,
                binding=binding, effLtv=eff_ltv, baseRate=0, surcharge=0, discount=0, netRate=0,
                premium=0, insured=approved, mp=mp, ads=ads,
                dscrInsured=noi / ads if ads > 0 else 0,
                dscrBase=noi / (approved * mc) if (mc > 0 and approved > 0) else 0,
                minDscr=min_dscr,
                ltvImpliedDscr=noi / (ltv_max_loan * mc) if (mc > 0 and ltv_max_loan > 0) else 0,
                dscrImpliedLtv=dscr_max_loan / lv if lv > 0 else 0,
                provTax=0, appFee=0, effAmort=amort, downPayment=property_value - approved, lendingValue=lv)

# ─── Acquisition defaults (resetAll + HTML selects) ───────────────────────────
ACQ_DEFAULTS = dict(
    purchasePrice=1175000, units=6, avgSqft=1000, rollValue=969400, acqCostPct=2.5, brokerFeePct=1,
    rentMode='avg', currentRent=1000, currentRentAnnual=72000, vacPct=5, otherInc=0, opexPct=30,
    opexBasis='egi',
    targetRent=1450, rentGrowth=2, indexGrowth=2, holding=5, optVacPct=5, optOpexPct=30, turnover=17.5,
    ccaOn=False, ccaRate=4, bldgPct=80,
    renoCost=20000, pctReno=50, cfkMonths=0, pctCfk=20,
    acqFinType='cmhc', program='mli-market', mliPts=50, buildingType='standard', purpose='purchase',
    targetLtv=85, rate=4.25, term=5, amort=40, province='qc',
    schlCapRate=0, schlOpexPct=None, schlVacancy=None,
    acqConvLtv=75, acqConvDscr=1.20,
    privPct=100, privRate=14, contingency=50,
    renewalRate=4.25, takeout=False,
    capDrift=0.25, newRate=4.25, margTax=26.5, cgTaxRate=33, commPct=2, closing=1500,
    refiFinType='cmhc', refiProg='mli-market', refiMliPts=50, refiCapRatePct=4.75, refiLtv=75,
    refiRate=4.25, refiTerm=5, refiAmort=40, refiClosing=3000, refiBrokerPct=1,
    refiConvLtv=75, refiConvDscr=1.20,
)

def build_projection(p):
    """Port of buildProjection() (index.html / projection.ts). Loan modeled in
    phases (renewal at every term boundary; optional Section-K take-out at the
    first boundary). Reno spend phased with unit turnover, normalized to
    complete by exit; private bridge drawn pro-rata, interest on drawn balance."""
    n = p['holdingIntR']
    phase = dict(principal=p['insured'], rate=p['rate'], amortYears=p['effAmort'],
                 startYear=0, ads=p['ads'], termYears=p['term'])
    rows = []
    takeout_info = None
    weightedRev = totalDS = totalInterest = principalRepaidTotal = privIntTotal = 0.0

    def frac_at(y):
        if y <= 0: return 0.0
        return 1 - (1 - p['tr']) ** y if p['tr'] > 0 else 1.0
    fracN = frac_at(n)
    def reno_at(y):
        return p['totalReno'] * (frac_at(y) - frac_at(y - 1)) / fracN if fracN > 0 else 0.0

    priv_bal = p['cfk'] * p['privFrac']
    reno_equity_spread = 0.0
    # DPA (CCA): pool dégressif (part bâtiment + rénos capitalisées), plafonnée au revenu
    cca_pool = p.get('ccaPool0', 0.0) or 0.0
    cca_total = tax_annual_total = 0.0
    for y in range(1, n + 1):
        # Deux pistes: loyers cibles au taux MARCHÉ, loyers en place à l'INDEXATION
        growthM = (1 + p['rentGrowth'] / 100) ** y
        growthI = (1 + p['indexGrowth'] / 100) ** y
        frac = frac_at(y)
        griY = p['targetRent'] * growthM * p['units'] * 12
        egiY = griY * (1 - p['optVacPct'] / 100) + p['otherInc']
        noiY = egiY - (griY if p['opexBasis'] == 'rbp' else egiY) * p['optOpexPct'] / 100
        noiCurrY = p['noi'] * growthI
        blendedNOI = noiY * frac + noiCurrY * (1 - frac)
        balPrev = loan_balance(phase['principal'], phase['rate'], phase['amortYears'],
                               (y - 1 - phase['startYear']) * 12)
        balEnd = loan_balance(phase['principal'], phase['rate'], phase['amortYears'],
                              (y - phase['startYear']) * 12)
        principalY = balPrev - balEnd
        interestY = phase['ads'] - principalY
        renoY = reno_at(y)
        drawY = 0.0 if takeout_info else renoY * p['privFrac']
        privIntY = 0.0 if (takeout_info or p['holding'] <= 0) else (priv_bal + drawY / 2) * p['privRate'] / 100
        priv_bal += drawY
        renoEquityY = renoY - drawY
        reno_equity_spread += renoEquityY
        taxableY = blendedNOI - interestY - privIntY
        ccaY = taxY = 0.0
        if p.get('ccaOn'):
            cca_pool += renoY
            ccaY = min(cca_pool * p['ccaRate'] / 100, max(0.0, taxableY))
            cca_pool -= ccaY
            cca_total += ccaY
            taxY = max(0.0, taxableY - ccaY) * p['margTax'] / 100
            tax_annual_total += taxY
        cf = blendedNOI - phase['ads'] - p['annualReserve'] - privIntY - renoEquityY
        row = dict(y=y, frac=frac, blendedNOI=blendedNOI, ads=phase['ads'], adsRate=phase['rate'],
                   phaseStartY=phase['startYear'], phaseTermY=phase['termYears'],
                   balEnd=balEnd, interestY=interestY, principalY=principalY, privIntY=privIntY,
                   renoY=renoY, renoEquityY=renoEquityY, taxableY=taxableY, ccaY=ccaY, taxY=taxY,
                   cf=cf, event=None, cashOut=0.0)
        weightedRev += blendedNOI; totalDS += phase['ads']; totalInterest += interestY
        principalRepaidTotal += principalY; privIntTotal += privIntY
        if y == phase['startYear'] + phase['termYears'] and y < n:
            if p['takeout'] and not takeout_info:
                refiVal = blendedNOI / (p['refiCapRatePct'] / 100) if p['refiCapRatePct'] > 0 else 0
                if p['refiFinType'] == 'cmhc':
                    refi = compute_cmhc(p['refiProg'], p['buildingType'], p['purpose'], p['refiMliPts'],
                                        p['refiLtv'], p['refiRate'], p['refiTerm'], p['refiAmort'],
                                        p['province'], refiVal, blendedNOI, p['units'], None, True)
                else:
                    refi = compute_conventional(p['refiConvLtv'], p['refiRate'], p['refiAmort'],
                                                p['refiConvDscr'], refiVal, blendedNOI, 0)
                refiBrokerFee = refi['approvedLoan'] * p['refiBrokerPct'] / 100
                netCashOut = (refi['approvedLoan'] - balEnd - p['refiClosing'] - refi['provTax']
                              - refiBrokerFee - priv_bal)
                takeout_info = dict(year=y, refi=refi, refiVal=refiVal, netCashOut=netCashOut,
                                    refiBrokerFee=refiBrokerFee, balAtTakeout=balEnd,
                                    noiAtTakeout=blendedNOI, privRepaid=priv_bal,
                                    refiPhaseFees=refi['premium'] + refi['provTax'] + refi['appFee']
                                    + refiBrokerFee + p['refiClosing'])
                priv_bal = 0.0
                row['event'] = 'takeout'; row['cashOut'] = netCashOut
                phase = dict(principal=refi['insured'], rate=p['refiRate'], amortYears=refi['effAmort'],
                             startYear=y, ads=refi['ads'], termYears=p['refiTerm'])
            else:
                remAmort = max(1, phase['amortYears'] - phase['termYears'])
                phase = dict(principal=balEnd, rate=p['renewalRate'], amortYears=remAmort, startYear=y,
                             ads=balEnd * mortgage_constant(p['renewalRate'], remAmort),
                             termYears=p['term'])
                row['event'] = 'renewal'
        row['privBalY'] = priv_bal
        rows.append(row)
    balanceAtExit = loan_balance(phase['principal'], phase['rate'], phase['amortYears'],
                                 (n - phase['startYear']) * 12)
    return dict(rows=rows, finalPhase=phase, balanceAtExit=balanceAtExit, weightedRev=weightedRev,
                totalDS=totalDS, totalInterest=totalInterest, principalRepaidTotal=principalRepaidTotal,
                privIntTotal=privIntTotal, privBalExit=priv_bal, renoEquitySpread=reno_equity_spread,
                ccaTotal=cca_total, taxAnnualTotal=tax_annual_total,
                takeoutInfo=takeout_info, holdingIntR=n)

def recalc(i):
    """Full acquisition-mode recalc. i = dict of inputs (see ACQ_DEFAULTS)."""
    o = {}
    holdingIntR = min(round(i['holding'] if i['holding'] > 0 else 1), 30)
    # A
    acqCostDollar = i['purchasePrice'] * i['acqCostPct'] / 100
    totalAcqCost = i['purchasePrice'] + acqCostDollar
    pricePerUnit = i['purchasePrice'] / i['units'] if i['units'] > 0 else 0
    premiumOverRoll = i['purchasePrice'] / i['rollValue'] - 1 if i['rollValue'] > 0 else 0
    # B
    if i['rentMode'] == 'annual':
        gri = i['currentRentAnnual']
        currentRent = gri / i['units'] / 12 if i['units'] > 0 else 0
    else:
        currentRent = i['currentRent']
        gri = currentRent * i['units'] * 12
    egi = gri * (1 - i['vacPct'] / 100) + i['otherInc']
    opex = (gri if i['opexBasis'] == 'rbp' else egi) * i['opexPct'] / 100
    noi = egi - opex
    capRate = noi / i['purchasePrice'] if i['purchasePrice'] > 0 else 0
    grm = i['purchasePrice'] / gri if gri > 0 else 0
    # D
    totalReno = i['renoCost'] * i['units'] * i['pctReno'] / 100
    cfk = i['cfkMonths'] * currentRent * i['units'] * i['pctCfk'] / 100
    totalInvest = totalReno + cfk
    # C
    def noi_at_rent(rent):
        g = rent * i['units'] * 12
        e = g * (1 - i['optVacPct'] / 100) + i['otherInc']
        return e - (g if i['opexBasis'] == 'rbp' else e) * i['optOpexPct'] / 100
    projRentExit = i['targetRent'] * (1 + i['rentGrowth'] / 100) ** i['holding']
    optGRI = projRentExit * i['units'] * 12
    optEGI = optGRI * (1 - i['optVacPct'] / 100) + i['otherInc']
    optOpex = (optGRI if i['opexBasis'] == 'rbp' else optEGI) * i['optOpexPct'] / 100
    optNOI = optEGI - optOpex
    # E — SCHL normalized NOI & economic value
    schlOpex = i['schlOpexPct'] if i['schlOpexPct'] is not None else i['opexPct']
    schlVac = i['schlVacancy'] if i['schlVacancy'] is not None else i['vacPct']
    schlEGI = gri * (1 - schlVac / 100) + i['otherInc']
    schlNOI = schlEGI - (gri if i['opexBasis'] == 'rbp' else schlEGI) * schlOpex / 100
    valeurEconomique = schlNOI / (i['schlCapRate'] / 100) if i['schlCapRate'] > 0 else i['purchasePrice']
    if i['acqFinType'] == 'cmhc':
        cmhc = compute_cmhc(i['program'], i['buildingType'], i['purpose'], i['mliPts'], i['targetLtv'],
                            i['rate'], i['term'], i['amort'], i['province'], i['purchasePrice'],
                            schlNOI, i['units'], valeurEconomique, False)
    else:
        cmhc = compute_conventional(i['acqConvLtv'], i['rate'], i['amort'], i['acqConvDscr'],
                                    i['purchasePrice'], schlNOI, valeurEconomique)
    # F
    renoEquity = totalInvest * (1 - i['privPct'] / 100)
    privAmt = totalInvest * i['privPct'] / 100
    privFrac = i['privPct'] / 100
    # G
    totalReserve = i['contingency'] * i['units'] * i['holding'] * 12
    annualReserve = totalReserve / i['holding'] if i['holding'] > 0 else 0
    # F2 — renewal / take-out
    renewalRate = i['renewalRate'] if i.get('renewalRate') else i['rate']
    hasRenewal = i['holding'] > i['term']
    takeout = bool(hasRenewal and i.get('takeout'))
    # Unified projection (phases, renewal/take-out, turnover-phased reno)
    tr = i['turnover'] / 100
    proj = build_projection(dict(
        holdingIntR=holdingIntR, holding=i['holding'], term=i['term'], rate=i['rate'],
        renewalRate=renewalRate, insured=cmhc['insured'], effAmort=cmhc['effAmort'], ads=cmhc['ads'],
        noi=noi, targetRent=i['targetRent'], rentGrowth=i['rentGrowth'],
        indexGrowth=i['indexGrowth'],
        ccaOn=i.get('ccaOn', False), ccaRate=i.get('ccaRate', 4),
        ccaPool0=i['purchasePrice'] * i.get('bldgPct', 80) / 100, margTax=i['margTax'],
        tr=tr, units=i['units'],
        otherInc=i['otherInc'], optVacPct=i['optVacPct'], optOpexPct=i['optOpexPct'],
        opexBasis=i['opexBasis'], annualReserve=annualReserve,
        totalReno=totalReno, cfk=cfk, privFrac=privFrac, privRate=i['privRate'],
        takeout=takeout, refiFinType=i['refiFinType'], refiProg=i['refiProg'],
        buildingType=i['buildingType'], purpose=i['purpose'], refiMliPts=i['refiMliPts'],
        refiLtv=i['refiLtv'], refiRate=i['refiRate'], refiTerm=i['refiTerm'],
        refiAmort=i['refiAmort'], province=i['province'], refiConvLtv=i['refiConvLtv'],
        refiConvDscr=i['refiConvDscr'], refiCapRatePct=i['refiCapRatePct'],
        refiClosing=i['refiClosing'], refiBrokerPct=i['refiBrokerPct']))
    privCost = proj['privIntTotal']
    privBalExit = proj['privBalExit']
    renoEquitySpread = proj['renoEquitySpread']
    # H — totalCash = équité engagée; cash à la fermeture exclut l'équité réno (étalée)
    brokerFee = cmhc['approvedLoan'] * i['brokerFeePct'] / 100
    totalCash = cmhc['downPayment'] + renoEquity + acqCostDollar + cmhc['provTax'] + brokerFee
    cfkEquityClose = cfk * (1 - privFrac)
    cashAtClose = cmhc['downPayment'] + acqCostDollar + cmhc['provTax'] + brokerFee + cfkEquityClose
    allInCost = totalAcqCost + totalInvest + brokerFee + cmhc['provTax'] + cmhc['appFee']
    yoc = optNOI / allInCost if allInCost > 0 else 0
    # I — sale
    fracExit = 1 - (1 - tr) ** i['holding'] if tr > 0 else 1
    noiExit = noi * (1 + i['indexGrowth'] / 100) ** i['holding']   # loyers en place indexés (TAL)
    exitNOI = optNOI * fracExit + noiExit * (1 - fracExit)
    exitCap = capRate + i['capDrift'] / 100 * i['holding']
    salePrice = exitNOI / exitCap if exitCap > 0 else 0
    balanceAtExit = proj['balanceAtExit']
    fp = proj['finalPhase']
    penalty = break_penalty(balanceAtExit, fp['rate'], i['newRate'],
                            holdingIntR - fp['startYear'], fp['termYears'])
    salesComm = salePrice * i['commPct'] / 100
    weightedRev = proj['weightedRev']
    totalDS = proj['totalDS']
    principalRepaid = proj['principalRepaidTotal']
    totalInterest = proj['totalInterest']
    grossProfit = salePrice - totalAcqCost - totalInvest
    cgTaxable = max(0, grossProfit - salesComm - i['closing'])
    cgTax = cgTaxable * i['cgTaxRate'] / 100
    rentalTaxable = max(0, weightedRev - totalInterest - privCost)
    ccaOn = bool(i.get('ccaOn'))
    if ccaOn:
        # Mode DPA: impôt locatif annuel (net de la DPA) + récupération à la sortie
        rentalTax = proj['taxAnnualTotal']
        recaptureTax = proj['ccaTotal'] * i['margTax'] / 100
    else:
        rentalTax = rentalTaxable * i['margTax'] / 100
        recaptureTax = 0.0
    rentalTaxTotal = rentalTax + recaptureTax
    netOpProfit = weightedRev - totalDS - privCost - totalReserve
    finFees = (cmhc['premium'] + cmhc['provTax'] + cmhc['appFee'] + brokerFee
               + (proj['takeoutInfo']['refiPhaseFees'] if proj['takeoutInfo'] else 0))
    netPreTax = grossProfit + netOpProfit + principalRepaid - finFees - penalty
    netPostTax = netPreTax - cgTax - salesComm - i['closing'] - rentalTaxTotal
    totalRet = netPostTax / totalCash if totalCash > 0 else 0
    annRet = ((1 + totalRet) ** (1 / i['holding']) - 1 if totalRet > -1 else -1) if i['holding'] > 0 else 0
    exitMult = (totalCash + netPostTax) / totalCash if totalCash > 0 else 0
    # K — refi (garder, à la sortie)
    refiPropVal = exitNOI / (i['refiCapRatePct'] / 100) if i['refiCapRatePct'] > 0 else 0
    if i['refiFinType'] == 'cmhc':
        refi = compute_cmhc(i['refiProg'], i['buildingType'], i['purpose'], i['refiMliPts'], i['refiLtv'],
                            i['refiRate'], i['refiTerm'], i['refiAmort'], i['province'], refiPropVal,
                            exitNOI, i['units'], None, True)
    else:
        refi = compute_conventional(i['refiConvLtv'], i['refiRate'], i['refiAmort'], i['refiConvDscr'],
                                    refiPropVal, exitNOI, 0)
    grossCashOut = refi['approvedLoan'] - balanceAtExit - penalty
    refiBrokerFee = refi['approvedLoan'] * i['refiBrokerPct'] / 100
    netCashOut = grossCashOut - i['refiClosing'] - refi['provTax'] - refiBrokerFee - privBalExit
    cashLeftIn = totalCash - netCashOut
    refiCF = exitNOI - refi['ads'] - annualReserve
    refiCoC = refiCF / cashLeftIn if cashLeftIn > 0 else None
    refiEquity = refiPropVal - refi['approvedLoan']
    capitalRecycledPct = netCashOut / totalCash if totalCash > 0 else 0
    # Tk — take-out sizing at the first term boundary (computed unconditionally,
    # mirrors the Excel block; equals proj['takeoutInfo'] when the take-out is on)
    _gTkM = (1 + i['rentGrowth'] / 100) ** i['term']       # loyers cibles: taux marché
    _gTkI = (1 + i['indexGrowth'] / 100) ** i['term']      # loyers en place: indexation
    _fracTk = 1 - (1 - tr) ** i['term'] if tr > 0 else 1
    tkNOI = noi_at_rent(i['targetRent'] * _gTkM) * _fracTk + noi * _gTkI * (1 - _fracTk)
    tkVal = tkNOI / (i['refiCapRatePct'] / 100) if i['refiCapRatePct'] > 0 else 0
    if i['refiFinType'] == 'cmhc':
        tk = compute_cmhc(i['refiProg'], i['buildingType'], i['purpose'], i['refiMliPts'], i['refiLtv'],
                          i['refiRate'], i['refiTerm'], i['refiAmort'], i['province'], tkVal,
                          tkNOI, i['units'], None, True)
    else:
        tk = compute_conventional(i['refiConvLtv'], i['refiRate'], i['refiAmort'], i['refiConvDscr'],
                                  tkVal, tkNOI, 0)
    tkBrokerFee = tk['approvedLoan'] * i['refiBrokerPct'] / 100
    soldeAuRetrait = loan_balance(cmhc['insured'], i['rate'], cmhc['effAmort'], i['term'] * 12)
    _fracN = (1 - (1 - tr) ** holdingIntR) if tr > 0 else 1
    _renoUpToTk = (totalReno * ((1 - (1 - tr) ** min(i['term'], holdingIntR)) if tr > 0 else 1) / _fracN
                   if _fracN > 0 else 0)
    privRembRetrait = privFrac * (cfk + _renoUpToTk)
    priseFondsTakeout = (tk['approvedLoan'] - soldeAuRetrait - i['refiClosing'] - tk['provTax']
                         - tkBrokerFee - privRembRetrait)
    # L — IRR vector (from the projection rows; t=0 = cash à la fermeture).
    # Mode DPA: impôt annuel déduit par année, récupération à la sortie.
    venteCF = [0.0] * (holdingIntR + 1)
    venteCF[0] = -cashAtClose
    for row in proj['rows']:
        venteCF[row['y']] = row['cf'] + row['cashOut'] - (row['taxY'] if ccaOn else 0.0)
    venteCF[holdingIntR] += (salePrice - salesComm - penalty - i['closing'] - balanceAtExit
                             - privBalExit - cgTax - (recaptureTax if ccaOn else rentalTax))
    irrVente = irr(venteCF)
    reconciliationEcart = sum(venteCF) - netPostTax
    # Rendements NON LEVIER (avant impôt, sans dette ni financement privé)
    unlevCF = [0.0] * (holdingIntR + 1)
    unlevCF[0] = -(totalAcqCost + cfk)
    for row in proj['rows']:
        unlevCF[row['y']] = row['blendedNOI'] - annualReserve - row['renoY']
    unlevCF[holdingIntR] += salePrice - salesComm - i['closing']
    irrUnlev = irr(unlevCF)
    _uin = sum(v for v in unlevCF if v >= 0)
    _uout = -sum(v for v in unlevCF if v < 0)
    unlevMult = _uin / _uout if _uout > 0 else 0
    # DSCR et rendement de la dette: minimum (année) et moyenne sur la détention
    minDscr = float('inf'); minDscrYear = 1; sumDscr = 0.0
    minDebtYield = float('inf'); minDyYear = 1; sumDy = 0.0
    for row in proj['rows']:
        d = row['blendedNOI'] / row['ads'] if row['ads'] > 0 else 0
        dy = row['blendedNOI'] / row['balEnd'] if row['balEnd'] > 0 else 0
        if d < minDscr: minDscr = d; minDscrYear = row['y']
        if dy < minDebtYield: minDebtYield = dy; minDyYear = row['y']
        sumDscr += d; sumDy += dy
    avgDscr = sumDscr / holdingIntR
    avgDebtYield = sumDy / holdingIntR
    ccaTotal = proj['ccaTotal']
    o.update(locals()); o.pop('i'); o.pop('o', None)
    return o

# ─── Development defaults ─────────────────────────────────────────────────────
DEV_DEFAULTS = dict(
    dLandPrice=300000, dLandClosingPct=2.5, dUnits=6, dGfaPerUnit=1000,
    dHardMode='per-sqft', dHardPerSqft=225, dHardPerUnit=225000, dDemolition=0, dSiteWork=0,
    dSoftPct=5, dArchEng=40000, dPermitsDev=15000, dLegalMkt=10000, dPropTaxConstr=8000, dDevFeePct=5,
    dContingencyPct=10, dContingencyBasis='hard',
    dConstrMonths=12, dLeaseUpMonths=6, dAvgDrawPct=55,
    dConstrFinType='conventional', dTakeoutFinType='cmhc',
    dConvTakeoutLtv=75, dConvTakeoutDscr=1.20,
    dCmhcProgram='mli-select', dMliPoints=70, dBuildingType='standard',
    dTargetLtc=75, dTargetLtv=85, dConstrRate=4.50, dPermRate=4.50,
    dTerm=10, dAmortization=45, dProvince='qc', dBrokerFeePct=1,
    dStabRent=1600, dOtherIncomeDev=0, dStabVacancy=5, dStabOpexPct=28,
    dExitCap=5.0, dTargetMargin=20, dCommissionDev=2, dCgTaxRateDev=33,
    dHoldYearsDev=10, dRentGrowthDev=2.5, dCapDriftDev=0, dMarginalTaxDev=25,
    dSaleClosingDev=5000, dSaleNewRateDev=4.50,
    dSchlVacancy=None, dSchlOpexPct=None, dSchlCapRate=0,
)

def recalc_dev(i):
    o = {}
    landClosingDollar = i['dLandPrice'] * i['dLandClosingPct'] / 100
    land = i['dLandPrice'] + landClosingDollar
    gfa = i['dGfaPerUnit'] * i['dUnits']
    hardBase = i['dHardPerUnit'] * i['dUnits'] if i['dHardMode'] == 'per-unit' else i['dHardPerSqft'] * gfa
    hard = hardBase + i['dDemolition'] + i['dSiteWork']
    softBase = hard * i['dSoftPct'] / 100 + i['dArchEng'] + i['dPermitsDev'] + i['dLegalMkt'] + i['dPropTaxConstr']
    devFee = (land + hard + softBase) * i['dDevFeePct'] / 100
    soft = softBase + devFee
    contingencyBase = hard + soft if i['dContingencyBasis'] == 'hardsoft' else hard
    contingency = contingencyBase * i['dContingencyPct'] / 100
    costExFin = land + hard + soft + contingency
    stabGRI = i['dStabRent'] * i['dUnits'] * 12
    stabEGI = stabGRI * (1 - i['dStabVacancy'] / 100) + i['dOtherIncomeDev']
    stabNOI = stabEGI * (1 - i['dStabOpexPct'] / 100)
    GDV = stabNOI / (i['dExitCap'] / 100) if i['dExitCap'] > 0 else 0
    schlVacDev = i['dSchlVacancy'] if i['dSchlVacancy'] is not None else i['dStabVacancy']
    schlOpexDev = i['dSchlOpexPct'] if i['dSchlOpexPct'] is not None else i['dStabOpexPct']
    schlEgiDev = stabGRI * (1 - schlVacDev / 100) + i['dOtherIncomeDev']
    schlNoiDev = schlEgiDev * (1 - schlOpexDev / 100)
    econValueDev = schlNoiDev / (i['dSchlCapRate'] / 100) if i['dSchlCapRate'] > 0 else GDV
    if i['dTakeoutFinType'] == 'cmhc':
        premiumPurpose = 'purchase' if i['dConstrFinType'] == 'conventional' else 'construction'
        takeout = compute_cmhc(i['dCmhcProgram'], i['dBuildingType'], premiumPurpose, i['dMliPoints'],
                               i['dTargetLtv'], i['dPermRate'], i['dTerm'], i['dAmortization'],
                               i['dProvince'], GDV, schlNoiDev, i['dUnits'], econValueDev, False)
    else:
        takeout = compute_conventional(i['dConvTakeoutLtv'], i['dPermRate'], i['dAmortization'],
                                       i['dConvTakeoutDscr'], GDV, schlNoiDev, econValueDev)
    carryMonths = i['dConstrMonths'] + i['dLeaseUpMonths']
    annualR = monthly_rate(i['dConstrRate']) * 12
    constrLoanPrelim = costExFin * i['dTargetLtc'] / 100
    financingCarry = constrLoanPrelim * (i['dAvgDrawPct'] / 100) * annualR * carryMonths / 12
    constrLoanLtc = (costExFin + financingCarry) * i['dTargetLtc'] / 100
    constrLoan = min(constrLoanLtc, takeout['approvedLoan'])
    brokerFee = constrLoan * i['dBrokerFeePct'] / 100
    TDC = costExFin + financingCarry + brokerFee + takeout['provTax']
    yocDev = stabNOI / TDC if TDC > 0 else 0
    devSpread = yocDev - i['dExitCap'] / 100
    profit = GDV - TDC
    marginOnCost = profit / TDC if TDC > 0 else 0
    equityRequired = TDC - constrLoan
    equityRecapture = takeout['approvedLoan'] - constrLoan
    cashLeftIn = TDC - takeout['approvedLoan']
    stabCashFlow = stabNOI - takeout['ads']
    cashOnCash = stabCashFlow / cashLeftIn if cashLeftIn > 0 else 0
    saleNetDev = GDV * (1 - i['dCommissionDev'] / 100)
    saleProfitPreDev = saleNetDev - TDC
    saleCgTaxDev = max(0, saleProfitPreDev) * i['dCgTaxRateDev'] / 100
    saleProfitPostDev = saleProfitPreDev - saleCgTaxDev
    saleRoEDev = saleProfitPostDev / equityRequired if equityRequired > 0 else 0
    preStabYears = carryMonths / 12
    constrGridYears = max(1, math.ceil(preStabYears))
    devHold = max(1, round(i['dHoldYearsDev']))
    exitCapPctDev = i['dExitCap'] + i['dCapDriftDev'] / 100 * devHold
    noiAtExit = stabNOI * (1 + i['dRentGrowthDev'] / 100) ** devHold if (exitCapPctDev > 0 and stabNOI > 0) else 0
    salePriceDev2 = noiAtExit / (exitCapPctDev / 100) if exitCapPctDev > 0 else 0
    balanceAtExitDev = loan_balance(takeout['insured'], i['dPermRate'], takeout['effAmort'], devHold * 12)
    salesCommDev2 = salePriceDev2 * i['dCommissionDev'] / 100
    penaltyDev = 0 if devHold >= i['dTerm'] else break_penalty(balanceAtExitDev, i['dPermRate'],
                                                               i['dSaleNewRateDev'], devHold, i['dTerm'])
    grossGainDev2 = salePriceDev2 - TDC
    cgTaxDevHold = max(0, grossGainDev2) * i['dCgTaxRateDev'] / 100
    sumOpNOIDev = 0
    holdCF = [0.0] * (constrGridYears + devHold + 1)
    holdCF[0] = -(equityRequired if equityRequired > 0 else 0)
    for k in range(1, devHold + 1):
        noiK = stabNOI * (1 + i['dRentGrowthDev'] / 100) ** k
        sumOpNOIDev += noiK
        holdCF[constrGridYears + k] = noiK - takeout['ads']
    principalRepaidDev = takeout['insured'] - balanceAtExitDev if takeout['insured'] > 0 else 0
    totalInterestDev = takeout['ads'] * devHold - principalRepaidDev
    rentalTaxDev = max(0, sumOpNOIDev - max(0, totalInterestDev)) * i['dMarginalTaxDev'] / 100
    netSaleProceedsDev = (salePriceDev2 - salesCommDev2 - penaltyDev - balanceAtExitDev
                          - cgTaxDevHold - rentalTaxDev - i['dSaleClosingDev'])
    holdCF[constrGridYears + devHold] += netSaleProceedsDev
    refiCF = holdCF[:]
    refiCF[constrGridYears] += equityRecapture
    irrDevRefi = irr(refiCF) if equityRequired > 0 else None
    totalCashInRefi = sum(refiCF[1:])
    totalRetDevRefi = totalCashInRefi / equityRequired - 1 if equityRequired > 0 else 0
    exitMultDevRefi = totalCashInRefi / equityRequired if equityRequired > 0 else 0
    targetMarginDec = i['dTargetMargin'] / 100
    maxTDC = GDV / (1 + targetMarginDec) if (GDV > 0 and (1 + targetMarginDec) > 0) else 0
    nonLandCost = hard + soft + contingency + financingCarry + brokerFee + takeout['provTax']
    residualLandGross = maxTDC - nonLandCost
    residualLand = residualLandGross / (1 + i['dLandClosingPct'] / 100) if i['dLandClosingPct'] > 0 else residualLandGross
    o.update(locals()); o.pop('i'); o.pop('o', None)
    return o

if __name__ == '__main__':
    r = recalc(dict(ACQ_DEFAULTS))
    print('── ACQUISITION (defaults) ──')
    for k in ['noi', 'capRate', 'optNOI', 'yoc', 'totalCash', 'allInCost', 'exitCap', 'salePrice',
              'balanceAtExit', 'penalty', 'weightedRev', 'netPreTax', 'netPostTax', 'annRet',
              'exitMult', 'netCashOut', 'cashLeftIn', 'refiCF', 'irrVente']:
        print(f'{k:>16}: {r[k] if not isinstance(r[k], float) else round(r[k], 4)}')
    c = r['cmhc']
    for k in ['approvedLoan', 'binding', 'effLtv', 'baseRate', 'netRate', 'premium', 'insured',
              'mp', 'ads', 'dscrInsured', 'dscrBase', 'provTax', 'downPayment']:
        print(f'cmhc.{k:>12}: {c[k] if not isinstance(c[k], float) else round(c[k], 4)}')
    rf = r['refi']
    for k in ['approvedLoan', 'binding', 'minDscr', 'ads', 'premium']:
        print(f'refi.{k:>12}: {rf[k] if not isinstance(rf[k], float) else round(rf[k], 4)}')
    d = recalc_dev(dict(DEV_DEFAULTS))
    print('── DÉVELOPPEMENT (defaults) ──')
    for k in ['land', 'hard', 'soft', 'contingency', 'costExFin', 'financingCarry', 'TDC', 'stabNOI',
              'GDV', 'constrLoan', 'equityRequired', 'yocDev', 'devSpread', 'profit', 'marginOnCost',
              'equityRecapture', 'cashLeftIn', 'stabCashFlow', 'cashOnCash', 'saleProfitPostDev',
              'saleRoEDev', 'residualLand', 'salePriceDev2', 'netSaleProceedsDev', 'irrDevRefi',
              'totalRetDevRefi', 'exitMultDevRefi']:
        v = d[k]
        print(f'{k:>18}: {v if not isinstance(v, float) else round(v, 4)}')
    t = d['takeout']
    for k in ['approvedLoan', 'binding', 'premium', 'insured', 'ads', 'effLtv']:
        print(f'takeout.{k:>12}: {t[k] if not isinstance(t[k], float) else round(t[k], 4)}')
