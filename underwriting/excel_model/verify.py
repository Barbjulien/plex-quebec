"""Verify the Excel workbook reproduces the JS engine (via engine.py oracle).
Patches input cells for each scenario, evaluates all formulas with `formulas`,
and compares named output cells to the oracle."""
import json, os, sys, shutil
import openpyxl, formulas
import engine

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = "/Users/julienberube/Documents/Claude/PlexQuebec/underwriting/Plex_Quebec_Underwriting_Model.xlsx"
NAMES = json.load(open(os.path.join(HERE, "names.json")))

ACQ_MAP = {
    "NOI": "noi", "CapRate": "capRate", "RBP": "gri", "EGI": "egi", "Depenses": "opex",
    "GRM": "grm", "FraisAcqMontant": "acqCostDollar", "CoutTotalAcq": "totalAcqCost",
    "PrixParUnite": "pricePerUnit", "PrimeSurRole": "premiumOverRoll",
    "LoyerProjSortie": "projRentExit", "NOIOptimise": "optNOI", "RenoTotal": "totalReno",
    "CFKTotal": "cfk", "InvestTotal": "totalInvest", "SchlNOI": "schlNOI",
    "ValeurEconomique": "valeurEconomique", "MontantPrive": "privAmt", "EquiteReno": "renoEquity",
    "CoutPrive": "privCost", "ReserveTotale": "totalReserve", "ReserveAnnuelle": "annualReserve",
    "FraisCourtier": "brokerFee", "TotalCashRequis": "totalCash", "CoutToutInclus": "allInCost",
    "YieldOnCost": "yoc", "FracSortie": "fracExit", "NOIReversion": "exitNOI",
    "CapSortie": "exitCap", "PrixVente": "salePrice", "SoldeSortie": "balanceAtExit",
    "Penalite": "penalty", "CommissionVente": "salesComm", "RevenusPonderes": "weightedRev",
    "ServiceDetteTotal": "totalDS", "CapitalRembourse": "principalRepaid", "GainBrut": "grossProfit",
    "ImpotGain": "cgTax", "ImpotLocatif": "rentalTax", "ProfitOperationnel": "netOpProfit",
    "FraisFinancement": "finFees", "ProfitAvantImpot": "netPreTax", "ProfitApresImpot": "netPostTax",
    "RendementTotal": "totalRet", "RendementAnnuel": "annRet", "MultipleSortie": "exitMult",
    "IRRVente": "irrVente", "ValeurRefi": "refiPropVal", "SortieFondsBrute": "grossCashOut",
    "FraisCourtierRefi": "refiBrokerFee", "PriseDeFondsNette": "netCashOut",
    "CashRestant": "cashLeftIn", "FluxPostRefi": "refiCF", "EquiteRestante": "refiEquity",
    "CapitalRecycle": "capitalRecycledPct",
    # Multi-term DCF + turnover-phased reno (2026-11-22)
    "CashFermeture": "cashAtClose", "EquiteCfkFermeture": "cfkEquityClose",
    "EquiteRenoEtalee": "renoEquitySpread", "SoldePriveSortie": "privBalExit",
    "InteretsTotaux": "totalInterest", "ReconciliationEcart": "reconciliationEcart",
    "TkNOI": "tkNOI", "TkValeur": "tkVal", "TkCourtier": "tkBrokerFee",
    "SoldeAuRetrait": "soldeAuRetrait", "PrivRembRetrait": "privRembRetrait",
    "PriseFondsTakeout": "priseFondsTakeout",
    # Deux pistes de croissance + DPA + non levier + couverture (2026-11-22)
    "DpaTotale": "ccaTotal", "Recuperation": "recaptureTax", "ImpotLocatifTotal": "rentalTaxTotal",
    "IRRNonLevier": "irrUnlev", "MultipleNonLevier": "unlevMult",
    "DscrMin": "minDscr", "DscrMinAnnee": "minDscrYear", "DscrMoyen": "avgDscr",
    "RendDetteMin": "minDebtYield", "RendDetteMinAnnee": "minDyYear", "RendDetteMoyen": "avgDebtYield",
}
TK_MAP = {
    "TkPret": "approvedLoan", "TkPrime": "premium", "TkAds": "ads",
    "TkHypoAssuree": "insured", "TkTaxePrime": "provTax",
}
CMHC_MAP = {
    "PretApprouve": "approvedLoan", "PlafondLtvAcq": "ltvMaxLoan", "PlafondDscrAcq": "dscrMaxLoan",
    "ContrainteAcq": "binding", "LtvEffAcq": "effLtv", "DscrMinAcq": "minDscr",
    "TauxBaseAcq": "baseRate", "SurchargeAcq": "surcharge", "RabaisAcq": "discount",
    "TauxNetAcq": "netRate", "PrimeSchl": "premium", "DroitsSchl": "appFee",
    "HypoAssuree": "insured", "PaiementMensuel": "mp", "ServiceDetteAnnuel": "ads",
    "DscrAssure": "dscrInsured", "DscrSurBase": "dscrBase", "TaxePrimeSchl": "provTax",
    "MiseDeFonds": "downPayment", "AmortEffAcq": "effAmort",
    "DscrImplLtvAcq": "ltvImpliedDscr", "LtvImplDscrAcq": "dscrImpliedLtv",
}
REFI_MAP = {
    "PretRefi": "approvedLoan", "DscrMinRefi": "minDscr", "ServiceDetteRefi": "ads",
    "PrimeRefi": "premium", "TaxePrimeRefi": "provTax", "ContrainteRefi": "binding",
    "LtvEffRefi": "effLtv", "DscrAssureRefi": "dscrInsured", "PaiementRefi": "mp",
}
DEV_MAP = {
    "DevTerrain": "land", "DevHard": "hard", "DevSoft": "soft", "DevContingence": "contingency",
    "DevCoutExFin": "costExFin", "DevCarry": "financingCarry", "DevTDC": "TDC",
    "DevNOIStab": "stabNOI", "DevGDV": "GDV", "DevSchlNOI": "schlNoiDev",
    "DevValeurEcon": "econValueDev", "DevPretConstr": "constrLoan",
    "DevFraisCourtier": "brokerFee", "DevEquiteRequise": "equityRequired", "DevYoC": "yocDev",
    "DevSpread": "devSpread", "DevProfit": "profit", "DevMargeSurCout": "marginOnCost",
    "DevEquiteRecup": "equityRecapture", "DevCashRestant": "cashLeftIn",
    "DevFluxStab": "stabCashFlow", "DevCoC": "cashOnCash", "DevVenteNette": "saleNetDev",
    "DevProfitVenteApres": "saleProfitPostDev", "DevRoEVente": "saleRoEDev",
    "DevMaxTDC": "maxTDC", "DevResiduelTerrain": "residualLand", "DevPrixVente": "salePriceDev2",
    "DevSoldeSortie": "balanceAtExitDev", "DevPenalite": "penaltyDev",
    "DevImpotGain": "cgTaxDevHold", "DevSommeNOI": "sumOpNOIDev",
    "DevImpotLocatif": "rentalTaxDev", "DevProduitNetVente": "netSaleProceedsDev",
    "DevIRR": "irrDevRefi", "DevRendTotal": "totalRetDevRefi", "DevMultiple": "exitMultDevRefi",
}
TAKEOUT_MAP = {
    "DevTkPret": "approvedLoan", "DevTkPrime": "premium", "DevTkHypoAssuree": "insured",
    "DevTkAds": "ads", "DevTkLtvEff": "effLtv", "DevTkTaxePrime": "provTax",
    "DevTkContrainte": "binding", "DevTkDscrMin": "minDscr", "DevTkDscrAssure": "dscrInsured",
}

def compute(path):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    base = os.path.basename(path).upper()
    out = {}
    for k, v in sol.items():
        ku = k.upper()
        for nm, (sheet, cell) in NAMES.items():
            if ku.endswith(f"'[{base}]{sheet.upper()}'!{cell}"):
                try:
                    out[nm] = v.value[0, 0]
                except Exception:
                    out[nm] = v.value if not hasattr(v, "value") else v.value
    return out

def close(a, b):
    if isinstance(b, str) or isinstance(a, str):
        return str(a) == str(b)
    if a is None or b is None:
        return a == b
    try:
        a = float(a); b = float(b)
    except (TypeError, ValueError):
        return False
    if abs(b) < 1e-6:
        return abs(a - b) < 1e-4
    return abs(a - b) / abs(b) < 1e-6

def check(tag, got, oracle, mapping, sub=None):
    fails = 0
    src = oracle[sub] if sub else oracle
    for nm, key in mapping.items():
        if nm not in got:
            print(f"  [{tag}] MISSING cell value: {nm}"); fails += 1; continue
        want = src.get(key)
        g = got[nm]
        # special cases
        if nm == "IRRVente" or nm == "DevIRR":
            if want is None:
                continue
        if not close(g, want):
            print(f"  [{tag}] MISMATCH {nm}: excel={g!r} oracle={want!r}"); fails += 1
    return fails

def patch(src, dst, changes):
    wb = openpyxl.load_workbook(src)
    for nm, val in changes.items():
        sheet, cell = NAMES[nm]
        wb[sheet][cell] = val
    wb.save(dst)

def run_scenario(label, xlsx_changes, acq_inputs, dev_inputs):
    path = os.path.join(HERE, f"variant_{label}.xlsx")
    patch(SRC, path, xlsx_changes)
    got = compute(path)
    r = engine.recalc(acq_inputs)
    dv = engine.recalc_dev(dev_inputs)
    fails = 0
    fails += check(label, got, r, ACQ_MAP)
    fails += check(label, got, r, CMHC_MAP, sub="cmhc")
    fails += check(label, got, r, REFI_MAP, sub="refi")
    fails += check(label, got, r, TK_MAP, sub="tk")
    fails += check(label, got, dv, DEV_MAP)
    fails += check(label, got, dv, TAKEOUT_MAP, sub="takeout")
    n = len(ACQ_MAP) + len(CMHC_MAP) + len(REFI_MAP) + len(TK_MAP) + len(DEV_MAP) + len(TAKEOUT_MAP)
    print(f"[{label}] {n - fails}/{n} OK, {fails} fails")
    return fails

total = 0

# ── Scenario 1: defaults ──
total += run_scenario("defaults", {}, dict(engine.ACQ_DEFAULTS), dict(engine.DEV_DEFAULTS))

# ── Scenario 2: broker RBP mode + MLI Select + construction + conventional refi ──
xlsx2 = {
    "BaseDepenses": "% RBP", "ModeLoyer": "Annuel total", "LoyerAnnuel": 78000,
    "Programme": "MLI Select", "PointsMLI": 70, "ButPret": "Construction",
    "Amortissement": 45, "SchlCapRate": 5.5, "SchlDepensesPct": 35, "SchlInoccupation": 3,
    "FinTypeRefi": "Conventionnel", "Detention": 7, "TauxRoulement": 20, "MoisCFK": 1,
    "NouveauTaux": 3.5, "Terme": 10,
}
acq2 = dict(engine.ACQ_DEFAULTS)
acq2.update(opexBasis="rbp", rentMode="annual", currentRentAnnual=78000, program="mli-select",
            mliPts=70, purpose="construction", amort=45, schlCapRate=5.5, schlOpexPct=35,
            schlVacancy=3, refiFinType="conventional", holding=7, turnover=20, cfkMonths=1,
            newRate=3.5, term=10)
dev2 = {
    "DevModeHard": "$ / unité", "DevBaseContingence": "Hard+Soft", "DevFinConstr": "SCHL",
    "DevPointsMLI": 100, "DevDetention": 12, "DevDeriveCap": 0.1, "DevSchlCapRate": 5.2,
    "DevSchlInocc": 3, "DevSchlOpex": 30,
}
devi2 = dict(engine.DEV_DEFAULTS)
devi2.update(dHardMode="per-unit", dContingencyBasis="hardsoft", dConstrFinType="cmhc",
             dMliPoints=100, dHoldYearsDev=12, dCapDriftDev=0.1, dSchlCapRate=5.2,
             dSchlVacancy=3, dSchlOpexPct=30)
total += run_scenario("branches", {**xlsx2, **dev2}, acq2, devi2)

# ── Scenario 3: conventional acquisition + conventional dev takeout, penalty paths ──
xlsx3 = {
    "FinTypeAcq": "Conventionnel", "LtvConvAcq": 70, "DscrConvAcq": 1.25,
    "Detention": 3, "NouveauTaux": 3.0, "TypeImmeuble": "Autre",
    "DevFinTakeout": "Conventionnel", "DevDetention": 5, "DevNouveauTaux": 3.5,
    "DevLtvConvTakeout": 70, "DevDscrConvTakeout": 1.3,
}
acq3 = dict(engine.ACQ_DEFAULTS)
acq3.update(acqFinType="conventional", acqConvLtv=70, acqConvDscr=1.25, holding=3,
            newRate=3.0, buildingType="other")
devi3 = dict(engine.DEV_DEFAULTS)
devi3.update(dTakeoutFinType="conventional", dHoldYearsDev=5, dSaleNewRateDev=3.5,
             dConvTakeoutLtv=70, dConvTakeoutDscr=1.3)
total += run_scenario("conv", xlsx3, acq3, devi3)

# ── Scenario 4: multi-term renewal (détention > terme), private 50 %, CFK ──
xlsx4 = {
    "Detention": 10, "TauxRenouvellement": 5.25, "PctPrive": 50, "MoisCFK": 2,
}
acq4 = dict(engine.ACQ_DEFAULTS)
acq4.update(holding=10, renewalRate=5.25, privPct=50, cfkMonths=2)
total += run_scenario("renewal", xlsx4, acq4, dict(engine.DEV_DEFAULTS))

# ── Scenario 5: equity take-out at the first renewal ──
xlsx5 = {
    "Detention": 10, "RetraitEquite": "Oui",
}
acq5 = dict(engine.ACQ_DEFAULTS)
acq5.update(holding=10, takeout=True)
total += run_scenario("takeout", xlsx5, acq5, dict(engine.DEV_DEFAULTS))

# ── Scenario 6: impôt annuel avec DPA (CCA), privé 50 %, CFK ──
xlsx6 = {
    "Detention": 10, "PctPrive": 50, "MoisCFK": 2, "UtiliserDPA": "Oui",
}
acq6 = dict(engine.ACQ_DEFAULTS)
acq6.update(holding=10, privPct=50, cfkMonths=2, ccaOn=True)
total += run_scenario("cca", xlsx6, acq6, dict(engine.DEV_DEFAULTS))

# ── Scenario 7: indexation < croissance marché (loss-to-lease qui s'élargit) ──
xlsx7 = {
    "TauxIndexation": 1, "CroissanceLoyer": 3, "Detention": 7, "TauxRenouvellement": 5.25,
}
acq7 = dict(engine.ACQ_DEFAULTS)
acq7.update(indexGrowth=1, rentGrowth=3, holding=7, renewalRate=5.25)
total += run_scenario("idx", xlsx7, acq7, dict(engine.DEV_DEFAULTS))

print("TOTAL FAILS:", total)
sys.exit(1 if total else 0)
