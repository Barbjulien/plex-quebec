"""Build the Plex Québec underwriting Excel workbook — live-formula port of
PlexQuebec/underwriting/index.html (engine of 2026-11-22, broker-calibrated):
multi-term DCF (renouvellement de terme + retrait d'équité au premier terme)
and turnover-phased renovations (décaissement réno au rythme du roulement,
financement privé tiré au prorata, intérêts sur le solde tiré).

Sheets: Sommaire · Lisez-moi · Intrants · Analyse · Flux · Développement · Tables SCHL
Every formula mirrors recalc() / recalcDev() in the JS, verified by verify.py.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

ACCENT = "3A5C7A"; GOLD = "B8893A"; CREAM = "F7F7F5"
INPUT_FILL = PatternFill("solid", fgColor="E7EEF5")
KEY_FILL = PatternFill("solid", fgColor="F3EBDD")
HEADER_FILL = PatternFill("solid", fgColor=ACCENT)
SUBHEAD_FILL = PatternFill("solid", fgColor="D8E1E9")
THIN = Side(style="thin", color="B9C4CE")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

F_CUR = '#,##0" $"'
F_CUR2 = '#,##0.00" $"'
F_PCT = "0.00%"
F_PCT1 = "0.0%"
F_X = '0.00"x"'
F_NUM = "#,##0"
F_NUM1 = "#,##0.0"
F_INP = "#,##0.00"      # raw percent-style inputs entered as whole numbers (85 = 85 %)
F_INPCT = '0.00" %"'

wb = Workbook()
NAMES = {}

def register(ws, row, col, nm):
    ref = f"'{ws.title}'!${get_column_letter(col)}${row}"
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
    NAMES[nm] = (ws.title, f"{get_column_letter(col)}{row}")

class W:
    """Sequential row writer: label (A), value (B), note (C)."""
    def __init__(self, ws, start=1, label_col=1, val_col=2, note_col=3):
        self.ws, self.r = ws, start
        self.lc, self.vc, self.nc = label_col, val_col, note_col

    def title(self, text):
        ws = self.ws
        c = ws.cell(self.r, self.lc, text)
        c.font = Font(bold=True, size=13, color="FFFFFF")
        for col in range(self.lc, self.nc + 1):
            ws.cell(self.r, col).fill = HEADER_FILL
        self.r += 1
        return self

    def section(self, text):
        ws = self.ws
        c = ws.cell(self.r, self.lc, text)
        c.font = Font(bold=True, size=11, color=ACCENT)
        for col in range(self.lc, self.nc + 1):
            ws.cell(self.r, col).fill = SUBHEAD_FILL
        self.r += 1
        return self

    def blank(self, n=1):
        self.r += n
        return self

    def note(self, text):
        c = self.ws.cell(self.r, self.lc, text)
        c.font = Font(italic=True, size=9, color="6B7682")
        self.r += 1
        return self

    def inp(self, label, value, nm=None, fmt=F_INP, note=None, dv=None):
        ws = self.ws
        ws.cell(self.r, self.lc, label)
        c = ws.cell(self.r, self.vc, value)
        c.fill = INPUT_FILL; c.border = BOX; c.number_format = fmt
        c.font = Font(bold=True, color="1F3A52")
        if note:
            n = ws.cell(self.r, self.nc, note); n.font = Font(size=9, color="6B7682")
        if nm: register(ws, self.r, self.vc, nm)
        if dv is not None: dv.add(c)
        self.r += 1
        return self

    def mirror(self, label, formula, fmt=F_INP, note="Défini sur l'onglet Sommaire"):
        ws = self.ws
        lab = ws.cell(self.r, self.lc, label)
        lab.font = Font(italic=True, color="8A94A0")
        c = ws.cell(self.r, self.vc, formula)
        c.number_format = fmt; c.font = Font(italic=True, color="8A94A0")
        n = ws.cell(self.r, self.nc, note); n.font = Font(size=9, italic=True, color="8A94A0")
        self.r += 1
        return self

    def out(self, label, formula, nm=None, fmt=F_CUR, key=False, note=None):
        ws = self.ws
        lab = ws.cell(self.r, self.lc, label)
        c = ws.cell(self.r, self.vc, formula)
        c.number_format = fmt
        if key:
            lab.font = Font(bold=True); c.font = Font(bold=True, color=ACCENT)
            for col in range(self.lc, self.nc + 1):
                ws.cell(self.r, col).fill = KEY_FILL
        if note:
            n = ws.cell(self.r, self.nc, note); n.font = Font(size=9, color="6B7682")
        if nm: register(ws, self.r, self.vc, nm)
        self.r += 1
        return self

def setwidths(ws, widths):
    for col, wd in widths.items():
        ws.column_dimensions[col].width = wd

def mkdv(ws, options):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    return dv

# ═════════════════════════════════ TABLES SCHL ═══════════════════════════════
tbl = wb.active
tbl.title = "Tables SCHL"
tbl.sheet_properties.tabColor = "8898A8"
setwidths(tbl, {"A": 30, "B": 14, "C": 15, "D": 15, "E": 15, "F": 15, "G": 42})

t = W(tbl, start=1)
t.title("Tables SCHL — barèmes (en vigueur 14 juillet 2025)")
t.blank()
t.section("Prime SCHL — taux de base selon le palier prêt ÷ prix d'achat")
hdr = ["LTV borne inf. (%)", "LTV jusqu'à (%)", "Standard · Achat", "Standard · Constr.",
       "Autre · Achat", "Autre · Constr."]
for j, h in enumerate(hdr):
    c = tbl.cell(t.r, 1 + j, h); c.font = Font(bold=True, size=9); c.fill = SUBHEAD_FILL
t.r += 1
prem_first = t.r
bands = [(0, 65, .026, .0325, .063, .0655), (65.001, 70, .0285, .0375, .066, .0685),
         (70.001, 75, .0335, .0425, .069, .0715), (75.001, 80, .0435, .050, .0705, .073),
         (80.001, 85, .0535, .060, .0775, .080), (85.001, 90, .059, .0675, .080, .0825),
         (90.001, 100, .0615, .070, .0875, .090)]
for row in bands:
    for j, v in enumerate(row):
        c = tbl.cell(t.r, 1 + j, v); c.border = BOX
        c.number_format = "0.000" if j == 0 else ("0" if j == 1 else "0.000%")
    t.r += 1
prem_last = t.r - 1
wb.defined_names["PrimeBornesInf"] = DefinedName(
    "PrimeBornesInf", attr_text=f"'Tables SCHL'!$A${prem_first}:$A${prem_last}")
wb.defined_names["PrimeTable"] = DefinedName(
    "PrimeTable", attr_text=f"'Tables SCHL'!$C${prem_first}:$F${prem_last}")
t.note("Palier déterminé par prêt ÷ prix d'achat (pratique courtier, validée Multi-Prêts 2026-07).")
t.blank()
t.section("Surcharge d'amortissement & droits")
t.out("Surcharge amortissement", "+0,25 % par tranche de 5 ans au-delà de 25 ans", fmt="@")
t.inp("Droits SCHL par unité", 150, nm="DroitsParUnite", fmt=F_CUR,
      note="Capitalisés dans le prêt total (prêt + prime + droits) — fiche Multi-Prêts 2026-07.")
t.blank()
t.section("Taxe provinciale sur la prime")
tax_first = t.r
for prov, rate in [("QC", 0.09975), ("ON", 0.08), ("MB", 0.07), ("SK", 0.06), ("Autre", 0.0)]:
    tbl.cell(t.r, 1, prov).border = BOX
    c = tbl.cell(t.r, 2, rate); c.border = BOX; c.number_format = "0.000%"
    t.r += 1
wb.defined_names["TaxeProvTable"] = DefinedName(
    "TaxeProvTable", attr_text=f"'Tables SCHL'!$A${tax_first}:$B${t.r - 1}")
t.blank()
t.section("MLI Select — paliers de points (référence)")
for row in [("Points", "Rabais prime", "LTV max", "Amort. max"),
            ("50–69", "10 %", "85 %", "40 ans"),
            ("70–99", "20 %", "90 %", "45 ans"),
            ("100+", "30 %", "95 %", "50 ans"),
            ("MLI Market", "—", "85 %", "40 ans")]:
    for j, v in enumerate(row):
        c = tbl.cell(t.r, 1 + j, v); c.border = BOX
        if t.r == t.r: c.font = Font(size=9)
    t.r += 1
t.blank()
t.section("DSCR minimum (fiche SCHL Standard Locatif)")
for txt in ["Type « Autre » : 1,20 (tous cas)",
            "MLI Select : 1,10",
            "MLI Market, 5–6 unités : 1,10 achat / 1,20 refinancement",
            "MLI Market, 7+ unités : 1,30 (terme < 10 ans) / 1,20 (terme ≥ 10 ans)"]:
    t.note(txt)

# ═════════════════════════════════ INTRANTS ══════════════════════════════════
inp = wb.create_sheet("Intrants")
inp.sheet_properties.tabColor = ACCENT
setwidths(inp, {"A": 38, "B": 16, "C": 52})

dv_mode = mkdv(inp, ["Moyen / unité", "Annuel total"])
dv_base = mkdv(inp, ["% EGI", "% RBP"])
dv_fin = mkdv(inp, ["SCHL", "Conventionnel"])
dv_prog = mkdv(inp, ["MLI Market", "MLI Select"])
dv_type = mkdv(inp, ["Standard", "Autre"])
dv_but = mkdv(inp, ["Achat", "Construction"])
dv_prov = mkdv(inp, ["QC", "ON", "MB", "SK", "Autre"])
dv_amort = mkdv(inp, ["25", "30", "35", "40", "45", "50"])
dv_terme = mkdv(inp, ["5", "10"])
dv_ouinon = mkdv(inp, ["Oui", "Non"])

w = W(inp, start=1)
w.title("Plex Québec — Sous-écriture d'acquisition · Intrants")
w.note("Cellules bleues = intrants. Tout le reste se calcule. Conventions : % saisis en nombres (85 = 85 %).")
w.blank()
w.section("A — Propriété & Acquisition")
w.mirror("Nom de la propriété", "=NomPropriete", fmt="General")
w.mirror("Prix d'achat", "=PrixAchat", fmt=F_CUR)
w.mirror("Nombre d'unités", "=Unites", fmt=F_NUM)
w.inp("Type d'unités", "6 x 4½", fmt="@")
w.inp("Superficie moyenne / unité (pi²)", 1000, nm="SuperficieMoy", fmt=F_NUM)
w.inp("Valeur au rôle", 969400, nm="ValeurRole", fmt=F_CUR)
w.inp("Frais d'acquisition (% du prix)", 2.5, nm="FraisAcqPct", fmt=F_INPCT,
      note="Mutation, notaire, inspection, etc.")
w.inp("Frais courtier hypothécaire (% du prêt)", 1, nm="FraisCourtierPct", fmt=F_INPCT)
w.blank()
w.section("B — Revenus actuels")
w.inp("Mode de saisie du loyer", "Moyen / unité", nm="ModeLoyer", fmt="@", dv=dv_mode,
      note="« Moyen / unité » utilise le loyer mensuel ; « Annuel total » utilise le RBP annuel.")
w.mirror("Loyer actuel moyen / unité / mois", "=LoyerActuel", fmt=F_CUR)
w.inp("Loyer annuel total (RBP)", 72000, nm="LoyerAnnuel", fmt=F_CUR)
w.inp("Taux d'inoccupation (%)", 5, nm="TauxInoccupation", fmt=F_INPCT)
w.inp("Autres revenus annuels", 0, nm="AutresRevenus", fmt=F_CUR)
w.inp("Base des dépenses", "% EGI", nm="BaseDepenses", fmt="@", dv=dv_base,
      note="% RBP = convention courtier (reproduit les fiches Multi-Prêts).")
w.inp("Dépenses d'exploitation (%)", 30, nm="DepensesPct", fmt=F_INPCT)
w.blank()
w.section("C — Revenus optimisés & détention")
w.mirror("Loyer cible / unité / mois", "=LoyerCible", fmt=F_CUR)
w.inp("Croissance loyers marché (% / an)", 2, nm="CroissanceLoyer", fmt=F_INPCT,
      note="Croissance du loyer CIBLE (marché).")
w.inp("Indexation loyers en place (% / an)", 2, nm="TauxIndexation", fmt=F_INPCT,
      note="Croissance des loyers des locataires EN PLACE (indexation TAL). Si < marché, le loss-to-lease s'élargit.")
w.mirror("Période de détention (ans)", "=Detention", fmt=F_NUM)
w.inp("Inoccupation optimisée (%)", 5, nm="InoccOptimisee", fmt=F_INPCT)
w.inp("Dépenses optimisées (%)", 30, nm="DepensesOptimisees", fmt=F_INPCT)
w.inp("Taux de roulement annuel (%)", 17.5, nm="TauxRoulement", fmt=F_INPCT,
      note="Transition loyers actuels → cibles (série géométrique).")
w.inp("Taux au renouvellement (%)", 4.25, nm="TauxRenouvellement", fmt=F_INPCT,
      note="Actif si détention > terme : le prêt se renouvelle à chaque fin de terme à ce taux.")
w.inp("Retrait d'équité au renouvellement", "Non", nm="RetraitEquite", fmt="@", dv=dv_ouinon,
      note="« Oui » : le 1er renouvellement devient un refi (Section K) ; retrait net compté dans le TRI.")
w.blank()
w.section("D — Rénovations & Cash for keys")
w.inp("Coût de rénovation / unité", 20000, nm="CoutRenoUnite", fmt=F_CUR,
      note="Décaissé au rythme du roulement des unités (pas à la fermeture) ; budget complété à la sortie.")
w.inp("% des unités à rénover", 50, nm="PctReno", fmt=F_INPCT)
w.inp("Cash for keys (mois de loyer)", 0, nm="MoisCFK", fmt=F_NUM1)
w.inp("% des unités avec CFK", 20, nm="PctCFK", fmt=F_INPCT)
w.blank()
w.section("E — Financement à l'acquisition")
w.inp("Type de financement", "SCHL", nm="FinTypeAcq", fmt="@", dv=dv_fin)
w.inp("Programme SCHL", "MLI Market", nm="Programme", fmt="@", dv=dv_prog)
w.inp("Points MLI Select", 50, nm="PointsMLI", fmt=F_NUM, note="Utilisé si programme = MLI Select.")
w.inp("Type d'immeuble", "Standard", nm="TypeImmeuble", fmt="@", dv=dv_type)
w.inp("But du prêt", "Achat", nm="ButPret", fmt="@", dv=dv_but)
w.mirror("LTV cible (%)", "=LtvCible", fmt=F_INPCT)
w.mirror("Taux d'intérêt (%)", "=TauxInteret", fmt=F_INPCT)
w.inp("Terme (ans)", 5, nm="Terme", fmt=F_NUM, dv=dv_terme)
w.inp("Amortissement (ans)", 40, nm="Amortissement", fmt=F_NUM, dv=dv_amort)
w.inp("Province", "QC", nm="Province", fmt="@", dv=dv_prov)
w.inp("Cap rate SCHL (valeur économique)", None, nm="SchlCapRate", fmt=F_INPCT,
      note="Laisser vide → valeur de prêt = prix d'achat.")
w.inp("Dépenses normalisées SCHL (%)", None, nm="SchlDepensesPct", fmt=F_INPCT,
      note="Laisser vide → dépenses de la section B.")
w.inp("Inoccupation normalisée SCHL (%)", None, nm="SchlInoccupation", fmt=F_INPCT,
      note="Laisser vide → inoccupation de la section B.")
w.inp("Conventionnel — LTV (%)", 75, nm="LtvConvAcq", fmt=F_INPCT,
      note="Utilisé si type = Conventionnel.")
w.inp("Conventionnel — DSCR minimum", 1.2, nm="DscrConvAcq", fmt=F_X)
w.blank()
w.section("F — Financement privé (rénovations)")
w.inp("% des rénos financé au privé", 100, nm="PctPrive", fmt=F_INPCT)
w.inp("Taux du financement privé (%)", 14, nm="TauxPrive", fmt=F_INPCT,
      note="Tiré au prorata de chaque décaissement ; intérêt sur le solde réellement tiré.")
w.blank()
w.section("G — Réserves & CapEx")
w.inp("Provision mensuelle / unité", 50, nm="ReserveUnitMois", fmt=F_CUR)
w.blank()
w.section("H — Scénario de vente")
w.mirror("Dérive du cap rate (% / an)", "=DeriveCapRate", fmt=F_INPCT)
w.inp("Taux du marché au bris (%)", 4.25, nm="NouveauTaux", fmt=F_INPCT,
      note="Pour la pénalité IRD si vente avant la fin du terme.")
w.inp("Taux marginal — revenu locatif (%)", 26.5, nm="TauxMarginal", fmt=F_INPCT)
w.inp("Taux effectif — gain en capital (%)", 33, nm="TauxGainCapital", fmt=F_INPCT)
w.inp("Impôt annuel avec DPA (CCA)", "Non", nm="UtiliserDPA", fmt="@", dv=dv_ouinon,
      note="« Oui » : impôt locatif annuel net de la DPA + récupération à la vente. « Non » : forfait à la sortie.")
w.inp("Taux DPA (%)", 4, nm="TauxDPA", fmt=F_INPCT, note="Catégorie 1 (bâtiment), dégressif.")
w.inp("% bâtiment (vs terrain)", 80, nm="PctBatiment", fmt=F_INPCT,
      note="Part du prix d'achat attribuable au bâtiment (seul le bâtiment s'amortit).")
w.inp("Commission de vente (%)", 2, nm="CommissionPct", fmt=F_INPCT)
w.inp("Frais de fermeture (vente)", 1500, nm="FraisFermeture", fmt=F_CUR,
      note="Saisir en valeur positive.")
w.blank()
w.section("K — Scénario de refinancement (garder)")
w.inp("Type de financement (refi)", "SCHL", nm="FinTypeRefi", fmt="@", dv=dv_fin)
w.inp("Programme SCHL (refi)", "MLI Market", nm="ProgrammeRefi", fmt="@", dv=dv_prog)
w.inp("Points MLI Select (refi)", 50, nm="PointsMLIRefi", fmt=F_NUM)
w.inp("Cap rate au refi (%)", 4.75, nm="CapRateRefi", fmt=F_INPCT)
w.inp("LTV cible refi (%)", 75, nm="LtvRefi", fmt=F_INPCT)
w.inp("Taux d'intérêt refi (%)", 4.25, nm="TauxRefi", fmt=F_INPCT)
w.inp("Terme refi (ans)", 5, nm="TermeRefi", fmt=F_NUM, dv=dv_terme)
w.inp("Amortissement refi (ans)", 40, nm="AmortRefi", fmt=F_NUM, dv=dv_amort)
w.inp("Frais de fermeture (refi)", 3000, nm="FraisFermetureRefi", fmt=F_CUR)
w.inp("Frais courtier refi (% du prêt)", 1, nm="CourtierRefiPct", fmt=F_INPCT)
w.inp("Conventionnel refi — LTV (%)", 75, nm="LtvConvRefi", fmt=F_INPCT)
w.inp("Conventionnel refi — DSCR minimum", 1.2, nm="DscrConvRefi", fmt=F_X)

# ═════════════════════════════════ ANALYSE ═══════════════════════════════════
ana = wb.create_sheet("Analyse")
ana.sheet_properties.tabColor = GOLD
setwidths(ana, {"A": 42, "B": 17, "C": 56})

a = W(ana, start=1)
a.title("Analyse — Acquisition")
a.blank()
a.section("Tableau de bord")
a.out("NOI actuel", "=NOI", fmt=F_CUR, key=True)
a.out("Cap rate à l'achat", "=CapRate", fmt=F_PCT)
a.out("NOI optimisé", "=NOIOptimise", fmt=F_CUR)
a.out("Yield on Cost", "=YieldOnCost", fmt=F_PCT)
a.out("Prêt approuvé", "=PretApprouve", fmt=F_CUR, key=True)
a.out("Cash requis", "=TotalCashRequis", fmt=F_CUR, key=True)
a.out("Profit net après impôt (vente)", "=ProfitApresImpot", fmt=F_CUR, key=True)
a.out("IRR (vente)", "=IRRVente", fmt=F_PCT)
a.out("Prise de fonds nette (refi)", "=PriseDeFondsNette", fmt=F_CUR, key=True)
a.blank()
a.section("A — Acquisition")
a.out("Frais d'acquisition ($)", "=PrixAchat*FraisAcqPct/100", nm="FraisAcqMontant")
a.out("Coût total d'acquisition", "=PrixAchat+FraisAcqMontant", nm="CoutTotalAcq")
a.out("Prix / unité", "=IF(Unites>0,PrixAchat/Unites,0)", nm="PrixParUnite")
a.out("Prime sur la valeur au rôle", "=IF(ValeurRole>0,PrixAchat/ValeurRole-1,0)",
      nm="PrimeSurRole", fmt=F_PCT)
a.blank()
a.section("B — Revenus actuels")
a.out("Loyer actuel effectif / unité / mois",
      '=IF(ModeLoyer="Annuel total",IF(Unites>0,LoyerAnnuel/Unites/12,0),LoyerActuel)',
      nm="LoyerActuelEff", fmt=F_CUR2)
a.out("Revenu brut potentiel (RBP)",
      '=IF(ModeLoyer="Annuel total",LoyerAnnuel,LoyerActuel*Unites*12)', nm="RBP")
a.out("Revenus effectifs bruts (EGI)", "=RBP*(1-TauxInoccupation/100)+AutresRevenus", nm="EGI")
a.out("Dépenses d'exploitation",
      '=IF(BaseDepenses="% RBP",RBP,EGI)*DepensesPct/100', nm="Depenses")
a.out("NOI actuel", "=EGI-Depenses", nm="NOI", key=True)
a.out("Cap rate à l'achat", "=IF(PrixAchat>0,NOI/PrixAchat,0)", nm="CapRate", fmt=F_PCT)
a.out("GRM", "=IF(RBP>0,PrixAchat/RBP,0)", nm="GRM", fmt=F_NUM1)
a.blank()
a.section("C — Revenus optimisés")
a.out("Loyer projeté à la sortie / unité",
      "=LoyerCible*(1+CroissanceLoyer/100)^Detention", nm="LoyerProjSortie", fmt=F_CUR2)
a.out("RBP optimisé (à la sortie)", "=LoyerProjSortie*Unites*12", nm="RBPOpt")
a.out("EGI optimisé", "=RBPOpt*(1-InoccOptimisee/100)+AutresRevenus", nm="EGIOpt")
a.out("Dépenses optimisées",
      '=IF(BaseDepenses="% RBP",RBPOpt,EGIOpt)*DepensesOptimisees/100', nm="DepensesOpt")
a.out("NOI optimisé", "=EGIOpt-DepensesOpt", nm="NOIOptimise", key=True)
a.blank()
a.section("D — Rénovations & Cash for keys")
a.out("Coût rénovations total", "=CoutRenoUnite*Unites*PctReno/100", nm="RenoTotal")
a.out("Cash for keys total", "=MoisCFK*LoyerActuelEff*Unites*PctCFK/100", nm="CFKTotal")
a.out("Investissement total", "=RenoTotal+CFKTotal", nm="InvestTotal", key=True)
a.out("Réno / pi²",
      "=IF(SuperficieMoy*Unites>0,RenoTotal/(SuperficieMoy*Unites),0)", nm="RenoParPi", fmt=F_CUR2)
a.blank()
a.section("E — Prêt à l'acquisition (SCHL ou conventionnel)")
a.out("Inoccupation normalisée SCHL (%)",
      "=IF(ISBLANK(SchlInoccupation),TauxInoccupation,SchlInoccupation)", nm="SchlVacEff", fmt=F_INPCT)
a.out("Dépenses normalisées SCHL (%)",
      "=IF(ISBLANK(SchlDepensesPct),DepensesPct,SchlDepensesPct)", nm="SchlOpexEff", fmt=F_INPCT)
a.out("EGI normalisé SCHL", "=RBP*(1-SchlVacEff/100)+AutresRevenus", nm="SchlEGI")
a.out("NOI normalisé SCHL",
      '=SchlEGI-IF(BaseDepenses="% RBP",RBP,SchlEGI)*SchlOpexEff/100', nm="SchlNOI",
      note="Base du dimensionnement DSCR du prêteur.")
a.out("Valeur économique SCHL",
      "=IF(ISBLANK(SchlCapRate),PrixAchat,IF(SchlCapRate>0,SchlNOI/(SchlCapRate/100),PrixAchat))",
      nm="ValeurEconomique")
a.out("Valeur de prêt (lending value)",
      "=IF(AND(ValeurEconomique>0,ValeurEconomique<PrixAchat),ValeurEconomique,PrixAchat)",
      nm="LvAcq")
a.out("Amortissement max du programme",
      '=IF(Programme="MLI Market",40,IF(PointsMLI>=100,50,IF(PointsMLI>=70,45,IF(PointsMLI>=50,40,25))))',
      nm="AmortMaxProg", fmt=F_NUM)
a.out("Amortissement effectif",
      '=IF(FinTypeAcq="SCHL",MIN(Amortissement,AmortMaxProg),Amortissement)',
      nm="AmortEffAcq", fmt=F_NUM)
a.out("LTV max du programme (%)",
      '=IF(Programme="MLI Market",85,IF(PointsMLI>=100,95,IF(PointsMLI>=70,90,85)))',
      nm="LtvMaxProg", fmt=F_NUM)
a.out("LTV plafond retenue (fraction)",
      '=IF(FinTypeAcq="SCHL",MIN(LtvCible,LtvMaxProg),LtvConvAcq)/100', nm="LtvCapAcq", fmt=F_PCT)
a.out("DSCR minimum requis (achat)",
      '=IF(FinTypeAcq="SCHL",IF(TypeImmeuble="Autre",1.2,IF(Programme="MLI Select",1.1,'
      'IF(Unites<=6,1.1,IF(Terme>=10,1.2,1.3)))),DscrConvAcq)', nm="DscrMinAcq", fmt=F_X)
a.out("Taux mensuel (comp. semi-annuelle)", "=(1+TauxInteret/200)^(1/6)-1",
      nm="TauxMensuelAcq", fmt="0.000000")
a.out("Constante hypothécaire",
      "=IF(TauxMensuelAcq=0,1/AmortEffAcq,TauxMensuelAcq*(1+TauxMensuelAcq)^(AmortEffAcq*12)"
      "/((1+TauxMensuelAcq)^(AmortEffAcq*12)-1)*12)", nm="ConstHypoAcq", fmt="0.0000%")
a.out("Plafond LTV (prêt max)", "=LvAcq*LtvCapAcq", nm="PlafondLtvAcq")
a.out("→ DSCR implicite à ce prêt",
      "=IF(AND(ConstHypoAcq>0,PlafondLtvAcq>0),SchlNOI/(PlafondLtvAcq*ConstHypoAcq),0)",
      nm="DscrImplLtvAcq", fmt=F_X)
a.out("Plafond DSCR (prêt max)",
      "=IF(ConstHypoAcq>0,SchlNOI/(DscrMinAcq*ConstHypoAcq),0)", nm="PlafondDscrAcq",
      note="Covenant dimensionné sur le prêt de base, AVANT prime (pratique courtier).")
a.out("→ % de la valeur à ce prêt", "=IF(LvAcq>0,PlafondDscrAcq/LvAcq,0)",
      nm="LtvImplDscrAcq", fmt=F_PCT)
a.out("PRÊT APPROUVÉ", "=MIN(PlafondLtvAcq,MAX(0,PlafondDscrAcq))", nm="PretApprouve", key=True)
a.out("Contrainte liante", '=IF(PlafondDscrAcq<=PlafondLtvAcq,"DSCR","LTV")',
      nm="ContrainteAcq", fmt="General")
a.out("% de financement (prêt ÷ prix)", "=IF(PrixAchat>0,PretApprouve/PrixAchat,0)",
      nm="LtvEffAcq", fmt=F_PCT, note="Détermine le palier de prime.")
a.out("DSCR (prêt de base)",
      "=IF(AND(ConstHypoAcq>0,PretApprouve>0),SchlNOI/(PretApprouve*ConstHypoAcq),0)",
      nm="DscrSurBase", fmt=F_X)
a.out("Colonne du barème de prime",
      '=IF(TypeImmeuble="Autre",2,0)+IF(ButPret="Construction",1,0)+1', nm="ColPrimeAcq", fmt=F_NUM)
a.out("Taux de base de la prime",
      '=IF(FinTypeAcq="SCHL",INDEX(PrimeTable,MATCH(LtvEffAcq*100,PrimeBornesInf,1),ColPrimeAcq),0)',
      nm="TauxBaseAcq", fmt="0.000%")
a.out("Surcharge amortissement",
      '=IF(FinTypeAcq="SCHL",MAX(0,INT((AmortEffAcq-25)/5))*0.0025,0)', nm="SurchargeAcq", fmt="0.000%")
a.out("Rabais MLI Select",
      '=IF(AND(FinTypeAcq="SCHL",Programme="MLI Select"),IF(PointsMLI>=100,0.3,'
      'IF(PointsMLI>=70,0.2,IF(PointsMLI>=50,0.1,0))),0)', nm="RabaisAcq", fmt="0%")
a.out("Taux net de la prime", "=MAX(0,(TauxBaseAcq+SurchargeAcq)*(1-RabaisAcq))",
      nm="TauxNetAcq", fmt="0.000%")
a.out("Prime SCHL nette", "=PretApprouve*TauxNetAcq", nm="PrimeSchl", key=True)
a.out("Droits SCHL (capitalisés)", '=IF(FinTypeAcq="SCHL",Unites*DroitsParUnite,0)', nm="DroitsSchl")
a.out("Hypothèque assurée (prêt + prime + droits)", "=PretApprouve+PrimeSchl+DroitsSchl",
      nm="HypoAssuree", key=True)
a.out("Paiement mensuel",
      "=IF(TauxMensuelAcq=0,HypoAssuree/(AmortEffAcq*12),HypoAssuree*TauxMensuelAcq"
      "*(1+TauxMensuelAcq)^(AmortEffAcq*12)/((1+TauxMensuelAcq)^(AmortEffAcq*12)-1))",
      nm="PaiementMensuel", fmt=F_CUR2)
a.out("Service de la dette annuel", "=PaiementMensuel*12", nm="ServiceDetteAnnuel", fmt=F_CUR2)
a.out("DSCR (hypothèque assurée)", "=IF(ServiceDetteAnnuel>0,SchlNOI/ServiceDetteAnnuel,0)",
      nm="DscrAssure", fmt=F_X, note="Indicatif seulement — tombe sous le minimum, c'est normal.")
a.out("TVQ / taxe provinciale sur la prime", "=PrimeSchl*VLOOKUP(Province,TaxeProvTable,2,FALSE)",
      nm="TaxePrimeSchl")
a.out("Mise de fonds", "=PrixAchat-PretApprouve", nm="MiseDeFonds", key=True)
a.blank()
a.section("F — Financement privé")
a.out("Montant financé au privé", "=InvestTotal*PctPrive/100", nm="MontantPrive")
a.out("Équité réno + CFK (engagée)", "=InvestTotal*(1-PctPrive/100)", nm="EquiteReno",
      note="Engagée au total ; la part réno se décaisse au rythme du roulement (onglet Flux).")
a.out("Coût d'intérêt privé (détention)",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$W$5:$W$34)",
      nm="CoutPrive", note="Intérêts sur le solde réellement tiré (colonne W de l'onglet Flux).")
a.out("Solde privé dû à la sortie", "=INDEX(Flux!$X$5:$X$34,DetentionInt)", nm="SoldePriveSortie",
      note="Égal au montant privé si aucun retrait ; 0 après un retrait d'équité (remboursé au refi).")
a.blank()
a.section("F2 — Renouvellement de terme & retrait d'équité")
a.out("Taux au renouvellement effectif (%)",
      "=IF(TauxRenouvellement>0,TauxRenouvellement,TauxInteret)", nm="TauxRenouvEff", fmt=F_INPCT)
a.out("Taux mensuel (renouvellement)", "=(1+TauxRenouvEff/200)^(1/6)-1",
      nm="TauxMensuelRenouv", fmt="0.000000")
a.out("Retrait d'équité actif ?",
      '=IF(AND(RetraitEquite="Oui",Detention>Terme,Terme<DetentionInt),TRUE,FALSE)',
      nm="TkActif", fmt="General",
      note="Le retrait exige détention > terme (le 1er renouvellement devient le refi).")
a.out("Année du retrait (9999 = aucun)", "=IF(TkActif,Terme,9999)", nm="AnneeRetrait", fmt=F_NUM)
a.blank()
a.section("G — Réserves")
a.out("Réserves totales (détention)", "=ReserveUnitMois*Unites*Detention*12", nm="ReserveTotale")
a.out("Réserve annuelle", "=IF(Detention>0,ReserveTotale/Detention,0)", nm="ReserveAnnuelle")
a.blank()
a.section("H — Cash requis & coût tout inclus")
a.out("Frais de courtier hypothécaire", "=PretApprouve*FraisCourtierPct/100", nm="FraisCourtier")
a.out("Équité CFK (payée à la fermeture)", "=CFKTotal*(1-PctPrive/100)", nm="EquiteCfkFermeture")
a.out("CASH À LA FERMETURE",
      "=MiseDeFonds+FraisAcqMontant+TaxePrimeSchl+FraisCourtier+EquiteCfkFermeture",
      nm="CashFermeture", key=True,
      note="Exclut l'équité réno (décaissée au rythme du roulement, dans les flux annuels).")
a.out("Équité réno décaissée (années 1 à détention)",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$Y$5:$Y$34)", nm="EquiteRenoEtalee")
a.out("ÉQUITÉ TOTALE ENGAGÉE (cash total requis)",
      "=MiseDeFonds+EquiteReno+FraisAcqMontant+TaxePrimeSchl+FraisCourtier",
      nm="TotalCashRequis", key=True,
      note="Dénominateur des rendements. Droits SCHL capitalisés — aucun cash pour ce poste.")
a.out("Coût tout inclus (base YoC)",
      "=CoutTotalAcq+InvestTotal+FraisCourtier+TaxePrimeSchl+DroitsSchl", nm="CoutToutInclus")
a.out("Yield on Cost", "=IF(CoutToutInclus>0,NOIOptimise/CoutToutInclus,0)",
      nm="YieldOnCost", fmt=F_PCT, key=True)
a.blank()
a.section("I — Scénario de vente")
a.out("Détention entière (grille IRR)", "=MIN(ROUND(IF(Detention>0,Detention,1),0),30)",
      nm="DetentionInt", fmt=F_NUM)
a.out("% unités optimisées (fin de grille)",
      "=IF(TauxRoulement>0,1-(1-TauxRoulement/100)^DetentionInt,1)", nm="FracDetInt", fmt=F_PCT1,
      note="Normalise l'étalement du budget réno (complété à la sortie).")
a.out("% unités optimisées à la sortie",
      "=IF(TauxRoulement>0,1-(1-TauxRoulement/100)^Detention,1)", nm="FracSortie", fmt=F_PCT1)
a.out("NOI actuel projeté à la sortie", "=NOI*(1+TauxIndexation/100)^Detention",
      nm="NOIActuelSortie", note="Loyers en place indexés (TAL), pas au taux marché.")
a.out("NOI à la réversion (roulement ajusté)",
      "=NOIOptimise*FracSortie+NOIActuelSortie*(1-FracSortie)", nm="NOIReversion", key=True)
a.out("Cap rate de sortie", "=CapRate+DeriveCapRate/100*Detention", nm="CapSortie", fmt=F_PCT)
a.out("Prix de vente estimé", "=IF(CapSortie>0,NOIReversion/CapSortie,0)", nm="PrixVente", key=True)
a.out("Prix de vente / unité", "=IF(Unites>0,PrixVente/Unites,0)", nm="PrixVenteUnite")
a.out("Solde hypothécaire à la sortie", "=INDEX(Flux!$Q$5:$Q$34,DetentionInt)", nm="SoldeSortie",
      note="Phase-aware : tient compte des renouvellements / du refi de retrait (onglet Flux).")
a.out("Taux de la phase active à la sortie (%)", "=INDEX(Flux!$K$5:$K$34,DetentionInt)",
      nm="PhaseTauxSortie", fmt=F_INPCT)
a.out("Début de la phase active (année)", "=INDEX(Flux!$J$5:$J$34,DetentionInt)",
      nm="PhaseDebutSortie", fmt=F_NUM)
a.out("Terme de la phase active (ans)", "=INDEX(Flux!$M$5:$M$34,DetentionInt)",
      nm="PhaseTermeSortie", fmt=F_NUM)
a.out("Pénalité de bris",
      "=IF(PhaseTermeSortie-(DetentionInt-PhaseDebutSortie)<=0,0,"
      "MAX(SoldeSortie*((1+PhaseTauxSortie/200)^(1/6)-1)*3,"
      "SoldeSortie*MAX(0,(PhaseTauxSortie-NouveauTaux)/100)"
      "*(PhaseTermeSortie-(DetentionInt-PhaseDebutSortie))))", nm="Penalite",
      note="MAX(3 mois d'intérêt, IRD) — mesurée contre le terme ACTIF (renouvelé).")
a.out("Commission de vente", "=PrixVente*CommissionPct/100", nm="CommissionVente")
a.out("Revenus pondérés (détention)",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$H$5:$H$34)", nm="RevenusPonderes")
a.out("Service de la dette total",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$O$5:$O$34)", nm="ServiceDetteTotal",
      note="Somme phase-aware (renouvellements / refi inclus).")
a.out("Capital remboursé (équité récupérée)",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*(Flux!$P$5:$P$34-Flux!$Q$5:$Q$34))",
      nm="CapitalRembourse")
a.out("Intérêts totaux payés", "=ServiceDetteTotal-CapitalRembourse", nm="InteretsTotaux")
a.out("Gain en capital brut", "=PrixVente-CoutTotalAcq-InvestTotal", nm="GainBrut")
a.out("Gain imposable (net des frais de vente)",
      "=MAX(0,GainBrut-CommissionVente-FraisFermeture)", nm="GainImposable")
a.out("Impôt sur le gain en capital", "=GainImposable*TauxGainCapital/100", nm="ImpotGain")
a.out("Revenu locatif imposable", "=MAX(0,RevenusPonderes-InteretsTotaux-CoutPrive)",
      nm="RevenuLocImposable")
a.out("Mode DPA actif ?", '=IF(UtiliserDPA="Oui",TRUE,FALSE)', nm="DpaActif", fmt="General")
a.out("DPA totale prise (détention)",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$AO$5:$AO$34)", nm="DpaTotale",
      note="Pool dégressif: part bâtiment + rénos capitalisées; plafonnée au revenu de l'année.")
a.out("Impôt sur le revenu locatif",
      "=IF(DpaActif,SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$AQ$5:$AQ$34),"
      "RevenuLocImposable*TauxMarginal/100)", nm="ImpotLocatif",
      note="Mode DPA: somme des impôts annuels (nets de DPA). Sinon: forfait à la sortie.")
a.out("Récupération d'amortissement (à la vente)",
      "=IF(DpaActif,DpaTotale*TauxMarginal/100,0)", nm="Recuperation")
a.out("Impôts locatifs totaux", "=ImpotLocatif+Recuperation", nm="ImpotLocatifTotal")
a.out("Profit opérationnel net", "=RevenusPonderes-ServiceDetteTotal-CoutPrive-ReserveTotale",
      nm="ProfitOperationnel")
a.out("Frais de financement (prime + taxes + droits + courtier)",
      "=PrimeSchl+TaxePrimeSchl+DroitsSchl+FraisCourtier"
      "+IF(TkActif,TkPrime+TkTaxePrime+TkDroits+TkCourtier+FraisFermetureRefi,0)",
      nm="FraisFinancement",
      note="Inclut les frais du refi de retrait quand le retrait d'équité est actif.")
a.out("PROFIT NET AVANT IMPÔT",
      "=GainBrut+ProfitOperationnel+CapitalRembourse-FraisFinancement-Penalite",
      nm="ProfitAvantImpot", key=True)
a.out("PROFIT NET APRÈS IMPÔT",
      "=ProfitAvantImpot-ImpotGain-CommissionVente-FraisFermeture-ImpotLocatifTotal",
      nm="ProfitApresImpot", key=True,
      note="Se réconcilie exactement avec le vecteur de flux IRR (onglet Flux).")
a.out("Rendement total (sur mise de fonds)", "=IF(TotalCashRequis>0,ProfitApresImpot/TotalCashRequis,0)",
      nm="RendementTotal", fmt=F_PCT)
a.out("Rendement annualisé",
      "=IF(Detention>0,IF(RendementTotal>-1,(1+RendementTotal)^(1/Detention)-1,-1),0)",
      nm="RendementAnnuel", fmt=F_PCT)
a.out("Multiple de sortie", "=IF(TotalCashRequis>0,(TotalCashRequis+ProfitApresImpot)/TotalCashRequis,0)",
      nm="MultipleSortie", fmt=F_X)
a.out("IRR (vente)", "=IRR(Flux!$AD$4:$AD$34)", nm="IRRVente", fmt=F_PCT)
a.out("Réconciliation (Σ flux − profit net)", "=SUM(Flux!$AD$4:$AD$34)-ProfitApresImpot",
      nm="ReconciliationEcart", fmt=F_CUR2,
      note="Doit être ≈ 0 : le vecteur IRR se réconcilie exactement avec le profit net.")
a.blank()
a.section("I2 — Rendements non levier & couverture de la dette")
a.out("TRI non levier", "=IRR(Flux!$AR$4:$AR$34)", nm="IRRNonLevier", fmt=F_PCT,
      note="Avant impôt, sans dette ni financement privé (colonne AR de l'onglet Flux).")
a.out("Multiple non levier",
      '=IF(SUMIF(Flux!$AR$4:$AR$34,"<0")<0,'
      'SUMIF(Flux!$AR$4:$AR$34,">0")/(-SUMIF(Flux!$AR$4:$AR$34,"<0")),0)',
      nm="MultipleNonLevier", fmt=F_X)
a.out("DSCR minimum (détention)", "=MIN(Flux!$AK$5:$AK$34)", nm="DscrMin", fmt=F_X,
      note="NOI blendé / service de dette de la phase active — l'année que le prêteur stresse.")
a.out("→ année du minimum", "=MATCH(DscrMin,Flux!$AK$5:$AK$34,0)", nm="DscrMinAnnee", fmt=F_NUM)
a.out("DSCR moyen (détention)",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$AE$5:$AE$34)/DetentionInt",
      nm="DscrMoyen", fmt=F_X)
a.out("Rendement de la dette minimum", "=MIN(Flux!$AL$5:$AL$34)", nm="RendDetteMin", fmt=F_PCT)
a.out("→ année du minimum", "=MATCH(RendDetteMin,Flux!$AL$5:$AL$34,0)", nm="RendDetteMinAnnee", fmt=F_NUM)
a.out("Rendement de la dette moyen",
      "=SUMPRODUCT((Flux!$A$5:$A$34<=DetentionInt)*Flux!$AH$5:$AH$34)/DetentionInt",
      nm="RendDetteMoyen", fmt=F_PCT)
a.blank()
a.section("K — Scénario de refinancement (garder)")
a.out("Valeur estimée au refi", "=IF(CapRateRefi>0,NOIReversion/(CapRateRefi/100),0)", nm="ValeurRefi")
a.out("Amortissement max (programme refi)",
      '=IF(ProgrammeRefi="MLI Market",40,IF(PointsMLIRefi>=100,50,IF(PointsMLIRefi>=70,45,'
      'IF(PointsMLIRefi>=50,40,25))))', nm="AmortMaxRefiProg", fmt=F_NUM)
a.out("Amortissement effectif (refi)",
      '=IF(FinTypeRefi="SCHL",MIN(AmortRefi,AmortMaxRefiProg),AmortRefi)', nm="AmortEffRefi", fmt=F_NUM)
a.out("LTV max (programme refi, %)",
      '=IF(ProgrammeRefi="MLI Market",85,IF(PointsMLIRefi>=100,95,IF(PointsMLIRefi>=70,90,85)))',
      nm="LtvMaxRefiProg", fmt=F_NUM)
a.out("LTV plafond retenue (refi, fraction)",
      '=IF(FinTypeRefi="SCHL",MIN(LtvRefi,LtvMaxRefiProg),LtvConvRefi)/100', nm="LtvCapRefi", fmt=F_PCT)
a.out("DSCR minimum requis (refi)",
      '=IF(FinTypeRefi="SCHL",IF(TypeImmeuble="Autre",1.2,IF(ProgrammeRefi="MLI Select",1.1,'
      'IF(Unites<=6,1.2,IF(TermeRefi>=10,1.2,1.3)))),DscrConvRefi)', nm="DscrMinRefi", fmt=F_X)
a.out("Taux mensuel (refi)", "=(1+TauxRefi/200)^(1/6)-1", nm="TauxMensuelRefi", fmt="0.000000")
a.out("Constante hypothécaire (refi)",
      "=IF(TauxMensuelRefi=0,1/AmortEffRefi,TauxMensuelRefi*(1+TauxMensuelRefi)^(AmortEffRefi*12)"
      "/((1+TauxMensuelRefi)^(AmortEffRefi*12)-1)*12)", nm="ConstHypoRefi", fmt="0.0000%")
a.out("Plafond LTV (refi)", "=ValeurRefi*LtvCapRefi", nm="PlafondLtvRefi")
a.out("Plafond DSCR (refi)", "=IF(ConstHypoRefi>0,NOIReversion/(DscrMinRefi*ConstHypoRefi),0)",
      nm="PlafondDscrRefi")
a.out("PRÊT REFI APPROUVÉ", "=MIN(PlafondLtvRefi,MAX(0,PlafondDscrRefi))", nm="PretRefi", key=True)
a.out("Contrainte liante (refi)", '=IF(PlafondDscrRefi<=PlafondLtvRefi,"DSCR","LTV")',
      nm="ContrainteRefi", fmt="General")
a.out("LTV effectif (refi)", "=IF(ValeurRefi>0,PretRefi/ValeurRefi,0)", nm="LtvEffRefi", fmt=F_PCT)
a.out("Taux de base de la prime (refi)",
      '=IF(FinTypeRefi="SCHL",INDEX(PrimeTable,MATCH(LtvEffRefi*100,PrimeBornesInf,1),ColPrimeAcq),0)',
      nm="TauxBaseRefi", fmt="0.000%")
a.out("Surcharge amortissement (refi)",
      '=IF(FinTypeRefi="SCHL",MAX(0,INT((AmortEffRefi-25)/5))*0.0025,0)', nm="SurchargeRefi", fmt="0.000%")
a.out("Rabais MLI Select (refi)",
      '=IF(AND(FinTypeRefi="SCHL",ProgrammeRefi="MLI Select"),IF(PointsMLIRefi>=100,0.3,'
      'IF(PointsMLIRefi>=70,0.2,IF(PointsMLIRefi>=50,0.1,0))),0)', nm="RabaisRefi", fmt="0%")
a.out("Taux net de la prime (refi)", "=MAX(0,(TauxBaseRefi+SurchargeRefi)*(1-RabaisRefi))",
      nm="TauxNetRefi", fmt="0.000%")
a.out("Prime SCHL nette (refi)", "=PretRefi*TauxNetRefi", nm="PrimeRefi")
a.out("Droits SCHL (refi, capitalisés)", '=IF(FinTypeRefi="SCHL",Unites*DroitsParUnite,0)',
      nm="DroitsRefi")
a.out("Hypothèque assurée (refi)", "=PretRefi+PrimeRefi+DroitsRefi", nm="HypoAssureeRefi")
a.out("Paiement mensuel (refi)",
      "=IF(TauxMensuelRefi=0,HypoAssureeRefi/(AmortEffRefi*12),HypoAssureeRefi*TauxMensuelRefi"
      "*(1+TauxMensuelRefi)^(AmortEffRefi*12)/((1+TauxMensuelRefi)^(AmortEffRefi*12)-1))",
      nm="PaiementRefi", fmt=F_CUR2)
a.out("Service de la dette annuel (refi)", "=PaiementRefi*12", nm="ServiceDetteRefi", fmt=F_CUR2)
a.out("DSCR sur hypothèque assurée (refi)",
      "=IF(ServiceDetteRefi>0,NOIReversion/ServiceDetteRefi,0)", nm="DscrAssureRefi", fmt=F_X)
a.out("TVQ / taxe sur la prime (refi)", "=PrimeRefi*VLOOKUP(Province,TaxeProvTable,2,FALSE)",
      nm="TaxePrimeRefi")
a.out("Sortie de fonds brute", "=PretRefi-SoldeSortie-Penalite", nm="SortieFondsBrute")
a.out("Frais de courtier (refi)", "=PretRefi*CourtierRefiPct/100", nm="FraisCourtierRefi")
a.out("PRISE DE FONDS NETTE",
      "=SortieFondsBrute-FraisFermetureRefi-TaxePrimeRefi-FraisCourtierRefi-SoldePriveSortie",
      nm="PriseDeFondsNette", key=True,
      note="Rembourse le solde privé réellement dû à la sortie (0 après un retrait d'équité).")
a.out("Cash restant dans l'actif", "=TotalCashRequis-PriseDeFondsNette", nm="CashRestant")
a.out("Flux opérationnel annuel (post-refi)", "=NOIReversion-ServiceDetteRefi-ReserveAnnuelle",
      nm="FluxPostRefi")
a.out("Cash-on-cash post-refi", '=IF(CashRestant>0,FluxPostRefi/CashRestant,"∞")',
      nm="CoCPostRefi", fmt=F_PCT)
a.out("Équité restante dans l'actif", "=ValeurRefi-PretRefi", nm="EquiteRestante")
a.out("Capital recyclé (% de la mise de fonds)",
      "=IF(TotalCashRequis>0,PriseDeFondsNette/TotalCashRequis,0)", nm="CapitalRecycle", fmt=F_PCT)
a.blank()
a.section("Tk — Refi de retrait au premier renouvellement (an = Terme)")
a.note("Dimensionné avec les paramètres de la Section K sur le NOI pondéré de l'année du terme.")
a.note("Calculé en tout temps ; n'entre dans le DCF que si « Retrait d'équité » = Oui et détention > terme.")
a.out("NOI pondéré à l'année du terme", "=INDEX(Flux!$H$5:$H$34,Terme)", nm="TkNOI")
a.out("Valeur estimée au retrait", "=IF(CapRateRefi>0,TkNOI/(CapRateRefi/100),0)", nm="TkValeur")
a.out("Plafond LTV (retrait)", "=TkValeur*LtvCapRefi", nm="TkPlafondLtv")
a.out("Plafond DSCR (retrait)", "=IF(ConstHypoRefi>0,TkNOI/(DscrMinRefi*ConstHypoRefi),0)",
      nm="TkPlafondDscr")
a.out("PRÊT DE RETRAIT APPROUVÉ", "=MIN(TkPlafondLtv,MAX(0,TkPlafondDscr))", nm="TkPret", key=True)
a.out("LTV effectif (retrait)", "=IF(TkValeur>0,TkPret/TkValeur,0)", nm="TkLtvEff", fmt=F_PCT)
a.out("Taux de base de la prime (retrait)",
      '=IF(FinTypeRefi="SCHL",INDEX(PrimeTable,MATCH(TkLtvEff*100,PrimeBornesInf,1),ColPrimeAcq),0)',
      nm="TkTauxBase", fmt="0.000%")
a.out("Taux net de la prime (retrait)", "=MAX(0,(TkTauxBase+SurchargeRefi)*(1-RabaisRefi))",
      nm="TkTauxNet", fmt="0.000%")
a.out("Prime SCHL (retrait)", "=TkPret*TkTauxNet", nm="TkPrime")
a.out("Droits SCHL (retrait)", '=IF(FinTypeRefi="SCHL",Unites*DroitsParUnite,0)', nm="TkDroits")
a.out("Hypothèque assurée (retrait)", "=TkPret+TkPrime+TkDroits", nm="TkHypoAssuree")
a.out("Paiement mensuel (retrait)",
      "=IF(TauxMensuelRefi=0,TkHypoAssuree/(AmortEffRefi*12),TkHypoAssuree*TauxMensuelRefi"
      "*(1+TauxMensuelRefi)^(AmortEffRefi*12)/((1+TauxMensuelRefi)^(AmortEffRefi*12)-1))",
      nm="TkPaiement", fmt=F_CUR2)
a.out("Service de dette annuel (retrait)", "=TkPaiement*12", nm="TkAds", fmt=F_CUR2)
a.out("TVQ / taxe sur la prime (retrait)", "=TkPrime*VLOOKUP(Province,TaxeProvTable,2,FALSE)",
      nm="TkTaxePrime")
a.out("Frais de courtier (retrait)", "=TkPret*CourtierRefiPct/100", nm="TkCourtier")
a.out("Solde hypothécaire au retrait (prêt original)",
      "=IF(TauxMensuelAcq=0,MAX(0,HypoAssuree-HypoAssuree/(AmortEffAcq*12)*Terme*12),"
      "MAX(0,HypoAssuree*(1+TauxMensuelAcq)^(Terme*12)-PaiementMensuel"
      "*((1+TauxMensuelAcq)^(Terme*12)-1)/TauxMensuelAcq))", nm="SoldeAuRetrait")
a.out("Solde privé remboursé au retrait",
      "=PctPrive/100*(CFKTotal+IF(FracDetInt>0,"
      "RenoTotal*IF(TauxRoulement>0,1-(1-TauxRoulement/100)^MIN(Terme,DetentionInt),1)/FracDetInt,0))",
      nm="PrivRembRetrait",
      note="Solde réellement tiré à cette date (tirages au rythme du roulement) — pas le montant engagé.")
a.out("RETRAIT NET (compté dans le TRI à l'an du terme)",
      "=TkPret-SoldeAuRetrait-FraisFermetureRefi-TkTaxePrime-TkCourtier-PrivRembRetrait",
      nm="PriseFondsTakeout", key=True)

# ═════════════════════════════════ FLUX ══════════════════════════════════════
flx = wb.create_sheet("Flux")
flx.sheet_properties.tabColor = GOLD
widths = {"A": 7, "B": 9, "C": 11, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12,
          "J": 9, "K": 9, "L": 9, "M": 9, "N": 10, "O": 12, "P": 13, "Q": 13, "R": 11,
          "S": 11, "T": 11, "U": 12, "V": 12, "W": 10, "X": 12, "Y": 11, "Z": 10,
          "AA": 12, "AB": 12, "AC": 13, "AD": 14, "AE": 8, "AF": 9, "AG": 9, "AH": 10,
          "AI": 11, "AJ": 12, "AK": 9, "AL": 9, "AM": 12, "AN": 12, "AO": 11, "AP": 12,
          "AQ": 11, "AR": 13}
setwidths(flx, widths)
flx.cell(1, 1, "Projections — Flux de trésorerie (années 1–30 ; " \
               "la détention active est définie par l'intrant « Période de détention »)")
flx.cell(1, 1).font = Font(bold=True, size=13, color="FFFFFF")
for col in range(1, 45):
    flx.cell(1, col).fill = HEADER_FILL
flx.cell(2, 1, "Prêt modélisé en PHASES : renouvellement à chaque fin de terme (taux au renouvellement), "
               "retrait d'équité optionnel au 1er terme. Rénos décaissées au rythme du roulement ; "
               "financement privé tiré au prorata. Seules les années ≤ détention comptent dans les "
               "sommes et le vecteur IRR (colonne AD).")
flx.cell(2, 1).font = Font(italic=True, size=9, color="6B7682")
headers = ["Année", "% opt.", "Loyer cible", "RBP opt.", "EGI opt.", "NOI opt.", "NOI actuel",
           "NOI pondéré", "EGI pondéré",
           "Début phase", "Taux phase", "Amort phase", "Terme phase", "Fin de terme",
           "Dette", "Solde début", "Solde fin", "Intérêts",
           "Réno décaissée", "Tirage privé", "Solde privé déb.", "Solde privé brut",
           "Int. privé", "Solde privé fin", "Équité réno",
           "Réserves", "CF opér.", "CF cumulé", "Retrait net", "Flux Vente (IRR)",
           "DSCR", "CoC", "YoC", "Debt yield", "Loyer moyen", "NOI/u/mois",
           "DSCR filtré", "Rend. dette filtré", "Revenu imposable", "Pool DPA début",
           "DPA année", "Pool DPA fin", "Impôt annuel (DPA)", "Flux non levier"]
for j, h in enumerate(headers):
    c = flx.cell(3, 1 + j, h); c.font = Font(bold=True, size=9, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="55789A"); c.alignment = Alignment(wrap_text=True)
flx.freeze_panes = "B4"
# t = 0 row
flx.cell(4, 1, 0).number_format = F_NUM
flx.cell(4, 28, "Cash à la fermeture →").font = Font(italic=True, size=9, color="6B7682")
c = flx.cell(4, 30, "=-CashFermeture"); c.number_format = F_CUR; c.font = Font(bold=True)
# t = 0 non levier: prix + frais d'acquisition + CFK (sans dette ni privé)
c = flx.cell(4, 44, "=-(CoutTotalAcq+CFKTotal)"); c.number_format = F_CUR; c.font = Font(bold=True)
for y in range(1, 31):
    r = 4 + y
    pr = r - 1
    A = f"$A{r}"
    flx.cell(r, 1, y).number_format = F_NUM
    put = lambda col, f, fmt: (flx.cell(r, col, f), setattr(flx.cell(r, col), "number_format", fmt))
    # Revenue blend (turnover curve)
    put(2, f"=IF(TauxRoulement>0,1-(1-TauxRoulement/100)^{A},1)", F_PCT1)
    put(3, f"=LoyerCible*(1+CroissanceLoyer/100)^{A}", F_NUM)
    put(4, f"=C{r}*Unites*12", F_NUM)
    put(5, f"=D{r}*(1-InoccOptimisee/100)+AutresRevenus", F_NUM)
    put(6, f'=E{r}-IF(BaseDepenses="% RBP",D{r},E{r})*DepensesOptimisees/100', F_NUM)
    put(7, f"=NOI*(1+TauxIndexation/100)^{A}", F_NUM)
    put(8, f"=F{r}*B{r}+G{r}*(1-B{r})", F_NUM)
    put(9, f"=E{r}*B{r}+EGI*(1+TauxIndexation/100)^{A}*(1-B{r})", F_NUM)
    # Loan phases: J début, K taux, L amort, M terme, N fin-de-terme (événement en FIN d'année)
    if y == 1:
        put(10, 0, F_NUM)
        put(11, "=TauxInteret", F_INPCT)
        put(12, "=AmortEffAcq", F_NUM)
        put(13, "=Terme", F_NUM)
        put(15, "=ServiceDetteAnnuel", F_NUM)
        put(16, "=HypoAssuree", F_NUM)
        put(21, "=CFKTotal*PctPrive/100", F_NUM)
    else:
        put(10, f'=IF(N{pr}<>"",$A{pr},J{pr})', F_NUM)
        put(11, f'=IF(N{pr}<>"",IF($A{pr}=AnneeRetrait,TauxRefi,TauxRenouvEff),K{pr})', F_INPCT)
        put(12, f'=IF(N{pr}<>"",IF($A{pr}=AnneeRetrait,AmortEffRefi,MAX(1,L{pr}-M{pr})),L{pr})', F_NUM)
        put(13, f'=IF(N{pr}<>"",IF($A{pr}=AnneeRetrait,TermeRefi,Terme),M{pr})', F_NUM)
        put(15, f'=IF(N{pr}<>"",IF($A{pr}=AnneeRetrait,TkAds,'
                f"Q{pr}*IF(TauxMensuelRenouv=0,1/MAX(1,L{pr}-M{pr}),"
                f"TauxMensuelRenouv*(1+TauxMensuelRenouv)^(MAX(1,L{pr}-M{pr})*12)"
                f"/((1+TauxMensuelRenouv)^(MAX(1,L{pr}-M{pr})*12)-1)*12)),O{pr})", F_NUM)
        put(16, f'=IF(AND(N{pr}<>"",$A{pr}=AnneeRetrait),TkHypoAssuree,Q{pr})', F_NUM)
        put(21, f"=X{pr}", F_NUM)
    put(14, f'=IF(AND({A}=J{r}+M{r},{A}<DetentionInt),IF({A}=AnneeRetrait,"Retrait","Renouv."),"")',
        "General")
    # Balance recurrence: (1+rm)^12 = (1+K/200)^2 (composition semi-annuelle)
    put(17, f"=IF(K{r}=0,MAX(0,P{r}-O{r}),MAX(0,P{r}*(1+K{r}/200)^2"
            f"-O{r}/12*((1+K{r}/200)^2-1)/((1+K{r}/200)^(1/6)-1)))", F_NUM)
    put(18, f"=O{r}-(P{r}-Q{r})", F_NUM)
    # Reno phased with turnover (normalized to complete by exit); private drawn pro-rata
    put(19, f"=IF({A}>DetentionInt,0,IF(FracDetInt>0,RenoTotal*(B{r}-IF({A}=1,0,B{pr}))/FracDetInt,0))",
        F_NUM)
    put(20, f"=IF({A}<=AnneeRetrait,S{r}*PctPrive/100,0)", F_NUM)
    put(22, f"=U{r}+T{r}", F_NUM)
    put(23, f"=IF(OR(Detention<=0,{A}>AnneeRetrait),0,(U{r}+T{r}/2)*TauxPrive/100)", F_NUM)
    put(24, f"=IF({A}=AnneeRetrait,0,V{r})", F_NUM)
    put(25, f"=S{r}-T{r}", F_NUM)
    # Flows
    put(26, "=ReserveAnnuelle", F_NUM)
    put(27, f"=H{r}-O{r}-Z{r}-W{r}-Y{r}", F_NUM)
    put(28, f'=IF({A}<=DetentionInt,SUM($AA$5:AA{r}),"")', F_NUM)
    put(29, f"=IF(AND(TkActif,{A}=AnneeRetrait),PriseFondsTakeout,0)", F_NUM)
    put(30, f"=IF({A}>DetentionInt,0,AA{r}+AC{r}-IF(DpaActif,AQ{r},0)+IF({A}=DetentionInt,"
            "PrixVente-CommissionVente-Penalite-FraisFermeture-SoldeSortie-SoldePriveSortie"
            "-ImpotGain-IF(DpaActif,Recuperation,ImpotLocatif),0))", F_NUM)
    # Metrics
    put(31, f"=IF(O{r}>0,H{r}/O{r},0)", F_X)
    put(32, f"=IF(TotalCashRequis>0,AA{r}/TotalCashRequis,0)", F_PCT1)
    put(33, f"=IF(CoutToutInclus>0,H{r}/CoutToutInclus,0)", F_PCT1)
    put(34, f"=IF(Q{r}>0,H{r}/Q{r},0)", F_PCT1)
    put(35, f"=B{r}*C{r}+(1-B{r})*LoyerActuelEff*(1+TauxIndexation/100)^{A}", F_NUM)
    put(36, f"=IF(Unites>0,H{r}/Unites/12,0)", F_NUM)
    # Helpers min/moyenne (sentinelle 1E9 hors détention)
    put(37, f"=IF({A}<=DetentionInt,AE{r},1000000000)", F_X)
    put(38, f"=IF({A}<=DetentionInt,AH{r},1000000000)", F_PCT1)
    # DPA (CCA): revenu imposable, pool dégressif (+ réno de l'année), DPA plafonnée au revenu
    put(39, f"=H{r}-R{r}-W{r}", F_NUM)
    put(40, f"=IF({A}=1,PrixAchat*PctBatiment/100,AP{pr})+S{r}", F_NUM)
    put(41, f"=IF(DpaActif,MIN(AN{r}*TauxDPA/100,MAX(0,AM{r})),0)", F_NUM)
    put(42, f"=AN{r}-AO{r}", F_NUM)
    put(43, f"=IF(DpaActif,MAX(0,AM{r}-AO{r})*TauxMarginal/100,0)", F_NUM)
    # Flux non levier: NOI blendé − réserves − réno décaissée (+ vente nette à la sortie)
    put(44, f"=IF({A}>DetentionInt,0,H{r}-Z{r}-S{r}"
            f"+IF({A}=DetentionInt,PrixVente-CommissionVente-FraisFermeture,0))", F_NUM)
# Waterfall cards (labels span A:D by overflow; values in col E)
wf = W(flx, start=37, label_col=1, val_col=5, note_col=6)
wf.section("Sortie — Vente (année = détention)")
wf.out("Prix de vente", "=PrixVente")
wf.out("Frais de clôture (commission + pénalité + fermeture)",
       "=-(CommissionVente+Penalite+FraisFermeture)")
wf.out("Remboursement hypothèque (phase active)", "=-SoldeSortie")
wf.out("Remboursement financement privé (solde dû)", "=-SoldePriveSortie")
wf.out("Impôts (gain + locatif)", "=-(ImpotGain+ImpotLocatif)")
wf.out("= Produit net de vente",
       "=PrixVente-CommissionVente-Penalite-FraisFermeture-SoldeSortie-SoldePriveSortie"
       "-ImpotGain-ImpotLocatif", key=True)
wf.blank()
wf.section("Renouvellement / retrait (si détention > terme)")
wf.out("Retrait net au 1er renouvellement (an = Terme)", "=IF(TkActif,PriseFondsTakeout,0)",
       note="0 si « Retrait d'équité » = Non.")
wf.blank()
wf.section("Sortie — Refi (garder)")
wf.out("Prêt refi", "=PretRefi")
wf.out("Remboursement hypothèque (phase active)", "=-SoldeSortie")
wf.out("Frais & pénalité (fermeture + TVQ + courtier + pénalité)",
       "=-(Penalite+TaxePrimeRefi+FraisCourtierRefi+FraisFermetureRefi)")
wf.out("Remboursement financement privé (solde dû)", "=-SoldePriveSortie")
wf.out("= Prise de fonds nette", "=PriseDeFondsNette", key=True)
wf.out("Flux opérationnel annuel (post-refi)", "=FluxPostRefi")

# ═════════════════════════════ DÉVELOPPEMENT ═════════════════════════════════
dev = wb.create_sheet("Développement")
dev.sheet_properties.tabColor = "6B8CAA"
setwidths(dev, {"A": 42, "B": 17, "C": 56, "E": 8, "F": 8, "G": 14, "H": 14, "I": 15})

dv_dmode = mkdv(dev, ["$ / pi²", "$ / unité"])
dv_dbase = mkdv(dev, ["Hard", "Hard+Soft"])
dv_dfin = mkdv(dev, ["SCHL", "Conventionnel"])
dv_dprog = mkdv(dev, ["MLI Market", "MLI Select"])
dv_dtype = mkdv(dev, ["Standard", "Autre"])
dv_dprov = mkdv(dev, ["QC", "ON", "MB", "SK", "Autre"])
dv_damort = mkdv(dev, ["25", "30", "35", "40", "45", "50"])
dv_dterme = mkdv(dev, ["5", "10"])

d = W(dev, start=1)
d.title("Mode Développement — construction neuve")
d.note("Cellules bleues = intrants. Le take-out permanent est dimensionné comme le prêt d'acquisition "
       "(double plafond LTV / DSCR, prime selon prêt ÷ GDV).")
d.blank()
d.section("DA — Terrain")
d.inp("Prix du terrain", 300000, nm="DevTerrainPrix", fmt=F_CUR)
d.inp("Frais de fermeture terrain (%)", 2.5, nm="DevFraisTerrainPct", fmt=F_INPCT)
d.inp("Nombre d'unités", 6, nm="DevUnites", fmt=F_NUM)
d.inp("Superficie brute / unité (pi²)", 1000, nm="DevGfaUnite", fmt=F_NUM)
d.blank()
d.section("DB — Coûts directs (hard)")
d.inp("Mode de saisie", "$ / pi²", nm="DevModeHard", fmt="@", dv=dv_dmode)
d.inp("Coût / pi²", 225, nm="DevHardPi", fmt=F_CUR)
d.inp("Coût / unité", 225000, nm="DevHardUnite", fmt=F_CUR)
d.inp("Démolition", 0, nm="DevDemolition", fmt=F_CUR)
d.inp("Travaux de site", 0, nm="DevSiteWork", fmt=F_CUR)
d.blank()
d.section("DC — Coûts indirects (soft)")
d.inp("Soft divers (% des hard)", 5, nm="DevSoftPct", fmt=F_INPCT)
d.inp("Architecture & ingénierie", 40000, nm="DevArchIng", fmt=F_CUR)
d.inp("Permis & redevances", 15000, nm="DevPermis", fmt=F_CUR)
d.inp("Légal & mise en marché", 10000, nm="DevLegalMkt", fmt=F_CUR)
d.inp("Taxes foncières (construction)", 8000, nm="DevTaxesConstr", fmt=F_CUR)
d.inp("Frais de promotion (% de terrain+hard+soft)", 5, nm="DevFraisPromoPct", fmt=F_INPCT)
d.blank()
d.section("DD — Contingence")
d.inp("Contingence (%)", 10, nm="DevContingencePct", fmt=F_INPCT)
d.inp("Base de la contingence", "Hard", nm="DevBaseContingence", fmt="@", dv=dv_dbase)
d.blank()
d.section("DE — Échéancier")
d.inp("Mois de construction", 12, nm="DevMoisConstr", fmt=F_NUM)
d.inp("Mois de location (lease-up)", 6, nm="DevMoisLocation", fmt=F_NUM)
d.inp("Tirage moyen du prêt (%)", 55, nm="DevTirageMoyPct", fmt=F_INPCT)
d.blank()
d.section("DF — Financement")
d.inp("Financement construction", "Conventionnel", nm="DevFinConstr", fmt="@", dv=dv_dfin,
      note="Si conventionnel, le take-out SCHL utilise le barème « Achat ».")
d.inp("Financement take-out", "SCHL", nm="DevFinTakeout", fmt="@", dv=dv_dfin)
d.inp("Programme SCHL (take-out)", "MLI Select", nm="DevProgramme", fmt="@", dv=dv_dprog)
d.inp("Points MLI Select", 70, nm="DevPointsMLI", fmt=F_NUM)
d.inp("Type d'immeuble", "Standard", nm="DevTypeImmeuble", fmt="@", dv=dv_dtype)
d.inp("LTC cible (%)", 75, nm="DevLtcCible", fmt=F_INPCT)
d.inp("LTV cible take-out (%)", 85, nm="DevLtvCible", fmt=F_INPCT)
d.inp("Taux construction (%)", 4.5, nm="DevTauxConstr", fmt=F_INPCT)
d.inp("Taux prêt permanent (%)", 4.5, nm="DevTauxPerm", fmt=F_INPCT,
      note="Laisser vide → taux construction.")
d.inp("Terme take-out (ans)", 10, nm="DevTerme", fmt=F_NUM, dv=dv_dterme)
d.inp("Amortissement take-out (ans)", 45, nm="DevAmort", fmt=F_NUM, dv=dv_damort)
d.inp("Province", "QC", nm="DevProvince", fmt="@", dv=dv_dprov)
d.inp("Frais courtier (% du prêt constr.)", 1, nm="DevCourtierPct", fmt=F_INPCT)
d.inp("Conv. take-out — LTV (%)", 75, nm="DevLtvConvTakeout", fmt=F_INPCT)
d.inp("Conv. take-out — DSCR min.", 1.2, nm="DevDscrConvTakeout", fmt=F_X)
d.inp("SCHL — inoccupation normalisée (%)", None, nm="DevSchlInocc", fmt=F_INPCT,
      note="Laisser vide → inoccupation stabilisée (DG).")
d.inp("SCHL — dépenses normalisées (%)", None, nm="DevSchlOpex", fmt=F_INPCT,
      note="Laisser vide → dépenses stabilisées (DG).")
d.inp("SCHL — cap rate valeur économique (%)", None, nm="DevSchlCapRate", fmt=F_INPCT,
      note="Laisser vide → valeur de prêt = GDV.")
d.blank()
d.section("DG — Revenus stabilisés")
d.inp("Loyer stabilisé / unité / mois", 1600, nm="DevLoyerStab", fmt=F_CUR)
d.inp("Autres revenus annuels", 0, nm="DevAutresRevenus", fmt=F_CUR)
d.inp("Inoccupation stabilisée (%)", 5, nm="DevInoccStab", fmt=F_INPCT)
d.inp("Dépenses d'exploitation (%)", 28, nm="DevOpexStab", fmt=F_INPCT)
d.blank()
d.section("DH — Valeur, sortie & détention")
d.inp("Cap rate de sortie (%)", 5, nm="DevCapSortie", fmt=F_INPCT)
d.inp("Marge cible (valeur résiduelle, %)", 20, nm="DevMargeCible", fmt=F_INPCT)
d.inp("Commission de vente (%)", 2, nm="DevCommission", fmt=F_INPCT)
d.inp("Taux effectif gain en capital (%)", 33, nm="DevTauxGainCap", fmt=F_INPCT)
d.inp("Détention post-stabilisation (ans)", 10, nm="DevDetention", fmt=F_NUM)
d.inp("Croissance des loyers (% / an)", 2.5, nm="DevCroissanceLoyer", fmt=F_INPCT)
d.inp("Dérive du cap rate (% / an)", 0, nm="DevDeriveCap", fmt=F_INPCT)
d.inp("Taux marginal — locatif (%)", 25, nm="DevTauxMarginal", fmt=F_INPCT)
d.inp("Frais de clôture à la vente", 5000, nm="DevFraisVente", fmt=F_CUR)
d.inp("Taux du marché au bris (%)", 4.5, nm="DevNouveauTaux", fmt=F_INPCT)
d.blank(2)
d.section("Pile de coûts")
d.out("Frais de fermeture terrain", "=DevTerrainPrix*DevFraisTerrainPct/100", nm="DevFraisTerrain")
d.out("Terrain (coût total)", "=DevTerrainPrix+DevFraisTerrain", nm="DevTerrain", key=True)
d.out("Superficie brute totale (pi²)", "=DevGfaUnite*DevUnites", nm="DevGFA", fmt=F_NUM)
d.out("Construction (hard, base)",
      '=IF(DevModeHard="$ / unité",DevHardUnite*DevUnites,DevHardPi*DevGFA)', nm="DevHardBase")
d.out("Total hard", "=DevHardBase+DevDemolition+DevSiteWork", nm="DevHard", key=True)
d.out("Soft (base)", "=DevHard*DevSoftPct/100+DevArchIng+DevPermis+DevLegalMkt+DevTaxesConstr",
      nm="DevSoftBase")
d.out("Frais de promotion", "=(DevTerrain+DevHard+DevSoftBase)*DevFraisPromoPct/100", nm="DevFraisPromo")
d.out("Total soft", "=DevSoftBase+DevFraisPromo", nm="DevSoft", key=True)
d.out("Contingence",
      '=IF(DevBaseContingence="Hard+Soft",DevHard+DevSoft,DevHard)*DevContingencePct/100',
      nm="DevContingence")
d.out("Coût ex-financement", "=DevTerrain+DevHard+DevSoft+DevContingence", nm="DevCoutExFin", key=True)
d.blank()
d.section("Revenus stabilisés & GDV")
d.out("RBP stabilisé", "=DevLoyerStab*DevUnites*12", nm="DevRBPStab")
d.out("EGI stabilisé", "=DevRBPStab*(1-DevInoccStab/100)+DevAutresRevenus", nm="DevEGIStab")
d.out("NOI stabilisé", "=DevEGIStab*(1-DevOpexStab/100)", nm="DevNOIStab", key=True)
d.out("GDV = NOI ÷ cap rate", "=IF(DevCapSortie>0,DevNOIStab/(DevCapSortie/100),0)",
      nm="DevGDV", key=True)
d.out("Inoccupation normalisée SCHL (%)",
      "=IF(ISBLANK(DevSchlInocc),DevInoccStab,DevSchlInocc)", nm="DevSchlVacEff", fmt=F_INPCT)
d.out("Dépenses normalisées SCHL (%)",
      "=IF(ISBLANK(DevSchlOpex),DevOpexStab,DevSchlOpex)", nm="DevSchlOpexEff", fmt=F_INPCT)
d.out("EGI normalisé SCHL", "=DevRBPStab*(1-DevSchlVacEff/100)+DevAutresRevenus", nm="DevSchlEGI")
d.out("NOI normalisé SCHL", "=DevSchlEGI*(1-DevSchlOpexEff/100)", nm="DevSchlNOI")
d.out("Valeur économique SCHL",
      "=IF(ISBLANK(DevSchlCapRate),DevGDV,IF(DevSchlCapRate>0,DevSchlNOI/(DevSchlCapRate/100),DevGDV))",
      nm="DevValeurEcon")
d.blank()
d.section("Prêt take-out permanent (SCHL ou conventionnel)")
d.out("Taux permanent effectif (%)",
      "=IF(ISBLANK(DevTauxPerm),DevTauxConstr,IF(DevTauxPerm>0,DevTauxPerm,DevTauxConstr))",
      nm="DevTauxPermEff", fmt=F_INPCT)
d.out("Valeur de prêt (lending value)",
      "=IF(AND(DevValeurEcon>0,DevValeurEcon<DevGDV),DevValeurEcon,DevGDV)", nm="DevTkLv")
d.out("Amortissement max du programme",
      '=IF(DevProgramme="MLI Market",40,IF(DevPointsMLI>=100,50,IF(DevPointsMLI>=70,45,'
      'IF(DevPointsMLI>=50,40,25))))', nm="DevTkAmortMax", fmt=F_NUM)
d.out("Amortissement effectif",
      '=IF(DevFinTakeout="SCHL",MIN(DevAmort,DevTkAmortMax),DevAmort)', nm="DevTkAmortEff", fmt=F_NUM)
d.out("LTV max du programme (%)",
      '=IF(DevProgramme="MLI Market",85,IF(DevPointsMLI>=100,95,IF(DevPointsMLI>=70,90,85)))',
      nm="DevTkLtvMax", fmt=F_NUM)
d.out("LTV plafond retenue (fraction)",
      '=IF(DevFinTakeout="SCHL",MIN(DevLtvCible,DevTkLtvMax),DevLtvConvTakeout)/100',
      nm="DevTkLtvCap", fmt=F_PCT)
d.out("DSCR minimum requis",
      '=IF(DevFinTakeout="SCHL",IF(DevTypeImmeuble="Autre",1.2,IF(DevProgramme="MLI Select",1.1,'
      'IF(DevUnites<=6,1.1,IF(DevTerme>=10,1.2,1.3)))),DevDscrConvTakeout)', nm="DevTkDscrMin", fmt=F_X)
d.out("Taux mensuel (comp. semi-annuelle)", "=(1+DevTauxPermEff/200)^(1/6)-1",
      nm="DevTkTauxMensuel", fmt="0.000000")
d.out("Constante hypothécaire",
      "=IF(DevTkTauxMensuel=0,1/DevTkAmortEff,DevTkTauxMensuel*(1+DevTkTauxMensuel)^(DevTkAmortEff*12)"
      "/((1+DevTkTauxMensuel)^(DevTkAmortEff*12)-1)*12)", nm="DevTkConstHypo", fmt="0.0000%")
d.out("Plafond LTV (prêt max)", "=DevTkLv*DevTkLtvCap", nm="DevTkPlafondLtv")
d.out("Plafond DSCR (prêt max)",
      "=IF(DevTkConstHypo>0,DevSchlNOI/(DevTkDscrMin*DevTkConstHypo),0)", nm="DevTkPlafondDscr")
d.out("PRÊT TAKE-OUT APPROUVÉ", "=MIN(DevTkPlafondLtv,MAX(0,DevTkPlafondDscr))",
      nm="DevTkPret", key=True)
d.out("Contrainte liante", '=IF(DevTkPlafondDscr<=DevTkPlafondLtv,"DSCR","LTV")',
      nm="DevTkContrainte", fmt="General")
d.out("LTV effectif (÷ GDV)", "=IF(DevGDV>0,DevTkPret/DevGDV,0)", nm="DevTkLtvEff", fmt=F_PCT)
d.out("Colonne du barème de prime",
      '=IF(DevTypeImmeuble="Autre",2,0)+IF(DevFinConstr="Conventionnel",0,1)+1',
      nm="DevColPrime", fmt=F_NUM,
      note="Construction conventionnelle → immeuble complété au take-out → barème « Achat ».")
d.out("Taux de base de la prime",
      '=IF(DevFinTakeout="SCHL",INDEX(PrimeTable,MATCH(DevTkLtvEff*100,PrimeBornesInf,1),DevColPrime),0)',
      nm="DevTkTauxBase", fmt="0.000%")
d.out("Surcharge amortissement",
      '=IF(DevFinTakeout="SCHL",MAX(0,INT((DevTkAmortEff-25)/5))*0.0025,0)', nm="DevTkSurcharge",
      fmt="0.000%")
d.out("Rabais MLI Select",
      '=IF(AND(DevFinTakeout="SCHL",DevProgramme="MLI Select"),IF(DevPointsMLI>=100,0.3,'
      'IF(DevPointsMLI>=70,0.2,IF(DevPointsMLI>=50,0.1,0))),0)', nm="DevTkRabais", fmt="0%")
d.out("Taux net de la prime", "=MAX(0,(DevTkTauxBase+DevTkSurcharge)*(1-DevTkRabais))",
      nm="DevTkTauxNet", fmt="0.000%")
d.out("Prime SCHL nette", "=DevTkPret*DevTkTauxNet", nm="DevTkPrime")
d.out("Droits SCHL (capitalisés)", '=IF(DevFinTakeout="SCHL",DevUnites*DroitsParUnite,0)',
      nm="DevTkDroits")
d.out("Hypothèque assurée", "=DevTkPret+DevTkPrime+DevTkDroits", nm="DevTkHypoAssuree")
d.out("Paiement mensuel",
      "=IF(DevTkTauxMensuel=0,DevTkHypoAssuree/(DevTkAmortEff*12),DevTkHypoAssuree*DevTkTauxMensuel"
      "*(1+DevTkTauxMensuel)^(DevTkAmortEff*12)/((1+DevTkTauxMensuel)^(DevTkAmortEff*12)-1))",
      nm="DevTkPaiement", fmt=F_CUR2)
d.out("Service de la dette annuel", "=DevTkPaiement*12", nm="DevTkAds", fmt=F_CUR2)
d.out("DSCR (hypothèque assurée)", "=IF(DevTkAds>0,DevSchlNOI/DevTkAds,0)", nm="DevTkDscrAssure", fmt=F_X)
d.out("TVQ / taxe sur la prime", "=DevTkPrime*VLOOKUP(DevProvince,TaxeProvTable,2,FALSE)",
      nm="DevTkTaxePrime")
d.blank()
d.section("Prêt de construction, carry & TDC")
d.out("Mois de portage (constr. + lease-up)", "=DevMoisConstr+DevMoisLocation",
      nm="DevCarryMois", fmt=F_NUM)
d.out("Taux mensuel construction", "=(1+DevTauxConstr/200)^(1/6)-1", nm="DevTauxMensuelConstr",
      fmt="0.000000")
d.out("Prêt constr. préliminaire (LTC)", "=DevCoutExFin*DevLtcCible/100", nm="DevPretConstrPrelim")
d.out("Carry d'intérêts (réserve capitalisée)",
      "=DevPretConstrPrelim*(DevTirageMoyPct/100)*(DevTauxMensuelConstr*12)*DevCarryMois/12",
      nm="DevCarry")
d.out("Plafond LTC (carry inclus)", "=(DevCoutExFin+DevCarry)*DevLtcCible/100", nm="DevPlafondLtc")
d.out("Prêt de construction", "=MIN(DevPlafondLtc,DevTkPret)", nm="DevPretConstr", key=True)
d.out("Frais de courtier hypothécaire", "=DevPretConstr*DevCourtierPct/100", nm="DevFraisCourtier")
d.out("COÛT TOTAL DE DÉVELOPPEMENT (TDC)",
      "=DevCoutExFin+DevCarry+DevFraisCourtier+DevTkTaxePrime", nm="DevTDC", key=True,
      note="Droits SCHL capitalisés dans le prêt take-out.")
d.blank()
d.section("Métriques de développement")
d.out("Yield on Cost", "=IF(DevTDC>0,DevNOIStab/DevTDC,0)", nm="DevYoC", fmt=F_PCT)
d.out("Spread de développement (YoC − cap)", "=DevYoC-DevCapSortie/100", nm="DevSpread", fmt=F_PCT,
      note="≥ +1,5 % = Go ; 0 à 1,5 % = limite ; < 0 = No-go.")
d.out("Profit (GDV − TDC)", "=DevGDV-DevTDC", nm="DevProfit", key=True)
d.out("Marge sur coût", "=IF(DevTDC>0,DevProfit/DevTDC,0)", nm="DevMargeSurCout", fmt=F_PCT)
d.out("ÉQUITÉ REQUISE (TDC − prêt constr.)", "=DevTDC-DevPretConstr", nm="DevEquiteRequise", key=True)
d.out("Équité récupérée à la conversion", "=DevTkPret-DevPretConstr", nm="DevEquiteRecup")
d.out("Cash restant dans l'actif", "=DevTDC-DevTkPret", nm="DevCashRestant")
d.out("Flux opérationnel stabilisé", "=DevNOIStab-DevTkAds", nm="DevFluxStab")
d.out("Cash-on-cash (détention)", "=IF(DevCashRestant>0,DevFluxStab/DevCashRestant,0)",
      nm="DevCoC", fmt=F_PCT)
d.blank()
d.section("Vente à l'achèvement (build-to-sell)")
d.out("Produit net de vente", "=DevGDV*(1-DevCommission/100)", nm="DevVenteNette")
d.out("Profit brut avant impôt", "=DevVenteNette-DevTDC", nm="DevProfitVenteAvant")
d.out("Impôt gain en capital", "=MAX(0,DevProfitVenteAvant)*DevTauxGainCap/100", nm="DevImpotVente")
d.out("PROFIT NET APRÈS IMPÔT", "=DevProfitVenteAvant-DevImpotVente", nm="DevProfitVenteApres", key=True)
d.out("RoE (vente)", "=IF(DevEquiteRequise>0,DevProfitVenteApres/DevEquiteRequise,0)",
      nm="DevRoEVente", fmt=F_PCT)
d.blank()
d.section("Valeur résiduelle du terrain")
d.out("TDC max admissible (GDV ÷ (1+marge))",
      "=IF(DevGDV>0,DevGDV/(1+DevMargeCible/100),0)", nm="DevMaxTDC")
d.out("Coûts non-terrain",
      "=DevHard+DevSoft+DevContingence+DevCarry+DevFraisCourtier+DevTkTaxePrime",
      nm="DevCoutsNonTerrain")
d.out("Valeur résiduelle brute", "=DevMaxTDC-DevCoutsNonTerrain", nm="DevResiduelBrut")
d.out("PRIX MAX DU TERRAIN",
      "=IF(DevFraisTerrainPct>0,DevResiduelBrut/(1+DevFraisTerrainPct/100),DevResiduelBrut)",
      nm="DevResiduelTerrain", key=True)
d.blank()
d.section("Détention multi-années puis vente (refi + vente)")
d.out("Années de grille pré-stabilisation", "=MAX(1,CEILING((DevMoisConstr+DevMoisLocation)/12,1))",
      nm="DevAnneesGrille", fmt=F_NUM)
d.out("Détention effective (ans, entier)", "=MAX(1,ROUND(DevDetention,0))",
      nm="DevDetentionEff", fmt=F_NUM)
d.out("Cap rate de sortie effectif (%)", "=DevCapSortie+DevDeriveCap/100*DevDetentionEff",
      nm="DevCapSortieEff", fmt=F_INPCT)
d.out("NOI à la sortie",
      "=IF(AND(DevCapSortieEff>0,DevNOIStab>0),DevNOIStab*(1+DevCroissanceLoyer/100)^DevDetentionEff,0)",
      nm="DevNOISortie")
d.out("Prix de vente estimé", "=IF(DevCapSortieEff>0,DevNOISortie/(DevCapSortieEff/100),0)",
      nm="DevPrixVente", key=True)
d.out("Solde hypothécaire à la vente",
      "=IF(DevTkTauxMensuel=0,MAX(0,DevTkHypoAssuree-DevTkHypoAssuree/(DevTkAmortEff*12)"
      "*DevDetentionEff*12),MAX(0,DevTkHypoAssuree*(1+DevTkTauxMensuel)^(DevDetentionEff*12)"
      "-DevTkPaiement*((1+DevTkTauxMensuel)^(DevDetentionEff*12)-1)/DevTkTauxMensuel))",
      nm="DevSoldeSortie")
d.out("Commission de vente", "=DevPrixVente*DevCommission/100", nm="DevCommVente")
d.out("Pénalité de bris",
      "=IF(DevDetentionEff>=DevTerme,0,MAX(DevSoldeSortie*DevTkTauxMensuel*3,"
      "DevSoldeSortie*MAX(0,(DevTauxPermEff-DevNouveauTaux)/100)*(DevTerme-DevDetentionEff)))",
      nm="DevPenalite")
d.out("Gain brut (prix − TDC)", "=DevPrixVente-DevTDC", nm="DevGainBrut")
d.out("Impôt gain en capital", "=MAX(0,DevGainBrut)*DevTauxGainCap/100", nm="DevImpotGain")
d.out("Somme des NOI (détention)", "=SUM($G$186:$G$231)", nm="DevSommeNOI")
d.out("Capital remboursé", "=IF(DevTkHypoAssuree>0,DevTkHypoAssuree-DevSoldeSortie,0)",
      nm="DevCapitalRemb")
d.out("Intérêts totaux", "=DevTkAds*DevDetentionEff-DevCapitalRemb", nm="DevInteretsTotaux")
d.out("Impôt revenu locatif",
      "=MAX(0,DevSommeNOI-MAX(0,DevInteretsTotaux))*DevTauxMarginal/100", nm="DevImpotLocatif")
d.out("PRODUIT NET DE VENTE",
      "=DevPrixVente-DevCommVente-DevPenalite-DevSoldeSortie-DevImpotGain-DevImpotLocatif"
      "-DevFraisVente", nm="DevProduitNetVente", key=True)
d.out("Total encaissé (refi + vente, t ≥ 1)", "=SUM($I$187:$I$231)", nm="DevTotalCashRefi")
d.out("IRR (refi + vente)", '=IF(DevEquiteRequise>0,IRR($I$186:$I$231),"—")', nm="DevIRR", fmt=F_PCT)
d.out("Rendement total", "=IF(DevEquiteRequise>0,DevTotalCashRefi/DevEquiteRequise-1,0)",
      nm="DevRendTotal", fmt=F_PCT)
d.out("Multiple de sortie", "=IF(DevEquiteRequise>0,DevTotalCashRefi/DevEquiteRequise,0)",
      nm="DevMultiple", fmt=F_X)

# Dev CF vector table (rows 186..231 → t = 0..45), columns E..I
CF_HDR_ROW = 185
assert d.r < CF_HDR_ROW, f"dev calc rows ({d.r}) overlap CF table at {CF_HDR_ROW}"
for j, h in enumerate(["t", "k (année stab.)", "NOI année k", "CF opérationnel", "Flux Refi+Vente (IRR)"]):
    c = dev.cell(CF_HDR_ROW, 5 + j, h)
    c.font = Font(bold=True, size=9, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="55789A")
    c.alignment = Alignment(wrap_text=True)
for tt in range(0, 46):
    r = 186 + tt
    dev.cell(r, 5, tt).number_format = F_NUM
    dev.cell(r, 6, f'=IF($E{r}>DevAnneesGrille,$E{r}-DevAnneesGrille,"")').number_format = F_NUM
    dev.cell(r, 7, f"=IF(AND($E{r}>DevAnneesGrille,$E{r}<=DevAnneesGrille+DevDetentionEff),"
                   f"DevNOIStab*(1+DevCroissanceLoyer/100)^($E{r}-DevAnneesGrille),0)").number_format = F_NUM
    dev.cell(r, 8, f"=IF(AND($E{r}>DevAnneesGrille,$E{r}<=DevAnneesGrille+DevDetentionEff),"
                   f"G{r}-DevTkAds,0)").number_format = F_NUM
    dev.cell(r, 9, f"=IF($E{r}=0,-MAX(0,DevEquiteRequise),H{r}+IF($E{r}=DevAnneesGrille,"
                   f"DevEquiteRecup,0)+IF($E{r}=DevAnneesGrille+DevDetentionEff,"
                   f"DevProduitNetVente,0))").number_format = F_NUM

# ═════════════════════════════════ SOMMAIRE ══════════════════════════════════
from openpyxl.formatting.rule import CellIsRule

som = wb.create_sheet("Sommaire")
som.sheet_properties.tabColor = GOLD
setwidths(som, {"A": 24, "B": 15, "C": 24, "D": 15, "E": 24, "F": 15, "G": 3, "H": 44})
som.sheet_view.showGridLines = False

def som_input(r, c_lab, label, value, nm, fmt):
    som.cell(r, c_lab, label).font = Font(size=10, color="3B4754")
    cell = som.cell(r, c_lab + 1, value)
    cell.fill = INPUT_FILL; cell.border = BOX; cell.number_format = fmt
    cell.font = Font(bold=True, color="1F3A52")
    register(som, r, c_lab + 1, nm)

def som_tile(r, c, label, formula, fmt, sub_label, sub_formula, sub_fmt):
    """3-row KPI tile anchored at (r, c): label / big value (merged 2 cols) / sub pair."""
    lab = som.cell(r, c, label)
    lab.font = Font(bold=True, size=9, color="6B7682")
    som.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
    v = som.cell(r + 1, c, formula)
    v.font = Font(bold=True, size=16, color=ACCENT); v.number_format = fmt
    sl = som.cell(r + 2, c, sub_label); sl.font = Font(size=9, color="6B7682")
    sv = som.cell(r + 2, c + 1, sub_formula)
    sv.font = Font(size=9, bold=True, color="3B4754"); sv.number_format = sub_fmt
    for rr in range(r, r + 3):
        for cc in (c, c + 1):
            cur = som.cell(rr, cc).border
            som.cell(rr, cc).fill = PatternFill("solid", fgColor="FCFBF8")

# Header
som.merge_cells("A1:H1")
h = som.cell(1, 1, "Sommaire du deal — Plex Québec")
h.font = Font(bold=True, size=15, color="FFFFFF"); h.fill = HEADER_FILL
for col in range(1, 9): som.cell(1, col).fill = HEADER_FILL
som.row_dimensions[1].height = 24
som.cell(2, 1, "Propriété :").font = Font(size=10, color="3B4754")
som.merge_cells("B2:D2")
pn = som.cell(2, 2, "6-plex — Exemple")
pn.fill = INPUT_FILL; pn.border = BOX; pn.font = Font(bold=True, color="1F3A52")
register(som, 2, 2, "NomPropriete")
som.cell(2, 5, "Préparé le :").font = Font(size=10, color="3B4754")
dte = som.cell(2, 6, None)
dte.fill = INPUT_FILL; dte.border = BOX; dte.number_format = "yyyy-mm-dd"

# Key inputs (the named cells live HERE; Intrants mirrors them)
sec = som.cell(4, 1, "Intrants clés — modifiez ici, tout le classeur se recalcule")
sec.font = Font(bold=True, size=11, color=ACCENT)
for col in range(1, 9): som.cell(4, col).fill = SUBHEAD_FILL
som_input(5, 1, "Prix d'achat", 1175000, "PrixAchat", F_CUR)
som_input(6, 1, "Nombre d'unités", 6, "Unites", F_NUM)
som_input(7, 1, "LTV cible (%)", 85, "LtvCible", F_INPCT)
som_input(5, 3, "Loyer actuel / unité / mois", 1000, "LoyerActuel", F_CUR)
som_input(6, 3, "Loyer cible / unité / mois", 1450, "LoyerCible", F_CUR)
som_input(7, 3, "Dérive du cap rate (% / an)", 0.25, "DeriveCapRate", F_INPCT)
som_input(5, 5, "Taux d'intérêt (%)", 4.25, "TauxInteret", F_INPCT)
som_input(6, 5, "Période de détention (ans)", 5, "Detention", F_NUM)
nte = som.cell(8, 1, "Le loyer actuel s'applique en mode « Moyen / unité » (voir Intrants, section B). "
                     "Tous les autres intrants : onglet Intrants.")
nte.font = Font(italic=True, size=9, color="6B7682")

# KPI tiles — 2 rows × 3 tiles
som_tile(10, 1, "NOI ACTUEL", "=NOI", F_CUR, "Cap rate à l'achat", "=CapRate", F_PCT)
som_tile(10, 3, "PRÊT APPROUVÉ", "=PretApprouve", F_CUR,
         '="Contrainte "&ContrainteAcq&" · DSCR assuré"', "=DscrAssure", F_X)
som_tile(10, 5, "CASH REQUIS", "=TotalCashRequis", F_CUR, "dont à la fermeture", "=CashFermeture", F_CUR)
som_tile(14, 1, "NOI OPTIMISÉ", "=NOIOptimise", F_CUR, "Yield on Cost", "=YieldOnCost", F_PCT)
som_tile(14, 3, "PROFIT NET APRÈS IMPÔT", "=ProfitApresImpot", F_CUR, "IRR (vente)", "=IRRVente", F_PCT)
som_tile(14, 5, "PRISE DE FONDS NETTE (REFI)", "=PriseDeFondsNette", F_CUR,
         "Cash-on-cash post-refi", "=CoCPostRefi", F_PCT)
som.conditional_formatting.add("C15",
    CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(bold=True, size=16, color="1E7A3C")))
som.conditional_formatting.add("C15",
    CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, size=16, color="B03030")))

# Summary table — three side-by-side label/value tables
sec = som.cell(18, 1, "Le deal en un tableau")
sec.font = Font(bold=True, size=11, color=ACCENT)
for col in range(1, 9): som.cell(18, col).fill = SUBHEAD_FILL

def som_table(start_row, col, title, rows):
    c = som.cell(start_row, col, title)
    c.font = Font(bold=True, size=10, color="FFFFFF")
    som.cell(start_row, col).fill = PatternFill("solid", fgColor="55789A")
    som.cell(start_row, col + 1).fill = PatternFill("solid", fgColor="55789A")
    r = start_row + 1
    for label, formula, fmt, key in rows:
        lab = som.cell(r, col, label); lab.font = Font(size=10)
        v = som.cell(r, col + 1, formula); v.number_format = fmt
        v.font = Font(size=10, bold=key, color=ACCENT if key else "17222B")
        if key:
            lab.font = Font(size=10, bold=True)
            som.cell(r, col).fill = KEY_FILL; som.cell(r, col + 1).fill = KEY_FILL
        for cc in (col, col + 1):
            som.cell(r, cc).border = BOX
        r += 1

som_table(19, 1, "Acquisition & revenus", [
    ("Prix d'achat", "=PrixAchat", F_CUR, False),
    ("Prix / unité", "=PrixParUnite", F_CUR, False),
    ("Prime sur la valeur au rôle", "=PrimeSurRole", F_PCT, False),
    ("Revenu brut potentiel (RBP)", "=RBP", F_CUR, False),
    ("Revenus effectifs (EGI)", "=EGI", F_CUR, False),
    ("Dépenses d'exploitation", "=-Depenses", F_CUR, False),
    ("NOI actuel", "=NOI", F_CUR, True),
    ("Cap rate à l'achat", "=CapRate", F_PCT, False),
    ("NOI optimisé", "=NOIOptimise", F_CUR, False),
    ("Yield on Cost", "=YieldOnCost", F_PCT, False),
    ("Investissement réno + CFK", "=InvestTotal", F_CUR, False),
])
som_table(19, 3, "Financement", [
    ("Prêt approuvé", "=PretApprouve", F_CUR, True),
    ("Contrainte liante", "=ContrainteAcq", "General", False),
    ("Hypothèque assurée", "=HypoAssuree", F_CUR, False),
    ("Prime SCHL nette", "=PrimeSchl", F_CUR, False),
    ("Paiement mensuel", "=PaiementMensuel", F_CUR2, False),
    ("DSCR (prêt de base)", "=DscrSurBase", F_X, False),
    ("DSCR (hypo assurée)", "=DscrAssure", F_X, False),
    ("% de financement (/ prix)", "=LtvEffAcq", F_PCT, False),
    ("Mise de fonds", "=MiseDeFonds", F_CUR, False),
    ("Cash requis total", "=TotalCashRequis", F_CUR, True),
])
som_table(19, 5, "Sortie & rendements (an = détention)", [
    ("Prix de vente estimé", "=PrixVente", F_CUR, False),
    ("Solde hypothécaire", "=-SoldeSortie", F_CUR, False),
    ("Pénalité de bris", "=-Penalite", F_CUR, False),
    ("Profit net après impôt", "=ProfitApresImpot", F_CUR, True),
    ("IRR (vente)", "=IRRVente", F_PCT, False),
    ("Rendement annualisé", "=RendementAnnuel", F_PCT, False),
    ("Multiple de sortie", "=MultipleSortie", F_X, False),
    ("Prise de fonds nette (refi)", "=PriseDeFondsNette", F_CUR, True),
    ("CoC post-refi", "=CoCPostRefi", F_PCT, False),
    ("Équité restante (refi)", "=EquiteRestante", F_CUR, False),
])

# Key assumptions footer
sec = som.cell(33, 1, "Hypothèses clés (définies sur Intrants)")
sec.font = Font(bold=True, size=11, color=ACCENT)
for col in range(1, 9): som.cell(33, col).fill = SUBHEAD_FILL
for (rr, cc, lab, f, fmt) in [
    (34, 1, "Inoccupation (%)", "=TauxInoccupation", F_INPCT),
    (35, 1, "Dépenses d'exploitation (%)", "=DepensesPct", F_INPCT),
    (34, 3, "Croissance des loyers (%/an)", "=CroissanceLoyer", F_INPCT),
    (35, 3, "Taux de roulement (%)", "=TauxRoulement", F_INPCT),
    (34, 5, "Cap rate de sortie", "=CapSortie", F_PCT),
    (35, 5, "Terme / amortissement", '=Terme&" ans / "&AmortEffAcq&" ans"', "General"),
]:
    som.cell(rr, cc, lab).font = Font(size=9, color="6B7682")
    v = som.cell(rr, cc + 1, f); v.number_format = fmt
    v.font = Font(size=9, bold=True, color="3B4754")

# ═════════════════════════════════ LISEZ-MOI ═════════════════════════════════
lis = wb.create_sheet("Lisez-moi", 0)
lis.sheet_properties.tabColor = GOLD
setwidths(lis, {"A": 48, "B": 20, "C": 70})
l = W(lis, start=1)
l.title("Plex Québec — Modèle de sous-écriture (version Excel)")
l.blank()
l.note("Transposition fidèle de l'outil web plexquebec.com/underwriting (moteur index.html, 2026-11-22).")
l.note("Chaque formule reproduit le moteur JavaScript au dollar près — vérifié automatiquement à la génération.")
l.blank()
l.section("Nouveautés du moteur 2026-11")
for txt in ["Prêt modélisé en PHASES : si détention > terme, renouvellement à chaque fin de terme au",
            "  « Taux au renouvellement » (paiement recalculé sur l'amortissement restant).",
            "Retrait d'équité optionnel au 1er renouvellement : refi dimensionné (Section K) sur le NOI",
            "  pondéré de l'an du terme ; retrait net compté dans le TRI à cette date (bloc Tk, Analyse).",
            "Rénovations ÉTALÉES au roulement des unités (même courbe que les revenus, budget complété",
            "  à la sortie) — le cash à la fermeture exclut l'équité réno, qui sort dans les flux annuels.",
            "Financement privé tiré au prorata de chaque décaissement ; intérêts sur le solde tiré.",
            "Pénalité de bris mesurée contre le terme ACTIF (renouvelé).",
            "Réconciliation profit net ↔ vecteur IRR : cellule « ReconciliationEcart » (≈ 0).",
            "Deux pistes de croissance : loyers cibles au taux MARCHÉ, loyers en place à l'INDEXATION (TAL).",
            "Impôt locatif : forfait à la sortie (défaut) OU annuel avec DPA (cat. 1) + récupération à la",
            "  vente (« UtiliserDPA » = Oui). Simplifications : pas de demi-année, récupération complète.",
            "Rendements non levier (TRI + multiple, avant impôt) et DSCR / rendement de la dette min & moyen.",
            "Omissions assumées (vivantes dans l'outil web seulement) : bande TRI ±1/±2 % au renouvellement,",
            "  TRI selon l'année de sortie, grille loyer cible × cap de sortie."]:
    l.note(txt)
l.blank()
l.section("Structure du classeur")
for txt in ["Sommaire — le deal en un coup d'œil : intrants clés (éditables), KPI et tableau synthèse.",
            "Intrants — les 50+ variables du modèle d'acquisition (cellules bleues).",
            "Analyse — NOI, prêt SCHL/conventionnel, cash requis, scénarios Vente et Refi.",
            "Flux — projection annuelle (30 ans), vecteur IRR, cascades de sortie.",
            "Développement — mode construction neuve (TDC, GDV, take-out, valeur résiduelle).",
            "Tables SCHL — barèmes de prime, taxes provinciales, droits, paliers MLI Select."]:
    l.note(txt)
l.blank()
l.section("Conventions calibrées courtier — NE PAS « corriger »")
for txt in ["Le covenant DSCR est dimensionné sur le PRÊT DE BASE (avant prime), pas sur l'hypothèque assurée.",
            "Le DSCR sur l'hypothèque assurée (~1,06x aux défauts) est affiché à titre indicatif seulement.",
            "Le palier de prime = prêt ÷ PRIX D'ACHAT (pas la valeur de prêt).",
            "Droits SCHL : 150 $/unité, capitalisés dans le prêt total (prêt + prime + droits).",
            "Base des dépenses « % RBP » : reproduit une analyse Multi-Prêts commerciale au dollar près",
            "  (validé : 114 Chemin des Patriotes, Rouville, 2026-07 — prêt 852 615 $, prime 3,60 %, droits 900 $).",
            "Composition semi-annuelle canadienne : taux mensuel = (1 + taux/2)^(1/6) − 1.",
            "Pénalité de bris = MAX(3 mois d'intérêt, IRD × solde × années restantes au terme)."]:
    l.note(txt)
l.blank()
l.section("Valeurs de contrôle (intrants par défaut) — pour vérifier l'intégrité du modèle")
checks = [("NOI actuel", 47880, F_CUR), ("Cap rate à l'achat", 0.0407, F_PCT),
          ("Prêt approuvé (contraint DSCR)", 841061, F_CUR), ("Prime SCHL nette (4,10 %)", 34484, F_CUR),
          ("DSCR hypothèque assurée", 1.06, F_X), ("Cash requis total", 375164, F_CUR),
          ("Cap rate de sortie (5 ans)", 0.0532, F_PCT), ("Profit net après impôt", -15483, F_CUR),
          ("IRR (vente)", -0.0086, F_PCT), ("Prise de fonds nette (refi)", 159123, F_CUR)]
for lbl, v, fmt in checks:
    lis.cell(l.r, 1, lbl)
    c = lis.cell(l.r, 2, v); c.number_format = fmt; c.font = Font(bold=True, color=ACCENT)
    l.r += 1
l.blank()
l.section("Synchronisation")
l.note("Le modèle web (index.html + port React /tools/underwriter) fait autorité. Toute modification du")
l.note("moteur web doit être reportée ici (et inversement), puis revalidée contre la fiche Multi-Prêts.")

wb._sheets = [wb[s] for s in
              ["Sommaire", "Lisez-moi", "Intrants", "Analyse", "Flux", "Développement", "Tables SCHL"]]
wb.active = wb["Sommaire"]
OUT = "/Users/julienberube/Documents/Claude/PlexQuebec/underwriting/Plex_Quebec_Underwriting_Model.xlsx"
wb.save(OUT)
import json, os
json.dump(NAMES, open(os.path.join(os.path.dirname(__file__), "names.json"), "w"))
print("saved", OUT, "·", len(NAMES), "named cells")
