# Plex Québec — Outil d'analyse d'acquisition

Outil interne d'analyse d'investissement pour l'acquisition de plex au Québec. Accessible directement dans le navigateur, sans serveur ni installation.

## Fichiers

| Fichier | Description |
|---|---|
| `index.html` | Application complète — ouvrir dans un navigateur |

## Utilisation

Ouvrir `index.html` directement dans Chrome ou Safari. Aucune connexion internet requise une fois la page chargée (sauf pour les polices Google Fonts au premier chargement).

---

## Ce que l'outil calcule

### Structure générale

L'outil est organisé en 4 onglets :

- **Paramètres** — Saisie des 47 variables du modèle
- **Analyse** — Sorties détaillées par section
- **Flux** — Projection des flux de trésorerie année par année
- **Guide** — Notes, benchmarks et tests de stress

Un bandeau de métriques clés est toujours visible en haut de page (NOI, cap rate, prêt approuvé, cash requis, profit net).

---

### Sections d'entrée

| Section | Variables clés |
|---|---|
| A — Propriété & Acquisition | Prix d'achat, unités, superficie, valeur au rôle, frais d'acquisition |
| B — Revenus actuels | Loyer actuel, inoccupation, autres revenus, dépenses d'exploitation |
| C — Revenus optimisés | Loyer cible, croissance annuelle, période de détention, taux de roulement |
| D — Rénovations & Cash for keys | Coût par unité, % unités à rénover, mois de loyer CFK |
| E — Financement SCHL | Programme (MLI Market / MLI Select), LTV, taux, terme, amortissement |
| F — Financement privé | % financé, taux, durée |
| G — Réserves & CapEx | Provision mensuelle par unité |
| H — Scénario de vente | Dérive du cap rate, commission, fiscalité |
| K — Scénario de refinancement | Paramètres SCHL indépendants pour le refi |

---

### Logique de calcul

#### Composition semi-annuelle (obligatoire au Canada)
```
Taux mensuel = (1 + taux_annuel / 2)^(1/6) - 1
```

#### Double plafond SCHL
Le prêt approuvé est le **minimum** entre :
- **Plafond LTV** = Valeur de prêt SCHL (valeur économique si fournie, sinon prix) × LTV effective (min. entre cible et max programme)
- **Plafond DSCR** = NOI ÷ (DSCR minimum × constante hypothécaire)

Conforme à la pratique des prêteurs (calibré sur une analyse Multi-Prêts commercial, 2026-07) : le covenant DSCR est dimensionné sur le **prêt de base** (avant prime). Le DSCR sur l'hypothèque assurée (prime capitalisée) est affiché à titre indicatif et tombe sous le minimum (~1.06 vs 1.10). La **contrainte liante** (LTV ou DSCR) est affichée en évidence.

#### Prime SCHL
```
Prime nette = Prêt × (taux_base + surcharge_amort) × (1 − rabais_MLI_Select)
```
- Palier de taux de base déterminé par **prêt ÷ prix d'achat** (pratique courtier confirmée)
- Surcharge amortissement : +0.25% par tranche de 5 ans au-delà de 25 ans
- Rabais MLI Select : 10% (50–69 pts), 20% (70–99 pts), 30% (100+ pts)
- LTV max MLI Select : 85% (50–69 pts), 90% (70–99 pts), 95% (100+ pts)
- Taxe provinciale sur prime : QC 9.975%, ON 8%, MB 7%, SK 6%
- Droits SCHL : 150 $/unité, capitalisés dans le prêt total (prêt + prime + droits)

#### Revenus pondérés (roulement des locataires)
La transition des loyers actuels vers les loyers cibles est modélisée par une série géométrique basée sur le taux de roulement annuel :
```
NOI pondéré année y = NOI_opt × (1 − (1 − tr)^y) + NOI_actuel × (1 − tr)^y
```

#### Pénalité de bris d'hypothèque
```
Pénalité = MAX(3 mois d'intérêt, IRD × solde × années restantes au terme)
```

#### Rendements (scénario de vente)
- Gain en capital imposable : gain brut moins frais de vente (commission + fermeture), au taux effectif saisi
- Revenu locatif imposable : revenus pondérés moins intérêts (hypothèque + financement privé)
- Le **profit net après impôt** se réconcilie exactement avec la somme du vecteur de flux utilisé pour l'IRR (frais de financement déduits, capital remboursé crédité)
- Rendement annualisé : `(1 + rendement_total)^(1/années) − 1`
- Multiple de sortie : `(mise de fonds + profit net) ÷ mise de fonds`
- Les frais de fermeture (vente et refi) se saisissent en **valeurs positives**

---

### Valeurs de référence (données par défaut)

Métriques produites par les valeurs par défaut de l'outil (mises à jour 2026-08-06 — le modèle a divergé de l'exemple `Plex_Underwriter.xlsx` d'origine : DSCR min 1.10 pour 5–6 unités selon la fiche Standard Locatif, dimensionnement calibré sur la pratique courtier, dérive du cap rate 0.25%/an) :

| Métrique | Valeur |
|---|---|
| Prix d'achat | 1 175 000 $ |
| Unités | 6 |
| NOI actuel | ~47 880 $ |
| Cap rate à l'achat | ~4.07% |
| Prêt approuvé (SCHL) | ~841 100 $ (contraint par DSCR, 1.10 sur prêt de base) |
| Prime SCHL nette (4.10%) | ~34 500 $ |
| DSCR sur hypo assurée | ~1.06x (indicatif) |
| Cash requis total | ~376 100 $ |
| Cap rate de sortie (5 ans) | ~5.32% |
| Profit net après impôt (5 ans) | ~−13 000 $ (≈ seuil de rentabilité) |
| IRR (vente) | ~−0.7% |
| Multiple de sortie | ~0.97x |
| Prise de fonds nette (refi) | ~159 000 $ |

Le scénario par défaut est volontairement neutre : c'est l'écart entre loyer actuel (1 000 $) et loyer cible (1 450 $) et les hypothèses de sortie qui créent (ou détruisent) la valeur.

---

## Branding

L'outil reprend l'identité visuelle de Plex Québec :
- Polices : **Fraunces** (titres) + **Inter Tight** (UI)
- Palette : fond crème `#f7f7f5`, accent acier `#3a5c7a`, or `#b8893a`
- Navigation collante avec flou de fond, identique à `index.html`

## Lien avec le modèle Excel

Ce fichier HTML était à l'origine la transposition web de `Plex_Underwriter.xlsx` (`/Documents/CRE Models/`). Le modèle web a depuis divergé et fait autorité : DSCR minimum selon la fiche SCHL Standard Locatif (1.10 achat / 1.20 refi pour 5–6 unités), dimensionnement calibré sur la pratique des courtiers commerciaux (validé au dollar près contre une analyse Multi-Prêts de 2026-07), et réconciliation exacte entre le profit net et le vecteur de flux IRR. Le fichier Excel n'a pas été mis à jour en conséquence.

**Note conventions courtier** : les analyses de courtiers expriment souvent les dépenses en % du revenu brut potentiel (RBP). Le champ « Base des dépenses » (section B) permet de choisir % EGI ou % RBP — en mode RBP, l'outil reproduit une analyse Multi-Prêts au dollar près (prêt total = prêt + prime + droits capitalisés, paiement identique).

## DCF multi-terme : renouvellement & retrait d'équité (2026-11-21)

Le moteur projette désormais le prêt en **phases** via `buildProjection()` (source unique des flux annuels, consommée par `recalc()` et `updateCashFlow()`). Quand la détention dépasse le terme, le prêt est renouvelé à chaque fin de terme au « Taux au renouvellement » (paiement recalculé sur l'amortissement restant — à taux inchangé, l'échéancier original est reproduit exactement). La case « Retrait d'équité au renouvellement » transforme le premier renouvellement en refinancement dimensionné par la Section K ; la prise de fonds nette entre dans le vecteur TRI à la date du retrait (le financement privé est remboursé à même les produits du refi et sa rampe d'intérêts s'arrête là). La pénalité de bris à la sortie est mesurée contre le terme **actif** (renouvelé). Une bande de sensibilité TRI (±1/±2 % sur le taux de l'événement de renouvellement — refi inclus) s'affiche dans l'analyse de vente. Simplification assumée : la prime SCHL du refi est calculée sur le plein prêt, pas seulement sur la majoration. La réconciliation profit net ↔ vecteur TRI reste exacte (vérifiée au cent près, retrait inclus).

## Rénovations étalées au roulement des unités (2026-11-22)

Le décaissement des rénovations suit désormais la **même courbe de roulement que les revenus** (fraction convertie à l'année y = 1 − (1 − roulement)^y) : la dépense réno de l'année y est proportionnelle aux unités qui roulent cette année-là, normalisée pour que le budget soit entièrement déployé à la sortie. Conséquences :

- **Cash à la fermeture** = mise de fonds + frais d'acquisition + TVQ prime + courtier + part équité du CFK (le CFK se paie à la clôture — il achète le roulement). L'équité réno sort dans les **flux annuels** (colonne « Réno (équité) » du tableau de flux) et entre dans le TRI aux dates réelles de décaissement (`venteCF[0]` = cash à la fermeture seulement).
- **Financement privé** : tiré au prorata de chaque décaissement (part CFK à la clôture, part réno au fil des travaux) ; les intérêts courent sur le **solde réellement tiré** (convention mi-année pour le tirage de l'année) — remplace l'ancienne rampe linéaire (y − 0,5)/détention. Au retrait d'équité, c'est le solde tiré à cette date qui est remboursé (pas le montant engagé total) ; les décaissements post-retrait sont 100 % équité. À la vente comme au refi de Section K, c'est le solde réellement dû à la sortie qui est remboursé (zéro après un retrait — corrige un double-comptage).
- **Cash-on-cash annuel** : le flux de chaque année déduit l'équité réno décaissée cette année-là — les premières années montrent le vrai drag de la phase travaux. **Yield on cost** : inchangé par construction (dénominateur = coût complet du projet, indépendant du calendrier) ; le port React utilise désormais le même dénominateur « all-in » que l'outil HTML (achat + frais acq + réno + courtier + frais SCHL).
- `totalCash` (« équité totale engagée », dénominateur du rendement net / multiple) garde sa définition ; la Section H affiche séparément le cash à la fermeture et l'équité réno étalée.

Réconciliation profit net ↔ Σ vecteur TRI vérifiée exacte (5 scénarios : base, renouvellement, retrait, CFK, privé 50 %) et parité HTML ↔ React au cent près (lignes annuelles, vecteur TRI, retrait inclus) — scripts de capture Playwright + tsx. L'invariant « renouvellement à taux inchangé reproduit l'échéancier original » tient toujours.

## Améliorations inspirées du modèle A.CRE (2026-11-22)

Après revue du modèle A.CRE All-in-One (multifamily), six améliorations retenues pour le contexte plex résidentiel québécois — appliquées aux trois moteurs :

1. **Deux pistes de croissance des loyers** : les loyers CIBLES croissent au taux marché (« Croissance loyers marché »), les loyers EN PLACE à l'« Indexation loyers en place » (TAL). Si indexation < marché, l'écart au marché (loss-to-lease) s'élargit avec le temps au lieu de rester constant. Défauts égaux (2 %/2 %) → aucun changement au comportement historique.
2. **TRI selon l'année de sortie** (outil web seulement) : vente simulée à la fin de chaque année — NOI blendé de l'année, cap dérivé, pénalité contre la phase active, impôts et réno décaissée cumulés ; à l'année d'un retrait d'équité, la vente REMPLACE le refi. Montre le moment optimal de sortie et les falaises de pénalité en cours de terme.
3. **Rendements non levier** : TRI + multiple avant impôt, sans dette ni financement privé (t0 = prix + frais acq + CFK ; années = NOI blendé − réserves − réno décaissée ; sortie = prix − commission − fermeture). Sépare la qualité de l'actif du montage financier.
4. **Impôt annuel avec DPA (CCA)** — toggle, défaut OFF (forfait à la sortie, comportement historique) : impôt locatif payé annuellement, réduit par la DPA (cat. 1 dégressive sur la part bâtiment + rénos capitalisées l'année du décaissement, plafonnée au revenu de l'année — pas de perte), récupération d'amortissement (Σ DPA) imposée au taux marginal à la vente. Le profit net total bouge peu ; c'est le TRI qui capte le report d'impôt. Simplifications documentées : pas de demi-année, récupération complète, pas de report de pertes. Intrants : taux DPA (4 %), % bâtiment (80 %).
5. **DSCR et rendement de la dette min/moyen** sur la détention (NOI blendé, phase-aware) avec l'année du minimum — le chiffre que le prêteur stresse.
6. **Grille de sensibilité TRI** (outil web seulement) : loyer cible (±10 %) × cap de sortie (±0,5 pt), 5×5, cellule de base ancrée sur le TRI du cas central.

Vérification : parité HTML ↔ React au cent près sur 6 scénarios (base, renouvellement, retrait, DPA, DPA+retrait, indexation 1 %/marché 3 %) ; réconciliation profit net == Σ vecteur TRI exacte dans les DEUX modes d'impôt ; les défauts reproduisent les chiffres d'avant le changement à l'identique.

**Modèle Excel `Plex_Quebec_Underwriting_Model.xlsx` : À JOUR (2026-11-22).** `excel_model/` (engine.py / build_xlsx.py / verify.py) a été reporté sur le moteur courant : l'onglet Flux modélise le prêt en **phases** (colonnes début/taux/amort/terme de phase, récurrence de solde, renouvellement à chaque fin de terme au « Taux au renouvellement »), un bloc « Tk » sur Analyse dimensionne le refi de retrait sur le NOI pondéré de l'an du terme (retrait net dans le vecteur IRR, colonne AD), et les rénos sont étalées au roulement (colonnes réno/tirage privé/solde privé/équité réno ; intérêts privés sur le solde tiré). Nouveaux intrants : « Taux au renouvellement » et « Retrait d'équité au renouvellement » (Intrants, section C) ; nouvelles cellules nommées : `CashFermeture`, `EquiteRenoEtalee`, `SoldePriveSortie`, bloc `Tk*`, `ReconciliationEcart` (Σ flux − profit net ≈ 0, vivante dans le classeur). La pénalité de bris est mesurée contre le terme actif. Le classeur porte aussi les améliorations A.CRE : deux pistes de croissance (TauxIndexation), mode DPA (colonnes AM–AQ de Flux : revenu imposable, pool dégressif, DPA plafonnée, impôt annuel ; `DpaActif`/`DpaTotale`/`Recuperation`/`ImpotLocatifTotal` sur Analyse), flux non levier (colonne AR + `IRRNonLevier`/`MultipleNonLevier`) et couverture min/moyenne (`DscrMin`/`DscrMinAnnee`/`DscrMoyen`, `RendDetteMin`/`RendDetteMinAnnee`/`RendDetteMoyen` via colonnes filtrées AK/AL). **Vérifié : 160 cellules nommées × 7 scénarios (défauts, branches RBP/MLI, conventionnel, renouvellement, retrait, DPA, indexation) = 1 120/1 120 contre l'oracle engine.py**, lui-même validé à l'identique contre le moteur JS (captures Playwright, au cent près). Omissions assumées (outil web seulement) : bande TRI ±1/±2 % au renouvellement, TRI selon l'année de sortie, grille loyer × cap. Pour mettre à jour après un changement de moteur : éditer `engine.py` + `build_xlsx.py`, puis `python3 build_xlsx.py && python3 verify.py`.
