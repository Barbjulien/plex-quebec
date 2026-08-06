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
