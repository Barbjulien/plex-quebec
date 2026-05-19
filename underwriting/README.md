# Plex Québec — Outil d'analyse d'acquisition

Outil interne d'analyse d'investissement pour l'acquisition de plex au Québec. Accessible directement dans le navigateur, sans serveur ni installation.

## Fichiers

| Fichier | Description |
|---|---|
| `underwriter.html` | Application complète — ouvrir dans un navigateur |

## Utilisation

Ouvrir `underwriter.html` directement dans Chrome ou Safari. Aucune connexion internet requise une fois la page chargée (sauf pour les polices Google Fonts au premier chargement).

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
- **Plafond LTV** = Valeur × LTV effective (min. entre cible et max programme)
- **Plafond DSCR** = NOI ÷ (DSCR minimum × constante hypothécaire)

La **contrainte liante** (LTV ou DSCR) est affichée en évidence.

#### Prime SCHL
```
Prime nette = Prêt × (taux_base + surcharge_amort) × (1 − rabais_MLI_Select)
```
- Surcharge amortissement : +0.25% par tranche de 5 ans au-delà de 25 ans
- Rabais MLI Select : 10% (50–69 pts), 20% (70–99 pts), 30% (100+ pts)
- TVQ sur prime : 9.975% au Québec

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
- Gain en capital imposable : 50% d'inclusion (règle canadienne)
- Rendement annualisé : `(1 + rendement_total)^(1/années) − 1`
- Multiple de sortie : `(mise de fonds + profit net) ÷ mise de fonds`

---

### Valeurs de référence (données par défaut)

Les valeurs par défaut correspondent à l'exemple du fichier `Plex_Underwriter.xlsx` situé dans `/Documents/CRE Models/` :

| Métrique | Valeur |
|---|---|
| Prix d'achat | 1 175 000 $ |
| Unités | 6 |
| NOI actuel | ~47 880 $ |
| Cap rate à l'achat | ~4.07% |
| Prêt approuvé (SCHL) | ~711 667 $ (contraint par DSCR) |
| Prime SCHL nette | ~23 841 $ |
| DSCR sur hypo assurée | ~1.26x |
| Cash requis total | ~495 086 $ |
| Profit net après impôt (5 ans) | ~263 836 $ |
| Rendement annualisé | ~8.9% |
| Multiple de sortie | ~1.53x |
| Prise de fonds nette (refi) | ~427 177 $ |

---

## Branding

L'outil reprend l'identité visuelle de Plex Québec :
- Polices : **Fraunces** (titres) + **Inter Tight** (UI)
- Palette : fond crème `#f7f7f5`, accent acier `#3a5c7a`, or `#b8893a`
- Navigation collante avec flou de fond, identique à `index.html`

## Lien avec le modèle Excel

Ce fichier HTML est la transposition web de `Plex_Underwriter.xlsx` (`/Documents/CRE Models/`). Toutes les formules ont été reproduites fidèlement en JavaScript vanilla, y compris la composition semi-annuelle canadienne, les grilles de primes SCHL, et la pondération des revenus par roulement des locataires.
