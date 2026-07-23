# GaylShapleyXLS — Algorithme de Gale-Shapley en VBA Excel

Implémentation de l'**algorithme de Gale-Shapley** (affectation stable) pour assigner des élèves à des projets directement dans Excel via des macros VBA.

---

## Feuilles Excel attendues

| Feuille | Rôle |
|---|---|
| `Préférences_Élèves` | Listes de vœux ordonnés des élèves |
| `Préférences_Projets` | Listes de classement des élèves par projet + capacités Min/Max |
| `Résultats` | Résultat de l'affectation |
| `Rapport_Satisfaction` | Rapport d'analyse de satisfaction |
| `Log_Affectation` | Journal pas à pas de l'algorithme |
| `Bilan_Performance` | Métriques agrégées |
| `Details_Suivi` | Listes de suivi (non affectés, sous-minimum) |
| `Affectations_par_Projet` | Rapport croisé par projet |
| `Dashboard` | Tableau de bord synthétique |

---

## Description des macros

### 1. `GenererDonneesDeTest`
Génère aléatoirement des données de test dans les feuilles de préférences. Utilise `System.Collections.ArrayList` pour un **Fisher-Yates shuffle** (mélange aléatoire). Demande à l'utilisateur le nombre d'élèves et de projets via `InputBox`.

### 2. `AffectationElevesProjets`
Implémentation principale de **Gale-Shapley côté projets** (les projets choisissent). La logique :
- Les élèves célibataires proposent à leurs projets préférés
- Le projet accepte provisoirement si une place est disponible
- Si plein, le projet compare le nouvel arrivant au "pire" élève actuel et l'évince si meilleur
- Termine quand plus aucun élève n'est célibataire

> ⚠️ **Variante non standard** : ce n'est pas le Gale-Shapley classique où l'élève parcourt sa liste progressivement — ici l'élève boucle sur **toute** sa liste d'un coup dans un `For j`. L'algorithme peut ne pas converger si un élève évincé reprend au début.

### 3. `AffectationPasAPas`
Même algorithme que `AffectationElevesProjets` mais avec journalisation détaillée dans `Log_Affectation` à chaque étape. Utilise `propositionsFaites` pour tracker le prochain vœu à soumettre. C'est la version la plus fidèle au Gale-Shapley classique.

### 4. `RapportSatisfactionEleves` / `CreerRapportSatisfaction`
Deux versions d'un même rapport : calcule pour chaque élève le rang du projet obtenu dans sa liste de vœux, puis produit la **moyenne** et l'**écart-type** des rangs.

$$s = \sqrt{\frac{\sum(x_i - \bar{x})^2}{n-1}}$$

*(écart-type d'échantillon)*

### 5. `CreerRapportProjets`
Rapport croisé par projet : calcule séparément la satisfaction des élèves et la "qualité d'équipe" du point de vue du projet, avec moyenne et écart-type pour chaque.

### 6. `BilanPerformanceAlgorithme`
Métriques globales :
- Taux d'affectation
- Répartition des vœux obtenus (1er, 2ème, 3ème)
- Taux de projets ayant atteint leur minimum de capacité
- Taux d'occupation global

Génère également la feuille `Details_Suivi` et appelle `CreerDashboard`.

### 7. Macros de diagnostic
Outils de débogage développés de manière itérative :

| Macro | Rôle |
|---|---|
| `DiagnostiqueAffectation` | Vérifie la lecture des préférences dans le log |
| `VerificationFinale_V2` | Affiche les données brutes pour "Élève 2" |
| `DebugRang` | Trace le calcul du rang pour "Élève 2" pas à pas |
| `VerifierToutesLesDonnees` | Relit toutes les données et affiche les rangs calculés |
| `TestArrayList` | Vérifie la disponibilité de `System.Collections.ArrayList` |

---

## Flux d'exécution recommandé

```
GenererDonneesDeTest
        ↓
AffectationPasAPas  (ou AffectationElevesProjets)
        ↓
CreerRapportSatisfaction
        ↓
CreerRapportProjets
        ↓
BilanPerformanceAlgorithme  →  Dashboard
```

---

## Problèmes connus / points d'attention

| # | Problème | Impact |
|---|---|---|
| 1 | **`Option Explicit` déclaré deux fois** (début du fichier et ~ligne 250) | Erreur de compilation |
| 2 | **`AffectationElevesProjets` : boucle `For j` non standard** — l'élève parcourt toute sa liste en une passe, sans mémoriser où il en était | Non-convergence potentielle si un élève évincé reprend depuis le début |
| 3 | **`System.Collections.ArrayList`** utilisé dans `GenererDonneesDeTest` — nécessite `.NET Framework`, non disponible sur toutes les machines | Erreur à l'exécution (`TestArrayList` permet de diagnostiquer) |
| 4 | **Macros de diagnostic codées en dur** sur "Élève 2" (`DebugRang`, `VerificationFinale_V2`) | Code de débogage à nettoyer avant livraison |
| 5 | **`CreerDashboard`** appelée dans `BilanPerformanceAlgorithme` mais non définie dans ce fichier | Erreur d'exécution si absente d'un autre module VBA |
| 6 | **Pas de gestion d'erreur** si un élève n'existe pas dans le classement d'un projet (`projetsPrefs(projetVise)(eleveActuel)` peut lever une erreur) | Runtime error possible sur données incomplètes |

---

## Prérequis

- Microsoft Excel avec macros activées
- `.NET Framework` installé (pour `System.Collections.ArrayList`)
- Feuilles listées ci-dessus créées au préalable (sauf `Details_Suivi` et `Verification_Finale` créées automatiquement)
