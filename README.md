# GaylShapleyXLS — Affectation d'équipes aux projets par l'algorithme de Gale-Shapley

Implémentation en **VBA Excel** de l'algorithme de Gale-Shapley pour affecter des **équipes d'élèves** à des projets de manière stable et optimale selon les préférences de chacun.

> **État du projet ✅** : Code compilé et fonctionnel sur la branche `dev`, avec dashboard KPI métier coloré.

---

## Contexte et problème résolu

Dans un contexte pédagogique (école d'ingénieurs, université), des élèves se regroupent en équipes et doivent être affectés à des projets. Chaque équipe classe les projets par ordre de préférence, et chaque projet classe les équipes candidates. Les projets ont des contraintes de capacité (nombre minimum et maximum d'équipes acceptées, fourchette de taille d'équipe acceptable).

L'objectif est de produire une **affectation stable** : il n'existe aucune paire (équipe, projet) où l'équipe préférerait ce projet à son affectation actuelle ET le projet préférerait cette équipe à au moins une équipe déjà affectée.

---

## L'algorithme de Gale-Shapley

### Principe général

L'algorithme de Gale-Shapley (1962, prix Nobel d'économie 2012) résout le **problème d'affectation stable**. Dans sa forme originale, il s'applique à des couples (mariage stable). Ici, il est adapté à un contexte *many-to-one* avec capacités : plusieurs équipes peuvent être affectées à un même projet.

### Adaptation utilisée : variante "équipes-proposantes"

Dans cette implémentation, ce sont les **équipes** qui font les propositions (et non les projets). Cela produit une affectation **optimale du côté des équipes** : aucune équipe ne peut obtenir un meilleur résultat dans aucune affectation stable.

### Déroulement pas à pas

**Initialisation :**
- Toutes les équipes sont « libres » (non affectées).
- Chaque équipe possède un compteur `propositionsFaites` initialisé à 0.
- Les projets incompatibles avec la taille d'une équipe sont **retirés au préalable** de sa liste de vœux.

**Itérations (tant qu'il reste des équipes libres) :**

1. On prend la première équipe libre de la file d'attente.
2. Elle propose au prochain projet de sa liste (celui qu'elle n'a pas encore sollicité).
3. **Cas A — Le projet a de la place** (nombre d'équipes acceptées < capacité max) :
   - Le projet accepte l'équipe **provisoirement**.
   - L'équipe sort de la file des libres.
4. **Cas B — Le projet est plein** :
   - Le projet compare le rang de la nouvelle équipe avec celui de la **pire équipe** actuellement acceptée.
   - **Si la nouvelle équipe est mieux classée** : la pire équipe est **évincée** (elle retourne dans la file des libres), et la nouvelle équipe est acceptée provisoirement.
   - **Si la nouvelle équipe est moins bien classée** : elle est **rejetée** et retourne en fin de file pour tenter son prochain vœu.
5. Si une équipe a épuisé toute sa liste de vœux compatibles, elle reste **non affectée**.

**Terminaison :**
L'algorithme se termine nécessairement car à chaque proposition, une équipe avance dans sa liste (elle ne reviendra jamais sur un projet déjà refusé ou quitté). La liste étant finie, la convergence est garantie.

### Propriété de stabilité

À la fin, l'affectation est **stable** : si une équipe E préfère un projet P à son projet actuel, alors P a déjà rencontré E et l'a rejetée au profit d'équipes qu'il classe mieux. P ne voudra donc jamais prendre E au détriment de ses affectés.

Cette stabilité est celle de Gale-Shapley : elle garantit l'absence de paire bloquante entre une équipe et un projet, mais elle ne garantit pas à elle seule que chaque projet atteigne son nombre minimum d'équipes.

Dans l'état actuel du projet, ce point n'est pas pénalisant : le choix métier est de favoriser les demandes des étudiants, donc les préférences des équipes priment sur le remplissage minimal des projets. Les minima restent affichés et suivis, mais ils ne bloquent pas encore l'affectation.

### Agrégation des préférences (score de Borda)

Les préférences d'une équipe sont calculées à partir des préférences individuelles de ses membres via un **score de Borda** :

1. Le 1er choix d'un membre rapporte N points, le 2e choix N-1 points, ..., le dernier choix 1 point.
2. Les points sont additionnés pour chaque projet.
3. Les projets sont ensuite triés par score décroissant, avec départage par indice le plus bas en cas d'égalité.

### Classement des équipes par les projets

Pour les données de test, le rang d'une équipe auprès d'un projet combine la **moyenne des rangs** et leur **dispersion** dans la liste individuelle des membres :

$$\text{score}(\text{équipe}, \text{projet}) = \text{moyenne} + \text{variance}$$

Un score **faible** = équipe enthousiaste et homogène = bien classée par le projet. Les équipes sont ensuite triées par score croissant pour obtenir le classement du projet.

---

## Structure du classeur Excel

### Feuilles alimentées automatiquement

| Feuille | Rôle | Créée/alimentée par |
|---|---|---|
| `Préférences_Élèves` | Listes de vœux ordonnés de chaque élève | `GenererDonneesDeTest` |
| `Équipes` | Composition des équipes (membres) | `GenererDonneesDeTest` |
| `Préférences_Équipes` | Préférences agrégées de chaque équipe (score de Borda) | `GenererDonneesDeTest` |
| `Préférences_Projets` | Capacités + classement des équipes par chaque projet | `GenererDonneesDeTest` |
| `Résultats` | Équipes affectées à chaque projet + statut capacité | Macros 2 ou 3 |
| `Log_Affectation` | Journal détaillé étape par étape | `AffectationEquipesPasAPas` |
| `Rapport_Satisfaction` | Rang du vœu obtenu par chaque équipe + statistiques | `CreerRapportSatisfaction` |
| `Bilan_Performance` | Métriques globales de l'affectation | `BilanPerformanceAlgorithme` |
| `Details_Suivi` | Listes équipes sans projet, projets sous-minimum | `BilanPerformanceAlgorithme` |
| `Affectations_par_Projet` | Vue détaillée projet → équipes → membres | `RemplirAffectationsParProjet` |

> Les feuilles manquantes sont **créées automatiquement** à la première exécution de chaque macro.

### Structure de `Préférences_Projets`

```
| Projet | MinEquipes | MaxEquipes | TailleMinEquipe | TailleMaxEquipe | Équipe 1 | Équipe 2 | ...
|--------|------------|------------|-----------------|-----------------|----------|----------|
| Proj A |     1      |     3      |        2        |        4        |    2     |    1     | ...
```

- **MinEquipes / MaxEquipes** : nombre d'équipes que le projet doit/peut accueillir. Le minimum est contrôlé dans les bilans, mais il n'est pas encore imposé comme contrainte bloquante dans l'affectation.
- **TailleMinEquipe / TailleMaxEquipe** : fourchette de taille d'équipe acceptable. Une équipe hors fourchette est automatiquement exclue de ce projet.
- Les colonnes suivantes contiennent le **rang** de chaque équipe (1 = équipe préférée).

---

## Description des macros

### 0. `InitialiserFeuillesFormulaires`
Prépare le classeur pour un mode de collecte **via formulaires en ligne** :
- crée (ou réinitialise) `Reponses_Eleves`
- crée (ou réinitialise) `Reponses_Commanditaires`
- crée `Instructions_Formulaires` avec le mode d'emploi

Les deux feuilles `Reponses_*` servent de destination de collage des exports CSV (Microsoft Forms, Google Forms, etc.).

### 0 bis. `GenererDonneesDepuisFormulaires`
Construit les feuilles d'entrée de l'algorithme à partir des réponses :
- lit `Reponses_Eleves` (équipe, élève, choix de projets)
- lit `Reponses_Commanditaires` (projet, capacités, tailles d'équipe)
- alimente automatiquement `Préférences_Élèves`, `Équipes`, `Préférences_Équipes`, `Préférences_Projets`

Le classement des équipes par projet reste calculé automatiquement avec la même logique que les données de test (moyenne + dispersion des rangs individuels).

### 0 ter. `InitialiserConfigurationImportFormulaires`
Prépare la feuille `Config_Import_Formulaires` qui contient les chemins des 2 fichiers Excel exportés depuis Microsoft Forms :
- fichier réponses élèves
- fichier réponses commanditaires

### 0 quater. `ImporterReponsesFormulairesAutomatique`
Importe automatiquement les réponses depuis ces fichiers Excel **sans copier-coller**, puis déclenche `GenererDonneesDepuisFormulaires`.

Si les chemins ne sont pas renseignés (ou invalides), la macro ouvre un sélecteur de fichier pour les choisir.

### 1. `GenererDonneesDeTest`
Génère l'ensemble des données de test en 4 étapes :
1. Préférences individuelles aléatoires de chaque élève (mélange Fisher-Yates).
2. Formation aléatoire des équipes (taille variable dans la fourchette saisie).
3. Calcul des préférences des équipes par **score de Borda** à partir des préférences individuelles.
4. Calcul des classements des projets (basés sur la moyenne des rangs individuels et leur dispersion).

Paramètres saisis via `InputBox` : nombre d'élèves, nombre de projets, taille min et max des équipes.

> Les feuilles `Préférences_Élèves`, `Équipes`, `Préférences_Équipes`, `Préférences_Projets` sont **créées automatiquement** si elles n'existent pas.

### 2. `AffectationEquipesProjets`
Exécute l'algorithme de Gale-Shapley complet en une passe, sans journalisation. Plus rapide que la version pas à pas.

Écrit les résultats dans `Résultats` :
- Équipes affectées à chaque projet
- Statut capacité (OK / MINIMUM NON ATTEINT)

### 3. `AffectationEquipesPasAPas`
Même algorithme avec journalisation détaillée de chaque étape dans `Log_Affectation` :
- Action de l'équipe (proposition)
- Décision du projet (acceptation provisoire, éviction, rejet)
- État courant du projet
- Liste des équipes encore libres

Produit les mêmes résultats que la macro 2 dans `Résultats`.

### 4. `CreerRapportSatisfaction`
Pour chaque équipe affectée, calcule le **rang** du projet obtenu dans sa liste de vœux agrégée. Produit en pied de tableau :

$$\bar{x} = \frac{1}{n}\sum_{i=1}^{n} r_i \qquad s = \sqrt{\frac{\sum_{i=1}^{n}(r_i - \bar{x})^2}{n-1}}$$

Un rang moyen proche de 1 indique une excellente satisfaction globale. L'écart-type mesure l'équité de la distribution.

### 5. `BilanPerformanceAlgorithme`
Génère deux feuilles de synthèse :

**`Bilan_Performance`** (indicateurs clés) :
- Taux d'affectation des équipes
- Satisfaction moyenne et écart-type des rangs obtenus
- Nombre d'équipes ayant obtenu leur 1er / 2ème / 3ème vœu
- Taux de projets ayant atteint leur minimum d'équipes

**`Details_Suivi`** (listes de suivi) :
- Équipes non affectées
- Projets n'ayant pas atteint leur minimum
- Projets sans aucune équipe

**`Dashboard_KPI_Metier`** (pilotage KPI avec seuils couleurs) :
- Statut automatique **VERT / JAUNE / ORANGE / ROUGE** selon seuils
- KPI équipes : satisfaction pondérée, satisfaction normalisée, top 1, top 3, frustration, équité
- KPI projets : atteinte des minima, projets vides, remplissage capacité, déficit minima
- KPI robustesse : tension capacitaire et marge de capacité

Politique de seuils appliquée :
- KPI où une valeur haute est meilleure (`higherIsBetter=True`) : `VERT >= seuil vert`, `JAUNE >= seuil jaune`, `ORANGE >= seuil orange`, sinon `ROUGE`.
- KPI où une valeur basse est meilleure (`higherIsBetter=False`) : `VERT <= seuil vert`, `JAUNE <= seuil jaune`, `ORANGE <= seuil orange`, sinon `ROUGE`.

### 6. `RemplirAffectationsParProjet`
Génère la feuille `Affectations_par_Projet` : vue détaillée avec, pour chaque projet, la liste des équipes affectées et la composition (membres) de chaque équipe.

Colonnes produites :

| Colonne | Contenu |
|---|---|
| **Projet** | Nom du projet (cellules fusionnées si plusieurs équipes) |
| **Statut Capacité** | OK ou MINIMUM NON ATTEINT |
| **Équipe** | Nom de chaque équipe affectée |
| **Membres** | Liste des élèves composant l'équipe |
| **Taille** | Nombre de membres |
| **Rang dans projet** | Classement de l'équipe selon le projet |

Couleurs :
- 🟢 Fond vert = projet avec minimum d'équipes atteint
- 🟡 Fond jaune = minimum non atteint
- 🔴 Fond rouge = aucune équipe affectée

### 7. `TestArrayList` (diagnostic)
Vérifie que `System.Collections.ArrayList` (.NET) est disponible sur la machine. Non utilisé par les macros principales, conservé à titre de diagnostic.

---

## Flux d'exécution recommandé

```
Mode A (simulation)
GenererDonneesDeTest
   │
   └──▶ AffectationEquipesPasAPas   (ou AffectationEquipesProjets)

Mode B (formulaires en ligne)
InitialiserFeuillesFormulaires
   │
   ├──▶ Option manuelle :
   │      Coller les exports de formulaires dans Reponses_Eleves / Reponses_Commanditaires
   │      puis GenererDonneesDepuisFormulaires
   │
   └──▶ Option automatique (recommandée) :
          InitialiserConfigurationImportFormulaires
          puis ImporterReponsesFormulairesAutomatique
       │
       └──▶ AffectationEquipesPasAPas   (ou AffectationEquipesProjets)
        │
        ├──▶ RemplirAffectationsParProjet   (vue détaillée projet → membres)
        │
        ├──▶ CreerRapportSatisfaction       (satisfaction par équipe)
        │
        └──▶ BilanPerformanceAlgorithme     (métriques globales)
```

> **Important :** Avant les macros d'affectation, il faut alimenter les 4 feuilles d'entrée via **une** des deux voies :
> - `GenererDonneesDeTest` (simulation)
> - `GenererDonneesDepuisFormulaires` (données réelles)

---

## Limites connues

| # | Limite | Impact |
|---|---|---|
| 1 | **Ex-aequo dans les classements** — deux équipes avec le même score sont départagées de façon déterministe par leur numéro | Comportement stable et reproductible |
| 2 | **Score de Borda** — peut produire des égalités de score ; l'implémentation les résout en faveur de l'indice le plus bas | Cas rare, impact marginal |
| 3 | **Équipes sans vœux compatibles** — si tous les projets sont incompatibles avec la taille d'une équipe, elle reste non affectée | Signalé dans `Details_Suivi` |

---

## À faire plus tard

- Mode "temps réel" sans fichier intermédiaire : l'étape suivante serait une connexion Power Query/Graph/Power Automate.

---

## Installation et intégration du code VBA

### Option 1 : À partir du fichier `.vb` (pour développeurs)

Le fichier `GaylShapley.vb` contient tout le code source VBA. Pour l'intégrer dans Excel :

**1. Ouvrir l'éditeur VBA**
   - Lancer Excel
   - Appuyer sur `Alt + F11` pour ouvrir l'**Éditeur VBA**
   - Ou via le menu : **Outils** → **Macros** → **Éditeur Visual Basic**

**2. Créer ou accéder au module**
   - Dans la fenêtre de gauche (**Explorateur de projets**), développer votre classeur
   - Clic droit sur **Modules** → **Insérer un module**
   - Une feuille blanche s'ouvre : vous avez un nouveau module

**3. Importer le code**
   - **Option A** : Copier-coller
     - Ouvrir `GaylShapley.vb` dans un éditeur de texte
     - Copier tout le contenu (Ctrl+A, Ctrl+C)
     - Coller dans le module VBA (Ctrl+V)
   
   - **Option B** : Importer directement (recommandé)
     - Dans l'éditeur VBA, **Fichier** → **Importer le fichier...**
     - Sélectionner `GaylShapley.vb`
     - Le module est créé automatiquement

**4. Sauvegarder en `.xlsm`**
   - Retourner à Excel (Alt+Q)
   - **Fichier** → **Enregistrer sous...**
   - Format : choisir **"Classeur Excel prenant en charge les macros (.xlsm)"**
   - Les macros sont maintenant intégrées et exécutables

### Option 2 : Utiliser directement le classeur packagé

Si le dépôt GitHub fournit un fichier `.xlsm` pré-intégré :
   - Télécharger le fichier `.xlsm`
   - S'assurer que **les macros sont activées** au moment de l'ouverture (Excel peut afficher une alerte de sécurité)
   - Les macros sont prêtes à utiliser

### Activation des macros

Certaines versions d'Excel demandent une confirmation lors de l'ouverture d'un classeur avec macros :
   - **Message de sécurité** : "Les macros ont été désactivées. Cliquez ici pour les activer."
   - Clic sur **"Activer les macros"** ou **"Options..."** → **"Activer ce contenu"**
   - Les macros restent activées tant que le fichier n'est pas déplacé

> **Note** : Pour exécuter les macros via le menu, aller à **Affichage** → **Macros** → **Afficher les macros** (ou `Alt + F8`), puis sélectionner la macro à exécuter.

---

## Corrections et améliorations récentes (v1.1+)

| Correction | Date | Impact |
|---|---|---|
| ✅ Levée de la limite de 26 projets | 2026-07 | Passage de noms lettres (A-Z) à numérotation (1, 2, 3, ...). Support illimité. |
| ✅ Décodage des noms de projets (Macro 1) | 2026-07 | Agrégation des préférences d'équipe compatible avec la nouvelle numérotation |
| ✅ Correction de la déclaration dupliquée | 2026-07 | Erreur de compilation `projLookup` supprimée |
| ✅ Documentation d'installation complète | 2026-07 | Section "Installation et intégration du code VBA" ajoutée |
| ✅ Mode formulaires en ligne | 2026-07 | Import des réponses élèves/commanditaires vers les feuilles d'entrée de l'algorithme |

---

## Prérequis

- Microsoft Excel (Windows) avec macros activées, sauvegardé en `.xlsm`
- Aucune dépendance externe : les macros utilisent uniquement `Scripting.Dictionary` (WScript natif Windows) et `Collection` (VBA natif)
