## TODO - Etat au 2026-07-30

L'algorithme actuel produit une affectation stable au sens Gale-Shapley avec equipes proposantes.
Les points ci-dessous sont alignes avec l'implementation VBA actuelle.

## Fait

1. Classement projet avec departage deterministe des ex aequo

- Regle en place :

```vb
If (scores(k) < scores(eq)) Or (scores(k) = scores(eq) And k < eq) Then rang = rang + 1
```

- Effet : ordre strict et reproductible.

Exemple correct de rangs :

| Equipe | Score | Rang |
| ------ | ----- | ---- |
| E1     | 1.50  | 1    |
| E2     | 1.50  | 2    |
| E3     | 2.00  | 3    |

2. Classement des equipes par projet base sur preferences individuelles

- Score utilise :

```text
Moyenne des rangs + Variance
```

- Effet : une equipe avec interet plus partage est mieux classee.

3. Preferences d'equipe agregees avec score de Borda

- Regle : 1er choix = N points, 2e = N-1, ..., dernier = 1.
- Effet : preference collective plus robuste qu'un vote majoritaire simple.

## A faire

4. Integrer MinEquipes dans l'algorithme (et pas seulement en controle post-affectation)

- Etat actuel : MinEquipes est verifie en sortie (statut et KPI), pas impose pendant l'affectation.
- Proposition : affectation en deux phases.
	- Phase 1 : viser les minima des projets.
	- Phase 2 : Gale-Shapley sur la capacite restante.

5. Ajouter une amelioration locale post-affectation

- Objectif : tester des echanges d'equipes qui ameliorent la satisfaction sans violer min/max ni contraintes de taille.

6. Ajouter un indicateur explicite de paires bloquantes

- Objectif : certifier formellement la stabilite observee sur chaque run.
- Sortie attendue : nombre de paires bloquantes (0 = stable).

## Optionnel (experimentation)

7. Variante avec liste d'attente projet (au lieu du rejet immediat)

- Objectif : reduire les effets d'ordre dans certains jeux de donnees.

8. Etude d'equite proposeurs/recepteurs

- Executer deux versions : equipes proposantes puis projets proposants.
- Comparer les KPI et choisir une politique metier.

## Priorites recommandees

1. Integrer MinEquipes dans le moteur d'affectation.
2. Mesurer les paires bloquantes apres chaque affectation.
3. Ajouter l'amelioration locale.
