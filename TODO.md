## TODO - État au 2026-07-30

L'algorithme actuel produit une affectation stable au sens Gale-Shapley avec équipes proposantes.
Les points ci-dessous sont alignés avec l'implémentation VBA actuelle.

## Fait

1. Classement projet avec départage déterministe des ex æquo

- Règle en place :

```vb
If (scores(k) < scores(eq)) Or (scores(k) = scores(eq) And k < eq) Then rang = rang + 1
```

- Effet : ordre strict et reproductible.

Exemple correct de rangs :

| Équipe | Score | Rang |
| ------ | ----- | ---- |
| E1     | 1.50  | 1    |
| E2     | 1.50  | 2    |
| E3     | 2.00  | 3    |

2. Classement des équipes par projet basé sur préférences individuelles

- Score utilisé :

```text
Moyenne des rangs + Variance
```

- Effet : une équipe avec intérêt plus partagé est mieux classée.

3. Préférences d'équipe agrégées avec score de Borda

- Règle : 1er choix = N points, 2e = N-1, ..., dernier = 1.
- Effet : préférence collective plus robuste qu'un vote majoritaire simple.

## À faire

4. Intégrer MinEquipes dans l'algorithme (et pas seulement en contrôle post-affectation)

- État actuel : MinEquipes est vérifié en sortie (statut et KPI), pas imposé pendant l'affectation.
- Proposition : affectation en deux phases.
	- Phase 1 : viser les minima des projets.
	- Phase 2 : Gale-Shapley sur la capacité restante.

5. Ajouter une amélioration locale post-affectation

- Objectif : tester des échanges d'équipes qui améliorent la satisfaction sans violer min/max ni contraintes de taille.

6. Ajouter un indicateur explicite de paires bloquantes

- Objectif : certifier formellement la stabilité observée sur chaque run.
- Sortie attendue : nombre de paires bloquantes (0 = stable).

## Optionnel (experimentation)

7. Variante avec liste d'attente projet (au lieu du rejet immédiat)

- Objectif : réduire les effets d'ordre dans certains jeux de données.

8. Étude d'équité proposeurs/récepteurs

- Exécuter deux versions : équipes proposantes puis projets proposants.
- Comparer les KPI et choisir une politique métier.

## Priorités recommandées

1. Intégrer MinEquipes dans le moteur d'affectation.
2. Mesurer les paires bloquantes après chaque affectation.
3. Ajouter l'amélioration locale.
