Dans votre cas, il faut distinguer **stabilité au sens Gale-Shapley** et **qualité globale de l'affectation**.

Votre algorithme produit déjà une affectation stable par rapport aux préférences utilisées :

* les équipes proposent ;
* les projets gardent provisoirement leurs meilleures équipes ;
* aucune équipe et aucun projet ne peuvent former une paire qui se préférerait mutuellement à leur affectation finale.

Cependant, plusieurs éléments réduisent la stabilité "pratique" ou la robustesse du résultat.

## 1. Éliminer les ex æquo dans les classements des projets

Aujourd'hui :

```vb
If scores(k) < scores(eq) Then rang = rang + 1
```

Deux équipes peuvent obtenir le même rang.

Exemple :

| Équipe | Score | Rang |
| ------ | ----- | ---- |
| E1     | 1.50  | 1    |
| E2     | 1.50  | 1    |
| E3     | 2.00  | 3    |

Quand un projet est plein, le choix de la pire équipe devient plus arbitraire.

### Amélioration

Créer un ordre total :

* score moyen
* puis taille d'équipe
* puis numéro d'équipe
* ou aléatoire reproductible

Ainsi chaque projet possède une préférence stricte.

***

## 2. Utiliser les préférences individuelles plutôt qu'une simple moyenne

Actuellement un projet classe les équipes selon :

```vb
score(eq) = moyenne des rangs des membres
```

Une moyenne masque les désaccords.

Exemple :

| Élève 1 | Élève 2 | Moyenne |
| ------- | ------- | ------- |
| 1       | 5       | 3       |
| 3       | 3       | 3       |

Les deux équipes sont équivalentes alors que leurs profils sont très différents.

### Alternative plus robuste

Calculer :

```text
Moyenne + Variance
```

ou

```text
Médiane des rangs
```

Ainsi les projets privilégient les équipes dont l'intérêt est partagé par tous les membres.

***

## 3. Utiliser un score de Borda

Votre vote majoritaire produit parfois des préférences d'équipes instables.

Exemple :

* 2 membres préfèrent Projet A
* 2 membres préfèrent Projet B

Le gagnant dépend alors de la règle de départage.

À la place :

```text
1er choix = N points
2e choix = N-1 points
...
Dernier choix = 1 point
```

Puis :

```vb
ScoreProjet = Somme des points de tous les membres
```

On obtient généralement des préférences d'équipe plus cohérentes.

***

## 4. Intégrer les minima dans l'algorithme

Actuellement :

```vb
MinEquipes
```

n'est utilisé qu'après coup.

Résultat :

```text
Projet A : minimum 2 -> reçoit 0
Projet B : minimum 2 -> reçoit 1
Projet C : reçoit tout
```

L'affectation est stable mais ne respecte pas les besoins des projets.

### Solution

Faire l'affectation en deux phases :

#### Phase 1

Atteindre les minima :

```text
chaque projet réserve ses places minimales
```

#### Phase 2

Lancer Gale-Shapley pour les places restantes.

Cela augmente fortement la stabilité organisationnelle.

***

## 5. Ajouter une étape d'amélioration locale

Après Gale-Shapley, chercher des échanges bénéfiques.

Exemple :

```text
Équipe A -> Projet 2
Équipe B -> Projet 4
```

Si :

* A préfère Projet 4
* B préfère Projet 2
* les contraintes restent respectées

alors échanger.

Cette phase correspond à une recherche locale ("local search") et améliore souvent la satisfaction moyenne sans casser les contraintes.

***

## 6. Remplacer le rejet immédiat par une liste d'attente

Aujourd'hui, un projet plein :

1. garde ses meilleures équipes ;
2. rejette immédiatement les autres.

Une variante consiste à :

```text
accepter provisoirement jusqu'à Max + marge
```

puis trier à la fin d'un cycle.

Cela réduit les effets liés à l'ordre des propositions.

***

## 7. Prendre en compte l'équité entre équipes

Gale-Shapley est favorable au côté qui propose.

Dans votre implémentation :

```text
Équipes = proposeurs
Projets = récepteurs
```

Les équipes sont avantagées.

Pour équilibrer :

* exécuter Gale-Shapley équipes → projets ;
* exécuter Gale-Shapley projets → équipes ;
* comparer les deux résultats ;
* choisir celui qui maximise un indicateur global.

***

## 8. Utiliser un critère de stabilité renforcée

Vous pouvez mesurer après affectation :

### Satisfaction moyenne

déjà présente.

### Satisfaction minimale

```text
quel est le pire rang obtenu ?
```

### Écart-type des rangs

déjà présent.

### Nombre de paires bloquantes

Vérifier :

```text
Existe-t-il une équipe E
et un projet P
qui se préfèrent mutuellement
à leur situation actuelle ?
```

S'il n'en existe aucune :

```text
Affectation stable
```

C'est le véritable critère théorique de stabilité.

***

## Ce qui apporterait le plus de valeur

Si je devais classer les améliorations :

1. **Utiliser les minima dans l'affectation**.
2. **Supprimer les ex æquo dans les classements projets**.
3. **Remplacer le vote majoritaire par un score de Borda**.
4. **Ajouter une phase d'amélioration locale après Gale-Shapley**.
5. **Mesurer explicitement les paires bloquantes** pour certifier la stabilité.

Ces cinq améliorations transformeraient votre outil d'un simple Gale-Shapley capacitaire vers un véritable système d'affectation robuste et exploitable dans un contexte pédagogique réel.
