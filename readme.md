# PingCastleDashboard

Génère un tableau de bord sur l'avancée du score Ping Castle à travers le temps en se basant sur les fichiers XML générés par Ping Castle.

Pour l'instant je me sert du `README.md` comme une liste des trucs à faire 😊

## Choses à faire

- Créer au moins 10 versions dummy de fichier XML PingCastle
- Ajouter le numéro de version du PingCastle à chaque report
- Penser à générer un dashboard différent par domaine
- Sortir le report du script principal pour plus de lisibilité

## Structure du dashboard

1. Page principale
  - Courbe avec l'évolution du nombre de points au total et sur chaque catégorie
2. Report #01
  - Informations en bref : numéro de version, date de l'audit, etc
  - Camembert avec la répartition des points entre les différentes catégories
  - Diagramme en barres qui montre l'évolution de chaque catégorie par rapport au rapport précédent
  - Tableau avec les points d'améliorations depuis le dernier rapport
  - Tableau avec les nouveaux points d'attention
  - Tableau de tous les points restant

## Liens utiles et inspirations

- <https://evotec.xyz/dashimo-pswritehtml-charting-icons-and-few-other-changes/>
- <https://github.com/vletoux/pingcastle>
