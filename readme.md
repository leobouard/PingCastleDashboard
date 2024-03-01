# PingCastleDashboard

Génère un tableau de bord sur l'avancée du score PingCastle à travers le temps en se basant sur les fichiers XML générés par PingCastle.

Pour se servir du script, c'est très simple : on balance tous les fichiers XML issus de PingCastle (vous pouvez mélanger les domaines, un dashboard par domaine sera généré) puis vous exécutez le script `PingCastleDashboard.ps1` et le tour est joué !

## Exemple

Basé sur des données fictives :

![Capture d'écran du dashboard](/data/illustation.png)

## Liens utiles et inspirations

- <https://github.com/vletoux/pingcastle>
- <https://github.com/EvotecIT/PSWriteHTML>
- <https://evotec.xyz/dashimo-pswritehtml-charting-icons-and-few-other-changes/>
- <https://evotec.xyz/dashimo-easy-table-conditional-formatting-and-more/>

## Remerciements

- [Thierry PLANTIVE](https://www.linkedin.com/in/thierry-plantive-764b5b93/) pour l'idée originale, les retours et les demandes d'améliorations des rapports
- [stefan-frenchies](https://github.com/stefan-frenchies) pour le code pour la récupération des informations stockées dans les fichiers XML
- [Vincent VUCCINO](https://www.linkedin.com/in/vincent-vuccino-7948762b/) pour le code pour la récupération des niveaux de maturité via le GitHub de PingCastle
