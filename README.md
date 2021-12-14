# Contacteur

Un logiciel de ligne de commande qui capte des informations en lien avec les élèves et les traite. Pour le moment, la version stable du logiciel produit seulement un fichier ODS avec les informations de contact des tuteurs des élèves d'un enseignant par le biais du site web EnCompass.

## Exécuter

Pour exécuter la version initiale du logiciel qui obtient que les informations de contact des tuteurs des élèves et les exporte à un fichier ODS, exécutez
```
cargo run --bin contacts
```
Pour exécuter la version expérimentale incomplète qui obtient les informations de contact des tuteurs des élèves, exporte les données de contact à un fichier ODS, charge les résultats d'évaluation d'un fichier ODS personnalisé et exporte tous les informations dans un fichier DB3, exécutez
```
cargo run --bin dev-experimental
```
Cette version détecte si les informations ont déjà été recueillies. Le but était d'automatiser l'envoie de courriels aux tuteurs suites aux évaluations. Il reste un peu de travail à faire pour implémenter cette fonctionnalité-là.