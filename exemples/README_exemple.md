# Fichier exemple

`donnees_equipements_exemple.csv` montre la structure attendue du tableau Excel.

Ce fichier contient intentionnellement quelques erreurs pour tester le validateur :

| Ligne | Problème |
|---|---|
| EQ-003 | Colonne "nom" vide |
| EQ-004 | Colonne "localisation" vide |
| EQ-005 | Date au mauvais format (AAAA-MM-JJ au lieu de JJ/MM/AAAA) |
| EQ-005 | Valeur numérique invalide ("quatre" au lieu d'un chiffre) |

Les autres lignes sont valides — le script devrait afficher ✓ OK pour elles.

> Note : ce fichier CSV est fourni comme référence visuelle.
> Le script Office Scripts s'exécute directement dans Excel Online —
> il faut copier les données dans une feuille Excel avant de lancer le script.
