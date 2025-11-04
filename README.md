# Validateur de données — Excel Online (Office Scripts)

Après avoir fait des macros VBA et un script Python pour détecter les doublons,
j'ai voulu essayer Office Scripts — l'équivalent de VBA mais pour Excel Online.

La syntaxe est du TypeScript, pas de VBA, mais la logique est très similaire.
Ce script parcourt un tableau Excel, détecte les erreurs de saisie courantes
et génère un rapport de validation directement dans un onglet du classeur.

## Ce que ça fait

- Vérifie que les colonnes obligatoires ne sont pas vides
- Valide le format des dates (JJ/MM/AAAA)
- Vérifie que les colonnes numériques contiennent bien des nombres
- Génère un onglet "Rapport" avec le résumé des erreurs trouvées
- Colore les cellules problématiques en rouge directement dans le tableau
- Permet de remettre à zéro la mise en forme (reset)

## Comment utiliser

1. Ouvrir le fichier Excel dans Excel Online (Microsoft 365)
2. Aller dans **Automatiser** → **Nouveau script**
3. Coller le contenu de `validateur.ts`
4. Cliquer **Exécuter**

## Fichier de test

`exemples/donnees_equipements.xlsx` — un tableau fictif d'équipements de laboratoire
avec quelques erreurs intentionnelles pour tester le validateur.

## Différences VBA vs Office Scripts que j'ai remarquées

| VBA | Office Scripts |
|---|---|
| `ActiveSheet` | `workbook.getActiveWorksheet()` |
| `.End(xlUp).Row` | `.getUsedRange().getRowCount()` |
| `Cells(i, j)` | `.getCell(i, j)` |
| `MsgBox "texte"` | `console.log("texte")` |
| `.Interior.Color = RGB(...)` | `.getFormat().getFill().setColor("#FF0000")` |

La logique est la même, juste la syntaxe change.

---
*Projet personnel — novembre / décembre 2025*
