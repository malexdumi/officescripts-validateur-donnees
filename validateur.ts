// validateur.ts
// v1.1 -- détection des cellules vides dans les colonnes obligatoires
//
// Maintenant que je sais lire le tableau, j'ajoute la vraie logique :
// vérifier que les colonnes importantes ne sont pas vides.
// Je colore les cellules problématiques en rouge — même idée
// que ce que je faisais en VBA avec Interior.Color.

function main(workbook: ExcelScript.Workbook) {

  const feuille = workbook.getActiveWorksheet();
  const plage = feuille.getUsedRange();

  if (!plage) {
    console.log("Feuille vide — rien à valider.");
    return;
  }

  const nbLignes = plage.getRowCount();
  const nbColonnes = plage.getColumnCount();

  // Colonnes obligatoires (index 0-based)
  // On suppose que les 4 premières colonnes sont obligatoires
  const colonnesObligatoires = [0, 1, 2, 3];

  let nbErreurs = 0;

  // Remettre d'abord toutes les cellules en blanc (reset)
  for (let ligne = 1; ligne < nbLignes; ligne++) {
    for (let col = 0; col < nbColonnes; col++) {
      feuille.getCell(ligne, col).getFormat().getFill().clear();
    }
  }

  // Vérifier les cellules vides dans les colonnes obligatoires
  for (let ligne = 1; ligne < nbLignes; ligne++) {
    for (const col of colonnesObligatoires) {
      const valeur = feuille.getCell(ligne, col).getValue();

      // Vide si null, undefined, ou chaîne vide après trim
      if (valeur === null || valeur === undefined || String(valeur).trim() === "") {
        // Colorier en rouge clair
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFCCCC");
        console.log(`  Cellule vide → ligne ${ligne + 1}, colonne ${col + 1}`);
        nbErreurs++;
      }
    }
  }

  if (nbErreurs === 0) {
    console.log("Aucune cellule vide trouvée dans les colonnes obligatoires.");
  } else {
    console.log(`\nTotal : ${nbErreurs} cellule(s) vide(s) marquée(s) en rouge.`);
  }
}
