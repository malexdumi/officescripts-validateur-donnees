// validateur.ts
// v1.2 -- correction bug lignes fantômes
//
// Problème découvert en testant : getUsedRange() retournait parfois
// des lignes "vides" en bas du tableau parce qu'une cellule avait
// été formatée sans contenir de valeur.
// Résultat : des faux positifs — des lignes vides signalées comme erreurs
// alors qu'elles ne font pas partie des données.
//
// Fix : avant de valider une ligne, vérifier qu'elle contient
// au moins une valeur non vide. Si toute la ligne est vide,
// on la saute — c'est une ligne fantôme.

function main(workbook: ExcelScript.Workbook) {

  const feuille = workbook.getActiveWorksheet();
  const plage = feuille.getUsedRange();

  if (!plage) {
    console.log("Feuille vide — rien à valider.");
    return;
  }

  const nbLignes = plage.getRowCount();
  const nbColonnes = plage.getColumnCount();
  const colonnesObligatoires = [0, 1, 2, 3];

  let nbErreurs = 0;
  let nbLignesFantomes = 0;

  // Reset mise en forme
  for (let ligne = 1; ligne < nbLignes; ligne++) {
    for (let col = 0; col < nbColonnes; col++) {
      feuille.getCell(ligne, col).getFormat().getFill().clear();
    }
  }

  for (let ligne = 1; ligne < nbLignes; ligne++) {

    // Vérifier si la ligne est entièrement vide (ligne fantôme)
    let ligneEstVide = true;
    for (let col = 0; col < nbColonnes; col++) {
      const val = feuille.getCell(ligne, col).getValue();
      if (val !== null && val !== undefined && String(val).trim() !== "") {
        ligneEstVide = false;
        break;
      }
    }

    // Sauter les lignes fantômes
    if (ligneEstVide) {
      nbLignesFantomes++;
      continue;
    }

    // Valider les colonnes obligatoires
    for (const col of colonnesObligatoires) {
      const valeur = feuille.getCell(ligne, col).getValue();
      if (valeur === null || valeur === undefined || String(valeur).trim() === "") {
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFCCCC");
        console.log(`  Cellule vide → ligne ${ligne + 1}, colonne ${col + 1}`);
        nbErreurs++;
      }
    }
  }

  if (nbLignesFantomes > 0) {
    console.log(`(${nbLignesFantomes} ligne(s) vide(s) ignorée(s) en fin de tableau)`);
  }

  if (nbErreurs === 0) {
    console.log("Aucune erreur détectée.");
  } else {
    console.log(`\nTotal : ${nbErreurs} cellule(s) vide(s) marquée(s).`);
  }
}
