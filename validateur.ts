// validateur.ts
// v1.3 -- ajout validation format des dates
//
// Nouvelle vérification : les colonnes de dates doivent
// respecter le format JJ/MM/AAAA.
// J'ai utilisé une expression régulière (regex) pour ça —
// c'est la première fois que j'en utilise une en TypeScript.
// En VBA on aurait utilisé IsDate(), ici c'est un peu différent
// parce qu'Excel Online stocke les dates comme des nombres.
// Il faut donc vérifier la valeur affichée, pas la valeur brute.

function estLigneVide(feuille: ExcelScript.Worksheet, ligne: number, nbColonnes: number): boolean {
  for (let col = 0; col < nbColonnes; col++) {
    const val = feuille.getCell(ligne, col).getValue();
    if (val !== null && val !== undefined && String(val).trim() !== "") {
      return false;
    }
  }
  return true;
}

function validerDate(valeur: string): boolean {
  // Format attendu : JJ/MM/AAAA
  // Regex : 2 chiffres / 2 chiffres / 4 chiffres
  const regex = /^\d{2}\/\d{2}\/\d{4}$/;
  if (!regex.test(valeur.trim())) {
    return false;
  }

  // Vérifier que les valeurs sont dans des plages logiques
  const parties = valeur.trim().split("/");
  const jour = parseInt(parties[0]);
  const mois = parseInt(parties[1]);
  const annee = parseInt(parties[2]);

  if (jour < 1 || jour > 31) return false;
  if (mois < 1 || mois > 12) return false;
  if (annee < 2000 || annee > 2100) return false;

  return true;
}

function main(workbook: ExcelScript.Workbook) {

  const feuille = workbook.getActiveWorksheet();
  const plage = feuille.getUsedRange();

  if (!plage) {
    console.log("Feuille vide — rien à valider.");
    return;
  }

  const nbLignes = plage.getRowCount();
  const nbColonnes = plage.getColumnCount();

  // Configuration des colonnes à valider
  const colonnesObligatoires = [0, 1, 2, 3];
  const colonnesDates = [4]; // index de la colonne date_inspection

  let nbErreurs = 0;

  // Reset
  for (let ligne = 1; ligne < nbLignes; ligne++) {
    for (let col = 0; col < nbColonnes; col++) {
      feuille.getCell(ligne, col).getFormat().getFill().clear();
    }
  }

  for (let ligne = 1; ligne < nbLignes; ligne++) {

    if (estLigneVide(feuille, ligne, nbColonnes)) continue;

    // Vérifier cellules vides
    for (const col of colonnesObligatoires) {
      const valeur = feuille.getCell(ligne, col).getValue();
      if (valeur === null || valeur === undefined || String(valeur).trim() === "") {
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFCCCC");
        nbErreurs++;
      }
    }

    // Vérifier format des dates
    for (const col of colonnesDates) {
      const valeur = String(feuille.getCell(ligne, col).getText()); // .getText() pour avoir la valeur affichée
      if (valeur.trim() !== "" && !validerDate(valeur)) {
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFE0CC");
        console.log(`  Date invalide → ligne ${ligne + 1} : "${valeur}" (attendu JJ/MM/AAAA)`);
        nbErreurs++;
      }
    }
  }

  if (nbErreurs === 0) {
    console.log("Aucune erreur détectée.");
  } else {
    console.log(`\nTotal : ${nbErreurs} erreur(s) détectée(s) et marquée(s).`);
  }
}
