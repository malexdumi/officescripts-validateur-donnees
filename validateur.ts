// validateur.ts
// v1.4 -- validation numérique + génération d'un onglet rapport
//
// Deux ajouts :
// 1. Vérifier que les colonnes numériques contiennent bien des nombres
// 2. Générer un onglet "Rapport de validation" avec le résumé des erreurs
//    (même idée que le rapport .txt en Python, mais directement dans Excel)

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
  const regex = /^\d{2}\/\d{2}\/\d{4}$/;
  if (!regex.test(valeur.trim())) return false;
  const parties = valeur.trim().split("/");
  const jour = parseInt(parties[0]);
  const mois = parseInt(parties[1]);
  const annee = parseInt(parties[2]);
  if (jour < 1 || jour > 31) return false;
  if (mois < 1 || mois > 12) return false;
  if (annee < 2000 || annee > 2100) return false;
  return true;
}

function genererRapport(workbook: ExcelScript.Workbook, erreurs: string[], nbLignesValidees: number) {

  // Supprimer l'onglet rapport s'il existe déjà
  const rapportExistant = workbook.getWorksheet("Rapport de validation");
  if (rapportExistant) {
    rapportExistant.delete();
  }

  // Créer un nouvel onglet
  const rapport = workbook.addWorksheet("Rapport de validation");

  // Titre
  rapport.getCell(0, 0).setValue("RAPPORT DE VALIDATION");
  rapport.getCell(0, 0).getFormat().getFont().setBold(true);
  rapport.getCell(0, 0).getFormat().getFont().setSize(14);
  rapport.getCell(0, 0).getFormat().getFont().setColor("#1A3A5C");

  // Résumé
  rapport.getCell(2, 0).setValue("Lignes validées");
  rapport.getCell(2, 1).setValue(nbLignesValidees);

  rapport.getCell(3, 0).setValue("Erreurs trouvées");
  rapport.getCell(3, 1).setValue(erreurs.length);

  if (erreurs.length === 0) {
    rapport.getCell(3, 1).getFormat().getFont().setColor("#008800");
  } else {
    rapport.getCell(3, 1).getFormat().getFont().setColor("#CC0000");
  }

  // Détail des erreurs
  if (erreurs.length > 0) {
    rapport.getCell(5, 0).setValue("Détail des erreurs :");
    rapport.getCell(5, 0).getFormat().getFont().setBold(true);

    for (let i = 0; i < erreurs.length; i++) {
      rapport.getCell(6 + i, 0).setValue(erreurs[i]);
    }
  } else {
    rapport.getCell(5, 0).setValue("Aucune erreur — données valides.");
    rapport.getCell(5, 0).getFormat().getFont().setColor("#008800");
  }

  // Ajuster la largeur de la colonne A
  rapport.getRange("A:A").getFormat().setColumnWidth(60);
}

function main(workbook: ExcelScript.Workbook) {

  const feuille = workbook.getActiveWorksheet();
  const plage = feuille.getUsedRange();

  if (!plage) {
    console.log("Feuille vide.");
    return;
  }

  const nbLignes = plage.getRowCount();
  const nbColonnes = plage.getColumnCount();

  const colonnesObligatoires = [0, 1, 2, 3];
  const colonnesDates = [4];
  const colonnesNumeriques = [5]; // ex: colonne "quantité" ou "valeur"

  const erreurs: string[] = [];
  let nbLignesValidees = 0;

  // Reset mise en forme
  for (let ligne = 1; ligne < nbLignes; ligne++) {
    for (let col = 0; col < nbColonnes; col++) {
      feuille.getCell(ligne, col).getFormat().getFill().clear();
    }
  }

  for (let ligne = 1; ligne < nbLignes; ligne++) {

    if (estLigneVide(feuille, ligne, nbColonnes)) continue;
    nbLignesValidees++;

    // Cellules vides
    for (const col of colonnesObligatoires) {
      const valeur = feuille.getCell(ligne, col).getValue();
      if (valeur === null || valeur === undefined || String(valeur).trim() === "") {
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFCCCC");
        erreurs.push(`Ligne ${ligne + 1}, col ${col + 1} : cellule obligatoire vide`);
      }
    }

    // Dates
    for (const col of colonnesDates) {
      const valeur = String(feuille.getCell(ligne, col).getText());
      if (valeur.trim() !== "" && !validerDate(valeur)) {
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFE0CC");
        erreurs.push(`Ligne ${ligne + 1}, col ${col + 1} : date invalide "${valeur}" (attendu JJ/MM/AAAA)`);
      }
    }

    // Valeurs numériques
    for (const col of colonnesNumeriques) {
      const valeur = feuille.getCell(ligne, col).getValue();
      if (valeur !== null && valeur !== undefined && String(valeur).trim() !== "") {
        if (isNaN(Number(valeur))) {
          feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFF0CC");
          erreurs.push(`Ligne ${ligne + 1}, col ${col + 1} : valeur numérique attendue, reçu "${valeur}"`);
        }
      }
    }
  }

  // Générer l'onglet rapport
  genererRapport(workbook, erreurs, nbLignesValidees);

  console.log(`Validation terminée : ${nbLignesValidees} ligne(s), ${erreurs.length} erreur(s).`);
  console.log("Voir l'onglet 'Rapport de validation' pour le détail.");
}
