// validateur.ts
// v1.5 -- colonne statut par ligne + fonction reset séparée
//
// Dernières améliorations :
// 1. Ajouter une colonne "Statut" à droite du tableau avec OK / ERREUR
//    pour voir d'un coup d'oeil quelles lignes ont des problèmes
// 2. Séparer le reset dans sa propre fonction — utile pour effacer
//    les marquages avant une nouvelle validation

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

// Remet toutes les cellules en blanc et efface la colonne statut
function resetMiseEnForme(feuille: ExcelScript.Worksheet, nbLignes: number, nbColonnes: number) {
  for (let ligne = 0; ligne < nbLignes; ligne++) {
    for (let col = 0; col < nbColonnes + 1; col++) {
      feuille.getCell(ligne, col).getFormat().getFill().clear();
      feuille.getCell(ligne, col).getFormat().getFont().setColor("#000000");
    }
  }
  // Effacer la colonne statut
  for (let ligne = 1; ligne < nbLignes; ligne++) {
    feuille.getCell(ligne, nbColonnes).setValue("");
  }
}

function genererRapport(workbook: ExcelScript.Workbook, erreurs: string[], nbLignesValidees: number) {
  const rapportExistant = workbook.getWorksheet("Rapport de validation");
  if (rapportExistant) {
    rapportExistant.delete();
  }

  const rapport = workbook.addWorksheet("Rapport de validation");

  rapport.getCell(0, 0).setValue("RAPPORT DE VALIDATION");
  rapport.getCell(0, 0).getFormat().getFont().setBold(true);
  rapport.getCell(0, 0).getFormat().getFont().setSize(14);
  rapport.getCell(0, 0).getFormat().getFont().setColor("#1A3A5C");

  rapport.getCell(2, 0).setValue("Lignes validées");
  rapport.getCell(2, 1).setValue(nbLignesValidees);
  rapport.getCell(3, 0).setValue("Erreurs trouvées");
  rapport.getCell(3, 1).setValue(erreurs.length);

  if (erreurs.length === 0) {
    rapport.getCell(3, 1).getFormat().getFont().setColor("#008800");
  } else {
    rapport.getCell(3, 1).getFormat().getFont().setColor("#CC0000");
  }

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

  rapport.getRange("A:A").getFormat().setColumnWidth(65);
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
  const colonnesNumeriques = [5];

  const erreurs: string[] = [];
  let nbLignesValidees = 0;

  // Reset complet avant de commencer
  resetMiseEnForme(feuille, nbLignes, nbColonnes);

  // Ajouter l'en-tête de la colonne statut
  feuille.getCell(0, nbColonnes).setValue("Statut");
  feuille.getCell(0, nbColonnes).getFormat().getFont().setBold(true);

  for (let ligne = 1; ligne < nbLignes; ligne++) {

    if (estLigneVide(feuille, ligne, nbColonnes)) continue;
    nbLignesValidees++;

    let ligneADesErreurs = false;

    // Cellules vides
    for (const col of colonnesObligatoires) {
      const valeur = feuille.getCell(ligne, col).getValue();
      if (valeur === null || valeur === undefined || String(valeur).trim() === "") {
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFCCCC");
        erreurs.push(`Ligne ${ligne + 1}, col ${col + 1} : cellule obligatoire vide`);
        ligneADesErreurs = true;
      }
    }

    // Dates
    for (const col of colonnesDates) {
      const valeur = String(feuille.getCell(ligne, col).getText());
      if (valeur.trim() !== "" && !validerDate(valeur)) {
        feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFE0CC");
        erreurs.push(`Ligne ${ligne + 1}, col ${col + 1} : date invalide "${valeur}"`);
        ligneADesErreurs = true;
      }
    }

    // Valeurs numériques
    for (const col of colonnesNumeriques) {
      const valeur = feuille.getCell(ligne, col).getValue();
      if (valeur !== null && valeur !== undefined && String(valeur).trim() !== "") {
        if (isNaN(Number(valeur))) {
          feuille.getCell(ligne, col).getFormat().getFill().setColor("#FFF0CC");
          erreurs.push(`Ligne ${ligne + 1}, col ${col + 1} : nombre attendu, reçu "${valeur}"`);
          ligneADesErreurs = true;
        }
      }
    }

    // Écrire le statut de la ligne
    const celluleStatut = feuille.getCell(ligne, nbColonnes);
    if (ligneADesErreurs) {
      celluleStatut.setValue("⚠ ERREUR");
      celluleStatut.getFormat().getFont().setColor("#CC0000");
      celluleStatut.getFormat().getFont().setBold(true);
    } else {
      celluleStatut.setValue("✓ OK");
      celluleStatut.getFormat().getFont().setColor("#008800");
    }
  }

  genererRapport(workbook, erreurs, nbLignesValidees);

  console.log(`Validation terminée : ${nbLignesValidees} ligne(s) vérifiée(s), ${erreurs.length} erreur(s).`);
  console.log("Voir l'onglet 'Rapport de validation'.");
}
