// validateur.ts
// Lecture basique du tableau actif et affichage dans la console
// -- Maria-Alexandra, novembre 2025
//
// Premier essai avec Office Scripts. Je voulais juste comprendre
// comment accéder aux cellules avant d'ajouter la vraie logique.
// C'est assez différent de VBA syntaxiquement mais la logique
// est la même : on boucle sur les lignes, on lit les valeurs.

function main(workbook: ExcelScript.Workbook) {

  // Récupérer la feuille active
  const feuille = workbook.getActiveWorksheet();

  // Récupérer la plage utilisée (équivalent de .UsedRange en VBA)
  const plagUtilisee = feuille.getUsedRange();

  if (!plagUtilisee) {
    console.log("La feuille est vide.");
    return;
  }

  const nbLignes = plagUtilisee.getRowCount();
  const nbColonnes = plagUtilisee.getColumnCount();

  console.log(`Tableau détecté : ${nbLignes} lignes x ${nbColonnes} colonnes`);
  console.log(`(en-tête inclus)`);

  // Lire et afficher l'en-tête
  const entete: string[] = [];
  for (let col = 0; col < nbColonnes; col++) {
    entete.push(String(feuille.getCell(0, col).getValue()));
  }
  console.log(`Colonnes : ${entete.join(", ")}`);

  // Afficher les 3 premières lignes de données
  console.log("\nAperçu des données :");
  for (let ligne = 1; ligne < Math.min(4, nbLignes); ligne++) {
    const valeurs: string[] = [];
    for (let col = 0; col < nbColonnes; col++) {
      valeurs.push(String(feuille.getCell(ligne, col).getValue()));
    }
    console.log(`  Ligne ${ligne} : ${valeurs.join(" | ")}`);
  }
}
