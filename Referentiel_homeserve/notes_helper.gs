/**
 * Ajoute les notes d'aide sur les cellules pertinentes.
 * À lancer une seule fois depuis l'éditeur Apps Script (les notes sont permanentes).
 *
 * Si tu veux les retirer plus tard, lance `retirerNotes`.
 */
function ajouterNotes() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('TRANSFERTS');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Feuille TRANSFERTS introuvable');
    return;
  }

  // Note sur la colonne CP (I2:I201)
  const noteCP =
    'Format CP accepté :\n\n' +
    '• Vide ou "Tous" → tous les CP de la filiale\n' +
    '• 33000 → un seul CP\n' +
    '• 33000, 33100, 33200 → liste (virgules)\n' +
    '• 33000-33999 → tranche\n' +
    '• 33000-33500, 35000 → mix\n\n' +
    'Le copier-coller depuis Excel (sauts de ligne) est accepté.';

  const rangeCP = sheet.getRange('I2:I201');
  // Appliquer la même note à toutes les cellules
  const notes = [];
  for (let i = 0; i < 200; i++) notes.push([noteCP]);
  rangeCP.setNotes(notes);

  // Note sur les colonnes Commercial (B et C)
  const noteCom = 'Choisissez parmi les commerciaux de la filiale sélectionnée (colonne A).';
  const rangeB = sheet.getRange('B2:B201');
  const rangeC = sheet.getRange('C2:C201');
  const notesCom = [];
  for (let i = 0; i < 200; i++) notesCom.push([noteCom]);
  rangeB.setNotes(notesCom);
  rangeC.setNotes(notesCom);

  // Note sur l'en-tête CP (rappel visible)
  sheet.getRange('I1').setNote(noteCP);

  SpreadsheetApp.getActive().toast(
    'Notes d\'aide ajoutées sur TRANSFERTS (colonnes B, C et I)',
    'Notes installées',
    5
  );
}

/**
 * Retire toutes les notes d'aide installées par ajouterNotes.
 */
function retirerNotes() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('TRANSFERTS');
  if (!sheet) return;

  sheet.getRange('B2:B201').clearNote();
  sheet.getRange('C2:C201').clearNote();
  sheet.getRange('I2:I201').clearNote();
  sheet.getRange('I1').clearNote();

  SpreadsheetApp.getActive().toast('Notes retirées', 'OK', 3);
}
