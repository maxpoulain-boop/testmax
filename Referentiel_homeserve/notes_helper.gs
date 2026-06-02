/**
 * Ajoute des notes d'aide sur les cellules pertinentes de TRANSFERTS et RECHERCHE.
 * À lancer une seule fois (ou après import Excel → Sheets).
 */
function ajouterNotes() {
  const ss = SpreadsheetApp.getActive();
  let count = 0;

  // ── TRANSFERTS ──────────────────────────────────────────────
  const transferts = ss.getSheetByName('TRANSFERTS');
  if (transferts) {
    const noteCP =
      'Format CP accepté :\n\n' +
      '• Vide ou "Tous" → tous les CP de la filiale\n' +
      '• 33000 → un seul CP\n' +
      '• 33000, 33100, 33200 → liste (virgules)\n' +
      '• 33000-33999 → tranche\n' +
      '• 33000-33500, 35000 → mix\n\n' +
      'Copier-coller depuis Excel (sauts de ligne) accepté.';

    const noteCom = 'Choisissez parmi les commerciaux de la filiale sélectionnée (colonne A).';

    const notesCP  = Array(200).fill([noteCP]);
    const notesCom = Array(200).fill([noteCom]);

    transferts.getRange('J2:J201').setNotes(notesCP);
    transferts.getRange('C2:C201').setNotes(notesCom);
    transferts.getRange('D2:D201').setNotes(notesCom);
    transferts.getRange('J1').setNote(noteCP);
    count++;
  }

  // ── RECHERCHE ────────────────────────────────────────────────
  const recherche = ss.getSheetByName('RECHERCHE');
  if (recherche) {
    recherche.getRange('C7').setNote(
      'Code postal à 5 chiffres (ex: 33000).\n' +
      'Doit correspondre exactement à la valeur dans AFFECTATIONS_COMMUNES.'
    );
    recherche.getRange('C9').setNote(
      'Optionnel. Si renseigné, filtre les commerciaux qui gèrent ce produit.\n' +
      'Laisser vide pour voir tous les commerciaux disponibles.'
    );
    count++;
  }

  SpreadsheetApp.getActive().toast(
    'Notes d\'aide ajoutées sur ' + count + ' feuille(s)',
    '✓ Notes installées',
    5
  );
}

/**
 * Retire toutes les notes ajoutées par ajouterNotes().
 */
function retirerNotes() {
  const ss = SpreadsheetApp.getActive();

  const transferts = ss.getSheetByName('TRANSFERTS');
  if (transferts) {
    transferts.getRange('C2:C201').clearNote();
    transferts.getRange('D2:D201').clearNote();
    transferts.getRange('J2:J201').clearNote();
    transferts.getRange('J1').clearNote();
  }

  const recherche = ss.getSheetByName('RECHERCHE');
  if (recherche) {
    recherche.getRange('C7').clearNote();
    recherche.getRange('C9').clearNote();
  }

  SpreadsheetApp.getActive().toast('Notes retirées', '✓ OK', 3);
}
