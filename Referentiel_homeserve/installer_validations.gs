/**
 * ============================================================
 * INSTALLER TOUTES LES LISTES DÉROULANTES — HomeServe Énergies Services
 * ============================================================
 *
 * Recrée toutes les validations de données nativement dans Google Sheets.
 * À relancer après tout import Excel → Sheets (les validations Excel ne survivent pas).
 *
 * UTILISATION :
 *   Menu ⚙️ Transferts → Installer les listes déroulantes
 *   — ou —
 *   Extensions > Apps Script → sélectionner installerToutesLesValidations → ▶️ Exécuter
 *
 * IDEMPOTENT : supprime puis recrée. Relançable à volonté.
 * ============================================================
 */

function installerToutesLesValidations() {
  const ss = SpreadsheetApp.getActive();
  let total = 0;
  const logs = [];

  try {
    total += installer_RECHERCHE_(ss, logs);
    total += installer_TRANSFERTS_(ss, logs);
    total += installer_AFFECTATIONS_(ss, logs);
    total += installer_PRODUITS_(ss, logs);
    total += installer_EXCEPTIONS_(ss, logs);
    total += installer_CONTACTS_(ss, logs);

    ss.toast(`${total} validations installées`, '✓ Listes déroulantes opérationnelles', 8);
    SpreadsheetApp.getUi().alert(
      'Installation terminée',
      logs.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Erreur lors de l\'installation',
      err.message + '\n\nVérifiez que toutes les feuilles existent (TRANSFERTS, RECHERCHE, PARAMETRES, etc.)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================================
// RECHERCHE
// ============================================================
function installer_RECHERCHE_(ss, logs) {
  const sheet = ss.getSheetByName('RECHERCHE');
  if (!sheet) { logs.push('✗ RECHERCHE manquante'); return 0; }

  sheet.getRange('A1:Z100').clearDataValidations();
  const params = ss.getSheetByName('PARAMETRES');
  let count = 0;

  // C6 = Filiale
  sheet.getRange('C6').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('B2:B16'), true)
      .setAllowInvalid(false).build()
  );
  count++;

  // C9 = Produit (filtre, optionnel)
  sheet.getRange('C9').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('C2:C25'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  logs.push(`✓ RECHERCHE : ${count} validations`);
  return count;
}

// ============================================================
// TRANSFERTS
// ============================================================
function installer_TRANSFERTS_(ss, logs) {
  const sheet = ss.getSheetByName('TRANSFERTS');
  if (!sheet) { logs.push('✗ TRANSFERTS manquante'); return 0; }

  sheet.getRange('A1:Z201').clearDataValidations();
  const params = ss.getSheetByName('PARAMETRES');
  let count = 0;

  // A = Filiale
  sheet.getRange('A2:A201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('B2:B16'), true)
      .setAllowInvalid(false).build()
  );
  count++;

  // B et C = Commerciaux filtrés par filiale (cascade via _DDL_TRANSFERTS)
  // Chaque ligne a sa propre plage source → 200 règles distinctes, mais on
  // les applique en 2 appels groupés (setDataValidations) au lieu de 400.
  const ddlSheet = ss.getSheetByName('_DDL_TRANSFERTS');
  if (ddlSheet) {
    const regles = [];
    for (let i = 2; i <= 201; i++) {
      const sourceRange = ddlSheet.getRange(i, 1, 1, 30); // A:AD de la ligne i
      regles.push([
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(sourceRange, true)
          .setAllowInvalid(true)
          .setHelpText('Choisissez parmi les commerciaux de la filiale sélectionnée (colonne A).')
          .build()
      ]);
    }
    sheet.getRange(2, 2, 200, 1).setDataValidations(regles); // colonne B en 1 appel
    sheet.getRange(2, 3, 200, 1).setDataValidations(regles); // colonne C en 1 appel
    count += 2;
  } else {
    logs.push('⚠ _DDL_TRANSFERTS manquante — colonnes B et C sans cascade');
  }

  // F = Motif
  sheet.getRange('F2:F201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('E2:E6'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  // G = Actif
  sheet.getRange('G2:G201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('D2:D3'), true)
      .setAllowInvalid(false).build()
  );
  count++;

  // H = Produit (liste incluant "Tous")
  sheet.getRange('H2:H201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('J2:J26'), true)
      .setAllowInvalid(false).build()
  );
  count++;

  logs.push(`✓ TRANSFERTS : ${count} validations (dont 200 cascades B+C)`);
  return count;
}

// ============================================================
// AFFECTATIONS_COMMUNES
// ============================================================
function installer_AFFECTATIONS_(ss, logs) {
  const sheet = ss.getSheetByName('AFFECTATIONS_COMMUNES');
  if (!sheet) { logs.push('✗ AFFECTATIONS_COMMUNES manquante'); return 0; }

  const params = ss.getSheetByName('PARAMETRES');
  const lastRow = sheet.getLastRow();
  const fin = lastRow + 50; // marge pour ajouts
  let count = 0;

  // A = Région
  sheet.getRange(2, 1, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('A2:A5'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  // B = Filiale
  sheet.getRange(2, 2, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('B2:B16'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  // G = Actif
  sheet.getRange(2, 7, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('D2:D3'), true)
      .setAllowInvalid(false).build()
  );
  count++;

  // J = Rôle
  sheet.getRange(2, 10, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('I2:I6'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  // K = Priorité
  sheet.getRange(2, 11, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('H2:H6'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  logs.push(`✓ AFFECTATIONS_COMMUNES : ${count} validations sur ${lastRow} lignes`);
  return count;
}

// ============================================================
// PRODUITS_COMMERCIAUX
// ============================================================
function installer_PRODUITS_(ss, logs) {
  const sheet = ss.getSheetByName('PRODUITS_COMMERCIAUX');
  if (!sheet) { logs.push('✗ PRODUITS_COMMERCIAUX manquante'); return 0; }

  const params = ss.getSheetByName('PARAMETRES');
  const lastRow = sheet.getLastRow();
  const fin = lastRow + 50;
  let count = 0;

  // A = Filiale
  sheet.getRange(2, 1, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('B2:B16'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  // C = Produit
  sheet.getRange(2, 3, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('C2:C25'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  // D = Actif
  sheet.getRange(2, 4, fin - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('D2:D3'), true)
      .setAllowInvalid(false).build()
  );
  count++;

  logs.push(`✓ PRODUITS_COMMERCIAUX : ${count} validations`);
  return count;
}

// ============================================================
// EXCEPTIONS_PRODUITS
// ============================================================
function installer_EXCEPTIONS_(ss, logs) {
  const sheet = ss.getSheetByName('EXCEPTIONS_PRODUITS');
  if (!sheet) { logs.push('✗ EXCEPTIONS_PRODUITS manquante'); return 0; }

  const params = ss.getSheetByName('PARAMETRES');
  let count = 0;

  sheet.getRange('A2:A201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('B2:B16'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  sheet.getRange('D2:D201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('C2:C25'), true)
      .setAllowInvalid(true).build()
  );
  count++;

  sheet.getRange('F2:F201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(params.getRange('D2:D3'), true)
      .setAllowInvalid(false).build()
  );
  count++;

  logs.push(`✓ EXCEPTIONS_PRODUITS : ${count} validations`);
  return count;
}

// ============================================================
// CONTACTS
// ============================================================
function installer_CONTACTS_(ss, logs) {
  const sheet = ss.getSheetByName('CONTACTS');
  if (!sheet) { logs.push('✗ CONTACTS manquante'); return 0; }

  sheet.getRange('C6').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['TEST', 'PROD'], true)
      .setAllowInvalid(false).build()
  );

  logs.push('✓ CONTACTS : 1 validation (mode TEST/PROD)');
  return 1;
}

// ============================================================
// TOUT EN UN (validations + notes)
// ============================================================
function installerToutComplet() {
  installerToutesLesValidations();
  if (typeof ajouterNotes === 'function') ajouterNotes();
}
