/**
 * ============================================================
 * INSTALLER TOUTES LES LISTES DÉROULANTES — HomeServe Énergies Services
 * ============================================================
 *
 * Recrée toutes les validations de données nativement dans Google Sheets,
 * après que la conversion depuis Excel les ait perdues.
 *
 * UTILISATION :
 *   1. Coller ce code dans Extensions > Apps Script (nouveau fichier "installerValidations")
 *   2. Sauvegarder (Ctrl+S)
 *   3. Sélectionner la fonction `installerToutesLesValidations` dans le menu déroulant
 *   4. Cliquer ▶️ Exécuter
 *   5. Autoriser le script si demandé
 *   6. Attendre 10-20 secondes, un toast confirme la fin
 *
 * Le script est IDEMPOTENT : tu peux le relancer à volonté, il commence par tout
 * supprimer puis tout recréer proprement.
 *
 * ============================================================
 */

function installerToutesLesValidations() {
  const ss = SpreadsheetApp.getActive();
  let totalCount = 0;
  let logs = [];

  try {
    totalCount += installer_RECHERCHE_(ss, logs);
    totalCount += installer_TRANSFERTS_(ss, logs);
    totalCount += installer_AFFECTATIONS_(ss, logs);
    totalCount += installer_PRODUITS_(ss, logs);
    totalCount += installer_EXCEPTIONS_(ss, logs);
    totalCount += installer_CONTACTS_(ss, logs);

    ss.toast(
      `${totalCount} validations installées avec succès`,
      '✓ Listes déroulantes opérationnelles',
      8
    );

    SpreadsheetApp.getUi().alert(
      'Installation terminée',
      logs.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Erreur lors de l\'installation',
      'Détail : ' + err.message + '\n\nVérifie que toutes les feuilles existent (TRANSFERTS, RECHERCHE, AFFECTATIONS_COMMUNES, PARAMETRES, etc.)',
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

  // Supprimer toutes les validations existantes
  sheet.getRange('A1:Z100').clearDataValidations();

  let count = 0;

  // C6 = Filiale
  const rangeFiliales = ss.getSheetByName('PARAMETRES').getRange('B2:B16');
  const ruleFiliale = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeFiliales, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C6').setDataValidation(ruleFiliale);
  count++;

  // C9 = Produit (filtre)
  const rangeProduits = ss.getSheetByName('PARAMETRES').getRange('C2:C25');
  const ruleProduit = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeProduits, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange('C9').setDataValidation(ruleProduit);
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

  sheet.getRange('A1:Z210').clearDataValidations();

  let count = 0;

  // A = Filiale
  const rangeFiliales = ss.getSheetByName('PARAMETRES').getRange('B2:B16');
  sheet.getRange('A2:A201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeFiliales, true)
      .setAllowInvalid(false)
      .build()
  );
  count++;

  // B et C = Commerciaux filtrés par filiale (cascade via _DDL_TRANSFERTS)
  // Chaque ligne i a sa propre plage source DDL!Ai:ADi
  // Donc on doit créer 200 validations distinctes
  const ddlSheet = ss.getSheetByName('_DDL_TRANSFERTS');
  if (ddlSheet) {
    for (let i = 2; i <= 201; i++) {
      const sourceRange = ddlSheet.getRange(i, 1, 1, 30); // A:AD pour la ligne i
      const ruleCom = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceRange, true)
        .setAllowInvalid(true)
        .setHelpText('Choisissez parmi les commerciaux de la filiale sélectionnée (colonne A).')
        .build();
      sheet.getRange(i, 2).setDataValidation(ruleCom);  // colonne B
      sheet.getRange(i, 3).setDataValidation(ruleCom);  // colonne C
    }
    count += 2; // logique : 2 colonnes
  }

  // F = Motif
  const rangeMotifs = ss.getSheetByName('PARAMETRES').getRange('E2:E6');
  sheet.getRange('F2:F201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeMotifs, true)
      .setAllowInvalid(true)
      .build()
  );
  count++;

  // G = Actif
  const rangeActif = ss.getSheetByName('PARAMETRES').getRange('D2:D3');
  sheet.getRange('G2:G201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeActif, true)
      .setAllowInvalid(false)
      .build()
  );
  count++;

  // H = Produit (avec Tous)
  const rangeProduitsT = ss.getSheetByName('PARAMETRES').getRange('J2:J26');
  sheet.getRange('H2:H201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeProduitsT, true)
      .setAllowInvalid(false)
      .build()
  );
  count++;

  // I = CP — pas de validation stricte, juste help text
  // (la validation custom Excel ne se traduit pas bien en Sheets)
  // On utilise les Notes (ajouterNotes) pour l'aide

  logs.push(`✓ TRANSFERTS : ${count} validations (dont 200 cascades pour B et C)`);
  return count;
}

// ============================================================
// AFFECTATIONS_COMMUNES
// ============================================================
function installer_AFFECTATIONS_(ss, logs) {
  const sheet = ss.getSheetByName('AFFECTATIONS_COMMUNES');
  if (!sheet) { logs.push('✗ AFFECTATIONS_COMMUNES manquante'); return 0; }

  const lastRow = sheet.getLastRow();
  const targetRange = `2:${lastRow + 50}`; // marge pour ajouts futurs

  // On applique sur A, B, G, J, K
  let count = 0;

  // A = Région
  const rangeRegions = ss.getSheetByName('PARAMETRES').getRange('A2:A5');
  sheet.getRange(`A${targetRange.replace(':', ':A')}`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeRegions, true)
      .setAllowInvalid(true)
      .build()
  );
  count++;

  // B = Filiale
  const rangeFiliales = ss.getSheetByName('PARAMETRES').getRange('B2:B16');
  sheet.getRange(`B${targetRange.replace(':', ':B')}`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeFiliales, true)
      .setAllowInvalid(true)
      .build()
  );
  count++;

  // G = Actif
  const rangeActif = ss.getSheetByName('PARAMETRES').getRange('D2:D3');
  sheet.getRange(`G${targetRange.replace(':', ':G')}`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeActif, true)
      .setAllowInvalid(false)
      .build()
  );
  count++;

  // J = Rôle
  const rangeRoles = ss.getSheetByName('PARAMETRES').getRange('I2:I6');
  sheet.getRange(`J${targetRange.replace(':', ':J')}`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeRoles, true)
      .setAllowInvalid(true)
      .build()
  );
  count++;

  // K = Priorité
  const rangePrio = ss.getSheetByName('PARAMETRES').getRange('H2:H6');
  sheet.getRange(`K${targetRange.replace(':', ':K')}`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangePrio, true)
      .setAllowInvalid(true)
      .build()
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

  const lastRow = sheet.getLastRow();
  let count = 0;

  // A = Filiale
  const rangeFiliales = ss.getSheetByName('PARAMETRES').getRange('B2:B16');
  sheet.getRange(2, 1, lastRow + 50 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeFiliales, true)
      .setAllowInvalid(true)
      .build()
  );
  count++;

  // C = Produit
  const rangeProduits = ss.getSheetByName('PARAMETRES').getRange('C2:C25');
  sheet.getRange(2, 3, lastRow + 50 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeProduits, true)
      .setAllowInvalid(true)
      .build()
  );
  count++;

  // D = Actif
  const rangeActif = ss.getSheetByName('PARAMETRES').getRange('D2:D3');
  sheet.getRange(2, 4, lastRow + 50 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeActif, true)
      .setAllowInvalid(false)
      .build()
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

  let count = 0;

  const rangeFiliales = ss.getSheetByName('PARAMETRES').getRange('B2:B16');
  sheet.getRange('A2:A201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeFiliales, true).setAllowInvalid(true).build()
  );
  count++;

  const rangeProduits = ss.getSheetByName('PARAMETRES').getRange('C2:C25');
  sheet.getRange('D2:D201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeProduits, true).setAllowInvalid(true).build()
  );
  count++;

  const rangeActif = ss.getSheetByName('PARAMETRES').getRange('D2:D3');
  sheet.getRange('F2:F201').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeActif, true).setAllowInvalid(false).build()
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

  // C6 = mode TEST/PROD
  sheet.getRange('C6').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['TEST', 'PROD'], true)
      .setAllowInvalid(false)
      .build()
  );

  logs.push(`✓ CONTACTS : 1 validation`);
  return 1;
}

// ============================================================
// BONUS : tout en 1 (validations + notes) si tu veux faire tout d'un coup
// ============================================================
function installerToutComplet() {
  installerToutesLesValidations();
  // Si tu as aussi le script notes_helper.gs avec ajouterNotes :
  if (typeof ajouterNotes === 'function') {
    ajouterNotes();
  }
}
