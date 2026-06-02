/**
 * ============================================================
 * INSTALLER TOUTES LES LISTES DÉROULANTES — HomeServe Énergies Services
 * ============================================================
 *
 * Recrée toutes les validations de données nativement dans Google Sheets.
 * À relancer après tout import Excel → Sheets (les validations Excel ne survivent pas).
 *
 * v2 — Plages nommées :
 *   - Les sources des listes sont centralisées dans l'objet PARAM ci-dessous.
 *   - installerPlagesNommees() crée des plages nommées (PARAM_FILIALES, etc.)
 *     dans le Sheet. Les validations lisent ces plages en priorité.
 *   - Pour agrandir une liste (ex: 16e filiale), il suffit d'étendre la plage
 *     nommée une seule fois (Données → Plages nommées) — plus besoin de toucher
 *     au code à 15 endroits.
 *   - Si une plage nommée n'existe pas, le code retombe sur l'adresse A1 figée.
 *
 * UTILISATION :
 *   Menu ⚙️ Transferts → Installer les listes déroulantes
 *
 * IDEMPOTENT : supprime puis recrée. Relançable à volonté.
 * ============================================================
 */

// === Source unique des listes de référence (onglet PARAMETRES) ===
// nom = nom de la plage nommée Google Sheets ; col = colonne (1-based) ; ligne = ligne de départ.
// La plage est calculée dynamiquement jusqu'à la dernière cellule non-vide de la colonne.
const PARAM = {
  regions:   { nom: 'PARAM_REGIONS',    col: 1,  ligne: 2 },
  filiales:  { nom: 'PARAM_FILIALES',   col: 2,  ligne: 2 },
  produits:  { nom: 'PARAM_PRODUITS',   col: 3,  ligne: 2 },
  actif:     { nom: 'PARAM_ACTIF',      col: 4,  ligne: 2 },
  motifs:    { nom: 'PARAM_MOTIFS',     col: 5,  ligne: 2 },
  priorites: { nom: 'PARAM_PRIORITES',  col: 6,  ligne: 2 },
  roles:     { nom: 'PARAM_ROLES',      col: 7,  ligne: 2 },
  produitsT: { nom: 'PARAM_PRODUITS_T', col: 8,  ligne: 2 },
};

/**
 * Retourne la plage d'une liste dans PARAMETRES en détectant la dernière ligne non-vide.
 * Garantit au moins 1 ligne même si la colonne est vide (évite une erreur de plage).
 */
function plageDynamique_(params, def) {
  const colData = params.getRange(def.ligne, def.col, params.getLastRow() - def.ligne + 2, 1).getValues();
  let derniereLigne = def.ligne;
  for (let i = 0; i < colData.length; i++) {
    if (colData[i][0] !== '' && colData[i][0] !== null) derniereLigne = def.ligne + i;
  }
  return params.getRange(def.ligne, def.col, derniereLigne - def.ligne + 1, 1);
}

/**
 * Retourne la plage source d'une liste : la plage nommée si elle existe,
 * sinon calcul dynamique de la dernière ligne non-vide dans PARAMETRES.
 */
function plageSource_(ss, cle) {
  const def = PARAM[cle];
  if (!def) throw new Error('Clé de liste inconnue : ' + cle);
  const nommee = ss.getRangeByName(def.nom);
  if (nommee) return nommee;
  const params = ss.getSheetByName('PARAMETRES');
  if (!params) throw new Error('Onglet PARAMETRES introuvable');
  return plageDynamique_(params, def);
}

/**
 * Crée (ou recrée) toutes les plages nommées en détectant dynamiquement
 * la dernière ligne non-vide dans chaque colonne de PARAMETRES.
 * Ajouter une filiale/produit dans PARAMETRES suffit — aucune modification du code.
 */
function installerPlagesNommees() {
  const ss = SpreadsheetApp.getActive();
  const params = ss.getSheetByName('PARAMETRES');
  if (!params) {
    SpreadsheetApp.getUi().alert('Onglet PARAMETRES introuvable — impossible de créer les plages nommées.');
    return 0;
  }

  // Supprimer les plages nommées qu'on gère, pour repartir propre
  const nomsGeres = Object.keys(PARAM).map(c => PARAM[c].nom);
  ss.getNamedRanges().forEach(nr => {
    if (nomsGeres.indexOf(nr.getName()) !== -1) nr.remove();
  });

  let count = 0;
  Object.keys(PARAM).forEach(cle => {
    const def = PARAM[cle];
    ss.setNamedRange(def.nom, plageDynamique_(params, def));
    count++;
  });

  ss.toast(count + ' plages nommées créées (détection automatique)', '✓ Plages nommées', 5);
  return count;
}

// ============================================================
// INSTALLATION DES VALIDATIONS
// ============================================================

function installerToutesLesValidations() {
  const ss = SpreadsheetApp.getActive();
  let total = 0;
  const logs = [];

  try {
    // S'assure que les plages nommées existent avant de bâtir les validations
    installerPlagesNommees();

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

/** Petit raccourci pour construire une règle "valeur dans une liste de référence". */
function regleListe_(ss, cle, allowInvalid) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(plageSource_(ss, cle), true)
    .setAllowInvalid(allowInvalid)
    .build();
}

// ============================================================
// RECHERCHE
// ============================================================
function installer_RECHERCHE_(ss, logs) {
  const sheet = ss.getSheetByName('RECHERCHE');
  if (!sheet) { logs.push('✗ RECHERCHE manquante'); return 0; }

  sheet.getRange('A1:Z100').clearDataValidations();
  let count = 0;

  sheet.getRange('C6').setDataValidation(regleListe_(ss, 'filiales', false)); // Filiale
  count++;
  sheet.getRange('C9').setDataValidation(regleListe_(ss, 'produits', true));  // Produit (filtre)
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
  let count = 0;

  // A = Région (→ responsable régional)
  sheet.getRange('A2:A201').setDataValidation(regleListe_(ss, 'regions', false));
  count++;

  // B = Filiale
  sheet.getRange('B2:B201').setDataValidation(regleListe_(ss, 'filiales', false));
  count++;

  // C et D = Commerciaux filtrés par filiale (cascade via _DDL_TRANSFERTS)
  // Chaque ligne a sa propre plage source → 200 règles distinctes, appliquées
  // en 2 appels groupés (setDataValidations) au lieu de 400.
  const ddlSheet = ss.getSheetByName('_DDL_TRANSFERTS');
  if (ddlSheet) {
    const regles = [];
    for (let i = 2; i <= 201; i++) {
      const sourceRange = ddlSheet.getRange(i, 1, 1, 30); // A:AD de la ligne i
      regles.push([
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(sourceRange, true)
          .setAllowInvalid(true)
          .build()
      ]);
    }
    sheet.getRange(2, 3, 200, 1).setDataValidations(regles); // colonne C en 1 appel
    sheet.getRange(2, 4, 200, 1).setDataValidations(regles); // colonne D en 1 appel
    count += 2;
  } else {
    logs.push('⚠ _DDL_TRANSFERTS manquante — colonnes C et D sans cascade');
  }

  sheet.getRange('G2:G201').setDataValidation(regleListe_(ss, 'motifs', true));     // G = Motif
  count++;
  sheet.getRange('H2:H201').setDataValidation(regleListe_(ss, 'actif', false));     // H = Actif
  count++;
  sheet.getRange('I2:I201').setDataValidation(regleListe_(ss, 'produitsT', false)); // I = Produit (+ Tous)
  count++;

  logs.push(`✓ TRANSFERTS : ${count} validations (dont 200 cascades C+D)`);
  return count;
}

// ============================================================
// AFFECTATIONS_COMMUNES
// ============================================================
function installer_AFFECTATIONS_(ss, logs) {
  const sheet = ss.getSheetByName('AFFECTATIONS_COMMUNES');
  if (!sheet) { logs.push('✗ AFFECTATIONS_COMMUNES manquante'); return 0; }

  const lastRow = sheet.getLastRow();
  const nbLignes = lastRow + 50 - 1; // marge pour ajouts futurs
  let count = 0;

  sheet.getRange(2, 1,  nbLignes, 1).setDataValidation(regleListe_(ss, 'regions',   true));  // A = Région
  count++;
  sheet.getRange(2, 2,  nbLignes, 1).setDataValidation(regleListe_(ss, 'filiales',  true));  // B = Filiale
  count++;
  sheet.getRange(2, 7,  nbLignes, 1).setDataValidation(regleListe_(ss, 'actif',     false)); // G = Actif
  count++;
  sheet.getRange(2, 10, nbLignes, 1).setDataValidation(regleListe_(ss, 'roles',     true));  // J = Rôle
  count++;
  sheet.getRange(2, 11, nbLignes, 1).setDataValidation(regleListe_(ss, 'priorites', true));  // K = Priorité
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
  const nbLignes = lastRow + 50 - 1;
  let count = 0;

  sheet.getRange(2, 1, nbLignes, 1).setDataValidation(regleListe_(ss, 'filiales', true)); // A = Filiale
  count++;
  sheet.getRange(2, 3, nbLignes, 1).setDataValidation(regleListe_(ss, 'produits', true)); // C = Produit
  count++;
  sheet.getRange(2, 4, nbLignes, 1).setDataValidation(regleListe_(ss, 'actif',    false));// D = Actif
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

  sheet.getRange('A2:A201').setDataValidation(regleListe_(ss, 'filiales', true)); // A = Filiale
  count++;
  sheet.getRange('D2:D201').setDataValidation(regleListe_(ss, 'produits', true)); // D = Produit
  count++;
  sheet.getRange('F2:F201').setDataValidation(regleListe_(ss, 'actif',    false));// F = Actif
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
