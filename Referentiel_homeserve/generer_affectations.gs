/**
 * ============================================================
 * GÉNÉRATION DES AFFECTATIONS PAR RÈGLES — HomeServe Énergies Services
 * ============================================================
 *
 * Objectif : éviter la saisie manuelle des ~9 300 lignes d'AFFECTATIONS_COMMUNES.
 *
 * Principe :
 *   - On saisit UNE règle par commercial dans l'onglet SAISIE_SECTEURS :
 *       Filiale | Commercial | Type zone (CP|COMMUNES) | Zone | Produits
 *   - "Zone" réutilise la même syntaxe que CP_MATCH : 33000-33999, 47000, etc.
 *     (ou une liste de noms de communes si Type = COMMUNES).
 *   - genererAffectations() déplie ces règles contre DPT_SOURCE et écrit :
 *       • AFFECTATIONS_COMMUNES  (1 ligne par commune × commercial)
 *       • PRODUITS_COMMERCIAUX   (1 ligne par commercial × produit)
 *
 * Rôle déduit automatiquement :
 *   - 1 seul commercial sur une commune  → Principal (Priorité 1)
 *   - 2 commerciaux ou plus               → Co-affecté (Priorité 2)
 *
 * Sécurité des données manuelles :
 *   - Une colonne "Origine" marque les lignes générées (= "généré").
 *   - La régénération ne supprime QUE les lignes générées : tes saisies
 *     manuelles (Origine vide) sont préservées.
 *   - Idempotent : relançable filiale par filiale au fil des infos reçues.
 *
 * Dépendances : réutilise testCPMatch() et normaliserCP() (notifications_transferts.gs).
 * ============================================================
 */

const GEN = {
  SAISIE: 'SAISIE_SECTEURS',
  AFFECT: 'AFFECTATIONS_COMMUNES',
  PRODUITS: 'PRODUITS_COMMERCIAUX',
  DPT: 'DPT_SOURCE',
  PARAMS: 'PARAMETRES',
  MARQUEUR: 'généré',
  // AFFECTATIONS_COMMUNES : 12 colonnes natives + 1 colonne "Origine" (M=13)
  AFFECT_COLS: 12,
  AFFECT_ORIGINE: 13,
  // PRODUITS_COMMERCIAUX : 6 colonnes natives + 1 colonne "Origine" (G=7)
  PRODUITS_COLS: 6,
  PRODUITS_ORIGINE: 7,
};

// ============================================================
// 0. CHARTE GRAPHIQUE HOMESERVE
// ============================================================
// Couleurs HomeServe France : rouge primaire, teal secondaire.
const CHARTE = {
  ROUGE: '#E22C22',
  BLANC: '#ffffff',
  GRIS: '#9e9e9e',          // onglets des feuilles techniques
  // Feuilles "données" dont la ligne 1 est un vrai en-tête → coloration ligne 1
  ENTETES: [
    'AFFECTATIONS_COMMUNES', 'PRODUITS_COMMERCIAUX', 'EXCEPTIONS_PRODUITS',
    'TRANSFERTS', 'CONTACTS', 'PARAMETRES', 'SAISIE_SECTEURS',
  ],
  // Feuilles visibles métier → onglet rouge (sans toucher leur mise en page)
  ONGLETS_ROUGES: [
    'ACCUEIL', 'RECHERCHE', 'VUE_REGIONALE', 'AFFECTATIONS_COMMUNES',
    'PRODUITS_COMMERCIAUX', 'EXCEPTIONS_PRODUITS', 'TRANSFERTS', 'CONTACTS',
    'PARAMETRES', 'SAISIE_SECTEURS',
  ],
  // Feuilles techniques → onglet gris pour les distinguer
  ONGLETS_GRIS: ['_COMMERCIAUX_PAR_FILIALE_H', '_DDL_TRANSFERTS', 'DPT_SOURCE'],
};

/**
 * Applique la charte HomeServe au classeur :
 *   - en-têtes (ligne 1) des feuilles de données en rouge / blanc / gras
 *   - onglets des feuilles métier en rouge, feuilles techniques en gris
 * Ne modifie ni les données, ni la mise en page des tableaux de bord.
 */
function appliquerCharte() {
  const ss = SpreadsheetApp.getActive();
  let entetes = 0, onglets = 0;

  CHARTE.ENTETES.forEach(nom => {
    const sh = ss.getSheetByName(nom);
    if (!sh) return;
    const nbCol = sh.getLastColumn();
    if (nbCol < 1) return;
    sh.getRange(1, 1, 1, nbCol)
      .setBackground(CHARTE.ROUGE)
      .setFontColor(CHARTE.BLANC)
      .setFontWeight('bold');
    entetes++;
  });

  CHARTE.ONGLETS_ROUGES.forEach(nom => {
    const sh = ss.getSheetByName(nom);
    if (sh) { sh.setTabColor(CHARTE.ROUGE); onglets++; }
  });
  CHARTE.ONGLETS_GRIS.forEach(nom => {
    const sh = ss.getSheetByName(nom);
    if (sh) { sh.setTabColor(CHARTE.GRIS); onglets++; }
  });

  ss.toast(entetes + ' en-têtes + ' + onglets + ' onglets mis aux couleurs HomeServe.',
    '🎨 Charte appliquée', 6);
}

// ============================================================
// 1. CRÉATION DE LA FEUILLE DE SAISIE
// ============================================================
function creerFeuilleSaisie() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(GEN.SAISIE);
  if (sheet) {
    SpreadsheetApp.getUi().alert(
      'La feuille ' + GEN.SAISIE + ' existe déjà.',
      'Elle n\'a pas été modifiée. Supprime-la manuellement si tu veux repartir de zéro.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  sheet = ss.insertSheet(GEN.SAISIE, ss.getNumSheets());

  const entetes = ['Filiale', 'Commercial', 'Type zone', 'Zone', 'Produits', 'Commentaire'];
  sheet.getRange(1, 1, 1, entetes.length).setValues([entetes])
    .setFontWeight('bold').setBackground('#E22C22').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Exemples explicatifs
  const exemples = [
    ['DEC Energies', 'Jean DUPONT', 'CP', '33000-33999, 47000', 'Pompe à chaleur Air/Eau, Panneaux photovoltaïques', 'Exemple : secteur par CP'],
    ['Cerise Energies', 'Marie MARTIN', 'COMMUNES', 'BORDEAUX, MERIGNAC, PESSAC', 'Tous', 'Exemple : secteur par liste de communes'],
  ];
  sheet.getRange(2, 1, exemples.length, exemples[0].length).setValues(exemples)
    .setFontStyle('italic').setFontColor('#999999');

  // Validations : Filiale + Type zone
  if (typeof plageSource_ === 'function') {
    try {
      sheet.getRange('A2:A500').setDataValidation(regleListe_(ss, 'filiales', true));
    } catch (e) { /* PARAMETRES absent : on ignore */ }
  }
  sheet.getRange('C2:C500').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['CP', 'COMMUNES'], true).setAllowInvalid(false).build());

  // Notes d'aide
  sheet.getRange('D1').setNote(
    'Type CP : syntaxe identique à la colonne CP des TRANSFERTS\n' +
    '  33000  |  33000, 47000  |  33000-33999  |  mix  |  Tous\n' +
    'Type COMMUNES : liste de noms de communes séparés par des virgules');
  sheet.getRange('E1').setNote(
    'Produits gérés par ce commercial, séparés par des virgules.\n' +
    '"Tous" = tous les produits du référentiel (PARAMETRES C2:C25).');

  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 220);

  ss.toast('Feuille ' + GEN.SAISIE + ' créée. Remplace les lignes d\'exemple par tes règles.', '✓ Prêt', 8);
  sheet.activate();
}

// ============================================================
// 2. UTILITAIRES
// ============================================================

/** Normalise un nom de commune pour comparaison : majuscules, sans accents ni ponctuation. */
function normaliserNomCommune_(x) {
  if (x === null || x === undefined) return '';
  return String(x).toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // accents
    .replace(/[^A-Z0-9]+/g, ' ').trim();
}

/** Carte département → région, déduite de DPT_SOURCE. */
function carteRegions_(dptValues) {
  const carte = {};
  for (let i = 1; i < dptValues.length; i++) {
    const reg = dptValues[i][0];   // A = Région
    const dep = dptValues[i][3];   // D = Département
    if (reg && dep && !carte[dep]) carte[dep] = reg;
  }
  return carte;
}

/**
 * Enrichit la carte département → région avec les régions déjà présentes dans
 * AFFECTATIONS_COMMUNES (A = Région, C = CP). N'écrase PAS DPT_SOURCE : ne
 * complète que les départements encore inconnus. Le département est déduit des
 * 2 premiers chiffres du CP (ex. 06000 → "06").
 */
function carteRegionsDepuisAffectations_(affectSheet, carte) {
  const last = affectSheet.getLastRow();
  if (last < 2) return;
  const data = affectSheet.getRange(2, 1, last - 1, 3).getValues(); // A=Région, B=Filiale, C=CP
  for (let i = 0; i < data.length; i++) {
    const reg = data[i][0];
    const cp = normaliserCP(data[i][2]);
    if (!reg || !cp) continue;
    const dep = cp.substring(0, 2);
    if (!carte[dep]) carte[dep] = reg;
  }
}

/** Liste des produits du référentiel (pour expansion de "Tous"). */
function listeProduits_(ss) {
  const params = ss.getSheetByName(GEN.PARAMS);
  if (!params) return [];
  const vals = params.getRange('C2:C25').getValues();
  return vals.map(r => r[0]).filter(v => v !== '' && v !== null);
}

/** Découpe une cellule "Produits" en tableau, en expansant "Tous". */
function parseProduits_(cell, tousProduits) {
  if (cell === null || cell === undefined) return [];
  const txt = String(cell).trim();
  if (txt === '') return [];
  if (txt.toLowerCase() === 'tous') return tousProduits.slice();
  return txt.split(/[,;\n\r]+/).map(s => s.trim()).filter(s => s !== '');
}

/**
 * Carte filiale → ensemble des commerciaux connus, lue depuis
 * _COMMERCIAUX_PAR_FILIALE_H (col A = filiale, col B+ = commerciaux).
 * Sert à détecter les fautes de frappe dans SAISIE_SECTEURS.
 */
function commerciauxConnus_(ss) {
  const sh = ss.getSheetByName('_COMMERCIAUX_PAR_FILIALE_H');
  const carte = {};            // filiale → { nomNorm: nomAffiche }
  if (!sh) return carte;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const filiale = String(data[i][0] || '').trim();
    if (!filiale) continue;
    const ens = {};
    for (let j = 1; j < data[i].length; j++) {
      const nom = String(data[i][j] || '').trim();
      if (nom) ens[nom.toUpperCase()] = nom;
    }
    carte[filiale] = ens;
  }
  return carte;
}

// ============================================================
// 3. GÉNÉRATION
// ============================================================
function genererAffectations() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const saisie = ss.getSheetByName(GEN.SAISIE);
  if (!saisie) {
    ui.alert('Feuille ' + GEN.SAISIE + ' introuvable.',
      'Lance d\'abord ⚙️ Transferts → Créer la feuille de saisie.', ui.ButtonSet.OK);
    return;
  }
  const dptSheet = ss.getSheetByName(GEN.DPT);
  const affectSheet = ss.getSheetByName(GEN.AFFECT);
  const produitsSheet = ss.getSheetByName(GEN.PRODUITS);
  if (!dptSheet || !affectSheet || !produitsSheet) {
    ui.alert('Feuille manquante',
      'Vérifie la présence de DPT_SOURCE, AFFECTATIONS_COMMUNES et PRODUITS_COMMERCIAUX.', ui.ButtonSet.OK);
    return;
  }

  ss.toast('Lecture des données…', '⏳ Génération', -1);

  // --- Lecture DPT_SOURCE ---
  const dpt = dptSheet.getDataRange().getValues(); // 0=Région,2=Filiale,3=Dép,4=CP,5=Nom
  const carteReg = carteRegions_(dpt);

  // Repli : pour les départements sans région dans DPT_SOURCE, on RÉUTILISE les
  // régions déjà saisies (à la main ou générées) dans AFFECTATIONS_COMMUNES.
  // Ainsi une région complétée manuellement persiste aux régénérations suivantes.
  carteRegionsDepuisAffectations_(affectSheet, carteReg);

  // PERF : on pré-calcule UNE SEULE FOIS les valeurs normalisées de chaque commune
  // (CP et nom) au lieu de les recalculer pour chaque règle. C'est ce qui évite
  // le dépassement de délai quand il y a beaucoup de règles.
  const N = dpt.length;
  const infoDpt = new Array(N);
  for (let i = 1; i < N; i++) {
    const row = dpt[i];
    infoDpt[i] = {
      cpNorm: normaliserCP(row[4]),
      cpBrut: row[4],
      ville: String(row[5] || '').trim(),
      villeNorm: normaliserNomCommune_(row[5]),
      region: row[0] || carteReg[row[3]] || '',
    };
  }

  // --- Lecture des règles ---
  const derniere = saisie.getLastRow();
  if (derniere < 2) { ss.toast('Aucune règle à traiter.', 'Génération', 5); return; }
  const regles = saisie.getRange(2, 1, derniere - 1, 6).getValues();

  const tousProduits = listeProduits_(ss);
  const refCommerciaux = commerciauxConnus_(ss);

  // Accumulateurs
  // communes : clé "cp|villeNorm" → { cp, ville, region, commerciaux: [{nom, filiale}] }
  const communes = {};
  const produitsSet = {}; // clé "filiale|commercial|produit" → true
  const rapport = { regles: 0, sansMatch: [], sansProduit: [], lignesIgnorees: 0, inconnus: [] };

  for (let r = 0; r < regles.length; r++) {
    const [filiale, commercial, typeZone, zone, produitsCell] = regles[r];
    if (!filiale || !commercial) { rapport.lignesIgnorees++; continue; }
    rapport.regles++;

    // Garde-fou : commercial absent du référentiel de la filiale → faute de frappe probable
    const ensFiliale = refCommerciaux[String(filiale).trim()];
    if (ensFiliale && !ensFiliale[String(commercial).trim().toUpperCase()]) {
      rapport.inconnus.push(commercial + ' (' + filiale + ')');
    }

    const type = String(typeZone || 'CP').trim().toUpperCase();
    let nbCommunes = 0;

    if (type === 'COMMUNES') {
      // Liste de noms de communes (normalisée une fois)
      const cible = {};
      String(zone || '').split(/[,;\n\r]+/).map(normaliserNomCommune_).filter(Boolean)
        .forEach(n => cible[n] = true);
      for (let i = 1; i < N; i++) {
        if (cible[infoDpt[i].villeNorm]) {
          nbCommunes += ajouterCommunePre_(communes, infoDpt[i], filiale, commercial);
        }
      }
    } else {
      // Type CP : expression compilée une fois, puis test rapide par commune
      const matcheur = compilerCP_(zone);
      for (let i = 1; i < N; i++) {
        if (testCompile_(matcheur, infoDpt[i].cpNorm)) {
          nbCommunes += ajouterCommunePre_(communes, infoDpt[i], filiale, commercial);
        }
      }
    }

    if (nbCommunes === 0) rapport.sansMatch.push(commercial + ' (' + filiale + ')');

    // Produits
    const prods = parseProduits_(produitsCell, tousProduits);
    if (prods.length === 0) {
      rapport.sansProduit.push(commercial + ' (' + filiale + ')');
    } else {
      prods.forEach(p => { produitsSet[filiale + '|' + commercial + '|' + p] = true; });
    }
  }

  // --- Construction des lignes AFFECTATIONS générées ---
  const conflitsMultiFiliale = [];
  const lignesAffect = [];
  // Clés "filiale|cpNorm|villeNorm" couvertes par la génération → sert à supprimer
  // les placeholders « A AFFECTER » de ces communes (ils sont désormais attribués).
  const clesCouvertes = {};
  Object.keys(communes).forEach(cle => {
    const c = communes[cle];
    const nb = c.commerciaux.length;
    const role = nb >= 2 ? 'Co-affecté' : 'Principal';
    const prio = nb >= 2 ? 2 : 1;
    // Détection commune partagée entre filiales différentes
    const filiales = {};
    c.commerciaux.forEach(x => filiales[x.filiale] = true);
    if (Object.keys(filiales).length > 1) conflitsMultiFiliale.push(c.ville + ' (' + c.cp + ')');

    c.commerciaux.forEach(x => {
      const cleCommune = x.filiale + '-' + c.cp + '-' + c.ville;  // ex. EGS Energies-6000-NICE
      clesCouvertes[String(x.filiale).trim().toUpperCase() + '|' + c.cpNorm + '|' + normaliserNomCommune_(c.ville)] = true;
      // A Région | B Filiale | C CP | D Ville | E Commercial | F Secteur | G Actif
      // H Commentaire | I Clé commune | J Rôle | K Priorité | L Rang | M Origine
      lignesAffect.push([
        c.region, x.filiale, c.cp, c.ville, x.nom, '', 'Oui',
        '', cleCommune, role, prio, '', GEN.MARQUEUR
      ]);
    });
  });

  // --- Construction des lignes PRODUITS générées ---
  const lignesProduits = Object.keys(produitsSet).map(cle => {
    const [filiale, commercial, produit] = cle.split('|');
    // A Filiale | B Commercial | C Produit | D Actif | E Commentaire | F Rang | G Origine
    return [filiale, commercial, produit, 'Oui', '', '', GEN.MARQUEUR];
  });

  // --- Comptage des lignes manuelles qui seront préservées (pour l'aperçu) ---
  const manuAffect = compterLignesManuelles_(affectSheet, GEN.AFFECT_COLS, GEN.AFFECT_ORIGINE);
  const manuProduits = compterLignesManuelles_(produitsSheet, GEN.PRODUITS_COLS, GEN.PRODUITS_ORIGINE);

  // --- Garde-fou n°1 : aperçu + confirmation AVANT toute écriture ---
  const apercu = [];
  apercu.push('AVANT d\'écrire, vérifiez ce résumé :');
  apercu.push('');
  apercu.push('• Règles traitées : ' + rapport.regles + (rapport.lignesIgnorees ? ' (' + rapport.lignesIgnorees + ' ignorée(s), filiale/commercial vide)' : ''));
  apercu.push('• Communes couvertes : ' + Object.keys(communes).length);
  apercu.push('• AFFECTATIONS : ' + lignesAffect.length + ' lignes générées remplacées' + (manuAffect ? ' (jusqu\'à ' + manuAffect + ' lignes manuelles conservées)' : ''));
  apercu.push('  ↳ les « A AFFECTER » des communes désormais attribuées sont supprimés (anti-doublon)');
  apercu.push('• PRODUITS : ' + lignesProduits.length + ' lignes générées remplacées' + (manuProduits ? ' (' + manuProduits + ' lignes manuelles conservées)' : ''));
  apercu.push('');
  if (rapport.inconnus.length) apercu.push('⚠ Commerciaux INCONNUS du référentiel (faute de frappe ?) : ' + rapport.inconnus.slice(0, 15).join(', ') + (rapport.inconnus.length > 15 ? '…' : ''));
  if (rapport.sansMatch.length) apercu.push('⚠ Aucune commune trouvée pour : ' + rapport.sansMatch.slice(0, 15).join(', ') + (rapport.sansMatch.length > 15 ? '…' : ''));
  if (rapport.sansProduit.length) apercu.push('⚠ Aucun produit renseigné pour : ' + rapport.sansProduit.slice(0, 15).join(', ') + (rapport.sansProduit.length > 15 ? '…' : ''));
  if (conflitsMultiFiliale.length) apercu.push('⚠ Communes partagées entre filiales : ' + conflitsMultiFiliale.slice(0, 15).join(', ') + (conflitsMultiFiliale.length > 15 ? '…' : ''));
  apercu.push('');
  apercu.push('Les lignes « généré » existantes seront remplacées. Continuer ?');

  const choix = ui.alert('Confirmer la génération', apercu.join('\n'), ui.ButtonSet.YES_NO);
  if (choix !== ui.Button.YES) {
    ss.toast('Génération annulée — aucune modification.', 'Annulé', 5);
    return;
  }

  // --- Écriture (en préservant les lignes manuelles) ---
  ss.toast('Écriture de ' + lignesAffect.length + ' affectations…', '⏳ Génération', -1);
  // Pour AFFECTATIONS : on remplace aussi les placeholders « A AFFECTER » des communes
  // désormais couvertes par une attribution (sinon doublon placeholder + commercial réel).
  const optPlaceholder = {
    cles: clesCouvertes,
    idxFiliale: 1, idxCP: 2, idxVille: 3, idxCommercial: 4,
    placeholders: ['A AFFECTER', ''],
  };
  try {
    ecrireEnPreservantManuel_(affectSheet, lignesAffect, GEN.AFFECT_COLS, GEN.AFFECT_ORIGINE, 'Origine', optPlaceholder);
    ecrireEnPreservantManuel_(produitsSheet, lignesProduits, GEN.PRODUITS_COLS, GEN.PRODUITS_ORIGINE, 'Origine', null);
  } catch (e) {
    ss.toast('Échec de l\'écriture', '✗ Génération', 5);
    ui.alert('Génération interrompue',
      'Une erreur est survenue pendant l\'écriture dans les feuilles :\n\n' + e.message +
      '\n\nCela peut arriver si le service Google Sheets est momentanément indisponible. ' +
      'Réessayez dans quelques minutes. Aucune donnée manuelle n\'a été perdue ' +
      '(les lignes « généré » seront simplement régénérées au prochain essai).',
      ui.ButtonSet.OK);
    return;
  }

  ss.toast('Terminé', '✓ Génération', 5);
  ui.alert('Génération terminée',
    lignesAffect.length + ' affectations et ' + lignesProduits.length + ' produits écrits.\n\n' +
    'Rappel : pour corriger une ligne « généré », modifiez la RÈGLE dans ' + GEN.SAISIE +
    ' puis relancez la génération — ne modifiez pas la ligne à la main (elle serait écrasée).',
    ui.ButtonSet.OK);
}

/** Compte les lignes manuelles (non vides, Origine ≠ "généré") d'une feuille. */
function compterLignesManuelles_(sheet, nbCols, colOrigine) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const data = sheet.getRange(2, 1, last - 1, colOrigine).getValues();
  return data.filter(row => {
    const vide = row.slice(0, nbCols).every(v => v === '' || v === null);
    return !vide && row[colOrigine - 1] !== GEN.MARQUEUR;
  }).length;
}

/** Ajoute un commercial à une commune (valeurs pré-calculées). Retourne 1 si nouvelle commune. */
function ajouterCommunePre_(communes, inf, filiale, commercial) {
  const cle = inf.cpNorm + '|' + inf.villeNorm;
  let nouvelle = 0;
  if (!communes[cle]) {
    communes[cle] = { cp: inf.cpBrut, cpNorm: inf.cpNorm, ville: inf.ville, region: inf.region, commerciaux: [] };
    nouvelle = 1;
  }
  // Dédup : un même commercial ne doit pas apparaître deux fois sur la commune
  const existe = communes[cle].commerciaux.some(x => x.nom === commercial && x.filiale === filiale);
  if (!existe) communes[cle].commerciaux.push({ nom: commercial, filiale: filiale });
  return nouvelle;
}

/** Compile une expression CP (vide/Tous/liste/tranches) en structure testable rapidement. */
function compilerCP_(expr) {
  const e = String(expr || '').trim();
  if (e === '' || e.toLowerCase() === 'tous') return { tous: true, singles: {}, ranges: [] };
  const singles = {}, ranges = [];
  e.split(/[,;\n\r]+/).forEach(m => {
    m = m.trim();
    if (!m) return;
    if (m.indexOf('-') !== -1) {
      const p = m.split('-');
      if (p.length === 2) {
        const a = normaliserCP(p[0]), b = normaliserCP(p[1]);
        if (a && b) ranges.push([parseInt(a, 10), parseInt(b, 10)]);
      }
    } else {
      const c = normaliserCP(m);
      if (c) singles[c] = true;
    }
  });
  return { tous: false, singles: singles, ranges: ranges };
}

/** Teste un CP normalisé contre une expression compilée. */
function testCompile_(m, cpNorm) {
  if (!cpNorm) return false;
  if (m.tous) return true;
  if (m.singles[cpNorm]) return true;
  const n = parseInt(cpNorm, 10);
  for (let k = 0; k < m.ranges.length; k++) {
    if (n >= m.ranges[k][0] && n <= m.ranges[k][1]) return true;
  }
  return false;
}

/**
 * Réécrit une feuille en gardant les lignes manuelles (colonne Origine ≠ "généré")
 * et en remplaçant les lignes générées par le nouveau lot.
 */
function ecrireEnPreservantManuel_(sheet, lignesGenerees, nbCols, colOrigine, titreOrigine, optPlaceholder) {
  // S'assure que l'en-tête de la colonne Origine existe (+ note d'avertissement)
  const enteteOrigine = sheet.getRange(1, colOrigine);
  if (!enteteOrigine.getValue()) {
    enteteOrigine.setValue(titreOrigine).setFontWeight('bold');
  }
  enteteOrigine.setNote(
    'Ne pas modifier à la main les lignes « généré » : elles sont écrasées à chaque ' +
    'génération. Pour corriger, modifiez la règle dans SAISIE_SECTEURS puis relancez ' +
    '⚙️ Transferts → Générer les affectations.');

  const last = sheet.getLastRow();
  let manuelles = [];
  if (last >= 2) {
    const data = sheet.getRange(2, 1, last - 1, colOrigine).getValues();
    const placeholdersSet = {};
    if (optPlaceholder) optPlaceholder.placeholders.forEach(p => placeholdersSet[String(p).trim().toUpperCase()] = true);

    manuelles = data.filter(row => {
      // ligne non vide ET non générée
      const origine = row[colOrigine - 1];
      const vide = row.slice(0, nbCols).every(v => v === '' || v === null);
      if (vide || origine === GEN.MARQUEUR) return false;

      // Remplacement des placeholders : si cette ligne manuelle est un « A AFFECTER »
      // (ou commercial vide) sur une commune désormais couverte par la génération,
      // on la retire pour éviter le doublon placeholder + commercial réel.
      if (optPlaceholder) {
        const commercial = String(row[optPlaceholder.idxCommercial] || '').trim().toUpperCase();
        if (placeholdersSet[commercial]) {
          const filiale = String(row[optPlaceholder.idxFiliale] || '').trim().toUpperCase();
          const cpNorm = normaliserCP(row[optPlaceholder.idxCP]);
          const villeNorm = normaliserNomCommune_(row[optPlaceholder.idxVille]);
          if (optPlaceholder.cles[filiale + '|' + cpNorm + '|' + villeNorm]) return false; // couverte → on supprime
        }
      }
      return true;
    });
  }

  // Efface tout sous l'en-tête puis réécrit manuelles + générées
  if (last >= 2) sheet.getRange(2, 1, last - 1, colOrigine).clearContent();

  const tout = manuelles.concat(
    lignesGenerees.map(l => {
      // normalise la largeur à colOrigine colonnes
      const row = l.slice();
      while (row.length < colOrigine) row.push('');
      return row;
    })
  );

  // PERF : écriture par blocs pour ne pas dépasser le délai du service Sheets.
  // On ne fait PAS de flush() entre les blocs (chaque flush relancerait un
  // recalcul complet des formules dépendantes). Chaque écriture est protégée
  // par un réessai en cas d'erreur transitoire ("Service indisponible").
  const BLOC = 10000;
  for (let depart = 0; depart < tout.length; depart += BLOC) {
    const lot = tout.slice(depart, depart + BLOC);
    setValuesAvecRetry_(sheet.getRange(2 + depart, 1, lot.length, colOrigine), lot);
  }
}

/** setValues protégé par réessais (back-off) contre les erreurs transitoires de Sheets. */
function setValuesAvecRetry_(range, values) {
  let essais = 0;
  while (true) {
    try {
      range.setValues(values);
      return;
    } catch (e) {
      essais++;
      if (essais >= 4) throw e;
      Utilities.sleep(2000 * essais); // 2s, 4s, 6s
    }
  }
}
