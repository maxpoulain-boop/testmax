/**
 * ============================================================
 * NOTIFICATIONS TRANSFERTS + CP_MATCH - HomeServe Énergies Services
 * ============================================================
 *
 * v2 — Ajoute la fonction CP_MATCH pour gérer les listes et tranches de CP
 *      dans la colonne CP de TRANSFERTS.
 *
 * Syntaxes acceptées dans la colonne CP :
 *   - (vide) ou "Tous"        → tous les CP
 *   - 33000                    → un seul CP
 *   - 33000, 33100, 33200      → liste séparée par virgules
 *   - 33000-33999              → tranche
 *   - 33000-33500, 35000       → mix tranche + valeur
 *   - copier-coller depuis Excel (sauts de ligne) → accepté
 *
 * ============================================================
 * INSTALLATION
 * ============================================================
 *
 * 1. Dans ton Google Sheet, va dans Extensions → Apps Script
 * 2. Sélectionne tout le code actuel et remplace-le par ce fichier
 * 3. Sauvegarde (Ctrl+S)
 * 4. Les déclencheurs déjà installés (setupTriggers) continuent de fonctionner.
 *    Pas besoin de relancer setupTriggers.
 * 5. La fonction CP_MATCH devient automatiquement utilisable dans les formules
 *    du Sheet (=CP_MATCH(cp, expression))
 *
 * ============================================================
 */

// === Configuration des feuilles et colonnes ===
const SHEET_TRANSFERTS = 'TRANSFERTS';
const SHEET_CONTACTS = 'CONTACTS';

const COL_FILIALE = 1;
const COL_REMPLACE = 2;
const COL_NOUVEAU = 3;
const COL_DATE_DEBUT = 4;
const COL_DATE_FIN = 5;
const COL_MOTIF = 6;
const COL_ACTIF = 7;
const COL_PRODUIT = 8;
const COL_CP = 9;
const COL_VILLE = 10;
const COL_COMMENTAIRE = 11;

const CELL_MODE = 'C6';
const CELL_EMAIL_TEST = 'C9';
const CELL_PROD_1 = 'C12';
const CELL_PROD_2 = 'C13';
const CELL_PROD_REGIONAL = 'C14';

const PROP_ROW_SENT = 'sent_validation_';
const PROP_ROW_REMINDER = 'sent_reminder_';
const PROP_LAST_STATE = 'last_state_';

// ============================================================
// FONCTION CUSTOM : CP_MATCH(cp_cherche, expression)
// ============================================================

/**
 * Vérifie si un CP correspond à une expression CP (vide, valeur, liste, tranche).
 *
 * @param {string|number} cpCherche  Le CP à tester (ex: "33500" ou 33500).
 * @param {string|range} expression  Le CP de référence dans TRANSFERTS.
 *                                   Peut être une cellule ou une plage entière.
 * @return TRUE/FALSE selon que le CP est couvert par l'expression.
 *         Si expression est une plage, renvoie un tableau de TRUE/FALSE.
 * @customfunction
 */
function CP_MATCH(cpCherche, expression) {
  const cp = normaliserCP(cpCherche);

  // Si expression est une plage (2D array), on map ligne par ligne
  if (Array.isArray(expression)) {
    return expression.map(function(row) {
      if (Array.isArray(row)) {
        return row.map(function(cell) { return testCPMatch(cp, cell); });
      }
      return testCPMatch(cp, row);
    });
  }

  return testCPMatch(cp, expression);
}

function normaliserCP(x) {
  if (x === null || x === undefined || x === '') return '';
  var s = String(x).trim();
  // Garde uniquement les chiffres
  s = s.replace(/\D/g, '');
  if (s === '') return '';
  // Padding à 5 chiffres si CP français < 10000 (ex: "1000" → "01000")
  while (s.length < 5) s = '0' + s;
  return s;
}

function testCPMatch(cp, expression) {
  if (cp === '') return false;
  if (expression === null || expression === undefined || expression === '') return true;

  var expr = String(expression).trim();
  if (expr === '' || expr.toLowerCase() === 'tous') return true;

  // Découpe sur virgule, point-virgule, saut de ligne, retour chariot, espace multiple
  var morceaux = expr.split(/[,;\n\r]+/);

  for (var i = 0; i < morceaux.length; i++) {
    var m = morceaux[i].trim();
    if (m === '') continue;

    // Tranche "33000-33999"
    if (m.indexOf('-') !== -1) {
      var parts = m.split('-');
      if (parts.length === 2) {
        var a = normaliserCP(parts[0]);
        var b = normaliserCP(parts[1]);
        if (a !== '' && b !== '') {
          var cpNum = parseInt(cp, 10);
          var aNum = parseInt(a, 10);
          var bNum = parseInt(b, 10);
          if (cpNum >= aNum && cpNum <= bNum) return true;
        }
      }
    } else {
      // Valeur exacte
      var v = normaliserCP(m);
      if (v === cp) return true;
    }
  }

  return false;
}

// ============================================================
// DÉCLENCHEUR onEdit : exécuté à chaque modification
// ============================================================

function onEdit(e) {
  if (!e || !e.range) return;
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_TRANSFERTS) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();

    if (col !== COL_ACTIF || row < 2) return;

    const newValue = e.value || '';
    const oldValue = e.oldValue || 'Non';

    Logger.log('Edit detecte ligne ' + row + ' : ' + oldValue + ' -> ' + newValue);

    if (newValue === 'Oui' && oldValue !== 'Oui') {
      envoyerEmailValidation(row);
    }

    if (oldValue === 'Oui' && newValue === 'Non') {
      envoyerEmailAnnulation(row);
    }
  } catch (err) {
    Logger.log('Erreur onEdit : ' + err.message);
  }
}

// ============================================================
// DÉCLENCHEUR QUOTIDIEN : rappels J-3 ouvrés
// ============================================================

function verifierRappelsQuotidiens() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_TRANSFERTS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const aujourdHui = nouvelleDate(new Date());

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const filiale = data[i][COL_FILIALE - 1];
    const actif = data[i][COL_ACTIF - 1];
    const dateDebut = data[i][COL_DATE_DEBUT - 1];

    if (!filiale || actif !== 'Oui' || !(dateDebut instanceof Date)) continue;

    const dateRappel = soustraireJoursOuvres(nouvelleDate(dateDebut), 3);

    if (memeJour(aujourdHui, dateRappel)) {
      const props = PropertiesService.getDocumentProperties();
      const key = PROP_ROW_REMINDER + row;
      if (props.getProperty(key) !== 'true') {
        envoyerEmailRappel(row);
        props.setProperty(key, 'true');
      }
    }
  }
}

// ============================================================
// FONCTIONS D'ENVOI
// ============================================================

function envoyerEmailValidation(row) {
  const infos = lireLigne(row);
  if (!infos.filiale) return;

  const destinataires = obtenirDestinataires();
  if (!destinataires.length) {
    Logger.log('Aucun destinataire configure');
    return;
  }

  const sujet = '[Transfert valide] ' + infos.filiale + ' - ' + infos.remplace + ' -> ' + infos.nouveau;
  const corps = construireCorpsEmail(infos, 'Transfert valide',
    "Un transfert vient d'etre valide dans le referentiel d'affectation commerciale.");

  MailApp.sendEmail({
    to: destinataires.join(','),
    subject: sujet,
    htmlBody: corps,
  });

  PropertiesService.getDocumentProperties().setProperty(PROP_ROW_SENT + row, 'true');
  Logger.log('Email validation envoye ligne ' + row);
}

function envoyerEmailRappel(row) {
  const infos = lireLigne(row);
  if (!infos.filiale) return;

  const destinataires = obtenirDestinataires();
  if (!destinataires.length) return;

  const sujet = '[Rappel J-3] Transfert ' + infos.filiale + ' commence le ' + formaterDate(infos.dateDebut);
  const corps = construireCorpsEmail(infos, 'Rappel : transfert imminent',
    'Ce transfert prendra effet dans 3 jours ouvres. Merci de verifier que tout est en ordre.');

  MailApp.sendEmail({
    to: destinataires.join(','),
    subject: sujet,
    htmlBody: corps,
  });

  Logger.log('Email rappel J-3 envoye ligne ' + row);
}

function envoyerEmailAnnulation(row) {
  const infos = lireLigne(row);
  if (!infos.filiale) return;

  const destinataires = obtenirDestinataires();
  if (!destinataires.length) return;

  const sujet = '[Transfert annule] ' + infos.filiale + ' - ' + infos.remplace + ' -> ' + infos.nouveau;
  const corps = construireCorpsEmail(infos, 'Transfert annule',
    "Un transfert precedemment valide vient d'etre desactive.");

  MailApp.sendEmail({
    to: destinataires.join(','),
    subject: sujet,
    htmlBody: corps,
  });

  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(PROP_ROW_SENT + row);
  props.deleteProperty(PROP_ROW_REMINDER + row);
  Logger.log('Email annulation envoye ligne ' + row);
}

// ============================================================
// UTILITAIRES
// ============================================================

function lireLigne(row) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_TRANSFERTS);
  const data = sheet.getRange(row, 1, 1, 11).getValues()[0];
  return {
    filiale: data[COL_FILIALE - 1] || '',
    remplace: data[COL_REMPLACE - 1] || '',
    nouveau: data[COL_NOUVEAU - 1] || '',
    dateDebut: data[COL_DATE_DEBUT - 1],
    dateFin: data[COL_DATE_FIN - 1],
    motif: data[COL_MOTIF - 1] || '',
    actif: data[COL_ACTIF - 1] || '',
    produit: data[COL_PRODUIT - 1] || 'Tous',
    cp: data[COL_CP - 1] || '',
    ville: data[COL_VILLE - 1] || '',
    commentaire: data[COL_COMMENTAIRE - 1] || '',
  };
}

function obtenirDestinataires() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CONTACTS);
  if (!sheet) return [];

  const mode = (sheet.getRange(CELL_MODE).getValue() || '').toString().toUpperCase();

  if (mode === 'TEST') {
    const email = sheet.getRange(CELL_EMAIL_TEST).getValue();
    return validerEmails([email]);
  }

  if (mode === 'PROD') {
    return validerEmails([
      sheet.getRange(CELL_PROD_1).getValue(),
      sheet.getRange(CELL_PROD_2).getValue(),
      sheet.getRange(CELL_PROD_REGIONAL).getValue(),
    ]);
  }

  return [];
}

function validerEmails(liste) {
  return liste
    .map(function(e) { return (e || '').toString().trim(); })
    .filter(function(e) { return e && e.indexOf('@') !== -1 && e.indexOf('ton.email') === -1; });
}

function construireCorpsEmail(infos, titre, accroche) {
  const dateDebutStr = formaterDate(infos.dateDebut);
  const dateFinStr = infos.dateFin ? formaterDate(infos.dateFin) : '(non precisee)';
  const zone = infos.cp ? infos.cp + (infos.ville ? ' ' + infos.ville : '') : 'Toutes communes';

  return '<div style="font-family: Calibri, Arial, sans-serif; color: #303030; max-width: 600px;">' +
    '<div style="background: #004990; color: white; padding: 16px;">' +
      '<h2 style="margin: 0;">' + titre + '</h2>' +
      '<p style="margin: 4px 0 0; opacity: 0.85;">Referentiel d\'affectation - HomeServe Energies Services</p>' +
    '</div>' +
    '<div style="background: #ffffff; border: 1px solid #d8d8d8; border-top: none; padding: 20px;">' +
      '<p>' + accroche + '</p>' +
      '<table style="border-collapse: collapse; width: 100%; margin-top: 12px;">' +
        ligneTable('Filiale', infos.filiale) +
        ligneTable('Commercial remplace', infos.remplace) +
        ligneTable('Nouveau commercial', infos.nouveau) +
        ligneTable('Date debut', dateDebutStr) +
        ligneTable('Date fin', dateFinStr) +
        ligneTable('Motif', infos.motif) +
        ligneTable('Produit', infos.produit) +
        ligneTable('Zone (CP)', zone) +
        (infos.commentaire ? ligneTable('Commentaire', infos.commentaire) : '') +
      '</table>' +
      '<p style="margin-top: 20px; font-size: 12px; color: #808080;">Email automatique - ne pas repondre.</p>' +
    '</div>' +
  '</div>';
}

function ligneTable(lbl, val) {
  return '<tr><td style="padding: 6px; background: #f4f7fb; font-weight: 600;">' + lbl + '</td>' +
         '<td style="padding: 6px;">' + val + '</td></tr>';
}

function formaterDate(d) {
  if (!(d instanceof Date)) return '(invalide)';
  const dd = ('0' + d.getDate()).slice(-2);
  const mm = ('0' + (d.getMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + d.getFullYear();
}

function nouvelleDate(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function memeJour(d1, d2) {
  return d1.getTime() === d2.getTime();
}

function soustraireJoursOuvres(date, n) {
  let resultat = new Date(date);
  let restants = n;
  while (restants > 0) {
    resultat.setDate(resultat.getDate() - 1);
    const jour = resultat.getDay();
    if (jour !== 0 && jour !== 6) restants--;
  }
  return resultat;
}

// ============================================================
// INSTALLATION ET TESTS
// ============================================================

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  ScriptApp.newTrigger('verifierRappelsQuotidiens')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  SpreadsheetApp.getActive().toast(
    'Declencheurs installes. Le script tournera automatiquement.',
    'Configuration OK',
    5
  );
}

function testEnvoiEmail() {
  const destinataires = obtenirDestinataires();
  if (!destinataires.length) {
    SpreadsheetApp.getUi().alert(
      'Aucun destinataire configure.\n\n' +
      'Verifie l\'onglet CONTACTS :\n' +
      '- C6 contient TEST ou PROD\n' +
      '- C9 contient un vrai email'
    );
    return;
  }

  const sujet = '[TEST] Notification transferts - HomeServe';
  const corps =
    '<div style="font-family: Calibri, Arial, sans-serif;">' +
      '<h2 style="color: #004990;">Test de configuration</h2>' +
      "<p>Cet email confirme que les notifications du referentiel d'affectation fonctionnent.</p>" +
      '<p><strong>Destinataires :</strong> ' + destinataires.join(', ') + '</p>' +
      '<p><strong>Mode :</strong> ' + SpreadsheetApp.getActive().getSheetByName(SHEET_CONTACTS).getRange(CELL_MODE).getValue() + '</p>' +
    '</div>';

  MailApp.sendEmail({ to: destinataires.join(','), subject: sujet, htmlBody: corps });
  SpreadsheetApp.getActive().toast('Email de test envoye a : ' + destinataires.join(', '), 'Test OK', 8);
}

function afficherDestinataires() {
  const destinataires = obtenirDestinataires();
  const mode = SpreadsheetApp.getActive().getSheetByName(SHEET_CONTACTS).getRange(CELL_MODE).getValue();
  SpreadsheetApp.getUi().alert(
    'Configuration actuelle',
    'Mode : ' + mode + '\n\nDestinataires :\n' + destinataires.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function reinitialiserHistorique() {
  const props = PropertiesService.getDocumentProperties();
  const all = props.getKeys();
  let count = 0;
  all.forEach(function(k) {
    if (k.indexOf(PROP_ROW_SENT) === 0 || k.indexOf(PROP_ROW_REMINDER) === 0 || k.indexOf(PROP_LAST_STATE) === 0) {
      props.deleteProperty(k);
      count++;
    }
  });
  SpreadsheetApp.getActive().toast(count + ' flags effaces', 'Reset', 5);
}

// ============================================================
// TESTS UNITAIRES CP_MATCH (optionnel, à lancer depuis l'éditeur)
// ============================================================

function testCPMatchUnit() {
  const tests = [
    // [cp, expression, attendu, description]
    ['33500', '', true, 'expression vide → toujours TRUE'],
    ['33500', 'Tous', true, 'expression "Tous" → TRUE'],
    ['33500', '33500', true, 'match exact'],
    ['33500', '33000', false, 'match exact différent'],
    ['33500', '33000, 33500, 34000', true, 'liste contenant'],
    ['33500', '33000, 34000', false, 'liste ne contenant pas'],
    ['33500', '33000-33999', true, 'tranche contenant'],
    ['33500', '34000-34999', false, 'tranche ne contenant pas'],
    ['33500', '33000-33500', true, 'tranche bornes incluses'],
    ['33500', '33000-33500, 35000', true, 'mix tranche + valeur'],
    ['33500', '33000\n34000\n33500', true, 'sauts de ligne'],
    ['33500', '  33500  ', true, 'avec espaces'],
    ['01200', '1000-2000', true, 'CP avec 0 devant'],
  ];

  const echecs = [];
  for (var i = 0; i < tests.length; i++) {
    var t = tests[i];
    var res = testCPMatch(normaliserCP(t[0]), t[1]);
    if (res !== t[2]) {
      echecs.push('FAIL: CP_MATCH("' + t[0] + '", "' + t[1] + '") = ' + res + ' (attendu ' + t[2] + ') — ' + t[3]);
    }
  }

  if (echecs.length === 0) {
    SpreadsheetApp.getUi().alert('Tests CP_MATCH', tests.length + ' tests passes avec succes', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('Tests CP_MATCH', echecs.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
