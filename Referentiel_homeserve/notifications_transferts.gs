/**
 * ============================================================
 * NOTIFICATIONS TRANSFERTS + CP_MATCH — HomeServe Énergies Services
 * ============================================================
 *
 * v5 — Améliorations de fond :
 *   - purgerProprietesObsoletes() : nettoie chaque jour les flags d'envoi
 *     orphelins (lignes supprimées, dates modifiées) → pas d'accumulation.
 *   - Listes de référence pilotées par plages nommées (voir installer_validations.gs).
 *
 * v4 — Corrections de revue :
 *   - Le handler de modification est renommé gererModificationActif (au lieu de
 *     onEdit) pour éviter la double exécution simple-trigger + installable.
 *     → APRÈS MISE À JOUR : relancer "Installer les déclencheurs" une fois.
 *   - Emails : échappement HTML des valeurs (noms, commentaires)
 *   - verifierRappelsQuotidiens : PropertiesService sorti de la boucle
 *
 * v3 :
 *   - Tracking des emails par clé contenu (filiale+remplace+nouveau+date)
 *     au lieu du numéro de ligne → résistant aux insertions/suppressions
 *   - Fonction onOpen : menu "⚙️ Transferts" visible dans le ruban
 *   - Fonction diagnostiquer() pour vérifier toute la configuration
 *
 * INSTALLATION
 * ============================================================
 * 1. Extensions → Apps Script → remplacer le code par ce fichier
 * 2. Sauvegarder (Ctrl+S)
 * 3. Lancer setupTriggers() une seule fois (menu ⚙️ Transferts → Installer déclencheurs)
 * 4. Renseigner les emails dans l'onglet CONTACTS
 * 5. Tester via menu ⚙️ Transferts → Tester envoi email
 * ============================================================
 */

// === Configuration feuilles / colonnes ===
const SHEET_TRANSFERTS  = 'TRANSFERTS';
const SHEET_CONTACTS    = 'CONTACTS';

const COL_FILIALE    = 1;
const COL_REMPLACE   = 2;
const COL_NOUVEAU    = 3;
const COL_DATE_DEBUT = 4;
const COL_DATE_FIN   = 5;
const COL_MOTIF      = 6;
const COL_ACTIF      = 7;
const COL_PRODUIT    = 8;
const COL_CP         = 9;
const COL_VILLE      = 10;
const COL_COMMENTAIRE= 11;

const CELL_MODE           = 'C6';
const CELL_EMAIL_TEST     = 'C9';
// Destinataires PROD : C12 jusqu'à C50 (plage lue dynamiquement, cellules vides ignorées)

// Préfixes pour PropertiesService — clés basées sur le contenu, pas le n° de ligne
const PROP_SENT     = 'sent_v_';
const PROP_REMINDER = 'sent_r_';

// ============================================================
// MENU PERSONNALISÉ
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Transferts')
    .addItem('🔧 Installer les déclencheurs', 'setupTriggers')
    .addItem('🏷️ Installer les plages nommées', 'installerPlagesNommees')
    .addItem('📋 Installer les listes déroulantes', 'installerToutesLesValidations')
    .addItem('📝 Ajouter les notes d\'aide', 'ajouterNotes')
    .addSeparator()
    .addItem('🗂️ Créer la feuille de saisie (secteurs)', 'creerFeuilleSaisie')
    .addItem('⚡ Générer les affectations depuis la saisie', 'genererAffectations')
    .addItem('🎨 Appliquer la charte HomeServe', 'appliquerCharte')
    .addSeparator()
    .addItem('📧 Tester l\'envoi d\'email', 'testEnvoiEmail')
    .addItem('👥 Voir les destinataires configurés', 'afficherDestinataires')
    .addSeparator()
    .addItem('🔍 Diagnostiquer la configuration', 'diagnostiquer')
    .addItem('🧹 Purger les flags emails orphelins', 'purgerProprietesObsoletesUI')
    .addItem('🔄 Réinitialiser l\'historique emails', 'reinitialiserHistorique')
    .addItem('🔒 Protéger les feuilles de référence', 'protegerFeuillesReference')
    .addToUi();
}

// ============================================================
// FONCTION CUSTOM : CP_MATCH(cp_cherche, expression)
// ============================================================

/**
 * Vérifie si un CP correspond à une expression CP (vide, valeur, liste, tranche).
 * @param {string|number} cpCherche  CP à tester (ex: "33500" ou 33500).
 * @param {string|range}  expression CP de référence dans TRANSFERTS.
 * @return TRUE/FALSE. Si expression est une plage, renvoie un tableau.
 * @customfunction
 */
function CP_MATCH(cpCherche, expression) {
  const cp = normaliserCP(cpCherche);
  if (Array.isArray(expression)) {
    return expression.map(function(row) {
      return Array.isArray(row)
        ? row.map(function(cell) { return testCPMatch(cp, cell); })
        : testCPMatch(cp, row);
    });
  }
  return testCPMatch(cp, expression);
}

function normaliserCP(x) {
  if (x === null || x === undefined || x === '') return '';
  let s = String(x).trim().replace(/\D/g, '');
  if (s === '') return '';
  while (s.length < 5) s = '0' + s;
  return s;
}

function testCPMatch(cp, expression) {
  if (cp === '') return false;
  if (expression === null || expression === undefined || expression === '') return true;

  const expr = String(expression).trim();
  if (expr === '' || expr.toLowerCase() === 'tous') return true;

  const morceaux = expr.split(/[,;\n\r]+/);
  for (let i = 0; i < morceaux.length; i++) {
    const m = morceaux[i].trim();
    if (!m) continue;
    if (m.indexOf('-') !== -1) {
      const parts = m.split('-');
      if (parts.length === 2) {
        const a = normaliserCP(parts[0]);
        const b = normaliserCP(parts[1]);
        if (a && b && parseInt(cp, 10) >= parseInt(a, 10) && parseInt(cp, 10) <= parseInt(b, 10)) return true;
      }
    } else {
      if (normaliserCP(m) === cp) return true;
    }
  }
  return false;
}

// ============================================================
// DÉCLENCHEUR DE MODIFICATION (installable)
// ============================================================
// IMPORTANT : ce handler ne doit PAS s'appeler onEdit, sinon Google le lance
// aussi comme "simple trigger" — qui n'a pas le droit d'envoyer des emails et
// génère une erreur à chaque édition. On lui donne donc un nom dédié et on le
// branche via setupTriggers() comme déclencheur installable.

function gererModificationActif(e) {
  if (!e || !e.range) return;
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_TRANSFERTS) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (col !== COL_ACTIF || row < 2) return;

    const newValue = e.value || '';
    const oldValue = e.oldValue || 'Non';

    if (newValue === 'Oui' && oldValue !== 'Oui') {
      envoyerEmailValidation(row);
    } else if (oldValue === 'Oui' && newValue !== 'Oui') {
      envoyerEmailAnnulation(row);
    }
  } catch (err) {
    Logger.log('Erreur gererModificationActif : ' + err.message);
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
  const props = PropertiesService.getDocumentProperties();

  for (let i = 0; i < data.length; i++) {
    const filiale  = data[i][COL_FILIALE - 1];
    const actif    = data[i][COL_ACTIF - 1];
    const dateDebut = data[i][COL_DATE_DEBUT - 1];

    if (!filiale || actif !== 'Oui' || !(dateDebut instanceof Date)) continue;

    const dateRappel = soustraireJoursOuvres(nouvelleDate(dateDebut), 3);
    if (!memeJour(aujourdHui, dateRappel)) continue;

    const infos = lireRangee(data[i]);
    const cle = cleTransfert(infos);

    if (props.getProperty(PROP_REMINDER + cle) !== 'true') {
      envoyerEmailRappelAvecInfos(infos);
      props.setProperty(PROP_REMINDER + cle, 'true');
    }
  }

  // Purge des flags orphelins (lignes supprimées, dates modifiées…)
  purgerProprietesObsoletes();
}

// ============================================================
// PURGE DES PROPRIÉTÉS ORPHELINES
// ============================================================
// Les flags d'envoi sont stockés par clé de contenu. Si une ligne est supprimée
// ou si sa date/commercial change, l'ancienne clé devient orpheline et resterait
// indéfiniment. Cette purge ne garde que les clés correspondant aux lignes
// actuellement présentes dans TRANSFERTS. Appelée chaque jour par le rappel,
// et disponible dans le menu.

function purgerProprietesObsoletes() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_TRANSFERTS);
  if (!sheet) return 0;

  const clesValides = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    data.forEach(row => {
      const infos = lireRangee(row);
      if (infos.filiale) clesValides[cleTransfert(infos)] = true;
    });
  }

  const props = PropertiesService.getDocumentProperties();
  let supprimes = 0;
  props.getKeys().forEach(k => {
    let cle = null;
    if (k.indexOf(PROP_SENT) === 0)          cle = k.slice(PROP_SENT.length);
    else if (k.indexOf(PROP_REMINDER) === 0) cle = k.slice(PROP_REMINDER.length);
    if (cle !== null && !clesValides[cle]) {
      props.deleteProperty(k);
      supprimes++;
    }
  });

  if (supprimes > 0) Logger.log('Purge : ' + supprimes + ' flag(s) orphelin(s) supprimé(s)');
  return supprimes;
}

/** Version appelable depuis le menu, avec retour visuel. */
function purgerProprietesObsoletesUI() {
  const supprimes = purgerProprietesObsoletes();
  SpreadsheetApp.getActive().toast(
    supprimes + ' flag(s) orphelin(s) supprimé(s)',
    '✓ Purge effectuée',
    5
  );
}

// ============================================================
// CLÉ DE TRACKING (basée sur le contenu, pas le n° de ligne)
// ============================================================

function cleTransfert(infos) {
  const dateStr = infos.dateDebut instanceof Date
    ? Utilities.formatDate(infos.dateDebut, Session.getScriptTimeZone(), 'yyyyMMdd')
    : String(infos.dateDebut || '');
  return (infos.filiale + '|' + infos.remplace + '|' + infos.nouveau + '|' + dateStr)
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9|_\-éàèêëîïôùûüç]/g, '')
    .slice(0, 250);
}

// ============================================================
// FONCTIONS D'ENVOI
// ============================================================

function envoyerEmailValidation(row) {
  const infos = lireLigne(row);
  if (!infos.filiale) return;

  const destinataires = obtenirDestinataires();
  if (!destinataires.length) {
    Logger.log('Aucun destinataire configuré — vérifiez l\'onglet CONTACTS');
    return;
  }

  const cle = cleTransfert(infos);
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(PROP_SENT + cle) === 'true') {
    Logger.log('Email validation déjà envoyé pour : ' + cle);
    return;
  }

  const sujet = '[Transfert validé] ' + infos.filiale + ' — ' + infos.remplace + ' → ' + infos.nouveau;
  MailApp.sendEmail({
    to: destinataires.join(','),
    subject: sujet,
    htmlBody: construireCorpsEmail(infos, 'Transfert validé',
      "Un transfert vient d'être validé dans le référentiel d'affectation commerciale."),
  });

  props.setProperty(PROP_SENT + cle, 'true');
  Logger.log('Email validation envoyé : ' + cle);
}

function envoyerEmailRappelAvecInfos(infos) {
  const destinataires = obtenirDestinataires();
  if (!destinataires.length) return;

  const sujet = '[Rappel J-3] Transfert ' + infos.filiale + ' commence le ' + formaterDate(infos.dateDebut);
  MailApp.sendEmail({
    to: destinataires.join(','),
    subject: sujet,
    htmlBody: construireCorpsEmail(infos, 'Rappel : transfert imminent',
      'Ce transfert prendra effet dans 3 jours ouvrés. Merci de vérifier que tout est en ordre.'),
  });
  Logger.log('Email rappel J-3 envoyé : ' + cleTransfert(infos));
}

function envoyerEmailAnnulation(row) {
  const infos = lireLigne(row);
  if (!infos.filiale) return;

  const destinataires = obtenirDestinataires();
  if (!destinataires.length) return;

  const sujet = '[Transfert annulé] ' + infos.filiale + ' — ' + infos.remplace + ' → ' + infos.nouveau;
  MailApp.sendEmail({
    to: destinataires.join(','),
    subject: sujet,
    htmlBody: construireCorpsEmail(infos, 'Transfert annulé',
      "Un transfert précédemment validé vient d'être désactivé."),
  });

  const cle = cleTransfert(infos);
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(PROP_SENT + cle);
  props.deleteProperty(PROP_REMINDER + cle);
  Logger.log('Email annulation envoyé : ' + cle);
}

// ============================================================
// UTILITAIRES
// ============================================================

function lireLigne(row) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_TRANSFERTS);
  const data = sheet.getRange(row, 1, 1, 11).getValues()[0];
  return lireRangee(data);
}

function lireRangee(data) {
  return {
    filiale    : data[COL_FILIALE - 1]     || '',
    remplace   : data[COL_REMPLACE - 1]    || '',
    nouveau    : data[COL_NOUVEAU - 1]     || '',
    dateDebut  : data[COL_DATE_DEBUT - 1],
    dateFin    : data[COL_DATE_FIN - 1],
    motif      : data[COL_MOTIF - 1]       || '',
    actif      : data[COL_ACTIF - 1]       || '',
    produit    : data[COL_PRODUIT - 1]     || 'Tous',
    cp         : data[COL_CP - 1]          || '',
    ville      : data[COL_VILLE - 1]       || '',
    commentaire: data[COL_COMMENTAIRE - 1] || '',
  };
}

function obtenirDestinataires() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CONTACTS);
  if (!sheet) return [];

  const mode = (sheet.getRange(CELL_MODE).getValue() || '').toString().toUpperCase();
  if (mode === 'TEST') {
    return validerEmails([sheet.getRange(CELL_EMAIL_TEST).getValue()]);
  }
  if (mode === 'PROD') {
    // Lit la colonne C à partir de la ligne 12 jusqu'à la dernière ligne non-vide.
    // Aucune limite : ajouter un destinataire = ajouter une ligne, sans toucher au code.
    // Les cellules vides intercalées sont simplement ignorées.
    const PREMIERE_LIGNE = 12;
    const derniere = sheet.getLastRow();
    if (derniere < PREMIERE_LIGNE) return [];
    const plage = sheet.getRange(PREMIERE_LIGNE, 3, derniere - PREMIERE_LIGNE + 1, 1).getValues();
    const emails = plage.map(r => r[0]).filter(v => v);
    return validerEmails(emails);
  }
  return [];
}

function validerEmails(liste) {
  return liste
    .map(e => (e || '').toString().trim())
    .filter(e => e && e.includes('@') && !e.includes('ton.email'));
}

function construireCorpsEmail(infos, titre, accroche) {
  const zone = infos.cp
    ? infos.cp + (infos.ville ? ' ' + infos.ville : '')
    : 'Toutes communes';

  return '<div style="font-family:Calibri,Arial,sans-serif;color:#303030;max-width:600px;">' +
    '<div style="background:#004990;color:white;padding:16px;">' +
      '<h2 style="margin:0;">' + titre + '</h2>' +
      '<p style="margin:4px 0 0;opacity:.85;">Référentiel d\'affectation — HomeServe Énergies Services</p>' +
    '</div>' +
    '<div style="background:#fff;border:1px solid #d8d8d8;border-top:none;padding:20px;">' +
      '<p>' + accroche + '</p>' +
      '<table style="border-collapse:collapse;width:100%;margin-top:12px;">' +
        ligneTable('Filiale',              infos.filiale) +
        ligneTable('Commercial remplacé',  infos.remplace) +
        ligneTable('Nouveau commercial',   infos.nouveau) +
        ligneTable('Date début',           formaterDate(infos.dateDebut)) +
        ligneTable('Date fin',             infos.dateFin ? formaterDate(infos.dateFin) : '(non précisée)') +
        ligneTable('Motif',                infos.motif) +
        ligneTable('Produit',              infos.produit) +
        ligneTable('Zone (CP)',            zone) +
        (infos.commentaire ? ligneTable('Commentaire', infos.commentaire) : '') +
      '</table>' +
      '<p style="margin-top:20px;font-size:12px;color:#808080;">Email automatique — ne pas répondre.</p>' +
    '</div>' +
  '</div>';
}

function ligneTable(lbl, val) {
  return '<tr><td style="padding:6px;background:#f4f7fb;font-weight:600;">' + echapperHtml(lbl) + '</td>' +
         '<td style="padding:6px;">' + echapperHtml(val) + '</td></tr>';
}

// Échappe les caractères HTML pour éviter qu'un nom ou commentaire (ex: "A < B")
// ne casse la mise en page de l'email.
function echapperHtml(val) {
  return String(val == null ? '' : val)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formaterDate(d) {
  if (!(d instanceof Date)) return '(invalide)';
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
}

function nouvelleDate(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function memeJour(d1, d2) {
  return d1.getTime() === d2.getTime();
}

function soustraireJoursOuvres(date, n) {
  const resultat = new Date(date);
  let restants = n;
  while (restants > 0) {
    resultat.setDate(resultat.getDate() - 1);
    const j = resultat.getDay();
    if (j !== 0 && j !== 6) restants--;
  }
  return resultat;
}

// ============================================================
// INSTALLATION ET TESTS
// ============================================================

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('gererModificationActif')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  ScriptApp.newTrigger('verifierRappelsQuotidiens')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  SpreadsheetApp.getActive().toast(
    'Déclencheurs installés. Notifications actives.',
    '✓ Configuration OK',
    5
  );
}

function testEnvoiEmail() {
  const destinataires = obtenirDestinataires();
  if (!destinataires.length) {
    SpreadsheetApp.getUi().alert(
      'Aucun destinataire configuré',
      'Vérifiez l\'onglet CONTACTS :\n' +
      '• C6 doit contenir TEST ou PROD\n' +
      '• C9 doit contenir votre vraie adresse email\n' +
      '• (Pas de "ton.email" dans l\'adresse)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const mode = ss.getSheetByName(SHEET_CONTACTS).getRange(CELL_MODE).getValue();
  MailApp.sendEmail({
    to: destinataires.join(','),
    subject: '[TEST] Notifications transferts — HomeServe',
    htmlBody: '<div style="font-family:Calibri,Arial,sans-serif;">' +
      '<h2 style="color:#004990;">✓ Test de configuration</h2>' +
      '<p>Cet email confirme que les notifications fonctionnent correctement.</p>' +
      '<p><strong>Mode :</strong> ' + mode + '</p>' +
      '<p><strong>Destinataires :</strong> ' + destinataires.join(', ') + '</p>' +
      '</div>',
  });
  ss.toast('Email de test envoyé à : ' + destinataires.join(', '), '✓ Test OK', 8);
}

function afficherDestinataires() {
  const destinataires = obtenirDestinataires();
  const mode = SpreadsheetApp.getActive().getSheetByName(SHEET_CONTACTS).getRange(CELL_MODE).getValue();
  SpreadsheetApp.getUi().alert(
    'Configuration actuelle',
    'Mode : ' + mode + '\n\nDestinataires :\n' +
    (destinataires.length ? destinataires.join('\n') : '(aucun — vérifiez l\'onglet CONTACTS)'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function reinitialiserHistorique() {
  const ui = SpreadsheetApp.getUi();
  const reponse = ui.alert(
    'Confirmer la réinitialisation',
    'Cela efface l\'historique des emails envoyés.\n' +
    'Les prochaines activations de transferts renverront des emails.\n\nContinuer ?',
    ui.ButtonSet.YES_NO
  );
  if (reponse !== ui.Button.YES) return;

  const props = PropertiesService.getDocumentProperties();
  let count = 0;
  props.getKeys().forEach(k => {
    if (k.startsWith(PROP_SENT) || k.startsWith(PROP_REMINDER)) {
      props.deleteProperty(k);
      count++;
    }
  });
  SpreadsheetApp.getActive().toast(count + ' flag(s) effacé(s)', '✓ Reset effectué', 5);
}

// ============================================================
// DIAGNOSTIC
// ============================================================

function diagnostiquer() {
  const ss = SpreadsheetApp.getActive();
  const lignes = [];
  let ok = true;

  // Feuilles requises
  const feuillesRequises = [
    'TRANSFERTS', 'CONTACTS', 'PARAMETRES', 'RECHERCHE',
    'AFFECTATIONS_COMMUNES', 'PRODUITS_COMMERCIAUX', 'EXCEPTIONS_PRODUITS',
    '_DDL_TRANSFERTS', '_COMMERCIAUX_PAR_FILIALE_H'
  ];
  feuillesRequises.forEach(nom => {
    if (ss.getSheetByName(nom)) {
      lignes.push('✓ Feuille ' + nom);
    } else {
      lignes.push('✗ Feuille manquante : ' + nom);
      ok = false;
    }
  });

  lignes.push('');

  // Contacts
  const contacts = ss.getSheetByName(SHEET_CONTACTS);
  if (contacts) {
    const mode = contacts.getRange(CELL_MODE).getValue();
    lignes.push('Mode email : ' + (mode || '(vide — définir TEST ou PROD en C6)'));
    const destinataires = obtenirDestinataires();
    if (destinataires.length) {
      lignes.push('✓ Destinataires : ' + destinataires.join(', '));
    } else {
      lignes.push('✗ Aucun destinataire valide — vérifiez l\'onglet CONTACTS');
      ok = false;
    }
  }

  lignes.push('');

  // Déclencheurs
  const triggers = ScriptApp.getProjectTriggers();
  const hasOnEdit = triggers.some(t => t.getHandlerFunction() === 'gererModificationActif');
  const hasQuotidien = triggers.some(t => t.getHandlerFunction() === 'verifierRappelsQuotidiens');
  lignes.push(hasOnEdit    ? '✓ Déclencheur de modification installé' : '✗ Déclencheur de modification manquant → lancer setupTriggers()');
  lignes.push(hasQuotidien ? '✓ Déclencheur quotidien installé'    : '✗ Déclencheur quotidien manquant → lancer setupTriggers()');
  if (!hasOnEdit || !hasQuotidien) ok = false;

  lignes.push('');

  // Transferts actifs
  const shTransferts = ss.getSheetByName(SHEET_TRANSFERTS);
  if (shTransferts) {
    const lastRow = shTransferts.getLastRow();
    if (lastRow > 1) {
      const data = shTransferts.getRange(2, COL_ACTIF, lastRow - 1, 1).getValues();
      const actifs = data.filter(r => r[0] === 'Oui').length;
      lignes.push('Transferts actifs : ' + actifs + ' / ' + (lastRow - 1));
    } else {
      lignes.push('Aucun transfert enregistré');
    }
  }

  const titre = ok ? '✓ Configuration OK' : '⚠ Configuration incomplète';
  SpreadsheetApp.getUi().alert(titre, lignes.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// PROTECTION DES FEUILLES DE RÉFÉRENCE
// ============================================================

function protegerFeuillesReference() {
  const ss = SpreadsheetApp.getActive();
  const feuillesAProteger = ['DPT_SOURCE', '_COMMERCIAUX_PAR_FILIALE_H', '_DDL_TRANSFERTS'];
  const moi = Session.getEffectiveUser();
  let count = 0;

  feuillesAProteger.forEach(nom => {
    const sh = ss.getSheetByName(nom);
    if (!sh) return;
    // Supprimer protections existantes
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const protection = sh.protect().setDescription('Feuille de référence — ne pas modifier manuellement');
    protection.addEditor(moi);
    protection.removeEditors(
      protection.getEditors().filter(e => e.getEmail() !== moi.getEmail())
    );
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
    count++;
  });

  // PARAMETRES : source de toutes les listes déroulantes. On la protège en
  // AVERTISSEMENT (et non en blocage), pour ne pas verrouiller l'administrateur
  // qui doit encore l'ajuster, tout en évitant les modifications accidentelles.
  const params = ss.getSheetByName('PARAMETRES');
  if (params) {
    params.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    params.protect()
      .setDescription('PARAMETRES — source des listes déroulantes. Modifier avec précaution.')
      .setWarningOnly(true);
    count++;
  }

  SpreadsheetApp.getActive().toast(
    count + ' feuille(s) protégée(s) : ' + feuillesAProteger.filter(n => ss.getSheetByName(n)).concat(params ? ['PARAMETRES (avertissement)'] : []).join(', '),
    '✓ Protection appliquée',
    6
  );
}

// ============================================================
// TESTS UNITAIRES CP_MATCH
// ============================================================

function testCPMatchUnit() {
  const tests = [
    ['33500', '',                true,  'expression vide → TRUE'],
    ['33500', 'Tous',            true,  '"Tous" → TRUE'],
    ['33500', '33500',           true,  'match exact'],
    ['33500', '33000',           false, 'valeur différente'],
    ['33500', '33000, 33500',    true,  'liste contenant'],
    ['33500', '33000, 34000',    false, 'liste ne contenant pas'],
    ['33500', '33000-33999',     true,  'tranche contenant'],
    ['33500', '34000-34999',     false, 'tranche ne contenant pas'],
    ['33500', '33000-33500',     true,  'borne incluse'],
    ['33500', '33000-33500, 35000', true, 'mix tranche + valeur'],
    ['33500', '33000\n34000\n33500', true, 'sauts de ligne'],
    ['01200', '1000-2000',       true,  'CP avec 0 devant'],
  ];

  const echecs = [];
  tests.forEach(t => {
    const res = testCPMatch(normaliserCP(t[0]), t[1]);
    if (res !== t[2]) echecs.push(`FAIL: CP_MATCH("${t[0]}", "${t[1]}") = ${res} (attendu ${t[2]}) — ${t[3]}`);
  });

  SpreadsheetApp.getUi().alert(
    'Tests CP_MATCH',
    echecs.length === 0
      ? tests.length + ' tests passés ✓'
      : echecs.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
