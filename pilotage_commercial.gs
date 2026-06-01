/**
 * PILOTAGE COMMERCIAL — Suivi hebdo des dossiers en signature
 * ------------------------------------------------------------
 * Onglets VISIBLES :
 *   - PILOTAGE  : 1 ligne / dossier, vue de travail unique et ÉDITABLE
 *   - DASHBOARD : synthèse + graphiques par commercial
 * Onglets MASQUÉS (moteur) :
 *   - IMPORT_CRM     : on colle le CSV ici
 *   - BASE_DOSSIERS  : stockage durable (CRM + technique + annotations manager)
 *   - HISTO_IMPORTS  : archive de tous les imports
 *
 * Flux hebdo :
 *   1. Coller le CSV dans IMPORT_CRM
 *   2. Menu "Pilotage commercial" > "Mettre à jour depuis IMPORT_CRM"
 *      -> conserve tes annotations, rafraîchit le CRM, recalcule alertes/mouvement,
 *         reconstruit PILOTAGE et DASHBOARD.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Pilotage commercial")
    .addItem("1. Initialiser le fichier", "initialiserPilotage")
    .addItem("2. Mettre à jour depuis IMPORT_CRM", "majDepuisImportCRM")
    .addItem("3. Rafraîchir PILOTAGE + DASHBOARD", "rafraichirTout")
    .addSeparator()
    .addItem("Trier PILOTAGE", "trierPilotage")
    .addToUi();
}

const CONFIG = {
  SHEET_IMPORT: "IMPORT_CRM",
  SHEET_BASE: "BASE_DOSSIERS",
  SHEET_PILOTAGE: "PILOTAGE",
  SHEET_DASHBOARD: "DASHBOARD",
  SHEET_HISTO: "HISTO_IMPORTS",

  // Colonnes telles qu'exportées du CRM (séparateur ; dans le CSV)
  CRM_HEADERS: [
    "charge_dossier", "step", "consentement", "tiers", "dossier",
    "commercial", "regie", "responsable", "typetravaux", "aidesutilisees",
    "daterebond", "datevisite", "datecontroledevis", "datetransodmcommercial",
    "datesigneprevuecial", "statutexportmpr", "datenotifmpr", "daterefusmpr",
    "datepos", "commpassagecial", "commretourcial", "commentairecial",
    "datesignfi", "datectrlcial"
  ],

  // Colonnes techniques internes (suivi du cycle de vie + mouvement)
  BASE_TECH_HEADERS: [
    "date_premier_import", "date_dernier_import", "nb_imports_vu",
    "present_dernier_import", "date_derniere_presence",
    "disparu_du_crm", "date_disparition_crm",
    "step_precedent", "nb_imports_meme_step", "date_changement_step", "mouvement"
  ],

  // Champs saisis par le manager (simplifiés). Stockés dans BASE, édités dans PILOTAGE.
  MANAGER_HEADERS: [
    "statut_manager", "niveau_chaleur", "blocage_principal",
    "prochaine_action", "date_cible_relance", "montant_devis",
    "probabilite_signature", "commentaire_manager", "date_maj_1to1"
  ],

  // Vue PILOTAGE (ordre des colonnes affichées)
  PILOTAGE_HEADERS: [
    "alerte",            // auto
    "commercial",        // crm
    "tiers",             // crm
    "dossier",           // crm
    "typetravaux",       // crm
    "step",              // crm (étape CRM)
    "mouvement",         // auto
    "age_dossier_j",     // auto
    "statut_manager",    // édit
    "niveau_chaleur",    // édit
    "blocage_principal", // édit
    "prochaine_action",  // édit
    "date_cible_relance",// édit
    "montant_devis",     // édit
    "probabilite_signature", // édit
    "prevision_ponderee",// auto
    "commentaire_manager",// édit
    "date_maj_1to1"      // édit
  ],

  // Listes déroulantes — c'est ici que tu ajustes les valeurs
  VALIDATIONS: {
    "statut_manager": [
      "", "À relancer", "En attente client", "En attente documents",
      "En attente banque", "Injoignable", "Signature proche", "Signé", "Perdu"
    ],
    "niveau_chaleur": ["Froid", "Tiède", "Chaud"],
    "blocage_principal": [
      "Aucun", "Prix / RAC", "Aides", "Financement", "Doute client",
      "Attente conjoint", "Injoignable", "Concurrent", "Technique", "Timing"
    ],
    "prochaine_action": [
      "", "Appeler le client", "Relancer devis", "Relancer documents",
      "Revoir l'offre", "Arbitrage manager", "Attendre retour",
      "Clôture (perdu)", "Signature attendue"
    ],
    "probabilite_signature": ["", "10%", "25%", "50%", "75%", "90%"]
  },

  // Statuts considérés comme "dossier clos" (sortent du pilotage actif)
  STATUTS_CLOS: ["Signé", "Perdu"],

  // Seuil de stagnation : nb d'imports au même stade avant alerte
  SEUIL_STAGNATION: 2,
  // Seuil d'ancienneté (jours) pour "vieux dossier"
  SEUIL_VIEUX_J: 30
};

/* ===================== ACTIONS MENU ===================== */

function initialiserPilotage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shImport = getOrCreateSheet_(ss, CONFIG.SHEET_IMPORT);
  const shBase = getOrCreateSheet_(ss, CONFIG.SHEET_BASE);
  const shPilote = getOrCreateSheet_(ss, CONFIG.SHEET_PILOTAGE);
  const shDash = getOrCreateSheet_(ss, CONFIG.SHEET_DASHBOARD);
  const shHisto = getOrCreateSheet_(ss, CONFIG.SHEET_HISTO);

  writeHeadersIfNeeded_(shImport, CONFIG.CRM_HEADERS);

  const baseHeaders = CONFIG.CRM_HEADERS
    .concat(CONFIG.BASE_TECH_HEADERS)
    .concat(CONFIG.MANAGER_HEADERS);
  writeHeadersIfNeeded_(shBase, baseHeaders);

  writeHeadersIfNeeded_(shHisto, ["date_import"].concat(CONFIG.CRM_HEADERS));
  writeHeadersIfNeeded_(shPilote, CONFIG.PILOTAGE_HEADERS);

  // Onglets moteur masqués, onglets de pilotage visibles
  hideSheet_(shImport);
  hideSheet_(shBase);
  hideSheet_(shHisto);
  shPilote.showSheet();
  shDash.showSheet();

  applyHeaderFormatting_(shImport);
  applyHeaderFormatting_(shBase);
  applyHeaderFormatting_(shHisto);
  applyHeaderFormatting_(shPilote);
}

function majDepuisImportCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initialiserPilotage();

  const shImport = ss.getSheetByName(CONFIG.SHEET_IMPORT);
  const shBase = ss.getSheetByName(CONFIG.SHEET_BASE);
  const shPilote = ss.getSheetByName(CONFIG.SHEET_PILOTAGE);
  const shHisto = ss.getSheetByName(CONFIG.SHEET_HISTO);

  const importData = shImport.getDataRange().getValues();
  if (importData.length < 2) {
    throw new Error("IMPORT_CRM ne contient aucune donnée à traiter.");
  }

  const importHeaders = normalizeHeaders_(importData[0]);
  const idxDossierImport = importHeaders.indexOf("dossier");
  if (idxDossierImport === -1) {
    throw new Error("La colonne 'dossier' est introuvable dans IMPORT_CRM.");
  }

  const importRows = importData.slice(1)
    .map(row => fitRowToHeaders_(row, importHeaders.length))
    .filter(row => String(row[idxDossierImport] || "").trim() !== "");

  if (importRows.length === 0) {
    throw new Error("Aucun dossier exploitable trouvé dans IMPORT_CRM.");
  }

  appendHisto_(shHisto, importHeaders, importRows);

  // 1. Récupérer les annotations manager fraîches saisies dans PILOTAGE
  const managerEdits = harvestManagerFromPilotage_(shPilote);

  // 2. Charger BASE
  const baseHeaders = normalizeHeaders_(
    shBase.getRange(1, 1, 1, shBase.getLastColumn()).getValues()[0]
  );
  const baseData = shBase.getLastRow() > 1
    ? shBase.getRange(2, 1, shBase.getLastRow() - 1, shBase.getLastColumn()).getValues()
    : [];

  const idxDossierBase = baseHeaders.indexOf("dossier");
  const baseMap = {};
  baseData.forEach(row => {
    const dossier = String(row[idxDossierBase] || "").trim();
    if (dossier) baseMap[dossier] = row.slice();
  });

  const presentSet = new Set();
  const today = new Date();

  // 3. Upsert depuis l'import + calcul du mouvement
  importRows.forEach(importRow => {
    const dossier = String(importRow[idxDossierImport] || "").trim();
    if (!dossier) return;
    presentSet.add(dossier);

    const isNew = !baseMap[dossier];
    const row = isNew ? new Array(baseHeaders.length).fill("") : baseMap[dossier].slice();

    const oldStep = isNew ? "" : String(getCell_(row, baseHeaders, "step") || "").trim();
    const newStep = String(importRow[importHeaders.indexOf("step")] || "").trim();

    // Copier les colonnes CRM
    CONFIG.CRM_HEADERS.forEach(col => {
      const idxI = importHeaders.indexOf(col);
      const idxB = baseHeaders.indexOf(col);
      if (idxI !== -1 && idxB !== -1) row[idxB] = importRow[idxI];
    });

    // Technique : présence
    setCell_(row, baseHeaders, "date_dernier_import", today);
    setCell_(row, baseHeaders, "date_derniere_presence", today);
    setCell_(row, baseHeaders, "present_dernier_import", "OUI");
    setCell_(row, baseHeaders, "disparu_du_crm", "NON");
    setCell_(row, baseHeaders, "date_disparition_crm", "");

    // Mouvement
    if (isNew) {
      setCell_(row, baseHeaders, "date_premier_import", today);
      setCell_(row, baseHeaders, "nb_imports_vu", 1);
      setCell_(row, baseHeaders, "step_precedent", "");
      setCell_(row, baseHeaders, "nb_imports_meme_step", 1);
      setCell_(row, baseHeaders, "date_changement_step", today);
      setCell_(row, baseHeaders, "mouvement", "🆕 Nouveau");
    } else {
      const nbImports = Number(getCell_(row, baseHeaders, "nb_imports_vu")) || 0;
      setCell_(row, baseHeaders, "nb_imports_vu", nbImports + 1);

      if (newStep && newStep !== oldStep) {
        setCell_(row, baseHeaders, "step_precedent", oldStep);
        setCell_(row, baseHeaders, "nb_imports_meme_step", 1);
        setCell_(row, baseHeaders, "date_changement_step", today);
        setCell_(row, baseHeaders, "mouvement", "✅ Avance");
      } else {
        const nbSame = (Number(getCell_(row, baseHeaders, "nb_imports_meme_step")) || 0) + 1;
        setCell_(row, baseHeaders, "nb_imports_meme_step", nbSame);
        setCell_(row, baseHeaders, "mouvement",
          nbSame >= CONFIG.SEUIL_STAGNATION ? "⚠️ Stagne" : "→ Stable");
      }
    }

    baseMap[dossier] = row;
  });

  // 4. Dossiers disparus du CRM
  Object.keys(baseMap).forEach(dossier => {
    if (!presentSet.has(dossier)) {
      const row = baseMap[dossier].slice();
      setCell_(row, baseHeaders, "present_dernier_import", "NON");
      setCell_(row, baseHeaders, "disparu_du_crm", "OUI");
      setCell_(row, baseHeaders, "mouvement", "❓ Hors CRM");
      if (!getCell_(row, baseHeaders, "date_disparition_crm")) {
        setCell_(row, baseHeaders, "date_disparition_crm", today);
      }
      baseMap[dossier] = row;
    }
  });

  // 5. Réappliquer les annotations manager fraîches (PILOTAGE prime sur l'ancien BASE)
  Object.keys(managerEdits).forEach(dossier => {
    if (!baseMap[dossier]) return;
    const edits = managerEdits[dossier];
    CONFIG.MANAGER_HEADERS.forEach(col => {
      if (Object.prototype.hasOwnProperty.call(edits, col)) {
        setCell_(baseMap[dossier], baseHeaders, col, edits[col]);
      }
    });
  });

  // 6. Réécrire BASE (trié par dossier)
  const finalBaseRows = Object.keys(baseMap).sort().map(d => baseMap[d]);
  rewriteSheet_(shBase, baseHeaders, finalBaseRows);
  applyHeaderFormatting_(shBase);

  // 7. Reconstruire PILOTAGE + DASHBOARD
  rebuildPilotage_(shBase, shPilote);
  rebuildDashboard_(shBase, ss.getSheetByName(CONFIG.SHEET_DASHBOARD));

  SpreadsheetApp.getUi().alert(
    "Mise à jour terminée.\n\n" +
    importRows.length + " dossiers traités depuis le CRM."
  );
}

function rafraichirTout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initialiserPilotage();

  const shBase = ss.getSheetByName(CONFIG.SHEET_BASE);
  const shPilote = ss.getSheetByName(CONFIG.SHEET_PILOTAGE);
  const shDash = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);

  // Sauver d'abord les éventuelles annotations en cours dans PILOTAGE
  const managerEdits = harvestManagerFromPilotage_(shPilote);
  if (Object.keys(managerEdits).length > 0) {
    const baseHeaders = normalizeHeaders_(
      shBase.getRange(1, 1, 1, shBase.getLastColumn()).getValues()[0]
    );
    const baseData = shBase.getLastRow() > 1
      ? shBase.getRange(2, 1, shBase.getLastRow() - 1, shBase.getLastColumn()).getValues()
      : [];
    const idxDossier = baseHeaders.indexOf("dossier");
    baseData.forEach(row => {
      const dossier = String(row[idxDossier] || "").trim();
      if (managerEdits[dossier]) {
        CONFIG.MANAGER_HEADERS.forEach(col => {
          if (Object.prototype.hasOwnProperty.call(managerEdits[dossier], col)) {
            setCell_(row, baseHeaders, col, managerEdits[dossier][col]);
          }
        });
      }
    });
    rewriteSheet_(shBase, baseHeaders, baseData);
  }

  rebuildPilotage_(shBase, shPilote);
  rebuildDashboard_(shBase, shDash);

  SpreadsheetApp.getUi().alert("PILOTAGE et DASHBOARD rafraîchis.");
}

function trierPilotage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPilote = ss.getSheetByName(CONFIG.SHEET_PILOTAGE);
  if (!shPilote) throw new Error("L'onglet PILOTAGE est introuvable.");
  trierPilotage_(shPilote);
  SpreadsheetApp.getUi().alert("PILOTAGE trié.");
}

/* ===================== CONSTRUCTION PILOTAGE ===================== */

function rebuildPilotage_(shBase, shPilote) {
  const baseHeaders = normalizeHeaders_(
    shBase.getRange(1, 1, 1, shBase.getLastColumn()).getValues()[0]
  );
  const baseRows = shBase.getLastRow() > 1
    ? shBase.getRange(2, 1, shBase.getLastRow() - 1, shBase.getLastColumn()).getValues()
    : [];

  const idxDossier = baseHeaders.indexOf("dossier");
  const output = [];

  baseRows.forEach(baseRow => {
    const dossier = String(baseRow[idxDossier] || "").trim();
    if (!dossier) return;

    const statut = String(getCell_(baseRow, baseHeaders, "statut_manager") || "").trim();
    const disparu = String(getCell_(baseRow, baseHeaders, "disparu_du_crm") || "").trim() === "OUI";

    // On sort du pilotage actif les dossiers clos ET disparus du CRM
    if (CONFIG.STATUTS_CLOS.indexOf(statut) !== -1 && disparu) return;

    const datePivot = pickDatePivot_(baseRow, baseHeaders);
    const age = calcDaysFrom_(datePivot);
    const mouvement = String(getCell_(baseRow, baseHeaders, "mouvement") || "");
    const montant = parseMontant_(getCell_(baseRow, baseHeaders, "montant_devis"));
    const proba = probaToNum_(getCell_(baseRow, baseHeaders, "probabilite_signature"));
    const prevision = (montant !== "" && proba !== "") ? Math.round(montant * proba) : "";

    const alerte = computeAlerte_(baseRow, baseHeaders, { age: age, mouvement: mouvement });

    const rowOut = CONFIG.PILOTAGE_HEADERS.map(header => {
      switch (header) {
        case "alerte": return alerte;
        case "mouvement": return mouvement;
        case "age_dossier_j": return age;
        case "prevision_ponderee": return prevision;
        case "montant_devis": return montant === "" ? "" : montant;
      }
      const idxB = baseHeaders.indexOf(header);
      return idxB !== -1 ? baseRow[idxB] : "";
    });

    output.push(rowOut);
  });

  rewriteSheet_(shPilote, CONFIG.PILOTAGE_HEADERS, output);
  setupPilotageValidation_(shPilote);
  applyPilotageFormatting_(shPilote);
  trierPilotage_(shPilote);
}

/**
 * Moteur d'alertes : renvoie UN badge selon la priorité.
 */
function computeAlerte_(baseRow, headers, calc) {
  const statut = String(getCell_(baseRow, headers, "statut_manager") || "").trim();
  if (CONFIG.STATUTS_CLOS.indexOf(statut) !== -1) return "✅";

  const disparu = String(getCell_(baseRow, headers, "disparu_du_crm") || "").trim() === "OUI";
  if (disparu) return "❓ À statuer";

  const prochaine = String(getCell_(baseRow, headers, "prochaine_action") || "").trim();
  if (!prochaine) return "⛔ Pas d'action";

  const dateRelance = getCell_(baseRow, headers, "date_cible_relance");
  const joursRelance = calcDaysFrom_(dateRelance);
  if (joursRelance !== "" && joursRelance > 0) return "🔴 Relance en retard";

  const chaleur = String(getCell_(baseRow, headers, "niveau_chaleur") || "").trim();
  const stagne = calc.mouvement === "⚠️ Stagne";
  if (chaleur === "Chaud" && stagne) return "🟠 Chaud bloqué";

  if (calc.age !== "" && calc.age > CONFIG.SEUIL_VIEUX_J && stagne) return "⚫ Vieux + stagne";

  return "✅";
}

function trierPilotage_(shPilote) {
  const lastRow = shPilote.getLastRow();
  const lastCol = shPilote.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return;

  const headers = normalizeHeaders_(shPilote.getRange(1, 1, 1, lastCol).getValues()[0]);
  const data = shPilote.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const idxCom = headers.indexOf("commercial");
  const idxAlerte = headers.indexOf("alerte");
  const idxAge = headers.indexOf("age_dossier_j");

  // Rang d'alerte : les vraies alertes remontent, ✅ en bas
  const alerteRank = a => (String(a || "").indexOf("✅") !== -1 ? 9 : 1);

  data.sort((a, b) => {
    const cA = String(a[idxCom] || "").toLowerCase().trim();
    const cB = String(b[idxCom] || "").toLowerCase().trim();
    if (cA < cB) return -1;
    if (cA > cB) return 1;

    const rA = alerteRank(a[idxAlerte]);
    const rB = alerteRank(b[idxAlerte]);
    if (rA !== rB) return rA - rB;

    const ageA = (idxAge !== -1 && a[idxAge] !== "") ? Number(a[idxAge]) : -1;
    const ageB = (idxAge !== -1 && b[idxAge] !== "") ? Number(b[idxAge]) : -1;
    return ageB - ageA; // du plus vieux au plus récent
  });

  shPilote.getRange(2, 1, data.length, lastCol).setValues(data);
}

/* ===================== DASHBOARD ===================== */

function rebuildDashboard_(shBase, shDash) {
  const baseHeaders = normalizeHeaders_(
    shBase.getRange(1, 1, 1, shBase.getLastColumn()).getValues()[0]
  );
  const baseRows = shBase.getLastRow() > 1
    ? shBase.getRange(2, 1, shBase.getLastRow() - 1, shBase.getLastColumn()).getValues()
    : [];

  // Agrégation par commercial (uniquement dossiers actifs)
  const stats = {}; // commercial -> {...}
  let totFroid = 0, totTiede = 0, totChaud = 0;

  baseRows.forEach(row => {
    const statut = String(getCell_(row, baseHeaders, "statut_manager") || "").trim();
    const disparu = String(getCell_(row, baseHeaders, "disparu_du_crm") || "").trim() === "OUI";
    const clos = CONFIG.STATUTS_CLOS.indexOf(statut) !== -1;
    if (clos && disparu) return; // dossier sorti du pilotage

    const com = String(getCell_(row, baseHeaders, "commercial") || "—").trim() || "—";
    if (!stats[com]) {
      stats[com] = {
        actifs: 0, chaud: 0, tiede: 0, froid: 0,
        alerte: 0, prevision: 0, signes: 0, sansAction: 0
      };
    }
    const s = stats[com];

    if (statut === "Signé") { s.signes++; return; }

    s.actifs++;

    const chaleur = String(getCell_(row, baseHeaders, "niveau_chaleur") || "").trim();
    if (chaleur === "Chaud") { s.chaud++; totChaud++; }
    else if (chaleur === "Tiède") { s.tiede++; totTiede++; }
    else if (chaleur === "Froid") { s.froid++; totFroid++; }

    const prochaine = String(getCell_(row, baseHeaders, "prochaine_action") || "").trim();
    if (!prochaine) s.sansAction++;

    const age = calcDaysFrom_(pickDatePivot_(row, baseHeaders));
    const mouvement = String(getCell_(row, baseHeaders, "mouvement") || "");
    const alerte = computeAlerte_(row, baseHeaders, { age: age, mouvement: mouvement });
    if (alerte.indexOf("✅") === -1) s.alerte++;

    const montant = parseMontant_(getCell_(row, baseHeaders, "montant_devis"));
    const proba = probaToNum_(getCell_(row, baseHeaders, "probabilite_signature"));
    if (montant !== "" && proba !== "") s.prevision += montant * proba;
  });

  // Construire le tableau de synthèse
  const headers = [
    "commercial", "dossiers_actifs", "chauds", "tièdes", "froids",
    "en_alerte", "sans_action", "prévision_€", "signés"
  ];
  const coms = Object.keys(stats).sort();
  const rows = coms.map(com => {
    const s = stats[com];
    return [
      com, s.actifs, s.chaud, s.tiede, s.froid,
      s.alerte, s.sansAction, Math.round(s.prevision), s.signes
    ];
  });

  // Ligne TOTAL
  if (rows.length > 0) {
    const tot = rows.reduce((acc, r) => {
      for (let i = 1; i < r.length; i++) acc[i] += Number(r[i]) || 0;
      return acc;
    }, ["TOTAL", 0, 0, 0, 0, 0, 0, 0, 0]);
    rows.push(tot);
  }

  // Réécriture
  shDash.clear();
  removeChartsFromSheet_(shDash);

  shDash.getRange(1, 1).setValue("TABLEAU DE BORD — Pilotage commercial")
    .setFontSize(14).setFontWeight("bold");
  shDash.getRange(2, 1).setValue("Dernière mise à jour : " + formatDateFR_(new Date()))
    .setFontStyle("italic");

  const startRow = 4;
  shDash.getRange(startRow, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#d9ead3");
  if (rows.length > 0) {
    shDash.getRange(startRow + 1, 1, rows.length, headers.length).setValues(rows);
    // Total en gras
    shDash.getRange(startRow + rows.length, 1, 1, headers.length)
      .setFontWeight("bold").setBackground("#fff2cc");
  }
  shDash.setFrozenRows(startRow);
  try { shDash.autoResizeColumns(1, headers.length); } catch (e) {}

  // Graphiques
  if (rows.length > 1) {
    const dataRows = rows.length - 1; // hors TOTAL
    const tableRange = shDash.getRange(startRow, 1, dataRows + 1, headers.length);

    // 1. Histogramme prévision € par commercial (col 1 + col 8)
    const previsionRange = shDash.getRange(startRow, 8, dataRows + 1, 1);
    const comRange = shDash.getRange(startRow, 1, dataRows + 1, 1);
    const chartPrevision = shDash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(comRange)
      .addRange(previsionRange)
      .setNumHeaders(1)
      .setOption("title", "Prévision € pondérée par commercial")
      .setOption("legend", { position: "none" })
      .setPosition(startRow + rows.length + 3, 1, 0, 0)
      .build();
    shDash.insertChart(chartPrevision);

    // 2. Camembert répartition chaleur (global)
    const heatHeaderRow = startRow + rows.length + 22;
    shDash.getRange(heatHeaderRow, 1, 3, 2).setValues([
      ["Chaud", totChaud],
      ["Tiède", totTiede],
      ["Froid", totFroid]
    ]);
    const heatRange = shDash.getRange(heatHeaderRow, 1, 3, 2);
    const chartHeat = shDash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(heatRange)
      .setOption("title", "Répartition de la chaleur (équipe)")
      .setOption("colors", ["#e06666", "#f6b26b", "#6fa8dc"])
      .setPosition(startRow + rows.length + 3, 6, 0, 0)
      .build();
    shDash.insertChart(chartHeat);
  }

  shDash.showSheet();
}

/* ===================== HELPERS ===================== */

function harvestManagerFromPilotage_(shPilote) {
  const result = {};
  if (!shPilote || shPilote.getLastRow() <= 1) return result;

  const headers = normalizeHeaders_(
    shPilote.getRange(1, 1, 1, shPilote.getLastColumn()).getValues()[0]
  );
  const idxDossier = headers.indexOf("dossier");
  if (idxDossier === -1) return result;

  const data = shPilote.getRange(2, 1, shPilote.getLastRow() - 1, shPilote.getLastColumn()).getValues();
  data.forEach(row => {
    const dossier = String(row[idxDossier] || "").trim();
    if (!dossier) return;
    const edits = {};
    CONFIG.MANAGER_HEADERS.forEach(col => {
      const idx = headers.indexOf(col);
      if (idx !== -1) edits[col] = row[idx];
    });
    result[dossier] = edits;
  });
  return result;
}

function appendHisto_(shHisto, importHeaders, importRows) {
  const histoHeaders = normalizeHeaders_(
    shHisto.getRange(1, 1, 1, shHisto.getLastColumn()).getValues()[0]
  );
  const now = new Date();
  const rowsToAppend = importRows.map(row => {
    const out = new Array(histoHeaders.length).fill("");
    out[0] = now;
    CONFIG.CRM_HEADERS.forEach(col => {
      const idxH = histoHeaders.indexOf(col);
      const idxI = importHeaders.indexOf(col);
      if (idxH !== -1 && idxI !== -1) out[idxH] = row[idxI];
    });
    return out;
  });
  if (rowsToAppend.length > 0) {
    shHisto.getRange(shHisto.getLastRow() + 1, 1, rowsToAppend.length, histoHeaders.length)
      .setValues(rowsToAppend);
  }
}

function setupPilotageValidation_(shPilote) {
  const lastCol = shPilote.getLastColumn();
  if (lastCol === 0) return;
  const headerRow = shPilote.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const maxRows = Math.max(shPilote.getMaxRows() - 1, 1);

  shPilote.getRange(2, 1, maxRows, lastCol).clearDataValidations();

  Object.keys(CONFIG.VALIDATIONS).forEach(colName => {
    const idx = headerRow.indexOf(colName);
    if (idx === -1) return;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.VALIDATIONS[colName], true)
      .setAllowInvalid(false)
      .build();
    shPilote.getRange(2, idx + 1, maxRows, 1).setDataValidation(rule);
  });
}

function applyPilotageFormatting_(shPilote) {
  const lastCol = shPilote.getLastColumn();
  const lastRow = shPilote.getMaxRows();
  if (lastCol === 0) return;

  applyHeaderFormatting_(shPilote);

  const headers = shPilote.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  shPilote.clearConditionalFormatRules();
  const rules = [];

  const dataRange = shPilote.getRange(2, 1, Math.max(lastRow - 1, 1), lastCol);

  // Couleur sur la colonne ALERTE selon le texte
  const idxAlerte = headers.indexOf("alerte") + 1;
  if (idxAlerte > 0) {
    const colAlerte = shPilote.getRange(2, idxAlerte, Math.max(lastRow - 1, 1), 1);
    [
      ["❓", "#d9d2e9"], ["⛔", "#f4cccc"], ["🔴", "#ea9999"],
      ["🟠", "#fce5cd"], ["⚫", "#d9d9d9"]
    ].forEach(([txt, color]) => {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(txt).setBackground(color).setRanges([colAlerte]).build());
    });
  }

  // Couleur sur la colonne CHALEUR
  const idxChaleur = headers.indexOf("niveau_chaleur") + 1;
  if (idxChaleur > 0) {
    const colChaleur = shPilote.getRange(2, idxChaleur, Math.max(lastRow - 1, 1), 1);
    [
      ["Chaud", "#e06666"], ["Tiède", "#f6b26b"], ["Froid", "#9fc5e8"]
    ].forEach(([txt, color]) => {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(txt).setBackground(color).setRanges([colChaleur]).build());
    });
  }

  // Couleur sur MOUVEMENT
  const idxMouv = headers.indexOf("mouvement") + 1;
  if (idxMouv > 0) {
    const colMouv = shPilote.getRange(2, idxMouv, Math.max(lastRow - 1, 1), 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Stagne").setBackground("#fce5cd").setRanges([colMouv]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("Avance").setBackground("#d9ead3").setRanges([colMouv]).build());
  }

  shPilote.setConditionalFormatRules(rules);

  // Format € pour montant et prévision
  ["montant_devis", "prevision_ponderee"].forEach(col => {
    const idx = headers.indexOf(col) + 1;
    if (idx > 0) {
      shPilote.getRange(2, idx, Math.max(lastRow - 1, 1), 1).setNumberFormat("#,##0 €");
    }
  });

  try { shPilote.autoResizeColumns(1, lastCol); } catch (e) {}
}

function applyHeaderFormatting_(sheet) {
  if (sheet.getLastColumn() === 0) return;
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight("bold").setBackground("#d9ead3");
  sheet.setFrozenRows(1);
}

function writeHeadersIfNeeded_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length))
    .getValues()[0].map(h => String(h || "").trim());
  const same = existing.length >= headers.length && headers.every((h, i) => existing[i] === h);
  if (!same) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function rewriteSheet_(sheet, headers, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function removeChartsFromSheet_(sheet) {
  sheet.getCharts().forEach(c => sheet.removeChart(c));
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function hideSheet_(sheet) {
  try { sheet.hideSheet(); } catch (e) {}
}

function normalizeHeaders_(headers) {
  return headers.map(h => String(h || "").trim());
}

function fitRowToHeaders_(row, len) {
  const out = row.slice(0, len);
  while (out.length < len) out.push("");
  return out;
}

function setCell_(row, headers, name, value) {
  const idx = headers.indexOf(name);
  if (idx !== -1) row[idx] = value;
}

function getCell_(row, headers, name) {
  const idx = headers.indexOf(name);
  return idx !== -1 ? row[idx] : "";
}

function pickDatePivot_(row, headers) {
  const candidates = ["datetransodmcommercial", "datecontroledevis", "datevisite"];
  for (let i = 0; i < candidates.length; i++) {
    const v = getCell_(row, headers, candidates[i]);
    if (v) return v;
  }
  return "";
}

/** Jours écoulés depuis une date (positif = passé). "" si non datable. */
function calcDaysFrom_(value) {
  const date = parseAnyDate_(value);
  if (!date) return "";
  const today = new Date();
  const d1 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const d2 = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.floor((d1 - d2) / (1000 * 60 * 60 * 24));
}

function parseAnyDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) return value;
  const str = String(value).trim();
  if (!str) return null;
  const fr = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (fr) {
    const date = new Date(Number(fr[3]), Number(fr[2]) - 1, Number(fr[1]));
    return isNaN(date) ? null : date;
  }
  const parsed = new Date(str);
  return isNaN(parsed) ? null : parsed;
}

function formatDateFR_(date) {
  const d = ("0" + date.getDate()).slice(-2);
  const m = ("0" + (date.getMonth() + 1)).slice(-2);
  return d + "/" + m + "/" + date.getFullYear();
}

function parseMontant_(value) {
  if (value === "" || value === null || value === undefined) return "";
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/[^\d.,-]/g, "").replace(/\s/g, "").replace(",", ".");
  const n = Number(cleaned);
  return isNaN(n) ? "" : n;
}

function probaToNum_(value) {
  const str = String(value || "").replace("%", "").trim();
  if (!str) return "";
  const n = Number(str);
  return isNaN(n) ? "" : n / 100;
}
