/**
 * ============================================================
 *  ENQUÊTES DE SATISFACTION - ZONES SANITAIRES DU BÉNIN
 *  Backend Google Apps Script (Web App)
 * ============================================================
 *  - doGet(e) sert les formulaires HTML et le tableau de bord admin
 *  - submitUsagers / submitPersonnels / submitCogecs ajoutent une ligne
 *    dans la feuille brute correspondante, en UN SEUL appel (rapide),
 *    en retrouvant les colonnes par leur intitulé exact en ligne 1
 *    (peu importe l'ordre).
 *  - Toutes les feuilles de traitement (T_*), de présentation (_V, _ZS)
 *    et le DASHBOARD se recalculent automatiquement.
 *  - getDashboardData(password, token) alimente le tableau de bord admin,
 *    y compris la carte des positions GPS des enquêteurs par jour.
 *    Si un token de session est fourni et correspond à un compte dont le
 *    niveau n'est pas national/administrateur, les données renvoyées
 *    sont filtrées SERVEUR selon la portée du compte (département /
 *    zone sanitaire / commune) — le navigateur ne reçoit jamais les
 *    données hors périmètre.
 *  - accountLogin / changePassword / getQrToken / refreshQrToken gèrent
 *    les comptes personnels (commune, zone, département, national,
 *    administrateur) et les 2 QR codes d'accès (Enquêteur, Personnel).
 * ============================================================
 */

var SHEET_USAGERS    = "👩‍👦 USAGERS";
var SHEET_PERSONNELS = "👩‍⚕️ PERSONNELS";
var SHEET_COGECS     = "👨‍👩‍👧‍👦 COGECS";
var T_USAGERS    = "👩‍👦 T_USAGERS";
var T_PERSONNELS = "👩‍⚕️ T_PERSONNELS";
var T_COGECS     = "👨‍👩‍👧‍👦 T_COGECS";
var KEY_COLUMN       = "Commune"; // colonne toujours renseignée, sert à repérer la 1ère ligne vide

// Mot de passe admin par défaut — CHANGEZ-LE : Extensions > Apps Script >
// Paramètres du projet (⚙️) > Propriétés du script > ajoutez ADMIN_PASSWORD
function getAdminPassword_() {
  var p = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return p || 'admin2025';
}

/**
 * API JSON pure : les pages HTML (Index, formulaires, dashboard) sont
 * maintenant hébergées ailleurs (GitHub Pages) et appellent ce endpoint
 * en fetch(POST). Ce script ne fait plus que lire/écrire le Google Sheet.
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: "API Satisfaction Zones Sanitaires en ligne. Utilisez POST." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    if (action === "submitUsagers")          out = submitUsagers(body.data);
    else if (action === "submitPersonnels")  out = submitPersonnels(body.data);
    else if (action === "submitCogecs")      out = submitCogecs(body.data);
    else if (action === "submitUsagersHopital")   out = submitUsagersHopital(body.data);
    else if (action === "submitPersonnelHopital") out = submitPersonnelHopital(body.data);
    else if (action === "submitCommunauteHopital") out = submitCommunauteHopital(body.data);
    else if (action === "submitActeursComHopital") out = submitActeursComHopital(body.data);
    else if (action === "checkAdminPassword") out = { ok: true, valid: checkAdminPassword(body.password) };
    else if (action === "getDashboardData")  out = getDashboardData(body.password, body.data && body.data.token);
    else if (action === "accountLogin")      out = handleAccountLogin(body.data.username, body.password);
    else if (action === "changePassword")    out = handleChangePassword(body.data.token, body.password, body.data.newPassword);
    else if (action === "changeUsername")    out = handleChangeUsername(body.data.token, body.password, body.data.newUsername);
    else if (action === "getQrToken")        out = handleGetQrToken(body.data.role);
    else if (action === "refreshQrToken")    out = handleRefreshQrToken(body.data.token, body.data.role);
    else if (action === "setQrCode")         out = handleSetQrCode(body.data.token, body.data.role, body.data.code);
    else if (action === "matchQrCode")       out = handleMatchQrCode(body.data.text);
    else if (action === "listAccounts")      out = handleListAccounts(body.data.token);
    else if (action === "setAccountActive")  out = handleSetAccountActive(body.data.token, body.data.username, body.data.active);
    else if (action === "updateAccount")     out = handleUpdateAccount(body.data.token, body.data.username, body.data.fields);
    else if (action === "deleteAccount")     out = handleDeleteAccount(body.data.token, body.data.username);
    else out = { ok: false, error: "Action inconnue : " + action };
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** {enteteExacte: indexColonne} — un seul appel réseau */
function buildHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || "").toString().trim();
    if (h) map[h] = i + 1;
  }
  return { map: map, lastCol: lastCol };
}

/**
 * Ligne suivante disponible. Optimisé : on se base sur getLastRow() (rapide,
 * pas de lecture de plage) plutôt que de scanner toute la colonne clé — sur
 * une feuille de plusieurs milliers de lignes, l'ancien scan complet pouvait
 * ralentir sensiblement chaque enregistrement.
 */
function getNextRow_(sheet, keyColIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  // Vérifie que la dernière ligne a bien une valeur en colonne clé (cas normal).
  var lastVal = sheet.getRange(lastRow, keyColIndex, 1, 1).getValue();
  if (lastVal !== "" && lastVal !== null) return lastRow + 1;
  // Repli rare (dernière ligne vide malgré getLastRow) : recherche la 1ère ligne
  // vide en partant de la fin, sans jamais lire au-delà de getLastRow().
  var values = sheet.getRange(2, keyColIndex, lastRow - 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] !== null) return i + 3;
  }
  return 2;
}

/**
 * Ajoute un enregistrement en UNE SEULE écriture (rapide).
 * data: objet {enteteExacte: valeur, ...}
 */
function appendRecord_(sheetName, data, keyColumnName) {
  keyColumnName = keyColumnName || KEY_COLUMN;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Feuille introuvable: " + sheetName);

    var hm = buildHeaderMap_(sheet);
    var keyCol = hm.map[keyColumnName];
    if (!keyCol) throw new Error("Colonne clé '" + keyColumnName + "' introuvable dans " + sheetName);

    var row = getNextRow_(sheet, keyCol);

    // construit un tableau complet de la ligne, puis UNE seule écriture
    var rowValues = new Array(hm.lastCol).fill("");
    for (var key in data) {
      if (!data.hasOwnProperty(key)) continue;
      var col = hm.map[key];
      if (col) rowValues[col - 1] = data[key];
    }
    sheet.getRange(row, 1, 1, hm.lastCol).setValues([rowValues]);
    SpreadsheetApp.flush();
    return { ok: true, row: row, sheet: sheetName };
  } finally {
    lock.releaseLock();
  }
}

function submitUsagers(data) {
  return appendRecord_(SHEET_USAGERS, data);
}

function submitPersonnels(data) {
  return appendRecord_(SHEET_PERSONNELS, data);
}

function submitCogecs(data) {
  return appendRecord_(SHEET_COGECS, data);
}

/* ============================================================
 *  ENQUÊTES HÔPITAUX — nouvelles feuilles, même classeur
 *  (formulaires distincts, aucune compilation/tableau de bord
 *  n'est construit pour cette base — collecte uniquement)
 * ============================================================ */
var SHEET_USAGERS_H    = "🏥 USAGERS_HOPITAL";
var SHEET_PERSONNELS_H = "🏥 PERSONNEL_HOPITAL";
var SHEET_COMMUNAUTE_H = "🏥 COMMUNAUTE_HOPITAL";
var SHEET_ACTEURS_H    = "🏥 ACTEURS_COM_HOPITAL";

var HEADERS_USAGERS_H    = ["dept", "zone", "fs", "Veuillez_enregistrer_sition_au_moins_90", "Enqueteur", "Sujet_enqu_t", "Hospitalisation", "A1a", "Comment_avez_vous_tr_acc_der_l_h_pital_", "Avez_vous_rencontr_re_v_hicule_ou_moto_", "L_acc_s_aux_b_timent_mobilit_r_duite_", "Une_fois_l_int_rie_neaux_d_orientation_", "La_propret_et_le_co_ils_au_rendez_vous_", "Avez_vous_des_suggestions_part", "Si_oui_lesquelles", "Avez_vous_des_diffic_au_sein_de_l_h_pital", "A1", "_3_Qu_est_ce_que_vous_n_avez_pas_aim", "A2", "A3", "_6_Qu_est_ce_que_vous_n_avez_pas_aim", "A3a", "A4", "_9_Qu_est_ce_que_vous_n_avez_pas_aim", "A4a", "A5", "_12_Qu_est_ce_que_vous_n_avez_pas_aim", "A5a", "A6", "_15_Qu_est_ce_que_vous_n_avez_pas_aim", "A6a", "C1", "_2_Dites_nous_un_peu_comment_a_a_t", "C2", "D1", "_5_Dites_nous_un_peu_comment_a_a_t", "D2", "E1", "B1", "B2", "B3", "G1", "G", "G2", "F1", "F2", "F3", "C3", "E2", "E3", "E4", "A7", "A8", "D2_Precise_001", "couleur_tenue", "C4", "C5", "H1", "H2", "Avant_qu_une_d_cisio_vous_t_consult_e_", "Avez_vous_pu_exprime_les_soins_propos_s_", "Vous_est_il_d_j_arriv_de_ref", "Si_oui_votre_d_cisi_ression_ni_jugement_", "Lors_de_votre_prise_en_charge_", "Avez_vous_effectivement_d_sign", "Si_oui_cette_person_essionnels_de_sant_", "Lors_de_votre_prise_en_charge", "Lors_de_votre_prise_en_charge_001", "Lors_de_votre_prise_en_charge__001", "H3", "H4", "H5", "H6", "H7", "C6", "H8", "AA1", "D1_Precise", "AA2", "D2_Precise", "AA3", "D3_Precise", "AA4", "D4_Precise", "I1", "Si_non_pourquoi", "Si_oui_pourquoi", "I2", "Si_non_pourquoi_001", "Si_oui_pourquoi_001", "I3", "I4", "I5", "Age", "Sexe", "departement", "commune", "Profession", "Instruction", "Raison_Choix", "Avez_vous_autre_chose_dire", "today", "_start-geopoint_latitude", "_start-geopoint_longitude", "_start-geopoint_altitude", "_start-geopoint_precision"];
var HEADERS_PERSONNELS_H = ["dept", "zs", "fs", "Formulaire_renseign_par", "Choisir_l_identifiant", "age", "Sexe", "Q1", "Q2", "Q3", "Avez_vous_re_u_une_f_evoirs_des_patients_", "Vous_sentez_vous_sou_evoirs_des_patients_", "Avez_vous_connaissance_des_dro", "Ces_droits_et_devoir_votre_tablissement_", "Pensez_vous_que_vos_onnement_de_travail_", "Avez_vous_des_difficult_s_dans", "Quelles_sont_les_pri_cice_de_vos_devoirs_", "Q4", "Q5", "Q6", "Q7", "Q8", "Recevez_vous_les_inf_d_lais_raisonnables_", "Quels_sont_les_canau_ut_ce_qui_s_applique", "Q9", "Les_coll_gues_de_vot_protocoles_de_soins_", "Les_coll_gues_assure_esoins_des_patients_", "Les_coll_gues_collab_ontinuit_des_soins_", "Q10", "Q11", "Q12", "Q13", "Q14", "Q15", "Q16", "Q17", "Q18", "Q19", "Si_oui_de_quelle_mani_re", "Quelle_autre_mani_re", "Quelles_sont_les_dif_cc_s_l_information", "Q20", "Q21", "Q22", "Vous_pensez_que_l_av_chnique_de_mes_soins", "Vous_pensez_que_l_im_d_erreurs_m_dicales", "Je_re_ois_r_guli_rem_ction_de_mon_service", "Les_plaintes_des_pat_difier_nos_pratiques", "Le_manque_de_temps_e_des_retours_patients", "Les_retours_des_pati_utiles_cliniquement", "Pouvez_vous_citer_un_exemple_r", "Si_oui_pr_ciser", "Selon_vous_quels_so_l_h_pital_de_Pob_", "_9_Quels_sont_les_pri_antage_les_patients_", "Q23", "Q24"];
var HEADERS_COMMUNAUTE_H = ["dept", "zone", "Formation_Sanitaire", "Enqueteur", "Comment_avez_vous_tr_acc_der_l_h_pital_", "Avez_vous_rencontr_re_v_hicule_ou_moto_", "L_acc_s_aux_b_timent_mobilit_r_duite_", "Une_fois_l_int_rie_neaux_d_orientation_", "La_propret_et_le_co_ils_au_rendez_vous_", "Avez_vous_des_suggestions_part", "Si_oui_lesquelles", "Comment_jugez_vous_e_accueil_au_portail_", "Comment_jugez_vous_e_l_dans_les_services_", "Comment_jugez_vous_e_au_service_de_sortie", "Comment_trouvez_vous_les_dirigeants_", "Comment_appr_ciez_vo_es_agents_examiner", "Comment_appr_ciez_vo_diguer_de_bons_soins", "Comment_appr_ciez_vo_e_bonne_prescription", "Comment_appr_ciez_vo_niquer_sur_les_actes", "Pensez_vous_que_les_la_prise_en_charge_", "Comment_trouvez_vous_rapport_avec_le_mal", "Pensez_vous_que_les_ous_servir_au_besoin", "Les_m_dicaments_pres_ibles_dans_le_centre", "Les_m_dicaments_non_pharmacies_proches_", "Comment_trouvez_vous_ifs_des_m_dicaments_", "Comment_trouvez_vous_ents_hospitalisation", "Comment_trouvez_vous_la_cour_de_l_h_pital", "Comment_trouvez_vous_alle_de_consultation", "Comment_trouvez_vous_les_toilettes", "Comment_trouvez_vous_la_buanderie", "Pensez_vous_que_les_onn_es_sur_les_soins", "Pensez_vous_qu_on_a_est_pas_souvent_seul", "Avez_vous_acc_s_l_eau", "Vous_a_t_on_une_fois_des_soins_domicile", "Vous_est_il_jamais_a_r_des_agents_de_sant", "Des_agents_vous_ont_eux_directement_vers", "Pensez_vous_que_les_h_pital_g_rent_bien", "Pensez_vous_que_dans_u_sein_de_l_h_pital_", "Si_vous_devrez_reven_s_motivants_ce_choix", "Age", "Sexe", "Profession", "Instruction", "Avez_vous_autre_chose_dire"];
var HEADERS_ACTEURS_H    = ["dept", "zone", "fs", "Enqueteur", "Que_pensez_vous_de_l_ette_structure_fs_", "Comment_appr_ciez_vo_ette_structure_fs_", "Selon_vous_les_soignant_respe", "Si_Oui_en_quoi", "Si_Non_en_quoi", "Les_usagers_de_cette_des_services_offerts", "Que_proposez_vous_ette_structure_fs_"];

/** À exécuter UNE FOIS depuis l'éditeur Apps Script (menu ▶ Exécuter) pour
 *  créer les 4 feuilles Hôpitaux avec leurs en-têtes, si elles n'existent
 *  pas déjà. Ne touche à aucune feuille existante des zones sanitaires. */
function setupFeuillesHopitaux(){
  var pairs = [
    [SHEET_USAGERS_H, HEADERS_USAGERS_H],
    [SHEET_PERSONNELS_H, HEADERS_PERSONNELS_H],
    [SHEET_COMMUNAUTE_H, HEADERS_COMMUNAUTE_H],
    [SHEET_ACTEURS_H, HEADERS_ACTEURS_H]
  ];
  pairs.forEach(function(p){ getOrCreateSheet_(p[0], p[1]); });
  Logger.log('Feuilles Hôpitaux prêtes : ' + pairs.map(function(p){return p[0];}).join(', '));
}

function submitUsagersHopital(data) {
  return appendRecord_(SHEET_USAGERS_H, data, 'dept');
}
function submitPersonnelHopital(data) {
  return appendRecord_(SHEET_PERSONNELS_H, data, 'dept');
}
function submitCommunauteHopital(data) {
  return appendRecord_(SHEET_COMMUNAUTE_H, data, 'dept');
}
function submitActeursComHopital(data) {
  return appendRecord_(SHEET_ACTEURS_H, data, 'dept');
}

/* ============================================================
 *  TABLEAU DE BORD ADMIN
 * ============================================================ */

function checkAdminPassword(password) {
  return password === getAdminPassword_();
}

/** Lit une colonne entière (par en-tête) en un seul appel, valeurs non vides seulement */
function readColumn_(sheet, headerMap, headerName) {
  var col = headerMap[headerName];
  if (!col) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var vals = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  return vals.map(function(r){ return r[0]; });
}

/** Compte les occurrences (ignore les vides) */
function countBy_(values) {
  var out = {};
  values.forEach(function(v){
    if (v === "" || v === null || v === undefined) return;
    var k = v.toString();
    out[k] = (out[k] || 0) + 1;
  });
  return out;
}

/**
 * Calcule, pour chaque ligne d'une feuille de traitement T_*, la moyenne des
 * SEULES colonnes de score (listées dans CATEGORY_MAP pour cette cible) — et
 * non plus de toutes les colonnes numériques de la ligne. L'ancienne version
 * incluait par erreur des colonnes comme l'Âge (un nombre brut, pas un score
 * 0/0.5/1), ce qui pouvait faire dépasser 100% à la satisfaction affichée.
 * Chaque valeur est en plus bornée à [0,1] par sécurité avant la moyenne.
 */
function computeRowScoresForCategories_(tSheet, categories, nRawRows) {
  if (!nRawRows || nRawRows <= 0) return [];
  var lastRow = tSheet.getLastRow();
  if (lastRow < 3) return [];
  var idxs = [];
  categories.forEach(function(cat){
    cat.cols.forEach(function(l){ idxs.push(columnLetterToIndex_(l) - 1); });
  });
  var maxIdx = Math.max.apply(null, idxs);
  var nRows = Math.min(nRawRows, lastRow - 2);
  if (nRows <= 0) return [];
  var range = tSheet.getRange(3, 1, nRows, maxIdx + 1).getValues();
  return range.map(function(row){
    var sum = 0, n = 0;
    idxs.forEach(function(ci){
      var v = row[ci];
      if (typeof v === "number") { sum += Math.max(0, Math.min(1, v)); n++; }
    });
    return n > 0 ? sum / n : null;
  });
}

/** Convertit un score 0..1 en pourcentage arrondi, jamais au-dessus de 100. */
function clampPct_(ratio) {
  return Math.min(100, Math.round(ratio * 1000) / 10);
}

function groupWithScore_(keys, scores) {
  var out = {}; // key -> {count, sum}
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (k === "" || k === null || k === undefined) continue;
    k = k.toString();
    var s = scores[i];
    if (!out[k]) out[k] = { count: 0, sum: 0, n: 0 };
    out[k].count++;
    if (typeof s === "number") { out[k].sum += s; out[k].n++; }
  }
  var result = [];
  for (var k in out) {
    result.push({
      label: k,
      count: out[k].count,
      avgScore: out[k].n > 0 ? clampPct_(out[k].sum / out[k].n) : null // en %
    });
  }
  result.sort(function(a,b){ return b.count - a.count; });
  return result;
}

/** Moyenne (en %) d'un tableau de scores 0..1, en ignorant les valeurs non numériques */
function avgOf_(scores) {
  var sum = 0, n = 0;
  scores.forEach(function(s){ if (typeof s === "number") { sum += s; n++; } });
  return n > 0 ? clampPct_(sum / n) : null;
}

/** Fusionne plusieurs {keys, scores} alignés en un seul groupWithScore_ */
function groupWithScoreMulti_(pairs) {
  var allKeys = [], allScores = [];
  pairs.forEach(function(p){
    allKeys = allKeys.concat(p.keys);
    allScores = allScores.concat(p.scores);
  });
  return groupWithScore_(allKeys, allScores);
}

/* ============================================================
 *  REGROUPEMENT DES QUESTIONS PAR CATÉGORIE (repris des onglets _V)
 * ============================================================ */
var CATEGORY_MAP = {
  usagers: [
    { name: '1. Accueil et orientation', cols: ["G","H","I","J"] },
    { name: '2. Information et communication', cols: ["K","L","M","N","O","P","Q"] },
    { name: '3. Qualité de l’examen et de l’écoute', cols: ["U","V","W","X"] },
    { name: '4. Soins et continuité', cols: ["Y","Z","AA","AB","AC","AD","AE"] },
    { name: '5. Respect, intimité et confidentialité', cols: ["AF","AG","AH","AI"] },
    { name: '6. Accessibilité financière et médicaments', cols: ["AJ","AK","AL","AM","AN","AO"] },
    { name: '7. Environnement et hygiène', cols: ["AQ","AR","AS","AT","AU","AV","AW"] },
    { name: '8. Équité, respect et compétence du personnel', cols: ["AX","AY","AZ","BA","BB","BC","BD","BE"] },
    { name: '9. Transparence et pratiques abusives', cols: ["BF","BG","BH","BI","BJ"] },
    { name: '10. Rapidité et réactivité de la prise en charge', cols: ["BK","BL","BM"] },
    { name: '11. Vaccination', cols: ["R","S","T","AP","BN"] }
  ],
  personnels: [
    { name: 'Ressources matérielles et infrastructures', cols: ["I","J","K"] },
    { name: 'Organisation du travail et charge professionnelle', cols: ["L","M","N"] },
    { name: 'Valorisation, reconnaissance et perspectives professionnelles', cols: ["O","P","Q"] },
    { name: 'Relations interpersonnelles et climat social', cols: ["R","S","T"] },
    { name: 'Protocoles, directives et communication institutionnelle', cols: ["U","V","W"] },
    { name: 'Compétences techniques et fiabilité de l’équipe', cols: ["X","Y","Z"] },
    { name: 'Supervision, encadrement et formation', cols: ["AA","AB","AC","AD","AE","AF"] },
    { name: 'Information et charte des droits des patients', cols: ["AG","AH","AI","AJ","AK","AL","AM","AN"] },
    { name: 'Pratiques éthiques et relation patient', cols: ["AO","AP","AQ","AR","AS"] },
    { name: 'Gouvernance et collaboration avec le COGECS', cols: ["AT","AU","AV","AW","AX"] }
  ],
  cogecs: [
    { name: '1. Gouvernance et gestion du centre de santé', cols: ["E","F","G","H","I","J"] },
    { name: '2. Collaboration interne avec le personnel soignant', cols: ["K","L","M","N"] },
    { name: '3. Relation avec la coordination de zone sanitaire', cols: ["O","P","Q"] },
    { name: '4. Capacités et moyens des membres du COGECS', cols: ["R","S","T"] }
  ]
};

function columnLetterToIndex_(letter) {
  var col = 0;
  for (var i = 0; i < letter.length; i++) col = col * 26 + (letter.charCodeAt(i) - 64);
  return col;
}

/** Pour chaque catégorie, la moyenne (par ligne brute) des seules colonnes de cette catégorie */
function computeCategoryScores_(tSheet, categories, nRawRows) {
  var out = {};
  if (!nRawRows || nRawRows <= 0) {
    categories.forEach(function(cat){ out[cat.name] = []; });
    return out;
  }
  var lastRow = tSheet.getLastRow();
  if (lastRow < 3) {
    categories.forEach(function(cat){ out[cat.name] = []; });
    return out;
  }
  var nRows = Math.min(nRawRows, lastRow - 2);
  if (nRows <= 0) {
    categories.forEach(function(cat){ out[cat.name] = []; });
    return out;
  }
  var lastCol = tSheet.getLastColumn();
  var range = tSheet.getRange(3, 1, nRows, lastCol).getValues();
  categories.forEach(function(cat){
    var idxs = cat.cols.map(function(l){ return columnLetterToIndex_(l) - 1; });
    out[cat.name] = range.map(function(row){
      var sum = 0, n = 0;
      idxs.forEach(function(ci){
        var v = row[ci];
        if (typeof v === "number") { sum += Math.max(0, Math.min(1, v)); n++; }
      });
      return n > 0 ? sum / n : null;
    });
  });
  return out;
}

/** Construit, pour une cible donnée, le détail par commune : effectif, âge moyen,
 *  répartition par sexe, et score moyen (%) pour chaque catégorie de questions. */
function buildCommuneReport_(sheetName, tSheetName, categories) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var tSheet = ss.getSheetByName(tSheetName);
  var hm = buildHeaderMap_(sheet).map;

  var communes = readColumn_(sheet, hm, "Commune");
  var ages = readColumn_(sheet, hm, "Âge");
  var sexes = readColumn_(sheet, hm, "Sexe");
  var n = communes.length;
  var catScores = computeCategoryScores_(tSheet, categories, n);
  var catNames = categories.map(function(c){ return c.name; });

  var byCommune = {};
  for (var i = 0; i < n; i++) {
    var commune = communes[i];
    if (commune === "" || commune === null || commune === undefined) continue;
    commune = commune.toString();
    if (!byCommune[commune]) {
      byCommune[commune] = { count: 0, ageSum: 0, ageN: 0, genderCounts: {}, catSum: {}, catN: {} };
      catNames.forEach(function(cn){ byCommune[commune].catSum[cn] = 0; byCommune[commune].catN[cn] = 0; });
    }
    var rec = byCommune[commune];
    rec.count++;
    if (typeof ages[i] === "number") { rec.ageSum += ages[i]; rec.ageN++; }
    var sexe = sexes[i];
    if (sexe !== "" && sexe !== null && sexe !== undefined) {
      var sk = sexe.toString();
      rec.genderCounts[sk] = (rec.genderCounts[sk] || 0) + 1;
    }
    catNames.forEach(function(cn){
      var s = catScores[cn][i];
      if (typeof s === "number") { rec.catSum[cn] += s; rec.catN[cn]++; }
    });
  }

  var result = {};
  for (var commune in byCommune) {
    var rec = byCommune[commune];
    var catPct = {};
    catNames.forEach(function(cn){
      catPct[cn] = rec.catN[cn] > 0 ? clampPct_(rec.catSum[cn] / rec.catN[cn]) : null;
    });
    var genderPct = {};
    for (var g in rec.genderCounts) genderPct[g] = Math.round((rec.genderCounts[g] / rec.count) * 1000) / 10;
    result[commune] = {
      count: rec.count,
      avgAge: rec.ageN > 0 ? Math.round((rec.ageSum / rec.ageN) * 10) / 10 : null,
      genderPct: genderPct,
      categoryScores: catPct
    };
  }
  return { categories: catNames, byCommune: result };
}

/** Comme buildCommuneReport_, mais regroupé par formation sanitaire (Structure).
 *  Chaque entrée garde sa Commune, ce qui permet de recalculer côté client
 *  n'importe quel niveau plus large (Zone sanitaire, Département, National)
 *  par simple moyenne pondérée, sans nouvel appel serveur. */
function buildStructureReport_(sheetName, tSheetName, categories) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var tSheet = ss.getSheetByName(tSheetName);
  var hm = buildHeaderMap_(sheet).map;

  var structures = readColumn_(sheet, hm, "Structure");
  var communes = readColumn_(sheet, hm, "Commune");
  var n = structures.length;
  var catScores = computeCategoryScores_(tSheet, categories, n);
  var catNames = categories.map(function(c){ return c.name; });

  var byStructure = {};
  for (var i = 0; i < n; i++) {
    var structure = structures[i];
    if (structure === "" || structure === null || structure === undefined) continue;
    structure = structure.toString();
    if (!byStructure[structure]) {
      byStructure[structure] = { count: 0, commune: (communes[i] || "").toString(), catSum: {}, catN: {} };
      catNames.forEach(function(cn){ byStructure[structure].catSum[cn] = 0; byStructure[structure].catN[cn] = 0; });
    }
    var rec = byStructure[structure];
    rec.count++;
    catNames.forEach(function(cn){
      var s = catScores[cn][i];
      if (typeof s === "number") { rec.catSum[cn] += s; rec.catN[cn]++; }
    });
  }

  var result = {};
  for (var structure in byStructure) {
    var rec = byStructure[structure];
    var catPct = {};
    catNames.forEach(function(cn){
      catPct[cn] = rec.catN[cn] > 0 ? clampPct_(rec.catSum[cn] / rec.catN[cn]) : null;
    });
    result[structure] = { commune: rec.commune, count: rec.count, categoryScores: catPct };
  }
  return { categories: catNames, byStructure: result };
}

/** Comme buildStructureReport_, mais regroupé par enquêteur. Ne s'applique
 *  qu'aux formulaires qui capturent l'identité de l'enquêteur (Usagers, COGECS) ;
 *  pour les autres, la colonne est absente et le résultat est simplement vide. */
function buildEnqueteurReport_(sheetName, tSheetName, categories) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var tSheet = ss.getSheetByName(tSheetName);
  var hm = buildHeaderMap_(sheet).map;

  var enqueteurs = readColumn_(sheet, hm, "Identification de l'enquêteur");
  var n = enqueteurs.length;
  var catScores = computeCategoryScores_(tSheet, categories, n);
  var catNames = categories.map(function(c){ return c.name; });

  var byEnqueteur = {};
  for (var i = 0; i < n; i++) {
    var enq = enqueteurs[i];
    if (enq === "" || enq === null || enq === undefined) continue;
    enq = enq.toString();
    if (!byEnqueteur[enq]) {
      byEnqueteur[enq] = { count: 0, catSum: {}, catN: {} };
      catNames.forEach(function(cn){ byEnqueteur[enq].catSum[cn] = 0; byEnqueteur[enq].catN[cn] = 0; });
    }
    var rec = byEnqueteur[enq];
    rec.count++;
    catNames.forEach(function(cn){
      var s = catScores[cn][i];
      if (typeof s === "number") { rec.catSum[cn] += s; rec.catN[cn]++; }
    });
  }

  var result = {};
  for (var enq2 in byEnqueteur) {
    var rec2 = byEnqueteur[enq2];
    var catPct = {};
    catNames.forEach(function(cn){
      catPct[cn] = rec2.catN[cn] > 0 ? clampPct_(rec2.catSum[cn] / rec2.catN[cn]) : null;
    });
    result[enq2] = { count: rec2.count, categoryScores: catPct };
  }
  return { categories: catNames, byEnqueteur: result };
}

/* ============================================================
 *  CARTE DES ENQUÊTEURS — positions GPS par jour de collecte
 *  (uniquement le formulaire Usagers capture le GPS)
 * ============================================================ */
function getEnqueteurPositions_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USAGERS);
  if (!sheet) return {};
  var hm = buildHeaderMap_(sheet).map;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var dates      = readColumn_(sheet, hm, "today");
  var enqueteurs = readColumn_(sheet, hm, "Identification de l'enquêteur");
  var communes   = readColumn_(sheet, hm, "Commune");
  var structures = readColumn_(sheet, hm, "Structure");
  var lats       = readColumn_(sheet, hm, "_start-geopoint_latitude");
  var lons       = readColumn_(sheet, hm, "_start-geopoint_longitude");
  var accs       = readColumn_(sheet, hm, "_start-geopoint_precision");
  var starts     = readColumn_(sheet, hm, "start"); // absent tant que les 2 colonnes ne sont pas ajoutées au classeur
  var ends       = readColumn_(sheet, hm, "end");

  var byDay = {};
  for (var i = 0; i < dates.length; i++) {
    var lat = lats[i], lon = lons[i];
    if (lat === "" || lon === "" || lat === null || lon === null || lat === undefined || lon === undefined) continue;
    if (isNaN(Number(lat)) || isNaN(Number(lon))) continue;
    var day = (dates[i] || "Date inconnue").toString();
    if (!byDay[day]) byDay[day] = [];
    var durationMin = null;
    if (starts[i] && ends[i]) {
      var sMs = new Date(starts[i]).getTime(), eMs = new Date(ends[i]).getTime();
      if (!isNaN(sMs) && !isNaN(eMs) && eMs > sMs) durationMin = Math.round((eMs - sMs) / 6000) / 10; // 1 décimale
    }
    byDay[day].push({
      enqueteur: (enqueteurs[i] || "Non renseigné").toString(),
      commune: (communes[i] || "").toString(),
      structure: (structures[i] || "").toString(),
      lat: Number(lat),
      lon: Number(lon),
      accuracy: (accs[i] !== "" && accs[i] !== null && accs[i] !== undefined) ? Math.round(Number(accs[i])) : null,
      durationMin: durationMin
    });
  }
  return byDay; // { "2026-07-19": [ {enqueteur, commune, structure, lat, lon, accuracy, durationMin}, ... ], ... }
}

function getDashboardData(password, sessionToken) {
  var account = null;
  var authOk = false;

  if (sessionToken) {
    var session = findSession_(sessionToken);
    if (session) {
      var acc = findAccountRow_(session.username);
      if (acc) { authOk = true; account = accountPublicView_(acc.data); }
    }
  }
  if (!authOk && password === getAdminPassword_()) {
    authOk = true; // compatibilité : mot de passe partagé seul, sans compte personnel
  }
  if (!authOk) {
    return { ok: false, error: "Authentification requise (connectez-vous depuis l'accueil)." };
  }

  var errors = [];
  function safe_(label, fn, fallback){
    try { return fn(); }
    catch (err) { errors.push(label + " : " + err.message); return fallback; }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var uSheet = ss.getSheetByName(SHEET_USAGERS);
  var cSheet = ss.getSheetByName(SHEET_COGECS);
  var pSheet = ss.getSheetByName(SHEET_PERSONNELS);
  var tuSheet = ss.getSheetByName(T_USAGERS);
  var tcSheet = ss.getSheetByName(T_COGECS);
  var tpSheet = ss.getSheetByName(T_PERSONNELS);

  var uHm = uSheet ? buildHeaderMap_(uSheet).map : {};
  var cHm = cSheet ? buildHeaderMap_(cSheet).map : {};
  var pHm = pSheet ? buildHeaderMap_(pSheet).map : {};

  var uCommunes   = safe_("uCommunes",   function(){ return readColumn_(uSheet, uHm, "Commune"); }, []);
  var uStructures = safe_("uStructures", function(){ return readColumn_(uSheet, uHm, "Structure"); }, []);
  var uEnqueteurs = safe_("uEnqueteurs", function(){ return readColumn_(uSheet, uHm, "Identification de l'enquêteur"); }, []);
  var uScores     = safe_("uScores",     function(){ return computeRowScoresForCategories_(tuSheet, CATEGORY_MAP.usagers, uCommunes.length); }, []);

  var cCommunes   = safe_("cCommunes",   function(){ return readColumn_(cSheet, cHm, "Commune"); }, []);
  var cStructures = safe_("cStructures", function(){ return readColumn_(cSheet, cHm, "Structure"); }, []);
  var cScores     = safe_("cScores",     function(){ return computeRowScoresForCategories_(tcSheet, CATEGORY_MAP.cogecs, cCommunes.length); }, []);

  var pCommunes   = safe_("pCommunes",   function(){ return readColumn_(pSheet, pHm, "Commune"); }, []);
  var pStructures = safe_("pStructures", function(){ return readColumn_(pSheet, pHm, "Structure"); }, []);
  var pScores     = safe_("pScores",     function(){ return computeRowScoresForCategories_(tpSheet, CATEGORY_MAP.personnels, pCommunes.length); }, []);

  var uCount = uCommunes.filter(function(v){return v!=="";}).length;
  var cCount = cCommunes.filter(function(v){return v!=="";}).length;
  var pCount = pCommunes.filter(function(v){return v!=="";}).length;

  var emptyCommuneReport = { categories: [], byCommune: {} };

  var result = {
    ok: true,
    counts: {
      usagers: uCount,
      cogecs: cCount,
      personnels: pCount,
      total: uCount + cCount + pCount
    },
    usagers: {
      byCommune: safe_("usagers.byCommune", function(){ return groupWithScore_(uCommunes, uScores); }, []),
      byStructure: safe_("usagers.byStructure", function(){ return groupWithScore_(uStructures, uScores); }, []),
      byEnqueteur: safe_("usagers.byEnqueteur", function(){ return groupWithScore_(uEnqueteurs, uScores); }, [])
    },
    cogecs: {
      byCommune: safe_("cogecs.byCommune", function(){ return groupWithScore_(cCommunes, cScores); }, []),
      byStructure: safe_("cogecs.byStructure", function(){ return groupWithScore_(cStructures, cScores); }, [])
    },
    personnels: {
      byCommune: safe_("personnels.byCommune", function(){ return groupWithScore_(pCommunes, pScores); }, []),
      byStructure: safe_("personnels.byStructure", function(){ return groupWithScore_(pStructures, pScores); }, [])
    },
    global: {
      byCible: [
        { label: "Usagers",            count: uCount, avgScore: safe_("avgOf uScores", function(){ return avgOf_(uScores); }, null) },
        { label: "Personnel soignant", count: pCount, avgScore: safe_("avgOf pScores", function(){ return avgOf_(pScores); }, null) },
        { label: "COGECS",             count: cCount, avgScore: safe_("avgOf cScores", function(){ return avgOf_(cScores); }, null) }
      ],
      byCommune: safe_("global.byCommune", function(){
        return groupWithScoreMulti_([
          { keys: uCommunes, scores: uScores },
          { keys: cCommunes, scores: cScores },
          { keys: pCommunes, scores: pScores }
        ]);
      }, [])
    },
    communeReport: {
      usagers:    safe_("communeReport.usagers",    function(){ return buildCommuneReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers); }, emptyCommuneReport),
      personnels: safe_("communeReport.personnels", function(){ return buildCommuneReport_(SHEET_PERSONNELS, T_PERSONNELS, CATEGORY_MAP.personnels); }, emptyCommuneReport),
      cogecs:     safe_("communeReport.cogecs",      function(){ return buildCommuneReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs); }, emptyCommuneReport)
    },
    structureReport: {
      usagers:    safe_("structureReport.usagers",    function(){ return buildStructureReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers); }, { categories: [], byStructure: {} }),
      personnels: safe_("structureReport.personnels", function(){ return buildStructureReport_(SHEET_PERSONNELS, T_PERSONNELS, CATEGORY_MAP.personnels); }, { categories: [], byStructure: {} }),
      cogecs:     safe_("structureReport.cogecs",      function(){ return buildStructureReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs); }, { categories: [], byStructure: {} })
    },
    enqueteurReport: {
      usagers: safe_("enqueteurReport.usagers", function(){ return buildEnqueteurReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers); }, { categories: [], byEnqueteur: {} }),
      cogecs:  safe_("enqueteurReport.cogecs",  function(){ return buildEnqueteurReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs); }, { categories: [], byEnqueteur: {} })
    },
    gpsByDay: safe_("gpsByDay", function(){ return getEnqueteurPositions_(); }, {}),
    errors: errors,
    generatedAt: new Date().toISOString()
  };

  // ---- Filtrage SERVEUR selon la portée du compte connecté (si un compte a été résolu) ----
  if (account) applyScopeToResult_(result, account);

  return result;
}

/* ============================================================
 *  COMPTES PERSONNELS (commune / zone / département / national / admin)
 *  + QR CODES D'ACCÈS (Enquêteur, Personnel)
 * ============================================================ */

var ACCOUNTS_SHEET_NAME = 'Comptes';
var SESSIONS_SHEET_NAME = 'Sessions';
var QRTOKENS_SHEET_NAME = 'QRTokens';
var SESSION_DURATION_MS = 12 * 60 * 60 * 1000; // 12 heures

// [identifiant, hash, niveau, département, zone sanitaire, commune, libellé]
var SEED_ACCOUNTS = [["commune.abomey", "82e73a3bc26ab1266e67ae26a157660b92db721b34b5f2b5c078d30194303f29", "commune", "Zou", "Djidja-Abomey-Agbangnizoun", "Abomey", "Commune de Abomey", "HBRPOI"], ["commune.abomey-calavi", "9ccc7b5ceb9278196c4bc9e2fd1abfdc06018c91181859806200d3d05f803c8e", "commune", "Atlantique", "Abomey-Calavi-So-ava", "Abomey-Calavi", "Commune de Abomey-Calavi", "G8F1CB"], ["commune.adja-ouere", "47aaee09ac655aeaf96d326d267e962d696f8e02780f44aa311f9caf7978912a", "commune", "Plateau", "Pobè-Kétou-Adja-ouèrè", "Adja-Ouèrè", "Commune de Adja-Ouèrè", "FNO6B9"], ["commune.adjarra", "bd636c339444304941d0949bd659bbc5819c676ffb17317e43a78245e8bcc3f1", "commune", "Ouémé", "Akpro-missérété-Avrankou-Adjarra", "Adjarra", "Commune de Adjarra", "M80O2R"], ["commune.adjohoun", "8e3afcec5eb7ad6d39cbb8c3699b1c18a78e89870e966bf59876a0f50c439267", "commune", "Ouémé", "Adjohoun-Bonou-Dangbo", "Adjohoun", "Commune de Adjohoun", "AK1VRJ"], ["commune.agbangnizoun", "69a96acc1743b031f70cc426ec13ed881ec5f3bf0fa7a7af175d4b0f235da5f1", "commune", "Zou", "Djidja-Abomey-Agbangnizoun", "Agbangnizoun", "Commune de Agbangnizoun", "NVGFYG"], ["commune.aguegues", "c19860cacc34da3c8b77692c0fed3a8b30f94ca2da02344a8b7a6edbf068127e", "commune", "Ouémé", "Porto-Novo-Sèmè-kpodji-Aguégués", "Aguegues", "Commune de Aguegues", "WWQC38"], ["commune.akpro-misserete", "0082975f401af7de065e49976f59663d0198f07a0a3ba5f012a219e5ab11a271", "commune", "Ouémé", "Akpro-missérété-Avrankou-Adjarra", "Akpro-Missereté", "Commune de Akpro-Missereté", "HYF9SX"], ["commune.allada", "f2c7f5cd6d7cd8ea40e17a0240b10e6302bab10e98384026f4500daa538e9fed", "commune", "Atlantique", "Allada-Toffo-Zè", "Allada", "Commune de Allada", "MECOSF"], ["commune.aplahoue", "b3db46f6612e7f62a0e50bf8deb2f03be82f89e09a52965685cae4b2b6f614ba", "commune", "Couffo", "Aplahoué-Djakotomey-Dogbo", "Aplahoué", "Commune de Aplahoué", "OGYR3X"], ["commune.athieme", "cc09ba0120cae9a00cb8413f906186c745f9ff3d83f1fec1945208a7391960b1", "commune", "Mono", "Lokossa-Athiémè", "Athieme", "Commune de Athieme", "KXWNRE"], ["commune.avrankou", "bf969a32b5c0d30a9a67284f4a759ff744794705cab449939f4fc77a5f8b201d", "commune", "Ouémé", "Akpro-missérété-Avrankou-Adjarra", "Avrankou", "Commune de Avrankou", "K8PK3Y"], ["commune.banikoara", "eafed6325706ad3b2466ff9ec1f8156aeaee58540b8d1376c7dbc3234a8fb612", "commune", "Alibori", "Banikoara", "Banikoara", "Commune de Banikoara", "R9OUDO"], ["commune.bante", "bf64a6ea967d2bed6117f996921f334935ee01bbc92c2310d405b51d4ac6956c", "commune", "Collines", "Savalou-Bantè", "Bante", "Commune de Bante", "CUZREN"], ["commune.bassila", "56e43a56dd21a99d728a069aff5baad953b25a6277ec45b88b28b4f0f2bce1c4", "commune", "Donga", "Bassila", "Bassila", "Commune de Bassila", "UN5Z3J"], ["commune.bembereke", "37c76e2530af41ae16e31a5cd1422e3f9fd3927a41c5bb2b852b4128a0e3545b", "commune", "Borgou", "Bembèrèkè-Sinendé", "Bembèrèkè", "Commune de Bembèrèkè", "QIP98Q"], ["commune.bohicon", "85d6dd4c413544c46e9c19b7121584b86a8ac592f58deae46eb4247b9b927e15", "commune", "Zou", "Bohicon-Za-kpota-Zogbodomey", "Bohicon", "Commune de Bohicon", "1ZXOI6"], ["commune.bonou", "66e32c5286ffbb01169427d2e098aed5b56877a0010d70b6ddff94aea41e6461", "commune", "Ouémé", "Adjohoun-Bonou-Dangbo", "Bonou", "Commune de Bonou", "5FDHJK"], ["commune.bopa", "67cde2c323b922512e39ea691cb01ba705526e90912352f22811398550f8576a", "commune", "Mono", "Comè-Grand popo-Houèyogbé-Bopa", "Bopa", "Commune de Bopa", "1EYY37"], ["commune.boukoumbe", "3cf392e368ad27b7dbee91e788c578a5e27e3dbdc712b1bc7cfc3f073234dc3c", "commune", "Atacora", "Natitingou-Boukoumbé-Toucountouna", "Boukoumbe", "Commune de Boukoumbe", "Q9AH8R"], ["commune.cobly", "6d564db2c3d6aff9787de0a4538099e05b72612140406e7469b8dbcbf6f1cc54", "commune", "Atacora", "Tanguiéta-Cobly-Matéri", "Cobly", "Commune de Cobly", "VHS1K3"], ["commune.come", "21a6abaaad3d0e530ea005f32dfb03a35cb7ff83109886dc511fcf60355871b9", "commune", "Mono", "Comè-Grand popo-Houèyogbé-Bopa", "Come", "Commune de Come", "AQ6L6G"], ["commune.copargo", "da0fd4afd5fae4e6d465dcf80430ab4cdca8c794c08f5ba50937101b3a0d581f", "commune", "Donga", "Djougou-Ouaké-Copargo", "Copargo", "Commune de Copargo", "T6MJXK"], ["commune.cotonou", "1424133d45678cac495e537fb09009cf166e41d2263b6f1fffc25766a8fa6a96", "commune", "Littoral", "Cotonou (selon arrondissement)", "Cotonou", "Commune de Cotonou", "87AU5B"], ["commune.cove", "d1fdb37ee662e878269da166641257a90df991a9c9b9ad18f8152330f49fcb27", "commune", "Zou", "Covè-Zagnanado-Ouinhi", "Covè", "Commune de Covè", "HXTPDP"], ["commune.dangbo", "8e8a2835962425e10654098a3b1b82269856318c52bd3f547de94f039bd77a43", "commune", "Ouémé", "Adjohoun-Bonou-Dangbo", "Dangbo", "Commune de Dangbo", "FF5E8I"], ["commune.dassa", "217e4ba955e8e67d9037414b4561563a453cb679734f54b32edbad49434a8001", "commune", "Collines", "Dassa-Glazoué", "Dassa", "Commune de Dassa", "I49KQ7"], ["commune.djakotome", "d68d9ed42d63399b08818e8e9f222b79ba7c94d845bdda1c5ff4b52816353120", "commune", "Couffo", "Aplahoué-Djakotomey-Dogbo", "Djakotome", "Commune de Djakotome", "1N8MTZ"], ["commune.djidja", "b3b81874649fbbd7295928261957711b73150a4fbc12872495f919ba5cb1179a", "commune", "Zou", "Djidja-Abomey-Agbangnizoun", "Djidja", "Commune de Djidja", "X272HP"], ["commune.djougou", "bcff0f59a128a9d07f8c3d1b21b4cb3db07b86186dd184a627f05ccfb56c0702", "commune", "Donga", "Djougou-Ouaké-Copargo", "Djougou", "Commune de Djougou", "OEVB9O"], ["commune.dogbo", "b42ecb9a75dded2c4ea5801b5691c1fd13fd79bf4141cafed713e5a3c4ec716e", "commune", "Couffo", "Aplahoué-Djakotomey-Dogbo", "Dogbo", "Commune de Dogbo", "OAEDOE"], ["commune.glazoue", "b47030a4b5acda9da70876fe4d99560d48ab2d2b64b20bbae07c5190e70db46c", "commune", "Collines", "Dassa-Glazoué", "Glazoue", "Commune de Glazoue", "CVE6PR"], ["commune.gogounou", "c4d6ee6a21defb4b7137dfcf308e4c76e5fd6f800f634ce87210317f3fab3cae", "commune", "Alibori", "Kandi-Gogounou-Ségbana", "Gogounou", "Commune de Gogounou", "5N8I4P"], ["commune.grand-popo", "68d79705fc6a537cf35d01f37ddfdabc4ed3edea2bbed8ab8e13a72d9156ac96", "commune", "Mono", "Comè-Grand popo-Houèyogbé-Bopa", "Grand-Popo", "Commune de Grand-Popo", "40MGG1"], ["commune.houeyogbe", "fcdcd911da85e660cbde51bab134c360341d1eb9dde8e526e0c23de52b2db3c7", "commune", "Mono", "Comè-Grand popo-Houèyogbé-Bopa", "Houeyogbe", "Commune de Houeyogbe", "W103DG"], ["commune.ifangni", "13baeb9e855c8619c9f8fb0f62669300039b64edfe3ec21d5c441b78b6213a8d", "commune", "Plateau", "Sakété-Ifangni", "Ifangni", "Commune de Ifangni", "DZVGPM"], ["commune.kalale", "d0130d8faa5cf0ef55a593b3129d31b8bc2cc5162445f278472186a53558c3b3", "commune", "Borgou", "Nikki-Kalalé-Pèrèrè", "Kalale", "Commune de Kalale", "M82I1L"], ["commune.kandi", "c51b5b9aa915160f86bac765b854511b44107c66dc9248f72480c0693eb1740e", "commune", "Alibori", "Kandi-Gogounou-Ségbana", "Kandi", "Commune de Kandi", "R3PE29"], ["commune.karimama", "31f5988ee4b5adbd145697998528d2ea7d12a47a6078c7ce0aeecaafb5e389d6", "commune", "Alibori", "Malanville-Karimama", "Karimama", "Commune de Karimama", "GD8AFP"], ["commune.klouekanme", "4e98f63df583f879b5cfe26caa77e5ca5b286b296a80f45e267d670cdb386004", "commune", "Couffo", "Klouékamè-Toviklin-Lalo", "Klouékanmè", "Commune de Klouékanmè", "K054NZ"], ["commune.kouande", "727e16d3932295f1d2861e244d83299549889bbce3fc8582acb1a4bbbe174197", "commune", "Atacora", "Kouandé-Péhunco-Kérou", "Kouande", "Commune de Kouande", "DKYAYQ"], ["commune.kpomasse", "b78f30ba7897a273bcc7ad765c23f7b8da113c565adb46192fb1c30a3f18b3f2", "commune", "Atlantique", "Ouidah-Kpomassè-Tori-Bossito", "Kpomasse", "Commune de Kpomasse", "3S195J"], ["commune.kerou", "85da4bab11f0cc4cb707451654453e32e03d95f8ea1bbc0fe69fa41bca58e088", "commune", "Atacora", "Kouandé-Péhunco-Kérou", "Kérou", "Commune de Kérou", "MSND8D"], ["commune.ketou", "3dd48efd6d975809e6ea4ae2fdb9c1e99547313f6d2a35bacd8b384a7a351d2a", "commune", "Plateau", "Pobè-Kétou-Adja-ouèrè", "Kétou", "Commune de Kétou", "UDD467"], ["commune.lalo", "f39ade1adb0f0a12b0f27c8819c8af02506e8205a6bccfcb1f1fc42aba715e6c", "commune", "Couffo", "Klouékamè-Toviklin-Lalo", "Lalo", "Commune de Lalo", "KD6FLE"], ["commune.lokossa", "7507aa57de7b1e85ce11bed4a384d84cc149b5efc8b0540c6713274a4fff45af", "commune", "Mono", "Lokossa-Athiémè", "Lokossa", "Commune de Lokossa", "EPZHPC"], ["commune.malanville", "ccde329f40c8803d4107eb1c2b438894a8ef79475edc2f0f5866a55168e7b417", "commune", "Alibori", "Malanville-Karimama", "Malanville", "Commune de Malanville", "F07UQN"], ["commune.materi", "440ce3ca21d4e142c324e1e32c97514370ae804f7df9f24bf07b8abd1487973d", "commune", "Atacora", "Tanguiéta-Cobly-Matéri", "Matéri", "Commune de Matéri", "UPQZIT"], ["commune.n-dali", "023e6387ba852b63ad72289eb03774159d2b67a28455ef8d7e560e87c9d68d9d", "commune", "Borgou", "N'Dali-Parakou", "N'Dali", "Commune de N'Dali", "3UEA3G"], ["commune.natitingou", "b73409dbced8040e382bd54653cb3a773e5273988ab1de44381c868c2754cd41", "commune", "Atacora", "Natitingou-Boukoumbé-Toucountouna", "Natitingou", "Commune de Natitingou", "E8N6QI"], ["commune.nikki", "2530048625640dce9a35eda144f7f6d0eae51fda9679d7e874ea4dbb089caa83", "commune", "Borgou", "Nikki-Kalalé-Pèrèrè", "Nikki", "Commune de Nikki", "WEPXSK"], ["commune.ouake", "5c7c29972454415fecf6e71e378d802ed787cf6d7365d98c243049c24f90251d", "commune", "Donga", "Djougou-Ouaké-Copargo", "Ouake", "Commune de Ouake", "28T7A9"], ["commune.ouesse", "c6249ac98c2344afd374a6c24e91b33ea474493ab7cdb6a3a49244f4d94c0cd8", "commune", "Collines", "Savè-Ouèssè", "Ouesse", "Commune de Ouesse", "TGIQHG"], ["commune.ouidah", "52c2d5a1fc6146f8dafc70b14d3841f26a8e0dbea47a91c22b8e5a1e6d8f3bac", "commune", "Atlantique", "Ouidah-Kpomassè-Tori-Bossito", "Ouidah", "Commune de Ouidah", "9JRSNV"], ["commune.ouinhi", "3029b6f27c162b228c0407842f00320a903409ec34e463906ef0b96752bf30ab", "commune", "Zou", "Covè-Zagnanado-Ouinhi", "Ouinhi", "Commune de Ouinhi", "NQ65QD"], ["commune.parakou", "c4b4e168b540049321cba76c38a1d60af0785ab2af6b0653d80a86841e03397f", "commune", "Borgou", "N'Dali-Parakou", "Parakou", "Commune de Parakou", "F1RCAV"], ["commune.pehunco", "b5c8ec6966124ab161e8fc3355e055fc23add9c4a3b0f8b9a7b65155e9920ec1", "commune", "Atacora", "Kouandé-Péhunco-Kérou", "Pehunco", "Commune de Pehunco", "IQK291"], ["commune.pobe", "369463d13bc05c4f5399d8b4912e78d01aedbdec0b69bf86ecc3e5060cd27d21", "commune", "Plateau", "Pobè-Kétou-Adja-ouèrè", "Pobè", "Commune de Pobè", "9AHEJ8"], ["commune.porto-novo", "8b4e0147c57da021f62c4930bc1cdf707a7a4a14c80bb281e77ab94e6c382adb", "commune", "Ouémé", "Porto-Novo-Sèmè-kpodji-Aguégués", "Porto-Novo", "Commune de Porto-Novo", "CX9J1I"], ["commune.perere", "59559efabcffe9ab9a8a1a460183eee39cfde88ccf6015c1dc8f460e502b4707", "commune", "Borgou", "Nikki-Kalalé-Pèrèrè", "Pèrèrè", "Commune de Pèrèrè", "CTXCWN"], ["commune.sakete", "1511f0b0ec458cf406e8c49c112b10263b0ba62bef1912fa5942cdcf768acebc", "commune", "Plateau", "Sakété-Ifangni", "Sakete", "Commune de Sakete", "PGW90J"], ["commune.savalou", "e6a15d92aa1e1ddf9fa4a68c7727dfbb4c0f6ee4aaf9b042a029fa4a25e53abc", "commune", "Collines", "Savalou-Bantè", "Savalou", "Commune de Savalou", "PKL0BL"], ["commune.save", "e4a23eea5c4295bcc219f6f039af3c57f51367e118a996b2e98f9a1e2b10d5a6", "commune", "Collines", "Savè-Ouèssè", "Savè", "Commune de Savè", "V0PRKG"], ["commune.segbana", "ed443e38700417929ef058499ed8fe3d33fd2b23968ac8fa1e9ee41da0fb94f3", "commune", "Alibori", "Kandi-Gogounou-Ségbana", "Segbana", "Commune de Segbana", "YC4OM3"], ["commune.seme-kpodji", "22cda553b63577581a486f170da142c9e82a92e55e103c9d527c4315ec3f2143", "commune", "Ouémé", "Porto-Novo-Sèmè-kpodji-Aguégués", "Seme-Kpodji", "Commune de Seme-Kpodji", "WTOOBM"], ["commune.sinende", "1ea3b6bc61dfeb1a6e72c65c702bee7c145d55a2db2db6c5ea51b1157d5effdc", "commune", "Borgou", "Bembèrèkè-Sinendé", "Sinende", "Commune de Sinende", "ZVRERW"], ["commune.so-ava", "a314a7a0a7be2be458341dfda1416d6599ac0db0ae8d494aa36773f6a0cc47ab", "commune", "Atlantique", "Abomey-Calavi-So-ava", "So-Ava", "Commune de So-Ava", "6Z8VBH"], ["commune.tanguieta", "be13464c77d139e5e169858144bf320d47a346a9722d1130ec96b5a3d2d81a09", "commune", "Atacora", "Tanguiéta-Cobly-Matéri", "Tanguieta", "Commune de Tanguieta", "QLQCG1"], ["commune.tchaourou", "b25a259c08971eafc7e1649847a1fdfaafb186f0fa200666fc4c235d880696a9", "commune", "Borgou", "Tchaourou", "Tchaourou", "Commune de Tchaourou", "WU16HY"], ["commune.toffo", "e8f522cda3167b036485502d8c4f98b3c3efe4c678568252473296cd69c09a3e", "commune", "Atlantique", "Allada-Toffo-Zè", "Toffo", "Commune de Toffo", "MQC1A7"], ["commune.tori-bossito", "8390ea72e90b57dccfa0db907d763c14708431d17a2e93ffd973190142ca4fe4", "commune", "Atlantique", "Ouidah-Kpomassè-Tori-Bossito", "Tori-Bossito", "Commune de Tori-Bossito", "8MX1EV"], ["commune.toukountouna", "0e78ade9a57ca2ff2ff88898517cacd716181b980dbbc3e530b9a02ecdaa8a08", "commune", "Atacora", "Natitingou-Boukoumbé-Toucountouna", "Toukountouna", "Commune de Toukountouna", "UHT6T0"], ["commune.toviklin", "a5e71148f1c46c1b7f7aef04de1ff093c31324a50ce3e30af48f835994170190", "commune", "Couffo", "Klouékamè-Toviklin-Lalo", "Toviklin", "Commune de Toviklin", "UZS9IM"], ["commune.za-kpota", "86f07f3df78901ea6a759b46d365868ec2ec8e0d9183e8efd1ef553760463d59", "commune", "Zou", "Bohicon-Za-kpota-Zogbodomey", "Za-Kpota", "Commune de Za-Kpota", "0YLTZ9"], ["commune.zagnanado", "41d64ee5054c859b8f36aaf6c5b0453287838426fc6cf0c7f298b4bb09fba28a", "commune", "Zou", "Covè-Zagnanado-Ouinhi", "Zagnanado", "Commune de Zagnanado", "ATSN1U"], ["commune.ze", "29e3b6db9683dedec30ea094dfd2e25613478c129b26bafa77afa65dd3282e41", "commune", "Atlantique", "Allada-Toffo-Zè", "Ze", "Commune de Ze", "322N64"], ["commune.zogbodomey", "f2d17d2f0d6ea91c19ecf48e07bcbf8f5d6a30c2db58eedfc200a85cf24c4d6c", "commune", "Zou", "Bohicon-Za-kpota-Zogbodomey", "Zogbodomey", "Commune de Zogbodomey", "KFS6VF"], ["zone.abomey-calavi-so-ava", "93aff0044c197be5aec397b98865b002fe632f57b161b519c14502c71467c4d0", "zone", "Atlantique", "Abomey-Calavi-So-ava", "", "Zone sanitaire Abomey-Calavi-So-ava", "PTOMJB"], ["zone.adjohoun-bonou-dangbo", "b19194f2b30134091697940fc445563c0037ae2ee3e7050c4334e661a8a8ff4b", "zone", "Ouémé", "Adjohoun-Bonou-Dangbo", "", "Zone sanitaire Adjohoun-Bonou-Dangbo", "CP4E30"], ["zone.akpro-misserete-avrankou-adjarra", "6a7b8bab1a24b8d74524e45a2a304919042be73b72023282a01124fcb5b0ca14", "zone", "Ouémé", "Akpro-missérété-Avrankou-Adjarra", "", "Zone sanitaire Akpro-missérété-Avrankou-Adjarra", "MY5ZPJ"], ["zone.allada-toffo-ze", "896bda3c617196462053f123fa31eab38bec3526b43203d5380b76ac705a2c62", "zone", "Atlantique", "Allada-Toffo-Zè", "", "Zone sanitaire Allada-Toffo-Zè", "AG1OL7"], ["zone.aplahoue-djakotomey-dogbo", "60e1049db5f806911494c6d858edf85dc4085aa022b77d39b2d055dbdaf1d2bf", "zone", "Couffo", "Aplahoué-Djakotomey-Dogbo", "", "Zone sanitaire Aplahoué-Djakotomey-Dogbo", "3D9PH3"], ["zone.banikoara", "518ef5eba7b8edf42d71f8596f360c97017d8e3f6603aed8d81f56b671b4b64a", "zone", "Alibori", "Banikoara", "", "Zone sanitaire Banikoara", "I379U2"], ["zone.bassila", "c021fc9d9aaafe7fbb304d00a89fbdacb7b7c470406f91f4d89560c92cc987b3", "zone", "Donga", "Bassila", "", "Zone sanitaire Bassila", "6192K4"], ["zone.bembereke-sinende", "c3f4df49cd383b154e31c496829dee8e852fceddea2fb9b578d93b2217a089c2", "zone", "Borgou", "Bembèrèkè-Sinendé", "", "Zone sanitaire Bembèrèkè-Sinendé", "2QPR75"], ["zone.bohicon-za-kpota-zogbodomey", "357e25f03c22431981cf6a956e55455267ceb4ea2b8013af9b7eac8be7a8f798", "zone", "Zou", "Bohicon-Za-kpota-Zogbodomey", "", "Zone sanitaire Bohicon-Za-kpota-Zogbodomey", "PR2ESP"], ["zone.come-grand-popo-houeyogbe-bopa", "433ae08ebe50d887f8c967bf40ce97fb0b2dc4be3edd1e456969d03059c7035c", "zone", "Mono", "Comè-Grand popo-Houèyogbé-Bopa", "", "Zone sanitaire Comè-Grand popo-Houèyogbé-Bopa", "RVU8FI"], ["zone.cotonou-selon-arrondissement", "d6eb20226562a1b447e582b6a3a40f24781215b088bcc555d075f7fd6609b219", "zone", "Littoral", "Cotonou (selon arrondissement)", "", "Zone sanitaire Cotonou (selon arrondissement)", "JOYJNE"], ["zone.cove-zagnanado-ouinhi", "bc77af79cc01313839e286c1967c1141416fd2d836f6cdebc8513c5486bcfb8a", "zone", "Zou", "Covè-Zagnanado-Ouinhi", "", "Zone sanitaire Covè-Zagnanado-Ouinhi", "00V830"], ["zone.dassa-glazoue", "e1d04e3183e2d5d445b0492642596075628effb9b542c448ae5248d99476cbbd", "zone", "Collines", "Dassa-Glazoué", "", "Zone sanitaire Dassa-Glazoué", "DN0YBY"], ["zone.djidja-abomey-agbangnizoun", "6e3bf8ab46965d17faf2f234cc15ce48934e57ed57b517c3a9892382cce6c470", "zone", "Zou", "Djidja-Abomey-Agbangnizoun", "", "Zone sanitaire Djidja-Abomey-Agbangnizoun", "4AWTY0"], ["zone.djougou-ouake-copargo", "2d9fc3f7ffef8b3c2bff4390cf10250ad8dfc1746c37e2b5112852e0491e689c", "zone", "Donga", "Djougou-Ouaké-Copargo", "", "Zone sanitaire Djougou-Ouaké-Copargo", "88O5OR"], ["zone.kandi-gogounou-segbana", "836f4983d30441de95457a8273b2bbce665ca0575f94a039788169db6d9e0f80", "zone", "Alibori", "Kandi-Gogounou-Ségbana", "", "Zone sanitaire Kandi-Gogounou-Ségbana", "15BYVZ"], ["zone.klouekame-toviklin-lalo", "ec06eb6d05bad5c0924439db67f40373de6a25d0802d71717e6b0c3a8405709c", "zone", "Couffo", "Klouékamè-Toviklin-Lalo", "", "Zone sanitaire Klouékamè-Toviklin-Lalo", "K3I8BZ"], ["zone.kouande-pehunco-kerou", "d5d089e286f65953f66026bcf9d1d36c226bc5c5591c8aae05ad012bae7c5df6", "zone", "Atacora", "Kouandé-Péhunco-Kérou", "", "Zone sanitaire Kouandé-Péhunco-Kérou", "BF1I3L"], ["zone.lokossa-athieme", "326193009399c33d0d9cfdb7a66d79e6f1de7e1a2c36bbe165f150962c35d6fe", "zone", "Mono", "Lokossa-Athiémè", "", "Zone sanitaire Lokossa-Athiémè", "DQYUN3"], ["zone.malanville-karimama", "df4fe3eb0c190f3ca440c55ab0ec86be22f95a8e23cb2159195e38f4007fb521", "zone", "Alibori", "Malanville-Karimama", "", "Zone sanitaire Malanville-Karimama", "UVYR0Q"], ["zone.n-dali-parakou", "5eeca57bf74c07fd49ee1f77c31494f4996a573c2dee8d9e033917fff06a4193", "zone", "Borgou", "N'Dali-Parakou", "", "Zone sanitaire N'Dali-Parakou", "F4B8DW"], ["zone.natitingou-boukoumbe-toucountouna", "807e7617d35863cc3c20bdeb35bff16d64167cc8afa22ec51e6581cfa07d946d", "zone", "Atacora", "Natitingou-Boukoumbé-Toucountouna", "", "Zone sanitaire Natitingou-Boukoumbé-Toucountouna", "OECBPM"], ["zone.nikki-kalale-perere", "47c490fde39da6c1390adcd9e05f8644f5fc8f642394384ee431f67485b1395a", "zone", "Borgou", "Nikki-Kalalé-Pèrèrè", "", "Zone sanitaire Nikki-Kalalé-Pèrèrè", "BJPI4H"], ["zone.ouidah-kpomasse-tori-bossito", "f77215277c860c9d05dcbc9d62ef530ddfe75527f358f71a1608c17ab615fe0c", "zone", "Atlantique", "Ouidah-Kpomassè-Tori-Bossito", "", "Zone sanitaire Ouidah-Kpomassè-Tori-Bossito", "N3QXKH"], ["zone.pobe-ketou-adja-ouere", "8b58c53c982952e1b909e4a56b73447f288c8bcde2337d0f422140e68a8ad9a7", "zone", "Plateau", "Pobè-Kétou-Adja-ouèrè", "", "Zone sanitaire Pobè-Kétou-Adja-ouèrè", "KTGBTY"], ["zone.porto-novo-seme-kpodji-aguegues", "67ca23d4dceb857e4cf1bd95ccbcfcb91548e51a4ee6f1cb39df4232a5e90081", "zone", "Ouémé", "Porto-Novo-Sèmè-kpodji-Aguégués", "", "Zone sanitaire Porto-Novo-Sèmè-kpodji-Aguégués", "ZMEPGT"], ["zone.sakete-ifangni", "a7fe18c0177dee50ff5eca51882a3f22f448e04051f8e2272ff4d4f916aee147", "zone", "Plateau", "Sakété-Ifangni", "", "Zone sanitaire Sakété-Ifangni", "HCW81X"], ["zone.savalou-bante", "7a12f2642b7a9a2b4e2e02a69284701a00a754a0075e37bebd51b232f1a7ea7d", "zone", "Collines", "Savalou-Bantè", "", "Zone sanitaire Savalou-Bantè", "E6VA05"], ["zone.save-ouesse", "dc271f7e793e58a29dd733ec52d1ef2018ee6e948a9ccca35bb93edd8be4fcbb", "zone", "Collines", "Savè-Ouèssè", "", "Zone sanitaire Savè-Ouèssè", "G1X3J1"], ["zone.tanguieta-cobly-materi", "3a32d6e68e9bfda67f79f9a366fc1dd8abe17a9eb7b501023ee6242871f9b2d5", "zone", "Atacora", "Tanguiéta-Cobly-Matéri", "", "Zone sanitaire Tanguiéta-Cobly-Matéri", "L7R843"], ["zone.tchaourou", "29402ee46aa9614c852c8220142314727ffede4957f00924d2854f05b75c352d", "zone", "Borgou", "Tchaourou", "", "Zone sanitaire Tchaourou", "1RUPFR"], ["dept.alibori", "9315781879a48242554146313ab5d7551424f2060fb32e9fb7b91fb498094b48", "departement", "Alibori", "", "", "Département Alibori", "2P3YVB"], ["dept.atacora", "68993e04040c15c214e501d8767fb7ce890697864ebfb7cfc96a387c214501cb", "departement", "Atacora", "", "", "Département Atacora", "5UL5NW"], ["dept.atlantique", "97179b1de356910a515e5faadfd32f8fe1f1559e26b6ad053fc60150907c7ec4", "departement", "Atlantique", "", "", "Département Atlantique", "QVRR9A"], ["dept.borgou", "cefd14c23754990513a1190b9e14900912832415b9059fd804497d6f5e117716", "departement", "Borgou", "", "", "Département Borgou", "7MFP05"], ["dept.collines", "83d003f5016f88cd6c91787368ef0c6dfbfde92e560de4f1d007c47870d16cf2", "departement", "Collines", "", "", "Département Collines", "9P452B"], ["dept.couffo", "4fb61224f3a1978aa34fa25b43efb66fd94a74198b17c9d3133aada19c885c79", "departement", "Couffo", "", "", "Département Couffo", "FSOZPT"], ["dept.donga", "fd482396087603bd80a6810c1d3e7bbdd155e3e8a38c1fe51a0b296db88f8f87", "departement", "Donga", "", "", "Département Donga", "X497W1"], ["dept.littoral", "ee950d9b1cdfc4d8b8faef4d5237fd6632bc5a6dfa5467cb782fd08fa1ca512d", "departement", "Littoral", "", "", "Département Littoral", "9VW3RT"], ["dept.mono", "d0cf840bb3ce6a1bd318cb852e850eef6811cd10f99b8bf5b3601119773663f7", "departement", "Mono", "", "", "Département Mono", "QOHMUH"], ["dept.oueme", "767da30c9bdfd6ac2b34a93b73979b74a36f81927d7fd380d2fcf9e5d42b47d1", "departement", "Ouémé", "", "", "Département Ouémé", "8LMN4R"], ["dept.plateau", "99f9b471fb528d9a787a58c40224628814cfca7c7576a951d497f691b3ed2ee5", "departement", "Plateau", "", "", "Département Plateau", "7SGMSO"], ["dept.zou", "1b31f34cc96961d1b2bf6113dc84bdfb298152ca61ed658e813b772a15555786", "departement", "Zou", "", "", "Département Zou", "XLTA8I"], ["national", "8c8092777e3f52626b12253e4acccce6f2472ceba90db219fb53c7bcccac99dd", "national", "", "", "", "Niveau National", "RCD9SI5G"], ["admin", "9f406a016d50635ba113c8320322d1e41ed519407d8be0ea247ea43a4d581533", "admin", "", "", "", "Administrateur", "AS442VLD"], ["mourede", "067ac18745bde46d2fcf358adcd1ea9ca17cdb9e8ac3ce45063b87a6fde60b14", "superviseur", "", "", "", "Mourede", "0IBYYURL"]];

// Commune -> {departement, zone} — découpage administratif officiel du Bénin
var REFERENTIEL_COMMUNES = {"Banikoara": {"departement": "Alibori", "zone": "Banikoara"}, "Gogounou": {"departement": "Alibori", "zone": "Kandi-Gogounou-Ségbana"}, "Kandi": {"departement": "Alibori", "zone": "Kandi-Gogounou-Ségbana"}, "Segbana": {"departement": "Alibori", "zone": "Kandi-Gogounou-Ségbana"}, "Karimama": {"departement": "Alibori", "zone": "Malanville-Karimama"}, "Malanville": {"departement": "Alibori", "zone": "Malanville-Karimama"}, "Kouande": {"departement": "Atacora", "zone": "Kouandé-Péhunco-Kérou"}, "Pehunco": {"departement": "Atacora", "zone": "Kouandé-Péhunco-Kérou"}, "Kérou": {"departement": "Atacora", "zone": "Kouandé-Péhunco-Kérou"}, "Natitingou": {"departement": "Atacora", "zone": "Natitingou-Boukoumbé-Toucountouna"}, "Boukoumbe": {"departement": "Atacora", "zone": "Natitingou-Boukoumbé-Toucountouna"}, "Toukountouna": {"departement": "Atacora", "zone": "Natitingou-Boukoumbé-Toucountouna"}, "Tanguieta": {"departement": "Atacora", "zone": "Tanguiéta-Cobly-Matéri"}, "Cobly": {"departement": "Atacora", "zone": "Tanguiéta-Cobly-Matéri"}, "Matéri": {"departement": "Atacora", "zone": "Tanguiéta-Cobly-Matéri"}, "Abomey-Calavi": {"departement": "Atlantique", "zone": "Abomey-Calavi-So-ava"}, "So-Ava": {"departement": "Atlantique", "zone": "Abomey-Calavi-So-ava"}, "Allada": {"departement": "Atlantique", "zone": "Allada-Toffo-Zè"}, "Toffo": {"departement": "Atlantique", "zone": "Allada-Toffo-Zè"}, "Ze": {"departement": "Atlantique", "zone": "Allada-Toffo-Zè"}, "Ouidah": {"departement": "Atlantique", "zone": "Ouidah-Kpomassè-Tori-Bossito"}, "Kpomasse": {"departement": "Atlantique", "zone": "Ouidah-Kpomassè-Tori-Bossito"}, "Tori-Bossito": {"departement": "Atlantique", "zone": "Ouidah-Kpomassè-Tori-Bossito"}, "Bembèrèkè": {"departement": "Borgou", "zone": "Bembèrèkè-Sinendé"}, "Sinende": {"departement": "Borgou", "zone": "Bembèrèkè-Sinendé"}, "N'Dali": {"departement": "Borgou", "zone": "N'Dali-Parakou"}, "Parakou": {"departement": "Borgou", "zone": "N'Dali-Parakou"}, "Nikki": {"departement": "Borgou", "zone": "Nikki-Kalalé-Pèrèrè"}, "Kalale": {"departement": "Borgou", "zone": "Nikki-Kalalé-Pèrèrè"}, "Pèrèrè": {"departement": "Borgou", "zone": "Nikki-Kalalé-Pèrèrè"}, "Tchaourou": {"departement": "Borgou", "zone": "Tchaourou"}, "Dassa": {"departement": "Collines", "zone": "Dassa-Glazoué"}, "Glazoue": {"departement": "Collines", "zone": "Dassa-Glazoué"}, "Savalou": {"departement": "Collines", "zone": "Savalou-Bantè"}, "Bante": {"departement": "Collines", "zone": "Savalou-Bantè"}, "Savè": {"departement": "Collines", "zone": "Savè-Ouèssè"}, "Ouesse": {"departement": "Collines", "zone": "Savè-Ouèssè"}, "Aplahoué": {"departement": "Couffo", "zone": "Aplahoué-Djakotomey-Dogbo"}, "Djakotome": {"departement": "Couffo", "zone": "Aplahoué-Djakotomey-Dogbo"}, "Dogbo": {"departement": "Couffo", "zone": "Aplahoué-Djakotomey-Dogbo"}, "Klouékanmè": {"departement": "Couffo", "zone": "Klouékamè-Toviklin-Lalo"}, "Toviklin": {"departement": "Couffo", "zone": "Klouékamè-Toviklin-Lalo"}, "Lalo": {"departement": "Couffo", "zone": "Klouékamè-Toviklin-Lalo"}, "Bassila": {"departement": "Donga", "zone": "Bassila"}, "Djougou": {"departement": "Donga", "zone": "Djougou-Ouaké-Copargo"}, "Ouake": {"departement": "Donga", "zone": "Djougou-Ouaké-Copargo"}, "Copargo": {"departement": "Donga", "zone": "Djougou-Ouaké-Copargo"}, "Cotonou": {"departement": "Littoral", "zone": "Cotonou (selon arrondissement)"}, "Athieme": {"departement": "Mono", "zone": "Lokossa-Athiémè"}, "Lokossa": {"departement": "Mono", "zone": "Lokossa-Athiémè"}, "Bopa": {"departement": "Mono", "zone": "Comè-Grand popo-Houèyogbé-Bopa"}, "Come": {"departement": "Mono", "zone": "Comè-Grand popo-Houèyogbé-Bopa"}, "Grand-Popo": {"departement": "Mono", "zone": "Comè-Grand popo-Houèyogbé-Bopa"}, "Houeyogbe": {"departement": "Mono", "zone": "Comè-Grand popo-Houèyogbé-Bopa"}, "Adjarra": {"departement": "Ouémé", "zone": "Akpro-missérété-Avrankou-Adjarra"}, "Akpro-Missereté": {"departement": "Ouémé", "zone": "Akpro-missérété-Avrankou-Adjarra"}, "Avrankou": {"departement": "Ouémé", "zone": "Akpro-missérété-Avrankou-Adjarra"}, "Adjohoun": {"departement": "Ouémé", "zone": "Adjohoun-Bonou-Dangbo"}, "Bonou": {"departement": "Ouémé", "zone": "Adjohoun-Bonou-Dangbo"}, "Dangbo": {"departement": "Ouémé", "zone": "Adjohoun-Bonou-Dangbo"}, "Aguegues": {"departement": "Ouémé", "zone": "Porto-Novo-Sèmè-kpodji-Aguégués"}, "Porto-Novo": {"departement": "Ouémé", "zone": "Porto-Novo-Sèmè-kpodji-Aguégués"}, "Seme-Kpodji": {"departement": "Ouémé", "zone": "Porto-Novo-Sèmè-kpodji-Aguégués"}, "Adja-Ouèrè": {"departement": "Plateau", "zone": "Pobè-Kétou-Adja-ouèrè"}, "Kétou": {"departement": "Plateau", "zone": "Pobè-Kétou-Adja-ouèrè"}, "Pobè": {"departement": "Plateau", "zone": "Pobè-Kétou-Adja-ouèrè"}, "Ifangni": {"departement": "Plateau", "zone": "Sakété-Ifangni"}, "Sakete": {"departement": "Plateau", "zone": "Sakété-Ifangni"}, "Abomey": {"departement": "Zou", "zone": "Djidja-Abomey-Agbangnizoun"}, "Agbangnizoun": {"departement": "Zou", "zone": "Djidja-Abomey-Agbangnizoun"}, "Djidja": {"departement": "Zou", "zone": "Djidja-Abomey-Agbangnizoun"}, "Bohicon": {"departement": "Zou", "zone": "Bohicon-Za-kpota-Zogbodomey"}, "Za-Kpota": {"departement": "Zou", "zone": "Bohicon-Za-kpota-Zogbodomey"}, "Zogbodomey": {"departement": "Zou", "zone": "Bohicon-Za-kpota-Zogbodomey"}, "Covè": {"departement": "Zou", "zone": "Covè-Zagnanado-Ouinhi"}, "Ouinhi": {"departement": "Zou", "zone": "Covè-Zagnanado-Ouinhi"}, "Zagnanado": {"departement": "Zou", "zone": "Covè-Zagnanado-Ouinhi"}};
// Formation sanitaire -> Commune (déduit des listes Commune/Arrondissement/Structure des formulaires)
var STRUCTURE_COMMUNE = {"CS Agbokpa": "Abomey", "CS Detohou": "Abomey", "CS Tangandji": "Abomey", "CS Djegbe (Abomey)": "Abomey", "CS Hounli": "Abomey", "CS Sehoun": "Abomey", "CS Adandokpodji": "Abomey", "CS Vidole": "Abomey", "CS Zounzonme": "Abomey", "CS Adanhondjigon": "Agbangnizoun", "CS Adingnigon": "Agbangnizoun", "CS Agbangnizoun": "Agbangnizoun", "CS Kinta": "Agbangnizoun", "CS Akodébakou": "Agbangnizoun", "MI Sinwe-Kpota": "Agbangnizoun", "CS Lissazounme": "Agbangnizoun", "CS Dovota": "Agbangnizoun", "CS Sahè": "Agbangnizoun", "CS Sinwe": "Agbangnizoun", "CS Tanve": "Agbangnizoun", "CS Kpoto": "Agbangnizoun", "CS Agondji": "Djidja", "CS Agouna": "Djidja", "CS Dan": "Djidja", "CS Djidja": "Djidja", "CS Dohouime": "Djidja", "CS Gobaix": "Djidja", "CS Outo": "Djidja", "CS Lobeta": "Djidja", "CS Monsourou": "Djidja", "CS Mougnon": "Djidja", "CS Oumgbega": "Djidja", "CS Saloudji": "Djidja", "CS Setto": "Djidja", "CS Ayogbé": "Djidja", "CS Sodohomè": "Bohicon", "CS Saclo": "Bohicon", "CS Gnidjazoun": "Bohicon", "CS Bohicon 2": "Bohicon", "Cs Lissèzoun": "Bohicon", "CS Ouassaho": "Bohicon", "CS Bohicon 1": "Bohicon", "CS Passagon": "Bohicon", "CS Avogbanna": "Bohicon", "CS Agongointo": "Bohicon", "CS Adjinagon": "Za-Kpota", "CS Za-kpota": "Za-Kpota", "CS Allahé": "Za-Kpota", "MI Za-Hla": "Za-Kpota", "CS Za-Tanta": "Za-Kpota", "CS Kpakpamè": "Za-Kpota", "CS Kpozoun": "Za-Kpota", "CS Houngomè": "Za-Kpota", "CS Assanlin": "Za-Kpota", "CS kpolokoué": "Za-Kpota", "CS zounzounmè": "Za-Kpota", "CS Tindji": "Za-Kpota", "CS Deme": "Zogbodomey", "CS Koussoukpa": "Zogbodomey", "CS Kpokissa": "Zogbodomey", "CS Massi": "Zogbodomey", "CS Cana 2": "Zogbodomey", "CS Zogbodome": "Zogbodomey", "MI Zoungbo-Bogon": "Zogbodomey", "CS Tanwé hessou": "Zogbodomey", "CS Akiza": "Zogbodomey", "CS Avlamè": "Zogbodomey", "CS Domè": "Zogbodomey", "CS Zoukou": "Zogbodomey", "CS Houin-Hounso": "Covè", "CS Covè": "Covè", "CS Lainta Cogbé": "Covè", "CS Naogon": "Covè", "CS Agonlin-Houegbo": "Zagnanado", "CS Baname": "Zagnanado", "CS Zakoumado": "Zagnanado", "MI Gossoé": "Zagnanado", "CS Don-Tan": "Zagnanado", "CS Dovi-Dove": "Zagnanado", "MI Dovi Zounnou": "Zagnanado", "CS Kpedekpo": "Zagnanado", "CS Zagnannado": "Zagnanado", "CS Dasso": "Ouinhi", "CS Ouinhi": "Ouinhi", "CS Aïzè": "Ouinhi", "CS Sagon": "Ouinhi", "MI Houédja": "Ouinhi", "MI Tévêdji": "Ouinhi", "CS Tohoues": "Ouinhi", "MI Akassa": "Ouinhi"};

function sha256Hex_(str){
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return raw.map(function(b){ return ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0'); }).join('');
}

function getOrCreateSheet_(name, headers){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

/** À exécuter UNE FOIS depuis l'éditeur Apps Script (menu ▶ Exécuter). */
function setupComptesEtQrTokens(){
  var comptes = getOrCreateSheet_(ACCOUNTS_SHEET_NAME,
    ['Identifiant','Hash','Niveau','Département','Zone Sanitaire','Commune','Libellé','DernièreModification','Actif','Mot de passe']);
  if (comptes.getLastRow() < 2){
    SEED_ACCOUNTS.forEach(function(r){ comptes.appendRow([r[0], r[1], r[2], r[3], r[4], r[5], r[6], '', 'OUI', r[7] || '']); });
  }
  getOrCreateSheet_(SESSIONS_SHEET_NAME, ['Token','Identifiant','Créée le']);
  var qr = getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle','Token','GénéréLe']);
  if (qr.getLastRow() < 2){
    qr.appendRow(['enqueteur', Utilities.getUuid(), new Date().toISOString()]);
    qr.appendRow(['personnel', Utilities.getUuid(), new Date().toISOString()]);
  }
  Logger.log('Comptes, Sessions et QRTokens prêts.');
}

/** Si vous avez DÉJÀ exécuté setupComptesEtQrTokens avant cette mise à jour (donc
 *  sans la colonne Actif), exécutez cette fonction UNE FOIS pour l'ajouter aux
 *  comptes existants sans les recréer. */
function migrationAjouterColonneActif(){
  var comptes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!comptes) return;
  var headers = comptes.getRange(1,1,1,comptes.getLastColumn()).getValues()[0];
  if (headers.indexOf('Actif') !== -1) { Logger.log('Colonne Actif déjà présente.'); return; }
  var col = comptes.getLastColumn() + 1;
  comptes.getRange(1, col).setValue('Actif');
  var lastRow = comptes.getLastRow();
  if (lastRow >= 2) comptes.getRange(2, col, lastRow - 1, 1).setValue('OUI');
  Logger.log('Colonne Actif ajoutée.');
}

/** Si vous avez DÉJÀ exécuté setupComptesEtQrTokens avant l'ajout de la colonne
 *  Mot de passe, exécutez cette fonction UNE FOIS. Elle ajoute la colonne et la
 *  remplit avec les mots de passe d'origine (ceux du CSV distribué) pour les
 *  comptes qui n'ont pas encore été changés — un compte déjà modifié par son
 *  utilisateur ne peut pas être reconstitué en clair (seul le hash est connu),
 *  sa cellule restera vide jusqu'au prochain changement de mot de passe. */
function migrationAjouterColonneMotDePasse(){
  var comptes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!comptes) return;
  var headers = comptes.getRange(1,1,1,comptes.getLastColumn()).getValues()[0];
  var col = headers.indexOf('Mot de passe');
  if (col === -1){
    col = comptes.getLastColumn() + 1;
    comptes.getRange(1, col).setValue('Mot de passe');
  } else {
    col = col + 1; // getRange est en base 1
  }
  var values = comptes.getDataRange().getValues();
  var byUser = {};
  SEED_ACCOUNTS.forEach(function(r){ byUser[r[0]] = r[7] || ''; });
  for (var i = 1; i < values.length; i++){
    var username = values[i][0];
    var currentHash = values[i][1];
    var seedRow = SEED_ACCOUNTS.filter(function(r){ return r[0] === username; })[0];
    // Ne renseigne le mot de passe d'origine que si le hash n'a pas changé
    // depuis (sinon on afficherait un mot de passe qui n'est plus le bon).
    if (seedRow && seedRow[1] === currentHash && byUser[username]){
      comptes.getRange(i + 1, col).setValue(byUser[username]);
    }
  }
  Logger.log('Colonne Mot de passe ajoutée / mise à jour.');
}

function findAccountRow_(username){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++){
    if (values[i][0] === username) return { row: i + 1, data: values[i] };
  }
  return null;
}

function accountPublicView_(data){
  return { username: data[0], level: data[2], departement: data[3], zone: data[4], commune: data[5], label: data[6],
    active: (data[8] === undefined || data[8] === '' || data[8] === 'OUI') };
}

function createSession_(username){
  var token = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET_NAME).appendRow([token, username, new Date().toISOString()]);
  return token;
}

function findSession_(token){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET_NAME);
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++){
    if (values[i][0] === token){
      var created = new Date(values[i][2]).getTime();
      if (Date.now() - created > SESSION_DURATION_MS) return null;
      return { row: i + 1, username: values[i][1] };
    }
  }
  return null;
}

/** action: accountLogin — data:{username}, password: le mot de passe personnel */
function handleAccountLogin(username, password){
  if (!username || !password) return { ok:false, error:'Identifiant et mot de passe requis.' };
  var found = findAccountRow_(username);
  if (!found) return { ok:false, error:'Identifiant inconnu.' };
  if (found.data[8] === 'NON') return { ok:false, error:'Ce compte a été désactivé par un administrateur.' };
  var hash = sha256Hex_(username + ':' + password);
  if (hash !== found.data[1]) return { ok:false, error:'Mot de passe incorrect.' };
  var token = createSession_(username);
  return { ok:true, token: token, account: accountPublicView_(found.data) };
}

/** Vrai si le token correspond à une session valide de niveau admin ou national */
function requireAdminSession_(token){
  var session = token ? findSession_(token) : null;
  if (!session) return { ok:false, error:'Session invalide ou expirée, reconnectez-vous.' };
  var acc = findAccountRow_(session.username);
  if (!acc || acc.data[2] !== 'admin'){
    return { ok:false, error:'Réservé au compte administrateur.' };
  }
  return { ok:true, acc: acc };
}

/** action: listAccounts — data:{token}. Ne renvoie jamais le hash. */
function handleListAccounts(token){
  var check = requireAdminSession_(token);
  if (!check.ok) return check;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++){
    var r = values[i];
    out.push({
      username: r[0], level: r[2], departement: r[3], zone: r[4], commune: r[5],
      label: r[6], lastModified: r[7] || '', active: (r[8] === undefined || r[8] === '' || r[8] === 'OUI')
    });
  }
  return { ok:true, accounts: out };
}

/** action: setAccountActive — data:{token, username, active} */
function handleSetAccountActive(token, username, active){
  var check = requireAdminSession_(token);
  if (!check.ok) return check;
  var found = findAccountRow_(username);
  if (!found) return { ok:false, error:'Compte introuvable.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  sheet.getRange(found.row, 9).setValue(active ? 'OUI' : 'NON');
  return { ok:true };
}

/** action: updateAccount — data:{token, username, fields:{level,departement,zone,commune,label}} */
function handleUpdateAccount(token, username, fields){
  var check = requireAdminSession_(token);
  if (!check.ok) return check;
  var found = findAccountRow_(username);
  if (!found) return { ok:false, error:'Compte introuvable.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  fields = fields || {};
  if (fields.level !== undefined) sheet.getRange(found.row, 3).setValue(fields.level);
  if (fields.departement !== undefined) sheet.getRange(found.row, 4).setValue(fields.departement);
  if (fields.zone !== undefined) sheet.getRange(found.row, 5).setValue(fields.zone);
  if (fields.commune !== undefined) sheet.getRange(found.row, 6).setValue(fields.commune);
  if (fields.label !== undefined) sheet.getRange(found.row, 7).setValue(fields.label);
  sheet.getRange(found.row, 8).setValue(new Date().toISOString());
  return { ok:true };
}

/** action: deleteAccount — data:{token, username} */
function handleDeleteAccount(token, username){
  var check = requireAdminSession_(token);
  if (!check.ok) return check;
  if (username === check.acc.data[0]) return { ok:false, error:'Vous ne pouvez pas supprimer votre propre compte connecté.' };
  var found = findAccountRow_(username);
  if (!found) return { ok:false, error:'Compte introuvable.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  sheet.deleteRow(found.row);
  return { ok:true };
}

/** action: changePassword — data:{token, newPassword}, password: l'ANCIEN mot de passe personnel */
function handleChangePassword(token, oldPassword, newPassword){
  var session = token ? findSession_(token) : null;
  if (!session) return { ok:false, error:'Session invalide ou expirée, reconnectez-vous.' };
  if (!newPassword || String(newPassword).length < 6){
    return { ok:false, error:'Le nouveau mot de passe doit contenir au moins 6 caractères.' };
  }
  var found = findAccountRow_(session.username);
  if (!found) return { ok:false, error:'Compte introuvable.' };
  var oldHash = sha256Hex_(session.username + ':' + oldPassword);
  if (oldHash !== found.data[1]) return { ok:false, error:'Ancien mot de passe incorrect.' };
  var newHash = sha256Hex_(session.username + ':' + newPassword);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  sheet.getRange(found.row, 2).setValue(newHash);
  sheet.getRange(found.row, 8).setValue(new Date().toISOString());
  sheet.getRange(found.row, 10).setValue(newPassword); // colonne "Mot de passe" (toujours à jour)
  return { ok:true };
}

/** Supprime toutes les sessions actives d'un identifiant (utilisé après un
 *  changement d'identifiant, pour ne pas laisser une session ouverte sous
 *  l'ancien nom traîner inutilement). */
function invalidateSessionsForUser_(username){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET_NAME);
  if (!sheet) return;
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--){
    if (values[i][1] === username) sheet.deleteRow(i + 1);
  }
}

/** action: changeUsername — data:{token, newUsername}, password: le mot de passe
 *  ACTUEL (pour vérifier l'identité avant de changer l'identifiant). */
function handleChangeUsername(token, currentPassword, newUsername){
  var session = token ? findSession_(token) : null;
  if (!session) return { ok:false, error:'Session invalide ou expirée, reconnectez-vous.' };
  var found = findAccountRow_(session.username);
  if (!found) return { ok:false, error:'Compte introuvable.' };
  var oldHash = sha256Hex_(session.username + ':' + currentPassword);
  if (oldHash !== found.data[1]) return { ok:false, error:'Mot de passe incorrect.' };

  newUsername = (newUsername || '').toString().trim();
  if (!newUsername) return { ok:false, error:'Identifiant vide.' };
  if (newUsername === session.username) return { ok:false, error:"C'est déjà votre identifiant actuel." };
  if (findAccountRow_(newUsername)) return { ok:false, error:'Cet identifiant est déjà utilisé par un autre compte.' };

  // Le mot de passe ne change pas, mais le hash dépend de l'identifiant : on le recalcule.
  var newHash = sha256Hex_(newUsername + ':' + currentPassword);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  sheet.getRange(found.row, 1).setValue(newUsername);
  sheet.getRange(found.row, 2).setValue(newHash);
  sheet.getRange(found.row, 8).setValue(new Date().toISOString());

  invalidateSessionsForUser_(session.username);
  var newToken = createSession_(newUsername);
  var updatedData = found.data.slice();
  updatedData[0] = newUsername;
  return { ok:true, token: newToken, account: accountPublicView_(updatedData) };
}

function findQrRow_(role){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QRTOKENS_SHEET_NAME);
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++){
    if (values[i][0] === role) return { row: i + 1, token: values[i][1] };
  }
  return null;
}

/** action: getQrToken — data:{role}. Public (pas de session requise) : le token n'ouvre
 *  l'accès qu'à un formulaire déjà public, il ne protège pas de donnée sensible. */
function handleGetQrToken(role){
  if (role !== 'enqueteur' && role !== 'personnel') return { ok:false, error:'Rôle inconnu.' };
  var row = findQrRow_(role);
  if (!row){
    var sheet = getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle','Token','GénéréLe']);
    var token = Utilities.getUuid();
    sheet.appendRow([role, token, new Date().toISOString()]);
    return { ok:true, token: token };
  }
  return { ok:true, token: row.token };
}

/** action: refreshQrToken — data:{token, role}. Réservé aux comptes national/admin. */
function handleRefreshQrToken(sessionToken, role){
  var session = sessionToken ? findSession_(sessionToken) : null;
  if (!session) return { ok:false, error:'Session invalide ou expirée, reconnectez-vous.' };
  var acc = findAccountRow_(session.username);
  if (!acc || acc.data[2] !== 'admin'){
    return { ok:false, error:"Seul le compte administrateur peut régénérer un QR code." };
  }
  if (role !== 'enqueteur' && role !== 'personnel') return { ok:false, error:'Rôle inconnu.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QRTOKENS_SHEET_NAME);
  var newToken = Utilities.getUuid();
  var row = findQrRow_(role);
  if (row){
    sheet.getRange(row.row, 2).setValue(newToken);
    sheet.getRange(row.row, 3).setValue(new Date().toISOString());
  } else {
    sheet.appendRow([role, newToken, new Date().toISOString()]);
  }
  return { ok:true, token: newToken };
}

/** action: setQrCode — data:{token, role, code}. Réservé aux comptes national/admin.
 *  Permet d'attribuer à une cible (enqueteur/personnel) un code qui vient d'être
 *  scanné depuis un QR EXTERNE (que l'application n'a pas généré elle-même) : ce
 *  code devient la nouvelle valeur attendue pour accorder l'accès à cette cible,
 *  à la place du code auto-généré. */
function handleSetQrCode(sessionToken, role, code){
  var session = sessionToken ? findSession_(sessionToken) : null;
  if (!session) return { ok:false, error:'Session invalide ou expirée, reconnectez-vous.' };
  var acc = findAccountRow_(session.username);
  if (!acc || acc.data[2] !== 'admin'){
    return { ok:false, error:"Seul le compte administrateur peut définir un code d'accès." };
  }
  if (role !== 'enqueteur' && role !== 'personnel') return { ok:false, error:'Rôle inconnu.' };
  code = (code || '').toString().trim();
  if (!code) return { ok:false, error:'Code vide ou non reconnu.' };
  var sheet = getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle','Token','GénéréLe']);
  var row = findQrRow_(role);
  if (row){
    sheet.getRange(row.row, 2).setNumberFormat('@').setValue(code);
    sheet.getRange(row.row, 3).setValue(new Date().toISOString());
  } else {
    sheet.appendRow([role, code, new Date().toISOString()]);
    sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@').setValue(code);
  }
  return { ok:true, token: code };
}

/** action: matchQrCode — data:{text}. Public (comme getQrToken). Reçoit le texte
 *  brut décodé d'un QR (qui peut être une URL générée par cette application, OU
 *  un code totalement externe / non généré par elle), et indique à quelle cible
 *  (le cas échéant) il correspond actuellement. Aucune correspondance = pas d'accès. */
function handleMatchQrCode(rawText){
  rawText = (rawText || '').toString().trim();
  if (!rawText) return { ok:false, error:'QR code vide.' };

  var candidate = rawText;
  var url = URL_(rawText);
  if (url) candidate = url.get('code') || url.get('t') || rawText;

  var rowEnqueteur = findQrRow_('enqueteur');
  var rowPersonnel = findQrRow_('personnel');
  var candidateStr = String(candidate).trim();
  if (rowEnqueteur && String(rowEnqueteur.token).trim() === candidateStr) return { ok:true, role:'enqueteur' };
  if (rowPersonnel && String(rowPersonnel.token).trim() === candidateStr) return { ok:true, role:'personnel' };
  return { ok:false, error:"Ce code ne correspond à aucune cible active." };
}

/** Mini-extracteur de paramètres de requête à partir d'une chaîne d'URL, sans
 *  dépendre de l'objet URL du navigateur (indisponible côté Apps Script). */
function URL_(text){
  var qIndex = text.indexOf('?');
  if (text.indexOf('http') !== 0 || qIndex === -1) return null;
  var query = text.substring(qIndex + 1);
  var params = {};
  query.split('&').forEach(function(pair){
    var kv = pair.split('=');
    if (kv[0]) params[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1] || '');
  });
  return { get: function(k){ return params[k] !== undefined ? params[k] : null; } };
}

/* ============================================================
 *  FILTRAGE SERVEUR DE getDashboardData SELON LA PORTÉE DU COMPTE
 * ============================================================ */
function communeInScope_(communeLabel, account){
  var lvl = account.level;
  if (lvl === 'admin' || lvl === 'national') return true;
  var info = REFERENTIEL_COMMUNES[communeLabel];
  if (!info) return false;
  if (lvl === 'departement') return info.departement === account.departement;
  if (lvl === 'zone') return info.zone === account.zone;
  if (lvl === 'commune') return communeLabel === account.commune;
  return true;
}
function structureInScope_(structureLabel, account){
  var commune = STRUCTURE_COMMUNE[structureLabel];
  if (!commune) return true; // structure inconnue de la table : on ne bloque pas
  return communeInScope_(commune, account);
}

/** Recalcule count/avgScore d'une cible à partir de ses lignes byCommune déjà filtrées */
function recomputeTotal_(byCommuneFiltered){
  var count = 0, weighted = 0, weightedN = 0;
  byCommuneFiltered.forEach(function(r){
    count += r.count;
    if (typeof r.avgScore === 'number'){ weighted += r.avgScore * r.count; weightedN += r.count; }
  });
  return { count: count, avgScore: weightedN > 0 ? Math.min(100, Math.round((weighted / weightedN) * 10) / 10) : null };
}

function applyScopeToResult_(result, account){
  if (account.level === 'admin' || account.level === 'national' || account.level === 'superviseur') return result; // accès complet

  ['usagers','personnels','cogecs'].forEach(function(cible){
    if (result[cible] && result[cible].byCommune){
      result[cible].byCommune = result[cible].byCommune.filter(function(r){ return communeInScope_(r.label, account); });
    }
    if (result[cible] && result[cible].byStructure){
      result[cible].byStructure = result[cible].byStructure.filter(function(r){ return structureInScope_(r.label, account); });
    }
    if (result[cible] && result[cible].byEnqueteur){
      result[cible].byEnqueteur = []; // non rattachable de façon fiable à une géographie précise
    }
  });

  if (result.global && result.global.byCommune){
    result.global.byCommune = result.global.byCommune.filter(function(r){ return communeInScope_(r.label, account); });
  }
  if (result.global && result.global.byCible){
    var tU = recomputeTotal_(result.usagers.byCommune);
    var tP = recomputeTotal_(result.personnels.byCommune);
    var tC = recomputeTotal_(result.cogecs.byCommune);
    result.global.byCible = [
      { label: "Usagers", count: tU.count, avgScore: tU.avgScore },
      { label: "Personnel soignant", count: tP.count, avgScore: tP.avgScore },
      { label: "COGECS", count: tC.count, avgScore: tC.avgScore }
    ];
    result.counts = { usagers: tU.count, personnels: tP.count, cogecs: tC.count, total: tU.count + tP.count + tC.count };
  }

  if (result.communeReport){
    ['usagers','personnels','cogecs'].forEach(function(cible){
      var cr = result.communeReport[cible];
      if (cr && cr.byCommune){
        Object.keys(cr.byCommune).forEach(function(commune){
          if (!communeInScope_(commune, account)) delete cr.byCommune[commune];
        });
      }
    });
  }

  if (result.structureReport){
    ['usagers','personnels','cogecs'].forEach(function(cible){
      var sr = result.structureReport[cible];
      if (sr && sr.byStructure){
        Object.keys(sr.byStructure).forEach(function(structure){
          if (!structureInScope_(structure, account)) delete sr.byStructure[structure];
        });
      }
    });
  }

  if (result.enqueteurReport){
    result.enqueteurReport.usagers = { categories: [], byEnqueteur: {} };
    result.enqueteurReport.cogecs = { categories: [], byEnqueteur: {} };
  }

  if (result.gpsByDay){
    Object.keys(result.gpsByDay).forEach(function(day){
      result.gpsByDay[day] = (result.gpsByDay[day] || []).filter(function(p){
        return !p.commune || communeInScope_(p.commune, account);
      });
    });
  }

  return result;
}
