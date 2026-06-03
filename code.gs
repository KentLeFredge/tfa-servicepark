// ============================================================
// TFA SERVICE PARK — Google Apps Script Backend
// Auth Steam OpenID 2.0 + Google Sheets
// ============================================================

var PROPS = PropertiesService.getScriptProperties();

// Composants gérés (ordre fixe)
var COMPONENTS = ['refroidissement', 'direction', 'transmission', 'suspension', 'chassis'];
var ZONES      = ['AVANT', 'ARRIERE', 'LATERAL_G', 'LATERAL_D', 'AERIEN'];

// ──────────────────────────────────────────────────────────────
// CONFIG SHEET — valeurs par défaut
// ──────────────────────────────────────────────────────────────

var CONFIG_DEFAULTS = {
  // Moteur physique
  SPEED_MIN_KMH:              8,
  SPEED_REF_CAR_KMH:        100,
  SPEED_REF_ENV_KMH:        150,
  SPEED_EXPONENT:           1.8,
  CAR_MULTIPLIER:           0.7,
  ENV_MULTIPLIER:           1.0,
  // Filtre artefact
  ARTEFACT_CAR_MAX_SPEED:   150,
  ARTEFACT_MAX_REL_Y:       0.3,
  // Seuils score → sévérité (pour affichage)
  SCORE_RIEN_MAX:            15,
  SCORE_LEGER_MAX:           20,
  SCORE_MODERE_MAX:          45,
  SCORE_SEVERE_MAX:          70,
  // Courbe pénalité résiduelle (exposant)
  PENALTY_CURVE:             0.7,
  // Pénalité max par composant (ballast kg ou restrictor)
  PENALTY_REFROIDISSEMENT_MAX: 12,   // restrictor
  PENALTY_DIRECTION_MAX:       20,   // ballast_kg
  PENALTY_TRANSMISSION_MAX:    8,    // restrictor
  PENALTY_SUSPENSION_MAX:      8,    // ballast_kg
  PENALTY_CHASSIS_MAX:         20,   // ballast_kg
  // Coût de réparation max par composant (minutes)
  REPAIR_COST_REFROIDISSEMENT: 40,
  REPAIR_COST_DIRECTION:       35,
  REPAIR_COST_TRANSMISSION:    30,
  REPAIR_COST_SUSPENSION:      20,
  REPAIR_COST_CHASSIS:         50,
  // Budget pilotes
  REPAIR_BUDGET_MIN:            60,
  // Dépassement budget → pénalité chrono
  OVERRUN_MAX_MIN:             15,
  OVERRUN_TO_SECONDS:           2,
  // Restrictor : pondération du 2e composant
  RESTRICTOR_SECOND_WEIGHT:    0.30,
};

var CONFIG_DESCRIPTIONS = {
  SPEED_MIN_KMH:              'Vitesse min (km/h) pour qu\'un choc cause des dégâts',
  SPEED_REF_CAR_KMH:          'Référence vitesse choc entre voitures (100 km/h = impact catastrophique)',
  SPEED_REF_ENV_KMH:          'Référence vitesse choc mur/env (150 km/h = sortie grave)',
  SPEED_EXPONENT:             'Exposant courbe énergie (1.8 = convexe, aggrave les gros chocs)',
  CAR_MULTIPLIER:             'Multiplicateur vitesse effective pour choc CAR (0.7 = moins sévère qu\'ENV)',
  ENV_MULTIPLIER:             'Multiplicateur vitesse effective pour choc ENV (1.0 = référence)',
  ARTEFACT_CAR_MAX_SPEED:     'Filtre artefact : vitesse CAR max au-delà = ignoré si |Y| < seuil',
  ARTEFACT_MAX_REL_Y:         'Filtre artefact : seuil |rel_Y| pour ignorer les téléportations',
  SCORE_RIEN_MAX:             'Score max pour "rien" (< 5 → pas de dégât affiché)',
  SCORE_LEGER_MAX:            'Score max pour sévérité légère',
  SCORE_MODERE_MAX:           'Score max pour sévérité modérée',
  SCORE_SEVERE_MAX:           'Score max pour sévérité sévère (au-delà = critique)',
  PENALTY_CURVE:              'Exposant courbe pénalité résiduelle (0.7 = progressive)',
  PENALTY_REFROIDISSEMENT_MAX:'Pénalité max refroidissement (RESTRICTOR)',
  PENALTY_DIRECTION_MAX:      'Pénalité max direction (BALLAST kg)',
  PENALTY_TRANSMISSION_MAX:   'Pénalité max transmission (RESTRICTOR)',
  PENALTY_SUSPENSION_MAX:     'Pénalité max suspension (BALLAST kg)',
  PENALTY_CHASSIS_MAX:        'Pénalité max châssis (BALLAST kg)',
  REPAIR_COST_REFROIDISSEMENT:'Coût réparation max refroidissement (min)',
  REPAIR_COST_DIRECTION:      'Coût réparation max direction (min)',
  REPAIR_COST_TRANSMISSION:   'Coût réparation max transmission (min)',
  REPAIR_COST_SUSPENSION:     'Coût réparation max suspension (min)',
  REPAIR_COST_CHASSIS:        'Coût réparation max châssis (min)',
  REPAIR_BUDGET_MIN:          'Budget total réparation alloué aux pilotes (min)',
  OVERRUN_MAX_MIN:            'Dépassement budget max autorisé (min)',
  OVERRUN_TO_SECONDS:         '1 min de dépassement = X secondes de pénalité chrono',
  RESTRICTOR_SECOND_WEIGHT:   'Poids du 2ème composant restrictor (0.30 = 30% du second)',
};

// Type de pénalité par composant (fixe, déduit de la mécanique)
var PENALTY_TYPE = {
  refroidissement: 'restrictor',
  direction:       'ballast_kg',
  transmission:    'restrictor',
  suspension:      'ballast_kg',
  chassis:         'ballast_kg',
};

function readConfigSheet() {
  var cfg = {};
  Object.keys(CONFIG_DEFAULTS).forEach(function(k) { cfg[k] = CONFIG_DEFAULTS[k]; });
  var sheet = getSheet('config');
  if (!sheet) return cfg;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0] || '').trim();
    var val = data[i][1];
    if (key && val !== '' && val !== null && val !== undefined) cfg[key] = val;
  }
  return cfg;
}

function ensureConfigSheet() {
  var ss    = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  var sheet = ss.getSheetByName('config');
  if (!sheet) {
    sheet = ss.insertSheet('config');
    sheet.appendRow(['key', 'value', 'description']);
    Object.keys(CONFIG_DEFAULTS).forEach(function(k) {
      sheet.appendRow([k, CONFIG_DEFAULTS[k], CONFIG_DESCRIPTIONS[k] || '']);
    });
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setColumnWidth(1, 250); sheet.setColumnWidth(2, 80); sheet.setColumnWidth(3, 400);
  }
  return sheet;
}

// ──────────────────────────────────────────────────────────────
// MOTEUR DE DÉGÂTS — Zone detection
// ──────────────────────────────────────────────────────────────

function detectZone(rx, ry, rz, speed) {
  // Tonneau / décollage
  if (Math.abs(ry) > 0.5 && speed > 30) return 'AERIEN';
  var az = Math.abs(rz), ax = Math.abs(rx);
  if (az >= ax) return rz >= 0 ? 'AVANT' : 'ARRIERE';
  return rx < 0 ? 'LATERAL_G' : 'LATERAL_D';
}

// ──────────────────────────────────────────────────────────────
// MOTEUR DE DÉGÂTS — Calcul scores depuis impacts
// ──────────────────────────────────────────────────────────────

function computeScoresFromImpacts(impacts, vehicleZones, cfg) {
  var scores = {};
  COMPONENTS.forEach(function(c) { scores[c] = 0; });

  // Fallback générique si véhicule inconnu
  if (!vehicleZones) {
    vehicleZones = {};
    ZONES.forEach(function(z) {
      vehicleZones[z] = {};
      COMPONENTS.forEach(function(c) { vehicleZones[z][c] = 0.35; });
    });
  }

  impacts.forEach(function(imp) {
    var speed = imp.speed || 0;
    var type  = imp.type  || 'ENV';
    var rx    = imp.rx    || 0;
    var ry    = imp.ry    || 0;
    var rz    = imp.rz    || 0;

    if (speed < (Number(cfg.SPEED_MIN_KMH) || 8)) return;

    // Filtre artefact téléportation
    var artMax = Number(cfg.ARTEFACT_CAR_MAX_SPEED) || 150;
    var artY   = Number(cfg.ARTEFACT_MAX_REL_Y)     || 0.3;
    if (type === 'CAR' && speed > artMax && Math.abs(ry) < artY) return;

    var mult    = (type === 'ENV') ? (Number(cfg.ENV_MULTIPLIER) || 1.0)
                                    : (Number(cfg.CAR_MULTIPLIER) || 0.7);
    var ref     = (type === 'ENV') ? (Number(cfg.SPEED_REF_ENV_KMH) || 150)
                                    : (Number(cfg.SPEED_REF_CAR_KMH) || 100);
    var exp     = Number(cfg.SPEED_EXPONENT) || 1.8;
    var speedEff= speed * mult;
    var energy  = Math.min(Math.pow(speedEff / ref, exp) * 100, 100);

    var zone        = detectZone(rx, ry, rz, speed);
    var zoneWeights = vehicleZones[zone] || {};

    COMPONENTS.forEach(function(comp) {
      var w = Number(zoneWeights[comp]) || 0;
      scores[comp] = Math.min(scores[comp] + energy * w, 100);
    });
  });

  // Arrondir les scores
  COMPONENTS.forEach(function(c) { scores[c] = Math.round(scores[c] * 10) / 10; });
  return scores;
}

function scoreToSeverity(score, cfg) {
  if (score < (Number(cfg.SCORE_RIEN_MAX)   || 5))  return 'intact';
  if (score < (Number(cfg.SCORE_LEGER_MAX)  || 20)) return 'leger';
  if (score < (Number(cfg.SCORE_MODERE_MAX) || 45)) return 'modere';
  if (score < (Number(cfg.SCORE_SEVERE_MAX) || 70)) return 'severe';
  return 'critique';
}

// Pénalité résiduelle depuis score (progressif)
// repair_pct : 0–100 (% du score réparé), pour UI binaire = 0 ou 100
function scoreToPenaltyValue(score, repairPct, penMax, cfg) {
  var repaired     = Math.max(0, Math.min(repairPct, score));
  var scoreResidual= Math.max(0, score - repaired);
  if (scoreResidual <= 0) return 0;
  var curve = Number(cfg.PENALTY_CURVE) || 0.7;
  return Math.round(Math.pow(scoreResidual / 100, curve) * penMax * 10) / 10;
}

// Coût de réparation (minutes)
function scoreToRepairCost(score, repairPct, repairCostMax, cfg) {
  if (score <= 0) return 0;
  var repaired = Math.max(0, Math.min(repairPct, score));
  return Math.round((repaired / score) * repairCostMax * (score / 100) * 10) / 10;
}

// ──────────────────────────────────────────────────────────────
// VEHICLE PROFILES — Lecture depuis Sheet
// ──────────────────────────────────────────────────────────────

function loadVehicleProfiles() {
  var sheet = getSheet('vehicle_profiles');
  if (!sheet) return {};
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var modelIdx= headers.indexOf('model_id');
  var zoneIdx = headers.indexOf('zone');
  var compIdx = {};
  COMPONENTS.forEach(function(c) { compIdx[c] = headers.indexOf(c); });

  var profiles = {};
  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var model = String(row[modelIdx] || '').trim();
    var zone  = String(row[zoneIdx]  || '').trim();
    if (!model || !zone) continue;
    if (!profiles[model]) profiles[model] = {};
    profiles[model][zone] = {};
    COMPONENTS.forEach(function(comp) {
      var idx = compIdx[comp];
      profiles[model][zone][comp] = (idx >= 0 && row[idx] !== '') ? (Number(row[idx]) || 0) : 0.35;
    });
  }
  return profiles;
}

// ──────────────────────────────────────────────────────────────
// ENTRY POINTS
// ──────────────────────────────────────────────────────────────

function doGet(e) {
  var action = e.parameter.action || '';
  if (action === 'steam_callback')        return handleSteamCallback(e);
  if (action === 'get_driver')            return withAuth(e, function(steamId) {
    var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
    return jsonResponse(getDriverData(steamId, stageId));
  });
  if (action === 'save_repairs')          return withAuth(e, function(steamId) {
    var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
    saveRepairs(steamId, stageId, JSON.parse(e.parameter.components || '[]'));
    return jsonResponse({ ok: true });
  });
  if (action === 'validate')              return withAuth(e, function(steamId) {
    validateRepairs(steamId, e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'));
    return jsonResponse({ ok: true });
  });
  if (action === 'admin_overview')        return withAdminToken(e, function() {
    return jsonResponse(getAdminOverview());
  });
  if (action === 'export_entry_list')     return withAdminToken(e, function() {
    return jsonResponse(exportEntryList());
  });
  if (action === 'get_config')            return withAdminToken(e, function() {
    return jsonResponse(getConfigForAdmin());
  });
  if (action === 'get_admins')            return withAdminToken(e, function() {
    return jsonResponse(getAdminsList());
  });
  if (action === 'fetch_acsm_results')    return withAdminToken(e, function() {
    return jsonResponse(fetchAcsmResults());
  });
  if (action === 'sync_acsm_championship') return withAdminToken(e, function() {
    return jsonResponse(syncAcsmChampionship());
  });
  if (action === 'admin_get_driver_repairs') return withAdminToken(e, function() {
    var guid    = e.parameter.driver_guid || '';
    var stageId = e.parameter.stage_id   || PROPS.getProperty('CURRENT_STAGE_ID');
    return jsonResponse(adminGetDriverRepairs(guid, stageId));
  });
  return jsonResponse({ error: 'unknown_action' });
}

function doPost(e) {
  var action = e.parameter.action || '';
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch(err) {}

  if (action === 'save_repairs')          return withAuth(e, function(steamId) {
    saveRepairs(steamId, e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'), body.components || []);
    return jsonResponse({ ok: true });
  });
  if (action === 'validate')              return withAuth(e, function(steamId) {
    validateRepairs(steamId, e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'));
    return jsonResponse({ ok: true });
  });
  if (action === 'import_championship')     return withAdminToken(e, function() {
    return jsonResponse(importChampionship(body));
  });
  if (action === 'import_session_json')     return withAdminToken(e, function() {
    var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
    return jsonResponse(importSessionJson(body, stageId));
  });
  if (action === 'import_vehicle_profiles') return withAdminToken(e, function() {
    return jsonResponse(importVehicleProfiles(body));
  });
  if (action === 'reset_stage')           return withAdminToken(e, function() {
    return jsonResponse(resetStage(body.stage_id || PROPS.getProperty('CURRENT_STAGE_ID')));
  });
  if (action === 'set_stage')             return withAdminToken(e, function() {
    PROPS.setProperty('CURRENT_STAGE_ID', body.stage_id);
    return jsonResponse({ ok: true, stage_id: body.stage_id });
  });
  if (action === 'purge_db')              return withAdminToken(e, function() {
    return jsonResponse(purgeDatabase());
  });
  if (action === 'save_config')           return withAdminToken(e, function() {
    return jsonResponse(saveConfigFromAdmin(body.config || {}));
  });
  if (action === 'save_admins')           return withAdminToken(e, function() {
    return jsonResponse(saveAdminsList(body.admins || []));
  });
  if (action === 'admin_set_repair')      return withAdminToken(e, function() {
    return jsonResponse(adminSetRepair(
      body.driver_guid, body.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'),
      body.component_id, body.repaired
    ));
  });
  if (action === 'admin_set_validated')   return withAdminToken(e, function() {
    return jsonResponse(adminSetValidated(
      body.driver_guid, body.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'),
      body.validated
    ));
  });
  return jsonResponse({ error: 'unknown_action' });
}

// ──────────────────────────────────────────────────────────────
// AUTH
// ──────────────────────────────────────────────────────────────

function withAuth(e, fn) {
  if (!validateToken(e.parameter.token || '', e.parameter.steam_id || ''))
    return jsonResponse({ error: 'unauthorized' });
  try { return fn(e.parameter.steam_id); } catch(err) { return jsonResponse({ error: err.message }); }
}

function withAdminToken(e, fn) {
  var token   = e.parameter.token    || '';
  var steamId = e.parameter.steam_id || '';
  if (!validateToken(token, steamId))    return jsonResponse({ error: 'unauthorized' });
  if (!isAdminSteamId(steamId))          return jsonResponse({ error: 'admin_forbidden' });
  try { return fn(steamId); } catch(err) { return jsonResponse({ error: err.message }); }
}

function isAdminSteamId(steamId) {
  // 1. Lire depuis la feuille "admins" (source principale)
  var sheet = getSheet('admins');
  if (sheet) {
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx   = headers.indexOf('steam_id');
    var actIdx  = headers.indexOf('active');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx] || '').trim() === String(steamId)) {
        // active = TRUE ou vide/absent → autorisé
        var active = data[i][actIdx];
        return active !== false && active !== 'FALSE' && active !== 0;
      }
    }
    return false;
  }
  // 2. Fallback Script Properties (séparateurs , ou ;)
  var admins = PROPS.getProperty('ADMIN_STEAM_IDS') || '';
  return admins.split(/[,;]/).map(function(s) { return s.trim(); }).indexOf(String(steamId)) !== -1;
}

function ensureAdminsSheet() {
  var ss    = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  var sheet = ss.getSheetByName('admins');
  if (!sheet) {
    sheet = ss.insertSheet('admins');
    sheet.appendRow(['steam_id', 'driver_name', 'active']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 70);
    // Migration automatique depuis Script Properties
    var existing = PROPS.getProperty('ADMIN_STEAM_IDS') || '';
    existing.split(/[,;]/).forEach(function(id) {
      id = id.trim();
      if (id) sheet.appendRow([id, '', true]);
    });
  }
  return sheet;
}

// ──────────────────────────────────────────────────────────────
// STEAM OPENID
// ──────────────────────────────────────────────────────────────

function handleSteamCallback(e) {
  var claimedId = e.parameter['openid.claimed_id'] || '';
  var steamId   = claimedId.replace(/.*\//, '');
  if (!/^\d{17}$/.test(steamId))            return HtmlService.createHtmlOutput('<h2>Auth failed: invalid Steam ID</h2>');
  if (!verifySteamSignature(e.parameters))  return HtmlService.createHtmlOutput('<h2>Auth failed: signature invalide</h2>');
  var token      = generateToken(steamId);
  var serviceUrl = PROPS.getProperty('SERVICE_PARK_URL');
  var isAdmin    = isAdminSteamId(steamId) ? '1' : '0';
  var redirectUrl = serviceUrl
    + '?token='    + encodeURIComponent(token)
    + '&steam_id=' + encodeURIComponent(steamId)
    + '&is_admin=' + isAdmin;
  return HtmlService.createHtmlOutput('<script>window.location.href=' + JSON.stringify(redirectUrl) + ';</script>');
}

function verifySteamSignature(params) {
  var postParams = [];
  for (var key in params) {
    if (key.indexOf('openid.') === 0) {
      var val = params[key][0];
      if (key === 'openid.mode') val = 'check_authentication';
      postParams.push(encodeURIComponent(key) + '=' + encodeURIComponent(val));
    }
  }
  var response = UrlFetchApp.fetch('https://steamcommunity.com/openid/login', {
    method: 'post', contentType: 'application/x-www-form-urlencoded',
    payload: postParams.join('&'), muteHttpExceptions: true
  });
  return response.getContentText().indexOf('is_valid:true') !== -1;
}

function generateToken(steamId) {
  var secret    = PROPS.getProperty('STEAM_SECRET');
  var timestamp = Math.floor(Date.now() / 1000).toString();
  var rawHmac   = Utilities.computeHmacSha256Signature(steamId + ':' + timestamp, secret);
  var hmacHex   = rawHmac.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  return hmacHex + ':' + timestamp;
}

function validateToken(token, steamId) {
  if (!token || !steamId) return false;
  var parts = token.split(':');
  if (parts.length !== 2) return false;
  if (Math.floor(Date.now() / 1000) - parseInt(parts[1], 10) > 14400) return false;
  var secret    = PROPS.getProperty('STEAM_SECRET');
  var rawHmac   = Utilities.computeHmacSha256Signature(steamId + ':' + parts[1], secret);
  var expectedHex = rawHmac.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  return parts[0] === expectedHex;
}

// ──────────────────────────────────────────────────────────────
// SHEETS HELPERS
// ──────────────────────────────────────────────────────────────

function getSheet(name) {
  return SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID')).getSheetByName(name);
}

// Ajoute une colonne manquante à driver_state si nécessaire
function ensureDriverStateColumn(colName) {
  var sheet   = getSheet('driver_state');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf(colName) === -1) {
    var newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue(colName);
  }
}

function getOrCreateSheet(name, headers) {
  var ss    = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

// ──────────────────────────────────────────────────────────────
// DRIVER DATA
// ──────────────────────────────────────────────────────────────

function getDriverData(steamId, stageId) {
  var cfg      = readConfigSheet();
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var driverRow= null;
  for (var i = 1; i < dsData.length; i++) {
    if (String(dsData[i][dsHeaders.indexOf('driver_guid')]) === String(steamId)) { driverRow = dsData[i]; break; }
  }
  if (!driverRow) throw new Error('driver_not_found');
  function dsVal(col) { var idx = dsHeaders.indexOf(col); return idx >= 0 ? driverRow[idx] : null; }

  // Composants depuis damage_components
  var dcSheet = getSheet('damage_components');
  var rawComponents = [];
  if (dcSheet) {
    var dcData   = dcSheet.getDataRange().getValues();
    var dcHeaders= dcData[0];
    for (var j = 1; j < dcData.length; j++) {
      var row = dcData[j];
      if (String(row[dcHeaders.indexOf('driver_guid')]) === String(steamId) &&
          String(row[dcHeaders.indexOf('stage_id')])    === String(stageId)) {
        var dc = (function(r) { return function(col) { var idx = dcHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
        rawComponents.push({
          id:         String(dc('component_id')),
          score:      Number(dc('score')) || 0,
          repair_pct: dc('repaired') === true || dc('repaired') === 'TRUE' ? 100 : 0,
        });
      }
    }
  }

  // Si aucun composant → afficher tous à 100% (intact)
  var components;
  if (rawComponents.length === 0) {
    components = COMPONENTS.map(function(c) {
      return { id: c, score: 0, severity: 'intact',
               ballast_kg: 0, restrictor: 0, repair_min: 0, repaired: false };
    });
  } else {
    // Calculer pénalités progressives depuis score
    var penMax = {
      refroidissement: Number(cfg.PENALTY_REFROIDISSEMENT_MAX) || 12,
      direction:       Number(cfg.PENALTY_DIRECTION_MAX)       || 20,
      transmission:    Number(cfg.PENALTY_TRANSMISSION_MAX)    || 8,
      suspension:      Number(cfg.PENALTY_SUSPENSION_MAX)      || 8,
      chassis:         Number(cfg.PENALTY_CHASSIS_MAX)         || 20,
    };
    var repairMax = {
      refroidissement: Number(cfg.REPAIR_COST_REFROIDISSEMENT) || 40,
      direction:       Number(cfg.REPAIR_COST_DIRECTION)       || 35,
      transmission:    Number(cfg.REPAIR_COST_TRANSMISSION)    || 30,
      suspension:      Number(cfg.REPAIR_COST_SUSPENSION)      || 20,
      chassis:         Number(cfg.REPAIR_COST_CHASSIS)         || 50,
    };

    // Index par id pour merge propre avec COMPONENTS
    var rawById = {};
    rawComponents.forEach(function(c) { rawById[c.id] = c; });

    components = COMPONENTS.map(function(compId) {
      var raw = rawById[compId];
      if (!raw) return { id: compId, score: 0, severity: 'intact',
                         ballast_kg: 0, restrictor: 0, repair_min: 0, repaired: false };
      var score      = raw.score;
      var repairPct  = raw.repair_pct;
      var sev        = scoreToSeverity(score, cfg);
      var penType    = PENALTY_TYPE[compId] || 'ballast_kg';
      var penVal     = scoreToPenaltyValue(score, repairPct, penMax[compId] || 0, cfg);
      var repairCost = scoreToRepairCost(score, repairPct, repairMax[compId] || 0, cfg);
      var repaired   = repairPct >= 100;
      return {
        id:         compId,
        score:      score,
        severity:   sev,
        ballast_kg: penType === 'ballast_kg' ? penVal : 0,
        restrictor: penType === 'restrictor' ? penVal : 0,
        repair_min: Math.ceil(repairCost),
        repaired:   repaired,
      };
    });
  }

  // Calculer totaux ballast/restrictor (règle restrictor : max + 30% du second)
  var ballastTotal = 0;
  var restrValues  = [];
  components.forEach(function(c) {
    if (!c.repaired) {
      ballastTotal += c.ballast_kg;
      if (c.restrictor > 0) restrValues.push(c.restrictor);
    }
  });
  restrValues.sort(function(a,b) { return b - a; });
  var w2  = Number(cfg.RESTRICTOR_SECOND_WEIGHT) || 0.3;
  var restrTotal = restrValues.length === 0 ? 0
                 : restrValues.length === 1 ? restrValues[0]
                 : restrValues[0] + restrValues[1] * w2;

  // Pénalité dépassement budget
  var repairUsed = Number(dsVal('repair_used_min')) || 0;
  var budget     = Number(cfg.REPAIR_BUDGET_MIN)    || 60;
  var overrunMax = Number(cfg.OVERRUN_MAX_MIN)      || 15;
  var overrun2s  = Number(cfg.OVERRUN_TO_SECONDS)   || 2;
  var overrun    = Math.min(Math.max(0, repairUsed - budget), overrunMax);

  // Résultats étape + étape suivante
  var stageResults = getStageResults(stageId);
  var driverResult = null;
  stageResults.forEach(function(r) {
    if (String(r.driver_guid) === String(steamId)) driverResult = r;
  });
  var nextStage = getNextStageInfo(stageId);

  return {
    driver_guid:       String(steamId),
    driver_name:       String(dsVal('driver_name') || ''),
    car_model:         String(dsVal('car_model')   || ''),
    skin:              String(dsVal('skin')         || ''),
    stage_id:          String(stageId),
    validated:         dsVal('validated') === true || dsVal('validated') === 'TRUE',
    repair_budget_min: budget,
    repair_used_min:   repairUsed,
    ballast_kg:        Math.round(Math.min(ballastTotal, 40) * 10) / 10,
    restrictor:        Math.round(Math.min(restrTotal, 12) * 10) / 10,
    penalty_seconds:   overrun * overrun2s,
    components:        components,
    best_lap_ms:       driverResult ? driverResult.best_lap_ms  : null,
    stage_pos:         driverResult ? driverResult.position     : null,
    stage_results:     stageResults,
    next_stage:        nextStage,
  };
}

// ──────────────────────────────────────────────────────────────
// STAGE RESULTS / NEXT STAGE
// ──────────────────────────────────────────────────────────────

function stageIdToIndex(stageId) {
  var evSheet = getSheet('championship_events');
  if (!evSheet) return -1;
  var data = evSheet.getDataRange().getValues();
  var h = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][h.indexOf('event_id')]) === String(stageId) ||
        'event_' + data[i][h.indexOf('event_index')] === stageId) {
      return Number(data[i][h.indexOf('event_index')]);
    }
  }
  return -1;
}

function getStageResults(stageId) {
  var srSheet = getSheet('stage_results');
  if (!srSheet) return [];
  var data = srSheet.getDataRange().getValues();
  var h    = data[0];
  var curIdx = stageIdToIndex(stageId);
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][h.indexOf('event_index')]) === curIdx) {
      var row = data[i];
      var dv  = (function(r) { return function(col) { var idx = h.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
      results.push({
        driver_guid: String(dv('driver_guid')  || ''),
        driver_name: String(dv('driver_name')  || ''),
        car_model:   String(dv('car_model')    || ''),
        skin:        String(dv('skin')         || ''),
        best_lap_ms: Number(dv('best_lap_ms')) || 0,
        laps:        Number(dv('laps'))        || 0,
        position:    Number(dv('position'))    || 0,
      });
    }
  }
  return results.sort(function(a,b) { return a.position - b.position; });
}

function getNextStageInfo(stageId) {
  var evSheet = getSheet('championship_events');
  if (!evSheet) return null;
  var data = evSheet.getDataRange().getValues();
  var h    = data[0];
  var curIdx = stageIdToIndex(stageId);
  if (curIdx < 0) return null;
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][h.indexOf('event_index')]) === curIdx + 1) {
      var row = data[i];
      var dv  = (function(r) { return function(col) { var idx = h.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
      return {
        event_index:     Number(dv('event_index')),
        event_id:        String(dv('event_id')      || ''),
        event_name:      String(dv('event_name')    || ''),
        track:           String(dv('track')          || ''),
        layout:          String(dv('layout')         || ''),
        weather_ambient: Number(dv('weather_ambient'))  || 0,
        weather_road:    Number(dv('weather_road'))     || 0,
        weather_wind:    Number(dv('weather_wind'))     || 0,
        cmwfx_type:      Number(dv('cmwfx_type'))       || 0,
        race_duration:   Number(dv('race_duration_min'))|| 0,
      };
    }
  }
  return null;
}

// ──────────────────────────────────────────────────────────────
// SAVE REPAIRS
// ──────────────────────────────────────────────────────────────

function saveRepairs(steamId, stageId, components) {
  var cfg      = readConfigSheet();
  var dcSheet  = getSheet('damage_components');
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];
  var guidIdx  = dcHeaders.indexOf('driver_guid');
  var stageIdx = dcHeaders.indexOf('stage_id');
  var compIdIdx= dcHeaders.indexOf('component_id');
  var repIdx   = dcHeaders.indexOf('repaired');

  var existingRows = {};
  for (var i = 1; i < dcData.length; i++) {
    if (String(dcData[i][guidIdx]) === String(steamId) && String(dcData[i][stageIdx]) === String(stageId))
      existingRows[String(dcData[i][compIdIdx])] = i + 1;
  }

  var penMax = {
    refroidissement: Number(cfg.PENALTY_REFROIDISSEMENT_MAX) || 12,
    direction:       Number(cfg.PENALTY_DIRECTION_MAX)       || 20,
    transmission:    Number(cfg.PENALTY_TRANSMISSION_MAX)    || 8,
    suspension:      Number(cfg.PENALTY_SUSPENSION_MAX)      || 8,
    chassis:         Number(cfg.PENALTY_CHASSIS_MAX)         || 20,
  };
  var repairMax = {
    refroidissement: Number(cfg.REPAIR_COST_REFROIDISSEMENT) || 40,
    direction:       Number(cfg.REPAIR_COST_DIRECTION)       || 35,
    transmission:    Number(cfg.REPAIR_COST_TRANSMISSION)    || 30,
    suspension:      Number(cfg.REPAIR_COST_SUSPENSION)      || 20,
    chassis:         Number(cfg.REPAIR_COST_CHASSIS)         || 50,
  };

  var totalBallast = 0, totalRestrictor = 0, totalRepairMin = 0;
  var restrValues  = [];

  for (var k = 0; k < components.length; k++) {
    var comp    = components[k];
    var repaired= comp.repaired === true;
    var score   = Number(comp.score) || 0;
    var repPct  = repaired ? 100 : 0;
    var penType = PENALTY_TYPE[comp.id] || 'ballast_kg';
    var penVal  = scoreToPenaltyValue(score, repPct, penMax[comp.id] || 0, cfg);
    var repCost = scoreToRepairCost(score, repPct, repairMax[comp.id] || 0, cfg);

    if (!repaired) {
      if (penType === 'ballast_kg') totalBallast += penVal;
      else                          restrValues.push(penVal);
      totalRepairMin += repCost;
    }
    if (existingRows[String(comp.id)]) dcSheet.getRange(existingRows[String(comp.id)], repIdx + 1).setValue(repaired);
  }

  // Règle restrictor : max + 30% du second
  restrValues.sort(function(a,b) { return b-a; });
  var w2 = Number(cfg.RESTRICTOR_SECOND_WEIGHT) || 0.3;
  totalRestrictor = restrValues.length === 0 ? 0
                  : restrValues.length === 1 ? restrValues[0]
                  : restrValues[0] + restrValues[1] * w2;

  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  for (var m = 1; m < dsData.length; m++) {
    if (String(dsData[m][dsHeaders.indexOf('driver_guid')]) === String(steamId)) {
      var r = m + 1;
      dsSheet.getRange(r, dsHeaders.indexOf('ballast_kg')     + 1).setValue(Math.min(totalBallast, 40));
      dsSheet.getRange(r, dsHeaders.indexOf('restrictor')     + 1).setValue(Math.min(totalRestrictor, 12));
      dsSheet.getRange(r, dsHeaders.indexOf('repair_used_min')+ 1).setValue(Math.ceil(totalRepairMin));
      dsSheet.getRange(r, dsHeaders.indexOf('last_updated')   + 1).setValue(new Date().toISOString());
      break;
    }
  }
}

// ──────────────────────────────────────────────────────────────
// VALIDATE REPAIRS
// ──────────────────────────────────────────────────────────────

function validateRepairs(steamId, stageId) {
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var guidIdx  = dsHeaders.indexOf('driver_guid');
  var valIdx   = dsHeaders.indexOf('validated');
  for (var i = 1; i < dsData.length; i++) {
    if (String(dsData[i][guidIdx]) === String(steamId)) {
      if (dsData[i][valIdx] === true || dsData[i][valIdx] === 'TRUE') throw new Error('already_validated');
      dsSheet.getRange(i + 1, valIdx + 1).setValue(true);
      dsSheet.getRange(i + 1, dsHeaders.indexOf('last_updated') + 1).setValue(new Date().toISOString());
      return;
    }
  }
  throw new Error('driver_not_found');
}

// ──────────────────────────────────────────────────────────────
// ADMIN — OVERVIEW
// ──────────────────────────────────────────────────────────────

function getAdminOverview() {
  var stageId  = PROPS.getProperty('CURRENT_STAGE_ID');
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var dcSheet  = getSheet('damage_components');
  var dcData   = dcSheet ? dcSheet.getDataRange().getValues() : [[]];
  var dcHeaders= dcData[0] || [];

  var dcIndex = {};
  for (var j = 1; j < dcData.length; j++) {
    var row  = dcData[j];
    var guid = String(row[dcHeaders.indexOf('driver_guid')]);
    var sid  = String(row[dcHeaders.indexOf('stage_id')]);
    if (sid !== stageId) continue;
    if (!dcIndex[guid]) dcIndex[guid] = [];
    dcIndex[guid].push({
      component_id: String(row[dcHeaders.indexOf('component_id')]),
      score:        Number(row[dcHeaders.indexOf('score')])    || 0,
      repaired:     row[dcHeaders.indexOf('repaired')] === true || row[dcHeaders.indexOf('repaired')] === 'TRUE'
    });
  }

  var stageResults  = getStageResults(stageId);
  var resultsByGuid = {};
  stageResults.forEach(function(r) { resultsByGuid[r.driver_guid] = r; });

  var drivers = [];
  for (var i = 1; i < dsData.length; i++) {
    var drow = dsData[i];
    var dv   = (function(r) { return function(col) { var idx = dsHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(drow);
    var guid = String(dv('driver_guid'));
    if (String(dv('car_model') || '') === 'tv_car') continue; // Exclure voitures de diffusion
    var comps= dcIndex[guid] || [];
    var res  = resultsByGuid[guid] || null;
    var damaged = comps.filter(function(c) { return c.score > 5; });
    drivers.push({
      driver_guid:     guid,
      driver_name:     String(dv('driver_name')     || ''),
      car_model:       String(dv('car_model')        || ''),
      validated:       dv('validated') === true || dv('validated') === 'TRUE',
      ballast_kg:      Number(dv('ballast_kg'))      || 0,
      restrictor:      Number(dv('restrictor'))      || 0,
      repair_used_min: Number(dv('repair_used_min')) || 0,
      last_updated:    String(dv('last_updated')     || ''),
      damaged_count:   damaged.length,
      repaired_count:  damaged.filter(function(c) { return c.repaired; }).length,
      components:      comps,
      best_lap_ms:     res ? res.best_lap_ms : null,
      stage_pos:       res ? res.position    : null,
    });
  }
  return { stage_id: stageId, drivers: drivers };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — IMPORT VEHICLE PROFILES (JSON)
// ──────────────────────────────────────────────────────────────

function importVehicleProfiles(data) {
  var ss    = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  var sheet = ss.getSheetByName('vehicle_profiles');
  if (!sheet) sheet = ss.insertSheet('vehicle_profiles');
  sheet.clearContents();

  var headers = ['model_id', 'pack', 'zone'].concat(COMPONENTS);
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  var rows    = [];
  var count   = 0;
  var packs   = data.packs || {};

  Object.keys(packs).forEach(function(packId) {
    var pack     = packs[packId];
    var vehicles = pack.vehicles || {};
    Object.keys(vehicles).forEach(function(modelId) {
      var vehicle = vehicles[modelId];
      var zones   = vehicle.zones || {};
      ZONES.forEach(function(zone) {
        var zw  = zones[zone] || {};
        var row = [modelId, packId, zone].concat(COMPONENTS.map(function(c) { return zw[c] !== undefined ? zw[c] : 0.35; }));
        rows.push(row);
        count++;
      });
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setColumnWidth(1, 300);
  return { ok: true, vehicles: count / ZONES.length, rows: count };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — IMPORT CHAMPIONSHIP JSON
// ──────────────────────────────────────────────────────────────

function importChampionship(data) {
  var ss = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));

  var evHeaders = ['event_index','event_id','event_name','track','layout',
                   'weather_ambient','weather_road','weather_wind','cmwfx_type','race_duration_min'];
  var evSheet = ss.getSheetByName('championship_events');
  if (!evSheet) evSheet = ss.insertSheet('championship_events');
  evSheet.clearContents();
  evSheet.appendRow(evHeaders);

  var srHeaders = ['event_index','driver_guid','driver_name','car_model','skin',
                   'best_lap_ms','laps','position','class_id'];
  var srSheet = ss.getSheetByName('stage_results');
  if (!srSheet) srSheet = ss.insertSheet('stage_results');
  srSheet.clearContents();
  srSheet.appendRow(srHeaders);

  var dcSheet = getOrCreateSheet('damage_components',
    ['driver_guid','stage_id','component_id','score','severity','ballast_kg','restrictor','repair_min','repaired']);

  ensureConfigSheet();
  ensureAdminsSheet();
  ensureDriverStateColumn('skin');
  ensureDriverStateColumn('team');
  var cfg             = readConfigSheet();
  var vehicleProfiles = loadVehicleProfiles();
  var events          = data.Events || [];
  var totalResults    = 0;
  var totalDamage     = 0;

  // ── Importer les entrants depuis Classes[].Entrants ──
  var stageIdForEntrants = PROPS.getProperty('CURRENT_STAGE_ID') || 'etape_01';
  var entrantCars = [];
  (data.Classes || []).forEach(function(cls) {
    var entrants = cls.Entrants || {};
    Object.keys(entrants).forEach(function(key) {
      var e = entrants[key];
      entrantCars.push({
        guid:       String(e.GUID || ''),
        name:       String(e.Name || ''),
        team:       String(e.Team || ''),
        model:      String(e.Model || ''),
        skin:       String(e.Skin  || ''),
        ballast:    0,
        restrictor: 0,
      });
    });
  });
  if (entrantCars.length) {
    // Réutilise importEntryList avec le format normalisé
    importEntryList(entrantCars.map(function(e) {
      return { GUID: e.guid, Name: e.name, Team: e.team, Model: e.model, Skin: e.skin };
    }), stageIdForEntrants);
  }

  // Supprimer les entrées damage_components liées à ce championnat
  var champEventIds = events.map(function(e, i) { return e.ID || ('event_' + i); });
  var dcData = dcSheet.getDataRange().getValues();
  var dcHdr  = dcData[0];
  for (var di = dcData.length - 1; di >= 1; di--) {
    if (champEventIds.indexOf(String(dcData[di][dcHdr.indexOf('stage_id')])) !== -1)
      dcSheet.deleteRow(di + 1);
  }

  for (var i = 0; i < events.length; i++) {
    var event  = events[i];
    var setup  = event.RaceSetup || {};

    // Météo de la course
    var raceWeather = null;
    var wMap = setup.Weather || {};
    Object.keys(wMap).forEach(function(wk) {
      if (wMap[wk].Sessions && wMap[wk].Sessions.indexOf('RACE') !== -1) raceWeather = wMap[wk];
    });
    if (!raceWeather) { var wk0 = Object.keys(wMap)[0]; if (wk0) raceWeather = wMap[wk0]; }

    var sessions     = setup.Sessions || {};
    var raceDuration = sessions.RACE ? (sessions.RACE.Time || 0) : 0;

    evSheet.appendRow([i, event.ID || '', event.Name || '', setup.Track || '', setup.TrackLayout || '',
      raceWeather ? raceWeather.BaseTemperatureAmbient : '',
      raceWeather ? raceWeather.BaseTemperatureRoad    : '',
      raceWeather ? Math.round((raceWeather.WindBaseSpeedMin + raceWeather.WindBaseSpeedMax) / 2) : '',
      raceWeather ? (raceWeather.CMWFXType || 0) : '', raceDuration]);

    var eventSess = event.Sessions || {};
    var raceSess  = eventSess.RACE  || {};
    var raceRes   = raceSess.Results|| {};
    var laps      = raceRes.Laps    || [];
    var cars      = raceRes.Cars    || [];
    var raceEvents= raceRes.Events  || [];

    if (!laps.length && !raceEvents.length) continue;

    // Index CarId → {guid, model, skin}
    var carMap = {};
    cars.forEach(function(car) {
      if (car.Driver) {
        carMap[car.CarId] = {
          guid:  car.Driver.Guid || '',
          model: car.Model || '',
          skin:  car.Skin  || '',
          class: car.ClassID || '',
        };
      }
    });

    // Index guid → model/skin depuis carMap
    var guidToModel = {};
    Object.keys(carMap).forEach(function(cid) {
      var c = carMap[cid];
      if (c.guid) guidToModel[c.guid] = { model: c.model, skin: c.skin };
    });

    // Résultats (meilleur tour)
    var bestLaps  = {}, lapCounts = {};
    laps.forEach(function(lap) {
      var guid = lap.DriverGuid;
      if (!guid || lap.LapTime > 1800000) return;
      if ((lap.CarModel || '') === 'tv_car') return;
      if (!bestLaps[guid] || lap.LapTime < bestLaps[guid].time) {
        bestLaps[guid] = { time: lap.LapTime, name: lap.DriverName || '', car: lap.CarModel || '' };
      }
      lapCounts[guid] = (lapCounts[guid] || 0) + 1;
    });
    var sorted = Object.keys(bestLaps).sort(function(a,b) { return bestLaps[a].time - bestLaps[b].time; });
    var srRows  = [];
    sorted.forEach(function(guid, pos) {
      var d    = bestLaps[guid];
      var info = guidToModel[guid] || {};
      srRows.push([i, guid, d.name, d.car, info.skin || '', d.time, lapCounts[guid] || 0, pos + 1, '']);
    });
    if (srRows.length) {
      srSheet.getRange(srSheet.getLastRow() + 1, 1, srRows.length, srHeaders.length).setValues(srRows);
      totalResults += srRows.length;
    }

    // Dégâts depuis collisions (avec profils véhicule + RelPosition)
    if (raceEvents.length > 0) {
      var stageId = event.ID || ('event_' + i);

      // Grouper impacts par guid
      var impactsByGuid = {};
      raceEvents.forEach(function(ev) {
        var cid  = ev.CarId;
        var info = carMap[cid];
        if (!info || !info.guid) return;
        var guid = info.guid;
        if (!impactsByGuid[guid]) impactsByGuid[guid] = [];
        var rel = ev.RelPosition || {};
        impactsByGuid[guid].push({
          speed: ev.ImpactSpeed || 0,
          type:  (ev.Type || 'ENV').indexOf('ENV') !== -1 ? 'ENV' : 'CAR',
          rx:    Number(rel.X || 0),
          ry:    Number(rel.Y || 0),
          rz:    Number(rel.Z || 0),
        });
      });

      var dcRows = [];
      Object.keys(impactsByGuid).forEach(function(guid) {
        var model   = (guidToModel[guid] || {}).model || '';
        var zones   = vehicleProfiles[model] || null;
        var scores  = computeScoresFromImpacts(impactsByGuid[guid], zones, cfg);

        COMPONENTS.forEach(function(compId) {
          var score = scores[compId];
          if (score < (Number(cfg.SCORE_RIEN_MAX) || 5)) return; // sous le seuil → ignoré
          var sev     = scoreToSeverity(score, cfg);
          var penType = PENALTY_TYPE[compId] || 'ballast_kg';
          var penVal  = scoreToPenaltyValue(score, 0, cfg['PENALTY_' + compId.toUpperCase() + '_MAX'] || 0, cfg);
          var repCost = scoreToRepairCost(score, 0, cfg['REPAIR_COST_' + compId.toUpperCase()] || 0, cfg);
          dcRows.push([guid, stageId, compId, score, sev,
            penType === 'ballast_kg' ? penVal : 0,
            penType === 'restrictor' ? penVal : 0,
            Math.ceil(repCost), false]);
          totalDamage++;
        });
      });
      if (dcRows.length) {
        dcSheet.getRange(dcSheet.getLastRow() + 1, 1, dcRows.length, 9).setValues(dcRows);
      }
    }
  }
  return { ok: true, events: events.length, results: totalResults, damage_entries: totalDamage };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — IMPORT SESSION CSV
// ──────────────────────────────────────────────────────────────

function importSessionCsv(csvText) {
  var cfg     = readConfigSheet();
  var stageId = PROPS.getProperty('CURRENT_STAGE_ID');

  // Index nom → guid depuis driver_state
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var nameToGuid = {}, nameToModel = {};
  for (var i = 1; i < dsData.length; i++) {
    var n = String(dsData[i][dsHeaders.indexOf('driver_name')] || '').toLowerCase().trim();
    var g = String(dsData[i][dsHeaders.indexOf('driver_guid')] || '');
    var m = String(dsData[i][dsHeaders.indexOf('car_model')]   || '');
    if (n && g) { nameToGuid[n] = g; nameToModel[n] = m; }
  }
  function findGuid(name) {
    var n = (name || '').toLowerCase().trim();
    if (nameToGuid[n]) return nameToGuid[n];
    var keys = Object.keys(nameToGuid);
    for (var k = 0; k < keys.length; k++) {
      if (keys[k].indexOf(n) !== -1 || n.indexOf(keys[k]) !== -1) return nameToGuid[keys[k]];
    }
    return null;
  }

  var lines = csvText.split('\n');

  // Race result
  var inRaceResult = false;
  var stageResults = [];
  for (var li = 0; li < lines.length; li++) {
    var line = lines[li].trim();
    if (line.indexOf('Race result') !== -1) { inRaceResult = true; continue; }
    if (inRaceResult && line.indexOf('Race laps') !== -1) { inRaceResult = false; break; }
    if (!inRaceResult) continue;
    var m = line.match(/^"(\d+)",\s*"[^"]*",\s*"[^"]*",\s*"[^"]*",\s*"([^"]+)",\s*"(\d+)",\s*"([^"]+)",\s*"([^"]+)"/);
    if (!m) continue;
    stageResults.push({ name: m[2].trim(), pos: Number(m[1]), laps: Number(m[3]), best_lap_str: m[5].trim() });
  }

  // Collisions
  var colRe = /"", "(.+?) reported contact with (environment|another vehicle .+?)\. Impact speed: ([\d.]+)"/;
  var impactsByName = {};
  for (var lj = 0; lj < lines.length; lj++) {
    var mc = lines[lj].match(colRe);
    if (!mc) continue;
    var dName = mc[1].trim();
    var cType = mc[2].trim();
    var speed = parseFloat(mc[3]);
    if (!impactsByName[dName]) impactsByName[dName] = [];
    // CSV n'a pas de RelPosition → approximation par type
    // ENV : choc frontal (AVANT), CAR : choc latéral (LATERAL_D approximatif)
    var zone = cType === 'environment' ? 'AVANT' : 'LATERAL_D';
    impactsByName[dName].push({ speed: speed, type: cType === 'environment' ? 'ENV' : 'CAR', zone: zone });
  }

  // Écrire stage_results
  var srSheet = getOrCreateSheet('stage_results',
    ['event_index','driver_guid','driver_name','car_model','skin','best_lap_ms','laps','position','class_id']);
  var srData = srSheet.getDataRange().getValues();
  var srH    = srData[0];
  for (var si = srData.length - 1; si >= 1; si--) {
    if (String(srData[si][srH.indexOf('event_index')]) === String(stageId)) srSheet.deleteRow(si + 1);
  }
  var inserted = 0, unmatched = [];
  stageResults.forEach(function(r) {
    var guid    = findGuid(r.name);
    var bestMs  = timeStrToMs(r.best_lap_str);
    var carModel= nameToModel[(r.name || '').toLowerCase().trim()] || '';
    if (!guid) { unmatched.push(r.name); }
    else {
      // Mettre à jour driver_name dans driver_state si vide
      var dsSheet2  = getSheet('driver_state');
      var dsData2   = dsSheet2.getDataRange().getValues();
      var dsH2      = dsData2[0];
      var gIdx2     = dsH2.indexOf('driver_guid');
      var nIdx2     = dsH2.indexOf('driver_name');
      for (var mi = 1; mi < dsData2.length; mi++) {
        if (String(dsData2[mi][gIdx2]) === String(guid)) {
          if (!String(dsData2[mi][nIdx2] || '').trim()) {
            dsSheet2.getRange(mi + 1, nIdx2 + 1).setValue(r.name);
          }
          break;
        }
      }
    }
    srSheet.appendRow([stageId, guid || '', r.name, carModel, '', bestMs, r.laps, r.pos, '']);
    inserted++;
  });

  // Écrire damage_components
  var dcSheet = getOrCreateSheet('damage_components',
    ['driver_guid','stage_id','component_id','score','severity','ballast_kg','restrictor','repair_min','repaired']);
  var dcData = dcSheet.getDataRange().getValues();
  var dcH    = dcData[0];
  for (var di = dcData.length - 1; di >= 1; di--) {
    if (String(dcData[di][dcH.indexOf('stage_id')]) === String(stageId)) dcSheet.deleteRow(di + 1);
  }

  var dmgInserted = 0;
  var dcRows = [];
  var vehicleProfiles = loadVehicleProfiles();

  Object.keys(impactsByName).forEach(function(name) {
    var guid = findGuid(name);
    if (!guid) return;
    var carModel = nameToModel[(name || '').toLowerCase().trim()] || '';
    var zones    = vehicleProfiles[carModel] || null;

    // Convertir les impacts CSV (avec zone approximative)
    var impacts = impactsByName[name].map(function(hit) {
      // Sans RelPosition exacte, on utilise la zone estimée
      var zoneWeights = zones ? (zones[hit.zone] || {}) : {};
      return { speed: hit.speed, type: hit.type, rx: 0, ry: 0, rz: hit.zone === 'AVANT' ? 1 : 0 };
    });

    var scores = computeScoresFromImpacts(impacts, zones, cfg);

    COMPONENTS.forEach(function(compId) {
      var score = scores[compId];
      if (score < (Number(cfg.SCORE_RIEN_MAX) || 5)) return;
      var sev     = scoreToSeverity(score, cfg);
      var penType = PENALTY_TYPE[compId] || 'ballast_kg';
      var penVal  = scoreToPenaltyValue(score, 0, cfg['PENALTY_' + compId.toUpperCase() + '_MAX'] || 0, cfg);
      var repCost = scoreToRepairCost(score, 0, cfg['REPAIR_COST_' + compId.toUpperCase()] || 0, cfg);
      dcRows.push([guid, stageId, compId, score, sev,
        penType === 'ballast_kg' ? penVal : 0,
        penType === 'restrictor' ? penVal : 0,
        Math.ceil(repCost), false]);
      dmgInserted++;
    });
  });

  if (dcRows.length) {
    dcSheet.getRange(dcSheet.getLastRow() + 1, 1, dcRows.length, 9).setValues(dcRows);
  }

  return { ok: true, stage_id: stageId, results: inserted, damage_entries: dmgInserted, unmatched: unmatched };
}

function timeStrToMs(str) {
  if (!str || str === '-') return 0;
  str = str.trim().replace(/^\'+/, '');
  var parts = str.split(':');
  if (parts.length === 2) return Math.round((parseFloat(parts[0]) * 60 + parseFloat(parts[1])) * 1000);
  return 0;
}

// ──────────────────────────────────────────────────────────────
// ADMIN — IMPORT ENTRY LIST
// ──────────────────────────────────────────────────────────────

function importEntryList(cars, stageId) {
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var existingGuids = {};
  for (var i = 1; i < dsData.length; i++)
    existingGuids[String(dsData[i][dsHeaders.indexOf('driver_guid')])] = i + 1;

  // Normalise le format d'entrée : supporte Cars[] (session JSON) et Classes[].Entrants (championnat)
  var normalised = [];
  cars.forEach(function(car) {
    // Format session JSON : { Driver: { Guid, Name }, Model, Skin, BallastKG, Restrictor }
    // Format championnat Classes.Entrants : { GUID, Name, Team, Model, Skin, Ballast, Restrictor }
    var driver = car.Driver || car.driver || {};
    normalised.push({
      guid:       String(driver.Guid || driver.guid || car.GUID || car.guid || ''),
      name:       String(driver.Name || driver.name || car.Name || car.name || ''),
      team:       String(driver.Team || driver.team || car.Team || car.team || ''),
      model:      String(car.Model   || car.model   || ''),
      skin:       String(car.Skin    || car.skin    || ''),
      ballast:    Number(car.BallastKG || car.Ballast || car.ballast || 0),
      restrictor: Number(car.Restrictor || car.restrictor || 0),
    });
  });

  var inserted = 0, updated = 0;
  normalised.forEach(function(e) {
    if (!e.guid || e.guid === 'undefined') return;
    if (e.model === 'tv_car') return; // Voiture de diffusion — exclure

    if (existingGuids[e.guid]) {
      var row = existingGuids[e.guid];
      dsSheet.getRange(row, dsHeaders.indexOf('driver_name') + 1).setValue(e.name);
      dsSheet.getRange(row, dsHeaders.indexOf('car_model')   + 1).setValue(e.model);
      if (dsHeaders.indexOf('skin') >= 0)
        dsSheet.getRange(row, dsHeaders.indexOf('skin') + 1).setValue(e.skin);
      if (dsHeaders.indexOf('team') >= 0)
        dsSheet.getRange(row, dsHeaders.indexOf('team') + 1).setValue(e.team);
      updated++;
    } else {
      var rowData = dsHeaders.map(function(h) {
        var map = { driver_guid: e.guid, driver_name: e.name, car_model: e.model,
          team: e.team, skin: e.skin,
          stage_id: stageId, validated: false, repair_used_min: 0,
          ballast_kg: 0, restrictor: 0, penalty_seconds: 0,
          last_updated: new Date().toISOString() };
        return map[h] !== undefined ? map[h] : '';
      });
      dsSheet.appendRow(rowData);
      inserted++;
    }
  });
  return { ok: true, inserted: inserted, updated: updated, total: normalised.length };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — RESET STAGE
// ──────────────────────────────────────────────────────────────

function resetStage(stageId) {
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  for (var i = 1; i < dsData.length; i++) {
    var r = i + 1;
    dsSheet.getRange(r, dsHeaders.indexOf('stage_id')        + 1).setValue(stageId);
    dsSheet.getRange(r, dsHeaders.indexOf('validated')       + 1).setValue(false);
    dsSheet.getRange(r, dsHeaders.indexOf('repair_used_min') + 1).setValue(0);
    dsSheet.getRange(r, dsHeaders.indexOf('ballast_kg')      + 1).setValue(0);
    dsSheet.getRange(r, dsHeaders.indexOf('restrictor')      + 1).setValue(0);
  }
  var dcSheet = getSheet('damage_components');
  if (dcSheet) {
    var dcData   = dcSheet.getDataRange().getValues();
    var dcHeaders= dcData[0];
    for (var j = 1; j < dcData.length; j++) {
      if (String(dcData[j][dcHeaders.indexOf('stage_id')]) === String(stageId))
        dcSheet.getRange(j + 1, dcHeaders.indexOf('repaired') + 1).setValue(false);
    }
  }
  return { ok: true, stage_id: stageId };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — EXPORT ENTRY LIST
// ──────────────────────────────────────────────────────────────

function exportEntryList() {
  ensureDriverStateColumn('skin');
  ensureDriverStateColumn('team');

  var cfg      = readConfigSheet();
  var stageId  = PROPS.getProperty('CURRENT_STAGE_ID') || '';

  // Index noms depuis stage_results (fallback si driver_name vide dans driver_state)
  var srSheet   = getSheet('stage_results');
  var nameByGuid= {};
  if (srSheet) {
    var srData = srSheet.getDataRange().getValues();
    var srH    = srData[0];
    for (var si = 1; si < srData.length; si++) {
      var sg = String(srData[si][srH.indexOf('driver_guid')] || '');
      var sn = String(srData[si][srH.indexOf('driver_name')] || '');
      if (sg && sn && !nameByGuid[sg]) nameByGuid[sg] = sn;
    }
  }

  // Lire driver_state
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];

  // Lire damage_components pour calculer les pénalités réelles
  var dcSheet  = getSheet('damage_components');
  var dcByGuid = {}; // guid → [composants]
  if (dcSheet) {
    var dcData   = dcSheet.getDataRange().getValues();
    var dcH      = dcData[0];
    for (var j = 1; j < dcData.length; j++) {
      var row  = dcData[j];
      var guid = String(row[dcH.indexOf('driver_guid')] || '');
      var sid  = String(row[dcH.indexOf('stage_id')]    || '');
      if (sid !== stageId) continue;
      if (!dcByGuid[guid]) dcByGuid[guid] = [];
      dcByGuid[guid].push({
        id:       String(row[dcH.indexOf('component_id')] || ''),
        score:    Number(row[dcH.indexOf('score')])        || 0,
        repaired: row[dcH.indexOf('repaired')] === true || row[dcH.indexOf('repaired')] === 'TRUE',
      });
    }
  }

  var penMax = {
    refroidissement: Number(cfg.PENALTY_REFROIDISSEMENT_MAX) || 12,
    direction:       Number(cfg.PENALTY_DIRECTION_MAX)       || 20,
    transmission:    Number(cfg.PENALTY_TRANSMISSION_MAX)    || 8,
    suspension:      Number(cfg.PENALTY_SUSPENSION_MAX)      || 8,
    chassis:         Number(cfg.PENALTY_CHASSIS_MAX)         || 20,
  };

  var cars   = [];
  var carIdx = 0;

  for (var i = 1; i < dsData.length; i++) {
    var drow = dsData[i];
    var dv   = (function(r, h) {
      return function(col) { var idx = h.indexOf(col); return idx >= 0 ? r[idx] : ''; };
    })(drow, dsHeaders);

    var guid  = String(dv('driver_guid') || '').trim();
    var model = String(dv('car_model')   || '').trim();
    if (!guid || model === 'tv_car') continue;

    // Calculer ballast/restrictor depuis damage_components (source de vérité)
    var comps        = dcByGuid[guid] || [];
    var totalBallast = 0;
    var restrValues  = [];

    comps.forEach(function(c) {
      if (c.repaired) return; // composant réparé → pas de pénalité
      var penType = PENALTY_TYPE[c.id] || 'ballast_kg';
      var penVal  = scoreToPenaltyValue(c.score, 0, penMax[c.id] || 0, cfg);
      if (penType === 'ballast_kg') totalBallast += penVal;
      else restrValues.push(penVal);
    });

    restrValues.sort(function(a,b) { return b-a; });
    var w2          = Number(cfg.RESTRICTOR_SECOND_WEIGHT) || 0.3;
    var totalRestr  = restrValues.length === 0 ? 0
                    : restrValues.length === 1 ? restrValues[0]
                    : restrValues[0] + restrValues[1] * w2;

    var ballastFinal    = Math.round(Math.min(totalBallast, 40));
    var restrictorFinal = Math.round(Math.min(totalRestr, 12));

    cars.push({
      BallastKG:  ballastFinal,
      CarId:      carIdx++,
      Driver: {
        Guid:      guid,
        GuidsList: [guid],
        Name:      (String(dv('driver_name') || '').trim() || nameByGuid[guid] || ''),
        Team:      String(dv('team')        || '').trim(),
        Nation:    ''
      },
      Model:      model,
      Restrictor: restrictorFinal,
      Skin:       String(dv('skin') || '').trim()
    });
  }

  return {
    Version:    7,
    Stage:      stageId,
    ExportedAt: new Date().toISOString(),
    Cars:       cars
  };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — CONFIG
// ──────────────────────────────────────────────────────────────

function getAdminsList() {
  ensureAdminsSheet();
  var sheet = getSheet('admins');
  var data  = sheet.getDataRange().getValues();
  var rows  = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    rows.push({
      steam_id:    String(data[i][0] || '').trim(),
      driver_name: String(data[i][1] || ''),
      active:      data[i][2] !== false && data[i][2] !== 'FALSE' && data[i][2] !== 0,
    });
  }
  return { ok: true, admins: rows };
}

function saveAdminsList(admins) {
  ensureAdminsSheet();
  var sheet = getSheet('admins');
  // Vider et réécrire (conserver en-tête)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  admins.forEach(function(a) {
    if (!a.steam_id) return;
    sheet.appendRow([a.steam_id.trim(), a.driver_name || '', a.active !== false]);
  });
  return { ok: true };
}

function getConfigForAdmin() {
  ensureConfigSheet();
  var sheet = getSheet('config');
  var data  = sheet.getDataRange().getValues();
  var rows  = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({ key: String(data[i][0] || ''), value: data[i][1], description: String(data[i][2] || '') });
  }
  return { ok: true, config: rows };
}

function saveConfigFromAdmin(configObj) {
  ensureConfigSheet();
  var sheet   = getSheet('config');
  var data    = sheet.getDataRange().getValues();
  var rowMap  = {};
  for (var i = 1; i < data.length; i++) rowMap[String(data[i][0])] = i + 1;
  Object.keys(configObj).forEach(function(key) {
    var val = configObj[key];
    if (rowMap[key]) sheet.getRange(rowMap[key], 2).setValue(val);
    else sheet.appendRow([key, val, CONFIG_DESCRIPTIONS[key] || '']);
  });
  return { ok: true };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — OVERRIDE RÉPARATIONS PAR PILOTE
// ──────────────────────────────────────────────────────────────

function adminGetDriverRepairs(guid, stageId) {
  var cfg      = readConfigSheet();
  var dcSheet  = getSheet('damage_components');
  if (!dcSheet) return { ok: true, components: [] };
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];
  var components = [];
  for (var i = 1; i < dcData.length; i++) {
    var row = dcData[i];
    if (String(row[dcHeaders.indexOf('driver_guid')]) !== String(guid)) continue;
    if (String(row[dcHeaders.indexOf('stage_id')])    !== String(stageId)) continue;
    var dc = (function(r) { return function(col) { var idx = dcHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
    var score   = Number(dc('score')) || 0;
    var repaired= dc('repaired') === true || dc('repaired') === 'TRUE';
    components.push({
      id:       String(dc('component_id')),
      score:    score,
      severity: scoreToSeverity(score, cfg),
      repaired: repaired,
      row:      i + 1,
    });
  }
  return { ok: true, components: components };
}

function adminSetRepair(guid, stageId, componentId, repaired) {
  var dcSheet  = getSheet('damage_components');
  if (!dcSheet) throw new Error('Feuille damage_components introuvable');
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];
  var guidIdx  = dcHeaders.indexOf('driver_guid');
  var stageIdx = dcHeaders.indexOf('stage_id');
  var compIdx  = dcHeaders.indexOf('component_id');
  var repIdx   = dcHeaders.indexOf('repaired');
  var found    = false;
  for (var i = 1; i < dcData.length; i++) {
    if (String(dcData[i][guidIdx])  === String(guid) &&
        String(dcData[i][stageIdx]) === String(stageId) &&
        String(dcData[i][compIdx])  === String(componentId)) {
      dcSheet.getRange(i + 1, repIdx + 1).setValue(repaired === true);
      found = true; break;
    }
  }
  if (!found) throw new Error('Composant introuvable : ' + componentId);
  // Recalculer les totaux driver_state
  _recalcDriverTotals(guid, stageId);
  return { ok: true };
}

function adminSetValidated(guid, stageId, validated) {
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var guidIdx  = dsHeaders.indexOf('driver_guid');
  var valIdx   = dsHeaders.indexOf('validated');
  for (var i = 1; i < dsData.length; i++) {
    if (String(dsData[i][guidIdx]) === String(guid)) {
      dsSheet.getRange(i + 1, valIdx + 1).setValue(validated === true);
      dsSheet.getRange(i + 1, dsHeaders.indexOf('last_updated') + 1).setValue(new Date().toISOString());
      return { ok: true };
    }
  }
  throw new Error('Pilote introuvable');
}

// Recalcule ballast/restrictor/repair_used dans driver_state depuis damage_components
function _recalcDriverTotals(guid, stageId) {
  var cfg      = readConfigSheet();
  var dcSheet  = getSheet('damage_components');
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];
  var penMax   = {
    refroidissement: Number(cfg.PENALTY_REFROIDISSEMENT_MAX) || 12,
    direction:       Number(cfg.PENALTY_DIRECTION_MAX)       || 20,
    transmission:    Number(cfg.PENALTY_TRANSMISSION_MAX)    || 8,
    suspension:      Number(cfg.PENALTY_SUSPENSION_MAX)      || 8,
    chassis:         Number(cfg.PENALTY_CHASSIS_MAX)         || 20,
  };
  var repairMax = {
    refroidissement: Number(cfg.REPAIR_COST_REFROIDISSEMENT) || 40,
    direction:       Number(cfg.REPAIR_COST_DIRECTION)       || 35,
    transmission:    Number(cfg.REPAIR_COST_TRANSMISSION)    || 30,
    suspension:      Number(cfg.REPAIR_COST_SUSPENSION)      || 20,
    chassis:         Number(cfg.REPAIR_COST_CHASSIS)         || 50,
  };
  var totalBallast = 0, restrValues = [], totalRepair = 0;
  for (var i = 1; i < dcData.length; i++) {
    var row = dcData[i];
    if (String(row[dcHeaders.indexOf('driver_guid')]) !== String(guid)) continue;
    if (String(row[dcHeaders.indexOf('stage_id')])    !== String(stageId)) continue;
    var compId  = String(row[dcHeaders.indexOf('component_id')]);
    var score   = Number(row[dcHeaders.indexOf('score')]) || 0;
    var repaired= row[dcHeaders.indexOf('repaired')] === true || row[dcHeaders.indexOf('repaired')] === 'TRUE';
    var repPct  = repaired ? 100 : 0;
    var penType = PENALTY_TYPE[compId] || 'ballast_kg';
    var penVal  = scoreToPenaltyValue(score, repPct, penMax[compId] || 0, cfg);
    var repCost = scoreToRepairCost(score, repPct, repairMax[compId] || 0, cfg);
    if (!repaired) {
      if (penType === 'ballast_kg') totalBallast += penVal;
      else restrValues.push(penVal);
      totalRepair += repCost;
    }
  }
  restrValues.sort(function(a,b) { return b-a; });
  var w2 = Number(cfg.RESTRICTOR_SECOND_WEIGHT) || 0.3;
  var totalRestr = restrValues.length === 0 ? 0
                 : restrValues.length === 1 ? restrValues[0]
                 : restrValues[0] + restrValues[1] * w2;
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  for (var m = 1; m < dsData.length; m++) {
    if (String(dsData[m][dsHeaders.indexOf('driver_guid')]) === String(guid)) {
      var r = m + 1;
      dsSheet.getRange(r, dsHeaders.indexOf('ballast_kg')     + 1).setValue(Math.min(totalBallast, 40));
      dsSheet.getRange(r, dsHeaders.indexOf('restrictor')     + 1).setValue(Math.min(totalRestr, 12));
      dsSheet.getRange(r, dsHeaders.indexOf('repair_used_min')+ 1).setValue(Math.ceil(totalRepair));
      dsSheet.getRange(r, dsHeaders.indexOf('last_updated')   + 1).setValue(new Date().toISOString());
      break;
    }
  }
}

// ──────────────────────────────────────────────────────────────
// ADMIN — PURGE BASE DE DONNÉES
// Vide driver_state, damage_components, stage_results
// Conserve : config, vehicle_profiles, championship_events
// ──────────────────────────────────────────────────────────────

function purgeDatabase() {
  var ss      = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  var purged  = [];
  var SHEETS_TO_PURGE = ['driver_state', 'damage_components', 'stage_results'];

  SHEETS_TO_PURGE.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1); // conserve la ligne d'en-têtes
    }
    purged.push(name);
  });

  return { ok: true, purged: purged };
}

// ──────────────────────────────────────────────────────────────
// ACSM FETCH — Sync direct depuis le serveur
// Script Properties requises :
//   ACSM_URL              : ex. https://acsm.revivalseries.fr
//   ACSM_CHAMPIONSHIP_ID  : UUID du championnat (ex. 7fd33cab-...)
// ──────────────────────────────────────────────────────────────

function getAcsmBase() {
  var url = PROPS.getProperty('ACSM_URL') || '';
  return url.replace(/\/$/, ''); // supprimer slash final
}

// Sync championnat complet depuis ACSM — teste plusieurs paths API
function syncAcsmChampionship() {
  var base    = getAcsmBase();
  var champId = PROPS.getProperty('ACSM_CHAMPIONSHIP_ID') || '';

  if (!base)    throw new Error('ACSM_URL non configuré dans Script Properties');
  if (!champId) throw new Error('ACSM_CHAMPIONSHIP_ID non configuré dans Script Properties');

  // ACSM v2.4 : essaie les paths connus
  var candidates = [
    base + '/api/championships/' + champId,
    base + '/api/championship/'  + champId,
    base + '/api/championships/' + champId + '/export',
  ];

  var response = null, usedUrl = '';
  for (var i = 0; i < candidates.length; i++) {
    try {
      var r = UrlFetchApp.fetch(candidates[i], { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) { response = r; usedUrl = candidates[i]; break; }
    } catch(e) {}
  }

  if (!response) {
    throw new Error('ACSM 404 sur tous les paths. URL de base : ' + base + ' | ID : ' + champId +
      '\nPaths testés : ' + candidates.join(' | '));
  }

  var data;
  try { data = JSON.parse(response.getContentText()); } catch(e) {
    throw new Error('Réponse non parseable depuis ' + usedUrl);
  }

  var result = importChampionship(data);
  result.source            = 'acsm_fetch';
  result.championship_name = data.Name || '';
  result.url_used          = usedUrl;
  return result;
}

// Fetch résultats de session depuis ACSM
// ACSM expose /api/results → liste des fichiers JSON de résultats
// On prend le plus récent pour l'étape courante
function fetchAcsmResults() {
  var base    = getAcsmBase();
  var stageId = PROPS.getProperty('CURRENT_STAGE_ID') || '';

  if (!base) throw new Error('ACSM_URL non configuré dans Script Properties');

  // 1. Récupérer la liste des résultats
  var listUrl  = base + '/api/results';
  var listResp = UrlFetchApp.fetch(listUrl, { muteHttpExceptions: true });
  if (listResp.getResponseCode() !== 200)
    throw new Error('Impossible de lister les résultats ACSM (' + listResp.getResponseCode() + ')');

  var resultsList;
  try { resultsList = JSON.parse(listResp.getContentText()); } catch(e) {
    throw new Error('Liste résultats non parseable');
  }

  // resultsList est un tableau de noms de fichiers ou d'objets
  // Format ACSM : tableau de strings ou [{Name, Date}]
  var files = [];
  if (Array.isArray(resultsList)) {
    resultsList.forEach(function(item) {
      if (typeof item === 'string') files.push(item);
      else if (item.Name)          files.push(item.Name);
      else if (item.Filename)      files.push(item.Filename);
    });
  }

  if (files.length === 0) throw new Error('Aucun résultat disponible sur ACSM');

  // 2. Prendre le fichier le plus récent (dernier de la liste, généralement trié par date)
  var latestFile = files[files.length - 1];

  // 3. Télécharger ce résultat
  var resultUrl  = base + '/results/download/' + latestFile;
  var resultResp = UrlFetchApp.fetch(resultUrl, { muteHttpExceptions: true });
  if (resultResp.getResponseCode() !== 200)
    throw new Error('Impossible de télécharger le résultat : ' + latestFile);

  var sessionData;
  try { sessionData = JSON.parse(resultResp.getContentText()); } catch(e) {
    throw new Error('Résultat non parseable');
  }

  // 4. Traiter comme un résultat de session JSON ACSM
  var result = importSessionJson(sessionData, stageId);
  result.source   = 'acsm_fetch';
  result.filename = latestFile;
  return result;
}

// Import d'un résultat de session JSON ACSM (format Results direct)
// Différent de importChampionship : c'est un résultat isolé, pas un championnat complet
function importSessionJson(sessionData, stageId) {
  var cfg             = readConfigSheet();
  var vehicleProfiles = loadVehicleProfiles();

  var cars      = sessionData.Cars   || [];
  var laps      = sessionData.Laps   || [];
  var events    = sessionData.Events || [];

  // Index guid → model/skin
  var guidToInfo = {};
  cars.forEach(function(car) {
    if (car.Driver && car.Driver.Guid) {
      guidToInfo[car.Driver.Guid] = { model: car.Model || '', skin: car.Skin || '',
                                       name: car.Driver.Name || '' };
    }
  });

  // Index carId → guid (pour les events)
  var carIdToGuid = {};
  cars.forEach(function(car) {
    if (car.CarId !== undefined && car.Driver && car.Driver.Guid)
      carIdToGuid[car.CarId] = car.Driver.Guid;
  });

  // Résultats (meilleur tour)
  var srSheet = getOrCreateSheet('stage_results',
    ['event_index','driver_guid','driver_name','car_model','skin','best_lap_ms','laps','position','class_id']);
  var srData = srSheet.getDataRange().getValues();
  var srH    = srData[0];
  for (var si = srData.length - 1; si >= 1; si--) {
    if (String(srData[si][srH.indexOf('event_index')]) === String(stageId)) srSheet.deleteRow(si + 1);
  }

  var bestLaps  = {}, lapCounts = {};
  laps.forEach(function(lap) {
    var guid = lap.DriverGuid;
    if (!guid || lap.LapTime > 1800000) return;
    if (!bestLaps[guid] || lap.LapTime < bestLaps[guid]) bestLaps[guid] = lap.LapTime;
    lapCounts[guid] = (lapCounts[guid] || 0) + 1;
  });

  var sorted = Object.keys(bestLaps).sort(function(a,b) { return bestLaps[a] - bestLaps[b]; });
  var srRows  = [];
  sorted.forEach(function(guid, pos) {
    var info = guidToInfo[guid] || {};
    srRows.push([stageId, guid, info.name || '', info.model || '', info.skin || '',
      bestLaps[guid], lapCounts[guid] || 0, pos + 1, '']);
  });
  if (srRows.length) srSheet.getRange(srSheet.getLastRow() + 1, 1, srRows.length, 9).setValues(srRows);

  // Dégâts depuis events (avec RelPosition — format JSON complet)
  var dcSheet = getOrCreateSheet('damage_components',
    ['driver_guid','stage_id','component_id','score','severity','ballast_kg','restrictor','repair_min','repaired']);
  var dcData = dcSheet.getDataRange().getValues();
  var dcH    = dcData[0];
  for (var di = dcData.length - 1; di >= 1; di--) {
    if (String(dcData[di][dcH.indexOf('stage_id')]) === String(stageId)) dcSheet.deleteRow(di + 1);
  }

  // Grouper impacts par guid
  var impactsByGuid = {};
  events.forEach(function(ev) {
    var guid = carIdToGuid[ev.CarId];
    if (!guid) return;
    if (!impactsByGuid[guid]) impactsByGuid[guid] = [];
    var rel = ev.RelPosition || {};
    impactsByGuid[guid].push({
      speed: ev.ImpactSpeed || 0,
      type:  (ev.Type || 'ENV').indexOf('ENV') !== -1 ? 'ENV' : 'CAR',
      rx: Number(rel.X || 0), ry: Number(rel.Y || 0), rz: Number(rel.Z || 0),
    });
  });

  var dcRows = [], totalDamage = 0;
  Object.keys(impactsByGuid).forEach(function(guid) {
    var info  = guidToInfo[guid] || {};
    var zones = vehicleProfiles[info.model || ''] || null;
    var scores= computeScoresFromImpacts(impactsByGuid[guid], zones, cfg);

    COMPONENTS.forEach(function(compId) {
      var score = scores[compId];
      if (score < (Number(cfg.SCORE_RIEN_MAX) || 5)) return;
      var sev     = scoreToSeverity(score, cfg);
      var penType = PENALTY_TYPE[compId] || 'ballast_kg';
      var penKey  = 'PENALTY_' + compId.toUpperCase() + '_MAX';
      var repKey  = 'REPAIR_COST_' + compId.toUpperCase();
      var penVal  = scoreToPenaltyValue(score, 0, Number(cfg[penKey]) || 0, cfg);
      var repCost = scoreToRepairCost(score, 0, Number(cfg[repKey]) || 0, cfg);
      dcRows.push([guid, stageId, compId, score, sev,
        penType === 'ballast_kg' ? penVal : 0,
        penType === 'restrictor' ? penVal : 0,
        Math.ceil(repCost), false]);
      totalDamage++;
    });
  });
  if (dcRows.length) dcSheet.getRange(dcSheet.getLastRow() + 1, 1, dcRows.length, 9).setValues(dcRows);

  return { ok: true, stage_id: stageId, results: srRows.length, damage_entries: totalDamage };
}

// ──────────────────────────────────────────────────────────────
// RESPONSE HELPER
// ──────────────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
