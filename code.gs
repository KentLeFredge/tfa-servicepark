// ============================================================
// TFA SERVICE PARK — Google Apps Script Backend
// Auth Steam OpenID 2.0 + Google Sheets
// ============================================================

var PROPS = PropertiesService.getScriptProperties();

// ──────────────────────────────────────────────────────────────
// DAMAGE ENGINE — Conversion collisions ACSM → composants
// Tous les seuils sont ajustables via Script Properties.
// ──────────────────────────────────────────────────────────────

/*
  PROPRIÉTÉS CONFIGURABLES (Script Properties) :
  ┌──────────────────────────────┬─────────┬────────────────────────────────────┐
  │ Propriété                    │ Défaut  │ Description                        │
  ├──────────────────────────────┼─────────┼────────────────────────────────────┤
  │ DMG_MIN_SPEED                │ 10      │ Vitesse min (km/h) pour dégât      │
  │ DMG_LEGER_MAX                │ 25      │ Seuil léger→modéré                 │
  │ DMG_MODERE_MAX               │ 50      │ Seuil modéré→sévère                │
  │ DMG_SEVERE_MAX               │ 80      │ Seuil sévère→critique              │
  │ DMG_BALLAST_LEGER            │ 5       │ Ballast (kg) sévérité légère       │
  │ DMG_BALLAST_MODERE           │ 10      │ Ballast sévérité modérée           │
  │ DMG_BALLAST_SEVERE           │ 20      │ Ballast sévérité sévère            │
  │ DMG_BALLAST_CRITIQUE         │ 30      │ Ballast sévérité critique          │
  │ DMG_RESTRICTOR_LEGER         │ 0       │ Restrictor sévérité légère         │
  │ DMG_RESTRICTOR_MODERE        │ 5       │ Restrictor sévérité modérée        │
  │ DMG_RESTRICTOR_SEVERE        │ 10      │ Restrictor sévérité sévère         │
  │ DMG_RESTRICTOR_CRITIQUE      │ 15      │ Restrictor sévérité critique       │
  │ DMG_REPAIR_LEGER             │ 8       │ Temps réparation (min) léger       │
  │ DMG_REPAIR_MODERE            │ 15      │ Temps réparation modéré            │
  │ DMG_REPAIR_SEVERE            │ 25      │ Temps réparation sévère            │
  │ DMG_REPAIR_CRITIQUE          │ 40      │ Temps réparation critique          │
  │ DMG_MERGE_MODE               │ worst   │ "worst" = pire choc / "sum" = cumul│
  └──────────────────────────────┴─────────┴────────────────────────────────────┘
*/

// Valeurs par défaut si la feuille config est absente ou incomplète
var CONFIG_DEFAULTS = {
  DMG_MIN_SPEED:           10,
  DMG_LEGER_MAX:           25,
  DMG_MODERE_MAX:          50,
  DMG_SEVERE_MAX:          80,
  DMG_BALLAST_LEGER:        5,
  DMG_BALLAST_MODERE:      10,
  DMG_BALLAST_SEVERE:      20,
  DMG_BALLAST_CRITIQUE:    30,
  DMG_RESTRICTOR_LEGER:     0,
  DMG_RESTRICTOR_MODERE:    5,
  DMG_RESTRICTOR_SEVERE:   10,
  DMG_RESTRICTOR_CRITIQUE: 15,
  DMG_REPAIR_LEGER:         8,
  DMG_REPAIR_MODERE:       15,
  DMG_REPAIR_SEVERE:       25,
  DMG_REPAIR_CRITIQUE:     40,
  DMG_MERGE_MODE:       'worst',
  REPAIR_BUDGET_MIN:       60,
};

var CONFIG_DESCRIPTIONS = {
  DMG_MIN_SPEED:           'Vitesse min (km/h) pour qu\'un choc cause des dégâts',
  DMG_LEGER_MAX:           'Seuil (km/h) léger → modéré',
  DMG_MODERE_MAX:          'Seuil (km/h) modéré → sévère',
  DMG_SEVERE_MAX:          'Seuil (km/h) sévère → critique',
  DMG_BALLAST_LEGER:       'Ballast (kg) — sévérité légère',
  DMG_BALLAST_MODERE:      'Ballast (kg) — sévérité modérée',
  DMG_BALLAST_SEVERE:      'Ballast (kg) — sévérité sévère',
  DMG_BALLAST_CRITIQUE:    'Ballast (kg) — sévérité critique',
  DMG_RESTRICTOR_LEGER:    'Restrictor — sévérité légère',
  DMG_RESTRICTOR_MODERE:   'Restrictor — sévérité modérée',
  DMG_RESTRICTOR_SEVERE:   'Restrictor — sévérité sévère',
  DMG_RESTRICTOR_CRITIQUE: 'Restrictor — sévérité critique',
  DMG_REPAIR_LEGER:        'Temps réparation (min) — légère',
  DMG_REPAIR_MODERE:       'Temps réparation (min) — modérée',
  DMG_REPAIR_SEVERE:       'Temps réparation (min) — sévère',
  DMG_REPAIR_CRITIQUE:     'Temps réparation (min) — critique',
  DMG_MERGE_MODE:          '"worst" = pire choc retenu / "sum" = cumul atténué',
  REPAIR_BUDGET_MIN:       'Budget temps de réparation alloué aux pilotes (min)',
};

// Lit la feuille config et retourne un objet key→value
function readConfigSheet() {
  var cfg = {};
  // Valeurs par défaut
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

// Crée la feuille config avec toutes les clés si elle n'existe pas
function ensureConfigSheet() {
  var ss    = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  var sheet = ss.getSheetByName('config');
  if (!sheet) {
    sheet = ss.insertSheet('config');
    sheet.appendRow(['key', 'value', 'description']);
    Object.keys(CONFIG_DEFAULTS).forEach(function(k) {
      sheet.appendRow([k, CONFIG_DEFAULTS[k], CONFIG_DESCRIPTIONS[k] || '']);
    });
    // Mise en forme basique
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 380);
  }
  return sheet;
}

function getDmgConfig() {
  var raw = readConfigSheet();
  return {
    minSpeed:  Number(raw.DMG_MIN_SPEED)  || 10,
    legerMax:  Number(raw.DMG_LEGER_MAX)  || 25,
    modereMax: Number(raw.DMG_MODERE_MAX) || 50,
    severeMax: Number(raw.DMG_SEVERE_MAX) || 80,
    ballast:   { leger:   Number(raw.DMG_BALLAST_LEGER)       || 5,
                 modere:  Number(raw.DMG_BALLAST_MODERE)      || 10,
                 severe:  Number(raw.DMG_BALLAST_SEVERE)      || 20,
                 critique:Number(raw.DMG_BALLAST_CRITIQUE)    || 30 },
    restrictor:{ leger:   Number(raw.DMG_RESTRICTOR_LEGER)    || 0,
                 modere:  Number(raw.DMG_RESTRICTOR_MODERE)   || 5,
                 severe:  Number(raw.DMG_RESTRICTOR_SEVERE)   || 10,
                 critique:Number(raw.DMG_RESTRICTOR_CRITIQUE) || 15 },
    repair:    { leger:   Number(raw.DMG_REPAIR_LEGER)        || 8,
                 modere:  Number(raw.DMG_REPAIR_MODERE)       || 15,
                 severe:  Number(raw.DMG_REPAIR_SEVERE)       || 25,
                 critique:Number(raw.DMG_REPAIR_CRITIQUE)     || 40 },
    mergeMode: String(raw.DMG_MERGE_MODE  || 'worst'),
    repairBudget: Number(raw.REPAIR_BUDGET_MIN) || 60,
  };
}

// Mappe RelPosition {X,Y,Z} → id composant
// Système de coordonnées AC : Z=avant, X=droite, Y=haut
function collisionToComponent(rel) {
  var x = rel.X || 0, y = rel.Y || 0, z = rel.Z || 0;
  // Avant prononcé → refroidissement (radiateur)
  if (z > 0.7)                          return 'refroidissement';
  // Avant latéral → direction
  if (z > 0.2 && Math.abs(x) > 0.25)   return 'direction';
  // Arrière → transmission
  if (z < -0.5)                         return 'transmission';
  // Côtés → suspension
  if (Math.abs(x) > 0.4)               return 'suspension';
  // Dessous → châssis
  if (y < -0.3)                         return 'chassis';
  // Centre/général → moteur
  return 'moteur';
}

// Vitesse → sévérité
function speedToSeverity(speed, cfg) {
  if (speed < cfg.minSpeed)  return null;         // ignoré
  if (speed < cfg.legerMax)  return 'leger';
  if (speed < cfg.modereMax) return 'modere';
  if (speed < cfg.severeMax) return 'severe';
  return 'critique';
}

// Sévérité → pénalités
function severityToPenalties(sev, cfg) {
  if (!sev || sev === 'intact' || sev === 'none') return { ballast_kg:0, restrictor:0, repair_min:0 };
  return {
    ballast_kg:  cfg.ballast[sev]   || 0,
    restrictor:  cfg.restrictor[sev]|| 0,
    repair_min:  cfg.repair[sev]    || 0,
  };
}

// Score (0–100) depuis vitesse d'impact : 100 = intact, 0 = destruction totale
function speedToScore(speed, cfg) {
  if (speed < cfg.minSpeed)  return 100;
  if (speed < cfg.legerMax)  return Math.round(100 - (speed - cfg.minSpeed) / (cfg.legerMax  - cfg.minSpeed) * 25);
  if (speed < cfg.modereMax) return Math.round(75  - (speed - cfg.legerMax)  / (cfg.modereMax - cfg.legerMax)  * 25);
  if (speed < cfg.severeMax) return Math.round(50  - (speed - cfg.modereMax) / (cfg.severeMax - cfg.modereMax) * 25);
  return Math.max(0, Math.round(25 - (speed - cfg.severeMax) / 20 * 25));
}

// Traite les collisions d'une session → damage_components par pilote
// collisions : tableau d'objets {Driver.Guid, ImpactSpeed, RelPosition, Type}
// retourne : { guid: { compId: { speed, severity, score, ...penalties } } }
function processCollisions(collisions, cfg) {
  var result = {};
  collisions.forEach(function(ev) {
    var guid  = ev.Driver ? ev.Driver.Guid : null;
    var speed = ev.ImpactSpeed || 0;
    var rel   = ev.RelPosition || { X:0, Y:0, Z:0 };
    var sev   = speedToSeverity(speed, cfg);
    if (!guid || !sev) return; // ignore chocs non-dommageables

    var compId = collisionToComponent(rel);
    if (!result[guid]) result[guid] = {};

    var existing = result[guid][compId];
    if (!existing) {
      result[guid][compId] = { speed: speed, severity: sev };
    } else if (cfg.mergeMode === 'sum') {
      // Accumulation : additionne les vitesses, recalcule sévérité
      var newSpeed = Math.min(existing.speed + speed * 0.5, 120); // atténue la somme
      result[guid][compId] = { speed: newSpeed, severity: speedToSeverity(newSpeed, cfg) || sev };
    } else {
      // Pire choc (défaut)
      if (speed > existing.speed) {
        result[guid][compId] = { speed: speed, severity: sev };
      }
    }
  });
  return result;
}

// Composants par défaut (100%, aucun dégât) si la feuille damage_components est vide
var DEFAULT_COMPONENTS = [
  { id:'refroidissement', name_fr:'Refroidissement', name_en:'Cooling',    icon:'🌡' },
  { id:'direction',       name_fr:'Direction',       name_en:'Steering',   icon:'🎮' },
  { id:'transmission',    name_fr:'Transmission',    name_en:'Transmission',icon:'⚙' },
  { id:'suspension',      name_fr:'Suspension',      name_en:'Suspension', icon:'🔧' },
  { id:'chassis',         name_fr:'Châssis',         name_en:'Chassis',    icon:'🛡' },
  { id:'moteur',          name_fr:'Moteur',          name_en:'Engine',     icon:'🔥' },
  { id:'aerodynamique',   name_fr:'Aérodynamique',   name_en:'Aero',       icon:'💨' },
];

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
  if (action === 'import_entry_list')     return withAdminToken(e, function() {
    return jsonResponse(importEntryList(body.cars || [], body.stage_id || PROPS.getProperty('CURRENT_STAGE_ID')));
  });
  if (action === 'import_championship')   return withAdminToken(e, function() {
    return jsonResponse(importChampionship(body));
  });
  if (action === 'import_session_csv')    return withAdminToken(e, function() {
    return jsonResponse(importSessionCsv(e.postData.contents || ''));
  });
  if (action === 'reset_stage')           return withAdminToken(e, function() {
    return jsonResponse(resetStage(body.stage_id || PROPS.getProperty('CURRENT_STAGE_ID')));
  });
  if (action === 'set_stage')             return withAdminToken(e, function() {
    PROPS.setProperty('CURRENT_STAGE_ID', body.stage_id);
    return jsonResponse({ ok: true, stage_id: body.stage_id });
  });
  if (action === 'save_config')           return withAdminToken(e, function() {
    return jsonResponse(saveConfigFromAdmin(body.config || {}));
  });
  return jsonResponse({ error: 'unknown_action' });
}

// ──────────────────────────────────────────────────────────────
// AUTH — TOKEN PILOTE
// ──────────────────────────────────────────────────────────────

function withAuth(e, fn) {
  if (!validateToken(e.parameter.token || '', e.parameter.steam_id || ''))
    return jsonResponse({ error: 'unauthorized' });
  try { return fn(e.parameter.steam_id); } catch(err) { return jsonResponse({ error: err.message }); }
}

// ──────────────────────────────────────────────────────────────
// AUTH — TOKEN ADMIN
// ──────────────────────────────────────────────────────────────

function withAdminToken(e, fn) {
  var token   = e.parameter.token    || '';
  var steamId = e.parameter.steam_id || '';
  if (!validateToken(token, steamId))    return jsonResponse({ error: 'unauthorized' });
  if (!isAdminSteamId(steamId))          return jsonResponse({ error: 'admin_forbidden' });
  try { return fn(steamId); } catch(err) { return jsonResponse({ error: err.message }); }
}

function isAdminSteamId(steamId) {
  var admins = PROPS.getProperty('ADMIN_STEAM_IDS') || '';
  return admins.split(',').map(function(s) { return s.trim(); }).indexOf(String(steamId)) !== -1;
}

// ──────────────────────────────────────────────────────────────
// STEAM OPENID AUTH
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

  return HtmlService.createHtmlOutput(
    '<script>window.location.href=' + JSON.stringify(redirectUrl) + ';</script>'
  );
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
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var driverRow= null;
  for (var i = 1; i < dsData.length; i++) {
    if (String(dsData[i][dsHeaders.indexOf('driver_guid')]) === String(steamId)) { driverRow = dsData[i]; break; }
  }
  if (!driverRow) throw new Error('driver_not_found');
  function dsVal(col) { var idx = dsHeaders.indexOf(col); return idx >= 0 ? driverRow[idx] : null; }

  // Composants de dégâts
  var dcSheet  = getSheet('damage_components');
  var components = [];
  if (dcSheet) {
    var dcData   = dcSheet.getDataRange().getValues();
    var dcHeaders= dcData[0];
    for (var j = 1; j < dcData.length; j++) {
      var row = dcData[j];
      if (String(row[dcHeaders.indexOf('driver_guid')]) === String(steamId) &&
          String(row[dcHeaders.indexOf('stage_id')])    === String(stageId)) {
        var dc = (function(r) { return function(col) { var idx = dcHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
        components.push({
          id: String(dc('component_id')), score: Number(dc('score')) || 0,
          severity: String(dc('severity') || 'none'),
          ballast_kg: Number(dc('ballast_kg')) || 0, restrictor: Number(dc('restrictor')) || 0,
          repair_min: Number(dc('repair_min')) || 0,
          repaired: dc('repaired') === true || dc('repaired') === 'TRUE'
        });
      }
    }
  }

  // Si aucun composant → afficher tous les composants à 100% (aucun dégât)
  if (components.length === 0) {
    components = DEFAULT_COMPONENTS.map(function(c) {
      return { id: c.id, name_fr: c.name_fr, name_en: c.name_en, icon: c.icon,
               score: 100, severity: 'intact', ballast_kg: 0, restrictor: 0, repair_min: 0, repaired: false };
    });
  }

  // Résultats de l'étape courante (depuis stage_results)
  var stageResults  = getStageResults(stageId);
  var driverResult  = null;
  var stageIdx      = stageIdToIndex(stageId);

  for (var k = 0; k < stageResults.length; k++) {
    if (String(stageResults[k].driver_guid) === String(steamId)) {
      driverResult = stageResults[k];
      break;
    }
  }

  // Étape suivante
  var nextStage = getNextStageInfo(stageId);

  return {
    driver_guid:       String(steamId),
    driver_name:       String(dsVal('driver_name') || ''),
    car_model:         String(dsVal('car_model')   || ''),
    skin:              String(dsVal('skin')         || ''),
    stage_id:          String(stageId),
    validated:         dsVal('validated') === true || dsVal('validated') === 'TRUE',
    repair_budget_min: 60,
    repair_used_min:   Number(dsVal('repair_used_min')) || 0,
    ballast_kg:        Number(dsVal('ballast_kg'))      || 0,
    restrictor:        Number(dsVal('restrictor'))      || 0,
    penalty_seconds:   Number(dsVal('penalty_seconds')) || 0,
    components:        components,
    // Résultats étape
    best_lap_ms:  driverResult ? driverResult.best_lap_ms  : null,
    stage_pos:    driverResult ? driverResult.position     : null,
    stage_results: stageResults,
    next_stage:    nextStage,
  };
}

// ──────────────────────────────────────────────────────────────
// STAGE RESULTS
// ──────────────────────────────────────────────────────────────

function stageIdToIndex(stageId) {
  // Essaie de trouver dans championship_events
  var evSheet = getSheet('championship_events');
  if (!evSheet) return -1;
  var data    = evSheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('event_id');
  var idxIdx  = headers.indexOf('event_index');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(stageId) ||
        'event_' + data[i][idxIdx] === stageId) return Number(data[i][idxIdx]);
  }
  return -1;
}

function getStageResults(stageId) {
  var srSheet = getSheet('stage_results');
  if (!srSheet) return [];
  var data    = srSheet.getDataRange().getValues();
  var headers = data[0];
  var evIdx   = headers.indexOf('event_index');
  var curIdx  = stageIdToIndex(stageId);
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][evIdx]) === curIdx) {
      var row = data[i];
      var dv  = (function(r) { return function(col) { var idx = headers.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
      results.push({
        driver_guid:  String(dv('driver_guid')  || ''),
        driver_name:  String(dv('driver_name')  || ''),
        car_model:    String(dv('car_model')     || ''),
        skin:         String(dv('skin')          || ''),
        best_lap_ms:  Number(dv('best_lap_ms'))  || 0,
        laps:         Number(dv('laps'))         || 0,
        position:     Number(dv('position'))     || 0,
        class_id:     String(dv('class_id')      || ''),
      });
    }
  }
  return results.sort(function(a, b) { return a.position - b.position; });
}

function getNextStageInfo(stageId) {
  var evSheet = getSheet('championship_events');
  if (!evSheet) return null;
  var data    = evSheet.getDataRange().getValues();
  var headers = data[0];
  var idxIdx  = headers.indexOf('event_index');
  var curIdx  = stageIdToIndex(stageId);
  if (curIdx < 0) return null;

  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][idxIdx]) === curIdx + 1) {
      var row = data[i];
      var dv  = (function(r) { return function(col) { var idx = headers.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
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
  var dcSheet  = getSheet('damage_components');
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];
  var guidIdx  = dcHeaders.indexOf('driver_guid'), stageIdx = dcHeaders.indexOf('stage_id');
  var compIdIdx= dcHeaders.indexOf('component_id'), repairedIdx = dcHeaders.indexOf('repaired');
  var existingRows = {};
  for (var i = 1; i < dcData.length; i++) {
    if (String(dcData[i][guidIdx]) === String(steamId) && String(dcData[i][stageIdx]) === String(stageId))
      existingRows[String(dcData[i][compIdIdx])] = i + 1;
  }
  var totalBallast = 0, totalRestrictor = 0, totalRepairMin = 0;
  for (var k = 0; k < components.length; k++) {
    var comp = components[k]; var repaired = comp.repaired === true;
    if (!repaired) {
      totalBallast    += Number(comp.ballast_kg)  || 0;
      totalRestrictor += Number(comp.restrictor)  || 0;
      totalRepairMin  += Number(comp.repair_min)  || 0;
    }
    if (existingRows[String(comp.id)]) dcSheet.getRange(existingRows[String(comp.id)], repairedIdx + 1).setValue(repaired);
  }
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  for (var m = 1; m < dsData.length; m++) {
    if (String(dsData[m][dsHeaders.indexOf('driver_guid')]) === String(steamId)) {
      var r = m + 1;
      dsSheet.getRange(r, dsHeaders.indexOf('ballast_kg')    + 1).setValue(totalBallast);
      dsSheet.getRange(r, dsHeaders.indexOf('restrictor')    + 1).setValue(totalRestrictor);
      dsSheet.getRange(r, dsHeaders.indexOf('repair_used_min')+ 1).setValue(totalRepairMin);
      dsSheet.getRange(r, dsHeaders.indexOf('last_updated')  + 1).setValue(new Date().toISOString());
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
  var guidIdx  = dsHeaders.indexOf('driver_guid'), validatedIdx = dsHeaders.indexOf('validated');
  for (var i = 1; i < dsData.length; i++) {
    if (String(dsData[i][guidIdx]) === String(steamId)) {
      if (dsData[i][validatedIdx] === true || dsData[i][validatedIdx] === 'TRUE') throw new Error('already_validated');
      dsSheet.getRange(i + 1, validatedIdx + 1).setValue(true);
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
      severity:     String(row[dcHeaders.indexOf('severity')]  || 'none'),
      ballast_kg:   Number(row[dcHeaders.indexOf('ballast_kg')]) || 0,
      restrictor:   Number(row[dcHeaders.indexOf('restrictor')]) || 0,
      repair_min:   Number(row[dcHeaders.indexOf('repair_min')]) || 0,
      repaired:     row[dcHeaders.indexOf('repaired')] === true || row[dcHeaders.indexOf('repaired')] === 'TRUE'
    });
  }

  var stageResults = getStageResults(stageId);
  var resultsByGuid = {};
  stageResults.forEach(function(r) { resultsByGuid[r.driver_guid] = r; });

  var drivers = [];
  for (var i = 1; i < dsData.length; i++) {
    var drow = dsData[i];
    var dv   = (function(r) { return function(col) { var idx = dsHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(drow);
    var guid = String(dv('driver_guid'));
    var comps= dcIndex[guid] || [];
    var res  = resultsByGuid[guid] || null;
    drivers.push({
      driver_guid:     guid,
      driver_name:     String(dv('driver_name')     || ''),
      car_model:       String(dv('car_model')        || ''),
      validated:       dv('validated') === true || dv('validated') === 'TRUE',
      ballast_kg:      Number(dv('ballast_kg'))      || 0,
      restrictor:      Number(dv('restrictor'))      || 0,
      repair_used_min: Number(dv('repair_used_min')) || 0,
      last_updated:    String(dv('last_updated')     || ''),
      damaged_count:   comps.filter(function(c) { return c.severity !== 'none' && c.severity !== 'intact'; }).length,
      repaired_count:  comps.filter(function(c) { return c.repaired; }).length,
      components:      comps,
      best_lap_ms:     res ? res.best_lap_ms : null,
      stage_pos:       res ? res.position    : null,
    });
  }
  return { stage_id: stageId, drivers: drivers };
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

  var inserted = 0, updated = 0;
  for (var k = 0; k < cars.length; k++) {
    var car    = cars[k];
    var driver = car.Driver || car.driver || {};
    var guid   = String(driver.Guid   || driver.guid   || '');
    var name   = String(driver.Name   || driver.name   || '');
    var model  = String(car.Model     || car.model     || '');
    var skin   = String(car.Skin      || car.skin      || '');
    if (!guid || guid === 'undefined') continue;

    if (existingGuids[guid]) {
      var row = existingGuids[guid];
      dsSheet.getRange(row, dsHeaders.indexOf('driver_name') + 1).setValue(name);
      dsSheet.getRange(row, dsHeaders.indexOf('car_model')   + 1).setValue(model);
      if (dsHeaders.indexOf('skin') >= 0) dsSheet.getRange(row, dsHeaders.indexOf('skin') + 1).setValue(skin);
      updated++;
    } else {
      var rowData = dsHeaders.map(function(h) {
        var map = { driver_guid: guid, driver_name: name, car_model: model,
          stage_id: stageId, validated: false, repair_used_min: 0,
          ballast_kg: 0, restrictor: 0, penalty_seconds: 0,
          last_updated: new Date().toISOString(), skin: skin };
        return map[h] !== undefined ? map[h] : '';
      });
      dsSheet.appendRow(rowData);
      inserted++;
    }
  }
  return { ok: true, inserted: inserted, updated: updated, total: cars.length };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — IMPORT CHAMPIONSHIP JSON
// ──────────────────────────────────────────────────────────────

function importChampionship(data) {
  var ss = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));

  // ── Feuille championship_events ──
  var evHeaders = ['event_index','event_id','event_name','track','layout',
                   'weather_ambient','weather_road','weather_wind','cmwfx_type','race_duration_min'];
  var evSheet = ss.getSheetByName('championship_events');
  if (!evSheet) evSheet = ss.insertSheet('championship_events');
  evSheet.clearContents();
  evSheet.appendRow(evHeaders);

  // ── Feuille stage_results ──
  var srHeaders = ['event_index','driver_guid','driver_name','car_model','skin',
                   'best_lap_ms','laps','position','class_id'];
  var srSheet = ss.getSheetByName('stage_results');
  if (!srSheet) srSheet = ss.insertSheet('stage_results');
  srSheet.clearContents();
  srSheet.appendRow(srHeaders);

  var events      = data.Events || [];
  var totalResults= 0;
  var totalDamage = 0;
  ensureConfigSheet(); // crée la feuille config avec les défauts si absente
  var cfg         = getDmgConfig();

  // ── Feuille damage_components (vider les entrées de ce championnat) ──
  var dcSheet  = getOrCreateSheet('damage_components',
    ['driver_guid','stage_id','component_id','score','severity','ballast_kg','restrictor','repair_min','repaired']);
  // Récupère les event IDs de ce championnat pour nettoyer les anciennes entrées
  var champEventIds = (data.Events || []).map(function(e, i) { return e.ID || ('event_' + i); });
  var dcData = dcSheet.getDataRange().getValues();
  var dcHeaders = dcData[0];
  var dcStageIdx = dcHeaders.indexOf('stage_id');
  // Supprime en partant du bas pour ne pas décaler les indices
  for (var di = dcData.length - 1; di >= 1; di--) {
    if (champEventIds.indexOf(String(dcData[di][dcStageIdx])) !== -1) {
      dcSheet.deleteRow(di + 1);
    }
  }

  for (var i = 0; i < events.length; i++) {
    var event  = events[i];
    var setup  = event.RaceSetup || {};

    // Trouver la météo de la course (Session "RACE")
    var raceWeather = null;
    var weatherMap  = setup.Weather || {};
    Object.keys(weatherMap).forEach(function(wKey) {
      var w = weatherMap[wKey];
      if (w.Sessions && w.Sessions.indexOf('RACE') !== -1) raceWeather = w;
    });
    // Fallback : première météo
    if (!raceWeather) {
      var wKeys = Object.keys(weatherMap);
      if (wKeys.length > 0) raceWeather = weatherMap[wKeys[0]];
    }

    var sessions     = setup.Sessions || {};
    var raceDuration = sessions.RACE ? (sessions.RACE.Time || 0) : 0;

    evSheet.appendRow([
      i,
      event.ID || '',
      event.Name || '',
      setup.Track || '',
      setup.TrackLayout || '',
      raceWeather ? raceWeather.BaseTemperatureAmbient : '',
      raceWeather ? raceWeather.BaseTemperatureRoad    : '',
      raceWeather ? Math.round((raceWeather.WindBaseSpeedMin + raceWeather.WindBaseSpeedMax) / 2) : '',
      raceWeather ? (raceWeather.CMWFXType || 0)       : '',
      raceDuration
    ]);

    // Résultats de la course RACE
    var eventSessions = event.Sessions || {};
    var raceSess      = eventSessions.RACE || {};
    var raceResults   = raceSess.Results   || {};
    var laps          = raceResults.Laps   || [];
    var cars          = raceResults.Cars   || [];

    if (laps.length === 0) continue;

    // Index Skin depuis Cars
    var skinMap = {};
    var classMap = {};
    cars.forEach(function(car) {
      if (car.Driver) {
        skinMap[car.Driver.Guid]  = car.Skin || '';
        classMap[car.Driver.Guid] = car.ClassID || '';
      }
    });

    // Meilleur tour par pilote (sans coupure de préférence)
    var bestLaps  = {};
    var lapCounts = {};
    laps.forEach(function(lap) {
      var guid = lap.DriverGuid;
      if (!guid) return;
      var time = lap.LapTime;
      // Ignorer les tours invalides (> 30 min, probablement erreur)
      if (time > 1800000) return;
      if (!bestLaps[guid] || time < bestLaps[guid].time) {
        bestLaps[guid] = { time: time, name: lap.DriverName || '', car: lap.CarModel || '' };
      }
      lapCounts[guid] = (lapCounts[guid] || 0) + 1;
    });

    // Trier par meilleur tour
    var sorted = Object.keys(bestLaps).sort(function(a, b) {
      return bestLaps[a].time - bestLaps[b].time;
    });

    var rowsBatch = [];
    sorted.forEach(function(guid, pos) {
      var d = bestLaps[guid];
      rowsBatch.push([
        i, guid, d.name, d.car,
        skinMap[guid]  || '',
        d.time,
        lapCounts[guid] || 0,
        pos + 1,
        classMap[guid] || ''
      ]);
    });

    if (rowsBatch.length > 0) {
      srSheet.getRange(srSheet.getLastRow() + 1, 1, rowsBatch.length, srHeaders.length).setValues(rowsBatch);
      totalResults += rowsBatch.length;
    }

    // ── Calcul des dégâts depuis collisions RACE ──
    var raceEvents    = raceResults.Events || [];
    if (raceEvents.length > 0) {
      var dmgByDriver = processCollisions(raceEvents, cfg);
      var eventStageId = event.ID || ('event_' + i);

      // Mapping guid → (car_model, skin) depuis Cars
      var carInfo = {};
      cars.forEach(function(car) {
        if (car.Driver) carInfo[car.Driver.Guid] = { car: car.Model || '', skin: car.Skin || '' };
      });

      Object.keys(dmgByDriver).forEach(function(guid) {
        var driverDmg = dmgByDriver[guid];
        Object.keys(driverDmg).forEach(function(compId) {
          var hit  = driverDmg[compId];
          var sev  = hit.severity || 'none';
          var pen  = severityToPenalties(sev, cfg);
          var score= speedToScore(hit.speed, cfg);
          dcSheet.appendRow([
            guid,
            eventStageId,
            compId,
            score,
            sev,
            pen.ballast_kg,
            pen.restrictor,
            pen.repair_min,
            false  // repaired
          ]);
          totalDamage++;
        });
      });
    }
  }

  return { ok: true, events: events.length, results: totalResults, damage_entries: totalDamage };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — CONFIG (lecture/écriture feuille config)
// ──────────────────────────────────────────────────────────────

function getConfigForAdmin() {
  ensureConfigSheet();
  var sheet   = getSheet('config');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows    = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      key:         String(data[i][0] || ''),
      value:       data[i][1],
      description: String(data[i][2] || ''),
    });
  }
  return { ok: true, config: rows };
}

function saveConfigFromAdmin(configObj) {
  // configObj = { key: value, ... }
  ensureConfigSheet();
  var sheet   = getSheet('config');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var keyIdx  = 0; // colonne A = key
  var valIdx  = 1; // colonne B = value

  // Index des lignes existantes
  var rowMap = {};
  for (var i = 1; i < data.length; i++) {
    rowMap[String(data[i][keyIdx])] = i + 1;
  }

  Object.keys(configObj).forEach(function(key) {
    var val = configObj[key];
    if (rowMap[key]) {
      sheet.getRange(rowMap[key], valIdx + 1).setValue(val);
    } else {
      sheet.appendRow([key, val, CONFIG_DESCRIPTIONS[key] || '']);
    }
  });

  return { ok: true };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — IMPORT SESSION CSV (stracker/ACSM export)
// ──────────────────────────────────────────────────────────────

function importSessionCsv(csvText) {
  var cfg     = getDmgConfig();
  var stageId = PROPS.getProperty('CURRENT_STAGE_ID');

  // ── 1. Construire l'index nom → GUID depuis driver_state ──
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var nameToGuid = {};
  var nameToRow  = {};
  for (var i = 1; i < dsData.length; i++) {
    var n = String(dsData[i][dsHeaders.indexOf('driver_name')] || '').toLowerCase().trim();
    var g = String(dsData[i][dsHeaders.indexOf('driver_guid')] || '');
    if (n && g) { nameToGuid[n] = g; nameToRow[g] = i + 1; }
  }

  function findGuid(name) {
    var n = (name || '').toLowerCase().trim();
    if (nameToGuid[n]) return nameToGuid[n];
    // Recherche partielle si nom exact non trouvé
    var keys = Object.keys(nameToGuid);
    for (var k = 0; k < keys.length; k++) {
      if (keys[k].indexOf(n) !== -1 || n.indexOf(keys[k]) !== -1) return nameToGuid[keys[k]];
    }
    return null;
  }

  var lines = csvText.split('\n');

  // ── 2. Parser "Race result" → positions + meilleur tour ──
  var inRaceResult = false;
  var stageResults = []; // { name, pos, laps, total_time, best_lap_ms }

  for (var li = 0; li < lines.length; li++) {
    var line = lines[li].trim();
    if (line.indexOf('Race result') !== -1) { inRaceResult = true; continue; }
    if (inRaceResult && line.indexOf('Race laps') !== -1) { inRaceResult = false; break; }
    if (!inRaceResult) continue;
    // Ligne de résultat : "1", "", "Team", "Vehicle", "Driver", "Laps", "Time", "Best lap", ...
    var m = line.match(/^"(\d+)",\s*"[^"]*",\s*"[^"]*",\s*"[^"]*",\s*"([^"]+)",\s*"(\d+)",\s*"([^"]+)",\s*"([^"]+)"/);
    if (!m) continue;
    var pos      = Number(m[1]);
    var drvName  = m[2].trim();
    var laps     = Number(m[3]);
    var bestLap  = m[5].trim(); // format MM:SS.mmm
    stageResults.push({ name: drvName, pos: pos, laps: laps, best_lap_str: bestLap });
  }

  // ── 3. Parser collisions ──
  // Format : "", "Driver reported contact with environment/another vehicle X. Impact speed: 12.34",
  var collisionsByName = {}; // name → [{ type, speed }]
  var colRe = /"", "(.+?) reported contact with (environment|another vehicle .+?)\. Impact speed: ([\d.]+)"/;

  for (var lj = 0; lj < lines.length; lj++) {
    var mc = lines[lj].match(colRe);
    if (!mc) continue;
    var dName = mc[1].trim();
    var cType = mc[2].trim();
    var speed = parseFloat(mc[3]);
    if (!collisionsByName[dName]) collisionsByName[dName] = [];
    collisionsByName[dName].push({ type: cType, speed: speed });
  }

  // ── 4. Calculer dégâts par pilote ──
  // Sans RelPosition : environment → avant (refroidissement/chassis), vehicle → côté (suspension/direction)
  function assignComponent(contactType, speed, cfg) {
    var isEnv = contactType === 'environment';
    if (isEnv) {
      // Choc mur : avant de la voiture
      if (speed >= cfg.severeMax)  return 'refroidissement'; // très fort → radiateur détruit
      if (speed >= cfg.modereMax)  return 'chassis';         // fort → châssis
      if (speed >= cfg.legerMax)   return 'direction';       // modéré → direction
      return 'direction';
    } else {
      // Choc voiture : côté
      if (speed >= cfg.severeMax)  return 'chassis';
      if (speed >= cfg.modereMax)  return 'suspension';
      return 'suspension';
    }
  }

  var dmgByName = {}; // name → { compId: { speed, severity } }
  Object.keys(collisionsByName).forEach(function(name) {
    var hits = collisionsByName[name];
    dmgByName[name] = {};
    hits.forEach(function(hit) {
      var sev = speedToSeverity(hit.speed, cfg);
      if (!sev) return;
      var compId   = assignComponent(hit.type, hit.speed, cfg);
      var existing = dmgByName[name][compId];
      if (!existing || hit.speed > existing.speed) {
        dmgByName[name][compId] = { speed: hit.speed, severity: sev };
      }
    });
  });

  // ── 5. Écrire stage_results ──
  var srSheet = getOrCreateSheet('stage_results',
    ['event_index','driver_guid','driver_name','car_model','skin','best_lap_ms','laps','position','class_id']);

  // Supprimer les entrées existantes pour ce stageId
  var srData = srSheet.getDataRange().getValues();
  var srHeaders = srData[0];
  var srStageIdx = srHeaders.indexOf('event_index');
  for (var si = srData.length - 1; si >= 1; si--) {
    if (String(srData[si][srStageIdx]) === String(stageId)) srSheet.deleteRow(si + 1);
  }

  var inserted = 0;
  stageResults.forEach(function(r) {
    var guid    = findGuid(r.name);
    var bestMs  = timeStrToMs(r.best_lap_str);
    // Récupérer car_model depuis driver_state
    var carModel = '';
    if (guid && nameToRow[guid]) {
      carModel = String(dsData[nameToRow[guid] - 1][dsHeaders.indexOf('car_model')] || '');
    }
    srSheet.appendRow([stageId, guid || '', r.name, carModel, '', bestMs, r.laps, r.pos, '']);
    inserted++;
  });

  // ── 6. Écrire damage_components ──
  var dcSheet = getOrCreateSheet('damage_components',
    ['driver_guid','stage_id','component_id','score','severity','ballast_kg','restrictor','repair_min','repaired']);

  // Supprimer les entrées existantes pour ce stageId
  var dcData = dcSheet.getDataRange().getValues();
  var dcHeaders = dcData[0];
  var dcStageIdx = dcHeaders.indexOf('stage_id');
  for (var di = dcData.length - 1; di >= 1; di--) {
    if (String(dcData[di][dcStageIdx]) === String(stageId)) dcSheet.deleteRow(di + 1);
  }

  var dmgInserted = 0;
  var dcRows = [];
  Object.keys(dmgByName).forEach(function(name) {
    var guid = findGuid(name);
    if (!guid) return;
    var driverDmg = dmgByName[name];
    Object.keys(driverDmg).forEach(function(compId) {
      var hit = driverDmg[compId];
      var sev = hit.severity;
      var pen = severityToPenalties(sev, cfg);
      var score = speedToScore(hit.speed, cfg);
      dcRows.push([guid, stageId, compId, score, sev, pen.ballast_kg, pen.restrictor, pen.repair_min, false]);
      dmgInserted++;
    });
  });
  if (dcRows.length > 0) {
    dcSheet.getRange(dcSheet.getLastRow() + 1, 1, dcRows.length, 9).setValues(dcRows);
  }

  return {
    ok:             true,
    stage_id:       stageId,
    results:        inserted,
    damage_entries: dmgInserted,
    unmatched:      stageResults.filter(function(r) { return !findGuid(r.name); }).map(function(r) { return r.name; }),
  };
}

// Convertit "MM:SS.mmm" → millisecondes
function timeStrToMs(str) {
  if (!str || str === '-') return 0;
  str = str.trim().replace(/'+/, '');
  var parts = str.split(':');
  if (parts.length === 2) {
    var mins = parseFloat(parts[0]) || 0;
    var secs = parseFloat(parts[1]) || 0;
    return Math.round((mins * 60 + secs) * 1000);
  }
  return 0;
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
  var dcSheet  = getSheet('damage_components');
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
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];
  var cars = [];
  for (var i = 1; i < dsData.length; i++) {
    var row = dsData[i];
    var dv  = (function(r) { return function(col) { var idx = dsHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
    var guid = String(dv('driver_guid') || '');
    if (!guid) continue;
    cars.push({
      BallastKG:  Number(dv('ballast_kg'))  || 0,
      CarId:      i - 1,
      Driver:     { Guid: guid, GuidsList: [guid], Name: String(dv('driver_name') || ''), Nation: '', Team: '' },
      Model:      String(dv('car_model')    || ''),
      Restrictor: Number(dv('restrictor'))  || 0,
      Skin:       String(dv('skin')         || '')
    });
  }
  return { Version: 7, Stage: PROPS.getProperty('CURRENT_STAGE_ID') || '', ExportedAt: new Date().toISOString(), Cars: cars };
}

// ──────────────────────────────────────────────────────────────
// RESPONSE HELPER
// ──────────────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
