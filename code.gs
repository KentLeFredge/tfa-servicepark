// ============================================================
// TFA SERVICE PARK — Google Apps Script Backend
// Auth Steam OpenID 2.0 + Google Sheets
// ============================================================

var PROPS = PropertiesService.getScriptProperties();

// ──────────────────────────────────────────────────────────────
// ENTRY POINTS
// ──────────────────────────────────────────────────────────────

function doGet(e) {
  var action = e.parameter.action || '';
  if (action === 'steam_callback')  return handleSteamCallback(e);
  if (action === 'get_driver')      return withAuth(e, function(steamId) {
    var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
    return jsonResponse(getDriverData(steamId, stageId));
  });
  if (action === 'save_repairs')    return withAuth(e, function(steamId) {
    var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
    saveRepairs(steamId, stageId, JSON.parse(e.parameter.components || '[]'));
    return jsonResponse({ ok: true });
  });
  if (action === 'validate')        return withAuth(e, function(steamId) {
    validateRepairs(steamId, e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'));
    return jsonResponse({ ok: true });
  });
  if (action === 'admin_overview')  return withAdminToken(e, function() {
    return jsonResponse(getAdminOverview());
  });
  if (action === 'export_entry_list') return withAdminToken(e, function() {
    return jsonResponse(exportEntryList());
  });
  return jsonResponse({ error: 'unknown_action' });
}

function doPost(e) {
  var action = e.parameter.action || '';
  if (action === 'save_repairs')    return withAuth(e, function(steamId) {
    var body = JSON.parse(e.postData.contents || '{}');
    saveRepairs(steamId, e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'), body.components || []);
    return jsonResponse({ ok: true });
  });
  if (action === 'validate')        return withAuth(e, function(steamId) {
    validateRepairs(steamId, e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID'));
    return jsonResponse({ ok: true });
  });
  if (action === 'import_entry_list') return withAdminToken(e, function() {
    var body = JSON.parse(e.postData.contents || '{}');
    return jsonResponse(importEntryList(body.cars || [], body.stage_id || PROPS.getProperty('CURRENT_STAGE_ID')));
  });
  if (action === 'reset_stage')     return withAdminToken(e, function() {
    var body = JSON.parse(e.postData.contents || '{}');
    return jsonResponse(resetStage(body.stage_id || PROPS.getProperty('CURRENT_STAGE_ID')));
  });
  if (action === 'set_stage')       return withAdminToken(e, function() {
    var body = JSON.parse(e.postData.contents || '{}');
    PROPS.setProperty('CURRENT_STAGE_ID', body.stage_id);
    return jsonResponse({ ok: true, stage_id: body.stage_id });
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
// AUTH — TOKEN ADMIN (token valide + steam_id dans ADMIN_STEAM_IDS)
// ──────────────────────────────────────────────────────────────

function withAdminToken(e, fn) {
  var token   = e.parameter.token    || '';
  var steamId = e.parameter.steam_id || '';
  if (!validateToken(token, steamId))       return jsonResponse({ error: 'unauthorized' });
  if (!isAdminSteamId(steamId))             return jsonResponse({ error: 'admin_forbidden' });
  try { return fn(steamId); } catch(err)  { return jsonResponse({ error: err.message }); }
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
  if (!/^\d{17}$/.test(steamId))       return HtmlService.createHtmlOutput('<h2>Auth failed: invalid Steam ID</h2>');
  if (!verifySteamSignature(e.parameters)) return HtmlService.createHtmlOutput('<h2>Auth failed: signature invalide</h2>');

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

  var dcSheet  = getSheet('damage_components');
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];
  var components = [];
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
  return {
    driver_guid: String(steamId), driver_name: String(dsVal('driver_name') || ''),
    car_model: String(dsVal('car_model') || ''), skin: String(dsVal('skin') || ''),
    stage_id: String(stageId),
    validated: dsVal('validated') === true || dsVal('validated') === 'TRUE',
    repair_budget_min: 45, repair_used_min: Number(dsVal('repair_used_min')) || 0,
    ballast_kg: Number(dsVal('ballast_kg')) || 0, restrictor: Number(dsVal('restrictor')) || 0,
    penalty_seconds: Number(dsVal('penalty_seconds')) || 0, components: components
  };
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
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];

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

  var drivers = [];
  for (var i = 1; i < dsData.length; i++) {
    var drow = dsData[i];
    var dv   = (function(r) { return function(col) { var idx = dsHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(drow);
    var guid = String(dv('driver_guid'));
    var comps= dcIndex[guid] || [];
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
      components:      comps
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
      if (dsHeaders.indexOf('skin') >= 0)
        dsSheet.getRange(row, dsHeaders.indexOf('skin') + 1).setValue(skin);
      updated++;
    } else {
      var rowData = dsHeaders.map(function(h) {
        var map = {
          driver_guid: guid, driver_name: name, car_model: model,
          stage_id: stageId, validated: false, repair_used_min: 0,
          ballast_kg: 0, restrictor: 0, penalty_seconds: 0,
          last_updated: new Date().toISOString(), skin: skin
        };
        return map[h] !== undefined ? map[h] : '';
      });
      dsSheet.appendRow(rowData);
      inserted++;
    }
  }
  return { ok: true, inserted: inserted, updated: updated, total: cars.length };
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
    dsSheet.getRange(r, dsHeaders.indexOf('stage_id')       + 1).setValue(stageId);
    dsSheet.getRange(r, dsHeaders.indexOf('validated')      + 1).setValue(false);
    dsSheet.getRange(r, dsHeaders.indexOf('repair_used_min')+ 1).setValue(0);
    dsSheet.getRange(r, dsHeaders.indexOf('ballast_kg')     + 1).setValue(0);
    dsSheet.getRange(r, dsHeaders.indexOf('restrictor')     + 1).setValue(0);
  }
  var dcSheet  = getSheet('damage_components');
  var dcData   = dcSheet.getDataRange().getValues();
  var dcHeaders= dcData[0];
  for (var j = 1; j < dcData.length; j++) {
    if (String(dcData[j][dcHeaders.indexOf('stage_id')]) === String(stageId))
      dcSheet.getRange(j + 1, dcHeaders.indexOf('repaired') + 1).setValue(false);
  }
  return { ok: true, stage_id: stageId };
}

// ──────────────────────────────────────────────────────────────
// ADMIN — EXPORT ENTRY LIST (format ACSM, ballast/restrictor mis à jour)
// ──────────────────────────────────────────────────────────────

function exportEntryList() {
  var dsSheet  = getSheet('driver_state');
  var dsData   = dsSheet.getDataRange().getValues();
  var dsHeaders= dsData[0];

  var cars = [];
  for (var i = 1; i < dsData.length; i++) {
    var row = dsData[i];
    var dv  = (function(r) { return function(col) { var idx = dsHeaders.indexOf(col); return idx >= 0 ? r[idx] : null; }; })(row);
    var guid     = String(dv('driver_guid')  || '');
    var name     = String(dv('driver_name')  || '');
    var model    = String(dv('car_model')    || '');
    var skin     = String(dv('skin')         || '');
    var ballast  = Number(dv('ballast_kg'))  || 0;
    var restrict = Number(dv('restrictor'))  || 0;
    if (!guid) continue;
    cars.push({
      BallastKG:  ballast,
      CarId:      i - 1,
      Driver: {
        Guid:      guid,
        GuidsList: [guid],
        Name:      name,
        Nation:    '',
        Team:      ''
      },
      Model:      model,
      Restrictor: restrict,
      Skin:       skin
    });
  }

  var stageId = PROPS.getProperty('CURRENT_STAGE_ID') || 'unknown';
  return {
    Version:  7,
    Stage:    stageId,
    ExportedAt: new Date().toISOString(),
    Cars:     cars
  };
}

// ──────────────────────────────────────────────────────────────
// RESPONSE HELPER
// ──────────────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
