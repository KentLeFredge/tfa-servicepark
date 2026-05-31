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

  // Steam callback after OpenID redirect
  if (action === 'steam_callback') {
    return handleSteamCallback(e);
  }

  // Read driver data (requires valid token)
  if (action === 'get_driver') {
    return withAuth(e, function(steamId) {
      var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
      var data = getDriverData(steamId, stageId);
      return jsonResponse(data);
    });
  }

  // Save repair choices (idempotent)
  if (action === 'save_repairs') {
    return withAuth(e, function(steamId) {
      var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
      var components = JSON.parse(e.parameter.components || '[]');
      saveRepairs(steamId, stageId, components);
      return jsonResponse({ ok: true });
    });
  }

  // Validate (irreversible)
  if (action === 'validate') {
    return withAuth(e, function(steamId) {
      var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
      validateRepairs(steamId, stageId);
      return jsonResponse({ ok: true });
    });
  }

  return jsonResponse({ error: 'unknown_action' }, 400);
}

function doPost(e) {
  var action = e.parameter.action || '';

  if (action === 'save_repairs') {
    return withAuth(e, function(steamId) {
      var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
      var body = JSON.parse(e.postData.contents || '{}');
      var components = body.components || [];
      saveRepairs(steamId, stageId, components);
      return jsonResponse({ ok: true });
    });
  }

  if (action === 'validate') {
    return withAuth(e, function(steamId) {
      var stageId = e.parameter.stage_id || PROPS.getProperty('CURRENT_STAGE_ID');
      validateRepairs(steamId, stageId);
      return jsonResponse({ ok: true });
    });
  }

  return jsonResponse({ error: 'unknown_action' }, 400);
}

// ──────────────────────────────────────────────────────────────
// STEAM OPENID AUTH
// ──────────────────────────────────────────────────────────────

function handleSteamCallback(e) {
  var params = e.parameters; // all params as arrays

  // 1. Extract steamid64 from openid.claimed_id
  // Format: https://steamcommunity.com/openid/id/76561198XXXXXXXXX
  var claimedId = e.parameter['openid.claimed_id'] || '';
  var steamId = claimedId.replace(/.*\//, '');

  if (!/^\d{17}$/.test(steamId)) {
    return HtmlService.createHtmlOutput('<h2>Auth failed: invalid Steam ID</h2>');
  }

  // 2. Validate signature with Steam
  var valid = verifySteamSignature(params);
  if (!valid) {
    return HtmlService.createHtmlOutput('<h2>Auth failed: signature invalide</h2>');
  }

  // 3. Generate token: HMAC:timestamp
  var token = generateToken(steamId);

  // 4. Redirect to service park with token
  var serviceUrl = PROPS.getProperty('SERVICE_PARK_URL');
  var redirectUrl = serviceUrl + '?token=' + encodeURIComponent(token) + '&steam_id=' + encodeURIComponent(steamId);

  return HtmlService.createHtmlOutput(
    '<script>window.location.href = ' + JSON.stringify(redirectUrl) + ';</script>'
  );
}

function verifySteamSignature(params) {
  // Rebuild the POST body for check_authentication
  // All openid.* params sent back as-is, with mode changed
  var postParams = [];
  for (var key in params) {
    if (key.indexOf('openid.') === 0) {
      var val = params[key][0]; // arrays → first element
      if (key === 'openid.mode') {
        val = 'check_authentication';
      }
      postParams.push(encodeURIComponent(key) + '=' + encodeURIComponent(val));
    }
  }
  var postBody = postParams.join('&');

  var response = UrlFetchApp.fetch('https://steamcommunity.com/openid/login', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: postBody,
    muteHttpExceptions: true
  });

  var text = response.getContentText();
  return text.indexOf('is_valid:true') !== -1;
}

// ──────────────────────────────────────────────────────────────
// TOKEN MANAGEMENT
// ──────────────────────────────────────────────────────────────

function generateToken(steamId) {
  var secret = PROPS.getProperty('STEAM_SECRET');
  var timestamp = Math.floor(Date.now() / 1000).toString();
  var payload = steamId + ':' + timestamp;
  var rawHmac = Utilities.computeHmacSha256Signature(payload, secret);
  var hmacHex = rawHmac.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
  return hmacHex + ':' + timestamp;
}

function validateToken(token, steamId) {
  if (!token || !steamId) return false;
  var parts = token.split(':');
  if (parts.length !== 2) return false;
  var hmacHex = parts[0];
  var timestamp = parts[1];

  // Check expiry (4 hours)
  var now = Math.floor(Date.now() / 1000);
  if (now - parseInt(timestamp, 10) > 14400) return false;

  // Recompute HMAC
  var secret = PROPS.getProperty('STEAM_SECRET');
  var payload = steamId + ':' + timestamp;
  var rawHmac = Utilities.computeHmacSha256Signature(payload, secret);
  var expectedHex = rawHmac.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');

  return hmacHex === expectedHex;
}

// Auth middleware helper
function withAuth(e, fn) {
  var token = e.parameter.token || '';
  var steamId = e.parameter.steam_id || '';
  if (!validateToken(token, steamId)) {
    return jsonResponse({ error: 'unauthorized' }, 401);
  }
  try {
    return fn(steamId);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ──────────────────────────────────────────────────────────────
// SHEETS HELPERS
// ──────────────────────────────────────────────────────────────

function getSheet(name) {
  var sheetId = PROPS.getProperty('SHEET_ID');
  var ss = SpreadsheetApp.openById(sheetId);
  return ss.getSheetByName(name);
}

// Returns object keyed by first column value
function sheetToMap(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    map[data[i][0]] = { row: i + 1, data: row };
  }
  return { headers: headers, map: map, raw: data };
}

function findRow(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) return i + 1;
  }
  return -1;
}

// ──────────────────────────────────────────────────────────────
// DRIVER DATA
// ──────────────────────────────────────────────────────────────

function getDriverData(steamId, stageId) {
  var dsSheet = getSheet('driver_state');
  var dcSheet = getSheet('damage_components');

  // --- driver_state ---
  var dsData = dsSheet.getDataRange().getValues();
  var dsHeaders = dsData[0];
  var driverRow = null;
  for (var i = 1; i < dsData.length; i++) {
    if (String(dsData[i][dsHeaders.indexOf('driver_guid')]) === String(steamId)) {
      driverRow = dsData[i];
      break;
    }
  }

  if (!driverRow) {
    throw new Error('driver_not_found:' + steamId);
  }

  function dsVal(col) {
    var idx = dsHeaders.indexOf(col);
    return idx >= 0 ? driverRow[idx] : null;
  }

  // --- damage_components for this driver + stage ---
  var dcData = dcSheet.getDataRange().getValues();
  var dcHeaders = dcData[0];
  var components = [];

  for (var j = 1; j < dcData.length; j++) {
    var row = dcData[j];
    var rowGuid = String(row[dcHeaders.indexOf('driver_guid')]);
    var rowStage = String(row[dcHeaders.indexOf('stage_id')]);
    if (rowGuid === String(steamId) && rowStage === String(stageId)) {
      function dcVal(col) {
        var idx = dcHeaders.indexOf(col);
        return idx >= 0 ? row[idx] : null;
      }
      components.push({
        id:          String(dcVal('component_id')),
        score:       Number(dcVal('score')) || 0,
        severity:    String(dcVal('severity') || 'none'),
        ballast_kg:  Number(dcVal('ballast_kg')) || 0,
        restrictor:  Number(dcVal('restrictor')) || 0,
        repair_min:  Number(dcVal('repair_min')) || 0,
        repaired:    dcVal('repaired') === true || dcVal('repaired') === 'TRUE'
      });
    }
  }

  return {
    driver_guid:       String(steamId),
    driver_name:       String(dsVal('driver_name') || ''),
    car_model:         String(dsVal('car_model') || ''),
    skin:              '',
    stage_id:          String(stageId),
    validated:         dsVal('validated') === true || dsVal('validated') === 'TRUE',
    repair_budget_min: 45,
    repair_used_min:   Number(dsVal('repair_used_min')) || 0,
    ballast_kg:        Number(dsVal('ballast_kg')) || 0,
    restrictor:        Number(dsVal('restrictor')) || 0,
    penalty_seconds:   Number(dsVal('penalty_seconds')) || 0,
    components:        components
  };
}

// ──────────────────────────────────────────────────────────────
// SAVE REPAIRS
// ──────────────────────────────────────────────────────────────

function saveRepairs(steamId, stageId, components) {
  var dcSheet = getSheet('damage_components');
  var dcData = dcSheet.getDataRange().getValues();
  var dcHeaders = dcData[0];

  var guidIdx      = dcHeaders.indexOf('driver_guid');
  var stageIdx     = dcHeaders.indexOf('stage_id');
  var compIdIdx    = dcHeaders.indexOf('component_id');
  var repairedIdx  = dcHeaders.indexOf('repaired');

  // Index existing rows for this driver+stage: key = component_id → sheet row number
  var existingRows = {};
  for (var i = 1; i < dcData.length; i++) {
    if (String(dcData[i][guidIdx]) === String(steamId) &&
        String(dcData[i][stageIdx]) === String(stageId)) {
      existingRows[String(dcData[i][compIdIdx])] = i + 1;
    }
  }

  var totalBallast = 0;
  var totalRestrictor = 0;
  var totalRepairMin = 0;

  for (var k = 0; k < components.length; k++) {
    var comp = components[k];
    var compId = String(comp.id);
    var repaired = comp.repaired === true;

    // Effective penalties: if repaired, no ballast/restrictor/time contribution
    if (!repaired) {
      totalBallast    += Number(comp.ballast_kg) || 0;
      totalRestrictor += Number(comp.restrictor) || 0;
      totalRepairMin  += Number(comp.repair_min) || 0;
    }

    // Update repaired flag in damage_components
    if (existingRows[compId]) {
      dcSheet.getRange(existingRows[compId], repairedIdx + 1).setValue(repaired);
    }
    // If component row doesn't exist, skip (rows created by race result import)
  }

  // Update driver_state totals
  var dsSheet = getSheet('driver_state');
  var dsData = dsSheet.getDataRange().getValues();
  var dsHeaders = dsData[0];

  var dsGuidIdx       = dsHeaders.indexOf('driver_guid');
  var dsBallastIdx    = dsHeaders.indexOf('ballast_kg');
  var dsRestrictIdx   = dsHeaders.indexOf('restrictor');
  var dsRepairUsedIdx = dsHeaders.indexOf('repair_used_min');
  var dsLastUpdIdx    = dsHeaders.indexOf('last_updated');

  for (var m = 1; m < dsData.length; m++) {
    if (String(dsData[m][dsGuidIdx]) === String(steamId)) {
      var rowNum = m + 1;
      dsSheet.getRange(rowNum, dsBallastIdx + 1).setValue(totalBallast);
      dsSheet.getRange(rowNum, dsRestrictIdx + 1).setValue(totalRestrictor);
      dsSheet.getRange(rowNum, dsRepairUsedIdx + 1).setValue(totalRepairMin);
      dsSheet.getRange(rowNum, dsLastUpdIdx + 1).setValue(new Date().toISOString());
      break;
    }
  }
}

// ──────────────────────────────────────────────────────────────
// VALIDATE REPAIRS
// ──────────────────────────────────────────────────────────────

function validateRepairs(steamId, stageId) {
  var dsSheet = getSheet('driver_state');
  var dsData = dsSheet.getDataRange().getValues();
  var dsHeaders = dsData[0];

  var guidIdx      = dsHeaders.indexOf('driver_guid');
  var validatedIdx = dsHeaders.indexOf('validated');
  var lastUpdIdx   = dsHeaders.indexOf('last_updated');

  for (var i = 1; i < dsData.length; i++) {
    if (String(dsData[i][guidIdx]) === String(steamId)) {
      if (dsData[i][validatedIdx] === true || dsData[i][validatedIdx] === 'TRUE') {
        throw new Error('already_validated');
      }
      dsSheet.getRange(i + 1, validatedIdx + 1).setValue(true);
      dsSheet.getRange(i + 1, lastUpdIdx + 1).setValue(new Date().toISOString());
      return;
    }
  }
  throw new Error('driver_not_found');
}

// ──────────────────────────────────────────────────────────────
// RESPONSE HELPERS
// ──────────────────────────────────────────────────────────────

function jsonResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
