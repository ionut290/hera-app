/**
 * Varga Cantieri - ponte di scrittura Google Sheet.
 * Riceve solo richieste firmate con la proprietà di script SYNC_SECRET.
 */

function doGet() {
  return jsonResponse_({ ok: true, service: "Varga Cantieri Google Sheet Sync" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    verifySecret_(body.secret);
    if (body.action === "createRegistrySpreadsheet") {
      if (!lock.tryLock(30000)) throw new Error("Creazione foglio già in corso. Riprova tra poco.");
      return createRegistrySpreadsheet_(body);
    }
    if (body.action === "syncRegistrySpreadsheet") {
      verifySpreadsheetAllowed_(body.spreadsheetId);
      if (!lock.tryLock(30000)) throw new Error("Foglio occupato. Riprova tra poco.");
      return syncRegistrySpreadsheet_(body);
    }
    if (!Array.isArray(body.headers) || !body.headers.length) throw new Error("Intestazioni mancanti.");
    if (body.action === "createSpreadsheet") {
      if (!lock.tryLock(30000)) throw new Error("Creazione foglio già in corso. Riprova tra poco.");
      return createSpreadsheet_(body);
    }
    verifySpreadsheetAllowed_(body.spreadsheetId);
    if (body.action !== "replaceRows") throw new Error("Azione non supportata.");
    if (!Array.isArray(body.rows)) throw new Error("Righe mancanti.");
    if (body.rows.length > 20000) throw new Error("Massimo 20.000 righe per sincronizzazione.");
    if (!lock.tryLock(30000)) throw new Error("Foglio occupato. Riprova tra poco.");

    var spreadsheet = SpreadsheetApp.openById(String(body.spreadsheetId || ""));
    var sheet = getSheetByGid_(spreadsheet, String(body.gid || "0"));
    var headers = body.headers.map(safeCell_);
    var rows = body.rows.map(function(row) {
      var source = Array.isArray(row) ? row : [];
      return headers.map(function(_, index) { return safeCell_(source[index]); });
    });

    ensureSheetSize_(sheet, Math.max(2, rows.length + 1), headers.length);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    writeRowsInChunks_(sheet, rows, headers.length);
    sheet.setFrozenRows(1);
    if (headers.length >= 2) {
      try { sheet.hideColumns(1, 2); } catch (_) { /* colonne già nascoste */ }
    }
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#0f766e")
      .setFontColor("#ffffff")
      .setWrap(true);
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      rowsWritten: rows.length,
      sheetName: sheet.getName(),
      operationId: String(body.operationId || "")
    });
  } catch (error) {
    return jsonResponse_({ ok: false, error: error && error.message ? error.message : String(error) });
  } finally {
    try { lock.releaseLock(); } catch (_) { /* lock non acquisito */ }
  }
}

function createRegistrySpreadsheet_(body) {
  var names = body.sheetNames || [];
  var spreadsheet = body.spreadsheetId ? SpreadsheetApp.openById(String(body.spreadsheetId)) : SpreadsheetApp.create("Varga Cantieri - Matrici dati");
  names.forEach(function(name, index) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = index === 0 && spreadsheet.getSheets().length === 1 && !spreadsheet.getSheets()[0].getLastRow() ? spreadsheet.getSheets()[0].setName(name) : spreadsheet.insertSheet(name);
    sheet.setFrozenRows(1);
  });
  return jsonResponse_({ ok: true, spreadsheetId: spreadsheet.getId(), sheetUrl: spreadsheet.getUrl(), sheetName: names[0] || "" });
}

function registryRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getDisplayValues(), headers = values.shift();
  return values.map(function(valuesRow) { var row = {}; headers.forEach(function(header, index) { row[header] = valuesRow[index]; }); return row; }).filter(function(row) { return String(row.RECORD_ID || "").trim(); });
}

function syncRegistrySpreadsheet_(body) {
  var spreadsheet = SpreadsheetApp.openById(String(body.spreadsheetId || ""));
  var incoming = {}, conflicts = [], written = 0;
  Object.keys(body.sheets || {}).forEach(function(name) {
    var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
    var remote = registryRows_(sheet), local = body.sheets[name] || [], byId = {};
    remote.forEach(function(row) { byId[String(row.RECORD_ID)] = row; });
    local.forEach(function(row) {
      var id = String(row.RECORD_ID || ""); if (!id) return;
      var old = byId[id], localVersion = Number(row.SYNC_VERSION) || 1, remoteVersion = Number(old && old.SYNC_VERSION) || 0;
      if (old && remoteVersion > localVersion && body.conflictPolicy !== "APP_WINS") {
        (incoming[name] || (incoming[name] = [])).push(old);
        conflicts.push({ sheet: name, recordId: id, resolution: "SHEET_NEWER" });
      } else {
        byId[id] = row; written += 1;
        if (old && remoteVersion === localVersion && String(old.UPDATED_AT || "") > String(row.UPDATED_AT || "") && body.conflictPolicy !== "APP_WINS") {
          byId[id] = old; (incoming[name] || (incoming[name] = [])).push(old);
          conflicts.push({ sheet: name, recordId: id, resolution: "LATEST_UPDATED_AT" });
        }
      }
    });
    // Le righe esistenti soltanto nel foglio sono importate, mai eliminate.
    remote.forEach(function(row) { if (!local.some(function(item) { return String(item.RECORD_ID) === String(row.RECORD_ID); })) (incoming[name] || (incoming[name] = [])).push(row); });
    var rows = Object.keys(byId).map(function(id) { return byId[id]; }), headerSet = {};
    rows.forEach(function(row) { Object.keys(row).forEach(function(key) { headerSet[key] = true; }); });
    local.forEach(function(row) { Object.keys(row).forEach(function(key) { headerSet[key] = true; }); });
    var technical = ["RECORD_ID", "UPDATED_AT", "UPDATED_BY", "SYNC_VERSION", "SYNC_SOURCE", "ROW_STATUS"];
    var headers = technical.concat(Object.keys(headerSet).filter(function(key) { return technical.indexOf(key) < 0; }));
    if (!headers.length) headers = technical;
    ensureSheetSize_(sheet, Math.max(2, rows.length + 1), headers.length); sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff").setWrap(true);
    writeRowsInChunks_(sheet, rows.map(function(row) { return headers.map(function(header) { return safeCell_(row[header]); }); }), headers.length); sheet.setFrozenRows(1);
  });
  var log = spreadsheet.getSheetByName("LOG_SINCRONIZZAZIONE");
  if (log) log.appendRow([new Date(), body.requestedByEmail || body.requestedByUid || "", body.registry || "", written, conflicts.length, "SYNC_OK"]);
  SpreadsheetApp.flush();
  return jsonResponse_({ ok: true, incoming: incoming, conflicts: conflicts, rowsWritten: written, spreadsheetId: spreadsheet.getId(), sheetUrl: spreadsheet.getUrl() });
}

function createSpreadsheet_(body) {
  var commessaId = String(body.commessaId || "").trim();
  if (!commessaId) throw new Error("ID commessa mancante.");
  var propertyKey = "COMMESSA_SHEET_" + digestKey_(commessaId);
  var properties = PropertiesService.getScriptProperties();
  var existingId = properties.getProperty(propertyKey);
  var spreadsheet;
  if (existingId) {
    try { spreadsheet = SpreadsheetApp.openById(existingId); } catch (_) { properties.deleteProperty(propertyKey); }
  }
  if (spreadsheet) {
    var existingSheet = spreadsheet.getSheets()[0];
    return jsonResponse_({ ok: true, spreadsheetId: spreadsheet.getId(), sheetUrl: spreadsheet.getUrl(), gid: String(existingSheet.getSheetId()), sheetName: existingSheet.getName(), alreadyExists: true });
  }
  var headers = body.headers.map(safeCell_);
  if (headers[0] !== "SYNC_KEY" || headers[1] !== "IMPIANTO_KEY") throw new Error("Colonne tecniche mancanti o non valide.");
  var safeName = String(body.commessaName || "Senza nome").replace(/[\\/]/g, "-").trim();
  spreadsheet = SpreadsheetApp.create("Varga Cantieri - " + safeName);
  properties.setProperty(propertyKey, spreadsheet.getId());
  var sheet = spreadsheet.getSheets()[0];
  ensureSheetSize_(sheet, 2, headers.length);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff").setWrap(true);
  sheet.setFrozenRows(1);
  try { sheet.hideColumns(1, 2); } catch (_) { /* colonne già nascoste */ }
  SpreadsheetApp.flush();
  return jsonResponse_({ ok: true, spreadsheetId: spreadsheet.getId(), sheetUrl: spreadsheet.getUrl(), gid: String(sheet.getSheetId()), sheetName: sheet.getName() });
}

function digestKey_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map(function(byte) { return (byte + 256).toString(16).slice(-2); }).join("").slice(0, 40);
}

function verifySecret_(provided) {
  var expected = PropertiesService.getScriptProperties().getProperty("SYNC_SECRET");
  if (!expected) throw new Error("SYNC_SECRET non configurato nello script.");
  if (String(provided || "") !== String(expected)) throw new Error("Richiesta non autorizzata.");
}

function verifySpreadsheetAllowed_(spreadsheetId) {
  var id = String(spreadsheetId || "").trim();
  if (!/^[A-Za-z0-9_-]+$/.test(id)) throw new Error("ID Google Sheet non valido.");
  var configured = PropertiesService.getScriptProperties().getProperty("ALLOWED_SPREADSHEET_IDS") || "";
  var allowed = configured.split(/[;,\s]+/).map(function(value) { return value.trim(); }).filter(String);
  if (allowed.length && allowed.indexOf(id) === -1) throw new Error("Google Sheet non autorizzato.");
}

function getSheetByGid_(spreadsheet, gid) {
  var numericGid = Number(gid);
  var sheets = spreadsheet.getSheets();
  for (var index = 0; index < sheets.length; index += 1) {
    if (sheets[index].getSheetId() === numericGid) return sheets[index];
  }
  return sheets[0];
}

function ensureSheetSize_(sheet, rows, columns) {
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < columns) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
}

function writeRowsInChunks_(sheet, rows, columnCount) {
  var chunkSize = 500;
  for (var start = 0; start < rows.length; start += chunkSize) {
    var chunk = rows.slice(start, start + chunkSize);
    sheet.getRange(start + 2, 1, chunk.length, columnCount).setValues(chunk);
  }
}

function safeCell_(value) {
  if (value === null || typeof value === "undefined") return "";
  if (value instanceof Date) return value;
  if (typeof value === "number" || typeof value === "boolean") return value;
  var text = String(value);
  // Evita formule iniettate da dati esterni senza alterare i normali valori.
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
