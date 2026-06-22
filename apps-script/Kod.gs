/**
 * Tapo měřené zásuvky → Google Sheet (více zařízení)
 * Jednou denně projde Gmail, najde Tapo e-maily s exportem spotřeby,
 * podle "Device Name" v těle mailu zařadí data do správného listu
 * (Skimmer / Gardena / …), naparsuje hodinová data a zapíše je.
 *
 * Hlavní funkce: importTapo()  — spouští se denním triggerem (setupTrigger).
 */

// ─── Konfigurace ──────────────────────────────────────────────────────────────
var CONFIG = {
  GMAIL_QUERY: 'from:no-reply@email.tp-link.com subject:(Energy Monitoring Data) has:attachment newer_than:2d',
  ATTACHMENT_MATCH: 'Energy Usage',   // hledaná příloha (kWh data)
  LABEL_DONE: 'tapo-zpracovano',

  // Mapování: část názvu zařízení (z těla mailu "Device Name: …") → název listu.
  // Tapo do těla píše např. "Device Name: Skimmer zásuvka".
  // Přidej sem další zásuvky, jak budeš měřit víc spotřebičů.
  DEVICE_MAP: [
    { match: 'skimmer', sheet: 'Skimmer' },
    { match: 'gardena', sheet: 'Gardena' },
    { match: 'dešťov', sheet: 'Gardena' },   // kdyby zásuvka měla v názvu "dešťovka"
    { match: 'destov', sheet: 'Gardena' },
  ],
  SHEET_FALLBACK: 'Ostatni',   // když Device Name nesedí na žádné pravidlo
};

// ─── Hlavní funkce ────────────────────────────────────────────────────────────
function importTapo() {
  var threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, 20);
  if (!threads.length) { Logger.log('Žádný nový Tapo e-mail.'); return; }

  var label = getOrCreateLabel_(CONFIG.LABEL_DONE);
  var processed = 0;

  threads.forEach(function (thread) {
    var labels = thread.getLabels().map(function (l) { return l.getName(); });
    if (labels.indexOf(CONFIG.LABEL_DONE) !== -1) return;

    thread.getMessages().forEach(function (msg) {
      var deviceName = extractDeviceName_(msg.getPlainBody());
      var sheetName = mapDeviceToSheet_(deviceName);

      msg.getAttachments().forEach(function (att) {
        if ((att.getName() || '').indexOf(CONFIG.ATTACHMENT_MATCH) === -1) return;
        var rows = parseXlsAttachment_(att);
        if (rows.length) {
          writeRows_(sheetName, rows);
          processed++;
          Logger.log('Zařízení "' + deviceName + '" → list ' + sheetName + ' (' + rows.length + ' řádků)');
        }
      });
    });
    thread.addLabel(label);
  });

  if (!processed) Logger.log('Nic nového ke zpracování.');
}

// ─── Device Name z těla mailu ─────────────────────────────────────────────────
function extractDeviceName_(body) {
  if (!body) return '';
  var m = body.match(/Device\s*Name\s*:\s*(.+)/i);
  return m ? m[1].trim() : '';
}

function mapDeviceToSheet_(deviceName) {
  var lower = (deviceName || '').toLowerCase();
  for (var i = 0; i < CONFIG.DEVICE_MAP.length; i++) {
    if (lower.indexOf(CONFIG.DEVICE_MAP[i].match) !== -1) return CONFIG.DEVICE_MAP[i].sheet;
  }
  return CONFIG.SHEET_FALLBACK;
}

// ─── Parsování .xls přes konverzi na Google Sheet (řeší binární BIFF) ─────────
function parseXlsAttachment_(att) {
  var out = [], tempId = null;
  try {
    var blob = att.copyBlob();
    var f = Drive.Files.insert({ title: 'tapo_tmp_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true });
    tempId = f.id;
    var ss = SpreadsheetApp.openById(tempId);
    var sheet = ss.getSheetByName('Day') || ss.getSheets()[0];
    var values = sheet.getDataRange().getValues();
    values.forEach(function (r) {
      if (!r || r.length < 2) return;
      var m = String(r[0]).trim().match(/^(\d{4})\/(\d{2})\/(\d{2})\s+(\d{1,2}):(\d{2}):(\d{2})\s*$/);
      if (!m) return;
      var kwh = parseFloat(String(r[1]).replace(',', '.'));
      if (isNaN(kwh)) return;
      out.push([m[1] + '-' + m[2] + '-' + m[3], parseInt(m[4], 10), kwh]);
    });
  } catch (e) {
    Logger.log('Chyba parsování: ' + e);
  } finally {
    if (tempId) { try { Drive.Files.remove(tempId); } catch (e2) {} }
  }
  return out;
}

// ─── Zápis do listu s deduplikací podle datum+hodina ──────────────────────────
function writeRows_(sheetName, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); sheet.appendRow(['datum', 'hodina', 'kWh']); }

  var existing = {};
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) existing[data[i][0] + '|' + data[i][1]] = i + 1;

  var toAppend = [];
  rows.forEach(function (r) {
    var key = r[0] + '|' + r[1];
    if (existing[key]) sheet.getRange(existing[key], 3).setValue(r[2]);
    else { toAppend.push(r); existing[key] = true; }
  });
  if (toAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 3).setValues(toAppend);
  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  }
}

// ─── Pomocné ──────────────────────────────────────────────────────────────────
function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

// Spusť JEDNOU ručně: vytvoří denní trigger (~6:00).
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importTapo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importTapo').timeBased().everyDays(1).atHour(6).create();
  Logger.log('Denní trigger nastaven na ~6:00.');
}
