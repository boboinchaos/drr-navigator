/**
 * DRR Resource Navigator — free backend (Google Apps Script)
 *
 * What it does for the website:
 *  - Keeps the shared list of resources advocates add, "Resource info update" fixes,
 *    and which resources applied (in a Google Sheet called "DRR Navigator Data").
 *  - Saves intakes to Airtable (HHP base > Navigator Intakes).
 *  - Emails the PDF (or saves a Gmail draft) from the Gmail account that owns this script.
 *
 * Setup (one time, about 10 minutes):
 *  1. Project Settings (gear icon) > Script properties > add:
 *       AIRTABLE_TOKEN  your Airtable personal access token (airtable.com/create/tokens,
 *                       scope data.records:write, access to the HHP base)
 *  2. Pick "setup" in the function menu at the top and press Run. Allow the permissions.
 *  3. Deploy > New deployment > type "Web app" > Execute as: Me > Who has access: Anyone > Deploy.
 *     Copy the Web app URL (ends in /exec) and send it to Claude.
 */

var AIRTABLE_BASE  = 'appGhhxjRKBaz6j7z';
var AIRTABLE_TABLE = 'tblyBGFzwEXnuIuQV';
var COLLECTIONS    = ['resources', 'corrections', 'outcomes'];

function setup() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('DRR Navigator Data');
  props.setProperty('SHEET_ID', ss.getId());
  COLLECTIONS.forEach(function (c) {
    var sh = ss.getSheetByName(c) || ss.insertSheet(c);
    if (sh.getLastRow() === 0) sh.appendRow(['id', 'createdAt', 'data']);
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  // Touch Gmail and the network once so all permissions are granted now.
  GmailApp.getAliases();
  UrlFetchApp.getRequest('https://api.airtable.com');
  if (!props.getProperty('PASSCODE')) Logger.log('Add a PASSCODE script property before sharing the site.');
  if (!props.getProperty('AIRTABLE_TOKEN')) Logger.log('Add an AIRTABLE_TOKEN script property to save intakes to Airtable.');
  Logger.log('Ready. Data sheet: ' + ss.getUrl());
}

function sheet_(name) {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Run setup first');
  return SpreadsheetApp.openById(id).getSheetByName(name);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var p = e.parameter || {};
    if (p.action === 'list') {
      if (COLLECTIONS.indexOf(p.col) < 0) return out_({ ok: false, code: 'bad_request' });
      var sh = sheet_(p.col), n = sh.getLastRow();
      if (n < 2) return out_({ ok: true, docs: [] });
      var start = Math.max(2, n - 1999);
      var rows = sh.getRange(start, 1, n - start + 1, 3).getValues();
      var docs = [];
      for (var i = rows.length - 1; i >= 0; i--) {
        try { docs.push({ id: String(rows[i][0]), data: JSON.parse(rows[i][2]) }); } catch (x) {}
      }
      return out_({ ok: true, docs: docs });
    }
    return out_({ ok: true, service: 'DRR Resource Navigator' });
  } catch (err) {
    return out_({ ok: false, code: 'server_error', message: String(err) });
  }
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, code: 'bad_request' }); }
  var props = PropertiesService.getScriptProperties();
  // No passcode (Jason's choice, 2026-11-12). Daily caps stop abuse of the public link instead.
  function overCap_(kind, max) {
    var key = 'cap_' + kind + '_' + Utilities.formatDate(new Date(), 'America/Chicago', 'yyyyMMdd');
    var c = CacheService.getScriptCache(); var n = Number(c.get(key) || 0);
    if (n >= max) return true; c.put(key, String(n + 1), 21600); return false;
  }
  try {
    if (req.action === 'check') return out_({ ok: true });

    if (req.action === 'add') {
      if (overCap_('add', 300)) return out_({ ok: false, code: 'too_many' });
      if (COLLECTIONS.indexOf(req.col) < 0) return out_({ ok: false, code: 'bad_request' });
      var json = JSON.stringify(req.doc || {});
      if (json.length > 45000) return out_({ ok: false, code: 'too_big' });
      var lock = LockService.getScriptLock(); lock.waitLock(20000);
      try {
        var id = Utilities.getUuid().slice(0, 18);
        sheet_(req.col).appendRow([id, new Date().toISOString(), json]);
      } finally { lock.releaseLock(); }
      return out_({ ok: true, id: id });
    }

    if (req.action === 'airtable') {
      if (overCap_('airtable', 150)) return out_({ ok: false, code: 'too_many' });
      var token = props.getProperty('AIRTABLE_TOKEN');
      if (!token) return out_({ ok: false, code: 'no_airtable' });
      var res = UrlFetchApp.fetch('https://api.airtable.com/v0/' + AIRTABLE_BASE + '/' + AIRTABLE_TABLE, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ records: req.records, typecast: true })
      });
      var code = res.getResponseCode();
      if (code >= 300) return out_({ ok: false, code: 'airtable_error', message: res.getContentText().slice(0, 300) });
      return out_({ ok: true });
    }

    if (req.action === 'mail') {
      if (overCap_('mail', 60)) return out_({ ok: false, code: 'too_many' });
      var to = (req.to || []).join(','), opts = { name: req.fromName || 'DRR Resource Navigator' };
      if (!to) return out_({ ok: false, code: 'bad_request' });
      if (req.cc && req.cc.length) opts.cc = req.cc.join(',');
      opts.attachments = (req.attachments || []).map(function (a) {
        return Utilities.newBlob(Utilities.base64Decode(a.content), a.mimeType || 'application/pdf', a.filename || 'plan.pdf');
      });
      if (req.draft) GmailApp.createDraft(to, req.subject || 'Help Action Plan', req.body || '', opts);
      else GmailApp.sendEmail(to, req.subject || 'Help Action Plan', req.body || '', opts);
      return out_({ ok: true });
    }

    return out_({ ok: false, code: 'bad_request' });
  } catch (err) {
    return out_({ ok: false, code: 'server_error', message: String(err) });
  }
}
