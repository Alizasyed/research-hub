/**
 * PDA Field Data Hub
 * A Google Apps Script web app where the team adds field data (recordings,
 * note photos, typed notes). A background queue transcribes it (ElevenLabs or
 * Gemini) and writes summaries against the study's seven assumptions.
 * Everything is saved in one shared Google Drive folder.
 *
 * Script properties (Project Settings > Script properties):
 *   GEMINI_API_KEY  Key from Google AI Studio (free tier works).
 *   GEMINI_MODEL    optional. Defaults to gemini-3.8-flash; switches itself if Google retires a model.
 *   ELEVENLABS_API_KEY  optional. Adds ElevenLabs Scribe as a transcription option.
 *   ROOT_FOLDER_ID, SHEET_ID, PROC_FOLDER_ID  written by setup().
 */

var DEFAULT_MODEL = 'gemini-3.8-flash';
var HEADERS = ['id', 'created', 'addedBy', 'date', 'respondent', 'kind', 'segment',
  'interviewer', 'location', 'language', 'status', 'claimedBy', 'folderId', 'files',
  'transcriptDocId', 'summaryDocId', 'done', 'updated', 'notes', 'jobs'];
var PART_SEC = 300;          // audio sent per transcription call
var RUN_BUDGET_MS = 270000;  // stop picking up work after 4.5 minutes per run

var ASSUMPTIONS = [
  'Wants a loan? (free, documentless self-registration produces real credit demand)',
  'Government route? (merchants will register through a government portal rather than suppliers, payment providers or POS companies)',
  'Software-ready file? (an automated process can make a loan file bank-ready)',
  'Bank approval? (banks will approve a high share of prepared applications)',
  'Data sharing? (merchants will share the transaction, inventory and repayment data the assessment needs)',
  'Real business? (registrations represent live, distinct businesses, not duplicates or dormant ones)',
  'Loan size? (the planned average loan size matches what merchants need)'
];

var JOURNEY = ['Set up and run the business', 'Register and formalise', 'Sell and get paid',
  'Buy stock', 'Need money', 'Decide how to borrow', 'Apply',
  'Bank assessment and decision', 'Receive the money and repay'];

// ---------- web app ----------

// ---------- web API for the website (GitHub Pages) ----------
// Deploy as a web app: Execute as Me, Who has access: Anyone.
// Every call must carry the team code saved in the TEAM_CODE script property.

var USER_ = '';
function me_() {
  if (USER_) return USER_;
  try { return Session.getActiveUser().getEmail(); } catch (e) { return ''; }
}

var API = ['getConfig', 'listSessions', 'createSession', 'updateSession', 'addFile', 'addTypedNotes', 'claim', 'setStatus',
  'enqueue', 'requestSummary', 'retryJob', 'cancelJob', 'processQueue', 'saveText', 'uploadUrl', 'ping'];

function doGet() {
  return ContentService.createTextOutput('Field Data Hub API is running. Open the website to use it.');
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var code = PropertiesService.getScriptProperties().getProperty('TEAM_CODE');
    if (!code) throw new Error('Add a TEAM_CODE script property first (any word the team will type to sign in).');
    if (String(req.code || '').trim().toLowerCase() !== code.trim().toLowerCase()) throw new Error('WRONG_CODE');
    if (API.indexOf(req.fn) < 0) throw new Error('Unknown action ' + req.fn);
    USER_ = String(req.who || '').trim().slice(0, 60) || 'Someone';
    out = { ok: true, data: globalThis[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function ping() { return 'ok'; }

/** Start a Drive upload; the browser then sends the bytes straight to Drive. */
function uploadUrl(name, mime, folderId, origin) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true&fields=id,name,mimeType,size', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime || 'application/octet-stream', Origin: origin },
    payload: JSON.stringify({ name: name, parents: [folderId] })
  });
  if (res.getResponseCode() !== 200) throw new Error('Drive upload could not start (' + res.getResponseCode() + ').');
  var h = res.getAllHeaders();
  return h.Location || h.location;
}

/** Run once from the editor. Creates the shared folder and the log sheet. */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty('ROOT_FOLDER_ID');
  var root = rootId ? DriveApp.getFolderById(rootId) : DriveApp.createFolder('PDA Field Data');
  props.setProperty('ROOT_FOLDER_ID', root.getId());
  var sheetId = props.getProperty('SHEET_ID');
  if (!sheetId) {
    var ss = SpreadsheetApp.create('PDA Field Data log');
    var sh = ss.getSheets()[0];
    sh.setName('Sessions');
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    DriveApp.getFileById(ss.getId()).moveTo(root);
    props.setProperty('SHEET_ID', ss.getId());
  }
  procFolder_();
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'processQueue'; });
  if (!has) ScriptApp.newTrigger('processQueue').timeBased().everyMinutes(5).create();
  Logger.log('Background queue: on (checks every 5 minutes).');
  Logger.log('Folder: ' + root.getUrl());
  Logger.log('Share this folder with the team as Editors.');
  if (!props.getProperty('GEMINI_API_KEY')) Logger.log('Add GEMINI_API_KEY under Project Settings > Script properties.');
}

function getConfig() {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty('ROOT_FOLDER_ID');
  if (!rootId) throw new Error('The hub is not set up yet. Run setup() in the script editor.');
  return {
    me: me_(),
    rootUrl: 'https://drive.google.com/drive/folders/' + rootId,
    hasKey: !!props.getProperty('GEMINI_API_KEY'),
    hasEleven: !!props.getProperty('ELEVENLABS_API_KEY'),
    procFolderId: procFolder_().getId()
  };
}

// ---------- sheet helpers ----------

var SHEET_ = null;
function sheet_() {
  if (SHEET_) return SHEET_;
  var sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('Sessions');
  var head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (head.join('|') !== HEADERS.join('|')) {
    if (head[0] && head[0] !== 'id') { sh.insertRowBefore(1); }  // row 1 holds a session written over the header
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  SHEET_ = sh;
  return sh;
}

function procFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PROC_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.getFolderById(props.getProperty('ROOT_FOLDER_ID')).createFolder('_processing (temporary files)');
  props.setProperty('PROC_FOLDER_ID', f.getId());
  return f;
}

function readAll_() {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var head = values.shift();
  return values.map(function (r, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, j) {
      var v = r[j];
      if (v instanceof Date) v = h === 'date' ? Utilities.formatDate(v, 'Asia/Karachi', 'yyyy-MM-dd') : v.toISOString();
      o[h] = v;
    });
    o.files = o.files ? JSON.parse(o.files) : [];
    o.done = o.done ? JSON.parse(o.done) : [];
    o.jobs = o.jobs ? JSON.parse(o.jobs) : [];
    return o;
  });
}

function find_(id) {
  var s = readAll_().filter(function (x) { return x.id === id; })[0];
  if (!s) throw new Error('Session not found: ' + id);
  return s;
}

function write_(s) {
  s.updated = new Date().toISOString();
  var row = HEADERS.map(function (h) {
    var v = s[h];
    if (h === 'files' || h === 'done' || h === 'jobs') return JSON.stringify(v || []);
    return v === undefined || v === null ? '' : v;
  });
  sheet_().getRange(s._row, 1, 1, HEADERS.length).setValues([row]);
  return clean_(s);
}

function clean_(s) {
  var o = JSON.parse(JSON.stringify(s));
  o.folderUrl = s.folderId ? 'https://drive.google.com/drive/folders/' + s.folderId : '';
  o.transcriptUrl = s.transcriptDocId ? 'https://docs.google.com/document/d/' + s.transcriptDocId + '/edit' : '';
  o.summaryUrl = s.summaryDocId ? 'https://docs.google.com/document/d/' + s.summaryDocId + '/edit' : '';
  return o;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- session API (called from the page) ----------

function listSessions() {
  return readAll_().map(clean_).reverse();
}

function createSession(meta) {
  return withLock_(function () {
    var root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
    var code = String(meta.respondent || '').trim() || nextCode_(meta.kind);
    var folder = root.createFolder((meta.date || '') + ' ' + code + (meta.kind ? ' (' + meta.kind + ')' : ''));
    var id = Utilities.getUuid().slice(0, 8);
    // A row of only blanks is not appended by Sheets, so write the id first.
    var row = HEADERS.map(function (h) { return h === 'id' ? id : ''; });
    var sh = sheet_();
    sh.appendRow(row);
    SpreadsheetApp.flush();
    var s = {
      _row: sh.getLastRow(), id: id, created: new Date().toISOString(),
      addedBy: me_(), date: meta.date || '', respondent: code,
      kind: meta.kind || '', segment: meta.segment || '', interviewer: meta.interviewer || '',
      location: meta.location || '', language: meta.language || 'Urdu and English',
      status: 'New', claimedBy: '', folderId: folder.getId(), files: [], done: [], jobs: [],
      transcriptDocId: '', summaryDocId: '', notes: meta.notes || ''
    };
    return write_(s);
  });
}

var PREFIX = { Merchant: 'M', Institution: 'I', Housing: 'H', Other: 'O' };

/** Next free respondent code for a type, e.g. M-08. */
function nextCode_(kind) {
  var p = PREFIX[kind] || 'O', max = 0, re = new RegExp('^' + p + '-(\\d+)$');
  readAll_().forEach(function (x) { var m = String(x.respondent).match(re); if (m) max = Math.max(max, +m[1]); });
  var n = max + 1;
  return p + '-' + (n < 10 ? '0' : '') + n;
}

var EDITABLE = ['respondent', 'date', 'kind', 'segment', 'interviewer', 'location'];

/** Manual edit of session details from the page. */
function updateSession(id, fields) {
  return withLock_(function () {
    var s = find_(id), changed = false;
    EDITABLE.forEach(function (k) {
      if (fields[k] !== undefined && String(fields[k]) !== String(s[k])) { s[k] = String(fields[k]).trim(); changed = true; }
    });
    if (changed) {
      try { DriveApp.getFolderById(s.folderId).setName(s.date + ' ' + s.respondent + (s.kind ? ' (' + s.kind + ')' : '')); } catch (e) {}
    }
    return write_(s);
  });
}

var SEGMENTS = ['Borrowed formally', 'Declined or abandoned', 'Informal finance only', 'Registered, never borrowed',
  'Bank or microfinance', 'Association or chamber', 'Platform or POS', 'Government'];

/** Fill blank details (interviewer, segment, market) from the transcript or notes. */
function autofill_(id, text) {
  var s = find_(id);
  var blank = !s.interviewer || !s.segment || !s.location;
  if (!blank) return clean_(s);
  var j = {};
  if (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) {
    var prompt = 'From this research interview transcript or field notes, work out the session details. Return JSON with keys: ' +
      '"interviewers" (first names of the researchers asking the questions, as said in the introduction or notes, comma separated, "" if never said), ' +
      '"segment" (exactly one of: ' + SEGMENTS.join('; ') + '; or "" if unclear. For merchants: Borrowed formally = has taken a bank or microfinance loan; Declined or abandoned = applied and was turned down or gave up; Informal finance only = only family, committees, supplier credit or moneylenders; Registered, never borrowed = formally registered but never took a formal loan), ' +
      '"market" (market or area where the interview took place, "" if not said). Use only what the text says.\n\n' + text.slice(0, 30000);
    try { j = JSON.parse(gemini_([{ text: prompt }], { json: true, maxTokens: 1024 })); } catch (e) { j = {}; }
  }
  return withLock_(function () {
    s = find_(id);
    var who = String(s.addedBy || '').split('@')[0].replace(/[._\d]+/g, ' ').trim();
    if (!s.interviewer) s.interviewer = j.interviewers || (who ? who.charAt(0).toUpperCase() + who.slice(1) + ' (uploader)' : '');
    if (!s.segment && SEGMENTS.indexOf(j.segment) >= 0) s.segment = j.segment;
    if (!s.location && j.market) s.location = j.market;
    return write_(s);
  });
}

function addFile(id, file) {
  return withLock_(function () {
    var s = find_(id);
    s.files.push({ id: file.id, name: file.name, mime: file.mime, size: file.size, by: me_() });
    return write_(s);
  });
}

function addTypedNotes(id, text) {
  var s = find_(id);
  var name = 'Typed notes ' + Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd HHmm') + '.txt';
  var f = DriveApp.getFolderById(s.folderId).createFile(name, text, MimeType.PLAIN_TEXT);
  addFile(id, { id: f.getId(), name: name, mime: 'text/plain', size: text.length });
  return autofill_(id, text);
}

function claim(id, release) {
  return withLock_(function () {
    var s = find_(id);
    s.claimedBy = release ? '' : me_();
    return write_(s);
  });
}

function setStatus(id, status) {
  return withLock_(function () {
    var s = find_(id);
    s.status = status;
    return write_(s);
  });
}

// ---------- Gemini ----------

/** Build a Gemini request for UrlFetchApp.fetchAll. */
function geminiReq_(parts, opts) {
  opts = opts || {};
  var props = PropertiesService.getScriptProperties();
  var model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
  var cfg = { temperature: 0.2, maxOutputTokens: opts.maxTokens || 32768 };
  if (opts.json) cfg.responseMimeType = 'application/json';
  if (/^gemini-2\.5-flash/.test(model)) cfg.thinkingConfig = { thinkingBudget: opts.think || 0 };
  return {
    url: 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent',
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': props.getProperty('GEMINI_API_KEY') },
    payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: cfg })
  };
}

function geminiText_(res) {
  var data = JSON.parse(res.getContentText());
  var cand = (data.candidates || [])[0];
  if (!cand || !cand.content) throw new Error('Gemini returned no text (' + (cand && cand.finishReason || 'blocked') + ').');
  return cand.content.parts.map(function (p) { return p.text || ''; }).join('');
}

/** Send many requests at once; repeat only the ones that hit a rate limit or server error. */
function fetchAllRetry_(reqs, name) {
  var out = new Array(reqs.length), todo = reqs.map(function (_, i) { return i; });
  for (var attempt = 0; attempt < 4 && todo.length; attempt++) {
    if (attempt) Utilities.sleep(5000 * attempt);
    var res = UrlFetchApp.fetchAll(todo.map(function (i) { return reqs[i]; })), again = [];
    res.forEach(function (r, k) {
      var c = r.getResponseCode(), i = todo[k];
      out[i] = r;
      if (c === 429 || c >= 500) again.push(i);
    });
    todo = again;
  }
  out.forEach(function (r) {
    if (r.getResponseCode() !== 200) throw new Error(name + ' error ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 300));
  });
  return out;
}

function gemini_(parts, opts) {
  opts = opts || {};
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('No Gemini API key yet. Add GEMINI_API_KEY in Project Settings > Script properties.');
  var model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
  var cfg = { temperature: 0.2, maxOutputTokens: opts.maxTokens || 32768 };
  if (opts.json) cfg.responseMimeType = 'application/json';
  if (/^gemini-2\.5-flash/.test(model)) cfg.thinkingConfig = { thinkingBudget: opts.think || 0 };
  var body = { contents: [{ role: 'user', parts: parts }], generationConfig: cfg };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';
  var last;
  for (var attempt = 0; attempt < 3; attempt++) {
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body)
    });
    var code = res.getResponseCode();
    if (code === 200) {
      var data = JSON.parse(res.getContentText());
      var cand = (data.candidates || [])[0];
      if (!cand || !cand.content) throw new Error('Gemini returned no text (' + (cand && cand.finishReason || 'blocked') + ').');
      return cand.content.parts.map(function (p) { return p.text || ''; }).join('');
    }
    last = 'Gemini error ' + code + ': ' + res.getContentText().slice(0, 300);
    // Retired model: Google's error names the replacement, so switch to it and try again.
    var moved = code === 404 && res.getContentText().match(/use models\/([\w.\-]+)/);
    if (moved && moved[1] !== model) {
      model = moved[1];
      props.setProperty('GEMINI_MODEL', model);
      url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';
      delete cfg.thinkingConfig;
      continue;
    }
    if (code !== 429 && code < 500) break;
    Utilities.sleep(4000 * (attempt + 1));
  }
  throw new Error(last);
}

var PRIVACY = 'Privacy: replace personal names of respondents and their family with [name], and replace CNIC, phone and account numbers with [number]. Keep business types, places at market level, amounts and institution names.';

function transcriptPrompt_(lang, mask) {
  var out;
  if (lang === 'english') out = 'Write every turn in English. Translate Urdu and Punjabi speech faithfully; keep short Urdu terms that matter (udhaar, committee, BC, khata) in Roman Urdu.';
  else if (lang === 'original') out = 'Write every turn in the language spoken: Urdu in Urdu script, English in English.';
  else out = 'For every turn spoken in Urdu or Punjabi, write the original in Urdu script, then on the next line the English translation starting with "EN:". Turns spoken in English need no translation.';
  return [
    'This is a recording from a qualitative research interview in Islamabad, Pakistan, about how small shopkeepers and merchants run their business, register, get paid and borrow money.',
    'Transcribe the whole recording verbatim. Do not summarise or skip anything.',
    'Start each speaker turn on a new line as: [mm:ss] Speaker: text  (timestamps measured from the start of this clip).',
    'Label speakers by role: Interviewer, Note-taker, Respondent, or Other. If a role is unclear use Speaker 1, Speaker 2.',
    'Mark unclear speech as [inaudible]. Mark long silences or interruptions in brackets, e.g. [customer arrives].',
    out,
    mask ? PRIVACY : '',
    'Return only the transcript text.'
  ].join('\n');
}

function shiftTimes_(text, offsetSec) {
  return text.replace(/\[(\d{1,2}):(\d{2})(?::(\d{2}))?\]/g, function (m, a, b, c) {
    var sec = c !== undefined ? (+a) * 3600 + (+b) * 60 + (+c) : (+a) * 60 + (+b);
    sec += offsetSec;
    var h = Math.floor(sec / 3600), mi = Math.floor(sec % 3600 / 60), s = sec % 60;
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return '[' + (h ? h + ':' + pad(mi) : pad(mi)) + ':' + pad(s) + ']';
  });
}

/** Transcribe one compressed audio chunk sent by the page (base64 MP3). */
function transcribeChunk(id, base64, mime, offsetSec, lang, mask, engine) {
  if (engine === 'elevenlabs') return elevenChunk_(base64, mime, offsetSec || 0, lang, mask);
  var text = gemini_([
    { inline_data: { mime_type: mime || 'audio/mp3', data: base64 } },
    { text: transcriptPrompt_(lang, mask) }
  ], { maxTokens: 32768 });
  return shiftTimes_(text.trim(), offsetSec || 0);
}

function fmtClock_(sec) {
  sec = Math.round(sec);
  var h = Math.floor(sec / 3600), m = Math.floor(sec % 3600 / 60), s = sec % 60;
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  return '[' + (h ? h + ':' + pad(m) : pad(m)) + ':' + pad(s) + ']';
}

function elevenReq_(bytes) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('ELEVENLABS_API_KEY');
  if (!key) throw new Error('No ElevenLabs key yet. Add ELEVENLABS_API_KEY in Project Settings > Script properties.');
  return {
    url: 'https://api.elevenlabs.io/v1/speech-to-text', method: 'post', muteHttpExceptions: true, headers: { 'xi-api-key': key },
    payload: { model_id: props.getProperty('ELEVENLABS_MODEL') || 'scribe_v1', file: Utilities.newBlob(bytes, 'audio/mp3', 'chunk.mp3'),
      diarize: 'true', tag_audio_events: 'true', timestamps_granularity: 'word' }
  };
}

/** Turn ElevenLabs words into "[mm:ss] Speaker N: text" lines. */
function elevenText_(res, offsetSec) {
  var data = JSON.parse(res.getContentText());
  var lines = [], cur = null;
  (data.words || []).forEach(function (w) {
    var spk = w.speaker_id || 'speaker_0';
    if (!cur || cur.spk !== spk) {
      if (cur) lines.push(cur);
      cur = { spk: spk, start: w.start || 0, text: '' };
    }
    cur.text += w.type === 'audio_event' ? ' [' + String(w.text).replace(/[()\[\]]/g, '') + '] ' : w.text;
  });
  if (cur) lines.push(cur);
  var text = lines.map(function (l) {
    return fmtClock_(offsetSec + l.start) + ' Speaker ' + (parseInt(String(l.spk).replace(/\D/g, ''), 10) + 1 || 1) + ': ' + l.text.replace(/\s+/g, ' ').trim();
  }).join('\n');
  return text || data.text || '';
}

function polishPrompt_(text, lang, mask) {
  var out = lang === 'english' ? 'Rewrite every turn in English, translating faithfully.' :
    lang === 'original' ? 'Keep every turn in the language spoken.' :
    'Keep each turn as spoken; under every turn spoken in Urdu or Punjabi add a line starting "EN:" with its English translation.';
  return 'Below is a machine transcript of a research interview in Islamabad, with speaker numbers and timestamps. ' +
    'Replace "Speaker N" with a role (Interviewer, Note-taker, Respondent, Other) where the content makes the role clear; otherwise keep the number. ' +
    'Do not change, shorten or reorder the words or the timestamps. ' + out + ' ' + (mask ? PRIVACY : '') + ' Return only the transcript.\n\n' + text;
}

/** ElevenLabs Scribe: speaker-labelled transcript; Gemini (if set) adds roles, translation and masking. */
function elevenChunk_(base64, mime, offsetSec, lang, mask) {
  var text = elevenText_(fetchAllRetry_([elevenReq_(Utilities.base64Decode(base64))], 'ElevenLabs')[0], offsetSec);
  if (!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || !text) return text;
  return gemini_([{ text: polishPrompt_(text, lang, mask) }], { maxTokens: 32768 }).trim();
}

/** Transcribe several parts of one recording at the same time. Returns texts in order. */
function transcribeParts_(chunks, lang, mask, engine) {
  var hasGem = !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (engine !== 'elevenlabs') {
    var gres = fetchAllRetry_(chunks.map(function (c) {
      return geminiReq_([{ inline_data: { mime_type: 'audio/mp3', data: Utilities.base64Encode(c.bytes) } }, { text: transcriptPrompt_(lang, mask) }], { maxTokens: 32768 });
    }), 'Gemini');
    return gres.map(function (r, i) { return shiftTimes_(geminiText_(r).trim(), chunks[i].offset); });
  }
  var raw = fetchAllRetry_(chunks.map(function (c) { return elevenReq_(c.bytes); }), 'ElevenLabs')
    .map(function (r, i) { return elevenText_(r, chunks[i].offset); });
  if (!hasGem) return raw;
  var idx = [], reqs = [];
  raw.forEach(function (t, i) { if (t) { idx.push(i); reqs.push(geminiReq_([{ text: polishPrompt_(t, lang, mask) }], { maxTokens: 32768 })); } });
  if (!reqs.length) return raw;
  try {
    fetchAllRetry_(reqs, 'Gemini').forEach(function (r, k) { raw[idx[k]] = geminiText_(r).trim(); });
  } catch (e) {
    Logger.log('Polish skipped: ' + e.message);  // keep the ElevenLabs text rather than lose the part
  }
  return raw;
}

/** Read a photo of handwritten notes, or tidy typed notes, from Drive. */
function readNotes(id, fileId, mask) {
  var s = find_(id);
  var file = s.files.filter(function (f) { return f.id === fileId; })[0];
  if (!file) throw new Error('File not found in this session.');
  var blob = DriveApp.getFileById(fileId).getBlob();
  var tidy = 'Turn these into clean, readable field notes. Keep every point, number and quote; do not add anything that is not there. Group points under short headings where the content allows (the business and registering; selling, getting paid and buying stock; money and borrowing; documents; concept and card sort reactions; other). Write in English; keep Urdu quotes in Urdu script with an English translation. Mark words you cannot read as [illegible]. ' + (mask ? PRIVACY : '') + ' Return plain text with headings on their own lines starting with "## ".';
  var parts;
  if (/^image\//.test(file.mime) || file.mime === 'application/pdf') {
    parts = [{ inline_data: { mime_type: file.mime, data: Utilities.base64Encode(blob.getBytes()) } },
      { text: 'These are photos or scans of handwritten interview notes from a research interview with a shopkeeper in Islamabad. ' + tidy }];
  } else {
    parts = [{ text: 'These are typed interview notes from a research interview with a shopkeeper in Islamabad. ' + tidy + '\n\nNOTES:\n' + blob.getDataAsString() }];
  }
  return gemini_(parts, { maxTokens: 16384 }).trim();
}

// ---------- docs ----------

function styleDoc_(body) {
  body.editAsText().setFontFamily('DM Sans').setFontSize(10.5).setForegroundColor('#343436');
}

function heading_(body, text, level) {
  var p = body.appendParagraph(text);
  p.setHeading(level === 1 ? DocumentApp.ParagraphHeading.HEADING1 : DocumentApp.ParagraphHeading.HEADING2);
  p.editAsText().setFontFamily('Figtree').setForegroundColor('#073763').setBold(true).setFontSize(level === 1 ? 18 : 13);
  return p;
}

function metaLine_(s) {
  return [s.respondent, s.kind, s.segment, s.date, s.location, 'Interviewer: ' + (s.interviewer || 'n/a')]
    .filter(function (x) { return x; }).join('  ·  ');
}

function ensureTranscriptDoc_(s) {
  if (s.transcriptDocId) return DocumentApp.openById(s.transcriptDocId);
  var doc = DocumentApp.create(s.respondent + ' · Transcript and notes');
  var body = doc.getBody();
  body.clear();
  heading_(body, s.respondent + ' · Transcript and notes', 1);
  body.appendParagraph(metaLine_(s));
  styleDoc_(body);
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(s.folderId));
  s.transcriptDocId = doc.getId();
  return DocumentApp.openById(doc.getId());
}

/** Append a finished transcript or tidied notes for one source file. */
function saveText(id, fileId, title, text) {
  saveText_(id, fileId, title, text);
  return autofill_(id, text);
}

function saveText_(id, fileId, title, text) {
  return withLock_(function () {
    var s = find_(id);
    var doc = ensureTranscriptDoc_(s);
    var body = doc.getBody();
    heading_(body, title, 2);
    text.split(/\n/).forEach(function (line) {
      if (/^##\s+/.test(line)) {
        var h = body.appendParagraph(line.replace(/^##\s+/, ''));
        h.setHeading(DocumentApp.ParagraphHeading.HEADING3);
        h.editAsText().setFontFamily('Figtree').setBold(true).setForegroundColor('#073763');
      } else {
        var p = body.appendParagraph(line);
        p.editAsText().setFontFamily('DM Sans').setFontSize(10.5).setForegroundColor(/^EN:/.test(line) ? '#5b5b5f' : '#343436');
        if (/[؀-ۿ]/.test(line)) p.setLeftToRight(false);
      }
    });
    doc.saveAndClose();
    if (s.done.indexOf(fileId) < 0) s.done.push(fileId);
    if (['New', 'Transcribing', 'Processing'].indexOf(s.status) >= 0) s.status = 'Transcribed';
    return write_(s);
  });
}

function summarise(id, mask) {
  var s = find_(id);
  if (!s.transcriptDocId) throw new Error('Transcribe a recording or notes first.');
  var text = DocumentApp.openById(s.transcriptDocId).getBody().getText();
  var isInst = /institution/i.test(s.kind);
  var prompt = [
    'You are helping a research team analyse one qualitative interview from field research for the Pakistan Digital Authority in Islamabad. The study tests seven assumptions behind digitising lending to small merchants (SMEs) through a government registration portal and credit scheme.',
    'Interview: ' + metaLine_(s) + '. ' + (isInst ? 'This is an interview with an institution (bank, microfinance bank, association, platform or government body), not a merchant.' : 'This is an interview with a merchant.'),
    'The seven assumptions, with their short labels:\n- ' + ASSUMPTIONS.join('\n- '),
    'The journey steps: ' + JOURNEY.join('; ') + '.',
    'Using ONLY what is in the transcript and notes below, return JSON with these keys:',
    '"snapshot": array of {"label","value"} covering who the respondent is, business type, years running, staff and outlets, payment channels, segment as described (borrowed formally / declined or abandoned / informal finance only / registered, never borrowed), phone and internet use. Use "not covered" when absent.',
    '"takeaways": 3 to 5 short sentences, the most important things this interview tells the team.',
    '"journey": array of {"step","what_happened","pain"} for the journey steps the interview covers' + (isInst ? ', described from the institution\'s side (what they do at that step and where files stall)' : '') + '. Omit steps not discussed.',
    '"assumptions": array of 7 objects {"assumption" (the short label, e.g. "Wants a loan?"), "verdict" (one of "Supports", "Challenges", "Mixed", "Not covered"), "evidence" (1 to 3 sentences), "quote" (a short verbatim quote in English with its [mm:ss] timestamp if there is one, else "")}. A verdict says whether what was heard supports or challenges the assumption being true.',
    '"workarounds": array of strings: intermediaries, informal routes and workarounds mentioned.',
    '"documents": array of strings: documents asked for, held or missing.',
    '"concepts": array of strings: reactions to the registration route concepts and the data-sharing card sort, if covered.',
    '"quotes": array of {"quote","timestamp","why"}: up to 6 strong verbatim quotes in English.',
    '"followups": array of strings: gaps, contradictions or questions to check in later interviews.',
    'Be specific and factual. Do not invent anything. ' + (mask ? PRIVACY : ''),
    'TRANSCRIPT AND NOTES:\n' + text
  ].join('\n\n');
  var raw = gemini_([{ text: prompt }], { json: true, maxTokens: 16384, think: 2048 });
  var j;
  try { j = JSON.parse(raw); } catch (e) { throw new Error('The summary came back in an unexpected format. Try again.'); }
  return withLock_(function () {
    s = find_(id);
    var doc;
    if (s.summaryDocId) {
      doc = DocumentApp.openById(s.summaryDocId);
      doc.getBody().clear();
    } else {
      doc = DocumentApp.create(s.respondent + ' · Summary');
      DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(s.folderId));
      s.summaryDocId = doc.getId();
    }
    buildSummary_(doc.getBody(), s, j);
    doc.saveAndClose();
    s.status = 'Summarised';
    return write_(s);
  });
}

function table_(body, head, rows, widths) {
  var t = body.appendTable([head].concat(rows.length ? rows : [head.map(function () { return ''; })]));
  t.setBorderColor('#d0d4dc');
  for (var c = 0; c < head.length; c++) {
    var cell = t.getRow(0).getCell(c);
    cell.setBackgroundColor('#073763');
    cell.editAsText().setForegroundColor('#ffffff').setBold(true).setFontFamily('Figtree').setFontSize(10);
  }
  for (var r = 1; r < t.getNumRows(); r++) {
    for (var k = 0; k < head.length; k++) {
      var cl = t.getRow(r).getCell(k);
      if (r % 2 === 0) cl.setBackgroundColor('#F3F3F3');
      cl.editAsText().setFontFamily('DM Sans').setFontSize(9.5).setForegroundColor('#343436');
    }
  }
  if (widths) widths.forEach(function (w, i) { if (w) t.setColumnWidth(i, w); });
  return t;
}

function bullets_(body, items) {
  if (!items || !items.length) { body.appendParagraph('Not covered.'); return; }
  items.forEach(function (x) { body.appendListItem(String(x)).setGlyphType(DocumentApp.GlyphType.BULLET); });
}

function buildSummary_(body, s, j) {
  heading_(body, s.respondent + ' · Interview summary', 1);
  body.appendParagraph(metaLine_(s));
  heading_(body, 'Snapshot', 2);
  table_(body, ['', ''], (j.snapshot || []).map(function (x) { return [x.label || '', x.value || '']; }), [150, null]);
  heading_(body, 'Key takeaways', 2);
  bullets_(body, j.takeaways);
  heading_(body, 'What we heard on the seven assumptions', 2);
  var t = table_(body, ['Assumption', 'Verdict', 'Evidence', 'Quote'], (j.assumptions || []).map(function (a) {
    return [a.assumption || '', a.verdict || '', a.evidence || '', a.quote || ''];
  }), [105, 70, null, 150]);
  var colours = { Supports: '#d9ead3', Challenges: '#f4cccc', Mixed: '#fff2cc', 'Not covered': '#eeeeee' };
  for (var r = 1; r < t.getNumRows(); r++) {
    var v = t.getRow(r).getCell(1).getText();
    if (colours[v]) t.getRow(r).getCell(1).setBackgroundColor(colours[v]);
  }
  heading_(body, 'Journey', 2);
  table_(body, ['Step', 'What happened', 'Pain or stall'], (j.journey || []).map(function (x) {
    return [x.step || '', x.what_happened || '', x.pain || ''];
  }), [120, null, 160]);
  heading_(body, 'Intermediaries and workarounds', 2);
  bullets_(body, j.workarounds);
  heading_(body, 'Documents', 2);
  bullets_(body, j.documents);
  heading_(body, 'Concept and card sort reactions', 2);
  bullets_(body, j.concepts);
  heading_(body, 'Quotes', 2);
  table_(body, ['Quote', 'Time', 'Why it matters'], (j.quotes || []).map(function (q) {
    return [q.quote || '', q.timestamp || '', q.why || ''];
  }), [null, 55, 150]);
  heading_(body, 'Follow-ups for later interviews', 2);
  bullets_(body, j.followups);
  var first = body.getChild(0);
  if (first.getType() === DocumentApp.ElementType.PARAGRAPH && first.asParagraph().getText() === '' && body.getNumChildren() > 1) first.removeFromParent();
}


// ---------- background queue ----------
// The page uploads files, then calls enqueue(). processQueue() runs from the
// page straight away and from a 5-minute trigger, so work carries on after the
// tab is closed. Each job holds a lease so two runs never do the same part.

function jobOf_(s, fileId) { return (s.jobs || []).filter(function (j) { return j.fileId === fileId; })[0]; }

function pending_(s) { return (s.jobs || []).some(function (j) { return j.state === 'queued' || j.state === 'running'; }); }

/** job: {fileId, procId, name, kind: 'media'|'notes', engine, lang, mask, dur, auto} */
function enqueue(id, job) {
  return withLock_(function () {
    var s = find_(id);
    s.jobs = (s.jobs || []).filter(function (j) { return j.fileId !== job.fileId; });
    var parts = job.kind === 'media' ? Math.max(1, Math.ceil((job.dur || 1) / PART_SEC)) : 1;
    s.jobs.push({
      fileId: job.fileId, procId: job.procId || job.fileId, name: job.name, kind: job.kind,
      engine: job.engine || 'gemini', lang: job.lang || 'both', mask: job.mask !== false, dur: job.dur || 0,
      part: 0, parts: parts, workId: '', state: 'queued', lease: 0, tries: 0, err: '', auto: !!job.auto,
      by: me_()
    });
    if (s.status !== 'Reviewed') s.status = 'Processing';
    return write_(s);
  });
}

function requestSummary(id, mask) {
  return enqueue(id, { fileId: '__summary', name: 'Summary', kind: 'summary', mask: mask });
}

function retryJob(id, fileId) {
  return withLock_(function () {
    var s = find_(id), j = jobOf_(s, fileId);
    if (j) { j.state = 'queued'; j.tries = 0; j.err = ''; j.lease = 0; }
    s.status = 'Processing';
    return write_(s);
  });
}

function cancelJob(id, fileId) {
  return withLock_(function () {
    var s = find_(id);
    s.jobs = (s.jobs || []).filter(function (j) { return j.fileId !== fileId; });
    if (!pending_(s) && s.status === 'Processing') s.status = s.summaryDocId ? 'Summarised' : (s.transcriptDocId ? 'Transcribed' : 'New');
    return write_(s);
  });
}

function processQueue() {
  var t0 = Date.now(), did = 0;
  while (Date.now() - t0 < RUN_BUDGET_MS) {
    var L = leaseNext_();
    if (!L) break;
    runJob_(L, t0);
    did++;
  }
  return did;
}

function leaseNext_() {
  return withLock_(function () {
    var now = Date.now(), all = readAll_();
    for (var i = 0; i < all.length; i++) {
      var s = all[i];
      for (var k = 0; k < (s.jobs || []).length; k++) {
        var j = s.jobs[k];
        if (j.kind === 'summary' && s.jobs.some(function (x) { return x.kind !== 'summary' && (x.state === 'queued' || x.state === 'running'); })) continue;
        if ((j.state === 'queued' || j.state === 'running') && (j.lease || 0) < now) {
          j.state = 'running'; j.lease = now + 7 * 60000; j.key = Utilities.getUuid();
          write_(s);
          return { id: s.id, fileId: j.fileId, key: j.key };
        }
      }
    }
    return null;
  });
}

/** Update a job the caller still holds; returns false if the lease was lost. */
function saveJob_(L, patch) {
  return withLock_(function () {
    var s = find_(L.id), j = jobOf_(s, L.fileId);
    if (!j || j.key !== L.key) return false;
    for (var k in patch) j[k] = patch[k];
    if (!pending_(s) && s.status === 'Processing') s.status = s.summaryDocId ? 'Summarised' : (s.transcriptDocId ? 'Transcribed' : 'New');
    write_(s);
    return true;
  });
}

var BYTES_ = {};
function procBytes_(procId) {
  if (!BYTES_[procId]) BYTES_ = {}, BYTES_[procId] = DriveApp.getFileById(procId).getBlob().getBytes();
  return BYTES_[procId];
}

function runJob_(L, t0) {
  var s = find_(L.id), j = jobOf_(s, L.fileId);
  if (!j || j.key !== L.key) return;
  try {
    if (j.kind === 'summary') {
      summarise(L.id, j.mask);
      saveJob_(L, { state: 'done', lease: 0 });
      return;
    }
    if (j.kind === 'notes') {
      var notes = readNotes(L.id, j.fileId, j.mask);
      finishJob_(L, j, 'Notes: ' + j.name, notes);
      return;
    }
    var props = PropertiesService.getScriptProperties();
    var width = +(props.getProperty(j.engine === 'elevenlabs' ? 'ELEVENLABS_PARALLEL' : 'GEMINI_PARALLEL') || (j.engine === 'elevenlabs' ? 4 : 6));
    while (j.part < j.parts && Date.now() - t0 < RUN_BUDGET_MS - 90000) {
      var bytes = procBytes_(j.procId), total = bytes.length;
      var bps = total / Math.max(j.dur, 1), chunks = [];
      for (var p = j.part; p < Math.min(j.parts, j.part + width); p++) {
        var start = Math.floor(p * PART_SEC * bps);
        var end = p === j.parts - 1 ? total : Math.floor((p + 1) * PART_SEC * bps);
        chunks.push({ bytes: bytes.slice(start, end), offset: p * PART_SEC });
      }
      var text = transcribeParts_(chunks, j.lang, j.mask, j.engine).join('\n');
      if (!j.workId) {
        j.workId = procFolder_().createFile(s.respondent + ' ' + j.name + ' (in progress).txt', text, MimeType.PLAIN_TEXT).getId();
      } else {
        var wf = DriveApp.getFileById(j.workId);
        wf.setContent(wf.getBlob().getDataAsString() + '\n' + text);
      }
      j.part += chunks.length;
      if (!saveJob_(L, { part: j.part, workId: j.workId, lease: Date.now() + 7 * 60000 })) return;
    }
    if (j.part < j.parts) { saveJob_(L, { lease: 0 }); return; }  // next run continues
    var content = j.workId ? DriveApp.getFileById(j.workId).getBlob().getDataAsString() : '';
    var label = j.engine === 'elevenlabs' ? 'Transcript (ElevenLabs): ' : 'Transcript: ';
    var m = Math.round(j.dur / 60);
    finishJob_(L, j, label + j.name + ' (' + m + ' min)', content);
    try { DriveApp.getFileById(j.workId).setTrashed(true); } catch (e) {}
    if (j.procId !== j.fileId) { try { DriveApp.getFileById(j.procId).setTrashed(true); } catch (e) {} }
  } catch (e) {
    var tries = (j.tries || 0) + 1;
    saveJob_(L, tries >= 3 ? { state: 'error', err: String(e.message || e), tries: tries, lease: 0 }
      : { tries: tries, err: String(e.message || e), lease: Date.now() + 60000, state: 'queued' });
  }
}

function finishJob_(L, j, title, text) {
  saveText_(L.id, L.fileId, title, text);
  saveJob_(L, { state: 'done', lease: 0, err: '' });
  try { autofill_(L.id, text); } catch (e) {}
  if (j.auto) {
    var s = find_(L.id);
    var others = (s.jobs || []).some(function (x) { return x.kind !== 'summary' && (x.state === 'queued' || x.state === 'running'); });
    if (!others) requestSummary(L.id, j.mask);
  }
}
