/**
 * G-ONE 쉽게 배우기 · 자료실 백엔드 (GONE 스프레드시트에 연결된 Apps Script)
 *
 * 시트 구성
 *  - resources : 자료 목록 (아래 HEADERS 순서)
 *  - content   : 사이트 문구 (key · value · updatedAt) — 관리자가 연필로 고친 글만 저장됩니다
 *  - 설정       : A열 키 / B열 값. adminPassword 행의 값 = 관리자 비밀번호
 *  - log       : 관리 기록
 *
 * 배포: 배포 → 배포 관리 → (연필) → 버전: 새 버전 → 배포  (웹 앱 주소는 그대로 유지)
 */

var SHEET = 'resources';
var CONTENT = 'content';
var HEADERS = ['id', 'order', 'category', 'type', 'title', 'description', 'url', 'thumbnail',
               'duration', 'tags', 'featured', 'isNew', 'visible', 'views', 'createdAt', 'updatedAt'];
var EDITABLE = ['order', 'category', 'type', 'title', 'description', 'url', 'thumbnail',
                'duration', 'tags', 'featured', 'isNew', 'visible'];

/* ---------- 웹 요청 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    switch (p.action) {
      case 'view': return json(addView_(p.id));
      case 'list':
      default:     return json({ ok: true, items: list_(), content: content_(), updated: new Date().toISOString() });
    }
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

// 브라우저에서 Content-Type: text/plain 으로 JSON 문자열을 보냅니다 (CORS 사전 요청 회피)
function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkKey_(body.key);
    switch (body.action) {
      case 'login':      return json({ ok: true });
      case 'adminList':  return json({ ok: true, items: adminList_(), content: content_() });
      case 'add':        return json(add_(body.item || {}));
      case 'update':     return json(update_(body.id, body.item || {}));
      case 'hide':       return json(update_(body.id, { visible: false }));
      case 'show':       return json(update_(body.id, { visible: true }));
      case 'delete':     return json(remove_(body.id));
      case 'reorder':    return json(reorder_(body.ids || []));
      case 'setContent': return json(setContent_(body.name, body.value));
      default:           return json({ ok: false, error: '알 수 없는 요청입니다.' });
    }
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- 읽기 ---------- */
function list_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('list');
  if (hit) return JSON.parse(hit);
  var items = readAll_()
    .filter(function (r) { return String(r.visible).toUpperCase() !== 'FALSE'; })
    .map(pub_)
    .sort(byOrder_);
  cache.put('list', JSON.stringify(items), 60); // 1분 캐시 (바꾸면 즉시 지움)
  return items;
}
function adminList_() {
  return readAll_().map(function (r) {
    var o = pub_(r);
    o.visible = String(r.visible).toUpperCase() !== 'FALSE';
    o.createdAt = r.createdAt; o.updatedAt = r.updatedAt;
    return o;
  }).sort(byOrder_);
}
function pub_(r) {
  return {
    id: String(r.id), order: Number(r.order) || 999, category: r.category || 'etc', type: r.type || 'video',
    title: r.title, description: r.description, url: r.url, thumbnail: r.thumbnail, duration: String(r.duration || ''),
    tags: r.tags, featured: bool_(r.featured), isNew: bool_(r.isNew), views: Number(r.views) || 0
  };
}
function byOrder_(a, b) { return (a.order || 999) - (b.order || 999); }

function readAll_() {
  var sh = sheet_(), v = sh.getDataRange().getDisplayValues(); // 1:30 같은 값이 시간으로 바뀌지 않게 보이는 그대로 읽음
  var head = v.shift();
  return v.filter(function (row) { return row[0] !== ''; }).map(function (row) {
    var o = {}; head.forEach(function (h, i) { o[h] = row[i]; }); return o;
  });
}

function content_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('content');
  if (hit) return JSON.parse(hit);
  var sh = contentSheet_(), out = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (String(r[0]).trim()) out[String(r[0]).trim()] = String(r[1]);
    });
  }
  var s = JSON.stringify(out);
  if (s.length < 90000) cache.put('content', s, 300);
  return out;
}

/* ---------- 쓰기 ---------- */
function addView_(id) {
  if (!id) return { ok: false, error: 'id가 없습니다.' };
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sh = sheet_(), row = findRow_(sh, id);
    if (!row) return { ok: false, error: '자료를 찾지 못했습니다.' };
    var col = HEADERS.indexOf('views') + 1;
    var n = (Number(sh.getRange(row, col).getValue()) || 0) + 1;
    sh.getRange(row, col).setValue(n);
    clear_();
    return { ok: true, views: n };
  } finally { lock.releaseLock(); }
}

function add_(item) {
  if (!item.title) throw new Error('제목을 입력해 주세요.');
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sh = sheet_(), now = new Date();
    var id = String(item.id || Utilities.getUuid().slice(0, 8)).trim();
    if (findRow_(sh, id)) throw new Error('같은 id가 이미 있습니다: ' + id);
    var rec = {
      id: id, order: Number(item.order) || sh.getLastRow(), category: item.category || 'etc', type: item.type || 'video',
      title: item.title, description: item.description || '', url: item.url || '', thumbnail: item.thumbnail || '',
      duration: item.duration || '', tags: item.tags || '', featured: bool_(item.featured),
      isNew: item.isNew === undefined ? true : bool_(item.isNew), visible: item.visible === undefined ? true : bool_(item.visible),
      views: 0, createdAt: now, updatedAt: now
    };
    var row = sh.getLastRow() + 1;
    sh.getRange(row, HEADERS.indexOf('duration') + 1).setNumberFormat('@');
    sh.getRange(row, 1, 1, HEADERS.length).setValues([HEADERS.map(function (h) { return rec[h]; })]);
    if (rec.featured) onlyFeatured_(sh, id);
    clear_();
    log_('add', id, item.title);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

function update_(id, patch) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sh = sheet_(), row = findRow_(sh, id);
    if (!row) throw new Error('자료를 찾지 못했습니다.');
    Object.keys(patch).forEach(function (k) {
      if (EDITABLE.indexOf(k) < 0) return;
      var c = HEADERS.indexOf(k) + 1, v = patch[k];
      if (k === 'featured' || k === 'isNew' || k === 'visible') v = bool_(v);
      if (k === 'order') v = Number(v) || 999;
      if (k === 'duration') sh.getRange(row, c).setNumberFormat('@');
      sh.getRange(row, c).setValue(v);
    });
    if (patch.featured !== undefined && bool_(patch.featured)) onlyFeatured_(sh, id);
    sh.getRange(row, HEADERS.indexOf('updatedAt') + 1).setValue(new Date());
    clear_();
    log_('update', id, JSON.stringify(patch).slice(0, 500));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function remove_(id) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sh = sheet_(), row = findRow_(sh, id);
    if (!row) throw new Error('자료를 찾지 못했습니다.');
    var title = sh.getRange(row, HEADERS.indexOf('title') + 1).getValue();
    sh.deleteRow(row);
    clear_();
    log_('delete', id, title);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function reorder_(ids) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sh = sheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return { ok: true };
    var idv = sh.getRange(2, 1, n, 1).getValues();
    var col = HEADERS.indexOf('order') + 1;
    var ord = sh.getRange(2, col, n, 1).getValues();
    ids.forEach(function (id, i) {
      for (var r = 0; r < n; r++) if (String(idv[r][0]) === String(id)) ord[r][0] = i + 1;
    });
    sh.getRange(2, col, n, 1).setValues(ord);
    clear_();
    log_('reorder', '', ids.join(','));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function setContent_(name, value) {
  name = String(name || '').trim();
  if (!/^[\w.\-]{1,60}$/.test(name)) throw new Error('문구 이름이 올바르지 않습니다.');
  value = String(value == null ? '' : value);
  if (value.length > 5000) throw new Error('글이 너무 깁니다.');
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sh = contentSheet_(), row = 0;
    if (sh.getLastRow() > 1) {
      var keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === name) { row = i + 2; break; }
    }
    if (value.trim() === '') { if (row) sh.deleteRow(row); }          // 비우면 기본 문구로 되돌림
    else if (row) sh.getRange(row, 2, 1, 2).setValues([[value, new Date()]]);
    else sh.appendRow([name, value, new Date()]);
    CacheService.getScriptCache().remove('content');
    log_('content', name, value.slice(0, 200));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- 도우미 ---------- */
function sheet_() {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}
function contentSheet_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(CONTENT);
  if (!sh) {
    sh = ss.insertSheet(CONTENT);
    sh.appendRow(['key', 'value', 'updatedAt']);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eefc');
    sh.setFrozenRows(1);
    sh.setColumnWidth(2, 480);
  }
  return sh;
}
function findRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function onlyFeatured_(sh, id) {  // 추천(TODAY'S PICK)은 하나만
  var n = sh.getLastRow() - 1; if (n < 1) return;
  var idv = sh.getRange(2, 1, n, 1).getValues(), col = HEADERS.indexOf('featured') + 1;
  sh.getRange(2, col, n, 1).setValues(idv.map(function (r) { return [String(r[0]) === String(id)]; }));
}
function clear_() { CacheService.getScriptCache().remove('list'); }
function checkKey_(key) {
  var real = getSetting_('adminPassword');
  if (!real) throw new Error("'설정' 시트에 adminPassword 값이 없습니다.");
  if (String(key || '').trim() !== real) { Utilities.sleep(800); throw new Error('관리자 비밀번호가 맞지 않습니다.'); }
}
// '설정' 시트(A열 키, B열 값)에서 값을 읽습니다
function getSetting_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName('설정');
  if (!sh || sh.getLastRow() < 1) return '';
  var v = sh.getRange(1, 1, sh.getLastRow(), 2).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === name) return String(v[i][1]).trim();
  return '';
}
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function log_(what, id, note) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('log') || ss.insertSheet('log');
  if (sh.getLastRow() === 0) sh.appendRow(['time', 'action', 'id', 'note']);
  sh.appendRow([new Date(), what, id, note]);
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 처음 한 번: 시트 머리글 만들기 (자료는 사이트 관리 화면에서 등록) ---------- */
function setup() {
  var sh = sheet_();
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e8eefc');
    sh.setFrozenRows(1);
  }
  sh.getRange(2, HEADERS.indexOf('duration') + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  contentSheet_();
}
