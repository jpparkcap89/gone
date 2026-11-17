/**
 * G-ONE 쉽게 배우기 · 자료실 백엔드 (Google Apps Script + Google Sheets)
 *
 * 시트 구성
 *  - resources : 자료 목록 (아래 HEADERS 순서)
 *  - log       : 조회·등록 기록 (선택)
 *
 * 스크립트 속성 (프로젝트 설정 → 스크립트 속성)
 *  - ADMIN_KEY : 관리자 등록용 비밀 키 (사이트 ?admin 화면에서 입력)
 *
 * 처음 한 번: 편집기에서 setup() 실행 → 시트와 예시 데이터가 만들어집니다.
 * 배포: 배포 → 새 배포 → 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → /exec 주소를 index.html API_URL에.
 */

// 자료실 DB 스프레드시트 ID (독립형 스크립트로 연결)
var SHEET_ID = '1_05KaUU1pN2BeV9HtZtWt2_xesPO3alZBE_fGorC0Po';
var SHEET = 'resources';
var HEADERS = ['id', 'order', 'category', 'type', 'title', 'description', 'url', 'thumbnail',
               'duration', 'tags', 'featured', 'isNew', 'visible', 'views', 'createdAt', 'updatedAt'];

/* ---------- 웹 요청 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    switch (p.action) {
      case 'view': return json(addView_(p.id));
      case 'list':
      default:     return json({ ok: true, items: list_(), updated: new Date().toISOString() });
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
      case 'add':    return json(add_(body.item || {}));
      case 'update': return json(update_(body.id, body.item || {}));
      case 'hide':   return json(update_(body.id, { visible: false }));
      default:       return json({ ok: false, error: '알 수 없는 요청입니다.' });
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
  var rows = readAll_().filter(function (r) { return r.visible !== false && String(r.visible).toUpperCase() !== 'FALSE'; });
  rows.sort(function (a, b) { return (Number(a.order) || 999) - (Number(b.order) || 999); });
  var items = rows.map(function (r) {
    return {
      id: String(r.id), order: Number(r.order) || 999, category: r.category || 'etc', type: r.type || 'video',
      title: r.title, description: r.description, url: r.url, thumbnail: r.thumbnail, duration: String(r.duration || ''),
      tags: r.tags, featured: bool_(r.featured), isNew: bool_(r.isNew), views: Number(r.views) || 0
    };
  });
  cache.put('list', JSON.stringify(items), 60); // 1분 캐시 (조회수는 view 요청 때 즉시 반영)
  return items;
}

function readAll_() {
  var sh = sheet_(), v = sh.getDataRange().getValues();
  var head = v.shift();
  return v.filter(function (row) { return row[0] !== ''; }).map(function (row) {
    var o = {}; head.forEach(function (h, i) { o[h] = row[i]; }); return o;
  });
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
    CacheService.getScriptCache().remove('list');
    return { ok: true, views: n };
  } finally { lock.releaseLock(); }
}

function add_(item) {
  if (!item.title) throw new Error('제목을 입력해 주세요.');
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sh = sheet_(), now = new Date();
    var id = item.id || Utilities.getUuid().slice(0, 8);
    var order = item.order || sh.getLastRow(); // 맨 뒤에
    var rec = {
      id: id, order: order, category: item.category || 'etc', type: item.type || 'video', title: item.title,
      description: item.description || '', url: item.url || '', thumbnail: item.thumbnail || '', duration: item.duration || '',
      tags: item.tags || '', featured: false, isNew: true, visible: true, views: 0, createdAt: now, updatedAt: now
    };
    sh.appendRow(HEADERS.map(function (h) { return rec[h]; }));
    CacheService.getScriptCache().remove('list');
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
      var c = HEADERS.indexOf(k);
      if (c > 0 && k !== 'views' && k !== 'createdAt') sh.getRange(row, c + 1).setValue(patch[k]);
    });
    sh.getRange(row, HEADERS.indexOf('updatedAt') + 1).setValue(new Date());
    CacheService.getScriptCache().remove('list');
    log_('update', id, JSON.stringify(patch));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- 도우미 ---------- */
function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}
function findRow_(sh, id) {
  var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function checkKey_(key) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!real) throw new Error('스크립트 속성 ADMIN_KEY가 설정되지 않았습니다.');
  if (key !== real) throw new Error('관리자 키가 맞지 않습니다.');
}
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function log_(what, id, note) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('log') || ss.insertSheet('log');
  if (sh.getLastRow() === 0) sh.appendRow(['time', 'action', 'id', 'note']);
  sh.appendRow([new Date(), what, id, note]);
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 처음 한 번 실행: 시트 만들기 + 예시 자료 ---------- */
function setup() {
  var sh = sheet_();
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e8eefc');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() > 1) return; // 이미 자료가 있으면 건너뜀
  var now = new Date();
  var seed = [
    ['m0', 1, 'messenger', 'video', '브리티 메신저 에센셜 편', 'GOE메신저 쓰던 선생님을 위한 빠른 전환: 자동 로그인, 선생님 찾기, 쪽지, 대화, 알림까지 한 번에.', '1:30', '에센셜,자동 로그인,쪽지,전환', true, true],
    ['m1', 2, 'messenger', 'video', '브리티 메신저 1편 · 시작하기와 자동 로그인', '설치, 로그인, 자동 로그인 3단계, 화면 구성을 알아봅니다.', '1:35', '자동 로그인,설치,로그인', false, false],
    ['m2', 3, 'messenger', 'video', '브리티 메신저 2편 · 연락처·조직도로 선생님 찾기', '내 상태 바꾸기, 조직도, 검색, 여러 명에게 한 번에 보내기.', '1:19', '조직도,연락처,검색', false, false],
    ['m3', 4, 'messenger', 'video', '브리티 메신저 3편 · 대화하기', '1:1과 그룹 대화, 메시지 예약 발신, 이모티콘, 공지방.', '1:32', '대화,예약 발신,공지방', false, false],
    ['m4', 5, 'messenger', 'video', '브리티 메신저 4편 · 쪽지 보내고 확인하기', '쪽지 쓰기, 받는 사람 고르기, 읽음 확인, 발신취소, 답장·전달.', '1:31', '쪽지,수신확인,회수,발신취소', false, false],
    ['m5', 6, 'messenger', 'video', '브리티 메신저 5편 · 워크스페이스·채널', '워크스페이스와 채널 만들기, 멤버 초대, 공지·파일·링크 탭.', '1:31', '워크스페이스,채널,협업', false, false],
    ['m6', 7, 'messenger', 'video', '브리티 메신저 6편 · 환경설정', '알림, 키워드 알림, 메시지 전송 키, 파일 저장 위치, 글꼴, 스킨.', '1:31', '환경설정,알림,줄바꿈', false, false],
    ['c1', 8, 'calendar', 'video', 'G-ONE 일정 관리 1부', '캘린더 만들기, 일정 등록(기관 공개·참석자), 상세 등록으로 여러 명 공유.', '2:57', '일정,캘린더,기관 공개', false, false],
    ['c2', 9, 'calendar', 'video', 'G-ONE 일정 관리 2부', 'Year 보기, 할 일, 캘린더 공유·구독, ICS 내보내기·가져오기.', '2:05', '일정,구독,ICS,할 일', false, false],
    ['c3', 10, 'calendar', 'video', '브리티 일정 관리 TIP', '날짜 클릭 간편 등록부터 상세 등록, 수정·삭제, 할 일까지.', '3:05', '일정,할 일,TIP', false, false],
    ['g1', 11, 'guide', 'pdf', '브리티 메신저 첫걸음 (PDF)', 'GOE메신저 쓰던 선생님을 위한 기본 사용법 7쪽 안내서. 기능 비교표와 3주 전환 로드맵 포함.', '7쪽', 'PDF,안내서,전환,비교표', false, true]
  ];
  seed.forEach(function (s) {
    var rec = { id: s[0], order: s[1], category: s[2], type: s[3], title: s[4], description: s[5], url: '', thumbnail: '',
                duration: s[6], tags: s[7], featured: s[8], isNew: s[9], visible: true, views: 0, createdAt: now, updatedAt: now };
    sh.appendRow(HEADERS.map(function (h) { return rec[h]; }));
  });
  sh.autoResizeColumns(1, HEADERS.length);
}
