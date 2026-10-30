/**
 * 🍇 포도알 공부 인증 — Google Sheets 백엔드 (Apps Script)
 * ---------------------------------------------------------------------------
 * 선생님 페이지(index.html)가 집계한 결과를 구글 시트에 쌓고,
 * 학생 페이지(student.html)가 그 현황을 읽어 갑니다.
 *
 * 시트 두 장을 씁니다 (없으면 자동으로 만들고 머리글도 넣습니다):
 *   students : id | 이름 | 닉네임들 | rewarded
 *   records  : 학생id | 날짜 | 출처 | 시각
 *
 * 쓰기(doPost)는 선생님만 아는 WRITE_KEY를 요구하고,
 * 조회(doGet)는 누구나 할 수 있게 열어 둡니다. (학생이 자기 포도송이를 봐야 하므로)
 *
 * ===========================================================================
 *  배포 방법 (5분)
 * ===========================================================================
 *  1. 구글 드라이브에서 새 스프레드시트를 하나 만듭니다. 이름은 아무거나.
 *     (시트 탭은 만들지 않아도 됩니다 — 첫 요청 때 자동으로 생깁니다)
 *
 *  2. 그 스프레드시트에서  확장 프로그램 → Apps Script  를 엽니다.
 *
 *  3. 기본으로 들어있는 코드를 다 지우고 이 파일 내용을 전부 붙여넣습니다.
 *
 *  4. 아래 WRITE_KEY 를 선생님만 아는 값으로 바꿉니다.
 *     예) var WRITE_KEY = 'oreum-grape-2026-mk';
 *     ※ 학생에게 알려 주는 조회 링크에는 이 키가 들어가지 않습니다.
 *
 *  5. 저장(💾) 후  배포 → 새 배포  를 누릅니다.
 *       - 유형 선택(⚙️) → 웹 앱
 *       - 설명         : 포도알 v1  (아무거나)
 *       - 실행 계정    : 나
 *       - 액세스 권한  : 모든 사용자            ← 중요! (학생이 로그인 없이 봐야 함)
 *     배포를 누르면 권한 승인을 한 번 요구합니다.
 *     "이 앱은 Google에서 확인하지 않았습니다" 경고가 나오면
 *     고급 → (안전하지 않음) 이동 → 허용 을 누르면 됩니다. (내가 만든 내 스크립트)
 *
 *  6. 끝나면 나오는  웹 앱 URL  을 복사합니다.
 *       https://script.google.com/macros/s/AKfy....../exec
 *     이 주소를 index.html → 설정 탭 → "구글 시트 연동" 칸에 붙여넣고,
 *     그 아래 쓰기 키 칸에는 4번에서 정한 WRITE_KEY 를 넣습니다.
 *
 *  7. 학생에게는 index.html 설정 탭의 "학생 조회 링크 복사" 버튼으로 나오는
 *     student.html?api=... 주소만 알려 주세요.
 *
 *  ※ 코드를 고친 뒤에는  배포 → 배포 관리 → ✏️(수정) → 버전: 새 버전 → 배포
 *     를 해야 반영됩니다. URL은 그대로 유지됩니다.
 *  ※ 동작 확인: 웹 앱 URL을 브라우저 주소창에 그대로 붙여 넣으면
 *     {"ok":true,...} JSON 이 보여야 합니다.
 * ===========================================================================
 */

/** 선생님만 아는 쓰기 키 — 반드시 바꾸세요 */
var WRITE_KEY = 'CHANGE-ME-포도알';

var SHEET_STUDENTS = 'students';
var SHEET_RECORDS  = 'records';
var HEAD_STUDENTS  = ['id', '이름', '닉네임들', 'rewarded'];
var HEAD_RECORDS   = ['학생id', '날짜', '출처', '시각'];

/** 설정(한 송이 알 수 등)은 시트를 더럽히지 않게 스크립트 속성에 둡니다 */
var PROP_CONFIG  = 'grapeConfig';
var PROP_SYNCED  = 'grapeSyncedAt';


/* ───────────────────────── 조회 (누구나) ───────────────────────── */

/**
 * GET → 전체 현황 JSON
 * 선생님 페이지의 localStorage 구조와 똑같은 모양으로 돌려줘서
 * student.html 이 그대로 쓸 수 있게 합니다.
 *
 *  {ok:true, updatedAt, config:{beadsPerBunch},
 *   students:[{id,name,aliases,rewarded}],
 *   records:{ s1:{ '2026-10-18':{src:'photo',at:'15:24'} } } }
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var students = readStudents_(ss);
    var records  = readRecords_(ss);

    var counts = {};
    for (var sid in records) counts[sid] = Object.keys(records[sid]).length;

    return json_({
      ok: true,
      updatedAt: PropertiesService.getScriptProperties().getProperty(PROP_SYNCED) || '',
      config: getConfig_(),
      students: students,
      records: records,
      counts: counts
    });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}


/* ───────────────────────── 쓰기 (선생님만) ───────────────────────── */

/**
 * POST → 집계 결과 반영
 *
 * 보내는 모양 (Content-Type 은 text/plain — Apps Script 의 CORS 사전요청을 피하려고):
 *  {
 *    key:    '쓰기 키',
 *    mode:   'merge' | 'replace',      // 기본 merge
 *    config: {beadsPerBunch:20},
 *    students:[{id,name,aliases:[],rewarded:0}],
 *    records: [{studentId,date,src,at}]   // 또는 {s1:{'2026-10-18':{src,at}}}
 *  }
 *
 *  students 는 보내 준 명단으로 통째로 맞춥니다 (이름·닉네임·보상 수정과 삭제가 반영됨).
 *  records 는
 *    merge   : 학생id+날짜가 이미 있으면 무시하고, 없는 것만 새 행으로 붙입니다. (기본)
 *    replace : records 시트를 비우고 보내 준 것으로 다시 씁니다.
 *              (선생님 페이지에서 인증을 지운 걸 시트에도 반영할 때)
 *
 * 돌려주는 모양: {ok:true, added, skipped, students, records, mode}
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = {};
    try {
      body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    } catch (parseErr) {
      return json_({ ok: false, error: '보낸 내용을 읽지 못했습니다 (JSON 형식 오류)' });
    }

    if (!WRITE_KEY || WRITE_KEY === 'CHANGE-ME-포도알') {
      return json_({ ok: false, error: 'Apps Script 의 WRITE_KEY 를 아직 바꾸지 않았습니다' });
    }
    if (String(body.key || '') !== WRITE_KEY) {
      return json_({ ok: false, error: '쓰기 키가 맞지 않습니다' });
    }

    lock.waitLock(25000);

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var mode = body.mode === 'replace' ? 'replace' : 'merge';

    // 1) 명단 — 통째로 맞춤
    var students = Array.isArray(body.students) ? body.students : [];
    writeStudents_(ss, students);

    // 2) 기록
    var incoming = normalizeRecords_(body.records);     // [{studentId,date,src,at}]
    var known = {};
    for (var i = 0; i < students.length; i++) known[String(students[i].id)] = true;
    incoming = incoming.filter(function (r) { return known[r.studentId]; });

    var sh = sheet_(ss, SHEET_RECORDS, HEAD_RECORDS);
    var added = 0, skipped = 0;

    if (mode === 'replace') {
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_RECORDS.length).clearContent();
      var rows = dedupe_(incoming);
      if (rows.length) {
        sh.getRange(2, 1, rows.length, HEAD_RECORDS.length).setValues(rows.map(recordRow_));
      }
      added = rows.length;
    } else {
      var have = readRecords_(ss);
      var fresh = [];
      var seen = {};
      for (var j = 0; j < incoming.length; j++) {
        var r = incoming[j];
        var k = r.studentId + '|' + r.date;
        if (seen[k] || (have[r.studentId] && have[r.studentId][r.date])) { skipped++; continue; }
        seen[k] = true;
        fresh.push(r);
      }
      if (fresh.length) {
        sh.getRange(sh.getLastRow() + 1, 1, fresh.length, HEAD_RECORDS.length)
          .setValues(fresh.map(recordRow_));
      }
      added = fresh.length;
    }

    // 3) 설정 + 동기화 시각
    var props = PropertiesService.getScriptProperties();
    if (body.config && typeof body.config === 'object') {
      props.setProperty(PROP_CONFIG, JSON.stringify({
        beadsPerBunch: Math.max(1, Math.min(100, parseInt(body.config.beadsPerBunch, 10) || 20))
      }));
    }
    var now = new Date().toISOString();
    props.setProperty(PROP_SYNCED, now);

    return json_({
      ok: true,
      mode: mode,
      added: added,
      skipped: skipped,
      students: students.length,
      records: Math.max(0, sh.getLastRow() - 1),
      updatedAt: now
    });

  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


/* ───────────────────────── 시트 입출력 ───────────────────────── */

/** 시트를 찾거나 머리글과 함께 새로 만든다 */
function sheet_(ss, name, head) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // 날짜·시각이 구글 시트의 날짜값으로 바뀌지 않게 열 전체를 '일반 텍스트'로
    if (name === SHEET_RECORDS) {
      sh.getRange(1, 2, sh.getMaxRows(), 3).setNumberFormat('@');
    }
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** students 시트 → [{id,name,aliases,rewarded}] */
function readStudents_(ss) {
  var sh = sheet_(ss, SHEET_STUDENTS, HEAD_STUDENTS);
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var vals = sh.getRange(2, 1, n, HEAD_STUDENTS.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][0] || '').trim();
    if (!id) continue;
    out.push({
      id: id,
      name: String(vals[i][1] || '').trim() || id,
      aliases: String(vals[i][2] || '').split(',').map(trim_).filter(Boolean),
      rewarded: parseInt(vals[i][3], 10) || 0
    });
  }
  return out;
}

/** [{id,name,aliases,rewarded}] → students 시트 (통째로 다시 씀) */
function writeStudents_(ss, students) {
  var sh = sheet_(ss, SHEET_STUDENTS, HEAD_STUDENTS);
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_STUDENTS.length).clearContent();
  }
  if (!students.length) return;
  var rows = students.map(function (s) {
    return [
      String(s.id || ''),
      String(s.name || ''),
      (Array.isArray(s.aliases) ? s.aliases : []).join(', '),
      parseInt(s.rewarded, 10) || 0
    ];
  }).filter(function (r) { return r[0]; });
  if (rows.length) sh.getRange(2, 1, rows.length, HEAD_STUDENTS.length).setValues(rows);
}

/** records 시트 → {학생id: {날짜: {src,at}}} */
function readRecords_(ss) {
  var sh = sheet_(ss, SHEET_RECORDS, HEAD_RECORDS);
  var n = sh.getLastRow() - 1;
  if (n < 1) return {};
  var vals = sh.getRange(2, 1, n, HEAD_RECORDS.length).getValues();
  var out = {};
  for (var i = 0; i < vals.length; i++) {
    var sid  = String(vals[i][0] || '').trim();
    var date = asDate_(vals[i][1]);
    if (!sid || !date) continue;
    if (!out[sid]) out[sid] = {};
    if (out[sid][date]) continue;                       // 하루 1알 — 먼저 적힌 것만
    out[sid][date] = { src: String(vals[i][2] || '').trim() || 'photo', at: asTime_(vals[i][3]) };
  }
  return out;
}

function recordRow_(r) {
  return [r.studentId, r.date, r.src, r.at];
}

/** 선생님 페이지가 보낸 records 를 [{studentId,date,src,at}] 로 통일 */
function normalizeRecords_(src) {
  var out = [];
  if (Array.isArray(src)) {
    for (var i = 0; i < src.length; i++) {
      var r = src[i] || {};
      var d = asDate_(r.date);
      if (!r.studentId || !d) continue;
      out.push({
        studentId: String(r.studentId),
        date: d,
        src: String(r.src || 'photo'),
        at: String(r.at || '')
      });
    }
  } else if (src && typeof src === 'object') {
    for (var sid in src) {
      var byDate = src[sid] || {};
      for (var date in byDate) {
        var dd = asDate_(date);
        if (!dd) continue;
        var info = byDate[date] || {};
        out.push({
          studentId: String(sid),
          date: dd,
          src: String(info.src || 'photo'),
          at: String(info.at || '')
        });
      }
    }
  }
  out.sort(function (a, b) {
    return a.studentId < b.studentId ? -1 : a.studentId > b.studentId ? 1
         : a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
  });
  return out;
}

function dedupe_(list) {
  var seen = {}, out = [];
  for (var i = 0; i < list.length; i++) {
    var k = list[i].studentId + '|' + list[i].date;
    if (seen[k]) continue;
    seen[k] = true;
    out.push(list[i]);
  }
  return out;
}


/* ───────────────────────── 작은 도구들 ───────────────────────── */

/** 날짜값인지 — instanceof 대신 모양으로 판단 (구글 시트가 돌려주는 값이 realm을 넘나들 수 있음) */
function isDateLike_(v) {
  return !!v && typeof v === 'object' && typeof v.getFullYear === 'function';
}

/** 셀 값이 문자열이든 구글 시트 날짜값이든 'YYYY-MM-DD' 로 */
function asDate_(v) {
  if (isDateLike_(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  }
  // 'YYYY-MM-DD' 로 시작하면 뒤에 시각 등이 붙어 있어도 날짜만 떼어 쓴다
  var m = String(v == null ? '' : v).trim().match(/^(\d{4})\D(\d{1,2})\D(\d{1,2})(?:\D|$)/);
  if (!m) return '';
  return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
}

/** 셀 값이 시간값이면 'HH:mm' 으로, 아니면 그대로 */
function asTime_(v) {
  if (isDateLike_(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Seoul', 'HH:mm');
  }
  return String(v == null ? '' : v).trim();
}

function pad2_(n) { return String(n).length < 2 ? '0' + n : String(n); }
function trim_(s) { return String(s).trim(); }

function getConfig_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PROP_CONFIG);
    var c = raw ? JSON.parse(raw) : {};
    return { beadsPerBunch: parseInt(c.beadsPerBunch, 10) || 20 };
  } catch (err) {
    return { beadsPerBunch: 20 };
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ───────────────────────── 점검용 (스크립트 편집기에서 실행) ───────────────────────── */

/** ▶ 실행 버튼으로 돌려 보면 시트 두 장이 머리글과 함께 만들어집니다 */
function 시트_준비() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, SHEET_STUDENTS, HEAD_STUDENTS);
  sheet_(ss, SHEET_RECORDS, HEAD_RECORDS);
  Logger.log('students / records 시트 준비 완료');
}

/** ▶ 지금 시트에 무엇이 들어 있는지 로그로 확인 */
function 현황_확인() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = readStudents_(ss);
  var rc = readRecords_(ss);
  var total = 0;
  for (var k in rc) total += Object.keys(rc[k]).length;
  Logger.log('학생 ' + st.length + '명 · 포도알 ' + total + '개 · 한 송이 ' + getConfig_().beadsPerBunch + '알');
  Logger.log(JSON.stringify(st));
}
