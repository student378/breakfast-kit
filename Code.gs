// ═══════════════════════════════════════════════════════════════
//  천원의 아침밥 신청·관리 — Google Apps Script (배포용 kit v1.0)
//  원본: 유한대학교 학생복지팀 운영본 v10 을 학교별 설정형으로 바꾼 것
//
//  ▶ 설치 (자세한 내용은 INSTALL.md)
//   1. 새 구글 스프레드시트 → 확장 프로그램 → Apps Script
//   2. 기본 Code.gs 내용을 모두 지우고 이 파일을 붙여넣기 → 저장
//   3. 편집기 위쪽 함수 목록에서 setupSheets 선택 → 실행 (처음 한 번, 권한 허용)
//      → 신청목록·식단·설정·학과목록·승인현황 시트가 만들어집니다.
//   4. 스프레드시트를 새로고침하면 위쪽에 「🍱 천원의 아침밥」 메뉴가 생깁니다.
//      → 「🔑 관리자 비밀번호 설정」에서 비밀번호를 정하세요. (코드·화면에 저장되지 않음)
//   5. 「설정」 시트에서 학교명·학번 자리수·학년·계좌 등을, 「학과목록」 시트에 학과를 입력
//   6. 배포 → 새 배포 → 유형 「웹 앱」, 실행 계정 「나」, 액세스 「모든 사용자」
//      → 받은 주소(…/exec)를 config.js 에 붙여넣기
//   7. (권장) 트리거 → keepAlive → 시간 기반 → 5분마다
//
//  ▶ 코드를 고친 뒤 반영: 배포 → 배포 관리 → 연필(편집) → 버전 "새 버전" → 배포 (주소 유지)
//
//  ▶ 학교마다 달라지는 값은 모두 「설정」·「학과목록」 시트에 있습니다. 코드는 고치지 않아도 됩니다.
//    「설정」 시트에 값이 비어 있으면 아래 SCHOOL_DEFAULTS 값을 씁니다.
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = 'Asia/Seoul';

// ── 설정값 (필요 시 여기서만 조정) ──
const CFG = {
  ENFORCE_TODAY_ONLY : true,   // 신청 날짜는 '오늘'만 허용
  ENFORCE_START_TIME : true,   // 관리자 설정의 신청 시작 시간 이전 차단
  ENFORCE_END_TIME   : true,   // 관리자 설정의 신청 종료 시간 이후 차단 (계속 열어두려면 false)
  MAX_PER_DAY        : 1000,   // 하루 신청 상한 (비정상 대량 신청 방지)
  STATS_CACHE_SEC    : 5,      // getStats 캐시 시간(초)
  TOKEN_TTL          : 21600,  // 관리자 세션 유지(초) — 캐시 최대 6시간, 사용 중이면 자동 연장
  MAX_FAIL           : 10,     // 로그인 연속 실패 허용 횟수
  LOCK_SEC           : 600,    // 초과 시 잠금(초)

  // ── 비식별화 (사업 종료 후 1회 실행) ──
  //  대상 기간은 「설정」 시트의 anonFrom / anonTo 에 입력합니다.
  ANON_CONFIRM       : false           // anonymizeRun 실행 직전에만 true 로 변경 (되돌릴 수 없음)
};

// ── 「설정」 시트에 값이 없을 때 쓰는 기본값 ──
//   (기존 운영 시트에 아래 키가 없어도 지금과 같이 동작하도록 학번 9자리·1~3학년을 기본값으로 둠)
const SCHOOL_DEFAULTS = {
  schoolName       : '',
  schoolNameEn     : '',
  orgName          : '',
  contact          : '',
  idDigits         : '9',
  grades           : '1학년,2학년,3학년',
  payMethods       : '계좌이체,현금지급',
  privacyRetention : '1년(해당 사업 종료 시까지)',
  privacyNote      : '사업 종료 후 성명·학번은 삭제하고, 학과·학년별 통계로만 활용합니다.'
};

// 처음 설치할 때 「설정」 시트에 만드는 행 — [키, 값, 설명]
const SETTING_ROWS = [
  ['schoolName',      '○○대학교',           '학교명 (학생·관리자 화면 머리글, 실적 점검표)'],
  ['schoolNameEn',    '',                    '영문 학교명 (비우면 표시 안 함)'],
  ['orgName',         '',                    '운영 부서 (예: 학생처 학생복지팀)'],
  ['contact',         '',                    '문의처 (예: 학생복지팀 02-000-0000) — 학생 화면 아래에 표시'],
  ['idDigits',        '9',                   '학번 자리수 — 숫자 하나(9), 여러 개(8,10), 범위(8-10)'],
  ['grades',          '1학년,2학년,3학년',     '학년 선택지 (쉼표로 구분)'],
  ['payMethods',      '계좌이체,현금지급',      '송금 방법 — 계좌이체·현금지급 중 사용할 것 (쉼표로 구분)'],
  ['privacyRetention','1년(해당 사업 종료 시까지)', '개인정보 보유·이용 기간 — 학교 개인정보 담당 부서 기준으로 입력'],
  ['privacyNote',     '사업 종료 후 성명·학번은 삭제하고, 학과·학년별 통계로만 활용합니다.', '보유 기간 아래 안내 문구'],
  ['startTime',       '08:30',               '신청 시작 시간 (HH:MM)'],
  ['endTime',         '10:30',               '신청 종료 시간 (HH:MM)'],
  ['soldoutManual',   'false',               '수동 판매 종료 (true / false) — 관리자 화면에서 바꿈'],
  ['soldoutMsg',      '금일 천원의 아침밥 판매가 종료되었습니다. 감사합니다.', '판매 종료 안내 문구'],
  ['bank',            '',                    '입금 은행'],
  ['bankNo',          '',                    '계좌번호'],
  ['bankOwner',       '',                    '예금주'],
  ['bankAmt',         '1,000원',              '결제 금액 표기'],
  ['anonFrom',        '',                    '비식별화 대상 이용일 시작 (YYYY-MM-DD) — 사업 종료 후 사용'],
  ['anonTo',          '',                    '비식별화 대상 이용일 끝 (YYYY-MM-DD)']
];

const APPLY_HEADERS = ['접수번호','신청일시','이용일','학과','학년','학번','성명','신청시간','메뉴','결제방법','개인정보동의'];

const PUBLIC_SETTING_KEYS = ['bank','bankNo','bankOwner','bankAmt','startTime','endTime','soldoutManual','soldoutMsg',
                             'schoolName','schoolNameEn','orgName','contact','idDigits','grades','payMethods','privacyRetention','privacyNote'];
// 관리자 화면에서 저장할 수 있지만 학생 화면에는 보내지 않는 키
const PRIVATE_SETTING_KEYS = ['anonFrom','anonTo'];

// 토큰이 있어야 실행되는 동작
const PROTECTED = { getAll:1, getLive:1, saveSettings:1, saveMenu:1, saveDepts:1, getAdminSettings:1, checkToken:1, adminLogout:1,
                    getApprovals:1, setApproval:1, setApprovalBulk:1, clearApprovalDate:1 };

// ─────────────────────────────────────────────────────────────
//  doGet
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  const p        = (e && e.parameter) ? e.parameter : {};
  const action   = p.action || 'getAll';      // v8과 동일하게 기본값 유지 — 단, getAll 은 토큰 필요
  const callback = p.callback;
  let result;

  try {
    if (action === 'adminLogin') {
      result = bfLogin_(p.pw);
    } else if (PROTECTED[action] && !bfTokenOk_(p.token)) {
      result = { status:'unauthorized' };
    } else {
      switch (action) {
        case 'save':         result = saveData(p);                 break;
        case 'getAll':       result = getAllData();                break;
        case 'getLive':      result = getLive_(p.date);            break;
        case 'getStats':     result = getStats(p.date);            break;
        case 'getSettings':  result = getSettings();               break;
        case 'saveSettings': result = bfWithLock_(function(){ return saveSettings(p); }); break;
        case 'getAdminSettings': result = getAdminSettings_();     break;
        case 'saveDepts':    result = bfWithLock_(function(){ return saveDepts_(p); }); break;
        case 'getMenu':      result = getMenuData();               break;
        case 'saveMenu':     result = bfWithLock_(function(){ return saveMenuData(p); }); break;
        // ── 관리자 세션 / 승인 ──
        case 'checkToken':   result = { status:'ok' };             break;
        case 'adminLogout':  CacheService.getScriptCache().remove('bf_tok_' + String(p.token)); result = { status:'ok' }; break;
        case 'getApprovals': result = { status:'ok', keys: bfGetApprovalKeys_() }; break;
        case 'setApproval':       result = bfWithLock_(function(){ return bfSetApprovals_(p.date, [p.sid], p.on === '1'); }); break;
        case 'setApprovalBulk':   result = bfWithLock_(function(){ return bfSetApprovals_(p.date, String(p.sids || '').split(','), p.on === '1'); }); break;
        case 'clearApprovalDate': result = bfWithLock_(function(){ return bfClearDate_(p.date); }); break;
        default:             result = { status:'error', message:'Unknown action: ' + action };
      }
    }
  } catch(err) {
    result = { status:'error', message: err.toString() };
  }

  const json = JSON.stringify(result);
  if (callback && /^[A-Za-z0-9_.$]+$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
//  saveData — 신청 저장 (LockService 직렬화 + 서버 검증)
// ─────────────────────────────────────────────────────────────
function cleanText_(s, maxLen) {
  return String(s == null ? '' : s).replace(/[<>"'`&;=\u0000-\u001f]/g, '').trim().slice(0, maxLen || 100);
}

function toMin_(t, def) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(t || '').trim());
  return m ? (parseInt(m[1], 10) * 60 + parseInt(m[2], 10)) : def;
}

function saveData(p) {
  // ── 입력 검사 ──
  const date = String(p.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date))        return { status:'invalid', message:'날짜 형식이 올바르지 않습니다.' };
  const school    = schoolConf_();
  const studentId = String(p.studentId || '').trim();
  if (!/^\d+$/.test(studentId) || school.idLens.indexOf(studentId.length) < 0)
                                                return { status:'invalid', message:'학번은 ' + school.idLabel + ' 숫자로 입력해주세요.' };
  const name = cleanText_(p.name, 20);
  if (name.length < 2)                          return { status:'invalid', message:'성명을 확인해주세요.' };
  const dept = cleanText_(p.dept, 40);
  if (!dept)                                    return { status:'invalid', message:'학과를 확인해주세요.' };
  if (school.depts.length && school.depts.indexOf(dept) < 0)
                                                return { status:'invalid', message:'학과 목록에 없는 학과입니다. 화면을 새로고침한 뒤 다시 선택해주세요.' };
  const grade = cleanText_(p.grade, 10);
  if (school.grades.indexOf(grade) < 0)         return { status:'invalid', message:'학년을 확인해주세요.' };
  const payMethod = cleanText_(p.payMethod, 10);
  if (school.payMethods.indexOf(payMethod) < 0) return { status:'invalid', message:'송금 방법을 확인해주세요.' };
  const menu = cleanText_(p.menu, 100);
  if (String(p.agree || '') !== 'Y')            return { status:'invalid', message:'개인정보 수집·이용에 동의해야 신청할 수 있습니다.' };

  const now     = new Date();
  const todayD  = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  if (CFG.ENFORCE_TODAY_ONLY && date !== todayD) {
    return { status:'invalid', message:'신청은 해당 날짜 당일에만 가능합니다.' };
  }

  // 신청 시간 (관리자 설정)
  const settings = getSettingsMap();
  const nowMin   = parseInt(Utilities.formatDate(now, TZ, 'H'), 10) * 60 + parseInt(Utilities.formatDate(now, TZ, 'm'), 10);
  const startMin = toMin_(settings['startTime'], 8 * 60 + 25);
  const endMin   = toMin_(settings['endTime'],  10 * 60 + 30);
  if (CFG.ENFORCE_START_TIME && nowMin < startMin) {
    return { status:'invalid', message:'신청 시간이 아닙니다. ' + settings['startTime'] + ' 이후에 신청해주세요.' };
  }
  if (CFG.ENFORCE_END_TIME && nowMin > endMin) {
    return { status:'invalid', message:'신청 시간이 종료되었습니다.' };
  }
  if (settings['soldoutManual'] === 'true') {
    return { status:'soldout', message:'판매가 종료되었습니다.' };
  }

  // 운영일 / 수량
  const menuSheet = SS.getSheetByName('식단');
  let capacity = 100;
  if (menuSheet) {
    // getDisplayValues() — 날짜 셀 Date 자동변환 방지
    const found = menuSheet.getDataRange().getDisplayValues().slice(1)
      .find(r => String(r[0]).trim() === date);
    if (!found) return { status:'invalid', message:'신청 가능한 운영일이 아닙니다.' };
    capacity = parseInt(found[2]) || 100;
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { status:'busy', message:'접속이 집중되고 있습니다. 잠시 후 자동으로 재시도됩니다.' };
  }

  try {
    const sheet = getOrCreateSheet('신청목록');
    ensureHeaders(sheet, APPLY_HEADERS);
    const headers = ensureColumns_(sheet, APPLY_HEADERS);   // 기존 시트에 없는 열(개인정보동의)은 오른쪽에 추가
    const cols    = readDateSid_(sheet, headers);           // 이용일·학번 두 열만 읽음

    // ① 중복 체크
    let currentSold = 0;
    for (let i = 0; i < cols.dates.length; i++) {
      if (cols.dates[i] !== date) continue;
      if (cols.sids[i] === studentId) return { status:'duplicate', message:'이미 해당 날짜에 신청한 학번입니다.' };
      currentSold++;
    }

    // ② 수량 — 초과 신청 허용(현장 승인으로 관리), 단 하루 상한은 둠
    if (currentSold >= CFG.MAX_PER_DAY) {
      return { status:'soldout', message:'오늘 신청이 마감되었습니다.' };
    }

    // ③ 저장 — 신청시간은 서버 시각 사용 (기기 시계 오차 방지, 관리자 정렬 기준)
    const timestamp = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss');
    const applyTime = Utilities.formatDate(now, TZ, 'HH:mm:ss');
    const today     = Utilities.formatDate(now, TZ, 'yyyyMMdd');
    const seq       = String(currentSold + 1).padStart(4, '0');
    const receiptNo = 'BF' + today + '-' + seq;

    const rec = { '접수번호':receiptNo, '신청일시':timestamp, '이용일':date, '학과':dept, '학년':grade, '학번':studentId,
                  '성명':name, '신청시간':applyTime, '메뉴':menu, '결제방법':payMethod, '개인정보동의':'동의 ' + timestamp };
    const row = headers.map(h => (rec[h] !== undefined ? rec[h] : ''));   // 열 순서가 바뀌어 있어도 제목 기준으로 저장

    // 값이 날짜/시간/숫자로 자동 변환되지 않도록, 저장할 행을 먼저 '텍스트' 서식으로 만든 뒤 같은 범위에 기록
    // (appendRow 는 '09:01:02' 같은 문자열을 시간 값으로 바꿔 저장해 관리자 화면에 1899-12-30… 으로 나타났음)
    const newRow = sheet.getLastRow() + 1;
    if (newRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), newRow - sheet.getMaxRows());
    sheet.getRange(newRow, 1, 1, row.length).setNumberFormat('@').setValues([row]);

    try { CacheService.getScriptCache().remove('bf_stats_' + date); } catch(e) {}

    return { status:'ok', receiptNo, applyTime, remain: Math.max(0, capacity - currentSold - 1), overCapacity: (currentSold >= capacity) };

  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  getAllData  (토큰 필요)
// ─────────────────────────────────────────────────────────────
function getAllData() {
  return getAllDataRaw(getOrCreateSheet('신청목록')).map(r => ({
    receiptNo: r['접수번호'] || '',
    timestamp: r['신청일시'] ? String(r['신청일시']) : '',
    date:      r['이용일']   || '',
    dept:      r['학과']     || '',
    grade:     r['학년']     || '',
    studentId: String(r['학번'] || ''),
    name:      r['성명']     || '',
    applyTime: r['신청시간'] || '',
    menu:      r['메뉴']     || '',
    payMethod: r['결제방법'] || ''
  }));
}

function getAllDataRaw(sheet) {
  if (!sheet) sheet = getOrCreateSheet('신청목록');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0].map(h => String(h).trim());
  // 신청시간: 시트에 시간 값으로 저장된 행(예전 저장 방식)도 화면에 보이는 시각 그대로 읽기 위해 표시값을 함께 읽음
  const iT = headers.indexOf('신청시간');
  const dispT = (iT >= 0) ? sheet.getRange(2, iT + 1, rows.length - 1, 1).getDisplayValues() : [];
  return rows.slice(1).map((row, r) => {
    const o = {};
    headers.forEach((h, i) => { o[h] = row[i]; });
    // 이용일이 날짜 값으로 저장돼 있어도 'yyyy-MM-dd' 문자열로 통일
    if (o['이용일'] instanceof Date) o['이용일'] = Utilities.formatDate(o['이용일'], TZ, 'yyyy-MM-dd');
    if (iT >= 0) o['신청시간'] = normTime_(row[iT], dispT[r] ? dispT[r][0] : '');
    return o;
  });
}

// 신청시간을 'HH:mm:ss' 로 통일
//  - 문자열 '9:01:02' / '09:01' → 자리수 보정
//  - 시간 값(시트가 자동 변환한 경우) → 시트에 표시된 값('오후 9:01:02', 'AM 9:01:02' 등)을 24시간제로 변환
//  - 하루 중 비율 숫자(0~1) → 시각 계산
function normTime_(v, disp) {
  const pad = n => String(n).padStart(2, '0');
  if (typeof v === 'string') {
    const m = /^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$/.exec(v);
    if (m) return pad(m[1]) + ':' + m[2] + ':' + pad(m[3] || 0);
  }
  if (typeof v === 'number' && v >= 0 && v < 1) {
    const sec = Math.round(v * 86400) % 86400;
    return pad(Math.floor(sec / 3600)) + ':' + pad(Math.floor(sec % 3600 / 60)) + ':' + pad(sec % 60);
  }
  const d = String(disp == null ? '' : disp).trim();
  const m2 = /(오전|오후|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|am|pm)?/.exec(d);
  if (m2) {
    let h = parseInt(m2[2], 10);
    const ap = (m2[1] || m2[5] || '').toLowerCase();
    if (ap === '오후' || ap === 'pm') { if (h < 12) h += 12; }
    else if (ap === '오전' || ap === 'am') { if (h === 12) h = 0; }
    return pad(h) + ':' + m2[3] + ':' + pad(m2[4] || 0);
  }
  return (v == null) ? '' : String(v);
}

// ─────────────────────────────────────────────────────────────
//  getStats (공개) — 5초 캐시
// ─────────────────────────────────────────────────────────────
function getStats(date) {
  date = String(date == null ? '' : date).trim();
  const cacheable = /^\d{4}-\d{2}-\d{2}$/.test(date);
  const cache = CacheService.getScriptCache();
  const key = 'bf_stats_' + date;
  if (cacheable) {
    const hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch(e) {} }
  }
  const r = computeStats_(date);
  if (cacheable) { try { cache.put(key, JSON.stringify(r), CFG.STATS_CACHE_SEC); } catch(e) {} }
  return r;
}

function computeStats_(date) {
  const menuSheet = SS.getSheetByName('식단');
  let capacity = 100;
  if (menuSheet) {
    const found = menuSheet.getDataRange().getDisplayValues().slice(1)
      .find(r => String(r[0]).trim() === date);
    if (found) capacity = parseInt(found[2]) || 100;
  }

  const settings      = getSettingsMap();
  const soldoutManual = settings['soldoutManual'] === 'true';
  const soldoutMsg    = settings['soldoutMsg'] || '금일 천원의 아침밥 판매가 종료되었습니다. 감사합니다.';

  if (soldoutManual) {
    return { status:'ok', date, total:capacity, sold:capacity, remain:0, soldoutManual:true, soldoutMsg };
  }

  const sheet  = getOrCreateSheet('신청목록');
  const sold   = (sheet.getLastRow() < 2) ? 0
               : readDateSid_(sheet, headersOf_(sheet)).dates.filter(d => d === date).length;
  const remain = Math.max(0, capacity - sold);

  return { status:'ok', date, total:capacity, sold, remain, soldoutManual:false, soldoutMsg };
}

// ─────────────────────────────────────────────────────────────
//  설정 읽기/쓰기  — Date 자동변환 방지
// ─────────────────────────────────────────────────────────────
function getSettings() {
  const all = getSettingsMap();
  const data = {};
  PUBLIC_SETTING_KEYS.forEach(k => {                                                    // 공개 키만 반환
    if (all[k] !== undefined && all[k] !== '') data[k] = all[k];
    else if (SCHOOL_DEFAULTS[k] !== undefined) data[k] = SCHOOL_DEFAULTS[k];
  });
  data.depts = getDepts_();
  return { status:'ok', data };
}

// 관리자 화면용 — 공개 키 + 비공개 키(비식별화 기간)
function getAdminSettings_() {
  const r = getSettings();
  const all = getSettingsMap();
  PRIVATE_SETTING_KEYS.forEach(k => { r.data[k] = all[k] || ''; });
  return r;
}

// 학교 기준값 — 신청 검증에 사용
function schoolConf_() {
  const m = getSettingsMap();
  const v = k => (m[k] !== undefined && m[k] !== '') ? m[k] : SCHOOL_DEFAULTS[k];
  const list = k => String(v(k) || '').split(',').map(x => x.trim()).filter(Boolean);
  const idLens = parseDigits_(v('idDigits'));
  return {
    idLens : idLens.length ? idLens : [9],
    idLabel: digitsLabel_(idLens.length ? idLens : [9]),
    grades : list('grades'),
    payMethods: list('payMethods'),
    depts  : getDepts_()
  };
}

// '9' → [9] / '8,10' → [8,10] / '8-10' → [8,9,10]  (1~20자리만 인정)
function parseDigits_(t) {
  const out = [];
  String(t || '').split(',').forEach(part => {
    const m = /^\s*(\d{1,2})\s*(?:-|~)\s*(\d{1,2})\s*$/.exec(part);
    if (m) { for (let n = +m[1]; n <= +m[2]; n++) out.push(n); return; }
    const n = parseInt(String(part).trim(), 10);
    if (n) out.push(n);
  });
  return out.filter((n, i, a) => n >= 1 && n <= 20 && a.indexOf(n) === i).sort((a, b) => a - b);
}
function digitsLabel_(lens) {
  if (lens.length === 1) return lens[0] + '자리';
  const contiguous = lens.every((n, i) => i === 0 || n === lens[i - 1] + 1);
  return contiguous ? (lens[0] + '~' + lens[lens.length - 1] + '자리') : (lens.join('·') + '자리');
}

// 「학과목록」 시트 A열 (2행부터). 시트가 없거나 비어 있으면 [] → 학과 검사 생략
function getDepts_() {
  const sh = SS.getSheetByName('학과목록');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues()
    .map(r => String(r[0]).trim()).filter(Boolean)
    .filter((d, i, a) => a.indexOf(d) === i);
}

// 관리자 화면 「학과 목록 저장」 — 줄바꿈으로 구분한 목록 전체를 교체
function saveDepts_(p) {
  const list = String(p.depts || '').split('\n').map(d => cleanText_(d, 40)).filter(Boolean)
    .filter((d, i, a) => a.indexOf(d) === i).slice(0, 300);
  if (!list.length) return { status:'invalid', message:'학과가 하나 이상 있어야 합니다.' };
  const sh = getOrCreateSheet('학과목록');
  if (sh.getLastRow() === 0) sh.getRange(1, 1).setValue('학과명');
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).clearContent();
  sh.getRange(2, 1, list.length, 1).setNumberFormat('@').setValues(list.map(d => [d]));
  return { status:'ok', count: list.length };
}

function getSettingsMap() {
  const sheet = SS.getSheetByName('설정');
  if (!sheet) return {};
  const map = {};
  sheet.getDataRange().getDisplayValues().slice(1).forEach(r => {
    if (r[0]) map[String(r[0]).trim()] = String(r[1] || '').trim();
  });
  return map;
}

function saveSettings(p) {
  const sheet = getOrCreateSheet('설정');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['key', 'value']);
    sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const rows = sheet.getDataRange().getValues();

  PUBLIC_SETTING_KEYS.concat(PRIVATE_SETTING_KEYS).forEach(key => {
    if (p[key] === undefined) return;
    const val = String(p[key] || '').slice(0, 300);
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === key) {
        sheet.getRange(i+1, 2).setNumberFormat('@').setValue(val);
        found = true; break;
      }
    }
    if (!found) {
      const nr = sheet.getLastRow() + 1;
      sheet.getRange(nr, 1).setValue(key);
      sheet.getRange(nr, 2).setNumberFormat('@').setValue(val);
    }
  });
  return { status:'ok' };
}

// ─────────────────────────────────────────────────────────────
//  식단 읽기 (공개) — 스프레드시트 '식단' 시트를 직접 읽음
//  형식: A열=YYYY-MM-DD, B열=메뉴명(콤마 구분), C열=수량
// ─────────────────────────────────────────────────────────────
function getMenuData() {
  const sheet = SS.getSheetByName('식단');
  if (!sheet) return { status:'ok', data:{} };

  const data = {};
  sheet.getDataRange().getDisplayValues().slice(1).forEach(r => {
    const date = String(r[0]).trim();
    const menu = String(r[1]).trim();
    const qty  = parseInt(r[2]) || 100;
    if (date && menu && /^\d{4}-\d{2}-\d{2}$/.test(date)) {
      data[date] = { menus: menu.split(',').map(m => m.trim()), qty };
    }
  });
  return { status:'ok', data };
}

// ─────────────────────────────────────────────────────────────
//  식단 저장 (토큰 필요)
//  관리자 화면은 식단을 5일씩 나눠(chunk) 순서대로 보낸다.
//   - isFirst === 'true' 인 첫 묶음에서만 기존 식단을 지우고
//   - 이후 묶음은 이어서 추가한다.   (chunkIndex 가 없는 단발 요청은 v8과 동일하게 전체 교체)
// ─────────────────────────────────────────────────────────────
function saveMenuData(p) {
  try {
    const sheet = getOrCreateSheet('식단');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['date', 'menu', 'qty']);
      sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    const chunked = (p.chunkIndex !== undefined);
    const isFirst = !chunked || String(p.isFirst) === 'true';
    if (isFirst) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    }

    const data = JSON.parse(p.menuJson || '{}');
    const rows = [];
    Object.keys(data).sort().forEach(date => {
      const info = data[date] || {};
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !Array.isArray(info.menus)) return;
      const menus = info.menus.map(m => String(m).trim()).filter(Boolean);
      if (!menus.length) return;
      const qty = Math.min(9999, Math.max(1, parseInt(info.qty, 10) || 100));
      rows.push([date, menus.join(', '), qty]);
    });

    if (rows.length > 0) {
      const start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, rows.length, 2).setNumberFormat('@');   // 날짜·메뉴가 날짜/수식으로 변환되지 않도록
      sheet.getRange(start, 1, rows.length, 3).setValues(rows);
    }
    return { status:'ok', saved: rows.length };
  } catch(err) {
    return { status:'error', message: err.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  유틸
// ─────────────────────────────────────────────────────────────
function getOrCreateSheet(name) {
  return SS.getSheetByName(name) || SS.insertSheet(name);
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

function headersOf_(sheet) {
  const lc = sheet.getLastColumn();
  if (lc < 1) return [];
  return sheet.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h).trim());
}

// 필요한 열 제목이 없으면 맨 오른쪽에 추가하고, 최종 제목 배열을 반환
function ensureColumns_(sheet, need) {
  const headers = headersOf_(sheet);
  need.forEach(h => {
    if (headers.indexOf(h) < 0) {
      headers.push(h);
      sheet.getRange(1, headers.length).setValue(h)
        .setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
    }
  });
  return headers;
}

// '이용일'·'학번' 두 열만 읽어 정규화 (중복 확인·수량 계산용)
function readDateSid_(sheet, headers) {
  const last = sheet.getLastRow();
  const iD = headers.indexOf('이용일'), iS = headers.indexOf('학번');
  if (last < 2 || iD < 0 || iS < 0) return { dates:[], sids:[] };
  const dates = sheet.getRange(2, iD + 1, last - 1, 1).getValues().map(r => bfNorm_(r[0]));
  const sids  = sheet.getRange(2, iS + 1, last - 1, 1).getValues().map(r => bfNorm_(r[0]));
  return { dates, sids };
}

// ─────────────────────────────────────────────────────────────
//  getLive (토큰 필요) — 관리자 실시간 화면: 선택 날짜 신청자 + 승인 + 현황
// ─────────────────────────────────────────────────────────────
function getLive_(date) {
  date = String(date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { status:'error', message:'bad date' };
  const rows = getAllData().filter(r => bfNorm_(r.date) === date);
  const keys = bfGetApprovalKeys_().filter(k => k.indexOf(date + '_') === 0);
  return { status:'ok', date, rows, keys, stats: getStats(date) };
}

// ─────────────────────────────────────────────────────────────
//  비식별화 — 사업 종료 후 편집기에서 실행
//   1) anonymizePreview 실행 → 실행 로그에서 대상 건수 확인 (시트는 바뀌지 않음)
//   2) CFG.ANON_CONFIRM 을 true 로 저장 → anonymizeRun 실행 → 다시 false 로 저장
//  처리 내용 (「설정」 시트 anonFrom ~ anonTo 이용일 대상)
//   - '신청목록' 성명 → 빈칸, 학번 → 'DEL-행번호' (원래 학번과 연결할 수 없는 일련번호)
//   - '승인현황' 학번도 같은 일련번호로 바꿔 승인 여부(이용 인원 통계)는 유지
//   - 학과·학년·이용일·신청시간·결제방법은 그대로 남음
// ─────────────────────────────────────────────────────────────
function anonRange_() {
  const m = getSettingsMap();
  return [String(m['anonFrom'] || '').trim(), String(m['anonTo'] || '').trim()];
}
function anonymizePreview() {
  const r = anonRange_();
  Logger.log(JSON.stringify(anonymize_(r[0], r[1], true)));
}
function anonymizeRun() {
  if (CFG.ANON_CONFIRM !== true) throw new Error('되돌릴 수 없는 작업입니다. CFG.ANON_CONFIRM 을 true 로 바꿔 저장한 뒤 실행하세요.');
  const r = anonRange_();
  Logger.log(JSON.stringify(anonymize_(r[0], r[1], false)));
}

function anonymize_(from, to, dryRun) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) throw new Error('「설정」 시트의 anonFrom / anonTo (YYYY-MM-DD) 를 확인하세요');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SS.getSheetByName('신청목록');
    if (!sheet || sheet.getLastRow() < 2) return { dryRun, rows:0, approvalsKept:0, approvalsDeleted:0 };
    const headers = headersOf_(sheet);
    const iD = headers.indexOf('이용일'), iS = headers.indexOf('학번'), iN = headers.indexOf('성명');
    if (iD < 0 || iS < 0 || iN < 0) throw new Error('신청목록 열 제목(이용일/학번/성명)을 찾을 수 없습니다');

    const n = sheet.getLastRow() - 1;
    const dCol = sheet.getRange(2, iD + 1, n, 1).getValues();
    const sCol = sheet.getRange(2, iS + 1, n, 1).getValues();
    const nCol = sheet.getRange(2, iN + 1, n, 1).getValues();
    const map = {};            // '이용일_원래학번' → 새 일련번호
    let rows = 0;
    for (let i = 0; i < n; i++) {
      const d = bfNorm_(dCol[i][0]);
      if (!d || d < from || d > to) continue;
      const sid = bfNorm_(sCol[i][0]);
      if (sid.indexOf('DEL-') === 0 && !nCol[i][0]) continue;     // 이미 처리된 행
      const newId = 'DEL-' + String(i + 2).padStart(6, '0');        // 시트 행번호 기반 → 행마다 고유
      if (sid) map[d + '_' + sid] = newId;
      sCol[i][0] = newId; nCol[i][0] = ''; rows++;
    }

    // 승인현황: 대상 기간 행은 새 일련번호로 바꾸고, 신청목록에 없는(짝 없는) 학번은 삭제
    const ash = bfSheet_();
    const alast = ash.getLastRow();
    const aVals = alast >= 2 ? ash.getRange(2, 1, alast - 1, 3).getValues() : [];
    const kept = [];
    let approvalsDeleted = 0, approvalsKept = 0;
    aVals.forEach(r => {
      const d = bfNorm_(r[0]), sid = bfNorm_(r[1]);
      if (!d || d < from || d > to || sid.indexOf('DEL-') === 0) { kept.push([d, sid, r[2]]); return; }
      const newId = map[d + '_' + sid];
      if (newId) { kept.push([d, newId, r[2]]); approvalsKept++; }
      else approvalsDeleted++;
    });

    if (!dryRun) {
      sheet.getRange(2, iS + 1, n, 1).setNumberFormat('@').setValues(sCol);
      sheet.getRange(2, iN + 1, n, 1).setValues(nCol);
      if (alast >= 2) {
        ash.getRange(2, 1, alast - 1, 3).clearContent();
        if (kept.length) {
          ash.getRange(2, 1, kept.length, 2).setNumberFormat('@');
          ash.getRange(2, 1, kept.length, 3).setValues(kept);
        }
        if (alast - 1 > kept.length) ash.deleteRows(2 + kept.length, alast - 1 - kept.length);
      }
    }
    return { dryRun, from, to, rows, approvalsKept, approvalsDeleted };
  } finally {
    lock.releaseLock();
  }
}

function bfWithLock_(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return { status:'busy' }; }
  try { return fn(); } finally { lock.releaseLock(); }
}

// ─────────────────────────────────────────────────────────────
//  관리자 로그인 / 세션 토큰
// ─────────────────────────────────────────────────────────────
function bfSafeEq_(a, b) {                                   // 길이·위치와 무관하게 끝까지 비교
  a = String(a); b = String(b);
  let diff = a.length ^ b.length;
  const n = Math.max(a.length, b.length);
  for (let i = 0; i < n; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function bfLogin_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get('bf_fail') || '0', 10);
  if (fails >= CFG.MAX_FAIL) return { status:'locked' };
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
  if (!real) return { status:'error', message:'스크립트 속성 ADMIN_PW가 설정되지 않았습니다' };
  if (bfSafeEq_(String(pw || ''), real)) {
    cache.remove('bf_fail');
    const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    cache.put('bf_tok_' + token, '1', CFG.TOKEN_TTL);
    return { status:'ok', token: token };
  }
  cache.put('bf_fail', String(fails + 1), CFG.LOCK_SEC);
  return { status:'fail' };
}

function bfTokenOk_(token) {
  token = String(token || '');
  if (token.length < 32) return false;
  const cache = CacheService.getScriptCache();
  if (!cache.get('bf_tok_' + token)) return false;
  cache.put('bf_tok_' + token, '1', CFG.TOKEN_TTL);          // 사용 중이면 만료 시간 연장
  return true;
}

// 로그인 잠금이 걸렸을 때 편집기에서 직접 실행해 해제
function bfClearLoginLock() { CacheService.getScriptCache().remove('bf_fail'); }

// ─────────────────────────────────────────────────────────────
//  승인현황 시트 (여러 기기에서 승인 상태 공유)
// ─────────────────────────────────────────────────────────────
function bfSheet_() {
  let sh = SS.getSheetByName('승인현황');
  if (!sh) {
    sh = SS.insertSheet('승인현황');
    sh.getRange(1, 1, 1, 3).setValues([['날짜', '학번', '승인시각']]).setFontWeight('bold').setBackground('#e0e7ff');
    sh.getRange(1, 1, sh.getMaxRows(), 2).setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}

function bfNorm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

function bfReadRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues().map(function (r) { return [bfNorm_(r[0]), bfNorm_(r[1])]; });
}

function bfGetApprovalKeys_() {
  return bfReadRows_(bfSheet_()).filter(function (r) { return r[0] && r[1]; }).map(function (r) { return r[0] + '_' + r[1]; });
}

// on=true: 승인 추가(이미 있으면 무시) / on=false: 승인 취소
function bfSetApprovals_(date, sids, on) {
  date = bfNorm_(date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { status:'error', message:'bad date' };
  sids = sids.map(bfNorm_).filter(function (s) { return s && s.length <= 20; });
  if (!sids.length) return { status:'ok', changed:0 };
  const sh = bfSheet_();
  const rows = bfReadRows_(sh);
  const have = {};
  rows.forEach(function (r, i) { if (r[0] === date) have[r[1]] = i + 2; });   // 학번 → 시트 행번호

  if (on) {
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const add = [];
    sids.forEach(function (s) { if (!have[s]) { add.push([date, s, now]); have[s] = -1; } });
    if (add.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, add.length, 2).setNumberFormat('@');
      sh.getRange(start, 1, add.length, 3).setValues(add);
    }
    return { status:'ok', changed: add.length };
  }
  const del = [];
  sids.forEach(function (s) { if (have[s] > 0) del.push(have[s]); });
  del.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });   // 아래 행부터 삭제
  return { status:'ok', changed: del.length };
}

function bfClearDate_(date) {
  date = bfNorm_(date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { status:'error', message:'bad date' };
  const sh = bfSheet_();
  const rows = bfReadRows_(sh);
  let n = 0;
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === date) { sh.deleteRow(i + 2); n++; }
  }
  return { status:'ok', changed: n };
}

// ─────────────────────────────────────────────────────────────
//  keepAlive — Cold Start 방지 (5분마다 자동 실행)
// ─────────────────────────────────────────────────────────────
function keepAlive() {
  SpreadsheetApp.getActiveSpreadsheet().getName();
}

// ─────────────────────────────────────────────────────────────
//  설치 — 처음 한 번 실행 (이미 있는 시트·값은 건드리지 않음)
// ─────────────────────────────────────────────────────────────
function setupSheets() {
  const head = (sh, cols, color) => {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground(color || '#1d4ed8').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  };
  // 신청목록
  let sh = SS.getSheetByName('신청목록');
  if (!sh) { sh = SS.insertSheet('신청목록'); head(sh, APPLY_HEADERS); sh.getRange(1, 1, sh.getMaxRows(), APPLY_HEADERS.length).setNumberFormat('@'); }
  // 식단
  sh = SS.getSheetByName('식단');
  if (!sh) { sh = SS.insertSheet('식단'); head(sh, ['date', 'menu', 'qty']); sh.getRange(1, 1, sh.getMaxRows(), 2).setNumberFormat('@'); }
  // 설정 — 없는 키만 추가
  sh = SS.getSheetByName('설정');
  if (!sh) { sh = SS.insertSheet('설정'); head(sh, ['키', '값', '설명'], '#0f766e'); }
  const have = getSettingsMap();
  const add = SETTING_ROWS.filter(r => have[r[0]] === undefined);
  if (add.length) {
    const start = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(start, 1, add.length, 3).setNumberFormat('@').setValues(add);
    sh.getRange(start, 2, add.length, 1).setBackground('#fff9c4');
  }
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 420);
  // 학과목록 — 비어 있으면 예시 3개
  sh = SS.getSheetByName('학과목록');
  if (!sh) {
    sh = SS.insertSheet('학과목록'); head(sh, ['학과명'], '#0f766e');
    sh.getRange(2, 1, 3, 1).setNumberFormat('@').setValues([['(예시) 컴퓨터공학과'], ['(예시) 경영학과'], ['(예시) 간호학과']]);
    sh.setColumnWidth(1, 260);
  }
  // 승인현황
  bfSheet_();
  // 기본 시트(시트1) 정리 — 비어 있을 때만
  const s1 = SS.getSheetByName('시트1') || SS.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && SS.getSheets().length > 1) SS.deleteSheet(s1);
  try { SpreadsheetApp.getUi().alert('설치가 끝났습니다.\n「설정」 시트의 노란 칸과 「학과목록」 시트를 학교에 맞게 고쳐 주세요.\n관리자 비밀번호는 위쪽 「🍱 천원의 아침밥」 메뉴에서 정합니다.'); } catch (e) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🍱 천원의 아침밥')
    .addItem('① 시트 설치 (처음 한 번)', 'setupSheets')
    .addItem('🔑 관리자 비밀번호 설정', 'setAdminPassword')
    .addItem('🔓 로그인 잠금 해제', 'bfClearLoginLock')
    .addToUi();
}

function setAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('관리자 비밀번호 설정', '관리자 화면에서 쓸 비밀번호를 입력하세요 (8자 이상).\n코드나 화면 파일에는 저장되지 않고 스크립트 속성에만 저장됩니다.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = String(r.getResponseText() || '');
  if (pw.length < 8) { ui.alert('8자 이상으로 입력해주세요.'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PW', pw);
  ui.alert('저장되었습니다.');
}
