/* ============================================================
   부활절 콘서트 — 참가 신청 / 식사 수령(체크인) 백엔드
   ------------------------------------------------------------
   동작:
   - 참가 신청  : ?action=register&name=..&phone=..&org=..&count=..
   - 티켓 조회  : ?action=lookup&t=CLF-0001
   - 식사 수령  : ?action=checkin&t=CLF-0001
   응답은 JSONP(callback 파라미터)로 돌려줍니다.

   ▣ 배포 방법 (한 번만)
   1) https://sheets.new 로 새 구글 스프레드시트 생성
   2) 메뉴 [확장 프로그램] → [Apps Script]
   3) 기본 Code.gs 내용을 지우고 이 파일 전체를 붙여넣기 → 저장
   4) 우측 상단 [배포] → [새 배포] → 유형 [웹 앱]
      - 실행 계정: 나
      - 액세스 권한: "모든 사용자(Anyone)"
      → [배포], 권한 승인
   5) 표시되는 "웹 앱 URL"(.../exec)을 복사해
      홈페이지의 config.js → APPS_SCRIPT_URL 에 붙여넣기
   ============================================================ */

var SHEET_NAME = '신청자';
var PREFIX = 'CLF-';   // 티켓 번호 접두사 (Christian Leaders Fellowship)

function doGet(e){
  var p = (e && e.parameter) ? e.parameter : {};
  var action = p.action || '';
  var out;
  try {
    if (action === 'register')      out = register(p);
    else if (action === 'lookup')   out = lookup(p.t);
    else if (action === 'checkin')  out = checkin(p.t);
    else                            out = { ok:false, error:'unknown action' };
  } catch (err) {
    out = { ok:false, error: String(err) };
  }
  return reply(out, p.callback);
}

function reply(obj, callback){
  var json = JSON.stringify(obj);
  if (callback){
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh){
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['신청시각','티켓번호','신청유형','성함','연락처','소속','인원','식사수령','수령시각']);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, 3).getValue() !== '신청유형') {
    // 기존 시트(8컬럼)에 신청유형 컬럼을 C열로 삽입
    sh.insertColumnBefore(3);
    sh.getRange(1, 3).setValue('신청유형');
  }
  return sh;
}

function register(p){
  var name  = (p.name  || '').toString().trim();
  var phone = (p.phone || '').toString().trim();
  var org   = (p.org   || '').toString().trim();
  var type  = (p.type  || '').toString().trim() || '개인';
  var count = parseInt(p.count, 10); if (!count || count < 1) count = 1; if (count > 200) count = 200;
  if (!name || !phone) return { ok:false, error:'성함과 연락처는 필수입니다.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = getSheet();
    var seq = sh.getLastRow();                 // 헤더 포함 행 수 = 다음 일련번호
    var ticket = PREFIX + ('0000' + seq).slice(-4);
    sh.appendRow([new Date(), ticket, type, name, phone, org, count, '', '']);
    return { ok:true, ticket:ticket, name:name, count:count, type:type };
  } finally {
    lock.releaseLock();
  }
}

function findRow(t){
  var sh = getSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (String(data[i][1]) === String(t)) return { sh:sh, idx:i, row:i+1, rec:data[i] };
  }
  return null;
}

function lookup(t){
  if (!t) return { ok:false, error:'no ticket' };
  var f = findRow(t);
  if (!f) return { ok:false, error:'not found' };
  var r = f.rec;
  return { ok:true, ticket:t, type:r[2], name:r[3], org:r[5], count:r[6], checkedIn: !!r[7], at: r[8] ? fmt(r[8]) : '' };
}

function checkin(t){
  if (!t) return { ok:false, error:'no ticket' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var f = findRow(t);
    if (!f) return { ok:false, error:'not found' };
    var r = f.rec;
    if (r[7]) return { ok:true, already:true, type:r[2], name:r[3], org:r[5], count:r[6], at: r[8] ? fmt(r[8]) : '' };
    var now = new Date();
    f.sh.getRange(f.row, 8).setValue('수령');
    f.sh.getRange(f.row, 9).setValue(now);
    return { ok:true, already:false, type:r[2], name:r[3], org:r[5], count:r[6], at: fmt(now) };
  } finally {
    lock.releaseLock();
  }
}

function fmt(d){
  try { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'M/d HH:mm'); }
  catch (e){ return String(d); }
}
