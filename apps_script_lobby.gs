// ================================================================
// 연애학개론 미팅 로비 (베타) — Apps Script 백엔드
// ================================================================
// 【설치 방법】
// 1) 연애학개론 시트의 Apps Script 프로젝트에서 새 파일 추가 (예: Lobby.gs)
//    → 이 파일 내용 전체를 붙여넣기
// 2) 기존 doGet(e) 함수의 "맨 첫 줄"에 아래 두 줄 추가:
//      var lobbyRes = lobbyRouter_(e);
//      if (lobbyRes) return lobbyRes;
// 3) 배포 → 배포 관리 → 연필(수정) → 버전 "새 버전" → 배포 (URL 유지)
//
// ※ 시트 탭(로비_분반/로비_신청)은 처음 호출될 때 자동 생성됩니다.
// ※ SMS는 프로젝트에 이미 있는 SOLAPI_API_KEY/SECRET/SENDER 상수를 자동으로
//    사용합니다. 없으면 아래 LOBBY_SOLAPI_* 에 직접 채워주세요.
// ================================================================

const LOBBY_ADMIN_KEY = 'love101-admin';   // make.html의 관리 키와 동일해야 함 (원하면 바꾸고 make.html도 같이)
const LOBBY_SHEET_LOBBIES = '로비_분반';
const LOBBY_SHEET_JOINS = '로비_신청';

// 프로젝트에 SOLAPI 상수가 없을 때만 사용되는 예비 설정
const LOBBY_SOLAPI_KEY = '';
const LOBBY_SOLAPI_SECRET = '';
const LOBBY_SOLAPI_SENDER = '';

// ---------------- 라우터 ----------------
function lobbyRouter_(e) {
  var action = e && e.parameter && e.parameter.action;
  if (!action || String(action).indexOf('lobby') !== 0) return null;
  var out;
  try {
    var d = {};
    if (e.parameter.d) {
      d = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(e.parameter.d)).getDataAsString('UTF-8'));
    }
    var isAdmin = e.parameter.key === LOBBY_ADMIN_KEY;
    switch (action) {
      case 'lobbyList':      out = lobbyList_(false); break;
      case 'lobbyAdminList': out = isAdmin ? lobbyList_(true) : { error: 'auth' }; break;
      case 'lobbyJoin':      out = lobbyJoin_(d); break;
      case 'lobbyCreate':    out = isAdmin ? lobbyCreate_(d) : { error: 'auth' }; break;
      case 'lobbyExclude':   out = isAdmin ? lobbyExclude_(d) : { error: 'auth' }; break;
      case 'lobbyConfirm':   out = isAdmin ? lobbyConfirm_(d) : { error: 'auth' }; break;
      case 'lobbyClose':     out = isAdmin ? lobbyClose_(d) : { error: 'auth' }; break;
      default: out = { error: 'unknown action' };
    }
  } catch (err) {
    out = { error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------- 시트 헬퍼 ----------------
function lobbySheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === LOBBY_SHEET_LOBBIES) {
      sh.appendRow(['ID', '제목', '유형', '설명', '정원남', '정원여', '상태', '링크', '생성시각']);
    } else {
      sh.appendRow(['시각', '분반ID', '이름', '성별', '학교', '나이', '전화', '카톡', '인스타', '입금정보', '상태']);
    }
  }
  return sh;
}

function lobbyRows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

// ---------------- 목록 조회 ----------------
function lobbyList_(withMembers) {
  var lsh = lobbySheet_(LOBBY_SHEET_LOBBIES);
  var jsh = lobbySheet_(LOBBY_SHEET_JOINS);
  var joins = lobbyRows_(jsh);
  var lobbies = lobbyRows_(lsh).map(function (r, i) {
    var id = String(r[0]);
    var members = [];
    var m = 0, f = 0;
    joins.forEach(function (j, ji) {
      if (String(j[1]) !== id) return;
      if (String(j[10]) !== 'active') return;
      if (String(j[3]) === '남자') m++; else if (String(j[3]) === '여자') f++;
      if (withMembers) {
        members.push({
          row: ji + 2, name: j[2], gender: j[3], school: j[4], age: j[5],
          phone: String(j[6]), kakao: j[7], insta: j[8], payment: j[9]
        });
      }
    });
    var lobby = {
      id: id, row: i + 2, title: r[1], type: r[2], desc: r[3],
      capM: Number(r[4]) || 0, capF: Number(r[5]) || 0,
      m: m, f: f, status: String(r[6] || 'open')
    };
    if (withMembers) { lobby.members = members; lobby.link = r[7]; }
    return lobby;
  });
  return { lobbies: lobbies };
}

// ---------------- 참여 신청 (동시성 잠금) ----------------
function lobbyJoin_(d) {
  if (!d.lobbyId || !d.name || !d.gender || !d.phone) return { error: 'missing' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var lsh = lobbySheet_(LOBBY_SHEET_LOBBIES);
    var rows = lobbyRows_(lsh);
    var idx = -1, lobby = null;
    rows.forEach(function (r, i) { if (String(r[0]) === String(d.lobbyId)) { idx = i; lobby = r; } });
    if (!lobby) return { error: 'notfound' };
    var status = String(lobby[6] || 'open');
    if (status !== 'open') return { error: 'closed' };

    var capM = Number(lobby[4]) || 0, capF = Number(lobby[5]) || 0;
    var jsh = lobbySheet_(LOBBY_SHEET_JOINS);
    var joins = lobbyRows_(jsh);
    var m = 0, f = 0, phone = String(d.phone).replace(/[^0-9]/g, '');
    for (var i = 0; i < joins.length; i++) {
      var j = joins[i];
      if (String(j[1]) !== String(d.lobbyId) || String(j[10]) !== 'active') continue;
      if (String(j[6]).replace(/[^0-9]/g, '') === phone) return { error: 'dup' };
      if (String(j[3]) === '남자') m++; else if (String(j[3]) === '여자') f++;
    }
    var isMale = d.gender === '남자';
    if (isMale && m >= capM) return { error: 'full' };
    if (!isMale && f >= capF) return { error: 'full' };

    jsh.appendRow([
      new Date(), String(d.lobbyId), d.name, d.gender, d.school || '', d.age || '',
      phone, d.kakao || '', d.insta || '', d.payment || '', 'active'
    ]);
    if (isMale) m++; else f++;
    if (m >= capM && f >= capF) lsh.getRange(idx + 2, 7).setValue('full');
    return { ok: true, m: m, f: f, capM: capM, capF: capF };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- 운영진: 분반 개설 ----------------
function lobbyCreate_(d) {
  if (!d.title) return { error: 'missing title' };
  var lsh = lobbySheet_(LOBBY_SHEET_LOBBIES);
  var id = 'L' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMddHHmmss');
  var capM = Number(d.capM) || 3, capF = Number(d.capF) || 3;
  lsh.appendRow([id, d.title, d.type || (capM + ':' + capF), d.desc || '', capM, capF, 'open', '', new Date()]);
  return { ok: true, id: id };
}

// ---------------- 운영진: 신청자 제외 ----------------
function lobbyExclude_(d) {
  if (!d.joinRow) return { error: 'missing joinRow' };
  var jsh = lobbySheet_(LOBBY_SHEET_JOINS);
  jsh.getRange(Number(d.joinRow), 11).setValue('excluded');
  var lobbyId = String(jsh.getRange(Number(d.joinRow), 2).getValue());
  // 자리가 다시 생기므로 full이었으면 open으로 되돌림
  var lsh = lobbySheet_(LOBBY_SHEET_LOBBIES);
  lobbyRows_(lsh).forEach(function (r, i) {
    if (String(r[0]) === lobbyId && String(r[6]) === 'full') {
      lsh.getRange(i + 2, 7).setValue('open');
    }
  });
  return { ok: true };
}

// ---------------- 운영진: 확정 + 오픈채팅 링크 SMS 발송 ----------------
function lobbyConfirm_(d) {
  if (!d.lobbyId || !d.link) return { error: 'missing lobbyId/link' };
  var data = lobbyList_(true);
  var lobby = null;
  data.lobbies.forEach(function (l) { if (l.id === String(d.lobbyId)) lobby = l; });
  if (!lobby) return { error: 'notfound' };
  if (!lobby.members.length) return { error: 'no members' };

  var text = '[연애학개론] "' + lobby.title + '" 미팅이 확정됐어요! 🎉\n\n' +
    '아래 오픈채팅방에 입장해주세요:\n' + d.link + '\n\n' +
    '즐거운 만남 되세요 ♡\n- 연애학개론 교무처\n인스타 @_posting2';

  var sent = 0, failed = [];
  lobby.members.forEach(function (mem) {
    if (lobbySms_(mem.phone, text)) sent++;
    else failed.push(mem.name + '(' + mem.phone + ')');
  });

  var lsh = lobbySheet_(LOBBY_SHEET_LOBBIES);
  lsh.getRange(lobby.row, 7).setValue('confirmed');
  lsh.getRange(lobby.row, 8).setValue(d.link);
  return { ok: true, sent: sent, failed: failed, total: lobby.members.length, smsText: text };
}

// ---------------- 운영진: 분반 마감 ----------------
function lobbyClose_(d) {
  if (!d.lobbyId) return { error: 'missing lobbyId' };
  var lsh = lobbySheet_(LOBBY_SHEET_LOBBIES);
  var done = false;
  lobbyRows_(lsh).forEach(function (r, i) {
    if (String(r[0]) === String(d.lobbyId)) { lsh.getRange(i + 2, 7).setValue('closed'); done = true; }
  });
  return done ? { ok: true } : { error: 'notfound' };
}

// ---------------- SMS (솔라피) ----------------
function lobbyCreds_() {
  try {
    if (typeof SOLAPI_API_KEY !== 'undefined' && SOLAPI_API_KEY) {
      return { key: SOLAPI_API_KEY, secret: SOLAPI_API_SECRET, sender: SOLAPI_SENDER };
    }
  } catch (e) {}
  return { key: LOBBY_SOLAPI_KEY, secret: LOBBY_SOLAPI_SECRET, sender: LOBBY_SOLAPI_SENDER };
}

function lobbySms_(to, text) {
  var c = lobbyCreds_();
  if (!c.key || !c.secret || !c.sender) return false;
  try {
    var date = new Date().toISOString();
    var salt = Utilities.getUuid();
    var sig = Utilities.computeHmacSha256Signature(date + salt, c.secret)
      .map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; })
      .join('');
    var msg = {
      to: String(to).replace(/[^0-9]/g, ''),
      from: String(c.sender).replace(/[^0-9]/g, ''),
      text: text,
      type: text.length > 45 ? 'LMS' : 'SMS'
    };
    if (msg.type === 'LMS') msg.subject = '연애학개론';
    var res = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'HMAC-SHA256 apiKey=' + c.key + ', date=' + date + ', salt=' + salt + ', signature=' + sig },
      payload: JSON.stringify({ message: msg }),
      muteHttpExceptions: true
    });
    return res.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}
