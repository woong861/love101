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
// ※ SMS는 이 프로젝트의 기존 sendSolapiLMS_() 함수(스크립트 속성 인증)를 재사용합니다.
// ================================================================

const LOBBY_ADMIN_KEY = 'love101-admin';   // make.html의 관리 키와 동일해야 함 (원하면 바꾸고 make.html도 같이)
const LOBBY_SHEET_LOBBIES = '로비_분반';
const LOBBY_SHEET_JOINS = '로비_신청';

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

// ---------------- 전화번호 헬퍼 ----------------
// 시트가 숫자로 저장해 앞 0이 사라진 경우 복원 (10~11자리 한국 휴대폰 가정)
function lobbyFixPhone_(v) {
  var p = String(v == null ? '' : v).replace(/[^0-9]/g, '');
  if (p && p.charAt(0) !== '0') p = '0' + p;
  return p;
}
function lobbyNormPhone_(v) {
  return lobbyFixPhone_(v);
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
          phone: lobbyFixPhone_(j[6]), kakao: j[7], insta: j[8], payment: j[9]
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
      if (lobbyNormPhone_(j[6]) === lobbyNormPhone_(phone)) return { error: 'dup' };
      if (String(j[3]) === '남자') m++; else if (String(j[3]) === '여자') f++;
    }
    var isMale = d.gender === '남자';
    if (isMale && m >= capM) return { error: 'full' };
    if (!isMale && f >= capF) return { error: 'full' };

    jsh.appendRow([
      new Date(), String(d.lobbyId), d.name, d.gender, d.school || '', d.age || '',
      "'" + phone, d.kakao || '', d.insta || '', d.payment || '', 'active'
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

// ---------------- 운영진: 확정 + SMS 발송 (링크 / 대표자 교환 병행) ----------------
// mode: 'link' = 오픈채팅 링크 전원 발송 (기본) / 'rep' = 각 팀 대표끼리 카톡 ID 교환
function lobbyConfirm_(d) {
  if (!d.lobbyId) return { error: 'missing lobbyId' };
  var mode = d.mode === 'rep' ? 'rep' : 'link';
  if (mode === 'link' && !d.link) return { error: 'missing link' };
  var data = lobbyList_(true);
  var lobby = null;
  data.lobbies.forEach(function (l) { if (l.id === String(d.lobbyId)) lobby = l; });
  if (!lobby) return { error: 'notfound' };
  if (!lobby.members.length) return { error: 'no members' };

  var sent = 0, failed = [];
  var send = function (mem, text) {
    if (lobbySms_(mem.phone, text)) sent++;
    else failed.push(mem.name + '(' + mem.phone + ')');
  };

  var linkValue, summary, repInfo = null;

  if (mode === 'link') {
    var text = '[연애학개론] "' + lobby.title + '" 미팅이 확정됐어요! 🎉\n\n' +
      '아래 오픈채팅방에 입장해주세요:\n' + d.link + '\n\n' +
      '즐거운 만남 되세요 ♡\n- 연애학개론 교무처\n인스타 @_posting2';
    lobby.members.forEach(function (mem) { send(mem, text); });
    linkValue = d.link;
    summary = text;
  } else {
    // 대표자: 성별별 첫 신청자
    var repM = null, repF = null;
    lobby.members.forEach(function (mem) {
      if (!repM && mem.gender === '남자') repM = mem;
      if (!repF && mem.gender === '여자') repF = mem;
    });
    if (!repM || !repF) return { error: 'rep mode needs both genders' };

    lobby.members.forEach(function (mem) {
      var text;
      if (mem === repM) {
        text = '[연애학개론] "' + lobby.title + '" 미팅이 확정됐어요! 🎉\n\n' +
          mem.name + '님이 남자팀 대표예요.\n' +
          '상대팀 대표 카톡 ID: ' + (repF.kakao || repF.phone) + '\n\n' +
          '먼저 친추하고 인사한 뒤, 6명 단체 카톡방을 만들어주세요 ♡\n' +
          '- 연애학개론 교무처\n문의 인스타 @_posting2';
      } else if (mem === repF) {
        text = '[연애학개론] "' + lobby.title + '" 미팅이 확정됐어요! 🎉\n\n' +
          mem.name + '님이 여자팀 대표예요.\n' +
          '상대팀 대표 카톡 ID: ' + (repM.kakao || repM.phone) + '\n' +
          '곧 연락이 올 거예요. 단체방이 만들어지면 팀원들을 초대해주세요 ♡\n' +
          '- 연애학개론 교무처\n문의 인스타 @_posting2';
      } else {
        text = '[연애학개론] "' + lobby.title + '" 미팅이 확정됐어요! 🎉\n\n' +
          '각 팀 대표끼리 연결 중이에요.\n곧 단체 카톡방에 초대될 예정입니다 ♡\n' +
          '- 연애학개론 교무처\n문의 인스타 @_posting2';
      }
      send(mem, text);
    });
    repInfo = { m: repM.name, f: repF.name };
    linkValue = '(대표자 교환) 남 ' + repM.name + ' ↔ 여 ' + repF.name;
    summary = linkValue;
  }

  var lsh = lobbySheet_(LOBBY_SHEET_LOBBIES);
  lsh.getRange(lobby.row, 7).setValue('confirmed');
  lsh.getRange(lobby.row, 8).setValue(linkValue);
  return { ok: true, mode: mode, sent: sent, failed: failed, total: lobby.members.length, smsText: summary, reps: repInfo };
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
// 이 프로젝트의 기존 sendSolapiLMS_() (스크립트 속성 인증)를 그대로 재사용
function lobbySms_(to, text) {
  try {
    sendSolapiLMS_(to, text);
    return true;
  } catch (e) {
    return false;
  }
}
