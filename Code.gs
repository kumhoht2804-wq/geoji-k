/**
 * 거지키우기_k 온라인 서버 (Google Apps Script)
 * 제작자 KSW
 *
 * 이 코드를 구글 시트의 [확장 프로그램 → Apps Script]에 그대로 붙여넣고
 * [배포 → 새 배포 → 웹 앱]으로 배포하세요.
 * 시트 탭(users / saves / board)은 첫 요청 때 자동으로 만들어져요.
 */

const SHEETS = {
  users: ['nick', 'hash', 'salt', 'token', 'created'],
  saves: ['nick', 'data', 'updated'],
  board: ['nick', 'total', 'stage', 'route', 'rebirth', 'updated', 'chBest', 'endings'],
};

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(SHEETS[name]);
    s.setFrozenRows(1);
  } else if (s.getLastColumn() < SHEETS[name].length) {
    s.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);  // 새 칸 머리글 자동 추가
  }
  return s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || '';
  if (a === 'board') return out_(board_());
  if (a === 'ping') return out_({ ok: true, msg: '거지키우기 서버 작동 중 👍' });
  if (a === 'game') {
    // 실행기(PC)가 최신 게임을 받아가는 곳
    const html = HtmlService.createHtmlOutputFromFile('game').getContent();
    const ver = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, html, Utilities.Charset.UTF_8)).slice(0, 10);
    return out_({ ok: true, ver: ver, html: html });
  }
  // 그 외(주소만 열었을 때) → 게임 화면 (핸드폰에서 이 주소로 바로 플레이)
  return HtmlService.createHtmlOutputFromFile('game')
    .setTitle('거지키우기_k')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: '잘못된 요청' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (b.action) {
      case 'register': return out_(register_(b));
      case 'login':    return out_(login_(b));
      case 'save':     return out_(save_(b));
      case 'load':     return out_(load_(b));
      case 'board':    return out_(board_());
      case 'ping':     return out_({ ok: true });
    }
    return out_({ ok: false, err: '알 수 없는 요청' });
  } catch (x) {
    return out_({ ok: false, err: String(x) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */
function hash_(pw, salt) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}
function findRow_(s, nick) {
  const last = s.getLastRow();
  if (last < 2) return -1;
  const v = s.getRange(2, 1, last - 1, 1).getValues();
  const k = String(nick).toLowerCase();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).toLowerCase() === k) return i + 2;
  return -1;
}
function cleanNick_(n) {
  n = String(n || '').trim();
  if (!/^[가-힣a-zA-Z0-9_]{2,12}$/.test(n)) throw '닉네임은 한글/영문/숫자 2~12자로 해주세요';
  return n;
}
function auth_(b) {
  const s = sheet_('users');
  const r = findRow_(s, b.nick || '');
  if (r < 0) throw '로그인이 필요해요';
  const row = s.getRange(r, 1, 1, 5).getValues()[0];
  if (!b.token || row[3] !== b.token) throw '로그인이 만료됐어요. 다시 로그인해 주세요';
  return row[0];
}
function upsert_(s, nick, values) {
  const r = findRow_(s, nick);
  if (r < 0) s.appendRow([nick].concat(values));
  else s.getRange(r, 2, 1, values.length).setValues([values]);
}

/* ---------- actions ---------- */
function register_(b) {
  const nick = cleanNick_(b.nick);
  const pw = String(b.pw || '');
  if (pw.length < 4) throw '비밀번호는 4자 이상으로 해주세요';
  const s = sheet_('users');
  if (findRow_(s, nick) > 0) throw '이미 있는 닉네임이에요';
  const salt = Utilities.getUuid(), token = Utilities.getUuid();
  s.appendRow([nick, hash_(pw, salt), salt, token, new Date()]);
  return { ok: true, nick: nick, token: token };
}

function login_(b) {
  const s = sheet_('users');
  const r = findRow_(s, String(b.nick || '').trim());
  if (r < 0) throw '없는 닉네임이에요';
  const row = s.getRange(r, 1, 1, 5).getValues()[0];
  if (hash_(String(b.pw || ''), row[2]) !== row[1]) throw '비밀번호가 틀렸어요';
  const token = Utilities.getUuid();
  s.getRange(r, 4).setValue(token);
  return { ok: true, nick: row[0], token: token };
}

function save_(b) {
  const nick = auth_(b);
  const data = String(b.data || '');
  if (data.length > 45000) throw '저장 데이터가 너무 커요';
  upsert_(sheet_('saves'), nick, [data, new Date()]);
  const m = b.meta || {};
  const total = Math.max(0, Number(m.total) || 0);
  upsert_(sheet_('board'), nick, [total, String(m.stage || '').slice(0, 20), String(m.route || '').slice(0, 10), Number(m.rebirth) || 0, new Date(),
    Math.max(0, Math.min(999, Number(m.chBest) || 0)), Math.max(0, Math.min(10, Number(m.endings) || 0))]);
  return { ok: true, at: Date.now() };
}

function load_(b) {
  const nick = auth_(b);
  const s = sheet_('saves');
  const r = findRow_(s, nick);
  if (r < 0) return { ok: true, data: null };
  const row = s.getRange(r, 1, 1, 3).getValues()[0];
  return { ok: true, data: row[1], updated: row[2] };
}

function board_() {
  const s = sheet_('board');
  const last = s.getLastRow();
  if (last < 2) return { ok: true, list: [] };
  const cols = Math.max(6, Math.min(8, s.getLastColumn()));
  const v = s.getRange(2, 1, last - 1, cols).getValues();
  const list = v.map(r => ({ nick: r[0], total: Number(r[1]) || 0, stage: r[2], route: r[3], rebirth: Number(r[4]) || 0,
      chBest: Number(r[6]) || 0, endings: Number(r[7]) || 0 }))
    .sort((a, b) => (b.rebirth - a.rebirth) || (b.total - a.total))
    .slice(0, 300);
  return { ok: true, list: list };
}
