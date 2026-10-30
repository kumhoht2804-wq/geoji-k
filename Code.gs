/**
 * 거지키우기_k 온라인 서버 (Google Apps Script)
 * 제작자 KSW
 *
 * 이 코드를 구글 시트의 [확장 프로그램 → Apps Script]에 그대로 붙여넣고
 * [배포 → 새 배포 → 웹 앱]으로 배포하세요.
 * 시트 탭(users / saves / board / events)은 첫 요청 때 자동으로 만들어져요.
 */

const SHEETS = {
  users: ['nick', 'hash', 'salt', 'token', 'created'],
  saves: ['nick', 'data', 'updated'],
  board: ['nick', 'total', 'stage', 'route', 'rebirth', 'updated', 'chBest', 'endings', 'meta'],
  events: ['to', 'from', 'type', 'amount', 'created', 'done', 'note'],
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
      case 'ping':     return out_({ ok: true, pvp: true });
      case 'attack':   return out_(attack_(b));
      case 'inbox':    return out_({ ok: true, events: inbox_(auth_(b)) });
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
    Math.max(0, Math.min(999, Number(m.chBest) || 0)), Math.max(0, Math.min(30, Number(m.endings) || 0)),
    JSON.stringify(m).slice(0, 3000)]);   // 앞으로 추가되는 랭킹 정보는 여기 통째로 저장 → 서버 재배포 불필요
  return { ok: true, at: Date.now(), events: inbox_(nick) };
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
  const cols = Math.max(6, Math.min(9, s.getLastColumn()));
  const v = s.getRange(2, 1, last - 1, cols).getValues();
  const list = v.map(r => {
      let extra = {};
      try { extra = r[8] ? JSON.parse(r[8]) : {}; } catch (e) {}
      return Object.assign(extra, { nick: r[0], total: Number(r[1]) || 0, stage: r[2], route: r[3], rebirth: Number(r[4]) || 0,
        chBest: Number(r[6]) || 0, endings: Number(r[7]) || 0 });
    })
    .sort((a, b) => (b.rebirth - a.rebirth) || (b.total - a.total))
    .slice(0, 300);
  return { ok: true, list: list };
}

/* ---------- PvP (뜯기 / 고소) ---------- */
const PVP = {
  rob: { cd: 10 * 60 * 1000, pct: 0.08 },   // 뜯기: 상대 현금 8%, 쿨타임 10분
  sue: { cd: 15 * 60 * 1000 },              // 고소: 상대 3분간 수입 -30%, 쿨타임 15분
};
function score_(rebirth, total) { return (Number(rebirth) || 0) * 1e18 + (Number(total) || 0); }
function boardRow_(nick) {
  const s = sheet_('board'), r = findRow_(s, nick);
  if (r < 0) return null;
  const v = s.getRange(r, 1, 1, 5).getValues()[0];
  return { nick: v[0], score: score_(v[4], v[1]) };
}
function saveOf_(nick) {
  const s = sheet_('saves'), r = findRow_(s, nick);
  if (r < 0) return null;
  try { return JSON.parse(s.getRange(r, 2).getValue()); } catch (e) { return null; }
}
function attack_(b) {
  const me = auth_(b), type = b.type, P = PVP[type];
  if (!P) throw '알 수 없는 공격';
  const target = String(b.target || '');
  if (target.toLowerCase() === me.toLowerCase()) throw '자기 자신은 안 돼요';
  const T = boardRow_(target), M = boardRow_(me);
  if (!T) throw '없는 사람이에요';
  const ev = sheet_('events'), now = Date.now(), last = ev.getLastRow();
  const rows = last >= 2 ? ev.getRange(2, 1, last - 1, 7).getValues() : [];
  for (const r of rows) {
    const t = new Date(r[4]).getTime();
    if (String(r[1]).toLowerCase() === me.toLowerCase() && r[2] === type && now - t < P.cd)
      throw `쿨타임 중이에요 (${Math.ceil((P.cd - (now - t)) / 60000)}분 남음)`;
  }
  const revenge = rows.some(r => String(r[0]).toLowerCase() === me.toLowerCase() && String(r[1]).toLowerCase() === target.toLowerCase()
    && now - new Date(r[4]).getTime() < 30 * 60 * 1000);
  // v1.7: 순위 제한 없음 (아래 순위 공격은 게임 쪽에서 비용 2배·성공률 -20% 페널티)
  const vs = saveOf_(target) || {};
  let amount = 0, note = '';
  if (type === 'rob') {
    if ((Number(vs.dogEnd) || 0) > now && Math.random() < 0.5) note = 'blocked';
    else amount = Math.floor(Math.max(0, Number(vs.money) || 0) * Math.min(0.18, Math.max(P.pct, Number(b.pct) || P.pct)));
  }
  if (type === 'sue' && (Number(vs.lawyerEnd) || 0) > now) note = 'lawyer';
  ev.appendRow([T.nick, me, type, amount, new Date(), false, note]);
  return { ok: true, type: type, amount: amount, note: note, target: T.nick, revenge: revenge };
}
function inbox_(nick) {
  const ev = sheet_('events'), last = ev.getLastRow();
  if (last < 2) return [];
  const rows = ev.getRange(2, 1, last - 1, 7).getValues(), out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]).toLowerCase() === nick.toLowerCase() && r[5] !== true) {
      out.push({ from: r[1], type: r[2], amount: Number(r[3]) || 0, note: r[6], at: new Date(r[4]).getTime() });
      ev.getRange(i + 2, 6).setValue(true);
    }
  }
  return out;
}
