/**
 * 다육이 소원 공작소 — 공유 재고 서버 (Google Apps Script)
 *
 * 이 스크립트를 구글 시트의 [확장 프로그램 → Apps Script]에 붙여 넣고
 * [배포 → 새 배포 → 웹 앱]으로 배포하면, 모든 태블릿이 이 시트 하나로 재고를 함께 써요.
 *
 * 시트 탭(없으면 자동으로 만들어요)
 *   - 재고 현황 : 종류별 준비 수량, 조정, 나눠준 수, 남은 수 (자동 갱신)
 *   - 기록      : 추천·조정·취소가 한 줄씩 쌓여요
 *   - 설정      : 회차(재고 초기화할 때마다 1씩 올라가요)
 */

// ▼ 앱(index.html)의 APP_KEY와 똑같이 맞춰 주세요.
const APP_KEY = 'dayuk-2026';
// ▼ 종류별 처음 준비 수량
const TOTAL = 40;

const KINDS = [
  // [앱 내부 이름, 품종, 바구니 번호, 아이용 별명]
  ['star',  '수',             1, '반짝 보석이'],
  ['bean',  '애플캔디',        2, '새콤 사과사탕'],
  ['rose',  '어메이징그레이스', 3, '장미 요정'],
  ['brave', '엘린',            4, '뾰족 용감이'],
  ['bunny', '미니염좌',         5, '동글 나무'],
];
const KIND = {}; KINDS.forEach(k => KIND[k[0]] = { real: k[1], basket: k[2], name: k[3] });

const LOG = '기록', CFG = '설정', SUM = '재고 현황';
const LOG_HEAD = ['ID', '시각', '회차', '구분', '바구니', '품종', '수량 변화', '종류코드', '요청ID'];
const TZ = 'Asia/Seoul';

/* ---------- 입구 ---------- */
function doGet() {
  return ContentService.createTextOutput('다육이 재고 서버가 켜져 있어요. (OK)');
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); }
  catch (err) { return out_({ error: 'bad_request' }); }
  try { return out_(handle_(body)); }
  catch (err) { return out_({ error: 'server_error', message: String(err && err.message || err) }); }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 처리 ---------- */
function handle_(b) {
  if (b.key !== APP_KEY) return { error: 'bad_key' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 읽기는 잠그지 않아요(빠르게). 시트가 아직 없을 때만 잠그고 만들어요.
  if (b.action === 'state') {
    if (!ready_(ss)) withLock_(() => setup_(ss));
    return { state: state_(ss) };
  }

  // 쓰기는 한 번에 하나씩만 (여러 태블릿이 동시에 눌러도 마지막 1개를 두 번 주지 않아요)
  return withLock_(() => {
    setup_(ss);
    const round = getRound_(ss);
    const res = {};
    switch (b.action) {
      case 'claim': Object.assign(res, claim_(ss, round, b)); break;
      case 'adjust': {
        const k = b.kind, d = Math.trunc(Number(b.delta) || 0);
        if (!KIND[k]) return { error: 'bad_kind' };
        if (d) addRow_(ss, round, '조정', k, d, '');
        break;
      }
      case 'setLeft': {
        const k = b.kind, v = Math.max(0, Math.trunc(Number(b.value)));
        if (!KIND[k] || isNaN(v)) return { error: 'bad_value' };
        const d = v - left_(rows_(ss, round), k);
        if (d) addRow_(ss, round, '조정', k, d, '');
        break;
      }
      case 'undo': undo_(ss, round, String(b.id || '')); break;
      case 'reset': setRound_(ss, round + 1); break;
      default: return { error: 'bad_action' };
    }
    writeSummary_(ss);
    res.state = state_(ss);
    return res;
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { error: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}

/* 추천 1개 확정: 서버에서 다시 세어 남아 있는 종류만 줘요 */
function claim_(ss, round, b) {
  const all = readLog_(ss);
  // 같은 요청이 인터넷 문제로 다시 오면, 화분을 또 빼지 않고 처음 결과를 돌려줘요.
  if (b.reqId) {
    const done = all.find(r => r.req === b.reqId && (r.type === '추천' || r.type === '추천(취소됨)'));
    if (done) return { kind: done.type === '추천' ? done.kind : null, id: done.id, repeated: true };
  }
  const rows = all.filter(r => r.round === round && (r.type === '추천' || r.type === '조정'));
  const order = (Array.isArray(b.order) ? b.order : []).filter(k => KIND[k]);
  KINDS.forEach(k => { if (order.indexOf(k[0]) < 0) order.push(k[0]); }); // 혹시 빠진 종류도 끝에 넣어요
  for (const k of order) {
    if (left_(rows, k) > 0) {
      const id = addRow_(ss, round, '추천', k, -1, String(b.reqId || ''));
      return { kind: k, id: id };
    }
  }
  return { kind: null };
}

function undo_(ss, round, id) {
  if (!id) return;
  const sh = ss.getSheetByName(LOG);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === id && vals[i][3] === '추천' && Number(vals[i][2]) === round) {
      sh.getRange(i + 1, 4).setValue('추천(취소됨)');
      sh.getRange(i + 1, 7).setValue(0);
      return;
    }
  }
}

/* ---------- 시트 읽기/쓰기 ---------- */
function ready_(ss) { return !!(ss.getSheetByName(LOG) && ss.getSheetByName(CFG) && ss.getSheetByName(SUM)); }

function setup_(ss) {
  if (!ss.getSheetByName(LOG)) {
    const sh = ss.insertSheet(LOG);
    sh.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName(CFG)) {
    const sh = ss.insertSheet(CFG);
    sh.getRange(1, 1, 2, 2).setValues([['회차', 1], ['안내', '재고 초기화를 누르면 회차가 1씩 올라가요. 지난 회차 기록은 개수에 들어가지 않아요.']]);
  }
  if (!ss.getSheetByName(SUM)) {
    ss.insertSheet(SUM, 0);
    writeSummary_(ss);
  }
}

function getRound_(ss) { return Number(ss.getSheetByName(CFG).getRange(1, 2).getValue()) || 1; }
function setRound_(ss, r) { ss.getSheetByName(CFG).getRange(1, 2).setValue(r); }

function readLog_(ss) {
  const vals = ss.getSheetByName(LOG).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const v = vals[i];
    if (!v[0]) continue;
    out.push({ id: String(v[0]), at: v[1], round: Number(v[2]), type: String(v[3]), delta: Number(v[6]) || 0, kind: String(v[7]), req: String(v[8] || '') });
  }
  return out;
}

function rows_(ss, round) {
  return readLog_(ss).filter(r => r.round === round && (r.type === '추천' || r.type === '조정'));
}

function left_(rows, k) {
  let supply = TOTAL, used = 0;
  rows.forEach(r => {
    if (r.kind !== k) return;
    if (r.type === '조정') supply += r.delta;
    else if (r.type === '추천') used += 1;
  });
  return supply - used;
}

function addRow_(ss, round, type, k, delta, req) {
  const id = 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const K = KIND[k];
  ss.getSheetByName(LOG).appendRow([id, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), round, type, K.basket + '번', K.real, delta, k, req]);
  return id;
}

function state_(ss) {
  const round = getRound_(ss);
  const picks = rows_(ss, round).map(r => r.type === '조정' ? { id: r.id, kind: r.kind, add: r.delta } : { id: r.id, kind: r.kind });
  return { round: round, total: TOTAL, picks: picks };
}

function writeSummary_(ss) {
  const sh = ss.getSheetByName(SUM);
  const round = getRound_(ss);
  const rows = rows_(ss, round);
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const table = [['바구니', '품종', '아이용 별명', '준비 수량', '추가/조정', '나눠준 수', '남은 수', '마지막 업데이트']];
  let T = [0, 0, 0, 0];
  KINDS.forEach((k, i) => {
    const adj = rows.filter(r => r.kind === k[0] && r.type === '조정').reduce((a, r) => a + r.delta, 0);
    const used = rows.filter(r => r.kind === k[0] && r.type === '추천').length;
    const left = Math.max(0, TOTAL + adj - used);
    table.push([k[2] + '번', k[1], k[3], TOTAL, adj, used, left, i === 0 ? now + ' (회차 ' + round + ')' : '']);
    T = [T[0] + TOTAL, T[1] + adj, T[2] + used, T[3] + left];
  });
  table.push(['', '합계', '', T[0], T[1], T[2], T[3], '']);
  sh.getRange(1, 1, table.length, table[0].length).setValues(table);
  sh.getRange(1, 1, 1, table[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}
