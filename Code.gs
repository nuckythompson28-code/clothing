/**
 * SOLTRI 피복 관리 — Google Apps Script 백엔드
 *
 * NAS server.py(포트 8588)와 동일한 API를 구글시트 위에 구현.
 * 프론트(GitHub Pages index.html)는 이 웹앱 URL만 입력하면 그대로 동작한다.
 *
 * ── 설치 ────────────────────────────────────────────
 * 1. 새 Google 스프레드시트 생성 (이름: 피복관리DB 등)
 * 2. 확장 프로그램 → Apps Script → 이 파일 내용 전체 붙여넣기
 * 3. 배포 → 새 배포 → 유형: 웹 앱
 *    - 실행: 나 / 액세스 권한: 모든 사용자(익명)
 * 4. 발급된 /exec URL을 앱 설정 화면에 입력
 *
 * 시트는 첫 호출 때 자동 생성된다 (setup 불필요).
 */

const API_KEY = 'soltri2026';
const PDF_FOLDER_NAME = '피복관리_PDF';

// ─── 시트 스키마 (헤더 = API 필드명) ───────────────────
const SCHEMA = {
  items:        ['id', 'name', 'size', 'category', 'stock', 'totalIn', 'totalOut'],
  emps:         ['empId', 'name', 'dept', 'rank', 'joinDate', 'status', 'retiredDate'],
  history:      ['id', 'type', 'empId', 'name', 'dept', 'rank', 'date', 'time', 'items', 'refId', 'note', 'sigData', 'backfilled'],
  stock_in_log: ['id', 'date', 'time', 'vendor', 'note', 'items', 'totalQty'],
  stock_adj:    ['id', 'type', 'typeName', 'itemId', 'itemName', 'itemSize', 'before', 'after', 'delta', 'note', 'date', 'time'],
  depts:        ['name', 'order'],
  pdf_config:   ['key', 'value'],
};
// 문자열로 강제할 컬럼 (사이즈 "95"→95 같은 자동 숫자화 방지용 역변환)
const STR_COLS = {
  items: ['id', 'name', 'size', 'category'],
  emps: ['empId', 'name', 'dept', 'rank', 'joinDate', 'status', 'retiredDate'],
  history: ['id', 'type', 'empId', 'name', 'dept', 'rank', 'date', 'time', 'items', 'refId', 'note', 'sigData'],
  stock_in_log: ['id', 'date', 'time', 'vendor', 'note', 'items'],
  stock_adj: ['id', 'type', 'typeName', 'itemId', 'itemName', 'itemSize', 'note', 'date', 'time'],
  depts: ['name'],
  pdf_config: ['key', 'value'],
};
const DATE_COLS = ['date', 'joinDate', 'retiredDate'];
const TIME_COLS = ['time'];

// ─── 시트 헬퍼 ────────────────────────────────────────
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    seedDefaults_(name, sh);
  }
  return sh;
}

function seedDefaults_(name, sh) {
  if (name === 'pdf_config') {
    const d = [
      ['문서제목', '피복지급의뢰서'], ['영문회사명', 'SOLTRI Co., Ltd.'],
      ['헤더배경색', '#1a3353'], ['헤더글자색', '#ffffff'], ['테이블헤더색', '#e8eef5'],
      ['결재라인표시', 'Y'], ['결재라인', '직원,부서장,이사'], ['문서번호접두어', 'UNI'],
    ];
    sh.getRange(2, 1, d.length, 2).setValues(d);
  }
  if (name === 'depts') {
    const d = ['성형1', '성형2', '가공1', '가공2', '후가공', '품질', '연구소', '영업', '관리'].map((n, i) => [n, i]);
    sh.getRange(2, 1, d.length, 2).setValues(d);
  }
}

// 셀 값 정규화: Date→문자열, 문자열 컬럼은 String 강제
function norm_(table, col, v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    if (TIME_COLS.indexOf(col) > -1) return Utilities.formatDate(v, tz, 'HH:mm');
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  if ((STR_COLS[table] || []).indexOf(col) > -1) return String(v);
  return v;
}

// 시트 전체 → 객체 배열
function readTable_(table) {
  const sh = sheet_(table);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = SCHEMA[table];
  const vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  return vals.map(row => {
    const o = {};
    cols.forEach((c, i) => { o[c] = norm_(table, c, row[i]); });
    return o;
  }).filter(o => String(o[cols[0]]) !== '');
}

// 객체 → 시트 행 배열
function toRow_(table, obj) {
  return SCHEMA[table].map(c => (obj[c] === undefined || obj[c] === null) ? '' : obj[c]);
}

function appendRow_(table, obj) {
  const sh = sheet_(table);
  sh.appendRow(toRow_(table, obj));
}

// id(첫 컬럼)로 행 번호 찾기 (없으면 -1)
function findRow_(table, id) {
  const sh = sheet_(table);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function updateRow_(table, rowNum, obj) {
  const sh = sheet_(table);
  sh.getRange(rowNum, 1, 1, SCHEMA[table].length).setValues([toRow_(table, obj)]);
}

function getRowObj_(table, rowNum) {
  const sh = sheet_(table);
  const cols = SCHEMA[table];
  const row = sh.getRange(rowNum, 1, 1, cols.length).getValues()[0];
  const o = {};
  cols.forEach((c, i) => { o[c] = norm_(table, c, row[i]); });
  return o;
}

// ─── 유틸 ─────────────────────────────────────────────
function nowDate_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function nowTime_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'); }
function genId_(suffix) { return Date.now().toString(36) + (suffix || ''); }
function parseItems_(s) { try { return JSON.parse(s) || []; } catch (e) { return []; } }

// ─── getAll ───────────────────────────────────────────
function actionGetAll() {
  const items = readTable_('items').map(r => ({
    id: r.id, name: r.name, size: r.size, category: r.category,
    stock: Number(r.stock) || 0, totalIn: Number(r.totalIn) || 0, totalOut: Number(r.totalOut) || 0,
  }));
  const history = readTable_('history').map(r => ({
    id: r.id, type: r.type, empId: r.empId, name: r.name, dept: r.dept, rank: r.rank,
    date: r.date, time: r.time, items: parseItems_(r.items),
    refId: r.refId, note: r.note, sigData: r.sigData, backfilled: !!Number(r.backfilled),
  })).sort((a, b) =>
    (b.date + ' ' + b.time + ' ' + b.id).localeCompare(a.date + ' ' + a.time + ' ' + a.id)
  );
  const pdfConfig = {};
  readTable_('pdf_config').forEach(r => { pdfConfig[r.key] = r.value; });
  const depts = readTable_('depts')
    .sort((a, b) => (Number(a.order) || 0) - (Number(b.order) || 0) || String(a.name).localeCompare(b.name))
    .map(r => r.name);
  const employees = readTable_('emps');
  const adjLog = readTable_('stock_adj').map(r => ({
    id: r.id, type: r.type, typeName: r.typeName, itemId: r.itemId, itemName: r.itemName,
    itemSize: r.itemSize, before: Number(r.before) || 0, after: Number(r.after) || 0,
    delta: Number(r.delta) || 0, note: r.note, date: r.date, time: r.time,
  })).sort((a, b) => (b.date + ' ' + b.time).localeCompare(a.date + ' ' + a.time));
  const stockInLog = readTable_('stock_in_log').map(r => ({
    id: r.id, date: r.date, time: r.time, vendor: r.vendor, note: r.note,
    items: parseItems_(r.items), totalQty: Number(r.totalQty) || 0,
  })).sort((a, b) =>
    (b.date + ' ' + b.time + ' ' + b.id).localeCompare(a.date + ' ' + a.time + ' ' + a.id)
  );
  return { items, history, pdfConfig, depts, employees, adjLog, stockInLog };
}

// ─── 지급 ─────────────────────────────────────────────
function actionDistribute(d) {
  const rid = d.recordId || genId_();
  const items = d.items || [];
  if (!items.length) return { error: 'items empty' };
  if (findRow_('history', rid) > -1) return { ok: true, recordId: rid, duplicate: true };
  for (const it of items) {
    const q = parseInt(it.qty, 10) || 0;
    if (q <= 0) continue;
    const rowNum = findRow_('items', it.id);
    if (rowNum < 0) return { error: '품목 없음: ' + it.id };
    const cur = getRowObj_('items', rowNum);
    cur.stock = (Number(cur.stock) || 0) - q;
    cur.totalOut = (Number(cur.totalOut) || 0) + q;
    updateRow_('items', rowNum, cur);
  }
  appendRow_('history', {
    id: rid, type: 'dist', empId: d.empId || '', name: d.name || '', dept: d.dept || '',
    rank: d.rank || '', date: d.date || nowDate_(), time: d.time || nowTime_(),
    items: JSON.stringify(items), refId: '', note: '', sigData: d.sigData || '', backfilled: 0,
  });
  return { ok: true, recordId: rid };
}

// ─── 반납 ─────────────────────────────────────────────
function actionReturn(d) {
  const rid = d.recordId || genId_('r');
  const items = d.items || [];
  if (!items.length) return { error: 'items empty' };
  if (findRow_('history', rid) > -1) return { ok: true, recordId: rid, duplicate: true };
  for (const it of items) {
    const q = parseInt(it.qty, 10) || 0;
    if (q <= 0) continue;
    const rowNum = findRow_('items', it.id);
    if (rowNum < 0) continue;
    const cur = getRowObj_('items', rowNum);
    cur.stock = (Number(cur.stock) || 0) + q;
    cur.totalOut = Math.max(0, (Number(cur.totalOut) || 0) - q);
    updateRow_('items', rowNum, cur);
  }
  appendRow_('history', {
    id: rid, type: 'return', empId: d.empId || '', name: d.name || '', dept: d.dept || '',
    rank: d.rank || '', date: d.date || nowDate_(), time: d.time || nowTime_(),
    items: JSON.stringify(items), refId: d.refId || '', note: '', sigData: d.sigData || '', backfilled: 0,
  });
  return { ok: true, recordId: rid };
}

// ─── 입고 ─────────────────────────────────────────────
function actionStockIn(d) {
  const rid = d.recordId || genId_('i');
  let items = d.items || [];
  if (!items.length) {
    const q = parseInt(d.qty, 10) || 0;
    if (!d.itemId || q <= 0) return { error: 'invalid' };
    items = [{ id: d.itemId, qty: q }];
  }
  if (findRow_('stock_in_log', rid) > -1) return { ok: true, recordId: rid, duplicate: true };
  const vendor = (d.vendor || '').trim();
  const note = (d.note || '').trim();
  const results = [];
  let totalQty = 0;
  for (const it of items) {
    const iid = it.id || it.itemId;
    const q = parseInt(it.qty, 10) || 0;
    if (!iid || q <= 0) continue;
    const rowNum = findRow_('items', iid);
    if (rowNum < 0) { results.push({ id: iid, error: '품목 없음' }); continue; }
    const cur = getRowObj_('items', rowNum);
    cur.stock = (Number(cur.stock) || 0) + q;
    cur.totalIn = (Number(cur.totalIn) || 0) + q;
    updateRow_('items', rowNum, cur);
    results.push({ id: iid, name: cur.name, size: cur.size, qty: q, newStock: cur.stock, added: q });
    totalQty += q;
  }
  if (totalQty > 0) {
    appendRow_('stock_in_log', {
      id: rid, date: d.date || nowDate_(), time: d.time || nowTime_(),
      vendor: vendor, note: note, items: JSON.stringify(results), totalQty: totalQty,
    });
  }
  return { ok: true, recordId: rid, items: results, count: results.length, totalQty: totalQty, vendor: vendor, note: note };
}

// ─── 재고조정 ─────────────────────────────────────────
function actionStockAdj(d) {
  const rid = d.recordId || genId_('a');
  if (!d.itemId) return { error: 'invalid itemId' };
  if (findRow_('stock_adj', rid) > -1) return { ok: true, recordId: rid, duplicate: true };
  const typeNames = { loss: '분실', damage: '파손', audit: '실사조정', other: '기타' };
  const adjType = d.adjType || 'other';
  const before = parseInt(d.before, 10) || 0;
  const after = parseInt(d.after, 10) || 0;
  const rowNum = findRow_('items', d.itemId);
  if (rowNum > -1) {
    const cur = getRowObj_('items', rowNum);
    cur.stock = after;
    updateRow_('items', rowNum, cur);
  }
  appendRow_('stock_adj', {
    id: rid, type: adjType, typeName: d.typeName || typeNames[adjType] || '기타',
    itemId: d.itemId, itemName: d.itemName || '', itemSize: d.itemSize || '',
    before: before, after: after, delta: after - before, note: d.note || '',
    date: d.date || nowDate_(), time: d.time || nowTime_(),
  });
  return { ok: true, recordId: rid };
}

// ─── 입고 이력 수정 ───────────────────────────────────
function actionStockInUpdate(d) {
  if (!d.recordId) return { error: 'recordId 필요' };
  const rowNum = findRow_('stock_in_log', d.recordId);
  if (rowNum < 0) return { error: '이력 없음' };
  const rec = getRowObj_('stock_in_log', rowNum);
  const oldItems = parseItems_(rec.items);
  const oldMap = {}, newMap = {};
  oldItems.forEach(it => { oldMap[it.id] = parseInt(it.qty, 10) || 0; });
  (d.items || []).forEach(it => { newMap[it.id] = parseInt(it.qty, 10) || 0; });
  for (const iid in newMap) {
    if (!(iid in oldMap)) return { error: '신규 품목 추가는 비지원' };
    if (newMap[iid] < 0) return { error: 'qty는 0 이상' };
  }
  const rebuilt = [];
  let totalQty = 0;
  for (const it of oldItems) {
    const oldQ = parseInt(it.qty, 10) || 0;
    const newQ = (it.id in newMap) ? newMap[it.id] : oldQ;
    const diff = newQ - oldQ;
    if (diff !== 0) {
      const irow = findRow_('items', it.id);
      if (irow > -1) {
        const cur = getRowObj_('items', irow);
        cur.stock = (Number(cur.stock) || 0) + diff;
        cur.totalIn = (Number(cur.totalIn) || 0) + diff;
        updateRow_('items', irow, cur);
      }
    }
    const ni = JSON.parse(JSON.stringify(it));
    ni.qty = newQ; ni.added = newQ;
    rebuilt.push(ni);
    totalQty += newQ;
  }
  rec.vendor = (d.vendor || '').trim();
  rec.note = (d.note || '').trim();
  rec.items = JSON.stringify(rebuilt);
  rec.totalQty = totalQty;
  updateRow_('stock_in_log', rowNum, rec);
  return { ok: true, recordId: d.recordId, totalQty: totalQty };
}

// ─── 입고 이력 삭제 ───────────────────────────────────
function actionStockInDelete(d) {
  if (!d.recordId) return { error: 'recordId 필요' };
  const rowNum = findRow_('stock_in_log', d.recordId);
  if (rowNum < 0) return { error: '이력 없음' };
  const rec = getRowObj_('stock_in_log', rowNum);
  for (const it of parseItems_(rec.items)) {
    const q = parseInt(it.qty, 10) || 0;
    if (!it.id || q === 0) continue;
    const irow = findRow_('items', it.id);
    if (irow > -1) {
      const cur = getRowObj_('items', irow);
      cur.stock = (Number(cur.stock) || 0) - q;
      cur.totalIn = (Number(cur.totalIn) || 0) - q;
      updateRow_('items', irow, cur);
    }
  }
  sheet_('stock_in_log').deleteRow(rowNum);
  return { ok: true, recordId: d.recordId };
}

// ─── 재고조정 이력 수정 ───────────────────────────────
function actionStockAdjUpdate(d) {
  if (!d.recordId) return { error: 'recordId 필요' };
  const rowNum = findRow_('stock_adj', d.recordId);
  if (rowNum < 0) return { error: '이력 없음' };
  const typeNames = { loss: '분실', damage: '파손', audit: '실사조정', other: '기타' };
  const rec = getRowObj_('stock_adj', rowNum);
  const newAfter = parseInt(d.after, 10) || 0;
  const diff = newAfter - (Number(rec.after) || 0);
  if (diff !== 0) {
    const irow = findRow_('items', rec.itemId);
    if (irow > -1) {
      const cur = getRowObj_('items', irow);
      cur.stock = (Number(cur.stock) || 0) + diff;
      updateRow_('items', irow, cur);
    }
  }
  rec.type = d.adjType || 'other';
  rec.typeName = d.typeName || typeNames[rec.type] || '기타';
  rec.after = newAfter;
  rec.delta = newAfter - (Number(rec.before) || 0);
  rec.note = d.note || '';
  updateRow_('stock_adj', rowNum, rec);
  return { ok: true, recordId: d.recordId };
}

// ─── 재고조정 이력 삭제 ───────────────────────────────
function actionStockAdjDelete(d) {
  if (!d.recordId) return { error: 'recordId 필요' };
  const rowNum = findRow_('stock_adj', d.recordId);
  if (rowNum < 0) return { error: '이력 없음' };
  const rec = getRowObj_('stock_adj', rowNum);
  const delta = Number(rec.delta) || 0;
  if (delta !== 0) {
    const irow = findRow_('items', rec.itemId);
    if (irow > -1) {
      const cur = getRowObj_('items', irow);
      cur.stock = (Number(cur.stock) || 0) - delta;
      updateRow_('items', irow, cur);
    }
  }
  sheet_('stock_adj').deleteRow(rowNum);
  return { ok: true, recordId: d.recordId };
}

// ─── 레거시 백필 ──────────────────────────────────────
function actionBackfillBatch(d) {
  const records = d.records || [];
  if (!records.length) return { error: 'no records' };
  const itemsTable = readTable_('items');
  const success = [], failed = [];
  const rows = [];
  for (const r of records) {
    try {
      const rid = r.id || genId_('h');
      const name = r.empName || r.name || '';
      if (!name) throw new Error('empName 누락');
      const norm = (r.items || []).map(it => {
        let iid = it.id;
        if (!iid) {
          const m = itemsTable.find(x => x.name === (it.name || '') && (x.size === (it.size || 'FREE') || (x.size === 'FREE' && !it.size)));
          iid = m ? m.id : '';
        }
        return { id: iid, name: it.name || '', size: it.size || '', qty: parseInt(it.qty, 10) || 0 };
      });
      rows.push(toRow_('history', {
        id: rid, type: 'backfill', empId: '', name: name, dept: r.dept || '', rank: r.rank || '',
        date: r.date || '', time: r.time || '00:00', items: JSON.stringify(norm),
        refId: '', note: r.note || 'legacy', sigData: '', backfilled: 1,
      }));
      success.push(rid);
    } catch (e) {
      failed.push({ id: r.id || '', error: String(e) });
    }
  }
  if (rows.length) {
    const sh = sheet_('history');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, SCHEMA.history.length).setValues(rows);
  }
  return { ok: true, success: success, failed: failed, appended: success.length };
}

// ─── PDF 업로드 (Google Drive) ────────────────────────
function actionUploadPdf(d) {
  let filename = (d.filename || '').trim();
  let payload = (d.data || '').trim();
  if (!filename) return { error: 'filename 필수' };
  if (!payload) return { error: 'data 필수' };
  if (payload.indexOf(',') > -1) payload = payload.split(',')[1];
  let safe = filename.replace(/[^\w\-. 가-힣]/g, '_');
  if (!/\.pdf$/i.test(safe)) safe += '.pdf';
  const tz = Session.getScriptTimeZone();
  const sub = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  // 루트 폴더(피복관리_PDF) → 월별 하위 폴더
  let root;
  const ri = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
  root = ri.hasNext() ? ri.next() : DriveApp.createFolder(PDF_FOLDER_NAME);
  let folder;
  const fi = root.getFoldersByName(sub);
  folder = fi.hasNext() ? fi.next() : root.createFolder(sub);
  let bytes;
  try { bytes = Utilities.base64Decode(payload); }
  catch (e) { return { error: 'base64 디코드 실패: ' + e }; }
  // 멱등성: 동명 + 동일 크기 → 같은 파일로 간주
  const ex = folder.getFilesByName(safe);
  if (ex.hasNext()) {
    const f = ex.next();
    if (f.getSize() === bytes.length) {
      return { ok: true, url: f.getUrl(), path: sub + '/' + safe, size: bytes.length, duplicate: true };
    }
    safe = safe.replace(/\.pdf$/i, '') + '_' + Utilities.formatDate(new Date(), tz, 'HHmmss') + '.pdf';
  }
  const file = folder.createFile(Utilities.newBlob(bytes, 'application/pdf', safe));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { ok: true, url: file.getUrl(), path: sub + '/' + safe, size: bytes.length };
}

// ─── 직원 관리 ────────────────────────────────────────
function actionUpdateEmp(d) {
  const empId = (d.empId || '').trim();
  if (!empId) return { error: 'empId 필요' };
  const name = (d.name || '').trim();
  if (!name) return { error: '성명 필수' };
  const status = d.status || 'active';
  if (status !== 'active' && status !== 'retired') return { error: 'status는 active 또는 retired' };
  const obj = {
    empId: empId, name: name, dept: (d.dept || '').trim(), rank: (d.rank || '').trim(),
    joinDate: (d.joinDate || '').trim(), status: status, retiredDate: (d.retiredDate || '').trim(),
  };
  const rowNum = findRow_('emps', empId);
  if (rowNum > -1) updateRow_('emps', rowNum, obj);
  else appendRow_('emps', obj);
  return { ok: true, empId: empId };
}

function actionDeleteEmp(d) {
  const empId = (d.empId || '').trim();
  if (!empId) return { error: 'empId 필요' };
  const rowNum = findRow_('emps', empId);
  if (rowNum > -1) sheet_('emps').deleteRow(rowNum);
  return { ok: true };
}

function actionUpdateEmpStatus(d) {
  const status = d.status;
  if (status !== 'active' && status !== 'retired') return { error: 'invalid status' };
  let rowNum = -1;
  if (d.empId) rowNum = findRow_('emps', d.empId);
  else if (d.empName) {
    const emps = readTable_('emps');
    const m = emps.find(e => e.name === d.empName);
    if (m) rowNum = findRow_('emps', m.empId);
  } else return { error: 'empId 또는 empName 필요' };
  if (rowNum < 0) return { error: '직원 없음' };
  const rec = getRowObj_('emps', rowNum);
  rec.status = status;
  rec.retiredDate = d.retiredDate || '';
  updateRow_('emps', rowNum, rec);
  return { ok: true, updated: rec.name };
}

// ─── 데이터 일괄 주입 (NAS → Sheets 마이그레이션 전용) ──
function actionImportAll(d) {
  const table = d.table;
  if (!SCHEMA[table]) return { error: 'unknown table: ' + table };
  const rows = d.rows || [];
  const sh = sheet_(table);
  if (d.replace) {
    const last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  }
  if (rows.length) {
    const vals = rows.map(o => toRow_(table, o));
    sh.getRange(sh.getLastRow() + 1, 1, vals.length, SCHEMA[table].length).setValues(vals);
  }
  return { ok: true, table: table, imported: rows.length };
}

// ─── HTTP 엔트리 ──────────────────────────────────────
function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== API_KEY) return json_({ error: 'unauthorized' });
  if (p.action === 'getAll') return json_(actionGetAll());
  return json_({ error: 'unknown action: ' + (p.action || '') });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  if (body.key !== API_KEY) return json_({ error: 'unauthorized' });
  const action = body.action || '';
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return json_({ error: '서버 사용 중 — 잠시 후 재시도' });
  }
  try {
    switch (action) {
      case 'distribute':               return json_(actionDistribute(body));
      case 'return':                   return json_(actionReturn(body));
      case 'stockIn':                  return json_(actionStockIn(body));
      case 'stockAdj':                 return json_(actionStockAdj(body));
      case 'stockInUpdate':            return json_(actionStockInUpdate(body));
      case 'stockInDelete':            return json_(actionStockInDelete(body));
      case 'stockAdjUpdate':           return json_(actionStockAdjUpdate(body));
      case 'stockAdjDelete':           return json_(actionStockAdjDelete(body));
      case 'distributeBackfillBatch':  return json_(actionBackfillBatch(body));
      case 'updateEmpStatus':          return json_(actionUpdateEmpStatus(body));
      case 'updateEmp':                return json_(actionUpdateEmp(body));
      case 'deleteEmp':                return json_(actionDeleteEmp(body));
      case 'uploadPdf':                return json_(actionUploadPdf(body));
      case 'importAll':                return json_(actionImportAll(body));
      default:                         return json_({ error: 'unknown action: ' + action });
    }
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
