/**
 * 三信家商餐飲科 進貨驗收（給管理者用平板點貨）
 * 跟訂貨系統共用同一份試算表：讀「叫貨紀錄」（③ 叫貨單，已確定的），驗收結果寫進「驗收紀錄」「驗收明細」分頁。
 * 部署成網頁應用程式：執行身分「我」、存取權「只有我自己」（用你的 Google 帳號登入才打得開，不用另外的帳號密碼）。
 */

const SPREADSHEET_ID = '1qIQe4STafrhzQk4I7NLkZQxJGe5hozf0g04zvJ-qZ_I';   // 訂貨系統的試算表
const SITE_TITLE = '三信家商餐飲科 進貨驗收';
const PHOTO_FOLDER_NAME = '三信家商餐飲科進貨銷貨單';   // 銷貨單照片放在雲端硬碟這個資料夾（每週一個子資料夾）
const SHEET = { purchase: '叫貨紀錄', checks: '驗收紀錄', lines: '驗收明細' };
const CHECK_HEADERS = ['驗收編號', '送貨日期（起）', '送貨日期（迄）', '廠商', '送貨組別', '驗收時間', '品項數', '問題數', '銷貨單照片', '備註', '驗收人'];
const LINE_HEADERS = ['驗收編號', '廠商', '送貨組別', '班級', '食材名稱', '應到', '實到', '狀態', '備註', '處理狀態', '處理時間', '處理備註'];
// 照班級點貨的廠商（其他廠商點品項總量）
const PER_CLASS = /^[ABD]/;
const ISSUE_STATES = ['待處理', '已補送', '扣款', '不用補'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(SITE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------
// 讀這一週要驗收的叫貨單
// ---------------------------------------------------------------------

/** monday：'yyyy-MM-dd'（那週星期一）。回傳每一張 ③ 叫貨單（廠商＋送貨組別）和驗收狀態 */
function getWeek(monday) {
  const from = parseDate_(monday) || mondayOf_(today_());
  const to = addDays_(from, 4);
  const f = dateStr_(from), t = dateStr_(to);
  const rows = readTable_(SHEET.purchase, ['送貨日期', '廠商', '食材名稱']).rows.map(function (r) {
    return {
      date: purchaseDate_(r['送貨日期']), group: String(r['送貨組別'] || '').trim(), vendor: String(r['廠商'] || '').trim(),
      className: String(r['上課班級'] || '').trim(), teacher: String(r['老師'] || '').trim(), name: String(r['食材名稱'] || '').trim(),
      qty: Number(r['叫貨量']) || 0, unit: String(r['叫貨單位'] || '').trim(), kind: String(r['類型'] || '訂單').trim(),
      note: String(r['備註'] || '').trim(), status: String(r['狀態'] || '').trim(),
    };
  }).filter(function (r) { return r.date >= f && r.date <= t && r.vendor && r.name; });

  const drafts = {}, by = {};
  rows.forEach(function (r) {
    if (r.status === '草稿') { drafts[r.vendor] = true; return; }
    const k = r.vendor + '|' + r.group;
    (by[k] = by[k] || []).push(r);
  });
  const checked = readChecks_(f);
  const batches = Object.keys(by).sort().map(function (k) {
    const list = by[k], vendor = list[0].vendor, group = list[0].group;
    const id = checkId_(f, vendor, group);
    const b = { id: id, vendor: vendor, group: group, label: vendor + (group ? '（送貨組別 ' + group + '，星期' + group.split('').join('、') + '）' : ''),
      perClass: PER_CLASS.test(vendor), lines: batchLines_(list, vendor), check: checked[id] || null };
    return b;
  });
  return { from: f, to: t, batches: batches, drafts: Object.keys(drafts).filter(function (v) { return !batches.some(function (b) { return b.vendor === v; }); }) };
}

/** 一張叫貨單要點的品項：A／B／D 照班級（各班明細），其他照品項總量（跟 ③ 叫貨單的合計一樣，含在 ③ 直接改過的合計） */
function batchLines_(list, vendor) {
  const live = list.filter(function (r) { return r.kind !== '合計' && r.qty > 0; });
  const ovr = {};
  list.forEach(function (r) { if (r.kind === '合計') ovr[r.name] = r; });
  if (PER_CLASS.test(vendor)) {
    const out = live.map(function (r) {
      const cls = r.kind === '補庫存' ? '補庫存' : r.className || '（沒填班級）';
      const show = vendorQty_(r.qty, r.unit, vendor, r.name);
      return { key: cls + '|' + r.name + '|' + r.unit, cls: cls, teacher: r.kind === '補庫存' ? '' : r.teacher, date: r.date, name: r.name,
        expect: show, note: r.note, ovr: ovr[r.name] ? '合計改成 ' + (ovr[r.name].qty > 0 ? fmtQty_(ovr[r.name].qty) + ' ' + ovr[r.name].unit : '不叫') : '' };
    });
    // 被直接改成「不叫」的品項不用點
    return out.filter(function (x) { return !(ovr[x.name] && !(ovr[x.name].qty > 0)); })
      .sort(function (a, b) { return (a.date + a.cls).localeCompare(b.date + b.cls, 'zh-Hant') || a.name.localeCompare(b.name, 'zh-Hant'); });
  }
  const tot = {};
  live.forEach(function (r) {
    if (ovr[r.name]) return;
    const cv = convertUnit_(r.qty, r.unit, vendor, r.name), k = r.name + '|' + cv.unit;
    (tot[k] = tot[k] || { name: r.name, unit: cv.unit, qty: 0 }).qty += cv.qty;
  });
  Object.keys(ovr).forEach(function (n) {
    const o = ovr[n];
    if (!(o.qty > 0)) return;
    const cv = convertUnit_(o.qty, o.unit, vendor, n);
    tot[n + '|' + cv.unit] = { name: n, unit: cv.unit, qty: cv.qty };
  });
  return Object.keys(tot).map(function (k) {
    const t = tot[k];
    return { key: '|' + t.name + '|' + t.unit, cls: '', name: t.name, expect: fmtTotal_(t.qty, t.unit, t.name) };
  }).sort(function (a, b) { return a.name.localeCompare(b.name, 'zh-Hant'); });
}

function readChecks_(from) {
  const out = {};
  const sh = ss_().getSheetByName(SHEET.checks);
  if (!sh || sh.getLastRow() < 2) return out;
  const lines = linesById_();
  readTable_(SHEET.checks, ['驗收編號']).rows.forEach(function (r) {
    if (purchaseDate_(r['送貨日期（起）']) !== from) return;
    const id = String(r['驗收編號']).trim();
    out[id] = { time: fmtDateTime_(r['驗收時間']), items: Number(r['品項數']) || 0, issues: Number(r['問題數']) || 0,
      photos: String(r['銷貨單照片'] || '').split(/\n+/).filter(Boolean), note: String(r['備註'] || ''), lines: lines[id] || [] };
  });
  return out;
}

function linesById_() {
  const out = {};
  const sh = ss_().getSheetByName(SHEET.lines);
  if (!sh || sh.getLastRow() < 2) return out;
  readTable_(SHEET.lines, ['驗收編號']).rows.forEach(function (r) {
    const id = String(r['驗收編號']).trim();
    (out[id] = out[id] || []).push({ cls: String(r['班級'] || ''), name: String(r['食材名稱'] || ''),
      got: String(r['實到'] || ''), status: String(r['狀態'] || ''), note: String(r['備註'] || '') });
  });
  return out;
}

function checkId_(from, vendor, group) { return 'R' + from.replace(/-/g, '') + '-' + vendor + (group ? '-' + group : ''); }

// ---------------------------------------------------------------------
// 存驗收結果
// ---------------------------------------------------------------------

/**
 * p = { from, to, vendor, group, note, photos: [url], items: [{ cls, name, expect, got, status, note }] }
 * status：到齊／少送／沒送／送錯／多送。同一張叫貨單重新驗收會蓋掉上次的（已經處理過的問題保留處理狀態）。
 */
function saveCheck(p) {
  const from = dateStr_(parseDate_(p.from)), to = dateStr_(parseDate_(p.to));
  const vendor = clean_(p.vendor, 30), group = clean_(p.group, 10);
  if (!from || !vendor) throw new Error('資料不完整，請重新整理。');
  const items = (p.items || []).map(function (x) {
    const st = ['到齊', '少送', '沒送', '送錯', '多送'].indexOf(x.status) >= 0 ? x.status : '';
    if (!st) throw new Error('「' + x.name + '」還沒點：請按 ✓ 或標記少送、沒送、送錯。');
    if ((st === '少送' || st === '多送') && !String(x.got || '').trim()) throw new Error('「' + x.name + '」' + st + '，請填實際送來多少。');
    return { cls: clean_(x.cls, 30), name: clean_(x.name, 40), expect: clean_(x.expect, 30), got: st === '到齊' ? clean_(x.expect, 30) : st === '沒送' ? '0' : clean_(x.got, 30),
      status: st, note: clean_(x.note, 100) };
  });
  const id = checkId_(from, vendor, group);
  const issues = items.filter(function (x) { return x.status !== '到齊'; });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.checks, CHECK_HEADERS);
    ensureSheet_(SHEET.lines, LINE_HEADERS);
    // 上次驗收過的問題，處理狀態保留
    const keep = {};
    const lt = readTable_(SHEET.lines, ['驗收編號']);
    const oldLines = lt.rows.filter(function (r) { return String(r['驗收編號']).trim() === id; });
    oldLines.forEach(function (r) {
      keep[String(r['班級']) + '|' + String(r['食材名稱'])] = { st: r['處理狀態'], t: r['處理時間'], n: r['處理備註'] };
    });
    if (oldLines.length) deleteRows_(lt.sheet, oldLines.map(function (r) { return r._row; }));
    const ct = readTable_(SHEET.checks, ['驗收編號']);
    const oldChecks = ct.rows.filter(function (r) { return String(r['驗收編號']).trim() === id; });
    if (oldChecks.length) deleteRows_(ct.sheet, oldChecks.map(function (r) { return r._row; }));
    appendRows_(SHEET.checks, [{ '驗收編號': id, '送貨日期（起）': from, '送貨日期（迄）': to, '廠商': vendor, '送貨組別': group,
      '驗收時間': new Date(), '品項數': items.length, '問題數': issues.length, '銷貨單照片': (p.photos || []).join('\n'),
      '備註': clean_(p.note, 200), '驗收人': Session.getActiveUser().getEmail() || '' }]);
    appendRows_(SHEET.lines, items.map(function (x) {
      const k = keep[x.cls + '|' + x.name];
      const issue = x.status !== '到齊';
      return { '驗收編號': id, '廠商': vendor, '送貨組別': group, '班級': x.cls, '食材名稱': x.name, '應到': x.expect,
        '實到': x.got, '狀態': x.status, '備註': x.note,
        '處理狀態': issue ? (k && k.st && k.st !== '' ? k.st : '待處理') : '', '處理時間': issue && k ? k.t || '' : '', '處理備註': issue && k ? k.n || '' : '' };
    }));
  } finally {
    lock.releaseLock();
  }
  let mail = '';
  if (issues.length) {
    try {
      const mailTo = Session.getEffectiveUser().getEmail();
      if (mailTo) {
        MailApp.sendEmail({ to: mailTo, subject: '【進貨驗收】' + vendor + ' 有 ' + issues.length + ' 項問題（' + rangeText_(from, to) + '）',
          htmlBody: '<p>' + escHtml_(vendor + (group ? ' 送貨組別 ' + group : '') + '，' + rangeText_(from, to)) + '：</p><ul>' +
            issues.map(function (x) { return '<li>' + escHtml_((x.cls ? x.cls + ' ' : '') + x.name + '：' + issueText_(x)) + '</li>'; }).join('') + '</ul>' +
            '<p>到進貨驗收網站的「問題追蹤」複製給廠商的文字、記錄處理結果。</p>' });
        mail = '已寄 Email 提醒。';
      }
    } catch (e) { mail = 'Email 寄送失敗：' + e.message; }
  }
  return { id: id, items: items.length, issues: issues.length, mail: mail };
}

function issueText_(x) {
  if (x.status === '沒送') return '沒送（應到 ' + x.expect + '）' + (x.note ? '，' + x.note : '');
  if (x.status === '送錯') return '送錯（應到 ' + x.expect + '）' + (x.note ? '，' + x.note : '');
  return x.status + '：應到 ' + x.expect + '，實到 ' + x.got + (x.note ? '，' + x.note : '');
}

// ---------------------------------------------------------------------
// 銷貨單照片
// ---------------------------------------------------------------------

/** data：base64（不含 data: 開頭），回傳 { url, name } */
function uploadPhoto(data, mime, from, vendor) {
  if (!data) throw new Error('沒有照片。');
  const week = dateStr_(parseDate_(from) || today_());
  const root = folderNamed_(PHOTO_FOLDER_NAME, null);
  const sub = folderNamed_(week + ' 那週', root);
  const name = week + ' ' + clean_(vendor, 30) + ' 銷貨單 ' + Utilities.formatDate(new Date(), tz_(), 'HHmmss') + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mime || 'image/jpeg', name);
  const f = sub.createFile(blob);
  return { url: f.getUrl(), name: name, id: f.getId() };
}

function folderNamed_(name, parent) {
  const it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : (parent ? parent.createFolder(name) : DriveApp.createFolder(name));
}

// ---------------------------------------------------------------------
// 問題追蹤
// ---------------------------------------------------------------------

/** all = true 連處理完的一起列（最近 60 天） */
function getIssues(all) {
  const sh = ss_().getSheetByName(SHEET.lines);
  if (!sh || sh.getLastRow() < 2) return { list: [], states: ISSUE_STATES };
  const checks = {};
  if (ss_().getSheetByName(SHEET.checks)) readTable_(SHEET.checks, ['驗收編號']).rows.forEach(function (r) {
    checks[String(r['驗收編號']).trim()] = { from: purchaseDate_(r['送貨日期（起）']), to: purchaseDate_(r['送貨日期（迄）']),
      time: fmtDateTime_(r['驗收時間']), photos: String(r['銷貨單照片'] || '').split(/\n+/).filter(Boolean) };
  });
  const since = dateStr_(addDays_(today_(), -60));
  const list = readTable_(SHEET.lines, ['驗收編號', '狀態']).rows.filter(function (r) {
    const st = String(r['狀態'] || '').trim();
    if (!st || st === '到齊') return false;
    const done = String(r['處理狀態'] || '').trim() && String(r['處理狀態']).trim() !== '待處理';
    const c = checks[String(r['驗收編號']).trim()];
    return all ? !c || c.from >= since : !done;
  }).map(function (r) {
    const id = String(r['驗收編號']).trim(), c = checks[id] || {};
    return { row: r._row, id: id, vendor: String(r['廠商'] || ''), group: String(r['送貨組別'] || ''), from: c.from || '', to: c.to || '', time: c.time || '',
      photos: c.photos || [], cls: String(r['班級'] || ''), name: String(r['食材名稱'] || ''), expect: String(r['應到'] || ''), got: String(r['實到'] || ''),
      status: String(r['狀態'] || ''), note: String(r['備註'] || ''), state: String(r['處理狀態'] || '待處理') || '待處理',
      doneTime: fmtDateTime_(r['處理時間']), doneNote: String(r['處理備註'] || '') };
  });
  return { list: list, states: ISSUE_STATES };
}

/** 記錄問題的處理結果：rows = 驗收明細的列號 */
function setIssues(rows, state, note) {
  if (ISSUE_STATES.indexOf(state) < 0) throw new Error('處理狀態不正確。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.lines, ['驗收編號', '處理狀態']);
    (rows || []).forEach(function (row) {
      const r = t.rows.find(function (x) { return x._row === Number(row); });
      if (!r) return;
      setCell_(t, r._row, '處理狀態', state);
      setCell_(t, r._row, '處理時間', state === '待處理' ? '' : new Date());
      if (note !== undefined && note !== null) setCell_(t, r._row, '處理備註', clean_(note, 100));
    });
  } finally {
    lock.releaseLock();
  }
  return getIssues(false);
}

// ---------------------------------------------------------------------
// 單位換算（跟訂貨系統 ③ 叫貨單一樣）
// ---------------------------------------------------------------------

const GRAMS = { 'g': 1, '公斤': 1000, '斤': 600, '兩': 37.5, 'T': 5, 't': 3 };   // 1 大T＝5 g、1 小T＝3 g（跟訂貨系統一樣）
const COUNT_UNITS = { '個': '顆', '粒': '顆', '顆': '顆', '根': '條', '支': '條', '條': '條' };
const EGG_GRAMS = 50;
function isEgg_(name) { return String(name || '').trim() === '雞蛋'; }
function gramVendor_(v) { return /^[EF]/.test(v || ''); }
function convertUnit_(qty, unit, vendor, name) {
  if (GRAMS[unit] && isEgg_(name)) return { qty: Math.round(qty * GRAMS[unit] / EGG_GRAMS * 100) / 100, unit: '顆' };
  if (!GRAMS[unit]) return { qty: qty, unit: COUNT_UNITS[unit] || unit };
  if (gramVendor_(vendor)) return { qty: Math.round(qty * GRAMS[unit] * 100) / 100, unit: 'g' };
  const target = /^[AB]/.test(vendor || '') ? '兩' : '公斤';
  return { qty: Math.round(qty * GRAMS[unit] / GRAMS[target] * 100) / 100, unit: target };
}
function vendorQty_(qty, unit, vendor, name) {
  if (gramVendor_(vendor) && GRAMS[unit] && !isEgg_(name)) return Math.ceil(Number(qty) * GRAMS[unit] - 1e-6) + ' g';
  const cv = convertUnit_(Number(qty), unit, vendor, name);
  if (cv.unit === '兩' || (cv.unit === '顆' && isEgg_(name) && GRAMS[unit])) return fmtTotal_(cv.qty, cv.unit, name);
  return fmtQty_(qty) + ' ' + unit;
}
function fmtTotal_(qty, unit, name) {
  if (unit === '兩') {
    const n = Math.ceil(qty - 1e-6), jin = Math.floor(n / 16), liang = n % 16;
    return jin ? jin + ' 斤' + (liang ? ' ' + liang + ' 兩' : '') : n + ' 兩';
  }
  if (unit === '顆' && isEgg_(name)) return Math.ceil(qty - 1e-6) + ' 顆';
  if (unit === 'g') return Math.ceil(qty - 1e-6) + ' g';
  return fmtQty_(qty) + ' ' + unit;
}
function fmtQty_(n) { return String(Math.round(Number(n) * 100) / 100); }

// ---------------------------------------------------------------------
// 試算表小工具（跟訂貨系統一樣，照表頭名稱讀寫）
// ---------------------------------------------------------------------

let SS_CACHE_ = null;
function ss_() {
  if (!SS_CACHE_) SS_CACHE_ = retry_(function () { return SpreadsheetApp.openById(SPREADSHEET_ID); });
  return SS_CACHE_;
}

function readTable_(name, required) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('試算表找不到「' + name + '」分頁。');
  const lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
  const values = lastRow ? retry_(function () { return sh.getRange(1, 1, lastRow, lastCol).getValues(); }) : [[]];
  const cols = {};
  values[0].forEach(function (h, i) { const k = String(h).trim(); if (k && cols[k] === undefined) cols[k] = i + 1; });
  (required || []).forEach(function (k) { if (cols[k] === undefined) throw new Error('分頁「' + name + '」缺少「' + k + '」欄位。'); });
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].every(function (v) { return v === '' || v === null; })) continue;
    const o = { _row: i + 1 };
    Object.keys(cols).forEach(function (k) { o[k] = values[i][cols[k] - 1]; });
    rows.push(o);
  }
  return { sheet: sh, cols: cols, rows: rows };
}

function ensureSheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function appendRows_(name, objs) {
  if (!objs.length) return;
  const sh = ss_().getSheetByName(name);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const data = objs.map(function (o) { return headers.map(function (h) { return o[h] === undefined ? '' : o[h]; }); });
  const start = sh.getLastRow() + 1;
  retry_(function () { sh.getRange(start, 1, data.length, headers.length).setValues(data); });
}

function setCell_(table, row, header, value) {
  const c = table.cols[header];
  if (c === undefined) throw new Error('分頁缺少「' + header + '」欄位。');
  table.sheet.getRange(row, c).setValue(value);
}

/** 從下往上刪，列號才不會跑掉 */
function deleteRows_(sh, rows) {
  rows.slice().sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
}

function retry_(fn) {
  let lastErr;
  for (let i = 0; i < 3; i++) {
    try { return fn(); } catch (e) {
      lastErr = e;
      if (!/Service|服務|無法正常運作|timed out|逾時|Internal error|內部錯誤|try again|稍後再試/i.test(String(e && e.message || e))) throw e;
      Utilities.sleep(800 * (i + 1));
    }
  }
  throw new Error('Google 試算表暫時沒有回應，請稍等一下再試一次。（' + (lastErr && lastErr.message) + '）');
}

function tz_() { return Session.getScriptTimeZone() || 'Asia/Taipei'; }
function today_() {
  const s = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd').split('-');
  return new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]));
}
function addDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function mondayOf_(d) { return addDays_(d, -((d.getDay() + 6) % 7)); }
function parseDate_(s) {
  const m = String(s || '').trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return d.getMonth() === Number(m[2]) - 1 ? d : null;
}
function toDate_(v) { return v instanceof Date ? new Date(v.getFullYear(), v.getMonth(), v.getDate()) : parseDate_(v); }
function dateStr_(d) { return d ? d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) : ''; }
function purchaseDate_(v) { const d = toDate_(v); return d ? dateStr_(d) : String(v || '').trim(); }
function fmtDateTime_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy/MM/dd HH:mm') : String(v || ''); }
function rangeText_(from, to) {
  const f = function (s) { const p = s.split('-'); return Number(p[1]) + '/' + Number(p[2]); };
  return f(from) + (to && to !== from ? '～' + f(to) : '');
}
function clean_(v, maxLen) {
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().replace(/^[=+\-@]+/, '').trim().slice(0, maxLen);
}
function escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
}
