/**
 * 三信家商餐飲科 進貨驗收（給管理者用平板點貨）
 * 跟訂貨系統共用同一份試算表：讀「叫貨紀錄」（③ 叫貨單，已確定的），驗收結果寫進「驗收紀錄」「驗收明細」分頁。
 * 部署成網頁應用程式：執行身分「我」、存取權「只有我自己」（用你的 Google 帳號登入才打得開，不用另外的帳號密碼）。
 */

const SPREADSHEET_ID = '1qIQe4STafrhzQk4I7NLkZQxJGe5hozf0g04zvJ-qZ_I';   // 訂貨系統的試算表
const SITE_TITLE = '三信家商餐飲科 進貨驗收';
const PHOTO_FOLDER_NAME = '三信家商餐飲科進貨銷貨單';   // 銷貨單照片放在雲端硬碟這個資料夾（每週一個子資料夾）
const SHEET = { purchase: '叫貨紀錄', checks: '驗收紀錄', lines: '驗收明細', slip: '銷貨單明細', prices: '單價表', items: '食材總表' };
const CHECK_HEADERS = ['驗收編號', '送貨日期（起）', '送貨日期（迄）', '廠商', '送貨組別', '驗收時間', '品項數', '問題數', '銷貨單照片', '備註', '驗收人', '銷貨單金額', '銷貨單合計'];
const SLIP_HEADERS = ['驗收編號', '送貨日期（起）', '廠商', '食材名稱', '銷貨單品名', '數量', '單位', '單價', '金額', '貨品編號'];
const PRICE_HEADERS = ['廠商', '食材名稱', '單位', '單價', '銷貨單品名', '更新日期', '驗收編號', '貨品編號'];
const LINE_HEADERS = ['驗收編號', '廠商', '送貨組別', '班級', '食材名稱', '應到', '實到', '狀態', '備註', '處理狀態', '處理時間', '處理備註'];
// 照班級點貨的廠商（其他廠商點品項總量）
const PER_CLASS = /^[ABD]/;
const ISSUE_STATES = ['待處理', '已補送', '扣款', '不用補'];
// 同一家廠商、送兩個地方：訂貨系統裡分開記（C烘焙蛋／C中餐蛋），驗收時合成一張，分「送烘焙教室」「送中餐教室」兩段
const SITE_VENDORS = { 'C烘焙蛋': { base: 'C蛋商', site: '送烘焙教室' }, 'C中餐蛋': { base: 'C蛋商', site: '送中餐教室' } };

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
    const sv = SITE_VENDORS[r.vendor];
    if (r.status === '草稿') { drafts[sv ? sv.base : r.vendor] = true; return; }
    const k = sv ? sv.base + '|' : r.vendor + '|' + r.group;
    (by[k] = by[k] || []).push(r);
  });
  const checked = readChecks_(f);
  const batches = Object.keys(by).sort().map(function (k) {
    const list = by[k], sv = SITE_VENDORS[list[0].vendor];
    const vendor = sv ? sv.base : list[0].vendor, group = sv ? '' : list[0].group;
    const id = checkId_(f, vendor, group);
    const b = { id: id, vendor: vendor, group: group, label: vendor + (group ? '（送貨組別 ' + group + '，星期' + group.split('').join('、') + '）' : ''),
      perClass: PER_CLASS.test(vendor), lines: sv ? siteLines_(list) : batchLines_(list, vendor), check: checked[id] || null };
    return b;
  });
  return { from: f, to: t, batches: batches, drafts: Object.keys(drafts).filter(function (v) { return !batches.some(function (b) { return b.vendor === v; }); }) };
}

/** 送兩個地方的廠商：各地方分開算總量，班級欄放「送烘焙教室／送中餐教室」 */
function siteLines_(list) {
  const out = [];
  Object.keys(SITE_VENDORS).map(function (v) { return [v, SITE_VENDORS[v].site]; })
    .sort(function (a, b) { return a[1].localeCompare(b[1], 'zh-Hant'); }).forEach(function (p) {
      batchLines_(list.filter(function (r) { return r.vendor === p[0]; }), p[0]).forEach(function (x) {
        out.push(Object.assign({}, x, { key: p[1] + x.key, cls: p[1] }));
      });
    });
  return out;
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
  const lines = linesById_(), slips = slipsById_();
  readTable_(SHEET.checks, ['驗收編號']).rows.forEach(function (r) {
    if (purchaseDate_(r['送貨日期（起）']) !== from) return;
    const id = String(r['驗收編號']).trim();
    out[id] = { time: fmtDateTime_(r['驗收時間']), items: Number(r['品項數']) || 0, issues: Number(r['問題數']) || 0,
      photos: String(r['銷貨單照片'] || '').split(/\n+/).filter(Boolean), note: String(r['備註'] || ''), lines: lines[id] || [],
      amount: Number(r['銷貨單金額']) || 0, slip: slips[id] ? { rows: slips[id], total: r['銷貨單合計'] === '' || r['銷貨單合計'] === undefined ? '' : String(r['銷貨單合計']) } : null };
  });
  return out;
}

function slipsById_() {
  const out = {};
  const sh = ss_().getSheetByName(SHEET.slip);
  if (!sh || sh.getLastRow() < 2) return out;
  readTable_(SHEET.slip, ['驗收編號']).rows.forEach(function (r) {
    const id = String(r['驗收編號']).trim();
    (out[id] = out[id] || []).push({ name: String(r['食材名稱'] || ''), slipName: String(r['銷貨單品名'] || ''), qty: String(r['數量'] === undefined ? '' : r['數量']),
      unit: String(r['單位'] || ''), price: String(r['單價'] === undefined ? '' : r['單價']), amount: String(r['金額'] === undefined ? '' : r['金額']),
      code: String(r['貨品編號'] || '') });
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
 * p = { from, to, vendor, group, note, photos: [url], items: [{ cls, name, expect, got, status, note }],
 *       slip: { total, rows: [{ name（食材名稱，空白＝叫貨單沒有的）, slipName, qty, unit, price, amount }] }（選填，銷貨單金額） }
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
  const slip = (p.slip && p.slip.rows || []).map(function (x) {
    return { name: clean_(x.name, 40), slipName: clean_(x.slipName, 60), qty: num_(x.qty), unit: clean_(x.unit, 10), price: num_(x.price), amount: num_(x.amount),
      code: clean_(x.code, 20).toUpperCase() };
  }).filter(function (x) { return (x.name || x.slipName) && (x.qty !== '' || x.price !== '' || x.amount !== ''); });
  const slipSum = slip.reduce(function (a, x) { return a + (Number(x.amount) || 0); }, 0);
  const slipTotal = p.slip ? num_(p.slip.total) : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.checks, CHECK_HEADERS);
    ensureSheet_(SHEET.lines, LINE_HEADERS);
    ensureSheet_(SHEET.slip, SLIP_HEADERS);
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
      '備註': clean_(p.note, 200), '驗收人': Session.getActiveUser().getEmail() || '',
      '銷貨單金額': slip.length ? Math.round(slipSum * 100) / 100 : '', '銷貨單合計': slipTotal }]);
    const st = readTable_(SHEET.slip, ['驗收編號']);
    const oldSlip = st.rows.filter(function (r) { return String(r['驗收編號']).trim() === id; });
    if (oldSlip.length) deleteRows_(st.sheet, oldSlip.map(function (r) { return r._row; }));
    if (slip.length) {
      appendRows_(SHEET.slip, slip.map(function (x) {
        return { '驗收編號': id, '送貨日期（起）': from, '廠商': vendor, '食材名稱': x.name, '銷貨單品名': x.slipName, '數量': x.qty, '單位': x.unit, '單價': x.price, '金額': x.amount,
          '貨品編號': x.code };
      }));
      savePrices_(vendor, from, id, slip);
    }
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
  return { id: id, items: items.length, issues: issues.length, mail: mail, amount: slip.length ? Math.round(slipSum * 100) / 100 : 0 };
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
// 第 2 階段：讀銷貨單（Google 雲端硬碟內建的文字辨識，免費）→ 品名、數量、單價、金額
// ---------------------------------------------------------------------

/**
 * photos：這次拍的銷貨單照片網址；names：這張叫貨單的食材名稱。
 * 回傳 { text, rows: { 食材名稱: { slipName, qty, unit, price, amount, sure } }, extras: [{ slipName, qty, unit, price, amount }], total }
 */
function readSlip(photos, vendor, names) {
  const ids = (photos || []).map(fileIdOf_).filter(Boolean);
  if (!ids.length) throw new Error('請先拍銷貨單。');
  const text = ids.slice(0, 6).map(ocrText_).join('\n');
  return parseSlip_(text, slipKeys_(clean_(vendor, 30), (names || []).map(function (n) { return clean_(n, 40); }).filter(Boolean)));
}

function fileIdOf_(url) {
  const m = String(url || '').match(/\/d\/([\w-]{20,})/) || String(url || '').match(/[?&]id=([\w-]{20,})/);
  return m ? m[1] : '';
}

/** 照片 → 文字：轉成 Google 文件（會做中文文字辨識）、匯出純文字，再把暫存的文件丟到垃圾桶 */
function ocrText_(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  const boundary = 'slip' + Date.now();
  const meta = JSON.stringify({ name: '[銷貨單辨識暫存]', mimeType: 'application/vnd.google-apps.document' });
  const head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
    '\r\n--' + boundary + '\r\nContent-Type: ' + (blob.getContentType() || 'image/jpeg') + '\r\n\r\n').getBytes();
  const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const res = retry_(function () {
    return UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id&ocrLanguage=zh-TW', {
      method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload: head.concat(blob.getBytes(), tail), headers: auth, muteHttpExceptions: true });
  });
  if (res.getResponseCode() >= 300) throw new Error('銷貨單辨識失敗（' + res.getResponseCode() + '），請再試一次或手動填。');
  const docId = JSON.parse(res.getContentText()).id;
  try {
    const out = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId + '/export?mimeType=text/plain', { headers: auth, muteHttpExceptions: true });
    return out.getResponseCode() < 300 ? out.getContentText('UTF-8') : '';
  } finally {
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
  }
}

/** 要找的字：食材名稱、食材總表的「別名」、這家廠商以前銷貨單上的品名（單價表記下來的）。回傳 [{ key, name }]，長的先比 */
function slipKeys_(vendor, names) {
  const want = {}, keys = [], codes = {};
  names.forEach(function (n) { want[n] = true; });
  const add = function (k, n) { k = normKey_(k); if (k && want[n] && !keys.some(function (x) { return x.key === k; })) keys.push({ key: k, name: n }); };
  names.forEach(function (n) { add(n, n); });
  if (ss_().getSheetByName(SHEET.prices)) readTable_(SHEET.prices, ['廠商', '食材名稱']).rows.forEach(function (r) {
    if (String(r['廠商']).trim() !== vendor) return;
    const n = String(r['食材名稱']).trim();
    add(r['銷貨單品名'], n);
    const c = String(r['貨品編號'] || '').trim().toUpperCase();
    if (c && want[n]) codes[c] = n;
  });
  if (ss_().getSheetByName(SHEET.items)) {
    const t = readTable_(SHEET.items, ['食材名稱']);
    if (t.cols['別名'] !== undefined) t.rows.forEach(function (r) {
      const n = String(r['食材名稱']).trim();
      if (want[n]) String(r['別名'] || '').split(/[,，、]+/).forEach(function (a) { add(a, n); });
    });
  }
  keys.sort(function (a, b) { return b.key.length - a.key.length; });
  keys.codes = codes;   // 廠商的貨品編號 → 食材名稱（以前存過的）
  return keys;
}

function normKey_(s) {
  return String(s || '').replace(/[\uFF01-\uFF5E]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[\s　()（）\[\]【】*＊]/g, '').toLowerCase();
}

const SLIP_UNITS = /(公斤|kg|台斤|斤|兩|包|箱|盒|顆|粒|罐|瓶|把|袋|條|支|個|片|桶|打|盤|籃|件|組|g|克)/i;
const SLIP_UNIT_SAME = { '台斤': '斤', 'kg': '公斤', 'KG': '公斤', 'Kg': '公斤', '克': 'g', 'G': 'g' };
const SLIP_CODE = /(?:^|\s)(?:\d{1,2}\s+)?([A-Za-z]{1,3}\d{1,5}(?:-\d{1,3})?)(?=\s)/;   // 貨品編號，例如 B01-2、A28、L06
const SLIP_TOTAL = /(合計|總計|總金額|金額合計|應收|小計|總額|本單)/;

function parseSlip_(text, keys) {
  const lines = String(text || '').split(/\r?\n/).map(function (l) {
    return l.replace(/[\uFF01-\uFF5E]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/\u3000/g, ' ').trim();
  }).filter(Boolean);
  const rows = {}, extras = [];
  let total = '';
  const nums = function (s) {
    return (s.match(/\d+(?:,\d{3})*(?:\.\d+)?/g) || []).map(function (x) { return Number(x.replace(/,/g, '')); });
  };
  const codes = keys.codes || {};
  const findKey = function (l) {
    const cm = l.match(SLIP_CODE), code = cm ? cm[1].toUpperCase() : '';
    if (code && codes[code]) {   // 以前存過的貨品編號：最準
      const nm = codes[code], rest = l.slice(l.indexOf(cm[1]) + cm[1].length).replace(/^\s*[^\d\s]+/, '');
      return { k: { key: '', name: nm }, code: code, rest: rest, partial: false };
    }
    const n = normKey_(l);
    for (let i = 0; i < keys.length; i++) {
      const at = n.indexOf(keys[i].key);
      if (at < 0) continue;
      // 品名前後還連著別的中文字（例如「二砂糖」找到「砂糖」）：可能是別的東西，標「請確認」
      // 括號裡的是規格（例如「高麗菜(初秋)」），不算
      const m = normKey_(l.replace(/[(（][^)）]*[)）]/g, ' ')), at2 = m.indexOf(keys[i].key), t = at2 >= 0 ? m : n, p = at2 >= 0 ? at2 : at;
      const cjk = /[\u4e00-\u9fff]/, before = t.charAt(p - 1), after = t.charAt(p + keys[i].key.length);
      return { k: keys[i], at: at, partial: cjk.test(before) || cjk.test(after), code: code };
    }
    return null;
  };
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (/單價/.test(l) && /(數量|品名)/.test(l)) continue;   // 表頭
    if (SLIP_TOTAL.test(l)) {
      let ns = nums(l.replace(/.*?(合計|總計|總金額|金額合計|應收|小計|總額|本單)/, ''));
      if (!ns.length && lines[i + 1] && !findKey(lines[i + 1])) ns = nums(lines[i + 1]);
      if (ns.length) total = ns[ns.length - 1];
      continue;
    }
    const hit = findKey(l);
    let rest = hit ? (hit.rest !== undefined ? hit.rest : afterKey_(l, hit.k.key)) : l;
    let ns = nums(rest);
    if (hit && !ns.length && lines[i + 1] && !findKey(lines[i + 1]) && !SLIP_TOTAL.test(lines[i + 1])) { rest = lines[i + 1]; ns = nums(rest); i++; }
    const u0 = (rest.match(SLIP_UNITS) || [])[1] || '', u = SLIP_UNIT_SAME[u0] || u0;
    const v = slipNumbers_(ns);
    if (hit) {
      const r = rows[hit.k.name];
      const slipName = slipNameOf_(l);
      if (!r) rows[hit.k.name] = { slipName: slipName, qty: v.qty, unit: u, price: v.price, amount: v.amount, sure: v.sure && !hit.partial, partial: !!hit.partial, code: hit.code || '' };
      else if (r.unit === u && v.qty !== '') {   // 同一樣分好幾行（例如分兩箱）：數量、金額加起來
        r.qty = round2_((Number(r.qty) || 0) + v.qty);
        if (v.amount !== '') r.amount = round2_((Number(r.amount) || 0) + v.amount);
      }
    } else if (ns.length >= 2 && /[\u4e00-\u9fff]{2,}/.test(l) && !/(日期|電話|傳真|統一編號|統編|地址|單號|編號|客戶|頁)/.test(l)) {
      const nm = slipNameOf_(l), cm = l.match(SLIP_CODE);
      if (nm) extras.push({ slipName: nm, qty: v.qty, unit: u, price: v.price, amount: v.amount, sure: v.sure, code: cm ? cm[1].toUpperCase() : '' });
    }
  }
  return { text: lines.slice(0, 300).join('\n'), rows: rows, extras: extras.slice(0, 40), total: total };
}

/** 一行裡的品名：去掉序號、貨品編號、數字、單位 */
function slipNameOf_(l) {
  return l.replace(SLIP_CODE, ' ').replace(/\d+(?:,\d{3})*(?:\.\d+)?/g, ' ').replace(new RegExp(SLIP_UNITS.source, 'gi'), ' ')
    .replace(/[:：|]/g, ' ').replace(/\s+/g, ' ').trim().split(' ')[0].slice(0, 40);
}

function afterKey_(line, key) {
  // 品名後面的字（品名前面常有貨號，不要算進數量）
  let acc = '';
  for (let i = 0; i < line.length; i++) {
    acc += line[i];
    if (normKey_(acc).indexOf(key) >= 0) return line.slice(i + 1);
  }
  return line;
}

/** 一行裡的數字 → 數量、單價、金額：有三個數字且「數量×單價≈金額」最準 */
function slipNumbers_(ns) {
  const out = { qty: '', price: '', amount: '', sure: false };
  if (!ns.length) return out;
  for (let a = 0; a < ns.length; a++) for (let b = a + 1; b < ns.length; b++) for (let c = b + 1; c < ns.length; c++) {
    if (ns[c] > 0 && Math.abs(ns[a] * ns[b] - ns[c]) <= Math.max(1, ns[c] * 0.02)) return { qty: ns[a], price: ns[b], amount: ns[c], sure: true };
  }
  out.qty = ns[0];
  if (ns.length >= 3) { out.price = ns[ns.length - 2]; out.amount = ns[ns.length - 1]; }
  else if (ns.length === 2) { out.amount = ns[1]; out.price = ns[0] ? round2_(ns[1] / ns[0]) : ''; }
  return out;
}

/** 單價表：每家廠商每樣食材（同一單位）留最新的單價，之後算成本用 */
function savePrices_(vendor, from, id, slip) {
  const list = slip.filter(function (x) { return x.name && Number(x.price) > 0; });
  if (!list.length) return;
  ensureSheet_(SHEET.prices, PRICE_HEADERS);
  const t = readTable_(SHEET.prices, ['廠商', '食材名稱']);
  const add = [];
  list.forEach(function (x) {
    const r = t.rows.find(function (y) { return String(y['廠商']).trim() === vendor && String(y['食材名稱']).trim() === x.name && String(y['單位'] || '').trim() === x.unit; });
    if (r) {
      if (purchaseDate_(r['更新日期']) > from) return;   // 已經有更新的價格
      setCell_(t, r._row, '單價', x.price); setCell_(t, r._row, '更新日期', from); setCell_(t, r._row, '驗收編號', id);
      if (x.slipName) setCell_(t, r._row, '銷貨單品名', x.slipName);
      if (x.code && t.cols['貨品編號'] !== undefined) setCell_(t, r._row, '貨品編號', x.code);
    } else add.push({ '廠商': vendor, '食材名稱': x.name, '單位': x.unit, '單價': x.price, '銷貨單品名': x.slipName, '更新日期': from, '驗收編號': id, '貨品編號': x.code });
  });
  if (add.length) appendRows_(SHEET.prices, add);
}

function num_(v) {
  const s = String(v === undefined || v === null ? '' : v).replace(/[,，\s]/g, '');
  if (s === '') return '';
  const n = Number(s);
  if (!isFinite(n) || n < 0) throw new Error('金額或數量「' + v + '」不是數字。');
  return round2_(n);
}
function round2_(n) { return Math.round(Number(n) * 100) / 100; }

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

const GRAMS = { 'g': 1, '公斤': 1000, '斤': 600, '兩': 37.5, 'T': 15, 't': 5 };   // 1 大T＝15 g、1 小T＝5 g（跟訂貨系統一樣）
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
  } else if (sh.getLastColumn()) {
    // 舊的分頁缺新欄位：補在最後面
    const have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    const miss = headers.filter(function (h) { return have.indexOf(h) < 0; });
    if (miss.length) sh.getRange(1, have.length + 1, 1, miss.length).setValues([miss]).setFontWeight('bold');
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
