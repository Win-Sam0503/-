/**
 * 三信家商餐飲科 進貨驗收（給管理者用平板點貨）
 *
 * © 2026 翁煒程 版權所有。
 * 本系統之程式碼、介面設計與內容均受中華民國著作權法保護，僅供三信家商餐飲科教職員內部使用。
 * 未經著作權人書面同意，禁止以任何形式重製、拷貝、修改、轉載或散佈本系統之程式碼、畫面及網址，
 * 亦不得提供予校外人士使用。違者將依法追究相關法律責任。
 *
 * 跟訂貨系統共用同一份試算表：讀「叫貨紀錄」（③ 叫貨單，已確定的），驗收結果寫進「驗收紀錄」「驗收明細」分頁。
 * 部署成網頁應用程式：執行身分「我」。打開後要用訂貨系統的「管理者」帳號、密碼登入（跟訂貨系統同一組密碼）。
 */

const SPREADSHEET_ID = '1qIQe4STafrhzQk4I7NLkZQxJGe5hozf0g04zvJ-qZ_I';   // 訂貨系統的試算表
const SITE_TITLE = '三信家商餐飲科 進貨驗收';
const PHOTO_FOLDER_NAME = '三信家商餐飲科進貨銷貨單';   // 銷貨單照片放在雲端硬碟這個資料夾（每週一個子資料夾）
const SHEET = { purchase: '叫貨紀錄', checks: '驗收紀錄', lines: '驗收明細', slip: '銷貨單明細', prices: '單價表', items: '食材總表', drafts: '驗收暫存' };
const DRAFT_HEADERS = ['驗收編號', '暫存時間', '內容'];
const SLIP_REC = '銷貨單紀錄', SLIP_REC_HEADERS = ['驗收編號', '送貨日期（起）', '廠商', '送貨組別', '銷貨單照片', '銷貨單金額', '銷貨單合計', '備註', '更新時間', '更新人'];
const CHECK_HEADERS = ['驗收編號', '送貨日期（起）', '送貨日期（迄）', '廠商', '送貨組別', '驗收時間', '品項數', '問題數', '銷貨單照片', '備註', '驗收人', '銷貨單金額', '銷貨單合計', '簽名人', '驗收人簽名'];
const SLIP_HEADERS = ['驗收編號', '送貨日期（起）', '廠商', '食材名稱', '銷貨單品名', '數量', '單位', '單價', '金額', '貨品編號', '第幾張', '銷貨單日期', '類型'];   // 類型：食材／器材（每張銷貨單選一類）   // 第幾張：照紙本一張一張打的
const PRICE_HEADERS = ['廠商', '食材名稱', '單位', '單價', '銷貨單品名', '更新日期', '驗收編號', '貨品編號', '換算'];   // 換算：1 個「單位」等於多少，例如「190 顆」「1000 g」
const LINE_HEADERS = ['驗收編號', '廠商', '送貨組別', '班級', '食材名稱', '應到', '實到', '狀態', '備註', '處理狀態', '處理時間', '處理備註', '照片'];
// 照班級點貨的廠商（其他廠商點品項總量）
const PER_CLASS = /^[ABD]/;
const ISSUE_STATES = ['待處理', '已補送', '換貨', '退貨', '扣款', '不用補'];
const CHECK_STATUS = ['到齊', '少送', '沒送', '送錯', '多送', '品質不好', '缺貨', '其他', '停送'];   // 停送：颱風假、停課，廠商沒送也不用送（不算問題、不算錢）
// 同一家廠商、送兩個地方：訂貨系統裡分開記（C烘焙蛋／C中餐蛋），驗收時合成一張，分「送烘焙教室」「送中餐教室」兩段
const SITE_VENDORS = { 'C烘焙蛋': { base: 'C蛋商', site: '送烘焙教室' }, 'C中餐蛋': { base: 'C蛋商', site: '送中餐教室' } };

// ---------------------------------------------------------------------
// 登入：用訂貨系統的帳號密碼（「帳號設定」分頁），只有管理者可以進來
// ---------------------------------------------------------------------
const ADMIN_ACCOUNTS = ['管理者'];          // 跟訂貨系統一樣：內建管理者＋「帳號設定」管理者欄打勾的
const RCV_SESSION_HOURS = 12;               // 登入後幾小時內不用再登入（平板放著隔天要重登）
const RCV_MAX_FAILS = 5, RCV_LOCK_MINUTES = 15;
let CURRENT_USER_ = '';                      // 這次呼叫是誰（驗收人、浮水印用）

function rcvLogin(account, password) {
  account = clean_(account, 30);
  password = String(password == null ? '' : password);
  if (!account || !password) throw new Error('請輸入帳號和密碼。');
  const cache = CacheService.getScriptCache(), failKey = 'rfail:' + account;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= RCV_MAX_FAILS) throw new Error('密碼錯誤次數過多，請 ' + RCV_LOCK_MINUTES + ' 分鐘後再試。');
  const t = readTable_('帳號設定', ['帳號', '密碼', '啟用']);
  const acc = t.rows.find(function (r) { return String(r['帳號']).trim() === account; });
  const ok = acc && isTrue_(acc['啟用']) && String(acc['密碼']) === password;
  const admin = ok && (ADMIN_ACCOUNTS.indexOf(account) >= 0 || (t.cols['管理者'] !== undefined && isTrue_(acc['管理者'])));
  rcvLog_(account, ok ? (admin ? '成功（驗收系統）' : '不是管理者（驗收系統）') : '失敗（驗收系統）');
  if (!ok) {
    cache.put(failKey, String(fails + 1), RCV_LOCK_MINUTES * 60);
    throw new Error('帳號或密碼不正確（跟訂貨系統同一組，有分大小寫）。');
  }
  if (!admin) throw new Error('驗收系統只有管理者可以登入。');
  cache.remove(failKey);
  const token = Utilities.getUuid();
  cache.put('rsess:' + token, account, RCV_SESSION_HOURS * 3600);
  return { token: token, account: account };
}
function rcvLogout(token) { CacheService.getScriptCache().remove('rsess:' + String(token || '')); return true; }
function rcvLog_(account, result) {
  try { appendRows_('登入資訊', [{ '登入時間': new Date(), '帳號': account, '結果': result }]); } catch (e) { console.error(e.message); }
}
function isTrue_(v) { return v === true || String(v).trim().toUpperCase() === 'TRUE' || String(v).trim() === '是'; }   // 跟訂貨系統一樣

/** 網頁呼叫伺服器一律經過這裡：先檢查登入，再呼叫對應的功能（功能名稱後面有 _，網頁不能直接呼叫） */
const API_ = {
  getWeek: function (a) { return getWeek_.apply(null, a); }, saveCheck: function (a) { return saveCheck_.apply(null, a); },
  saveSlip: function (a) { return saveSlip_.apply(null, a); }, deleteSlip: function (a) { return deleteSlip_.apply(null, a); },
  uploadPhoto: function (a) { return uploadPhoto_.apply(null, a); }, readSlip: function (a) { return readSlip_.apply(null, a); },
  getMonth: function (a) { return getMonth_.apply(null, a); }, saveMonth: function (a) { return saveMonth_.apply(null, a); },
  saveDraftRemote: function (a) { return saveDraftRemote_.apply(null, a); }, dropDraftRemote: function (a) { return dropDraftRemote_.apply(null, a); },
  getIssues: function (a) { return getIssues_.apply(null, a); }, setIssues: function (a) { return setIssues_.apply(null, a); },
  whoAmI: function () { return CURRENT_USER_; },
  hxGet: function (a) { return hxGet_.apply(null, a); }, hxCandidates: function (a) { return hxCandidates_.apply(null, a); },
  hxSaveInvoice: function (a) { return hxSaveInvoice_.apply(null, a); }, hxDeleteInvoice: function (a) { return hxDeleteInvoice_.apply(null, a); },
  hxSetDone: function (a) { return hxSetDone_.apply(null, a); }, hxStats: function (a) { return hxStats_.apply(null, a); },
  hxSaveCats: function (a) { return hxSaveCats_.apply(null, a); },
  hxBudgets: function () { return hxBudgets_(); }, hxSaveBudget: function (a) { return hxSaveBudgets_([a[0]]); },
  hxSaveBudgets: function (a) { return hxSaveBudgets_.apply(null, a); },
  hxItems: function () { return hxItems_(); }, hxPrintData: function (a) { return hxPrintData_.apply(null, a); }, hxOrders: function (a) { return hxOrders_.apply(null, a); },
  hxDeleteBudget: function (a) { return hxDeleteBudget_.apply(null, a); },
};
function api(token, fn, args) {
  const account = CacheService.getScriptCache().get('rsess:' + String(token || ''));
  if (!account) throw new Error('請重新登入。');
  if (!API_[fn]) throw new Error('沒有這個功能：' + fn);
  CURRENT_USER_ = account;
  return API_[fn](args || []);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(SITE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------
// 讀這一週要驗收的叫貨單
// ---------------------------------------------------------------------

/** monday：'yyyy-MM-dd'（那週星期一）。回傳每一張 ③ 叫貨單（廠商＋送貨組別）和驗收狀態 */
function getWeek_(monday) {
  const from = parseDate_(monday) || mondayOf_(today_());
  const to = addDays_(from, 4);
  const f = dateStr_(from), t = dateStr_(to);
  const recs = readSlipRecs_();
  const out = weekBatches_(purchaseRows_(f, t), f, readChecks_(f), readPrices_(), recs);
  out.extras = extraSlips_(recs, f, t);
  const drafts = readDrafts_();
  out.batches.forEach(function (b) { delete b.list; b.draft = drafts[b.id] || null; });
  return out;
}

/** 叫貨紀錄裡送貨日期在 f～t（'yyyy-MM-dd'）的 */
function purchaseRows_(f, t) {
  return readTable_(SHEET.purchase, ['送貨日期', '廠商', '食材名稱']).rows.map(function (r) {
    return {
      date: purchaseDate_(r['送貨日期']), group: String(r['送貨組別'] || '').trim(), vendor: String(r['廠商'] || '').trim(),
      className: String(r['上課班級'] || '').trim(), teacher: String(r['老師'] || '').trim(), name: String(r['食材名稱'] || '').trim(),
      qty: Number(r['叫貨量']) || 0, unit: String(r['叫貨單位'] || '').trim(), kind: String(r['類型'] || '訂單').trim(),
      note: String(r['備註'] || '').trim(), status: String(r['狀態'] || '').trim(),
    };
  }).filter(function (r) { return r.date >= f && r.date <= t && r.vendor && r.name; });
}

/** 一週（f＝星期一）的叫貨紀錄 → 一張張 ③ 叫貨單（b.list 是原始的叫貨紀錄） */
function weekBatches_(rows, f, checked, prices, recs) {
  const t = dateStr_(addDays_(parseDate_(f), 4));
  rows = rows.filter(function (r) { return r.date >= f && r.date <= t; });
  const drafts = {}, by = {};
  rows.forEach(function (r) {
    const sv = SITE_VENDORS[r.vendor];
    if (r.status === '草稿') { drafts[sv ? sv.base : r.vendor] = true; return; }
    const k = sv ? sv.base + '|' : r.vendor + '|' + r.group;
    (by[k] = by[k] || []).push(r);
  });
  const batches = Object.keys(by).sort().map(function (k) {
    const list = by[k], sv = SITE_VENDORS[list[0].vendor];
    const vendor = sv ? sv.base : list[0].vendor, group = sv ? '' : list[0].group;
    const id = checkId_(f, vendor, group);
    const b = { id: id, vendor: vendor, group: group, label: vendor + (group ? '（送貨組別 ' + group + '，星期' + group.split('').join('、') + '）' : ''),
      perClass: PER_CLASS.test(vendor), lines: sv ? siteLines_(list) : batchLines_(list, vendor), check: checked[id] || null,
      totals: nameTotals_(list), prices: prices[vendor] || {}, list: list };
    // 銷貨單（照片＋金額）：「銷貨單」分頁存的優先；舊資料是跟驗收一起存的
    const rec = recs && recs[id], c = b.check;
    b.slip = rec ? { rows: (c && c.slip ? c.slip.rows : null) || slipsFor_(id), total: rec.total, amount: rec.amount, photos: rec.photos, note: rec.note, time: rec.time, by: rec.by }
      : c && ((c.slip && c.slip.rows.length) || c.photos.length) ? { rows: c.slip ? c.slip.rows : [], total: c.slip ? c.slip.total : '', amount: c.amount, photos: c.photos, note: '', time: c.time, by: '', legacy: true } : null;
    return b;
  });
  return { from: f, to: t, batches: batches, drafts: Object.keys(drafts).filter(function (v) { return !batches.some(function (b) { return b.vendor === v; }); }) };
}

/** 每樣食材這張叫貨單一共叫多少（同一樣有好幾種單位就不算）：{ 食材名稱: { qty, unit } }，給銷貨單金額「手動填」先帶數量 */
function nameTotals_(list) {
  const tot = {};
  const byVendor = {};
  list.forEach(function (r) { (byVendor[r.vendor] = byVendor[r.vendor] || []).push(r); });
  Object.keys(byVendor).forEach(function (v) {
    const rows = byVendor[v], ovr = {};
    rows.forEach(function (r) { if (r.kind === '合計') ovr[r.name] = r; });
    const add = function (n, qty, unit) {
      const cv = convertUnit_(qty, unit, v, n), t = tot[n];
      if (!t) tot[n] = { qty: cv.qty, unit: cv.unit };
      else if (t.unit === cv.unit) t.qty = Math.round((t.qty + cv.qty) * 100) / 100;
      else t.mixed = true;
    };
    rows.forEach(function (r) { if (r.kind !== '合計' && r.qty > 0 && !ovr[r.name]) add(r.name, r.qty, r.unit); });
    Object.keys(ovr).forEach(function (n) { if (ovr[n].qty > 0) add(n, ovr[n].qty, ovr[n].unit); });
  });
  Object.keys(tot).forEach(function (n) { if (tot[n].mixed) delete tot[n]; });
  return tot;
}

/** 單價表：{ 廠商: { 食材名稱: { price, unit, date } } }（同一樣有好幾種單位時取最新的） */
function readPrices_() {
  const out = {};
  if (!ss_().getSheetByName(SHEET.prices)) return out;
  readTable_(SHEET.prices, ['廠商', '食材名稱', '單價']).rows.forEach(function (r) {
    const v = String(r['廠商']).trim(), n = String(r['食材名稱']).trim(), pr = Number(r['單價']);
    if (!v || !n || !(pr > 0)) return;
    const d = purchaseDate_(r['更新日期']), m = out[v] = out[v] || {};
    const cv = parseQty_(r['換算']);
    if (!m[n] || d >= m[n].date) m[n] = { price: pr, unit: String(r['單位'] || '').trim(), date: d, conv: cv && cv.qty > 0 ? cv : null };
  });
  return out;
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
      return { key: r.date + '|' + cls + '|' + r.teacher + '|' + r.name + '|' + r.unit, cls: cls, teacher: r.kind === '補庫存' ? '' : r.teacher, date: r.date, name: r.name,
        expect: show, note: r.note, ovr: ovr[r.name] ? '合計改成 ' + (ovr[r.name].qty > 0 ? fmtQty_(ovr[r.name].qty) + ' ' + ovr[r.name].unit : '不叫') : '' };
    });
    // 被直接改成「不叫」的品項不用點
    // 同班同一樣食材不同天（或同一天叫兩次）要各點各的：key 帶日期、老師，完全一樣的再加 #2、#3
    const seen = {};
    out.forEach(function (x) { const n = seen[x.key] = (seen[x.key] || 0) + 1; if (n > 1) x.key += '#' + n; });
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
      amount: Number(r['銷貨單金額']) || 0, signer: String(r['簽名人'] || ''), signature: String(r['驗收人簽名'] || ''), slip: slips[id] ? { rows: slips[id], total: r['銷貨單合計'] === '' || r['銷貨單合計'] === undefined ? '' : String(r['銷貨單合計']) } : null };
  });
  return out;
}

let SLIPS_MEMO_ = null;
function slipsFor_(id) { if (!SLIPS_MEMO_) SLIPS_MEMO_ = slipsById_(); return SLIPS_MEMO_[id] || []; }

function slipsById_() {
  const out = {};
  const sh = ss_().getSheetByName(SHEET.slip);
  if (!sh || sh.getLastRow() < 2) return out;
  readTable_(SHEET.slip, ['驗收編號']).rows.forEach(function (r) {
    const id = String(r['驗收編號']).trim();
    (out[id] = out[id] || []).push({ name: String(r['食材名稱'] || ''), slipName: String(r['銷貨單品名'] || ''), qty: String(r['數量'] === undefined ? '' : r['數量']),
      unit: String(r['單位'] || ''), price: String(r['單價'] === undefined ? '' : r['單價']), amount: String(r['金額'] === undefined ? '' : r['金額']),
      code: String(r['貨品編號'] || ''), sheet: Number(r['第幾張']) || 0, date: r['銷貨單日期'] instanceof Date ? dateStr_(r['銷貨單日期']) : String(r['銷貨單日期'] || '').replace(/^'/, ''),
      kind: String(r['類型'] || '').trim() === '器材' ? '器材' : '食材' });
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
      got: String(r['實到'] || ''), status: String(r['狀態'] || ''), note: String(r['備註'] || ''), photos: photoList_(r['照片']) });
  });
  return out;
}

function checkId_(from, vendor, group) { return 'R' + from.replace(/-/g, '') + '-' + vendor + (group ? '-' + group : ''); }

// ---------------------------------------------------------------------
// 存驗收結果
// ---------------------------------------------------------------------

/** 銷貨單紀錄：{ 驗收編號: { photos, amount, total, note, time, by } } */
function readSlipRecs_() {
  const out = {};
  if (!ss_().getSheetByName(SLIP_REC)) return out;
  readTable_(SLIP_REC, ['驗收編號']).rows.forEach(function (r) {
    const id = String(r['驗收編號']).trim();
    if (id) out[id] = { id: id, date: purchaseDate_(r['送貨日期（起）']), vendor: String(r['廠商'] || '').trim(), group: String(r['送貨組別'] || '').trim(),
      photos: String(r['銷貨單照片'] || '').split(/\n+/).filter(Boolean), amount: Number(r['銷貨單金額']) || 0,
      total: r['銷貨單合計'] === '' || r['銷貨單合計'] === undefined ? '' : String(r['銷貨單合計']), note: String(r['備註'] || ''), time: fmtDateTime_(r['更新時間']), by: String(r['更新人'] || '') };
  });
  return out;
}

/**
 * 「💰 銷貨單」分頁：存銷貨單照片和金額（跟驗收點貨分開，還沒驗收也可以先存）。
 * p = { from, vendor, group, photos: [url], note, slip: { total, rows: [{ name, slipName, qty, unit, price, amount, code, noPrice }] } }
 */
function saveSlip_(p) {
  // 臨時銷貨單（不在叫貨單裡）：編號 S 開頭，日期是送貨那天
  const extra = !!p.extra, vendor = clean_(p.vendor, 30), group = extra ? '' : clean_(p.group, 10);
  const from = dateStr_(parseDate_(extra ? p.date : p.from));
  if (!vendor) throw new Error('請填廠商（不在系統裡的商店也可以直接打名字）。');
  if (!from) throw new Error(extra ? '請選送貨日期。' : '資料不完整，請重新整理。');
  const id = extra ? (/^S\d{8}-/.test(String(p.id || '')) ? clean_(p.id, 60) : 'S' + from.replace(/-/g, '') + '-' + vendor + '-' + Utilities.formatDate(new Date(), tz_(), 'HHmmss')) : checkId_(from, vendor, group);
  const photos = (p.photos || []).filter(function (u) { return /^https:\/\/(drive|docs)\.google\.com\//.test(String(u)); }).slice(0, 30);
  const slip = (p.slip && p.slip.rows || []).map(function (x) {
    const q4 = num_(x.qty) === '' ? '' : Math.round(Number(String(x.qty).replace(/[,，\s]/g, '')) * 10000) / 10000;   // 斤兩換成斤會有 4 位小數（3 兩＝0.1875 斤）
    return { name: clean_(x.name, 40), slipName: clean_(x.slipName, 60), qty: q4, unit: clean_(x.unit, 10), price: num_(x.price), amount: num_(x.amount),
      code: normCode_(clean_(x.code, 20)), noPrice: !!x.noPrice, sheet: Number(x.sheet) || '', date: parseDate_(x.date) ? dateStr_(parseDate_(x.date)) : '', kind: x.kind === '器材' ? '器材' : '食材' };
  }).filter(function (x) { return (x.name || x.slipName) && (x.qty !== '' || x.price !== '' || x.amount !== ''); });
  const sum = Math.round(slip.reduce(function (a, x) { return a + (Number(x.amount) || 0); }, 0) * 100) / 100;
  const total = p.slip ? num_(p.slip.total) : '';
  ensureSheet_(SLIP_REC, SLIP_REC_HEADERS);
  ensureSheet_(SHEET.slip, SLIP_HEADERS);
  const lock = waitLock_();
  try {
    const rt = readTable_(SLIP_REC, ['驗收編號']);
    const old = rt.rows.filter(function (r) { return String(r['驗收編號']).trim() === id; });
    if (old.length) deleteRows_(rt.sheet, old.map(function (r) { return r._row; }));
    appendRows_(SLIP_REC, [{ '驗收編號': id, '送貨日期（起）': from, '廠商': vendor, '送貨組別': group, '銷貨單照片': photos.join('\n'),
      '銷貨單金額': slip.length ? sum : total !== '' ? Number(total) : '', '銷貨單合計': total, '備註': clean_(p.note, 200), '更新時間': new Date(), '更新人': CURRENT_USER_ || '' }]);
    const st = readTable_(SHEET.slip, ['驗收編號']);
    const oldSlip = st.rows.filter(function (r) { return String(r['驗收編號']).trim() === id; });
    if (oldSlip.length) deleteRows_(st.sheet, oldSlip.map(function (r) { return r._row; }));
    if (slip.length) {
      appendRows_(SHEET.slip, slip.map(function (x) {
        return { '驗收編號': id, '送貨日期（起）': from, '廠商': vendor, '食材名稱': x.name, '銷貨單品名': x.slipName, '數量': x.qty, '單位': x.unit, '單價': x.price, '金額': x.amount, '貨品編號': x.code,
          '第幾張': x.sheet, '銷貨單日期': x.date ? "'" + x.date : '', '類型': x.kind };
      }));
      savePrices_(vendor, from, id, slip);
    }
  } finally {
    lock.releaseLock();
  }
  return { id: id, amount: slip.length ? sum : total !== '' ? Number(total) : '', rows: slip.length, photos: photos.length };   // 只填合計：金額＝合計
}

/** 刪掉一張臨時銷貨單（叫貨單上的不能刪，只能清空） */
function deleteSlip_(id) {
  id = clean_(id, 60);
  if (!/^S\d{8}-/.test(id)) throw new Error('只有「臨時銷貨單」可以刪除。');
  const lock = waitLock_();
  try {
    [SLIP_REC, SHEET.slip].forEach(function (n) {
      if (!ss_().getSheetByName(n)) return;
      const t = readTable_(n, ['驗收編號']);
      deleteRows_(t.sheet, t.rows.filter(function (r) { return String(r['驗收編號']).trim() === id; }).map(function (r) { return r._row; }));
    });
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** 臨時銷貨單（S 開頭），送貨日期在 f～t：[{ id, vendor, date, slip }] */
function extraSlips_(recs, f, t, vendor) {
  return Object.keys(recs).filter(function (id) { const r = recs[id]; return /^S\d{8}-/.test(id) && r.date >= f && r.date <= t && (!vendor || r.vendor === vendor); })
    .map(function (id) { const r = recs[id]; return { id: id, vendor: r.vendor, date: r.date, slip: { rows: slipsFor_(id), total: r.total, amount: r.amount, photos: r.photos, note: r.note, time: r.time, by: r.by } }; })
    .sort(function (a, b) { return a.date.localeCompare(b.date) || a.vendor.localeCompare(b.vendor, 'zh-Hant'); });
}

/**
 * p = { from, to, vendor, group, note, photos: [url], items: [{ cls, name, expect, got, status, note }],
 *       slip: { total, rows: [{ name（食材名稱，空白＝叫貨單沒有的）, slipName, qty, unit, price, amount }] }（選填，銷貨單金額） }
 * status：到齊／少送／沒送／送錯／多送。同一張叫貨單重新驗收會蓋掉上次的（已經處理過的問題保留處理狀態）。
 */
function saveCheck_(p) {
  const from = dateStr_(parseDate_(p.from)), to = dateStr_(parseDate_(p.to));
  const vendor = clean_(p.vendor, 30), group = clean_(p.group, 10);
  if (!from || !vendor) throw new Error('資料不完整，請重新整理。');
  const items = (p.items || []).map(function (x) {
    const st = CHECK_STATUS.indexOf(x.status) >= 0 ? x.status : '';
    if (!st) throw new Error('「' + x.name + '」還沒點：請按 ✓ 或標記少送、沒送、送錯。');
    if ((st === '少送' || st === '多送') && !String(x.got || '').trim()) throw new Error('「' + x.name + '」' + st + '，請填實際送來多少。');
    if (st === '其他' && !String(x.note || '').trim()) throw new Error('「' + x.name + '」選了「其他」，請在備註寫是什麼狀況。');
    const got = st === '到齊' ? clean_(x.expect, 30) : st === '沒送' || st === '缺貨' || st === '停送' ? '0'
      : st === '品質不好' || st === '其他' ? clean_(x.got, 30) || clean_(x.expect, 30) : clean_(x.got, 30);
    return { cls: clean_(x.cls, 30), name: clean_(x.name, 40), expect: clean_(x.expect, 30), got: got, status: st, note: clean_(x.note, 100),
      photos: (x.photos || []).filter(function (u) { return /^https:\/\/(drive|docs)\.google\.com\//.test(String(u)); }).slice(0, 6) };
  });
  const id = checkId_(from, vendor, group);
  const issues = items.filter(function (x) { return x.status !== '到齊' && x.status !== '停送'; });
  const allStopped = items.length && items.every(function (x) { return x.status === '停送'; });
  // 驗收人簽名：新簽的存成圖片放雲端硬碟；沒重簽就沿用上次的
  const signer = clean_(p.signer, 20);
  let sigUrl = /^https:\/\/(drive|docs)\.google\.com\//.test(String(p.signatureUrl || '')) ? String(p.signatureUrl) : '';
  if (p.signature) sigUrl = saveImage_(p.signature, 'image/png', from, vendor, '驗收簽名' + (signer ? ' ' + signer : ''), 'png').url;
  if (!sigUrl && !allStopped) throw new Error('請在「驗收人簽名」簽名。');   // 整張停送（沒有貨）不用簽
  const slip = (p.slip && p.slip.rows || []).map(function (x) {
    return { name: clean_(x.name, 40), slipName: clean_(x.slipName, 60), qty: num_(x.qty), unit: clean_(x.unit, 10), price: num_(x.price), amount: num_(x.amount),
      code: normCode_(clean_(x.code, 20)), noPrice: !!x.noPrice };   // noPrice：這一行含別的品項（廠商開在一起），單價不準，不進單價表
  }).filter(function (x) { return (x.name || x.slipName) && (x.qty !== '' || x.price !== '' || x.amount !== ''); });
  const slipSum = slip.reduce(function (a, x) { return a + (Number(x.amount) || 0); }, 0);
  const slipTotal = p.slip ? num_(p.slip.total) : '';
  const withSlip = p.slip !== undefined || p.photos !== undefined;   // 新版：銷貨單金額、照片在「銷貨單」分頁另外存，驗收不動它們
  const lock = waitLock_();
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
    const old = oldChecks[0] || {};
    if (oldChecks.length) deleteRows_(ct.sheet, oldChecks.map(function (r) { return r._row; }));
    appendRows_(SHEET.checks, [{ '驗收編號': id, '送貨日期（起）': from, '送貨日期（迄）': to, '廠商': vendor, '送貨組別': group,
      '驗收時間': new Date(), '品項數': items.length, '問題數': issues.length, '銷貨單照片': withSlip ? (p.photos || []).join('\n') : old['銷貨單照片'] || '',
      '備註': clean_(p.note, 200), '驗收人': CURRENT_USER_ || Session.getActiveUser().getEmail() || '',
      '銷貨單金額': withSlip ? (slip.length ? Math.round(slipSum * 100) / 100 : '') : old['銷貨單金額'] === undefined ? '' : old['銷貨單金額'],
      '銷貨單合計': withSlip ? slipTotal : old['銷貨單合計'] === undefined ? '' : old['銷貨單合計'], '簽名人': signer, '驗收人簽名': sigUrl }]);
    dropDraft_(id);
    if (withSlip) {
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
    }
    appendRows_(SHEET.lines, items.map(function (x) {
      const k = keep[x.cls + '|' + x.name];
      const issue = x.status !== '到齊';
      return { '驗收編號': id, '廠商': vendor, '送貨組別': group, '班級': x.cls, '食材名稱': x.name, '應到': x.expect,
        '實到': x.got, '狀態': x.status, '備註': x.note, '照片': x.photos.join('\n'),
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
          htmlBody: '<p>' + escHtml_(vendor + (group ? ' 送貨組別 ' + group : '') + '，' + rangeText_(from, to) + (signer ? '（驗收人：' + signer + '）' : '')) + '：</p><ul>' +
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
  if (x.status === '缺貨') return '廠商缺貨（應到 ' + x.expect + '）' + (x.note ? '，' + x.note : '');
  if (x.status === '其他') return '其他：' + x.note + (x.got !== x.expect ? '（實到 ' + x.got + '）' : '');
  if (x.status === '品質不好') return '品質不好（' + x.got + '）' + (x.note ? '，' + x.note : '') + (x.photos && x.photos.length ? '（有拍照）' : '');
  return x.status + '：應到 ' + x.expect + '，實到 ' + x.got + (x.note ? '，' + x.note : '');
}

// ---------------------------------------------------------------------
// 銷貨單照片
// ---------------------------------------------------------------------

/** data：base64（不含 data: 開頭），回傳 { url, name } */
function uploadPhoto_(data, mime, from, vendor, what) {
  if (!data) throw new Error('沒有照片。');
  return saveImage_(data, mime || 'image/jpeg', from, vendor, clean_(what, 40) || '銷貨單', 'jpg');
}

function photoList_(v) { return String(v || '').split(/\n+/).filter(Boolean); }

/** 圖片存到雲端硬碟「進貨銷貨單」資料夾的那週子資料夾 */
function saveImage_(data, mime, from, vendor, what, ext) {
  const week = dateStr_(parseDate_(from) || today_());
  const sub = weekFolder_(week);
  const name = week + ' ' + clean_(vendor, 30) + ' ' + what + ' ' + Utilities.formatDate(new Date(), tz_(), 'HHmmss') + '.' + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(String(data).replace(/^data:[^,]*,/, '')), mime, name);
  const f = sub.createFile(blob);
  return { url: f.getUrl(), name: name, id: f.getId() };
}

/** 那週的照片資料夾：記住資料夾編號，不用每張照片都在整個雲端硬碟找一次（找資料夾很慢） */
function weekFolder_(week) {
  const cache = CacheService.getScriptCache(), key = 'phfolder:' + week;
  const id = cache.get(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const props = PropertiesService.getScriptProperties();
  let root = null;
  const rid = props.getProperty('PHOTO_ROOT_ID');
  if (rid) { try { root = DriveApp.getFolderById(rid); if (root.isTrashed()) root = null; } catch (e) { root = null; } }
  if (!root) { root = folderNamed_(PHOTO_FOLDER_NAME, null); props.setProperty('PHOTO_ROOT_ID', root.getId()); }
  const sub = folderNamed_(week + ' 那週', root);
  cache.put(key, sub.getId(), 6 * 3600);
  return sub;
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
function readSlip_(photos, vendor, names) {
  const ids = (photos || []).map(fileIdOf_).filter(Boolean);
  if (!ids.length) throw new Error('請先拍銷貨單。');
  const text = ocrTexts_(ids.slice(0, 6)).join('\n');
  return parseSlip_(text, slipKeys_(clean_(vendor, 30), (names || []).map(function (n) { return clean_(n, 40); }).filter(Boolean)));
}

function fileIdOf_(url) {
  const m = String(url || '').match(/\/d\/([\w-]{20,})/) || String(url || '').match(/[?&]id=([\w-]{20,})/);
  return m ? m[1] : '';
}

/** 照片 → 文字：轉成 Google 文件（會做中文文字辨識）、匯出純文字，再把暫存的文件丟到垃圾桶。好幾張同時送，不用一張一張排隊 */
function ocrTexts_(fileIds) {
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const reqs = fileIds.map(function (fileId, i) {
    const blob = DriveApp.getFileById(fileId).getBlob();
    const boundary = 'slip' + Date.now() + i;
    const meta = JSON.stringify({ name: '[銷貨單辨識暫存]', mimeType: 'application/vnd.google-apps.document' });
    const head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
      '\r\n--' + boundary + '\r\nContent-Type: ' + (blob.getContentType() || 'image/jpeg') + '\r\n\r\n').getBytes();
    const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
    return { url: 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id&ocrLanguage=zh-TW', method: 'post',
      contentType: 'multipart/related; boundary=' + boundary, payload: head.concat(blob.getBytes(), tail), headers: auth, muteHttpExceptions: true };
  });
  let res = UrlFetchApp.fetchAll(reqs);
  // 失敗的（Google 忙碌）等一下再送一次
  const again = [];
  res.forEach(function (r, i) { if (r.getResponseCode() >= 300) again.push(i); });
  if (again.length) {
    Utilities.sleep(1500);
    const r2 = UrlFetchApp.fetchAll(again.map(function (i) { return reqs[i]; }));
    again.forEach(function (i, k) { res[i] = r2[k]; });
  }
  const docIds = res.map(function (r) { return r.getResponseCode() < 300 ? JSON.parse(r.getContentText()).id : ''; });
  if (!docIds.some(Boolean)) throw new Error('銷貨單辨識失敗（' + res[0].getResponseCode() + '），請再試一次或手動填。');
  try {
    const ok = docIds.filter(Boolean);
    const outs = UrlFetchApp.fetchAll(ok.map(function (id) { return { url: 'https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text/plain', headers: auth, muteHttpExceptions: true }; }));
    return outs.map(function (o) { return o.getResponseCode() < 300 ? o.getContentText('UTF-8') : ''; });
  } finally {
    docIds.filter(Boolean).forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
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
    const c = normCode_(r['貨品編號']);
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
const SLIP_CODE = /(?:^|\s)(?:\d{1,2}\s+)?([A-Za-z]{1,3}\d{1,5}(?:[-~～]\d{1,3}){0,2})(?=\s)/;   // 貨品編號，例如 B01-2、A28、L06、A116~02、A005-9-3
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
  // 第 qi 個數字的單位：緊接在後面的（「0.1 kg 650 65」），沒有就取緊接在前面的（旺來興「袋 2 590」）；品名規格裡的「22kg」「1.66斤」不會被當成數量的單位
  const unitNear = function (s, qi) {
    const re = /\d+(?:,\d{3})*(?:\.\d+)?/g, after = new RegExp('^\\s*' + SLIP_UNITS.source, 'i'), before = new RegExp(SLIP_UNITS.source + '\\s*$', 'i');
    let m, k = 0;
    while ((m = re.exec(s))) {
      if (k++ !== qi) continue;
      const a = s.slice(m.index + m[0].length).match(after);
      if (a) return a[1];
      const b = s.slice(0, m.index).match(before);
      return b ? b[1] : '';
    }
    return '';
  };
  const codes = keys.codes || {};
  const findKey = function (l) {
    const cm = l.match(SLIP_CODE), code = cm ? normCode_(cm[1]) : '';
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
    const v = slipNumbers_(ns);
    const u0 = (v.qi !== undefined && unitNear(rest, v.qi)) || (rest.match(SLIP_UNITS) || [])[1] || '', u = SLIP_UNIT_SAME[u0] || SLIP_UNIT_SAME[u0.toLowerCase()] || u0;
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
      if (nm) extras.push({ slipName: nm, qty: v.qty, unit: u, price: v.price, amount: v.amount, sure: v.sure, code: cm ? normCode_(cm[1]) : '' });
    }
  }
  return { text: lines.slice(0, 300).join('\n'), rows: rows, extras: extras.slice(0, 40), total: total };
}

/** 貨品編號統一寫法：大寫，「~」「～」都當成「-」 */
function normCode_(c) { return String(c || '').trim().toUpperCase().replace(/[~～]/g, '-'); }

/** 一行裡的品名：去掉序號、貨品編號、數字、單位 */
function slipNameOf_(l) {
  const parts = l.replace(/^\s*\d{1,2}\s*[.、]/, ' ').replace(SLIP_CODE, ' ').replace(/[A-Za-z]?\d{5,}/g, ' ')   // 序號、貨號、條碼
    .replace(/^([^(（]*?)[(（][^)）]*[)）]\s*(?=[\u4e00-\u9fff])/, '$1')   // 品名前面的括號，例如「(PET瓶)公賣局米酒」
    .replace(/\d+(?:,\d{3})*(?:\.\d+)?/g, ' ').replace(new RegExp(SLIP_UNITS.source, 'gi'), ' ')
    .replace(/[:：|]/g, ' ').replace(/([\u4e00-\u9fff]) (?=[\u4e00-\u9fff])/g, '$1').replace(/\s+/g, ' ').trim().split(' ');
  const nm = parts.filter(function (x) { return /[\u4e00-\u9fff]{2,}|[A-Za-z]{3,}/.test(x); })[0] || parts[0] || '';
  return nm.replace(/^(.{2,}?)[(（].*$/, '$1').slice(0, 40);
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
  const ok = function (a, b, c) { return ns[c] > 0 && Math.abs(ns[a] * ns[b] - ns[c]) <= Math.max(0.6, ns[c] * 0.01); };
  // 先找連在一起的「數量 單價 金額」（從後面找，前面規格欄的數字比較不會誤判）
  for (let i = ns.length - 3; i >= 0; i--) if (ok(i, i + 1, i + 2)) return { qty: ns[i], price: ns[i + 1], amount: ns[i + 2], sure: true, qi: i };
  for (let a = 0; a < ns.length; a++) for (let b = a + 1; b < ns.length; b++) for (let c = b + 1; c < ns.length; c++) {
    if (ns[c] > 0 && Math.abs(ns[a] * ns[b] - ns[c]) <= Math.max(0.6, ns[c] * 0.01)) return { qty: ns[a], price: ns[b], amount: ns[c], sure: true, qi: a };
  }
  out.qty = ns[0]; out.qi = 0;
  if (ns.length >= 3) { out.price = ns[ns.length - 2]; out.amount = ns[ns.length - 1]; }
  else if (ns.length === 2) { out.amount = ns[1]; out.price = ns[0] ? round2_(ns[1] / ns[0]) : ''; }
  return out;
}

/** 單價表：每家廠商每樣食材（同一單位）留最新的單價，之後算成本用 */
function savePrices_(vendor, from, id, slip) {
  const list = slip.filter(function (x) { return x.name && Number(x.price) > 0 && !x.noPrice; });
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
// 月結對帳：一家廠商一個月送了多少錢（有填銷貨單金額用填的，沒填用「實到數量 × 單價表」估），跟發票／對帳單比
// ---------------------------------------------------------------------

const MONTH_HEADERS = ['廠商', '月份', '系統金額', '發票金額', '差額', '對帳時間', '發票照片', '備註'];

/** vendor：叫貨單上的廠商（A菜商、C蛋商、E烘焙乾貨…）；ym：'yyyy-MM' */
function getMonth_(vendor, ym) {
  const m = String(ym || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) throw new Error('月份不正確。');
  const start = new Date(Number(m[1]), Number(m[2]) - 1, 1), end = new Date(Number(m[1]), Number(m[2]), 0);
  const ms = dateStr_(start), me = dateStr_(end);
  const first = mondayOf_(start);
  const rows = purchaseRows_(dateStr_(first), me), prices = readPrices_(), recs = readSlipRecs_();
  const vendors = {};
  rows.forEach(function (r) { const sv = SITE_VENDORS[r.vendor]; vendors[sv ? sv.base : r.vendor] = true; });
  extraSlips_(recs, ms, me).forEach(function (x) { vendors[x.vendor] = true; });
  vendor = clean_(vendor, 30);
  const out = { vendor: vendor, ym: ym, vendors: Object.keys(vendors).sort(), entries: [], record: null };
  if (!vendor) return out;
  const pr = prices[vendor] || {};
  for (let d = first; d <= end; d = addDays_(d, 7)) {
    const f = dateStr_(d), wk = weekBatches_(rows, f, readChecks_(f), prices, recs);
    wk.batches.filter(function (b) { return b.vendor === vendor; }).forEach(function (b) {
      const inMonth = b.list.filter(function (r) { return r.date >= ms && r.date <= me; });
      if (!inMonth.length) return;
      const dates = b.list.map(function (r) { return r.date; }).sort(), whole = inMonth.length === b.list.length;
      // 點總量的（乾貨、蛋）一週只送一次：照第一天算在哪個月
      if (!b.perClass && dates[0] < ms) return;
      const e = { id: b.id, label: b.label, date: inMonth.map(function (r) { return r.date; }).sort()[0], checked: !!b.check, items: [], amount: 0, estimated: false, noPrice: [] };
      // 照片明細：「銷貨單」分頁拍的照片、驗收時有問題品項的照片
      e.photos = b.slip ? (b.slip.photos || []) : [];
      e.itemPhotos = b.check ? (b.check.lines || []).filter(function (l) { return l.photos && l.photos.length; })
        .map(function (l) { return { name: (l.cls ? l.cls + ' ' : '') + l.name, status: l.status, note: l.note, photos: l.photos }; }) : [];
      const lineDate = {};
      b.lines.forEach(function (x) { lineDate[x.cls + '|' + x.name] = x.date || dates[0]; });
      if (b.slip && whole && b.slip.amount > 0) {
        // 填了銷貨單金額：照填的（只填合計的沒有品項，列一行合計）
        b.slip.rows.forEach(function (x) {
          e.items.push({ cls: '', name: x.name || x.slipName, got: x.qty + ' ' + x.unit, price: x.price, amount: Number(x.amount) || 0, kind: x.kind || '食材' });
        });
        if (!b.slip.rows.length) e.items.push({ cls: '', name: '銷貨單合計（只填合計，沒有品項）', got: '', price: '', amount: b.slip.amount });
        e.fromSlip = true;
        e.amount = b.slip.amount;
      } else if (b.check) {
        b.check.lines.forEach(function (l) {
          const dt = lineDate[l.cls + '|' + l.name] || dates[0];
          if (dt < ms || dt > me) return;
          e.items.push(estimateItem_(l.cls, l.name, l.got, pr[l.name], e));
        });
      } else {
        // 還沒驗收：照叫貨量估
        const tot = nameTotals_(b.perClass ? inMonth : b.list);
        Object.keys(tot).forEach(function (n) {
          const t0 = tot[n];
          e.items.push(estimateItem_('', n, t0.unit === '兩' ? fmtTotal_(t0.qty, '兩', n) : fmtQty_(t0.qty) + ' ' + t0.unit, pr[n], e));
        });
      }
      e.amount = round2_(e.amount);
      out.entries.push(e);
    });
  }
  // 臨時銷貨單（不在叫貨單裡的送貨）：照填的金額
  extraSlips_(recs, ms, me, vendor).forEach(function (x) {
    out.entries.push({ id: x.id, label: vendor + '（臨時銷貨單）', date: x.date, checked: true, extra: true, fromSlip: true, amount: round2_(x.slip.amount || 0), estimated: false, noPrice: [],
      photos: x.slip.photos, itemPhotos: [], items: x.slip.rows.map(function (r) { return { cls: '', name: r.name || r.slipName, got: r.qty + ' ' + r.unit, price: r.price, amount: Number(r.amount) || 0, kind: r.kind || '食材' }; }) });
  });
  out.entries.sort(function (a, b) { return a.date.localeCompare(b.date) || a.label.localeCompare(b.label); });
  if (ss_().getSheetByName('月結對帳')) {
    const r = readTable_('月結對帳', ['廠商', '月份']).rows.filter(function (x) { return String(x['廠商']).trim() === vendor && monthStr_(x['月份']) === ym; })[0];
    if (r) out.record = { invoice: r['發票金額'], system: r['系統金額'], time: fmtDateTime_(r['對帳時間']), photos: String(r['發票照片'] || '').split(/\n+/).filter(Boolean), note: String(r['備註'] || '') };
  }
  return out;
}

function monthStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM');
  return String(v || '').trim().replace(/^'/, '').replace(/\//g, '-').slice(0, 7);
}

/** 一項的金額 ≈ 數量 × 單價表的價格（單位不同時用重量換算） */
function estimateItem_(cls, name, got, p, e) {
  const q = parseQty_(got), it = { cls: cls, name: name, got: got, price: p ? p.price : '', amount: null, unit: p ? p.unit : '' };
  if (q && q.qty === 0) { it.amount = 0; return it; }
  if (!p) { e.noPrice.push(name); return it; }
  const qty = qtyIn_(q, p);
  if (qty === null) { e.noPrice.push(name + '（單位對不上）'); return it; }
  it.amount = round2_(qty * p.price);
  e.amount += it.amount; e.estimated = true;
  return it;
}

/** 數量 q（{qty, unit}）換成單價表那一項的單位：同單位、重量互換，或用「換算」（例如 1 K＝190 顆、1 罐＝1000 g）。換不了回傳 null */
function qtyIn_(q, p) {
  if (!q) return null;
  if (q.unit === p.unit) return q.qty;
  if (GRAMS[q.unit] && GRAMS[p.unit]) return q.qty * GRAMS[q.unit] / GRAMS[p.unit];
  const c = p.conv;
  if (c) {
    if (q.unit === c.unit) return q.qty / c.qty;
    if (GRAMS[q.unit] && GRAMS[c.unit]) return q.qty * GRAMS[q.unit] / GRAMS[c.unit] / c.qty;
  }
  return null;
}

/** 「1 斤 8 兩」「2000 g」「10 顆」「2斤半」「半斤」「0」→ { qty, unit } */
function parseQty_(s) {
  s = String(s || '').replace(/\s+/g, '');
  if (s === '' ) return null;
  if (/^0+(\.0+)?$/.test(s)) return { qty: 0, unit: '' };
  let m = s.match(/^半(斤|公斤)$/);
  if (m) return { qty: 0.5, unit: m[1] };
  m = s.match(/^([\d.]+)斤(半|([\d.]+)兩)?$/);
  if (m) return { qty: Number(m[1]) + (m[2] === '半' ? 0.5 : m[3] ? Number(m[3]) / 16 : 0), unit: '斤' };
  m = s.match(/^([\d.]+)(公斤|kg)半?$/i);
  if (m) return { qty: Number(m[1]) + (/半$/.test(s) ? 0.5 : 0), unit: '公斤' };
  m = s.match(/^([\d.]+)([^\d.]+)$/);
  if (m) return { qty: Number(m[1]), unit: ({ '克': 'g', 'G': 'g', '台斤': '斤', 'KG': '公斤', 'kg': '公斤' })[m[2]] || m[2] };
  return null;
}

/** 存對帳結果（同一家同一個月存一筆，重存會蓋掉） */
function saveMonth_(p) {
  const vendor = clean_(p.vendor, 30), ym = String(p.ym || '');
  if (!vendor || !/^\d{4}-\d{2}$/.test(ym)) throw new Error('資料不完整，請重新整理。');
  const sys = num_(p.system), inv = num_(p.invoice);
  if (inv === '') throw new Error('請填發票（或對帳單）上的金額。');
  const lock = waitLock_();
  try {
    ensureSheet_('月結對帳', MONTH_HEADERS);
    const t = readTable_('月結對帳', ['廠商', '月份']);
    const old = t.rows.filter(function (x) { return String(x['廠商']).trim() === vendor && monthStr_(x['月份']) === ym; });
    if (old.length) deleteRows_(t.sheet, old.map(function (r) { return r._row; }));
    appendRows_('月結對帳', [{ '廠商': vendor, '月份': "'" + ym, '系統金額': sys, '發票金額': inv, '差額': sys === '' ? '' : round2_(inv - sys),
      '對帳時間': new Date(), '發票照片': (p.photos || []).join('\n'), '備註': clean_(p.note, 200) }]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// ---------------------------------------------------------------------
// 暫時儲存：點到一半先存到試算表（換平板、清掉瀏覽器資料也不會不見），完成驗收後刪掉
// ---------------------------------------------------------------------

function readDrafts_() {
  const out = {};
  if (!ss_().getSheetByName(SHEET.drafts)) return out;
  readTable_(SHEET.drafts, ['驗收編號', '內容']).rows.forEach(function (r) {
    const t = r['暫存時間'] instanceof Date ? r['暫存時間'] : null;
    out[String(r['驗收編號']).trim()] = { ts: t ? t.getTime() : 0, at: t ? Utilities.formatDate(t, tz_(), 'M/d HH:mm') : '', data: String(r['內容'] || '') };
  });
  return out;
}

/** id：驗收編號；data：JSON 字串（點到哪裡、實到、備註、照片、銷貨單金額…，不含簽名） */
function saveDraftRemote_(id, data) {
  id = String(id || '').trim();
  if (!/^R\d{8}-/.test(id)) throw new Error('驗收編號不正確，請重新整理。');
  data = String(data || '');
  if (data.length > 45000) throw new Error('內容太多，存不進去（照片請少一點）。');
  try { JSON.parse(data); } catch (e) { throw new Error('內容格式不正確。'); }
  const lock = waitLock_();
  try {
    ensureSheet_(SHEET.drafts, DRAFT_HEADERS);
    dropDraft_(id);
    const now = new Date();
    appendRows_(SHEET.drafts, [{ '驗收編號': id, '暫存時間': now, '內容': data }]);
    return { ts: now.getTime(), at: Utilities.formatDate(now, tz_(), 'M/d HH:mm') };
  } finally {
    lock.releaseLock();
  }
}

/** 「清掉、重新點」：刪掉試算表裡的暫存 */
function dropDraftRemote_(id) {
  id = String(id || '').trim();
  if (!/^R\d{8}-/.test(id)) return { ok: true };
  const lock = waitLock_();
  try { dropDraft_(id); } finally { lock.releaseLock(); }
  return { ok: true };
}

function dropDraft_(id) {
  if (!ss_().getSheetByName(SHEET.drafts)) return;
  const t = readTable_(SHEET.drafts, ['驗收編號']);
  const old = t.rows.filter(function (r) { return String(r['驗收編號']).trim() === id; });
  if (old.length) deleteRows_(t.sheet, old.map(function (r) { return r._row; }));
}

// ---------------------------------------------------------------------
// 問題追蹤
// ---------------------------------------------------------------------

/** all = true 連處理完的一起列（最近 60 天） */
function getIssues_(all) {
  const sh = ss_().getSheetByName(SHEET.lines);
  if (!sh || sh.getLastRow() < 2) return { list: [], states: ISSUE_STATES, pending: draftIssues_() };
  const checks = {};
  if (ss_().getSheetByName(SHEET.checks)) readTable_(SHEET.checks, ['驗收編號']).rows.forEach(function (r) {
    checks[String(r['驗收編號']).trim()] = { from: purchaseDate_(r['送貨日期（起）']), to: purchaseDate_(r['送貨日期（迄）']),
      time: fmtDateTime_(r['驗收時間']), photos: String(r['銷貨單照片'] || '').split(/\n+/).filter(Boolean) };
  });
  const recs = readSlipRecs_();   // 「銷貨單」分頁存的照片
  Object.keys(checks).forEach(function (id) { if (recs[id] && recs[id].photos.length) checks[id].photos = recs[id].photos; });
  const since = dateStr_(addDays_(today_(), -60));
  const list = readTable_(SHEET.lines, ['驗收編號', '狀態']).rows.filter(function (r) {
    const st = String(r['狀態'] || '').trim();
    if (!st || st === '到齊' || st === '停送') return false;
    const done = String(r['處理狀態'] || '').trim() && String(r['處理狀態']).trim() !== '待處理';
    const c = checks[String(r['驗收編號']).trim()];
    return all ? !c || c.from >= since : !done;
  }).map(function (r) {
    const id = String(r['驗收編號']).trim(), c = checks[id] || {};
    return { row: r._row, id: id, vendor: String(r['廠商'] || ''), group: String(r['送貨組別'] || ''), from: c.from || '', to: c.to || '', time: c.time || '',
      photos: c.photos || [], cls: String(r['班級'] || ''), name: String(r['食材名稱'] || ''), expect: String(r['應到'] || ''), got: String(r['實到'] || ''),
      status: String(r['狀態'] || ''), note: String(r['備註'] || ''), state: String(r['處理狀態'] || '待處理') || '待處理',
      doneTime: fmtDateTime_(r['處理時間']), doneNote: String(r['處理備註'] || ''), itemPhotos: photoList_(r['照片']) };
  });
  return { list: list, states: ISSUE_STATES, pending: draftIssues_() };
}

/** 暫時儲存（還沒完成驗收）裡標了缺貨、少送…的品項，先列在問題追蹤，免得忘記 */
function draftIssues_() {
  const out = [], drafts = readDrafts_();
  Object.keys(drafts).forEach(function (id) {
    const m = id.match(/^R(\d{4})(\d{2})(\d{2})-(.+)$/);
    if (!m) return;
    let d = null;
    try { d = JSON.parse(drafts[id].data); } catch (e) { return; }
    const items = [];
    Object.keys(d && d.marks || {}).forEach(function (k) {
      const x = d.marks[k], st = x && x.status;
      if (!st || st === '到齊' || st === '停送') return;
      const p = k.split('|');
      items.push({ cls: p[0] || '', name: p[1] || k, status: st, got: String(x.got || ''), note: String(x.note || '') });
    });
    if (!items.length) return;
    const rest = m[4].split('-');
    out.push({ id: id, from: m[1] + '-' + m[2] + '-' + m[3], vendor: rest[0], group: rest.slice(1).join('-'), at: drafts[id].at, items: items });
  });
  return out.sort(function (a, b) { return a.from < b.from ? -1 : a.from > b.from ? 1 : 0; });
}

/** 記錄問題的處理結果：rows = 驗收明細的列號 */
function setIssues_(rows, state, note) {
  if (ISSUE_STATES.indexOf(state) < 0) throw new Error('處理狀態不正確。');
  const lock = waitLock_();
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
  return getIssues_(false);
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
// 核銷專區：發票（一張發票包含哪幾次送貨）→ 拆到各經費類別 → 各類別經手人確認「已核銷」→ 統計表
// 跟驗收（點貨）分開：只讀驗收的結果，不改驗收紀錄
// ---------------------------------------------------------------------

const HX_INVOICE_HEADERS = ['發票編號', '廠商', '發票號碼', '發票日期', '發票金額', '涵蓋方式', '起', '迄', '驗收編號', '系統金額', '發票照片', '備註', '建立時間', '建立人'];
const HX_ALLOC_HEADERS = ['發票編號', '類別', '金額', '經手人', '狀態', '核銷日期', '核銷人', '明細', '經費編號'];   // 經費編號：「C30033=18400、C30034=9200」   // 明細：這一類填的好幾筆，例如「150+80+20」
const HX_LINE_HEADERS = ['發票編號', '發票號碼', '發票日期', '金額', '照片', '備註'];
// 另外叫料：不在叫貨單裡的（臨時電話加叫、系統外的商店、活動用、老師先買後報帳），一項一列，金額算進那一類
const HX_EXTRA_HEADERS = ['發票編號', '日期', '品名', '金額', '類別', '購買人', '備註'];   // 一筆核銷可以有好幾張發票（「發票編號」是這筆核銷的編號）
const HX_CAT_HEADERS = ['類別', '經手人', '排序'];
const HX_BUDGET_HEADERS = ['預算編號', '類別', '名稱', '起', '迄', '預算', '備註', '更新時間', '更新人'];
const HX_DEFAULT_CATS = ['國中技藝班（餐旅）', '國中技藝班（食品）', '特教班', '實習實作', '餐飲科課費', '國中生涯實作（三信）', '國中生涯實作（國中）', '未未成果展'];
const HX_STOCK_CAT = '餐飲科課費';   // 補庫存（沒有班級、大家共用的）先帶這一類

/** 類別和經手人：[{ name, handler }]（第一次用會建「核銷類別」分頁，放預設的 8 類） */
function hxCats_() {
  const sh = ss_().getSheetByName('核銷類別');
  if (!sh || sh.getLastRow() < 2) {
    ensureSheet_('核銷類別', HX_CAT_HEADERS);
    if (ss_().getSheetByName('核銷類別').getLastRow() < 2) appendRows_('核銷類別', HX_DEFAULT_CATS.map(function (n, i) { return { '類別': n, '經手人': '', '排序': i + 1 }; }));
  }
  return readTable_('核銷類別', ['類別']).rows.filter(function (r) { return String(r['類別']).trim(); })
    .sort(function (a, b) { return (Number(a['排序']) || 99) - (Number(b['排序']) || 99); })
    .map(function (r) { return { name: String(r['類別']).trim(), handler: String(r['經手人'] || '').trim() }; });
}

/** 存類別設定：list = [{ name, handler }]，照順序 */
function hxSaveCats_(list) {
  list = (list || []).map(function (x) { return { name: clean_(x.name, 30), handler: clean_(x.handler, 30) }; }).filter(function (x) { return x.name; });
  if (!list.length) throw new Error('至少要有一個類別。');
  const seen = {};
  list.forEach(function (x) { if (seen[x.name]) throw new Error('類別「' + x.name + '」重複了。'); seen[x.name] = true; });
  const lock = waitLock_();
  try {
    ensureSheet_('核銷類別', HX_CAT_HEADERS);
    const t = readTable_('核銷類別', ['類別']);
    if (t.rows.length) deleteRows_(t.sheet, t.rows.map(function (r) { return r._row; }));
    appendRows_('核銷類別', list.map(function (x, i) { return { '類別': x.name, '經手人': x.handler, '排序': i + 1 }; }));
  } finally {
    lock.releaseLock();
  }
  return hxCats_();
}

function hxReadInvoices_() {
  const out = [], by = {};
  if (ss_().getSheetByName('核銷發票')) readTable_('核銷發票', ['發票編號']).rows.forEach(function (r) {
    const id = String(r['發票編號']).trim();
    if (!id) return;
    const d = r['發票日期'] instanceof Date ? dateStr_(r['發票日期']) : String(r['發票日期'] || '').replace(/^'/, '');
    by[id] = { id: id, vendor: String(r['廠商'] || ''), no: String(r['發票號碼'] || '').replace(/^'/, ''), date: d, amount: Number(r['發票金額']) || 0,
      mode: String(r['涵蓋方式'] || ''), from: r['起'] instanceof Date ? dateStr_(r['起']) : String(r['起'] || '').replace(/^'/, ''),
      to: r['迄'] instanceof Date ? dateStr_(r['迄']) : String(r['迄'] || '').replace(/^'/, ''),
      checks: String(r['驗收編號'] || '').split(/\n+/).filter(Boolean), system: r['系統金額'] === '' ? '' : Number(r['系統金額']),
      photos: String(r['發票照片'] || '').split(/\n+/).filter(Boolean), note: String(r['備註'] || ''), created: fmtDateTime_(r['建立時間']), by: String(r['建立人'] || ''), allocs: [], lines: [], extras: [] };
    out.push(by[id]);
  });
  if (ss_().getSheetByName('核銷分配')) readTable_('核銷分配', ['發票編號', '類別']).rows.forEach(function (r) {
    const inv = by[String(r['發票編號']).trim()];
    if (!inv) return;
    const amt = Number(r['金額']) || 0, parts = String(r['明細'] === undefined ? '' : r['明細']).replace(/^'/, '').split('+').map(Number).filter(function (n) { return !isNaN(n) && n; });
    inv.allocs.push({ cat: String(r['類別'] || ''), amount: amt, parts: parts.length || String(r['明細'] || '').replace(/^'/, '') === '0' ? parts : [amt], handler: String(r['經手人'] || ''), status: String(r['狀態'] || '待核銷'),
      doneDate: r['核銷日期'] instanceof Date ? dateStr_(r['核銷日期']) : String(r['核銷日期'] || ''), doneBy: String(r['核銷人'] || ''),
      codes: String(r['經費編號'] || '').split(/[、,，\n]+/).map(function (x) { const m = x.trim().match(/^([A-Za-z0-9-]+)=([\d.]+)$/); return m ? { code: m[1], amount: Number(m[2]) } : null; }).filter(Boolean) });
  });
  if (ss_().getSheetByName('核銷發票明細')) readTable_('核銷發票明細', ['發票編號']).rows.forEach(function (r) {
    const inv = by[String(r['發票編號']).trim()];
    if (!inv) return;
    inv.lines.push({ no: String(r['發票號碼'] || '').replace(/^'/, ''), date: r['發票日期'] instanceof Date ? dateStr_(r['發票日期']) : String(r['發票日期'] || '').replace(/^'/, ''),
      amount: Number(r['金額']) || 0, photos: String(r['照片'] || '').split(/\n+/).filter(Boolean), note: String(r['備註'] || '') });
  });
  if (ss_().getSheetByName('核銷另外叫料')) readTable_('核銷另外叫料', ['發票編號']).rows.forEach(function (r) {
    const inv = by[String(r['發票編號']).trim()];
    if (!inv) return;
    inv.extras.push({ date: r['日期'] instanceof Date ? dateStr_(r['日期']) : String(r['日期'] || '').replace(/^'/, ''), name: String(r['品名'] || ''), amount: Number(r['金額']) || 0,
      cat: String(r['類別'] || ''), buyer: String(r['購買人'] || ''), note: String(r['備註'] || '') });
  });
  // 舊的（只有一張發票、還沒有明細的）
  out.forEach(function (inv) { if (!inv.lines.length) inv.lines.push({ no: inv.no, date: inv.date, amount: inv.amount, photos: inv.photos }); });
  return out.sort(function (a, b) { return b.date.localeCompare(a.date) || b.created.localeCompare(a.created); });
}

/** 核銷首頁：類別、廠商、發票（可以篩廠商、發票月份 'yyyy-MM'） */
function hxGet_(vendor, ym) {
  const cats = hxCats_();
  let list = hxReadInvoices_();
  const vendors = {};
  list.forEach(function (x) { vendors[x.vendor] = true; });
  // 最近三個月有叫貨的廠商也列進選單
  const now = today_();
  readTable_(SHEET.purchase, []).rows.forEach(function (r) {
    const d = parseDate_(purchaseDate_(r['送貨日期']));
    if (!d || d < addDays_(now, -100)) return;
    const v = String(r['廠商'] || '').trim(), sv = SITE_VENDORS[v];
    if (v) vendors[sv ? sv.base : v] = true;
  });
  const rc = readSlipRecs_();   // 臨時銷貨單的廠商（例如全聯）也列進選單
  Object.keys(rc).forEach(function (id) { if (/^S\d{8}-/.test(id) && rc[id].vendor) vendors[rc[id].vendor] = true; });
  const budgets = hxBudgets_(list), codes = hxCodes_(list);
  if (vendor) list = list.filter(function (x) { return x.vendor === vendor; });
  if (ym) list = list.filter(function (x) { return x.date.slice(0, 7) === ym; });
  return { cats: cats, vendors: Object.keys(vendors).filter(Boolean).sort(), invoices: list, budgets: budgets, codes: codes, me: CURRENT_USER_ };
}

/** 這家廠商在 from～to 之間的每一次送貨（驗收）：金額、各班金額，已經算進哪張發票 */
function hxCandidates_(vendor, from, to, exceptId) {
  vendor = clean_(vendor, 30);
  const f = parseDate_(from), t = parseDate_(to);
  if (!vendor || !f || !t) throw new Error('請選廠商和日期。');
  if (t < f) throw new Error('迄日不能早於起日。');
  if ((t - f) / 86400000 > 190) throw new Error('一次最多查半年。');
  const used = {};
  hxReadInvoices_().forEach(function (inv) { if (inv.id !== exceptId) inv.checks.forEach(function (c) { used[c] = inv.no || inv.id; }); });
  const out = [];
  for (let m = new Date(f.getFullYear(), f.getMonth(), 1); m <= t; m = new Date(m.getFullYear(), m.getMonth() + 1, 1)) {
    const ym = Utilities.formatDate(m, tz_(), 'yyyy-MM');
    getMonth_(vendor, ym).entries.forEach(function (e) {
      if (e.date < dateStr_(f) || e.date > dateStr_(t)) return;
      const byCls = {};
      const byKind = {};
      e.items.forEach(function (x) { const k = x.cls || ''; byCls[k] = round2_((byCls[k] || 0) + (Number(x.amount) || 0)); const kd = x.kind || '食材'; byKind[kd] = round2_((byKind[kd] || 0) + (Number(x.amount) || 0)); });
      out.push({ key: e.id + '@' + ym, id: e.id, date: e.date, label: e.label, checked: e.checked, amount: e.amount, estimated: e.estimated,
        noPrice: e.noPrice.length, byCls: byCls, byKind: byKind, items: e.items, usedBy: used[e.id + '@' + ym] || '' });
    });
  }
  return out.sort(function (a, b) { return a.date.localeCompare(b.date); });
}

/**
 * 新增／修改一筆核銷：p = { id?, vendor, lines:[{no, date, amount, photos}]（好幾張發票）, mode, from, to, checks:[key], system, note,
 *   allocs:[{cat, parts:[金額…]}]（每一類可以填好幾筆） }
 */
function hxSaveInvoice_(p) {
  const vendor = clean_(p.vendor, 30);
  if (!vendor) throw new Error('請填廠商（不在系統裡的商店也可以直接打名字）。');
  const lines = (p.lines || []).map(function (l) { return { no: clean_(l.no, 30), date: parseDate_(l.date) ? dateStr_(parseDate_(l.date)) : '', amount: round2_(Number(l.amount) || 0), photos: (l.photos || []).slice(0, 20), note: clean_(l.note, 100) }; })
    .filter(function (l) { return l.no || l.amount || l.photos.length; });
  if (!lines.length) throw new Error('請至少填一張發票。');
  lines.forEach(function (l, i) {
    if (!l.date) throw new Error('第 ' + (i + 1) + ' 張發票沒有填日期。');
    if (!(l.amount > 0)) throw new Error('第 ' + (i + 1) + ' 張發票沒有填金額。');
  });
  const amount = round2_(lines.reduce(function (n, l) { return n + l.amount; }, 0));
  p.date = lines.map(function (l) { return l.date; }).sort().pop();   // 統計表照最後一張發票的月份
  p.no = lines.map(function (l) { return l.no; }).filter(Boolean).join('、');
  p.photos = [].concat.apply([], lines.map(function (l) { return l.photos; }));
  const cats = hxCats_(), handler = {};
  cats.forEach(function (c) { handler[c.name] = c.handler; });
  const extras = (p.extras || []).map(function (x) {
    return { date: parseDate_(x.date) ? dateStr_(parseDate_(x.date)) : '', name: clean_(x.name, 60), amount: round2_(Number(x.amount) || 0), cat: clean_(x.cat, 30), buyer: clean_(x.buyer, 30), note: clean_(x.note, 100) };
  }).filter(function (x) { return x.name || x.amount; });
  extras.forEach(function (x, i) {
    if (!x.name) throw new Error('另外叫料第 ' + (i + 1) + ' 項沒有填品名。');
    if (!x.amount) throw new Error('另外叫料「' + x.name + '」沒有填金額。');
    if (!x.cat) throw new Error('另外叫料「' + x.name + '」沒有選類別。');
  });
  // 每一類＝自己填的幾筆（明細）＋另外叫料這一類的金額
  const byCat = {};
  (p.allocs || []).forEach(function (a) {
    const c = clean_(a.cat, 30), parts = (a.parts || [a.amount]).map(Number).filter(function (n) { return !isNaN(n) && n; });
    const codes = (a.codes || []).map(function (x) { return { code: clean_(x.code, 20).replace(/[^A-Za-z0-9-]/g, ''), amount: round2_(Number(x.amount) || 0) }; }).filter(function (x) { return x.code && x.amount > 0; });
    if (c) { const o = byCat[c] = byCat[c] || { parts: [], extra: 0, codes: [] }; o.parts = o.parts.concat(parts); o.codes = (o.codes || []).concat(codes); }
  });
  extras.forEach(function (x) { (byCat[x.cat] = byCat[x.cat] || { parts: [], extra: 0 }).extra = round2_(byCat[x.cat].extra + x.amount); });
  const allocs = Object.keys(byCat).map(function (c) {
    const o = byCat[c];
    return { cat: c, parts: o.parts, codes: o.codes || [], amount: round2_(o.parts.reduce(function (n, x) { return n + x; }, 0) + o.extra) };
  }).filter(function (a) { return a.amount; });
  if (!allocs.length) throw new Error('請把發票金額拆到類別。');
  const sum = round2_(allocs.reduce(function (n, a) { return n + a.amount; }, 0));
  if (Math.abs(sum - amount) >= 1) throw new Error('各類別合計 $' + sum + '，跟發票合計 $' + amount + ' 不一樣。');
  allocs.forEach(function (a) { if (handler[a.cat] === undefined) throw new Error('沒有「' + a.cat + '」這個類別。'); });
  hxEnsureSheets_();   // 建分頁、補欄位放在鎖外面（第一次比較慢，不要卡住別人）
  const lock = waitLock_();
  try {
    let id = clean_(p.id, 40), old = {};
    if (id) {
      const ti = readTable_('核銷發票', ['發票編號']), ta = readTable_('核銷分配', ['發票編號', '類別']);
      const ri = ti.rows.filter(function (r) { return String(r['發票編號']).trim() === id; });
      if (!ri.length) throw new Error('找不到這張發票，可能被刪掉了。');
      const ra = ta.rows.filter(function (r) { return String(r['發票編號']).trim() === id; });
      ra.forEach(function (r) { old[String(r['類別'])] = r; });   // 已核銷的類別：金額沒改就保留核銷狀態
      deleteRows_(ta.sheet, ra.map(function (r) { return r._row; }));
      deleteRows_(ti.sheet, ri.map(function (r) { return r._row; }));
      ['核銷發票明細', '核銷另外叫料'].forEach(function (n) {
        const tl = readTable_(n, ['發票編號']);
        deleteRows_(tl.sheet, tl.rows.filter(function (r) { return String(r['發票編號']).trim() === id; }).map(function (r) { return r._row; }));
      });
    } else {
      id = 'HX' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    }
    appendRows_('核銷發票', [{ '發票編號': id, '廠商': vendor, '發票號碼': "'" + clean_(p.no, 300), '發票日期': "'" + dateStr_(parseDate_(p.date)), '發票金額': amount,
      '涵蓋方式': p.mode === 'range' ? '日期區間' : '勾選', '起': p.from ? "'" + p.from : '', '迄': p.to ? "'" + p.to : '', '驗收編號': (p.checks || []).join('\n'),
      '系統金額': p.system === '' || p.system === undefined ? '' : round2_(p.system), '發票照片': (p.photos || []).join('\n'), '備註': clean_(p.note, 200),
      '建立時間': new Date(), '建立人': CURRENT_USER_ }]);
    appendRows_('核銷分配', allocs.map(function (a) {
      const o = old[a.cat], keep = o && String(o['狀態']) === '已核銷' && Math.abs((Number(o['金額']) || 0) - a.amount) < 0.01;
      return { '發票編號': id, '類別': a.cat, '金額': a.amount, '經手人': handler[a.cat] || '', '狀態': keep ? '已核銷' : '待核銷',
        '核銷日期': keep ? o['核銷日期'] : '', '核銷人': keep ? o['核銷人'] : '', '明細': "'" + (a.parts.length ? a.parts.join('+') : '0'),
        '經費編號': a.codes.map(function (x) { return x.code + '=' + x.amount; }).join('、') };
    }));
    if (extras.length) appendRows_('核銷另外叫料', extras.map(function (x) {
      return { '發票編號': id, '日期': x.date ? "'" + x.date : '', '品名': x.name, '金額': x.amount, '類別': x.cat, '購買人': x.buyer, '備註': x.note };
    }));
    appendRows_('核銷發票明細', lines.map(function (l) {
      return { '發票編號': id, '發票號碼': "'" + l.no, '發票日期': "'" + l.date, '金額': l.amount, '照片': l.photos.join('\n'), '備註': l.note };
    }));
    return { id: id };
  } finally {
    lock.releaseLock();
  }
}

/** 核銷用到的分頁：沒有就建、缺欄位就補；檢查過記 6 小時，之後不再每次檢查 */
function hxEnsureSheets_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('hxSheetsOk3')) return;
  ensureSheet_('核銷發票', HX_INVOICE_HEADERS);
  ensureSheet_('核銷分配', HX_ALLOC_HEADERS);
  ensureSheet_('核銷發票明細', HX_LINE_HEADERS);
  ensureSheet_('核銷另外叫料', HX_EXTRA_HEADERS);
  cache.put('hxSheetsOk3', '1', 6 * 3600);
}

// ---------------------------------------------------------------------
// 另外叫料快速 key：食材總表的品名（打字跳出來選）、從老師的訂單匯入
// ---------------------------------------------------------------------

/**
 * 列印請購單：這筆核銷勾的送貨，照銷貨單上每張選的「食材／器材」分開
 * 回傳 { vendor, date, amount, no, kinds: { 食材: { amount, items: [{ name, unit, qty, amount }] }, 器材: {…} } }
 */
function hxPrintData_(id) {
  const inv = hxReadInvoices_().filter(function (x) { return x.id === id; })[0];
  if (!inv) throw new Error('找不到這筆核銷，可能被刪掉了。');
  const kinds = {}, add = function (kind, name, qty, unit, amount) {
    const k = kinds[kind] = kinds[kind] || { amount: 0, items: [] };
    const n = String(name || '').trim() || '（沒有品名）', u = String(unit || '').trim();
    let it = k.items.filter(function (x) { return x.name === n && x.unit === u; })[0];
    if (!it) { it = { name: n, unit: u, qty: 0, amount: 0 }; k.items.push(it); }
    const q = Number(qty); if (q > 0) it.qty = Math.round((it.qty + q) * 10000) / 10000;
    it.amount = round2_(it.amount + (Number(amount) || 0)); k.amount = round2_(k.amount + (Number(amount) || 0));
  };
  const months = {};
  inv.checks.forEach(function (key) { const m = String(key).split('@'); if (m[1]) (months[m[1]] = months[m[1]] || []).push(m[0]); });
  Object.keys(months).forEach(function (ym) {
    getMonth_(inv.vendor, ym).entries.forEach(function (e) {
      if (months[ym].indexOf(e.id) < 0) return;
      e.items.forEach(function (x) {
        const m = String(x.got || '').trim().match(/^([\d.]+)\s*(.*)$/);
        add(x.kind || '食材', x.name, m ? m[1] : '', m ? m[2] : '', x.amount);
      });
    });
  });
  inv.extras.forEach(function (x) { add('食材', x.name, '', '', x.amount); });   // 另外叫料當食材
  return { id: inv.id, vendor: inv.vendor, date: inv.date, amount: inv.amount, no: inv.no, kinds: kinds };
}

/** 食材總表啟用中的品名：[[名稱, 別名, 廠商]]（快取 1 小時） */
function hxItems_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('hxItems2');
  if (hit) return JSON.parse(hit);
  const out = [];
  if (ss_().getSheetByName(SHEET.items)) readTable_(SHEET.items, ['食材名稱']).rows.forEach(function (r) {
    const n = String(r['食材名稱'] || '').trim();
    if (!n || (r['啟用'] !== undefined && String(r['啟用']).trim() !== '' && !isTrue_(r['啟用']))) return;
    out.push([n, String(r['別名'] || '').trim(), String(r['廠商'] || '').trim(), String(r['單位'] || '').trim()]);
  });
  try { cache.put('hxItems2', JSON.stringify(out), 3600); } catch (e) {}
  return out;
}

/**
 * 老師的訂單（訂單總表＋歷史訂單），使用日期在 from～to：
 * [{ no, date, teacher, cls, project, dept, items: [{ name, qty, unit, vendor, bought }] }]（bought＝已經透過廠商叫貨）
 */
function hxOrders_(from, to, q) {
  const f = parseDate_(from), t = parseDate_(to);
  if (!f || !t) throw new Error('請選日期。');
  if ((t - f) / 86400000 > 92) throw new Error('日期範圍最多 3 個月。');
  const fs = dateStr_(f), ts = dateStr_(t), by = {}, words = String(q || '').trim().toLowerCase().split(/\s+/).filter(Boolean);
  ['訂單總表', '歷史訂單'].forEach(function (name) {
    if (!ss_().getSheetByName(name)) return;
    readTable_(name, ['訂單編號', '食材名稱']).rows.forEach(function (r) {
      const d = purchaseDate_(r['使用日期']), st = String(r['狀態'] || '').trim();
      if (d < fs || d > ts || st === '已取消') return;
      const no = String(r['訂單編號'] || '').trim(), n = String(r['食材名稱'] || '').trim();
      if (!no || !n) return;
      const o = by[no] = by[no] || { no: no, date: d, teacher: String(r['代訂對象老師'] || '').trim() || String(r['操作老師'] || '').trim(),
        cls: String(r['上課班級'] || '').trim(), project: String(r['實作項目'] || '').trim(), dept: String(r['科別'] || '').trim(), items: [] };
      o.items.push({ name: n, qty: r['數量'] === '' ? '' : Number(r['數量']) || r['數量'], unit: String(r['單位'] || '').trim(), vendor: String(r['廠商'] || '').trim(),
        bought: st === '已叫貨' && !isTrue_(r['不叫貨']) && String(r['廠商'] || '').trim() !== '自備' });
    });
  });
  let list = Object.keys(by).map(function (k) { return by[k]; });
  if (words.length) list = list.filter(function (o) {
    const hay = [o.teacher, o.cls, o.project, o.dept, o.items.map(function (x) { return x.name; }).join(' ')].join(' ').toLowerCase();
    return words.every(function (w) { return hay.indexOf(w) >= 0; });
  });
  list.sort(function (a, b) { return b.date.localeCompare(a.date) || a.cls.localeCompare(b.cls); });
  return { list: list.slice(0, 80), more: Math.max(0, list.length - 80) };
}

// ---------------------------------------------------------------------
// 經費編號（例如實習實作 C30033）：核銷時勾編號，這裡算每個編號用了多少、用在哪幾張發票
// ---------------------------------------------------------------------
function hxCodes_(invoices) {
  if (!ss_().getSheetByName('經費編號')) return [];
  const list = readTable_('經費編號', ['編號', '金額']).rows.filter(function (r) { return String(r['編號']).trim(); }).map(function (r) {
    return { code: String(r['編號']).trim(), name: String(r['實習課程名稱'] || r['名稱'] || '').trim(), amount: Number(r['金額']) || 0, vendors: String(r['廠商'] || '').trim(),
      month: String(r['開發票月份'] || '').trim(), cat: String(r['類別'] || '').trim(), group: String(r['組別'] || '').trim(), note: String(r['備註'] || '').trim(), used: 0, done: 0, uses: [] };
  });
  const by = {};
  list.forEach(function (c) { by[c.code] = c; });
  (invoices || hxReadInvoices_()).forEach(function (inv) {
    inv.allocs.forEach(function (a) {
      (a.codes || []).forEach(function (x) {
        const c = by[x.code];
        if (!c) return;
        c.used = round2_(c.used + x.amount);
        if (a.status === '已核銷') c.done = round2_(c.done + x.amount);
        c.uses.push({ id: inv.id, no: inv.no, vendor: inv.vendor, date: inv.date, amount: x.amount, status: a.status });
      });
    });
  });
  list.forEach(function (c) { c.left = round2_(c.amount - c.used); });
  return list;
}

// ---------------------------------------------------------------------
// 預算：每一類可以有好幾筆（各自的起迄日），已用＝這段期間發票日期落在裡面、分到這一類的金額
// ---------------------------------------------------------------------

/** [{ id, cat, name, from, to, amount, note, done, pending, used, left }]（invoices 不給就自己讀） */
function hxBudgets_(invoices) {
  if (!ss_().getSheetByName('核銷預算')) return [];
  const ds = function (v) { return v instanceof Date ? dateStr_(v) : String(v || '').replace(/^'/, ''); };
  const list = readTable_('核銷預算', ['預算編號', '類別']).rows.filter(function (r) { return String(r['預算編號']).trim(); }).map(function (r) {
    return { id: String(r['預算編號']).trim(), cat: String(r['類別'] || '').trim(), name: String(r['名稱'] || ''), from: ds(r['起']), to: ds(r['迄']),
      amount: Number(r['預算']) || 0, note: String(r['備註'] || ''), done: 0, pending: 0, n: 0 };
  });
  if (!list.length) return [];
  (invoices || hxReadInvoices_()).forEach(function (inv) {
    inv.allocs.forEach(function (a) {
      list.forEach(function (b) {
        if (b.cat !== a.cat || inv.date < b.from || inv.date > b.to) return;
        if (a.status === '已核銷') b.done = round2_(b.done + a.amount); else b.pending = round2_(b.pending + a.amount);
        b.n++;
      });
    });
  });
  list.forEach(function (b) { b.used = round2_(b.done + b.pending); b.left = round2_(b.amount - b.used); });
  return list.sort(function (a, b) { return b.from.localeCompare(a.from) || a.cat.localeCompare(b.cat, 'zh-Hant'); });
}

/** 新增或修改預算（一次好幾類）：list = [{ id?, cat, name, from, to, amount, note }] */
function hxSaveBudgets_(list) {
  const rows = (list || []).map(function (b) {
    const cat = clean_(b.cat, 30), name = clean_(b.name, 40), f = parseDate_(b.from), t = parseDate_(b.to), amt = Number(String(b.amount || '').replace(/[,，\s$]/g, ''));
    if (!cat) throw new Error('請選類別。');
    if (!f || !t) throw new Error('「' + cat + '」請填起訖日期。');
    if (t < f) throw new Error('「' + cat + '」迄日不能比起日早。');
    if (!(amt > 0)) throw new Error('「' + cat + '」請填核定金額。');
    return { id: clean_(b.id, 40), row: { '類別': cat, '名稱': name, '起': "'" + dateStr_(f), '迄': "'" + dateStr_(t), '預算': amt, '備註': clean_(b.note, 100), '更新時間': new Date(), '更新人': CURRENT_USER_ || '' } };
  });
  if (!rows.length) return hxBudgets_();
  ensureSheet_('核銷預算', HX_BUDGET_HEADERS);
  const lock = waitLock_();
  try {
    const tb = readTable_('核銷預算', ['預算編號']), add = [];
    rows.forEach(function (x, i) {
      if (x.id) {
        const r = tb.rows.filter(function (y) { return String(y['預算編號']).trim() === x.id; })[0];
        if (!r) throw new Error('「' + x.row['類別'] + '」這一期找不到，可能被別人刪掉了，請重新整理。');
        Object.keys(x.row).forEach(function (h) { setCell_(tb, r._row, h, x.row[h]); });
      } else {
        x.row['預算編號'] = 'BG' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss') + ('0' + i).slice(-2) + Math.floor(Math.random() * 90 + 10);
        add.push(x.row);
      }
    });
    if (add.length) appendRows_('核銷預算', add);
  } finally {
    lock.releaseLock();
  }
  return hxBudgets_();
}

function hxDeleteBudget_(id) {
  id = clean_(id, 40);
  if (!ss_().getSheetByName('核銷預算')) return [];
  const lock = waitLock_();
  try {
    const tb = readTable_('核銷預算', ['預算編號']);
    deleteRows_(tb.sheet, tb.rows.filter(function (r) { return String(r['預算編號']).trim() === id; }).map(function (r) { return r._row; }));
  } finally {
    lock.releaseLock();
  }
  return hxBudgets_();
}

/** 等別人存完再存：最多等 30 秒，等不到就請使用者稍等再按（不要顯示 Google 的英文／技術訊息） */
function waitLock_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('系統正在儲存別的資料（可能是另一台平板或手機），請等幾秒再按一次。');
  return lock;
}

function hxDeleteInvoice_(id) {
  id = clean_(id, 40);
  const lock = waitLock_();
  try {
    ['核銷分配', '核銷發票明細', '核銷另外叫料', '核銷發票'].forEach(function (n) {
      if (!ss_().getSheetByName(n)) return;
      const t = readTable_(n, ['發票編號']);
      deleteRows_(t.sheet, t.rows.filter(function (r) { return String(r['發票編號']).trim() === id; }).map(function (r) { return r._row; }));
    });
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** 某張發票的某一類：經手人確認「已核銷」（done=false 改回待核銷） */
function hxSetDone_(id, cat, done) {
  const t = readTable_('核銷分配', ['發票編號', '類別']);
  const r = t.rows.filter(function (x) { return String(x['發票編號']).trim() === id && String(x['類別']) === cat; })[0];
  if (!r) throw new Error('找不到這筆。');
  setCell_(t, r._row, '狀態', done ? '已核銷' : '待核銷');
  setCell_(t, r._row, '核銷日期', done ? new Date() : '');
  setCell_(t, r._row, '核銷人', done ? CURRENT_USER_ : '');
  return true;
}

/** 統計表：from、to 是 'yyyy-MM'。照發票日期的月份 × 類別 × 廠商加總，分已核銷／待核銷 */
function hxStats_(from, to) {
  const cats = hxCats_().map(function (c) { return c.name; });
  const rows = {};
  hxReadInvoices_().forEach(function (inv) {
    const ym = inv.date.slice(0, 7);
    if ((from && ym < from) || (to && ym > to)) return;
    inv.allocs.forEach(function (a) {
      if (cats.indexOf(a.cat) < 0) cats.push(a.cat);
      const k = [ym, a.cat, inv.vendor].join('|');
      const o = rows[k] = rows[k] || { ym: ym, cat: a.cat, vendor: inv.vendor, amount: 0, done: 0, n: 0 };
      o.amount = round2_(o.amount + a.amount); o.n++;
      if (a.status === '已核銷') o.done = round2_(o.done + a.amount);
    });
  });
  return { cats: cats, rows: Object.keys(rows).map(function (k) { return rows[k]; }) };
}

// ---------------------------------------------------------------------
// 試算表小工具（跟訂貨系統一樣，照表頭名稱讀寫）
// ---------------------------------------------------------------------

let SS_CACHE_ = null;
function ss_() {
  if (!SS_CACHE_) SS_CACHE_ = retry_(function () { return SpreadsheetApp.openById(SPREADSHEET_ID); });
  return SS_CACHE_;
}

// 同一次動作裡同一個分頁只讀一次（月結、核銷會一週一週重複讀，讀太多次 Google 會回 RESOURCE_EXHAUSTED）；有寫入就清掉重讀
let TABLE_MEMO_ = {};
function tableMemoClear_() { TABLE_MEMO_ = {}; SLIPS_MEMO_ = null; }
function readTable_(name, required) {
  const hit = TABLE_MEMO_[name];
  if (hit) {
    (required || []).forEach(function (k) { if (hit.cols[k] === undefined) throw new Error('分頁「' + name + '」缺少「' + k + '」欄位。'); });
    return { sheet: hit.sheet, cols: hit.cols, rows: hit.rows.map(function (o) { return Object.assign({}, o); }) };
  }
  const t = readTableNow_(name, required);
  TABLE_MEMO_[name] = t;
  return { sheet: t.sheet, cols: t.cols, rows: t.rows.map(function (o) { return Object.assign({}, o); }) };
}
function readTableNow_(name, required) {
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
  tableMemoClear_();
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
  tableMemoClear_();
  if (!objs.length) return;
  const sh = ss_().getSheetByName(name);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const data = objs.map(function (o) { return headers.map(function (h) { return o[h] === undefined ? '' : o[h]; }); });
  const start = sh.getLastRow() + 1;
  retry_(function () { sh.getRange(start, 1, data.length, headers.length).setValues(data); });
}

function setCell_(table, row, header, value) {
  tableMemoClear_();
  const c = table.cols[header];
  if (c === undefined) throw new Error('分頁缺少「' + header + '」欄位。');
  table.sheet.getRange(row, c).setValue(value);
}

/** 從下往上刪，列號才不會跑掉 */
function deleteRows_(sh, rows) {
  tableMemoClear_();
  // 連在一起的列一次刪（一列一列刪很慢，存檔時會一直佔著鎖）
  const list = rows.slice().sort(function (a, b) { return b - a; });
  for (let i = 0; i < list.length;) {
    let j = i;
    while (j + 1 < list.length && list[j + 1] === list[j] - 1) j++;
    sh.deleteRows(list[j], j - i + 1);
    i = j + 1;
  }
}

function retry_(fn) {
  let lastErr;
  for (let i = 0; i < 3; i++) {
    try { return fn(); } catch (e) {
      lastErr = e;
      const m = String(e && e.message || e);
      if (!/Service|服務|無法正常運作|timed out|逾時|Internal error|內部錯誤|try again|稍後再試|RESOURCE_EXHAUSTED|儲存空間|storage/i.test(m)) throw e;
      Utilities.sleep((/RESOURCE_EXHAUSTED|儲存空間|storage/i.test(m) ? 2000 : 800) * (i + 1));
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
