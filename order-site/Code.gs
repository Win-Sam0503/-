/**
 * 老師訂貨網站：Google Apps Script 後端
 *
 * 這支程式掛在「訂貨與庫存管理系統」試算表上（擴充功能 > Apps Script），
 * 網站讀取「商品主檔」顯示商品，老師送出的訂單寫入「客戶訂單」，
 * 新老師自動加入「客戶」分頁。部署方式見 README.md。
 */

// 若程式不是從試算表內建立的，請填入試算表 ID（網址 /d/ 與 /edit 之間那段）。
const SPREADSHEET_ID = '';

const SITE_TITLE = '老師訂貨網站';
const MAX_QTY_PER_ITEM = 999;
const MAX_ITEMS_PER_ORDER = 50;
const MY_ORDERS_LIMIT = 20;

const SHEET = {
  products: '商品主檔',
  orders: '客戶訂單',
  customers: '客戶',
};

// 商品主檔欄位（從 0 起算）
const P = { id: 0, name: 1, category: 2, unit: 3, price: 6, safety: 8, available: 15, listed: 18 };
// 客戶訂單欄位（從 1 起算，供 getRange 使用）
const O = { no: 1, date: 2, customer: 3, productId: 4, name: 5, unit: 6, qty: 7, price: 8, amount: 9, status: 10, shipDate: 11, note: 12, phone: 13, wantDate: 14 };
// 客戶欄位（從 1 起算）
const C = { code: 1, name: 2, contact: 3, phone: 4, address: 5, payment: 6, note: 7 };

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(SITE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('找不到分頁「' + name + '」，請確認試算表分頁名稱沒有被修改。');
  return sh;
}

/** 回傳網站上架的商品清單。 */
function getProducts() {
  return readProducts_()
    .filter(function (p) { return p.listed; })
    .map(function (p) {
      return { id: p.id, name: p.name, category: p.category, unit: p.unit, price: p.price, stock: p.stock };
    });
}

function readProducts_() {
  const sh = sheet_(SHEET.products);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, P.listed + 1).getValues();
  const out = [];
  rows.forEach(function (r) {
    const id = String(r[P.id]).trim();
    if (!id) return;
    const available = Number(r[P.available]) || 0;
    const safety = Number(r[P.safety]) || 0;
    let stock = '有貨';
    if (available <= 0) stock = '需調貨';
    else if (available < safety) stock = '庫存少';
    out.push({
      id: id,
      name: String(r[P.name]),
      category: String(r[P.category] || '其他'),
      unit: String(r[P.unit]),
      price: Number(r[P.price]) || 0,
      stock: stock,
      listed: r[P.listed] === true,
    });
  });
  return out;
}

/**
 * 建立訂單。
 * payload: { name, school, phone, address, wantDate, note, items: [{ id, qty }] }
 */
function submitOrder(payload) {
  const order = validateOrder_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const customerName = upsertCustomer_(order);
    const orderNo = nextOrderNo_();
    writeOrderLines_(orderNo, customerName, order);
    SpreadsheetApp.flush();
    const total = order.items.reduce(function (s, it) { return s + it.qty * it.price; }, 0);
    return { orderNo: orderNo, customer: customerName, total: total, items: order.items };
  } finally {
    lock.releaseLock();
  }
}

function validateOrder_(p) {
  if (!p || typeof p !== 'object') throw new Error('訂單資料不完整。');
  const name = clean_(p.name, 30);
  const school = clean_(p.school, 40);
  const phone = digits_(p.phone);
  const address = clean_(p.address, 100);
  const note = clean_(p.note, 200);
  if (!name) throw new Error('請填寫姓名。');
  if (phone.length < 8 || phone.length > 10) throw new Error('請填寫正確的聯絡電話。');
  if (!address) throw new Error('請填寫送貨地址。');

  let wantDate = null;
  if (p.wantDate) {
    const m = String(p.wantDate).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) throw new Error('希望到貨日格式不正確。');
    wantDate = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (wantDate < today_()) throw new Error('希望到貨日不能早於今天。');
  }

  if (!Array.isArray(p.items) || p.items.length === 0) throw new Error('購物車是空的。');
  if (p.items.length > MAX_ITEMS_PER_ORDER) throw new Error('一張訂單最多 ' + MAX_ITEMS_PER_ORDER + ' 個品項。');

  const catalog = {};
  readProducts_().forEach(function (prod) { if (prod.listed) catalog[prod.id] = prod; });
  const merged = {};
  p.items.forEach(function (it) {
    const id = String(it && it.id || '').trim();
    const qty = Number(it && it.qty);
    const prod = catalog[id];
    if (!prod) throw new Error('商品「' + id + '」目前無法訂購，請重新整理頁面。');
    if (!Number.isInteger(qty) || qty < 1 || qty > MAX_QTY_PER_ITEM) {
      throw new Error('「' + prod.name + '」的數量需為 1 到 ' + MAX_QTY_PER_ITEM + ' 的整數。');
    }
    merged[id] = (merged[id] || 0) + qty;
  });
  const items = Object.keys(merged).map(function (id) {
    const prod = catalog[id];
    return { id: id, name: prod.name, unit: prod.unit, price: prod.price, qty: merged[id] };
  });

  return { name: name, school: school, phone: phone, address: address, note: note, wantDate: wantDate, items: items };
}

/** 依電話找老師，找不到就新增到「客戶」分頁。回傳客戶名稱。 */
function upsertCustomer_(order) {
  const sh = sheet_(SHEET.customers);
  const last = sh.getLastRow();
  const rows = last >= 2 ? sh.getRange(2, 1, last - 1, C.note).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    if (digits_(rows[i][C.phone - 1]) === order.phone && rows[i][C.name - 1]) {
      // 地址有變就更新
      if (order.address && rows[i][C.address - 1] !== order.address) {
        sh.getRange(i + 2, C.address).setValue(order.address);
      }
      return String(rows[i][C.name - 1]);
    }
  }
  const displayName = order.school ? order.name + '老師（' + order.school + '）' : order.name + '老師';
  let maxT = 0;
  rows.forEach(function (r) {
    const m = String(r[C.code - 1]).match(/^T(\d+)$/);
    if (m) maxT = Math.max(maxT, Number(m[1]));
  });
  const code = 'T' + ('000' + (maxT + 1)).slice(-3);
  const row = lastRowInColumn_(sh, C.code) + 1;
  sh.getRange(row, C.phone).setNumberFormat('@');
  sh.getRange(row, 1, 1, C.note).setValues([[code, displayName, order.name, order.phone, order.address, '', '網站訂貨自動建立']]);
  return displayName;
}

function nextOrderNo_() {
  const prefix = 'SO-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-';
  const sh = sheet_(SHEET.orders);
  const last = lastRowInColumn_(sh, O.no);
  let max = 0;
  if (last >= 2) {
    sh.getRange(2, O.no, last - 1, 1).getValues().forEach(function (r) {
      const s = String(r[0]);
      if (s.indexOf(prefix) === 0) max = Math.max(max, Number(s.slice(prefix.length)) || 0);
    });
  }
  return prefix + ('00' + (max + 1)).slice(-3);
}

/**
 * 寫入訂單明細。品名、單位、單價、金額由試算表第 2 列的 ARRAYFORMULA 自動帶出，
 * 所以這裡只寫輸入欄位，跳過 E、F、H、I 欄，避免蓋掉公式。
 */
function writeOrderLines_(orderNo, customerName, order) {
  const sh = sheet_(SHEET.orders);
  const start = lastRowInColumn_(sh, O.no) + 1;
  const n = order.items.length;
  const needRows = start + n - 1 - sh.getMaxRows();
  if (needRows > 0) sh.insertRowsAfter(sh.getMaxRows(), needRows);

  const date = today_();
  const school = order.school ? '學校/單位：' + order.school + '；' : '';
  const note = '網站訂單；' + school + '地址：' + order.address + (order.note ? '；備註：' + order.note : '');

  sh.getRange(start, O.no, n, 4).setValues(order.items.map(function (it) {
    return [orderNo, date, customerName, it.id];
  }));
  sh.getRange(start, O.qty, n, 1).setValues(order.items.map(function (it) { return [it.qty]; }));
  sh.getRange(start, O.phone, n, 1).setNumberFormat('@');
  sh.getRange(start, O.status, n, 5).setValues(order.items.map(function () {
    return ['待處理', '', note, order.phone, order.wantDate || ''];
  }));
}

/** 用電話 + 姓名查詢自己的訂單。 */
function getMyOrders(phone, name) {
  const ph = digits_(phone);
  const nm = clean_(name, 30);
  if (ph.length < 8 || !nm) throw new Error('請輸入姓名與下單時的聯絡電話。');

  const sh = sheet_(SHEET.orders);
  const last = lastRowInColumn_(sh, O.no);
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, O.wantDate).getValues();
  const tz = Session.getScriptTimeZone();
  const byNo = {};
  const order = [];
  rows.forEach(function (r) {
    if (digits_(r[O.phone - 1]) !== ph) return;
    if (String(r[O.customer - 1]).indexOf(nm) !== 0) return;
    const no = String(r[O.no - 1]);
    if (!byNo[no]) {
      byNo[no] = {
        orderNo: no,
        date: fmtDate_(r[O.date - 1], tz),
        wantDate: fmtDate_(r[O.wantDate - 1], tz),
        items: [],
        total: 0,
        statuses: [],
      };
      order.push(no);
    }
    const o = byNo[no];
    const status = String(r[O.status - 1]);
    o.items.push({
      name: String(r[O.name - 1]),
      unit: String(r[O.unit - 1]),
      qty: Number(r[O.qty - 1]) || 0,
      amount: Number(r[O.amount - 1]) || 0,
      status: status,
      shipDate: fmtDate_(r[O.shipDate - 1], tz),
    });
    if (status !== '已取消') o.total += Number(r[O.amount - 1]) || 0;
    o.statuses.push(status);
  });

  return order.reverse().slice(0, MY_ORDERS_LIMIT).map(function (no) {
    const o = byNo[no];
    const s = o.statuses;
    o.status = s.every(function (x) { return x === '已取消'; }) ? '已取消'
      : s.some(function (x) { return x === '待處理'; }) ? '處理中'
      : '已出貨';
    delete o.statuses;
    return o;
  });
}

// ---- 小工具 ----

function lastRowInColumn_(sh, col) {
  const maxRows = sh.getMaxRows();
  const values = sh.getRange(1, col, maxRows, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) return i + 1;
  }
  return 0;
}

function clean_(v, maxLen) {
  // 去掉開頭的 = + - @，避免被試算表當成公式執行
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().replace(/^[=+\-@]+/, '').slice(0, maxLen);
}

function digits_(v) {
  return String(v == null ? '' : v).replace(/\D/g, '');
}

function today_() {
  const s = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd').split('-');
  return new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]));
}

function fmtDate_(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy/MM/dd') : String(v || '');
}
