/**
 * 餐飲科老師訂貨網站：Google Apps Script 後端
 *
 * 資料全部存在試算表「訂貨表2.0」：
 *   帳號設定   老師帳號、密碼、啟用
 *   登入資訊   每次登入的紀錄
 *   食材總表   可訂購的食材
 *   廠商設定   廠商的送貨組別與星期代碼
 *   訂單總表   老師送出的訂單，一個食材一列
 *   待審核食材 老師提交、食材總表裡沒有的食材
 *
 * 程式依「標題名稱」找欄位，欄位順序調整不影響，但請不要改標題文字。
 * 部署方式見 README.md。
 */

// 試算表 ID（網址 /d/ 與 /edit 之間那段）。
const SPREADSHEET_ID = '1qIQe4STafrhzQk4I7NLkZQxJGe5hozf0g04zvJ-qZ_I';

const SITE_TITLE = '餐飲科訂貨系統';
const ADMIN_ACCOUNTS = ['管理者'];          // 這些帳號登入後會看到管理功能
const DEPARTMENTS = ['中餐', '西餐', '烘焙', '飲調', '果雕', '刀功'];
// 上課時間格式：「上午／下午 幾點:幾分」，分鐘每 5 分鐘，例如「下午 1:05」
const CLASS_TIME_PATTERN = /^(上午|下午) (1[0-2]|[1-9]):[0-5][05]$/;
// 鎖單：每週一 12:00 截止「下週」的訂單；截止後最早只能訂下下週的課。管理者不受限制。
const LOCK_WEEKDAY = 1;                     // 截止星期（1 = 星期一 … 5 = 星期五）
const LOCK_HOUR = 12;
const LOCK_MINUTE = 0;
const TOOL_ID = '器具';                     // 使用器具在訂單總表的「食材ID」與「廠商」
const ITEMS_CACHE_MINUTES = 10;             // 食材清單暫存時間（加快登入）；核准新食材時會立即更新
const MAX_DAYS_AHEAD = 30;                  // 上課日期最晚可選幾天後
const MAX_ITEMS_PER_ORDER = 80;
const SESSION_HOURS = 6;                    // 登入有效時間
const MAX_LOGIN_FAILS = 5;                  // 連續輸錯幾次就暫停登入
const LOCK_MINUTES = 10;
const ORDER_STATUSES = ['待處理', '已叫貨', '已到貨', '已取消'];
const PENDING_ID = '待審核';

const SHEET = {
  accounts: '帳號設定',
  logins: '登入資訊',
  items: '食材總表',
  vendors: '廠商設定',
  orders: '訂單總表',
  pending: '待審核食材',
};

const WEEKDAY = ['日', '一', '二', '三', '四', '五', '六'];
// 試算表標題的別名：訂單總表用「使用日期」或「上課日期」都可以
const HEADER_ALIASES = { '上課日期': '使用日期', '使用日期': '上課日期' };

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(SITE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =====================================================================
// 登入
// =====================================================================

function login(account, password) {
  account = clean_(account, 30);
  password = String(password == null ? '' : password);
  if (!account || !password) throw new Error('請輸入帳號和密碼。');

  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + account;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_LOGIN_FAILS) {
    throw new Error('密碼錯誤次數過多，請 ' + LOCK_MINUTES + ' 分鐘後再試。');
  }

  const acc = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']).rows
    .find(function (r) { return String(r['帳號']).trim() === account; });
  const ok = acc && isTrue_(acc['啟用']) && String(acc['密碼']) === password;
  logLogin_(account, ok ? '成功' : '失敗');

  if (!ok) {
    cache.put(failKey, String(fails + 1), LOCK_MINUTES * 60);
    throw new Error('帳號或密碼不正確，或帳號未啟用。');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid();
  cache.put('session:' + token, account, SESSION_HOURS * 3600);
  return { token: token, account: account, isAdmin: ADMIN_ACCOUNTS.indexOf(account) >= 0 };
}

/** 登入並一次載入網站資料（少一次來回，登入比較快）。 */
function loginAndLoad(account, password) {
  const res = login(account, password);
  res.data = getBootstrap(res.token);
  return res;
}

function logout(token) {
  CacheService.getScriptCache().remove('session:' + token);
  return true;
}

function logLogin_(account, result) {
  try {
    appendRows_(SHEET.logins, [{ '登入時間': new Date(), '帳號': account, '結果': result }],
      { '登入時間': 'yyyy/mm/dd hh:mm:ss' });
  } catch (e) {
    console.error('登入紀錄寫入失敗：' + e.message);
  }
}

/** 檢查登入狀態，回傳帳號名稱。 */
function auth_(token, needAdmin) {
  const account = token ? CacheService.getScriptCache().get('session:' + token) : null;
  if (!account) throw new Error('SESSION_EXPIRED');
  if (needAdmin && ADMIN_ACCOUNTS.indexOf(account) < 0) throw new Error('只有管理者可以使用這個功能。');
  return account;
}

// =====================================================================
// 老師：讀取資料、下單、查詢、取消
// =====================================================================

/** 登入後一次載入網站需要的資料。 */
function getBootstrap(token) {
  const account = auth_(token);
  const items = cachedItems_();
  const vendors = [];
  items.forEach(function (it) { if (it.vendor && vendors.indexOf(it.vendor) < 0) vendors.push(it.vendor); });
  vendors.sort();

  const teachers = readTable_(SHEET.accounts, ['帳號', '啟用']).rows
    .filter(function (r) { return isTrue_(r['啟用']); })
    .map(function (r) { return String(r['帳號']).trim(); })
    .filter(function (a) { return a && ADMIN_ACCOUNTS.indexOf(a) < 0; });

  const depts = DEPARTMENTS.slice();
  items.forEach(function (it) { if (it.dept !== '通用' && depts.indexOf(it.dept) < 0) depts.push(it.dept); });

  const isAdmin = ADMIN_ACCOUNTS.indexOf(account) >= 0;
  const now = new Date();
  return {
    account: account,
    isAdmin: isAdmin,
    items: items,
    vendors: vendors,
    teachers: teachers,
    departments: depts,
    minDate: dateStr_(isAdmin ? today_() : earliestClassDate_(now)),
    maxDate: dateStr_(addDays_(today_(), MAX_DAYS_AHEAD)),
    lockRule: lockRuleText_(),
    statuses: ORDER_STATUSES,
  };
}

/** 食材清單（暫存在伺服器，避免每次登入都讀整張食材總表）。 */
function cachedItems_() {
  const cache = CacheService.getScriptCache();
  const meta = cache.get('items:n');
  if (meta) {
    const keys = [];
    for (let i = 0; i < Number(meta); i++) keys.push('items:' + i);
    const parts = cache.getAll(keys);
    if (keys.every(function (k) { return parts[k] != null; })) {
      try { return JSON.parse(keys.map(function (k) { return parts[k]; }).join('')); } catch (e) {}
    }
  }
  const items = readItems_();
  const json = JSON.stringify(items);
  const CHUNK = 90000;   // 每個暫存值上限 100KB
  const put = {};
  let n = 0;
  for (let i = 0; i < json.length; i += CHUNK) put['items:' + (n++)] = json.slice(i, i + CHUNK);
  put['items:n'] = String(n);
  try { cache.putAll(put, ITEMS_CACHE_MINUTES * 60); } catch (e) { console.error(e.message); }
  return items;
}

function clearItemsCache_() {
  CacheService.getScriptCache().remove('items:n');
}

function readItems_() {
  return readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']).rows
    .filter(function (r) {
      return String(r['食材ID']).trim() && (r['啟用'] === undefined || isTrue_(r['啟用']));
    })
    .map(function (r) {
      return {
        id: String(r['食材ID']).trim(),
        name: String(r['食材名稱']).trim(),
        alias: String(r['別名'] || '').trim(),
        vendor: String(r['廠商'] || '').trim(),
        dept: String(r['科別'] || '通用').trim() || '通用',
        unit: String(r['單位'] || '').trim(),
        group: String(r['適用組別'] || '全部').trim() || '全部',
      };
    });
}

function lockRuleText_() {
  return '每週' + WEEKDAY[LOCK_WEEKDAY] + ' ' + LOCK_HOUR + ':' + ('0' + LOCK_MINUTE).slice(-2) + ' 截止下週的訂單';
}

/**
 * 送出訂單。
 * payload: {
 *   useDate: 'yyyy-MM-dd', dept, classTime, room, className, project, onBehalfOf,
 *   items: [{ id, qty, unit, note }],                       // 食材總表內的食材
 *   newItems: [{ name, vendor, qty, unit, note }]           // 找不到、需審核的食材
 * }
 */
function submitOrder(token, payload) {
  const account = auth_(token);
  const o = validateOrder_(payload, account, ADMIN_ACCOUNTS.indexOf(account) >= 0);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const orderNo = nextOrderNo_();
    const now = new Date();
    const base = {
      '訂單編號': orderNo,
      '送出時間': now,
      '使用日期': o.useDate,
      '星期': '星期' + WEEKDAY[o.useDate.getDay()],
      '操作老師': account,
      '代訂對象老師': o.onBehalfOf,
      '科別': o.dept,
      '上課時間': o.classTime,
      '上課教室': o.room,
      '上課班級': o.className,
      '實作項目': o.project,
      '狀態': '待處理',
    };
    const rows = o.lines.map(function (l) {
      return Object.assign({}, base, {
        '食材ID': l.id, '食材名稱': l.name, '廠商': l.vendor, '送貨組別': l.group,
        '數量': l.qty, '單位': l.unit, '備註': l.note,
      });
    });
    appendRows_(SHEET.orders, rows, { '送出時間': 'yyyy/mm/dd hh:mm:ss', '使用日期': 'yyyy-mm-dd' });

    const pending = o.lines.filter(function (l) { return l.id === PENDING_ID; }).map(function (l) {
      return {
        '新增時間': now, '提交老師': account, '食材名稱': l.name, '廠商': l.vendor,
        '單位': l.unit, '對應訂單編號': orderNo, '審核狀態': '待審核',
      };
    });
    if (pending.length) appendRows_(SHEET.pending, pending, { '新增時間': 'yyyy/mm/dd hh:mm:ss' });

    SpreadsheetApp.flush();
    const tools = o.lines.filter(function (l) { return l.id === TOOL_ID; }).length;
    return { orderNo: orderNo, useDate: dateStr_(o.useDate), count: o.lines.length - tools, tools: tools, pending: pending.length };
  } finally {
    lock.releaseLock();
  }
}

function validateOrder_(p, account, isAdmin) {
  if (!p || typeof p !== 'object') throw new Error('訂單資料不完整。');
  const useDate = parseDate_(p.useDate);
  if (!useDate) throw new Error('請選擇上課日期。');
  const min = isAdmin ? today_() : earliestClassDate_(new Date());
  const max = addDays_(today_(), MAX_DAYS_AHEAD);
  if (useDate < min) {
    throw new Error(isAdmin ? '上課日期不能早於今天。'
      : '已過訂貨截止（' + lockRuleText_() + '），目前最早只能訂 ' + dateStr_(min) + ' 的課。');
  }
  if (useDate > max) throw new Error('上課日期最晚只能選 ' + dateStr_(max) + '（30 天內）。');
  const dow = useDate.getDay();
  if (dow === 0 || dow === 6) throw new Error('上課日期是週末，廠商不送貨，請改選平日。');

  const dept = clean_(p.dept, 20);
  const classTime = clean_(p.classTime, 12);
  const className = clean_(p.className, 30);
  const project = clean_(p.project, 60);
  if (!dept) throw new Error('請選擇科別。');
  if (!CLASS_TIME_PATTERN.test(classTime)) throw new Error('請選擇上課時間（上午／下午、幾點、幾分）。');
  if (!className) throw new Error('請填寫上課班級。');
  if (!project) throw new Error('請填寫實作項目。');

  let onBehalfOf = clean_(p.onBehalfOf, 30);
  if (onBehalfOf === account) onBehalfOf = '';
  if (onBehalfOf) {
    const valid = readTable_(SHEET.accounts, ['帳號', '啟用']).rows.some(function (r) {
      return String(r['帳號']).trim() === onBehalfOf && isTrue_(r['啟用']);
    });
    if (!valid) throw new Error('代訂對象「' + onBehalfOf + '」不在帳號名單中。');
  }

  const catalog = {};
  readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']).rows.forEach(function (r) {
    const id = String(r['食材ID']).trim();
    if (id && (r['啟用'] === undefined || isTrue_(r['啟用']))) catalog[id] = r;
  });
  const vendorGroups = readVendorGroups_();

  const lines = [];
  (Array.isArray(p.items) ? p.items : []).forEach(function (it) {
    const id = String(it && it.id || '').trim();
    const r = catalog[id];
    if (!r) throw new Error('食材「' + id + '」已停用或不存在，請重新整理頁面。');
    const name = String(r['食材名稱']).trim();
    const vendor = String(r['廠商']).trim();
    lines.push({
      id: id, name: name, vendor: vendor,
      qty: qty_(it.qty, name), unit: unit_(it.unit, r['單位'], name), note: clean_(it.note, 100),
      group: deliveryGroup_(vendor, dow, vendorGroups),
    });
  });
  (Array.isArray(p.newItems) ? p.newItems : []).forEach(function (it) {
    const name = clean_(it && it.name, 40);
    const vendor = clean_(it && it.vendor, 20);
    if (!name) throw new Error('新增食材請填寫名稱。');
    if (!vendor) throw new Error('請選擇「' + name + '」的廠商。');
    lines.push({
      id: PENDING_ID, name: name, vendor: vendor,
      qty: qty_(it.qty, name), unit: unit_(it.unit, '', name), note: clean_(it.note, 100),
      group: deliveryGroup_(vendor, dow, vendorGroups),
    });
  });

  (Array.isArray(p.tools) ? p.tools : []).forEach(function (t) {
    const name = clean_(t && t.name, 40);
    if (!name) throw new Error('請填寫使用器具的名稱。');
    lines.push({
      id: TOOL_ID, name: name, vendor: TOOL_ID, group: '',
      qty: qty_(t.qty, name), unit: '', note: clean_(t.note, 100),
    });
  });

  if (!lines.length) throw new Error('請至少選擇一項食材或器具。');
  if (lines.length > MAX_ITEMS_PER_ORDER) throw new Error('一張訂單最多 ' + MAX_ITEMS_PER_ORDER + ' 項食材。');

  return {
    useDate: useDate, dept: dept, classTime: classTime, room: clean_(p.room, 30),
    className: className, project: project, onBehalfOf: onBehalfOf, lines: lines,
  };
}

function qty_(v, name) {
  const n = Number(v);
  if (!isFinite(n) || n <= 0 || n > 9999) throw new Error('「' + name + '」的數量需大於 0。');
  return Math.round(n * 100) / 100;
}

function unit_(v, fallback, name) {
  const u = clean_(v, 10) || String(fallback || '').trim();
  if (!u) throw new Error('請填寫「' + name + '」的單位。');
  return u;
}

/** 讀「廠商設定」：{ 廠商: [星期代碼, ...] } */
function readVendorGroups_() {
  const map = {};
  try {
    readTable_(SHEET.vendors, ['廠商', '星期代碼']).rows.forEach(function (r) {
      const v = String(r['廠商']).trim();
      const code = String(r['星期代碼']).trim();
      if (!v || !code) return;
      (map[v] = map[v] || []).push(code);
    });
  } catch (e) {
    console.error(e.message);
  }
  return map;
}

/** 依上課日期的星期決定送貨組別；廠商設定沒有列的廠商，週一二為「一二」、週三四五為「三四五」。 */
function deliveryGroup_(vendor, dow, vendorGroups) {
  const ch = WEEKDAY[dow];
  const codes = vendorGroups[vendor] || [];
  for (let i = 0; i < codes.length; i++) {
    if (codes[i].indexOf(ch) >= 0) return codes[i];
  }
  return dow <= 2 ? '一二' : '三四五';
}

function nextOrderNo_() {
  // 編號是送出的秒數；同一秒有兩張訂單時，等到下一秒再編號（呼叫端已持有鎖）
  const props = PropertiesService.getScriptProperties();
  const last = props.getProperty('lastOrderNo');
  let no = 'O' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss');
  for (let i = 0; i < 20 && no <= (last || ''); i++) {
    Utilities.sleep(200);
    no = 'O' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss');
  }
  if (no <= (last || '')) throw new Error('系統忙碌中，請稍後再送出一次。');
  props.setProperty('lastOrderNo', no);
  return no;
}

/** 我送出的、或別人代我訂的訂單（最近 60 天的上課日期起）。 */
function getMyOrders(token) {
  const account = auth_(token);
  const since = addDays_(today_(), -60);
  const rows = readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows.filter(function (r) {
    if (String(r['操作老師']).trim() !== account && String(r['代訂對象老師'] || '').trim() !== account) return false;
    const d = toDate_(r['使用日期']);
    return !d || d >= since;
  });
  return groupOrders_(rows, account).sort(function (a, b) { return a.orderNo < b.orderNo ? 1 : -1; });
}

function cancelOrder(token, orderNo) {
  const account = auth_(token);
  const isAdmin = ADMIN_ACCOUNTS.indexOf(account) >= 0;
  return updateOrderStatus_(orderNo, '已取消', function (rows) {
    if (isAdmin) return;
    const r = rows[0];
    if (String(r['操作老師']).trim() !== account && String(r['代訂對象老師'] || '').trim() !== account) {
      throw new Error('只能取消自己的訂單。');
    }
    if (rows.some(function (x) { return String(x['狀態']).trim() !== '待處理'; })) {
      throw new Error('這張訂單已經在處理，無法自行取消，請聯絡管理者。');
    }
    if (isLocked_(toDate_(r['使用日期']))) {
      throw new Error('這張訂單的上課週已過訂貨截止時間，無法取消，請聯絡管理者。');
    }
  });
}

// =====================================================================
// 管理者
// =====================================================================

/** 指定上課日期的所有訂單。 */
function adminGetOrders(token, from, to) {
  auth_(token, true);
  let start = parseDate_(from);
  let end = parseDate_(to) || start;
  if (!start) throw new Error('請選擇日期。');
  if (end < start) { const t = start; start = end; end = t; }
  const rows = readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows.filter(function (r) {
    const d = toDate_(r['使用日期']);
    return d && d >= start && d <= end;
  });
  return groupOrders_(rows, null).sort(function (a, b) {
    return a.useDate === b.useDate ? (a.orderNo < b.orderNo ? -1 : 1) : (a.useDate < b.useDate ? -1 : 1);
  });
}

function adminSetStatus(token, orderNo, status) {
  auth_(token, true);
  if (ORDER_STATUSES.indexOf(status) < 0) throw new Error('狀態不正確。');
  return updateOrderStatus_(orderNo, status, null);
}

function adminGetPending(token) {
  auth_(token, true);
  return readTable_(SHEET.pending, ['食材名稱', '審核狀態']).rows
    .filter(function (r) { const s = String(r['審核狀態'] || '').trim(); return !s || s === '待審核'; })
    .map(function (r) {
      return {
        row: r._row,
        time: fmtDateTime_(r['新增時間']),
        teacher: String(r['提交老師'] || ''),
        name: String(r['食材名稱'] || ''),
        vendor: String(r['廠商'] || ''),
        unit: String(r['單位'] || ''),
        orderNo: String(r['對應訂單編號'] || ''),
      };
    });
}

/**
 * 審核新食材。
 * decision: 'approve' 加入食材總表；'map' 對應到既有食材（edits.existingId）；'reject' 退回。
 * edits: { name, vendor, unit, existingId }
 */
function adminReviewItem(token, row, decision, edits) {
  auth_(token, true);
  edits = edits || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const pend = readTable_(SHEET.pending, ['食材名稱', '審核狀態']);
    const rec = pend.rows.find(function (r) { return r._row === Number(row); });
    if (!rec) throw new Error('找不到這筆待審核資料，請重新整理。');
    const status = String(rec['審核狀態'] || '').trim();
    if (status && status !== '待審核') throw new Error('這筆資料已經審核過了。');
    const origName = String(rec['食材名稱']).trim();
    const orderNo = String(rec['對應訂單編號'] || '').trim();
    let result;

    if (decision === 'reject') {
      result = '已退回';
    } else if (decision === 'approve') {
      const name = clean_(edits.name, 40) || origName;
      const vendor = clean_(edits.vendor, 20) || String(rec['廠商'] || '').trim();
      const unit = clean_(edits.unit, 10) || String(rec['單位'] || '').trim();
      if (!vendor) throw new Error('請選擇廠商。');
      const items = readTable_(SHEET.items, ['食材ID', '食材名稱']);
      if (items.rows.some(function (r) { return String(r['食材名稱']).trim() === name; })) {
        throw new Error('食材總表已經有「' + name + '」，請改用「對應既有食材」。');
      }
      let max = 0;
      items.rows.forEach(function (r) {
        const m = String(r['食材ID']).match(/^I(\d+)$/);
        if (m) max = Math.max(max, Number(m[1]));
      });
      const id = 'I' + ('0000' + (max + 1)).slice(-4);
      appendRows_(SHEET.items, [{
        '食材ID': id, '食材名稱': name, '別名': name === origName ? '' : origName, '廠商': vendor,
        '科別': '通用', '單位': unit, '適用組別': '全部', '啟用': true,
      }]);
      clearItemsCache_();
      fixOrderLines_(orderNo, origName, { id: id, name: name, vendor: vendor });
      result = '已核准（' + id + '）';
    } else if (decision === 'map') {
      const id = String(edits.existingId || '').trim();
      const r = readTable_(SHEET.items, ['食材ID', '食材名稱']).rows.find(function (x) { return String(x['食材ID']).trim() === id; });
      if (!r) throw new Error('找不到食材 ' + id + '。');
      fixOrderLines_(orderNo, origName, { id: id, name: String(r['食材名稱']).trim(), vendor: String(r['廠商'] || '').trim() });
      result = '已對應（' + id + '）';
    } else {
      throw new Error('審核動作不正確。');
    }

    setCell_(pend, rec._row, '審核狀態', result);
    return result;
  } finally {
    lock.releaseLock();
  }
}

/** 把訂單中「待審核」的食材換成正式的食材 ID、名稱、廠商。 */
function fixOrderLines_(orderNo, origName, to) {
  if (!orderNo) return;
  const t = readTable_(SHEET.orders, ['訂單編號', '食材ID', '食材名稱']);
  const vendorGroups = readVendorGroups_();
  t.rows.forEach(function (r) {
    if (String(r['訂單編號']).trim() !== orderNo) return;
    if (String(r['食材ID']).trim() !== PENDING_ID || String(r['食材名稱']).trim() !== origName) return;
    setCell_(t, r._row, '食材ID', to.id);
    setCell_(t, r._row, '食材名稱', to.name);
    if (to.vendor && t.cols['廠商'] !== undefined) {
      setCell_(t, r._row, '廠商', to.vendor);
      const d = toDate_(r['使用日期']);
      if (d && t.cols['送貨組別'] !== undefined) setCell_(t, r._row, '送貨組別', deliveryGroup_(to.vendor, d.getDay(), vendorGroups));
    }
  });
}

// =====================================================================
// 共用：訂單
// =====================================================================

function groupOrders_(rows, account) {
  const map = {};
  const list = [];
  rows.forEach(function (r) {
    const no = String(r['訂單編號']).trim();
    if (!no) return;
    if (!map[no]) {
      map[no] = {
        orderNo: no,
        submitted: fmtDateTime_(r['送出時間']),
        useDate: dateStr_(toDate_(r['使用日期'])),
        weekday: String(r['星期'] || ''),
        teacher: String(r['操作老師'] || ''),
        onBehalfOf: String(r['代訂對象老師'] || ''),
        dept: String(r['科別'] || ''),
        classTime: String(r['上課時間'] || ''),
        room: String(r['上課教室'] || ''),
        className: String(r['上課班級'] || ''),
        project: String(r['實作項目'] || ''),
        items: [],
      };
      list.push(map[no]);
    }
    map[no].items.push({
      id: String(r['食材ID'] || ''),
      name: String(r['食材名稱'] || ''),
      vendor: String(r['廠商'] || ''),
      group: String(r['送貨組別'] || ''),
      qty: Number(r['數量']) || 0,
      unit: String(r['單位'] || ''),
      note: String(r['備註'] || ''),
      status: String(r['狀態'] || ''),
    });
  });
  list.forEach(function (o) {
    const s = o.items.map(function (i) { return i.status; });
    o.status = s.every(function (x) { return x === s[0]; }) ? s[0] : '部分處理';
    o.locked = isLocked_(parseDate_(o.useDate));
    o.canCancel = !!account && o.status === '待處理' && !o.locked;
  });
  return list;
}

function updateOrderStatus_(orderNo, status, check) {
  orderNo = String(orderNo || '').trim();
  if (!orderNo) throw new Error('缺少訂單編號。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '狀態']);
    const rows = t.rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo; });
    if (!rows.length) throw new Error('找不到訂單 ' + orderNo + '。');
    if (check) check(rows);
    rows.forEach(function (r) { setCell_(t, r._row, '狀態', status); });
    SpreadsheetApp.flush();
    return status;
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// 共用：試算表讀寫（依標題找欄位）
// =====================================================================

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('找不到分頁「' + name + '」，請確認分頁名稱。');
  return sh;
}

/** 讀整張表，回傳 { sheet, cols: {標題: 欄號(1起)}, rows: [{標題: 值, _row: 列號}] } */
function readTable_(name, required) {
  const sh = sheet_(name);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const values = lastRow ? sh.getRange(1, 1, lastRow, lastCol).getValues() : [[]];
  const cols = {};
  values[0].forEach(function (h, i) {
    const k = String(h).trim();
    if (k && cols[k] === undefined) cols[k] = i + 1;
  });
  Object.keys(HEADER_ALIASES).forEach(function (k) {
    if (cols[k] === undefined && cols[HEADER_ALIASES[k]] !== undefined) cols[k] = cols[HEADER_ALIASES[k]];
  });
  (required || []).forEach(function (k) {
    if (cols[k] === undefined) throw new Error('分頁「' + name + '」缺少「' + k + '」欄位。');
  });
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].every(function (v) { return v === '' || v === null; })) continue;
    const o = { _row: i + 1 };
    Object.keys(cols).forEach(function (k) { o[k] = values[i][cols[k] - 1]; });
    rows.push(o);
  }
  return { sheet: sh, cols: cols, rows: rows };
}

/** 依標題名稱新增多列；formats 指定某些欄位的數字格式。 */
function appendRows_(name, objs, formats) {
  if (!objs.length) return;
  const sh = sheet_(name);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const start = sh.getLastRow() + 1;
  const needRows = start + objs.length - 1 - sh.getMaxRows();
  if (needRows > 0) sh.insertRowsAfter(sh.getMaxRows(), needRows);
  const val = function (o, h) {
    const v = o[h] !== undefined ? o[h] : o[HEADER_ALIASES[h]];
    return v === undefined || v === null ? '' : v;
  };
  const data = objs.map(function (o) {
    return headers.map(function (h) { return val(o, h); });
  });
  // 文字欄位設成純文字格式，避免「0912」「1/2」被轉成數字或日期
  headers.forEach(function (h, i) {
    const fmt = formats && (formats[h] || formats[HEADER_ALIASES[h]]);
    const isText = objs.some(function (o) { const v = val(o, h); return typeof v === 'string' && v !== ''; });
    if (fmt) sh.getRange(start, i + 1, objs.length, 1).setNumberFormat(fmt);
    else if (isText) sh.getRange(start, i + 1, objs.length, 1).setNumberFormat('@');
  });
  sh.getRange(start, 1, objs.length, headers.length).setValues(data);
}

function setCell_(table, row, header, value) {
  const c = table.cols[header];
  if (c === undefined) throw new Error('分頁缺少「' + header + '」欄位。');
  table.sheet.getRange(row, c).setValue(value);
}

// =====================================================================
// 小工具
// =====================================================================

function tz_() { return Session.getScriptTimeZone() || 'Asia/Taipei'; }

function clean_(v, maxLen) {
  // 去掉開頭的 = + - @，避免被試算表當成公式
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().replace(/^[=+\-@]+/, '').trim().slice(0, maxLen);
}

function isTrue_(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE' || String(v).trim() === '是';
}

function today_(now) {
  const s = Utilities.formatDate(now || new Date(), tz_(), 'yyyy-MM-dd').split('-');
  return new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]));
}

function addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

function parseDate_(s) {
  const m = String(s || '').trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return d.getMonth() === Number(m[2]) - 1 ? d : null;
}

function toDate_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  return parseDate_(v);
}

function dateStr_(d) {
  if (!d) return '';
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

function fmtDateTime_(v) {
  return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy/MM/dd HH:mm') : String(v || '');
}

// ---- 鎖單 ----

/** 本週（星期一開始）的訂貨截止時間，例如本週一 12:00。 */
function currentDeadline_(now) {
  const today = today_(now);
  const monday = addDays_(today, -((today.getDay() + 6) % 7));
  const d = addDays_(monday, LOCK_WEEKDAY - 1);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), LOCK_HOUR, LOCK_MINUTE, 0);
}

/** 老師目前最早可以訂的上課日期：截止前是下週一，截止後是下下週一。 */
function earliestClassDate_(now) {
  const today = today_(now);
  const nextMonday = addDays_(today, 7 - ((today.getDay() + 6) % 7));
  return now < currentDeadline_(now) ? nextMonday : addDays_(nextMonday, 7);
}

/** 這個上課日期的訂單是否已鎖住（已過截止，老師不能再取消）。 */
function isLocked_(classDate) {
  return !classDate || classDate < earliestClassDate_(new Date());
}
