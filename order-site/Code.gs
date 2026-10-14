/**
 * 三信家商餐飲科訂貨系統：Google Apps Script 後端
 *
 * © 2026 翁煒程 版權所有。
 * 本系統之程式碼、介面設計與內容均受中華民國著作權法保護，僅供三信家商餐飲科教職員內部使用。
 * 未經著作權人書面同意，禁止以任何形式重製、拷貝、修改、轉載或散佈本系統之程式碼、畫面及網址，
 * 亦不得提供予校外人士使用。違者將依法追究相關法律責任。
 *
 * 資料全部存在試算表「訂貨表2.0」：
 *   帳號設定   老師帳號、密碼、啟用
 *   登入資訊   每次登入的紀錄
 *   食材總表   可訂購的食材
 *   廠商設定   廠商的送貨組別與星期代碼
 *   訂單總表   老師送出的訂單，一個食材一列
 *   待審核食材 老師提交、食材總表裡沒有的食材
 *   範本       老師各自的配方範本
 *   單位       老師新增的單位（管理者審核）
 *
 * 程式依「標題名稱」找欄位，欄位順序調整不影響，但請不要改標題文字。
 * 部署方式見 README.md。
 */

// 試算表 ID（網址 /d/ 與 /edit 之間那段）。
const SPREADSHEET_ID = '1qIQe4STafrhzQk4I7NLkZQxJGe5hozf0g04zvJ-qZ_I';

const SITE_TITLE = '三信家商餐飲科訂貨系統';
const ADMIN_ACCOUNTS = ['管理者'];          // 這些帳號登入後會看到管理功能
const DEPARTMENTS = ['中餐', '西餐', '日料', '烘焙', '飲調', '果雕', '刀功'];
// 上課時間格式：「上午／下午 幾點:幾分」，分鐘每 5 分鐘，例如「下午 1:05」
const CLASS_TIME_PATTERN = /^(上午|下午|晚上) (1[0-2]|[1-9]):[0-5][05]$/;
const SESSIONS = ['上午', '下午', '晚上'];
// 鎖單：每週一 12:00 截止「下週」的訂單；截止後最早只能訂下下週的課。管理者不受限制。
const LOCK_WEEKDAY = 1;                     // 截止星期（1 = 星期一 … 5 = 星期五）
const LOCK_HOUR = 12;
const LOCK_MINUTE = 0;
const TOOL_ID = '器具';                     // 使用器具在訂單總表的「食材ID」與「廠商」
const ITEMS_CACHE_MINUTES = 10;             // 食材清單暫存時間（加快登入）；核准新食材時會立即更新
const MINE_DAYS = 60;                       // 「我的訂單」預設顯示最近幾天
const MAX_DAYS_AHEAD = 30;                  // 上課日期最晚可選幾天後
const MAX_ITEMS_PER_ORDER = 80;
const SESSION_HOURS = 6;                    // 登入有效時間
const MAX_LOGIN_FAILS = 5;
const DEFAULT_PASSWORD = '0000';            // 用這個密碼登入時，要先改密碼才能使用
const MIN_PASSWORD_LENGTH = 4;                  // 連續輸錯幾次就暫停登入
const LOCK_MINUTES = 10;
const ORDER_STATUSES = ['待處理', '已審核', '已叫貨', '已到貨', '已取消'];
// 單位選單：叫貨單位、配方單位。老師只能從這些單位選，避免同一個單位有好幾種寫法。
const ORDER_UNITS = ['斤', '兩', '公斤', 'g', '顆', '個', '粒', '條', '根', '把', '支', '朵', '片', '塊', '隻', '尾', '付', '打', '包', '袋', '盒', '罐', '瓶', '桶', '箱', '公升'];
const RECIPE_UNITS = ['T', 't', '杯', 'ml', 'cc'];
const UNIT_SYNONYMS = { '公克': 'g', '克': 'g', 'G': 'g', 'kg': '公斤', 'KG': '公斤', 'Kg': '公斤', '台斤': '斤', '大匙': 'T', '小匙': 't', 'CC': 'cc', '毫升': 'ml', 'L': '公升' };
const MAX_GROUPS = 99;

// ---- 自動化（排程）----
const NOTIFY_EMAILS = [];                    // 收通知的 Email；空白就寄給部署這個網站的帳號
const AUTO_FOLDER_NAME = '三信家商餐飲科訂貨單';   // 自動產生的廠商訂貨單存在雲端硬碟這個資料夾
const DIGEST_HOUR = 7;                       // 每日摘要信幾點寄
const VENDOR_ORDER_HOUR = 12;                // 每週一幾點產生下週廠商訂貨單（配合 12:00 截止）
const LINE_BIND_MINUTES = 30;                // LINE 綁定碼有效時間
const LINE_RESERVE = 20;                     // LINE 本月剩下的則數少於這個數字時，不再私訊老師（留給管理者通知）
const REMIND_HOUR = 15;                     // 每週五幾點在 LINE 群組提醒還沒訂貨的老師
const MAINT_HOUR = 23;                       // 每週日幾點做每週維護（先備份、再封存舊訂單）
const ARCHIVE_AFTER_DAYS = 14;               // 上課日期過了幾天的訂單搬到「歷史訂單」
const BACKUP_KEEP = 4;                       // 試算表備份保留幾份（每週一份）
const BACKUP_FOLDER_NAME = '三信家商餐飲科訂貨系統備份';
const BACKUP_PREFIX = '訂貨表備份 ';
const HISTORY_FOLDER_NAME = '三信家商餐飲科歷史訂單';   // 每個月一個檔案
const HISTORY_KEEP_DAYS = 400;               // 「歷史訂單」分頁保留約 13 個月（更早的已存在資料夾的月份檔案）
const IMPORT_MAX_MB = 10;                    // 上傳檔案大小上限
const PENDING_ID = '待審核';
// 廠商寫「自備」的食材（例如水、冰塊）：學校自己有，不跟廠商叫貨，但請料單上一定要印出來
const SELF_VENDOR = '自備';
// 同一家廠商依訂單科別分開叫貨：食材總表照舊寫「E乾貨」「C蛋商」，
// 烘焙科的訂單算「E烘焙乾貨」「C烘焙蛋」，其他科算「E中餐乾貨」「C中餐蛋」。
const VENDOR_SPLIT = {
  'E乾貨': { baking: 'E烘焙乾貨', other: 'E中餐乾貨' },
  'C蛋商': { baking: 'C烘焙蛋', other: 'C中餐蛋' },
};
const BAKING_DEPTS = ['烘焙'];

const SHEET = {
  accounts: '帳號設定',
  logins: '登入資訊',
  items: '食材總表',
  vendors: '廠商設定',
  orders: '訂單總表',
  pending: '待審核食材',
  templates: '範本',
  units: '單位',
  history: '歷史訂單',
  schedule: '課表',
  myLessons: '老師課表',
  skips: '不用訂貨',
  errors: '錯誤紀錄',
  edits: '修改紀錄',
  guoji: '國技排班',
};
const GUOJI_HEADERS = ['日期', '時段', '班級', '核心主題', '實作名稱', '老師', '教室', '備註', '需要訂貨', '匯入時間'];
const EDIT_HEADERS = ['時間', '訂單編號', '修改人', '內容'];
const ERROR_HEADERS = ['時間', '帳號', '功能', '類型', '訊息', '畫面'];
const ERROR_KEEP = 2000;                      // 錯誤紀錄最多保留幾列
const MY_LESSON_HEADERS = ['老師', '類型', '星期', '日期', '時段', '課程', '班級', '教室', '建立時間'];
const SKIP_HEADERS = ['老師', '日期', '時段', '課程', '標記時間'];

// 需要你在試算表手動新增的欄位（沒有的話網站照常運作，只是對應功能不會存檔）
const OPTIONAL_COLUMNS = {
  '訂單總表': ['組數', '不叫貨', '段落', '審核人'],
  '食材總表': ['舊編號'],
};
const UNIT_HEADERS = ['單位', '類型', '狀態', '提交老師', '提交時間', '審核人'];
const TEMPLATE_HEADERS = ['老師', '範本名稱', '科別', '實作項目', '類型', '食材ID', '名稱', '廠商', '一組數量', '單位', '備註', '段落', '不叫貨', '建立時間'];

const WEEKDAY = ['日', '一', '二', '三', '四', '五', '六'];
// 試算表標題的別名：訂單總表用「使用日期」或「上課日期」都可以
const HEADER_ALIASES = { '上課日期': '使用日期', '使用日期': '上課日期' };

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(SITE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // 「加到主畫面」後像 App 一樣全螢幕打開（Apps Script 不允許設定主畫面名稱，加的時候可以自己改）
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
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

  const accRows = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']).rows;
  let acc = accRows.find(function (r) { return String(r['帳號']).trim() === account; });
  if (!acc) {
    // 打字差一點也找得到：異體字（尙／尚）、空白、全形字
    const key = accountKey_(account);
    const same = accRows.filter(function (r) { return isTrue_(r['啟用']) && accountKey_(r['帳號']) === key; });
    if (same.length === 1) acc = same[0];
  }
  if (acc) account = String(acc['帳號']).trim();   // 之後一律用試算表上的寫法
  const ok = acc && isTrue_(acc['啟用']) && String(acc['密碼']) === password;
  logLogin_(account, ok ? '成功' : '失敗');

  if (!ok) {
    cache.put(failKey, String(fails + 1), LOCK_MINUTES * 60);
    throw new Error(acc && isTrue_(acc['啟用'])
      ? '密碼不正確（有分大小寫）。忘記密碼請找管理者重設。'
      : '找不到這個帳號，或帳號未啟用。帳號是你的名字，請確認有沒有打錯字或異體字（例如「尚／尙」「黃／黄」）。');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid();
  cache.put('session:' + token, account, SESSION_HOURS * 3600);
  const mustChange = password === DEFAULT_PASSWORD;
  if (mustChange) cache.put('mustchange:' + token, '1', SESSION_HOURS * 3600);
  return { token: token, account: account, isAdmin: ADMIN_ACCOUNTS.indexOf(account) >= 0, mustChange: mustChange };
}

// 常見的異體字、簡體字 → 台灣常用寫法（只用來比對帳號，不會改試算表）
const NAME_VARIANTS_ = {
  '尙': '尚', '黄': '黃', '吴': '吳', '吕': '呂', '温': '溫', '峯': '峰', '啓': '啟', '眞': '真', '淸': '清', '靑': '青',
  '姸': '妍', '羣': '群', '綫': '線', '着': '著', '裏': '裡', '敎': '教', '衞': '衛', '册': '冊', '爲': '為', '恒': '恆',
  '晋': '晉', '凉': '涼', '叙': '敘', '庄': '莊', '甯': '寧', '邨': '村', '鄕': '鄉', '冲': '沖', '刘': '劉', '陈': '陳',
  '张': '張', '杨': '楊', '赵': '趙', '许': '許', '郑': '鄭', '谢': '謝', '叶': '葉', '苏': '蘇',
};
function accountKey_(name) {
  return String(name == null ? '' : name)
    .replace(/[\uFF01-\uFF5E]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })   // 全形英數 → 半形
    .replace(/[\s\u3000]+/g, '').toLowerCase()
    .replace(/./g, function (c) { return NAME_VARIANTS_[c] || c; });
}

/**
 * 老師自己改密碼。
 * 規則：至少 MIN_PASSWORD_LENGTH 碼、不能跟舊密碼一樣、不能是預設密碼。
 */
function changePassword(token, oldPassword, newPassword) {
  const account = auth_(token, false, true);
  oldPassword = String(oldPassword == null ? '' : oldPassword);
  newPassword = String(newPassword == null ? '' : newPassword);
  if (/\s/.test(newPassword)) throw new Error('新密碼不能有空白。');
  if (newPassword.length < MIN_PASSWORD_LENGTH) throw new Error('新密碼至少要 ' + MIN_PASSWORD_LENGTH + ' 碼。');
  if (newPassword.length > 30) throw new Error('新密碼最多 30 碼。');
  if (newPassword === DEFAULT_PASSWORD) throw new Error('新密碼不能是預設密碼 ' + DEFAULT_PASSWORD + '。');
  if (newPassword === oldPassword) throw new Error('新密碼不能跟目前的密碼一樣。');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']);
    const r = t.rows.find(function (x) { return String(x['帳號']).trim() === account; });
    if (!r || !isTrue_(r['啟用'])) throw new Error('找不到你的帳號，請聯絡管理者。');
    if (String(r['密碼']) !== oldPassword) throw new Error('目前的密碼不正確。');
    // 純文字格式，避免「0123」被試算表變成數字 123
    const cell = t.sheet.getRange(r._row, t.cols['密碼']);
    cell.setNumberFormat('@');
    cell.setValue(newPassword);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove('mustchange:' + token);
  logLogin_(account, '改密碼');
  return true;
}

/** 登入並一次載入網站資料（少一次來回，登入比較快）。 */
function loginAndLoad(account, password) {
  const res = login(account, password);
  if (!res.mustChange) res.data = getBootstrap(res.token);   // 要先改密碼的，改完才載入資料
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
function auth_(token, needAdmin, allowMustChange) {
  const cache = CacheService.getScriptCache();
  const account = token ? cache.get('session:' + token) : null;
  if (!account) throw new Error('SESSION_EXPIRED');
  // 還在用預設密碼的，除了改密碼以外什麼都不能做
  if (!allowMustChange && cache.get('mustchange:' + token)) throw new Error('MUST_CHANGE_PASSWORD');
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
    vendorSplit: VENDOR_SPLIT,
    selfVendor: SELF_VENDOR,
    bakingDepts: BAKING_DEPTS,
    teachers: teachers,
    departments: depts,
    minDate: dateStr_(isAdmin ? today_() : earliestClassDate_(now)),
    maxDate: dateStr_(addDays_(today_(), MAX_DAYS_AHEAD)),
    lockRule: lockRuleText_(),
    statuses: ORDER_STATUSES,
    orderUnits: unitLists_().order,
    recipeUnits: unitLists_().recipe,
    pendingUnits: unitLists_().pending,
    missingSetup: isAdmin ? missingSetup_() : [],
  };
}

/** 檢查試算表是否已加上新功能需要的欄位和分頁，回傳缺少的項目（只給管理者看）。 */
function missingSetup_() {
  const out = [];
  Object.keys(OPTIONAL_COLUMNS).forEach(function (name) {
    let cols = {};
    try { cols = readHeaders_(name); } catch (e) { out.push('分頁「' + name + '」'); return; }
    OPTIONAL_COLUMNS[name].forEach(function (c) { if (cols[c] === undefined) out.push('「' + name + '」的「' + c + '」欄'); });
  });
  if (!ss_().getSheetByName(SHEET.templates)) out.push('「範本」分頁');
  if (!ss_().getSheetByName(SHEET.units)) out.push('「單位」分頁');
  return out;
}

function readHeaders_(name) {
  const sh = sheet_(name);
  const cols = {};
  sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].forEach(function (h, i) {
    const k = String(h).trim();
    if (k && cols[k] === undefined) cols[k] = i + 1;
  });
  return cols;
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
        code: String(r['舊編號'] || '').trim(),
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
    const res = writeOrder_(o, orderNo, new Date(), account, account, {});
    SpreadsheetApp.flush();
    return res;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 老師修改自己還沒截止、還是「待處理」的訂單（整張重寫，訂單編號不變），並用 LINE 通知管理者。
 * payload 跟 submitOrder 一樣。
 */
function updateOrder(token, orderNo, payload) {
  const account = auth_(token);
  const isAdmin = ADMIN_ACCOUNTS.indexOf(account) >= 0;
  orderNo = String(orderNo || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let res, before;
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '狀態']);
    const rows = t.rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo; });
    if (!rows.length) throw new Error('找不到訂單 ' + orderNo + '，可能已被封存或刪除。');
    const r0 = rows[0];
    const operator = String(r0['操作老師'] || '').trim();
    const forWho = String(r0['代訂對象老師'] || '').trim();
    if (!isAdmin && operator !== account && forWho !== account) throw new Error('只能修改自己的訂單。');
    const active = rows.filter(function (r) { return String(r['狀態']).trim() !== '已取消'; });
    if (!active.length) throw new Error('這張訂單已取消，請重新訂一張。');
    if (active.some(function (r) { return String(r['狀態']).trim() !== '待處理'; })) {
      throw new Error('這張訂單已經在處理，無法修改，請聯絡管理者。');
    }
    if (!isAdmin && isLocked_(toDate_(r0['使用日期']))) throw new Error('這張訂單的上課週已過訂貨截止時間，無法修改，請聯絡管理者。');

    // 代訂對象老師自己改的時候，保留原本的操作老師和代訂對象
    const p = Object.assign({}, payload);
    if (operator !== account && !isAdmin) p.onBehalfOf = '';
    const o = validateOrder_(p, operator === account || isAdmin ? account : operator, isAdmin);
    if (operator !== account && !isAdmin) o.onBehalfOf = forWho;

    before = groupOrders_(rows, null)[0];
    const pendingNames = {};
    readTable_(SHEET.pending, ['食材名稱']).rows.forEach(function (x) {
      if (String(x['對應訂單編號'] || '').trim() === orderNo) pendingNames[String(x['食材名稱']).trim()] = true;
    });
    deleteRows_(t.sheet, rows.map(function (r) { return r._row; }));
    const submitted = r0['送出時間'] instanceof Date ? r0['送出時間'] : new Date();
    res = writeOrder_(o, orderNo, submitted, operator || account, account, pendingNames);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  try {
    const changes = diffOrders_(before, orderOf_(orderNo));
    if (!changes.length) return Object.assign(res, { updated: true });
    const text = (before ? before.className + '「' + before.project + '」：' : '') + changes.join('、');
    logOrderEdit_(orderNo, account, changes);
    // 不馬上傳 LINE（省則數），記下來寫進隔天的每日摘要
    const props = PropertiesService.getScriptProperties();
    const edits = JSON.parse(props.getProperty('ORDER_EDITS') || '[]');
    edits.push({ t: fmtDateTime_(new Date()), who: account, text: text, no: orderNo });
    props.setProperty('ORDER_EDITS', JSON.stringify(edits.slice(-50)));
  } catch (e) {
    console.error(e.message);
  }
  res.updated = true;
  return res;
}

/** 寫入一張訂單的各列（新訂單和修改共用）。skipPending：已經在待審核清單裡的新食材名稱。 */
function writeOrder_(o, orderNo, submittedAt, operator, account, skipPending) {
  const now = new Date();
  registerUnits_(o.newUnits, account);
  const base = {
    '訂單編號': orderNo,
    '送出時間': submittedAt,
    '使用日期': o.useDate,
    '星期': '星期' + WEEKDAY[o.useDate.getDay()],
    '操作老師': operator,
    '代訂對象老師': o.onBehalfOf,
    '科別': o.dept,
    '上課時間': o.classTime,
    '上課教室': o.room,
    '上課班級': o.className,
    '實作項目': o.project,
    '組數': o.groups || '',
    '狀態': '待處理',
  };
  const rows = o.lines.map(function (l) {
    return Object.assign({}, base, {
      '食材ID': l.id, '食材名稱': l.name, '廠商': l.vendor, '送貨組別': l.group,
      '數量': l.qty, '單位': l.unit, '備註': l.note,
      '不叫貨': l.noOrder ? true : '', '段落': l.section || '',
    });
  });
  appendRows_(SHEET.orders, rows, { '送出時間': 'yyyy/mm/dd hh:mm:ss', '使用日期': 'yyyy-mm-dd' });

  const newOnes = o.lines.filter(function (l) { return l.id === PENDING_ID; });
  const pending = newOnes.filter(function (l) { return !skipPending[l.name]; }).map(function (l) {
    return {
      '新增時間': now, '提交老師': account, '食材名稱': l.name, '廠商': l.catVendor || l.vendor,
      '單位': l.unit, '對應訂單編號': orderNo, '審核狀態': '待審核',
    };
  });
  if (pending.length) appendRows_(SHEET.pending, pending, { '新增時間': 'yyyy/mm/dd hh:mm:ss' });
  const tools = o.lines.filter(function (l) { return l.id === TOOL_ID; }).length;
  return { orderNo: orderNo, useDate: dateStr_(o.useDate), count: o.lines.length - tools, tools: tools, pending: newOnes.length };
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
  if (!CLASS_TIME_PATTERN.test(classTime)) throw new Error('請選擇上課時間（上午／下午／晚上、幾點、幾分）。');
  if (!className) throw new Error('請填寫上課班級。');
  if (!project) throw new Error('請填寫實作項目。');
  let groups = '';
  if (p.groups !== '' && p.groups !== null && p.groups !== undefined) {
    groups = Number(p.groups);
    if (!Number.isInteger(groups) || groups < 1 || groups > MAX_GROUPS) throw new Error('組數需為 1 到 ' + MAX_GROUPS + ' 的整數。');
  }

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

  const newUnits = checkNewUnits_(p.newUnits);
  const lines = [];
  (Array.isArray(p.items) ? p.items : []).forEach(function (it) {
    const id = String(it && it.id || '').trim();
    const r = catalog[id];
    if (!r) throw new Error('食材「' + id + '」已停用或不存在，請重新整理頁面。');
    const name = String(r['食材名稱']).trim();
    const vendor = String(r['廠商']).trim();
    lines.push({
      id: id, name: name, vendor: orderVendor_(vendor, dept), catVendor: vendor,
      qty: qty_(it.qty, name), unit: unit_(it.unit, r['單位'], name, newUnits), note: clean_(it.note, 100),
      group: deliveryGroup_(vendor, dow, vendorGroups),
      noOrder: !!it.noOrder || vendor === SELF_VENDOR, section: clean_(it.section, 20),
    });
  });
  (Array.isArray(p.newItems) ? p.newItems : []).forEach(function (it) {
    const name = clean_(it && it.name, 40);
    const vendor = clean_(it && it.vendor, 20);
    if (!name) throw new Error('新增食材請填寫名稱。');
    if (!vendor) throw new Error('請選擇「' + name + '」的廠商。');
    lines.push({
      id: PENDING_ID, name: name, vendor: orderVendor_(vendor, dept), catVendor: vendor,
      qty: qty_(it.qty, name), unit: unit_(it.unit, '', name, newUnits), note: clean_(it.note, 100),
      group: deliveryGroup_(vendor, dow, vendorGroups),
      noOrder: !!it.noOrder || vendor === SELF_VENDOR, section: clean_(it.section, 20),
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
    className: className, project: project, onBehalfOf: onBehalfOf, groups: groups, lines: lines,
    newUnits: newUnits.filter(function (u) { return lines.some(function (l) { return l.unit === u; }); }),
  };
}

function qty_(v, name) {
  const n = Number(v);
  if (!isFinite(n) || n <= 0 || n > 99999) throw new Error('「' + name + '」的數量需大於 0。');
  return Math.round(n * 100) / 100;
}

function normUnit_(u) {
  u = String(u == null ? '' : u).trim();
  return UNIT_SYNONYMS[u] || u;
}

function unit_(v, fallback, name, extra) {
  const u = normUnit_(clean_(v, 10)) || normUnit_(fallback);
  if (!u) throw new Error('請選擇「' + name + '」的單位。');
  const ok = unitLists_();
  if (ok.all.indexOf(u) < 0 && (!extra || extra.indexOf(u) < 0)) {
    throw new Error('「' + name + '」的單位「' + u + '」不在選單中，請重新選擇，或用「＋ 新增單位」。');
  }
  return u;
}

// ---- 老師新增的單位（「單位」分頁，管理者審核） ----

let UNIT_CACHE_ = null;   // 同一次執行內重複使用

/** { order: [...], recipe: [...], pending: [...], all: [...] }：內建單位＋核准的單位；審核中的也可以用 */
function unitLists_() {
  if (UNIT_CACHE_) return UNIT_CACHE_;
  const order = ORDER_UNITS.slice(), recipe = RECIPE_UNITS.slice(), pending = [];
  readCustomUnits_().forEach(function (r) {
    const u = r.unit;
    if (!u || order.indexOf(u) >= 0 || recipe.indexOf(u) >= 0 || pending.indexOf(u) >= 0) return;
    if (r.status === '核准') (r.type === '配方' ? recipe : order).push(u);
    else if (r.status === '待審核') pending.push(u);
  });
  UNIT_CACHE_ = { order: order, recipe: recipe, pending: pending, all: order.concat(recipe, pending) };
  return UNIT_CACHE_;
}

function readCustomUnits_() {
  if (!ss_().getSheetByName(SHEET.units)) return [];
  return readTable_(SHEET.units, ['單位', '狀態']).rows.map(function (r) {
    return {
      row: r._row, unit: normUnit_(r['單位']), type: String(r['類型'] || '').trim(),
      status: String(r['狀態'] || '').trim() || '待審核', teacher: String(r['提交老師'] || ''),
      time: fmtDateTime_(r['提交時間']),
    };
  });
}

/** 檢查老師這次新增的單位，回傳整理過的清單（還沒寫入試算表）。 */
function checkNewUnits_(list) {
  const out = [];
  (Array.isArray(list) ? list : []).forEach(function (v) {
    const u = normUnit_(clean_(v, 10));
    if (!u || out.indexOf(u) >= 0) return;
    if (/[0-9０-９]/.test(u)) throw new Error('新增的單位「' + u + '」不能包含數字。');
    if (unitLists_().all.indexOf(u) >= 0) return;   // 已經有了
    if (readCustomUnits_().some(function (r) { return r.unit === u && (r.status === '退回' || r.status === '停用'); })) {
      throw new Error('單位「' + u + '」已被管理者退回或停用，請改用清單中的單位。');
    }
    out.push(u);
  });
  if (out.length && !ss_().getSheetByName(SHEET.units)) throw new Error('試算表還沒有「單位」分頁，暫時不能新增單位，請聯絡管理者。');
  return out;
}

function registerUnits_(units, account) {
  if (!units.length) return;
  appendRows_(SHEET.units, units.map(function (u) {
    return { '單位': u, '類型': '', '狀態': '待審核', '提交老師': account, '提交時間': new Date() };
  }), { '提交時間': 'yyyy/mm/dd hh:mm' });
  UNIT_CACHE_ = null;
}

function adminGetPendingUnits(token) {
  auth_(token, true);
  return readCustomUnits_().filter(function (r) { return r.status === '待審核'; });
}

/** decision: 'approve'（type: '叫貨' 或 '配方'）或 'reject' */
function adminReviewUnit(token, row, decision, type) {
  const admin = auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.units, ['單位', '狀態']);
    const r = t.rows.find(function (x) { return x._row === Number(row); });
    if (!r) throw new Error('找不到這個單位，請重新整理。');
    if (String(r['狀態'] || '').trim() && String(r['狀態']).trim() !== '待審核') throw new Error('這個單位已經審核過了。');
    if (decision === 'approve') {
      setCell_(t, r._row, '類型', type === '配方' ? '配方' : '叫貨');
      setCell_(t, r._row, '狀態', '核准');
    } else if (decision === 'reject') {
      setCell_(t, r._row, '狀態', '退回');
    } else {
      throw new Error('審核動作不正確。');
    }
    if (t.cols['審核人'] !== undefined) setCell_(t, r._row, '審核人', admin);
    UNIT_CACHE_ = null;
    return decision === 'approve' ? '已核准' : '已退回';
  } finally {
    lock.releaseLock();
  }
}

/** 管理者：所有單位（內建＋自訂，含停用的），給「單位管理」用 */
function adminListUnits(token) {
  auth_(token, true);
  return { order: ORDER_UNITS.slice(), recipe: RECIPE_UNITS.slice(),
    custom: readCustomUnits_().filter(function (r) { return r.status !== '待審核'; }) };
}

/** 管理者直接新增單位（不用審核）。type: '叫貨' | '配方' */
function adminAddUnit(token, unit, type) {
  const admin = auth_(token, true);
  const u = normUnit_(clean_(unit, 10));
  if (!u) throw new Error('請輸入單位。');
  if (/[0-9０-９]/.test(u)) throw new Error('單位不能包含數字。');
  if (u.length > 4) throw new Error('單位最多 4 個字。');
  type = type === '配方' ? '配方' : '叫貨';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    UNIT_CACHE_ = null;
    const typed = clean_(unit, 10);
    if (ORDER_UNITS.indexOf(u) >= 0 || RECIPE_UNITS.indexOf(u) >= 0) {
      throw new Error(typed !== u ? '「' + typed + '」就是系統內建的「' + u + '」，老師打「' + typed + '」會自動換成「' + u + '」，不用新增。' : '「' + u + '」是系統內建的單位，已經可以用了。');
    }
    const sh = ensureSheet_(SHEET.units, UNIT_HEADERS);
    const t = readTable_(SHEET.units, ['單位', '狀態']);
    const same = t.rows.filter(function (r) { return normUnit_(r['單位']) === u; });
    if (same.some(function (r) { return String(r['狀態']).trim() === '核准'; })) throw new Error('「' + u + '」已經在單位清單裡了。');
    if (same.length) {
      // 老師提交過（待審核／退回／停用）：直接改成核准
      const r = same[0];
      setCell_(t, r._row, '類型', type);
      setCell_(t, r._row, '狀態', '核准');
      if (t.cols['審核人'] !== undefined) setCell_(t, r._row, '審核人', admin);
    } else {
      appendRows_(SHEET.units, [{ '單位': u, '類型': type, '狀態': '核准', '提交老師': admin, '提交時間': new Date(), '審核人': admin }],
        { '提交時間': 'yyyy/mm/dd hh:mm' });
    }
    SpreadsheetApp.flush();
    UNIT_CACHE_ = null;
  } finally {
    lock.releaseLock();
  }
  return { unit: u, type: type };
}

/** 管理者停用／恢復自訂單位（內建單位不能停用）。舊訂單不受影響。 */
function adminSetUnitActive(token, row, active) {
  const admin = auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.units, ['單位', '狀態']);
    const r = t.rows.find(function (x) { return x._row === Number(row); });
    if (!r) throw new Error('找不到這個單位，請重新整理。');
    setCell_(t, r._row, '狀態', active ? '核准' : '停用');
    if (active && !String(r['類型'] || '').trim()) setCell_(t, r._row, '類型', '叫貨');
    if (t.cols['審核人'] !== undefined) setCell_(t, r._row, '審核人', admin);
    SpreadsheetApp.flush();
    UNIT_CACHE_ = null;
  } finally {
    lock.releaseLock();
  }
  return adminListUnits(token);
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

/** 訂單上要寫的廠商：E乾貨、C蛋商依科別分成烘焙／中餐兩種。 */
function orderVendor_(vendor, dept) {
  const split = VENDOR_SPLIT[String(vendor || '').trim()];
  if (!split) return vendor;
  return BAKING_DEPTS.indexOf(String(dept || '').trim()) >= 0 ? split.baking : split.other;
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

/** 我送出的、或別人代我訂的訂單（預設最近 60 天的上課日期起；days 最多 400，「顯示更早的訂單」用）。 */
function getMyOrders(token, days) {
  const account = auth_(token);
  days = Math.min(Math.max(Number(days) || MINE_DAYS, MINE_DAYS), 400);
  const since = addDays_(today_(), -days);
  const rows = withHistory_(readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows, since).filter(function (r) {
    if (String(r['操作老師']).trim() !== account && String(r['代訂對象老師'] || '').trim() !== account) return false;
    const d = toDate_(r['使用日期']);
    return !d || d >= since;
  });
  return groupOrders_(rows, account).sort(function (a, b) { return a.orderNo < b.orderNo ? 1 : -1; });
}

/**
 * 「上次這堂課訂了什麼」：同一位老師、同一個班級、同一個時段，上課日期在 date 之前最近的一張訂單（含歷史訂單、一年內）。
 * 同一個星期幾的優先。找不到回傳 null。
 */
function findLastLessonOrder(token, date, session, className) {
  const account = auth_(token);
  const before = parseDate_(date);
  if (!before || SESSIONS.indexOf(session) < 0) return null;
  const classes = classKeys_(className);
  if (!classes.length) return null;
  const since = addDays_(before, -365);
  const rows = withHistory_(readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows, since).filter(function (r) {
    if (String(r['操作老師']).trim() !== account && String(r['代訂對象老師'] || '').trim() !== account) return false;
    const d = toDate_(r['使用日期']);
    return d && d >= since && d < before;
  });
  const cands = groupOrders_(rows, account).filter(function (o) {
    if (o.status === '已取消' || sessionOf_(o.classTime) !== session) return false;
    if (!o.items.some(function (it) { return it.status !== '已取消'; })) return false;
    return classKeys_(o.className).some(function (c) { return classes.indexOf(c) >= 0; });
  });
  if (!cands.length) return null;
  const wd = before.getDay();
  cands.sort(function (a, b) {
    const sa = parseDate_(a.useDate).getDay() === wd ? 1 : 0, sb = parseDate_(b.useDate).getDay() === wd ? 1 : 0;
    return sa !== sb ? sb - sa : (a.useDate < b.useDate ? 1 : -1);
  });
  return cands[0];
}

/** 班級「109/110」「109、110」→ ['109','110']（全形數字轉半形） */
function classKeys_(s) {
  return String(s || '').replace(/[\uFF10-\uFF19]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .split(/[\/、,，\s]+/).map(function (x) { return x.trim(); }).filter(Boolean);
}

function cancelOrder(token, orderNo) {
  const account = auth_(token);
  const isAdmin = ADMIN_ACCOUNTS.indexOf(account) >= 0;
  const res = cancelOrderChecked_(account, isAdmin, orderNo);
  try { logOrderEdit_(String(orderNo || '').trim(), account, ['取消整張訂單']); } catch (e) { console.error(e.message); }
  if (isAdmin) notifyTeacherOfEdit_(String(orderNo || '').trim(), account, ['取消整張訂單']);
  return res;
}

function cancelOrderChecked_(account, isAdmin, orderNo) {
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
  const rows = withHistory_(readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows, start).filter(function (r) {
    const d = toDate_(r['使用日期']);
    return d && d >= start && d <= end;
  });
  const list = groupOrders_(rows, null).sort(function (a, b) {
    return a.useDate === b.useDate ? (a.orderNo < b.orderNo ? -1 : 1) : (a.useDate < b.useDate ? -1 : 1);
  });
  const edits = orderEditsFor_(list.map(function (o) { return o.orderNo; }));
  list.forEach(function (o) { o.edits = edits[o.orderNo] || []; });
  return list;
}

/** 管理者改上課資訊：{ className, room, project, groups }，只改有給的欄位。沒有要改的回傳 null。 */
function orderInfoSet_(info) {
  if (!info || typeof info !== 'object') return null;
  const set = {};
  if (info.className !== undefined) {
    const v = clean_(info.className, 30);
    if (!v) throw new Error('上課班級不能空白。');
    set['上課班級'] = v;
  }
  if (info.room !== undefined) set['上課教室'] = clean_(info.room, 20);
  if (info.project !== undefined) {
    const v = clean_(info.project, 60);
    if (!v) throw new Error('實作項目不能空白。');
    set['實作項目'] = v;
  }
  if (info.groups !== undefined) {
    if (info.groups === '' || info.groups === null) set['組數'] = '';
    else {
      const g = Number(info.groups);
      if (!Number.isInteger(g) || g < 1 || g > MAX_GROUPS) throw new Error('組數需為 1 到 ' + MAX_GROUPS + ' 的整數。');
      set['組數'] = g;
    }
  }
  return Object.keys(set).length ? set : null;
}

// ---- 修改紀錄：老師修改／取消訂單、管理者修改明細，每次記一列（時間、修改人、改了什麼） ----

/** 讀一張訂單（目前的訂單總表） */
function orderOf_(orderNo) {
  const rows = readTable_(SHEET.orders, ['訂單編號']).rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo; });
  return rows.length ? groupOrders_(rows, null)[0] : null;
}

/** 比對修改前後：上課資訊、刪除／新增的品項、數量、有貨。回傳每一項變更的文字。 */
function diffOrders_(a, b) {
  if (!a || !b) return [];
  const lines = [];
  const day = function (v) { const d = parseDate_(v); return d ? shortDateTw_(d) : String(v || ''); };
  [['useDate', '日期', day], ['classTime', '上課時間'], ['className', '班級'], ['room', '教室'], ['project', '實作項目'], ['groups', '組數'], ['dept', '科別']]
    .forEach(function (f) {
      const x = String(a[f[0]] == null ? '' : a[f[0]]), y = String(b[f[0]] == null ? '' : b[f[0]]);
      if (x !== y) lines.push(f[1] + ' ' + (f[2] ? f[2](x) : x || '（空白）') + ' → ' + (f[2] ? f[2](y) : y || '（空白）'));
    });
  const num = function (n) { return String(Math.round(n * 100) / 100); };
  const q = function (e, n) { return e.tool ? '×' + num(n) : ' ' + num(n) + e.unit; };
  const agg = function (o) {
    const m = {};
    o.items.forEach(function (it) {
      if (it.status === '已取消') return;
      const tool = it.id === TOOL_ID;
      const k = (tool ? '器具|' : '') + it.name + '|' + (tool ? '' : it.unit);
      const e = m[k] = m[k] || { label: (tool ? '器具「' + it.name + '」' : it.name), unit: tool ? '' : it.unit, tool: tool, qty: 0, noOrder: it.noOrder };
      e.qty += it.qty;
    });
    return m;
  };
  const A = agg(a), B = agg(b);
  Object.keys(A).forEach(function (k) {
    const x = A[k], y = B[k];
    if (!y) { lines.push('刪除 ' + x.label + q(x, x.qty)); return; }
    if (x.qty !== y.qty) lines.push(x.label + q(x, x.qty) + ' →' + q(y, y.qty));
    if (!x.noOrder !== !y.noOrder) lines.push(x.label + (y.noOrder ? ' 改成有貨' : ' 改成要叫貨'));
  });
  Object.keys(B).forEach(function (k) {
    if (!A[k]) lines.push('新增 ' + B[k].label + q(B[k], B[k].qty));
  });
  return lines;
}

function logOrderEdit_(orderNo, who, changes) {
  if (!changes.length) return;   // 按了儲存但沒改東西，不記
  ensureSheet_(SHEET.edits, EDIT_HEADERS);
  appendRows_(SHEET.edits, [{ '時間': new Date(), '訂單編號': orderNo, '修改人': who, '內容': changes.join('；') || '（內容沒有變）' }],
    { '時間': 'yyyy/mm/dd hh:mm:ss' });
}

/** { 訂單編號: [{ time, who, text }] }，依時間排序 */
function orderEditsFor_(orderNos) {
  const out = {};
  const sh = ss_().getSheetByName(SHEET.edits);
  if (!sh || sh.getLastRow() < 2 || !orderNos.length) return out;
  const want = {};
  orderNos.forEach(function (n) { want[n] = true; });
  readTable_(SHEET.edits, ['時間', '訂單編號']).rows.forEach(function (r) {
    const no = String(r['訂單編號']).trim();
    if (!want[no]) return;
    (out[no] = out[no] || []).push({ time: fmtDateTime_(errTime_(r['時間'])), who: String(r['修改人'] || ''), text: String(r['內容'] || '') });
  });
  Object.keys(out).forEach(function (k) { out[k].sort(function (x, y) { return x.time < y.time ? -1 : 1; }); });
  return out;
}

function adminSetStatus(token, orderNo, status) {
  const admin = auth_(token, true);
  if (ORDER_STATUSES.indexOf(status) < 0) throw new Error('狀態不正確。');
  const res = updateOrderStatus_(orderNo, status, null, status === '已審核' ? admin : null);
  if (status === '已取消') {
    try { logOrderEdit_(String(orderNo || '').trim(), admin, ['取消整張訂單']); } catch (e) { console.error(e.message); }
    return { result: res, notify: notifyTeacherOfEdit_(String(orderNo || '').trim(), admin, ['取消整張訂單']) };
  }
  return { result: res, notify: '' };
}

/**
 * 管理者改到老師的訂單（修改明細、取消）時，LINE 私訊那位老師（代訂的話傳給被代訂的老師）。
 * 老師沒綁定 LINE、管理者改自己的單、本月則數快用完時不傳。回傳結果文字（給畫面顯示）。
 */
function notifyTeacherOfEdit_(orderNo, admin, changes) {
  try {
    if (!changes || !changes.length) return '';
    let o = orderOf_(orderNo);
    if (!o) {
      const rows = withHistory_([], null).filter(function (r) { return String(r['訂單編號']).trim() === orderNo; });
      o = rows.length ? groupOrders_(rows, null)[0] : null;
    }
    if (!o) return '';
    const teacher = o.onBehalfOf || o.teacher;
    if (!teacher || teacher === admin) return '';
    const uid = lineUserOf_(teacher);
    if (!uid) return teacher + ' 老師沒有綁定 LINE，沒有通知。';
    const q = lineQuota_();
    if (q && q.limit && q.limit - q.used - 1 < LINE_RESERVE) return '本月 LINE 則數快用完，沒有通知 ' + teacher + ' 老師。';
    const d = parseDate_(o.useDate);
    const text = '【訂單修改通知】' + teacher + ' 老師您好，管理者修改了您的訂單：\n' +
      (d ? shortDateTw_(d) + ' ' : '') + o.className + '「' + o.project + '」\n' +
      changes.map(function (c) { return '・' + c; }).join('\n') +
      '\n\n有問題請找管理者。' + (webAppUrl_() ? '\n' + webAppUrl_() : '');
    return linePush_(uid, text) ? '已用 LINE 通知 ' + teacher + ' 老師。' : 'LINE 通知 ' + teacher + ' 老師失敗。';
  } catch (e) {
    console.error('通知老師失敗：' + e.message);
    return '';
  }
}

/**
 * 管理者修改訂單明細。
 * changes: [{ row, qty, unit, note, noOrder, cancel, kind, itemId, name }]，row 是訂單總表的列號（由 adminGetOrders 提供）
 *   kind = 'tool'：改成使用器具（name = 器具名稱）；kind = 'food' + itemId：改成食材總表的這樣食材
 */
function adminUpdateOrder(token, orderNo, changes, info) {
  const admin = auth_(token, true);
  orderNo = String(orderNo || '').trim();
  if (!Array.isArray(changes)) changes = [];
  const infoSet = orderInfoSet_(info);
  if (!changes.length && !infoSet) throw new Error('沒有要修改的內容。');
  let changesDone = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '數量', '單位', '狀態']);
    const byRow = {};
    t.rows.forEach(function (r) { if (String(r['訂單編號']).trim() === orderNo) byRow[r._row] = r; });
    if (!Object.keys(byRow).length) throw new Error('找不到訂單 ' + orderNo + '。');
    const before = groupOrders_(Object.keys(byRow).map(function (k) { return byRow[k]; }), null)[0];
    let catalog = null, vendorGroups = null;
    // 食材總表沒有的食材：先新增到食材總表，再當一般食材處理
    changes.forEach(function (c) {
      if (c.kind === 'food' && !c.itemId && c.newItem && !c.cancel) c.itemId = createItem_(c.newItem);
    });
    // 新增的品項：先用這張訂單的上課資訊新增一列，再跟其他列一樣套用修改
    const adds = changes.filter(function (c) { return c.add && !c.cancel; });
    if (adds.length) {
      const any = t.rows.find(function (r) { return String(r['訂單編號']).trim() === orderNo; });
      const keep = ['訂單編號', '送出時間', '使用日期', '星期', '操作老師', '代訂對象老師', '科別', '上課時間', '上課教室', '上課班級', '實作項目', '組數', '狀態'];
      const status = ['已取消'].indexOf(String(any['狀態']).trim()) >= 0 ? '待處理' : String(any['狀態']).trim();
      const rowsToAdd = adds.map(function (c) {
        const o = {};
        keep.forEach(function (k) { o[k] = any[k]; });
        o['狀態'] = status;
        o['食材ID'] = TOOL_ID; o['食材名稱'] = clean_(c.name, 40) || '新品項'; o['廠商'] = TOOL_ID; o['數量'] = 1;
        return o;
      });
      const start = t.sheet.getLastRow() + 1;
      appendRows_(SHEET.orders, rowsToAdd, { '送出時間': 'yyyy/mm/dd hh:mm:ss', '使用日期': 'yyyy-mm-dd' });
      SpreadsheetApp.flush();
      const t2 = readTable_(SHEET.orders, ['訂單編號', '數量', '單位', '狀態']);
      adds.forEach(function (c, i) {
        const r = t2.rows.find(function (x) { return x._row === start + i; });
        byRow[r._row] = r; c.row = r._row;
      });
      t.sheet = t2.sheet; t.cols = t2.cols;
    }
    changes.forEach(function (c) {
      if (c.add && c.cancel) return;
      const r = byRow[Number(c.row)];
      if (!r) throw new Error('訂單明細已變動，請重新整理後再改。');
      let name = String(r['食材名稱'] || '');
      if (c.cancel) {
        setCell_(t, r._row, '狀態', '已取消');
        return;
      }
      if (c.restore) {
        // 復原誤刪的品項：狀態改回這張訂單其他品項的狀態
        const other = Object.keys(byRow).map(function (k) { return String(byRow[k]['狀態']).trim(); })
          .find(function (st) { return st && st !== '已取消'; });
        setCell_(t, r._row, '狀態', other || '待處理');
        return;
      }
      let isTool = String(r['食材ID']).trim() === TOOL_ID;
      let catUnit = r['單位'];
      // 食材、器具互換，或換成另一樣食材（例如檔案帶入時對錯）
      if (c.kind === 'tool') {
        const toolName = clean_(c.name, 40) || name;
        if (!isTool || toolName !== name) {
          setCell_(t, r._row, '食材ID', TOOL_ID);
          setCell_(t, r._row, '食材名稱', toolName);
          if (t.cols['廠商'] !== undefined) setCell_(t, r._row, '廠商', TOOL_ID);
          if (t.cols['送貨組別'] !== undefined) setCell_(t, r._row, '送貨組別', '');
          if (t.cols['單位'] !== undefined) setCell_(t, r._row, '單位', '');
          if (t.cols['不叫貨'] !== undefined) setCell_(t, r._row, '不叫貨', '');
        }
        isTool = true; name = toolName;
      } else if (c.kind === 'food' && c.itemId && String(c.itemId).trim() !== String(r['食材ID']).trim()) {
        if (!catalog) {
          catalog = {};
          readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']).rows.forEach(function (x) {
            const id = String(x['食材ID']).trim();
            if (id && (x['啟用'] === undefined || isTrue_(x['啟用']))) catalog[id] = x;
          });
          vendorGroups = readVendorGroups_();
        }
        const it = catalog[String(c.itemId).trim()];
        if (!it) throw new Error('食材「' + c.itemId + '」已停用或不存在。');
        const vendor = String(it['廠商']).trim();
        const d = toDate_(r['使用日期']);
        name = String(it['食材名稱']).trim();
        setCell_(t, r._row, '食材ID', String(c.itemId).trim());
        setCell_(t, r._row, '食材名稱', name);
        if (t.cols['廠商'] !== undefined) setCell_(t, r._row, '廠商', orderVendor_(vendor, r['科別']));
        if (vendor === SELF_VENDOR) c.noOrder = true;
        if (t.cols['送貨組別'] !== undefined && d) setCell_(t, r._row, '送貨組別', deliveryGroup_(vendor, d.getDay(), vendorGroups));
        isTool = false; catUnit = it['單位'] || '';
      }
      setCell_(t, r._row, '數量', qty_(c.qty, name));
      if (!isTool) setCell_(t, r._row, '單位', unit_(c.unit, catUnit, name));
      if (t.cols['備註'] !== undefined) setCell_(t, r._row, '備註', clean_(c.note, 100));
      if (t.cols['不叫貨'] !== undefined) setCell_(t, r._row, '不叫貨', c.noOrder && !isTool ? true : '');
    });
    if (t.cols['審核人'] !== undefined) {
      Object.keys(byRow).forEach(function (row) { setCell_(t, Number(row), '審核人', admin); });
    }
    if (infoSet) {
      // 上課資訊（班級、教室、實作項目、組數）整張訂單每一列都改，包含剛新增的品項
      SpreadsheetApp.flush();
      const t2 = readTable_(SHEET.orders, ['訂單編號']);
      t2.rows.forEach(function (r) {
        if (String(r['訂單編號']).trim() !== orderNo) return;
        Object.keys(infoSet).forEach(function (k) { if (t2.cols[k] !== undefined) setCell_(t2, r._row, k, infoSet[k]); });
      });
    }
    SpreadsheetApp.flush();
    try { changesDone = diffOrders_(before, orderOf_(orderNo)); logOrderEdit_(orderNo, admin, changesDone); } catch (e) { console.error(e.message); }
  } finally {
    lock.releaseLock();
  }
  return { ok: true, notify: notifyTeacherOfEdit_(orderNo, admin, changesDone) };   // 傳 LINE 放在鎖外面，不讓別人等
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
/** 新增一筆食材到食材總表，回傳新的 I 編號。名稱重複時丟出 dupMessage。 */
function createItem_(it, dupMessage) {
  const name = clean_(it.name, 40);
  const vendor = clean_(it.vendor, 20);
  if (!name) throw new Error('請填寫食材名稱。');
  if (!vendor) throw new Error('請選擇「' + name + '」的廠商。');
  const items = readTable_(SHEET.items, ['食材ID', '食材名稱']);
  if (items.rows.some(function (r) { return String(r['食材名稱']).trim() === name; })) {
    throw new Error(dupMessage || '食材總表已經有「' + name + '」。');
  }
  let max = 0;
  items.rows.forEach(function (r) {
    const m = String(r['食材ID']).match(/^I(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  const id = 'I' + ('0000' + (max + 1)).slice(-4);
  const unit = it.unit ? normUnit_(clean_(it.unit, 10)) : '';
  appendRows_(SHEET.items, [{
    '食材ID': id, '食材名稱': name, '別名': clean_(it.alias, 100), '廠商': vendor,
    '科別': clean_(it.dept, 20) || '通用', '單位': unit, '適用組別': '全部', '啟用': true,
  }]);
  clearItemsCache_();
  return id;
}

/** 管理者在網站上新增食材（不用打開試算表）。item: { name, alias, vendor, dept, unit } */
function adminAddItem(token, item) {
  auth_(token, true);
  item = item || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const id = createItem_(item);
    SpreadsheetApp.flush();
    return cachedItems_().find(function (x) { return x.id === id; }) || { id: id, name: clean_(item.name, 40) };
  } finally {
    lock.releaseLock();
  }
}

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
      const id = createItem_({ name: name, alias: name === origName ? '' : origName, vendor: vendor, unit: unit },
        '食材總表已經有「' + name + '」，請改用「對應既有食材」。');
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
      setCell_(t, r._row, '廠商', orderVendor_(to.vendor, r['科別']));
      const d = toDate_(r['使用日期']);
      if (d && t.cols['送貨組別'] !== undefined) setCell_(t, r._row, '送貨組別', deliveryGroup_(to.vendor, d.getDay(), vendorGroups));
    }
  });
}

// =====================================================================
// 範本（每位老師各自的）
// =====================================================================

function templateSheet_() {
  const sh = ss_().getSheetByName(SHEET.templates);
  if (!sh) throw new Error('試算表還沒有「範本」分頁，請管理者先新增。');
  return sh;
}

/** 我的範本：[{ name, dept, project, created, items: [{ type, id, name, vendor, perGroup, unit, note, section, noOrder }] }] */
function getTemplates(token) {
  const account = auth_(token);
  if (!ss_().getSheetByName(SHEET.templates)) return [];
  const map = {};
  const list = [];
  readTable_(SHEET.templates, ['老師', '範本名稱']).rows.forEach(function (r) {
    if (String(r['老師']).trim() !== account) return;
    const name = String(r['範本名稱']).trim();
    if (!name) return;
    if (!map[name]) {
      map[name] = { name: name, dept: String(r['科別'] || ''), project: String(r['實作項目'] || ''), created: fmtDateTime_(r['建立時間']), items: [] };
      list.push(map[name]);
    }
    map[name].items.push({
      type: String(r['類型'] || '食材'),
      id: String(r['食材ID'] || ''),
      name: String(r['名稱'] || ''),
      vendor: String(r['廠商'] || ''),
      perGroup: Number(r['一組數量']) || 0,
      unit: String(r['單位'] || ''),
      note: String(r['備註'] || ''),
      section: String(r['段落'] || ''),
      noOrder: isTrue_(r['不叫貨']),
    });
  });
  return list.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
}

/**
 * 存成範本（同名的會覆蓋）。
 * t: { name, dept, project, items: [{ type: '食材'|'新食材'|'器具', id, name, vendor, perGroup, unit, note, section, noOrder }] }
 */
function saveTemplate(token, t) {
  const account = auth_(token);
  const name = clean_(t && t.name, 30);
  if (!name) throw new Error('請填寫範本名稱。');
  const items = Array.isArray(t.items) ? t.items : [];
  if (!items.length) throw new Error('範本至少要有一項食材或器具。');
  if (items.length > MAX_ITEMS_PER_ORDER) throw new Error('範本最多 ' + MAX_ITEMS_PER_ORDER + ' 項。');
  const now = new Date();
  const rows = items.map(function (it) {
    const type = ['食材', '新食材', '器具'].indexOf(it.type) >= 0 ? it.type : '食材';
    const nm = clean_(it.name, 40);
    const per = Number(it.perGroup);
    if (!nm) throw new Error('範本有品項沒有名稱。');
    if (!isFinite(per) || per <= 0) throw new Error('「' + nm + '」一組的數量需大於 0。');
    return {
      '老師': account, '範本名稱': name, '科別': clean_(t.dept, 20), '實作項目': clean_(t.project, 60),
      '類型': type, '食材ID': type === '食材' ? clean_(it.id, 20) : '', '名稱': nm, '廠商': clean_(it.vendor, 20),
      '一組數量': Math.round(per * 10000) / 10000, '單位': type === '器具' ? '' : normUnit_(clean_(it.unit, 10)),
      '備註': clean_(it.note, 100), '段落': clean_(it.section, 20), '不叫貨': it.noOrder ? true : '', '建立時間': now,
    };
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    templateSheet_();
    deleteTemplateRows_(account, name);
    appendRows_(SHEET.templates, rows, { '建立時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
    return name;
  } finally {
    lock.releaseLock();
  }
}

function deleteTemplate(token, name) {
  const account = auth_(token);
  name = clean_(name, 30);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    templateSheet_();
    return deleteTemplateRows_(account, name);
  } finally {
    lock.releaseLock();
  }
}

function deleteTemplateRows_(account, name) {
  const t = readTable_(SHEET.templates, ['老師', '範本名稱']);
  const rows = t.rows.filter(function (r) {
    return String(r['老師']).trim() === account && String(r['範本名稱']).trim() === name;
  }).map(function (r) { return r._row; }).sort(function (a, b) { return b - a; });
  rows.forEach(function (row) { t.sheet.deleteRow(row); });   // 從下往上刪，列號才不會跑掉
  return rows.length;
}

// =====================================================================
// 自動化：截止後自動產生廠商訂貨單、每日摘要（Email＋LINE 通知）
// =====================================================================
//
// 第一次使用：在 Apps Script 編輯器選函式「setupAutomation」按「執行」並授權，
// 之後就會每週一中午、每天早上自動執行。也可以在網站「廠商訂貨單」頁管理。

const AUTO_HANDLERS = ['autoVendorOrders', 'dailyDigest', 'weeklyMaintenance', 'remindTeachers'];

/** 在編輯器執行一次：授權並建立排程。 */
function setupAutomation() {
  setTriggers_(true);
  return automationStatus_();
}

function setTriggers_(enabled) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (AUTO_HANDLERS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  if (!enabled) return;
  ScriptApp.newTrigger('autoVendorOrders').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(VENDOR_ORDER_HOUR).inTimezone(tz_()).create();
  ScriptApp.newTrigger('dailyDigest').timeBased()
    .everyDays(1).atHour(DIGEST_HOUR).inTimezone(tz_()).create();
  ScriptApp.newTrigger('weeklyMaintenance').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(MAINT_HOUR).inTimezone(tz_()).create();
  ScriptApp.newTrigger('remindTeachers').timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(REMIND_HOUR).inTimezone(tz_()).create();
}

function automationStatus_() {
  const handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  const props = PropertiesService.getScriptProperties();
  return {
    vendorTrigger: handlers.indexOf('autoVendorOrders') >= 0,
    digestTrigger: handlers.indexOf('dailyDigest') >= 0,
    maintTrigger: handlers.indexOf('weeklyMaintenance') >= 0,
    remindTrigger: handlers.indexOf('remindTeachers') >= 0,
    line: lineStatus_(),
    remindHour: REMIND_HOUR,
    lineReserve: LINE_RESERVE,
    historyFolderUrl: props.getProperty('HIST_FOLDER_URL') || '',
    historyMonths: Object.keys(JSON.parse(props.getProperty('HIST_MONTHS') || '{}')).length,
    lastHistText: props.getProperty('LAST_HIST_TEXT') || '',
    errors7d: recentSystemErrors_(addDays_(new Date(), -7)).length,
    scheduleCount: scheduleCount_(),
    lastRemindTime: props.getProperty('LAST_REMIND_TIME') || '',
    lastMaintTime: props.getProperty('LAST_MAINT_TIME') || '',
    lastMaintText: props.getProperty('LAST_MAINT_TEXT') || '',
    maintHour: MAINT_HOUR,
    archiveDays: ARCHIVE_AFTER_DAYS,
    backupKeep: BACKUP_KEEP,
    emails: notifyEmails_(),
    lineTokenSet: !!props.getProperty('LINE_TOKEN'),
    lineBound: !!props.getProperty('LINE_TO'),
    webhookUrl: webAppUrl_(),
    lastVendorUrl: props.getProperty('LAST_VENDOR_URL') || '',
    lastVendorTime: props.getProperty('LAST_VENDOR_TIME') || '',
    vendorHour: VENDOR_ORDER_HOUR,
    digestHour: DIGEST_HOUR,
  };
}

/** 排程呼叫的函式只接受真正的排程觸發，避免被網頁直接呼叫。 */
function fromTrigger_(e) {
  if (!e || !e.triggerUid) return false;
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getUniqueId() === String(e.triggerUid); });
}

// 排程：每週一中午產生下週的廠商訂貨單
function autoVendorOrders(e) {
  if (!fromTrigger_(e)) throw new Error('這個功能只能由排程執行。');
  logged_('每週訂貨單', function () { runVendorOrders_(); });
}

// 排程：每天早上寄摘要（沒有新東西就不寄）
function dailyDigest(e) {
  if (!fromTrigger_(e)) throw new Error('這個功能只能由排程執行。');
  logged_('每日摘要', function () { runDigest_(false); });
}

// 排程：每週日晚上先備份試算表、再把舊訂單搬到「歷史訂單」
function weeklyMaintenance(e) {
  if (!fromTrigger_(e)) throw new Error('這個功能只能由排程執行。');
  logged_('每週備份封存', function () { runMaintenance_(); });
}

// 排程：每週五下午用 LINE 私訊提醒下週還沒訂貨的老師
function remindTeachers(e) {
  if (!fromTrigger_(e)) throw new Error('這個功能只能由排程執行。');
  logged_('週五訂貨提醒', function () { runReminder_(true); });
}

// ---- 錯誤紀錄：老師在網站上遇到的錯誤、排程失敗，記到「錯誤紀錄」分頁，系統錯誤會放進每日摘要 ----

/** 排程執行失敗時記下來（再丟出去，Apps Script 的執行紀錄也看得到） */
function logged_(what, fn) {
  try {
    return fn();
  } catch (err) {
    logError_('（排程）', what, '系統錯誤', err && err.message ? err.message : String(err), '');
    throw err;
  }
}

/** 系統錯誤＝程式或 Google 服務出問題；操作提示＝給老師看的中文說明（例如「請選擇日期」） */
function errorType_(message) {
  const m = String(message || '');
  if (/TypeError|ReferenceError|RangeError|SyntaxError|is not defined|is not a function|Cannot read|Cannot set|undefined|Exception|Service|服務|逾時|timed? ?out|權限|permission|authori[sz]|Too many|quota|配額|伺服器|NetworkError|網路|網頁錯誤|failed|無法連線/i.test(m)) return '系統錯誤';
  return /[\u4e00-\u9fff]/.test(m) ? '操作提示' : '系統錯誤';
}

function logError_(account, fn, type, message, view) {
  try {
    const cache = CacheService.getScriptCache();
    const key = 'err:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, account + '|' + fn + '|' + message, Utilities.Charset.UTF_8));
    if (cache.get(key)) return;   // 同一個人、同一個錯誤 10 分鐘內只記一次
    cache.put(key, '1', 600);
    const sh = ensureSheet_(SHEET.errors, ERROR_HEADERS);
    appendRows_(SHEET.errors, [{ '時間': new Date(), '帳號': clean_(account, 30), '功能': clean_(fn, 40), '類型': type,
      '訊息': clean_(message, 300), '畫面': clean_(view, 100) }], { '時間': 'yyyy/mm/dd hh:mm:ss' });
    const extra = sh.getLastRow() - 1 - ERROR_KEEP;
    if (extra > 200) sh.deleteRows(2, extra);
  } catch (e) {
    console.error('錯誤紀錄寫入失敗：' + e.message);
  }
}

/** 網頁回報錯誤（google.script.run 失敗、網頁程式錯誤）。沒登入也可以回報。 */
function logClientError(token, fn, message, view) {
  const account = (token && CacheService.getScriptCache().get('session:' + token)) || '（未登入）';
  const msg = String(message || '');
  if (!msg || /SESSION_EXPIRED|MUST_CHANGE_PASSWORD/.test(msg)) return false;
  logError_(account, String(fn || ''), errorType_(msg), msg, String(view || ''));
  return true;
}

function errTime_(v) {
  if (v instanceof Date) return v;
  const d = parseDateTime_(String(v || '').replace(/-/g, '/'));
  return d || toDate_(v);
}

/** since 之後的系統錯誤（每日摘要、自動化畫面用） */
function recentSystemErrors_(since) {
  const sh = ss_().getSheetByName(SHEET.errors);
  if (!sh || sh.getLastRow() < 2) return [];
  return readTable_(SHEET.errors, ['時間', '類型']).rows.filter(function (r) {
    const t = errTime_(r['時間']);
    return String(r['類型']).trim() === '系統錯誤' && t && t > since;
  }).map(function (r) {
    return { time: fmtDateTime_(errTime_(r['時間'])), account: String(r['帳號'] || ''), fn: String(r['功能'] || ''), message: String(r['訊息'] || '') };
  });
}

// ---- 管理者在網站上操作 ----

function adminGetAutomation(token) {
  auth_(token, true);
  return automationStatus_();
}

function adminSetAutomation(token, enabled) {
  auth_(token, true);
  setTriggers_(!!enabled);
  return automationStatus_();
}

function adminRunVendorNow(token) {
  auth_(token, true);
  return runVendorOrders_();
}

function adminRunMaintenanceNow(token) {
  auth_(token, true);
  return runMaintenance_();
}

/** send = false 只預覽提醒內容，不傳送 */
function adminRunReminderNow(token, send) {
  auth_(token, true);
  return runReminder_(!!send);
}

function adminRunDigestNow(token) {
  auth_(token, true);
  return runDigest_(true);
}

/** 儲存 LINE 官方帳號的 Channel access token（先確認可以用） */
function adminSetLineToken(token, lineToken) {
  auth_(token, true);
  lineToken = String(lineToken || '').trim();
  const props = PropertiesService.getScriptProperties();
  if (!lineToken) {
    props.deleteProperty('LINE_TOKEN');
    props.deleteProperty('LINE_TO');
    return automationStatus_();
  }
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
    headers: { Authorization: 'Bearer ' + lineToken }, muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('LINE 不接受這組 Channel access token，請確認複製完整（錯誤碼 ' + res.getResponseCode() + '）。');
  props.setProperty('LINE_TOKEN', lineToken);
  return automationStatus_();
}

/** 產生一次性綁定碼：用 LINE 傳「綁定 123456」給官方帳號 */
function adminLineBindCode(token) {
  auth_(token, true);
  if (!PropertiesService.getScriptProperties().getProperty('LINE_TOKEN')) throw new Error('請先儲存 Channel access token。');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('lineBind', code, LINE_BIND_MINUTES * 60);
  return { code: code, minutes: LINE_BIND_MINUTES, webhookUrl: webAppUrl_() };
}

function adminTestNotify(token) {
  auth_(token, true);
  return notify_('【測試】三信家商餐飲科訂貨系統通知',
    '<p>這是測試通知。收到這封信表示 Email 通知設定正常。</p>',
    '【測試】三信家商餐飲科訂貨系統：收到這則訊息表示 LINE 通知設定正常。');
}

// ---- LINE webhook：只處理綁定 ----
// LINE Developers 後台的 Webhook URL 填這個網站的網址（網頁應用程式網址）。
function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents || '{}');
    const props = PropertiesService.getScriptProperties();
    const lineToken = props.getProperty('LINE_TOKEN');
    (body.events || []).forEach(function (ev) {
      if (ev.type !== 'message' || !ev.message || ev.message.type !== 'text' || !lineToken) return;
      const userId = ev.source && ev.source.userId;
      const m = String(ev.message.text).trim().match(/^綁定\s*(群組)?\s*(\d{6})$/);
      if (!m) {
        // 選單按鈕（我的未訂貨課程、我的訂單、截止時間說明）：免費回覆；群組裡不回
        if (ev.source && ev.source.type === 'user') {
          const ans = lineAutoReply_(ev.message.text, userId);
          if (ans) lineReply_(ev.replyToken, ans);
        }
        return;
      }
      const cache = CacheService.getScriptCache();
      const code = cache.get('lineBind');
      const teacher = !m[1] && cache.get('tbind:' + m[2]);
      if (teacher && ev.source && ev.source.type === 'user' && userId) {
        // 老師個人綁定
        const old = props.getProperty('LU:' + teacher);
        if (old) props.deleteProperty('LA:' + old);
        const prevAcc = props.getProperty('LA:' + userId);
        if (prevAcc) props.deleteProperty('LU:' + prevAcc);
        props.setProperty('LU:' + teacher, userId);
        props.setProperty('LA:' + userId, teacher);
        cache.remove('tbind:' + m[2]);
        lineReply_(ev.replyToken, teacher + ' 老師，綁定完成！之後有課還沒訂貨時，會在截止前提醒您。\n點下方選單「我的未訂貨課程」隨時可以查詢。');
        return;
      }
      let reply;
      if (code && code === m[2] && m[1]) {
        // 「綁定群組 123456」：在老師的 LINE 群組裡傳，之後訂貨提醒會傳到這個群組
        reply = '訂貨提醒已改成私訊老師個人，不用綁定群組了。請老師登入網站 →「我的課」→「綁定我的 LINE」。';
      } else if (code && code === m[2]) {
        const to = ev.source && (ev.source.groupId || ev.source.roomId || ev.source.userId);
        props.setProperty('LINE_TO', to);
        cache.remove('lineBind');
        reply = '綁定完成！之後訂貨系統的通知會傳到這裡。';
      } else {
        reply = '綁定碼不正確或已過期，請到網站重新產生。';
      }
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + lineToken },
        payload: JSON.stringify({ replyToken: ev.replyToken, messages: [{ type: 'text', text: reply }] }),
      });
    });
  } catch (err) {
    console.error('LINE webhook：' + err.message);
    logError_('（LINE）', 'LINE 回覆', '系統錯誤', err.message, '');
  }
  return ContentService.createTextOutput('OK');
}

// ---- 通知 ----

function notifyEmails_() {
  if (NOTIFY_EMAILS.length) return NOTIFY_EMAILS.slice();
  const me = Session.getEffectiveUser().getEmail();
  return me ? [me] : [];
}

function webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** 寄 Email，並在有綁定 LINE 時推播。回傳各管道的結果。 */
function notify_(subject, html, lineText) {
  const out = { email: '', line: '' };
  const to = notifyEmails_();
  try {
    if (to.length) {
      MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: html + '<p style="color:#888;font-size:12px">' + SITE_TITLE + ' 自動通知</p>' });
      out.email = '已寄到 ' + to.join('、');
    } else {
      out.email = '沒有收件人';
    }
  } catch (e) {
    out.email = 'Email 寄送失敗：' + e.message;
  }
  const props = PropertiesService.getScriptProperties();
  const lineToken = props.getProperty('LINE_TOKEN'), lineTo = props.getProperty('LINE_TO');
  if (lineToken && lineTo) {
    try {
      const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + lineToken },
        payload: JSON.stringify({ to: lineTo, messages: [{ type: 'text', text: String(lineText).slice(0, 4900) }] }),
      });
      out.line = res.getResponseCode() === 200 ? '已傳到 LINE' : 'LINE 傳送失敗（錯誤碼 ' + res.getResponseCode() + '）';
    } catch (e) {
      out.line = 'LINE 傳送失敗：' + e.message;
    }
  } else {
    out.line = '尚未綁定 LINE';
  }
  return out;
}

// ---- 每週維護：備份＋封存 ----

function runMaintenance_() {
  const props = PropertiesService.getScriptProperties();
  const out = { backup: null, archived: 0, orders: 0, error: '' };
  try {
    out.backup = backupSpreadsheet_();
  } catch (e) {
    out.error = '備份失敗，這次先不封存舊訂單：' + e.message;
  }
  if (!out.error) {
    try {
      const a = archiveOldOrders_();
      out.archived = a.rows; out.orders = a.orders; out.before = a.before;
    } catch (e) {
      out.error = '封存舊訂單失敗：' + e.message;
    }
  }
  if (!out.error) {
    try {
      out.history = exportHistoryMonths_();
    } catch (e) {
      out.error = '歷史訂單存檔失敗：' + e.message;
    }
  }
  const lines = [];
  if (out.backup) lines.push('已備份試算表（保留最近 ' + BACKUP_KEEP + ' 份' + (out.backup.removed ? '，刪除 ' + out.backup.removed + ' 份舊備份' : '') + '）');
  if (!out.error) lines.push(out.orders ? '已把 ' + out.orders + ' 張舊訂單（上課日期 ' + out.before + ' 以前）搬到「' + SHEET.history + '」' : '沒有需要封存的舊訂單');
  if (out.history) lines.push.apply(lines, historyLines_(out.history));
  if (out.error) lines.push('⚠ ' + out.error);
  const text = lines.join('；');
  props.setProperty('LAST_MAINT_TIME', fmtDateTime_(new Date()));
  props.setProperty('LAST_MAINT_TEXT', text);
  out.notify = notify_('【每週維護】' + (out.error ? '有問題，請查看' : '完成'),
    '<ul>' + lines.map(function (l) { return '<li>' + escHtml_(l) + '</li>'; }).join('') + '</ul>' +
      (out.backup ? '<p><a href="' + out.backup.url + '">打開這週的備份</a>（雲端硬碟資料夾「' + escHtml_(BACKUP_FOLDER_NAME) + '」）</p>' : '') +
      (out.history && out.history.files.length ? '<p><a href="' + out.history.folderUrl + '">打開歷史訂單資料夾</a></p>' : ''),
    '【每週維護】\n' + lines.join('\n'));
  out.text = text;
  return out;
}

// ---- 歷史訂單：每個月存成一個試算表，放在雲端硬碟「三信家商餐飲科歷史訂單」資料夾 ----

function folderNamed_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/**
 * 1. 「歷史訂單」分頁裡已經全部封存的月份（月底過了 ARCHIVE_AFTER_DAYS 天），每個月存成一個檔案「歷史訂單 2026年10月」。
 * 2. 上課日期早於 HISTORY_KEEP_DAYS 天、而且已經存檔的列，從「歷史訂單」分頁刪掉（刪之前先核對月份檔案，缺的補進去）。
 * 已存檔的月份記在 HIST_MONTHS（{ '2026-10': 檔案ID }），不會重複建檔。
 */
function exportHistoryMonths_() {
  const props = PropertiesService.getScriptProperties();
  const folder = folderNamed_(HISTORY_FOLDER_NAME);
  props.setProperty('HIST_FOLDER_URL', folder.getUrl());
  const out = { files: [], removed: 0, folderUrl: folder.getUrl() };
  const sh = ss_().getSheetByName(SHEET.history);
  if (!sh || sh.getLastRow() < 2) return out;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const t = readTable_(SHEET.history, ['訂單編號', '使用日期']);
    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const raw = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
    const byMonth = {};
    t.rows.forEach(function (r) {
      const d = toDate_(r['使用日期']);
      if (!d) return;   // 沒有上課日期的列不動
      const key = Utilities.formatDate(d, tz_(), 'yyyy-MM');
      (byMonth[key] = byMonth[key] || []).push({ no: String(r['訂單編號']).trim(), d: d, row: r._row, values: raw[r._row - 1] });
    });
    const done = JSON.parse(props.getProperty('HIST_MONTHS') || '{}');
    const today = today_();

    Object.keys(byMonth).sort().forEach(function (key) {
      if (done[key]) return;
      const ym = key.split('-').map(Number);
      const monthEnd = new Date(ym[0], ym[1], 0);
      if (addDays_(monthEnd, ARCHIVE_AFTER_DAYS + 1) > today) return;   // 這個月還有訂單沒封存完
      const file = newHistoryFile_(ym[0], ym[1], header, byMonth[key].map(function (x) { return x.values; }), folder);
      done[key] = file.getId();
      props.setProperty('HIST_MONTHS', JSON.stringify(done));   // 每存完一個月就記下來，中途失敗也不會重複建檔
      out.files.push({ month: ym[0] + '年' + ym[1] + '月', orders: countOrders_(byMonth[key]), url: file.getUrl() });
    });

    const cutoff = addDays_(today, -HISTORY_KEEP_DAYS);
    const del = [];
    Object.keys(byMonth).forEach(function (key) {
      const old = byMonth[key].filter(function (x) { return x.d < cutoff; });
      if (!done[key] || !old.length) return;
      // 刪之前核對：月份檔案裡沒有的訂單先補進去（檔案被刪掉就重建）
      let ss2 = null;
      try { ss2 = SpreadsheetApp.openById(done[key]); } catch (e) { ss2 = null; }
      if (!ss2) {
        const ym = key.split('-').map(Number);
        const f = newHistoryFile_(ym[0], ym[1], header, byMonth[key].map(function (x) { return x.values; }), folder);
        done[key] = f.getId();
        props.setProperty('HIST_MONTHS', JSON.stringify(done));
      } else {
        const s2 = ss2.getSheets()[0];
        const head2 = s2.getRange(1, 1, 1, Math.max(s2.getLastColumn(), 1)).getValues()[0].map(function (h) { return String(h).trim(); });
        const noCol = head2.indexOf('訂單編號');
        const have = {};
        if (noCol >= 0 && s2.getLastRow() > 1) {
          s2.getRange(2, noCol + 1, s2.getLastRow() - 1, 1).getValues().forEach(function (v) { have[String(v[0]).trim()] = true; });
        }
        const missing = old.filter(function (x) { return !have[x.no]; });
        if (missing.length) {
          s2.getRange(s2.getLastRow() + 1, 1, missing.length, lastCol).setValues(missing.map(function (x) { return x.values; }));
        }
      }
      old.forEach(function (x) { del.push(x.row); });
    });
    if (del.length) {
      SpreadsheetApp.flush();
      deleteRows_(sh, del);
      SpreadsheetApp.flush();
    }
    out.removed = del.length;
  } finally {
    lock.releaseLock();
  }
  props.setProperty('LAST_HIST_TEXT', historyLines_(out).join('；'));
  return out;
}

function newHistoryFile_(year, month, header, rows, folder) {
  const file = SpreadsheetApp.create('歷史訂單 ' + year + '年' + month + '月');
  const s = file.getSheets()[0];
  s.setName('訂單');
  const all = [header].concat(rows);
  s.getRange(1, 1, all.length, header.length).setValues(all);
  s.getRange(1, 1, 1, header.length).setFontWeight('bold');
  s.setFrozenRows(1);
  header.forEach(function (h, i) {
    h = String(h).trim();
    if (rows.length && /日期$/.test(h)) s.getRange(2, i + 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    if (rows.length && /時間$/.test(h) && h !== '上課時間') s.getRange(2, i + 1, rows.length, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  });
  DriveApp.getFileById(file.getId()).moveTo(folder);
  return file;
}

function countOrders_(list) {
  const seen = {};
  list.forEach(function (x) { seen[x.no] = true; });
  return Object.keys(seen).length;
}

function historyLines_(h) {
  const lines = [];
  if (h.files.length) lines.push('已存檔 ' + h.files.map(function (f) { return f.month + '（' + f.orders + ' 張）'; }).join('、') + ' 到雲端硬碟「' + HISTORY_FOLDER_NAME + '」');
  if (h.removed) lines.push('「' + SHEET.history + '」分頁刪除 ' + h.removed + ' 列超過 ' + HISTORY_KEEP_DAYS + ' 天、已存檔的資料');
  if (!lines.length) lines.push('沒有需要存檔的月份（每個月結束約 2 週後才會存檔）');
  return lines;
}

/** 管理者：立即整理歷史訂單 */
function adminRunHistoryExport(token) {
  auth_(token, true);
  const h = exportHistoryMonths_();
  return { text: historyLines_(h).join('；'), folderUrl: h.folderUrl, files: h.files, removed: h.removed };
}

/** 複製整份試算表到備份資料夾，只留最近 BACKUP_KEEP 份。 */
function backupSpreadsheet_() {
  const src = DriveApp.getFileById(ss_().getId());
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
  const copy = src.makeCopy(BACKUP_PREFIX + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), folder);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(BACKUP_PREFIX) === 0) files.push(f);
  }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  let removed = 0;
  files.slice(BACKUP_KEEP).forEach(function (f) {
    if (f.getId() === copy.getId()) return;
    f.setTrashed(true);
    removed++;
  });
  return { url: copy.getUrl(), name: copy.getName(), removed: removed };
}

/** 上課日期早於 (今天 − ARCHIVE_AFTER_DAYS) 的訂單整張搬到「歷史訂單」分頁。 */
function archiveOldOrders_() {
  const before = addDays_(today_(), -ARCHIVE_AFTER_DAYS);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '使用日期']);
    // 一張訂單的上課日期都一樣；以訂單為單位判斷，避免一張單被拆開
    const newest = {};
    t.rows.forEach(function (r) {
      const no = String(r['訂單編號']).trim();
      const d = toDate_(r['使用日期']);
      if (!no) return;
      if (newest[no] === null) return;
      newest[no] = !d ? null : (!newest[no] || d > newest[no] ? d : newest[no]);   // 沒有日期的不封存
    });
    const old = t.rows.filter(function (r) {
      const no = String(r['訂單編號']).trim();
      return no && newest[no] && newest[no] < before;
    });
    if (!old.length) return { rows: 0, orders: 0, before: dateStr_(before) };

    historySheet_(t.sheet);
    const already = {};
    readTable_(SHEET.history, ['訂單編號']).rows.forEach(function (r) { already[String(r['訂單編號']).trim()] = true; });
    // 上次搬到一半（已寫入歷史、還沒刪除）的訂單不重複寫入
    const toWrite = old.filter(function (r) { return !already[String(r['訂單編號']).trim()]; }).map(function (r) {
      const o = {};
      Object.keys(t.cols).forEach(function (k) { o[k] = r[k]; });
      return o;
    });
    appendRows_(SHEET.history, toWrite);
    SpreadsheetApp.flush();

    deleteRows_(t.sheet, old.map(function (r) { return r._row; }));
    SpreadsheetApp.flush();
    const orders = {};
    old.forEach(function (r) { orders[String(r['訂單編號']).trim()] = true; });
    return { rows: old.length, orders: Object.keys(orders).length, before: dateStr_(before) };
  } finally {
    lock.releaseLock();
  }
}

/** 取得「歷史訂單」分頁；沒有就建立，標題跟訂單總表一樣（缺的欄位補在後面）。 */
function historySheet_(ordersSheet) {
  const headers = ordersSheet.getRange(1, 1, 1, Math.max(ordersSheet.getLastColumn(), 1)).getValues()[0]
    .map(function (h) { return String(h).trim(); }).filter(Boolean);
  let sh = ss_().getSheetByName(SHEET.history);
  if (!sh) {
    sh = ss_().insertSheet(SHEET.history);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  const have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
  const missing = headers.filter(function (h) { return have.indexOf(h) < 0 && have.indexOf(HEADER_ALIASES[h]) < 0; });
  if (missing.length) {
    if (sh.getMaxColumns() < have.length + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), have.length + missing.length - sh.getMaxColumns());
    sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh;
}

/** 刪除多列（連續的列一次刪），從下往上刪列號才不會跑掉。 */
function deleteRows_(sh, rows) {
  rows = rows.slice().sort(function (a, b) { return b - a; });
  // 試算表至少要留一列沒凍結的列
  if (sh.getMaxRows() - sh.getFrozenRows() <= rows.length) sh.insertRowsAfter(sh.getMaxRows(), 1);
  let i = 0;
  while (i < rows.length) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++;
    sh.deleteRows(rows[j], j - i + 1);
    i = j + 1;
  }
}

/**
 * 「歷史訂單」分頁裡上課日期在 since 之後的列。
 * 歷史訂單是依封存順序往下加的，日期大致由舊到新：先只讀「使用日期」一欄，找到第一筆 since 之後的列，
 * 再只讀那一列到最後，資料多的時候比整張讀快很多。
 */
function historyRowsSince_(sh, since) {
  if (!since) return readTable_(SHEET.history, ['訂單編號']).rows;
  const lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  let dateIdx = header.indexOf('使用日期');
  if (dateIdx < 0) dateIdx = header.indexOf(HEADER_ALIASES['使用日期']);
  if (dateIdx < 0 || header.indexOf('訂單編號') < 0) return readTable_(SHEET.history, ['訂單編號']).rows;
  const dates = retry_(function () { return sh.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues(); });
  let first = -1;
  for (let i = 0; i < dates.length; i++) {
    const d = toDate_(dates[i][0]);
    if (!d || d >= since) { first = i; break; }
  }
  if (first < 0) return [];
  const values = retry_(function () { return sh.getRange(2 + first, 1, lastRow - 1 - first, lastCol).getValues(); });
  const cols = {};
  header.forEach(function (k, i) { if (k && cols[k] === undefined) cols[k] = i; });
  Object.keys(HEADER_ALIASES).forEach(function (k) {
    if (cols[k] === undefined && cols[HEADER_ALIASES[k]] !== undefined) cols[k] = cols[HEADER_ALIASES[k]];
  });
  const out = [];
  values.forEach(function (v, i) {
    if (v.every(function (x) { return x === '' || x === null; })) return;
    const o = { _row: 2 + first + i };
    Object.keys(cols).forEach(function (k) { o[k] = v[cols[k]]; });
    out.push(o);
  });
  return out;
}

/** 目前的訂單再加上「歷史訂單」裡上課日期在 since 之後的列（標記 _archived）。 */
function withHistory_(rows, since) {
  const sh = ss_().getSheetByName(SHEET.history);
  if (!sh || sh.getLastRow() < 2) return rows;
  const current = {};
  rows.forEach(function (r) { current[String(r['訂單編號']).trim()] = true; });
  const hist = historyRowsSince_(sh, since).filter(function (r) {
    if (current[String(r['訂單編號']).trim()]) return false;
    const d = toDate_(r['使用日期']);
    return !since || !d || d >= since;
  });
  hist.forEach(function (r) { r._archived = true; });
  return rows.concat(hist);
}

// ---- 檔案匯入：Excel、Word、照片（OCR）→ 表格與文字 ----
//
// 需要在 Apps Script 編輯器左邊「服務 ＋」加入「Drive API」。
// 檔案先轉成 Google 試算表／文件讀內容，讀完立刻刪掉，不會留在雲端硬碟。

const IMPORT_TYPES_ = {
  sheet: /\.(xlsx|xls|xlsm|csv|ods)$/i,
  doc: /\.(docx|doc|odt|rtf|txt)$/i,
  ocr: /\.(jpe?g|png|gif|bmp|webp|heic|heif|pdf)$/i,
};

function importFile(token, file) {
  auth_(token);
  if (!file || !file.data) throw new Error('沒有收到檔案。');
  const name = clean_(file.name, 100) || '上傳檔案';
  const kind = IMPORT_TYPES_.sheet.test(name) ? 'sheet' : IMPORT_TYPES_.doc.test(name) ? 'doc'
    : (IMPORT_TYPES_.ocr.test(name) || /^image\//.test(file.mimeType || '')) ? 'ocr' : '';
  if (!kind) throw new Error('不支援這種檔案。請上傳 Excel（.xlsx）、Word（.docx）、照片或 PDF。');
  const bytes = Utilities.base64Decode(String(file.data));
  if (bytes.length > IMPORT_MAX_MB * 1024 * 1024) throw new Error('檔案太大（上限 ' + IMPORT_MAX_MB + ' MB）。');
  if (typeof Drive === 'undefined') {
    throw new Error('伺服器還沒開啟「Drive API」服務。請管理者到 Apps Script 編輯器左邊「服務 ＋」加入 Drive API，再重新部署。');
  }
  const blob = Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', name);
  const target = kind === 'sheet' ? MimeType.GOOGLE_SHEETS : MimeType.GOOGLE_DOCS;
  const id = convertFile_(blob, name, target, kind === 'ocr');
  try {
    return kind === 'sheet' ? readConvertedSheet_(id) : readConvertedDoc_(id, kind === 'ocr');
  } finally {
    try { Drive.Files.remove(id); } catch (e) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e2) {} }
  }
}

/** 用 Drive API 把檔案轉成 Google 格式（照片／PDF 會做文字辨識），回傳新檔案 ID。 */
function convertFile_(blob, name, targetMime, ocr) {
  return retry_(function () {
    if (Drive.Files.create) {   // Drive API v3
      const f = Drive.Files.create({ name: '[匯入暫存] ' + name, mimeType: targetMime }, blob, ocr ? { ocrLanguage: 'zh-TW' } : {});
      return f.id;
    }
    const f = Drive.Files.insert({ title: '[匯入暫存] ' + name, mimeType: targetMime }, blob,   // Drive API v2
      ocr ? { ocr: true, ocrLanguage: 'zh-TW', convert: true } : { convert: true });
    return f.id;
  });
}

function readConvertedSheet_(id) {
  const ss = SpreadsheetApp.openById(id);
  const tables = ss.getSheets().map(function (sh) {
    const r = Math.min(sh.getLastRow(), 300), c = Math.min(sh.getLastColumn(), 40);
    return { name: sh.getName(), rows: r && c ? sh.getRange(1, 1, r, c).getDisplayValues() : [] };
  }).filter(function (t) { return t.rows.length; });
  return { kind: 'sheet', tables: tables, text: '' };
}

function readConvertedDoc_(id, ocr) {
  const body = DocumentApp.openById(id).getBody();
  const tables = [], lines = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    const type = el.getType();
    if (type === DocumentApp.ElementType.TABLE) {
      const tb = el.asTable(), rows = [];
      for (let r = 0; r < tb.getNumRows() && r < 300; r++) {
        const row = tb.getRow(r), cells = [];
        for (let c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c).getText().replace(/\s*\n\s*/g, ' ').trim());
        rows.push(cells);
      }
      tables.push({ name: '表格' + (tables.length + 1), rows: rows });
    } else if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
      el.getText().split('\n').forEach(function (l) { if (l.trim()) lines.push(l.trim()); });
    }
  }
  return { kind: ocr ? 'ocr' : 'doc', tables: tables, text: lines.slice(0, 600).join('\n') };
}

// ---- 訂貨提醒（依「課表」分頁） ----

/** 讀課表：[{ day: 0~4（週一~週五）, session: '上午'|'下午', room, course, periods, className, teacher, note }]，只取要訂貨、有老師的課 */
function readSchedule_() {
  const sh = ss_().getSheetByName(SHEET.schedule);
  if (!sh || sh.getLastRow() < 2) return [];
  return readTable_(SHEET.schedule, ['星期', '時段', '老師']).rows.map(function (r) {
    return {
      day: '一二三四五'.indexOf(String(r['星期']).replace(/星期|週/g, '').trim()),
      session: String(r['時段']).trim(),
      room: String(r['教室'] || '').trim(),
      course: String(r['課程'] || '').trim(),
      className: String(r['班級'] || '').trim(),
      teacher: String(r['老師']).trim(),
      need: r['需要訂貨'] === undefined || r['需要訂貨'] === '' || isTrue_(r['需要訂貨']),
    };
  }).filter(function (x) {
    return x.need && x.teacher && x.day >= 0 && SESSIONS.indexOf(x.session) >= 0 && !/國技|預留/.test(x.course);
  });
}

/** 自動建立需要的分頁（第一次使用時） */
function ensureSheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readMyLessonRows_() {
  const sh = ss_().getSheetByName(SHEET.myLessons);
  if (!sh || sh.getLastRow() < 2) return [];
  return readTable_(SHEET.myLessons, ['老師', '類型', '時段', '課程']).rows.map(function (r) {
    const d = toDate_(r['日期']);
    return {
      row: r._row, teacher: String(r['老師']).trim(), type: String(r['類型']).trim() === '單次' ? '單次' : '每週',
      day: '一二三四五'.indexOf(String(r['星期'] || '').replace(/星期|週/g, '').trim()), date: d ? dateStr_(d) : '',
      session: String(r['時段']).trim(), course: String(r['課程'] || '').trim(),
      className: String(r['班級'] || '').trim(), room: String(r['教室'] || '').trim(),
    };
  }).filter(function (x) { return x.teacher && SESSIONS.indexOf(x.session) >= 0; });
}

/** from～to 之間每一天的課（學校課表＋老師自訂），teacher = null 是全部老師 */
function lessonsBetween_(from, to, teacher) {
  const sched = readSchedule_(), mine = readMyLessonRows_();
  const out = [];
  // 國技班：每週輪流的老師，照「國技排班」分頁的日期
  readGuoji_().forEach(function (g) {
    if (!g.need || !g.teacher || (teacher && g.teacher !== teacher)) return;
    const d = parseDate_(g.date);
    if (!d || d < from || d > to || d.getDay() === 0 || d.getDay() === 6) return;
    out.push({ teacher: g.teacher, date: g.date, session: g.session, course: g.practice || g.topic || '國技實作', className: g.className, room: g.room, source: '國技', row: g.row });
  });
  for (let d = new Date(from); d <= to; d = addDays_(d, 1)) {
    const idx = d.getDay() - 1;   // 週一 = 0
    if (idx < 0 || idx > 4) continue;
    const ds = dateStr_(d);
    sched.forEach(function (x) {
      if (x.day === idx && (!teacher || x.teacher === teacher)) {
        out.push({ teacher: x.teacher, date: ds, session: x.session, course: x.course, className: x.className, room: x.room, source: '課表' });
      }
    });
    mine.forEach(function (x) {
      if (teacher && x.teacher !== teacher) return;
      if ((x.type === '每週' && x.day === idx) || (x.type === '單次' && x.date === ds)) {
        out.push({ teacher: x.teacher, date: ds, session: x.session, course: x.course, className: x.className, room: x.room, source: '自訂', row: x.row });
      }
    });
  }
  return out;
}

function sessionOf_(classTime) {
  const t = String(classTime || '');
  return t.indexOf('晚上') === 0 ? '晚上' : t.indexOf('下午') === 0 ? '下午' : '上午';
}

/** 每堂課加上 status：ordered（同老師、同一天、同時段有訂單）／skipped（按了不用訂貨）／todo */
function lessonsWithStatus_(from, to, teacher) {
  const lessons = lessonsBetween_(from, to, teacher);
  const have = {}, skip = {};
  ordersBetween_(from, to).forEach(function (o) {
    if (o.status === '已取消') return;
    have[(o.onBehalfOf || o.teacher) + '|' + o.useDate + '|' + sessionOf_(o.classTime)] = true;
  });
  const sh = ss_().getSheetByName(SHEET.skips);
  if (sh && sh.getLastRow() >= 2) {
    readTable_(SHEET.skips, ['老師', '日期', '時段']).rows.forEach(function (r) {
      const d = toDate_(r['日期']);
      if (d) skip[String(r['老師']).trim() + '|' + dateStr_(d) + '|' + String(r['時段']).trim()] = true;
    });
  }
  // 同一位老師同一天同時段有兩堂課時，只要一張訂單（也只算一次）
  const seen = {};
  return lessons.filter(function (x) {
    const k = x.teacher + '|' + x.date + '|' + x.session + '|' + x.course;
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  }).map(function (x) {
    const k = x.teacher + '|' + x.date + '|' + x.session;
    x.status = have[k] ? 'ordered' : skip[k] ? 'skipped' : 'todo';
    return x;
  });
}

// ---- 老師：我的課（未訂貨課程、不用訂貨、自訂上課時間） ----

/** 我接下來可以訂貨的課（最早可訂日期～30 天內）、學校課表、我自訂的課 */
/** 「我的課」看誰的課：老師只能看自己；管理者可以傳 as 看任何一位老師（課表上的名字也可以） */
function lessonAccount_(token, as) {
  const account = auth_(token);
  as = clean_(as, 30);
  if (!as || as === account) return account;
  if (ADMIN_ACCOUNTS.indexOf(account) < 0) throw new Error('只有管理者可以查看別的老師的課。');
  return as;
}

function getMyLessons(token, as) {
  const account = lessonAccount_(token, as);
  const now = new Date();
  const from = earliestClassDate_(now);
  const to = addDays_(today_(), MAX_DAYS_AHEAD);
  const firstWeekEnd = addDays_(from, 4);
  const deadline = addDays_(from, -7 + (LOCK_WEEKDAY - 1));
  const lessons = lessonsWithStatus_(from, to, account).sort(function (a, b) {
    return a.date === b.date ? SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session) : (a.date < b.date ? -1 : 1);
  });
  const todoNext = lessons.filter(function (x) { return x.status === 'todo' && parseDate_(x.date) <= firstWeekEnd; }).length;
  return {
    lessons: lessons,
    nextWeek: { from: dateStr_(from), to: dateStr_(firstWeekEnd), deadline: shortDateTw_(deadline) + ' ' + LOCK_HOUR + ':' + ('0' + LOCK_MINUTE).slice(-2), todo: todoNext },
    school: readSchedule_().filter(function (x) { return x.teacher === account; }),
    mine: readMyLessonRows_().filter(function (x) { return x.teacher === account; }),
    teacher: account,
    schoolRows: ADMIN_ACCOUNTS.indexOf(auth_(token)) >= 0 ? scheduleRowsOf_(account) : null,
    guojiRows: ADMIN_ACCOUNTS.indexOf(auth_(token)) >= 0 ? guojiRowsOf_(account) : null,
    line: { enabled: !!PropertiesService.getScriptProperties().getProperty('LINE_TOKEN'), bound: !!lineUserOf_(account) },
  };
}

/** 標記／取消「不用訂貨」：date = yyyy-mm-dd，session = 上午／下午／晚上 */
function setNoOrder(token, date, session, course, skip, as) {
  const account = lessonAccount_(token, as);
  const d = parseDate_(date);
  if (!d || SESSIONS.indexOf(session) < 0) throw new Error('資料不正確，請重新整理。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SHEET.skips, SKIP_HEADERS);
    const t = readTable_(SHEET.skips, ['老師', '日期', '時段']);
    const same = t.rows.filter(function (r) {
      const rd = toDate_(r['日期']);
      return String(r['老師']).trim() === account && rd && dateStr_(rd) === dateStr_(d) && String(r['時段']).trim() === session;
    });
    if (skip && !same.length) {
      appendRows_(SHEET.skips, [{ '老師': account, '日期': d, '時段': session, '課程': clean_(course, 40), '標記時間': new Date() }],
        { '日期': 'yyyy-mm-dd', '標記時間': 'yyyy/mm/dd hh:mm' });
    }
    if (!skip && same.length) deleteRows_(sh, same.map(function (r) { return r._row; }));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getMyLessons(token, as);
}

/** 老師自己新增上課時間（例如夜間部）。l: { type: '每週'|'單次', weekday: '一'~'五', date, session, course, className, room } */
function addMyLesson(token, l, as) {
  const account = lessonAccount_(token, as);
  l = l || {};
  const type = l.type === '單次' ? '單次' : '每週';
  const session = String(l.session || '').trim();
  const course = clean_(l.course, 30);
  if (SESSIONS.indexOf(session) < 0) throw new Error('請選擇上課時段（上午／下午／晚上）。');
  if (!course) throw new Error('請填寫課程或實習項目。');
  let weekday = '', date = '';
  if (type === '每週') {
    weekday = String(l.weekday || '').trim();
    if ('一二三四五'.indexOf(weekday) < 0 || weekday.length !== 1) throw new Error('請選擇星期幾（週一～週五）。');
  } else {
    date = parseDate_(l.date);
    if (!date) throw new Error('請選擇日期。');
    if (date.getDay() === 0 || date.getDay() === 6) throw new Error('日期是週末，廠商不送貨，請改選平日。');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.myLessons, MY_LESSON_HEADERS);
    appendRows_(SHEET.myLessons, [{
      '老師': account, '類型': type, '星期': weekday, '日期': date || '', '時段': session,
      '課程': course, '班級': clean_(l.className, 30), '教室': clean_(l.room, 20), '建立時間': new Date(),
    }], { '日期': 'yyyy-mm-dd', '建立時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getMyLessons(token, as);
}

function deleteMyLesson(token, row, as) {
  const account = lessonAccount_(token, as);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.myLessons, ['老師']);
    const r = t.rows.find(function (x) { return x._row === Number(row); });
    if (!r || String(r['老師']).trim() !== account) throw new Error('找不到這堂課，請重新整理。');
    deleteRows_(t.sheet, [r._row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getMyLessons(token, as);
}

// ---- 管理者：修改學校課表（「課表」分頁） ----

// ---- 國技排班（國技班每週輪流上課的老師、實作名稱） ----

function readGuoji_() {
  const sh = ss_().getSheetByName(SHEET.guoji);
  if (!sh || sh.getLastRow() < 2) return [];
  return readTable_(SHEET.guoji, ['日期', '時段', '班級', '老師']).rows.map(function (r) {
    const d = toDate_(r['日期']);
    return {
      row: r._row, date: d ? dateStr_(d) : '', session: String(r['時段']).trim(), className: String(r['班級']).trim(),
      topic: String(r['核心主題'] || '').trim(), practice: String(r['實作名稱'] || '').trim(), teacher: String(r['老師']).trim(),
      room: String(r['教室'] || '').trim(), note: String(r['備註'] || '').trim(),
      need: r['需要訂貨'] === undefined || r['需要訂貨'] === '' || isTrue_(r['需要訂貨']),
    };
  }).filter(function (g) { return g.date && SESSIONS.indexOf(g.session) >= 0; });
}

/** 管理者：某位老師接下來的國技排班（含已取消的） */
function guojiRowsOf_(teacher) {
  const from = addDays_(today_(), -7);
  return readGuoji_().filter(function (g) { return g.teacher === teacher && parseDate_(g.date) >= from; })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
}

/**
 * 管理者匯入國技排班表。rows: [{ date, session, className, topic, practice, teacher, room, note }]
 * 檔案裡有的班級，原本的排班整個換成新的（其他班級不動）。
 */
function adminImportGuoji(token, rows) {
  auth_(token, true);
  if (!Array.isArray(rows) || !rows.length) throw new Error('沒有可以匯入的排班。');
  if (rows.length > 3000) throw new Error('一次最多匯入 3000 週。');
  const clean = rows.map(function (x, i) {
    const d = parseDate_(x.date);
    if (!d) throw new Error('第 ' + (i + 1) + ' 筆的日期不正確。');
    const session = String(x.session || '').trim();
    if (SESSIONS.indexOf(session) < 0) throw new Error('第 ' + (i + 1) + ' 筆的時段不正確。');
    const className = clean_(x.className, 30), teacher = clean_(x.teacher, 30);
    if (!className || !teacher) throw new Error('第 ' + (i + 1) + ' 筆缺少班級或老師。');
    return { '日期': d, '時段': session, '班級': className, '核心主題': clean_(x.topic, 30), '實作名稱': clean_(x.practice, 60),
      '老師': teacher, '教室': clean_(x.room, 20), '備註': clean_(x.note, 100), '需要訂貨': true, '匯入時間': new Date() };
  });
  const classes = {};
  clean.forEach(function (r) { classes[r['班級']] = true; });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let removed = 0;
  try {
    const sh = ensureSheet_(SHEET.guoji, GUOJI_HEADERS);
    if (sh.getLastRow() >= 2) {
      const old = readTable_(SHEET.guoji, ['班級']).rows.filter(function (r) { return classes[String(r['班級']).trim()]; });
      removed = old.length;
      if (old.length) deleteRows_(sh, old.map(function (r) { return r._row; }));
    }
    appendRows_(SHEET.guoji, clean, { '日期': 'yyyy-mm-dd', '匯入時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const accounts = {};
  readTable_(SHEET.accounts, ['帳號']).rows.forEach(function (r) { accounts[String(r['帳號']).trim()] = true; });
  const noAccount = {};
  clean.forEach(function (r) { if (!accounts[r['老師']]) noAccount[r['老師']] = true; });
  return { added: clean.length, removed: removed, classes: Object.keys(classes).length, noAccount: Object.keys(noAccount) };
}

/** 管理者：某一週的國技課換老師、取消或恢復。action: 'teacher' | 'cancel' | 'restore' */
function adminUpdateGuoji(token, row, expect, action, teacherName, as) {
  const admin = auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.guoji, ['日期', '班級', '老師']);
    const r = t.rows.find(function (x) { return x._row === Number(row); });
    const d = r ? toDate_(r['日期']) : null;
    if (!r || !expect || !d || dateStr_(d) !== expect.date || String(r['班級']).trim() !== expect.className) {
      throw new Error('排班已經被改過，請重新整理再試一次。');
    }
    const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd');
    const note = String(r['備註'] || '').trim();
    if (action === 'teacher') {
      const nt = clean_(teacherName, 30);
      if (!nt) throw new Error('請選擇老師。');
      setCell_(t, r._row, '老師', nt);
      setCell_(t, r._row, '備註', (admin + ' ' + stamp + ' 由' + String(r['老師']).trim() + '換成' + nt + (note ? '；' + note : '')).slice(0, 100));
    } else if (action === 'cancel') {
      setCell_(t, r._row, '需要訂貨', false);
      setCell_(t, r._row, '備註', (admin + ' ' + stamp + ' 取消' + (note ? '；' + note : '')).slice(0, 100));
    } else if (action === 'restore') {
      setCell_(t, r._row, '需要訂貨', true);
      setCell_(t, r._row, '備註', note.replace(/^\S+ \d{4}\/\d{2}\/\d{2} 取消；?/, ''));
    } else {
      throw new Error('不支援的操作。');
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getMyLessons(token, as);
}

/** 某位老師在「課表」分頁的所有列（包含已取消的），給管理者修改用 */
function scheduleRowsOf_(teacher) {
  const sh = ss_().getSheetByName(SHEET.schedule);
  if (!sh || sh.getLastRow() < 2) return [];
  return readTable_(SHEET.schedule, ['星期', '時段', '老師']).rows.filter(function (r) {
    return String(r['老師']).trim() === teacher;
  }).map(scheduleRowOut_).sort(function (a, b) {
    return a.weekday === b.weekday ? SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session) : '一二三四五'.indexOf(a.weekday) - '一二三四五'.indexOf(b.weekday);
  });
}

function scheduleRowOut_(r) {
  const course = String(r['課程'] || '').trim();
  const need = r['需要訂貨'] === undefined || r['需要訂貨'] === '' || isTrue_(r['需要訂貨']);
  return {
    row: r._row, weekday: String(r['星期']).replace(/星期|週/g, '').trim(), session: String(r['時段']).trim(),
    room: String(r['教室'] || '').trim(), course: course, period: String(r['節次'] || '').trim(),
    className: String(r['班級'] || '').trim(), teacher: String(r['老師']).trim(),
    need: need, skipped: /國技|預留/.test(course), note: String(r['備註'] || '').trim(),
  };
}

/** 課表上有、但帳號設定沒有的老師名字（管理者的「查看老師」選單用，方便找出打錯的名字） */
function scheduleOnlyTeachers_() {
  const sh = ss_().getSheetByName(SHEET.schedule);
  if (!sh || sh.getLastRow() < 2) return [];
  const accounts = {};
  readTable_(SHEET.accounts, ['帳號']).rows.forEach(function (r) { accounts[String(r['帳號']).trim()] = true; });
  const names = {};
  readTable_(SHEET.schedule, ['老師']).rows.forEach(function (r) {
    const n = String(r['老師']).trim();
    if (n && !accounts[n]) names[n] = true;
  });
  readGuoji_().forEach(function (g) { if (g.teacher && !accounts[g.teacher]) names[g.teacher] = true; });
  return Object.keys(names).sort();
}

function adminScheduleTeachers(token) {
  auth_(token, true);
  return scheduleOnlyTeachers_();
}

function scheduleFields_(f) {
  f = f || {};
  const weekday = String(f.weekday || '').replace(/星期|週/g, '').trim();
  if (weekday.length !== 1 || '一二三四五'.indexOf(weekday) < 0) throw new Error('請選擇星期幾（一～五）。');
  const session = String(f.session || '').trim();
  if (SESSIONS.indexOf(session) < 0) throw new Error('請選擇時段（上午／下午／晚上）。');
  const course = clean_(f.course, 30);
  if (!course) throw new Error('請填寫課程。');
  const teacher = clean_(f.teacher, 30);
  if (!teacher) throw new Error('請填寫老師。');
  return { '星期': weekday, '時段': session, '教室': clean_(f.room, 20), '課程': course, '節次': clean_(f.period, 10),
    '班級': clean_(f.className, 30), '老師': teacher };
}

/**
 * 管理者修改「課表」分頁的一列。
 * action：'edit'（fields 改內容）、'cancel'（永久取消：不再列入我的課和提醒，資料保留）、'restore'（恢復）。
 * expect：畫面上看到的星期、時段、老師，跟試算表不一樣（別人改過）就不改。
 */
function adminUpdateSchedule(token, row, expect, action, fields, as) {
  const admin = auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.schedule, ['星期', '時段', '老師']);
    const r = t.rows.find(function (x) { return x._row === Number(row); });
    const cur = r ? scheduleRowOut_(r) : null;
    if (!cur || !expect || cur.weekday !== expect.weekday || cur.session !== expect.session || cur.teacher !== expect.teacher) {
      throw new Error('課表已經被改過，請重新整理再試一次。');
    }
    const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd');
    let set;
    if (action === 'cancel') {
      set = { '需要訂貨': false, '備註': (admin + ' ' + stamp + ' 取消' + (cur.note ? '；' + cur.note : '')).slice(0, 100) };
    } else if (action === 'restore') {
      set = { '需要訂貨': true, '備註': cur.note.replace(/^\S+ \d{4}\/\d{2}\/\d{2} 取消；?/, '') };
    } else {
      set = scheduleFields_(fields);
      if (fields && fields.note !== undefined) set['備註'] = clean_(fields.note, 100);
    }
    Object.keys(set).forEach(function (k) {
      if (t.cols[k] === undefined) throw new Error('「課表」分頁缺少「' + k + '」欄位。');
      setCell_(t, r._row, k, set[k]);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getMyLessons(token, as);
}

/** 管理者在「課表」分頁新增一堂課 */
function adminAddSchedule(token, fields, as) {
  auth_(token, true);
  const set = scheduleFields_(fields);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.schedule, ['星期', '時段', '教室', '課程', '節次', '班級', '老師', '需要訂貨', '備註']);
    set['需要訂貨'] = true;
    set['備註'] = clean_(fields && fields.note, 100);
    appendRows_(SHEET.schedule, [set]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getMyLessons(token, as);
}

function scheduleCount_() {
  try { return readSchedule_().length; } catch (e) { return 0; }
}

/**
 * 找出下週（週一～週五）課表上有課、但還沒有訂單的老師。
 * 有訂單 = 同一位老師（代訂的算被代訂的老師）、同一天、同一個上午／下午，而且沒有取消。
 */
function runReminder_(send) {
  if (!readSchedule_().length && !readMyLessonRows_().length) throw new Error('「課表」分頁沒有資料。');
  // 提醒的是「下一個截止」管的那一週：例如週五 10/16 提醒的是 10/26～10/30（10/19 週一中午截止）
  const now = new Date();
  const from = earliestClassDate_(now);
  const r = { from: from, to: addDays_(from, 4) };
  const deadline = addDays_(from, -7 + (LOCK_WEEKDAY - 1));
  const lessons = lessonsWithStatus_(r.from, r.to, null);
  const missing = {};
  let total = 0;
  lessons.forEach(function (x) {
    if (x.status !== 'todo') return;
    (missing[x.teacher] = missing[x.teacher] || []).push('週' + WEEKDAY[parseDate_(x.date).getDay()] + x.session + ' ' + x.course + (x.className ? ' ' + x.className : ''));
    total++;
  });
  const range = shortDateTw_(r.from) + '～' + shortDateTw_(r.to);
  const teachers = Object.keys(missing);
  const dl = shortDateTw_(deadline) + ' ' + LOCK_HOUR + ':' + ('0' + LOCK_MINUTE).slice(-2);
  const link = webAppUrl_() ? webAppUrl_() + '?page=lessons' : '';
  const personal = function (t) {
    return '【訂貨提醒】' + t + ' 老師您好，' + range + ' 還有 ' + missing[t].length + ' 堂課沒訂貨：\n' +
      missing[t].map(function (x) { return '・' + x; }).join('\n') +
      '\n\n請在 ' + dl + ' 前上網訂貨；不用訂貨的課請在「我的課」按「不用訂貨」。' + (link ? '\n' + link : '');
  };
  const bound = teachers.filter(function (t) { return !!lineUserOf_(t); });
  const unbound = teachers.filter(function (t) { return !lineUserOf_(t); });
  const summaryOf = function (verb, list) {
    return teachers.length
      ? '【訂貨提醒】' + range + ' 還沒訂貨：' + teachers.length + ' 位老師、' + total + ' 堂課（' + dl + ' 截止）。\n' +
        verb + '：' + (list.join('、') || '（無）') + '\n未綁定 LINE：' + (unbound.join('、') || '（無）')
      : '【訂貨提醒】' + range + ' 課表上的課都已經訂貨或標記不用訂貨了。';
  };
  const out = { text: summaryOf('會私訊', bound), teachers: teachers.length, lessons: total, sent: false, to: '', preview: bound.length ? personal(bound[0]) : (teachers.length ? personal(teachers[0]) : ''), bound: bound, unbound: unbound };
  if (!send) return out;

  const props = PropertiesService.getScriptProperties();
  const lineToken = props.getProperty('LINE_TOKEN');
  if (!lineToken) throw new Error('還沒設定 LINE，請先在「自動化」儲存 Channel access token。');
  // 自動煞車：本月剩下的則數不夠就不私訊老師
  const q = lineQuota_();
  let sentTo = [], skipped = [];
  bound.forEach(function (t) {
    if (q && q.limit && q.limit - q.used - sentTo.length - 1 < LINE_RESERVE) { skipped.push(t); return; }
    if (linePush_(lineUserOf_(t), personal(t))) sentTo.push(t); else skipped.push(t);
  });
  let adminText = summaryOf('已私訊', sentTo);
  if (skipped.length) adminText += '\n⚠ 本月 LINE 則數快用完（或傳送失敗），沒有私訊：' + skipped.join('、') + '，請請他們看網站「我的課」。';
  if (q && q.limit) adminText += '\n本月 LINE 已用約 ' + (q.used + sentTo.length + 1) + '／' + q.limit + ' 則。';
  if (props.getProperty('LINE_TO')) linePush_(props.getProperty('LINE_TO'), adminText);
  props.setProperty('LAST_REMIND_TIME', fmtDateTime_(new Date()));
  out.text = adminText;
  out.sent = true;
  out.to = sentTo.length + ' 位老師';
  out.sentTo = sentTo; out.skipped = skipped;
  return out;
}

// ---- LINE：老師個人綁定、則數、免費回覆、圖文選單 ----

function lineUserOf_(account) { return PropertiesService.getScriptProperties().getProperty('LU:' + account) || ''; }
function lineAccountOf_(userId) { return PropertiesService.getScriptProperties().getProperty('LA:' + userId) || ''; }

/** 主動推播一則（會算則數）。成功回傳 true。 */
function linePush_(to, text) {
  const lineToken = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!lineToken || !to) return false;
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + lineToken },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] }),
  });
  return res.getResponseCode() === 200;
}

/** 回覆（免費，不算則數） */
function lineReply_(replyToken, text) {
  const lineToken = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!lineToken || !replyToken) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + lineToken },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] }),
  });
}

/** 本月則數：{ limit, used }（查不到回傳 null） */
function lineQuota_() {
  const lineToken = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!lineToken) return null;
  try {
    const h = { headers: { Authorization: 'Bearer ' + lineToken }, muteHttpExceptions: true };
    const a = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/quota', h);
    const b = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/quota/consumption', h);
    if (a.getResponseCode() !== 200 || b.getResponseCode() !== 200) return null;
    const qa = JSON.parse(a.getContentText()), qb = JSON.parse(b.getContentText());
    return { limit: qa.type === 'limited' ? Number(qa.value) : 0, used: Number(qb.totalUsage) || 0 };
  } catch (e) {
    return null;
  }
}

/** 官方帳號的加好友連結（第一次查完記起來） */
function lineFriendUrl_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('LINE_BASIC_ID');
  const lineToken = props.getProperty('LINE_TOKEN');
  if (!id && lineToken) {
    try {
      const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', { headers: { Authorization: 'Bearer ' + lineToken }, muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        id = JSON.parse(res.getContentText()).basicId || '';
        if (id) props.setProperty('LINE_BASIC_ID', id);
      }
    } catch (e) {}
  }
  return id ? 'https://line.me/R/ti/p/' + encodeURIComponent(id) : '';
}

/** 老師：產生自己的 LINE 綁定碼 */
function myLineBindCode(token) {
  const account = auth_(token);
  if (!PropertiesService.getScriptProperties().getProperty('LINE_TOKEN')) throw new Error('管理者還沒設定 LINE 官方帳號。');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('tbind:' + code, account, LINE_BIND_MINUTES * 60);
  return { code: code, minutes: LINE_BIND_MINUTES, friendUrl: lineFriendUrl_() };
}

function myLineUnbind(token) {
  const account = auth_(token);
  const props = PropertiesService.getScriptProperties();
  const uid = props.getProperty('LU:' + account);
  if (uid) { props.deleteProperty('LA:' + uid); props.deleteProperty('LU:' + account); }
  return true;
}

/** 管理者：哪些老師綁了 LINE、本月則數 */
function lineStatus_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const bound = Object.keys(props).filter(function (k) { return k.indexOf('LU:') === 0; }).map(function (k) { return k.slice(3); }).sort();
  return { teachersBound: bound, quota: lineQuota_(), richMenu: props.RICH_MENU_ID || '' };
}

/** 老師在官方帳號按選單（或打字）時的免費回覆 */
function lineAutoReply_(text, userId) {
  const t = String(text || '').trim();
  const account = lineAccountOf_(userId);
  const url = webAppUrl_();
  if (/截止|規則/.test(t) && t.length <= 10) {
    const now = new Date(), from = earliestClassDate_(now), dl = currentDeadline_(now) > now ? currentDeadline_(now) : addDays_(currentDeadline_(now), 7);
    return '【訂貨規則】' + lockRuleText_() + '。\n' +
      '下一個截止：' + shortDateTw_(dl) + ' ' + LOCK_HOUR + ':' + ('0' + LOCK_MINUTE).slice(-2) + '\n' +
      '目前最早可以訂：' + shortDateTw_(from) + ' 的課，最晚可以訂 ' + MAX_DAYS_AHEAD + ' 天內的課。' + (url ? '\n' + url : '');
  }
  if (/網站|訂貨系統/.test(t) && t.length <= 10) return url ? '訂貨網站：' + url : '訂貨網站還沒部署。';
  const wantLessons = /未訂貨|還沒訂|我的課/.test(t) && t.length <= 12;
  const wantOrders = /我的訂單|訂單狀態/.test(t) && t.length <= 12;
  if (!wantLessons && !wantOrders) return '';
  if (!account) return '還沒綁定帳號，請先登入訂貨網站 →「我的課」→「綁定我的 LINE」，再把綁定碼傳給我。' + (url ? '\n' + url : '');
  if (wantLessons) {
    const from = earliestClassDate_(new Date()), to = addDays_(today_(), MAX_DAYS_AHEAD);
    const todo = lessonsWithStatus_(from, to, account).filter(function (x) { return x.status === 'todo'; })
      .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session); });
    if (!todo.length) return account + ' 老師，接下來 ' + MAX_DAYS_AHEAD + ' 天的課都已經訂貨或標記不用訂貨了。';
    return account + ' 老師，還有 ' + todo.length + ' 堂課沒訂貨：\n' + todo.slice(0, 20).map(function (x) {
      return '・' + shortDateTw_(parseDate_(x.date)) + x.session + ' ' + x.course + (x.className ? ' ' + x.className : '');
    }).join('\n') + (todo.length > 20 ? '\n…' : '') + (url ? '\n\n去訂貨：' + url + '?page=lessons' : '');
  }
  const since = addDays_(today_(), -7);
  const rows = readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows.filter(function (r) {
    const d = toDate_(r['使用日期']);
    return (String(r['操作老師']).trim() === account || String(r['代訂對象老師'] || '').trim() === account) && d && d >= since;
  });
  const orders = groupOrders_(rows, account).sort(function (a, b) { return a.useDate < b.useDate ? -1 : 1; });
  if (!orders.length) return account + ' 老師，最近沒有訂單。' + (url ? '\n' + url : '');
  return account + ' 老師最近的訂單：\n' + orders.slice(0, 15).map(function (o) {
    return '・' + shortDateTw_(parseDate_(o.useDate)) + ' ' + o.className + '「' + o.project + '」' + o.status;
  }).join('\n') + (url ? '\n\n' + url : '');
}

/** 管理者：建立官方帳號圖文選單（4 個按鈕）並設為預設 */
function adminSetupRichMenu(token) {
  auth_(token, true);
  const props = PropertiesService.getScriptProperties();
  const lineToken = props.getProperty('LINE_TOKEN');
  if (!lineToken) throw new Error('請先儲存 Channel access token。');
  const url = webAppUrl_();
  if (!url) throw new Error('請先部署網站，再回來設定選單。');
  const auth = { Authorization: 'Bearer ' + lineToken };
  const w = 2500, h = 843, cw = Math.floor(w / 4);
  const areas = [
    { action: { type: 'message', label: '我的未訂貨課程', text: '我的未訂貨課程' } },
    { action: { type: 'uri', label: '打開訂貨網站', uri: url } },
    { action: { type: 'message', label: '我的訂單', text: '我的訂單' } },
    { action: { type: 'message', label: '截止時間說明', text: '截止時間說明' } },
  ].map(function (a, i) { a.bounds = { x: i * cw, y: 0, width: i === 3 ? w - 3 * cw : cw, height: h }; return a; });
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: auth,
    payload: JSON.stringify({ size: { width: w, height: h }, selected: true, name: '訂貨系統選單', chatBarText: '訂貨選單', areas: areas }),
  });
  if (res.getResponseCode() !== 200) throw new Error('建立選單失敗：' + res.getContentText().slice(0, 200));
  const id = JSON.parse(res.getContentText()).richMenuId;
  const png = Utilities.base64Decode(HtmlService.createHtmlOutputFromFile('RichMenu').getContent().replace(/\s+/g, ''));
  const up = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/' + id + '/content', {
    method: 'post', contentType: 'image/png', muteHttpExceptions: true, headers: auth, payload: png,
  });
  if (up.getResponseCode() !== 200) throw new Error('上傳選單圖片失敗：' + up.getContentText().slice(0, 200));
  const def = UrlFetchApp.fetch('https://api.line.me/v2/bot/user/all/richmenu/' + id, { method: 'post', muteHttpExceptions: true, headers: auth });
  if (def.getResponseCode() !== 200) throw new Error('設定預設選單失敗：' + def.getContentText().slice(0, 200));
  // 刪掉舊的選單
  const old = props.getProperty('RICH_MENU_ID');
  if (old && old !== id) UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/' + old, { method: 'delete', muteHttpExceptions: true, headers: auth });
  props.setProperty('RICH_MENU_ID', id);
  return automationStatus_();
}

/** 只傳 LINE（沒綁定就不傳）。 */
function notifyLine_(text) {
  const props = PropertiesService.getScriptProperties();
  const lineToken = props.getProperty('LINE_TOKEN'), lineTo = props.getProperty('LINE_TO');
  if (!lineToken || !lineTo) return '尚未綁定 LINE';
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + lineToken },
    payload: JSON.stringify({ to: lineTo, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] }),
  });
  return res.getResponseCode() === 200 ? '已傳到 LINE' : 'LINE 傳送失敗（錯誤碼 ' + res.getResponseCode() + '）';
}

// ---- 廠商訂貨單 ----

/** 「下週」：今天之後的第一個星期一到星期五（今天是星期一也算下週一） */
function nextWeekRange_() {
  const today = today_();
  const monday = addDays_(today, ((8 - today.getDay()) % 7) || 7);
  return { from: monday, to: addDays_(monday, 4) };
}

const GRAMS_ = { 'g': 1, '公斤': 1000, '斤': 600, '兩': 37.5 };
function convertUnit_(qty, unit, vendor) {
  if (!GRAMS_[unit]) return { qty: qty, unit: unit };
  const target = /^[AB]/.test(vendor || '') ? '斤' : '公斤';
  return { qty: Math.round(qty * GRAMS_[unit] / GRAMS_[target] * 100) / 100, unit: target };
}

function ordersBetween_(from, to) {
  const rows = readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows.filter(function (r) {
    const d = toDate_(r['使用日期']);
    return d && d >= from && d <= to;
  });
  return groupOrders_(rows, null);
}

/** 依廠商整理：{ 廠商: { totals: [...], lines: [...] } }（不含已取消、有貨、器具） */
function vendorBook_(orders) {
  const by = {};
  orders.forEach(function (o) {
    if (o.status === '已取消') return;
    o.items.forEach(function (it) {
      if (it.id === TOOL_ID || it.noOrder || it.status === '已取消') return;
      const v = by[it.vendor] = by[it.vendor] || { totals: {}, lines: [] };
      const cv = convertUnit_(it.qty, it.unit, it.vendor);
      const k = it.name + '|' + cv.unit;
      v.totals[k] = v.totals[k] || { name: it.name, unit: cv.unit, qty: 0 };
      v.totals[k].qty = Math.round((v.totals[k].qty + cv.qty) * 100) / 100;
      v.lines.push([o.useDate, it.group, o.className, it.name, it.qty, it.unit, o.onBehalfOf || o.teacher, it.note]);
    });
  });
  Object.keys(by).forEach(function (v) {
    by[v].totals = Object.keys(by[v].totals).sort().map(function (k) { return by[v].totals[k]; });
    by[v].lines.sort(function (a, b) { return (a[0] + a[2] + a[3]).localeCompare(b[0] + b[2] + b[3]); });
  });
  return by;
}

function runVendorOrders_() {
  const r = nextWeekRange_();
  const range = shortDateTw_(r.from) + '～' + shortDateTw_(r.to);
  const book = vendorBook_(ordersBetween_(r.from, r.to));
  const vendors = Object.keys(book).sort();
  if (!vendors.length) {
    const n = notify_('【訂貨單】' + range + ' 沒有需要叫貨的品項',
      '<p>' + range + ' 目前沒有需要跟廠商叫貨的品項，所以沒有產生訂貨單。</p>',
      '【訂貨單】' + range + ' 沒有需要叫貨的品項，沒有產生訂貨單。');
    return { url: '', vendors: [], notify: n };
  }

  const name = '廠商訂貨單 ' + dateStr_(r.from) + '～' + dateStr_(r.to);
  const book_ss = SpreadsheetApp.create(name);
  const first = book_ss.getSheets()[0];
  vendors.forEach(function (v) {
    const sh = book_ss.insertSheet(v.slice(0, 90));
    const d = book[v];
    const rows = [[v + ' 訂貨單', '', '', '', '', '', '', ''], [range, '', '', '', '', '', '', ''], ['', '', '', '', '', '', '', ''],
      ['【食材合計】', '', '', '', '', '', '', ''], ['食材', '單位', '合計', '', '', '', '', '']];
    d.totals.forEach(function (t) { rows.push([t.name, t.unit, t.qty, '', '', '', '', '']); });
    rows.push(['', '', '', '', '', '', '', '']);
    rows.push(['【明細】', '', '', '', '', '', '', '']);
    const detailHead = rows.length + 1;
    rows.push(['上課日期', '組別', '上課班級', '食材', '數量', '單位', '老師', '備註']);
    d.lines.forEach(function (l) { rows.push(l); });
    sh.getRange(1, 1, rows.length, 8).setNumberFormat('@').setValues(rows.map(function (row) {
      return row.map(function (c) { return c === '' ? '' : String(c); });
    }));
    sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    sh.getRange(5, 1, 1, 3).setFontWeight('bold').setBackground('#eeeeee');
    sh.getRange(detailHead, 1, 1, 8).setFontWeight('bold').setBackground('#eeeeee');
    sh.setColumnWidth(1, 140); sh.setColumnWidth(4, 160); sh.setColumnWidth(8, 160);
  });
  book_ss.deleteSheet(first);

  const file = DriveApp.getFileById(book_ss.getId());
  const folders = DriveApp.getFoldersByName(AUTO_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(AUTO_FOLDER_NAME);
  file.moveTo(folder);
  const url = book_ss.getUrl();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_VENDOR_URL', url);
  props.setProperty('LAST_VENDOR_TIME', fmtDateTime_(new Date()));

  const summary = vendors.map(function (v) { return v + ' ' + book[v].totals.length + ' 項'; });
  const n = notify_('【訂貨單】' + range + ' 廠商訂貨單已產生',
    '<p>' + range + ' 的訂貨已截止，廠商訂貨單已經產生：</p><ul>' +
      summary.map(function (x) { return '<li>' + escHtml_(x) + '</li>'; }).join('') + '</ul>' +
      '<p><a href="' + url + '">打開廠商訂貨單（Google 試算表）</a></p>' +
      '<p>每家廠商一個分頁，上面是食材合計、下面是各班明細。</p>',
    '【訂貨單已產生】' + range + '\n' + summary.join('\n') + '\n\n' + url);
  return { url: url, vendors: summary, notify: n };
}

// ---- 每日摘要 ----

function runDigest_(force) {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const since = props.getProperty('DIGEST_LAST') ? new Date(Number(props.getProperty('DIGEST_LAST'))) : new Date(now.getTime() - 24 * 3600 * 1000);
  const all = groupOrders_(readTable_(SHEET.orders, ['訂單編號']).rows, null);
  const newOrders = all.filter(function (o) {
    const t = parseDateTime_(o.submitted);
    return t && t > since;
  });
  const pendItems = readTable_(SHEET.pending, ['食材名稱', '審核狀態']).rows.filter(function (r) {
    const s = String(r['審核狀態'] || '').trim(); return !s || s === '待審核';
  }).length;
  const pendUnits = readCustomUnits_().filter(function (r) { return r.status === '待審核'; }).length;
  const r = nextWeekRange_();
  const nextWeek = all.filter(function (o) {
    const d = parseDate_(o.useDate);
    return d && d >= r.from && d <= r.to && o.status !== '已取消';
  });

  const edits = JSON.parse(props.getProperty('ORDER_EDITS') || '[]');
  const errors = recentSystemErrors_(since);
  if (!force && !newOrders.length && !pendItems && !pendUnits && !edits.length && !errors.length) {
    props.setProperty('DIGEST_LAST', String(now.getTime()));
    return { sent: false, reason: '沒有新訂單或待審核項目，今天不寄。' };
  }

  const day = shortDateTw_(today_());
  const lineOrders = newOrders.slice(0, 15).map(function (o) {
    return '・' + (o.onBehalfOf || o.teacher) + ' ' + o.className + ' ' + o.project + '（' + shortDateTw_(parseDate_(o.useDate)) + (o.status === '已取消' ? '，已取消' : '') + '）';
  });
  const editLines = edits.slice(-10).map(function (e) { return '・' + e.who + ' 修改 ' + e.no + '（' + e.text + '）'; });
  const lineText = '【訂貨每日摘要 ' + day + '】\n' +
    (edits.length ? '老師修改了 ' + edits.length + ' 張訂單：\n' + editLines.join('\n') + (edits.length > 10 ? '\n…' : '') + '\n' : '') +
    '新訂單 ' + newOrders.length + ' 張' + (lineOrders.length ? '：\n' + lineOrders.join('\n') + (newOrders.length > 15 ? '\n…' : '') : '') + '\n' +
    '待審核：新食材 ' + pendItems + '、新單位 ' + pendUnits + '\n' +
    '下週（' + shortDateTw_(r.from) + '～' + shortDateTw_(r.to) + '）已收到 ' + nextWeek.length + ' 張訂單' +
    (errors.length ? '\n⚠ 系統錯誤 ' + errors.length + ' 筆（詳見試算表「錯誤紀錄」分頁）：\n' + errors.slice(-3).map(function (e) {
      return '・' + e.account + ' ' + e.fn + '：' + e.message.slice(0, 60);
    }).join('\n') : '') +
    (webAppUrl_() ? '\n\n' + webAppUrl_() : '');
  const html = '<h3>訂貨每日摘要 ' + escHtml_(day) + '</h3>' +
    '<p><b>新訂單 ' + newOrders.length + ' 張</b>（' + escHtml_(fmtDateTime_(since)) + ' 之後送出）</p>' +
    (newOrders.length ? '<table border="1" cellpadding="4" style="border-collapse:collapse;font-size:14px"><tr><th>送出時間</th><th>老師</th><th>班級</th><th>實作項目</th><th>上課日期</th><th>品項</th><th>狀態</th></tr>' +
      newOrders.map(function (o) {
        return '<tr><td>' + escHtml_(o.submitted) + '</td><td>' + escHtml_(o.onBehalfOf || o.teacher) + '</td><td>' + escHtml_(o.className) + '</td><td>' +
          escHtml_(o.project) + '</td><td>' + escHtml_(o.useDate) + '</td><td>' + o.items.length + '</td><td>' + escHtml_(o.status) + '</td></tr>';
      }).join('') + '</table>' : '') +
    '<p><b>待審核：</b>新食材 ' + pendItems + ' 項、新單位 ' + pendUnits + ' 個</p>' +
    '<p><b>下週</b>（' + escHtml_(shortDateTw_(r.from) + '～' + shortDateTw_(r.to)) + '）已收到 ' + nextWeek.length + ' 張訂單；' + escHtml_(lockRuleText_()) + '。</p>' +
    (webAppUrl_() ? '<p><a href="' + webAppUrl_() + '">打開訂貨網站</a></p>' : '');
  const editHtml = edits.length ? '<p><b>老師修改的訂單 ' + edits.length + ' 張</b></p><ul>' +
    edits.map(function (e) { return '<li>' + escHtml_(e.t + ' ' + e.who + ' 修改 ' + e.no + '：' + e.text) + '</li>'; }).join('') + '</ul>' : '';
  const errHtml = errors.length ? '<p><b style="color:#b91c1c">系統錯誤 ' + errors.length + ' 筆</b>（試算表「錯誤紀錄」分頁有完整紀錄）</p><ul>' +
    errors.slice(-20).map(function (e) { return '<li>' + escHtml_(e.time + ' ' + e.account + ' ' + e.fn + '：' + e.message) + '</li>'; }).join('') + '</ul>' : '';
  const n = notify_('【訂貨每日摘要】' + day + '：新訂單 ' + newOrders.length + ' 張' + (edits.length ? '、修改 ' + edits.length + ' 張' : '') + (errors.length ? '、系統錯誤 ' + errors.length + ' 筆' : ''), html + editHtml + errHtml, lineText);
  props.setProperty('DIGEST_LAST', String(now.getTime()));
  props.deleteProperty('ORDER_EDITS');
  return { sent: true, newOrders: newOrders.length, pendItems: pendItems, pendUnits: pendUnits, errors: errors.length, notify: n };
}

function shortDateTw_(d) {
  return d ? (d.getMonth() + 1) + '/' + d.getDate() + '（' + WEEKDAY[d.getDay()] + '）' : '';
}

function parseDateTime_(s) {
  const m = String(s || '').match(/^(\d{4})\/(\d{2})\/(\d{2}) (\d{2}):(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) : null;
}

function escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
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
        groups: r['組數'] === '' || r['組數'] === undefined ? '' : Number(r['組數']) || '',
        reviewer: String(r['審核人'] || ''),
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
      noOrder: isTrue_(r['不叫貨']),
      section: String(r['段落'] || ''),
      row: r._archived ? null : r._row,
    });
    if (r._archived) map[no].archived = true;
  });
  list.forEach(function (o) {
    // 狀態以沒取消的品項為準（管理者刪掉個別品項時，訂單本身不算「部分處理」）
    let s = o.items.map(function (i) { return i.status; });
    const active = s.filter(function (x) { return x !== '已取消'; });
    if (active.length) s = active;
    o.status = s.every(function (x) { return x === s[0]; }) ? s[0] : '部分處理';
    o.archived = !!o.archived;
    o.locked = o.archived || isLocked_(parseDate_(o.useDate));
    o.canCancel = !!account && o.status === '待處理' && !o.locked;
    o.canEdit = o.canCancel;
  });
  return list;
}

function updateOrderStatus_(orderNo, status, check, reviewer) {
  orderNo = String(orderNo || '').trim();
  if (!orderNo) throw new Error('缺少訂單編號。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '狀態']);
    const rows = t.rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo; });
    if (!rows.length) throw new Error('找不到訂單 ' + orderNo + '。');
    if (check) check(rows);
    rows.forEach(function (r) {
      // 管理者已個別取消的品項保持取消
      if (String(r['狀態']).trim() === '已取消' && rows.some(function (x) { return String(x['狀態']).trim() !== '已取消'; })) return;
      setCell_(t, r._row, '狀態', status);
      if (reviewer && t.cols['審核人'] !== undefined) setCell_(t, r._row, '審核人', reviewer);
    });
    SpreadsheetApp.flush();
    return status;
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// 共用：試算表讀寫（依標題找欄位）
// =====================================================================

// 同一次執行只開一次試算表（重複 openById 容易出現「試算表服務無法正常運作」）
let SS_CACHE_ = null;
function ss_() {
  if (!SS_CACHE_) {
    SS_CACHE_ = retry_(function () {
      return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    });
  }
  return SS_CACHE_;
}

/** Google 服務暫時忙碌時自動重試（最多 3 次） */
function retry_(fn) {
  let lastErr;
  for (let i = 0; i < 3; i++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      const m = String(e && e.message || e);
      if (!/Service|服務|無法正常運作|timed out|逾時|Internal error|內部錯誤|try again|稍後再試/i.test(m)) throw e;
      Utilities.sleep(800 * (i + 1));
    }
  }
  throw new Error('Google 試算表暫時沒有回應，請稍等一下再試一次。（' + (lastErr && lastErr.message) + '）');
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
  const values = lastRow ? retry_(function () { return sh.getRange(1, 1, lastRow, lastCol).getValues(); }) : [[]];
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
  retry_(function () { sh.getRange(start, 1, objs.length, headers.length).setValues(data); });
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
