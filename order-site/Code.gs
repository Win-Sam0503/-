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
const ADMIN_ACCOUNTS = ['管理者'];          // 這些帳號登入後會看到管理功能（其他老師可以在網站上設成管理者，存在「帳號設定」的「管理者」欄）
const DEPARTMENTS = ['中餐', '西餐', '日料', '烘焙', '飲調', '果雕', '刀功'];
// 上課時間：只選時段「上午／下午／晚上」（夜間＝晚上）；以前的訂單有「下午 1:05」這種幾點幾分，也還是看得懂
const CLASS_TIME_PATTERN = /^(上午|下午|晚上)( (1[0-2]|[1-9]):[0-5][05])?$/;
const SESSIONS = ['上午', '下午', '晚上'];
// 課表類別：總課表分頁籤；空白＝日間部。彈性課程可以設開始、結束日期（只有那幾週）
const SCHEDULE_KINDS = ['日間部', '夜間部', '彈性課程'];
const SCHEDULE_HEADERS = ['星期', '時段', '教室', '課程', '節次', '班級', '老師', '需要訂貨', '備註', '類別', '開始日期', '結束日期'];
// 鎖單：每週一 12:00 截止「下週」的訂單；截止後最早只能訂下下週的課。管理者不受限制。
// 這三個是預設值；管理者可以在「系統設定 → 訂貨截止時間」改（存在指令碼屬性 LOCK_RULE），程式裡一律用 lockRule_()
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
const LINE_BIND_MINUTES = 30;                // LINE 綁定碼有效時間
const LINE_RESERVE = 20;                     // LINE 本月剩下的則數少於這個數字時，不再私訊老師（留給管理者通知）
const REMIND_HOUR = 15;                     // 截止前一個上課日（截止週一＝週五）幾點用 LINE 提醒還沒訂貨的老師
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
  purchase: '叫貨紀錄',
  prints: '列印紀錄',
  flags: '訂單異常',
  contacts: '廠商聯絡',
  holidays: '放假日',
  subs: '科別替換',
  pickups: '跨教室領料',
  audits: '抽查紀錄',
  places: '放置位置',
  hidden: '老師刪除',
  issuesDone: '異常處理',
};
// 科別替換：同一個叫法不同科是不同東西（例如烘焙課的「牛奶」是保久乳、飲調課的是鮮奶(全)）
const SUB_HEADERS = ['科別', '食材', '換成', '備註'];
// 跨教室領料單：管理者自己加的（系統自動抓不到的，例如臨時跟別間教室借的食材）
const PICKUP_PLACES = ['烘焙教室', '中餐教室', '飲調教室'];   // 跨教室領料只有這三間
const PICKUP_HEADERS = ['使用日期', '上課時間', '班級', '老師', '科別', '教室', '領取地點', '食材名稱', '數量', '單位', '備註', '建立者', '建立時間'];
// 抽查紀錄：管理者隨機抽查分好的食材，一班一個抽查編號、每樣食材一列
const AUDIT_HEADERS = ['抽查編號', '抽查時間', '時間點', '冰箱', '訂單編號', '上課日期', '上課時間', '班級', '科別', '老師', '食材名稱', '清單數量', '單位', '實際數量', '結果', '封條', '說明', '照片', '抽查人', '陪同人', '抽法'];
const AUDIT_TIMINGS = ['分貨後、上課前', '下課後（剩料）'];
const AUDIT_RESULTS = ['正常', '少了', '多了', '沒有', '不確定'];
const AUDIT_SEALS = ['完整', '破了', '沒封', ''];
const AUDIT_FOLDER_NAME = '三信家商餐飲科抽查照片';
// 放置位置：校隊分好貨後登記每張訂單的東西放在哪裡（同一張訂單可以登記好幾次，以最後一次為準）
const PLACE_VENDORS = /^[ABCD]/;   // 要冰的才登記：A 菜、B 肉、C 蛋、D 海鮮（E 乾貨、F 飲調不用）
const PLACE_HEADERS = ['訂單編號', '上課日期', '班級', '老師', '放置位置', '照片', '登記人', '登記時間'];
const PLACE_FOLDER_NAME = '三信家商餐飲科放置照片';
// 老師刪除：老師從「我的訂單」刪掉的訂單（老師看不到，管理者照樣看得到、標「老師已刪除」）
const HIDDEN_HEADERS = ['訂單編號', '老師', '刪除時間', '原狀態'];
// 異常紀錄：訂單彙整的「⚠ 異常紀錄」看過、勾「已處理」的事件
const ISSUE_DONE_HEADERS = ['事件ID', '處理人', '處理時間'];
// 叫貨紀錄：學生扣完庫存後，管理者確定的「實際叫貨量」（老師訂的量不變）
const PURCHASE_HEADERS = ['送貨日期', '送貨組別', '廠商', '訂單編號', '上課班級', '老師', '食材名稱', '食材ID', '需求量', '需求單位', '叫貨量', '叫貨單位', '類型', '備註', '狀態', '確定時間', '確認人'];
const PURCHASE_MAX_ROWS = 3000;
const GUOJI_HEADERS = ['日期', '時段', '班級', '核心主題', '實作名稱', '老師', '教室', '備註', '需要訂貨', '匯入時間'];
const EDIT_HEADERS = ['時間', '訂單編號', '修改人', '內容'];
// 列印紀錄：管理者在訂單彙整印過的料單，每張訂單一列（最後一次列印）
const PRINT_HEADERS = ['訂單編號', '列印時間', '列印人'];
const FLAG_HEADERS = ['訂單編號', '說明', '標記人', '標記時間', '狀態', '解決時間', '解決人', '食材'];   // 狀態：異常／已解決；食材空白＝整張訂單
// 廠商聯絡：傳真、電話，印在叫貨單上
const CONTACT_HEADERS = ['廠商', '傳真', '電話', '聯絡人', '備註'];
// 放假日（國定假日、補假、校慶補假…）：這幾天的課不算未訂貨、不提醒
const HOLIDAY_HEADERS = ['開始日期', '結束日期', '名稱'];
const ERROR_HEADERS = ['時間', '帳號', '功能', '類型', '訊息', '畫面'];
const ERROR_KEEP = 2000;                      // 錯誤紀錄最多保留幾列
const MY_LESSON_HEADERS = ['老師', '類型', '星期', '日期', '時段', '課程', '班級', '教室', '建立時間'];
// 「不用訂貨」分頁也放管理者手動標的「已叫料」（標記欄）
const SKIP_HEADERS = ['老師', '日期', '時段', '課程', '標記時間', '標記', '備註', '標記人'];

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

function doGet(e) {
  // 一鍵更新後的自動測試：程式讀得到、畫面產生得出來就回 OK（不含任何資料）
  if (e && e.parameter && e.parameter.health) {
    let ok = false;
    try { ok = HtmlService.createHtmlOutputFromFile('Index').getContent().length > 1000; } catch (err) { ok = false; }
    return ContentService.createTextOutput(ok ? 'ORDER_SITE_OK' : 'ORDER_SITE_FAIL');
  }
  // 網頁沒有用到範本語法，直接送出（不經過範本處理，比較快）
  return HtmlService.createHtmlOutputFromFile('Index')
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
  return { token: token, account: account, isAdmin: isAdmin_(account), mustChange: mustChange };
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
  const account = auth_(token, false, true, true);
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
/** 是不是管理者：內建的「管理者」帳號，或「帳號設定」裡「管理者」欄打勾、而且啟用的老師 */
function isAdmin_(account) {
  if (!account) return false;
  if (ADMIN_ACCOUNTS.indexOf(account) >= 0) return true;
  return extraAdmins_().indexOf(account) >= 0;
}
function extraAdmins_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('extraAdmins');
  if (hit) return JSON.parse(hit);
  const t = readTable_(SHEET.accounts, ['帳號']);
  const list = t.cols['管理者'] === undefined ? [] : t.rows.filter(function (r) {
    return isTrue_(r['管理者']) && isTrue_(r['啟用']);
  }).map(function (r) { return String(r['帳號']).trim(); }).filter(Boolean);
  cache.put('extraAdmins', JSON.stringify(list), 300);
  return list;
}

/** 管理者：目前的管理者和可以設定的老師 */
function adminListAdmins(token) {
  return listAdmins_(auth_(token, true));
}
function listAdmins_(me) {
  const t = readTable_(SHEET.accounts, ['帳號', '啟用']);
  const teachers = t.rows.filter(function (r) { return isTrue_(r['啟用']); }).map(function (r) { return String(r['帳號']).trim(); })
    .filter(function (a) { return a && ADMIN_ACCOUNTS.indexOf(a) < 0 && crewAccounts_().indexOf(a) < 0; });
  return { builtin: ADMIN_ACCOUNTS.slice(), extra: extraAdmins_(), teachers: teachers, me: me };
}

/** 管理者：把某位老師設成管理者（on = true）或取消 */
function adminSetAdmin(token, account, on) {
  const me = auth_(token, true);
  account = clean_(account, 30);
  if (ADMIN_ACCOUNTS.indexOf(account) >= 0) throw new Error('「' + account + '」是內建的管理者，不能取消。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(sheet_(SHEET.accounts), ['管理者']);
    const t = readTable_(SHEET.accounts, ['帳號']);
    const r = t.rows.find(function (x) { return String(x['帳號']).trim() === account; });
    if (!r) throw new Error('「帳號設定」裡找不到「' + account + '」。');
    if (on && !isTrue_(r['啟用'])) throw new Error('「' + account + '」的帳號沒有啟用。');
    setCell_(t, r._row, '管理者', !!on);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove('extraAdmins');
  const out = listAdmins_(me);
  out.selfRemoved = !on && account === me;
  return out;
}

function auth_(token, needAdmin, allowMustChange, allowCrew) {
  const cache = CacheService.getScriptCache();
  const account = token ? cache.get('session:' + token) : null;
  if (!account) throw new Error('SESSION_EXPIRED');
  // 還在用預設密碼的，除了改密碼以外什麼都不能做
  if (!allowMustChange && cache.get('mustchange:' + token)) throw new Error('MUST_CHANGE_PASSWORD');
  if (needAdmin && !isAdmin_(account)) throw new Error('只有管理者可以使用這個功能。');
  // 校隊帳號只能登記放置位置、改密碼，不能訂貨、看訂單
  if (!needAdmin && !allowCrew && !isAdmin_(account) && isCrew_(account)) throw new Error('校隊帳號只能使用「放置登記」。');
  return account;
}

/** 校隊帳號：「帳號設定」的「校隊」欄打勾、而且啟用 */
function isCrew_(account) {
  return !!account && crewAccounts_().indexOf(account) >= 0;
}
function crewAccounts_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('crewAccounts');
  if (hit) return JSON.parse(hit);
  const t = readTable_(SHEET.accounts, ['帳號']);
  const list = t.cols['校隊'] === undefined ? [] : t.rows.filter(function (r) {
    return isTrue_(r['校隊']) && isTrue_(r['啟用']);
  }).map(function (r) { return String(r['帳號']).trim(); }).filter(Boolean);
  cache.put('crewAccounts', JSON.stringify(list), 300);
  return list;
}

// =====================================================================
// 老師：讀取資料、下單、查詢、取消
// =====================================================================

/** 登入後一次載入網站需要的資料。 */
// 打開網站時要讀的、不常變的資料（老師名單、單位、科別替換、廠商聯絡、放假日）：暫存 5 分鐘，
// 網站上改到這幾張分頁時自動清掉（appendRows_／setCell_／deleteRows_）；直接在試算表改的，最慢 5 分鐘後生效
const BOOT_CACHE_KEY = 'boot:static:v1';
const BOOT_SHEETS = ['帳號設定', '單位', '科別替換', '廠商聯絡', '放假日'];
function bootStatic_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(BOOT_CACHE_KEY);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const u = unitLists_();
  const st = {
    teachers: readTable_(SHEET.accounts, ['帳號', '啟用']).rows
      .filter(function (r) { return isTrue_(r['啟用']); })
      .map(function (r) { return String(r['帳號']).trim(); })
      .filter(function (a) { return a && ADMIN_ACCOUNTS.indexOf(a) < 0 && crewAccounts_().indexOf(a) < 0; }),
    units: { order: u.order, recipe: u.recipe, pending: u.pending },
    deptSubs: readDeptSubs_(),
    missingSetup: missingSetup_(),
    vendorContacts: readVendorContacts_(),
    holidays: holidaysBetween_(addDays_(today_(), -7), addDays_(today_(), MAX_DAYS_AHEAD + 30)),
    prices: orderPrices_(),
  };
  try { const j = JSON.stringify(st); if (j.length < 90000) cache.put(BOOT_CACHE_KEY, j, 300); } catch (e) {}
  return st;
}
/**
 * 單價表（進貨驗收網站記的）：{ 食材名稱: [{ v: 廠商, u: 單位, p: 單價, c: { q, u } 換算（1 個單位＝多少）, d: 更新日期 }] }
 * 老師訂貨時顯示單價和大約金額。同一家同一單位只留最新的。
 */
function orderPrices_() {
  const out = {};
  try {
    if (!ss_().getSheetByName('單價表')) return out;
    const seen = {};
    readTable_('單價表', ['廠商', '食材名稱', '單價']).rows.forEach(function (r) {
      const n = String(r['食材名稱']).trim(), v = String(r['廠商']).trim(), u = String(r['單位'] || '').trim(), p = Number(r['單價']);
      if (!n || !(p > 0)) return;
      const d = r['更新日期'] instanceof Date ? dateStr_(toDate_(r['更新日期'])) : String(r['更新日期'] || '').trim();
      const m = String(r['換算'] || '').trim().match(/^([\d.]+)\s*([^\d\s].*)$/);
      const x = { v: v, u: u, p: p, c: m && Number(m[1]) > 0 ? { q: Number(m[1]), u: m[2].trim() } : null, d: d };
      const k = v + '|' + n + '|' + u;
      if (seen[k] && seen[k].d >= d) return;
      if (seen[k]) out[n].splice(out[n].indexOf(seen[k]), 1);
      seen[k] = x;
      (out[n] = out[n] || []).push(x);
    });
  } catch (e) { console.error('orderPrices_ ' + e.message); }
  return out;
}

function bootDirty_(name) {
  if (BOOT_SHEETS.indexOf(String(name)) < 0) return;
  try { CacheService.getScriptCache().remove(BOOT_CACHE_KEY); } catch (e) {}
}

function getBootstrap(token) {
  const account = auth_(token, false, false, true);
  if (!isAdmin_(account) && isCrew_(account)) return { account: account, isCrew: true, isAdmin: false, items: [], today: dateStr_(today_()), placeEnabled: placeEnabled_() };
  const items = cachedItems_();
  const vendors = [];
  items.forEach(function (it) { if (it.vendor && vendors.indexOf(it.vendor) < 0) vendors.push(it.vendor); });
  vendors.sort();

  const st = bootStatic_();
  const teachers = st.teachers;

  const depts = DEPARTMENTS.slice();
  items.forEach(function (it) { if (it.dept !== '通用' && depts.indexOf(it.dept) < 0) depts.push(it.dept); });

  const isAdmin = isAdmin_(account);
  const now = new Date();
  return {
    account: account,
    isAdmin: isAdmin,
    isOwner: ADMIN_ACCOUNTS.indexOf(account) >= 0,
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
    deadline: nextDeadline_(now),
    statuses: ORDER_STATUSES,
    orderUnits: st.units.order,
    recipeUnits: st.units.recipe,
    pendingUnits: st.units.pending,
    deptSubs: st.deptSubs,
    missingSetup: isAdmin ? st.missingSetup : [],
    vendorContacts: isAdmin ? st.vendorContacts : {},
    holidays: st.holidays,
    prices: st.prices || {},
    placeEnabled: placeEnabled_(),
  };
}

// ---- 教學影片：放在雲端硬碟「訂貨系統教學影片」資料夾；檔名開頭「管理者」只有管理者看得到 ----
const HELP_FOLDER = '訂貨系統教學影片';

/** 教學影片清單（登入頁也可以看，不用 token；管理者多看管理者教學和分享狀態）。refresh：管理者重新讀資料夾 */
function getHelpVideos(token, refresh) {
  let admin = false;
  if (token) { try { admin = isAdmin_(auth_(token)); } catch (e) { admin = false; } }
  const cache = CacheService.getScriptCache();
  let list = admin && refresh ? null : JSON.parse(cache.get('helpVideos') || 'null');
  if (!list) {
    list = [];
    const it = DriveApp.getFoldersByName(HELP_FOLDER);
    if (it.hasNext()) {
      const files = it.next().getFiles();
      while (files.hasNext()) {
        const f = files.next();
        if (!/^video\//.test(f.getMimeType())) continue;
        let shared = true;
        try { shared = f.getSharingAccess() !== DriveApp.Access.PRIVATE; } catch (e) { /* 共用雲端硬碟 */ }
        list.push({ id: f.getId(), name: f.getName().replace(/\.[^.]+$/, ''), shared: shared });
      }
    }
    list.sort(function (a, b) { return a.name < b.name ? -1 : a.name > b.name ? 1 : 0; });
    cache.put('helpVideos', JSON.stringify(list), 600);
  }
  return {
    admin: admin, folder: HELP_FOLDER,
    videos: list.filter(function (v) { return admin || !/^管理者/.test(v.name); })
      .map(function (v) { return admin ? v : { id: v.id, name: v.name }; }),
  };
}

/** 廠商聯絡資料：{ 廠商: { fax, phone, person, note } } */
function readVendorContacts_() {
  const sh = ss_().getSheetByName(SHEET.contacts);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  readTable_(SHEET.contacts, ['廠商']).rows.forEach(function (r) {
    const v = String(r['廠商']).trim();
    if (v) out[v] = { fax: String(r['傳真'] || '').trim(), phone: String(r['電話'] || '').trim(), person: String(r['聯絡人'] || '').trim(), note: String(r['備註'] || '').trim() };
  });
  return out;
}

/** 管理者儲存一家廠商的傳真、電話、聯絡人、備註，回傳全部廠商的聯絡資料 */
function adminSaveVendorContact(token, vendor, f) {
  auth_(token, true);
  const v = clean_(vendor, 40);
  if (!v) throw new Error('沒有廠商名稱。');
  f = f || {};
  const set = { '傳真': clean_(f.fax, 30), '電話': clean_(f.phone, 30), '聯絡人': clean_(f.person, 30), '備註': clean_(f.note, 100) };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SHEET.contacts, CONTACT_HEADERS);
    ensureHeaders_(sh, CONTACT_HEADERS);
    sh.getRange('B:C').setNumberFormat('@');   // 電話號碼開頭的 0 不要被吃掉
    const t = readTable_(SHEET.contacts, ['廠商']);
    const r = t.rows.find(function (x) { return String(x['廠商']).trim() === v; });
    if (r) Object.keys(set).forEach(function (k) { setCell_(t, r._row, k, set[k]); });
    else { set['廠商'] = v; appendRows_(SHEET.contacts, [set]); }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return readVendorContacts_();
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

/** 訂貨截止規則 { wd: 星期(1～5), h: 時, m: 分 }：系統設定存的，沒有就用預設 */
function lockRule_() {
  return memo_('lockRule', function () {
    try {
      const r = JSON.parse(PropertiesService.getScriptProperties().getProperty('LOCK_RULE') || 'null');
      if (r && r.wd >= 1 && r.wd <= 5 && r.h >= 0 && r.h <= 23 && r.m >= 0 && r.m <= 59) return { wd: r.wd, h: r.h, m: r.m };
    } catch (e) {}
    return { wd: LOCK_WEEKDAY, h: LOCK_HOUR, m: LOCK_MINUTE };
  });
}
function lockTime_() { const r = lockRule_(); return r.h + ':' + ('0' + r.m).slice(-2); }
// 截止那天整點產生訂貨單：12:00 截止就 12 點跑；12:30 截止就 13 點跑（不要在截止前就產生）
function vendorOrderHour_() { const r = lockRule_(); return r.m > 0 ? Math.min(23, r.h + 1) : r.h; }
// 訂貨提醒在截止前一個上課日（截止週一 → 週五，截止週三 → 週二）
function remindWeekday_() { const wd = lockRule_().wd; return wd === 1 ? 5 : wd - 1; }

function lockRuleText_() {
  return '每週' + WEEKDAY[lockRule_().wd] + ' ' + lockTime_() + ' 截止下週的訂單';
}

/** 管理者：設定訂貨截止時間（星期一～五、幾點幾分）；有啟用排程的話，訂貨單和提醒的排程跟著改 */
function adminSetLockRule(token, rule) {
  auth_(token, true);
  const wd = Number(rule && rule.wd), h = Number(rule && rule.h), m = Number(rule && rule.m);
  if (!(wd >= 1 && wd <= 5 && Math.floor(wd) === wd)) throw new Error('截止星期只能選星期一～星期五。');
  if (!(h >= 0 && h <= 23 && Math.floor(h) === h) || !(m >= 0 && m <= 59 && Math.floor(m) === m)) throw new Error('截止時間不正確。');
  const props = PropertiesService.getScriptProperties();
  if (wd === LOCK_WEEKDAY && h === LOCK_HOUR && m === LOCK_MINUTE) props.deleteProperty('LOCK_RULE');
  else props.setProperty('LOCK_RULE', JSON.stringify({ wd: wd, h: h, m: m }));
  memoClear_();
  const handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (handlers.indexOf('autoVendorOrders') >= 0 || handlers.indexOf('remindTeachers') >= 0) setTriggers_(true);
  return automationStatus_();
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
  const o = validateOrder_(payload, account, isAdmin_(account));

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
  const isAdmin = isAdmin_(account);
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
  return { orderNo: orderNo, useDate: dateStr_(o.useDate), count: o.lines.length - tools, tools: tools, pending: newOnes.length, learned: o.learned || [] };
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
  if (!CLASS_TIME_PATTERN.test(classTime)) throw new Error('請選擇上課時間（上午／下午／夜間）。');
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
  const learn = learnIndex_(catalog), learned = [];
  (Array.isArray(p.newItems) ? p.newItems : []).forEach(function (it) {
    const name = clean_(it && it.name, 40);
    const vendor = clean_(it && it.vendor, 20);
    if (!name) throw new Error('新增食材請填寫名稱。');
    // 自動學習：名稱和正式食材一樣，或是管理者以前對應過的叫法（只對到一種食材），直接換成正式食材，不用再審核。
    // 老師在確認視窗按過「不是，我要新的」（keepNew）就照老師的。
    const hit = it.keepNew ? null : learnMatch_(learn, name);
    if (hit) {
      const r = catalog[hit], realVendor = String(r['廠商']).trim(), realName = String(r['食材名稱']).trim();
      if (learnKey_(name) !== learnKey_(realName)) learned.push({ from: name, to: realName });
      lines.push({
        id: hit, name: realName, vendor: orderVendor_(realVendor, dept), catVendor: realVendor,
        qty: qty_(it.qty, realName), unit: unit_(it.unit, r['單位'], realName, newUnits), note: clean_(it.note, 100),
        group: deliveryGroup_(realVendor, dow, vendorGroups),
        noOrder: !!it.noOrder || realVendor === SELF_VENDOR, section: clean_(it.section, 20),
      });
      return;
    }
    if (!vendor) throw new Error('請選擇「' + name + '」的廠商。');
    lines.push({
      id: PENDING_ID, name: name, vendor: orderVendor_(vendor, dept), catVendor: vendor,
      qty: qty_(it.qty, name), unit: unit_(it.unit, '', name, newUnits), note: clean_(it.note, 100),
      group: deliveryGroup_(vendor, dow, vendorGroups),
      noOrder: !!it.noOrder || vendor === SELF_VENDOR, section: clean_(it.section, 20),
    });
  });

  // 科別替換：這一科老師選的（或新增的）食材換成另一項（例如烘焙課的牛奶 → 保久乳）
  const subs = readDeptSubs_()[dept];
  if (subs) {
    const byName = {};
    Object.keys(catalog).forEach(function (id) { byName[learnKey_(catalog[id]['食材名稱'])] = id; });
    lines.forEach(function (l, i) {
      if (l.id === TOOL_ID) return;
      const to = subs[learnKey_(l.name)], id = to && byName[learnKey_(to)];
      if (!id || id === l.id) return;
      const r = catalog[id], realVendor = String(r['廠商']).trim(), realName = String(r['食材名稱']).trim();
      const prev = learned.find(function (x) { return x.to === l.name; });   // 牛奶 → 鮮奶(全) → 保久乳：合成一筆
      if (prev) { prev.to = realName; prev.dept = dept; } else learned.push({ from: l.name, to: realName, dept: dept });
      lines[i] = Object.assign({}, l, { id: id, name: realName, vendor: orderVendor_(realVendor, dept), catVendor: realVendor,
        group: deliveryGroup_(realVendor, dow, vendorGroups), noOrder: l.noOrder || realVendor === SELF_VENDOR });
    });
  }
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
    className: className, project: project, onBehalfOf: onBehalfOf, groups: groups, lines: lines, learned: learned,
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
let VENDOR_GROUPS_CACHE_ = null;   // 同一次執行只讀一次
function readVendorGroups_() {
  if (VENDOR_GROUPS_CACHE_) return VENDOR_GROUPS_CACHE_;
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
    return map;   // 讀不到時不快取，下次再試
  }
  VENDOR_GROUPS_CACHE_ = map;
  return map;
}

/** 訂單上要寫的廠商：E乾貨、C蛋商依科別分成烘焙／中餐兩種。 */
function orderVendor_(vendor, dept) {
  const split = VENDOR_SPLIT[String(vendor || '').trim()];
  if (!split) return vendor;
  return BAKING_DEPTS.indexOf(String(dept || '').trim()) >= 0 ? split.baking : split.other;
}

/**
 * 依上課日期的星期決定送貨組別，照「廠商設定」分頁（目前菜商、肉商、海鮮分一二／三四五）。
 * 廠商設定沒有列的廠商不分組別（回傳空白），進貨單一週合在一起。
 */
function deliveryGroup_(vendor, dow, vendorGroups) {
  const ch = WEEKDAY[dow];
  const codes = vendorGroups[String(vendor || '').trim()] || [];
  for (let i = 0; i < codes.length; i++) {
    if (codes[i].indexOf(ch) >= 0) return codes[i];
  }
  return '';
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
  const list = groupOrders_(rows, account).sort(function (a, b) { return a.orderNo < b.orderNo ? 1 : -1; });
  // 管理者改過、退單的紀錄（含原因），給老師看
  const edits = orderEditsFor_(list.map(function (o) { return o.orderNo; }));
  list.forEach(function (o) {
    o.adminNotes = (edits[o.orderNo] || []).filter(function (e) { return isAdmin_(e.who); });
  });
  const hidden = hiddenOrders_();
  const shown = list.filter(function (o) { return !hidden[o.orderNo]; });
  const places = placesFor_(shown.map(function (o) { return o.orderNo; }));
  shown.forEach(function (o) { o.place = places[o.orderNo] || null; });
  return shown;
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
  const isAdmin = isAdmin_(account);
  const res = cancelOrderChecked_(account, isAdmin, orderNo);
  try { logOrderEdit_(String(orderNo || '').trim(), account, ['取消整張訂單']); } catch (e) { console.error(e.message); }
  if (isAdmin) notifyTeacherOfEdit_(String(orderNo || '').trim(), account, ['取消整張訂單']);
  return res;
}

// ---- 異常紀錄：標記異常、老師刪除、退單、修改、取消、抽查異常，集中在一張表 ----

/**
 * 管理者：from～to（事件發生的日期）的異常紀錄，最新的在前。countOnly：只回傳還沒處理的件數（不讀訂單，比較快）
 * 每筆：{ id, time, type, orderNo, useDate, className, project, teacher, who, note, done }
 */
function adminGetIssues(token, from, to, countOnly) {
  auth_(token, true);
  const start = parseDate_(from) || addDays_(today_(), -30), end = addDays_(parseDate_(to) || today_(), 1);
  const out = [];
  const tm = function (v) { const d = v instanceof Date ? v : errTime_(v); return d instanceof Date && !isNaN(d) ? d : null; };
  const inRange = function (d) { return d && d >= start && d < end; };
  const add = function (d, type, orderNo, who, note, extra) {
    if (!inRange(d)) return;
    out.push(Object.assign({ id: type + '|' + orderNo + '|' + Utilities.formatDate(d, tz_(), 'yyyyMMddHHmmss') + (extra && extra.item ? '|' + extra.item : ''),
      t: d.getTime(), time: fmtDateTime_(d), type: type, orderNo: orderNo, who: who, note: note }, extra || {}));
  };
  const has = function (name) { const sh = ss_().getSheetByName(name); return sh && sh.getLastRow() >= 2; };
  if (has(SHEET.flags)) readTable_(SHEET.flags, ['訂單編號']).rows.forEach(function (r) {
    const item = String(r['食材'] || '').trim();
    add(tm(r['標記時間']), '標記異常', String(r['訂單編號']).trim(), String(r['標記人'] || ''), (item ? item + '：' : '') + String(r['說明'] || ''),
      { item: item, resolved: String(r['狀態']).trim() === '已解決' });
  });
  if (has(SHEET.hidden)) readTable_(SHEET.hidden, ['訂單編號']).rows.forEach(function (r) {
    add(tm(r['刪除時間']), '老師刪除', String(r['訂單編號']).trim(), String(r['老師'] || ''), '刪除時狀態：' + String(r['原狀態'] || ''));
  });
  if (has(SHEET.edits)) readTable_(SHEET.edits, ['時間', '訂單編號']).rows.forEach(function (r) {
    const text = String(r['內容'] || ''), who = String(r['修改人'] || ''), no = String(r['訂單編號']).trim();
    if (/老師刪除訂單/.test(text)) return;   // 「老師刪除」分頁已經有
    const type = /(^|；)退單/.test(text) ? '退單' : /取消整張訂單|整張取消/.test(text) ? '取消' : '修改';
    add(tm(r['時間']), type, no, who, (type === '修改' || type === '取消' ? (isAdmin_(who) ? '管理者' : '老師') + '：' : '') + text);
  });
  if (has(SHEET.audits)) {
    const seen = {};
    readTable_(SHEET.audits, ['抽查編號']).rows.forEach(function (r) {
      const bad = ['少了', '多了', '沒有'].indexOf(String(r['結果']).trim()) >= 0, seal = String(r['封條']).trim() === '破了';
      if (!bad && !seal) return;
      const id = String(r['抽查編號']).trim();
      const d = tm(r['抽查時間']);
      const note = (seal ? '封條破了；' : '') + (bad ? String(r['食材名稱'] || '') + ' ' + String(r['結果']) + '（清單 ' + r['清單數量'] + '、實際 ' + r['實際數量'] + ' ' + String(r['單位'] || '') + '）' : '');
      if (seen[id]) { seen[id].note += '；' + note.replace(/^封條破了；/, ''); return; }
      add(d, '抽查異常', String(r['訂單編號']).trim(), String(r['抽查人'] || ''), note,
        { useDate: toDate_(r['上課日期']) ? dateStr_(toDate_(r['上課日期'])) : '', className: String(r['班級'] || ''), teacher: String(r['老師'] || ''), auditId: id });
      seen[id] = out[out.length - 1];
    });
  }
  const done = {};
  if (has(SHEET.issuesDone)) readTable_(SHEET.issuesDone, ['事件ID']).rows.forEach(function (r) { done[String(r['事件ID']).trim()] = String(r['處理人'] || '') + ' ' + fmtDateTime_(r['處理時間']); });
  out.forEach(function (x) { x.done = done[x.id] || (x.resolved ? '已解決' : ''); });
  if (countOnly) return out.filter(function (x) { return !x.done; }).length;
  // 補上訂單的上課日期、班級、實作項目、老師
  const want = {};
  out.forEach(function (x) { want[x.orderNo] = true; });
  const meta = {};
  groupOrders_(withHistory_(readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows, addDays_(start, -120)).filter(function (r) { return want[String(r['訂單編號']).trim()]; }), null)
    .forEach(function (o) { meta[o.orderNo] = o; });
  out.forEach(function (x) {
    const o = meta[x.orderNo];
    if (!o) return;
    x.useDate = x.useDate || o.useDate; x.className = x.className || o.className; x.project = o.project; x.teacher = x.teacher || o.onBehalfOf || o.teacher;
  });
  return out.sort(function (a, b) { return b.t - a.t; });
}

/** 管理者：異常紀錄勾「已處理」（on）或取消 */
function adminSetIssueDone(token, id, on) {
  const who = auth_(token, true);
  id = String(id || '').slice(0, 200);
  if (!id) throw new Error('缺少事件。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(ensureSheet_(SHEET.issuesDone, ISSUE_DONE_HEADERS), ISSUE_DONE_HEADERS);
    const t = readTable_(SHEET.issuesDone, ['事件ID']);
    const rows = t.rows.filter(function (r) { return String(r['事件ID']).trim() === id; });
    if (on && !rows.length) appendRows_(SHEET.issuesDone, [{ '事件ID': id, '處理人': who, '處理時間': new Date() }], { '處理時間': 'yyyy/mm/dd hh:mm' });
    if (!on && rows.length) deleteRows_(t.sheet, rows.map(function (r) { return r._row; }));
  } finally {
    lock.releaseLock();
  }
  return { done: on ? who + ' ' + fmtDateTime_(new Date()) : '' };
}

/** 老師刪掉的訂單：{ 訂單編號: { by, time } } */
function hiddenOrders_() {
  return memo_('hiddenOrders', function () {
    const out = {};
    const sh = ss_().getSheetByName(SHEET.hidden);
    if (!sh || sh.getLastRow() < 2) return out;
    readTable_(SHEET.hidden, ['訂單編號']).rows.forEach(function (r) {
      out[String(r['訂單編號']).trim()] = { by: String(r['老師'] || ''), time: fmtDateTime_(r['刪除時間']) };
    });
    return out;
  });
}

/**
 * 老師刪除自己的訂單（老師那邊看不到，管理者照樣看得到）。
 * 待處理、還沒截止 → 先取消再刪（規則跟取消一樣）；已取消 → 直接刪；已審核、已叫貨 → 不能刪，請找管理者。
 */
function deleteMyOrder(token, orderNo) {
  const account = auth_(token);
  orderNo = clean_(orderNo, 30);
  const rows = readTable_(SHEET.orders, ['訂單編號']).rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo; });
  if (!rows.length) throw new Error('找不到這張訂單（上課日期太久以前的訂單不能刪除）。');
  const r0 = rows[0];
  if (String(r0['操作老師']).trim() !== account && String(r0['代訂對象老師'] || '').trim() !== account) throw new Error('只能刪除自己的訂單。');
  if (hiddenOrders_()[orderNo]) return { ok: true };
  const live = rows.filter(function (x) { return String(x['狀態']).trim() !== '已取消'; });
  const was = live.length ? String(live[0]['狀態']).trim() : '已取消';
  if (live.length) {
    if (live.some(function (x) { return String(x['狀態']).trim() !== '待處理'; })) throw new Error('這張訂單已經在處理（已審核或已叫貨），不能自己刪除，請聯絡管理者。');
    cancelOrderChecked_(account, false, orderNo);   // 已過截止會擋下來
    try { logOrderEdit_(orderNo, account, ['老師刪除訂單']); } catch (e) { console.error(e.message); }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(ensureSheet_(SHEET.hidden, HIDDEN_HEADERS), HIDDEN_HEADERS);
    appendRows_(SHEET.hidden, [{ '訂單編號': orderNo, '老師': account, '刪除時間': new Date(), '原狀態': was }], { '刪除時間': 'yyyy/mm/dd hh:mm' });
  } finally {
    lock.releaseLock();
  }
  return { ok: true, cancelled: !!live.length };
}

/** 老師：把自己所有「已取消」的訂單一次從我的訂單移除（管理者照樣看得到） */
function deleteMyCancelled(token) {
  const account = auth_(token);
  const hidden = hiddenOrders_();
  const rows = readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows.filter(function (r) {
    return String(r['操作老師']).trim() === account || String(r['代訂對象老師'] || '').trim() === account;
  });
  const nos = groupOrders_(rows, account).filter(function (o) { return o.status === '已取消' && !hidden[o.orderNo]; }).map(function (o) { return o.orderNo; });
  if (!nos.length) return { count: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(ensureSheet_(SHEET.hidden, HIDDEN_HEADERS), HIDDEN_HEADERS);
    const now = new Date();
    appendRows_(SHEET.hidden, nos.map(function (no) { return { '訂單編號': no, '老師': account, '刪除時間': now, '原狀態': '已取消' }; }), { '刪除時間': 'yyyy/mm/dd hh:mm' });
  } finally {
    lock.releaseLock();
  }
  return { count: nos.length };
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
/** lite：廠商訂貨單（①②③④）用，不需要課表比對、修改紀錄、列印紀錄，少讀幾個分頁比較快 */
function adminGetOrders(token, from, to, lite) {
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
  if (lite) return list;
  const nos = list.map(function (o) { return o.orderNo; });
  const edits = orderEditsFor_(nos), printed = printedFor_(nos), flags = flagsFor_(nos), places = placesFor_(nos), hidden = hiddenOrders_();
  list.forEach(function (o) { o.edits = edits[o.orderNo] || []; o.printed = printed[o.orderNo] || null; o.flag = flags.orders[o.orderNo] || null; o.itemFlags = flags.items[o.orderNo] || {}; o.place = places[o.orderNo] || null; o.deleted = hidden[o.orderNo] || null; });
  try { matchLessons_(list, start, end); } catch (e) { console.error('matchLessons_ ' + e.message); }
  return list;
}

/**
 * 每張訂單對照課表：o.lesson = { match: [課], near: [同一週這位老師的其他課（含叫料狀態）], known: 課表上有沒有這位老師 }。
 * 老師的日期、時段填錯時，管理者一眼就看得出來是哪一堂課。
 */
function matchLessons_(list, start, end) {
  if (!list.length) return;
  const mon = addDays_(start, -((start.getDay() + 6) % 7)), fri = addDays_(end, (5 - end.getDay() + 7) % 7);
  const lessons = lessonsBetween_(mon, fri, null);
  const st = lessonStatusMaps_(mon, fri);
  const teachers = {};
  readSchedule_().forEach(function (x) { teachers[x.teacher] = true; });
  readMyLessonRows_().forEach(function (x) { teachers[x.teacher] = true; });
  lessons.forEach(function (x) { teachers[x.teacher] = true; });
  const weekOf = function (ds) { const d = parseDate_(ds); return dateStr_(addDays_(d, -((d.getDay() + 6) % 7))); };
  const brief = function (x) {
    const k = x.teacher + '|' + x.date + '|' + x.session;
    return { date: x.date, session: x.session, course: x.course, room: x.room, className: x.className, source: x.source,
      status: st.have[k] ? 'ordered' : st.manual[k] ? 'manual' : st.skip[k] ? 'skipped' : holidayOf_(x.date) ? 'holiday' : 'todo' };
  };
  list.forEach(function (o) {
    if (isReturn_(o.project)) { o.lesson = null; return; }   // 還料單不對課表
    if (o.status === '已取消') return;
    const who = o.onBehalfOf || o.teacher, ses = sessionOf_(o.classTime), wk = weekOf(o.useDate);
    const mine = lessons.filter(function (x) { return x.teacher === who; });
    const match = mine.filter(function (x) { return x.date === o.useDate && x.session === ses; });
    o.lesson = { teacher: who, session: ses, known: !!teachers[who], match: match.map(brief),
      near: match.length ? [] : mine.filter(function (x) { return weekOf(x.date) === wk; }).map(brief) };
  });
}

/** 這些訂單最後一次在訂單彙整列印的時間和列印人：{ 訂單編號: { time, who } } */
/** 還沒解決的異常標記：{ orders: { 訂單編號: { note, by, time } }, items: { 訂單編號: { 食材名稱: { note, by, time } } } } */
function flagsFor_(orderNos) {
  const out = { orders: {}, items: {} };
  const sh = ss_().getSheetByName(SHEET.flags);
  if (!sh || sh.getLastRow() < 2 || !orderNos.length) return out;
  const want = {};
  orderNos.forEach(function (n) { want[n] = true; });
  readTable_(SHEET.flags, ['訂單編號', '狀態']).rows.forEach(function (r) {
    const no = String(r['訂單編號']).trim();
    if (!want[no] || String(r['狀態']).trim() !== '異常') return;
    const f = { note: String(r['說明'] || ''), by: String(r['標記人'] || ''), time: fmtDateTime_(r['標記時間']) }, item = String(r['食材'] || '').trim();
    if (item) (out.items[no] = out.items[no] || {})[item] = f; else out.orders[no] = f;
  });
  return out;
}

/** 管理者：某一張訂單的某一樣食材勾「異常」（on）或取消，只有管理者看得到、不影響叫貨 */
function adminSetItemFlag(token, orderNo, item, on, note) {
  const who = auth_(token, true);
  orderNo = clean_(orderNo, 30); item = clean_(item, 40); note = clean_(note, 200);
  if (!orderNo || !item) throw new Error('缺少訂單編號或食材。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(ensureSheet_(SHEET.flags, FLAG_HEADERS), FLAG_HEADERS);
    const t = readTable_(SHEET.flags, ['訂單編號', '狀態', '食材']);
    const open = t.rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo && String(r['食材']).trim() === item && String(r['狀態']).trim() === '異常'; });
    if (!on) { open.forEach(function (r) { setCells_(t, r._row, { '狀態': '已解決', '解決時間': new Date(), '解決人': who }); }); return { flag: null }; }
    if (open.length) setCells_(t, open[0]._row, { '說明': note, '標記人': who, '標記時間': new Date() });
    else appendRows_(SHEET.flags, [{ '訂單編號': orderNo, '食材': item, '說明': note, '標記人': who, '標記時間': new Date(), '狀態': '異常' }], { '標記時間': 'yyyy/mm/dd hh:mm' });
    return { flag: { note: note, by: who, time: fmtDateTime_(new Date()) } };
  } finally {
    lock.releaseLock();
  }
}

/** 管理者：標記訂單異常（note）或解除（resolve = true）。不改訂單狀態，叫貨照常 */
function adminSetFlag(token, orderNo, note, resolve) {
  const who = auth_(token, true);
  orderNo = clean_(orderNo, 30);
  if (!orderNo) throw new Error('缺少訂單編號。');
  note = clean_(note, 200);
  if (!resolve && !note) throw new Error('請寫異常說明，例如：數量看起來不對、要問老師。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.flags, FLAG_HEADERS);
    ensureHeaders_(sheet_(SHEET.flags), FLAG_HEADERS);
    const t = readTable_(SHEET.flags, ['訂單編號', '狀態']);
    const open = t.rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo && !String(r['食材'] || '').trim() && String(r['狀態']).trim() === '異常'; });
    if (resolve) {
      open.forEach(function (r) { setCells_(t, r._row, { '狀態': '已解決', '解決時間': new Date(), '解決人': who }); });
      if (note) open.forEach(function (r) { setCell_(t, r._row, '說明', String(r['說明'] || '') + '（解決：' + note + '）'); });
      return { flag: null };
    }
    if (open.length) {
      setCells_(t, open[0]._row, { '說明': note, '標記人': who, '標記時間': new Date() });
    } else {
      appendRows_(SHEET.flags, [{ '訂單編號': orderNo, '說明': note, '標記人': who, '標記時間': new Date(), '狀態': '異常' }], { '標記時間': 'yyyy/mm/dd hh:mm' });
    }
    return { flag: { note: note, by: who, time: fmtDateTime_(new Date()) } };
  } finally {
    lock.releaseLock();
  }
}

function printedFor_(orderNos) {
  const out = {};
  const sh = ss_().getSheetByName(SHEET.prints);
  if (!sh || sh.getLastRow() < 2 || !orderNos.length) return out;
  const want = {};
  orderNos.forEach(function (n) { want[n] = true; });
  readTable_(SHEET.prints, ['訂單編號', '列印時間']).rows.forEach(function (r) {
    const no = String(r['訂單編號']).trim();
    if (want[no]) out[no] = { time: fmtDateTime_(r['列印時間']), who: String(r['列印人'] || '') };
  });
  return out;
}

/** 管理者在訂單彙整按了列印：記下這幾張訂單的列印時間（同一張只留最後一次） */
function adminMarkPrinted(token, orderNos) {
  const admin = auth_(token, true);
  const nos = (Array.isArray(orderNos) ? orderNos : []).map(function (n) { return String(n || '').trim(); }).filter(Boolean);
  if (!nos.length) return {};
  if (nos.length > 500) throw new Error('一次最多記錄 500 張。');
  const now = new Date();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.prints, PRINT_HEADERS);
    const t = readTable_(SHEET.prints, ['訂單編號']);
    const at = {};
    t.rows.forEach(function (r) { at[String(r['訂單編號']).trim()] = r._row; });
    const add = [];
    nos.forEach(function (no) {
      if (at[no]) { setCell_(t, at[no], '列印時間', now); setCell_(t, at[no], '列印人', admin); }
      else { add.push({ '訂單編號': no, '列印時間': now, '列印人': admin }); at[no] = -1; }
    });
    if (add.length) appendRows_(SHEET.prints, add, { '列印時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const out = {};
  nos.forEach(function (no) { out[no] = { time: fmtDateTime_(now), who: admin }; });
  return out;
}

/** 管理者改上課資訊：{ className, room, project, groups, classTime, useDate }，只改有給的欄位。沒有要改的回傳 null。 */
function orderInfoSet_(info) {
  if (!info || typeof info !== 'object') return null;
  const set = {};
  if (info.classTime !== undefined) {
    const v = clean_(info.classTime, 12);
    if (!CLASS_TIME_PATTERN.test(v)) throw new Error('請選擇上課時間（上午／下午／夜間）。');
    set['上課時間'] = v;
  }
  if (info.useDate !== undefined) {
    const d = parseDate_(info.useDate);
    if (!d) throw new Error('上課日期不正確。');
    if (d.getDay() === 0 || d.getDay() === 6) throw new Error('上課日期是週末，廠商不送貨，請改選平日。');
    set['使用日期'] = d;
    set['星期'] = '星期' + WEEKDAY[d.getDay()];
  }
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
  const removed = [], added = [];
  Object.keys(A).forEach(function (k) {
    const x = A[k], y = B[k];
    if (!y) { removed.push(x); return; }
    if (x.qty !== y.qty) lines.push(x.label + q(x, x.qty) + ' →' + q(y, y.qty));
    if (!x.noOrder !== !y.noOrder) lines.push(x.label + (y.noOrder ? ' 改成有貨' : ' 改成要叫貨'));
  });
  Object.keys(B).forEach(function (k) { if (!A[k]) added.push(B[k]); });
  // 同一樣食材只換了單位：寫成「洋蔥 3斤 → 3公斤」
  removed.forEach(function (x) {
    const i = added.findIndex(function (y) { return !x.tool && !y.tool && y.label === x.label; });
    if (i >= 0) { lines.push(x.label + q(x, x.qty) + ' → ' + q(added[i], added[i].qty).trim()); added.splice(i, 1); }
    else lines.push('刪除 ' + x.label + q(x, x.qty));
  });
  added.forEach(function (y) { lines.push('新增 ' + y.label + q(y, y.qty)); });
  return lines;
}

function logOrderEdit_(orderNo, who, changes, reason) {
  if (!changes.length) return;   // 按了儲存但沒改東西，不記
  reason = clean_(reason, 100);
  ensureSheet_(SHEET.edits, EDIT_HEADERS);
  appendRows_(SHEET.edits, [{ '時間': new Date(), '訂單編號': orderNo, '修改人': who, '內容': (reason ? '原因：' + reason + '；' : '') + changes.join('；') }],
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

// ---- 確定訂貨（實際叫貨量） ----

function purchaseRange_(from, to) {
  let start = parseDate_(from), end = parseDate_(to) || start;
  if (!start) throw new Error('請選擇日期。');
  if (end < start) { const t = start; start = end; end = t; }
  if ((end - start) / 86400000 > 62) throw new Error('一次最多確定兩個月內的叫貨。');
  return { from: dateStr_(start), to: dateStr_(end) };
}

function purchaseDate_(v) {
  const d = toDate_(v);
  return d ? dateStr_(d) : String(v || '').trim();
}

/** 讀這段期間的叫貨（草稿和已確定）：[{ date, group, vendor, orderNo, className, teacher, name, id, need, needUnit, qty, unit, kind, note, status, time, by }] */
function adminGetPurchase(token, from, to) {
  auth_(token, true);
  const r = purchaseRange_(from, to);
  const sh = ss_().getSheetByName(SHEET.purchase);
  if (!sh || sh.getLastRow() < 2) return [];
  return readTable_(SHEET.purchase, ['送貨日期', '廠商', '食材名稱', '叫貨量']).rows.filter(function (x) {
    const d = purchaseDate_(x['送貨日期']);
    return d >= r.from && d <= r.to;
  }).map(function (x) {
    return {
      date: purchaseDate_(x['送貨日期']), group: String(x['送貨組別'] || ''), vendor: String(x['廠商'] || ''),
      orderNo: String(x['訂單編號'] || ''), className: String(x['上課班級'] || ''), teacher: String(x['老師'] || ''),
      name: String(x['食材名稱'] || ''), id: String(x['食材ID'] || ''),
      need: x['需求量'] === '' ? '' : Number(x['需求量']), needUnit: String(x['需求單位'] || ''),
      qty: Number(x['叫貨量']) || 0, unit: String(x['叫貨單位'] || ''), kind: String(x['類型'] || '訂單'),
      note: String(x['備註'] || ''), status: String(x['狀態'] || '').trim() === '草稿' ? '草稿' : '已確定',
      time: fmtDateTime_(x['確定時間']), by: String(x['確認人'] || ''),
    };
  });
}

/**
 * 儲存一家廠商的實際叫貨量：這段期間這家廠商的叫貨紀錄整批換成 rows（可以重複修改）。
 * confirm = false：先存草稿，不改訂單狀態、不通知。
 * confirm = true：正式確定；訂單裡所有要叫貨的廠商都確定了，訂單（待處理、已審核）才改成「已叫貨」；
 *                 並用 LINE／Email 傳這家廠商的叫貨摘要給管理者。
 * rows: [{ date, group, vendor, orderNo, className, teacher, name, id, need, needUnit, qty, unit, kind: '訂單'|'補庫存'|'合計', note }]
 * kind '合計'：在 ③ 叫貨單直接改的合計（這個送貨組別這個食材給廠商的總量，0 ＝ 不叫），各班明細不變。
 * quiet：只是在 ③ 改合計，不再傳 LINE／Email。
 */
function adminSavePurchase(token, from, to, vendor, rows, confirm, quiet) {
  const admin = auth_(token, true);
  const r = purchaseRange_(from, to);
  vendor = clean_(vendor, 30);
  if (!vendor) throw new Error('缺少廠商。');
  if (!Array.isArray(rows)) throw new Error('叫貨資料不正確。');
  if (rows.length > PURCHASE_MAX_ROWS) throw new Error('品項太多，請縮短日期範圍。');
  const status = confirm ? '已確定' : '草稿';
  const now = new Date();
  const out = rows.map(function (x, i) {
    const date = String(x && x.date || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date < r.from || date > r.to) throw new Error('第 ' + (i + 1) + ' 筆的送貨日期不在這段期間。');
    const name = clean_(x.name, 40);
    if (!name) throw new Error('第 ' + (i + 1) + ' 筆沒有食材名稱。');
    if (clean_(x.vendor, 30) !== vendor) throw new Error('「' + name + '」不是' + vendor + '的品項。');
    const qty = Number(x.qty);
    if (!isFinite(qty) || qty < 0 || qty > 99999) throw new Error('「' + name + '」的叫貨量不正確。');
    const unit = clean_(x.unit, 10);
    if (qty > 0 && !unit) throw new Error('「' + name + '」沒有填叫貨單位。');
    const need = x.need === '' || x.need === null || x.need === undefined ? '' : Number(x.need);
    const add = x.kind === '補庫存' || x.kind === '合計';
    return {
      '送貨日期': date, '送貨組別': clean_(x.group, 20), '廠商': vendor,
      '訂單編號': add ? '' : clean_(x.orderNo, 30), '上課班級': add ? '' : clean_(x.className, 30), '老師': add ? '' : clean_(x.teacher, 30),
      '食材名稱': name, '食材ID': clean_(x.id, 20),
      '需求量': isFinite(need) ? need : '', '需求單位': clean_(x.needUnit, 10),
      '叫貨量': Math.round(qty * 100) / 100, '叫貨單位': unit, '類型': add ? x.kind : '訂單',
      '備註': clean_(x.note, 60), '狀態': status, '確定時間': now, '確認人': admin,
    };
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let changedOrders = 0;
  try {
    // 1. 叫貨紀錄：這段期間這家廠商的舊紀錄刪掉，換成這次的
    const sh = ensureSheet_(SHEET.purchase, PURCHASE_HEADERS);
    ensureHeaders_(sh, PURCHASE_HEADERS);
    let confirmed = {};   // 這段期間已確定的廠商
    if (sh.getLastRow() >= 2) {
      const all = readTable_(SHEET.purchase, ['送貨日期', '廠商']).rows.filter(function (x) {
        const d = purchaseDate_(x['送貨日期']);
        return d >= r.from && d <= r.to;
      });
      all.forEach(function (x) {
        const v = String(x['廠商']).trim();
        if (v !== vendor && String(x['狀態'] || '').trim() !== '草稿') confirmed[v] = true;
      });
      const old = all.filter(function (x) { return String(x['廠商']).trim() === vendor; }).map(function (x) { return x._row; });
      if (old.length) deleteRows_(sh, old);
    }
    appendRows_(SHEET.purchase, out, { '確定時間': 'yyyy/MM/dd HH:mm', '需求量': '0.##', '叫貨量': '0.##' });
    if (confirm) confirmed[vendor] = true;

    // 2. 確定時：訂單裡要叫貨的廠商全部確定了，待處理、已審核的訂單才改成「已叫貨」（個別取消的品項保持取消）
    if (confirm) {
      const t = readTable_(SHEET.orders, ['訂單編號', '使用日期', '狀態']);
      const start = parseDate_(r.from), end = parseDate_(r.to);
      const byNo = {};
      t.rows.forEach(function (x) {
        const d = toDate_(x['使用日期']);
        if (!d || d < start || d > end) return;
        (byNo[String(x['訂單編號']).trim()] = byNo[String(x['訂單編號']).trim()] || []).push(x);
      });
      const col = t.cols['狀態'];
      const lastRow = t.sheet.getLastRow();
      const statusVals = lastRow >= 2 ? t.sheet.getRange(2, col, lastRow - 1, 1).getValues() : [];
      Object.keys(byNo).forEach(function (no) {
        const live = byNo[no].filter(function (x) { return String(x['狀態']).trim() !== '已取消'; });
        if (!live.some(function (x) { return ['待處理', '已審核'].indexOf(String(x['狀態']).trim()) >= 0; })) return;
        const need = {};
        live.forEach(function (x) {
          if (String(x['食材ID']).trim() === TOOL_ID || isTrue_(x['不叫貨'])) return;
          need[String(x['廠商']).trim()] = true;
        });
        const vs = Object.keys(need);
        if (!vs.length || vs.indexOf(vendor) < 0 || !vs.every(function (v) { return confirmed[v]; })) return;
        live.forEach(function (x) { statusVals[x._row - 2][0] = '已叫貨'; });
        changedOrders++;
      });
      if (changedOrders) t.sheet.getRange(2, col, statusVals.length, 1).setValues(statusVals);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  if (!confirm || quiet) return { saved: out.length, orders: changedOrders, notify: null, status: status };

  // 3. 叫貨摘要傳給管理者
  const range = shortDateTw_(parseDate_(r.from)) + (r.from !== r.to ? '～' + shortDateTw_(parseDate_(r.to)) : '');
  const c = { n: 0, diff: 0, add: 0, zero: 0 };
  out.forEach(function (x) {
    if (x['類型'] === '合計') return;
    if (x['叫貨量'] > 0) c.n++; else c.zero++;
    if (x['類型'] === '補庫存') c.add++;
    else if (x['叫貨量'] > 0 && (x['叫貨量'] !== x['需求量'] || x['叫貨單位'] !== x['需求單位'])) c.diff++;
  });
  const line = vendor + '：叫 ' + c.n + ' 筆' + [c.diff ? '跟老師訂的不同 ' + c.diff + ' 筆' : '', c.add ? '補庫存 ' + c.add + ' 筆' : '', c.zero ? '不叫（用庫存）' + c.zero + ' 筆' : '']
    .filter(Boolean).map(function (x) { return '，' + x; }).join('');
  let notify = null;
  try {
    notify = notify_('【叫貨確定】' + vendor + ' ' + range,
      '<p>' + escHtml_(admin) + ' 已確定 ' + escHtml_(range) + ' 的叫貨。</p><p>' + escHtml_(line) + '</p>' +
        (changedOrders ? '<p>' + changedOrders + ' 張訂單的廠商都確定了，改成「已叫貨」。</p>' : '') +
        '<p>明細在試算表「' + SHEET.purchase + '」分頁。</p>',
      null /* 只寄 Email，不傳 LINE（省則數）*/);
  } catch (e) { console.error(e.message); }
  return { saved: out.length, orders: changedOrders, notify: notify, status: status };
}

/** 管理者退單給老師：整張取消，LINE 告訴老師退單原因，請老師用「再訂一次」修改後重新送出 */
function adminReturnOrder(token, orderNo, reason, notify) {
  const admin = auth_(token, true);
  orderNo = String(orderNo || '').trim();
  reason = clean_(reason, 100);
  if (!reason) throw new Error('請填寫退單原因，老師才知道要改什麼。');
  updateOrderStatus_(orderNo, '已取消', function (rows) {
    if (rows.every(function (r) { return String(r['狀態']).trim() === '已取消'; })) throw new Error('這張訂單已經取消了。');
  }, null);
  try { logOrderEdit_(orderNo, admin, ['退單'], reason); } catch (e) { console.error(e.message); }
  return { ok: true, notify: notify === false ? '沒有通知老師。' : notifyTeacherOfEdit_(orderNo, admin, ['退單'], reason, 'return') };
}

/**
 * 管理者改到老師的訂單（修改明細、取消）時，LINE 私訊那位老師（代訂的話傳給被代訂的老師）。
 * 老師沒綁定 LINE、管理者改自己的單、本月則數快用完時不傳。回傳結果文字（給畫面顯示）。
 */
function notifyTeacherOfEdit_(orderNo, admin, changes, reason, kind) {
  try {
    if ((!changes || !changes.length) && !reason) return '';
    changes = changes || [];
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
    const ret = kind === 'return';
    const text = (ret ? '【退單通知】' + teacher + ' 老師您好，管理者退回了您的訂單：\n' : '【訂單修改通知】' + teacher + ' 老師您好，管理者修改了您的訂單：\n') +
      (d ? shortDateTw_(d) + ' ' : '') + o.className + '「' + o.project + '」\n' +
      (reason ? (ret ? '退單原因：' : '修改原因：') + reason + '\n' : '') +
      (ret ? '' : changes.map(function (c) { return '・' + c; }).join('\n') + '\n') +
      (ret ? '\n請到「我的訂單」找到這張，按「再訂一次」修改後重新送出（已過截止的話請找管理者）。' : '\n有問題請找管理者。') +
      (lineUrl_('mine') ? '\n' + lineUrl_('mine') : '');
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
function adminUpdateOrder(token, orderNo, changes, info, reason, notify) {
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
        // 老師打的新叫法（待審核）被改成正式食材：自動記成別名，下次老師這樣打就直接對到
        if (String(r['食材ID']).trim() === PENDING_ID && name) { try { addAlias_(String(c.itemId).trim(), name); } catch (e) { console.error('addAlias_ ' + e.message); } }
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
      const vg = readVendorGroups_();
      t2.rows.forEach(function (r) {
        if (String(r['訂單編號']).trim() !== orderNo) return;
        Object.keys(infoSet).forEach(function (k) { if (t2.cols[k] !== undefined) setCell_(t2, r._row, k, infoSet[k]); });
        // 改了上課日期：送貨組別照新的星期重算（菜商、肉商、海鮮分一二／三四五）
        if (infoSet['使用日期'] && t2.cols['送貨組別'] !== undefined && String(r['食材ID']).trim() !== TOOL_ID) {
          setCell_(t2, r._row, '送貨組別', deliveryGroup_(groupVendor_(String(r['廠商'] || '').trim()), infoSet['使用日期'].getDay(), vg));
        }
      });
      if (infoSet['使用日期'] && t2.cols['使用日期'] !== undefined) {
        t2.rows.forEach(function (r) { if (String(r['訂單編號']).trim() === orderNo) t2.sheet.getRange(r._row, t2.cols['使用日期']).setNumberFormat('yyyy-mm-dd'); });
      }
    }
    SpreadsheetApp.flush();
    try { changesDone = diffOrders_(before, orderOf_(orderNo)); logOrderEdit_(orderNo, admin, changesDone, reason); } catch (e) { console.error(e.message); }
  } finally {
    lock.releaseLock();
  }
  // 傳 LINE 放在鎖外面，不讓別人等；管理者沒勾「LINE 通知老師」就不傳（notify === false）
  return { ok: true, notify: changesDone.length && notify !== false ? notifyTeacherOfEdit_(orderNo, admin, changesDone, clean_(reason, 100)) : changesDone.length ? '沒有通知老師。' : '' };
}

/**
 * 待審核食材，依「食材名稱」合併：同一個新食材在好幾張訂單裡，只要審核一次。
 * 也列出訂單裡還是「待審核」、但審核紀錄已經處理過（之前退回或重複送審）的食材（orphan = true），
 * 讓管理者補對應，不然廠商訂貨單的合計會分成兩項。
 */
function adminGetPending(token) {
  auth_(token, true);
  const groups = {}, list = [];
  const g = function (name) {
    if (!groups[name]) { groups[name] = { name: name, rows: [], teachers: [], orderNos: [], depts: [], lines: 0, time: '', vendor: '', unit: '' }; list.push(groups[name]); }
    return groups[name];
  };
  const add = function (arr, v) { if (v && arr.indexOf(v) < 0) arr.push(v); };
  readTable_(SHEET.pending, ['食材名稱', '審核狀態']).rows.forEach(function (r) {
    const st = String(r['審核狀態'] || '').trim();
    if (st && st !== '待審核') return;
    const name = String(r['食材名稱'] || '').trim();
    if (!name) return;
    const x = g(name);
    x.rows.push(r._row);
    add(x.teachers, String(r['提交老師'] || '').trim());
    add(x.orderNos, String(r['對應訂單編號'] || '').trim());
    if (!x.time) { x.time = fmtDateTime_(r['新增時間']); x.vendor = String(r['廠商'] || ''); x.unit = String(r['單位'] || ''); }
  });
  readTable_(SHEET.orders, ['訂單編號', '食材ID', '食材名稱']).rows.forEach(function (r) {
    if (String(r['食材ID']).trim() !== PENDING_ID || String(r['狀態'] || '').trim() === '已取消') return;
    const name = String(r['食材名稱'] || '').trim();
    if (!name) return;
    const x = g(name);
    x.lines++;
    add(x.orderNos, String(r['訂單編號']).trim());
    add(x.depts, String(r['科別'] || '').trim());
    add(x.teachers, String(r['代訂對象老師'] || r['操作老師'] || '').trim());
    if (!x.vendor) { x.vendor = String(r['廠商'] || ''); x.unit = String(r['單位'] || ''); }
  });
  return list.map(function (x) {
    return {
      row: x.rows[0] || 0, name: x.name, time: x.time, teacher: x.teachers.join('、'), vendor: x.vendor, unit: x.unit,
      orderNo: x.orderNos.join('、'), orders: x.orderNos.length, lines: x.lines, records: x.rows.length, orphan: !x.rows.length, depts: x.depts,
    };
  }).sort(function (a, b) { return a.orphan === b.orphan ? 0 : a.orphan ? 1 : -1; });
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

/**
 * 審核（依食材名稱）：row = 待審核食材的列號；沒有審核紀錄、只在訂單裡的（orphan）用 name。
 * 核准／對應會更新「所有」訂單裡同名的待審核品項，同名的其他審核紀錄也一起標記，不用一張一張審。
 * 核准時食材總表已經有這個名稱，就直接對應過去（不再報錯）。
 */
function adminReviewItem(token, row, decision, edits, name) {
  auth_(token, true);
  edits = edits || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const pend = readTable_(SHEET.pending, ['食材名稱', '審核狀態']);
    const isOpen = function (r) { const st = String(r['審核狀態'] || '').trim(); return !st || st === '待審核'; };
    let origName, first = null;
    if (Number(row)) {
      first = pend.rows.find(function (r) { return r._row === Number(row); });
      if (!first) throw new Error('找不到這筆待審核資料，請重新整理。');
      if (!isOpen(first)) throw new Error('這筆資料已經審核過了，請重新整理。');
      origName = String(first['食材名稱']).trim();
    } else {
      origName = String(name || '').trim();
      if (!origName) throw new Error('缺少食材名稱。');
    }
    const recs = pend.rows.filter(function (r) { return isOpen(r) && String(r['食材名稱']).trim() === origName; });
    let result, lines = 0;

    if (decision === 'reject') {
      if (!recs.length) throw new Error('這個食材已經不在審核清單，請改用「對應既有食材」或到訂單彙整修改。');
      result = '已退回';
    } else if (decision === 'approve') {
      const newName = clean_(edits.name, 40) || origName;
      const vendor = clean_(edits.vendor, 20) || String(first && first['廠商'] || '').trim();
      const unit = clean_(edits.unit, 10) || String(first && first['單位'] || '').trim();
      if (!vendor) throw new Error('請選擇廠商。');
      const exist = readTable_(SHEET.items, ['食材ID', '食材名稱']).rows.find(function (x) { return String(x['食材名稱']).trim() === newName; });
      if (exist) {
        const id = String(exist['食材ID']).trim();
        lines = fixPendingLines_(origName, { id: id, name: newName, vendor: String(exist['廠商'] || '').trim() });
        result = '已對應（' + id + '，食材總表本來就有）' + (addAlias_(id, origName) ? '，「' + origName + '」已加入別名' : '');
      } else {
        const id = createItem_({ name: newName, alias: newName === origName ? '' : origName, vendor: vendor, unit: unit });
        lines = fixPendingLines_(origName, { id: id, name: newName, vendor: vendor });
        result = '已核准（' + id + '）';
      }
    } else if (decision === 'map') {
      const id = String(edits.existingId || '').trim();
      const r = readTable_(SHEET.items, ['食材ID', '食材名稱']).rows.find(function (x) { return String(x['食材ID']).trim() === id; });
      if (!r) throw new Error('找不到食材 ' + id + '。');
      lines = fixPendingLines_(origName, { id: id, name: String(r['食材名稱']).trim(), vendor: String(r['廠商'] || '').trim() });
      result = '已對應（' + id + '）' + (addAlias_(id, origName) ? '，「' + origName + '」已加入別名' : '');
    } else {
      throw new Error('審核動作不正確。');
    }

    recs.forEach(function (r) { setCell_(pend, r._row, '審核狀態', result); });
    return result + (lines ? '，已更新 ' + lines + ' 筆訂單品項' : '');
  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理者：乾貨的進貨科別。食材總表的廠商寫「E乾貨」＝兩科都進（照訂單科別分開叫）；
 * 寫「E烘焙乾貨」＝只有烘焙在進（中餐課訂的也跟烘焙叫、到烘焙教室拿）；「E中餐乾貨」反過來。
 * 也回傳過去（訂單總表＋歷史訂單）烘焙課、其他課各訂過幾次，給管理者參考。
 */
function adminGetDryGoods(token, base) {
  auth_(token, true);
  base = base || 'E乾貨';
  const sp = VENDOR_SPLIT[base];
  if (!sp) throw new Error('「' + base + '」沒有分科。');
  const hist = {};
  [SHEET.orders, SHEET.history].forEach(function (name) {
    if (!ss_().getSheetByName(name)) return;
    readTable_(name, ['食材ID', '科別']).rows.forEach(function (r) {
      const id = String(r['食材ID'] || '').trim();
      if (!id || String(r['狀態'] || '').trim() === '已取消') return;
      const h = hist[id] = hist[id] || { baking: 0, other: 0 };
      if (BAKING_DEPTS.indexOf(String(r['科別'] || '').trim()) >= 0) h.baking++; else h.other++;
    });
  });
  const list = [];
  readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']).rows.forEach(function (r) {
    if (r['啟用'] !== undefined && !isTrue_(r['啟用'])) return;
    const v = String(r['廠商'] || '').trim();
    const mode = v === base ? 'both' : v === sp.baking ? 'baking' : v === sp.other ? 'other' : '';
    if (!mode) return;
    const id = String(r['食材ID']).trim();
    list.push({ id: id, name: String(r['食材名稱']).trim(), mode: mode, hist: hist[id] || { baking: 0, other: 0 } });
  });
  return { base: base, baking: sp.baking, other: sp.other, list: list.sort(function (a, b) { return a.name.localeCompare(b.name, 'zh-Hant'); }) };
}

/** 管理者：改乾貨的進貨科別。changes: [{ id, mode: 'both'|'baking'|'other' }] */
function adminSetDryGoods(token, base, changes) {
  auth_(token, true);
  base = base || 'E乾貨';
  const sp = VENDOR_SPLIT[base];
  if (!sp) throw new Error('「' + base + '」沒有分科。');
  const to = { both: base, baking: sp.baking, other: sp.other };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']);
    const byId = {};
    t.rows.forEach(function (r) { byId[String(r['食材ID']).trim()] = r; });
    let n = 0;
    (changes || []).forEach(function (c) {
      const r = byId[String(c && c.id || '').trim()], v = to[c && c.mode];
      if (!r || !v) return;
      const cur = String(r['廠商'] || '').trim();
      if ([base, sp.baking, sp.other].indexOf(cur) < 0 || cur === v) return;
      setCell_(t, r._row, '廠商', v);
      n++;
    });
    // 還沒叫貨的訂單（今天以後的課、狀態待處理或已審核）跟著換廠商，不然這週已經送出的還是跟原本那科叫
    let lines = 0;
    const moved = {};
    (changes || []).forEach(function (c) { if (to[c && c.mode]) moved[String(c.id).trim()] = to[c.mode]; });
    if (n) {
      clearItemsCache_();
      const o = readTable_(SHEET.orders, ['訂單編號', '食材ID', '廠商']);
      const today = today_();
      o.rows.forEach(function (r) {
        const v = moved[String(r['食材ID'] || '').trim()];
        const d = toDate_(r['使用日期']);
        const st = String(r['狀態'] || '').trim();
        if (!v || !d || d < today || (st !== '待處理' && st !== '已審核')) return;
        const nv = orderVendor_(v, r['科別']);
        if (String(r['廠商'] || '').trim() === nv) return;
        setCell_(o, r._row, '廠商', nv);
        lines++;
      });
    }
    return { changed: n, lines: lines, data: adminGetDryGoods(token, base) };
  } finally {
    lock.releaseLock();
  }
}

/** 訂單上可以用的廠商（食材總表的廠商，分科的換成兩家：E乾貨 → E烘焙乾貨、E中餐乾貨），不含自備 */
function orderVendors_() {
  const out = [];
  cachedItems_().forEach(function (it) {
    const v = String(it.vendor || '').trim();
    if (!v || v === SELF_VENDOR) return;
    (VENDOR_SPLIT[v] ? [VENDOR_SPLIT[v].baking, VENDOR_SPLIT[v].other] : [v]).forEach(function (x) { if (out.indexOf(x) < 0) out.push(x); });
  });
  return out;
}

/** 送貨組別用的廠商名稱（E烘焙乾貨、E中餐乾貨 照 E乾貨 的設定） */
function groupVendor_(v) {
  const base = Object.keys(VENDOR_SPLIT).find(function (k) { return VENDOR_SPLIT[k].baking === v || VENDOR_SPLIT[k].other === v; });
  return base || v;
}

/**
 * 管理者（② 訂貨單）：臨時轉單。某張（orderNo 也可以是陣列：好幾張）訂單的某個食材，原本的廠商沒有賣，改跟別家叫
 * （例如中餐課的乳酪絲改跟烘焙乾貨叫、送到烘焙教室；火腿肉商沒有，改跟菜商或烘焙乾貨叫）。送貨組別跟著新廠商重算。
 * remember：以後這個食材都跟這家叫（食材總表的廠商改掉，還沒叫貨的訂單也跟著改）。
 */
function adminMoveLineVendor(token, orderNo, name, fromVendor, toVendor, remember) {
  auth_(token, true);
  const nos = {};
  (Array.isArray(orderNo) ? orderNo : [orderNo]).forEach(function (n) { n = String(n || '').trim(); if (n) nos[n] = true; });
  name = String(name || '').trim();
  fromVendor = String(fromVendor || '').trim(); toVendor = String(toVendor || '').trim();
  if (!toVendor || toVendor === fromVendor) throw new Error('請選另一家廠商。');
  if (orderVendors_().indexOf(toVendor) < 0) throw new Error('沒有「' + toVendor + '」這家廠商。');
  const vendorGroups = readVendorGroups_();
  let lines = 0, id = '';
  const moved = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '食材名稱', '廠商']);
    t.rows.forEach(function (r) {
      if (!nos[String(r['訂單編號']).trim()] || String(r['食材名稱']).trim() !== name || String(r['廠商']).trim() !== fromVendor) return;
      if (String(r['狀態'] || '').trim() === '已取消') return;
      setCell_(t, r._row, '廠商', toVendor);
      const d = toDate_(r['使用日期']);
      const g = d ? deliveryGroup_(groupVendor_(toVendor), d.getDay(), vendorGroups) : '';
      if (t.cols['送貨組別'] !== undefined) setCell_(t, r._row, '送貨組別', g);
      moved.push({ orderNo: String(r['訂單編號']).trim(), group: g });
      id = id || String(r['食材ID'] || '').trim();
      lines++;
    });
  } finally {
    lock.releaseLock();
  }
  if (!lines) throw new Error('找不到這筆訂單品項（可能已經被修改或取消），請重新產生。');
  const keep = !!(remember && id && id !== PENDING_ID);
  const more = keep ? setItemVendor_(id, toVendor) : [];
  return { lines: lines, id: id, moved: moved, remembered: keep, more: more.length, moreMoved: more };
}

/**
 * 管理者：一次改很多食材的廠商（乾貨進貨科別「貼上清單」）。changes: [{ id, vendor }]，
 * vendor 可以是訂單上的廠商（E烘焙乾貨、E中餐乾貨、F飲調…）或分科前的名稱（E乾貨＝兩科都進）。
 * 食材總表、訂單總表各讀一次；還沒叫貨的訂單（今天以後、待處理或已審核）跟著改。
 */
function adminSetItemVendors(token, changes) {
  auth_(token, true);
  const ok = orderVendors_().concat(Object.keys(VENDOR_SPLIT));
  const want = {};
  (changes || []).forEach(function (c) {
    const id = String(c && c.id || '').trim(), v = String(c && c.vendor || '').trim();
    if (!id) return;
    if (ok.indexOf(v) < 0) throw new Error('沒有「' + v + '」這家廠商。');
    want[id] = v;
  });
  const vendorGroups = readVendorGroups_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // 整欄讀、整欄寫（一格一格寫幾百格太慢）
    const column = function (table, header) {
      const n = table.sheet.getLastRow() - 1, c = table.cols[header];
      const range = table.sheet.getRange(2, c, Math.max(n, 1), 1);
      return { range: range, values: n > 0 ? range.getValues() : [], dirty: false };
    };
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']);
    const iv = column(t, '廠商');
    let items = 0;
    t.rows.forEach(function (r) {
      const v = want[String(r['食材ID']).trim()];
      if (!v || String(r['廠商'] || '').trim() === v) return;
      iv.values[r._row - 2][0] = v; iv.dirty = true;
      items++;
    });
    if (iv.dirty) { iv.range.setValues(iv.values); clearItemsCache_(); }
    const o = readTable_(SHEET.orders, ['訂單編號', '食材ID', '廠商']);
    const ov = column(o, '廠商'), og = o.cols['送貨組別'] !== undefined ? column(o, '送貨組別') : null;
    const today = today_();
    let lines = 0;
    o.rows.forEach(function (x) {
      const v = want[String(x['食材ID'] || '').trim()];
      if (!v) return;
      const d = toDate_(x['使用日期']), st = String(x['狀態'] || '').trim();
      if (!d || d < today || (st !== '待處理' && st !== '已審核')) return;
      const nv = orderVendor_(v, x['科別']);
      if (String(x['廠商'] || '').trim() === nv) return;
      ov.values[x._row - 2][0] = nv; ov.dirty = true;
      if (og) { og.values[x._row - 2][0] = deliveryGroup_(groupVendor_(nv), d.getDay(), vendorGroups); og.dirty = true; }
      lines++;
    });
    if (ov.dirty) ov.range.setValues(ov.values);
    if (og && og.dirty) og.range.setValues(og.values);
    return { items: items, lines: lines };
  } finally {
    lock.releaseLock();
  }
}

/** 食材總表的廠商改成 vendor（訂單上的廠商名稱，例如 E烘焙乾貨、A菜商），還沒叫貨的訂單（今天以後、待處理或已審核）也跟著改。回傳改了的訂單品項 [{ orderNo, vendor, group }] */
function setItemVendor_(id, vendor) {
  const vendorGroups = readVendorGroups_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']);
    const r = t.rows.find(function (x) { return String(x['食材ID']).trim() === id; });
    if (!r) return [];
    if (String(r['廠商'] || '').trim() !== vendor) { setCell_(t, r._row, '廠商', vendor); clearItemsCache_(); }
    const o = readTable_(SHEET.orders, ['訂單編號', '食材ID', '廠商']);
    const today = today_();
    const out = [];
    o.rows.forEach(function (x) {
      if (String(x['食材ID'] || '').trim() !== id) return;
      const d = toDate_(x['使用日期']), st = String(x['狀態'] || '').trim();
      if (!d || d < today || (st !== '待處理' && st !== '已審核')) return;
      const nv = orderVendor_(vendor, x['科別']);
      if (String(x['廠商'] || '').trim() === nv) return;
      const g = deliveryGroup_(groupVendor_(nv), d.getDay(), vendorGroups);
      setCell_(o, x._row, '廠商', nv);
      if (o.cols['送貨組別'] !== undefined) setCell_(o, x._row, '送貨組別', g);
      out.push({ orderNo: String(x['訂單編號']).trim(), vendor: nv, group: g });
    });
    return out;
  } finally {
    lock.releaseLock();
  }
}

/** 科別替換：{ 科別: { 食材(比對用): 換成的食材名稱 } } */
function readDeptSubs_() {
  const out = {};
  try {
    if (!ss_().getSheetByName(SHEET.subs)) return out;
    readTable_(SHEET.subs, ['科別', '食材', '換成']).rows.forEach(function (r) {
      const d = String(r['科別'] || '').trim(), from = String(r['食材'] || '').trim(), to = String(r['換成'] || '').trim();
      if (d && from && to) (out[d] = out[d] || {})[learnKey_(from)] = to;
    });
  } catch (e) { console.error('readDeptSubs_ ' + e.message); }
  return out;
}

function listDeptSubs_() {
  if (!ss_().getSheetByName(SHEET.subs)) return [];
  return readTable_(SHEET.subs, ['科別', '食材', '換成']).rows.map(function (r) {
    return { row: r._row, dept: String(r['科別'] || '').trim(), from: String(r['食材'] || '').trim(), to: String(r['換成'] || '').trim(), note: String(r['備註'] || '') };
  }).filter(function (x) { return x.dept && x.from && x.to; });
}

/** 管理者：科別替換清單 */
function adminGetDeptSubs(token) { auth_(token, true); return listDeptSubs_(); }

/** 管理者：新增或修改一條科別替換（同科別同食材就覆蓋） */
function adminSetDeptSub(token, dept, from, to, note) {
  auth_(token, true);
  dept = clean_(dept, 20); from = clean_(from, 40); to = clean_(to, 40);
  if (!dept || !from || !to) throw new Error('請填科別、食材、換成。');
  if (learnKey_(from) === learnKey_(to)) throw new Error('「食材」和「換成」不能一樣。');
  const ok = cachedItems_().some(function (it) { return it.name === to; });
  if (!ok) throw new Error('食材總表沒有「' + to + '」（或已停用），請用正式名稱。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.subs, SUB_HEADERS);
    const t = readTable_(SHEET.subs, ['科別', '食材', '換成']);
    const r = t.rows.find(function (x) { return String(x['科別']).trim() === dept && learnKey_(x['食材']) === learnKey_(from); });
    if (r) { setCell_(t, r._row, '換成', to); if (t.cols['備註'] !== undefined) setCell_(t, r._row, '備註', clean_(note, 60)); }
    else appendRows_(SHEET.subs, [{ '科別': dept, '食材': from, '換成': to, '備註': clean_(note, 60) }]);
  } finally {
    lock.releaseLock();
  }
  return listDeptSubs_();
}

function adminDeleteDeptSub(token, row) {
  auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName(SHEET.subs);
    if (!sh || !(Number(row) >= 2)) throw new Error('找不到這條規則，請重新整理。');
    deleteRows_(sh, [Number(row)]);
  } finally {
    lock.releaseLock();
  }
  return listDeptSubs_();
}

/** 管理者：這段期間手動加的跨教室領料 */
function listPickups_(start, end) {
  if (!ss_().getSheetByName(SHEET.pickups)) return [];
  return readTable_(SHEET.pickups, ['使用日期', '食材名稱']).rows.map(function (r) {
    return { row: r._row, date: dateStr_(toDate_(r['使用日期'])), time: String(r['上課時間'] || ''), className: String(r['班級'] || ''),
      teacher: String(r['老師'] || ''), dept: String(r['科別'] || ''), room: String(r['教室'] || ''), place: String(r['領取地點'] || '').trim(),
      name: String(r['食材名稱'] || '').trim(), qty: r['數量'], unit: String(r['單位'] || ''), note: String(r['備註'] || '') };
  }).filter(function (x) {
    const d = parseDate_(x.date);
    return x.name && x.place && d && d >= start && d <= end;
  });
}

function adminGetPickups(token, from, to) {
  auth_(token, true);
  const start = parseDate_(from), end = parseDate_(to) || start;
  if (!start) throw new Error('請選擇日期。');
  return listPickups_(start, end);
}

function adminAddPickup(token, rec) {
  const me = auth_(token, true);
  rec = rec || {};
  const d = parseDate_(rec.date);
  if (!d) throw new Error('請選擇日期。');
  const time = clean_(rec.time, 10);
  if (time && !CLASS_TIME_PATTERN.test(time)) throw new Error('上課時間請選上午／下午／夜間。');
  const place = clean_(rec.place, 20), name = clean_(rec.name, 40);
  if (!place) throw new Error('請填領取地點。');
  if (PICKUP_PLACES.indexOf(place) < 0) throw new Error('領取地點只能選烘焙教室、中餐教室或飲調教室。');
  if (!name) throw new Error('請填食材。');
  const qty = Number(rec.qty);
  if (!(qty > 0)) throw new Error('數量要大於 0。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.pickups, PICKUP_HEADERS);
    appendRows_(SHEET.pickups, [{ '使用日期': d, '上課時間': time, '班級': clean_(rec.className, 20), '老師': clean_(rec.teacher, 20),
      '科別': clean_(rec.dept, 20), '教室': clean_(rec.room, 20), '領取地點': place, '食材名稱': name, '數量': qty,
      '單位': clean_(rec.unit, 10), '備註': clean_(rec.note, 60), '建立者': me, '建立時間': new Date() }]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

function adminDeletePickup(token, row) {
  auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName(SHEET.pickups);
    if (!sh || !(Number(row) >= 2)) throw new Error('找不到這筆，請重新產生。');
    deleteRows_(sh, [Number(row)]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

/**
 * 管理者：食材改名／合併，食材總表、還沒叫貨的訂單（待處理、已審核）、老師範本、今天以後的叫貨紀錄一起改，
 * 舊名字自動加成別名（老師打舊名字也找得到）。直接在試算表改名的話，已經送出的訂單不會跟著改。
 * adminRenameItem(id, 新名稱)；adminMergeItem(被合併的 id, 留下來的 id)：被合併的那項停用。
 */
function adminRenameItem(token, id, newName) {
  auth_(token, true);
  id = String(id || '').trim(); newName = clean_(newName, 40);
  if (!newName) throw new Error('請填新名稱。');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let res;
  try {
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱']);
    const r = t.rows.find(function (x) { return String(x['食材ID']).trim() === id; });
    if (!r) throw new Error('找不到食材 ' + id + '。');
    const oldName = String(r['食材名稱']).trim();
    if (oldName === newName) throw new Error('名稱沒有改。');
    if (t.rows.some(function (x) { return x !== r && String(x['食材名稱']).trim() === newName; })) {
      throw new Error('食材總表已經有「' + newName + '」，兩項是同一樣東西的話請用「合併」。');
    }
    setCell_(t, r._row, '食材名稱', newName);
    clearItemsCache_();
    res = relinkItem_(id, id, newName, null);
    res.alias = addAlias_(id, oldName);
    res.oldName = oldName;
  } finally {
    lock.releaseLock();
  }
  return res;
}

function adminMergeItem(token, fromId, toId) {
  auth_(token, true);
  fromId = String(fromId || '').trim(); toId = String(toId || '').trim();
  if (!fromId || !toId || fromId === toId) throw new Error('請選兩項不同的食材。');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let res;
  try {
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱', '廠商']);
    const a = t.rows.find(function (x) { return String(x['食材ID']).trim() === fromId; });
    const b = t.rows.find(function (x) { return String(x['食材ID']).trim() === toId; });
    if (!a || !b) throw new Error('找不到食材，請重新整理。');
    const aName = String(a['食材名稱']).trim(), bName = String(b['食材名稱']).trim();
    if (t.cols['啟用'] !== undefined) setCell_(t, a._row, '啟用', false);
    clearItemsCache_();
    res = relinkItem_(fromId, toId, bName, String(b['廠商'] || '').trim());
    res.alias = [aName].concat(String(a['別名'] || '').split(/[,，、]+/)).map(function (x) { return String(x).trim(); })
      .filter(Boolean).filter(function (x) { return addAlias_(toId, x); });
    res.fromName = aName; res.toName = bName;
  } finally {
    lock.releaseLock();
  }
  return res;
}

/** 把 fromId 的訂單（待處理、已審核）、範本、今天以後的叫貨紀錄換成 toId／name（vendor 有給就連廠商、送貨組別一起換）。回傳各改了幾筆 */
function relinkItem_(fromId, toId, name, vendor) {
  const out = { orders: 0, templates: 0, purchase: 0 };
  const vendorGroups = readVendorGroups_();
  const o = readTable_(SHEET.orders, ['訂單編號', '食材ID', '食材名稱']);
  o.rows.forEach(function (x) {
    if (String(x['食材ID'] || '').trim() !== fromId) return;
    const st = String(x['狀態'] || '').trim();
    if (st !== '待處理' && st !== '已審核') return;
    if (toId !== fromId) setCell_(o, x._row, '食材ID', toId);
    if (String(x['食材名稱']).trim() !== name) setCell_(o, x._row, '食材名稱', name);
    if (vendor && o.cols['廠商'] !== undefined) {
      const nv = orderVendor_(vendor, x['科別']), d = toDate_(x['使用日期']);
      if (String(x['廠商'] || '').trim() !== nv) {
        setCell_(o, x._row, '廠商', nv);
        if (d && o.cols['送貨組別'] !== undefined) setCell_(o, x._row, '送貨組別', deliveryGroup_(groupVendor_(nv), d.getDay(), vendorGroups));
      }
    }
    out.orders++;
  });
  const tsh = ss_().getSheetByName(SHEET.templates);
  if (tsh && tsh.getLastRow() >= 2) {
    const t = readTable_(SHEET.templates, ['食材ID', '名稱']);
    t.rows.forEach(function (x) {
      if (String(x['食材ID'] || '').trim() !== fromId) return;
      if (toId !== fromId) setCell_(t, x._row, '食材ID', toId);
      setCell_(t, x._row, '名稱', name);
      out.templates++;
    });
  }
  const psh = ss_().getSheetByName(SHEET.purchase);
  if (psh && psh.getLastRow() >= 2) {
    const today = dateStr_(today_());
    const p = readTable_(SHEET.purchase, ['送貨日期', '食材名稱', '食材ID']);
    p.rows.forEach(function (x) {
      if (String(x['食材ID'] || '').trim() !== fromId || purchaseDate_(x['送貨日期']) < today) return;
      if (toId !== fromId) setCell_(p, x._row, '食材ID', toId);
      setCell_(p, x._row, '食材名稱', name);
      out.purchase++;
    });
  }
  return out;
}

/** 比對用：去掉空白、括號，英文小寫（跟網頁的 normName 一樣） */
function learnKey_(t) { return String(t || '').replace(/[\s()（）\[\]【】]/g, '').toLowerCase(); }

/** 正式食材的名稱、別名索引：{ name: {key: [id]}, alias: {key: [id]} } */
function learnIndex_(catalog) {
  const idx = { name: {}, alias: {} };
  const put = function (m, k, id) { if (k && (m[k] = m[k] || []).indexOf(id) < 0) m[k].push(id); };
  Object.keys(catalog).forEach(function (id) {
    const r = catalog[id];
    if (id === PENDING_ID || id === TOOL_ID) return;
    put(idx.name, learnKey_(r['食材名稱']), id);
    String(r['別名'] || '').split(/[,，、]+/).forEach(function (a) { put(idx.alias, learnKey_(a), id); });
  });
  return idx;
}

/** 名稱完全一樣優先；別名只有對到「一種」食材才算（例如「糖」同時是細砂糖和二砂的別名，就不自動換，讓老師自己選） */
function learnMatch_(idx, name) {
  const k = learnKey_(name);
  if (!k) return '';
  const n = idx.name[k];
  if (n && n.length === 1) return n[0];
  if (n && n.length > 1) return '';
  const a = idx.alias[k];
  return a && a.length === 1 ? a[0] : '';
}

/** 食材總表現在寫「自備」的食材ID（例如水）：舊訂單當時記成別家廠商的，也一律不叫貨 */
let SELF_IDS_ = null;   // 同一次執行只讀一次
function selfItemIds_() {
  if (SELF_IDS_) return SELF_IDS_;
  const cache = CacheService.getScriptCache();
  try { const c = cache.get('selfIds'); if (c) return SELF_IDS_ = JSON.parse(c); } catch (e) {}
  const out = SELF_IDS_ = {};
  readTable_(SHEET.items, ['食材ID', '廠商']).rows.forEach(function (r) {
    if (String(r['廠商']).trim() === SELF_VENDOR) out[String(r['食材ID']).trim()] = true;
  });
  try { cache.put('selfIds', JSON.stringify(out), 600); } catch (e) {}   // 10 分鐘；食材總表改成自備後最晚 10 分鐘生效
  return out;
}

/** 管理者：系統學到的叫法（食材總表的「別名」）。[{ id, name, vendor, aliases: [], shared: {別名: [其他食材名稱]} }] */
function adminGetAliases(token) {
  auth_(token, true);
  const t = readTable_(SHEET.items, ['食材ID', '食材名稱']);
  if (t.cols['別名'] === undefined) return { ok: false, list: [] };
  const owners = {};
  const list = [];
  t.rows.forEach(function (r) {
    if (r['啟用'] !== undefined && !isTrue_(r['啟用'])) return;
    const aliases = String(r['別名'] || '').split(/[,，、]+/).map(function (a) { return a.trim(); }).filter(Boolean);
    if (!aliases.length) return;
    const x = { id: String(r['食材ID']).trim(), name: String(r['食材名稱']).trim(), vendor: String(r['廠商'] || '').trim(), aliases: aliases };
    aliases.forEach(function (a) { (owners[learnKey_(a)] = owners[learnKey_(a)] || []).push(x.name); });
    list.push(x);
  });
  list.forEach(function (x) {
    x.shared = {};
    x.aliases.forEach(function (a) {
      const o = owners[learnKey_(a)].filter(function (n) { return n !== x.name; });
      if (o.length) x.shared[a] = o;
    });
  });
  return { ok: true, list: list.sort(function (a, b) { return a.name.localeCompare(b.name, 'zh-Hant'); }) };
}

/**
 * 管理者：自己新增叫法（例如細砂糖 ← 砂糖、細糖、糖），老師打這些字會直接找到這項。
 * 叫法本身是另一個正式食材的名稱就不加（要用「食材改名／合併」）；已經是別的食材的叫法會加，但會提醒（兩項都有就不自動換）。
 */
function adminAddAliases(token, name, text) {
  auth_(token, true);
  name = clean_(name, 40);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const out = { added: [], skipped: [], shared: [] };
  try {
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱']);
    if (t.cols['別名'] === undefined) throw new Error('食材總表沒有「別名」欄。');
    const live = t.rows.filter(function (x) { return x['啟用'] === undefined || isTrue_(x['啟用']); });
    const r = live.find(function (x) { return String(x['食材名稱']).trim() === name; });
    if (!r) throw new Error('食材總表沒有「' + name + '」（或已停用），請從清單選正式名稱。');
    const id = String(r['食材ID']).trim();
    const list = String(text || '').split(/[,，、;；\s]+/).map(function (x) { return clean_(x, 40); }).filter(Boolean)
      .filter(function (x, i, a) { return a.indexOf(x) === i; });
    if (!list.length) throw new Error('請填要新增的叫法。');
    list.forEach(function (a) {
      if (learnKey_(a) === learnKey_(name)) { out.skipped.push(a + '（就是正式名稱）'); return; }
      const other = live.find(function (x) { return x !== r && learnKey_(x['食材名稱']) === learnKey_(a); });
      if (other) { out.skipped.push(a + '（是另一項食材的名稱，請用「食材改名／合併」把「' + a + '」合併到「' + name + '」）'); return; }
      const also = live.filter(function (x) { return x !== r && String(x['別名'] || '').split(/[,，、]+/).some(function (y) { return learnKey_(y) === learnKey_(a); }); })
        .map(function (x) { return String(x['食材名稱']).trim(); });
      if (addAlias_(id, a)) { out.added.push(a); if (also.length) out.shared.push(a + '（「' + also.join('、') + '」也有這個叫法）'); }
      else out.skipped.push(a + '（已經有了）');
    });
  } finally {
    lock.releaseLock();
  }
  out.list = adminGetAliases(token);
  return out;
}

/** 管理者：刪掉一個學錯的叫法 */
function adminRemoveAlias(token, id, alias) {
  auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱']);
    if (t.cols['別名'] === undefined) throw new Error('食材總表沒有「別名」欄。');
    const r = t.rows.find(function (x) { return String(x['食材ID']).trim() === String(id || '').trim(); });
    if (!r) throw new Error('找不到食材 ' + id + '，請重新整理。');
    const list = String(r['別名'] || '').split(/[,，、]+/).map(function (a) { return a.trim(); }).filter(Boolean);
    const left = list.filter(function (a) { return a !== String(alias || '').trim(); });
    if (left.length === list.length) throw new Error('「' + alias + '」已經不在「' + r['食材名稱'] + '」的別名裡。');
    t.sheet.getRange(r._row, t.cols['別名']).setNumberFormat('@').setValue(left.join(','));
    clearItemsCache_();
    return adminGetAliases(token);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 把老師的叫法加到食材的「別名」（例如「全蛋」對應到「雞蛋」後，下次搜尋「全蛋」就找得到雞蛋）。
 * 名稱或別名已經有就不加；食材總表沒有「別名」欄就略過。回傳有沒有加。
 */
function addAlias_(id, alias) {
  alias = clean_(alias, 40);
  if (!alias) return false;
  const t = readTable_(SHEET.items, ['食材ID', '食材名稱']);
  if (t.cols['別名'] === undefined) return false;
  const r = t.rows.find(function (x) { return String(x['食材ID']).trim() === id; });
  if (!r || String(r['食材名稱']).trim() === alias) return false;
  // 這個叫法本身就是另一個正式食材的名稱（例如把「糖粉」對應成細砂糖）：不學，避免以後把糖粉都換掉
  if (t.rows.some(function (x) { return x !== r && (x['啟用'] === undefined || isTrue_(x['啟用'])) && learnKey_(x['食材名稱']) === learnKey_(alias); })) return false;
  const cur = String(r['別名'] || '').trim();
  const list = cur ? cur.split(/[,，、\s]+/).filter(Boolean) : [];
  if (list.indexOf(alias) >= 0) return false;
  list.push(alias);
  const v = list.join(',');
  if (v.length > 200) return false;
  t.sheet.getRange(r._row, t.cols['別名']).setNumberFormat('@').setValue(v);
  clearItemsCache_();
  return true;
}

/** 把「所有」訂單中同名的「待審核」食材換成正式的食材 ID、名稱、廠商（送貨組別跟著廠商重算）。回傳更新幾筆。 */
function fixPendingLines_(origName, to) {
  const t = readTable_(SHEET.orders, ['訂單編號', '食材ID', '食材名稱']);
  const vendorGroups = readVendorGroups_();
  let n = 0;
  t.rows.forEach(function (r) {
    if (String(r['食材ID']).trim() !== PENDING_ID || String(r['食材名稱']).trim() !== origName) return;
    n++;
    const v = { '食材ID': to.id, '食材名稱': to.name };
    if (to.vendor && t.cols['廠商'] !== undefined) {
      v['廠商'] = orderVendor_(to.vendor, r['科別']);
      const d = toDate_(r['使用日期']);
      if (d && t.cols['送貨組別'] !== undefined) v['送貨組別'] = deliveryGroup_(to.vendor, d.getDay(), vendorGroups);
    }
    setCells_(t, r._row, v);
  });
  return n;
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
// 之後就會在訂貨截止那天、每天早上自動執行。也可以在網站「系統設定」頁管理。

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
  const days = [null, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
  ScriptApp.newTrigger('autoVendorOrders').timeBased()
    .onWeekDay(days[lockRule_().wd]).atHour(vendorOrderHour_()).inTimezone(tz_()).create();
  ScriptApp.newTrigger('dailyDigest').timeBased()
    .everyDays(1).atHour(DIGEST_HOUR).inTimezone(tz_()).create();
  ScriptApp.newTrigger('weeklyMaintenance').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(MAINT_HOUR).inTimezone(tz_()).create();
  ScriptApp.newTrigger('remindTeachers').timeBased()
    .onWeekDay(days[remindWeekday_()]).atHour(REMIND_HOUR).inTimezone(tz_()).create();
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
    remindDay: WEEKDAY[remindWeekday_()],
    lockRule: lockRule_(),
    lockDefault: { wd: LOCK_WEEKDAY, h: LOCK_HOUR, m: LOCK_MINUTE },
    lockDay: WEEKDAY[lockRule_().wd],
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
    emailsCustom: !!PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS'),
    lineTokenSet: !!props.getProperty('LINE_TOKEN'),
    lineBound: !!props.getProperty('LINE_TO'),
    webhookUrl: webAppUrl_(),
    siteUrlSaved: !!props.getProperty('SITE_URL'),
    autoUrl: autoUrl_(),
    lastVendorUrl: props.getProperty('LAST_VENDOR_URL') || '',
    lastVendorTime: props.getProperty('LAST_VENDOR_TIME') || '',
    vendorHour: vendorOrderHour_(),
    digestHour: DIGEST_HOUR,
  };
}

/** 排程呼叫的函式只接受真正的排程觸發，避免被網頁直接呼叫。 */
function fromTrigger_(e) {
  if (!e || !e.triggerUid) return false;
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getUniqueId() === String(e.triggerUid); });
}

// 排程：訂貨截止那天（預設週一中午）產生下週的廠商訂貨單
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

// 排程：截止前一個上課日（預設週五）下午用 LINE 私訊提醒下週還沒訂貨的老師
function remindTeachers(e) {
  if (!fromTrigger_(e)) throw new Error('這個功能只能由排程執行。');
  logged_('訂貨提醒', function () { runReminder_(true); });
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
      if (/^驗證碼$/.test(String(ev.message.text).trim())) {
        const codeAns = updateCodeReply_(ev, props);
        if (codeAns) { lineReply_(ev.replyToken, codeAns); return; }
      }
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
  const saved = String(PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS') || '').split(',').filter(Boolean);
  if (saved.length) return saved;   // 「系統設定 → Email 通知」設定的
  if (NOTIFY_EMAILS.length) return NOTIFY_EMAILS.slice();
  const me = Session.getEffectiveUser().getEmail();
  return me ? [me] : [];
}

/**
 * 傳到 LINE 的網站連結：加上 openExternalBrowser=1，LINE 會直接用手機的瀏覽器（Safari／Chrome）打開，
 * 不用 LINE 內建瀏覽器（內建瀏覽器開 Apps Script 網站常出現「目前無法開啟這個檔案」）。
 */
function lineUrl_(page) {
  const u = webAppUrl_();
  return u ? u + '?' + (page ? 'page=' + page + '&' : '') + 'openExternalBrowser=1' : '';
}

/** 網站網址：優先用管理者在「自動化」設定的網址（部署很多個時，Apps Script 自己回報的可能是舊的） */
function webAppUrl_() {
  const saved = PropertiesService.getScriptProperties().getProperty('SITE_URL');
  if (saved) return saved;
  return autoUrl_();
}

function autoUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** 管理者設定網站網址（LINE 通知、選單、摘要裡的連結都用這個）；空白 = 改回自動判斷 */
/** 管理者：設定收通知的 Email（最多 5 個，逗號分開）；空白＝改回寄給部署網站的帳號 */
function adminSetNotifyEmails(token, text) {
  auth_(token, true);
  const list = String(text || '').split(/[\s,，、;；]+/).map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean)
    .filter(function (x, i, a) { return a.indexOf(x) === i; });
  if (list.length > 5) throw new Error('最多 5 個 Email。');
  list.forEach(function (x) { if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)) throw new Error('「' + x + '」不是正確的 Email。'); });
  const props = PropertiesService.getScriptProperties();
  if (list.length) props.setProperty('NOTIFY_EMAILS', list.join(',')); else props.deleteProperty('NOTIFY_EMAILS');
  return automationStatus_();
}

/** 管理者：只寄一封測試信（不傳 LINE，不花則數） */
function adminTestEmail(token) {
  auth_(token, true);
  return notify_('【測試】三信家商餐飲科訂貨系統 Email 通知',
    '<p>這是測試信。收到這封信表示 Email 通知設定正常，之後的每日摘要、叫貨確定等通知都會寄到這裡。</p>', null);
}

function adminSetSiteUrl(token, url) {
  auth_(token, true);
  url = String(url || '').trim().replace(/[?#].*$/, '');
  const props = PropertiesService.getScriptProperties();
  if (!url) {
    props.deleteProperty('SITE_URL');
  } else {
    if (!/^https:\/\/script\.google\.com\/(a\/macros\/[^\/]+|macros)\/s\/[\w-]+\/exec$/.test(url)) {
      throw new Error('網址格式不對，請貼上「管理部署」裡的網頁應用程式網址（https://script.google.com/macros/s/…/exec）。');
    }
    props.setProperty('SITE_URL', url);
  }
  return automationStatus_();
}

// =====================================================================
// 一鍵更新程式（管理者用手機也能更新）
// 保護：① 只有內建「管理者」帳號能用；② 每次更新／還原都要輸入傳到管理者 LINE 的驗證碼（錯 3 次鎖 1 小時）；
// ③ 只從固定的 GitHub 儲存庫讀程式，token 只有唯讀權限；④ 只更新到畫面上看到的那一版（版本代碼要一致）；
// ⑤ Apps Script 會擋語法錯誤，部署後自動測試網站，打不開就自動還原；⑥ 每次都記在「更新紀錄」分頁並用 LINE 通知。
// 需要：script.google.com/home/usersettings 打開「Google Apps Script API」，appsscript.json 有 script.projects、script.deployments 權限。
// =====================================================================
const UPDATE_REPO = 'win-sam0503/-';
const UPDATE_DIR = 'order-site';
const UPDATE_DEFAULT_BRANCH = 'claude/hi-ksvoju';
const UPDATE_FILES = { Code: ['Code.gs', 'SERVER_JS'], Index: ['Index.html', 'HTML'], RichMenu: ['RichMenu.html', 'HTML'] };
const UPDATE_LOG_HEADERS = ['時間', '動作', '版本', '說明', '結果', '操作人'];

/** 只有內建的管理者帳號（不含後來設定的管理者）可以更新程式 */
function updateOwner_(token) {
  const account = auth_(token, true);
  if (ADMIN_ACCOUNTS.indexOf(account) < 0) throw new Error('只有內建的「管理者」帳號可以更新程式。');
  return account;
}

function gh_(path, raw) {
  const tk = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!tk) throw new Error('還沒設定 GitHub token，請先在「更新程式」的設定貼上。');
  const res = UrlFetchApp.fetch('https://api.github.com/repos/' + UPDATE_REPO + '/' + path, {
    headers: { Authorization: 'Bearer ' + tk, Accept: raw ? 'application/vnd.github.raw' : 'application/vnd.github+json',
      'User-Agent': 'order-site-updater', 'X-GitHub-Api-Version': '2022-11-28' },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code === 401) throw new Error('GitHub token 不正確或已過期，請重新產生一組貼上。');
  if (code === 403) throw new Error('GitHub 拒絕存取（token 沒有這個儲存庫的「Contents 唯讀」權限，或暫時超過次數）。');
  if (code === 404) throw new Error('GitHub 上找不到（' + path.split('?')[0] + '）：請確認分支名稱和 token 的儲存庫權限。');
  if (code >= 300) throw new Error('GitHub 回應錯誤（' + code + '）。');
  return raw ? res.getContentText('UTF-8') : JSON.parse(res.getContentText('UTF-8'));
}

function gas_(method, path, body) { return gasAt_(ScriptApp.getScriptId(), method, path, body); }

/** Apps Script API：scriptId 空白＝建立新專案（POST /v1/projects） */
function gasAt_(scriptId, method, path, body) {
  const opt = { method: method, contentType: 'application/json', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  if (body) opt.payload = JSON.stringify(body);
  const res = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects' + (scriptId ? '/' + scriptId : '') + path, opt);
  const code = res.getResponseCode(), text = res.getContentText('UTF-8');
  if (code >= 300) {
    let m = '';
    try { m = JSON.parse(text).error.message; } catch (e) { m = text.slice(0, 300); }
    if (/not enabled the Apps Script API/i.test(m)) {
      throw new Error('還沒打開「Google Apps Script API」：請用電腦到 script.google.com/home/usersettings 打開，等幾分鐘再試。');
    }
    if (/has not been used in project|is disabled/i.test(m)) {
      const link = (m.match(/https:\/\/console\.\S+?(?=[\s.]*$|\s)/) || [])[0] || '';
      const proj = (m.match(/project (\d+)/) || [])[1] || '';
      throw new Error('個人設定的 Apps Script API 已開啟，但這個程式所屬的 Google Cloud 專案' + (proj ? '（' + proj + '）' : '') +
        '還沒啟用 Apps Script API。請用電腦打開：' + (link || 'https://console.cloud.google.com/apis/library/script.googleapis.com' + (proj ? '?project=' + proj : '')) +
        ' 按「啟用」，等 5 分鐘再試。打不開或沒有權限時，要把 Apps Script 換到自己的 Google Cloud 專案（見說明）。');
    }
    if (code === 403) throw new Error('程式沒有「管理 Apps Script 專案」的權限：請用電腦貼上新的 appsscript.json，再執行一次授權。（' + m + '）');
    throw new Error(m || ('Apps Script API 錯誤（' + code + '）'));
  }
  return text ? JSON.parse(text) : {};
}

/** 現在網站用的部署 ID（從「自動化 → 網站網址」取） */
function updateDeploymentId_() {
  const m = String(webAppUrl_() || '').match(/\/s\/([\w-]+)\/exec/);
  if (!m) throw new Error('找不到網站的部署 ID，請先在「自動化 → 網站網址」貼上正在使用的網址。');
  return m[1];
}

function updateState_() {
  const p = PropertiesService.getScriptProperties();
  return { sha: p.getProperty('UPDATE_SHA') || '', at: p.getProperty('UPDATE_AT') || '', version: Number(p.getProperty('UPDATE_VERSION') || 0),
    prevVersion: Number(p.getProperty('UPDATE_PREV_VERSION') || 0), prevSha: p.getProperty('UPDATE_PREV_SHA') || '' };
}

function updateLog_(action, version, note, result, who) {
  try {
    ensureSheet_('更新紀錄', UPDATE_LOG_HEADERS);
    appendRows_('更新紀錄', [{ '時間': new Date(), '動作': action, '版本': version, '說明': String(note || '').slice(0, 200), '結果': result, '操作人': who }],
      { '時間': 'yyyy/mm/dd hh:mm' });
  } catch (e) { console.error('updateLog_ ' + e.message); }
  // 結果在網站上和「更新紀錄」看得到，不另外推播（省 LINE 則數）；只有驗證碼錯太多次被鎖定時通知
  if (/^拒絕/.test(String(result))) linePush_(PropertiesService.getScriptProperties().getProperty('LINE_TO'), '【訂貨系統更新】' + action + '：' + result + (version ? '（' + version + '）' : '') + (note ? '\n' + note : ''));
}

/** 管理者：更新程式的狀態（目前版本、GitHub 上的版本） */
function adminUpdateInfo(token, branch) {
  updateOwner_(token);
  const p = PropertiesService.getScriptProperties();
  branch = clean_(branch, 100) || p.getProperty('UPDATE_BRANCH') || UPDATE_DEFAULT_BRANCH;
  const out = { hasToken: !!p.getProperty('GITHUB_TOKEN'), lineReady: !!(p.getProperty('LINE_TOKEN') && p.getProperty('LINE_TO')),
    branch: branch, current: updateState_(), repo: UPDATE_REPO, lockedUntil: Number(p.getProperty('UPDATE_LOCK_UNTIL') || 0) };
  if (!out.hasToken) return out;
  out.receive = { url: p.getProperty('RECEIVE_URL') || '', sha: p.getProperty('RECEIVE_SHA') || '', at: p.getProperty('RECEIVE_AT') || '' };
  try {
    out.branches = gh_('branches?per_page=100').map(function (b) { return b.name; });
    const rc = gh_('commits?sha=' + encodeURIComponent(branch) + '&path=' + RECEIVE_DIR + '&per_page=1')[0];
    if (rc) out.receive.latest = { sha: rc.sha, message: String(rc.commit.message || '').split('\n')[0].slice(0, 120), date: rc.commit.committer ? rc.commit.committer.date : '' };
    out.commits = gh_('commits?sha=' + encodeURIComponent(branch) + '&path=' + UPDATE_DIR + '&per_page=10').map(function (c) {
      return { sha: c.sha, message: String(c.commit.message || '').split('\n')[0].slice(0, 120), date: c.commit.committer ? c.commit.committer.date : '' };
    });
  } catch (e) {
    out.error = e.message;
  }
  return out;
}

/** 管理者：存 GitHub token（空白＝不改；"-"＝刪除）和分支 */
function adminSaveUpdateSettings(token, ghToken, branch) {
  const who = updateOwner_(token);
  const p = PropertiesService.getScriptProperties();
  const t = String(ghToken || '').trim();
  if (t === '-') { p.deleteProperty('GITHUB_TOKEN'); updateLog_('刪除 GitHub token', '', '', '完成', who); }
  else if (t) {
    if (!/^(github_pat_|ghp_)\w+$/.test(t)) throw new Error('這看起來不像 GitHub token（應該是 github_pat_ 開頭）。');
    p.setProperty('GITHUB_TOKEN', t);
    updateLog_('設定 GitHub token', '', '', '完成', who);
  }
  const b = clean_(branch, 100);
  if (b) p.setProperty('UPDATE_BRANCH', b);
  return adminUpdateInfo(token, b);
}

/**
 * 管理者：要更新或還原前，先產生 6 位數驗證碼（5 分鐘有效）。
 * 預設不主動推播（推播會算 LINE 則數）：管理者在綁定通知的 LINE 聊天室傳「驗證碼」，webhook 免費回覆。
 * push = true：收不到回覆時改用推播（算 1 則）。
 */
function adminRequestUpdateCode(token, action, sha, push) {
  updateOwner_(token);
  const p = PropertiesService.getScriptProperties();
  if (Number(p.getProperty('UPDATE_LOCK_UNTIL') || 0) > Date.now()) throw new Error('驗證碼錯太多次，已鎖定 1 小時，請稍後再試。');
  if (action !== 'update' && action !== 'rollback' && action !== 'receive') throw new Error('動作不正確。');
  const to = p.getProperty('LINE_TO');
  if (!to || !p.getProperty('LINE_TOKEN')) throw new Error('要先設定 LINE 通知（自動化 → LINE 官方帳號通知），才收得到驗證碼。');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('updCode', JSON.stringify({ code: code, action: action, sha: String(sha || ''), tries: 0 }), 300);
  if (!push) return { sent: false };
  if (!linePush_(to, updateCodeText_(action, sha, code))) {
    throw new Error('驗證碼傳送失敗，請確認 LINE 設定或本月則數。');
  }
  return { sent: true };
}

function updateCodeText_(action, sha, code) {
  const what = action === 'update' ? '更新程式到版本 ' + String(sha || '').slice(0, 7)
    : action === 'receive' ? '建立／更新進貨驗收網站到版本 ' + String(sha || '').slice(0, 7) : '還原成上一版';
  return '【訂貨系統】' + what + ' 的驗證碼：' + code + '\n5 分鐘內有效。\n如果不是你按的，請不要告訴任何人，並馬上修改管理者密碼。';
}

/** LINE 傳「驗證碼」：只有綁定成管理者通知的那個聊天室（LINE_TO）會回覆，回覆不算則數 */
function updateCodeReply_(ev, props) {
  const src = ev.source && (ev.source.groupId || ev.source.roomId || ev.source.userId);
  if (!src || src !== props.getProperty('LINE_TO')) return null;
  const s = JSON.parse(CacheService.getScriptCache().get('updCode') || 'null');
  if (!s) return '目前沒有要更新或還原的動作。請先到網站「更新程式」按「⟳ 更新到…」，再傳「驗證碼」。';
  return updateCodeText_(s.action, s.sha, s.code);
}

function verifyUpdateCode_(action, sha, code, who) {
  const cache = CacheService.getScriptCache(), p = PropertiesService.getScriptProperties();
  if (Number(p.getProperty('UPDATE_LOCK_UNTIL') || 0) > Date.now()) throw new Error('驗證碼錯太多次，已鎖定 1 小時，請稍後再試。');
  const s = JSON.parse(cache.get('updCode') || 'null');
  if (!s || s.action !== action || (action !== 'rollback' && s.sha !== sha)) throw new Error('驗證碼已過期或不是這個動作的，請重新按「傳驗證碼」。');
  if (String(code || '').trim() !== s.code) {
    s.tries++;
    if (s.tries >= 3) {
      cache.remove('updCode');
      p.setProperty('UPDATE_LOCK_UNTIL', String(Date.now() + 3600 * 1000));
      updateLog_(action === 'update' ? '更新程式' : '還原', '', '驗證碼錯 3 次，鎖定 1 小時', '拒絕', who);
      throw new Error('驗證碼錯 3 次，已鎖定 1 小時。');
    }
    cache.put('updCode', JSON.stringify(s), 300);
    throw new Error('驗證碼不對，還可以再試 ' + (3 - s.tries) + ' 次。');
  }
  cache.remove('updCode');
}

/** 部署後測試：打開網站的健康檢查（?health=1），新程式能讀取、畫面能產生才算成功 */
function updateHealthy_() {
  try {
    // 帶上自己的授權，網站設成「只限學校帳號」也打得開
    const res = UrlFetchApp.fetch('https://script.google.com/macros/s/' + updateDeploymentId_() + '/exec?health=1', {
      muteHttpExceptions: true, followRedirects: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    return res.getResponseCode() === 200 && /ORDER_SITE_OK/.test(res.getContentText());
  } catch (e) {
    return false;
  }
}

function deployVersion_(version, note) {
  gas_('put', '/deployments/' + updateDeploymentId_(), { deploymentConfig: { scriptId: ScriptApp.getScriptId(), versionNumber: version,
    manifestFileName: 'appsscript', description: String(note).slice(0, 100) } });
}

/** 管理者：輸入 LINE 驗證碼後，把網站更新成 GitHub 上 branch 的 sha 那一版 */
function adminRunUpdate(token, branch, sha, code) {
  const who = updateOwner_(token);
  branch = clean_(branch, 100) || UPDATE_DEFAULT_BRANCH;
  if (!/^[0-9a-f]{40}$/.test(String(sha || ''))) throw new Error('版本不正確，請重新整理。');
  verifyUpdateCode_('update', sha, code, who);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('有人正在更新，請稍等一下再試。');
  const short = sha.slice(0, 7);
  try {
    const head = gh_('commits?sha=' + encodeURIComponent(branch) + '&path=' + UPDATE_DIR + '&per_page=1')[0];
    if (!head || head.sha !== sha) throw new Error('GitHub 上的最新版本跟畫面上的不一樣，請按「重新整理」確認後再更新。');
    const msg = String(head.commit.message || '').split('\n')[0].slice(0, 80);
    const src = {};
    Object.keys(UPDATE_FILES).forEach(function (name) {
      try { src[name] = gh_('contents/' + UPDATE_DIR + '/' + UPDATE_FILES[name][0] + '?ref=' + sha, true); } catch (e) { if (name !== 'RichMenu') throw e; }
    });
    if (!/function doGet\s*\(/.test(src.Code || '') || src.Code.length < 50000) throw new Error('GitHub 上的 Code.gs 看起來不完整，沒有更新。');
    if (!/<\/html>\s*$/i.test(src.Index || '') || src.Index.length < 50000) throw new Error('GitHub 上的 Index.html 看起來不完整，沒有更新。');
    const manifest = gh_('contents/' + UPDATE_DIR + '/appsscript.json?ref=' + sha, true);
    const before = gas_('get', '/content');
    const files = JSON.parse(JSON.stringify(before.files || []));
    const cur = files.find(function (f) { return f.name === 'appsscript'; });
    const scopes = function (j) { try { return (JSON.parse(j).oauthScopes || []).slice().sort().join(' '); } catch (e) { return '?'; } };
    if (cur && scopes(cur.source) !== scopes(manifest)) throw new Error('這次更新需要新的權限（appsscript.json 改了），請用電腦更新並授權一次。');
    Object.keys(src).forEach(function (name) {
      const f = files.find(function (x) { return x.name === name; });
      if (f) f.source = src[name]; else files.push({ name: name, type: UPDATE_FILES[name][1], source: src[name] });
    });
    if (cur) cur.source = manifest;
    try {
      gas_('put', '/content', { files: files });
    } catch (e) {
      throw new Error('新程式存不進 Apps Script（' + e.message + '），網站維持原本的版本。');
    }
    const ver = gas_('post', '/versions', { description: 'GitHub ' + short + '：' + msg }).versionNumber;
    const dep = gas_('get', '/deployments/' + updateDeploymentId_());
    const prev = dep.deploymentConfig && dep.deploymentConfig.versionNumber || 0;
    deployVersion_(ver, '一鍵更新 ' + short + ' ' + msg);
    if (!updateHealthy_()) {
      // 新版本打不開：網站和排程都換回原本的程式
      if (prev) deployVersion_(prev, '自動還原（' + short + ' 測試失敗）');
      gas_('put', '/content', { files: before.files });
      updateLog_('更新程式', short, msg, '失敗，已自動還原', who);
      throw new Error('新版本部署後網站打不開，已經自動還原成原本的版本。請把這個狀況告訴 Claude。');
    }
    const p = PropertiesService.getScriptProperties(), st = updateState_();
    p.setProperties({ UPDATE_PREV_VERSION: String(prev), UPDATE_PREV_SHA: st.sha, UPDATE_SHA: sha, UPDATE_VERSION: String(ver),
      UPDATE_AT: fmtDateTime_(new Date()), UPDATE_BRANCH: branch });
    updateLog_('更新程式', short, msg, '成功（Apps Script 版本 ' + ver + '）', who);
    return { ok: true, version: ver, message: msg, info: adminUpdateInfo(token, branch) };
  } catch (e) {
    if (!/已經自動還原/.test(e.message)) updateLog_('更新程式', short, e.message, '沒有更新', who);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

const RECEIVE_DIR = 'receive-site';

/**
 * 管理者：輸入 LINE 驗證碼後，建立（第一次）或更新「進貨驗收」網站：
 * 在你的 Google 帳號新增一個 Apps Script 專案、貼上 GitHub 上 receive-site 的程式、部署成「只有我自己」打得開的網頁應用程式。
 */
function adminReceiveSite(token, branch, sha, code) {
  const who = updateOwner_(token);
  branch = clean_(branch, 100) || UPDATE_DEFAULT_BRANCH;
  if (!/^[0-9a-f]{40}$/.test(String(sha || ''))) throw new Error('版本不正確，請重新整理。');
  verifyUpdateCode_('receive', sha, code, who);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('有人正在更新，請稍等一下再試。');
  const p = PropertiesService.getScriptProperties(), short = sha.slice(0, 7);
  try {
    const head = gh_('commits?sha=' + encodeURIComponent(branch) + '&path=' + RECEIVE_DIR + '&per_page=1')[0];
    if (!head || head.sha !== sha) throw new Error('GitHub 上的最新版本跟畫面上的不一樣，請按「重新整理」確認後再做。');
    const msg = String(head.commit.message || '').split('\n')[0].slice(0, 80);
    const get = function (f) { return gh_('contents/' + RECEIVE_DIR + '/' + f + '?ref=' + sha, true); };
    const files = [{ name: 'appsscript', type: 'JSON', source: get('appsscript.json') },
      { name: 'Code', type: 'SERVER_JS', source: get('Code.gs') }, { name: 'Index', type: 'HTML', source: get('Index.html') }];
    if (!/function doGet\s*\(/.test(files[1].source) || !/<\/html>\s*$/i.test(files[2].source)) throw new Error('GitHub 上的進貨驗收程式看起來不完整，沒有更新。');
    let id = p.getProperty('RECEIVE_SCRIPT_ID'), created = false;
    if (!id) {
      id = gasAt_('', 'post', '', { title: '進貨驗收' }).scriptId;
      p.setProperty('RECEIVE_SCRIPT_ID', id);
      created = true;
    }
    gasAt_(id, 'put', '/content', { files: files });
    const ver = gasAt_(id, 'post', '/versions', { description: 'GitHub ' + short + '：' + msg }).versionNumber;
    const cfg = { deploymentConfig: { scriptId: id, versionNumber: ver, manifestFileName: 'appsscript', description: '進貨驗收 ' + short } };
    let dep = null, depId = p.getProperty('RECEIVE_DEPLOY_ID');
    if (depId) { try { dep = gasAt_(id, 'put', '/deployments/' + depId, cfg); } catch (e) { dep = null; } }
    if (!dep) { dep = gasAt_(id, 'post', '/deployments', cfg.deploymentConfig); depId = dep.deploymentId; }
    const web = (dep.entryPoints || []).filter(function (x) { return x.entryPointType === 'WEB_APP'; })[0];
    const url = web && web.webApp && web.webApp.url || 'https://script.google.com/macros/s/' + depId + '/exec';
    p.setProperties({ RECEIVE_DEPLOY_ID: depId, RECEIVE_URL: url, RECEIVE_SHA: sha, RECEIVE_AT: fmtDateTime_(new Date()) });
    updateLog_(created ? '建立進貨驗收網站' : '更新進貨驗收網站', short, msg, '成功（版本 ' + ver + '）', who);
    return { ok: true, created: created, url: url, info: adminUpdateInfo(token, branch) };
  } catch (e) {
    updateLog_('進貨驗收網站', short, e.message, '沒有完成', who);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** 管理者：輸入 LINE 驗證碼後，還原成上一版（網站和排程用的程式都換回去） */
function adminRollbackUpdate(token, code) {
  const who = updateOwner_(token);
  const st = updateState_();
  if (!st.prevVersion) throw new Error('沒有可以還原的上一版。');
  verifyUpdateCode_('rollback', '', code, who);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('有人正在更新，請稍等一下再試。');
  try {
    const old = gas_('get', '/content?versionNumber=' + st.prevVersion);
    gas_('put', '/content', { files: old.files });
    deployVersion_(st.prevVersion, '還原到版本 ' + st.prevVersion);
    PropertiesService.getScriptProperties().setProperties({ UPDATE_VERSION: String(st.prevVersion), UPDATE_SHA: st.prevSha,
      UPDATE_PREV_VERSION: '0', UPDATE_PREV_SHA: '', UPDATE_AT: fmtDateTime_(new Date()) });
    updateLog_('還原', st.prevSha ? st.prevSha.slice(0, 7) : '版本 ' + st.prevVersion, '', '成功', who);
    return { ok: true, version: st.prevVersion, info: adminUpdateInfo(token) };
  } finally {
    lock.releaseLock();
  }
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
  if (lineText === null) {
    out.line = '只寄 Email（不傳 LINE，省則數）';
  } else if (lineToken && lineTo) {
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
      out.pendArchived = archivePendingReviews_();
    } catch (e) {
      out.pendError = '待審核食材紀錄封存失敗：' + e.message;   // 不影響其他維護
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
  if (out.pendArchived) lines.push('已把 ' + out.pendArchived.rows + ' 筆審核完超過 ' + PENDING_KEEP_DAYS + ' 天的「待審核食材」紀錄搬到雲端硬碟「' + HISTORY_FOLDER_NAME + '」的「' + PENDING_ARCHIVE_NAME + '」');
  if (out.pendError) lines.push('⚠ ' + out.pendError);
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

// ---- 待審核食材：審核完（已核准、已對應、已退回）超過 PENDING_KEEP_DAYS 天的紀錄，搬到雲端硬碟的「待審核食材紀錄」檔案 ----
const PENDING_KEEP_DAYS = 30;
const PENDING_ARCHIVE_NAME = '待審核食材紀錄';
function archivePendingReviews_() {
  const sh = ss_().getSheetByName(SHEET.pending);
  if (!sh || sh.getLastRow() < 2) return null;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const t = readTable_(SHEET.pending, ['食材名稱', '審核狀態']);
    const cut = addDays_(today_(), -PENDING_KEEP_DAYS);
    const old = t.rows.filter(function (r) {
      const st = String(r['審核狀態'] || '').trim(), d = r['新增時間'] instanceof Date ? r['新增時間'] : null;
      return st && st !== '待審核' && d && d < cut;
    });
    if (!old.length) return null;
    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const raw = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
    const props = PropertiesService.getScriptProperties();
    let file = null;
    const id = props.getProperty('PEND_ARCHIVE_ID');
    if (id) { try { file = SpreadsheetApp.openById(id); } catch (e) { file = null; } }
    if (!file) {
      file = SpreadsheetApp.create(PENDING_ARCHIVE_NAME);
      DriveApp.getFileById(file.getId()).moveTo(folderNamed_(HISTORY_FOLDER_NAME));
      file.getSheets()[0].getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      file.getSheets()[0].setFrozenRows(1);
      props.setProperty('PEND_ARCHIVE_ID', file.getId());
    }
    const dst = file.getSheets()[0];
    const vals = old.map(function (r) { return raw[r._row - 1]; });
    dst.getRange(dst.getLastRow() + 1, 1, vals.length, header.length).setValues(vals);
    SpreadsheetApp.flush();
    deleteRows_(sh, old.map(function (r) { return r._row; }));   // 先寫進紀錄檔，成功了才刪
    return { rows: old.length, url: file.getUrl() };
  } finally {
    lock.releaseLock();
  }
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
  try { bootDirty_(sh.getName()); } catch (e) {}
  memoClear_();
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
// 有開「Drive API」服務就用它；沒開也可以（直接呼叫 Drive 網路 API）。
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
    if (typeof Drive === 'undefined') return convertViaRest_(blob, name, targetMime, ocr);   // 沒開「Drive API」服務：直接呼叫 Drive 網路 API
    if (Drive.Files.create) {   // Drive API v3
      const f = Drive.Files.create({ name: '[匯入暫存] ' + name, mimeType: targetMime }, blob, ocr ? { ocrLanguage: 'zh-TW' } : {});
      return f.id;
    }
    const f = Drive.Files.insert({ title: '[匯入暫存] ' + name, mimeType: targetMime }, blob,   // Drive API v2
      ocr ? { ocr: true, ocrLanguage: 'zh-TW', convert: true } : { convert: true });
    return f.id;
  });
}

/** 不用進階服務，直接用 Drive API v3 上傳並轉檔（權限跟 DriveApp 一樣） */
function convertViaRest_(blob, name, targetMime, ocr) {
  const boundary = 'import' + Date.now();
  const meta = JSON.stringify({ name: '[匯入暫存] ' + name, mimeType: targetMime });
  const head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
    '\r\n--' + boundary + '\r\nContent-Type: ' + (blob.getContentType() || 'application/octet-stream') + '\r\n\r\n').getBytes();
  const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id' + (ocr ? '&ocrLanguage=zh-TW' : ''), {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    payload: head.concat(blob.getBytes(), tail),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code >= 300) {
    let m = '';
    try { m = JSON.parse(res.getContentText()).error.message; } catch (e) { m = res.getContentText().slice(0, 200); }
    throw new Error('檔案轉換失敗（' + code + '）：' + m);
  }
  return JSON.parse(res.getContentText()).id;
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
  return memo_('readSchedule_', readScheduleRaw_).map(function (o) { return Object.assign({}, o); });
}
function readScheduleRaw_() {
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
      kind: scheduleKind_(r['類別']),
      from: toDate_(r['開始日期']) ? dateStr_(toDate_(r['開始日期'])) : '',
      to: toDate_(r['結束日期']) ? dateStr_(toDate_(r['結束日期'])) : '',
    };
  }).filter(function (x) {
    return x.need && x.teacher && x.day >= 0 && SESSIONS.indexOf(x.session) >= 0 && !/國技|預留/.test(x.course);
  });
}

function scheduleKind_(v) {
  const k = String(v || '').trim();
  return SCHEDULE_KINDS.indexOf(k) >= 0 ? k : SCHEDULE_KINDS[0];
}

/** 自動建立需要的分頁（第一次使用時） */
/** 分頁缺少的標題補在最右邊（舊版建立的分頁也能存新欄位） */
function ensureHeaders_(sh, headers) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const have = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const missing = headers.filter(function (h) { return have.indexOf(h) < 0; });
  if (!missing.length) return;
  const start = have.filter(Boolean).length ? lastCol + 1 : 1;
  sh.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold');
}

// ---- 今日抽查 ----

/**
 * 管理者：從這一天有領料的訂單隨機抽 n 張（系統抽，不是管理者挑）。
 * 不含已取消、還料單、只有器具／有貨／自備的訂單。scope = 要抽的廠商類別，例如 'ABCD'（預設：菜肉蛋海鮮）
 */
function adminAuditPick(token, date, n, scope, all) {
  auth_(token, true);
  const d = parseDate_(date);
  if (!d) throw new Error('請選擇日期。');
  n = Math.max(1, Math.min(10, Number(n) || 3));
  const cats = String(scope || 'ABCD').toUpperCase().replace(/[^A-F]/g, '') || 'ABCD';
  const pool = adminGetOrders(token, dateStr_(d), dateStr_(d), true).filter(function (o) {
    if (o.status === '已取消' || isReturn_(o.project)) return false;
    return o.items.some(function (it) {
      return it.id !== TOOL_ID && !it.noOrder && it.status !== '已取消' && it.vendor !== SELF_VENDOR && cats.indexOf(String(it.vendor || '').charAt(0).toUpperCase()) >= 0;
    });
  });
  if (all) {   // 指定抽查：列出這天全部，照時段、班級排，讓管理者自己勾
    const ses = function (t) { const i = SESSIONS.indexOf(sessionOf_(t)); return i < 0 ? 9 : i; };
    pool.sort(function (a, b) { return ses(a.classTime) - ses(b.classTime) || String(a.className).localeCompare(String(b.className), 'zh-Hant', { numeric: true }); });
    const ps = placesFor_(pool.map(function (o) { return o.orderNo; }));
    pool.forEach(function (o) { o.place = ps[o.orderNo] || null; });
    return { total: pool.length, orders: pool };
  }
  for (let i = pool.length - 1; i > 0; i--) {   // 洗牌
    const j = Math.floor(Math.random() * (i + 1)), t = pool[i];
    pool[i] = pool[j]; pool[j] = t;
  }
  const picked = pool.slice(0, n), places = placesFor_(picked.map(function (o) { return o.orderNo; }));
  picked.forEach(function (o) { o.place = places[o.orderNo] || null; });
  return { total: pool.length, orders: picked };
}

/** 管理者：存抽查結果。照片存到雲端硬碟「三信家商餐飲科抽查照片」，表上放連結 */
function adminSaveAudit(token, a) {
  const who = auth_(token, true);
  a = a || {};
  const timing = AUDIT_TIMINGS.indexOf(a.timing) >= 0 ? a.timing : AUDIT_TIMINGS[0];
  const partner = clean_(a.partner, 30);
  const how = a.how === '指定' ? '指定' : '隨機';
  const list = (a.orders || []).filter(function (o) { return o && (o.items || []).length; });
  if (!list.length) throw new Error('沒有要存的抽查結果。');
  if (list.length > 10) throw new Error('一次最多存 10 班。');
  const now = new Date(), stamp = Utilities.formatDate(now, tz_(), 'yyyyMMdd-HHmmss');
  const rows = [], ids = [];
  list.forEach(function (o, k) {
    const id = 'AU' + stamp + '-' + (k + 1);
    ids.push(id);
    const urls = (o.photos || []).slice(0, 3).map(function (p, i) { return saveAuditPhoto_(p, id + '-' + (i + 1)); }).filter(Boolean);
    const seal = AUDIT_SEALS.indexOf(o.seal) >= 0 ? o.seal : '';
    o.items.slice(0, 80).forEach(function (it, i) {
      const actual = it.actual === '' || it.actual === null || it.actual === undefined ? '' : Number(it.actual);
      rows.push({
        '抽查編號': id, '抽查時間': now, '時間點': timing, '冰箱': clean_(o.fridge, 30), '訂單編號': clean_(o.orderNo, 30),
        '上課日期': clean_(o.useDate, 12), '上課時間': clean_(o.classTime, 20), '班級': clean_(o.className, 20), '科別': clean_(o.dept, 10),
        '老師': clean_(o.teacher, 20), '食材名稱': clean_(it.name, 40), '清單數量': Number(it.qty) || 0, '單位': clean_(it.unit, 10),
        '實際數量': isNaN(actual) ? '' : actual, '結果': AUDIT_RESULTS.indexOf(it.result) >= 0 ? it.result : '不確定',
        '封條': i === 0 ? seal : '', '說明': clean_([i === 0 ? o.note : '', it.note].filter(Boolean).join('；'), 200),
        '照片': i === 0 ? urls.join('\n') : '', '抽查人': who, '陪同人': partner, '抽法': how,
      });
    });
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(ensureSheet_(SHEET.audits, AUDIT_HEADERS), AUDIT_HEADERS);
    appendRows_(SHEET.audits, rows, { '抽查時間': 'yyyy/mm/dd hh:mm' });
  } finally {
    lock.releaseLock();
  }
  return { ids: ids, rows: rows.length };
}

function saveAuditPhoto_(dataUrl, name) {
  const m = String(dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + (m[1] === 'image/png' ? '.png' : '.jpg'));
  if (blob.getBytes().length > 4 * 1024 * 1024) throw new Error('照片太大，請重拍。');
  return folderNamed_(AUDIT_FOLDER_NAME).createFile(blob).getUrl();
}

/** 管理者：抽查紀錄（照抽查編號合成一班一筆），from～to 是抽查日期 */
function adminGetAudits(token, from, to) {
  auth_(token, true);
  const start = parseDate_(from), end = parseDate_(to) || start;
  if (!start) throw new Error('請選擇日期。');
  const sh = ss_().getSheetByName(SHEET.audits);
  if (!sh || sh.getLastRow() < 2) return [];
  const endNext = addDays_(end, 1), by = {}, out = [];
  readTable_(SHEET.audits, ['抽查編號', '抽查時間']).rows.forEach(function (r) {
    const raw = r['抽查時間'], t = raw instanceof Date ? raw : toDate_(raw);
    if (!t || t < start || t >= endNext) return;
    const id = String(r['抽查編號']).trim();
    if (!by[id]) {
      by[id] = { id: id, time: fmtDateTime_(t), timing: String(r['時間點'] || ''), fridge: String(r['冰箱'] || ''), orderNo: String(r['訂單編號'] || ''),
        useDate: (toDate_(r['上課日期']) ? dateStr_(toDate_(r['上課日期'])) : String(r['上課日期'] || '')), classTime: String(r['上課時間'] || ''),
        className: String(r['班級'] || ''), dept: String(r['科別'] || ''), teacher: String(r['老師'] || ''), seal: '', note: '', photos: [],
        by: String(r['抽查人'] || ''), partner: String(r['陪同人'] || ''), how: String(r['抽法'] || '隨機'), items: [] };
      out.push(by[id]);
    }
    const g = by[id];
    if (r['封條']) g.seal = String(r['封條']);
    if (r['照片']) g.photos = g.photos.concat(String(r['照片']).split(/\s+/).filter(Boolean));
    g.items.push({ name: String(r['食材名稱'] || ''), qty: r['清單數量'], unit: String(r['單位'] || ''), actual: r['實際數量'],
      result: String(r['結果'] || ''), note: String(r['說明'] || '') });
  });
  return out.sort(function (a, b) { return a.time < b.time ? 1 : a.time > b.time ? -1 : 0; });
}

// ---- 校隊：放置登記 ----

/** 放置登記開放了沒（系統設定的開關，預設關閉；指令碼屬性 PLACE_ENABLED） */
function placeEnabled_() {
  return memo_('placeEnabled', function () { return PropertiesService.getScriptProperties().getProperty('PLACE_ENABLED') === '1'; });
}
function placeGate_() {
  if (!placeEnabled_()) throw new Error('放置登記目前還沒開放，請等管理者通知。');
}

/** 管理者：開放／關閉放置登記 */
function adminSetPlaceEnabled(token, on) {
  auth_(token, true);
  const props = PropertiesService.getScriptProperties();
  if (on) props.setProperty('PLACE_ENABLED', '1'); else props.deleteProperty('PLACE_ENABLED');
  memoClear_();
  return { placeEnabled: placeEnabled_() };
}

/** 每張訂單最後一次登記的放置位置：{ 訂單編號: { place, photo, by, time } } */
function placesFor_(orderNos) {
  const out = {};
  if (!placeEnabled_()) return out;   // 放置登記沒開放：老師、抽查都不顯示位置
  const sh = ss_().getSheetByName(SHEET.places);
  if (!sh || sh.getLastRow() < 2 || !orderNos.length) return out;
  const want = {};
  orderNos.forEach(function (n) { want[n] = true; });
  readTable_(SHEET.places, ['訂單編號', '放置位置']).rows.forEach(function (r) {
    const no = String(r['訂單編號']).trim();
    if (!want[no]) return;
    const t = r['登記時間'] instanceof Date ? r['登記時間'] : toDate_(r['登記時間']);
    const prev = out[no];
    if (prev && prev._t && t && t < prev._t) return;
    out[no] = { place: String(r['放置位置'] || ''), photo: String(r['照片'] || ''), by: String(r['登記人'] || ''), time: fmtDateTime_(t || r['登記時間']), _t: t };
  });
  Object.keys(out).forEach(function (k) { delete out[k]._t; });
  return out;
}

// 校隊練習、假日實習：老師照平常訂貨，在班級或實作項目寫「校隊」「假日」
function placeKind_(o) {
  const t = String(o.className || '') + ' ' + String(o.project || '');
  return /校隊/.test(t) ? 'team' : /假日/.test(t) ? 'holiday' : '';
}

/**
 * 校隊（或管理者）：要冰的食材（A～D）的訂單（不含金額），和目前登記的位置。
 * kind 空白 = date 這一天的一般課（不含校隊練習、假日實習）；'team' / 'holiday' = 今天起兩週內的校隊練習／假日實習
 */
function crewGetOrders(token, date, kind) {
  const account = auth_(token, false, false, true);
  if (!isAdmin_(account) && !isCrew_(account)) throw new Error('只有校隊和管理者可以使用放置登記。');
  placeGate_();
  kind = kind === 'team' || kind === 'holiday' ? kind : '';
  const from = kind ? today_() : parseDate_(date);
  if (!from) throw new Error('請選擇日期。');
  const to = kind ? addDays_(from, 14) : from;
  const rows = withHistory_(readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows, from).filter(function (r) {
    const x = toDate_(r['使用日期']);
    return x && x >= from && x <= to;
  });
  const list = groupOrders_(rows, null).filter(function (o) {
    return o.status !== '已取消' && !isReturn_(o.project) && placeKind_(o) === kind;
  }).map(function (o) {
    const items = o.items.filter(function (it) { return it.id !== TOOL_ID && !it.noOrder && it.status !== '已取消' && PLACE_VENDORS.test(String(it.vendor || '')); })
      .map(function (it) { return { name: it.name, qty: it.qty, unit: it.unit, vendor: String(it.vendor || '').charAt(0) }; });
    return { orderNo: o.orderNo, useDate: o.useDate, classTime: o.classTime, className: o.className, dept: o.dept, room: o.room,
      project: o.project, teacher: o.onBehalfOf || o.teacher, items: items };
  }).filter(function (o) { return o.items.length; });
  const places = placesFor_(list.map(function (o) { return o.orderNo; }));
  list.forEach(function (o) { o.place = places[o.orderNo] || null; });
  // 最近用過的位置（給校隊快速點選）
  let recent = [];
  const sh = ss_().getSheetByName(SHEET.places);
  if (sh && sh.getLastRow() >= 2) {
    const all = readTable_(SHEET.places, ['放置位置']).rows.map(function (r) { return String(r['放置位置'] || '').trim(); }).filter(Boolean).reverse();
    all.forEach(function (p) { if (recent.indexOf(p) < 0 && recent.length < 12) recent.push(p); });
  }
  return { orders: list.sort(function (a, b) {
    return String(a.useDate).localeCompare(String(b.useDate)) || String(a.classTime).localeCompare(String(b.classTime)) || String(a.className).localeCompare(String(b.className), 'zh-Hant', { numeric: true });
  }), recent: recent };
}

/** 校隊（或管理者）：登記這張訂單的東西放在哪裡，可以附一張照片 */
function crewSetPlace(token, orderNo, place, photo) {
  const account = auth_(token, false, false, true);
  if (!isAdmin_(account) && !isCrew_(account)) throw new Error('只有校隊和管理者可以使用放置登記。');
  placeGate_();
  orderNo = clean_(orderNo, 30); place = clean_(place, 40);
  if (!orderNo) throw new Error('缺少訂單編號。');
  if (!place) throw new Error('請填放在哪裡，例如：西餐 B 冰箱中層。');
  const r = readTable_(SHEET.orders, ['訂單編號']).rows.find(function (x) { return String(x['訂單編號']).trim() === orderNo; });
  if (!r) throw new Error('找不到這張訂單。');
  let url = '';
  if (photo) {
    const m = String(photo).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
    if (m) {
      const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], orderNo + '-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss') + '.jpg');
      if (blob.getBytes().length > 4 * 1024 * 1024) throw new Error('照片太大，請重拍。');
      url = folderNamed_(PLACE_FOLDER_NAME).createFile(blob).getUrl();
    }
  }
  const d = toDate_(r['使用日期']);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(ensureSheet_(SHEET.places, PLACE_HEADERS), PLACE_HEADERS);
    appendRows_(SHEET.places, [{ '訂單編號': orderNo, '上課日期': d ? dateStr_(d) : '', '班級': String(r['上課班級'] || ''),
      '老師': String(r['代訂對象老師'] || '').trim() || String(r['操作老師'] || ''), '放置位置': place, '照片': url, '登記人': account, '登記時間': new Date() }],
      { '登記時間': 'yyyy/mm/dd hh:mm' });
  } finally {
    lock.releaseLock();
  }
  return { place: place, photo: url, by: account, time: fmtDateTime_(new Date()) };
}

/** 管理者：校隊帳號清單 */
function adminListCrew(token) {
  auth_(token, true);
  const t = readTable_(SHEET.accounts, ['帳號', '啟用']);
  if (t.cols['校隊'] === undefined) return [];
  return t.rows.filter(function (r) { return isTrue_(r['校隊']); }).map(function (r) {
    return { account: String(r['帳號']).trim(), enabled: isTrue_(r['啟用']) };
  });
}

/** 管理者：新增校隊帳號（預設密碼 0000，第一次登入要改密碼）、停用／啟用、重設密碼 */
function adminSetCrew(token, account, action) {
  auth_(token, true);
  account = clean_(account, 30);
  if (!account) throw new Error('請輸入校隊的名字。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(sheet_(SHEET.accounts), ['校隊']);
    const t = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']);
    const r = t.rows.find(function (x) { return String(x['帳號']).trim() === account; });
    if (action === 'add') {
      if (r && !isTrue_(r['校隊'])) throw new Error('「' + account + '」已經是老師的帳號，請換一個名字（例如加上「校隊」）。');
      if (r) setCells_(t, r._row, { '啟用': true });
      else appendRows_(SHEET.accounts, [{ '帳號': account, '密碼': DEFAULT_PASSWORD, '啟用': true, '校隊': true }]);
      if (r) setCell_(t, r._row, '校隊', true);
    } else {
      if (!r || !isTrue_(r['校隊'])) throw new Error('找不到校隊帳號「' + account + '」。');
      if (action === 'disable') setCell_(t, r._row, '啟用', false);
      else if (action === 'enable') setCell_(t, r._row, '啟用', true);
      else if (action === 'reset') {
        const cell = t.sheet.getRange(r._row, t.cols['密碼']);
        cell.setNumberFormat('@');
        cell.setValue(DEFAULT_PASSWORD);
      } else throw new Error('不明的動作。');
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove('crewAccounts');
  try { CacheService.getScriptCache().remove(BOOT_CACHE_KEY); } catch (e) {}
  return adminListCrew(token);
}

function ensureSheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { CacheService.getScriptCache().remove(BOOT_CACHE_KEY); } catch (e) {}   // 「缺少分頁」提醒要重算
  }
  return sh;
}

function readMyLessonRows_() {
  return memo_('readMyLessonRows_', readMyLessonRowsRaw_).map(function (o) { return Object.assign({}, o); });
}
function readMyLessonRowsRaw_() {
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

/** 放假日：[{ row, from, to, name }]（日期是 yyyy-MM-dd，一次執行只讀一次） */
let holidayCache_ = null;
function readHolidays_() {
  if (holidayCache_) return holidayCache_;
  const sh = ss_().getSheetByName(SHEET.holidays);
  holidayCache_ = !sh || sh.getLastRow() < 2 ? [] : readTable_(SHEET.holidays, ['開始日期']).rows.map(function (r) {
    const a = toDate_(r['開始日期']), b = toDate_(r['結束日期']) || a;
    return a ? { row: r._row, from: dateStr_(a), to: dateStr_(b < a ? a : b), name: String(r['名稱'] || '').trim() || '放假' } : null;
  }).filter(Boolean).sort(function (x, y) { return x.from < y.from ? -1 : 1; });
  return holidayCache_;
}
/** 這天放假的話回傳名稱，不然回傳空字串 */
function holidayOf_(ds) {
  const h = readHolidays_().find(function (x) { return ds >= x.from && ds <= x.to; });
  return h ? h.name : '';
}
/** from～to 之間的放假日（給畫面顯示） */
function holidaysBetween_(from, to) {
  const a = dateStr_(from), b = dateStr_(to);
  return readHolidays_().filter(function (x) { return x.to >= a && x.from <= b; });
}

/** 管理者：from～to 還沒取消的訂單（颱風假、臨時停課要一起取消的），給畫面預覽 */
function adminHolidayOrders(token, from, to) {
  auth_(token, true);
  const a = parseDate_(from), b = parseDate_(to) || a;
  if (!a) return [];
  const map = {};
  readTable_(SHEET.orders, ['訂單編號', '使用日期', '狀態']).rows.forEach(function (r) {
    const d = toDate_(r['使用日期']);
    if (!d || d < a || d > b || String(r['狀態']).trim() === '已取消') return;
    const no = String(r['訂單編號']).trim();
    map[no] = map[no] || { orderNo: no, date: dateStr_(d), className: String(r['上課班級'] || ''), teacher: String(r['代訂對象老師'] || r['操作老師'] || ''),
      project: String(r['實作項目'] || ''), status: String(r['狀態'] || '') };
  });
  return Object.keys(map).map(function (k) { return map[k]; }).sort(function (x, y) { return x.date < y.date ? -1 : x.date > y.date ? 1 : 0; });
}

/** 管理者：新增放假日（from～to，名稱），回傳全部放假日。cancel = true：這幾天的訂單一起取消（颱風假、臨時停課） */
function adminAddHoliday(token, from, to, name, cancel) {
  const who = auth_(token, true);
  const a = parseDate_(from), b = parseDate_(to) || a;
  if (!a) throw new Error('請選擇放假日期。');
  if (b < a) throw new Error('結束日期不能早於開始日期。');
  const n = clean_(name, 30);
  if (!n) throw new Error('請填寫放假名稱，例如：國慶日、校慶補假。');
  let canceled = 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.holidays, HOLIDAY_HEADERS);
    appendRows_(SHEET.holidays, [{ '開始日期': a, '結束日期': b, '名稱': n }], { '開始日期': 'yyyy-mm-dd', '結束日期': 'yyyy-mm-dd' });
    if (cancel) {
      const t = readTable_(SHEET.orders, ['訂單編號', '使用日期', '狀態']), nos = {};
      t.rows.forEach(function (r) {
        const d = toDate_(r['使用日期']);
        if (!d || d < a || d > b || String(r['狀態']).trim() === '已取消') return;
        setCell_(t, r._row, '狀態', '已取消');
        nos[String(r['訂單編號']).trim()] = true;
      });
      Object.keys(nos).forEach(function (no) { logOrderEdit_(no, who, ['整張取消（' + n + '停課）'], n); });
      SpreadsheetApp.flush();
      canceled = Object.keys(nos).length;
    }
  } finally {
    lock.releaseLock();
  }
  holidayCache_ = null;
  const list = readHolidays_();
  list.canceled = canceled;
  return cancel ? { list: list, canceled: canceled } : list;
}

/** 管理者：刪除放假日（expect 是畫面上看到的開始日期，避免刪錯列） */
function adminDeleteHoliday(token, row, expectFrom) {
  auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName(SHEET.holidays);
    const t = readTable_(SHEET.holidays, ['開始日期']);
    const r = t.rows.find(function (x) { return x._row === Number(row); });
    if (!r || !toDate_(r['開始日期']) || dateStr_(toDate_(r['開始日期'])) !== expectFrom) throw new Error('放假日已經被改過，請重新整理再試一次。');
    deleteRows_(sh, [r._row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  holidayCache_ = null;
  return readHolidays_();
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
      if ((x.from && ds < x.from) || (x.to && ds > x.to)) return;   // 彈性課程：只有開始～結束日期那幾週
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
  const st = lessonStatusMaps_(from, to), have = st.have, skip = st.skip, manual = st.manual;
  // 同一位老師同一天同時段有兩堂課時，只要一張訂單（也只算一次）
  const seen = {};
  return lessons.filter(function (x) {
    const k = x.teacher + '|' + x.date + '|' + x.session + '|' + x.course;
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  }).map(function (x) {
    const k = x.teacher + '|' + x.date + '|' + x.session;
    const hol = holidayOf_(x.date);
    x.status = have[k] || manual[k] ? 'ordered' : skip[k] ? 'skipped' : hol ? 'holiday' : 'todo';
    if (have[k]) x.orders = have[k].map(function (o) { return { project: o.project, status: o.status, orderNo: o.orderNo }; });   // 我的課顯示訂了什麼（例如巧克力餅乾）
    if (hol) x.holiday = hol;
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
  if (!isAdmin_(account)) throw new Error('只有管理者可以查看別的老師的課。');
  return as;
}

function getMyLessons(token, as) {
  const account = lessonAccount_(token, as);
  const now = new Date();
  const from = earliestClassDate_(now);
  const to = addDays_(today_(), MAX_DAYS_AHEAD);
  const firstWeekEnd = addDays_(from, 4);
  const deadline = addDays_(from, -7 + (lockRule_().wd - 1));
  const lessons = lessonsWithStatus_(from, to, account).sort(function (a, b) {
    return a.date === b.date ? SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session) : (a.date < b.date ? -1 : 1);
  });
  const todoNext = lessons.filter(function (x) { return x.status === 'todo' && parseDate_(x.date) <= firstWeekEnd; }).length;
  return {
    lessons: lessons,
    nextWeek: { from: dateStr_(from), to: dateStr_(firstWeekEnd), deadline: shortDateTw_(deadline) + ' ' + lockTime_(), todo: todoNext },
    school: readSchedule_().filter(function (x) { return x.teacher === account; }),
    mine: readMyLessonRows_().filter(function (x) { return x.teacher === account; }),
    holidays: holidaysBetween_(from, to),
    teacher: account,
    schoolRows: isAdmin_(auth_(token)) ? scheduleRowsOf_(account) : null,
    guojiRows: isAdmin_(auth_(token)) ? guojiRowsOf_(account) : null,
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
  return memo_('readGuoji_', readGuojiRaw_).map(function (o) { return Object.assign({}, o); });
}
function readGuojiRaw_() {
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
    kind: scheduleKind_(r['類別']),
    from: toDate_(r['開始日期']) ? dateStr_(toDate_(r['開始日期'])) : '',
    to: toDate_(r['結束日期']) ? dateStr_(toDate_(r['結束日期'])) : '',
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
  const out = { '星期': weekday, '時段': session, '教室': clean_(f.room, 20), '課程': course, '節次': clean_(f.period, 10),
    '班級': clean_(f.className, 30), '老師': teacher };
  if (f.kind !== undefined) {
    const kind = String(f.kind || '').trim() || SCHEDULE_KINDS[0];
    if (SCHEDULE_KINDS.indexOf(kind) < 0) throw new Error('課表類別不正確。');
    out['類別'] = kind;
  }
  if (f.from !== undefined || f.to !== undefined) {
    const a = f.from ? parseDate_(f.from) : null, b = f.to ? parseDate_(f.to) : null;
    if ((f.from && !a) || (f.to && !b)) throw new Error('開始或結束日期不正確。');
    if (a && b && b < a) throw new Error('結束日期不能早於開始日期。');
    out['開始日期'] = a || '';
    out['結束日期'] = b || '';
  }
  return out;
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
    ensureHeaders_(sheet_(SHEET.schedule), SCHEDULE_HEADERS);
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

/** 管理者：總課表（「課表」分頁全部的課，含已取消的），依教室 × 星期 × 時段排成格子 */
function adminGetSchedule(token, weekOf) {
  auth_(token, true);
  const sh = ss_().getSheetByName(SHEET.schedule);
  const rows = !sh || sh.getLastRow() < 2 ? [] : readTable_(SHEET.schedule, ['星期', '時段', '老師']).rows.map(scheduleRowOut_);
  const accounts = {};
  readTable_(SHEET.accounts, ['帳號']).rows.forEach(function (r) { accounts[String(r['帳號']).trim()] = true; });
  rows.forEach(function (x) { x.noAccount = !!x.teacher && !accounts[x.teacher]; });
  const out = { rows: rows };
  const base = parseDate_(weekOf);
  if (!base) return out;
  // 這一週（週一～週五）每堂課的叫料狀態，再加上國技排班、老師自己新增的課
  const mon = addDays_(base, -((base.getDay() + 6) % 7)), fri = addDays_(mon, 4);
  const dates = [0, 1, 2, 3, 4].map(function (i) { return dateStr_(addDays_(mon, i)); });
  const st = lessonStatusMaps_(mon, fri);
  const status = function (teacher, date, session) {
    const k = teacher + '|' + date + '|' + session;
    const hol = holidayOf_(date);
    const r = st.have[k] ? { status: 'ordered', orders: st.have[k] } : st.manual[k] ? { status: 'ordered', manual: st.manual[k] } : st.skip[k] ? { status: 'skipped' } : hol ? { status: 'holiday', holiday: hol } : { status: 'todo' };
    if (st.manual[k] || st.skip[k]) r.marked = true;   // 有手動標記（可以改回自動）
    return r;
  };
  rows.forEach(function (x) {
    if (!x.need || x.skipped) return;
    const i = '一二三四五'.indexOf(x.weekday);
    if (i < 0) return;
    if ((x.from && dates[i] < x.from) || (x.to && dates[i] > x.to)) { x.outOfWeek = true; return; }
    const r = status(x.teacher, dates[i], x.session);
    x.status = r.status;
    x.orders = r.orders || [];
    x.manual = r.manual || null;
    x.marked = !!r.marked;
    x.holiday = r.holiday || '';
    x.date = dates[i];
  });
  out.extras = lessonsBetween_(mon, fri, null).filter(function (x) { return x.source !== '課表'; }).map(function (x) {
    const r = status(x.teacher, x.date, x.session);
    return { source: x.source, teacher: x.teacher, weekday: '一二三四五'.charAt(dates.indexOf(x.date)), session: x.session, course: x.course,
      className: x.className, room: x.room, status: r.status, orders: r.orders || [], noAccount: !!x.teacher && !accounts[x.teacher],
      manual: r.manual || null, marked: !!r.marked, date: x.date, holiday: r.holiday || '',
      kind: x.source === '自訂' && x.session === '晚上' ? '夜間部' : '日間部' };
  });
  out.week = { from: dates[0], to: dates[4], dates: dates, holidays: dates.map(holidayOf_) };
  out.holidays = readHolidays_();
  out.weekOrders = st.list;
  return out;
}

/**
 * 某段期間的訂單和標記：have[老師|日期|時段] = [{ orderNo, project, status }]、
 * skip[...] = true（不用訂貨）、manual[...] = { note, by }（管理者手動標「已叫料」）
 */
// 還料單（還別的老師借的材料，不是上課用）：實作項目開頭是【還料】，不對課表、不算這堂課已叫料
const RETURN_TAG = '【還料】';
function isReturn_(project) { return String(project || '').indexOf(RETURN_TAG) === 0; }

function lessonStatusMaps_(from, to) {
  const have = {}, skip = {}, manual = {};
  const rows = withHistory_(readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows, from).filter(function (r) {
    const d = toDate_(r['使用日期']);
    return d && d >= from && d <= to;
  });
  const list = [];
  groupOrders_(rows, null).forEach(function (o) {
    if (o.status === '已取消' || isReturn_(o.project)) return;
    list.push({ teacher: o.onBehalfOf || o.teacher, date: o.useDate, session: sessionOf_(o.classTime), project: o.project, className: o.className, orderNo: o.orderNo });
    const k = (o.onBehalfOf || o.teacher) + '|' + o.useDate + '|' + sessionOf_(o.classTime);
    (have[k] = have[k] || []).push({ orderNo: o.orderNo, project: o.project, status: o.status, className: o.className });
  });
  const sh = ss_().getSheetByName(SHEET.skips);
  if (sh && sh.getLastRow() >= 2) {
    readTable_(SHEET.skips, ['老師', '日期', '時段']).rows.forEach(function (r) {
      const d = toDate_(r['日期']);
      if (!d) return;
      const k = String(r['老師']).trim() + '|' + dateStr_(d) + '|' + String(r['時段']).trim();
      if (String(r['標記'] || '').trim() === '已叫料') manual[k] = { note: String(r['備註'] || '').trim(), by: String(r['標記人'] || '').trim() };
      else skip[k] = true;
    });
  }
  return { have: have, skip: skip, manual: manual, list: list };
}

/**
 * 管理者在總課表手動改某一週某堂課的叫料狀態（例如老師訂單日期填錯，系統對不起來）。
 * mark：'auto'（依訂單自動判斷）／'ordered'（已叫料）／'skipped'（不用訂貨）
 */
function adminSetLessonMark(token, teacher, date, session, course, mark, note, weekOf) {
  const admin = auth_(token, true);
  const t0 = clean_(teacher, 30), d = parseDate_(date);
  if (!t0 || !d || SESSIONS.indexOf(session) < 0) throw new Error('資料不正確，請重新整理。');
  if (['auto', 'ordered', 'skipped'].indexOf(mark) < 0) throw new Error('狀態不正確。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SHEET.skips, SKIP_HEADERS);
    ensureHeaders_(sh, SKIP_HEADERS);
    const t = readTable_(SHEET.skips, ['老師', '日期', '時段']);
    const same = t.rows.filter(function (r) {
      const rd = toDate_(r['日期']);
      return String(r['老師']).trim() === t0 && rd && dateStr_(rd) === dateStr_(d) && String(r['時段']).trim() === session;
    });
    if (same.length) deleteRows_(sh, same.map(function (r) { return r._row; }));
    if (mark !== 'auto') {
      appendRows_(SHEET.skips, [{ '老師': t0, '日期': d, '時段': session, '課程': clean_(course, 40), '標記時間': new Date(),
        '標記': mark === 'ordered' ? '已叫料' : '不用訂貨', '備註': clean_(note, 100), '標記人': admin }],
        { '日期': 'yyyy-mm-dd', '標記時間': 'yyyy/mm/dd hh:mm' });
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return adminGetSchedule(token, weekOf);
}

/**
 * 管理者在總課表上改一堂課：fields 是新的內容，fields.need = false 是永久取消、true 是上課中。
 * expect 跟 adminUpdateSchedule 一樣，別人改過就不改。
 */
function adminGridSchedule(token, row, expect, fields, weekOf) {
  const admin = auth_(token, true);
  const set = scheduleFields_(fields);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_(sheet_(SHEET.schedule), SCHEDULE_HEADERS);
    const t = readTable_(SHEET.schedule, ['星期', '時段', '老師']);
    const r = t.rows.find(function (x) { return x._row === Number(row); });
    const cur = r ? scheduleRowOut_(r) : null;
    if (!cur || !expect || cur.weekday !== expect.weekday || cur.session !== expect.session || cur.teacher !== expect.teacher) {
      throw new Error('課表已經被改過，請重新整理再試一次。');
    }
    let note = fields.note === undefined ? cur.note : clean_(fields.note, 100);
    const need = fields.need === undefined ? cur.need : !!fields.need;
    if (cur.need && !need) {
      note = (admin + ' ' + Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd') + ' 取消' + (note ? '；' + note : '')).slice(0, 100);
    } else if (!cur.need && need) {
      note = note.replace(/^\S+ \d{4}\/\d{2}\/\d{2} 取消；?/, '');
    }
    set['需要訂貨'] = need;
    set['備註'] = note;
    Object.keys(set).forEach(function (k) {
      if (t.cols[k] === undefined) throw new Error('「課表」分頁缺少「' + k + '」欄位。');
      setCell_(t, r._row, k, set[k]);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return adminGetSchedule(token, weekOf);
}

/** 管理者在總課表的空格新增一堂課，回傳新的總課表 */
function adminGridAddSchedule(token, fields, weekOf) {
  auth_(token, true);
  addScheduleRow_(fields);
  return adminGetSchedule(token, weekOf);
}

/** 管理者在「課表」分頁新增一堂課 */
function adminAddSchedule(token, fields, as) {
  auth_(token, true);
  addScheduleRow_(fields);
  return getMyLessons(token, as);
}

function addScheduleRow_(fields) {
  const set = scheduleFields_(fields);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SHEET.schedule, SCHEDULE_HEADERS);
    ensureHeaders_(sh, SCHEDULE_HEADERS);
    set['需要訂貨'] = true;
    set['備註'] = clean_(fields && fields.note, 100);
    appendRows_(SHEET.schedule, [set], { '開始日期': 'yyyy-mm-dd', '結束日期': 'yyyy-mm-dd' });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
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
  const deadline = addDays_(from, -7 + (lockRule_().wd - 1));
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
  const dl = shortDateTw_(deadline) + ' ' + lockTime_();
  const link = lineUrl_('lessons');
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
  const url = lineUrl_('');
  if (/截止|規則/.test(t) && t.length <= 10) {
    const now = new Date(), from = earliestClassDate_(now), dl = currentDeadline_(now) > now ? currentDeadline_(now) : addDays_(currentDeadline_(now), 7);
    return '【訂貨規則】' + lockRuleText_() + '。\n' +
      '下一個截止：' + shortDateTw_(dl) + ' ' + lockTime_() + '\n' +
      '目前最早可以訂：' + shortDateTw_(from) + ' 的課，最晚可以訂 ' + MAX_DAYS_AHEAD + ' 天內的課。' + (url ? '\n' + url : '');
  }
  if (/網站|訂貨系統/.test(t) && t.length <= 10) return url ? '訂貨網站：' + url : '訂貨網站還沒部署。';
  if (/教學|影片|怎麼用|使用說明|操作說明/.test(t) && t.length <= 10) {
    return url ? '▶ 教學影片（不用登入就可以看，用瀏覽器打開）：\n' + lineUrl_('help') : '訂貨網站還沒部署。';
  }
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
    }).join('\n') + (todo.length > 20 ? '\n…' : '') + (url ? '\n\n去訂貨：' + lineUrl_('lessons') : '');
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

/** 管理者：建立官方帳號圖文選單（6 個按鈕，上下兩排）並設為預設 */
function adminSetupRichMenu(token) {
  auth_(token, true);
  const props = PropertiesService.getScriptProperties();
  const lineToken = props.getProperty('LINE_TOKEN');
  if (!lineToken) throw new Error('請先儲存 Channel access token。');
  const url = webAppUrl_();
  if (!url) throw new Error('請先部署網站，再回來設定選單。');
  const auth = { Authorization: 'Bearer ' + lineToken };
  const w = 2500, h = 1686, cw = Math.floor(w / 3), rh = h / 2;
  const areas = [
    { action: { type: 'message', label: '我的未訂貨課程', text: '我的未訂貨課程' } },
    { action: { type: 'uri', label: '打開訂貨網站', uri: lineUrl_('') } },
    { action: { type: 'message', label: '我的訂單', text: '我的訂單' } },
    { action: { type: 'message', label: '截止時間說明', text: '截止時間說明' } },
    { action: { type: 'uri', label: '教學影片', uri: lineUrl_('help') } },
    { action: { type: 'uri', label: '我的課表', uri: lineUrl_('lessons') } },
  ].map(function (a, i) {
    const c = i % 3, r = Math.floor(i / 3);
    a.bounds = { x: c * cw, y: r * rh, width: c === 2 ? w - 2 * cw : cw, height: rh };
    return a;
  });
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

const GRAMS_ = { 'g': 1, '公斤': 1000, '斤': 600, '兩': 37.5, 'T': 15, 't': 5 };   // 1 大T＝15 g、1 小T＝5 g（跟網頁的 GRAMS 一樣）
// 合計時計數單位同一種的算在一起（個／顆／粒→顆，根／支／條→條）
const COUNT_UNITS_ = { '個': '顆', '粒': '顆', '顆': '顆', '根': '條', '支': '條', '條': '條' };
// 給廠商的合計：菜商、肉商的重量用「兩」加總、顯示成「1 斤 4 兩」；雞蛋用克訂的換成顆（1 顆 50 克）；其他廠商換成公斤
const EGG_GRAMS_ = 50;
function isEgg_(name) { return String(name || '').trim() === '雞蛋'; }
function convertUnit_(qty, unit, vendor, name) {
  if (GRAMS_[unit] && isEgg_(name)) return { qty: Math.round(qty * GRAMS_[unit] / EGG_GRAMS_ * 100) / 100, unit: '顆' };
  if (!GRAMS_[unit]) return { qty: qty, unit: COUNT_UNITS_[unit] || unit };
  const target = /^[AB]/.test(vendor || '') ? '兩' : '公斤';
  return { qty: Math.round(qty * GRAMS_[unit] / GRAMS_[target] * 100) / 100, unit: target };
}
/** 給廠商的明細（單筆）：菜商、肉商的重量換成斤兩、雞蛋的克換成顆，其他照老師原本的單位 */
function vendorLineQty_(qty, unit, vendor, name) {
  const cv = convertUnit_(Number(qty), unit, vendor, name);
  if (cv.unit === '兩' || (cv.unit === '顆' && isEgg_(name) && GRAMS_[unit])) return totalText_(cv.qty, cv.unit, name);
  return { qty: qty, unit: unit };
}
/** 合計顯示：{ qty, unit }；兩無條件進位到整兩、滿 16 兩進成斤，雞蛋進位到整顆 */
function totalText_(qty, unit, name) {
  if (unit === '兩') {
    const n = Math.ceil(qty - 1e-6), jin = Math.floor(n / 16), liang = n % 16;
    return jin ? { qty: jin + '斤' + (liang ? liang + '兩' : ''), unit: liang ? '斤兩' : '斤' } : { qty: String(n), unit: '兩' };
  }
  if (unit === '顆' && isEgg_(name)) return { qty: String(Math.ceil(qty - 1e-6)), unit: '顆' };
  return { qty: qty, unit: unit };
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
      const cv = convertUnit_(it.qty, it.unit, it.vendor, it.name);
      const k = it.name + '|' + cv.unit;
      v.totals[k] = v.totals[k] || { name: it.name, unit: cv.unit, qty: 0 };
      v.totals[k].qty = Math.round((v.totals[k].qty + cv.qty) * 100) / 100;
      const lq = vendorLineQty_(it.qty, it.unit, it.vendor, it.name);
      const line = [o.useDate, it.group, o.className, it.name, lq.qty, lq.unit, o.onBehalfOf || o.teacher, it.note];
      line.block = [o.useDate, o.className, o.orderNo, it.group].join('|');   // 同一張訂單、同一送貨組別 = 明細合併成一塊
      v.lines.push(line);
    });
  });
  Object.keys(by).forEach(function (v) {
    by[v].totals = Object.keys(by[v].totals).sort().map(function (k) { return by[v].totals[k]; });
    by[v].lines.sort(function (a, b) { return a.block === b.block ? String(a[3]).localeCompare(String(b[3])) : (a.block < b.block ? -1 : 1); });
  });
  return by;
}

/** 明細：同一張訂單的上課日期、組別、班級、老師合併成一格，訂單和訂單之間畫粗線 */
function mergeDetailBlocks_(sh, firstRow, lines) {
  let start = 0;
  for (let i = 1; i <= lines.length; i++) {
    if (i < lines.length && lines[i].block === lines[start].block) continue;
    const row = firstRow + start, n = i - start;
    sh.getRange(row, 1, 1, 8).setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    if (n > 1) [1, 2, 3, 7].forEach(function (c) { sh.getRange(row, c, n, 1).merge(); });
    start = i;
  }
  if (lines.length) sh.getRange(firstRow, 1, lines.length, 8).setVerticalAlignment('middle');
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
    d.totals.forEach(function (t) { const x = totalText_(t.qty, t.unit, t.name); rows.push([t.name, x.unit, x.qty, '', '', '', '', '']); });
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
    mergeDetailBlocks_(sh, detailHead + 1, d.lines);
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
    (lineUrl_('') ? '\n\n' + lineUrl_('') : '');
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
  const n = notify_('【訂貨每日摘要】' + day + '：新訂單 ' + newOrders.length + ' 張' + (edits.length ? '、修改 ' + edits.length + ' 張' : '') + (errors.length ? '、系統錯誤 ' + errors.length + ' 筆' : ''), html + editHtml + errHtml, null);   // 只寄 Email，不傳 LINE（省則數）
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
  const selfIds = rows.length ? selfItemIds_() : {};
  const vendorGroups = readVendorGroups_();
  // 舊訂單可能所有廠商都寫了送貨組別；廠商設定沒有分組別的廠商一律不分
  const groupOf = function (r) {
    const v = String(r['廠商'] || '').trim();
    return vendorGroups[v] ? String(r['送貨組別'] || '') : '';
  };
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
    const iid = String(r['食材ID'] || '').trim(), self = !!selfIds[iid];
    map[no].items.push({
      id: String(r['食材ID'] || ''),
      name: String(r['食材名稱'] || ''),
      vendor: self ? SELF_VENDOR : String(r['廠商'] || ''),
      group: groupOf(r),
      qty: Number(r['數量']) || 0,
      unit: String(r['單位'] || ''),
      note: String(r['備註'] || ''),
      status: String(r['狀態'] || ''),
      noOrder: isTrue_(r['不叫貨']) || self,
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

let SHEETS_ = {};   // 同一次執行裡分頁物件只找一次
function sheet_(name) {
  const sh = SHEETS_[name] || (SHEETS_[name] = ss_().getSheetByName(name));
  if (!sh) { delete SHEETS_[name]; throw new Error('找不到分頁「' + name + '」，請確認分頁名稱。'); }
  return sh;
}

// 課表、老師課表、國技排班：同一次執行讀一次就好（「我的課」一次要用好幾遍）；網站寫入任何分頁時清掉
let MEMO_ = {};
function memo_(key, fn) { return MEMO_.hasOwnProperty(key) ? MEMO_[key] : (MEMO_[key] = fn()); }
function memoClear_() { MEMO_ = {}; }

/** 讀整張表，回傳 { sheet, cols: {標題: 欄號(1起)}, rows: [{標題: 值, _row: 列號}] } */
function readTable_(name, required) {
  const sh = sheet_(name);
  // 一次讀完（以前是先問最後一列、最後一欄再讀，要連 Google 三次）
  const values = retry_(function () { return sh.getDataRange().getValues(); });
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
  bootDirty_(name);
  memoClear_();
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

/** 同一列改好幾格：相鄰的欄一次寫（少連幾次 Google） */
function setCells_(table, row, values) {
  const list = Object.keys(values).map(function (h) {
    const c = table.cols[h];
    if (c === undefined) throw new Error('分頁缺少「' + h + '」欄位。');
    return { c: c, v: values[h] };
  }).sort(function (a, b) { return a.c - b.c; });
  try { bootDirty_(table.sheet.getName()); } catch (e) {}
  memoClear_();
  for (let i = 0; i < list.length;) {
    let j = i;
    while (j + 1 < list.length && list[j + 1].c === list[j].c + 1) j++;
    table.sheet.getRange(row, list[i].c, 1, j - i + 1).setValues([list.slice(i, j + 1).map(function (x) { return x.v; })]);
    i = j + 1;
  }
}

function setCell_(table, row, header, value) {
  const c = table.cols[header];
  if (c === undefined) throw new Error('分頁缺少「' + header + '」欄位。');
  try { bootDirty_(table.sheet.getName()); } catch (e) {}
  memoClear_();
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

/** 本週（星期一開始）的訂貨截止時間，例如本週一 12:00（系統設定可以改）。 */
function currentDeadline_(now) {
  const today = today_(now);
  const monday = addDays_(today, -((today.getDay() + 6) % 7));
  const r = lockRule_();
  const d = addDays_(monday, r.wd - 1);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), r.h, r.m, 0);
}

/** 下一個訂貨截止：{ at: 毫秒, label: '10/19（一）12:00', from, to: 這次截止的是哪一週的課 } */
function nextDeadline_(now) {
  const r = lockRule_();
  let d = currentDeadline_(now);
  if (now >= d) d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 7, r.h, r.m, 0);
  const monday = addDays_(new Date(d.getFullYear(), d.getMonth(), d.getDate()), 8 - r.wd);
  return {
    at: d.getTime(), label: shortDateTw_(d) + ' ' + lockTime_(),
    from: dateStr_(monday), to: dateStr_(addDays_(monday, 4)),
  };
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
