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
const SCHEDULE_HEADERS = ['星期', '時段', '教室', '課程', '節次', '班級', '老師', '需要訂貨', '備註', '類別', '開始日期', '結束日期', '隔週'];
/** 隔週上課：「隔週」欄填有上課那週的週一（空白＝每週都上）。ds 那週跟它差奇數週＝這週不上 */
function altWeekOff_(anchor, ds) {
  if (!anchor) return false;
  const mon = function (s) { const d = parseDate_(s); return d ? addDays_(d, -((d.getDay() + 6) % 7)) : null; };
  const a = mon(anchor), b = mon(ds);
  if (!a || !b) return false;
  return Math.abs(Math.round((b - a) / (7 * 86400000))) % 2 === 1;
}
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
// 同義的單位合併成一種（個、粒→顆；根、支→條；附→付；cc→ml），選單只列合併後的
const ORDER_UNITS = ['斤', '兩', '公斤', 'g', '顆', '條', '把', '朵', '片', '塊', '隻', '尾', '付', '打', '包', '袋', '盒', '罐', '瓶', '桶', '箱', '公升'];
const RECIPE_UNITS = ['T', 't', '杯', 'ml'];
const UNIT_SYNONYMS = { '公克': 'g', '克': 'g', 'G': 'g', 'kg': '公斤', 'KG': '公斤', 'Kg': '公斤', '台斤': '斤', '大匙': 'T', '小匙': 't', 'CC': 'ml', 'cc': 'ml', '毫升': 'ml', 'L': '公升',
  '個': '顆', '粒': '顆', '根': '條', '支': '條', '附': '付', '副': '付' };
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
  if (CUR_STUDENT_) throw new Error('校隊學生用 LINE 登入的不能改組別帳號的密碼，請找管理者。');
  oldPassword = String(oldPassword == null ? '' : oldPassword);
  newPassword = String(newPassword == null ? '' : newPassword);
  if (/\s/.test(newPassword)) throw new Error('新密碼不能有空白。');
  if (newPassword.length < MIN_PASSWORD_LENGTH) throw new Error('新密碼至少要 ' + MIN_PASSWORD_LENGTH + ' 碼。');
  if (newPassword.length > 30) throw new Error('新密碼最多 30 碼。');
  if (newPassword === DEFAULT_PASSWORD) throw new Error('新密碼不能是預設密碼 ' + DEFAULT_PASSWORD + '。');
  if (newPassword === oldPassword) throw new Error('新密碼不能跟目前的密碼一樣。');

  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']);
    const r = t.rows.find(function (x) { return String(x['帳號']).trim() === account; });
    if (!r || !isTrue_(r['啟用'])) throw new Error('找不到你的帳號，請聯絡管理者。');
    const viaReset = !!CacheService.getScriptCache().get('pwresettok:' + token);   // LINE「忘記密碼」連結登入的：不用舊密碼
    if (!viaReset && String(r['密碼']) !== oldPassword) throw new Error('目前的密碼不正確。');
    if (viaReset && String(r['密碼']) === newPassword) throw new Error('新密碼不能跟原本的密碼一樣。');
    // 純文字格式，避免「0123」被試算表變成數字 123
    const cell = t.sheet.getRange(r._row, t.cols['密碼']);
    cell.setNumberFormat('@');
    cell.setValue(newPassword);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const wasReset = !!CacheService.getScriptCache().get('pwresettok:' + token);
  CacheService.getScriptCache().remove('mustchange:' + token);
  CacheService.getScriptCache().remove('pwresettok:' + token);
  logLogin_(account, wasReset ? '改密碼（LINE 忘記密碼）' : '改密碼');
  return true;
}

/** 登入並一次載入網站資料（少一次來回，登入比較快）。 */
function loginAndLoad(account, password) {
  const res = login(account, password);
  if (!res.mustChange) res.data = getBootstrap(res.token);   // 要先改密碼的，改完才載入資料
  return res;
}

/**
 * LINE 管理者選單的「補訂貨」「訂貨系統」：官方帳號回一個一次性的登入連結（10 分鐘內、只能用一次），
 * 打開就直接登入，不用打帳號密碼。連結只會傳給綁定 LINE 的管理者。
 */
const LINE_LOGIN_MINUTES = 10;
function lineLoginLink_(account, page) {
  const code = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  CacheService.getScriptCache().put('llink:' + code, account, LINE_LOGIN_MINUTES * 60);
  const u = lineUrl_(page || '');
  return u ? u + '&lt=' + code : '';
}
/** 驗收系統的一次性登入碼：兩個網站是不同的程式，透過試算表「LINE登入碼」分頁交給驗收系統（10 分鐘、用一次） */
function rcvLoginCode_(account) {
  const code = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  const lock = scriptLock_(); lock.waitLock(20000);
  try {
    const sh = ensureSheet_('LINE登入碼', ['碼', '帳號', '系統', '期限']);
    const t = readTable_('LINE登入碼', ['碼', '期限']), now = Date.now();
    const old = t.rows.filter(function (r) { return !(r['期限'] instanceof Date) || r['期限'].getTime() < now; }).map(function (r) { return r._row; });
    if (old.length) deleteRows_(sh, old);   // 過期的順便清掉
    appendRows_('LINE登入碼', [{ '碼': code, '帳號': account, '系統': '驗收', '期限': new Date(now + LINE_LOGIN_MINUTES * 60000) }]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return code;
}
function loginByLink(code) {
  code = String(code || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 64);
  const cache = CacheService.getScriptCache(), account = code ? cache.get('llink:' + code) : null;
  if (!account) throw new Error('這個登入連結已經過期或用過了，請在 LINE 再按一次，或用帳號密碼登入。');
  cache.remove('llink:' + code);
  if (!isAdmin_(account)) throw new Error('只有管理者可以用 LINE 連結登入。');
  const token = Utilities.getUuid();
  cache.put('session:' + token, account, SESSION_HOURS * 3600);
  try { logLogin_(account, '成功（LINE 連結）'); } catch (e) {}
  return { token: token, account: account, isAdmin: true, mustChange: false, data: getBootstrap(token) };
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
  const lock = scriptLock_();
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

// 校隊學生用 LINE 連結登入組別共用帳號（中餐校隊…）時，這次請求是哪個學生（訂單、修改紀錄記名字）
var CUR_STUDENT_ = '';
function auth_(token, needAdmin, allowMustChange, allowCrew) {
  const cache = CacheService.getScriptCache();
  const account = token ? cache.get('session:' + token) : null;
  if (!account) throw new Error('SESSION_EXPIRED');
  CUR_STUDENT_ = cache.get('sessstu:' + token) || '';
  // 還在用預設密碼的，除了改密碼以外什麼都不能做
  if (!allowMustChange && cache.get('mustchange:' + token)) throw new Error('MUST_CHANGE_PASSWORD');
  if (needAdmin && !isAdmin_(account)) throw new Error('只有管理者可以使用這個功能。');
  // 校隊帳號只能登記放置位置、改密碼，不能訂貨、看訂單
  // 「可訂貨」打勾的校隊可以幫老師代訂（送出時一定要選代訂對象老師）
  if (!needAdmin && !allowCrew && !isAdmin_(account) && isCrew_(account) && !crewCanOrder_(account)) throw new Error('校隊帳號只能使用「放置登記」。');
  return account;
}

/** 校隊帳號：「帳號設定」的「校隊」欄打勾、而且啟用 */
function isCrew_(account) {
  return !!account && crewAccounts_().indexOf(account) >= 0;
}
function crewAccounts_() { return crewInfo_().all; }
/** 校隊帳號而且「可訂貨」打勾：可以登入訂貨網幫老師代訂 */
function crewCanOrder_(account) { return !!account && crewInfo_().order.indexOf(account) >= 0; }
function crewInfo_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('crewAccounts2');
  if (hit) return JSON.parse(hit);
  const t = readTable_(SHEET.accounts, ['帳號']);
  const rows = t.cols['校隊'] === undefined ? [] : t.rows.filter(function (r) { return isTrue_(r['校隊']) && isTrue_(r['啟用']) && String(r['帳號']).trim(); });
  const info = {
    all: rows.map(function (r) { return String(r['帳號']).trim(); }),
    order: rows.filter(function (r) { return isTrue_(r['可訂貨']); }).map(function (r) { return String(r['帳號']).trim(); })
  };
  cache.put('crewAccounts2', JSON.stringify(info), 300);
  return info;
}

// =====================================================================
// 老師：讀取資料、下單、查詢、取消
// =====================================================================

/** 登入後一次載入網站需要的資料。 */
// 打開網站時要讀的、不常變的資料（老師名單、單位、科別替換、廠商聯絡、放假日）：暫存 5 分鐘，
// 網站上改到這幾張分頁時自動清掉（appendRows_／setCell_／deleteRows_）；直接在試算表改的，最慢 5 分鐘後生效
const BOOT_CACHE_KEY = 'boot:static:v1';
const BOOT_SHEETS = ['帳號設定', '單位', '科別替換', '廠商聯絡', '放假日', '班級人數', '單位換算'];
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
    classSizes: readClassSizes_(),
    unitConv: readUnitConv_(),
    unitHist: itemUnitHist_(),
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
      if (/^手動/.test(String(r['驗收編號'] || ''))) x.m = 1;   // 管理者在成本頁自己填的
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
  tblCacheDrop_(name);
  if (DATA_VER_SHEETS_.indexOf(String(name)) >= 0) dataBump_();
  if (BOOT_SHEETS.indexOf(String(name)) < 0) return;
  try { CacheService.getScriptCache().remove(BOOT_CACHE_KEY); } catch (e) {}
}

function getBootstrap(token) {
  fastReads_();
  const account = auth_(token, false, false, true);
  if (!isAdmin_(account) && isCrew_(account) && !crewCanOrder_(account)) return { account: account, isCrew: true, isAdmin: false, items: [], today: dateStr_(today_()), placeEnabled: placeEnabled_() };
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
    crewOrder: !isAdmin && crewCanOrder_(account),
    student: (token && CacheService.getScriptCache().get('sessstu:' + token)) || '',
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
    unitMerge: unitMerge_(),
    deptSubs: st.deptSubs,
    missingSetup: isAdmin ? st.missingSetup : [],
    vendorContacts: isAdmin ? st.vendorContacts : {},
    holidays: st.holidays,
    prices: st.prices || {},
    placeEnabled: placeEnabled_(),
    auditEnabled: PropertiesService.getScriptProperties().getProperty('AUDIT_ENABLED') === '1',
    classSizes: st.classSizes || {},
    unitConv: st.unitConv || {},
    unitHist: st.unitHist || {},
    notes: notesFor_(account),
  };
}

// ---- 單位換算（算料單成本用）：老師訂「把、支、包」，銷貨單是「斤」時，管理者填 1 把 ＝ 幾斤 ----
const UNIT_CONV_SHEET = '單位換算', UNIT_CONV_HEADERS = ['食材名稱', '單位', '等於', '換算單位', '更新時間', '更新人'];
/** { 食材名稱: [{ u: '把', q: 0.3, u2: '斤' }] }（1 u ＝ q u2） */
function readUnitConv_() {
  const out = {};
  if (!ss_().getSheetByName(UNIT_CONV_SHEET)) return out;
  readTable_(UNIT_CONV_SHEET, ['食材名稱', '單位', '等於', '換算單位']).rows.forEach(function (r) {
    const n = String(r['食材名稱']).trim(), u = String(r['單位']).trim(), q = Number(r['等於']), u2 = String(r['換算單位']).trim();
    if (n && u && u2 && q > 0) (out[n] = out[n] || []).push({ u: u, q: q, u2: u2 });
  });
  return out;
}
/**
 * 管理者：自己找到的單價直接填進單價表（例如廠商報價單、網路價格）。
 * 同一家同一樣同一單位只留一筆；price 空白＝刪掉手填的那筆（銷貨單記的不動）。
 * 之後那家廠商的銷貨單有這一樣，成本會自動改用銷貨單的單價。回傳最新的單價表（給畫面用）
 */
function adminSavePrice(token, vendor, name, unit, price, spec) {
  const admin = auth_(token, true);
  vendor = clean_(vendor, 30); name = clean_(name, 40); unit = clean_(unit, 10);
  // 規格（選填）：一個單位有多少，例如「1000 g」「25 包」，算成本時用來換老師訂的單位
  spec = String(spec == null ? '' : spec).replace(/\s+/g, ' ').trim().slice(0, 20);
  if (spec && !/^[\d.]+\s*[^\d\s.]+/.test(spec)) throw new Error('規格請寫「數字＋單位」，例如 1000 g、25 包。');
  if (spec) spec = spec.replace(/^([\d.]+)\s*/, '$1 ');
  if (!vendor || !name || !unit) throw new Error('資料不完整（廠商、品名、單位都要有）。');
  const raw = String(price == null ? '' : price).replace(/[,，$＄\s]/g, ''), p = Number(raw);
  if (raw !== '' && !(p > 0 && p < 1000000)) throw new Error('單價請填大於 0 的數字。');
  const H = ['廠商', '食材名稱', '單位', '單價', '銷貨單品名', '更新日期', '驗收編號', '貨品編號', '換算'];
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_('單價表', H);
    const t = readTable_('單價表', ['廠商', '食材名稱', '單價']);
    const same = t.rows.filter(function (r) { return String(r['廠商']).trim() === vendor && String(r['食材名稱']).trim() === name && String(r['單位'] || '').trim() === unit; });
    const tag = '手動（' + admin + '）', today = dateStr_(new Date());
    if (raw === '') {
      const del = same.filter(function (r) { return /^手動/.test(String(r['驗收編號'] || '')); }).map(function (r) { return r._row; });
      if (del.length) deleteRows_(sh, del);
    } else if (same.length) {
      const r = same[same.length - 1];
      setCell_(t, r._row, '單價', p); setCell_(t, r._row, '更新日期', today); setCell_(t, r._row, '驗收編號', tag);
      if (spec && t.cols['換算']) setCell_(t, r._row, '換算', spec);
    } else appendRows_('單價表', [{ '廠商': vendor, '食材名稱': name, '單位': unit, '單價': p, '更新日期': today, '驗收編號': tag, '換算': spec }]);
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove(BOOT_CACHE_KEY); } catch (e) {}
  } finally {
    lock.releaseLock();
  }
  return orderPrices_();
}

/** 管理者：存一筆換算（q 空白＝刪掉）。回傳全部換算 */
function adminSaveUnitConv(token, name, unit, q, unit2) {
  const admin = auth_(token, true);
  name = clean_(name, 40); unit = clean_(unit, 10); unit2 = clean_(unit2, 10);
  if (!name || !unit || !unit2) throw new Error('資料不完整。');
  const raw = String(q == null ? '' : q).trim(), n = Number(raw);
  if (raw !== '' && !(n > 0 && n < 100000)) throw new Error('請填大於 0 的數字。');
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(UNIT_CONV_SHEET, UNIT_CONV_HEADERS);
    const t = readTable_(UNIT_CONV_SHEET, ['食材名稱', '單位']);
    // 同一樣、同一個單位的重量換算只留一筆（改成「1 袋＝22 公斤」時，舊的「1 袋＝22 g」要一起刪掉，不然會用到舊的）
    const isW = function (u) { return !!UNIT_G_[String(u || '').trim()]; };
    const old = t.rows.filter(function (r) { const u2 = String(r['換算單位'] || '').trim(); return String(r['食材名稱']).trim() === name && String(r['單位']).trim() === unit && (u2 === unit2 || (isW(u2) && isW(unit2))); }).map(function (r) { return r._row; });
    if (old.length) deleteRows_(sh, old);
    if (raw !== '') appendRows_(UNIT_CONV_SHEET, [{ '食材名稱': name, '單位': unit, '等於': n, '換算單位': unit2, '更新時間': new Date(), '更新人': admin }], { '更新時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
    bootDirty_(UNIT_CONV_SHEET);
  } finally {
    lock.releaseLock();
  }
  return readUnitConv_();
}

/** 管理者：一次存好幾筆單位換算 list: [{ name, unit, q, unit2 }] */
function adminSaveUnitConvs(token, list) {
  const admin = auth_(token, true);
  list = (list || []).map(function (x) { return { n: clean_(x && x.name, 40), u: clean_(x && x.unit, 10), q: Number(x && x.q), u2: clean_(x && x.unit2, 10) }; })
    .filter(function (x) { return x.n && x.u && x.u2 && x.q > 0 && x.q < 100000; });
  if (!list.length) throw new Error('沒有要存的換算（數字要大於 0）。');
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(UNIT_CONV_SHEET, UNIT_CONV_HEADERS);
    const t = readTable_(UNIT_CONV_SHEET, ['食材名稱', '單位']);
    const wk = function (u2) { return UNIT_G_[u2] ? '#w' : u2; };   // 重量單位當同一種：新的「＝22 公斤」會蓋掉舊的「＝22 g」
    const want = {}; list.forEach(function (x) { want[x.n + '|' + x.u + '|' + wk(x.u2)] = true; });
    const old = t.rows.filter(function (r) { return want[String(r['食材名稱']).trim() + '|' + String(r['單位']).trim() + '|' + wk(String(r['換算單位'] || '').trim())]; }).map(function (r) { return r._row; });
    if (old.length) deleteRows_(sh, old);
    const now = new Date();
    appendRows_(UNIT_CONV_SHEET, list.map(function (x) { return { '食材名稱': x.n, '單位': x.u, '等於': x.q, '換算單位': x.u2, '更新時間': now, '更新人': admin }; }), { '更新時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
    bootDirty_(UNIT_CONV_SHEET);
  } finally {
    lock.releaseLock();
  }
  return readUnitConv_();
}

// ---- 班級人數：管理者設定，老師在我的課、訂貨頁、料單看得到（不能改） ----
const CLASS_SIZE_SHEET = '班級人數', CLASS_SIZE_HEADERS = ['班級', '人數', '備註', '更新時間', '更新人', '隱藏', '課程'];   // 隱藏：管理者刪掉的班級；課程：選修課同一班不同課程人數不同（空白＝整班）

function classSizeKey_(cls, course) { return String(cls || '').replace(/\s+/g, '') + '|' + String(course || '').trim(); }

/** { 班級: 人數 }；選修課另外放在 courseSizes：{ 「314/316|創意饅頭製作」: 36 } */
function readClassSizes_() {
  const out = {}, course = {};
  if (ss_().getSheetByName(CLASS_SIZE_SHEET)) readTable_(CLASS_SIZE_SHEET, ['班級', '人數']).rows.forEach(function (r) {
    const c = String(r['班級']).trim(), n = Number(r['人數']), co = String(r['課程'] || '').trim();
    if (!c || !(n > 0)) return;
    if (co) course[classSizeKey_(c, co)] = n; else out[c] = n;
  });
  out.__course = course;
  return out;
}

/**
 * 管理者：班級人數清單＋課表、國技排班、最近訂單裡出現過、還沒設定的班級。
 * 課表上同一個班有好幾門課（選修，例如 314/316 有創意饅頭製作、巧克力製作…）：每門課各列一行。
 */
function adminGetClassSizes(token) {
  auth_(token, true);
  const rows = {}, hidden = {};
  const put = function (c, co, size, note) { const k = classSizeKey_(c, co); if (!rows[k]) rows[k] = { className: c, course: co || '', size: '', note: '' }; if (size !== undefined) { rows[k].size = size; rows[k].note = note || ''; } };
  if (ss_().getSheetByName(CLASS_SIZE_SHEET)) readTable_(CLASS_SIZE_SHEET, ['班級', '人數']).rows.forEach(function (r) {
    const c = String(r['班級']).trim(), co = String(r['課程'] || '').trim(); if (!c) return;
    if (isTrue_(r['隱藏'])) { hidden[classSizeKey_(c, co)] = true; return; }
    put(c, co, Number(r['人數']) || '', String(r['備註'] || ''));
  });
  const add = function (v) { String(v || '').split(/[\/、,，\s]+/).forEach(function (c) { c = c.trim(); if (c && c.length <= 12) put(c, ''); }); };
  const sched = readSchedule_(), byCls = {};
  sched.forEach(function (x) { add(x.className); const c = String(x.className || '').trim(); if (c && x.course) (byCls[c] = byCls[c] || {})[x.course] = true; });
  Object.keys(byCls).forEach(function (c) { const cs = Object.keys(byCls[c]); if (cs.length > 1) cs.forEach(function (co) { put(c, co); }); });   // 選修：同一班好幾門課
  readGuoji_().forEach(function (g) { add(g.className); });
  const from = addDays_(today_(), -60);
  readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows.forEach(function (r) { const d = toDate_(r['使用日期']); if (d && d >= from) add(r['上課班級']); });
  Object.keys(hidden).forEach(function (k) { delete rows[k]; });
  return Object.keys(rows).map(function (k) { return rows[k]; });
}

/** 管理者：存班級人數。list = [{ className, course, size, note }]，人數空白＝不存；deleted = [{ className, course }]：刪掉（課表上有也不再列出來） */
function adminSaveClassSizes(token, list, deleted) {
  const admin = auth_(token, true);
  const rows = [], seen = {};
  (deleted || []).forEach(function (x) {
    if (typeof x === 'string') x = { className: x, course: '' };
    const c = clean_(x.className, 20), co = clean_(x.course, 40), k = classSizeKey_(c, co);
    if (c && !seen[k]) { seen[k] = true; rows.push({ '班級': c, '課程': co, '人數': '', '備註': '', '更新時間': new Date(), '更新人': admin, '隱藏': true }); }
  });
  (list || []).forEach(function (x) {
    const c = clean_(x.className, 20), co = clean_(x.course, 40), k = classSizeKey_(c, co), raw = String(x.size == null ? '' : x.size).trim();
    if (!c || seen[k]) return;
    if (raw === '') return;
    const n = Number(raw);
    if (!Number.isInteger(n) || n < 1 || n > 200) throw new Error('「' + c + (co ? ' ' + co : '') + '」的人數要是 1～200 的整數。');
    seen[k] = true;
    rows.push({ '班級': c, '課程': co, '人數': n, '備註': clean_(x.note, 50), '更新時間': new Date(), '更新人': admin });
  });
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(CLASS_SIZE_SHEET, CLASS_SIZE_HEADERS);
    ensureHeaders_(sh, CLASS_SIZE_HEADERS);
    const old = readTable_(CLASS_SIZE_SHEET, ['班級']).rows.map(function (r) { return r._row; });
    if (old.length) deleteRows_(sh, old);
    if (rows.length) appendRows_(CLASS_SIZE_SHEET, rows, { '班級': '@', '更新時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
    bootDirty_(CLASS_SIZE_SHEET);
  } finally {
    lock.releaseLock();
  }
  return { saved: rows.filter(function (r) { return !r['隱藏']; }).length, deleted: rows.filter(function (r) { return r['隱藏']; }).length };
}

/** 訂單是哪一門課（對課表：同老師、同一天、同時段），料單的選修人數用 */
function attachCourse_(list, teacher) {
  if (!list.length) return;
  const ds = list.map(function (o) { return o.useDate; }).filter(Boolean).sort();
  const from = parseDate_(ds[0]), to = parseDate_(ds[ds.length - 1]);
  if (!from || !to) return;
  const les = lessonsBetween_(from, to, teacher || null), by = {};
  les.forEach(function (x) { const k = x.teacher + '|' + x.date + '|' + x.session; if (!by[k]) by[k] = x.course; });
  list.forEach(function (o) { if (!o.course) o.course = by[(o.onBehalfOf || o.teacher) + '|' + o.useDate + '|' + sessionOf_(o.classTime)] || ''; });
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

/**
 * ➕ 快速補訂貨（LINE 管理者選單打開的小網頁）：某一天送貨多叫一兩樣，加到那天那家廠商的叫貨單（類型「補庫存」，不算哪位老師）。
 * rows: [{ id, name, vendor, qty, unit, note }]。那天那家已經「確定叫貨」的，補的也直接是已確定（叫貨單看得到，要另外通知廠商）；
 * 還沒確定的存成草稿，跟著「② 訂貨單」一起確定。回傳每家廠商補了什麼、要不要打電話。
 */
function adminQuickPurchase(token, date, rows) {
  const admin = auth_(token, true);
  const d = parseDate_(date);
  if (!d) throw new Error('請選送貨日期。');
  if (d < addDays_(today_(), -1) || d > addDays_(today_(), 60)) throw new Error('送貨日期只能選今天起兩個月內。');
  if (d.getDay() === 0 || d.getDay() === 6) throw new Error('送貨日期是週末，廠商不送貨。');
  const ds = dateStr_(d);
  if (!Array.isArray(rows) || !rows.length) throw new Error('請至少填一樣。');
  if (rows.length > 30) throw new Error('一次最多 30 樣。');
  const items = {};
  readItems_().forEach(function (it) { items[it.id] = it; });
  const okVendors = orderVendors_().concat(Object.keys(VENDOR_SPLIT).reduce(function (a, k) { return a.concat([VENDOR_SPLIT[k].baking, VENDOR_SPLIT[k].other]); }, []));
  const list = rows.map(function (x, i) {
    const it = items[String(x && x.id || '').trim()];
    const name = it ? it.name : clean_(x && x.name, 40);
    if (!name) throw new Error('第 ' + (i + 1) + ' 行沒有選食材。');
    let vendor = clean_(x.vendor, 30) || (it ? it.vendor : '');
    if (VENDOR_SPLIT[vendor]) throw new Error('「' + name + '」要選跟哪一家叫（' + VENDOR_SPLIT[vendor].other + '／' + VENDOR_SPLIT[vendor].baking + '）。');
    if (!vendor || okVendors.indexOf(vendor) < 0) throw new Error('「' + name + '」的廠商不正確。');
    const qty = Number(x.qty);
    if (!(qty > 0 && qty <= 99999)) throw new Error('「' + name + '」的數量不正確。');
    const unit = clean_(x.unit, 10);
    if (!unit) throw new Error('「' + name + '」沒有填單位。');
    return { id: it ? it.id : '', name: name, vendor: vendor, qty: Math.round(qty * 100) / 100, unit: unit, note: clean_(x.note, 40), addItem: !it && !!x.addItem };
  });
  const vg = readVendorGroups_(), contacts = readVendorContacts_(), now = new Date(), out = {};
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    ensureSheet_(SHEET.purchase, PURCHASE_HEADERS);
    // 食材總表沒有的：勾了「也加到食材總表」就新增（以後可以直接選）
    list.forEach(function (x) {
      if (!x.addItem) return;
      try { x.id = createItem_({ name: x.name, vendor: x.vendor, unit: x.unit }); x.added = true; } catch (e) { x.addErr = e.message; }
    });
    const have = {};
    readTable_(SHEET.purchase, ['送貨日期', '廠商']).rows.forEach(function (r) {
      if (purchaseDate_(r['送貨日期']) !== ds) return;
      const v = String(r['廠商'] || '').trim(), st = String(r['狀態'] || '').trim() === '草稿' ? '草稿' : '已確定';
      if (!have[v] || st === '已確定') have[v] = st;
    });
    const add = list.map(function (x) {
      const st = have[x.vendor] === '已確定' ? '已確定' : '草稿';
      const o = out[x.vendor] = out[x.vendor] || { vendor: x.vendor, status: st, phone: (contacts[x.vendor] || {}).phone || '', items: [] };
      o.items.push(x.name + ' ' + x.qty + ' ' + x.unit + (x.added ? '（已加到食材總表）' : x.addErr ? '（' + x.addErr + '）' : ''));
      return { '送貨日期': ds, '送貨組別': deliveryGroup_(x.vendor, d.getDay(), vg), '廠商': x.vendor, '訂單編號': '', '上課班級': '', '老師': '',
        '食材名稱': x.name, '食材ID': x.id, '需求量': '', '需求單位': '', '叫貨量': x.qty, '叫貨單位': x.unit, '類型': '補庫存',
        '備註': ('補訂貨' + (x.note ? '：' + x.note : '')).slice(0, 60), '狀態': st, '確定時間': now, '確認人': admin };
    });
    appendRows_(SHEET.purchase, add);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { date: ds, vendors: Object.keys(out).map(function (k) { return out[k]; }) };
}

/** 管理者儲存一家廠商的傳真、電話、聯絡人、備註，回傳全部廠商的聯絡資料 */
function adminSaveVendorContact(token, vendor, f) {
  auth_(token, true);
  const v = clean_(vendor, 40);
  if (!v) throw new Error('沒有廠商名稱。');
  f = f || {};
  const set = { '傳真': clean_(f.fax, 30), '電話': clean_(f.phone, 30), '聯絡人': clean_(f.person, 30), '備註': clean_(f.note, 100) };
  const lock = scriptLock_();
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

/** 管理者：全部食材（含停用的），停用／恢復用。回傳 [{ id, name, alias, vendor, dept, active, n }]（n＝最近 60 天被訂幾次） */
function adminItemsAll(token) {
  auth_(token, true);
  const since = addDays_(today_(), -60), cnt = {};
  readTable_(SHEET.orders, ['食材ID']).rows.forEach(function (r) {
    const d = r['使用日期'] instanceof Date ? toDate_(r['使用日期']) : parseDate_(String(r['使用日期'] || ''));
    if (d && d >= since) { const id = String(r['食材ID']).trim(); cnt[id] = (cnt[id] || 0) + 1; }
  });
  return readTable_(SHEET.items, ['食材ID', '食材名稱']).rows.filter(function (r) { return String(r['食材ID']).trim() && String(r['食材名稱']).trim(); }).map(function (r) {
    const id = String(r['食材ID']).trim();
    return { id: id, name: String(r['食材名稱']).trim(), alias: String(r['別名'] || '').trim(), vendor: String(r['廠商'] || '').trim(), dept: String(r['科別'] || '').trim(),
      active: r['啟用'] === undefined || isTrue_(r['啟用']), n: cnt[id] || 0 };
  });
}
/** 管理者：停用／恢復食材（ids 可以好幾個）。停用的老師搜不到、不能訂；舊訂單、成本不受影響 */
function adminSetItemsActive(token, ids, on) {
  auth_(token, true);
  ids = [].concat(ids || []).map(function (x) { return String(x || '').trim(); }).filter(Boolean).slice(0, 300);
  if (!ids.length) throw new Error('請選食材。');
  const lock = scriptLock_();
  lock.waitLock(30000);
  try {
    ensureHeaders_(sheet_(SHEET.items), ['啟用']);
    const t = readTable_(SHEET.items, ['食材ID', '食材名稱']);
    let n = 0;
    t.rows.forEach(function (r) { if (ids.indexOf(String(r['食材ID']).trim()) >= 0) { setCell_(t, r._row, '啟用', !!on); n++; } });
    if (!n) throw new Error('找不到這些食材，請重新整理。');
    SpreadsheetApp.flush();
    clearItemsCache_();
  } finally {
    lock.releaseLock();
  }
  return adminItemsAll(token);
}
function clearItemsCache_() {
  CacheService.getScriptCache().remove('items:n');
}

/**
 * 每樣食材老師平常用什麼單位訂：{ 食材ID: { 單位: 次數 } }（訂單總表＋歷史訂單，取消的不算）。
 * 訂貨畫面用來提醒打錯單位（小黃瓜 3 片 → 平常是條）
 */
/**
 * 管理者：每樣食材的單位建議。照訂單（含歷史）用過的單位（同義的合併），挑最常用的 1～3 個：
 * 用過 2 次以上、而且佔這樣食材 10% 以上的；預設單位一定放進去。回傳 [{ id, name, vendor, current, hist: [[單位, 次數]], suggest }]
 */
function adminSuggestItemUnits(token, mode) {
  auth_(token, true);
  // mode＝'cost'（成本好算）：菜、肉、海鮮（A、B、D 開頭，銷貨單照斤算）只留重量單位＋已經有換算的個數單位，老師就不會再訂「顆、條」換不過去
  const conv = mode === 'cost' ? readUnitConv_() : null;
  const hist = {};
  [SHEET.orders, SHEET.history].forEach(function (name) {
    if (!ss_().getSheetByName(name)) return;
    readTable_(name, ['食材ID', '單位']).rows.forEach(function (r) {
      const id = String(r['食材ID'] || '').trim(), u = normUnit_(r['單位']);
      if (!id || !u || String(r['狀態'] || '').trim() === '已取消') return;
      const h = hist[id] = hist[id] || {};
      h[u] = (h[u] || 0) + 1;
    });
  });
  return readTable_(SHEET.items, ['食材ID', '食材名稱']).rows.filter(function (r) {
    return String(r['食材ID']).trim() && (r['啟用'] === undefined || isTrue_(r['啟用']));
  }).map(function (r) {
    const id = String(r['食材ID']).trim(), h = hist[id] || {}, total = Object.keys(h).reduce(function (t, u) { return t + h[u]; }, 0);
    const list = Object.keys(h).map(function (u) { return [u, h[u]]; }).sort(function (a, b) { return b[1] - a[1]; });
    const def = normUnit_(r['單位']);
    let sug = list.filter(function (x) { return x[1] >= 2 && x[1] >= total * 0.1; }).slice(0, 3).map(function (x) { return x[0]; });
    if (def && sug.indexOf(def) < 0) sug = [def].concat(sug).slice(0, 3);
    if (!sug.length && list.length) sug = [list[0][0]];
    const vend = String(r['廠商'] || '').trim();
    if (conv && /^[ABD]/.test(vend)) {
      const nm = String(r['食材名稱']).trim(), cv = conv[nm] || [];
      const hasConv = function (u) { return cv.some(function (c) { return normUnit_(c.u) === u && UNIT_G_[c.u2]; }); };
      let w = list.map(function (x) { return x[0]; }).filter(function (u) { return UNIT_G_[u] || hasConv(u); }).slice(0, 4);
      if (!w.some(function (u) { return UNIT_G_[u]; })) w = ['斤'].concat(w).slice(0, 4);   // 至少要有一個重量單位
      if (w.indexOf('g') < 0 && w.length < 4) w.push('g');
      sug = w;
    }
    return { id: id, name: String(r['食材名稱']).trim(), vendor: String(r['廠商'] || '').trim(), current: itemUnits_(r['可用單位']), hist: list, total: total, suggest: sug };
  }).sort(function (a, b) { return b.total - a.total || a.name.localeCompare(b.name, 'zh-Hant'); });
}
/** 管理者：存「可用單位」changes: [{ id, units: '顆、斤' }]（空白＝不限制） */
function adminSetItemUnits(token, changes) {
  auth_(token, true);
  const ok = unitLists_().all, want = {};
  (changes || []).forEach(function (c) {
    const id = String(c && c.id || '').trim(); if (!id) return;
    const us = itemUnits_(c.units);
    us.forEach(function (u) { if (ok.indexOf(u) < 0) throw new Error('沒有「' + u + '」這個單位（先到單位管理新增）。'); });
    want[id] = us.join('、');
  });
  const lock = scriptLock_(); lock.waitLock(30000);
  let n = 0;
  try {
    ensureHeaders_(sheet_(SHEET.items), ['可用單位']);
    const t = readTable_(SHEET.items, ['食材ID', '可用單位']);
    const col = t.cols['可用單位'], range = t.sheet.getRange(2, col, Math.max(t.sheet.getLastRow() - 1, 1), 1), vals = range.getValues();
    t.rows.forEach(function (r) {
      const id = String(r['食材ID']).trim();
      if (!(id in want) || String(r['可用單位'] || '') === want[id]) return;
      vals[r._row - 2][0] = want[id]; n++;
    });
    if (n) { range.setValues(vals); clearItemsCache_(); }
  } finally { lock.releaseLock(); }
  return { saved: n };
}

function itemUnitHist_() {
  const out = {};
  try {
    [SHEET.orders, SHEET.history].forEach(function (name) {
      if (!ss_().getSheetByName(name)) return;
      readTable_(name, ['食材ID', '單位']).rows.forEach(function (r) {
        const id = String(r['食材ID'] || '').trim(), u = String(r['單位'] || '').trim();
        if (!id || !u || String(r['狀態'] || '').trim() === '已取消') return;
        const h = out[id] = out[id] || {};
        h[u] = (h[u] || 0) + 1;
      });
    });
  } catch (e) { console.error('itemUnitHist_ ' + e.message); }
  return out;
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
        units: itemUnits_(r['可用單位']),
        group: String(r['適用組別'] || '全部').trim() || '全部',
        code: String(r['舊編號'] || '').trim(),
        prefer: String(r['建議改用'] || '').trim(),   // 學校統一用別樣（例如酥油→羅亞無鹽）：老師加入時提醒，一定要用要填原因
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
/**
 * 送出前檢查重複：同一位老師（代訂就看被代訂的）、同一天、同一時段、同一班已經有沒取消的訂單。
 * list: [{ useDate, classTime, className, onBehalfOf }]，回傳 [{ orderNo, useDate, session, className, project, submitted }]
 */
function checkDuplicateOrders(token, list) {
  const account = auth_(token);
  const want = {};
  (list || []).slice(0, 20).forEach(function (x) {
    const who = clean_(x.onBehalfOf, 30) || account;
    want[[who, String(x.useDate || '').trim(), sessionOf_(x.classTime), String(x.className || '').trim()].join('|')] = true;
  });
  const out = {};
  readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows.forEach(function (r) {
    if (String(r['狀態']).trim() === '已取消') return;
    const no = String(r['訂單編號']).trim();
    if (out[no]) return;
    const who = String(r['代訂對象老師'] || '').trim() || String(r['操作老師'] || '').trim();
    const d = toDate_(r['使用日期']), ses = sessionOf_(r['上課時間']), cls = String(r['上課班級'] || '').trim();
    if (!d || !want[[who, dateStr_(d), ses, cls].join('|')]) return;
    out[no] = { orderNo: no, useDate: dateStr_(d), session: ses, className: cls, project: String(r['實作項目'] || ''), submitted: fmtDateTime_(r['送出時間']) };
  });
  return Object.keys(out).map(function (k) { return out[k]; });
}

function submitOrder(token, payload) {
  const account = auth_(token);
  const o = validateOrder_(payload, account, isAdmin_(account));

  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const orderNo = nextOrderNo_();
    const res = writeOrder_(o, orderNo, new Date(), account, account, {});
    SpreadsheetApp.flush();
    res.line = lineSyncInfo_(account);   // 送出成功畫面：告訴老師可以在 LINE 查看這張料單
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
  const lock = scriptLock_();
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
    p.keepStudent = String(r0['操作學生'] || '').trim();   // 校隊學生訂的單，管理者改了也保留是誰訂的
    if (operator !== account && !isAdmin) p.onBehalfOf = '';
    const o = validateOrder_(p, operator === account || isAdmin ? account : operator, isAdmin);
    o.keepStudent = p.keepStudent;
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
  const student = CUR_STUDENT_ || o.keepStudent || '';
  if (student) { ensureHeaders_(sheet_(SHEET.orders), ['操作學生']); base['操作學生'] = student; }
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
  const className = normClassList_(clean_(p.className, 30));
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
  if (!isAdmin_(account) && isCrew_(account)) {
    if (!onBehalfOf) throw new Error('校隊是幫老師代訂，請在「代訂對象老師」選是哪位老師的課。');
    if (isCrew_(onBehalfOf)) throw new Error('代訂對象要選老師，不能選校隊帳號。');
  }
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
      qty: qty_(it.qty, name), unit: allowUnit_(r, unit_(it.unit, r['單位'], name, newUnits), name), note: clean_(it.note, 100),
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
        qty: qty_(it.qty, realName), unit: allowUnit_(r, unit_(it.unit, r['單位'], realName, newUnits), realName), note: clean_(it.note, 100),
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
/** 給廠商、算成本用的單位：再套管理者設的合併（老師打「隻」，廠商訂貨單看到「條」） */
function vendUnit_(u) { u = normUnit_(u); return unitMerge_()[u] || u; }
/** 管理者自己設的單位合併（例如 隻 → 支）：存在指令碼屬性 UNIT_MERGE，{ 被合併的: 留下來的 } */
let UNIT_MERGE_ = null;
function unitMerge_() {
  if (UNIT_MERGE_) return UNIT_MERGE_;
  try { UNIT_MERGE_ = JSON.parse(PropertiesService.getScriptProperties().getProperty('UNIT_MERGE') || '{}') || {}; } catch (e) { UNIT_MERGE_ = {}; }
  return UNIT_MERGE_;
}
/** 管理者：單位合併清單（內建的＋自己設的） */
function adminGetUnitMerge(token) {
  auth_(token, true);
  return { builtin: UNIT_SYNONYMS, custom: unitMerge_() };
}
/**
 * 管理者：設定「from 當成 to」（remove＝true 取消）。fixOrders＝true：還沒叫貨的訂單（待處理、已審核）、範本裡的 from 一起改成 to。
 * 改完清掉開網頁的快取，老師重新整理就只看得到 to。
 */
function adminSetUnitMerge(token, from, to, remove, fixOrders) {
  const admin = auth_(token, true);
  from = clean_(from, 10); to = clean_(to, 10);
  const lock = scriptLock_(); lock.waitLock(20000);
  const out = { orders: 0, templates: 0 };
  try {
    const m = Object.assign({}, unitMerge_());
    if (remove) { delete m[from]; }
    else {
      if (!from || !to) throw new Error('請選兩個單位。');
      if (from === to) throw new Error('兩個單位一樣。');
      if (UNIT_SYNONYMS[from]) throw new Error('「' + from + '」系統已經當成「' + UNIT_SYNONYMS[from] + '」了。');
      const t2 = UNIT_SYNONYMS[to] || m[to] || to;   // 留下來的如果也被合併過，直接指到最後那個
      if (t2 === from) throw new Error('「' + to + '」已經合併到「' + from + '」，不能反過來再合併。');
      m[from] = t2;
      Object.keys(m).forEach(function (k) { if (m[k] === from) m[k] = t2; });   // 以前合併到 from 的，改指到新的
    }
    PropertiesService.getScriptProperties().setProperty('UNIT_MERGE', JSON.stringify(m));
    UNIT_MERGE_ = m; UNIT_CACHE_ = null;
    try { CacheService.getScriptCache().remove(BOOT_CACHE_KEY); } catch (e) {}
    if (false && !remove && fixOrders) {   // 現在不改老師的訂單：老師打什麼就是什麼，廠商訂貨單、成本才換
      const target = m[from];
      const o = readTable_(SHEET.orders, ['訂單編號', '單位']);
      o.rows.forEach(function (x) {
        const st = String(x['狀態'] || '').trim();
        if ((st === '待處理' || st === '已審核') && String(x['單位'] || '').trim() === from) { setCell_(o, x._row, '單位', target); out.orders++; }
      });
      const tsh = ss_().getSheetByName(SHEET.templates);
      if (tsh && tsh.getLastRow() >= 2) {
        const t = readTable_(SHEET.templates, ['名稱']);
        if (t.cols['單位'] !== undefined) t.rows.forEach(function (x) { if (String(x['單位'] || '').trim() === from) { setCell_(t, x._row, '單位', target); out.templates++; } });
      }
      if (out.orders && ss_().getSheetByName(SHEET.edits)) { try { appendRows_(SHEET.edits, [{ '時間': new Date(), '訂單編號': '（單位合併）', '修改人': admin, '內容': '單位「' + from + '」合併到「' + target + '」，還沒叫貨的 ' + out.orders + ' 行一起改' }]); } catch (e) {} }
    }
  } finally { lock.releaseLock(); }
  out.custom = unitMerge_();
  return out;
}

/** 食材總表「可用單位」（、隔開，空白＝不限制）→ 合併同義後的陣列 */
function itemUnits_(v) {
  const out = [];
  String(v || '').split(/[、,，\/\s]+/).forEach(function (u) { u = normUnit_(u); if (u && out.indexOf(u) < 0) out.push(u); });
  return out;
}
/** 這樣食材有設「可用單位」：單位要在裡面 */
function allowUnit_(r, unit, name) {
  const ok = itemUnits_(r && r['可用單位']);
  if (ok.length && ok.indexOf(unit) < 0) throw new Error('「' + name + '」只能用這些單位：' + ok.join('、') + '（管理者在「🥕 食材管理 → 每樣食材的單位」設定）。');
  return unit;
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
  UNIT_CACHE_ = { order: order, recipe: recipe, pending: pending, all: order.concat(recipe, pending) };   // 合併的單位老師照樣可以選（後端換算）
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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

/**
 * 好幾班一起上：「308.309.110」「308、309、110」「308 309 110」「308，309」都統一成「308/309/110」。
 * 每一段都要像班級（3 位數字，前面可以有 夜／國技／輪調… 幾個字）才合併，其他照原樣（例如「國技26」「選手培訓」）。
 */
function normClassList_(s) {
  s = String(s || '').trim();
  const parts = s.split(/[\/.．。、,，;；\s+＋&]+/).map(function (x) { return x.trim(); }).filter(Boolean);
  if (parts.length < 2 || !parts.every(function (x) { return /^[一-鿿]{0,3}\d{2,3}班?$/.test(x); })) return s;
  return parts.map(function (x) { return x.replace(/班$/, ''); }).join('/');
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
  try { attachSlipPrices_(shown); } catch (e) { console.error('attachSlipPrices_ ' + e.message); }
  try { attachCourse_(shown, account); } catch (e) { console.error('attachCourse_ ' + e.message); }
  return shown;
}

/**
 * 料單成本：每個品項帶上「那一週、那家廠商銷貨單上的單價」it.sp = { u: 單位, p: 單價 }（驗收網站填的「銷貨單明細」）。
 * 那週沒填的不帶，畫面再用單價表最新的估。
 */
function attachSlipPrices_(list) {
  if (!list.length || !ss_().getSheetByName('銷貨單明細')) return;
  const mon = function (d) { return d ? dateStr_(addDays_(d, -((d.getDay() + 6) % 7))) : ''; };
  const weeks = {};
  list.forEach(function (o) { weeks[mon(parseDate_(o.useDate))] = true; });
  const np = {}, slipG = {}, last = {}, cmap = readCostMap_(), map = slipPriceMap_(weeks, np, slipG, last), match = slipMatcher_(map, np, cmap);
  const ac = autoUnitConv_(weeks, match, slipG, mon);
  // 別週算得出來的「1 條 ≈ 幾公克」：同一家、同一樣，好幾週平均（那週分攤不出來時用）
  const acHist = {};
  Object.keys(ac).forEach(function (k) {
    const vk = k.split('|').slice(1).join('|');
    Object.keys(ac[k]).forEach(function (u) { const h = (acHist[vk] = acHist[vk] || {}); h[u] = h[u] || { g: 0, n: 0 }; h[u].g += ac[k][u]; h[u].n++; });
  });
  list.forEach(function (o) {
    const wk = mon(parseDate_(o.useDate));
    (o.items || []).forEach(function (it) {
      const r = match(wk, it.vendor, it.name);
      if (!r || r.np || r.sug) {
        const sub = cmap[String(it.vendor || '').trim() + '|' + it.name] || cmap['|' + it.name] || '';
        const lp = last[String(it.name || '').trim()] || (sub && sub !== '✗' ? last[sub] : null);
        if (lp) it.lp = lp;
      }
      if (!r) return;
      if (r.np) { it.spNP = r.np; return; }
      if (r.sug) { it.spSug = r.sug; return; }
      it.sp = { u: r.u, p: r.p, wk: wk, sn: r.k.split('|').slice(2).join('|'), sv: r.k.split('|')[1] };
      if (r.via) it.sp.via = r.via;
      if (r.src) it.sp.src = r.src;
      if (r.c) it.sp.c = r.c;
      // 菜商用台斤算、老師用條／隻／盒訂：照那週銷貨單分攤出來的「1 條 ≈ 幾斤」
      const c = ac[r.k], gu = UNIT_G_[r.u];
      const iu = unitSyn_(it.unit);
      if (c && gu && !UNIT_G_[iu] && c[iu]) it.sp.ac = { u: String(it.unit).trim(), n: c[iu] / gu };
      else if (gu && !UNIT_G_[iu]) {
        const h = (acHist[r.k.split('|').slice(1).join('|')] || {})[iu];
        if (h && h.n) it.sp.ac = { u: String(it.unit).trim(), n: h.g / h.n / gu, hist: h.n };
      }
    });
  });
}

/** 同義的個數單位當成同一種（算成本用）：個、粒、顆 → 顆；支、根、隻、尾、條 → 條 */
const UNIT_SYN_ = { '個': '顆', '粒': '顆', '顆': '顆', '支': '條', '根': '條', '隻': '條', '尾': '條', '條': '條', '附': '付', '副': '付', 'cc': 'ml', 'CC': 'ml' };
function unitSyn_(u) { u = String(u || '').trim(); u = unitMerge_()[u] || u; return UNIT_SYN_[u] || u; }   // 管理者設的合併（隻 → 支）也算同一個
const UNIT_G_ = { 'g': 1, '公克': 1, '克': 1, '公斤': 1000, 'kg': 1000, 'KG': 1000, 'Kg': 1000, '斤': 600, '台斤': 600, '兩': 37.5 };

/** 找訂單品項對到哪一樣銷貨單品項：回傳 { k: 「週|廠商|銷貨單品名」, u, p, via } 或 { np: 單價空白的說明 } 或 null */
function slipMatcher_(map, np, cmap) {
  const byVend = {}, byName = {};   // 那週各廠商銷貨單上有哪些品名、同一個品名在哪幾家
  Object.keys(map).forEach(function (k) {
    const a = k.split('|'), wv = a[0] + '|' + a[1], wn = a[0] + '|' + a.slice(2).join('|');
    (byVend[wv] = byVend[wv] || []).push(a.slice(2).join('|'));
    (byName[wn] = byName[wn] || []).push(a[1]);
  });
  // 訂單上的廠商跟銷貨單上的不一定一樣：蛋分烘焙／中餐，銷貨單是同一家
  // 乾貨：E中餐乾貨（銘哲）、E烘焙乾貨（旺來興）是不同家；訂單上「E乾貨」（通用的）兩家都可能送，兩家都找
  const alias = function (v) { return /^C.*蛋/.test(v) ? ['C蛋商', v] : v === 'E乾貨' ? ['E乾貨', 'E中餐乾貨', 'E烘焙乾貨'] : /^E/.test(v) ? [v, 'E乾貨'] : [v]; };
  const hit = function (k, via) { const x = map[k]; return { k: k, u: x.u, p: x.p, via: via || '', src: x.src, c: x.c }; };
  const memo = {};
  return function (wk, vendor, name) {
    name = String(name || '').trim();
    const mk = wk + '|' + vendor + '|' + name;
    if (mk in memo) return memo[mk];
    const vs = alias(String(vendor || '').trim());
    // 設定的成本對應（訂單廠商、同一家乾貨／蛋商的別家、或「每家」都算）
    let sub = '';
    for (let i = 0; i < vs.length && !sub; i++) sub = cmap[vs[i] + '|' + name] || '';
    sub = sub || cmap['|' + name] || '';
    const noAuto = sub === '✗';   // 管理者標了「✗ 不要自動對應」
    let r = null;
    for (let i = 0; i < vs.length && !r; i++) {
      if (map[wk + '|' + vs[i] + '|' + name]) r = hit(wk + '|' + vs[i] + '|' + name);
      // 成本對應：廠商送的是別樣（叫中薑、送嫩薑），用那一樣的單價
      else if (sub && !noAuto && map[wk + '|' + vs[i] + '|' + sub]) r = hit(wk + '|' + vs[i] + '|' + sub, sub);
    }
    // 訂單上的廠商跟銷貨單不一樣：那週只有一家廠商的銷貨單有「一模一樣」的名稱，就用那一家的單價
    if (!r && !noAuto) {
      const others = (byName[wk + '|' + name] || []).filter(function (v) { return vs.indexOf(v) < 0; });
      if (others.length === 1) r = hit(wk + '|' + others[0] + '|' + name, '自動（' + others[0] + '）');
    }
    // 那週銷貨單有這一樣，只是單價沒填：告訴管理者去補
    for (let i = 0; i < vs.length && !r; i++) {
      if (np[wk + '|' + vs[i] + '|' + name] || (sub && np[wk + '|' + vs[i] + '|' + sub])) r = { np: vs[i] + '｜' + wk };
    }
    // 名稱相近的（叫「蔥」、銷貨單是「青蔥」）只當「建議」，管理者按 ✓ 才算成本，免得對錯
    if (!r && !noAuto) {
      for (let i = 0; i < vs.length; i++) {
        const info = {}, nm = guessSlipName_(name, byVend[wk + '|' + vs[i]] || [], info);
        // 有把握的（同義詞、只差括號）直接算，標「自動相近」讓管理者核對；沒把握的才當建議
        if (nm && info.sure) { r = hit(wk + '|' + vs[i] + '|' + nm, '自動相近：' + nm); break; }
        if (nm) { const x = map[wk + '|' + vs[i] + '|' + nm]; r = { sug: { to: nm, sv: vs[i], u: x.u, p: x.p } }; break; }
      }
    }
    return (memo[mk] = r);
  };
}

/**
 * 自動單位換算：菜商銷貨單用台斤，老師用條／隻／盒訂。
 * 同一週、同一樣（對到同一個銷貨單品項）：銷貨單總重 − 老師用重量單位訂的，剩下的照「條」數平分，
 * 得到「1 條 ≈ 幾公克」。那週用重量以外的單位只有一種才算（條、顆混在一起分不出來）。
 * 回傳 { 銷貨單key: { 條: 公克 } }
 */
function autoUnitConv_(weeks, match, slipG, mon) {
  const agg = {}, uc = readUnitConv_();
  // 條、付、片混在一起：先換成同一種再分。來源：品名裡寫的「(1付=2片)」、系統設定「單位換算」裡兩邊都不是重量的（1 付 ＝ 2 片）
  const toBase = function (name, u) {
    name = String(name || '');
    const m = name.match(/1\s*([^\d\s=＝()（）]+?)\s*[=＝]\s*([\d.]+)\s*([^\d\s()（）]+)/);
    if (m && m[1] === u && Number(m[2]) > 0) return { b: m[3], f: Number(m[2]) };
    const c = (uc[name.trim()] || []).filter(function (x) { return unitSyn_(x.u) === u && !UNIT_G_[x.u2]; })[0];
    return c ? { b: unitSyn_(c.u2), f: c.q } : { b: u, f: 1 };
  };
  readTable_(SHEET.orders, ['使用日期', '食材名稱', '廠商', '數量', '單位']).rows.forEach(function (r) {
    const d = r['使用日期'] instanceof Date ? toDate_(r['使用日期']) : parseDate_(String(r['使用日期'] || ''));
    const wk = mon(d);
    if (!weeks[wk]) return;
    if (/取消|退/.test(String(r['狀態'] || '')) || String(r['不叫貨'] || '').trim()) return;
    const q = Number(r['數量']), u = unitSyn_(r['單位']);
    if (!(q > 0) || !u) return;
    const m = match(wk, String(r['廠商'] || '').trim(), r['食材名稱']);
    if (!m || m.np || !slipG[m.k]) return;
    const a = agg[m.k] = agg[m.k] || { g: 0, n: {}, alias: {} };
    if (UNIT_G_[u]) { a.g += q * UNIT_G_[u]; return; }
    const t = toBase(r['食材名稱'], u);
    a.n[t.b] = (a.n[t.b] || 0) + q * t.f;
    if (t.b !== u) a.alias[u] = t;
  });
  const out = {};
  Object.keys(agg).forEach(function (k) {
    const a = agg[k], us = Object.keys(a.n), left = slipG[k] - a.g;
    if (us.length !== 1 || !(left > 0)) return;
    const per = left / a.n[us[0]];
    out[k] = { [us[0]]: per };
    Object.keys(a.alias).forEach(function (u) { out[k][u] = per * a.alias[u].f; });   // 1 付＝2 片：付的重量＝片的兩倍
  });
  return out;
}

/** 常見的同義叫法（老師叫的 → 銷貨單上可能的名稱），自動對應時優先用 */
const COST_SYN = {
  // 菜商銷貨單的簡寫、錯字
  '紅蘿蔔': ['紅K', '紅k'], '白蘿蔔': ['白K', '白k', '菜頭'], '馬鈴薯': ['馬K', '馬k', '洋芋'], '紅辣椒': ['大辣椒', '大辣', '辣椒'], '辣椒': ['大辣椒', '大辣', '紅辣椒'],
  '杏鮑菇': ['杏包菇'], '金桔': ['金秸', '金吉'], '大黑豆干': ['大黑干', '黑豆乾', '黑豆干'], '紅色火龍果': ['紅龍果'], '白色火龍果': ['白龍果'],
  '板豆腐': ['大漢板豆腐'], '中華嫩豆腐': ['中華豆腐'], '新鮮豆包': ['生豆包', '豆包'], '乾魷魚身': ['乾魷魚'], '大蒜苗': ['蒜苗'], '青蒜苗': ['蒜苗'],
  '豆芽菜': ['綠豆芽'], '綠豆芽': ['豆芽菜'], '新鮮香菇': ['香菇'], '山東大白菜': ['大白菜'], '結球白菜': ['大白菜'],
  // 常見的同義叫法
  '蔥': ['青蔥'], '青蔥': ['蔥'], '嫩薑': ['子薑', '薑'], '中薑': ['子薑', '嫩薑', '老薑', '薑'], '生薑': ['老薑', '子薑', '薑'], '老薑': ['薑'], '薑': ['老薑', '子薑', '嫩薑'],
  '蒜': ['蒜頭', '蒜仁'], '蒜頭': ['蒜仁', '蒜'], '蒜仁': ['蒜頭', '蒜'], '香菜': ['芫荽'], '芫荽': ['香菜'], '高麗菜': ['甘藍'], '甘藍': ['高麗菜'], '地瓜': ['番薯'], '番薯': ['地瓜'],
  '洋芋': ['馬鈴薯'], '番茄': ['牛番茄', '蕃茄'], '蕃茄': ['牛番茄', '番茄'], '玉米粒': ['罐頭玉米粒'], '雞蛋': ['蛋'], '蛋': ['雞蛋'],
  '雞腿肉': ['去骨雞腿肉'], '大雞腿': ['雞腿'], '細砂糖': ['特砂糖', '白砂糖', '砂糖'], '砂糖': ['細砂糖', '特砂糖', '白砂糖'], '白糖': ['白砂糖', '細砂糖'],
  '核桃': ['核桃仁', '碎核桃'], '胡桃仁': ['核桃仁'], '雞粉': ['康寶雞粉'], '味醂': ['味之琳', '味霖'], '奶水': ['三花奶水'], '粗柴魚片': ['粗柴魚'], '咖哩塊': ['日本咖喱片']
};
/** 自動對應：在同一週、同一家廠商銷貨單的品名裡，找跟老師叫的名稱最像、而且只有一個的；找不到回傳 '' */
function guessSlipName_(name, cands, info) {
  info = info || {};
  info.sure = false;
  if (!cands.length) return '';
  const core = function (t) { return learnKey_(String(t || '').replace(/[(（\[【][^)）\]】]*[)）\]】]/g, '')); };
  const o = core(name);
  if (!o) return '';
  const syn = (COST_SYN[o] || []);
  for (let i = 0; i < syn.length; i++) { const hit = cands.filter(function (c) { return core(c) === syn[i]; }); if (hit.length) { info.sure = true; return hit[0]; } }
  const same = cands.filter(function (c) { return core(c) === o; });   // 只差括號、空白（無鹽奶油(安佳) ↔ 無鹽奶油）
  if (same.length === 1) { info.sure = true; return same[0]; }
  if (same.length) return '';
  const near = cands.filter(function (c) { const k = core(c); return k && (k.indexOf(o) >= 0 || o.indexOf(k) >= 0) && Math.min(k.length, o.length) >= 2; });
  return near.length === 1 ? near[0] : '';
}

/** 成本對應：{ 「廠商|訂的品項」: 銷貨單品項 }（廠商空白＝每家都適用） */
const COST_MAP_SHEET = '成本對應', COST_MAP_HEADERS = ['廠商', '訂的品項', '銷貨單品項', '更新時間', '更新人'];
function readCostMap_() {
  const out = {};
  if (!ss_().getSheetByName(COST_MAP_SHEET)) return out;
  readTable_(COST_MAP_SHEET, ['訂的品項', '銷貨單品項']).rows.forEach(function (r) {
    const a = String(r['訂的品項']).trim(), b = String(r['銷貨單品項']).trim();
    if (a && b) out[String(r['廠商'] || '').trim() + '|' + a] = b;
  });
  return out;
}
/** 管理者：存／刪一筆成本對應（slipName 空白＝刪掉）。回傳全部 [{ vendor, order, slip }] */
function adminSaveCostMap(token, vendor, orderName, slipName) {
  const admin = auth_(token, true);
  vendor = clean_(vendor, 30); orderName = clean_(orderName, 40); slipName = clean_(slipName, 40);
  if (!orderName) throw new Error('資料不完整。');
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(COST_MAP_SHEET, COST_MAP_HEADERS);
    const t = readTable_(COST_MAP_SHEET, ['訂的品項']);
    const old = t.rows.filter(function (r) { return String(r['廠商'] || '').trim() === vendor && String(r['訂的品項']).trim() === orderName; }).map(function (r) { return r._row; });
    if (old.length) deleteRows_(sh, old);
    if (slipName) appendRows_(COST_MAP_SHEET, [{ '廠商': vendor, '訂的品項': orderName, '銷貨單品項': slipName, '更新時間': new Date(), '更新人': admin }], { '更新時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return adminGetCostMap(token);
}
/** 管理者：一次存好幾筆成本對應 list: [{ vendor, order, slip }]（slip＝'✗' 代表不要自動對應） */
function adminSaveCostMaps(token, list) {
  const admin = auth_(token, true);
  list = (list || []).map(function (x) { return { v: clean_(x && x.vendor, 30), o: clean_(x && x.order, 40), s: clean_(x && x.slip, 40) }; }).filter(function (x) { return x.o && x.s; });
  if (!list.length) throw new Error('沒有要存的對應。');
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(COST_MAP_SHEET, COST_MAP_HEADERS);
    const t = readTable_(COST_MAP_SHEET, ['訂的品項']);
    const want = {}; list.forEach(function (x) { want[x.v + '|' + x.o] = true; });
    const old = t.rows.filter(function (r) { return want[String(r['廠商'] || '').trim() + '|' + String(r['訂的品項']).trim()]; }).map(function (r) { return r._row; });
    if (old.length) deleteRows_(sh, old);
    const now = new Date();
    appendRows_(COST_MAP_SHEET, list.map(function (x) { return { '廠商': x.v, '訂的品項': x.o, '銷貨單品項': x.s, '更新時間': now, '更新人': admin }; }), { '更新時間': 'yyyy/mm/dd hh:mm' });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { saved: list.length };
}
function adminGetCostMap(token) {
  auth_(token, true);
  const m = readCostMap_();
  return Object.keys(m).map(function (k) { const p = k.split('|'); return { vendor: p[0], order: p.slice(1).join('|'), slip: m[k] }; })
    .sort(function (a, b) { return (a.vendor + a.order).localeCompare(b.vendor + b.order, 'zh-Hant'); });
}
/** 管理者：這幾週每家廠商銷貨單上有哪些品項（給「沒有成本」挑對應用）：{ 廠商: [{ n, u, p }] } */
function adminSlipItems(token, from, to) {
  auth_(token, true);
  const f = parseDate_(from), t = parseDate_(to);
  if (!f || !t || !ss_().getSheetByName('銷貨單明細')) return {};
  const weeks = {};
  for (let d = addDays_(f, -((f.getDay() + 6) % 7)); d <= t; d = addDays_(d, 7)) weeks[dateStr_(d)] = true;
  const map = slipPriceMap_(weeks), out = {};
  Object.keys(map).forEach(function (k) {
    const p = k.split('|'), v = p[1], n = p.slice(2).join('|');
    if (!out[v]) out[v] = {};
    out[v][n] = { n: n, u: map[k].u, p: map[k].p };
  });
  Object.keys(out).forEach(function (v) { out[v] = Object.keys(out[v]).sort(function (a, b) { return a.localeCompare(b, 'zh-Hant'); }).map(function (n) { return out[v][n]; }); });
  return out;
}

/**
 * 💰 成本「依銷貨單總額分攤」：每週每家廠商銷貨單的食材金額合計（器材不算）。
 * 週＝銷貨單日期（沒有就用送貨日期）那週的週一。回傳 [{ wk, v, amt, n }]
 */
function adminSlipTotals(token, from, to) {
  auth_(token, true);
  const f = parseDate_(from), t = parseDate_(to);
  if (!f || !t || !ss_().getSheetByName('銷貨單明細')) return [];
  const mon = function (d) { return d ? dateStr_(addDays_(d, -((d.getDay() + 6) % 7))) : ''; };
  const lo = mon(f), hi = mon(t), out = {};
  readTable_('銷貨單明細', ['送貨日期（起）', '廠商']).rows.forEach(function (r) {
    if (String(r['類型'] || '').trim() === '器材') return;
    const sd = r['銷貨單日期'] instanceof Date ? toDate_(r['銷貨單日期']) : parseDate_(String(r['銷貨單日期'] || '').replace(/^'/, '')), d0 = toDate_(r['送貨日期（起）']);
    // 算成本照「這批貨是哪一週的課用的」：叫貨單的銷貨單（R 開頭）照叫貨那週（週五先送下週的貨也算下週）；臨時銷貨單照銷貨單日期
    const wk = /^R\d{8}-/.test(String(r['驗收編號'] || '')) && d0 ? mon(d0) : mon(sd || d0);
    if (!wk || wk < lo || wk > hi) return;
    const a = Number(r['金額']), q = Number(r['數量']), p = Number(r['單價']);
    const amt = a > 0 ? a : (q > 0 && p > 0 ? q * p : 0);
    const v = String(r['廠商'] || '').trim(), k = wk + '|' + v;
    if (!out[k]) out[k] = { wk: wk, v: v, amt: 0, n: 0, blank: 0, cls: {} };
    out[k].amt += amt; out[k].n++; if (!amt) out[k].blank++;
    // 菜商一班一張：驗收時有選班級的，記那一班的金額（成本直接算給那一班）
    const cls = String(r['上課班級'] || '').replace(/^'/, '').trim();
    if (cls) { const ck = cls + '|' + String(r['老師'] || '').trim(); out[k].cls[ck] = (out[k].cls[ck] || 0) + amt; }
  });
  return Object.keys(out).map(function (k) { const x = out[k]; x.amt = Math.round(x.amt); Object.keys(x.cls).forEach(function (c) { x.cls[c] = Math.round(x.cls[c]); }); return x; });
}

/** 銷貨單的單價：{ 「週一|廠商|品名（換成食材總表正式名稱）」: { u, p } }，只讀 weeks 裡的週 */
function slipPriceMap_(weeks, noPrice, grams, last) {
  const mon = function (d) { return d ? dateStr_(addDays_(d, -((d.getDay() + 6) % 7))) : ''; };
  // 銷貨單上的品名（廠商的叫法，例如「大骨高湯罐」）用食材總表的別名換成正式名稱（高湯罐），才對得到老師的訂單
  const official = {}, itemRows = readTable_(SHEET.items, ['食材名稱']).rows;
  itemRows.forEach(function (r) { const n = String(r['食材名稱'] || '').trim(); if (n) official[learnKey_(n)] = n; });   // 正式名稱優先
  itemRows.forEach(function (r) {
    const n = String(r['食材名稱'] || '').trim(); if (!n) return;
    String(r['別名'] || '').split(/[,，、]+/).forEach(function (a) { const k = learnKey_(a); if (k && !official[k]) official[k] = n; });
  });
  const toName = function (v) {
    const k = learnKey_(v); if (!k) return '';
    const k2 = learnKey_(String(v).replace(/[(（\[【][^)）\]】]*[)）\]】]/g, ''));   // 紅龍果(紅肉) → 紅龍果
    return official[k] || (k2 && official[k2]) || '';
  };
  const map = {}, rows = {};
  readTable_('銷貨單明細', ['送貨日期（起）', '廠商', '食材名稱', '單價']).rows.forEach(function (r) {
    const raw = String(r['食材名稱'] || '').trim(), p = Number(r['單價']);
    const name = toName(raw) || toName(r['銷貨單品名']) || raw;
    if (!name) return;
    const sd = r['銷貨單日期'] instanceof Date ? toDate_(r['銷貨單日期']) : parseDate_(String(r['銷貨單日期'] || '').replace(/^'/, '')), d0 = toDate_(r['送貨日期（起）']);
    // 算成本照「這批貨是哪一週的課用的」：叫貨單的銷貨單（R 開頭）照叫貨那週（週五先送下週的貨也算下週）；臨時銷貨單照銷貨單日期
    const wk = /^R\d{8}-/.test(String(r['驗收編號'] || '')) && d0 ? mon(d0) : mon(sd || d0);
    // 最近一次的單價（任何一週、任何一家）：那週銷貨單沒有這一樣時拿來估
    if (last && p > 0 && wk) {
      const lp = last[name];
      if (!lp || wk >= lp.wk) last[name] = { p: p, u: unitSyn_(r['單位']), wk: wk, v: String(r['廠商']).trim() };
    }
    if (!weeks[wk]) return;
    const key = wk + '|' + String(r['廠商']).trim() + '|' + name;
    const sq = Number(r['數量']), sg = UNIT_G_[String(r['單位'] || '').trim()];
    if (grams && sq > 0 && sg) grams[key] = (grams[key] || 0) + sq * sg;   // 那週銷貨單上這一樣總共幾公克（自動單位換算用）
    if (!(p > 0)) { if (noPrice) noPrice[key] = true; return; }   // 有這一樣、單價沒填
    const amt = Number(r['金額']);
    const pw = Number(r['每單位重量']), pwu = String(r['重量單位'] || '').trim();
    (rows[key] = rows[key] || []).push({ ref: { row: r._row, id: String(r['驗收編號'] || ''), n: raw, q: r['數量'] === '' ? '' : sq, u: String(r['單位'] || '').trim(), p: p, a: r['金額'] === '' ? '' : amt, d: sd ? dateStr_(sd) : '' },   // 成本頁可以直接改這幾行
      u: unitSyn_(r['單位']), p: p, q: sq > 0 ? sq : 0, a: amt > 0 ? amt : (sq > 0 ? p * sq : 0),
      c: pw > 0 && UNIT_G_[pwu] && !UNIT_G_[String(r['單位'] || '').trim()] ? { u: pwu, q: pw } : null });   // 1 罐＝100 g（銷貨單填的規格），老師用 g 訂也算得到
  });
  // 同一週同一樣常常有好幾筆（每班一筆、有的單價是金額÷數量算的有小數）：
  // 單價＝那週加總的金額 ÷ 加總的數量（重量單位先換成公克再換回最常用的單位），不是隨便拿最後一筆
  Object.keys(rows).forEach(function (key) { map[key] = weekPrice_(rows[key]); });
  return map;
}
/**
 * 管理者：成本頁直接改銷貨單明細的數量／單價／金額（不用跳到驗收網）。
 * edits = [{ row, id, name, qty, price, amount }]：先核對那一行的驗收編號、品名還一樣才改（別人剛好改過、行數移動就不改）。
 * 改完重算那張銷貨單的「銷貨單金額」，記到修改紀錄。
 */
function adminEditSlipRows(token, edits) {
  const who = auth_(token, true);
  edits = (edits || []).slice(0, 60);
  if (!edits.length) throw new Error('沒有要改的。');
  const numOrBlank = function (v) { const s = String(v === undefined || v === null ? '' : v).replace(/[,，\s]/g, ''); if (s === '') return ''; const n = Number(s); if (isNaN(n) || n < 0) throw new Error('「' + v + '」不是數字。'); return Math.round(n * 10000) / 10000; };
  const lock = scriptLock_(); lock.waitLock(20000);
  const done = [], ids = {}, logs = [];
  try {
    memoClear_();
    const t = readTable_('銷貨單明細', ['驗收編號', '食材名稱', '數量', '單價', '金額']);
    const byRow = {}; t.rows.forEach(function (r) { byRow[r._row] = r; });
    edits.forEach(function (e) {
      const r = byRow[Number(e.row)];
      if (!r || String(r['驗收編號']).trim() !== String(e.id || '').trim() || String(r['食材名稱']).trim() !== String(e.name || '').trim()) throw new Error('「' + e.name + '」那一行剛被改過或移動了，請按「查詢」重新整理再改。');
      const q = numOrBlank(e.qty), p = numOrBlank(e.price), a = numOrBlank(e.amount);
      const old = [r['數量'], r['單價'], r['金額']].map(function (v) { return v === '' ? '空白' : String(v); });
      if (String(r['數量']) === String(q) && String(r['單價']) === String(p) && String(r['金額']) === String(a)) return;
      setCells_(t, r._row, { '數量': q, '單價': p, '金額': a === '' ? '' : Math.round(a) });
      ids[String(r['驗收編號']).trim()] = 1;
      logs.push({ '時間': new Date(), '訂單編號': String(r['驗收編號']).trim(), '修改人': who, '內容': '成本頁改銷貨單「' + e.name + '」數量 ' + old[0] + '→' + (q === '' ? '空白' : q) + '、單價 ' + old[1] + '→' + (p === '' ? '空白' : p) + '、金額 ' + old[2] + '→' + (a === '' ? '空白' : Math.round(a)) });
      done.push(e.name);
    });
    // 那張銷貨單的金額合計跟著更新（驗收網、核銷看到的才一致）
    if (Object.keys(ids).length && ss_().getSheetByName('銷貨單紀錄')) {
      memoClear_();
      const t2 = readTable_('銷貨單明細', ['驗收編號', '金額']), sum = {};
      t2.rows.forEach(function (r) { const id = String(r['驗收編號']).trim(); if (ids[id]) sum[id] = (sum[id] || 0) + (Number(r['金額']) || 0); });
      const rec = readTable_('銷貨單紀錄', ['驗收編號']);
      if (rec.cols['銷貨單金額']) rec.rows.forEach(function (r) { const id = String(r['驗收編號']).trim(); if (ids[id] && sum[id] !== undefined) setCell_(rec, r._row, '銷貨單金額', Math.round(sum[id] * 100) / 100); });
    }
    if (logs.length && ss_().getSheetByName(SHEET.edits)) { try { appendRows_(SHEET.edits, logs); } catch (e) {} }
  } finally { lock.releaseLock(); }
  return { n: done.length, names: done };
}
/** 一週同一樣的好幾筆 → { u, p, src: { n 筆數, q 總量, a 總金額, lo 最低, hi 最高 } } */
function weekPrice_(list) {
  const groups = {};   // 重量單位一組（換成公克），其他照單位分組
  list.forEach(function (x) { const g = UNIT_G_[x.u] ? '#w' : x.u; (groups[g] = groups[g] || []).push(x); });
  const amtOf = function (xs) { return xs.reduce(function (t, x) { return t + x.a; }, 0); };
  const gk = Object.keys(groups).sort(function (a, b) { return amtOf(groups[b]) - amtOf(groups[a]) || groups[b].length - groups[a].length; })[0];
  const xs = groups[gk];
  let pu = xs[0].u;
  if (gk === '#w') {   // 用金額最多的那個重量單位當單價單位（菜商通常是斤／台斤）
    const by = {}; xs.forEach(function (x) { by[x.u] = (by[x.u] || 0) + x.a + 0.001; });
    pu = Object.keys(by).sort(function (a, b) { return by[b] - by[a]; })[0];
  }
  const toPu = function (x) { return gk === '#w' ? x.q * UNIT_G_[x.u] / UNIT_G_[pu] : x.q; };
  const unitP = function (x) { return gk === '#w' ? x.p * UNIT_G_[pu] / UNIT_G_[x.u] : x.p; };
  const withQ = xs.filter(function (x) { return x.q > 0 && x.a > 0; });
  const ps = xs.map(unitP);
  let p, q = 0, a = 0;
  if (withQ.length) {
    withQ.forEach(function (x) { q += toPu(x); a += x.a; });
    p = a / q;
  } else p = ps.reduce(function (t, v) { return t + v; }, 0) / ps.length;   // 都沒填數量：單價平均
  const r2 = function (v) { return Math.round(v * 100) / 100; };
  const cv = xs.filter(function (x) { return x.c; }).pop();
  return { u: pu, p: r2(p), c: cv ? cv.c : null, src: { n: xs.length, q: r2(q), a: Math.round(a), lo: r2(Math.min.apply(null, ps)), hi: r2(Math.max.apply(null, ps)),
    rows: list.filter(function (x) { return x.ref; }).slice(0, 40).map(function (x) { return x.ref; }) } };
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

function cancelOrder(token, orderNo, reason) {
  const account = auth_(token);
  const isAdmin = isAdmin_(account);
  reason = clean_(reason, 80);
  const res = cancelOrderChecked_(account, isAdmin, orderNo);
  const what = ['取消整張訂單' + (reason ? '（原因：' + reason + '）' : '')];
  try { logOrderEdit_(String(orderNo || '').trim(), account, what); } catch (e) { console.error(e.message); }
  if (isAdmin) notifyTeacherOfEdit_(String(orderNo || '').trim(), account, what);
  return res;
}

// ---- 📜 老師操作紀錄：老師說有訂、系統卻沒有時，查他到底做了什麼 ----

/**
 * 管理者：某位老師 from～to（操作的日期）做過的事，照時間排：登入、送出失敗（沒填完整、錯誤）、送出訂單、修改、取消、刪除、退單。
 * 資料來源：登入資訊、錯誤紀錄、訂單總表＋歷史訂單、修改紀錄、老師刪除（都是系統自動寫的，老師改不到）。
 */
function adminTeacherTrail(token, teacher, from, to) {
  fastReads_();
  auth_(token, true);
  teacher = clean_(teacher, 30);
  if (!teacher) throw new Error('請選老師。');
  const start = parseDate_(from) || addDays_(today_(), -14), end = addDays_(parseDate_(to) || today_(), 1);
  if (end <= start) throw new Error('結束日期要在開始日期之後。');
  if ((end - start) / 86400000 > 120) throw new Error('一次最多查 120 天。');
  const out = [], has = function (name) { const sh = ss_().getSheetByName(name); return sh && sh.getLastRow() >= 2; };
  const tm = function (v) { const d = v instanceof Date ? v : errTime_(v); return d instanceof Date && !isNaN(d) ? d : null; };
  const add = function (d, type, text, extra) {
    if (!d || d < start || d >= end) return;
    out.push(Object.assign({ t: d.getTime(), time: Utilities.formatDate(d, tz_(), 'yyyy/MM/dd HH:mm:ss'), type: type, text: text }, extra || {}));
  };
  const mine = function (v) { return String(v || '').trim() === teacher; };
  if (has(SHEET.logins)) readTable_(SHEET.logins, ['登入時間', '帳號']).rows.forEach(function (r) {
    if (mine(r['帳號'])) add(tm(r['登入時間']), /成功/.test(String(r['結果'])) ? '登入' : '登入失敗', String(r['結果'] || ''));
  });
  if (has(SHEET.errors)) readTable_(SHEET.errors, ['時間', '帳號']).rows.forEach(function (r) {
    if (!mine(r['帳號'])) return;
    const fn = String(r['功能'] || ''), msg = String(r['訊息'] || '').replace(/^Error:\s*/, '');
    const send = /送出訂單|submitOrder/.test(fn), unsent = fn === '料單沒送出';
    add(tm(r['時間']), unsent ? '📝 填了沒送出' : send ? '⚠ 送出失敗' : '系統訊息', (send || unsent ? '' : fn + '：') + msg, { page: String(r['畫面'] || '').split('｜')[0] });
  });
  // 訂單：送出時間在這段期間的（也列上課日期在這段期間、但比較早送出的，方便對）
  const orders = groupOrders_(withHistory_(readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows, addDays_(start, -60)), null)
    .filter(function (o) { return mine(o.teacher) || mine(o.onBehalfOf); });
  const byNo = {};
  orders.forEach(function (o) {
    byNo[o.orderNo] = o;
    const live = o.items.filter(function (i) { return i.status !== '已取消'; }).length;
    const desc = '上課 ' + shortDateTw_(parseDate_(o.useDate)) + ' ' + (o.classTime || '') + '・' + (o.className || '') + '・' + (o.project || '（沒填實作項目）') + '・' + (live || o.items.length) + ' 項・現在狀態：' + o.status +
      (o.onBehalfOf ? (mine(o.teacher) ? '（幫 ' + o.onBehalfOf + ' 代訂）' : '（' + o.teacher + ' 代訂）') : '') + (o.student ? '・訂的學生：' + o.student : '');
    add(parseDateTime_(o.submitted) || tm(o.submitted), '✅ 送出訂單', desc, { orderNo: o.orderNo });
  });
  if (has(SHEET.edits)) readTable_(SHEET.edits, ['時間', '訂單編號']).rows.forEach(function (r) {
    const no = String(r['訂單編號']).trim(), who = String(r['修改人'] || ''), text = String(r['內容'] || '');
    if (!mine(who) && !byNo[no]) return;
    const type = /(^|；)退單/.test(text) ? '退單' : /取消整張訂單|整張取消/.test(text) ? '取消訂單' : /老師刪除訂單/.test(text) ? '刪除訂單' : '修改訂單';
    add(tm(r['時間']), type, (mine(who) ? '' : who + '：') + text, { orderNo: no });
  });
  if (has(SHEET.skips)) readTable_(SHEET.skips, ['老師', '日期', '時段']).rows.forEach(function (r) {
    if (!mine(r['老師']) || String(r['標記'] || '').trim() === '已叫料') return;
    const by = String(r['標記人'] || '').trim();
    add(tm(r['標記時間']), '不用訂貨', '上課 ' + shortDateTw_(toDate_(r['日期'])) + ' ' + String(r['時段'] || '') + ' ' + String(r['課程'] || '') + (String(r['備註'] || '').trim() ? '・' + String(r['備註']).trim() : '・（沒寫原因）') + (by && by !== teacher ? '（' + by + ' 標記）' : ''));
  });
  if (has(SHEET.hidden)) readTable_(SHEET.hidden, ['訂單編號']).rows.forEach(function (r) {
    if (mine(r['老師'])) add(tm(r['刪除時間']), '刪除訂單', '老師自己刪除（刪除時狀態：' + String(r['原狀態'] || '') + '）', { orderNo: String(r['訂單編號']).trim() });
  });
  out.sort(function (a, b) { return a.t - b.t; });
  const n = function (ty) { return out.filter(function (x) { return x.type === ty; }).length; };
  return { teacher: teacher, from: dateStr_(start), to: dateStr_(addDays_(end, -1)), events: out,
    summary: { logins: n('登入'), sent: n('✅ 送出訂單'), failed: n('⚠ 送出失敗'), unsent: n('📝 填了沒送出'), edits: n('修改訂單'), cancels: n('取消訂單') + n('刪除訂單') } };
}

// ---- 異常紀錄：標記異常、老師刪除、退單、修改、取消、抽查異常，集中在一張表 ----

/**
 * 管理者：from～to（事件發生的日期）的異常紀錄，最新的在前。countOnly：只回傳還沒處理的件數（不讀訂單，比較快）
 * 每筆：{ id, time, type, orderNo, useDate, className, project, teacher, who, note, done }
 */
function adminGetIssues(token, from, to, countOnly) {
  fastReads_();
  auth_(token, true);
  // 用「上課日期」篩（找不到訂單的才用發生時間）；沒給日期＝前後 30 天
  const start = parseDate_(from) || addDays_(today_(), -30), end = addDays_(parseDate_(to) || addDays_(today_(), 30), 1);
  const out = [];
  const tm = function (v) { const d = v instanceof Date ? v : errTime_(v); return d instanceof Date && !isNaN(d) ? d : null; };
  const inRange = function (d) { return d && d >= start && d < end; };
  const evFrom = addDays_(start, -180), evTo = addDays_(end, 60);   // 事件大概發生在上課前後這段時間
  const add = function (d, type, orderNo, who, note, extra) {
    if (!d || d < evFrom || d >= evTo) return;
    out.push(Object.assign({ id: type + '|' + orderNo + '|' + Utilities.formatDate(d, tz_(), 'yyyyMMddHHmmss') + (extra && extra.item ? '|' + extra.item : ''),
      t: d.getTime(), time: fmtDateTime_(d), type: type, orderNo: orderNo, who: who, note: note }, extra || {}));
  };
  const has = function (name) { const sh = ss_().getSheetByName(name); return sh && sh.getLastRow() >= 2; };
  if (has(SHEET.flags)) readTable_(SHEET.flags, ['訂單編號']).rows.forEach(function (r) {
    const item = String(r['食材'] || '').trim();
    add(tm(r['標記時間']), '標記異常', String(r['訂單編號']).trim(), String(r['標記人'] || ''), String(r['說明'] || ''),
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
  // 補上訂單的上課日期、時段、班級、實作項目、老師、送出時間
  const want = {};
  out.forEach(function (x) { want[x.orderNo] = true; });
  const meta = {};
  groupOrders_(withHistory_(readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows, start).filter(function (r) { return want[String(r['訂單編號']).trim()]; }), null)
    .forEach(function (o) { meta[o.orderNo] = o; });
  out.forEach(function (x) {
    const o = meta[x.orderNo];
    if (o) {
      x.useDate = x.useDate || o.useDate; x.className = x.className || o.className; x.project = o.project; x.teacher = x.teacher || o.onBehalfOf || o.teacher;
      x.classTime = o.classTime; x.submitted = o.submitted;
    }
    if (x.type === '標記異常' && x.item) {   // 「香菜 3 公斤（異常）：說明」
      const it = o && o.items.filter(function (i) { return i.name === x.item; })[0];
      x.note = x.item + (it ? ' ' + it.qty + ' ' + it.unit : '') + '（異常）' + (x.note ? '：' + x.note : '');
    }
  });
  markResent_(out);
  const list = out.filter(function (x) { const u = parseDate_(x.useDate); return u ? inRange(u) : inRange(new Date(x.t)); });
  if (countOnly) return list.filter(function (x) { return !x.done; }).length;
  return list.sort(function (a, b) { return String(b.useDate || '').localeCompare(String(a.useDate || '')) || b.t - a.t; });
}

/**
 * 退單後老師重新送單：同一位老師、退單之後送出、上課日期或實作項目或班級一樣、沒取消的訂單。
 * 找到就在退單那筆寫 resent，舊單的退單、取消、標記異常都自動算已處理。
 */
function markResent_(out) {
  const returned = out.filter(function (x) { return x.type === '退單' && x.teacher; });
  if (!returned.length) return;
  const who = {};
  returned.forEach(function (x) { who[x.teacher] = true; });
  const sent = {};
  const rows = readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows.filter(function (r) {
    const t = String(r['代訂對象老師'] || '').trim() || String(r['操作老師'] || '').trim();
    if (!who[t]) return false;
    const d = r['送出時間'] instanceof Date ? r['送出時間'] : errTime_(r['送出時間']);
    if (d instanceof Date && !isNaN(d)) sent[String(r['訂單編號']).trim()] = d.getTime();
    return true;
  });
  const cands = groupOrders_(rows, null).filter(function (o) { return o.status !== '已取消'; });
  returned.forEach(function (x) {
    const hit = cands.filter(function (o) {
      return o.orderNo !== x.orderNo && (o.onBehalfOf || o.teacher) === x.teacher && (sent[o.orderNo] || 0) > x.t &&
        (o.useDate === x.useDate || (x.project && o.project === x.project) || (x.className && o.className === x.className));
    }).sort(function (a, b) { return sent[a.orderNo] - sent[b.orderNo]; })[0];
    if (!hit) return;
    x.resent = { orderNo: hit.orderNo, useDate: hit.useDate, time: fmtDateTime_(new Date(sent[hit.orderNo])) };
    out.forEach(function (y) {
      if (y.orderNo !== x.orderNo || ['退單', '取消', '標記異常'].indexOf(y.type) < 0) return;
      y.autoDone = true;
      if (!y.done) y.done = '老師已重新送單';
    });
  });
}

/** 管理者：異常紀錄勾「已處理」（on）或取消 */
function adminSetIssueDone(token, id, on) {
  const who = auth_(token, true);
  id = String(id || '').slice(0, 200);
  if (!id) throw new Error('缺少事件。');
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  fastReads_();
  auth_(token, true);
  let start = parseDate_(from);
  let end = parseDate_(to) || start;
  if (!start) throw new Error('請選擇日期。');
  if (end < start) { const t = start; start = end; end = t; }
  const all = readTable_(SHEET.orders, ['訂單編號', '使用日期']).rows;
  const rows = withHistory_(all, start).filter(function (r) {
    const d = toDate_(r['使用日期']);
    return d && d >= start && d <= end;
  });
  const list = groupOrders_(rows, null).sort(function (a, b) {
    return a.useDate === b.useDate ? (a.orderNo < b.orderNo ? -1 : 1) : (a.useDate < b.useDate ? -1 : 1);
  });
  if (lite) return list;
  try { markOddQty_(list, all); } catch (e) { console.error('markOddQty_ ' + e.message); }
  const nos = list.map(function (o) { return o.orderNo; });
  const edits = orderEditsFor_(nos), printed = printedFor_(nos), flags = flagsFor_(nos), places = placesFor_(nos), hidden = hiddenOrders_();
  list.forEach(function (o) { o.edits = edits[o.orderNo] || []; o.printed = printed[o.orderNo] || null; o.flag = flags.orders[o.orderNo] || null; o.itemFlags = flags.items[o.orderNo] || {}; o.place = places[o.orderNo] || null; o.deleted = hidden[o.orderNo] || null; });
  try { matchLessons_(list, start, end); } catch (e) { console.error('matchLessons_ ' + e.message); }
  try { attachSlipPrices_(list); } catch (e) { console.error('attachSlipPrices_ ' + e.message); }
  try { attachCourse_(list, null); } catch (e) { console.error('attachCourse_ ' + e.message); }
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

/**
 * 數量、單位可能打錯（例如 5 顆打成 35 顆、2 斤打成 2 公斤）：跟這項食材以前的訂單比。
 * 比平常（中位數）多 3 倍以上，或大家幾乎都用別的單位 → it.qtyWarn 寫原因，o.qtyWarn＝幾項。只提醒管理者，不擋老師送單。
 */
const QTY_GRAMS_ = { 'g': 1, '公斤': 1000, '斤': 600, '兩': 37.5 };
function markOddQty_(list, rows) {
  const stat = {};
  rows.forEach(function (r) {
    if (String(r['狀態']).trim() === '已取消') return;
    const name = String(r['食材名稱'] || '').trim(), unit = String(r['單位'] || '').trim(), q = Number(r['數量']) || 0;
    if (!name || !unit || q <= 0) return;
    const s = stat[name] || (stat[name] = { units: {}, n: 0, vals: {} });
    s.n++;
    s.units[unit] = (s.units[unit] || 0) + 1;
    const g = QTY_GRAMS_[unit], k = g ? 'g' : unit;
    (s.vals[k] || (s.vals[k] = [])).push(g ? q * g : q);
  });
  const median = function (a) { a = a.slice().sort(function (x, y) { return x - y; }); const m = a.length >> 1; return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2; };
  const fmt = function (v) { return String(Math.round(v * 100) / 100); };
  list.forEach(function (o) {
    if (o.status === '已取消') return;
    let n = 0;
    o.items.forEach(function (it) {
      const s = stat[String(it.name).trim()];
      if (!s || it.status === '已取消' || !(it.qty > 0)) return;
      const why = [];
      // 單位：至少 5 筆、8 成以上用同一種單位，這張卻用別的
      const top = Object.keys(s.units).sort(function (a, b) { return s.units[b] - s.units[a]; })[0];
      if (s.n >= 5 && top !== it.unit && s.units[top] / s.n >= 0.8) why.push('平常用「' + top + '」');
      // 數量：同一種單位（重量換成克比）至少 3 筆，超過中位數 3 倍
      const g = QTY_GRAMS_[it.unit], vals = s.vals[g ? 'g' : it.unit];
      if (vals && vals.length >= 3) {
        const m = median(vals), v = g ? it.qty * g : it.qty;
        if (m > 0 && v >= m * 3) why.push('平常約 ' + fmt(g ? m / g : m) + ' ' + it.unit);
      }
      if (why.length) { it.qtyWarn = why.join('，'); n++; }
    });
    if (n) o.qtyWarn = n;
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  appendRows_(SHEET.edits, [{ '時間': new Date(), '訂單編號': orderNo, '修改人': who, '內容': (CUR_STUDENT_ ? '［' + CUR_STUDENT_ + '］' : '') + (reason ? '原因：' + reason + '；' : '') + changes.join('；') }],
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

/** 管理者：勾選的訂單一次審核（只改還是「待處理」的品項；別人剛改過狀態的不動） */
function adminApproveMany(token, orderNos) {
  const admin = auth_(token, true);
  const want = {};
  (orderNos || []).slice(0, 300).forEach(function (no) { want[String(no).trim()] = true; });
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '狀態']);
    const done = {};
    const nums = t.rows.filter(function (r) {
      const no = String(r['訂單編號']).trim();
      if (!want[no] || String(r['狀態']).trim() !== '待處理') return false;
      done[no] = true;
      return true;
    }).map(function (r) { return r._row; }).sort(function (a, b) { return a - b; });
    if (!nums.length) return { done: [] };
    try { bootDirty_(t.sheet.getName()); } catch (e) {}
    memoClear_();
    // 同一張訂單的列通常連在一起：連續的列一次寫，比一格一格快很多
    const writeCol = function (h, v) {
      const c = t.cols[h];
      if (c === undefined) return;
      for (let i = 0; i < nums.length;) {
        let j = i;
        while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++;
        t.sheet.getRange(nums[i], c, j - i + 1, 1).setValues(nums.slice(i, j + 1).map(function () { return [v]; }));
        i = j + 1;
      }
    };
    writeCol('狀態', '已審核');
    writeCol('審核人', admin);
    SpreadsheetApp.flush();
    return { done: Object.keys(done) };
  } finally {
    lock.releaseLock();
  }
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

  const lock = scriptLock_();
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
      if (changedOrders) { t.sheet.getRange(2, col, statusVals.length, 1).setValues(statusVals); dataBump_(); }
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
/**
 * 管理者退單。noClass＝這堂課不用上（月考、放假…）：通知寫「不用再訂」，這堂課自動標成「不用訂貨」（不會再提醒）。
 */
function adminReturnOrder(token, orderNo, reason, notify, noClass) {
  const admin = auth_(token, true);
  orderNo = String(orderNo || '').trim();
  reason = clean_(reason, 100);
  if (!reason) throw new Error('請填寫退單原因，老師才知道要改什麼。');
  updateOrderStatus_(orderNo, '已取消', function (rows) {
    if (rows.every(function (r) { return String(r['狀態']).trim() === '已取消'; })) throw new Error('這張訂單已經取消了。');
  }, null);
  try { logOrderEdit_(orderNo, admin, [noClass ? '退單（這堂課不用上，不用再訂）' : '退單'], reason); } catch (e) { console.error(e.message); }
  let marked = '';
  if (noClass) {
    try {
      const o = orderOf_(orderNo), d = o && parseDate_(o.useDate), teacher = o && (o.onBehalfOf || o.teacher), session = o && sessionOf_(o.classTime);
      if (d && teacher) {
        const sh = ensureSheet_(SHEET.skips, SKIP_HEADERS);
        ensureHeaders_(sh, SKIP_HEADERS);
        const t = readTable_(SHEET.skips, ['老師', '日期', '時段']);
        const same = t.rows.some(function (r) { const rd = toDate_(r['日期']); return String(r['老師']).trim() === teacher && rd && dateStr_(rd) === dateStr_(d) && String(r['時段']).trim() === session; });
        if (!same) appendRows_(SHEET.skips, [{ '老師': teacher, '日期': d, '時段': session, '課程': clean_(o.project, 40), '標記時間': new Date(), '備註': '不用訂貨原因：' + reason + '（退單）', '標記人': admin }],
          { '日期': 'yyyy-mm-dd', '標記時間': 'yyyy/mm/dd hh:mm' });
        marked = '這堂課已標成「不用訂貨」。';
      }
    } catch (e) { console.error('退單標不用訂貨 ' + e.message); }
  }
  return { ok: true, marked: marked, notify: notify === false ? '沒有通知老師。' : notifyTeacherOfEdit_(orderNo, admin, ['退單'], reason, noClass ? 'return-noclass' : 'return') };
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
    const ret = kind === 'return' || kind === 'return-noclass', noClass = kind === 'return-noclass';
    const text = (ret ? '【退單通知】' + teacher + ' 老師您好，管理者退回了您的訂單：\n' : '【訂單修改通知】' + teacher + ' 老師您好，管理者修改了您的訂單：\n') +
      (d ? shortDateTw_(d) + ' ' : '') + o.className + '「' + o.project + '」\n' +
      (reason ? (ret ? '退單原因：' : '修改原因：') + reason + '\n' : '') +
      (ret ? '' : changes.map(function (c) { return '・' + c; }).join('\n') + '\n') +
      (noClass ? '\n這堂課不用上課，不需要再訂貨，系統已經幫您標成「不用訂貨」，不會再提醒。有問題請找管理者。'
        : ret ? '\n請到「我的訂單」找到這張，按「再訂一次」修改後重新送出（已過截止的話請找管理者）。' : '\n有問題請找管理者。') +
      (!noClass && lineUrl_('mine') ? '\n' + lineUrl_('mine') : '');
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
  const lock = scriptLock_();
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
    if (!groups[name]) { groups[name] = { name: name, rows: [], teachers: [], orderNos: [], depts: [], uses: [], lines: 0, time: '', vendor: '', unit: '' }; list.push(groups[name]); }
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
    // 每位老師要多少（審核時點開看）
    x.uses.push({
      orderNo: String(r['訂單編號']).trim(), teacher: String(r['代訂對象老師'] || r['操作老師'] || '').trim(),
      date: dateStr_(toDate_(r['使用日期'])), session: String(r['上課時間'] || '').trim(), className: String(r['上課班級'] || '').trim(),
      project: String(r['實作項目'] || '').trim(), groups: String(r['組數'] || '').trim(),
      qty: String(r['數量'] === undefined ? '' : r['數量']).trim(), unit: String(r['單位'] || '').trim(), note: String(r['備註'] || '').trim(),
    });
  });
  return list.map(function (x) {
    x.uses.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    return {
      row: x.rows[0] || 0, name: x.name, time: x.time, teacher: x.teachers.join('、'), vendor: x.vendor, unit: x.unit,
      orderNo: x.orderNos.join('、'), orders: x.orderNos.length, lines: x.lines, records: x.rows.length, orphan: !x.rows.length, depts: x.depts, uses: x.uses,
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
    if (ov.dirty) { ov.range.setValues(ov.values); dataBump_(); }
    if (og && og.dirty) og.range.setValues(og.values);
    return { items: items, lines: lines };
  } finally {
    lock.releaseLock();
  }
}

/** 食材總表的廠商改成 vendor（訂單上的廠商名稱，例如 E烘焙乾貨、A菜商），還沒叫貨的訂單（今天以後、待處理或已審核）也跟著改。回傳改了的訂單品項 [{ orderNo, vendor, group }] */
function setItemVendor_(id, vendor) {
  const vendorGroups = readVendorGroups_();
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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

/**
 * 授權跳不出來、或出現「你沒有呼叫 ○○ 的權限」時：在 Apps Script 編輯器選這個函式按「執行」。
 * 有任何一項權限沒勾到（例如「連線至外部服務」），會跳出授權視窗補勾；全部都有就直接執行完畢。
 */
function authorizeAll() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  return '全部權限都有了';
}

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
    receiveUrl: props.getProperty('RECEIVE_URL') || '',
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
  const stu = token ? CacheService.getScriptCache().get('sessstu:' + token) : '';
  logError_(account, String(fn || ''), fn === '料單沒送出' ? '操作紀錄' : errorType_(msg), (stu ? '［' + stu + '］' : '') + msg, String(view || ''));   // 沒送出的料單不算系統錯誤
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
      // 全形數字、空白換成半形；常見錯字（榜定、邦定、幫定）當成「綁定」，不然對不到、官方帳號完全不回
      ev.message.text = String(ev.message.text).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
        .replace(/\u3000/g, ' ').replace(/^\s*(榜定|邦定|幫定|绑定)/, '綁定');
      const userId = ev.source && ev.source.userId;
      if (/^驗證碼$/.test(String(ev.message.text).trim())) {
        const codeAns = updateCodeReply_(ev, props);
        if (codeAns) { lineReply_(ev.replyToken, codeAns); return; }
        // 不是綁定管理者通知的聊天室：告訴他要去哪裡傳（不回驗證碼）
        const to = String(props.getProperty('LINE_TO') || '');
        lineReply_(ev.replyToken, !to ? '還沒有綁定管理者通知的聊天室，收不到驗證碼。請到訂貨網「系統設定 → LINE 官方帳號通知」按「產生綁定碼」，在要收通知的聊天室傳「綁定 ○○○○○○」。'
          : '驗證碼只會回覆在「綁定管理者通知」的那個' + (/^[CR]/.test(to) ? '群組' : '個人聊天室') + '，這裡不會回。請到那裡傳「驗證碼」；或在網站按「收不到？改用推播傳送」。');
        return;
      }
      const cm = String(ev.message.text).trim().match(/^校隊\s*(\d{6})\s*(.{0,30})$/);
      if (cm) { lineReply_(ev.replyToken, lineCrewBind_(cm[1], ev.source)); return; }
      // 校隊登記名字、組別：「我是 中餐 王同學」
      if (/^我是\s*\S/.test(String(ev.message.text).trim()) && ev.source && ev.source.type === 'user' && lineIsCrew_(userId)) { lineReply_(ev.replyToken, lineCrewName_(String(ev.message.text).trim(), userId)); return; }
      const m = String(ev.message.text).trim().match(/^綁定\s*(群組)?\s*(\d{6})$/);
      if (!m) {
        // 選單按鈕（我的未訂貨課程、我的訂單、截止時間說明）：免費回覆；群組裡不回
        if (ev.source && ev.source.type === 'user') {
          const adm = lineAdminReply_(ev.message.text, userId);   // 管理者：課表、料單、補訂貨（卡片，回覆不算則數）
          if (adm) { lineReplyMsgs_(ev.replyToken, adm); return; }
          const tl = lineTodayLessonsMsgs_(ev.message.text, userId);   // 老師：今天的課＋每堂的料單
          if (tl) { lineReplyMsgs_(ev.replyToken, tl); return; }
          const fp = lineForgotPw_(ev.message.text, userId);   // 忘記密碼：回一次性的重設連結
          if (fp) { lineReplyMsgs_(ev.replyToken, fp); return; }
          const om = lineMyOrdersMsgs_(ev.message.text, userId);   // 老師：我的訂單（卡片列出每張的材料）
          if (om) { lineReplyMsgs_(ev.replyToken, om); return; }
          const ans = lineAutoReply_(ev.message.text, userId);
          if (ans) lineReply_(ev.replyToken, ans);
          // 看起來是要綁定（有 6 位數字或「綁」字）但格式不對：告訴他怎麼打，不要完全沒反應
          else if (/\d{6}|綁/.test(ev.message.text)) lineReply_(ev.replyToken, '看不懂這句。要綁定請傳「綁定 123456」（綁定兩個字＋空格＋網站上的 6 位數字）；校隊傳「校隊 123456」。');
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
        if (isAdmin_(teacher)) { lineLinkAdminMenu_(userId); lineReply_(ev.replyToken, teacher + '，綁定完成！您是管理者，下方選單會換成管理者版：補訂貨、訂貨系統、驗收系統、查看課表、查看料單、查看叫貨單（選單沒換的話，關掉聊天室再打開）。'); return; }
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
  // 綁定管理者通知的聊天室；或內建管理者帳號本人的個人聊天（管理者選單「🔑 驗證碼」）
  const owner = ev.source && ev.source.type === 'user' && ADMIN_ACCOUNTS.indexOf(lineAccountOf_(ev.source.userId)) >= 0;
  if (!src || (src !== props.getProperty('LINE_TO') && !owner)) return null;
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
  const lock = scriptLock_();
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
    let ver;
    try { ver = gas_('post', '/versions', { description: 'GitHub ' + short + '：' + msg }).versionNumber; }
    catch (e) { try { gas_('put', '/content', { files: before.files }); } catch (e2) {} throw versionLimitError_(e, '訂貨系統'); }   // 建不了版本：程式換回原本的，網站照舊
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
 * 在你的 Google 帳號新增一個 Apps Script 專案、貼上 GitHub 上 receive-site 的程式、部署成網頁應用程式（誰都打得開網址，但要用管理者帳號密碼登入）。
 */
function adminReceiveSite(token, branch, sha, code) {
  const who = updateOwner_(token);
  branch = clean_(branch, 100) || UPDATE_DEFAULT_BRANCH;
  if (!/^[0-9a-f]{40}$/.test(String(sha || ''))) throw new Error('版本不正確，請重新整理。');
  verifyUpdateCode_('receive', sha, code, who);
  const lock = scriptLock_();
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
    let old = null;
    try { old = gasAt_(id, 'get', '/content'); } catch (e) {}
    gasAt_(id, 'put', '/content', { files: files });
    let ver;
    try { ver = gasAt_(id, 'post', '/versions', { description: 'GitHub ' + short + '：' + msg }).versionNumber; }
    catch (e) { if (old && old.files) { try { gasAt_(id, 'put', '/content', { files: old.files }); } catch (e2) {} } throw versionLimitError_(e, '進貨驗收'); }
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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

  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  // 要查的日期都還在訂單總表裡（比總表最早的上課日期晚）：不用讀歷史訂單（省 1～3 秒）
  if (since) {
    let min = null;
    rows.forEach(function (r) { const d = toDate_(r['使用日期']); if (d && (!min || d < min)) min = d; });
    if (min && since > min) return rows;
  }
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
  // 瀏覽器沒給檔案類型（或給 octet-stream）時照副檔名補；Drive 轉檔看類型，類型錯會回「Invalid Value」
  const EXT_MIME = { xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xls: 'application/vnd.ms-excel', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', csv: 'text/csv', ods: 'application/vnd.oasis.opendocument.spreadsheet',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', doc: 'application/msword', odt: 'application/vnd.oasis.opendocument.text', rtf: 'application/rtf', txt: 'text/plain',
    jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', pdf: 'application/pdf', heic: 'image/heic', heif: 'image/heif' };
  const ext = (name.match(/\.(\w+)$/) || [])[1] ? name.match(/\.(\w+)$/)[1].toLowerCase() : '';
  let mime = String(file.mimeType || '').trim();
  if (!mime || mime === 'application/octet-stream') mime = EXT_MIME[ext] || 'application/octet-stream';
  if (/heic|heif/i.test(mime)) throw new Error('這張照片是 iPhone 的 HEIC 格式，Google 沒辦法辨識文字。請改用截圖，或在 iPhone「設定 → 相機 → 格式」選「最相容」再拍一次。');
  const blob = Utilities.newBlob(bytes, mime, name);
  const target = kind === 'sheet' ? MimeType.GOOGLE_SHEETS : MimeType.GOOGLE_DOCS;
  const id = convertFile_(blob, name, target, kind === 'ocr');
  try {
    return kind === 'sheet' ? readConvertedSheet_(id) : readConvertedDoc_(id, kind === 'ocr');
  } finally {
    try { Drive.Files.remove(id); } catch (e) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e2) {} }
  }
}

/**
 * 用 Drive API 把檔案轉成 Google 格式（照片／PDF 會做文字辨識），回傳新檔案 ID。
 * 文字辨識的語言提示：先試 zh-TW，Google 回「Invalid Value」再試 zh，最後不給語言（Google 自己判斷）。
 */
function convertFile_(blob, name, targetMime, ocr) {
  const langs = ocr ? ['zh-TW', 'zh', ''] : [''];
  let last = null;
  for (let i = 0; i < langs.length; i++) {
    try {
      return retry_(function () {
        if (typeof Drive === 'undefined') return convertViaRest_(blob, name, targetMime, ocr, langs[i]);   // 沒開「Drive API」服務：直接呼叫 Drive 網路 API
        if (Drive.Files.create) {   // Drive API v3
          const f = Drive.Files.create({ name: '[匯入暫存] ' + name, mimeType: targetMime }, blob, langs[i] ? { ocrLanguage: langs[i] } : {});
          return f.id;
        }
        const f = Drive.Files.insert({ title: '[匯入暫存] ' + name, mimeType: targetMime }, blob,   // Drive API v2
          ocr ? (langs[i] ? { ocr: true, ocrLanguage: langs[i], convert: true } : { ocr: true, convert: true }) : { convert: true });
        return f.id;
      });
    } catch (e) {
      last = e;
      if (!/Invalid Value|invalid/i.test(e.message)) throw e;
    }
  }
  throw new Error('Google 沒辦法讀這個檔案（' + (blob.getContentType() || '未知類型') + '）。請改存成 .xlsx／.docx，或用截圖、JPG 照片再試。（' + (last ? last.message : '') + '）');
}

/** 不用進階服務，直接用 Drive API v3 上傳並轉檔（權限跟 DriveApp 一樣） */
function convertViaRest_(blob, name, targetMime, ocr, lang) {
  const boundary = 'import' + Date.now();
  const meta = JSON.stringify({ name: '[匯入暫存] ' + name, mimeType: targetMime });
  const head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
    '\r\n--' + boundary + '\r\nContent-Type: ' + (blob.getContentType() || 'application/octet-stream') + '\r\n\r\n').getBytes();
  const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id' + (ocr && lang ? '&ocrLanguage=' + encodeURIComponent(lang) : ''), {
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
      alt: toDate_(r['隔週']) ? dateStr_(toDate_(r['隔週'])) : '',
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
  memoClear_();
  try { tblCacheDrop_(sh.getName()); } catch (e) {}
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
  const lock = scriptLock_();
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

/** 管理者：選單要不要顯示「抽查」（沒在用就藏起來） */
function adminSetAuditEnabled(token, on) {
  auth_(token, true);
  const props = PropertiesService.getScriptProperties();
  if (on) props.setProperty('AUDIT_ENABLED', '1'); else props.deleteProperty('AUDIT_ENABLED');
  try { CacheService.getScriptCache().remove(BOOT_CACHE_KEY); } catch (e) {}
  return { auditEnabled: on ? true : false };
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
  const lock = scriptLock_();
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

/** 管理者：老師帳號清單（不含校隊；不回傳密碼，只說還是不是預設密碼） */
function adminListTeachers(token) {
  auth_(token, true);
  const t = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']);
  return t.rows.filter(function (r) { return String(r['帳號']).trim() && !isTrue_(r['校隊']); }).map(function (r) {
    const a = String(r['帳號']).trim();
    return { account: a, enabled: isTrue_(r['啟用']), admin: ADMIN_ACCOUNTS.indexOf(a) >= 0 || isTrue_(r['管理者']), builtin: ADMIN_ACCOUNTS.indexOf(a) >= 0, isDefault: String(r['密碼']) === DEFAULT_PASSWORD };
  }).sort(function (a, b) { return (b.enabled - a.enabled) || a.account.localeCompare(b.account, 'zh-Hant'); });
}

/**
 * 管理者：看全部帳號的密碼。要再打一次自己的密碼才給（避免別人用開著的電腦偷看），每次查看都記在錯誤紀錄（操作紀錄）。
 */
function adminRevealPasswords(token, myPassword) {
  const me = auth_(token, true);
  const t = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']);
  const mine = t.rows.find(function (x) { return String(x['帳號']).trim() === me; });
  if (!mine || String(mine['密碼']) !== String(myPassword || '')) {
    logError_(me, '查看老師密碼', '操作紀錄', '密碼打錯，沒有顯示', '系統設定');
    throw new Error('你的密碼不正確，沒有顯示。');
  }
  logError_(me, '查看老師密碼', '操作紀錄', '查看了全部帳號的密碼', '系統設定');
  const out = {};
  t.rows.forEach(function (r) { const a = String(r['帳號']).trim(); if (a) out[a] = String(r['密碼']); });
  return out;
}

/**
 * 管理者：新增老師帳號（可以一次好幾位，用「、」或空白隔開；預設密碼 0000，第一次登入要自己改）、停用／啟用、重設密碼（改回 0000）。
 * 帳號就是老師的名字（登入、代訂、課表都用這個名字）。
 */
function adminSetTeacher(token, account, action) {
  const me = auth_(token, true);
  const lock = scriptLock_();
  lock.waitLock(20000);
  const out = { added: [], existed: [] };
  try {
    const t = readTable_(SHEET.accounts, ['帳號', '密碼', '啟用']);
    const find = function (a) { return t.rows.find(function (x) { return String(x['帳號']).trim() === a; }); };
    if (action === 'add') {
      const names = String(account || '').split(/[、,，\s\n]+/).map(function (n) { return clean_(n, 30); }).filter(Boolean);
      if (!names.length) throw new Error('請輸入老師的名字。');
      const add = [];
      names.forEach(function (n) {
        if (/[0-9]{3,}|[<>"'=]/.test(n)) throw new Error('「' + n + '」不像名字，請檢查。');
        const r = find(n) || t.rows.find(function (x) { return accountKey_(x['帳號']) === accountKey_(n); });
        if (r) {
          if (isTrue_(r['校隊'])) throw new Error('「' + n + '」已經是校隊帳號，老師請換一個名字。');
          if (!isTrue_(r['啟用'])) setCell_(t, r._row, '啟用', true);
          out.existed.push(String(r['帳號']).trim());
        } else if (add.indexOf(n) < 0) add.push(n);
      });
      if (add.length) {
        appendRows_(SHEET.accounts, add.map(function (n) { return { '帳號': n, '密碼': DEFAULT_PASSWORD, '啟用': true }; }));
        // 密碼欄設成純文字，0000 才不會變成數字 0
        const t2 = readTable_(SHEET.accounts, ['帳號', '密碼']);
        t2.rows.forEach(function (x) { if (add.indexOf(String(x['帳號']).trim()) >= 0) { const c = t2.sheet.getRange(x._row, t2.cols['密碼']); c.setNumberFormat('@'); c.setValue(DEFAULT_PASSWORD); } });
      }
      out.added = add;
    } else {
      account = clean_(account, 30);
      const r = find(account);
      if (!r || isTrue_(r['校隊'])) throw new Error('找不到老師帳號「' + account + '」。');
      if (action === 'disable') {
        if (ADMIN_ACCOUNTS.indexOf(account) >= 0) throw new Error('「' + account + '」是內建的管理者，不能停用。');
        if (account === me) throw new Error('不能停用自己的帳號。');
        setCell_(t, r._row, '啟用', false);
      } else if (action === 'enable') setCell_(t, r._row, '啟用', true);
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
  try { CacheService.getScriptCache().removeAll([BOOT_CACHE_KEY, 'extraAdmins']); } catch (e) {}
  out.list = adminListTeachers(token);
  return out;
}

// ---- 💌 給老師的留言：管理者在系統設定寫，老師登入時跳出來（存在指令碼屬性，不用另開分頁） ----
const NOTES_KEY_ = 'LOGIN_NOTES';
function notesAll_() {
  try { const v = JSON.parse(PropertiesService.getScriptProperties().getProperty(NOTES_KEY_) || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; }
}
/** 這個帳號今天看得到的留言（在期間內、對象是全部或有他） */
function notesFor_(account) {
  const t = dateStr_(today_()), seen = notesSeen_(account);
  return notesAll_().filter(function (n) {
    return (!n.start || n.start <= t) && (!n.end || n.end >= t) && (!n.to || !n.to.length || n.to.indexOf(account) >= 0);
  }).map(function (n) { return { id: n.id, text: n.text, from: n.from, start: n.start, end: n.end, style: n.style || '', seen: seen.indexOf(n.id) >= 0 }; });
}
// 看過哪幾則記在伺服器（換電腦、換手機也不會再跳）
function notesSeen_(account) {
  try { const v = JSON.parse(PropertiesService.getScriptProperties().getProperty('NOTE_SEEN_' + account) || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; }
}
/** 老師按「收到」：記下看過了（只留還存在的留言） */
function markNotesSeen(token, ids) {
  const account = auth_(token, false, false, true);
  const live = notesAll_().map(function (n) { return n.id; });
  const seen = notesSeen_(account).concat((ids || []).map(String)).filter(function (id, i, a) { return live.indexOf(id) >= 0 && a.indexOf(id) === i; });
  PropertiesService.getScriptProperties().setProperty('NOTE_SEEN_' + account, JSON.stringify(seen));
  return seen.length;
}
/** 管理者：全部留言 */
function adminListNotes(token) {
  auth_(token, true);
  return notesAll_();
}
/** 管理者：新增／修改留言 n = { id?, text, from, start, end, to: [帳號]（空＝全部老師）, style } */
function adminSaveNote(token, n) {
  const account = auth_(token, true);
  n = n || {};
  const text = String(n.text || '').replace(/\r/g, '').trim().slice(0, 300);
  if (!text) throw new Error('請輸入留言內容。');
  const day = function (v) { v = String(v || '').trim(); if (!v) return ''; if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error('日期格式不對。'); return v; };
  const start = day(n.start), end = day(n.end);
  if (start && end && end < start) throw new Error('結束日期不能早於開始日期。');
  const to = (Array.isArray(n.to) ? n.to : []).map(function (x) { return clean_(x, 30); }).filter(Boolean).slice(0, 100);
  const style = ['', 'xmas', 'cny', 'moon', 'teacher', 'thanks', 'cheer'].indexOf(String(n.style || '')) >= 0 ? String(n.style || '') : '';
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    let list = notesAll_();
    const id = String(n.id || '') || 'N' + Date.now() + Math.random().toString(36).slice(2, 6);
    const o = { id: id, text: text, from: clean_(n.from, 30), start: start, end: end, to: to, style: style, by: account, at: fmtDateTime_(new Date()) };
    const i = list.findIndex(function (x) { return x.id === id; });
    if (i >= 0) list[i] = o; else list.unshift(o);
    // 太多就把已經過期最久的刪掉（屬性一格最多 9KB）
    const t = dateStr_(today_());
    while (list.length > 30 || JSON.stringify(list).length > 8500) {
      let k = -1;
      list.forEach(function (x, j) { if (x.end && x.end < t && (k < 0 || x.end < list[k].end)) k = j; });
      if (k < 0) k = list.length - 1;
      if (list[k].id === id) throw new Error('留言太多了，請先刪掉幾則舊的。');
      list.splice(k, 1);
    }
    PropertiesService.getScriptProperties().setProperty(NOTES_KEY_, JSON.stringify(list));
    return list;
  } finally {
    lock.releaseLock();
  }
}
function adminDeleteNote(token, id) {
  auth_(token, true);
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const list = notesAll_().filter(function (x) { return x.id !== String(id || ''); });
    PropertiesService.getScriptProperties().setProperty(NOTES_KEY_, JSON.stringify(list));
    return list;
  } finally {
    lock.releaseLock();
  }
}

/** 管理者：校隊帳號清單 */
function adminListCrew(token) {
  auth_(token, true);
  const t = readTable_(SHEET.accounts, ['帳號', '啟用']);
  if (t.cols['校隊'] === undefined) return [];
  return t.rows.filter(function (r) { return isTrue_(r['校隊']); }).map(function (r) {
    return { account: String(r['帳號']).trim(), enabled: isTrue_(r['啟用']), canOrder: isTrue_(r['可訂貨']) };
  });
}

/** 管理者：新增校隊帳號（預設密碼 0000，第一次登入要改密碼）、停用／啟用、重設密碼、刪除（要先停用） */
function adminSetCrew(token, account, action) {
  auth_(token, true);
  account = clean_(account, 30);
  if (!account) throw new Error('請輸入校隊的名字。');
  const lock = scriptLock_();
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
      if (action === 'order-on' || action === 'order-off') {
        ensureHeaders_(sheet_(SHEET.accounts), ['可訂貨']);
        const t2 = readTable_(SHEET.accounts, ['帳號']);
        setCell_(t2, r._row, '可訂貨', action === 'order-on');
      } else if (action === 'disable') setCell_(t, r._row, '啟用', false);
      else if (action === 'enable') setCell_(t, r._row, '啟用', true);
      else if (action === 'delete') {
        if (isTrue_(r['啟用'])) throw new Error('請先按「停用」再刪除「' + account + '」。');
        deleteRows_(t.sheet, [r._row]);
      } else if (action === 'reset') {
        const cell = t.sheet.getRange(r._row, t.cols['密碼']);
        cell.setNumberFormat('@');
        cell.setValue(DEFAULT_PASSWORD);
      } else throw new Error('不明的動作。');
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().removeAll(['crewAccounts', 'crewAccounts2']);
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
      if (altWeekOff_(x.alt, ds)) return;   // 隔週上課：這週不上
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
function lessonsWithStatus_(from, to, teacher, st0) {
  const lessons = lessonsBetween_(from, to, teacher);
  const st = st0 || lessonStatusMaps_(from, to), have = st.have, skip = st.skip, manual = st.manual;
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
    if (x.status === 'skipped' && typeof skip[k] === 'string') x.skipNote = skip[k];
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
  fastReads_();
  const account = lessonAccount_(token, as);
  const now = new Date();
  const from = earliestClassDate_(now);
  const to = addDays_(today_(), MAX_DAYS_AHEAD);
  const firstWeekEnd = addDays_(from, 4);
  const deadline = addDays_(from, -7 + (lockRule_().wd - 1));
  const wk0 = addDays_(today_(), -((today_().getDay() + 6) % 7));   // 這週一：未對應的訂單從這週開始找（鎖單的這週也算）
  const st = lessonStatusMaps_(wk0 < from ? wk0 : from, to);
  const lessons = lessonsWithStatus_(from, to, account, st).sort(function (a, b) {
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
    unmatched: unmatchedOrders_(account, wk0 < from ? wk0 : from, to, st),
  };
}

/** 班級對不對得上：「213/214」和「213」算對得上（拆開來有一樣的） */
function sameClass_(a, b) {
  const sp = function (v) { return String(v || '').split(/[\/、,，\s]+/).map(function (x) { return x.trim(); }).filter(Boolean); };
  const x = sp(a), y = sp(b);
  if (!x.length || !y.length) return true;   // 沒填班級的不算錯
  return x.some(function (v) { return y.indexOf(v) >= 0; });
}

/**
 * 老師的訂單對不到課表（我的課「未對應」）：reason 'time'＝那天那個時段沒有這位老師的課；'class'＝時間對得到但班級不一樣。
 * cands：同一週這位老師的課，給老師選「其實是哪一堂」。
 */
function unmatchedOrders_(teacher, from, to, st) {
  const les = lessonsBetween_(from, to, teacher);
  const weekOf = function (ds) { const d = parseDate_(ds); return dateStr_(addDays_(d, -((d.getDay() + 6) % 7))); };
  const out = [];
  st.list.forEach(function (o) {
    if (o.teacher !== teacher || o.status === '已取消') return;
    const same = les.filter(function (x) { return x.date === o.date && x.session === o.session; });
    const reason = !same.length ? 'time' : same.some(function (x) { return sameClass_(x.className, o.className); }) ? '' : 'class';
    if (!reason) return;
    const wk = weekOf(o.date);
    // 時間對不到、但管理者已經在總課表把那週的課手動標成「已叫料」：管理者處理過了，不再列
    if (reason === 'time' && les.some(function (x) { return weekOf(x.date) === wk && st.manual[teacher + '|' + x.date + '|' + x.session]; })) return;
    const seen = {};
    out.push({ orderNo: o.orderNo, date: o.date, session: o.session, classTime: o.classTime, className: o.className, room: o.room, project: o.project, status: o.status,
      reason: reason, week: wk, locked: isLocked_(parseDate_(o.date)),
      cands: les.filter(function (x) {
        const k = x.date + '|' + x.session + '|' + x.className;
        if (weekOf(x.date) !== wk || seen[k]) return false;
        return (seen[k] = true);
      }).map(function (x) { return { date: x.date, session: x.session, course: x.course, className: x.className, room: x.room }; })
        .sort(function (a, b) { return a.date === b.date ? SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session) : (a.date < b.date ? -1 : 1); }) });
  });
  return out.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
}

/**
 * 老師把訂單對到正確的課：只改上課資訊（日期、星期、時段、班級、教室），食材數量不動；鎖單後也可以（管理者看得到修改紀錄）。
 * 改日期只能在同一週，而且訂單還沒叫貨（已叫貨的改日期會對不上叫貨單，請管理者處理）。
 */
function fixOrderLesson(token, orderNo, lesson, as) {
  const account = lessonAccount_(token, as);
  orderNo = clean_(orderNo, 30);
  lesson = lesson || {};
  const d = parseDate_(lesson.date), session = String(lesson.session || '').trim();
  if (!orderNo || !d || SESSIONS.indexOf(session) < 0) throw new Error('資料不正確，請重新整理。');
  // 只能選自己課表上的課
  const mon = addDays_(d, -((d.getDay() + 6) % 7));
  const L = lessonsBetween_(d, d, account).filter(function (x) { return x.session === session && String(x.className || '') === String(lesson.className || '') && String(x.room || '') === String(lesson.room || ''); })[0];
  if (!L) throw new Error('課表上找不到這堂課，請重新整理。');
  const lock = scriptLock_();
  lock.waitLock(20000);
  let before, changes = [];
  try {
    const t = readTable_(SHEET.orders, ['訂單編號', '狀態']);
    const rows = t.rows.filter(function (r) { return String(r['訂單編號']).trim() === orderNo; });
    if (!rows.length) throw new Error('找不到訂單 ' + orderNo + '。');
    const r0 = rows[0];
    if (String(r0['操作老師'] || '').trim() !== account && String(r0['代訂對象老師'] || '').trim() !== account && !isAdmin_(auth_(token))) throw new Error('只能修改自己的訂單。');
    const active = rows.filter(function (r) { return String(r['狀態']).trim() !== '已取消'; });
    if (!active.length) throw new Error('這張訂單已取消。');
    before = groupOrders_(rows, null)[0];
    const od = toDate_(r0['使用日期']);
    const set = {};
    if (!od || dateStr_(od) !== dateStr_(d)) {
      if (!od || dateStr_(addDays_(od, -((od.getDay() + 6) % 7))) !== dateStr_(mon)) throw new Error('只能對到同一週的課；要換到別週請聯絡管理者。');
      if (active.some(function (r) { return String(r['狀態']).trim() !== '待處理' && String(r['狀態']).trim() !== '已審核'; })) throw new Error('這張訂單已經叫貨，改日期會對不上叫貨單，請聯絡管理者。');
      set['使用日期'] = d; set['星期'] = '星期' + WEEKDAY[d.getDay()];
    }
    if (sessionOf_(r0['上課時間']) !== session) set['上課時間'] = session;
    if (String(r0['上課班級'] || '').trim() !== String(L.className || '').trim() && L.className) set['上課班級'] = clean_(L.className, 30);
    if (String(r0['上課教室'] || '').trim() !== String(L.room || '').trim() && L.room) set['上課教室'] = clean_(L.room, 20);
    if (!Object.keys(set).length) return { ok: true, changed: false };
    const vg = set['使用日期'] ? readVendorGroups_() : null;
    rows.forEach(function (r) {
      Object.keys(set).forEach(function (k) { if (t.cols[k] !== undefined) setCell_(t, r._row, k, set[k]); });
      if (set['使用日期'] && t.cols['送貨組別'] !== undefined && String(r['食材ID']).trim() !== TOOL_ID) {
        setCell_(t, r._row, '送貨組別', deliveryGroup_(groupVendor_(String(r['廠商'] || '').trim()), d.getDay(), vg));
      }
      if (set['使用日期'] && t.cols['使用日期'] !== undefined) t.sheet.getRange(r._row, t.cols['使用日期']).setNumberFormat('yyyy-mm-dd');
    });
    SpreadsheetApp.flush();
    try { changes = diffOrders_(before, orderOf_(orderNo)); logOrderEdit_(orderNo, account, changes, '老師在我的課對應課表'); } catch (e) { console.error(e.message); }
  } finally {
    lock.releaseLock();
  }
  return { ok: true, changed: true, changes: changes };
}

/** 標記／取消「不用訂貨」：date = yyyy-mm-dd，session = 上午／下午／晚上 */
function setNoOrder(token, date, session, course, skip, as, reason) {
  const account = lessonAccount_(token, as), who = auth_(token);
  reason = clean_(reason, 80);
  const d = parseDate_(date);
  if (!d || SESSIONS.indexOf(session) < 0) throw new Error('資料不正確，請重新整理。');
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SHEET.skips, SKIP_HEADERS);
    const t = readTable_(SHEET.skips, ['老師', '日期', '時段']);
    const same = t.rows.filter(function (r) {
      const rd = toDate_(r['日期']);
      return String(r['老師']).trim() === account && rd && dateStr_(rd) === dateStr_(d) && String(r['時段']).trim() === session;
    });
    if (skip && !same.length) {
      ensureHeaders_(sh, SKIP_HEADERS);
      appendRows_(SHEET.skips, [{ '老師': account, '日期': d, '時段': session, '課程': clean_(course, 40), '標記時間': new Date(), '備註': reason ? '不用訂貨原因：' + reason : '', '標記人': who }],
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
    alt: toDate_(r['隔週']) ? dateStr_(toDate_(r['隔週'])) : '',
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
  if (f.alt !== undefined) {   // 週期：每週／隔週（altDate：有上課的那一週，任一天都可以）
    if (f.alt === '隔週') {
      const d = parseDate_(f.altDate);
      if (!d) throw new Error('隔週上課：請選「有上課的那一週」（那週任一天都可以）。');
      out['隔週'] = addDays_(d, -((d.getDay() + 6) % 7));
    } else out['隔週'] = '';
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
  const lock = scriptLock_();
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
  fastReads_();
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
    if ((x.from && dates[i] < x.from) || (x.to && dates[i] > x.to) || altWeekOff_(x.alt, dates[i])) { x.outOfWeek = true; return; }
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
    list.push({ teacher: o.onBehalfOf || o.teacher, date: o.useDate, session: sessionOf_(o.classTime), project: o.project, className: o.className, orderNo: o.orderNo,
      classTime: o.classTime, room: o.room, status: o.status });
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
      else skip[k] = String(r['備註'] || '').replace(/^不用訂貨原因：/, '').trim() || true;   // 有寫原因就帶原因
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
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
  const lock = scriptLock_();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SHEET.schedule, SCHEDULE_HEADERS);
    ensureHeaders_(sh, SCHEDULE_HEADERS);
    set['需要訂貨'] = true;
    set['備註'] = clean_(fields && fields.note, 100);
    appendRows_(SHEET.schedule, [set], { '開始日期': 'yyyy-mm-dd', '結束日期': 'yyyy-mm-dd', '隔週': 'yyyy-mm-dd' });
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

/** 回覆好幾則（文字或卡片），回覆不算則數 */
function lineReplyMsgs_(replyToken, msgs) {
  const lineToken = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!lineToken || !replyToken) return;
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + lineToken },
    payload: JSON.stringify({ replyToken: replyToken, messages: msgs.slice(0, 5) }),
  });
  if (res.getResponseCode() !== 200) console.error('lineReplyMsgs_ ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
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
  let cc = null; try { cc = JSON.parse(props.CREW_CODE || 'null'); } catch (e) {}
  return { teachersBound: bound, quota: lineQuota_(), richMenu: props.RICH_MENU_ID || '',
    crewBound: Object.keys(props).filter(function (k) { return k.indexOf('CL:') === 0; }).length, crewCode: cc && cc.until > Date.now() ? { code: cc.code, until: Utilities.formatDate(new Date(cc.until), tz_(), 'MM/dd HH:mm') } : null };
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
function adminSetupRichMenu(token, adminPng, crewPng, teacherPng) {
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
    { action: { type: 'message', label: '今天的課', text: '今天的課' } },
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
  // 老師版圖片：網頁畫好傳上來的（七彩）；沒有就用舊的 RichMenu 檔
  const tb = String(teacherPng || '').replace(/^data:image\/(png|jpeg);base64,/, ''), tMime = /^data:image\/jpeg/.test(String(teacherPng || '')) ? 'image/jpeg' : 'image/png';
  const png = tb.length > 1000 && tb.length < 1300000 ? Utilities.base64Decode(tb) : Utilities.base64Decode(HtmlService.createHtmlOutputFromFile('RichMenu').getContent().replace(/\s+/g, ''));
  const up = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/' + id + '/content', {
    method: 'post', contentType: tb.length > 1000 && tb.length < 1300000 ? tMime : 'image/png', muteHttpExceptions: true, headers: auth, payload: png,
  });
  if (up.getResponseCode() !== 200) throw new Error('上傳選單圖片失敗：' + up.getContentText().slice(0, 200));
  const def = UrlFetchApp.fetch('https://api.line.me/v2/bot/user/all/richmenu/' + id, { method: 'post', muteHttpExceptions: true, headers: auth });
  if (def.getResponseCode() !== 200) throw new Error('設定預設選單失敗：' + def.getContentText().slice(0, 200));
  // 刪掉舊的選單
  const old = props.getProperty('RICH_MENU_ID');
  if (old && old !== id) UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/' + old, { method: 'delete', muteHttpExceptions: true, headers: auth });
  props.setProperty('RICH_MENU_ID', id);
  if (adminPng) lineSetupAdminMenu_(adminPng);
  if (crewPng) lineSetupCrewMenu_(crewPng);
  return automationStatus_();
}
// ---- 校隊：加官方 LINE、傳「校隊 綁定碼」，選單只有看料單、看課表（唯讀，不能訂貨）----
/** 管理者：產生校隊綁定碼（7 天內有效，全部校隊同學用同一個） */
function adminCrewLineCode(token) {
  auth_(token, true);
  throw new Error('校隊綁定已經改成「每人一組碼」：請在下面的校隊名單新增名字，系統會給每個人自己的碼。');
}
/** 管理者：清掉全部校隊的 LINE（換屆、畢業）。選單換回一般版 */
function adminCrewLineClear(token) {
  auth_(token, true);
  const props = PropertiesService.getScriptProperties(), all = props.getProperties(), lineToken = props.getProperty('LINE_TOKEN');
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('CL:') !== 0) return;
    if (lineToken) { try { UrlFetchApp.fetch('https://api.line.me/v2/bot/user/' + k.slice(3) + '/richmenu', { method: 'delete', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + lineToken } }); } catch (e) {} }
    props.deleteProperty(k);
  });
  props.deleteProperty('CREW_CODE');
  const list = crewRoster_(); list.forEach(function (x) { x.uid = ''; x.t = ''; }); crewRosterSave_(list);   // 名單留著，要用時再給新的碼
  return automationStatus_();
}
const CREW_TEAMS = ['中餐', '烘焙', '飲調', '咖啡', '餐飲服務', '西餐'];
/** 「中餐 王同學」「烘焙校隊 李同學」→ { team, name }（組別認不出來就是空的） */
function crewWho_(s) {
  s = String(s || '').replace(/[，,、:：]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 30);
  let team = '';
  CREW_TEAMS.forEach(function (t) { if (!team && s.indexOf(t) >= 0) team = t; });
  const name = s.replace(/(中餐|烘焙|飲調|咖啡|餐飲服務|西餐)\s*(校隊|組)?/, '').replace(/^我是\s*/, '').trim().slice(0, 12);
  return { team: team, name: name };
}
/** 校隊 LINE：{ t: 綁定時間, team, name }（舊的只存時間） */
function crewLineInfo_(userId) {
  const v = userId ? PropertiesService.getScriptProperties().getProperty('CL:' + userId) : null;
  if (!v) return null;
  try { const o = JSON.parse(v); if (o && typeof o === 'object') return { t: o.t || '', team: o.team || '', name: o.name || '', rid: o.rid || '' }; } catch (e) {}
  return { t: v, team: '', name: '', rid: '' };
}
function crewLineSave_(userId, info) {
  PropertiesService.getScriptProperties().setProperty('CL:' + userId, JSON.stringify({ t: info.t, team: info.team || '', name: info.name || '' }));
}
// ---- 校隊名單：管理者先新增名字，每人一組綁定碼（7 天、用一次），確定是誰綁的 ----
const CREW_CODE_DAYS = 7;
function crewRoster_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('CREW_ROSTER') || '[]') || []; } catch (e) { return []; } }
function crewRosterSave_(list) { PropertiesService.getScriptProperties().setProperty('CREW_ROSTER', JSON.stringify(list)); }
function crewNewCode_(list) {
  const used = {}; list.forEach(function (x) { if (x.code) used[x.code] = true; });
  let c; do { c = String(Math.floor(100000 + Math.random() * 900000)); } while (used[c]);
  return c;
}
function lineCrewBind_(code, src) {
  if (!src || src.type !== 'user' || !src.userId) return '請用個人聊天室（不是群組）傳「校隊 綁定碼」。';
  const lock = scriptLock_(); lock.waitLock(20000);
  try {
    const list = crewRoster_(), now = Date.now();
    const r = list.find(function (x) { return x.code && x.code === code; });
    if (!r || !(r.until > now)) return '校隊綁定碼不正確或已過期，請找管理者拿新的（每個人的碼都不一樣，不能用別人的）。';
    // 這支 LINE 之前綁過別人：先解除那個名字
    list.forEach(function (x) { if (x.uid === src.userId && x.id !== r.id) { x.uid = ''; x.t = ''; } });
    if (r.uid && r.uid !== src.userId) PropertiesService.getScriptProperties().deleteProperty('CL:' + r.uid);   // 同一個名字換手機
    r.uid = src.userId; r.t = Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm'); r.code = ''; r.until = 0;
    crewRosterSave_(list);
    PropertiesService.getScriptProperties().setProperty('CL:' + src.userId, JSON.stringify({ t: r.t, team: r.team, name: r.name, rid: r.id }));
  } finally { lock.releaseLock(); }
  lineLinkMenu_(src.userId, 'CREW_RICH_MENU_ID');
  const me = crewLineInfo_(src.userId);
  return '校隊綁定完成，' + (me.team ? me.team + '校隊 ' : '') + me.name + '！\n下方選單按「🗓 總課表」可以打開網頁，看總課表和每張料單（今天、明天、這週、下週或指定日期都可以）；也可以看今天／明天料單、課表、「🔍 查老師」（選單沒換的話，關掉聊天室再打開）。';
}
/** 組別的共用訂貨帳號：「中餐校隊」「烘焙校隊」…（帳號設定裡有、而且啟用才算） */
function crewTeamAccount_(team) {
  if (!team) return '';
  const name = team + '校隊';
  const r = readTable_(SHEET.accounts, ['帳號', '啟用']).rows.find(function (x) { return String(x['帳號']).trim() === name && isTrue_(x['啟用']); });
  return r ? name : '';
}
/** 名單上這位（LINE userId）能不能用組別帳號訂貨：{ account, label } 或 null */
function crewOrderOf_(uid) {
  const me = crewLineInfo_(uid);
  if (!me || !me.rid) return null;
  const r = crewRoster_().find(function (x) { return x.id === me.rid && x.uid === uid; });
  if (!r || !r.canOrder) return null;
  const acc = crewTeamAccount_(r.team);
  return acc ? { account: acc, label: r.team + '校隊 ' + r.name, name: r.name } : null;
}
/** LINE 校隊訂貨連結（10 分鐘、用一次）：打開就登入組別共用帳號，訂單記學生名字 */
function crewOrderLink_(uid) {
  const code = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  CacheService.getScriptCache().put('colink:' + code, uid, LINE_LOGIN_MINUTES * 60);
  const u = lineUrl_('');
  return u ? u + '&co=' + code : '';
}
function loginByCrewLink(code) {
  code = String(code || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 64);
  const cache = CacheService.getScriptCache(), uid = code ? cache.get('colink:' + code) : null;
  if (!uid) throw new Error('這個訂貨連結已經過期或用過了，請在 LINE 再按一次「🛒 校隊訂貨」。');
  cache.remove('colink:' + code);
  const c = crewOrderOf_(uid);
  if (!c) throw new Error('你沒有訂貨權限（或組別帳號停用了），請找管理者。');
  const token = Utilities.getUuid();
  cache.put('session:' + token, c.account, SESSION_HOURS * 3600);
  cache.put('sessstu:' + token, c.name, SESSION_HOURS * 3600);
  try { logLogin_(c.account, '成功（LINE 校隊：' + c.label + '）'); } catch (e) {}
  CUR_STUDENT_ = c.name;
  const data = getBootstrap(token);
  data.student = c.name;
  return { token: token, account: c.account, isAdmin: false, mustChange: false, data: data };
}
/** 管理者：校隊名單（每人狀態：等綁定＋碼／已綁定）＋沒在名單上、以前用共用碼綁的 */
function adminCrewRoster(token) {
  auth_(token, true);
  const now = Date.now(), list = crewRoster_(), props = PropertiesService.getScriptProperties().getProperties();
  const out = list.map(function (x) {
    return { id: x.id, team: x.team, name: x.name, bound: !!(x.uid && props['CL:' + x.uid]), t: x.t || '', canOrder: !!x.canOrder,
      code: x.code && x.until > now ? x.code : '', until: x.code && x.until > now ? Utilities.formatDate(new Date(x.until), tz_(), 'MM/dd') : '', expired: !!(x.code && !(x.until > now)) };
  });
  const inRoster = {}; list.forEach(function (x) { if (x.uid) inRoster[x.uid] = true; });
  Object.keys(props).forEach(function (k) {
    if (k.indexOf('CL:') !== 0 || inRoster[k.slice(3)]) return;
    const me = crewLineInfo_(k.slice(3));
    out.push({ id: 'L:' + k.slice(3), team: me.team, name: me.name || '（沒有名字）', bound: true, t: me.t, legacy: true });
  });
  return out.sort(function (a, b) { return (CREW_TEAMS.indexOf(a.team) + 99 * !a.team) - (CREW_TEAMS.indexOf(b.team) + 99 * !b.team) || a.name.localeCompare(b.name, 'zh-Hant'); });
}
/**
 * 管理者：名單操作。action：'add'（names 可以好幾個，「、」隔開）、'code'（重新產生這個人的碼）、'save'（改組別、名字）、'remove'（移除，LINE 也解除）
 */
function adminCrewRosterEdit(token, action, id, team, names) {
  auth_(token, true);
  if (CREW_TEAMS.indexOf(String(team || '')) < 0) team = '';
  const lock = scriptLock_(); lock.waitLock(20000);
  try {
    const list = crewRoster_(), until = Date.now() + CREW_CODE_DAYS * 86400000, props = PropertiesService.getScriptProperties();
    const unlink = function (uid) {
      if (!uid) return;
      const lineToken = props.getProperty('LINE_TOKEN');
      if (lineToken) { try { UrlFetchApp.fetch('https://api.line.me/v2/bot/user/' + uid + '/richmenu', { method: 'delete', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + lineToken } }); } catch (e) {} }
      props.deleteProperty('CL:' + uid);
    };
    if (action === 'import') {
      // 一次貼整份名單：每一行「組別 名字 名字…」或「組別：名字、名字」；同組已經有的跳過，不會重複
      let cur = '', added = 0, skipped = 0;
      String(names || '').split(/\n+/).forEach(function (line) {
        line = String(line).replace(/[（(][^）)]*[）)]/g, ' ').replace(/[，,、:：\t]/g, ' ').replace(/\s+/g, ' ').trim();
        if (!line) return;
        const tm = CREW_TEAMS.filter(function (t) { return line.indexOf(t) === 0; }).sort(function (a, b) { return b.length - a.length; })[0];
        if (tm) { cur = tm; line = line.slice(tm.length).replace(/^\s*(校隊|組)/, '').trim(); }
        if (!cur) throw new Error('第一行要先寫組別（' + CREW_TEAMS.join('、') + '）。');
        line.split(' ').map(function (n) { return clean_(n.replace(/^\d+$/, ''), 12); }).filter(Boolean).forEach(function (n) {
          if (list.some(function (x) { return x.team === cur && x.name === n; })) { skipped++; return; }
          list.push({ id: Utilities.getUuid().slice(0, 8), team: cur, name: n, code: crewNewCode_(list), until: until, uid: '', t: '' });
          added++;
        });
      });
      if (!added && !skipped) throw new Error('沒有讀到名字。');
      crewRosterSave_(list);
      return { list: adminCrewRoster(token), imported: { added: added, skipped: skipped } };
    }
    if (action === 'add') {
      if (!team) throw new Error('請選組別。');
      const ns = String(names || '').split(/[、,，\s\n]+/).map(function (n) { return clean_(n, 12); }).filter(Boolean);
      if (!ns.length) throw new Error('請輸入名字。');
      ns.forEach(function (n) {
        if (list.some(function (x) { return x.team === team && x.name === n; })) throw new Error(team + '校隊已經有「' + n + '」了。');
        list.push({ id: Utilities.getUuid().slice(0, 8), team: team, name: n, code: crewNewCode_(list), until: until, uid: '', t: '' });
      });
    } else if (String(id || '').indexOf('L:') === 0) {   // 以前用共用碼綁的
      const uid = String(id).slice(2).replace(/[^A-Za-z0-9]/g, '');
      if (action === 'remove') unlink(uid);
      else if (action === 'save') { const me = crewLineInfo_(uid); if (me) crewLineSave_(uid, { t: me.t, team: team, name: clean_(names, 12) }); }
      else throw new Error('以前用共用碼綁的請先移除，再加到名單、給新的碼。');
    } else {
      const r = list.find(function (x) { return x.id === id; });
      if (!r) throw new Error('找不到這位校隊，請重新整理。');
      if (action === 'code') { r.code = crewNewCode_(list); r.until = until; }
      else if (action === 'order-on' || action === 'order-off') {
        if (action === 'order-on' && !crewTeamAccount_(r.team)) throw new Error('帳號設定裡沒有「' + r.team + '校隊」這個帳號（或停用了），請先建立，校隊才能用它訂貨。');
        r.canOrder = action === 'order-on';
      }
      else if (action === 'save') {
        r.team = team || r.team; r.name = clean_(names, 12) || r.name;
        if (r.uid && props.getProperty('CL:' + r.uid)) props.setProperty('CL:' + r.uid, JSON.stringify({ t: r.t, team: r.team, name: r.name, rid: r.id }));
      } else if (action === 'remove') { unlink(r.uid); list.splice(list.indexOf(r), 1); }
      else throw new Error('動作不正確。');
    }
    crewRosterSave_(list);
  } finally { lock.releaseLock(); }
  return adminCrewRoster(token);
}
/** 校隊改名字、組別：「我是 中餐 王同學」 */
function lineCrewName_(text, userId) {
  const me = crewLineInfo_(userId);
  if (!me) return null;
  if (me.rid) return '你是「' + (me.team ? me.team + '校隊 ' : '') + me.name + '」，名字由管理者設定；要改請找管理者。';
  const w = crewWho_(text);
  if (!w.name && !w.team) return '請傳「我是 中餐 王同學」（組別：中餐／烘焙／飲調／西餐）。';
  crewLineSave_(userId, { t: me.t, team: w.team || me.team, name: w.name || me.name });
  const n = crewLineInfo_(userId);
  return '已登記：' + (n.team ? n.team + '校隊 ' : '') + (n.name || '（沒有名字）') + '。' + (n.team ? '' : '\n組別還沒登記，請傳「我是 中餐 ' + (n.name || '王同學') + '」。');
}
/** 校隊網頁查詢的連結（12 小時內可以重複打開；管理者移除這位校隊就失效） */
const CREW_VIEW_HOURS = 12;
function crewViewLink_(userId) {
  const code = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  CacheService.getScriptCache().put('cvlink:' + code, userId, CREW_VIEW_HOURS * 3600);
  const u = lineUrl_('');
  return u ? u + '&cv=' + code : '';
}
/** 網頁：校隊用 LINE 連結打開查詢頁（不用帳號密碼，只能看） */
function crewViewLogin(code) {
  code = String(code || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 64);
  const uid = code ? CacheService.getScriptCache().get('cvlink:' + code) : null;
  const me = uid ? crewLineInfo_(uid) : null;
  if (!me) throw new Error('這個連結已經過期（' + CREW_VIEW_HOURS + ' 小時），或你已經不是校隊了。請在 LINE 再按一次「🗓 總課表」。');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('cview:' + token, uid, CREW_VIEW_HOURS * 3600);
  return { token: token, team: me.team, name: me.name, label: (me.team ? me.team + '校隊 ' : '校隊 ') + (me.name || '') };
}
/** 誰可以看校隊查詢頁：LINE 校隊（cview）、校隊帳號、管理者 */
function crewViewer_(token) {
  const cache = CacheService.getScriptCache(), t = String(token || '');
  const uid = t ? cache.get('cview:' + t) : null;
  if (uid) {
    const me = crewLineInfo_(uid);
    if (!me) throw new Error('你已經不是校隊了（管理者移除了），請找管理者。');
    return (me.team ? me.team + '校隊 ' : '校隊 ') + (me.name || '');
  }
  const account = t ? cache.get('session:' + t) : null;
  if (!account) throw new Error('SESSION_EXPIRED');
  if (!isAdmin_(account) && !isCrew_(account)) throw new Error('只有校隊和管理者可以看。');
  return account;
}
/**
 * 校隊查詢：一週（週一～週五）每堂課、叫料了沒、每張料單的食材。mon：那週的週一 'yyyy-MM-dd'
 * 只能看，不回傳成本、單價。
 */
function crewWeek(token, mon, to) {
  fastReads_();
  const who = crewViewer_(token);
  let m = parseDate_(mon) || today_(), fri;
  const t2 = parseDate_(to);
  if (t2) {   // 指定日期：從 mon 到 to（最多 31 天）
    if (t2 < m) throw new Error('結束日期要在開始日期之後。');
    if ((t2 - m) / 86400000 > 31) throw new Error('一次最多查 31 天，請把日期縮短一點。');
    fri = t2;
  } else {
    m = addDays_(m, -((m.getDay() + 6) % 7));
    fri = addDays_(m, 4);
  }
  // 科別（中餐／烘焙／飲調）：有料單照料單的科別，沒有就照教室、課程名稱猜
  const deptOf = function (room, course, os) {
    const d = (os || []).map(function (o) { return String(o.dept || ''); }).filter(Boolean)[0];
    if (d) return /烘焙|西點/.test(d) ? '烘焙' : /飲|調|咖啡/.test(d) ? '飲調' : /中餐|中式|西餐|餐/.test(d) ? '中餐' : d;
    const t = String(room || '') + ' ' + String(course || '');
    return /烘|麵包|西點|蛋糕|點心/.test(t) ? '烘焙' : /飲調|飲料|調酒|咖啡|吧檯|茶|調[一二三四五六]|迦/.test(t) ? '飲調' : /中|餐|廚|烹|料理/.test(t) ? '中餐' : '';   // 「烹調」不算飲調
  };
  const lessons = lessonsWithStatus_(m, fri, null).map(function (x) {
    return { date: x.date, session: x.session, teacher: x.teacher, className: x.className || '', course: x.course || '', room: x.room || '', status: x.status, holiday: x.holiday || '',
      dept: deptOf(x.room, x.course, x.orders), orders: (x.orders || []).map(function (o) { return o.orderNo; }) };
  });
  const orders0 = ordersBetween_(m, fri).filter(function (o) { return o.status !== '已取消'; });
  const ded = deductNotes_(orders0);   // 扣貨扣掉的：旁邊標「扣貨已扣掉」，校隊才不會以為廠商沒送
  const orders = orders0.map(function (o) {
    return { orderNo: o.orderNo, date: o.useDate, session: o.classTime, teacher: o.onBehalfOf || o.teacher, className: o.className || '', project: o.project || '', groups: o.groups || '', room: o.room || '', status: o.status, dept: o.dept || '',
      items: o.items.filter(function (i) { return i.status !== '已取消' && i.name; }).map(function (i) { return { name: i.name, qty: Math.round(Number(i.qty) * 100) / 100, unit: i.unit || '', vendor: i.vendor || '', noOrder: !!i.noOrder, tool: i.id === TOOL_ID, note: i.note || '', ded: ded[o.orderNo + '|' + i.name] || '' }; }) };
  });
  return { who: who, mon: dateStr_(m), fri: dateStr_(fri), lessons: lessons, orders: orders, holidays: holidaysBetween_(m, fri) };
}
function lineIsCrew_(userId) { return !!(userId && PropertiesService.getScriptProperties().getProperty('CL:' + userId)); }
function lineLinkMenu_(userId, key) {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty(key), lineToken = props.getProperty('LINE_TOKEN');
  if (!id || !lineToken || !userId) return;
  try { UrlFetchApp.fetch('https://api.line.me/v2/bot/user/' + userId + '/richmenu/' + id, { method: 'post', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + lineToken } }); } catch (e) {}
}
/** 校隊選單（2×2）：今天料單、明天料單、今天課表、明天課表 */
function lineSetupCrewMenu_(png) {
  const props = PropertiesService.getScriptProperties(), lineToken = props.getProperty('LINE_TOKEN');
  const b64 = String(png || '').replace(/^data:image\/(png|jpeg);base64,/, ''), mime = /^data:image\/jpeg/.test(String(png || '')) ? 'image/jpeg' : 'image/png';   // 七彩的用 JPG（LINE 選單圖片最大 1 MB）
  if (!lineToken || b64.length < 1000 || b64.length > 1300000) return;
  const auth = { Authorization: 'Bearer ' + lineToken };
  const w = 2500, h = 1686, cw = Math.floor(w / 3), rh = h / 2;
  const areas = [['今天料單', '今天料單'], ['明天料單', '明天料單'], ['查老師', '查老師'], ['今天課表', '今天課表'], ['明天課表', '明天課表'], ['總課表', '總課表']].map(function (a, i) {
    const c = i % 3, r = Math.floor(i / 3);
    return { bounds: { x: c * cw, y: r * rh, width: c === 2 ? w - 2 * cw : cw, height: rh }, action: { type: 'message', label: a[0], text: a[1] } };
  });
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu', { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: auth,
    payload: JSON.stringify({ size: { width: w, height: h }, selected: true, name: '訂貨系統校隊選單', chatBarText: '校隊選單', areas: areas }) });
  if (res.getResponseCode() !== 200) throw new Error('建立校隊選單失敗：' + res.getContentText().slice(0, 200));
  const id = JSON.parse(res.getContentText()).richMenuId;
  const up = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/' + id + '/content', { method: 'post', contentType: mime, muteHttpExceptions: true, headers: auth, payload: Utilities.base64Decode(b64) });
  if (up.getResponseCode() !== 200) throw new Error('上傳校隊選單圖片失敗：' + up.getContentText().slice(0, 200));
  const old = props.getProperty('CREW_RICH_MENU_ID');
  if (old && old !== id) UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/' + old, { method: 'delete', muteHttpExceptions: true, headers: auth });
  props.setProperty('CREW_RICH_MENU_ID', id);
  const all = props.getProperties();
  Object.keys(all).forEach(function (k) { if (k.indexOf('CL:') === 0) lineLinkMenu_(k.slice(3), 'CREW_RICH_MENU_ID'); });
}
/** 管理者專用選單（2×3）：補訂貨、訂貨系統、驗收系統／查看課表、查看料單、查看叫貨單。圖片在瀏覽器畫好傳上來 */
function lineSetupAdminMenu_(png) {
  const props = PropertiesService.getScriptProperties(), lineToken = props.getProperty('LINE_TOKEN');
  const b64 = String(png || '').replace(/^data:image\/(png|jpeg);base64,/, ''), mime = /^data:image\/jpeg/.test(String(png || '')) ? 'image/jpeg' : 'image/png';   // 七彩的用 JPG（LINE 選單圖片最大 1 MB）
  if (!lineToken || b64.length < 1000 || b64.length > 1300000) return;
  const auth = { Authorization: 'Bearer ' + lineToken };
  const w = 2500, h = 1686, cw = Math.floor(w / 3), rh = h / 2;
  const areas = [
    { action: { type: 'message', label: '還沒叫料的老師', text: '還沒叫料' } },
    { action: { type: 'message', label: '下週訂貨進度', text: '訂貨進度' } },
    { action: { type: 'message', label: '驗證碼', text: '驗證碼' } },
    { action: { type: 'message', label: '訂貨網', text: '訂貨系統' } },
    { action: { type: 'message', label: '驗收網', text: '驗收系統' } },
    { action: { type: 'message', label: '快速補訂貨', text: '補訂貨' } },
  ].map(function (a, i) { const c = i % 3, r = Math.floor(i / 3); a.bounds = { x: c * cw, y: r * rh, width: c === 2 ? w - 2 * cw : cw, height: rh }; return a; });
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu', { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: auth,
    payload: JSON.stringify({ size: { width: w, height: h }, selected: true, name: '訂貨系統管理者選單', chatBarText: '管理者選單', areas: areas }) });
  if (res.getResponseCode() !== 200) throw new Error('建立管理者選單失敗：' + res.getContentText().slice(0, 200));
  const id = JSON.parse(res.getContentText()).richMenuId;
  const up = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/' + id + '/content', { method: 'post', contentType: mime, muteHttpExceptions: true, headers: auth, payload: Utilities.base64Decode(b64) });
  if (up.getResponseCode() !== 200) throw new Error('上傳管理者選單圖片失敗：' + up.getContentText().slice(0, 200));
  const old = props.getProperty('ADMIN_RICH_MENU_ID');
  if (old && old !== id) UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/' + old, { method: 'delete', muteHttpExceptions: true, headers: auth });
  props.setProperty('ADMIN_RICH_MENU_ID', id);
  // 已經綁定 LINE 的管理者，選單換成管理者版
  const all = props.getProperties();
  Object.keys(all).forEach(function (k) { if (k.indexOf('LU:') === 0 && isAdmin_(k.slice(3))) lineLinkAdminMenu_(all[k]); });
}
function lineLinkAdminMenu_(userId) {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('ADMIN_RICH_MENU_ID'), lineToken = props.getProperty('LINE_TOKEN');
  if (!id || !lineToken || !userId) return;
  try { UrlFetchApp.fetch('https://api.line.me/v2/bot/user/' + userId + '/richmenu/' + id, { method: 'post', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + lineToken } }); } catch (e) {}
}

// ---- 管理者在 LINE 查：課表、料單、補訂貨（卡片，回覆不算則數）----
const FX_BRAND = '#8a2b1a', FX_MUTED = '#78716c', FX_BAD = '#b42318', FX_OK = '#067647';
function fxT_(t, o) { return Object.assign({ type: 'text', text: String(t == null || t === '' ? ' ' : t).slice(0, 2000), wrap: true }, o || {}); }
function fxB_(title, sub, rows, footer) {
  const b = { type: 'bubble', size: 'mega',
    header: { type: 'box', layout: 'vertical', backgroundColor: FX_BRAND, paddingAll: '14px', contents: [fxT_(title, { color: '#ffffff', weight: 'bold', size: 'lg' })].concat(sub ? [fxT_(sub, { color: '#ffffffcc', size: 'sm', margin: 'xs' })] : []) },
    body: { type: 'box', layout: 'vertical', spacing: 'md', paddingAll: '14px', contents: rows.length ? rows : [fxT_('（沒有）', { color: FX_MUTED })] } };
  if (footer && footer.length) b.footer = { type: 'box', layout: footer.length <= 3 ? 'horizontal' : 'vertical', spacing: 'xs', paddingAll: '8px', contents: footer };
  return b;
}
function fxBtn_(label, text) { return { type: 'button', style: 'link', height: 'sm', action: { type: 'message', label: String(label).slice(0, 20), text: text } }; }
function fxFlexMsg_(alt, bubbles) { return { type: 'flex', altText: String(alt).slice(0, 380), contents: bubbles.length === 1 ? bubbles[0] : { type: 'carousel', contents: bubbles.slice(0, 12) } }; }
/** 「今天」「明天」「11/12」「2026-11-12」→ Date；沒寫＝今天 */
function lineDateArg_(s) {
  s = String(s || '').trim();
  const t = today_();
  if (!s || /^(今天|今日)$/.test(s)) return t;
  if (/^(明天|明日)$/.test(s)) return addDays_(t, 1);
  if (s === '後天') return addDays_(t, 2);
  if (/^(昨天|昨日)$/.test(s)) return addDays_(t, -1);
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (m) return new Date(+m[1], m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[\/.月-](\d{1,2})日?$/);
  if (m) { const d = new Date(t.getFullYear(), m[1] - 1, +m[2]); return d < addDays_(t, -180) ? new Date(t.getFullYear() + 1, m[1] - 1, +m[2]) : d; }
  return null;
}
let LINE_VIEW_CREW_ = false;   // 這次回覆是給校隊看（快速按鈕不放叫貨單）
/** 老師名單（帳號設定裡啟用的老師，不含管理者、校隊） */
function lineTeacherNames_() {
  return readTable_(SHEET.accounts, ['帳號', '啟用']).rows
    .filter(function (r) { return isTrue_(r['啟用']); })
    .map(function (r) { return String(r['帳號']).trim(); })
    .filter(function (a) { return a && ADMIN_ACCOUNTS.indexOf(a) < 0 && crewAccounts_().indexOf(a) < 0; });
}
/** 🔍 查老師：這週＋下週每一堂課（叫料了沒、做什麼）；沒打名字就列這兩週有課的老師讓他點 */
function lineTeacherMsgs_(q) {
  const t0 = today_(), mon = addDays_(t0, -((t0.getDay() + 6) % 7)), fri2 = addDays_(mon, 11);
  const all = lessonsWithStatus_(mon, fri2, null);
  const names = lineTeacherNames_();
  all.forEach(function (x) { if (x.teacher && names.indexOf(x.teacher) < 0) names.push(x.teacher); });
  const qr = function (list) { return { items: list.slice(0, 13).map(function (n) { return { type: 'action', action: { type: 'message', label: String(n).slice(0, 20), text: '老師 ' + n } }; }) }; };
  if (!q) {
    const busy = {};
    all.forEach(function (x) { if (x.teacher) busy[x.teacher] = (busy[x.teacher] || 0) + 1; });
    const list = Object.keys(busy).sort(function (a, b) { return busy[b] - busy[a] || a.localeCompare(b); });
    if (!list.length) return [{ type: 'text', text: '這週和下週沒有課。' }];
    return [{ type: 'text', text: '要查哪位老師？點下面的名字，或直接打「老師 王小明」。\n這週和下週有課的老師：' + list.join('、'), quickReply: qr(list) }];
  }
  let hit = names.filter(function (n) { return n === q; });
  if (!hit.length) hit = names.filter(function (n) { return n.indexOf(q) >= 0 || q.indexOf(n) >= 0; });
  if (!hit.length) return [{ type: 'text', text: '找不到「' + q + '」老師，名字要跟訂貨系統的帳號一樣。打「查老師」可以看名單。' }];
  if (hit.length > 1) return [{ type: 'text', text: '有好幾位：' + hit.join('、') + '，請點一位。', quickReply: qr(hit) }];
  const name = hit[0];
  const list = all.filter(function (x) { return x.teacher === name; })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session); });
  const ST = { ordered: ['✓ 有料單', FX_OK], skipped: ['不用訂', FX_MUTED], holiday: ['放假', FX_MUTED], todo: ['✗ 還沒訂', FX_BAD] };
  const nextMon = dateStr_(addDays_(mon, 7));
  const rows = [];
  ['這週', '下週'].forEach(function (lab, wi) {
    const part = list.filter(function (x) { return wi ? x.date >= nextMon : x.date < nextMon; });
    rows.push(fxT_(lab + '（' + shortDateTw_(addDays_(mon, wi * 7)) + ' 起）', { size: 'xs', color: FX_MUTED, weight: 'bold' }));
    if (!part.length) { rows.push(fxT_('沒有課', { size: 'sm', color: FX_MUTED })); return; }
    part.forEach(function (x) {
      const st = ST[x.status] || ['', FX_MUTED];
      rows.push({ type: 'box', layout: 'horizontal', contents: [
        { type: 'box', layout: 'vertical', flex: 6, contents: [fxT_(shortDateTw_(parseDate_(x.date)) + ' ' + (x.session === '晚上' ? '夜間' : x.session) + '　' + (x.className || ''), { size: 'sm', weight: 'bold' }),
          fxT_(x.course + (x.orders && x.orders.length ? '「' + x.orders.map(function (o) { return o.project; }).join('、') + '」' : '') + (x.holiday ? '（' + x.holiday + '）' : ''), { size: 'xs', color: FX_MUTED })] },
        fxT_(st[0], { size: 'sm', color: st[1], align: 'end', flex: 3, gravity: 'center', weight: x.status === 'todo' ? 'bold' : 'regular' })] });
    });
    if (!wi) rows.push({ type: 'separator' });
  });
  const todo = list.filter(function (x) { return x.status === 'todo'; }).length;
  return [fxFlexMsg_(name + ' 老師：這週和下週 ' + list.length + ' 堂課', [fxB_('👩‍🍳 ' + name + ' 老師', '這週＋下週 ' + list.length + ' 堂課' + (todo ? '・' + todo + ' 堂還沒訂' : ''), rows,
    [fxBtn_('看他的料單（7 天內）', '料單 ' + name), fxBtn_('查別的老師', '查老師')])])];
}
function lineAdminReply_(text, userId) {
  const t = String(text || '').trim();
  if (/^(❗\s*)?(還沒叫料|未叫料|還沒訂貨|沒叫料的老師|還沒叫料的老師)(\s*下下週)?$/.test(t) && isAdmin_(lineAccountOf_(userId))) return lineTodoTeachersMsgs_(/下下週/.test(t));
  if (/^(📊\s*)?(下週)?訂貨進度(\s*下下週)?$/.test(t) && isAdmin_(lineAccountOf_(userId))) return lineWeekProgressMsgs_(/下下週/.test(t));
  const m = t.match(/^(今天|明天|後天|昨天)?\s*(課表|料單|叫貨單)\s*(.{0,20})$/);
  const help = /^管理者說明$/.test(t), go = /^(補訂貨|訂貨系統)$/.test(t);
  if (t === '驗收系統' && isAdmin_(lineAccountOf_(userId))) {
    const ru = PropertiesService.getScriptProperties().getProperty('RECEIVE_URL'), acc = lineAccountOf_(userId);
    if (!ru) return [{ type: 'text', text: '進貨驗收網站還沒部署（到訂貨網「⟳ 更新程式」按「更新進貨驗收」）。' }];
    const code = rcvLoginCode_(acc);
    return [fxFlexMsg_('驗收系統：按按鈕直接登入', [fxB_('✅ 驗收系統', '不用打帳號密碼，直接登入成「' + acc + '」',
      [fxT_('打開後就是登入好的進貨驗收（驗收、銷貨單、核銷）。', { size: 'sm' }), fxT_('⚠ 這個連結 ' + LINE_LOGIN_MINUTES + ' 分鐘內有效、只能用一次，不要轉傳給別人。', { size: 'xs', color: FX_BAD })],
      [{ type: 'button', style: 'primary', color: '#475467', action: { type: 'uri', label: '打開（直接登入）', uri: ru + '?openExternalBrowser=1&lt=' + code } }])])];
  }
  // 🔍 查老師：這週＋下週有沒有課、叫料了沒（校隊、管理者都可以）
  const tq = t.match(/^(?:🔍\s*)?(?:查老師|查詢老師|找老師|老師)\s*(.{0,12})$/) || t.match(/^(?:查\s*)?([一-鿿]{2,4})\s*老師$/);
  const crewHelp = /^校隊說明$/.test(t);
  // 📱 網頁查詢：校隊打開查詢網頁（上週／這週／下週的課表、每張料單）
  if (/^(📱\s*|🗓\s*)?(網頁查詢|網頁|查詢網頁|總課表|看總課表|課表網頁)$/.test(t)) {
    const a1 = lineAccountOf_(userId);
    if (!lineIsCrew_(userId)) return a1 && isAdmin_(a1) ? [{ type: 'text', text: '「網頁查詢」是給校隊用的；管理者請用「訂貨系統」打開網站看總課表、訂單彙整。' }] : null;
    const link = crewViewLink_(userId), me = crewLineInfo_(userId);
    if (!link) return [{ type: 'text', text: '訂貨網站還沒部署。' }];
    return [fxFlexMsg_('校隊總課表', [fxB_('🗓 總課表', (me.team ? me.team + '校隊 ' : '') + (me.name || ''),
      [fxT_('按下面的按鈕打開總課表：一天一欄，看每堂課的老師、班級、教室、有沒有叫料；可以選中餐／烘焙／飲調、老師、日期，點一堂課看整張料單。', { size: 'sm', wrap: true }),
       fxT_('連結 ' + CREW_VIEW_HOURS + ' 小時內有效，過了再按一次選單就好。只能看，不能訂貨。', { size: 'xs', color: FX_MUTED, wrap: true })].concat(me.name && me.team ? [] : [fxT_('⚠ 還沒登記名字或組別：請傳「我是 中餐 王同學」。', { size: 'xs', color: FX_BAD, wrap: true })]),
      [{ type: 'button', style: 'primary', color: '#8a2b1a', action: { type: 'uri', label: '🗓 打開總課表', uri: link } }].concat(crewOrderOf_(userId) ? [{ type: 'button', style: 'secondary', action: { type: 'message', label: '🛒 校隊訂貨', text: '校隊訂貨' } }] : []))])];
  }
  // 🛒 校隊訂貨：名單上打勾的學生，用組別共用帳號訂貨（訂單記學生名字）
  if (/^(🛒\s*)?校隊訂貨$/.test(t)) {
    if (!lineIsCrew_(userId)) return null;
    const c = crewOrderOf_(userId);
    if (!c) return [{ type: 'text', text: '你沒有訂貨權限，要訂貨請找管理者在校隊名單幫你打勾。' }];
    const link = crewOrderLink_(userId);
    if (!link) return [{ type: 'text', text: '訂貨網站還沒部署。' }];
    return [fxFlexMsg_('校隊訂貨', [fxB_('🛒 校隊訂貨', c.label, [fxT_('打開後直接登入「' + c.account + '」，送出的訂單會記上你的名字（' + c.name + '）。', { size: 'sm', wrap: true }),
      fxT_('⚠ 連結 ' + LINE_LOGIN_MINUTES + ' 分鐘內有效、只能用一次，不要轉傳給別人。', { size: 'xs', color: FX_BAD, wrap: true })],
      [{ type: 'button', style: 'primary', color: '#b45309', action: { type: 'uri', label: '打開訂貨網（直接登入）', uri: link } }])])];
  }
  if (tq || crewHelp) {
    const acc0 = lineAccountOf_(userId);
    if (!(acc0 && isAdmin_(acc0)) && !lineIsCrew_(userId)) return tq && /^(查老師|查詢老師|找老師)/.test(t) ? [{ type: 'text', text: '這個功能要先綁定（校隊請找管理者拿「校隊綁定碼」）。' }] : null;
    if (crewHelp) return [{ type: 'text', text: '【校隊 LINE 功能】\n・🗓 總課表：打開網頁看上週、這週、下週的課表和每張料單（最好用）。\n・🔍 查老師：點老師名字，看他這週和下週每一堂課、叫料了沒、做什麼。也可以直接打「老師 王小明」。\n・今天料單／明天料單：每張料單的食材和數量，左右滑動看。也可以打「料單 11/20」「料單 王小明」。\n・今天課表／明天課表：每堂課哪位老師、哪一班。也可以打「課表 11/20」。\n只能看，不能訂貨。' }];
    LINE_VIEW_CREW_ = !(acc0 && isAdmin_(acc0));
    return lineTeacherMsgs_(String(tq[1] || '').replace(/老師$/, '').trim());
  }
  if (!m && !help && !go) return null;
  const account = lineAccountOf_(userId);
  // 校隊：只能看料單、課表（唯讀）
  if ((!account || !isAdmin_(account)) && lineIsCrew_(userId) && m && (m[2] === '課表' || m[2] === '料單')) LINE_VIEW_CREW_ = true;
  else if (!account || !isAdmin_(account)) return m || go && t === '補訂貨' ? [{ type: 'text', text: '這個功能只有管理者可以用。' }] : null;
  else LINE_VIEW_CREW_ = false;
  if (go) {
    const link = lineLoginLink_(account, t === '補訂貨' ? 'quick' : '');
    if (!link) return [{ type: 'text', text: '訂貨網站還沒部署。' }];
    return [fxFlexMsg_(t + '：按按鈕直接登入', [fxB_(t === '補訂貨' ? '➕ 補訂貨' : '🛒 訂貨系統', '不用打帳號密碼，直接登入成「' + account + '」',
      [fxT_(t === '補訂貨' ? '打開是很簡單的補訂貨頁：選送貨日期、打食材、數量，按送出，就加到那天那家廠商的叫貨單。' : '打開後就是登入好的訂貨系統。', { size: 'sm' }),
        fxT_('⚠ 這個連結 ' + LINE_LOGIN_MINUTES + ' 分鐘內有效、只能用一次，不要轉傳給別人。', { size: 'xs', color: FX_BAD })],
      [{ type: 'button', style: 'primary', color: FX_BRAND, action: { type: 'uri', label: '打開（直接登入）', uri: link } }])])];
  }
  if (help) return [{ type: 'text', text: '【管理者 LINE 功能】\n・補訂貨、訂貨系統：回一個登入連結，按了直接登入（10 分鐘內、用一次）。\n・驗收系統：一樣回一個登入連結，按了直接登入進貨驗收。\n・叫貨單：今天送貨的各廠商叫貨內容，也可以傳「叫貨單 11/12」。\n・今天課表／明天課表：每堂課哪位老師、哪一班、叫料了沒。也可以傳「課表 11/12」查某一天。\n・今天料單／明天料單：每張料單的食材和數量，左右滑動看。也可以傳「料單 11/12」或「料單 王小明」（那位老師接下來 7 天的）。\n・補訂貨：直接打開訂貨網站下單（代訂、補訂）。\n這些都是回覆，不花 LINE 則數。' }];
  const arg = (m[3] || '').trim();
  if (m[2] === '課表') {
    const d = lineDateArg_(m[1] || arg);
    if (!d) return [{ type: 'text', text: '日期看不懂，請傳「課表 11/12」這樣的格式。' }];
    return lineLessonsMsgs_(d);
  }
  if (m[2] === '叫貨單') {
    const dd = lineDateArg_(m[1] || arg);
    if (!dd) return [{ type: 'text', text: '日期看不懂，請傳「叫貨單 11/12」這樣的格式。' }];
    return linePurchaseMsgs_(dd);
  }
  const d = lineDateArg_(m[1] || arg);
  if (d) return lineOrdersMsgs_(d, d, '');
  return lineOrdersMsgs_(today_(), addDays_(today_(), 6), arg);   // 料單 王小明
}
function lineLessonsMsgs_(d) {
  const ds = dateStr_(d);
  const nav = [fxBtn_('◀ 前一天', '課表 ' + dateStr_(addDays_(d, -1))), fxBtn_('這天料單', '料單 ' + ds), fxBtn_('後一天 ▶', '課表 ' + dateStr_(addDays_(d, 1)))];
  const list = lessonsWithStatus_(d, d, null).sort(function (a, b) { return SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session) || String(a.teacher).localeCompare(String(b.teacher)); });
  const title = '📅 ' + shortDateTw_(d) + ' 課表';
  if (!list.length) return [fxFlexMsg_(title + '：沒有課', [fxB_(title, holidayOf_(ds) ? '放假：' + holidayOf_(ds) : '這天沒有課', [], nav)])];
  const ST = { ordered: ['✓ 已叫料', FX_OK], skipped: ['不用訂', FX_MUTED], holiday: ['放假', FX_MUTED], todo: ['✗ 未叫料', FX_BAD] };
  const todo = list.filter(function (x) { return x.status === 'todo'; }).length;
  const rows = [];
  SESSIONS.forEach(function (ses) {
    const part = list.filter(function (x) { return x.session === ses; });
    if (!part.length) return;
    if (rows.length) rows.push({ type: 'separator' });
    rows.push(fxT_(ses === '晚上' ? '夜間' : ses, { size: 'xs', color: FX_MUTED, weight: 'bold' }));
    part.forEach(function (x) {
      const st = ST[x.status] || ['', FX_MUTED];
      rows.push({ type: 'box', layout: 'horizontal', contents: [
        { type: 'box', layout: 'vertical', flex: 6, contents: [fxT_(x.teacher + '　' + (x.className || ''), { size: 'sm', weight: 'bold' }), fxT_(x.course + (x.orders && x.orders.length ? '「' + x.orders.map(function (o) { return o.project; }).join('、') + '」' : ''), { size: 'xs', color: FX_MUTED })] },
        fxT_(st[0], { size: 'sm', color: st[1], align: 'end', flex: 3, gravity: 'center', weight: x.status === 'todo' ? 'bold' : 'regular' })] });
    });
  });
  if (rows.length > 60) { rows.length = 60; rows.push(fxT_('…課太多了，其他請到網站「總課表」看。', { size: 'xs', color: FX_MUTED })); }
  return [fxFlexMsg_(title + '：' + list.length + ' 堂' + (todo ? '，' + todo + ' 堂未叫料' : ''), [fxB_(title, list.length + ' 堂課' + (todo ? '・' + todo + ' 堂未叫料' : '・都叫料了'), rows, nav)])];
}
function lineOrdersMsgs_(from, to, teacher) {
  const orders = ordersBetween_(from, to).filter(function (o) { return o.status !== '已取消' && (!teacher || o.teacher === teacher || o.onBehalfOf === teacher); })
    .sort(function (a, b) { return a.useDate < b.useDate ? -1 : a.useDate > b.useDate ? 1 : SESSIONS.indexOf(a.classTime) - SESSIONS.indexOf(b.classTime); });
  const one = dateStr_(from) === dateStr_(to);
  const title = '📋 ' + (teacher ? teacher + ' 的料單（' + shortDateTw_(from) + '～' + shortDateTw_(to) + '）' : shortDateTw_(from) + ' 料單');
  if (!orders.length) return [{ type: 'text', text: (teacher ? teacher + ' 接下來 7 天沒有料單（名字要跟帳號一樣）。' : shortDateTw_(from) + ' 沒有料單。') }];
  const ded = deductNotes_(orders.slice(0, 12));
  const bubbles = orders.slice(0, 12).map(function (o) {
    const its = o.items.filter(function (i) { return i.status !== '已取消' && i.name; });
    const rows = its.slice(0, 25).map(function (i) {
      const dn = ded[o.orderNo + '|' + i.name];
      return { type: 'box', layout: 'horizontal', contents: [fxT_(i.name + (i.noOrder ? '（有貨）' : dn ? '（' + dn + '）' : ''), { size: 'sm', flex: 6, color: i.noOrder ? FX_MUTED : '#111111' }), fxT_((Math.round(i.qty * 100) / 100) + ' ' + i.unit, { size: 'sm', align: 'end', flex: 3 })] };
    });
    if (its.length > 25) rows.push(fxT_('…還有 ' + (its.length - 25) + ' 項', { size: 'xs', color: FX_MUTED }));
    return fxB_((o.onBehalfOf || o.teacher) + '　' + o.className, (one ? '' : shortDateTw_(parseDate_(o.useDate)) + ' ') + (o.classTime === '晚上' ? '夜間' : o.classTime) + '・' + (o.project || '（沒填實作項目）') + (o.groups ? '・' + o.groups + ' 組' : '') + '・' + o.status,
      [fxT_(its.length + ' 項食材', { size: 'xs', color: FX_MUTED })].concat(rows));
  });
  const ds0 = dateStr_(from);
  const qr = one ? { items: [['◀ 前一天', '料單 ' + dateStr_(addDays_(from, -1))], ['後一天 ▶', '料單 ' + dateStr_(addDays_(from, 1))], ['這天課表', '課表 ' + ds0]].concat(LINE_VIEW_CREW_ ? [] : [['這天叫貨單', '叫貨單 ' + ds0]]).map(function (x) { return { type: 'action', action: { type: 'message', label: x[0], text: x[1] } }; }) } : null;
  const msgs = [{ type: 'text', text: title.replace(/^📋 /, '') + '：共 ' + orders.length + ' 張' + (orders.length > 1 ? '，左右滑動看每一張' : '') + (orders.length > 12 ? '（先列 12 張，其他請到網站「訂單彙整」看）' : '') }, fxFlexMsg_(title, bubbles)];
  if (qr) msgs[msgs.length - 1].quickReply = qr;
  return msgs;
}
/** 叫貨單：那一天送貨的，每家廠商一張卡片（同一樣食材加總） */
function linePurchaseMsgs_(d) {
  const ds = dateStr_(d);
  const qr = { items: [['◀ 前一天', '叫貨單 ' + dateStr_(addDays_(d, -1))], ['後一天 ▶', '叫貨單 ' + dateStr_(addDays_(d, 1))], ['這天料單', '料單 ' + ds], ['這天課表', '課表 ' + ds]].map(function (x) { return { type: 'action', action: { type: 'message', label: x[0], text: x[1] } }; }) };
  let rows = [];
  try { if (ss_().getSheetByName(SHEET.purchase)) rows = readTable_(SHEET.purchase, ['送貨日期', '廠商', '食材名稱']).rows; } catch (e) { rows = []; }
  rows = rows.filter(function (r) { const x = r['送貨日期'] instanceof Date ? r['送貨日期'] : parseDate_(String(r['送貨日期'] || '').replace(/^'/, '')); return x && dateStr_(x) === ds && Number(r['叫貨量']) > 0; });
  if (!rows.length) return [Object.assign({ type: 'text', text: shortDateTw_(d) + ' 沒有要送貨的叫貨單（或還沒產生／還沒確定）。' }, { quickReply: qr })];
  const by = {}, order = [];
  rows.forEach(function (r) {
    const v = String(r['廠商'] || '').trim(), k = String(r['食材名稱'] || '').trim() + '|' + String(r['叫貨單位'] || '').trim();
    const g = by[v] = by[v] || (order.push(v), { items: {}, keys: [], draft: false, groups: {} });
    if (!g.items[k]) { g.items[k] = 0; g.keys.push(k); }
    g.items[k] += Number(r['叫貨量']) || 0;
    if (/草稿/.test(String(r['狀態'] || ''))) g.draft = true;
    if (String(r['送貨組別'] || '').trim()) g.groups[String(r['送貨組別']).trim()] = true;
  });
  order.sort();
  const bubbles = order.map(function (v) {
    const g = by[v];
    const items = g.keys.map(function (k) { const p = k.split('|'); return { type: 'box', layout: 'horizontal', contents: [fxT_(p[0], { size: 'sm', flex: 6 }), fxT_((Math.round(g.items[k] * 100) / 100) + ' ' + p[1], { size: 'sm', align: 'end', flex: 3 })] }; });
    if (items.length > 30) { items.length = 30; items.push(fxT_('…還有 ' + (g.keys.length - 30) + ' 樣，請到網站「廠商訂貨單」看。', { size: 'xs', color: FX_MUTED })); }
    return fxB_('🚚 ' + v, shortDateTw_(d) + ' 送貨・' + g.keys.length + ' 樣' + (Object.keys(g.groups).length ? '・組別 ' + Object.keys(g.groups).join('、') : '') + (g.draft ? '・⚠ 還是草稿' : ''), items);
  });
  return [Object.assign({ type: 'text', text: shortDateTw_(d) + ' 叫貨單：' + order.length + ' 家廠商' + (order.length > 1 ? '，左右滑動看每一家' : '') }, {}), Object.assign(fxFlexMsg_(shortDateTw_(d) + ' 叫貨單', bubbles.slice(0, 12)), { quickReply: qr })];
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

// 1 大T＝15 g、1 小T＝5 g、1 杯＝240 g、1 ml≈1 g（配方單位給廠商時換成重量；老師的料單照原本的單位）（跟網頁的 GRAMS 一樣）
const GRAMS_ = { 'g': 1, '公斤': 1000, '斤': 600, '兩': 37.5, 'T': 15, 't': 5, '杯': 240, 'ml': 1, 'cc': 1 };
// 合計時計數單位同一種的算在一起（個／顆／粒→顆，根／支／條→條）
const COUNT_UNITS_ = { '個': '顆', '粒': '顆', '顆': '顆', '根': '條', '支': '條', '條': '條' };
// 給廠商的合計：菜商、肉商的重量用「兩」加總、顯示成「1 斤 4 兩」；雞蛋用克訂的換成顆（1 顆 50 克）；其他廠商換成公斤
const EGG_GRAMS_ = 50;
function isEgg_(name) { return String(name || '').trim() === '雞蛋'; }
function convertUnit_(qty, unit, vendor, name) {
  unit = vendUnit_(unit);   // 管理者設的單位合併（隻 → 條）
  if (GRAMS_[unit] && isEgg_(name)) return { qty: Math.round(qty * GRAMS_[unit] / EGG_GRAMS_ * 100) / 100, unit: '顆' };
  if (!GRAMS_[unit]) return { qty: qty, unit: COUNT_UNITS_[unit] || unit };
  const target = /^[AB]/.test(vendor || '') ? '兩' : /^[EF]/.test(vendor || '') ? 'g' : '公斤';   // 乾貨、飲調換成克（跟網頁一樣）
  return { qty: Math.round(qty * GRAMS_[unit] / GRAMS_[target] * 100) / 100, unit: target };
}
/** 給廠商的明細（單筆）：菜商、肉商的重量換成斤兩、雞蛋的克換成顆，其他照老師原本的單位 */
function vendorLineQty_(qty, unit, vendor, name) {
  const cv = convertUnit_(Number(qty), unit, vendor, name);
  if (cv.unit === '兩' || (cv.unit === '顆' && isEgg_(name) && GRAMS_[unit])) return totalText_(cv.qty, cv.unit, name);
  return { qty: qty, unit: vendUnit_(unit) };
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
        student: String(r['操作學生'] || '').trim(),
        items: [],
      };
      list.push(map[no]);
    }
    const iid = String(r['食材ID'] || '').trim(), self = !!selfIds[iid];
    map[no].items.push({
      d: map[no].useDate,   // 上課日期：估成本時用「那時候」的單價（雞蛋 10 月 840、11 月 880 要分開看）
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
  const lock = scriptLock_();
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
      if (!/Service|服務|無法正常運作|timed out|逾時|Internal error|內部錯誤|try again|稍後再試|RESOURCE_EXHAUSTED|儲存空間|storage/i.test(m)) throw e;
      if (/RESOURCE_EXHAUSTED|儲存空間|storage/i.test(m) && i >= 1) break;   // Google 忙碌：只再試一次，不要卡很久
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
function memoClear_() { MEMO_ = {}; TABLE_RAW_ = {}; }
// 只讀的查詢（訂單彙整、我的課、總課表、基本資料…）：同一次執行裡同一張表只跟 Google 讀一次（每讀一次整張表要 1～3 秒）。
// 寫入的工具（appendRows_、setCell_、deleteRows_…）都會 memoClear_() 清掉，所以不會讀到舊的
var TABLE_RAW_ = {}, TABLE_MEMO_ON_ = false;
function fastReads_() { TABLE_MEMO_ON_ = true; }

/** 讀整張表，回傳 { sheet, cols: {標題: 欄號(1起)}, rows: [{標題: 值, _row: 列號}] } */
function readTable_(name, required) {
  const sh = sheet_(name);
  // 一次讀完（以前是先問最後一列、最後一欄再讀，要連 Google 三次）
  let values = (TABLE_MEMO_ON_ && TABLE_RAW_[name]) || tblCacheGet_(name);
  if (!values) { values = retry_(function () { return sh.getDataRange().getValues(); }); tblCachePut_(name, values); }
  if (TABLE_MEMO_ON_) TABLE_RAW_[name] = values;
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

/** 管理者：🔎 查單價（成本頁） */
function adminPriceSearch(token, q, vendor) {
  auth_(token, true);
  return priceSearch_(q, vendor);
}
/** 📈 單價漲跌一覽：銷貨單上每一樣（同一家）的全部單價紀錄（每週），畫面照月份排 */
function adminPriceTrend(token) {
  auth_(token, true);
  return priceSearch_('', '', true);
}
/**
 * 🔎 查單價：從「銷貨單明細」找食材（名稱或銷貨單品名有關鍵字，空白隔開＝都要有）。
 * 同一樣、同一家一組：最近一次單價、每週單價（近 26 週）、最高最低。回傳最多 40 組，最近買的排前面。
 * 訂貨系統和驗收系統都用同一份（兩邊程式各一份，改的時候兩邊一起改）。
 */
function priceSearch_(q, vendor, all) {
  const words = String(q || '').toLowerCase().replace(/[（）()]/g, ' ').split(/\s+/).filter(Boolean);
  vendor = String(vendor || '').trim();
  if (!ss_().getSheetByName('銷貨單明細')) return { groups: [], vendors: [] };
  const tz = tz_();
  const toD = function (v) {
    if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    const m = String(v || '').replace(/^'/, '').match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
    return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  };
  const ds = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
  const monOf = function (d) { const x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); x.setDate(x.getDate() - (x.getDay() + 6) % 7); return ds(x); };
  const G = { 'g': 1, '公克': 1, '克': 1, '公斤': 1000, 'kg': 1000, 'KG': 1000, 'Kg': 1000, '斤': 600, '台斤': 600, '兩': 37.5 };
  const groups = {}, vens = {};
  readTable_('銷貨單明細', ['廠商', '食材名稱', '單價']).rows.forEach(function (r) {
    if (String(r['類型'] || '').trim() === '器材') return;
    const vv = String(r['廠商'] || '').trim();
    if (vv && Number(r['單價']) > 0) vens[vv] = (vens[vv] || 0) + 1;
    if (vendor && vv !== vendor) return;
    if (!words.length && !vendor && !all) return;
    const name = String(r['食材名稱'] || '').trim(), sn = String(r['銷貨單品名'] || '').trim();
    const hay = (name + ' ' + sn).toLowerCase();
    if (!name || !words.every(function (w) { return hay.indexOf(w) >= 0; })) return;
    const p = Number(r['單價']);
    if (!(p > 0)) return;
    const d = toD(r['銷貨單日期']) || toD(r['送貨日期（起）']);
    if (!d) return;
    const v = String(r['廠商'] || '').trim(), u = String(r['單位'] || '').trim(), qn = Number(r['數量']) > 0 ? Number(r['數量']) : 0;
    const k = name + '|' + v;
    const g = groups[k] = groups[k] || { name: name, vendor: v, alias: {}, rows: [] };
    if (sn && sn !== name) g.alias[sn] = true;
    g.rows.push({ d: d, u: u, p: p, q: qn, a: Number(r['金額']) > 0 ? Number(r['金額']) : p * qn });
  });
  const r2 = function (x) { return Math.round(x * 100) / 100; };
  const since = new Date(); since.setDate(since.getDate() - 7 * 26);
  const out = Object.keys(groups).map(function (k) {
    const g = groups[k];
    // 主要單位：重量單位算一組（換成最常用的那個），其他照單位；用金額最多的一組
    const byU = {};
    g.rows.forEach(function (x) { const key = G[x.u] ? '#w' : x.u; byU[key] = (byU[key] || 0) + (x.a || x.p) + 0.001; });
    const main = Object.keys(byU).sort(function (a, b) { return byU[b] - byU[a]; })[0];
    let rows = g.rows.filter(function (x) { return (G[x.u] ? '#w' : x.u) === main; });
    let pu = rows[0].u;
    if (main === '#w') { const c = {}; rows.forEach(function (x) { c[x.u] = (c[x.u] || 0) + 1; }); pu = Object.keys(c).sort(function (a, b) { return c[b] - c[a]; })[0]; }
    const unitP = function (x) { return main === '#w' ? x.p * G[pu] / G[x.u] : x.p; };
    const qty = function (x) { return main === '#w' ? x.q * G[x.u] / G[pu] : x.q; };
    rows.sort(function (a, b) { return a.d - b.d; });
    const wk = {};
    rows.forEach(function (x) { const w = monOf(x.d); (wk[w] = wk[w] || []).push(x); });
    const weeks = Object.keys(wk).sort().map(function (w) {
      const xs = wk[w], wq = xs.filter(function (x) { return x.q > 0 && x.a > 0; });
      let p;
      if (wq.length) { let a = 0, q = 0; wq.forEach(function (x) { a += x.a; q += qty(x); }); p = a / q; }
      else p = xs.reduce(function (t, x) { return t + unitP(x); }, 0) / xs.length;
      return { wk: w, p: r2(p), n: xs.length, q: r2(xs.reduce(function (t, x) { return t + qty(x); }, 0)) };
    });
    const last = rows[rows.length - 1], ps = rows.map(unitP);
    const recent = weeks.filter(function (w) { return w.wk >= ds(since); });
    return {
      name: g.name, vendor: g.vendor, alias: Object.keys(g.alias).slice(0, 5), unit: pu,
      last: { p: r2(unitP(last)), date: ds(last.d) },
      lo: r2(Math.min.apply(null, ps)), hi: r2(Math.max.apply(null, ps)), n: rows.length,
      other: g.rows.length - rows.length,   // 用別的單位買的筆數（例如有時候用箱）
      weeks: all ? weeks : (recent.length ? recent : weeks.slice(-6))
    };
  });
  out.sort(function (a, b) { return a.last.date < b.last.date ? 1 : a.last.date > b.last.date ? -1 : a.name.localeCompare(b.name, 'zh-Hant'); });
  const vendors = Object.keys(vens).sort(function (a, b) { return a.localeCompare(b, 'zh-Hant'); });
  if (all) return { groups: out, vendors: vendors };
  if (!words.length && !vendor) return { groups: [], vendors: vendors };
  const lim = vendor ? 400 : 40;   // 選廠商＝看這家全部品項的時價
  return { groups: out.slice(0, lim), more: Math.max(0, out.length - lim), vendors: vendors, vendor: vendor };
}


// ---- 🔄 自動同步：訂單有新增或修改，記一個「資料版本」；別台開著的訂單彙整、審核每 20 秒問一次，變了就自動更新 ----
const DATA_VER_SHEETS_ = ['訂單總表', '待審核食材', '歷史訂單'];
var DATA_BUMPED_ = false;   // 同一次執行只記一次（寫很多格時不要每格都記）
function dataBump_() {
  if (DATA_BUMPED_) return;
  DATA_BUMPED_ = true;
  try { CacheService.getScriptCache().put('ordDataVer', String(Date.now()), 21600); } catch (e) {}
}
function getDataVer(token) {
  return CacheService.getScriptCache().get('ordDataVer') || '';
}

/**
 * 💲 缺單價清單：最近 60 天到之後 45 天有人訂、可是銷貨單和單價表都沒有價錢的食材（同義的叫法、成本對應過的不算缺）。
 * 回傳 [{ name, vendor, units: [[單位, 次數]], n: 幾張料單, next: 最近一次上課日, unit: 食材總表的單位 }]
 */
function adminMissingPrices(token) {
  auth_(token, true);
  const has = {};
  if (ss_().getSheetByName('單價表')) readTable_('單價表', ['食材名稱', '單價']).rows.forEach(function (r) { if (Number(r['單價']) > 0) has[String(r['食材名稱']).trim()] = true; });
  if (ss_().getSheetByName('銷貨單明細')) readTable_('銷貨單明細', ['食材名稱', '單價']).rows.forEach(function (r) {
    if (!(Number(r['單價']) > 0)) return;
    has[String(r['食材名稱']).trim()] = true;
    const sn = String(r['銷貨單品名'] || '').trim(); if (sn) has[sn] = true;
  });
  const cmap = readCostMap_();
  const priced = function (vendor, name) {
    if (has[name]) return true;
    const m = cmap[vendor + '|' + name] || cmap['|' + name];
    if (m && m !== '✗' && has[m]) return true;
    if ((COST_SYN[name] || []).some(function (x) { return has[x]; })) return true;
    // 銷貨單上多了括號、多一兩個字（乾辣椒(進口)、紅蔥頭仁）：成本頁會自動對到，不算缺
    if (name.length < 2) return false;
    return hasKeys().some(function (k) { const base = k.replace(/[（(].*?[)）]/g, '').trim(); return base === name || (k.indexOf(name) === 0 && k.length - name.length <= 2); });
  };
  let keys = null;
  const hasKeys = function () { return keys || (keys = Object.keys(has)); };
  const today = today_(), from = addDays_(today, -60), to = addDays_(today, 45), out = {};
  readTable_(SHEET.orders, ['訂單編號', '使用日期', '食材名稱', '廠商']).rows.forEach(function (r) {
    const d = toDate_(r['使用日期']);
    if (!d || d < from || d > to) return;
    if (/取消|退/.test(String(r['狀態'] || '')) || String(r['不叫貨'] || '').trim()) return;
    const id = String(r['食材ID'] || '').trim(), name = String(r['食材名稱'] || '').trim(), vendor = String(r['廠商'] || '').trim();
    if (!name || id === TOOL_ID || vendor === TOOL_ID || vendor === SELF_VENDOR || !vendor) return;
    if (priced(vendor, name)) return;
    const k = vendor + '|' + name, x = out[k] = out[k] || { name: name, vendor: vendor, id: id, units: {}, orders: {}, next: '' };
    const u = String(r['單位'] || '').trim(); if (u) x.units[u] = (x.units[u] || 0) + 1;
    x.orders[String(r['訂單編號'])] = true;
    const ds = dateStr_(d);
    if (ds >= dateStr_(today) && (!x.next || ds < x.next)) x.next = ds;
  });
  const itemUnit = {};
  readTable_(SHEET.items, ['食材ID', '食材名稱']).rows.forEach(function (r) { itemUnit[String(r['食材名稱']).trim()] = String(r['單位'] || '').trim(); });
  return Object.keys(out).map(function (k) {
    const x = out[k];
    return { name: x.name, vendor: x.vendor, n: Object.keys(x.orders).length, next: x.next, unit: itemUnit[x.name] || '',
      units: Object.keys(x.units).map(function (u) { return [u, x.units[u]]; }).sort(function (a, b) { return b[1] - a[1]; }) };
  }).sort(function (a, b) { return (a.next ? 0 : 1) - (b.next ? 0 : 1) || b.n - a.n || a.vendor.localeCompare(b.vendor, 'zh-Hant'); });
}


// ---- LINE「我的訂單」：最近送出的訂單一張一張卡片，列出全部材料（回覆不算則數；網站送出後馬上查得到） ----
function lineSyncInfo_(account) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('LINE_TOKEN')) return { enabled: false };
  let oa = props.getProperty('LINE_BASIC_ID') || '';
  if (!oa) { try { lineFriendUrl_(); oa = props.getProperty('LINE_BASIC_ID') || ''; } catch (e) {} }
  return { enabled: true, bound: !!lineUserOf_(account), oa: oa };
}
function lineMyOrdersMsgs_(text, userId) {
  const t = String(text || '').trim();
  const todayOnly = /今天|今日/.test(t) && /料單|訂單|材料/.test(t) && t.length <= 12;
  if (!todayOnly && !(/我的訂單|訂單狀態|最近訂單|我的料單/.test(t) && t.length <= 12)) return null;
  const account = lineAccountOf_(userId);
  if (!account) return todayOnly ? [{ type: 'text', text: '還沒綁定帳號，請先登入訂貨網站 →「我的課」→「綁定我的 LINE」。' }] : null;
  const since = addDays_(today_(), -7), td = dateStr_(today_());
  const rows = readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows.filter(function (r) {
    const d = toDate_(r['使用日期']);
    return (String(r['操作老師']).trim() === account || String(r['代訂對象老師'] || '').trim() === account) && d && (todayOnly ? dateStr_(d) === td : d >= since);
  });
  const hidden = hiddenOrders_();   // 老師自己從「我的訂單」刪掉的不列
  const orders = groupOrders_(rows, account).filter(function (o) { return !hidden[o.orderNo] && (!todayOnly || o.status !== '已取消'); })
    .sort(function (a, b) { return String(b.submitted).localeCompare(String(a.submitted)); });   // 最新送出的排第一張
  const url = lineUrl_('mine');
  if (!orders.length) return [{ type: 'text', text: account + ' 老師，' + (todayOnly ? '今天沒有料單。' : '最近（上課日在 7 天前以後）沒有訂單。') + (url ? '\n' + url : '') }];
  if (todayOnly) { const places = placesFor_(orders.map(function (o) { return o.orderNo; })); orders.forEach(function (o) { o.place = places[o.orderNo] || null; }); }
  const fmtQ = function (q) { return String(Math.round(Number(q) * 100) / 100); };
  const bubbles = orders.slice(0, 10).map(function (o) {
    const live = o.items.filter(function (i) { return i.status !== '已取消'; });
    const food = live.filter(function (i) { return i.id !== TOOL_ID; }), tools = live.filter(function (i) { return i.id === TOOL_ID; });
    const line = function (i) {
      return { type: 'box', layout: 'horizontal', contents: [fxT_(i.name + (i.noOrder ? '（有貨）' : ''), { size: 'sm', flex: 6, color: i.noOrder ? FX_MUTED : '#111111' }),
        fxT_(fmtQ(i.qty) + ' ' + i.unit, { size: 'sm', flex: 3, align: 'end', color: '#111111' })] };
    };
    const rowsOut = [];
    food.slice(0, 35).forEach(function (i) { rowsOut.push(line(i)); });
    if (food.length > 35) rowsOut.push(fxT_('…還有 ' + (food.length - 35) + ' 樣，請到網站看。', { size: 'xs', color: FX_MUTED }));
    if (o.place) rowsOut.unshift(fxT_('📍 冷藏食材放在：' + o.place.place, { size: 'sm', color: FX_OK, weight: 'bold' }));
    if (tools.length) {
      rowsOut.push({ type: 'separator', margin: 'md' });
      rowsOut.push(fxT_('🍳 器具', { size: 'xs', color: FX_MUTED, weight: 'bold' }));
      tools.slice(0, 15).forEach(function (i) { rowsOut.push(line(i)); });
    }
    const st = o.status === '已取消' ? '❌ 已取消' : o.status === '待處理' ? '⏳ 待處理' : '✅ ' + o.status;
    const sub = st + '・' + food.length + ' 樣' + (tools.length ? '＋器具 ' + tools.length : '') + '\n' + o.orderNo + '・送出 ' + String(o.submitted).replace(/^\d{4}\//, '').replace(/(\d{1,2}:\d{2}):\d{2}$/, '$1');
    return fxB_(shortDateTw_(parseDate_(o.useDate)) + ' ' + (o.className || '') + '「' + (o.project || '') + '」', sub, rowsOut,
      url ? [{ type: 'button', style: 'link', height: 'sm', action: { type: 'uri', label: '打開我的訂單', uri: url } }] : null);
  });
  return [{ type: 'text', text: todayOnly ? account + ' 老師今天的料單（' + orders.length + ' 堂' + (orders.length > 1 ? '，左右滑動' : '') + '）：' : account + ' 老師最近送出的訂單（最新的在第一張' + (orders.length > 1 ? '，左右滑動看其他張' : '') + '）：' },
    fxFlexMsg_(account + ' 老師的訂單', bubbles)];
}

/**
 * 老師：今天（或指定日期）上課的料單，打開網站時跳出來看。
 * 自己訂的、別人代訂給自己的都算；取消的、老師自己刪掉的不算。放置登記有開放就附上位置。
 */
function getTodayOrders(token, dateStr) {
  const account = auth_(token);
  const day = parseDate_(dateStr) || today_(), ds = dateStr_(day);
  const rows = readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows.filter(function (r) {
    if (String(r['操作老師']).trim() !== account && String(r['代訂對象老師'] || '').trim() !== account) return false;
    const d = toDate_(r['使用日期']);
    return d && dateStr_(d) === ds;
  });
  const hidden = hiddenOrders_();
  const order = { '上午': 1, '下午': 2, '晚上': 3 };
  const list = groupOrders_(rows, account).filter(function (o) { return o.status !== '已取消' && !hidden[o.orderNo]; })
    .sort(function (a, b) { return (order[String(a.classTime).slice(0, 2)] || 9) - (order[String(b.classTime).slice(0, 2)] || 9) || String(a.classTime).localeCompare(String(b.classTime)); });
  const places = placesFor_(list.map(function (o) { return o.orderNo; })), ded = deductNotes_(list);
  list.forEach(function (o) { o.place = places[o.orderNo] || null; o.items.forEach(function (i) { i.ded = ded[o.orderNo + '|' + i.name] || ''; }); });
  return { date: ds, orders: list };
}


// ---- 🔑 忘記密碼：綁定 LINE 的老師在官方帳號傳「忘記密碼」，回一個一次性的連結（10 分鐘、用一次），打開就設定新密碼 ----
// LINE 帳號只有老師本人拿得到，所以用它來確認身分；回覆不算則數。沒綁定的請管理者「重設密碼」改回 0000。
const PW_RESET_MINUTES = 10;
function lineForgotPw_(text, userId) {
  const t = String(text || '').replace(/\s/g, '');
  if (!(/忘記密碼|忘了密碼|密碼忘了|密碼忘記|重設密碼|找回密碼/.test(t) && t.length <= 12)) return null;
  const account = lineAccountOf_(userId);
  if (!account) return [{ type: 'text', text: '這個 LINE 還沒綁定訂貨系統的帳號，沒辦法用 LINE 重設密碼。\n請找管理者在「系統設定 → 👤 老師帳號」按「重設密碼」（改回 0000），登入後再自己改。' }];
  const r = readTable_(SHEET.accounts, ['帳號', '啟用']).rows.find(function (x) { return String(x['帳號']).trim() === account; });
  if (!r || !isTrue_(r['啟用'])) return [{ type: 'text', text: '「' + account + '」這個帳號已經停用，請找管理者。' }];
  const code = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  CacheService.getScriptCache().put('pwreset:' + code, account, PW_RESET_MINUTES * 60);
  const u = lineUrl_('');
  if (!u) return [{ type: 'text', text: '訂貨網站還沒部署，請找管理者。' }];
  try { logLogin_(account, 'LINE 忘記密碼：已傳重設連結'); } catch (e) {}
  return [fxFlexMsg_('重設訂貨系統密碼', [fxB_('🔑 重設密碼', account + '（訂貨系統）', [
    fxT_('按下面的按鈕，會打開訂貨網站，直接請你設定新密碼（不用舊密碼）。', { size: 'sm' }),
    fxT_('⏱ ' + PW_RESET_MINUTES + ' 分鐘內有效、只能用一次。不是你本人要求的話，不用理它，原本的密碼不會變。', { size: 'xs', color: FX_MUTED })
  ], [{ type: 'button', style: 'primary', color: FX_BRAND, action: { type: 'uri', label: '設定新密碼', uri: u + '&rp=' + code } }])])];
}
/** 打開忘記密碼的連結：登入，但只能先設定新密碼 */
function loginByResetLink(code) {
  code = String(code || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 64);
  const cache = CacheService.getScriptCache(), account = code ? cache.get('pwreset:' + code) : null;
  if (!account) throw new Error('這個重設密碼連結已經過期或用過了。請在 LINE 再傳一次「忘記密碼」。');
  cache.remove('pwreset:' + code);
  const r = readTable_(SHEET.accounts, ['帳號', '啟用']).rows.find(function (x) { return String(x['帳號']).trim() === account; });
  if (!r || !isTrue_(r['啟用'])) throw new Error('這個帳號已經停用，請找管理者。');
  const token = Utilities.getUuid();
  cache.put('session:' + token, account, SESSION_HOURS * 3600);
  cache.put('mustchange:' + token, '1', SESSION_HOURS * 3600);   // 設定好新密碼之前，什麼都不能做
  cache.put('pwresettok:' + token, '1', 30 * 60);
  try { logLogin_(account, '成功（LINE 忘記密碼連結）'); } catch (e) {}
  return { token: token, account: account, mustChange: true, reset: true };
}


// ---- LINE「今天的課」：老師今天每一堂課（課表＋自訂＋國技）一張卡片，附上那堂的料單材料、放哪裡；沒訂、不用訂也寫出來（回覆不算則數） ----
function lineTodayLessonsMsgs_(text, userId) {
  const t = String(text || '').replace(/\s/g, '');
  if (!(/今天的?課|今日課|今天上什麼|今天有什麼課/.test(t) && t.length <= 10)) return null;
  const account = lineAccountOf_(userId);
  if (!account) return [{ type: 'text', text: '還沒綁定帳號，請先登入訂貨網站 →「我的課」→「綁定我的 LINE」，再把綁定碼傳給我。' }];
  const day = today_(), ds = dateStr_(day), url = lineUrl_('lessons');
  const wd = day.getDay();
  const lessons = (wd === 0 || wd === 6) ? [] : lessonsWithStatus_(day, day, account);
  // 今天的訂單（含別人代訂給我的）
  const rows = readTable_(SHEET.orders, ['訂單編號', '操作老師']).rows.filter(function (r) {
    if (String(r['操作老師']).trim() !== account && String(r['代訂對象老師'] || '').trim() !== account) return false;
    const d = toDate_(r['使用日期']); return d && dateStr_(d) === ds;
  });
  const hidden = hiddenOrders_();
  const orders = groupOrders_(rows, account).filter(function (o) { return o.status !== '已取消' && !hidden[o.orderNo]; });
  const places = placesFor_(orders.map(function (o) { return o.orderNo; }));
  const byNo = {}; orders.forEach(function (o) { o.place = places[o.orderNo] || null; byNo[o.orderNo] = o; });
  const used = {};
  const fmtQ = function (q) { return String(Math.round(Number(q) * 100) / 100); };
  const itemRows = function (o) {
    const live = o.items.filter(function (i) { return i.status !== '已取消'; });
    const food = live.filter(function (i) { return i.id !== TOOL_ID; }), tools = live.filter(function (i) { return i.id === TOOL_ID; });
    const dd = deductNotes_([o]);
    const line = function (i) { const dn = dd[o.orderNo + '|' + i.name]; return { type: 'box', layout: 'horizontal', contents: [fxT_(i.name + (i.noOrder ? '（有貨）' : dn ? '（' + dn + '）' : ''), { size: 'sm', flex: 6, color: i.noOrder ? FX_MUTED : '#111111' }), fxT_(fmtQ(i.qty) + ' ' + i.unit, { size: 'sm', flex: 3, align: 'end', color: '#111111' })] }; };
    const out = [fxT_('📄 「' + (o.project || '') + '」' + o.status, { size: 'sm', weight: 'bold', color: FX_BRAND })];
    if (o.place) out.push(fxT_('📍 冷藏食材放在：' + o.place.place, { size: 'sm', color: FX_OK, weight: 'bold' }));
    food.slice(0, 30).forEach(function (i) { out.push(line(i)); });
    if (food.length > 30) out.push(fxT_('…還有 ' + (food.length - 30) + ' 樣，請到網站看。', { size: 'xs', color: FX_MUTED }));
    if (tools.length) { out.push(fxT_('🍳 器具', { size: 'xs', color: FX_MUTED, weight: 'bold' })); tools.slice(0, 12).forEach(function (i) { out.push(line(i)); }); }
    return out;
  };
  const SO = { '上午': 1, '下午': 2, '晚上': 3 };
  const bubbles = lessons.slice().sort(function (a, b) { return (SO[a.session] || 9) - (SO[b.session] || 9); }).map(function (x) {
    let body = [];
    if (x.status === 'holiday') body.push(fxT_('📅 放假' + (x.holiday ? '：' + x.holiday : ''), { color: FX_MUTED }));
    else if (x.status === 'skipped') body.push(fxT_('✔ 不用訂貨' + (x.skipNote ? '（' + x.skipNote + '）' : ''), { color: FX_MUTED }));
    else if (x.status === 'todo') body.push(fxT_('⚠ 這堂課沒有訂貨', { color: FX_BAD, weight: 'bold' }));
    (x.orders || []).forEach(function (oo) { const o = byNo[oo.orderNo]; if (!o) return; used[o.orderNo] = true; if (body.length) body.push({ type: 'separator', margin: 'md' }); body = body.concat(itemRows(o)); });
    if (x.status === 'ordered' && !(x.orders || []).length) body.push(fxT_('✔ 已標記叫料', { color: FX_MUTED }));
    return fxB_(x.session + '　' + (x.course || ''), [x.className, x.room].filter(Boolean).join('・'), body);
  });
  // 今天的訂單對不到課表的（代課、補課）也列出來
  orders.filter(function (o) { return !used[o.orderNo]; }).forEach(function (o) {
    bubbles.push(fxB_((o.classTime || '') + '　' + (o.className || ''), '另外訂的料單', itemRows(o)));
  });
  if (!bubbles.length) return [{ type: 'text', text: account + ' 老師，今天（' + shortDateTw_(day) + '）沒有課。' }];
  const foot = url ? [{ type: 'button', style: 'link', height: 'sm', action: { type: 'uri', label: '打開我的課', uri: url } }] : null;
  if (foot) bubbles.forEach(function (b) { b.footer = { type: 'box', layout: 'horizontal', paddingAll: '8px', contents: foot }; });
  return [{ type: 'text', text: account + ' 老師，今天 ' + shortDateTw_(day) + ' 有 ' + bubbles.length + ' 堂' + (bubbles.length > 1 ? '（左右滑動）' : '') + '：' }, fxFlexMsg_(account + ' 老師今天的課', bubbles)];
}


/**
 * 扣貨扣掉的品項：照「叫貨紀錄」（② 訂貨單填的實際叫貨），每一班每一樣比老師訂的少就標出來。
 * 回傳 { '訂單編號|食材名稱': '扣貨已扣掉' | '扣貨扣掉 2 斤，只叫 1 斤' }
 */
function deductNotes_(orders) {
  const out = {}, want = {};
  (orders || []).forEach(function (o) { want[o.orderNo] = true; });
  const sh = ss_().getSheetByName(SHEET.purchase);
  if (!sh || sh.getLastRow() < 2 || !Object.keys(want).length) return out;
  const sum = {};
  readTable_(SHEET.purchase, ['訂單編號', '食材名稱', '叫貨量']).rows.forEach(function (r) {
    const no = String(r['訂單編號'] || '').trim();
    if (!want[no] || String(r['類型'] || '訂單').trim() !== '訂單') return;
    const k = no + '|' + String(r['食材名稱'] || '').trim();   // 單位不比（乾貨的需求會換成克）
    const x = sum[k] = sum[k] || { need: 0, qty: 0, unit: String(r['叫貨單位'] || '').trim(), nu: String(r['需求單位'] || '').trim(), mixed: false };
    x.need += Number(r['需求量']) || 0; x.qty += Number(r['叫貨量']) || 0;
    if (String(r['叫貨單位'] || '').trim() !== x.unit || String(r['需求單位'] || '').trim() !== x.nu) x.mixed = true;
  });
  const f = function (n) { return String(Math.round(n * 100) / 100); };
  Object.keys(sum).forEach(function (k) {
    const x = sum[k];
    if (!(x.need > 0)) return;
    if (x.qty <= 0) out[k] = '扣貨已扣掉';
    else if (!x.mixed && x.unit === x.nu && x.qty < x.need) out[k] = '扣貨扣掉 ' + f(x.need - x.qty) + ' ' + x.nu + '，只叫 ' + f(x.qty) + ' ' + x.unit;
  });
  return out;
}


// ---- ⚡ 少讀試算表：很少改的分頁（課表、單位換算…）讀過就放在 Google 暫存 5 分鐘，下一次直接拿（快 0.2～1 秒／分頁） ----
// 系統自己寫這些分頁時馬上丟掉暫存（bootDirty_）；這次執行有拿鎖（要存東西）就一律直接讀試算表，免得照舊的列號寫錯列。
// 直接在試算表手動改的，最慢 5 分鐘後生效；急的話系統設定按「🔄 重新讀取試算表」。
const TBL_CACHE_SHEETS_ = ['課表', '老師課表', '國技排班', '放假日', '不用訂貨', '單位換算', '成本對應', '科別替換', '班級人數', '廠商設定', '廠商聯絡', '單位', '範本', '修改紀錄', '老師刪除', '放置位置'];
const TBL_CACHE_SEC_ = 300, TBL_CHUNK_ = 90000;
var TBL_NO_CACHE_ = false;
function scriptLock_() { TBL_NO_CACHE_ = true; return LockService.getScriptLock(); }
function tblCacheOk_(name) { return !TBL_NO_CACHE_ && TBL_CACHE_SHEETS_.indexOf(String(name)) >= 0; }
function tblCacheGet_(name) {
  if (!tblCacheOk_(name)) return null;
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get('tbl:' + name + ':n'));
    if (!(n > 0)) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push('tbl:' + name + ':' + i);
    const got = c.getAll(keys); let s = '';
    for (let i = 0; i < n; i++) { if (got[keys[i]] == null) return null; s += got[keys[i]]; }
    return JSON.parse(s, function (k, v) { return v && typeof v === 'object' && v.__d !== undefined && Object.keys(v).length === 1 ? new Date(v.__d) : v; });
  } catch (e) { return null; }
}
function tblCachePut_(name, values) {
  if (!tblCacheOk_(name)) return;
  try {
    const s = JSON.stringify(values, function (k, v) { const raw = this[k]; return raw instanceof Date ? { __d: raw.getTime() } : v; });
    if (s.length > TBL_CHUNK_ * 40) return;   // 太大就不放
    const parts = {}; let n = 0;
    for (let i = 0; i < s.length; i += TBL_CHUNK_) parts['tbl:' + name + ':' + (n++)] = s.slice(i, i + TBL_CHUNK_);
    const c = CacheService.getScriptCache();
    c.putAll(parts, TBL_CACHE_SEC_);
    c.put('tbl:' + name + ':n', String(n), TBL_CACHE_SEC_);
  } catch (e) {}
}
function tblCacheDrop_(name) {
  if (TBL_CACHE_SHEETS_.indexOf(String(name)) < 0) return;
  try { CacheService.getScriptCache().remove('tbl:' + name + ':n'); } catch (e) {}
}
/** 管理者：直接在試算表改了課表、單位換算…，按這個讓網站馬上讀新的 */
function adminClearSheetCache(token) {
  auth_(token, true);
  const c = CacheService.getScriptCache();
  // 連食材總表、自備食材、管理者名單、教學影片清單一起清（登入狀態不動，不會把大家登出）
  c.removeAll(TBL_CACHE_SHEETS_.map(function (n) { return 'tbl:' + n + ':n'; }).concat([BOOT_CACHE_KEY, 'items:n', 'selfIds', 'extraAdmins', 'helpVideos']));
  return TBL_CACHE_SHEETS_.length;
}


// ---- 管理者 LINE：❗ 還沒叫料的老師、📊 下週訂貨進度（「下週」＝老師現在正在訂的那一週；回覆不算則數） ----
function adminOrderWeek_(later) {
  const first = earliestClassDate_(new Date());
  const mon = addDays_(first, -((first.getDay() + 6) % 7) + (later ? 7 : 0));
  return { mon: mon, fri: addDays_(mon, 4), from: later ? addDays_(first, -((first.getDay() + 6) % 7) + 7) : first };
}
function lineTodoTeachersMsgs_(later) {
  const w = adminOrderWeek_(later);
  const ls = lessonsWithStatus_(w.from, w.fri, null).filter(function (x) { return x.status === 'todo'; });
  const title = (later ? '下下週' : '下週') + ' ' + shortDateTw_(w.from) + '～' + shortDateTw_(w.fri);
  const dl = currentDeadline_(new Date()), dlText = '截止：' + shortDateTw_(dl > new Date() ? dl : addDays_(dl, 7)) + ' ' + lockTime_();
  const qr = { items: [['下下週', '還沒叫料 下下週'], ['下週', '還沒叫料'], ['📊 訂貨進度', '訂貨進度']].map(function (a) { return { type: 'action', action: { type: 'message', label: a[0], text: a[1] } }; }) };
  if (!ls.length) return [Object.assign({ type: 'text', text: '👍 ' + title + ' 的課都叫料了（或標了不用訂貨）。\n' + dlText }, { quickReply: qr })];
  const by = {};
  ls.forEach(function (x) { (by[x.teacher] = by[x.teacher] || []).push(x); });
  const SO = { '上午': 1, '下午': 2, '晚上': 3 };
  const names = Object.keys(by).sort(function (a, b) { return by[b].length - by[a].length || a.localeCompare(b, 'zh-Hant'); });
  const rows = [];
  names.forEach(function (n, i) {
    if (i) rows.push({ type: 'separator' });
    rows.push(fxT_(n + '（' + by[n].length + ' 堂）', { size: 'sm', weight: 'bold', color: FX_BAD }));
    by[n].sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (SO[a.session] || 9) - (SO[b.session] || 9); }).slice(0, 8).forEach(function (x) {
      rows.push(fxT_('・' + shortDateTw_(parseDate_(x.date)) + x.session + ' ' + (x.course || '') + (x.className ? ' ' + x.className : ''), { size: 'xs', color: '#333333' }));
    });
    if (by[n].length > 8) rows.push(fxT_('…還有 ' + (by[n].length - 8) + ' 堂', { size: 'xs', color: FX_MUTED }));
  });
  if (rows.length > 70) { rows.length = 70; rows.push(fxT_('…太多了，其他請到網站「總課表」看。', { size: 'xs', color: FX_MUTED })); }
  return [Object.assign(fxFlexMsg_('還沒叫料：' + names.length + ' 位老師、' + ls.length + ' 堂', [fxB_('❗ 還沒叫料的老師', title + '・' + names.length + ' 位、' + ls.length + ' 堂\n' + dlText, rows)]), { quickReply: qr })];
}
function lineWeekProgressMsgs_(later) {
  const w = adminOrderWeek_(later);
  const ls = lessonsWithStatus_(w.from, w.fri, null);
  const cnt = { ordered: 0, skipped: 0, todo: 0, holiday: 0 };
  ls.forEach(function (x) { cnt[x.status] = (cnt[x.status] || 0) + 1; });
  const need = ls.length - cnt.holiday, done = cnt.ordered + cnt.skipped, pct = need ? Math.round(done / need * 100) : 100;
  // 那週的訂單：待處理（還沒審核）幾張
  const orders = ordersBetween_(w.mon, w.fri).filter(function (o) { return o.status !== '已取消'; });
  const waiting = orders.filter(function (o) { return o.status === '待處理'; }).length;
  let pend = 0;
  try { readTable_(SHEET.pending, ['食材名稱']).rows.forEach(function (r) { const st = String(r['審核狀態'] || '').trim(); if (!st || st === '待審核') pend++; }); } catch (e) {}
  // 叫貨：哪幾家確定了、哪幾家還是草稿
  const vs = {};
  try {
    const sh = ss_().getSheetByName(SHEET.purchase);
    if (sh && sh.getLastRow() > 1) readTable_(SHEET.purchase, ['送貨日期', '廠商']).rows.forEach(function (r) {
      const d = toDate_(r['送貨日期']); if (!d || d < addDays_(w.mon, -4) || d > w.fri) return;
      const v = String(r['廠商'] || '').trim(); if (!v) return;
      if (String(r['狀態'] || '').trim() === '草稿') vs[v] = vs[v] || '草稿'; else vs[v] = '確定';
    });
  } catch (e) {}
  const conf = Object.keys(vs).filter(function (v) { return vs[v] === '確定'; }), draft = Object.keys(vs).filter(function (v) { return vs[v] === '草稿'; });
  const dl = currentDeadline_(new Date()), dlDate = dl > new Date() ? dl : addDays_(dl, 7);
  const bar = function (p) { const n = Math.round(p / 10); return '▓'.repeat(n) + '░'.repeat(10 - n); };
  const line = function (a, b, color) { return { type: 'box', layout: 'horizontal', contents: [fxT_(a, { size: 'sm', flex: 6 }), fxT_(b, { size: 'sm', flex: 4, align: 'end', weight: 'bold', color: color || '#111111' })] }; };
  const rows = [
    fxT_(bar(pct) + '  ' + pct + '%', { size: 'md', weight: 'bold', color: pct >= 100 ? FX_OK : FX_BRAND }),
    line('要上課', need + ' 堂'),
    line('✅ 已叫料', cnt.ordered + ' 堂', FX_OK),
    line('➖ 不用訂貨', cnt.skipped + ' 堂', FX_MUTED),
    line('❗ 還沒叫料', cnt.todo + ' 堂', cnt.todo ? FX_BAD : FX_OK),
  ].concat(cnt.holiday ? [line('📅 放假', cnt.holiday + ' 堂', FX_MUTED)] : []).concat([
    { type: 'separator' },
    line('⏳ 待處理訂單（還沒審核）', waiting + ' 張', waiting ? '#b54708' : FX_OK),
    line('🆕 待審核食材', pend + ' 樣', pend ? '#b54708' : FX_OK),
    { type: 'separator' },
    fxT_('🚚 叫貨：' + (conf.length ? '已確定 ' + conf.join('、') : '還沒確定任何一家') + (draft.length ? '\n📝 草稿：' + draft.join('、') : ''), { size: 'xs', color: '#333333' }),
    fxT_('截止：' + shortDateTw_(dlDate) + ' ' + lockTime_(), { size: 'xs', color: FX_MUTED }),
  ]);
  const qr = { items: [['❗ 還沒叫料', later ? '還沒叫料 下下週' : '還沒叫料'], [later ? '下週' : '下下週', later ? '訂貨進度' : '訂貨進度 下下週']].map(function (a) { return { type: 'action', action: { type: 'message', label: a[0], text: a[1] } }; }) };
  return [Object.assign(fxFlexMsg_('訂貨進度 ' + pct + '%', [fxB_('📊 ' + (later ? '下下週' : '下週') + '訂貨進度', shortDateTw_(w.from) + '～' + shortDateTw_(w.fri), rows)]), { quickReply: qr })];
}


/** Apps Script 每個專案最多 200 個版本：滿了就告訴管理者怎麼刪舊版本（Google 沒有開放用程式刪） */
function versionLimitError_(e, which) {
  const m = String(e && e.message || e);
  if (!/200 versions|limit of \d+ versions|maximum number of versions|版本/i.test(m)) return e;
  return new Error('「' + which + '」的 Apps Script 版本已經滿 200 個（Google 的上限），這次沒有更新，網站維持原本的版本。\n' +
    '請用電腦刪掉舊版本：打開 script.google.com →「' + which + '」專案 → 左邊時鐘圖示「專案記錄」→ 右上「大量刪除版本」→ 勾選舊的版本（留最新的 10 個就好；正在使用的版本不能刪）→ 刪除，再回來按一次更新。');
}
