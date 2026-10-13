/**
 * 古着物販アシスタント Ver.1 — 共通の設定と道具
 *
 * ★ スクールのスクリプトとは別ファイル。
 *   スクールの関数・定数と名前がぶつからないよう、名前はすべて ASST_ / asst で始める。
 *   商品管理シートの列は「見出しの名前」で探す（列の位置に依存しない）。
 */

// ===== シート名 =====
const ASST_ITEMS_SHEET = '商品管理';   // スクールのシート（読むだけ。段階1では書き込まない）
const ASST_SHEETS = {
  SETTINGS: '⚙️アシスタント設定',
  MANAGE: 'アシスタント管理',
  LOG: '定点観測ログ',
  PRICE_LOG: '値下げ履歴',
};

// ===== 商品のステータス =====
const ASST_STATUS = {
  WAIT: '7日間様子見',
  PENDING: '観測待ち',       // 7日を過ぎたが、まだ対応が決まっていない
  MARKDOWN: '定期値下げ',
  IMPROVE: '改善後観察',
  STOPPED: '値下げ停止',
  SOLD: '売却済み',
  OUT: '対象外',             // 再出品で置き換わった行・無料交換に出した行
};

// ===== 商品管理シートで使う列（見出しの名前）=====
const ASST_ITEM_COLS = {
  MGT: '管理番号',
  NAME: '商品名',
  LIST_PRICE: '出品価格',
  LIST_DATE: '出品日',
  SOLD_DATE: '売却日',
  VIEWS_7D: '7日後閲覧数',
  SEARCH_7D: '7日後検索数',
  LIKES: 'いいね数',
  DROPS: '値下げ回数',
  SALE_DONE: '売却完了',
  MEMO: 'メモ',
  SUPERSEDED: '再出品済み',
  EXCHANGED: '無料交換に出した',
};

// ===== 新しいシートの見出し =====
const ASST_MANAGE_HEADERS = [
  '管理番号', '商品名', '出品日', 'ステータス', '現在価格',
  '次回定点観測日', '最終観測日', '観測回数',
  '定期値下げ', '最終値下げ日', '次回値下げ予定日',
  '最終改善日', '改善内容', '確認メモ', '更新日時',
];
const ASST_LOG_HEADERS = [
  '計測日', '管理番号', '商品名', '経過日数',
  '閲覧数（合計）', '検索数（合計）', 'いいね数（合計）',
  '閲覧数（週の増加）', '検索数（週の増加）', 'いいね数（週の増加）',
  '閲覧率（%）', 'その時の価格',
  '前回から行った変更', 'システムのおすすめ', '実際に行った対応', 'メモ', '記録元',
];
const ASST_PRICE_HEADERS = [
  '日付', '管理番号', '商品名', '操作', '旧価格', '新価格', '値下げ額',
  '値下げ回数（累計）', '出品からの日数', '記録元',
];

// ===== 設定シートの項目（A列の項目名で値を探すので、項目名は変えない）=====
const ASST_SETTING = {
  OBS_DAYS: '観測の間隔（日）',
  DROP_DAYS: '値下げの間隔（日）',
  SEARCH_LOW: '検索数が少ない基準',
  VIEW_RATE_LOW: '閲覧率が低い基準（%）',
  LIKE_VIEWS_MIN: 'いいね判定に使う閲覧数',
  LIKES_LOW: 'いいねが少ない基準',
  REFRESH_HOUR: '毎朝の更新時刻（時）',
};
const ASST_SETTINGS_DEFAULTS = [
  [ASST_SETTING.OBS_DAYS, 7, '出品日から何日ごとに定点観測するか'],
  [ASST_SETTING.DROP_DAYS, 4, '定期値下げの通知を何日ごとに出すか'],
  [ASST_SETTING.SEARCH_LOW, 50, 'その週に増えた検索数がこの数未満なら「検索数が少ない」'],
  [ASST_SETTING.VIEW_RATE_LOW, 10, 'その週の閲覧率がこの%未満なら「閲覧率が低い」'],
  [ASST_SETTING.LIKE_VIEWS_MIN, 20, 'その週に増えた閲覧数がこの数以上で、'],
  [ASST_SETTING.LIKES_LOW, 0, '増えたいいねがこの数以下なら「いいねが少ない」'],
  [ASST_SETTING.REFRESH_HOUR, 6, '「今日の物販タスク」を作り直す時刻（0〜23）'],
];

// ===== 日付の道具（時刻を切り捨てた「日」で扱う）=====
function asstDay_(v) {
  if (!(v instanceof Date) || isNaN(v.getTime())) return null;
  return new Date(v.getFullYear(), v.getMonth(), v.getDate());
}

function asstToday_() {
  return asstDay_(new Date());
}

function asstAddDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

/** a から b まで何日か（b − a） */
function asstDaysBetween_(a, b) {
  return Math.round((asstDay_(b).getTime() - asstDay_(a).getTime()) / 86400000);
}

// ===== シートの道具 =====
function asstSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/** 1行目の見出し → 列の位置（0始まり）の対応表 */
function asstHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const map = {};
  if (lastCol < 1) return map;
  sheet.getRange(1, 1, 1, lastCol).getValues()[0].forEach(function(h, i) {
    const key = String(h || '').trim();
    if (key && !(key in map)) map[key] = i;
  });
  return map;
}

/** 見出しが全部そろっているか確認。足りなければ、その名前の一覧を返す */
function asstMissingHeaders_(map, names) {
  return names.filter(function(n) { return !(n in map); });
}

/**
 * 商品管理シートを読み、管理番号のある行を商品オブジェクトの配列で返す。
 * 読むだけで、商品管理シートには何も書き込まない。
 */
function asstReadItems_() {
  const sheet = asstSheet_(ASST_ITEMS_SHEET);
  if (!sheet) throw new Error('「' + ASST_ITEMS_SHEET + '」シートが見つかりません。');
  const map = asstHeaderMap_(sheet);
  const required = Object.keys(ASST_ITEM_COLS).map(function(k) { return ASST_ITEM_COLS[k]; });
  const missing = asstMissingHeaders_(map, required);
  if (missing.length) {
    throw new Error('商品管理シートに次の見出しが見つかりません：' + missing.join('、') +
      '\nスクールのシートの見出しが変わった可能性があります。');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const C = ASST_ITEM_COLS;
  const items = [];
  values.forEach(function(r) {
    const mgt = String(r[map[C.MGT]] || '').trim();
    if (!mgt) return;
    items.push({
      mgt: mgt,
      name: String(r[map[C.NAME]] || ''),
      listPrice: r[map[C.LIST_PRICE]],
      listDate: asstDay_(r[map[C.LIST_DATE]]),
      soldDate: r[map[C.SOLD_DATE]],
      views7d: r[map[C.VIEWS_7D]],
      search7d: r[map[C.SEARCH_7D]],
      likes: r[map[C.LIKES]],
      drops: Number(r[map[C.DROPS]]) || 0,
      saleDone: r[map[C.SALE_DONE]] === true,
      memo: String(r[map[C.MEMO]] || ''),
      superseded: r[map[C.SUPERSEDED]] === true,
      exchanged: r[map[C.EXCHANGED]] === true,
    });
  });
  return items;
}

/** 売れた商品か */
function asstIsSold_(item) {
  return item.saleDone || (item.soldDate !== '' && item.soldDate !== null && item.soldDate !== undefined);
}

/** 再出品で置き換わった行・無料交換に出した行か */
function asstIsOut_(item) {
  return item.superseded || item.exchanged;
}

/** 設定シートの値を読む。読めない項目は初期値を使う */
function asstGetSettings_() {
  const s = {};
  ASST_SETTINGS_DEFAULTS.forEach(function(row) { s[row[0]] = row[1]; });
  const sheet = asstSheet_(ASST_SHEETS.SETTINGS);
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
      const key = String(r[0] || '').trim();
      const val = Number(r[1]);
      if (key in s && r[1] !== '' && !isNaN(val)) s[key] = val;
    });
  }
  return s;
}

function asstIsNumber_(v) {
  return v !== '' && v !== null && v !== undefined && !isNaN(Number(v));
}
