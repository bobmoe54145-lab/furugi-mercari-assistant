/**
 * 古着物販アシスタント Ver.1 — 段階4：今日の物販タスク・毎朝の自動更新・A列の色分け
 *
 * - 「今日の物販タスク」シート：今日（と期限切れ）の定点観測・値下げを一覧にする。毎朝と、記録のたびに作り直す。
 * - 毎朝の自動更新：時間指定のトリガーで asstDailyRefresh を動かす（設定シートの時刻）。
 * - 色分け：商品管理シートの A列（管理番号）だけに条件付き書式を足す。
 *     濃い黄色 = 今日、定点観測が必要 / 濃い赤（白文字）= 今日、値下げが必要
 *   スクールの色ルール9個は消さず、中身も変えない（アシスタントのルールを先頭に置くだけ）。
 */

const ASST_TASKS_SHEET = '今日の物販タスク';
const ASST_FLAGS_SHEET = 'アシスタント色分け用';   // 色分けのための非表示シート
const ASST_TASK_WIDTH = 9;

// ===== メニュー =====

function asstRefreshTasks() {
  const ui = SpreadsheetApp.getUi();
  if (!asstSheet_(ASST_SHEETS.MANAGE)) {
    ui.alert('先にメニュー「初期設定」を実行してください。');
    return;
  }
  try {
    asstSyncItems_(asstReadItems_());
    const counts = asstRebuildTasks_();
    asstApplyHighlight_();
    asstSheet_(ASST_TASKS_SHEET).activate();
    SpreadsheetApp.getActive().toast(
      '📊 定点観測 ' + counts.obs + '件 ・ 💰 値下げ ' + counts.drop + '件', '今日の物販タスクを更新しました', 5);
  } catch (err) {
    ui.alert('❌ 更新できませんでした', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** メニュー「毎朝の自動更新を設定」。設定シートの時刻で毎日動くようにする（時刻を変えたら、もう一度押す） */
function asstInstallDaily() {
  const ui = SpreadsheetApp.getUi();
  if (!asstSheet_(ASST_SHEETS.MANAGE)) {
    ui.alert('先にメニュー「初期設定」を実行してください。');
    return;
  }
  let hour = Number(asstGetSettings_()[ASST_SETTING.REFRESH_HOUR]);
  if (!Number.isInteger(hour) || hour < 0 || hour > 23) hour = 6;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'asstDailyRefresh') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('asstDailyRefresh').timeBased().everyDays(1).atHour(hour).create();
  asstDailyRefresh();
  ui.alert('⏰ 毎朝の自動更新を設定しました',
    '毎朝 ' + hour + '時台に「今日の物販タスク」と色分けを自動で更新します。\n' +
    '（Google の仕組みで、' + hour + '時〜' + hour + '時59分のどこかで動きます）\n\n' +
    '時刻を変えたいときは、設定シートの「毎朝の更新時刻（時）」を書き換えてから、もう一度このメニューを押してください。',
    ui.ButtonSet.OK);
}

/** 毎朝トリガーから呼ばれる（画面は出さない） */
function asstDailyRefresh() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    asstSyncItems_(asstReadItems_());
  } finally {
    lock.releaseLock();
  }
  asstRebuildTasks_();
  asstApplyHighlight_();
}

// ===== 今日の物販タスク =====

/**
 * 「今日の物販タスク」シートと、色分け用の一覧を作り直す。
 * @return {{obs:number, drop:number}}
 */
function asstRebuildTasks_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = asstToday_();
  const manage = asstReadManage_();
  const H = manage.H;
  const logs = asstReadLogsByMgt_();
  const urls = asstReadItemUrls_();
  const dropsByMgt = {};
  asstReadItems_().forEach(function(it) { if (!(it.mgt in dropsByMgt)) dropsByMgt[it.mgt] = it.drops; });

  const obs = [], drop = [];
  manage.rows.forEach(function(r) {
    const status = r[H['ステータス']];
    if (status === ASST_STATUS.SOLD || status === ASST_STATUS.OUT) return;
    const mgt = String(r[H['管理番号']] || '').trim();
    const listDate = asstDay_(r[H['出品日']]);
    if (!mgt || !listDate) return;
    const nextObs = asstDay_(r[H['次回定点観測日']]);
    if (nextObs && nextObs <= today) obs.push({ r: r, mgt: mgt, date: nextObs, listDate: listDate });
    const nextDrop = asstDay_(r[H['次回値下げ予定日']]);
    if (status === ASST_STATUS.MARKDOWN && nextDrop && nextDrop <= today) {
      drop.push({ r: r, mgt: mgt, date: nextDrop, listDate: listDate });
    }
  });
  const byDate = function(a, b) { return a.date - b.date || a.mgt.localeCompare(b.mgt); };
  obs.sort(byDate);
  drop.sort(byDate);

  const late = function(d) {
    const n = asstDaysBetween_(d, today);
    return n > 0 ? '期限切れ（' + n + '日）' : '今日';
  };
  const link = function(mgt) {
    const url = urls[mgt];
    return url ? '=HYPERLINK("' + url.replace(/"/g, '""') + '","開く")' : '';
  };
  const yen = function(v) { return asstIsNumber_(v) ? Number(v) : ''; };

  const rows = [], styles = [];  // styles: 行ごとの見た目の種類
  const push = function(row, style) {
    while (row.length < ASST_TASK_WIDTH) row.push('');
    rows.push(row);
    styles.push(style || '');
  };

  push(['今日の物販タスク（' + asstFmt_(today) + '）', '', '', '', '', '', '', '', '更新：' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d HH:mm')], 'title');
  push(['📊 定点観測：' + obs.length + '件　💰 値下げ：' + drop.length + '件　合計：' + (obs.length + drop.length) + '件'], 'count');
  push(['行を選んでからメニュー「物販アシスタント」→「📊 定点観測を記録」「💰 値下げ」を押すと、その商品が選ばれた状態で開きます。'], 'hint');
  push([], '');

  push(['【本日の定点観測】' + obs.length + '件'], 'section');
  push(['管理番号', '商品名', '予定日', '観測', '状況', 'ステータス', '前回の記録', '現在価格', 'メルカリ'], 'header');
  if (!obs.length) push(['今日の定点観測はありません。'], 'none');
  obs.forEach(function(o) {
    const history = logs[o.mgt] || [];
    const last = history.length ? history[history.length - 1] : null;
    push([o.mgt, o.r[H['商品名']], o.date, asstDaysBetween_(o.listDate, o.date) + '日後', late(o.date),
      o.r[H['ステータス']],
      last ? asstFmt_(last.date).slice(5) + '（' + last.elapsed + '日後）閲覧' + last.views + '・検索' + last.search +
        '・いいね' + (last.likes === '' ? '-' : last.likes) : 'まだなし',
      yen(o.r[H['現在価格']]), link(o.mgt)], late(o.date) === '今日' ? 'item' : 'late');
  });
  push([], '');

  push(['【本日の値下げ】' + drop.length + '件'], 'section');
  push(['管理番号', '商品名', '予定日', '値下げ', '状況', '現在価格', '最終値下げ日', '出品から', 'メルカリ'], 'header');
  if (!drop.length) push(['今日の値下げはありません。'], 'none');
  drop.forEach(function(d) {
    const lastDrop = asstDay_(d.r[H['最終値下げ日']]);
    const count = (dropsByMgt[d.mgt] || 0) + 1;
    push([d.mgt, d.r[H['商品名']], d.date, count + '回目', late(d.date), yen(d.r[H['現在価格']]),
      lastDrop || '', asstDaysBetween_(d.listDate, today) + '日', link(d.mgt)], late(d.date) === '今日' ? 'item' : 'late');
  });

  let sheet = ss.getSheetByName(ASST_TASKS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ASST_TASKS_SHEET, 0);  // 一番左に置く
  }
  sheet.clear();
  sheet.getRange(1, 1, rows.length, ASST_TASK_WIDTH).setValues(rows);
  asstStyleTasks_(sheet, styles);

  asstWriteFlags_(obs.map(function(o) { return o.mgt; }), drop.map(function(d) { return d.mgt; }));
  return { obs: obs.length, drop: drop.length };
}

function asstStyleTasks_(sheet, styles) {
  const W = ASST_TASK_WIDTH;
  sheet.setFrozenRows(0);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 260);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 70);
  styles.forEach(function(s, i) {
    const r = sheet.getRange(i + 1, 1, 1, W);
    if (s === 'title') { r.setFontWeight('bold').setFontSize(14); sheet.getRange(i + 1, W).setFontWeight('normal').setFontSize(10).setFontColor('#5f6368'); }
    else if (s === 'count') r.setFontWeight('bold').setFontSize(13);
    else if (s === 'hint') r.setFontColor('#5f6368');
    else if (s === 'section') r.setFontWeight('bold').setBackground('#e8eaf6');
    else if (s === 'header') r.setFontWeight('bold').setBackground('#f1f3f4');
    else if (s === 'late') r.setFontColor('#b3261e');
    else if (s === 'none') r.setFontColor('#5f6368');
  });
  styles.forEach(function(s, i) {
    if (s === 'item' || s === 'late') {
      sheet.getRange(i + 1, 3).setNumberFormat('m/d');
    }
  });
  // 値下げの表の「現在価格」「最終値下げ日」、観測の表の「現在価格」
  let section = '';
  styles.forEach(function(s, i) {
    if (s === 'section') section = section ? 'drop' : 'obs';
    if (s !== 'item' && s !== 'late') return;
    if (section === 'obs') sheet.getRange(i + 1, 8).setNumberFormat('#,##0');
    else { sheet.getRange(i + 1, 6).setNumberFormat('#,##0'); sheet.getRange(i + 1, 7).setNumberFormat('m/d'); }
  });
}

// ===== A列の色分け =====

/** 色分け用の非表示シートに、今日の観測・値下げの管理番号を書く（A列=観測、B列=値下げ） */
function asstWriteFlags_(obsMgts, dropMgts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ASST_FLAGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ASST_FLAGS_SHEET);
    sheet.hideSheet();
  }
  sheet.clearContents();
  const n = Math.max(obsMgts.length, dropMgts.length);
  const rows = [['観測', '値下げ']];
  for (let i = 0; i < n; i++) rows.push([obsMgts[i] || '', dropMgts[i] || '']);
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

/**
 * 商品管理シートの A列に、アシスタントの色ルールを2つ入れる（すでにあれば入れ直す）。
 * スクールのルールはそのままの順番で後ろに残す。A列だけが対象なので、ほかの列の色は今までどおり。
 */
function asstApplyHighlight_() {
  const sheet = asstSheet_(ASST_ITEMS_SHEET);
  if (!sheet) return;
  const flags = "'" + ASST_FLAGS_SHEET + "'";
  const range = sheet.getRange(2, 1, Math.max(1, sheet.getMaxRows() - 1), 1);
  const obsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",COUNTIF(INDIRECT("' + flags + '!A:A"),$A2)>0)')
    .setBackground('#ffd600').setFontColor('#202124').setBold(true)
    .setRanges([range]).build();
  const dropRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",COUNTIF(INDIRECT("' + flags + '!B:B"),$A2)>0)')
    .setBackground('#c62828').setFontColor('#ffffff').setBold(true)
    .setRanges([range]).build();

  const others = sheet.getConditionalFormatRules().filter(function(rule) {
    const c = rule.getBooleanCondition();
    const values = c ? c.getCriteriaValues() : [];
    return !values.some(function(v) { return String(v).indexOf(ASST_FLAGS_SHEET) >= 0; });
  });
  // 観測を先に（観測 → 値下げ の順に作業するため、両方ある日は黄色を表示）
  sheet.setConditionalFormatRules([obsRule, dropRule].concat(others));
}

/**
 * メニューから観測・値下げ画面を開くときに、選んでいる行の管理番号を返す。
 * 今日の物販タスク・商品管理・アシスタント管理などのシートで行を選んでいれば、その商品で開く。
 */
function asstActiveMgt_() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const row = sheet.getActiveCell().getRow();
    if (row < 2) return '';
    let col = 1;
    if (sheet.getName() !== ASST_TASKS_SHEET) {
      const map = asstHeaderMap_(sheet);
      if (!('管理番号' in map)) return '';
      col = map['管理番号'] + 1;
    }
    return String(sheet.getRange(row, col).getValue() || '').trim();
  } catch (_) {
    return '';
  }
}
