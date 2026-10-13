/**
 * 古着物販アシスタント Ver.1 — 段階1：導入・初期設定・商品リストの同期
 *
 * ★ スクールの onOpen / onEdit とぶつからないよう、メニューは
 *   「インストール型トリガー」で asstOnOpen を呼んで追加する（asstInstall を1回だけ実行）。
 * ★ 段階1は商品管理シートを読むだけで、書き込まない。
 */

// ===== 導入（エディタから1回だけ実行）=====

/**
 * Apps Script のエディタで、この関数を選んで「実行」する（最初に1回だけ）。
 * スプレッドシートを開いたときに「物販アシスタント」メニューが出るようになる。
 */
function asstInstall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'asstOnOpen';
  });
  if (!exists) {
    ScriptApp.newTrigger('asstOnOpen').forSpreadsheet(ss).onOpen().create();
  }
  try { asstBuildMenu_(); } catch (_) {}
  Logger.log(exists
    ? 'すでに導入済みです。スプレッドシートを再読み込みしてください。'
    : '導入しました。スプレッドシートを再読み込みすると「物販アシスタント」メニューが出ます。');
}

/** スプレッドシートを開いたときにトリガーから呼ばれる */
function asstOnOpen(e) {
  try { asstBuildMenu_(); } catch (err) { console.warn('asstOnOpen:', err); }
}

// メニューの中身は Assistant_Menu.gs の asstBuildMenu_ にある

function asstOpenSettings() { asstActivate_(ASST_SHEETS.SETTINGS); }
function asstOpenManage() { asstActivate_(ASST_SHEETS.MANAGE); }

function asstActivate_(name) {
  const sheet = asstSheet_(name);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「' + name + '」シートがまだありません。先に「初期設定」を実行してください。');
    return;
  }
  sheet.activate();
}

// ===== 初期設定 =====

/**
 * メニュー「初期設定」。
 * 1. 新しいシートを作る（すでにあれば作らない）
 * 2. 今ある在庫をアシスタント管理シートに登録する（登録済みの商品はそのまま）
 * 3. 入力済みの7日後データを定点観測ログにコピーする（コピー済みの商品は飛ばす）
 * 何回実行しても、同じ商品が二重に登録されることはない。
 */
function asstInitialize() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🔧 物販アシスタントの初期設定',
    '次のことを行います。\n\n' +
    '・新しいシートを4つ作る（設定・アシスタント管理・定点観測ログ・値下げ履歴）\n' +
    '・今ある在庫をアシスタント管理シートに登録する\n' +
    '・入力済みの7日後データを定点観測ログにコピーする\n\n' +
    '商品管理シートは読むだけで、何も書き換えません。\n実行してよろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  let items;
  try {
    items = asstReadItems_();
  } catch (err) {
    ui.alert('❌ 初期設定を中止しました', String(err.message || err), ui.ButtonSet.OK);
    return;
  }

  const lock = LockService.getDocumentLock();
  let sync, logCount;
  try {
    lock.waitLock(15000);
    asstEnsureSheets_();
    sync = asstSyncItems_(items);
    logCount = asstImportDay7Log_(items);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }

  const manage = asstSheet_(ASST_SHEETS.MANAGE);
  if (manage) manage.activate();
  ui.alert(
    '✅ 初期設定が終わりました',
    'アシスタント管理に追加：' + sync.added + '件\n' +
    '（7日間様子見 ' + sync.byStatus[ASST_STATUS.WAIT] + '件・' +
    '定期値下げ ' + sync.byStatus[ASST_STATUS.MARKDOWN] + '件・' +
    '観測待ち ' + sync.byStatus[ASST_STATUS.PENDING] + '件）\n' +
    '定点観測ログに移した7日後データ：' + logCount + '件\n\n' +
    '「確認メモ」に文字が入っている商品は、内容を確認してください：' + sync.needsCheck + '件',
    ui.ButtonSet.OK
  );
}

/** 新しいシートを作る（すでにあるシートは触らない） */
function asstEnsureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let settings = ss.getSheetByName(ASST_SHEETS.SETTINGS);
  if (!settings) {
    settings = ss.insertSheet(ASST_SHEETS.SETTINGS);
    const rows = [['項目（名前は変えないでください）', '値', '説明']].concat(ASST_SETTINGS_DEFAULTS);
    settings.getRange(1, 1, rows.length, 3).setValues(rows);
    asstStyleHeader_(settings, 3);
    settings.setColumnWidth(1, 260);
    settings.setColumnWidth(2, 80);
    settings.setColumnWidth(3, 420);
    settings.getRange(2, 2, ASST_SETTINGS_DEFAULTS.length, 1).setBackground('#fff8e7');
  }

  asstEnsureTable_(ss, ASST_SHEETS.MANAGE, ASST_MANAGE_HEADERS, function(sheet) {
    sheet.setFrozenColumns(2);
    sheet.setColumnWidth(2, 260);
    sheet.setColumnWidth(14, 320);
  });
  asstEnsureTable_(ss, ASST_SHEETS.LOG, ASST_LOG_HEADERS, function(sheet) {
    sheet.setColumnWidth(3, 240);
  });
  asstEnsureTable_(ss, ASST_SHEETS.PRICE_LOG, ASST_PRICE_HEADERS, function(sheet) {
    sheet.setColumnWidth(3, 240);
  });
}

function asstEnsureTable_(ss, name, headers, styleFn) {
  if (ss.getSheetByName(name)) return;
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  asstStyleHeader_(sheet, headers.length);
  if (styleFn) styleFn(sheet);
}

function asstStyleHeader_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setFontWeight('bold')
    .setBackground('#e8eaf6')
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
}

// ===== 商品リストの同期 =====

/**
 * メニュー「商品リストを同期」。
 * 商品管理シートの変化（新しく出品した商品・売れた商品）をアシスタント管理シートに反映する。
 */
function asstSyncItems() {
  const ui = SpreadsheetApp.getUi();
  if (!asstSheet_(ASST_SHEETS.MANAGE)) {
    ui.alert('先にメニュー「初期設定」を実行してください。');
    return;
  }
  let result;
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
    result = asstSyncItems_(asstReadItems_());
  } catch (err) {
    ui.alert('❌ 同期できませんでした', String(err.message || err), ui.ButtonSet.OK);
    return;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
  SpreadsheetApp.getActive().toast(
    '追加 ' + result.added + '件 ・ 売却済みにした ' + result.sold + '件 ・ 対象外にした ' + result.out + '件',
    '🔄 商品リストを同期しました', 6
  );
}

/**
 * アシスタント管理シートを商品管理シートに合わせる。
 * - まだ登録されていない在庫 → 追加（日付とステータスを自動で決める）
 * - 売れた商品 → ステータスを「売却済み」
 * - 再出品で置き換わった行・無料交換 → ステータスを「対象外」
 * - 商品名・出品日 → 商品管理シートの内容に合わせる
 * 観測・値下げの日付やステータスは、ここでは上書きしない。
 */
function asstSyncItems_(items) {
  const sheet = asstSheet_(ASST_SHEETS.MANAGE);
  const H = asstIndex_(ASST_MANAGE_HEADERS);
  const width = ASST_MANAGE_HEADERS.length;
  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];

  const rowByMgt = {};
  rows.forEach(function(r, i) {
    const mgt = String(r[H['管理番号']] || '').trim();
    if (mgt) rowByMgt[mgt] = i;
  });

  const settings = asstGetSettings_();
  const today = asstToday_();
  const now = new Date();
  const result = { added: 0, sold: 0, out: 0, needsCheck: 0, byStatus: {} };
  Object.keys(ASST_STATUS).forEach(function(k) { result.byStatus[ASST_STATUS[k]] = 0; });

  const newRows = [];
  items.forEach(function(item) {
    if (!item.listDate) return;
    const idx = rowByMgt[item.mgt];

    if (idx === undefined) {
      if (asstIsSold_(item) || asstIsOut_(item)) return;  // 売れた商品は新しく登録しない
      const row = asstBuildInitialRow_(item, today, settings);
      row[H['更新日時']] = now;
      newRows.push(row);
      result.added++;
      result.byStatus[row[H['ステータス']]]++;
      if (row[H['確認メモ']]) result.needsCheck++;
      return;
    }

    const r = rows[idx];
    let changed = false;
    if (r[H['商品名']] !== item.name) { r[H['商品名']] = item.name; changed = true; }
    const curDate = asstDay_(r[H['出品日']]);
    if (!curDate || curDate.getTime() !== item.listDate.getTime()) { r[H['出品日']] = item.listDate; changed = true; }

    const status = r[H['ステータス']];
    if (asstIsSold_(item) && status !== ASST_STATUS.SOLD) {
      r[H['ステータス']] = ASST_STATUS.SOLD;
      result.sold++;
      changed = true;
    } else if (!asstIsSold_(item) && asstIsOut_(item) && status !== ASST_STATUS.OUT) {
      r[H['ステータス']] = ASST_STATUS.OUT;
      result.out++;
      changed = true;
    }
    if (changed) r[H['更新日時']] = now;
  });

  if (rows.length) sheet.getRange(2, 1, rows.length, width).setValues(rows);
  if (newRows.length) {
    sheet.getRange(rows.length + 2, 1, newRows.length, width).setValues(newRows);
  }
  asstFormatManage_(sheet);
  return result;
}

/**
 * 新しく登録する商品の1行を作る。
 * - 出品から7日未満 → 7日間様子見
 * - 値下げ回数が1回以上 → 定期値下げ（次回値下げ日 = メモの最後の日付 + 4日。過ぎていれば今日）
 * - それ以外で7日を過ぎている → 観測待ち（次の観測で対応を選ぶ）
 * 現在価格は、メモの最後の「旧価格→新価格」から読む。読めなければ出品価格を仮に入れる。
 */
function asstBuildInitialRow_(item, today, settings) {
  const H = asstIndex_(ASST_MANAGE_HEADERS);
  const row = new Array(ASST_MANAGE_HEADERS.length).fill('');
  const obsDays = settings[ASST_SETTING.OBS_DAYS];
  const dropDays = settings[ASST_SETTING.DROP_DAYS];
  const elapsed = asstDaysBetween_(item.listDate, today);
  const hasDay7 = asstIsNumber_(item.views7d) && asstIsNumber_(item.search7d);
  const memo = asstParseMemo_(item.memo, item.listDate, today);
  const notes = [];

  row[H['管理番号']] = item.mgt;
  row[H['商品名']] = item.name;
  row[H['出品日']] = item.listDate;

  // 次回定点観測日：出品日 + 7日の倍数のうち、今日以降でいちばん近い日
  let n = Math.max(1, Math.ceil(elapsed / obsDays));
  if (n === 1 && hasDay7) n = 2;  // 7日後の観測は済んでいる
  row[H['次回定点観測日']] = asstAddDays_(item.listDate, n * obsDays);
  if (hasDay7) {
    row[H['最終観測日']] = asstAddDays_(item.listDate, obsDays);
    row[H['観測回数']] = 1;
  } else {
    row[H['観測回数']] = 0;
  }

  // ステータスと値下げ
  if (elapsed < obsDays) {
    row[H['ステータス']] = ASST_STATUS.WAIT;
  } else if (item.drops >= 1) {
    row[H['ステータス']] = ASST_STATUS.MARKDOWN;
    row[H['定期値下げ']] = '対象';
    if (memo.lastDate) {
      row[H['最終値下げ日']] = memo.lastDate;
      const next = asstAddDays_(memo.lastDate, dropDays);
      row[H['次回値下げ予定日']] = next < today ? today : next;
    } else {
      row[H['次回値下げ予定日']] = today;
      notes.push('最終値下げ日をメモから読み取れませんでした（次回値下げ日を今日にしています）');
    }
  } else {
    row[H['ステータス']] = ASST_STATUS.PENDING;  // 次の観測で対応を選ぶ
  }

  // 現在価格
  if (memo.lastPrice) {
    row[H['現在価格']] = memo.lastPrice;
  } else {
    row[H['現在価格']] = asstIsNumber_(item.listPrice) ? Number(item.listPrice) : '';
    if (item.drops >= 1) notes.push('現在価格をメモから読み取れませんでした（出品価格を仮に入れています）');
  }

  row[H['確認メモ']] = notes.join(' / ');
  return row;
}

/**
 * メモ欄から「いちばん新しい日付」と「最後に書かれた値下げ後の価格」を読み取る。
 * 価格は「3480→2980」「2890円に値下げ」「3080値下げ」「2480円に変更」「10/8,2480円」の書き方に対応し、
 * メモの後ろにあるものを使う。
 * 例：「10/1,3980→3480 / 10/8,3480→2980」→ 日付 10/8、価格 2980
 */
function asstParseMemo_(memo, listDate, today) {
  const s = String(memo || '');
  let lastDate = null;
  const dateRe = /(\d{1,2})\/(\d{1,2})/g;
  let m;
  while ((m = dateRe.exec(s)) !== null) {
    const mo = Number(m[1]), d = Number(m[2]);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) continue;
    let dt = new Date(listDate.getFullYear(), mo - 1, d);
    if (dt < listDate) dt = new Date(listDate.getFullYear() + 1, mo - 1, d);  // 年をまたいだ場合
    if (dt > today) continue;
    if (!lastDate || dt > lastDate) lastDate = dt;
  }

  let lastPrice = null, lastPos = -1;
  const patterns = [
    { re: /([\d,]{3,7})\s*円?\s*→\s*([\d,]{3,7})/g, group: 2 },          // 3480→2980
    { re: /(\d[\d,]{2,6})\s*円?\s*[、,]?\s*(?:に|へ)?\s*(?:値下げ|変更)/g, group: 1 },  // 2890円に値下げ / 3080値下げ / 2480円に変更
    { re: /\d{1,2}\/\d{1,2}\s*[,、]?\s*([\d,]{3,7})\s*円/g, group: 1 },  // 10/8,2480円（その日の価格）
  ];
  patterns.forEach(function(p) {
    let pm;
    while ((pm = p.re.exec(s)) !== null) {
      const price = Number(pm[p.group].replace(/,/g, ''));
      if (price >= 300 && pm.index >= lastPos) { lastPrice = price; lastPos = pm.index; }
    }
  });
  return { lastDate: lastDate, lastPrice: lastPrice };
}

/** 見出しの名前 → 位置 の対応表（新しいシート用） */
function asstIndex_(headers) {
  const map = {};
  headers.forEach(function(h, i) { map[h] = i; });
  return map;
}

function asstFormatManage_(sheet) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  const H = asstIndex_(ASST_MANAGE_HEADERS);
  ['出品日', '次回定点観測日', '最終観測日', '最終値下げ日', '次回値下げ予定日', '最終改善日'].forEach(function(h) {
    sheet.getRange(2, H[h] + 1, n, 1).setNumberFormat('yyyy/mm/dd');
  });
  sheet.getRange(2, H['現在価格'] + 1, n, 1).setNumberFormat('#,##0');
  sheet.getRange(2, H['更新日時'] + 1, n, 1).setNumberFormat('yyyy/mm/dd hh:mm');
}

// ===== 7日後データの移行 =====

/**
 * 商品管理シートの「7日後閲覧数・7日後検索数・いいね数」を、
 * 定点観測ログの最初の記録（経過日数7）としてコピーする。売れた商品も含める。
 * すでにログに経過日数7の記録がある商品は飛ばす。
 */
function asstImportDay7Log_(items) {
  const sheet = asstSheet_(ASST_SHEETS.LOG);
  const H = asstIndex_(ASST_LOG_HEADERS);
  const width = ASST_LOG_HEADERS.length;
  const lastRow = sheet.getLastRow();

  const done = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, width).getValues().forEach(function(r) {
      if (Number(r[H['経過日数']]) === 7) done[String(r[H['管理番号']]).trim()] = true;
    });
  }

  const rows = [];
  items.forEach(function(item) {
    if (!item.listDate || done[item.mgt]) return;
    if (!asstIsNumber_(item.views7d) || !asstIsNumber_(item.search7d)) return;
    const views = Number(item.views7d), search = Number(item.search7d);
    const likes = asstIsNumber_(item.likes) ? Number(item.likes) : '';
    const row = new Array(width).fill('');
    row[H['計測日']] = asstAddDays_(item.listDate, 7);
    row[H['管理番号']] = item.mgt;
    row[H['商品名']] = item.name;
    row[H['経過日数']] = 7;
    row[H['閲覧数（合計）']] = views;
    row[H['検索数（合計）']] = search;
    row[H['いいね数（合計）']] = likes;
    row[H['閲覧数（週の増加）']] = views;
    row[H['検索数（週の増加）']] = search;
    row[H['いいね数（週の増加）']] = likes;
    row[H['閲覧率（%）']] = search > 0 ? Math.round(views / search * 1000) / 10 : '';
    row[H['その時の価格']] = asstIsNumber_(item.listPrice) ? Number(item.listPrice) : '';
    row[H['記録元']] = '初期設定（商品管理の7日後データから移行）';
    rows.push(row);
  });

  if (rows.length) {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, width).setValues(rows);
    sheet.getRange(start, H['計測日'] + 1, rows.length, 1).setNumberFormat('yyyy/mm/dd');
    sheet.getRange(start, H['その時の価格'] + 1, rows.length, 1).setNumberFormat('#,##0');
  }
  return rows.length;
}
