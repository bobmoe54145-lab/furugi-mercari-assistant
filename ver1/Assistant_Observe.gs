/**
 * 古着物販アシスタント Ver.1 — 段階2：定点観測
 *
 * - 観測画面：閲覧数・検索数・いいね数（出品からの合計）を入れると、
 *   前回からの増加・閲覧率・おすすめを表示し、選んだ対応と一緒に記録する。
 * - 定点観測ログ：1回の観測 = 1行。上書きしない。
 * - 商品管理シート：最初の観測（7日後）のときだけ、AD・AE・AG が空なら書き込む。上書きしない。
 * - 観測まとめ：ログから、1商品1行で 7日後・14日後・21日後… を横に並べて作り直す。
 */

const ASST_SUMMARY_SHEET = '観測まとめ';

// 観測後に選べる対応（先頭が初期値）
const ASST_ACTIONS = ['定期値下げ', '写真変更', '商品名変更', '商品説明変更', '複数項目変更', '今回は変更なし', '値下げ停止', 'その他'];
const ASST_IMPROVE_ACTIONS = ['写真変更', '商品名変更', '商品説明変更', '複数項目変更'];

// ===== メニューから開く =====

function asstOpenObservation() {
  const ui = SpreadsheetApp.getUi();
  if (!asstSheet_(ASST_SHEETS.MANAGE)) {
    ui.alert('先にメニュー「初期設定」を実行してください。');
    return;
  }
  try {
    asstSyncItems_(asstReadItems_());  // 売れた商品・新しい商品を反映してから開く
  } catch (err) {
    ui.alert('❌ 商品リストを読めませんでした', String(err.message || err), ui.ButtonSet.OK);
    return;
  }
  // 今日の物販タスクや商品管理で行を選んでいれば、その商品を選んだ状態で開く
  const selectMgt = typeof asstActiveMgt_ === 'function' ? asstActiveMgt_() : '';
  const html = HtmlService.createHtmlOutput(asstObserveHtml_(selectMgt)).setWidth(760).setHeight(680);
  ui.showModalDialog(html, '📊 定点観測を記録');
}

function asstRefreshSummary() {
  if (!asstSheet_(ASST_SHEETS.LOG)) {
    SpreadsheetApp.getUi().alert('先にメニュー「初期設定」を実行してください。');
    return;
  }
  const sheet = asstRebuildSummary_();
  sheet.activate();
  SpreadsheetApp.getActive().toast('観測まとめを作り直しました', '📈 観測まとめ', 4);
}

// ===== 観測画面に渡すデータ =====

/** 観測画面（ブラウザ側）から呼ばれる。日付は文字列で渡す */
function asstGetObservationData() {
  const settings = asstGetSettings_();
  const today = asstToday_();
  const manage = asstReadManage_();
  const logs = asstReadLogsByMgt_();
  const urls = asstReadItemUrls_();
  const drops = asstReadDropsByMgt_();
  const H = manage.H;

  const list = [];
  manage.rows.forEach(function(r) {
    const status = r[H['ステータス']];
    if (status === ASST_STATUS.SOLD || status === ASST_STATUS.OUT) return;
    const mgt = String(r[H['管理番号']]).trim();
    const listDate = asstDay_(r[H['出品日']]);
    if (!mgt || !listDate) return;
    const nextObs = asstDay_(r[H['次回定点観測日']]);
    const history = logs[mgt] || [];
    const last = history.length ? history[history.length - 1] : null;
    list.push({
      mgt: mgt,
      name: String(r[H['商品名']] || ''),
      status: status,
      listDate: asstFmt_(listDate),
      elapsed: asstDaysBetween_(listDate, today),
      price: r[H['現在価格']],
      nextObs: nextObs ? asstFmt_(nextObs) : '',
      nextObsDays: nextObs ? asstDaysBetween_(listDate, nextObs) : '',
      due: !!nextObs && nextObs <= today,
      improve: String(r[H['改善内容']] || ''),
      changes: asstChangesSince_(drops[mgt] || [], last ? last.date : listDate, r, H),
      url: urls[mgt] || '',
      last: last ? {
        date: asstFmt_(last.date), elapsed: last.elapsed,
        views: last.views, search: last.search, likes: last.likes,
        incViews: last.incViews, incSearch: last.incSearch, incLikes: last.incLikes,
      } : null,
    });
  });

  // 今日・期限切れの商品を先に、その中は予定日の早い順
  list.sort(function(a, b) {
    if (a.due !== b.due) return a.due ? -1 : 1;
    return String(a.nextObs).localeCompare(String(b.nextObs));
  });

  return {
    today: asstFmt_(today),
    actions: ASST_ACTIONS,
    improveActions: ASST_IMPROVE_ACTIONS,
    settings: {
      obsDays: settings[ASST_SETTING.OBS_DAYS],
      searchLow: settings[ASST_SETTING.SEARCH_LOW],
      viewRateLow: settings[ASST_SETTING.VIEW_RATE_LOW],
      likeViewsMin: settings[ASST_SETTING.LIKE_VIEWS_MIN],
      likesLow: settings[ASST_SETTING.LIKES_LOW],
    },
    items: list,
  };
}

// ===== 記録する =====

/**
 * 観測画面の「記録する」から呼ばれる。
 * @param {{mgt:string, views:number, search:number, likes:number, action:string, change:string, memo:string}} p
 */
function asstSaveObservation(p) {
  const views = asstToCount_(p.views), search = asstToCount_(p.search), likes = asstToCount_(p.likes);
  if (views === null || search === null || likes === null) {
    throw new Error('閲覧数・検索数・いいね数は、0以上の整数で入力してください。');
  }
  if (ASST_ACTIONS.indexOf(p.action) < 0) throw new Error('対応を選んでください。');

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const settings = asstGetSettings_();
    const obsDays = settings[ASST_SETTING.OBS_DAYS];
    const today = asstToday_();
    const manage = asstReadManage_();
    const H = manage.H;
    const idx = manage.rows.findIndex(function(r) { return String(r[H['管理番号']]).trim() === p.mgt; });
    if (idx < 0) throw new Error('アシスタント管理シートに ' + p.mgt + ' が見つかりません。');
    const row = manage.rows[idx];
    const listDate = asstDay_(row[H['出品日']]);
    const elapsed = asstDaysBetween_(listDate, today);
    if (elapsed < obsDays) {
      throw new Error('まだ出品から' + elapsed + '日目です。最初の' + obsDays + '日間は様子見期間なので記録できません。');
    }

    // 前回の記録との差（その週に増えた数）
    const history = asstReadLogsByMgt_()[p.mgt] || [];
    const prev = history.length ? history[history.length - 1] : null;
    if (prev && (views < prev.views || search < prev.search)) {
      throw new Error('前回の記録（閲覧数 ' + prev.views + '・検索数 ' + prev.search + '）より少ない数字です。\n' +
        'メルカリの数字は出品からの合計なので、前回以上になるはずです。入力を確かめてください。');
    }
    const inc = prev
      ? { views: views - prev.views, search: search - prev.search, likes: likes - prev.likes }
      : { views: views, search: search, likes: likes };
    const rate = inc.search > 0 ? Math.round(inc.views / inc.search * 1000) / 10 : '';
    const rec = asstRecommend_(inc, settings);

    // 定点観測ログに1行追加
    const log = asstSheet_(ASST_SHEETS.LOG);
    const L = asstIndex_(ASST_LOG_HEADERS);
    const logRow = new Array(ASST_LOG_HEADERS.length).fill('');
    logRow[L['計測日']] = today;
    logRow[L['管理番号']] = p.mgt;
    logRow[L['商品名']] = row[H['商品名']];
    logRow[L['経過日数']] = elapsed;
    logRow[L['閲覧数（合計）']] = views;
    logRow[L['検索数（合計）']] = search;
    logRow[L['いいね数（合計）']] = likes;
    logRow[L['閲覧数（週の増加）']] = inc.views;
    logRow[L['検索数（週の増加）']] = inc.search;
    logRow[L['いいね数（週の増加）']] = inc.likes;
    logRow[L['閲覧率（%）']] = rate;
    logRow[L['その時の価格']] = row[H['現在価格']];
    logRow[L['前回から行った変更']] = String(p.change || '').trim();
    logRow[L['システムのおすすめ']] = rec;
    logRow[L['実際に行った対応']] = p.action;
    logRow[L['メモ']] = String(p.memo || '').trim();
    logRow[L['記録元']] = '観測画面';
    const logAt = log.getLastRow() + 1;
    log.getRange(logAt, 1, 1, logRow.length).setValues([logRow]);
    log.getRange(logAt, L['計測日'] + 1).setNumberFormat('yyyy/mm/dd');
    log.getRange(logAt, L['その時の価格'] + 1).setNumberFormat('#,##0');

    // 最初の観測が「7日後」とみなせる日（7日後に一番近い）なら、
    // 商品管理の 7日後閲覧数・7日後検索数・いいね数 へ（空欄のときだけ）
    let wroteItems = false;
    if (!prev && Math.round(elapsed / obsDays) === 1) wroteItems = asstWriteDay7ToItems_(p.mgt, views, search, likes);

    // アシスタント管理の更新
    const before = row[H['ステータス']];
    const isImprove = ASST_IMPROVE_ACTIONS.indexOf(p.action) >= 0;
    row[H['最終観測日']] = today;
    row[H['観測回数']] = (Number(row[H['観測回数']]) || 0) + 1;
    row[H['次回定点観測日']] = isImprove
      ? asstAddDays_(today, obsDays)                                        // 変更後ちょうど7日で比べる
      : asstAddDays_(listDate, (Math.floor(elapsed / obsDays) + 1) * obsDays); // 出品日から7日ごとの次の日
    let dropToday = false;
    if (p.action === '定期値下げ') {
      row[H['ステータス']] = ASST_STATUS.MARKDOWN;
      row[H['定期値下げ']] = '対象';
      if (before !== ASST_STATUS.MARKDOWN || !asstDay_(row[H['次回値下げ予定日']])) {
        row[H['次回値下げ予定日']] = today;  // 定期値下げを選んだ日に1回目の値下げ
      }
      dropToday = asstDay_(row[H['次回値下げ予定日']]) <= today;
    } else if (isImprove) {
      row[H['ステータス']] = ASST_STATUS.IMPROVE;
      row[H['最終改善日']] = today;
      row[H['改善内容']] = p.action;
      row[H['定期値下げ']] = '一時停止';
      row[H['次回値下げ予定日']] = '';    // 変更後7日間は値下げしない
    } else if (p.action === '値下げ停止') {
      row[H['ステータス']] = ASST_STATUS.STOPPED;
      row[H['定期値下げ']] = '停止';
      row[H['次回値下げ予定日']] = '';
    } else {
      // 今回は変更なし・その他：定期値下げ中ならそのまま続ける
      if (before !== ASST_STATUS.MARKDOWN && before !== ASST_STATUS.STOPPED) {
        row[H['ステータス']] = ASST_STATUS.PENDING;
        row[H['次回値下げ予定日']] = '';
      }
    }
    row[H['更新日時']] = new Date();
    manage.sheet.getRange(idx + 2, 1, 1, row.length).setValues([row]);
    asstFormatManage_(manage.sheet);

    asstRebuildSummary_();
    if (typeof asstRebuildTasks_ === 'function') {
      try { asstRebuildTasks_(); } catch (err) { console.warn('asstRebuildTasks_:', err); }
    }

    return {
      recommendation: rec,
      wroteItems: wroteItems,
      dropToday: dropToday,
      nextObs: asstFmt_(asstDay_(row[H['次回定点観測日']])),
    };
  } finally {
    lock.releaseLock();
  }
}

/** 判定ルール（数字は「その週に増えた数」）。基準は設定シートの値 */
function asstRecommend_(inc, settings) {
  const rate = inc.search > 0 ? inc.views / inc.search * 100 : 0;
  if (inc.search < settings[ASST_SETTING.SEARCH_LOW]) return '商品名・検索キーワードの改善';
  if (rate < settings[ASST_SETTING.VIEW_RATE_LOW]) return '1枚目の写真・商品名の改善';
  if (inc.views >= settings[ASST_SETTING.LIKE_VIEWS_MIN] && inc.likes <= settings[ASST_SETTING.LIKES_LOW]) {
    return '価格・商品の見せ方の確認';
  }
  return '定期値下げ';
}

/** 商品管理の AD・AE・AG に、空欄のときだけ書き込む。1つでも書けたら true */
function asstWriteDay7ToItems_(mgt, views, search, likes) {
  const sheet = asstSheet_(ASST_ITEMS_SHEET);
  const map = asstHeaderMap_(sheet);
  const C = ASST_ITEM_COLS;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const ids = sheet.getRange(2, map[C.MGT] + 1, lastRow - 1, 1).getValues();
  const i = ids.findIndex(function(r) { return String(r[0]).trim() === mgt; });
  if (i < 0) return false;
  const rowNum = i + 2;
  let wrote = false;
  [[C.VIEWS_7D, views], [C.SEARCH_7D, search], [C.LIKES, likes]].forEach(function(pair) {
    const cell = sheet.getRange(rowNum, map[pair[0]] + 1);
    if (cell.getValue() === '') { cell.setValue(pair[1]); wrote = true; }
  });
  return wrote;
}

/**
 * 前回の観測から今日までに行った変更（値下げ履歴・改善内容）を文章にする。
 * 観測画面の「前回から行った変更」の初期値になる（画面で書き直せる）。
 */
function asstChangesSince_(drops, since, manageRow, H) {
  const parts = [];
  drops.forEach(function(d) {
    if (d.date >= since) parts.push('値下げ ' + d.from + '→' + d.to);
  });
  const improvedAt = asstDay_(manageRow[H['最終改善日']]);
  if (improvedAt && improvedAt >= since && manageRow[H['改善内容']]) parts.push(String(manageRow[H['改善内容']]));
  return parts.join('、');
}

/** 値下げ履歴の「値下げ」の行を 管理番号 → [{date, from, to}] にまとめる（1回だけ読む） */
function asstReadDropsByMgt_() {
  const out = {};
  const sheet = asstSheet_(ASST_SHEETS.PRICE_LOG);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const P = asstIndex_(ASST_PRICE_HEADERS);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, ASST_PRICE_HEADERS.length).getValues().forEach(function(r) {
    const mgt = String(r[P['管理番号']] || '').trim();
    const date = asstDay_(r[P['日付']]);
    if (!mgt || !date || r[P['操作']] !== '値下げ') return;
    (out[mgt] = out[mgt] || []).push({ date: date, from: r[P['旧価格']], to: r[P['新価格']] });
  });
  return out;
}

// ===== 観測まとめ =====

/**
 * 定点観測ログから「観測まとめ」シートを作り直す。
 * 1商品1行。各週の列に「合計（+その週に増えた数）」を表示する。
 */
function asstRebuildSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ASST_SUMMARY_SHEET);
  if (!sheet) sheet = ss.insertSheet(ASST_SUMMARY_SHEET);

  const logs = asstReadLogsByMgt_();
  const manage = asstReadManage_();
  const H = manage.H;
  const statusByMgt = {};
  const order = [];
  manage.rows.forEach(function(r) {
    const mgt = String(r[H['管理番号']]).trim();
    if (mgt) { statusByMgt[mgt] = r[H['ステータス']]; order.push(mgt); }
  });
  Object.keys(logs).forEach(function(mgt) { if (!(mgt in statusByMgt)) order.push(mgt); });

  // 経過日数 → 何週目か（7日後=1週目）。
  // 同じ週に2回記録があるときは、後の記録を次の週に回す（次の週も埋まっていれば後の記録を使う）
  const weeksByMgt = {};
  let maxWeek = 1;
  order.forEach(function(mgt) {
    const w = {};
    (logs[mgt] || []).forEach(function(e) {
      let n = Math.max(1, Math.round(e.elapsed / 7));
      if (w[n] && !w[n + 1]) n = n + 1;
      w[n] = e;
      if (n > maxWeek) maxWeek = n;
    });
    weeksByMgt[mgt] = w;
  });
  maxWeek = Math.min(maxWeek, 26);

  const header = ['管理番号', '商品名', 'ステータス', '観測回数'];
  for (let n = 1; n <= maxWeek; n++) {
    header.push((n * 7) + '日後\n閲覧', (n * 7) + '日後\n検索', (n * 7) + '日後\nいいね', (n * 7) + '日後\n閲覧率%');
  }
  const cell = function(total, inc, first) {
    if (total === '' || total === null || total === undefined) return '';
    return first || inc === '' ? String(total) : total + '（' + (inc >= 0 ? '+' : '') + inc + '）';
  };
  const rows = [];
  order.forEach(function(mgt) {
    const history = logs[mgt] || [];
    if (!history.length) return;
    const w = weeksByMgt[mgt];
    const row = [mgt, history[history.length - 1].name, statusByMgt[mgt] || '', history.length];
    for (let n = 1; n <= maxWeek; n++) {
      const e = w[n];
      if (!e) { row.push('', '', '', ''); continue; }
      const first = history.indexOf(e) === 0;
      row.push(cell(e.views, e.incViews, first), cell(e.search, e.incSearch, first),
        cell(e.likes, e.incLikes, first), e.rate);
    }
    rows.push(row);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  asstStyleHeader_(sheet, header.length);
  sheet.setFrozenColumns(2);
  sheet.setColumnWidth(2, 240);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    sheet.getRange(2, 5, rows.length, header.length - 4).setHorizontalAlignment('right');
  }
  return sheet;
}

// ===== 読み込みの道具 =====

/** アシスタント管理シートを読む */
function asstReadManage_() {
  const sheet = asstSheet_(ASST_SHEETS.MANAGE);
  const width = ASST_MANAGE_HEADERS.length;
  const lastRow = sheet.getLastRow();
  return {
    sheet: sheet,
    H: asstIndex_(ASST_MANAGE_HEADERS),
    rows: lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [],
  };
}

/** 定点観測ログを 管理番号 → 記録の配列（古い順）にまとめる */
function asstReadLogsByMgt_() {
  const sheet = asstSheet_(ASST_SHEETS.LOG);
  const out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  const L = asstIndex_(ASST_LOG_HEADERS);
  const num = function(v) { return asstIsNumber_(v) ? Number(v) : ''; };
  sheet.getRange(2, 1, sheet.getLastRow() - 1, ASST_LOG_HEADERS.length).getValues().forEach(function(r, i) {
    const mgt = String(r[L['管理番号']] || '').trim();
    const date = asstDay_(r[L['計測日']]);
    if (!mgt || !date) return;
    (out[mgt] = out[mgt] || []).push({
      order: i, date: date, name: String(r[L['商品名']] || ''), elapsed: Number(r[L['経過日数']]) || 0,
      views: num(r[L['閲覧数（合計）']]), search: num(r[L['検索数（合計）']]), likes: num(r[L['いいね数（合計）']]),
      incViews: num(r[L['閲覧数（週の増加）']]), incSearch: num(r[L['検索数（週の増加）']]),
      incLikes: num(r[L['いいね数（週の増加）']]), rate: num(r[L['閲覧率（%）']]),
    });
  });
  Object.keys(out).forEach(function(k) {
    out[k].sort(function(a, b) { return a.date - b.date || a.order - b.order; });
  });
  return out;
}

/** 商品管理の「メルカリ商品URL/画像」を 管理番号 → URL で返す（列がなければ空） */
function asstReadItemUrls_() {
  const sheet = asstSheet_(ASST_ITEMS_SHEET);
  const map = asstHeaderMap_(sheet);
  const out = {};
  const urlCol = map['メルカリ商品URL/画像'];
  if (urlCol === undefined || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().forEach(function(r) {
    const mgt = String(r[map[ASST_ITEM_COLS.MGT]] || '').trim();
    const url = String(r[urlCol] || '').trim();
    if (mgt && /^https?:\/\//.test(url)) out[mgt] = url;
  });
  return out;
}

function asstFmt_(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd') : '';
}

function asstToCount_(v) {
  const n = Number(v);
  return v !== '' && v !== null && Number.isInteger(n) && n >= 0 ? n : null;
}

// ===== 観測画面（HTML）=====

function asstObserveHtml_(selectMgt) {
  return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><style>' +
    'body{font-family:-apple-system,"Hiragino Sans",sans-serif;font-size:13px;color:#202124;margin:0;padding:12px 16px}' +
    'h3{font-size:13px;margin:14px 0 6px;color:#3c4043}' +
    'select,input,textarea{font:inherit;padding:6px;border:1px solid #c4c7c5;border-radius:6px;box-sizing:border-box}' +
    '#item{width:100%}.row{display:flex;gap:12px}.row>div{flex:1}' +
    '.box{background:#f8f9fa;border-radius:8px;padding:8px 12px;line-height:1.7}' +
    '.num{width:100%;font-size:16px;text-align:right}' +
    'table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid #e0e0e0;padding:4px 8px;text-align:right}' +
    'th:first-child,td:first-child{text-align:left}' +
    '.rec{font-size:15px;font-weight:600;padding:8px 12px;border-radius:8px;background:#e8f0fe;color:#174ea6}' +
    '.warn{color:#b3261e}.muted{color:#5f6368}.tag{display:inline-block;padding:0 8px;border-radius:10px;background:#fce8e6;color:#b3261e;font-size:12px}' +
    'button{font:inherit;padding:8px 20px;border-radius:6px;border:0;cursor:pointer}' +
    '#save{background:#1a73e8;color:#fff;font-weight:600}#save:disabled{background:#9aa0a6}' +
    '#close{background:#f1f3f4}#msg{margin-top:8px;white-space:pre-wrap}' +
    '</style></head><body>' +
    '<div id="loading" class="muted">読み込み中…</div>' +
    '<div id="main" style="display:none">' +
    '<select id="item"></select>' +
    '<div class="box" id="info" style="margin-top:8px"></div>' +
    '<h3>メルカリで見た数字（出品からの合計）</h3>' +
    '<div class="row"><div>閲覧数<input class="num" id="views" type="number" min="0"></div>' +
    '<div>検索数<input class="num" id="search" type="number" min="0"></div>' +
    '<div>いいね数<input class="num" id="likes" type="number" min="0"></div></div>' +
    '<div id="calc" style="margin-top:10px"></div>' +
    '<h3>前回から行った変更</h3><input id="change" style="width:100%" placeholder="例：値下げ 3980→3480、写真変更">' +
    '<div class="row" style="margin-top:6px"><div><h3>実際に行う対応</h3><select id="action" style="width:100%"></select></div>' +
    '<div><h3>メモ</h3><input id="memo" style="width:100%"></div></div>' +
    '<div style="margin-top:14px;display:flex;gap:8px"><button id="save">記録する</button><button id="close">閉じる</button></div>' +
    '<div id="msg"></div></div>' +
    '<script>var SELECT=' + JSON.stringify(String(selectMgt || '')).replace(/</g, '\\u003c') + ';\n' +
    asstObserveScript_() + '</script></body></html>';
}

function asstObserveScript_() {
  return [
    'var DATA=null,CUR=null;',
    'function $(id){return document.getElementById(id);}',
    'function txt(el,s){el.textContent=s;return el;}',
    'function load(selectMgt,doneMsg,dropMgt){google.script.run.withSuccessHandler(function(d){DATA=d;render(selectMgt);if(doneMsg){$("msg").className="";txt($("msg"),doneMsg);}if(dropMgt)dropButton(dropMgt);}).withFailureHandler(fail).asstGetObservationData();}',
    'function dropButton(mgt){var b=document.createElement("button");b.style.cssText="margin-top:8px;background:#d93025;color:#fff;font-weight:600";',
    '  txt(b,"💰 "+mgt+" をこのまま値下げする");b.addEventListener("click",function(){b.disabled=true;google.script.run.withFailureHandler(fail).asstOpenMarkdownFor(mgt);});',
    '  var d=document.createElement("div");d.appendChild(b);$("msg").appendChild(d);}',
    'function fail(e){$("loading").style.display="none";$("main").style.display="block";$("msg").className="warn";txt($("msg"),"エラー："+(e&&e.message?e.message:e));$("save").disabled=false;}',
    'function render(selectMgt){',
    '  $("loading").style.display="none";$("main").style.display="block";',
    '  var sel=$("item");sel.innerHTML="";',
    '  var g1=document.createElement("optgroup");g1.label="今日・期限切れ";',
    '  var g2=document.createElement("optgroup");g2.label="そのほかの商品";',
    '  DATA.items.forEach(function(it,i){var o=document.createElement("option");o.value=i;',
    '    o.textContent=it.mgt+"  "+it.name.slice(0,26)+"  （予定 "+it.nextObs.slice(5)+"・"+it.nextObsDays+"日後）";',
    '    (it.due?g1:g2).appendChild(o);});',
    '  if(g1.children.length)sel.appendChild(g1);if(g2.children.length)sel.appendChild(g2);',
    '  if(!DATA.items.length){txt($("info"),"記録できる商品がありません。");return;}',
    '  var idx=0;DATA.items.forEach(function(it,i){if(it.mgt===selectMgt)idx=i;});sel.value=idx;',
    '  var acts=$("action");acts.innerHTML="";DATA.actions.forEach(function(a){var o=document.createElement("option");o.value=a;o.textContent=a;acts.appendChild(o);});',
    '  pick();',
    '}',
    'function pick(){',
    '  CUR=DATA.items[$("item").value];var it=CUR;var info=$("info");info.innerHTML="";',
    '  function line(label,val){var d=document.createElement("div");var b=document.createElement("span");b.className="muted";txt(b,label+"：");d.appendChild(b);d.appendChild(document.createTextNode(val));info.appendChild(d);return d;}',
    '  var t=line("商品",it.mgt+"  "+it.name);',
    '  if(it.due){var tag=document.createElement("span");tag.className="tag";txt(tag,it.nextObs<DATA.today?"期限切れ":"今日");t.appendChild(document.createTextNode(" "));t.appendChild(tag);}',
    '  line("出品日",it.listDate+"（今日で"+it.elapsed+"日目）  ステータス："+it.status+"  現在価格："+(it.price===""?"-":Number(it.price).toLocaleString()+"円"));',
    '  line("前回の記録",it.last?(it.last.date+"（"+it.last.elapsed+"日後）  閲覧 "+it.last.views+" ・ 検索 "+it.last.search+" ・ いいね "+(it.last.likes===""?"-":it.last.likes)):"まだありません（今回が最初の記録）");',
    '  if(it.url){var d=line("メルカリ","");var a=document.createElement("a");a.href=it.url;a.target="_blank";txt(a,"商品ページを開く");d.appendChild(a);}',
    '  $("views").value="";$("search").value="";$("likes").value="";$("memo").value="";$("change").value=it.changes||"";',
    '  $("action").value=DATA.actions[0];txt($("msg"),"");$("msg").className="";',
    '  $("save").disabled=it.elapsed<DATA.settings.obsDays;',
    '  if($("save").disabled){$("msg").className="warn";txt($("msg"),"まだ様子見期間です（出品から"+DATA.settings.obsDays+"日間は記録できません）。");}',
    '  calc();$("views").focus();',
    '}',
    'function n(id){var v=$(id).value;return v===""?null:Number(v);}',
    'function recommend(inc){var s=DATA.settings;var rate=inc.search>0?inc.views/inc.search*100:0;',
    '  if(inc.search<s.searchLow)return"商品名・検索キーワードの改善";',
    '  if(rate<s.viewRateLow)return"1枚目の写真・商品名の改善";',
    '  if(inc.views>=s.likeViewsMin&&inc.likes<=s.likesLow)return"価格・商品の見せ方の確認";',
    '  return"定期値下げ";}',
    'function calc(){var c=$("calc");c.innerHTML="";var v=n("views"),s=n("search"),l=n("likes");',
    '  if(v===null||s===null||l===null){txt(c,"3つの数字を入れると、増えた数とおすすめが出ます。").className="muted";return;}',
    '  c.className="";var p=CUR.last;',
    '  var inc=p?{views:v-p.views,search:s-p.search,likes:l-(p.likes===""?0:p.likes)}:{views:v,search:s,likes:l};',
    '  if(p&&(inc.views<0||inc.search<0)){var w=document.createElement("div");w.className="warn";txt(w,"前回より少ない数字があります。メルカリの数字は出品からの合計です。");c.appendChild(w);return;}',
    '  var rate=inc.search>0?Math.round(inc.views/inc.search*1000)/10:"-";',
    '  var tb=document.createElement("table");',
    '  function tr(cells,th){var r=document.createElement("tr");cells.forEach(function(x){var e=document.createElement(th?"th":"td");txt(e,x);r.appendChild(e);});tb.appendChild(r);}',
    '  var improve=CUR.status==="改善後観察"&&p&&p.incViews!=="";',
    '  if(improve){tr(["","変更前の1週間","今回（変更後）"],true);',
    '    var pr=p.incSearch>0?Math.round(p.incViews/p.incSearch*1000)/10:"-";',
    '    tr(["閲覧数の増加",p.incViews,inc.views]);tr(["検索数の増加",p.incSearch,inc.search]);tr(["いいねの増加",p.incLikes,inc.likes]);tr(["閲覧率（%）",pr,rate]);}',
    '  else{tr(["",p?"前回からの増加":"最初の1週間","閲覧率（%）"],true);tr(["閲覧数",inc.views,rate]);tr(["検索数",inc.search,""]);tr(["いいね",inc.likes,""]);}',
    '  c.appendChild(tb);',
    '  var r=document.createElement("div");r.className="rec";r.style.marginTop="8px";txt(r,"おすすめ："+recommend(inc));c.appendChild(r);',
    '}',
    'function save(){var v=n("views"),s=n("search"),l=n("likes");',
    '  if(v===null||s===null||l===null){$("msg").className="warn";txt($("msg"),"閲覧数・検索数・いいね数を入れてください。");return;}',
    '  $("save").disabled=true;$("msg").className="muted";txt($("msg"),"記録しています…");var mgt=CUR.mgt;',
    '  google.script.run.withSuccessHandler(function(res){',
    '    var m="✅ "+mgt+" を記録しました。次回の観測日："+res.nextObs;',
    '    if(res.wroteItems)m+="\\n商品管理シートの 7日後閲覧数・7日後検索数・いいね数 にも書き込みました。";',
    '    if(res.dropToday)m+="\\n💰 今日がこの商品の値下げ日です。";',
    '    var next=null;DATA.items.forEach(function(it){if(!next&&it.due&&it.mgt!==mgt)next=it.mgt;});',
    '    load(next||mgt,m,res.dropToday?mgt:null);',
    '  }).withFailureHandler(fail).asstSaveObservation({mgt:mgt,views:v,search:s,likes:l,action:$("action").value,change:$("change").value,memo:$("memo").value});',
    '}',
    '$("item").addEventListener("change",pick);',
    '["views","search","likes"].forEach(function(id){$(id).addEventListener("input",calc);});',
    '$("save").addEventListener("click",save);',
    '$("close").addEventListener("click",function(){google.script.host.close();});',
    'load(SELECT||null);',
  ].join('\n');
}
