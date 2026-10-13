/**
 * 古着物販アシスタント Ver.1 — 段階5：商品登録（スクショ読み取り）
 *
 * - スクショを Google ドライブの文字読み取り（OCR・無料）にかけ、商品名と出品価格を下書きする。
 *   ★ Apps Script の「サービス」に「Drive API」を追加しておく必要がある。
 *   読み取り用の一時ファイルは、読んだらすぐ削除する。
 * - 登録：スクールの「新規行を追加」（insertNewRowForManualEntry）で行を作るか、
 *   商品名が空の既存の行を選んで、次の列に書き込む。
 *     商品名・ジャンル・系統・ブランド・出品価格・商品の状態・季節・仕入れ記録ID・メルカリ商品URL/画像
 *     （出品日は、画面で変えたときだけ書き込む）
 *   仕入れ記録ID からの自動入力（仕入れ日・仕入れ値・付帯費・仕入れ場所）は、
 *   スクールの applyPurchaseRecAutoFill_ をそのまま呼ぶ。
 */

const ASST_REG_COLS = {
  NAME: '商品名', GENRE: 'ジャンル', TASTE: '系統', BRAND: 'ブランド', PRICE: '出品価格',
  CONDITION: '商品の状態', SEASON: '季節', REC_ID: '仕入れ記録ID', URL: 'メルカリ商品URL/画像',
};

// プルダウンが読めなかったときの候補（スクールの初期値と同じ）
const ASST_REG_DEFAULT_LISTS = {
  'ジャンル': ['トップス', 'ボトムス', 'アウター', 'ワンピース', '小物', 'その他'],
  '系統': ['古着', 'ストリート', 'きれいめ', 'スポーツ', 'Y2K', 'その他'],
  '商品の状態': ['新品', '未使用', '目立った傷なし', 'やや傷あり'],
  '季節': ['春夏', '秋冬', 'オールシーズン'],
};

// ===== メニューから開く =====

function asstOpenRegister() {
  const ui = SpreadsheetApp.getUi();
  if (!asstSheet_(ASST_SHEETS.MANAGE)) {
    ui.alert('先にメニュー「初期設定」を実行してください。');
    return;
  }
  const html = HtmlService.createHtmlOutput(asstRegisterHtml_()).setWidth(760).setHeight(720);
  ui.showModalDialog(html, '📝 商品を登録');
}

// ===== 画面に渡すデータ =====

function asstGetRegisterData() {
  const sheet = asstSheet_(ASST_ITEMS_SHEET);
  const map = asstHeaderMap_(sheet);
  const C = ASST_REG_COLS;
  const missing = asstMissingHeaders_(map, [C.NAME, C.GENRE, C.TASTE, C.BRAND, C.PRICE, C.CONDITION, C.SEASON, '管理番号', '出品日']);
  if (missing.length) throw new Error('商品管理シートに次の見出しが見つかりません：' + missing.join('、'));

  const lists = {};
  [C.GENRE, C.TASTE, C.CONDITION, C.SEASON].forEach(function(h) {
    lists[h] = asstListFromValidation_(sheet, map[h] + 1) || ASST_REG_DEFAULT_LISTS[h];
  });

  // ブランドの候補（よく使う順）と、商品名が空の行（新規行を一括追加した行など）
  const lastRow = sheet.getLastRow();
  const values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
  const brandCount = {};
  const blanks = [];
  values.forEach(function(r) {
    const brand = String(r[map[C.BRAND]] || '').trim();
    if (brand) brandCount[brand] = (brandCount[brand] || 0) + 1;
    const mgt = String(r[map['管理番号']] || '').trim();
    if (mgt && !String(r[map[C.NAME]] || '').trim()) {
      const d = asstDay_(r[map['出品日']]);
      blanks.push({ mgt: mgt, listDate: d ? asstFmt_(d) : '' });
    }
  });
  const brands = Object.keys(brandCount).sort(function(a, b) { return brandCount[b] - brandCount[a]; });

  let last = {};
  try { last = JSON.parse(PropertiesService.getUserProperties().getProperty('ASST_LAST_REGISTER') || '{}'); } catch (_) {}

  return {
    today: asstFmt_(asstToday_()),
    lists: lists,
    brands: brands.slice(0, 300),
    blanks: blanks,
    purchases: map[C.REC_ID] === undefined ? [] : asstReadPurchaseRecords_(),
    hasRecId: map[C.REC_ID] !== undefined,
    hasUrl: map[C.URL] !== undefined,
    last: last,
    ocrReady: typeof Drive !== 'undefined',
  };
}

/** 商品管理の2行目のプルダウン（リスト形式）の選択肢を読む。読めなければ null */
function asstListFromValidation_(sheet, col) {
  try {
    const rule = sheet.getRange(2, col).getDataValidation();
    if (!rule) return null;
    const type = rule.getCriteriaType();
    const v = rule.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return v[0].map(String);
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      return v[0].getValues().map(function(r) { return String(r[0] || '').trim(); }).filter(String);
    }
  } catch (_) {}
  return null;
}

/** ダッシュボードの仕入れ記録テーブル（I18〜）から、ID・仕入れ日・場所・点数・出品済みを読む。新しい順 */
function asstReadPurchaseRecords_() {
  const dash = asstSheet_('月次ダッシュボード');
  if (!dash) return [];
  let end = 47;
  try {
    const dp = Number(PropertiesService.getDocumentProperties().getProperty('PURCHASE_REC_ROW_END'));
    if (dp && dp > end) end = dp;
  } catch (_) {}
  const rows = dash.getRange(18, 9, end - 18 + 1, 10).getValues();  // I〜R
  const out = [];
  rows.forEach(function(r) {
    const id = String(r[0] || '').trim();
    if (!id) return;
    const d = asstDay_(r[1]);
    out.push({ id: id, date: d ? asstFmt_(d) : '', place: String(r[2] || ''), qty: r[3], done: r[9] === true });
  });
  out.sort(function(a, b) { return String(b.date).localeCompare(String(a.date)) || b.id.localeCompare(a.id); });
  return out;
}

// ===== スクショの文字読み取り =====

/**
 * 画面から送られた画像（base64）を、Google ドライブの OCR で文字にする。
 * 一時的に Google ドキュメントを作り、文字を読んだらすぐ削除する。
 * @return {{name:string, price:(number|string), text:string}}
 */
function asstOcrImage(base64, mimeType) {
  if (typeof Drive === 'undefined') {
    throw new Error('文字読み取りの準備ができていません。Apps Script の「サービス」に「Drive API」を追加してください。');
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'image/png', 'asst-ocr');
  const title = 'アシスタント読み取り用（自動で削除されます）';
  let fileId = null;
  try {
    if (Drive.Files.create) {   // Drive API v3
      fileId = Drive.Files.create({ name: title, mimeType: 'application/vnd.google-apps.document' }, blob,
        { ocrLanguage: 'ja' }).id;
    } else {                    // Drive API v2
      fileId = Drive.Files.insert({ title: title, mimeType: 'application/vnd.google-apps.document' }, blob,
        { ocr: true, ocrLanguage: 'ja' }).id;
    }
    const text = DocumentApp.openById(fileId).getBody().getText();
    const parsed = asstParseListingText_(text, asstTitleWords_());
    return { name: parsed.name, price: parsed.price, joins: parsed.joins, text: text };
  } finally {
    if (fileId) {
      try { Drive.Files.remove(fileId); } catch (_) {
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch (__) {}
      }
    }
  }
}

/**
 * これまでの商品名（商品管理のB列）に出てくる単語の一覧。
 * 折り返された商品名をつなぐとき、空白を入れるかどうかの判断に使う。
 */
function asstTitleWords_() {
  const words = {};
  try {
    const sheet = asstSheet_(ASST_ITEMS_SHEET);
    const map = asstHeaderMap_(sheet);
    if (sheet.getLastRow() < 2 || map[ASST_REG_COLS.NAME] === undefined) return words;
    sheet.getRange(2, map[ASST_REG_COLS.NAME] + 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
      String(r[0] || '').split(/[\s　]+/).forEach(function(w) { if (w) words[w] = true; });
    });
  } catch (_) {}
  return words;
}

/**
 * 折り返された2行をつなぐ。空白を入れるかどうかを推測する。
 * - 英字・数字どうし → 空白あり（英単語は途中で折り返されない）
 * - つなげた形が過去の商品名にある単語 → 空白なし（例：異素＋材 → 異素材）
 * - 前後のどちらかが過去の商品名にある単語 → 空白あり（例：ブルー＋幾何学）
 * - どちらも知らない → 単語の途中で切れたとみなして空白なし
 */
function asstJoinWrapped_(a, b, words) {
  const left = a.split(' ').pop(), right = b.split(' ')[0];
  let space;
  if (/[A-Za-z0-9)）]$/.test(a) && /^[A-Za-z0-9(（]/.test(b)) space = true;
  else if (words[left + right]) space = false;
  else space = !!(words[left] || words[right]);
  return { text: a + (space ? ' ' : '') + b, note: '『' + left + '』＋『' + right + '』→ ' + (space ? '空白あり' : '空白なし') };
}

/**
 * 読み取った文字から、商品名と出品価格を探す（あくまで下書き。画面で確認・修正する）。
 * - 価格：最初に出てくる「¥2,980」「2,980円」など（300円以上）
 * - 商品名：「商品名」の次の行。なければ価格の上の行（折り返された行はつなげる。40文字まで）
 */
function asstParseListingText_(text, words) {
  words = words || {};
  const lines = String(text || '').split(/\r?\n/).map(function(s) { return s.replace(/\s+/g, ' ').trim(); })
    .filter(function(s) { return s; });

  let price = '', priceLine = -1;
  const priceRe = /[¥￥]\s*([\d,，]{3,11})|([\d,，]{3,11})\s*円/;
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(priceRe);
    if (!m) continue;
    const n = Number((m[1] || m[2]).replace(/[,，]/g, ''));
    if (n >= 300 && n <= 9999999) { price = n; priceLine = i; break; }
  }

  // 商品名ではない行（時刻・ボタン・価格・「VINTAGE / サイズ：XL」の行・パンくずリスト など）
  const noise = /^(\d{1,2}:\d{2}|[<＜‹←×✕]|.*(税込|送料込み|送料込|いいね|コメント|購入手続き|検索|お探し|シェア|出品者|商品の説明|商品の情報|カテゴリー|商品の状態|配送料|発送|ブランド|サイズ|プロモーション|ホーム|mercari|[>＞›]).*|[\d\s%.,¥￥円]+)$/i;
  let name = '';
  const joins = [];
  const labelIdx = lines.findIndex(function(s) { return /^商品名/.test(s); });
  if (labelIdx >= 0) {
    const rest = lines[labelIdx].replace(/^商品名[\s:：]*/, '');
    name = rest || (lines[labelIdx + 1] || '');
    name = name.replace(/\s*\d+\s*\/\s*40$/, '');  // 文字数カウンター「32/40」を除く
  } else if (priceLine > 0) {
    // 価格の上にある行を、上へ向かって集める。長い商品名は折り返されて複数行になるのでつなげる。
    // メルカリの商品名は40文字までなので、それを超える行は商品名ではない（パンくずリストなど）
    const parts = [];
    for (let i = priceLine - 1; i >= 0; i--) {
      if (noise.test(lines[i])) { if (parts.length) break; continue; }
      if ((lines[i] + parts.join('')).length > 40) break;
      parts.unshift(lines[i]);
    }
    name = parts[0] || '';
    for (let k = 1; k < parts.length; k++) {
      const j = asstJoinWrapped_(name, parts[k], words);
      name = j.text;
      joins.push(j.note);
    }
  }
  return { name: name.slice(0, 80), price: price, joins: joins };
}

// ===== 登録 =====

/**
 * 登録画面の「登録」から呼ばれる。
 * @param {{target:string, name:string, genre:string, taste:string, brand:string, price:(number|string),
 *          condition:string, season:string, recId:string, url:string, listDate:string}} p
 *   target = '' なら新しい行を作る。管理番号なら、その（商品名が空の）行に書き込む。
 */
function asstRegisterItem(p) {
  const name = String(p.name || '').trim();
  const price = asstToCount_(p.price);
  if (!name) throw new Error('商品名を入力してください。');
  if (price === null || price < ASST_MERCARI_MIN_PRICE) throw new Error('出品価格は ' + ASST_MERCARI_MIN_PRICE + '円以上の整数で入力してください。');
  ['genre', 'taste', 'condition', 'season'].forEach(function(k) {
    if (!String(p[k] || '').trim()) throw new Error('ジャンル・系統・商品の状態・季節を選んでください。');
  });

  const sheet = asstSheet_(ASST_ITEMS_SHEET);
  let mgt = String(p.target || '').trim();

  if (!mgt) {
    // スクールの「新規行を追加」と同じ処理で行を作る（管理番号・出品日・出品点数・数式・チェックボックス）
    if (typeof insertNewRowForManualEntry !== 'function') {
      throw new Error('スクールの「新規行を追加」の処理が見つかりません。');
    }
    insertNewRowForManualEntry();
    const map0 = asstHeaderMap_(sheet);
    mgt = String(sheet.getRange(2, map0['管理番号'] + 1).getValue() || '').trim();
    if (!mgt) throw new Error('新しい行の管理番号を読めませんでした。商品管理シートの2行目を確認してください。');
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  let rowNum, recFilled = false;
  try {
    const map = asstHeaderMap_(sheet);
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, map['管理番号'] + 1, lastRow - 1, 1).getValues();
    const i = ids.findIndex(function(r) { return String(r[0]).trim() === mgt; });
    if (i < 0) throw new Error('商品管理シートに ' + mgt + ' が見つかりません。');
    rowNum = i + 2;
    if (p.target && String(sheet.getRange(rowNum, map[ASST_REG_COLS.NAME] + 1).getValue() || '').trim()) {
      throw new Error(mgt + ' にはすでに商品名が入っています。別の行を選んでください。');
    }

    const C = ASST_REG_COLS;
    const set = function(header, value) {
      if (map[header] === undefined) return;
      sheet.getRange(rowNum, map[header] + 1).setValue(value);
    };
    set(C.NAME, name);
    set(C.GENRE, String(p.genre).trim());
    set(C.TASTE, String(p.taste).trim());
    set(C.BRAND, String(p.brand || '').trim());
    set(C.PRICE, price);
    set(C.CONDITION, String(p.condition).trim());
    set(C.SEASON, String(p.season).trim());
    const url = String(p.url || '').trim();
    if (url) set(C.URL, url);

    // 出品日：画面で変えたときだけ書き込む
    const listDateCell = sheet.getRange(rowNum, map['出品日'] + 1);
    const wanted = p.listDate ? new Date(String(p.listDate).replace(/-/g, '/')) : null;
    const current = asstDay_(listDateCell.getValue());
    if (wanted && !isNaN(wanted.getTime()) && (!current || current.getTime() !== asstDay_(wanted).getTime())) {
      listDateCell.setValue(asstDay_(wanted));
    }

    // 仕入れ記録ID：書き込んで、スクールの自動入力（仕入れ日・仕入れ値・付帯費・仕入れ場所）を呼ぶ
    const recId = String(p.recId || '').trim();
    if (recId && map[C.REC_ID] !== undefined) {
      set(C.REC_ID, recId);
      if (typeof applyPurchaseRecAutoFill_ === 'function') {
        recFilled = applyPurchaseRecAutoFill_(sheet, rowNum, 1, recId).filled > 0;
      }
    }

    try {
      PropertiesService.getUserProperties().setProperty('ASST_LAST_REGISTER', JSON.stringify({
        genre: p.genre, taste: p.taste, condition: p.condition, season: p.season, recId: recId,
      }));
    } catch (_) {}
  } finally {
    lock.releaseLock();
  }

  // アシスタント管理に追加（7日間様子見）、スクールの在庫アラートを更新
  try { asstSyncItems_(asstReadItems_()); } catch (err) { console.warn('asstSyncItems_:', err); }
  if (typeof refreshAlerts === 'function') {
    try { refreshAlerts(); } catch (err) { console.warn('refreshAlerts:', err); }
  }

  return {
    message: '✅ ' + mgt + '「' + name.slice(0, 30) + '」を登録しました（商品管理の' + rowNum + '行目）。' +
      (p.recId ? (recFilled ? '\n仕入れ日・仕入れ値・付帯費・仕入れ場所も入力しました。'
        : '\n⚠️ 仕入れ記録ID「' + p.recId + '」がダッシュボードに見つからず、仕入れ情報は入りませんでした。') : ''),
  };
}

// ===== 登録画面（HTML）=====

function asstRegisterHtml_() {
  return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><style>' +
    'body{font-family:-apple-system,"Hiragino Sans",sans-serif;font-size:13px;color:#202124;margin:0;padding:12px 16px}' +
    'h3{font-size:13px;margin:12px 0 6px;color:#3c4043}' +
    'select,input{font:inherit;padding:6px;border:1px solid #c4c7c5;border-radius:6px;box-sizing:border-box;width:100%}' +
    '.grid{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px 12px}.wide{grid-column:span 3}.full{grid-column:span 4}' +
    'label{display:block;color:#5f6368;font-size:12px;margin-bottom:2px}' +
    '#drop{border:2px dashed #9aa0a6;border-radius:10px;padding:16px;text-align:center;color:#5f6368;cursor:pointer}' +
    '#drop.over{border-color:#1a73e8;background:#e8f0fe}' +
    '.muted{color:#5f6368}.warn{color:#b3261e}.ok{color:#188038}' +
    'button{font:inherit;padding:8px 20px;border-radius:6px;border:0;cursor:pointer}' +
    '#save{background:#1a73e8;color:#fff;font-weight:600}#save:disabled{background:#9aa0a6}#close{background:#f1f3f4}' +
    '#same{background:#f1f3f4;padding:4px 12px;font-size:12px}' +
    'details{margin-top:6px}pre{white-space:pre-wrap;background:#f8f9fa;padding:8px;border-radius:6px;max-height:140px;overflow:auto;font-size:12px}' +
    '#msg{margin-top:10px;white-space:pre-wrap}#ocrmsg{white-space:pre-wrap}' +
    '</style></head><body>' +
    '<div id="loading" class="muted">読み込み中…</div>' +
    '<div id="main" style="display:none">' +
    '<div id="drop">📷 メルカリの出品画面のスクショをここにドラッグ＆ドロップ（またはクリックして選ぶ）' +
    '<input id="file" type="file" accept="image/*" style="display:none"></div>' +
    '<div id="ocrmsg" class="muted" style="margin-top:6px"></div>' +
    '<details id="rawbox" style="display:none"><summary class="muted">読み取った文字をすべて見る</summary><pre id="raw"></pre></details>' +
    '<h3>登録内容（読み取った内容は下書きです。確認・修正してから登録してください）</h3>' +
    '<div class="grid">' +
    '<div class="wide"><label>商品名</label><input id="name" maxlength="80"></div>' +
    '<div><label>出品価格（円）</label><input id="price" type="number" min="300"></div>' +
    '<div><label>ジャンル</label><select id="genre"></select></div>' +
    '<div><label>系統</label><select id="taste"></select></div>' +
    '<div><label>商品の状態</label><select id="condition"></select></div>' +
    '<div><label>季節</label><select id="season"></select></div>' +
    '<div class="wide"><label>ブランド（なければ空欄）</label><input id="brand" list="brands"><datalist id="brands"></datalist></div>' +
    '<div><label>出品日</label><input id="listDate" type="date"></div>' +
    '<div class="wide" id="recBox"><label>仕入れ記録ID</label><select id="recId"></select></div>' +
    '<div style="display:flex;align-items:flex-end"><button id="same">前回と同じ</button></div>' +
    '<div class="full" id="urlBox"><label>メルカリ商品URL（任意）</label><input id="url" placeholder="https://jp.mercari.com/item/..."></div>' +
    '<div class="full"><label>書き込む行</label><select id="target"></select></div>' +
    '</div>' +
    '<div style="margin-top:14px;display:flex;gap:8px"><button id="save">登録する</button><button id="close">閉じる</button></div>' +
    '<div id="msg"></div></div>' +
    '<script>' + asstRegisterScript_() + '</script></body></html>';
}

function asstRegisterScript_() {
  return [
    'var DATA=null;',
    'function $(id){return document.getElementById(id);}',
    'function txt(el,s){el.textContent=s;return el;}',
    'function opt(sel,value,label){var o=document.createElement("option");o.value=value;o.textContent=label;sel.appendChild(o);}',
    'function fail(e){$("loading").style.display="none";$("main").style.display="block";$("msg").className="warn";txt($("msg"),"エラー："+(e&&e.message?e.message:e));$("save").disabled=false;}',
    'function load(doneMsg){google.script.run.withSuccessHandler(function(d){DATA=d;render();if(doneMsg){$("msg").className="ok";txt($("msg"),doneMsg);}}).withFailureHandler(fail).asstGetRegisterData();}',
    'function render(){',
    '  $("loading").style.display="none";$("main").style.display="block";',
    '  [["genre","ジャンル"],["taste","系統"],["condition","商品の状態"],["season","季節"]].forEach(function(p){var s=$(p[0]);s.innerHTML="";opt(s,"","選んでください");DATA.lists[p[1]].forEach(function(v){opt(s,v,v);});});',
    '  var dl=$("brands");dl.innerHTML="";DATA.brands.forEach(function(b){var o=document.createElement("option");o.value=b;dl.appendChild(o);});',
    '  var r=$("recId");r.innerHTML="";opt(r,"","（なし）");',
    '  DATA.purchases.forEach(function(p){opt(r,p.id,p.id+"  "+p.date+"  "+p.place+(p.qty!==""?"  "+p.qty+"点":"")+(p.done?"（出品済み）":""));});',
    '  $("recBox").style.display=DATA.hasRecId?"":"none";$("urlBox").style.display=DATA.hasUrl?"":"none";',
    '  var t=$("target");t.innerHTML="";opt(t,"","新しい行を作る（スクールの「新規行を追加」と同じ）");',
    '  DATA.blanks.forEach(function(b){opt(t,b.mgt,b.mgt+" の行に書き込む（商品名が空の行・出品日 "+b.listDate+"）");});',
    '  ["name","price","brand","url"].forEach(function(id){$(id).value="";});',
    '  $("listDate").value=DATA.today.replace(/\\//g,"-");',
    '  if(!DATA.ocrReady){$("ocrmsg").className="warn";txt($("ocrmsg"),"文字の読み取りは、Apps Script の「サービス」に「Drive API」を追加すると使えます（手入力での登録はできます）。");}',
    '  $("save").disabled=false;$("name").focus();',
    '}',
    'function same(){var l=DATA.last||{};["genre","taste","condition","season","recId"].forEach(function(k){if(l[k]!==undefined&&l[k]!==null)$(k).value=l[k];});}',
    'function readFile(file){',
    '  if(!file||!/^image\\//.test(file.type)){$("ocrmsg").className="warn";txt($("ocrmsg"),"画像ファイルを選んでください。");return;}',
    '  if(!DATA.ocrReady){$("ocrmsg").className="warn";txt($("ocrmsg"),"Drive API が追加されていないため、読み取りできません。");return;}',
    '  $("ocrmsg").className="muted";txt($("ocrmsg"),"読み取っています…（数秒かかります）");',
    '  var fr=new FileReader();fr.onload=function(){shrink(fr.result,file.type,function(dataUrl,type){',
    '    var b64=dataUrl.split(",")[1];',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if(res.name)$("name").value=res.name;if(res.price!=="")$("price").value=res.price;',
    '      txt($("raw"),res.text||"");$("rawbox").style.display="block";',
    '      $("ocrmsg").className=res.name||res.price!==""?"ok":"warn";',
    '      txt($("ocrmsg"),res.name||res.price!==""?"✅ 読み取りました。商品名と価格が正しいか確認してください。"+(res.joins&&res.joins.length?"\\n折り返し："+res.joins.join(" / ")+"（推測です。空白の有無を確認してください）":""):"商品名・価格を見つけられませんでした。「読み取った文字をすべて見る」から写すか、手で入力してください。");',
    '    }).withFailureHandler(function(e){$("ocrmsg").className="warn";txt($("ocrmsg"),"読み取りに失敗しました："+(e&&e.message?e.message:e));}).asstOcrImage(b64,type);',
    '  });};fr.readAsDataURL(file);',
    '}',
    '// 大きすぎる画像は縮小してから送る（長い辺 2400px まで）',
    'function shrink(dataUrl,type,done){var img=new Image();img.onload=function(){var max=2400,w=img.width,h=img.height,s=Math.min(1,max/Math.max(w,h));',
    '  if(s>=1){done(dataUrl,type);return;}var c=document.createElement("canvas");c.width=Math.round(w*s);c.height=Math.round(h*s);',
    '  c.getContext("2d").drawImage(img,0,0,c.width,c.height);done(c.toDataURL("image/png"),"image/png");};img.src=dataUrl;}',
    'function save(){',
    '  $("save").disabled=true;$("msg").className="muted";txt($("msg"),"登録しています…");',
    '  var p={target:$("target").value,name:$("name").value,genre:$("genre").value,taste:$("taste").value,brand:$("brand").value,',
    '    price:$("price").value===""?"":Number($("price").value),condition:$("condition").value,season:$("season").value,',
    '    recId:$("recId").value,url:$("url").value,listDate:$("listDate").value};',
    '  google.script.run.withSuccessHandler(function(res){',
    '    $("rawbox").style.display="none";txt($("ocrmsg"),"");load(res.message+"\\n続けて次の商品を登録できます。");',
    '  }).withFailureHandler(fail).asstRegisterItem(p);',
    '}',
    'var drop=$("drop");',
    'drop.addEventListener("click",function(){$("file").click();});',
    '$("file").addEventListener("change",function(e){readFile(e.target.files[0]);e.target.value="";});',
    'drop.addEventListener("dragover",function(e){e.preventDefault();drop.className="over";});',
    'drop.addEventListener("dragleave",function(){drop.className="";});',
    'drop.addEventListener("drop",function(e){e.preventDefault();drop.className="";readFile(e.dataTransfer.files[0]);});',
    '$("same").addEventListener("click",same);',
    '$("save").addEventListener("click",save);',
    '$("close").addEventListener("click",function(){google.script.host.close();});',
    'load();',
  ].join('\n');
}
