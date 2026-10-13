/**
 * 古着物販アシスタント Ver.1 — 段階3：値下げ管理
 *
 * - 値下げ画面：今日が値下げ日の商品を順に表示。新価格は自分で入力する。
 * - 値下げ完了：アシスタント管理（現在価格・最終値下げ日・次回値下げ予定日）を更新し、
 *   商品管理の「値下げ回数」を +1、「メモ」の最後に改行して `10/12,2980→2480` を追記、値下げ履歴に記録する。
 * - 今回は見送る：次回値下げ予定日を4日後にする（定期値下げは続ける）。
 * - 定期値下げ停止：値下げ通知の対象から外す（7日ごとの観測は続ける）。
 * 商品管理シートに書き込むのは「値下げ回数」「メモ」の2列だけ。
 */

const ASST_MERCARI_MIN_PRICE = 300;  // メルカリの最低価格

// ===== メニューから開く =====

function asstOpenMarkdown() {
  // 今日の物販タスクや商品管理で行を選んでいれば、その商品を選んだ状態で開く
  asstShowMarkdownDialog_(typeof asstActiveMgt_ === 'function' ? asstActiveMgt_() : '');
}

/** 観測画面の「このまま値下げする」から呼ばれる（その商品を選んだ状態で開く） */
function asstOpenMarkdownFor(mgt) {
  asstShowMarkdownDialog_(String(mgt || ''));
}

function asstShowMarkdownDialog_(mgt) {
  const ui = SpreadsheetApp.getUi();
  if (!asstSheet_(ASST_SHEETS.MANAGE)) {
    ui.alert('先にメニュー「初期設定」を実行してください。');
    return;
  }
  try {
    asstSyncItems_(asstReadItems_());
  } catch (err) {
    ui.alert('❌ 商品リストを読めませんでした', String(err.message || err), ui.ButtonSet.OK);
    return;
  }
  const html = HtmlService.createHtmlOutput(asstMarkdownHtml_(mgt)).setWidth(720).setHeight(620);
  ui.showModalDialog(html, '💰 値下げ');
}

// ===== 値下げ画面に渡すデータ =====

function asstGetMarkdownData() {
  const settings = asstGetSettings_();
  const today = asstToday_();
  const manage = asstReadManage_();
  const H = manage.H;
  const itemsByMgt = {};
  asstReadItems_().forEach(function(it) { if (!(it.mgt in itemsByMgt)) itemsByMgt[it.mgt] = it; });
  const logs = asstReadLogsByMgt_();
  const history = asstReadPriceHistoryByMgt_();
  const urls = asstReadItemUrls_();

  const list = [];
  manage.rows.forEach(function(r) {
    if (r[H['ステータス']] !== ASST_STATUS.MARKDOWN) return;
    const mgt = String(r[H['管理番号']]).trim();
    const listDate = asstDay_(r[H['出品日']]);
    if (!mgt || !listDate) return;
    const item = itemsByMgt[mgt] || {};
    const nextDrop = asstDay_(r[H['次回値下げ予定日']]);
    const lastDrop = asstDay_(r[H['最終値下げ日']]);
    const obs = logs[mgt] || [];
    const last = obs.length ? obs[obs.length - 1] : null;
    list.push({
      mgt: mgt,
      name: String(r[H['商品名']] || ''),
      listDate: asstFmt_(listDate),
      elapsed: asstDaysBetween_(listDate, today),
      price: asstIsNumber_(r[H['現在価格']]) ? Number(r[H['現在価格']]) : null,
      listPrice: asstIsNumber_(item.listPrice) ? Number(item.listPrice) : null,
      drops: item.drops || 0,
      lastDrop: lastDrop ? asstFmt_(lastDrop) : '',
      nextDrop: nextDrop ? asstFmt_(nextDrop) : '',
      due: !!nextDrop && nextDrop <= today,
      url: urls[mgt] || '',
      lastObs: last ? (asstFmt_(last.date) + '（' + last.elapsed + '日後）  閲覧 ' + last.views +
        ' ・ 検索 ' + last.search + ' ・ いいね ' + (last.likes === '' ? '-' : last.likes)) : '',
      history: (history[mgt] || []).slice(-3).map(function(h) {
        return asstFmt_(h.date).slice(5) + ' ' + h.op + (h.op === '値下げ' ? ' ' + h.from + '→' + h.to : '');
      }),
    });
  });

  list.sort(function(a, b) {
    if (a.due !== b.due) return a.due ? -1 : 1;
    return String(a.nextDrop).localeCompare(String(b.nextDrop));
  });

  return {
    today: asstFmt_(today),
    dropDays: settings[ASST_SETTING.DROP_DAYS],
    minPrice: ASST_MERCARI_MIN_PRICE,
    items: list,
  };
}

// ===== 値下げ完了・見送り・停止 =====

/**
 * 値下げ画面のボタンから呼ばれる。
 * @param {{mgt:string, action:string, newPrice:(number|string)}} p  action = 'done' | 'skip' | 'stop'
 */
function asstApplyMarkdown(p) {
  if (['done', 'skip', 'stop'].indexOf(p.action) < 0) throw new Error('操作が正しくありません。');

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  let result;
  try {
    const settings = asstGetSettings_();
    const dropDays = settings[ASST_SETTING.DROP_DAYS];
    const today = asstToday_();
    const manage = asstReadManage_();
    const H = manage.H;
    const idx = manage.rows.findIndex(function(r) { return String(r[H['管理番号']]).trim() === p.mgt; });
    if (idx < 0) throw new Error('アシスタント管理シートに ' + p.mgt + ' が見つかりません。');
    const row = manage.rows[idx];
    if (row[H['ステータス']] !== ASST_STATUS.MARKDOWN) {
      throw new Error(p.mgt + ' は定期値下げ中ではありません（ステータス：' + row[H['ステータス']] + '）。');
    }
    const listDate = asstDay_(row[H['出品日']]);
    const oldPrice = asstIsNumber_(row[H['現在価格']]) ? Number(row[H['現在価格']]) : null;

    if (p.action === 'done') {
      const newPrice = asstToCount_(p.newPrice);
      if (newPrice === null || newPrice < ASST_MERCARI_MIN_PRICE) {
        throw new Error('新価格は ' + ASST_MERCARI_MIN_PRICE + '円以上の整数で入力してください。');
      }
      if (oldPrice !== null && newPrice >= oldPrice) {
        throw new Error('新価格が現在価格（' + oldPrice.toLocaleString() + '円）以上です。値下げ後の価格を入力してください。');
      }

      // 商品管理：値下げ回数 +1、メモに追記（この2列だけ）
      const written = asstWriteDropToItems_(p.mgt, today, oldPrice, newPrice);

      row[H['現在価格']] = newPrice;
      row[H['最終値下げ日']] = today;
      row[H['次回値下げ予定日']] = asstAddDays_(today, dropDays);
      asstAppendPriceLog_(today, p.mgt, row[H['商品名']], '値下げ', oldPrice, newPrice,
        written.drops, listDate ? asstDaysBetween_(listDate, today) : '');
      result = {
        message: '✅ ' + p.mgt + ' を ' + (oldPrice === null ? '' : oldPrice.toLocaleString() + '円 → ') +
          newPrice.toLocaleString() + '円 にしました（' + written.drops + '回目）。次回の値下げ日：' +
          asstFmt_(row[H['次回値下げ予定日']]),
      };
    } else if (p.action === 'skip') {
      row[H['次回値下げ予定日']] = asstAddDays_(today, dropDays);
      asstAppendPriceLog_(today, p.mgt, row[H['商品名']], '見送り', oldPrice, '', '',
        listDate ? asstDaysBetween_(listDate, today) : '');
      result = { message: '⏭ ' + p.mgt + ' は今回見送りました。次回の値下げ日：' + asstFmt_(row[H['次回値下げ予定日']]) };
    } else {
      row[H['ステータス']] = ASST_STATUS.STOPPED;
      row[H['定期値下げ']] = '停止';
      row[H['次回値下げ予定日']] = '';
      asstAppendPriceLog_(today, p.mgt, row[H['商品名']], '停止', oldPrice, '', '',
        listDate ? asstDaysBetween_(listDate, today) : '');
      result = { message: '⏹ ' + p.mgt + ' の定期値下げを止めました（7日ごとの観測は続きます）。' };
    }

    row[H['更新日時']] = new Date();
    manage.sheet.getRange(idx + 2, 1, 1, row.length).setValues([row]);
    asstFormatManage_(manage.sheet);
  } finally {
    lock.releaseLock();
  }

  // スクリプトの書き込みではスクールの onEdit が動かないので、売却分析を自分で更新する
  if (p.action === 'done' && typeof refreshSalesAnalysis === 'function') {
    try { refreshSalesAnalysis(); } catch (err) { console.warn('refreshSalesAnalysis:', err); }
  }
  if (typeof asstRebuildTasks_ === 'function') {
    try { asstRebuildTasks_(); } catch (err) { console.warn('asstRebuildTasks_:', err); }
  }
  return result;
}

/**
 * 商品管理の「値下げ回数」を +1 し、「メモ」の最後に改行して `10/12,2980→2480` を追記する。
 * ほかの列には触らない。
 */
function asstWriteDropToItems_(mgt, today, oldPrice, newPrice) {
  const sheet = asstSheet_(ASST_ITEMS_SHEET);
  const map = asstHeaderMap_(sheet);
  const C = ASST_ITEM_COLS;
  const lastRow = sheet.getLastRow();
  const ids = lastRow >= 2 ? sheet.getRange(2, map[C.MGT] + 1, lastRow - 1, 1).getValues() : [];
  const i = ids.findIndex(function(r) { return String(r[0]).trim() === mgt; });
  if (i < 0) throw new Error('商品管理シートに ' + mgt + ' が見つかりません。');
  const rowNum = i + 2;

  const dropsCell = sheet.getRange(rowNum, map[C.DROPS] + 1);
  const drops = (Number(dropsCell.getValue()) || 0) + 1;
  dropsCell.setValue(drops);

  const memoCell = sheet.getRange(rowNum, map[C.MEMO] + 1);
  const memo = String(memoCell.getValue() || '');
  const line = (today.getMonth() + 1) + '/' + today.getDate() + ',' +
    (oldPrice === null ? '' : oldPrice) + '→' + newPrice;
  memoCell.setValue(memo ? memo + '\n' + line : line);

  return { drops: drops };
}

function asstAppendPriceLog_(date, mgt, name, op, oldPrice, newPrice, dropsTotal, elapsed) {
  const sheet = asstSheet_(ASST_SHEETS.PRICE_LOG);
  const P = asstIndex_(ASST_PRICE_HEADERS);
  const row = new Array(ASST_PRICE_HEADERS.length).fill('');
  row[P['日付']] = date;
  row[P['管理番号']] = mgt;
  row[P['商品名']] = name;
  row[P['操作']] = op;
  row[P['旧価格']] = oldPrice === null ? '' : oldPrice;
  row[P['新価格']] = newPrice;
  row[P['値下げ額']] = (op === '値下げ' && oldPrice !== null) ? oldPrice - newPrice : '';
  row[P['値下げ回数（累計）']] = dropsTotal;
  row[P['出品からの日数']] = elapsed;
  row[P['記録元']] = '値下げ画面';
  const at = sheet.getLastRow() + 1;
  sheet.getRange(at, 1, 1, row.length).setValues([row]);
  sheet.getRange(at, P['日付'] + 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(at, P['旧価格'] + 1, 1, 3).setNumberFormat('#,##0');
}

/** 値下げ履歴を 管理番号 → [{date, op, from, to}]（古い順）にまとめる */
function asstReadPriceHistoryByMgt_() {
  const out = {};
  const sheet = asstSheet_(ASST_SHEETS.PRICE_LOG);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const P = asstIndex_(ASST_PRICE_HEADERS);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, ASST_PRICE_HEADERS.length).getValues().forEach(function(r) {
    const mgt = String(r[P['管理番号']] || '').trim();
    const date = asstDay_(r[P['日付']]);
    if (!mgt || !date) return;
    (out[mgt] = out[mgt] || []).push({ date: date, op: r[P['操作']], from: r[P['旧価格']], to: r[P['新価格']] });
  });
  return out;
}

// ===== 値下げ画面（HTML）=====

function asstMarkdownHtml_(selectMgt) {
  return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><style>' +
    'body{font-family:-apple-system,"Hiragino Sans",sans-serif;font-size:13px;color:#202124;margin:0;padding:12px 16px}' +
    'h3{font-size:13px;margin:14px 0 6px;color:#3c4043}' +
    'select,input{font:inherit;padding:6px;border:1px solid #c4c7c5;border-radius:6px;box-sizing:border-box}' +
    '#item{width:100%}.box{background:#f8f9fa;border-radius:8px;padding:8px 12px;line-height:1.7}' +
    '.muted{color:#5f6368}.warn{color:#b3261e}.tag{display:inline-block;padding:0 8px;border-radius:10px;background:#fce8e6;color:#b3261e;font-size:12px}' +
    '.price{display:flex;align-items:center;gap:12px;font-size:15px}' +
    '#newPrice{width:140px;font-size:18px;text-align:right}' +
    '.quick button{padding:4px 10px;margin-right:4px;background:#f1f3f4;border-radius:14px}' +
    'button{font:inherit;padding:8px 18px;border-radius:6px;border:0;cursor:pointer}' +
    '#done{background:#d93025;color:#fff;font-weight:600}#skip{background:#f1f3f4}#stop{background:#f1f3f4}#close{background:#fff;border:1px solid #dadce0}' +
    'button:disabled{opacity:.5;cursor:default}#msg{margin-top:10px;white-space:pre-wrap}' +
    '</style></head><body>' +
    '<div id="loading" class="muted">読み込み中…</div>' +
    '<div id="main" style="display:none">' +
    '<select id="item"></select>' +
    '<div class="box" id="info" style="margin-top:8px"></div>' +
    '<h3>新しい価格（メルカリで価格を変えてから「値下げ完了」を押してください）</h3>' +
    '<div class="price"><span id="cur"></span><span>→</span><input id="newPrice" type="number" min="300" step="10"><span>円</span><span id="diff" class="muted"></span></div>' +
    '<div class="quick" style="margin-top:8px"><span class="muted">金額の入力補助：</span>' +
    '<button data-d="100">−100</button><button data-d="200">−200</button><button data-d="300">−300</button><button data-d="500">−500</button></div>' +
    '<div style="margin-top:16px;display:flex;gap:8px;flex-wrap:wrap">' +
    '<button id="done">値下げ完了</button><button id="skip">今回は見送る</button><button id="stop">定期値下げ停止</button><button id="close">閉じる</button></div>' +
    '<div id="msg"></div></div>' +
    '<script>var SELECT=' + JSON.stringify(String(selectMgt || '')).replace(/</g, '\\u003c') + ';\n' + asstMarkdownScript_() + '</script></body></html>';
}

function asstMarkdownScript_() {
  return [
    'var DATA=null,CUR=null;',
    'function $(id){return document.getElementById(id);}',
    'function txt(el,s){el.textContent=s;return el;}',
    'function yen(v){return v===null||v===""?"-":Number(v).toLocaleString()+"円";}',
    'function load(selectMgt,doneMsg){google.script.run.withSuccessHandler(function(d){DATA=d;render(selectMgt);if(doneMsg){$("msg").className="";txt($("msg"),doneMsg);}}).withFailureHandler(fail).asstGetMarkdownData();}',
    'function fail(e){$("loading").style.display="none";$("main").style.display="block";$("msg").className="warn";txt($("msg"),"エラー："+(e&&e.message?e.message:e));buttons(true);}',
    'function buttons(on){["done","skip","stop"].forEach(function(id){$(id).disabled=!on;});}',
    'function render(selectMgt){',
    '  $("loading").style.display="none";$("main").style.display="block";',
    '  var sel=$("item");sel.innerHTML="";',
    '  var g1=document.createElement("optgroup");g1.label="今日・期限切れ（"+DATA.items.filter(function(i){return i.due;}).length+"件）";',
    '  var g2=document.createElement("optgroup");g2.label="そのほかの定期値下げ中の商品";',
    '  DATA.items.forEach(function(it,i){var o=document.createElement("option");o.value=i;',
    '    o.textContent=it.mgt+"  "+it.name.slice(0,26)+"  （"+yen(it.price)+"・予定 "+(it.nextDrop?it.nextDrop.slice(5):"-")+"）";',
    '    (it.due?g1:g2).appendChild(o);});',
    '  if(g1.children.length)sel.appendChild(g1);if(g2.children.length)sel.appendChild(g2);',
    '  if(!DATA.items.length){txt($("info"),"定期値下げ中の商品がありません。");buttons(false);return;}',
    '  var idx=0;DATA.items.forEach(function(it,i){if(it.mgt===selectMgt)idx=i;});sel.value=idx;pick();',
    '}',
    'function pick(){',
    '  CUR=DATA.items[$("item").value];var it=CUR;var info=$("info");info.innerHTML="";',
    '  function line(label,val){var d=document.createElement("div");var b=document.createElement("span");b.className="muted";txt(b,label+"：");d.appendChild(b);d.appendChild(document.createTextNode(val));info.appendChild(d);return d;}',
    '  var t=line("商品",it.mgt+"  "+it.name);',
    '  if(it.due){var tag=document.createElement("span");tag.className="tag";txt(tag,it.nextDrop<DATA.today?"期限切れ":"今日");t.appendChild(document.createTextNode(" "));t.appendChild(tag);}',
    '  line("出品日",it.listDate+"（今日で"+it.elapsed+"日目）  出品価格："+yen(it.listPrice));',
    '  line("値下げ",(it.drops)+"回  最終値下げ日："+(it.lastDrop||"-")+"  次回予定："+(it.nextDrop||"-"));',
    '  if(it.history.length)line("最近の履歴",it.history.join(" / "));',
    '  if(it.lastObs)line("前回の観測",it.lastObs);',
    '  if(it.url){var d=line("メルカリ","");var a=document.createElement("a");a.href=it.url;a.target="_blank";txt(a,"商品ページを開く");d.appendChild(a);}',
    '  $("stop").removeAttribute("data-sure");txt($("stop"),"定期値下げ停止");',
    '  txt($("cur"),"現在 "+yen(it.price));$("newPrice").value="";txt($("diff"),"");txt($("msg"),"");$("msg").className="";buttons(true);$("newPrice").focus();',
    '}',
    'function diff(){var v=$("newPrice").value;if(v===""||CUR.price===null){txt($("diff"),"");return;}var d=CUR.price-Number(v);txt($("diff"),d>0?"（−"+d.toLocaleString()+"円）":"");}',
    'function send(action){',
    '  var v=$("newPrice").value;',
    '  if(action==="done"){if(v===""){$("msg").className="warn";txt($("msg"),"新しい価格を入力してください。");return;}}',
    '  if(action==="stop"&&$("stop").getAttribute("data-sure")!=="1"){$("stop").setAttribute("data-sure","1");txt($("stop"),"本当に停止する");$("msg").className="warn";txt($("msg"),"もう一度押すと定期値下げを止めます（7日ごとの観測は続きます）。");return;}',
    '  buttons(false);$("msg").className="muted";txt($("msg"),"記録しています…");var mgt=CUR.mgt;',
    '  google.script.run.withSuccessHandler(function(res){',
    '    var next=null;DATA.items.forEach(function(it){if(!next&&it.due&&it.mgt!==mgt)next=it.mgt;});',
    '    load(next||"",res.message);',
    '  }).withFailureHandler(fail).asstApplyMarkdown({mgt:mgt,action:action,newPrice:v===""?"":Number(v)});',
    '}',
    '$("item").addEventListener("change",pick);',
    '$("newPrice").addEventListener("input",diff);',
    'Array.prototype.forEach.call(document.querySelectorAll(".quick button"),function(b){b.addEventListener("click",function(){if(CUR.price===null)return;$("newPrice").value=CUR.price-Number(b.getAttribute("data-d"));diff();});});',
    '$("done").addEventListener("click",function(){send("done");});',
    '$("skip").addEventListener("click",function(){send("skip");});',
    '$("stop").addEventListener("click",function(){send("stop");});',
    '$("close").addEventListener("click",function(){google.script.host.close();});',
    'load(SELECT);',
  ].join('\n');
}
