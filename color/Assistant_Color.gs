/**
 * 古着物販アシスタント — 観測日・値下げ日の色分け（商品管理シートの A〜B列）
 *
 *   黄色   ：出品から 7・14・21・28日目…（7日ごと）＝ 定点観測の日
 *   赤     ：出品から 11・15・19・23日目…（4日ごと）＝ 値下げの日
 *            メモに「値下げ停止」と書いた商品は赤くしない
 *   オレンジ：黄色と赤が同じ日（35日目・63日目…）
 *
 * 色は「在庫日数」の列から毎日自動で決まる（このスクリプトは毎日は動かない）。
 * 売却済み・売却完了・再出品で置き換わった行・無料交換に出した行には色を付けない。
 *
 * 使い方：Apps Script のエディタで asstSetupColors を選んで「実行」（最初に1回だけ）。
 *         色分けをやめたいときは asstRemoveColors を実行。
 * スクールの色ルールは消さず、中身も順番も変えない（このルール3つを先頭に置くだけ）。
 */

const ASST_COLOR_SHEET = '商品管理';
const ASST_COLOR_STOP_WORD = '値下げ停止';
const ASST_COLOR_MARK = 'MOD(';   // このスクリプトのルールの目印（スクールのルールには MOD が無い）

function asstSetupColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASST_COLOR_SHEET);
  if (!sheet) throw new Error('「' + ASST_COLOR_SHEET + '」シートが見つかりません。');

  // 列は見出しの名前で探す
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  const col = function(name) {
    const i = headers.indexOf(name);
    if (i < 0) throw new Error('商品管理シートに「' + name + '」の列が見つかりません。');
    return asstColLetter_(i + 1);
  };
  const A = col('管理番号'), Z = col('在庫日数'), AL = col('メモ');
  const AI = col('売却完了'), AP = col('再出品済み'), AR = col('無料交換に出した');

  const inStock = '$' + A + '2<>"",ISNUMBER($' + Z + '2),$' + AI + '2<>TRUE,$' + AP + '2<>TRUE,$' + AR + '2<>TRUE';
  const observe = '$' + Z + '2>=7,MOD($' + Z + '2,7)=0';
  const drop = '$' + Z + '2>=11,MOD($' + Z + '2-7,4)=0,NOT(REGEXMATCH(TO_TEXT($' + AL + '2),"' + ASST_COLOR_STOP_WORD + '"))';

  const range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, 2);  // A〜B列
  const rule = function(formula, bg, font) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(' + formula + ')')
      .setBackground(bg).setFontColor(font).setBold(true)
      .setRanges([range]).build();
  };
  const ours = [
    rule(inStock + ',' + observe + ',' + drop, '#ff9800', '#202124'),  // 両方：オレンジ
    rule(inStock + ',' + drop, '#c62828', '#ffffff'),                  // 値下げ：赤
    rule(inStock + ',' + observe, '#ffd600', '#202124'),               // 観測：黄色
  ];

  const others = sheet.getConditionalFormatRules().filter(function(r) { return !asstIsOurColorRule_(r); });
  sheet.setConditionalFormatRules(ours.concat(others));
  Logger.log('色分けを設定しました（スクールのルール ' + others.length + '個はそのまま）。');
}

/** 色分けをやめる（このスクリプトのルールだけ消す） */
function asstRemoveColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASST_COLOR_SHEET);
  const rules = sheet.getConditionalFormatRules();
  const others = rules.filter(function(r) { return !asstIsOurColorRule_(r); });
  sheet.setConditionalFormatRules(others);
  Logger.log('色分けを消しました（' + (rules.length - others.length) + '個）。');
}

/** このスクリプト（または前のバージョンのアシスタント）が入れたルールか */
function asstIsOurColorRule_(r) {
  const c = r.getBooleanCondition();
  if (!c) return false;
  return c.getCriteriaValues().some(function(v) {
    const s = String(v);
    return s.indexOf(ASST_COLOR_MARK) >= 0 || s.indexOf('アシスタント色分け用') >= 0;
  });
}

function asstColLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
