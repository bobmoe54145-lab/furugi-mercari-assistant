/**
 * 古着物販アシスタント Ver.1 — メニュー
 *
 * スプレッドシートを開いたとき（asstOnOpen）に「物販アシスタント」メニューを作る。
 * 段階が進むごとに、ここに項目を足していく。
 */
function asstBuildMenu_() {
  SpreadsheetApp.getUi().createMenu('物販アシスタント')
    .addItem('📋 今日の物販タスクを更新', 'asstRefreshTasks')
    .addItem('📝 商品を登録', 'asstOpenRegister')
    .addItem('📊 定点観測を記録', 'asstOpenObservation')
    .addItem('💰 値下げ', 'asstOpenMarkdown')
    .addItem('📈 観測まとめを更新', 'asstRefreshSummary')
    .addSeparator()
    .addItem('🔄 商品リストを同期', 'asstSyncItems')
    .addItem('⚙️ 設定シートを開く', 'asstOpenSettings')
    .addItem('📋 アシスタント管理シートを開く', 'asstOpenManage')
    .addSeparator()
    .addItem('⏰ 毎朝の自動更新を設定', 'asstInstallDaily')
    .addItem('🔧 初期設定（最初に1回）', 'asstInitialize')
    .addToUi();
}
