/**
 * スプレッドシートを開いたときに自動で呼ばれる
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('カード管理ツール')
    .addItem('最新動画取得', 'newUpdate')
    .addItem('カードデータ取得', 'fetchMtgCardDataJa')
    .addToUi();
}
