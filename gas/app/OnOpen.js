/**
 * スプレッドシートを開いたときに自動で呼ばれる
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('カード管理ツール') // メニュー名
    .addItem('Webアプリを開く', 'openWebApp')
    .addItem('データ確認ログ', 'logDatabase')
    .addItem('最新動画取得', 'newUpdate')
    .addItem('カードデータ取得', 'fetchMtgCardDataJa')
    .addToUi();
}

/**
 * WebアプリのURLをダイアログで表示
 */
function openWebApp() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(`WebアプリURL:\n${url}`);
}