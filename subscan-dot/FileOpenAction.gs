/**
 * ファイルを開いたときのイベントハンドラ
 */
function onOpen() {
  // Uiクラスを取得する
  var ui = SpreadsheetApp.getUi();
  // Uiクラスからメニューを作成する 
  var menu = ui.createMenu("ファイル読込");
  // メニューにアイテムを追加する
  menu.addItem("CSVファイル読込み", "fileOpen");
  // メニューをUiクラスに追加する
  menu.addToUi();
  // Uiクラスからメニューを作成する
  var writeMenu = ui.createMenu("データ抽出");
  // メニューにアイテムを追加する
  writeMenu.addItem("カスタムシート書き出し", "writeData");
  // メニューをUiクラスに追加する
  writeMenu.addToUi();
}

/**
 * ファイル読込み用HTML作成
 */
function fileOpen() {
  var html = HtmlService.createHtmlOutputFromFile("FileOpen");
  SpreadsheetApp.getUi().showModalDialog(html, "CSVファイル読込み");
}

/**
 * tradingHistシートクリア処理
 */
function tradingHistSheetClear(){
  // シート名指定でシートを取得する
  var thsh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("tradingHist");
  // シートのすべてをクリアする
  thsh.clear();
}

/**
 * ファイル読込み用HTML作成
 * @param {number} data - 読込みデータ
 * 
 * FileOpen.html内のJSに指定して呼び出す
 */
function readTextGASFileOpen(data){
  // 書き出し前にシートを全てクリアする
  tradingHistSheetClear();
  var thsh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("tradingHist");
  var csv = Utilities.parseCsv(data);
  // セルA1からCSVの内容を書き込んでいく
  thsh.getRange(1,1,csv.length,csv[0].length).setValues(csv);
  var lastRow = thsh.getLastRow();
  // 項目行を除くため行数調整する(lastRow - 1 ;だとNaNになるためデクリメントすること)
  lastRow--;
  // Uiクラスを使用して処理終了メッセージダイアログ(タイトルとOKボタン）を表示
  var ui = SpreadsheetApp.getUi();
  // ダイアログタイトル、メッセージと「OK」ボタンを表示(改行するときは「\n」を追加する)
  var title = "読み込み成功"
  var message = "tradingHistシートに" + lastRow + "件読み込みました。"
  ui.alert(title, message, ui.ButtonSet.OK);
}
