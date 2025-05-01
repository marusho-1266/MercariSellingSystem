// メルカリ販売管理システム メインコード

/**
 * 初期セットアップを行う関数
 */
function initialSetup() {
  // スプレッドシートのセットアップ
  createMasterSpreadsheet();
}

/**
 * マスタースプレッドシートを作成する関数
 */
function createMasterSpreadsheet() {
  // 既存のスプレッドシートを探す
  const properties = PropertiesService.getScriptProperties();
  let ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  
  if (!ssId) {
    // 新規スプレッドシート作成
    const ss = SpreadsheetApp.create('メルカリ販売管理システム');
    ssId = ss.getId();
    
    // スプレッドシートIDを保存
    properties.setProperty('MASTER_SPREADSHEET_ID', ssId);
    
    // 必要なシートを作成
    createRequiredSheets(ss);
  }
  
  return ssId;
}

/**
 * 必要なシートを作成する関数
 */
function createRequiredSheets(ss) {
  // 商品マスタ
  createSheet(ss, '商品マスタ', [
    ['商品ID', '商品名', 'カテゴリ', '仕入れ価格', '販売予定価格', '状態', '備考']
  ]);
  
  // 在庫管理
  createSheet(ss, '在庫管理', [
    ['商品ID', '在庫数', 'ステータス', '最終更新日']
  ]);
  
  // 販売管理
  createSheet(ss, '販売管理', [
    ['取引ID', '商品ID', '販売日', '販売価格', '販売手数料', '送料', '購入者情報', '取引ステータス']
  ]);
  
  // 財務管理
  createSheet(ss, '財務管理', [
    ['日付', '取引ID', '収入', '支出', '利益', '備考']
  ]);
}

/**
 * シートを作成するヘルパー関数
 */
function createSheet(ss, sheetName, headers) {
  const sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers[0].length);
}

/**
 * WebアプリケーションのGETリクエストを処理する関数
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('メルカリ販売管理システム')
      .setFaviconUrl('https://www.google.com/images/favicon.ico');
} 