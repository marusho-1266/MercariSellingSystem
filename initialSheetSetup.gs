// スプレッドシート仕様書（spreadsheet.md）に基づく全シート作成関数
function createAllSheetsFromSpec() {
  const properties = PropertiesService.getScriptProperties();
  let ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  let ss;
  if (!ssId) {
    ss = SpreadsheetApp.create('メルカリ販売管理システム-仕様準拠');
    ssId = ss.getId();
    properties.setProperty('MASTER_SPREADSHEET_ID', ssId);
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }

  // 商品マスタシート
  let productSheet = ss.getSheetByName('商品マスタ');
  if (!productSheet) productSheet = ss.insertSheet('商品マスタ');
  productSheet.clear();
  productSheet.appendRow(['商品ID', '商品名', 'カテゴリ', '仕入れ価格', '販売予定価格', '状態', '備考']);

  // 在庫管理シート
  let inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) inventorySheet = ss.insertSheet('在庫管理');
  inventorySheet.clear();
  inventorySheet.appendRow(['商品ID', '在庫数', 'ステータス', '最終更新日']);

  // 出品管理シート
  let listingSheet = ss.getSheetByName('出品管理');
  if (!listingSheet) listingSheet = ss.insertSheet('出品管理');
  listingSheet.clear();
  listingSheet.appendRow(['出品ID', '商品ID', '出品日', '出品価格', '出品数', 'ステータス', '備考']);

  // 販売管理シート
  let salesSheet = ss.getSheetByName('販売管理');
  if (!salesSheet) salesSheet = ss.insertSheet('販売管理');
  salesSheet.clear();
  salesSheet.appendRow(['取引ID', '出品ID', '商品ID', '販売日', '販売価格', '販売手数料', '送料', '購入者情報', '取引ステータス']);

  // 財務管理シート
  let financeSheet = ss.getSheetByName('財務管理');
  if (!financeSheet) financeSheet = ss.insertSheet('財務管理');
  financeSheet.clear();
  financeSheet.appendRow(['日付', '取引ID', '収入', '支出', '利益', '備考']);
} 