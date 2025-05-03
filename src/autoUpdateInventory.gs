// 在庫自動更新: 販売時の在庫数自動更新
function autoUpdateInventory(productId) {
  if (!productId) throw new Error('商品IDは必須です');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const salesSheet = ss.getSheetByName('販売管理');
  const inventorySheet = ss.getSheetByName('在庫管理');
  if (!salesSheet || !inventorySheet) throw new Error('必要なシートが存在しません');

  // 販売管理シートで該当商品IDの売上件数をカウント
  const salesData = salesSheet.getDataRange().getValues();
  let salesCount = 0;
  for (let i = 1; i < salesData.length; i++) {
    if (salesData[i][1] === productId) {
      salesCount++;
    }
  }

  // 在庫管理シートで該当商品IDの在庫数を更新
  const inventoryData = inventorySheet.getDataRange().getValues();
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === productId) {
      let originalStock = Number(inventoryData[i][1]);
      let newStock = Math.max(0, originalStock - salesCount);
      inventorySheet.getRange(i+1, 2).setValue(newStock);
      // 在庫数が0ならステータスも自動変更
      if (newStock === 0) {
        inventorySheet.getRange(i+1, 3).setValue('在庫切れ');
      }
      // 最終更新日も更新
      inventorySheet.getRange(i+1, 4).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
      break;
    }
  }
}

// テストデータ作成用関数（GAS上で動作確認用）
function createTestData() {
  const properties = PropertiesService.getScriptProperties();
  let ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  let ss;
  if (!ssId) {
    ss = SpreadsheetApp.create('メルカリ販売管理システム-テスト');
    ssId = ss.getId();
    properties.setProperty('MASTER_SPREADSHEET_ID', ssId);
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }

  // 商品マスタ
  let productSheet = ss.getSheetByName('商品マスタ');
  if (!productSheet) productSheet = ss.insertSheet('商品マスタ');
  productSheet.clear();
  productSheet.appendRow(['商品ID', '商品名', 'カテゴリ', '仕入れ価格', '販売予定価格', '状態', '備考']);
  productSheet.appendRow(['P001', 'テスト商品A', '家電', 1000, 2000, '新品', '']);
  productSheet.appendRow(['P002', 'テスト商品B', '本', 500, 1200, '中古', '']);

  // 在庫管理
  let inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) inventorySheet = ss.insertSheet('在庫管理');
  inventorySheet.clear();
  inventorySheet.appendRow(['商品ID', '在庫数', 'ステータス', '最終更新日']);
  inventorySheet.appendRow(['P001', 5, '出品中', '']);
  inventorySheet.appendRow(['P002', 2, '出品中', '']);

  // 販売管理
  let salesSheet = ss.getSheetByName('販売管理');
  if (!salesSheet) salesSheet = ss.insertSheet('販売管理');
  salesSheet.clear();
  salesSheet.appendRow(['取引ID', '商品ID', '販売日', '販売価格', '販売手数料', '送料', '購入者情報', '取引ステータス']);
  salesSheet.appendRow(['T001', 'P001', '2024/06/01', 2000, 200, 100, 'テスト太郎', '売約済み']);
  salesSheet.appendRow(['T002', 'P002', '2024/06/02', 1200, 100, 80, 'テスト花子', '売約済み']);
}

// Webアプリ用エントリポイント
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}

// 在庫一覧データ取得（商品マスタJOIN）
function getInventoryList() {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const inventorySheet = ss.getSheetByName('在庫管理');
  const productSheet = ss.getSheetByName('商品マスタ');
  if (!inventorySheet || !productSheet) throw new Error('必要なシートが存在しません');

  const inventoryData = inventorySheet.getDataRange().getValues();
  const productData = productSheet.getDataRange().getValues();
  const productMap = {};
  for (let i = 1; i < productData.length; i++) {
    productMap[productData[i][0]] = {
      商品名: productData[i][1],
      カテゴリ: productData[i][2]
    };
  }

  const result = [];
  for (let i = 1; i < inventoryData.length; i++) {
    const row = inventoryData[i];
    const productId = row[0];
    result.push({
      商品ID: productId,
      商品名: productMap[productId] ? productMap[productId].商品名 : '',
      カテゴリ: productMap[productId] ? productMap[productId].カテゴリ : '',
      在庫数: row[1],
      ステータス: row[2],
      最終更新日: row[3]
    });
  }
  return result;
} 