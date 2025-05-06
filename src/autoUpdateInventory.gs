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
    const pid = String(productData[i][0]).trim();
    if (!pid) continue;
    productMap[pid] = {
      商品名: String(productData[i][1]).trim(),
      カテゴリ: String(productData[i][2]).trim()
    };
  }

  const result = [];
  for (let i = 1; i < inventoryData.length; i++) {
    const row = inventoryData[i];
    const productId = String(row[0]).trim();
    if (!productId) continue;
    result.push({
      商品ID: productId,
      商品名: productMap[productId] ? productMap[productId].商品名 : '',
      カテゴリ: productMap[productId] ? productMap[productId].カテゴリ : '',
      在庫数: String(row[1]).trim(),
      ステータス: String(row[2]).trim(),
      最終更新日: String(row[3]).trim()
    });
  }
  Logger.log(result);
  return result;
}

// 販売一覧データ取得（商品マスタJOIN）
function getSalesList() {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const salesSheet = ss.getSheetByName('販売管理');
  const productSheet = ss.getSheetByName('商品マスタ');
  if (!salesSheet || !productSheet) throw new Error('必要なシートが存在しません');

  const salesData = salesSheet.getDataRange().getValues();
  if (salesData.length < 2) return [];
  const headers = salesData[0];
  const idx = (name) => headers.indexOf(name);

  const productData = productSheet.getDataRange().getValues();
  const productHeaders = productData[0];
  const productIdIdx = productHeaders.indexOf('商品ID');
  const productNameIdx = productHeaders.indexOf('商品名');
  const categoryIdx = productHeaders.indexOf('カテゴリ');
  const productMap = {};
  for (let i = 1; i < productData.length; i++) {
    productMap[productData[i][productIdIdx]] = {
      商品名: productData[i][productNameIdx],
      カテゴリ: productData[i][categoryIdx]
    };
  }

  const result = [];
  for (let i = 1; i < salesData.length; i++) {
    const row = salesData[i];
    // 空行スキップ
    if (row.every(cell => cell === '' || cell === null)) continue;
    const productId = row[idx('商品ID')];
    let saleDate = row[idx('販売日')];
    if (saleDate instanceof Date) {
      saleDate = Utilities.formatDate(saleDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    }
    result.push({
      取引ID: row[idx('取引ID')],
      出品ID: row[idx('出品ID')],
      商品ID: productId,
      商品名: productMap[productId] ? productMap[productId].商品名 : '',
      カテゴリ: productMap[productId] ? productMap[productId].カテゴリ : '',
      販売日: saleDate,
      販売価格: row[idx('販売価格')],
      販売手数料: row[idx('販売手数料')],
      送料: row[idx('送料')],
      購入者情報: row[idx('購入者情報')],
      取引ステータス: row[idx('取引ステータス')]
    });
  }
  return result;
}

// ①新規登録: 商品マスタ＆在庫管理に登録
function registerNewProduct(product, stock, status) {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  // 商品マスタ
  let productSheet = ss.getSheetByName('商品マスタ');
  if (!productSheet) productSheet = ss.insertSheet('商品マスタ');
  // 在庫管理
  let inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) inventorySheet = ss.insertSheet('在庫管理');

  // 商品ID発番
  const productId = 'P' + Date.now() + Math.floor(Math.random() * 1000);
  // 商品マスタ登録
  productSheet.appendRow([
    productId,
    product.商品名,
    product.カテゴリ,
    product.仕入れ価格,
    product.販売予定価格,
    product.状態 || '',
    product.備考 || ''
  ]);
  // ステータス選択（デフォルトは「出品可能」）
  const validStatus = ['出品可能', '仕入中', '在庫切れ'];
  let inventoryStatus = status && validStatus.includes(status) ? status : '出品可能';
  // 在庫管理登録
  inventorySheet.appendRow([
    productId,
    stock,
    inventoryStatus,
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  ]);
  return productId;
}

// ②出品登録: 在庫管理の在庫数減算＆出品管理に追加
function registerListing(productId, listingInfo) {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  // 在庫管理
  let inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) inventorySheet = ss.insertSheet('在庫管理');
  // 出品管理
  let listingSheet = ss.getSheetByName('出品管理');
  if (!listingSheet) listingSheet = ss.insertSheet('出品管理');
  // 出品ID発番
  const listingId = 'L' + Date.now() + Math.floor(Math.random() * 1000);
  // 在庫数減算
  const inventoryData = inventorySheet.getDataRange().getValues();
  for (let i = 1; i < inventoryData.length; i++) {
    if (String(inventoryData[i][0]).trim() === String(productId).trim()) {
      let stock = Number(inventoryData[i][1]);
      if (stock <= 0) throw new Error('在庫がありません');
      inventorySheet.getRange(i+1, 2).setValue(stock - 1);
      inventorySheet.getRange(i+1, 3).setValue('出品中');
      inventorySheet.getRange(i+1, 4).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
      break;
    }
  }
  // 出品管理追加
  listingSheet.appendRow([
    listingId,
    productId,
    listingInfo.出品日,
    listingInfo.出品価格,
    '出品中',
    listingInfo.備考 || ''
  ]);
  return listingId;
}

// ③販売登録: 出品管理の該当行を完了＆販売管理に追加
function registerSale(listingId, saleInfo) {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  // 出品管理
  let listingSheet = ss.getSheetByName('出品管理');
  if (!listingSheet) throw new Error('出品管理シートが存在しません');
  // 販売管理
  let salesSheet = ss.getSheetByName('販売管理');
  if (!salesSheet) salesSheet = ss.insertSheet('販売管理');
  // listingIdから商品ID取得＆ステータス完了
  const listingData = listingSheet.getDataRange().getValues();
  let productId = '';
  for (let i = 1; i < listingData.length; i++) {
    if (String(listingData[i][0]).trim() === String(listingId).trim()) {
      productId = listingData[i][1];
      listingSheet.getRange(i+1, 5).setValue('完了');
      break;
    }
  }
  if (!productId) throw new Error('該当する出品IDが見つかりません');
  // 販売管理追加
  const saleId = 'T' + Date.now() + Math.floor(Math.random() * 1000);
  salesSheet.appendRow([
    saleId,
    listingId,
    productId,
    saleInfo.販売日,
    saleInfo.販売価格,
    saleInfo.販売手数料,
    saleInfo.送料,
    saleInfo.購入者情報 || '',
    saleInfo.取引ステータス || '売約済み'
  ]);
  return saleId;
} 