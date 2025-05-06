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

// 商品IDが商品マスタに存在するかチェックする共通関数
function validateProductIdExists(productId) {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const productSheet = ss.getSheetByName('商品マスタ');
  if (!productSheet) throw new Error('商品マスタシートが存在しません');
  const productData = productSheet.getDataRange().getValues();
  for (let i = 1; i < productData.length; i++) {
    if (String(productData[i][0]).trim() === String(productId).trim()) {
      return true;
    }
  }
  throw new Error('商品IDが商品マスタに存在しません: ' + productId);
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
  // 在庫管理への登録処理は削除
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
      listingSheet.getRange(i+1, 6).setValue('完了');
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

// 仕入登録: 仕入管理シートに履歴追加、在庫数加算、ステータス自動遷移
function registerPurchase(purchase) {
  // 必須バリデーション
  validateProductIdExists(purchase.商品ID);
  if (!purchase.商品ID) throw new Error('商品IDは必須です');
  if (!purchase.仕入日) throw new Error('仕入日は必須です');
  if (isNaN(purchase.仕入数) || purchase.仕入数 <= 0 || !Number.isInteger(Number(purchase.仕入数))) throw new Error('仕入数は1以上の整数で入力してください');
  if (isNaN(purchase.仕入価格) || purchase.仕入価格 < 0) throw new Error('仕入価格は0以上の数値で入力してください');

  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);

  // 仕入管理シート
  let purchaseSheet = ss.getSheetByName('仕入管理');
  if (!purchaseSheet) throw new Error('仕入管理シートが存在しません');

  // 在庫管理シート
  let inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) throw new Error('在庫管理シートが存在しません');
  const inventoryData = inventorySheet.getDataRange().getValues();
  const inventoryHeaders = inventoryData[0];
  let inventoryRowIdx = -1;
  let currentStock = 0;
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === purchase.商品ID) {
      inventoryRowIdx = i + 1;
      currentStock = Number(inventoryData[i][1]);
      break;
    }
  }
  // 在庫管理に行がなければ新規追加
  if (inventoryRowIdx === -1) {
    // 初期在庫は仕入数、ステータスは在庫数>0なら「出品可能」それ以外は「仕入中」
    const newStock = Number(purchase.仕入数);
    const newStatus = newStock > 0 ? '出品可能' : '仕入中';
    inventorySheet.appendRow([
      purchase.商品ID,
      newStock,
      newStatus,
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    ]);
    inventoryRowIdx = inventorySheet.getLastRow();
    currentStock = newStock;
  }

  // 在庫数加算
  const newStock = currentStock + Number(purchase.仕入数);
  if (newStock < 0) throw new Error('在庫数が負の値になります');
  inventorySheet.getRange(inventoryRowIdx, 2).setValue(newStock);
  // ステータス自動遷移（在庫数>0なら「出品可能」）
  const statusColIdx = inventoryHeaders.indexOf('ステータス');
  if (statusColIdx !== -1) {
    const newStatus = newStock > 0 ? '出品可能' : '仕入中';
    inventorySheet.getRange(inventoryRowIdx, statusColIdx + 1).setValue(newStatus);
  }
  // 最終更新日も更新
  const dateColIdx = inventoryHeaders.indexOf('最終更新日');
  if (dateColIdx !== -1) {
    inventorySheet.getRange(inventoryRowIdx, dateColIdx + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  }

  // 仕入ID発番
  const purchaseId = 'S' + Date.now() + Math.floor(Math.random() * 1000);
  // 仕入管理シートに履歴追加
  purchaseSheet.appendRow([
    purchaseId,
    purchase.商品ID,
    purchase.仕入日,
    Number(purchase.仕入数),
    Number(purchase.仕入価格),
    purchase.ステータス || '完了',
    purchase.備考 || ''
  ]);
  return purchaseId;
}

// 仕入更新: 仕入管理シートの該当行編集、在庫数調整
function updatePurchase(purchaseId, updateFields) {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);

  // 仕入管理シート
  let purchaseSheet = ss.getSheetByName('仕入管理');
  if (!purchaseSheet) throw new Error('仕入管理シートが存在しません');
  const purchaseData = purchaseSheet.getDataRange().getValues();
  const purchaseHeaders = purchaseData[0];
  let rowIdx = -1;
  let oldProductId = '';
  let oldQty = 0;
  for (let i = 1; i < purchaseData.length; i++) {
    if (purchaseData[i][0] === purchaseId) {
      rowIdx = i + 1;
      oldProductId = purchaseData[i][1];
      oldQty = Number(purchaseData[i][3]);
      break;
    }
  }
  if (rowIdx === -1) throw new Error('該当する仕入IDが見つかりません');

  // 在庫管理シート
  let inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) throw new Error('在庫管理シートが存在しません');
  const inventoryData = inventorySheet.getDataRange().getValues();
  const inventoryHeaders = inventoryData[0];
  let invRowIdx = -1;
  let currentStock = 0;
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === oldProductId) {
      invRowIdx = i + 1;
      currentStock = Number(inventoryData[i][1]);
      break;
    }
  }
  if (invRowIdx === -1) throw new Error('該当する商品IDの在庫が見つかりません');

  validateProductIdExists(oldProductId);

  // 仕入数の増減分を計算
  let newQty = updateFields.仕入数 !== undefined ? Number(updateFields.仕入数) : oldQty;
  if (isNaN(newQty) || newQty <= 0 || !Number.isInteger(newQty)) throw new Error('仕入数は1以上の整数で入力してください');
  if (updateFields.仕入価格 !== undefined && (isNaN(Number(updateFields.仕入価格)) || Number(updateFields.仕入価格) < 0)) throw new Error('仕入価格は0以上の数値で入力してください');
  let diff = newQty - oldQty;
  // 在庫数調整
  if (currentStock + diff < 0) throw new Error('在庫数が負の値になります');
  inventorySheet.getRange(invRowIdx, 2).setValue(currentStock + diff);
  // ステータス自動遷移
  const statusColIdx = inventoryHeaders.indexOf('ステータス');
  if (statusColIdx !== -1) {
    const newStock = currentStock + diff;
    const newStatus = newStock > 0 ? '出品可能' : '仕入中';
    inventorySheet.getRange(invRowIdx, statusColIdx + 1).setValue(newStatus);
  }
  // 最終更新日も更新
  const dateColIdx = inventoryHeaders.indexOf('最終更新日');
  if (dateColIdx !== -1) {
    inventorySheet.getRange(invRowIdx, dateColIdx + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  }

  // 仕入管理シートの該当行を更新
  for (const key in updateFields) {
    const colIdx = purchaseHeaders.indexOf(key);
    if (colIdx !== -1) {
      purchaseSheet.getRange(rowIdx, colIdx + 1).setValue(updateFields[key]);
    }
  }
  return true;
}

// 仕入削除: 仕入管理シートの該当行削除、在庫数調整
function deletePurchase(purchaseId) {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);

  // 仕入管理シート
  let purchaseSheet = ss.getSheetByName('仕入管理');
  if (!purchaseSheet) throw new Error('仕入管理シートが存在しません');
  const purchaseData = purchaseSheet.getDataRange().getValues();
  const purchaseHeaders = purchaseData[0];
  let rowIdx = -1;
  let productId = '';
  let qty = 0;
  for (let i = 1; i < purchaseData.length; i++) {
    if (purchaseData[i][0] === purchaseId) {
      rowIdx = i + 1;
      productId = purchaseData[i][1];
      qty = Number(purchaseData[i][3]);
      break;
    }
  }
  if (rowIdx === -1) throw new Error('該当する仕入IDが見つかりません');

  // 在庫管理シート
  let inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) throw new Error('在庫管理シートが存在しません');
  const inventoryData = inventorySheet.getDataRange().getValues();
  const inventoryHeaders = inventoryData[0];
  let invRowIdx = -1;
  let currentStock = 0;
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === productId) {
      invRowIdx = i + 1;
      currentStock = Number(inventoryData[i][1]);
      break;
    }
  }
  if (invRowIdx === -1) throw new Error('該当する商品IDの在庫が見つかりません');

  validateProductIdExists(productId);

  // 在庫数減算
  const newStock = Math.max(0, currentStock - qty);
  if (newStock < 0) throw new Error('在庫数が負の値になります');
  inventorySheet.getRange(invRowIdx, 2).setValue(newStock);
  // ステータス自動遷移
  const statusColIdx = inventoryHeaders.indexOf('ステータス');
  if (statusColIdx !== -1) {
    const newStatus = newStock > 0 ? '出品可能' : '仕入中';
    inventorySheet.getRange(invRowIdx, statusColIdx + 1).setValue(newStatus);
  }
  // 最終更新日も更新
  const dateColIdx = inventoryHeaders.indexOf('最終更新日');
  if (dateColIdx !== -1) {
    inventorySheet.getRange(invRowIdx, dateColIdx + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  }

  // 仕入管理シートの該当行削除
  purchaseSheet.deleteRow(rowIdx);
  return true;
}

// 仕入リスト取得: 仕入管理シートの履歴一覧取得
function getPurchaseList() {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const purchaseSheet = ss.getSheetByName('仕入管理');
  if (!purchaseSheet) throw new Error('仕入管理シートが存在しません');
  const data = purchaseSheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    // 空行スキップ
    if (rowData.every(cell => cell === '' || cell === null)) continue;
    const row = {};
    headers.forEach((h, idx) => {
      let value = rowData[idx];
      // Date型は文字列化
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      }
      row[h] = value;
    });
    list.push(row);
  }
  return list;
} 