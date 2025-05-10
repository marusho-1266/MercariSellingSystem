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
  let isNewInventory = false;
  if (inventoryRowIdx === -1) {
    // 在庫管理への新規追加・加算処理は削除
    // ここでは何もしない
  }
  // 既存行がある場合も在庫数加算処理は削除
  // ここでは何もしない

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
  // validateProductIdExists(oldProductId); // 商品マスタ存在チェックはOK
  // 在庫管理の「該当する商品IDの在庫が見つかりません」エラーthrow部分は削除

  // ステータス変更前の値を取得
  const oldStatusColIdx = purchaseHeaders.indexOf('ステータス');
  const oldStatus = oldStatusColIdx !== -1 ? purchaseData[rowIdx-1][oldStatusColIdx] : '';
  let newStatus = updateFields.ステータス !== undefined ? updateFields.ステータス : oldStatus;
  let newQty = updateFields.仕入数 !== undefined ? Number(updateFields.仕入数) : oldQty;

  // ステータスが「仕入中」→「完了」に変わった場合のみ在庫管理に加算
  if (oldStatus === '仕入中' && newStatus === '完了') {
    // 在庫管理シートに行がなければ新規追加、あれば加算
    let found = false;
    for (let i = 1; i < inventoryData.length; i++) {
      if (inventoryData[i][0] === oldProductId) {
        // 既存行：在庫数加算
        const currentStock = Number(inventoryData[i][1]);
        const newStock = currentStock + newQty;
        inventorySheet.getRange(i+1, 2).setValue(newStock);
        // ステータス自動遷移
        const statusColIdx = inventoryHeaders.indexOf('ステータス');
        if (statusColIdx !== -1) {
          const status = newStock > 0 ? '出品可能' : '仕入中';
          inventorySheet.getRange(i+1, statusColIdx+1).setValue(status);
        }
        // 最終更新日も更新
        const dateColIdx = inventoryHeaders.indexOf('最終更新日');
        if (dateColIdx !== -1) {
          inventorySheet.getRange(i+1, dateColIdx+1).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
        }
        found = true;
        break;
      }
    }
    if (!found) {
      // 新規追加
      inventorySheet.appendRow([
        oldProductId,
        newQty,
        newQty > 0 ? '出品可能' : '仕入中',
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
      ]);
    }
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
  Logger.log('getPurchaseList開始');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const purchaseSheet = ss.getSheetByName('仕入管理');
  if (!purchaseSheet) throw new Error('仕入管理シートが存在しません');
  const purchaseData = purchaseSheet.getDataRange().getValues();
  const purchaseHeaders = purchaseData[0];
  
  Logger.log('仕入管理シート行数: ' + purchaseData.length);
  
  // 商品マスタから商品名を取得するための準備
  const productSheet = ss.getSheetByName('商品マスタ');
  if (!productSheet) throw new Error('商品マスタシートが存在しません');
  const productData = productSheet.getDataRange().getValues();
  const productHeaders = productData[0];
  const productIdIdx = productHeaders.indexOf('商品ID');
  const productNameIdx = productHeaders.indexOf('商品名');
  
  // 商品IDから商品名を取得しやすいようにマップを作成
  const productMap = {};
  for (let i = 1; i < productData.length; i++) {
    if (productData[i][productIdIdx]) {
      productMap[productData[i][productIdIdx]] = productData[i][productNameIdx];
    }
  }
  Logger.log('商品マスタから商品マップ作成: ' + Object.keys(productMap).length + '件');
  
  const result = [];
  const productIdColIdx = purchaseHeaders.indexOf('商品ID');
  
  if (productIdColIdx === -1) {
    Logger.log('仕入管理シートに商品ID列が見つかりません');
    throw new Error('仕入管理シートのフォーマットが不正です: 商品ID列がありません');
  }
  
  for (let i = 1; i < purchaseData.length; i++) {
    // 空行チェック（すべてのセルが空かnullかチェック）
    if (purchaseData[i].every(cell => cell === '' || cell === null)) {
      Logger.log('空行をスキップ: ' + i);
      continue;
    }
    
    const row = {};
    for (let j = 0; j < purchaseHeaders.length; j++) {
      row[purchaseHeaders[j]] = purchaseData[i][j];
    }
    
    // 商品名を追加
    const productId = purchaseData[i][productIdColIdx];
    Logger.log('処理中の商品ID: ' + productId);
    
    if (productId && productMap[productId]) {
      row['商品名'] = productMap[productId];
      Logger.log('商品名を設定: ' + row['商品名']);
    } else {
      row['商品名'] = '不明';
      Logger.log('商品名が見つからないため不明を設定');
    }
    
    // ログ出力
    Logger.log('仕入データ追加: ' + JSON.stringify(row));
    
    result.push(row);
  }
  
  Logger.log('取得結果件数: ' + result.length);
  return result;
} 