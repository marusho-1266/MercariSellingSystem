// 出品登録関数: 在庫から出品情報を登録し、在庫数を減算
function registerListing(listing) {
  // 必須バリデーション
  if (!listing.商品ID) throw new Error('商品IDは必須です');
  if (isNaN(listing.出品数) || listing.出品数 <= 0) throw new Error('出品数は1以上の数値で入力してください');
  if (isNaN(listing.出品価格) || listing.出品価格 <= 0) throw new Error('出品価格は1以上の数値で入力してください');

  // スプレッドシート取得
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);

  // 在庫管理シート取得
  const inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) throw new Error('在庫管理シートが存在しません');
  const inventoryData = inventorySheet.getDataRange().getValues();
  const inventoryHeaders = inventoryData[0];
  let inventoryRow = null;
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === listing.商品ID) {
      inventoryRow = inventoryData[i];
      var inventoryRowIdx = i + 1;
      break;
    }
  }
  if (!inventoryRow) throw new Error('該当する商品IDの在庫が見つかりません');
  const 在庫数 = Number(inventoryRow[inventoryHeaders.indexOf('在庫数')]);
  if (在庫数 < listing.出品数) throw new Error('在庫数が不足しています');

  // 在庫数減算
  const newStock = 在庫数 - listing.出品数;
  inventorySheet.getRange(inventoryRowIdx, inventoryHeaders.indexOf('在庫数') + 1).setValue(newStock);
  
  // 在庫数が0になったらステータスを「在庫切れ」に更新
  const statusColIdx = inventoryHeaders.indexOf('ステータス');
  if (statusColIdx !== -1 && newStock === 0) {
    inventorySheet.getRange(inventoryRowIdx, statusColIdx + 1).setValue('在庫切れ');
  }
  
  // 最終更新日も更新
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const updateColIdx = inventoryHeaders.indexOf('最終更新日');
  if (updateColIdx !== -1) {
    inventorySheet.getRange(inventoryRowIdx, updateColIdx + 1).setValue(dateStr);
  }

  // 出品管理シート取得
  const listingSheet = ss.getSheetByName('出品管理');
  if (!listingSheet) throw new Error('出品管理シートが存在しません');

  // 出品ID自動発番
  const listingId = 'L' + new Date().getTime();

  // 出品データ作成
  const row = [
    listingId,
    listing.商品ID,
    dateStr,
    Number(listing.出品価格),
    Number(listing.出品数),
    '出品中',
    listing.備考 || ''
  ];
  listingSheet.appendRow(row);
  return listingId;
}

// 出品一覧取得関数: 出品管理シートの全データを配列で返す
function getListingList(filterStatus) {
  try { // エラーハンドリングを追加
    Logger.log('getListingList started. filterStatus: ' + filterStatus);
    const properties = PropertiesService.getScriptProperties();
    const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
    if (!ssId) {
      Logger.log('Error: MASTER_SPREADSHEET_ID not found.');
      throw new Error('マスタースプレッドシートが未作成です');
    }
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName('出品管理');
    if (!sheet) {
      Logger.log('Error: 出品管理シートが存在しません');
      throw new Error('出品管理シートが存在しません');
    }
    
    // 商品マスタから商品名を取得するための準備
    const productSheet = ss.getSheetByName('商品マスタ');
    if (!productSheet) {
      Logger.log('Error: 商品マスタシートが存在しません');
      throw new Error('商品マスタシートが存在しません');
    }
    const productData = productSheet.getDataRange().getValues();
    const productHeaders = productData[0];
    const productIdIdx = productHeaders.indexOf('商品ID');
    const productNameIdx = productHeaders.indexOf('商品名');
    
    // 商品IDから商品名を取得しやすいようにマップを作成
    const productMap = {};
    for (let i = 1; i < productData.length; i++) {
      if (productData[i][productIdIdx]) {
        productMap[productData[i][productIdIdx]] = productData[i][productNameIdx];
        Logger.log('商品ID: ' + productData[i][productIdIdx] + ' → 商品名: ' + productData[i][productNameIdx]);
      }
    }
    Logger.log('商品マスタマップ作成完了: ' + Object.keys(productMap).length + '件');
    
    const data = sheet.getDataRange().getValues();
    Logger.log('Raw data length from 出品管理 sheet: ' + data.length);
    if (data.length < 2) {
      Logger.log('No data rows found (only header or empty). Returning empty list.');
      return []; // ヘッダーのみ、または空の場合は空配列を返す
    }
    const headers = data[0];
    Logger.log('Headers: ' + JSON.stringify(headers));
    const productIdColIdx = headers.indexOf('商品ID');
    const statusColIdx = headers.indexOf('ステータス');
    
    const list = [];
    for (let i = 1; i < data.length; i++) {
      const rowData = data[i];
      
      // 空行スキップ（すべてのセルが空かnullかチェック）
      if (rowData.every(cell => cell === '' || cell === null)) {
        Logger.log('空行をスキップ: ' + i);
        continue;
      }
      
      // ステータスでフィルタリング
      if (filterStatus && statusColIdx !== -1) {
        const status = rowData[statusColIdx];
        if (status !== filterStatus) {
          Logger.log(`行 ${i+1} はステータス「${status}」なので、フィルター「${filterStatus}」に一致せずスキップします`);
          continue;
        }
      }
      
      Logger.log('Processing rowData: ' + JSON.stringify(rowData)); // Log raw row data
      const row = {};
      headers.forEach((h, idx) => {
        let value = rowData[idx];
        // Explicitly convert Date objects to a reliable string format for transfer
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        }
        row[h] = value;
      });
      
      // 商品名を追加
      if (productIdColIdx !== -1) {
        const productId = rowData[productIdColIdx];
        if (productId && productMap[productId]) {
          row['商品名'] = productMap[productId];
          Logger.log('商品名を設定: ' + row['商品名'] + ' (商品ID: ' + productId + ')');
        } else {
          row['商品名'] = '不明';
          Logger.log('商品名が見つからないため不明を設定 (商品ID: ' + productId + ')');
        }
      }
      
      Logger.log('Constructed row object: ' + JSON.stringify(row)); // Log constructed object
      list.push(row);
    }
    // Log the final list just before returning
    try {
      Logger.log('Final list being returned: ' + list.length + ' 件');
    } catch (stringifyError) {
      Logger.log('Error stringifying final list: ' + stringifyError.message);
      // If stringify fails here, that's a big clue
    }
    return list;
  } catch (e) {
    Logger.log('Error in getListingList: ' + e.message + ' Stack: ' + e.stack);
    // フロントエンドの failureHandler でエラーをキャッチできるように再スロー
    throw new Error('出品一覧の取得中にエラーが発生しました: ' + e.message);
  }
}

// 出品中の出品ID＋商品名リスト取得（販売登録用）
function getListableListings() {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const listingSheet = ss.getSheetByName('出品管理');
  const productSheet = ss.getSheetByName('商品マスタ');
  if (!listingSheet || !productSheet) throw new Error('必要なシートが存在しません');

  const listingData = listingSheet.getDataRange().getValues();
  const productData = productSheet.getDataRange().getValues();
  const productMap = {};
  for (let i = 1; i < productData.length; i++) {
    productMap[productData[i][0]] = productData[i][1]; // 商品ID→商品名
  }
  const result = [];
  const statusIdx = listingData[0].indexOf('ステータス');
  const listingIdIdx = listingData[0].indexOf('出品ID');
  const productIdIdx = listingData[0].indexOf('商品ID');
  for (let i = 1; i < listingData.length; i++) {
    if (listingData[i][statusIdx] === '出品中') {
      result.push({
        出品ID: listingData[i][listingIdIdx],
        商品名: productMap[listingData[i][productIdIdx]] || ''
      });
    }
  }
  return result;
} 