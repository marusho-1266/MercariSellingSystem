// 在庫管理CRUD: 在庫新規登録
function createInventory(inventory) {
  // 必須バリデーション
  if (!inventory.商品ID) {
    throw new Error('商品IDは必須です');
  }
  if (isNaN(inventory.在庫数)) {
    throw new Error('在庫数は数値で入力してください');
  }
  // ステータスは任意だが、なければ「未出品」
  const status = inventory.ステータス || '未出品';
  const validStatus = ['未出品', '出品中', '売約済み', '発送済み'];
  if (!validStatus.includes(status)) {
    throw new Error('ステータスが不正です');
  }
  // 日付
  const updateDate = inventory.最終更新日 || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  // スプレッドシート取得
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('在庫管理');
  if (!sheet) throw new Error('在庫管理シートが存在しません');

  // 商品ID重複チェック
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === inventory.商品ID) {
      throw new Error('この商品IDの在庫はすでに登録されています');
    }
  }

  // 追加データ作成
  const row = [
    inventory.商品ID,
    Number(inventory.在庫数),
    status,
    updateDate
  ];
  sheet.appendRow(row);
  return inventory.商品ID;
}

// 在庫管理CRUD: 商品IDで在庫情報取得
function getInventoryByProductId(productId) {
  if (!productId) throw new Error('商品IDは必須です');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('在庫管理');
  if (!sheet) throw new Error('在庫管理シートが存在しません');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      const inventory = {};
      headers.forEach((h, idx) => inventory[h] = data[i][idx]);
      return inventory;
    }
  }
  return null;
}

// 在庫管理CRUD: 在庫情報の更新
function updateInventory(productId, updateFields) {
  if (!productId) throw new Error('商品IDは必須です');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('在庫管理');
  if (!sheet) throw new Error('在庫管理シートが存在しません');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      for (let key in updateFields) {
        const colIdx = headers.indexOf(key);
        if (colIdx !== -1) {
          sheet.getRange(i+1, colIdx+1).setValue(updateFields[key]);
        }
      }
      // 最終更新日も自動更新
      const dateColIdx = headers.indexOf('最終更新日');
      if (dateColIdx !== -1) {
        sheet.getRange(i+1, dateColIdx+1).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
      }
      return true;
    }
  }
  throw new Error('該当する商品IDが見つかりません');
}

// 在庫管理CRUD: 在庫情報削除
function deleteInventory(productId) {
  if (!productId) throw new Error('商品IDは必須です');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('在庫管理');
  if (!sheet) throw new Error('在庫管理シートが存在しません');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      sheet.deleteRow(i+1);
      return true;
    }
  }
  throw new Error('該当する商品IDが見つかりません');
} 