// 商品マスタCRUD: 商品登録
function createProduct(product) {
  // 必須バリデーション
  if (!product.商品名 || !product.カテゴリ) {
    throw new Error('商品名とカテゴリは必須です');
  }
  if (isNaN(product.仕入れ価格) || isNaN(product.販売予定価格)) {
    throw new Error('価格は数値で入力してください');
  }

  // スプレッドシート取得
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('商品マスタ');
  if (!sheet) throw new Error('商品マスタシートが存在しません');

  // 商品ID生成（タイムスタンプ＋ランダム）
  const productId = 'P' + Date.now() + Math.floor(Math.random() * 1000);

  // 既存商品ID重複チェック
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      throw new Error('商品ID重複エラー');
    }
  }

  // 追加データ作成
  const row = [
    productId,
    product.商品名,
    product.カテゴリ,
    Number(product.仕入れ価格),
    Number(product.販売予定価格),
    product.状態 || '',
    product.備考 || ''
  ];
  sheet.appendRow(row);
  return productId;
}

// 商品マスタCRUD: 商品IDで検索
function getProductById(productId) {
  if (!productId) throw new Error('商品IDは必須です');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('商品マスタ');
  if (!sheet) throw new Error('商品マスタシートが存在しません');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      const product = {};
      headers.forEach((h, idx) => product[h] = data[i][idx]);
      return product;
    }
  }
  return null;
}

// 商品マスタCRUD: 商品情報の更新
function updateProduct(productId, updateFields) {
  if (!productId) throw new Error('商品IDは必須です');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('商品マスタ');
  if (!sheet) throw new Error('商品マスタシートが存在しません');
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
      return true;
    }
  }
  throw new Error('該当する商品IDが見つかりません');
}

// 商品マスタCRUD: 商品削除
function deleteProduct(productId) {
  if (!productId) throw new Error('商品IDは必須です');
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('商品マスタ');
  if (!sheet) throw new Error('商品マスタシートが存在しません');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      sheet.deleteRow(i+1);
      return true;
    }
  }
  throw new Error('該当する商品IDが見つかりません');
}

// 商品登録＋在庫登録をまとめて行う関数（UIから呼び出し用）
function registerProductAndInventory(product, inventory) {
  try {
    // 商品マスタ登録
    var productId = createProduct(product);
    // 在庫管理登録
    inventory.商品ID = productId;
    createInventory(inventory);
    return productId;
  } catch (e) {
    throw new Error(e.message || e);
  }
} 