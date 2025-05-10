// 売上記録機能: 販売情報の記録
function createSalesRecord(sales) {
  // バリデーション
  if (!sales.商品ID) throw new Error('商品IDは必須です');
  if (!sales.販売日) throw new Error('販売日は必須です');
  if (isNaN(sales.販売価格)) throw new Error('販売価格は数値で入力してください');
  if (isNaN(sales.販売手数料)) throw new Error('販売手数料は数値で入力してください');
  if (isNaN(sales.送料)) throw new Error('送料は数値で入力してください');

  // スプレッドシート取得
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const salesSheet = ss.getSheetByName('販売管理');
  if (!salesSheet) throw new Error('販売管理シートが存在しません');
  const inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) throw new Error('在庫管理シートが存在しません');

  // 商品ID存在チェック＆在庫数取得
  const inventoryData = inventorySheet.getDataRange().getValues();
  let inventoryRow = -1;
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === sales.商品ID) {
      inventoryRow = i;
      break;
    }
  }
  if (inventoryRow === -1) throw new Error('該当する商品IDの在庫が見つかりません');

  // 在庫数減算＆ステータス更新
  let stock = Number(inventoryData[inventoryRow][1]);
  if (stock <= 0) throw new Error('在庫がありません');
  stock -= 1;
  inventorySheet.getRange(inventoryRow+1, 2).setValue(stock);
  // ステータス更新（売約済みに）
  inventorySheet.getRange(inventoryRow+1, 3).setValue('売約済み');
  // 最終更新日更新
  inventorySheet.getRange(inventoryRow+1, 4).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));

  // 取引ID生成
  const transactionId = 'T' + Date.now() + Math.floor(Math.random() * 1000);

  // 販売管理シートに記録
  const row = [
    transactionId,
    sales.商品ID,
    sales.販売日,
    Number(sales.販売価格),
    Number(sales.販売手数料),
    Number(sales.送料),
    sales.購入者情報 || '',
    sales.取引ステータス || '売約済み'
  ];
  salesSheet.appendRow(row);
  return transactionId;
}

// 取引ステータス管理: 取引ステータスの更新
function updateTransactionStatus(transactionId, newStatus) {
  const validStatus = ['出品中', '売約済み', '発送待ち', '評価待ち', '取引完了'];
  if (!transactionId) throw new Error('取引IDは必須です');
  if (!validStatus.includes(newStatus)) throw new Error('不正なステータスです');

  // スプレッドシート取得
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const salesSheet = ss.getSheetByName('販売管理');
  if (!salesSheet) throw new Error('販売管理シートが存在しません');
  const inventorySheet = ss.getSheetByName('在庫管理');
  if (!inventorySheet) throw new Error('在庫管理シートが存在しません');

  // 販売管理シートで該当取引IDを検索
  const salesData = salesSheet.getDataRange().getValues();
  const salesHeaders = salesData[0];
  let salesRow = -1;
  let productId = '';
  for (let i = 1; i < salesData.length; i++) {
    if (salesData[i][0] === transactionId) {
      salesRow = i;
      productId = salesData[i][1];
      break;
    }
  }
  if (salesRow === -1) throw new Error('該当する取引IDが見つかりません');

  // ステータス更新
  const statusColIdx = salesHeaders.indexOf('取引ステータス');
  if (statusColIdx === -1) throw new Error('取引ステータス列が見つかりません');
  salesSheet.getRange(salesRow+1, statusColIdx+1).setValue(newStatus);

  // 在庫管理シートも連動（商品IDで検索）
  if (productId) {
    const inventoryData = inventorySheet.getDataRange().getValues();
    const inventoryHeaders = inventoryData[0];
    for (let j = 1; j < inventoryData.length; j++) {
      if (inventoryData[j][0] === productId) {
        const invStatusColIdx = inventoryHeaders.indexOf('ステータス');
        if (invStatusColIdx !== -1) {
          inventorySheet.getRange(j+1, invStatusColIdx+1).setValue(newStatus);
        }
        break;
      }
    }
  }
  return true;
} 