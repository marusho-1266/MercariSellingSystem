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