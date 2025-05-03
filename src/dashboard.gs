// ダッシュボード用サマリーデータ取得
function getDashboardSummary() {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const salesSheet = ss.getSheetByName('販売管理');
  const inventorySheet = ss.getSheetByName('在庫管理');
  const productSheet = ss.getSheetByName('商品マスタ');
  if (!salesSheet || !inventorySheet || !productSheet) throw new Error('必要なシートが存在しません');

  // 販売管理集計
  const salesData = salesSheet.getDataRange().getValues();
  let totalSales = 0;
  let totalProfit = 0;
  let salesCount = 0;
  for (let i = 1; i < salesData.length; i++) {
    const price = Number(salesData[i][3]);
    const fee = Number(salesData[i][4]);
    const shipping = Number(salesData[i][5]);
    // 商品IDから仕入れ価格取得
    const productId = salesData[i][1];
    let cost = 0;
    const productData = productSheet.getDataRange().getValues();
    for (let j = 1; j < productData.length; j++) {
      if (productData[j][0] === productId) {
        cost = Number(productData[j][3]);
        break;
      }
    }
    totalSales += price;
    totalProfit += (price - cost - fee - shipping);
    salesCount++;
  }

  // 在庫管理集計
  const inventoryData = inventorySheet.getDataRange().getValues();
  let totalStock = 0;
  let outOfStockCount = 0;
  let listedCount = 0;
  for (let i = 1; i < inventoryData.length; i++) {
    const stock = Number(inventoryData[i][1]);
    const status = inventoryData[i][2];
    totalStock += stock;
    if (stock === 0) outOfStockCount++;
    if (status === '出品中') listedCount++;
  }

  return {
    totalSales,
    salesCount,
    totalProfit,
    totalStock,
    outOfStockCount,
    listedCount
  };
} 