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
  inventorySheet.getRange(inventoryRowIdx, inventoryHeaders.indexOf('在庫数') + 1).setValue(在庫数 - listing.出品数);
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
function getListingList() {
  const properties = PropertiesService.getScriptProperties();
  const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
  if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('出品管理');
  if (!sheet) throw new Error('出品管理シートが存在しません');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, idx) => row[h] = data[i][idx]);
    list.push(row);
  }
  return list;
} 