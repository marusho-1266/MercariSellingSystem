/**
 * 財務レポート関連の機能を提供するモジュール
 */

/**
 * 基本財務指標を取得する
 * @returns {Object} 財務指標のオブジェクト
 */
function getFinanceReport() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
    if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
    
    const ss = SpreadsheetApp.openById(ssId);
    const salesSheet = ss.getSheetByName('販売管理');
    const purchaseSheet = ss.getSheetByName('仕入管理');
    
    if (!salesSheet || !purchaseSheet) throw new Error('必要なシートが存在しません');
    
    // データ取得
    const salesData = salesSheet.getDataRange().getValues();
    const purchaseData = purchaseSheet.getDataRange().getValues();
    
    // ヘッダー行を除外
    const salesRows = salesData.slice(1);
    const purchaseRows = purchaseData.slice(1);
    
    // ヘッダーのインデックスを取得
    const salesHeaders = salesData[0];
    const purchaseHeaders = purchaseData[0];
    
    const salePriceIdx = salesHeaders.indexOf('販売価格');
    const saleFeeIdx = salesHeaders.indexOf('販売手数料');
    const shippingFeeIdx = salesHeaders.indexOf('送料');
    const saleStatusIdx = salesHeaders.indexOf('取引ステータス');
    
    const purchasePriceIdx = purchaseHeaders.indexOf('仕入価格');
    const purchaseStatusIdx = purchaseHeaders.indexOf('ステータス');
    
    // 集計
    let totalSales = 0;
    let totalFees = 0;
    let totalShipping = 0;
    let completedSales = 0;
    let pendingSales = 0;
    
    for (const row of salesRows) {
      // 取引完了の販売のみ集計
      if (row[saleStatusIdx] === '取引完了') {
        totalSales += Number(row[salePriceIdx]) || 0;
        totalFees += Number(row[saleFeeIdx]) || 0;
        totalShipping += Number(row[shippingFeeIdx]) || 0;
        completedSales++;
      } else {
        pendingSales++;
      }
    }
    
    // 仕入れコスト集計
    let totalPurchaseCost = 0;
    let totalPurchaseCount = 0;
    
    for (const row of purchaseRows) {
      if (row[purchaseStatusIdx] === '完了') {
        totalPurchaseCost += Number(row[purchasePriceIdx]) || 0;
        totalPurchaseCount++;
      }
    }
    
    // 利益計算
    const grossProfit = totalSales - totalFees - totalShipping;
    const netProfit = grossProfit - totalPurchaseCost;
    const profitMargin = totalSales > 0 ? (netProfit / totalSales) * 100 : 0;
    
    return {
      totalSales: totalSales,
      totalFees: totalFees,
      totalShipping: totalShipping,
      grossProfit: grossProfit,
      totalPurchaseCost: totalPurchaseCost,
      netProfit: netProfit,
      profitMargin: profitMargin.toFixed(2),
      completedSales: completedSales,
      pendingSales: pendingSales,
      totalPurchaseCount: totalPurchaseCount
    };
  } catch (e) {
    Logger.log('財務レポート取得エラー: ' + e.toString());
    throw new Error('財務レポートの取得に失敗しました: ' + e.message);
  }
}

/**
 * 期間別の売上データを取得する
 * @param {string} startDate - 開始日 (YYYY-MM-DD形式)
 * @param {string} endDate - 終了日 (YYYY-MM-DD形式)
 * @param {string} groupBy - グループ化の単位 ('day', 'week', 'month')
 * @returns {Array} 期間別売上データの配列
 */
function getPeriodSales(startDate = null, endDate = null, groupBy = 'month') {
  try {
    const properties = PropertiesService.getScriptProperties();
    const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
    if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
    
    const ss = SpreadsheetApp.openById(ssId);
    const salesSheet = ss.getSheetByName('販売管理');
    
    if (!salesSheet) throw new Error('販売管理シートが存在しません');
    
    // データ取得
    const salesData = salesSheet.getDataRange().getValues();
    
    // ヘッダー行を除外
    const salesRows = salesData.slice(1);
    const salesHeaders = salesData[0];
    
    const saleDateIdx = salesHeaders.indexOf('販売日');
    const salePriceIdx = salesHeaders.indexOf('販売価格');
    const saleFeeIdx = salesHeaders.indexOf('販売手数料');
    const shippingFeeIdx = salesHeaders.indexOf('送料');
    const saleStatusIdx = salesHeaders.indexOf('取引ステータス');
    
    // 日付フィルタリングの準備
    const start = startDate ? new Date(startDate) : new Date(new Date().getFullYear(), 0, 1); // 今年の1月1日
    const end = endDate ? new Date(endDate) : new Date(); // 今日
    
    // 期間別データを格納するオブジェクト
    const periodData = {};
    
    for (const row of salesRows) {
      // 販売日の取得と検証
      const saleDate = row[saleDateIdx] instanceof Date ? row[saleDateIdx] : new Date(row[saleDateIdx]);
      
      if (isNaN(saleDate.getTime())) continue; // 無効な日付はスキップ
      
      // 日付フィルタリング
      if (saleDate < start || saleDate > end) continue;
      
      // 取引完了の販売のみ集計
      if (row[saleStatusIdx] !== '取引完了') continue;
      
      // 期間キーの生成
      let periodKey;
      switch (groupBy) {
        case 'day':
          periodKey = Utilities.formatDate(saleDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          break;
        case 'week':
          // 週の初日（日曜日）を取得
          const firstDayOfWeek = new Date(saleDate);
          const day = saleDate.getDay();
          firstDayOfWeek.setDate(saleDate.getDate() - day);
          periodKey = Utilities.formatDate(firstDayOfWeek, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' week';
          break;
        case 'month':
        default:
          periodKey = Utilities.formatDate(saleDate, Session.getScriptTimeZone(), 'yyyy-MM');
          break;
      }
      
      // データ集計
      if (!periodData[periodKey]) {
        periodData[periodKey] = {
          period: periodKey,
          sales: 0,
          fees: 0,
          shipping: 0,
          count: 0
        };
      }
      
      periodData[periodKey].sales += Number(row[salePriceIdx]) || 0;
      periodData[periodKey].fees += Number(row[saleFeeIdx]) || 0;
      periodData[periodKey].shipping += Number(row[shippingFeeIdx]) || 0;
      periodData[periodKey].count++;
    }
    
    // 配列に変換して返す
    return Object.values(periodData).map(item => {
      return {
        ...item,
        profit: item.sales - item.fees - item.shipping
      };
    });
  } catch (e) {
    Logger.log('期間別売上データ取得エラー: ' + e.toString());
    throw new Error('期間別売上データの取得に失敗しました: ' + e.message);
  }
}

/**
 * カテゴリ別の売上データを取得する
 * @param {string} startDate - 開始日 (YYYY-MM-DD形式)
 * @param {string} endDate - 終了日 (YYYY-MM-DD形式)
 * @returns {Array} カテゴリ別売上データの配列
 */
function getCategorySales(startDate = null, endDate = null) {
  try {
    Logger.log('getCategorySales 開始: ' + startDate + ' ～ ' + endDate);
    
    const properties = PropertiesService.getScriptProperties();
    const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
    if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
    
    const ss = SpreadsheetApp.openById(ssId);
    const salesSheet = ss.getSheetByName('販売管理');
    const productSheet = ss.getSheetByName('商品マスタ');
    
    if (!salesSheet || !productSheet) throw new Error('必要なシートが存在しません');
    
    // データ取得
    const salesData = salesSheet.getDataRange().getValues();
    const productData = productSheet.getDataRange().getValues();
    
    // ヘッダー行を除外
    const salesRows = salesData.slice(1);
    const productRows = productData.slice(1);
    
    const salesHeaders = salesData[0];
    const productHeaders = productData[0];
    
    const saleDateIdx = salesHeaders.indexOf('販売日');
    const productIdIdx = salesHeaders.indexOf('商品ID');
    const salePriceIdx = salesHeaders.indexOf('販売価格');
    const saleFeeIdx = salesHeaders.indexOf('販売手数料');
    const shippingFeeIdx = salesHeaders.indexOf('送料');
    const saleStatusIdx = salesHeaders.indexOf('取引ステータス');
    
    const productIdMasterIdx = productHeaders.indexOf('商品ID');
    const categoryIdx = productHeaders.indexOf('カテゴリ');
    
    Logger.log('インデックス情報 - 販売日:' + saleDateIdx + ', 商品ID:' + productIdIdx + ', カテゴリ:' + categoryIdx);
    
    // 商品IDからカテゴリを取得するマップを作成
    const productCategoryMap = {};
    for (const row of productRows) {
      const productId = row[productIdMasterIdx];
      const category = row[categoryIdx] || 'その他'; // カテゴリがない場合は「その他」
      productCategoryMap[productId] = category;
    }
    
    Logger.log('商品カテゴリマップ作成完了: ' + Object.keys(productCategoryMap).length + '件');
    
    // 日付フィルタリングの準備
    const start = startDate ? new Date(startDate) : new Date(new Date().getFullYear(), 0, 1); // 今年の1月1日
    const end = endDate ? new Date(endDate) : new Date(); // 今日
    
    Logger.log('日付フィルター: ' + start.toISOString() + ' ～ ' + end.toISOString());
    
    // カテゴリ別データを格納するオブジェクト
    const categoryData = {};
    
    // デフォルトカテゴリを追加（データがなくても表示するため）
    const defaultCategories = ['衣類', '家電', '本・雑誌', 'ホビー', 'コスメ', 'その他'];
    defaultCategories.forEach(category => {
      categoryData[category] = {
        category: category,
        sales: 0,
        fees: 0,
        shipping: 0,
        count: 0
      };
    });
    
    let processedRows = 0;
    let matchedRows = 0;
    
    for (const row of salesRows) {
      processedRows++;
      
      // 販売日の取得と検証
      const saleDate = row[saleDateIdx] instanceof Date ? row[saleDateIdx] : new Date(row[saleDateIdx]);
      
      if (isNaN(saleDate.getTime())) {
        Logger.log('無効な日付をスキップ: ' + row[saleDateIdx]);
        continue; // 無効な日付はスキップ
      }
      
      // 日付フィルタリング
      if (saleDate < start || saleDate > end) continue;
      
      // 取引完了の販売のみ集計
      if (row[saleStatusIdx] !== '取引完了') continue;
      
      const productId = row[productIdIdx];
      const category = productCategoryMap[productId] || 'その他';
      
      matchedRows++;
      
      // データ集計
      if (!categoryData[category]) {
        categoryData[category] = {
          category: category,
          sales: 0,
          fees: 0,
          shipping: 0,
          count: 0
        };
      }
      
      categoryData[category].sales += Number(row[salePriceIdx]) || 0;
      categoryData[category].fees += Number(row[saleFeeIdx]) || 0;
      categoryData[category].shipping += Number(row[shippingFeeIdx]) || 0;
      categoryData[category].count++;
    }
    
    Logger.log('処理行数: ' + processedRows + '行中' + matchedRows + '行が条件に一致');
    
    // 配列に変換して返す
    const result = Object.values(categoryData).map(item => {
      return {
        ...item,
        profit: item.sales - item.fees - item.shipping
      };
    });
    
    // 売上が0のカテゴリを除外
    const filteredResult = result.filter(item => item.sales > 0 || item.category === 'その他');
    
    Logger.log('カテゴリ別集計結果: ' + filteredResult.length + '件のカテゴリ');
    
    // 売上金額でソート
    filteredResult.sort((a, b) => b.sales - a.sales);
    
    return filteredResult;
  } catch (e) {
    Logger.log('カテゴリ別売上データ取得エラー: ' + e.toString());
    throw new Error('カテゴリ別売上データの取得に失敗しました: ' + e.message);
  }
}

/**
 * 商品別のパフォーマンスデータを取得する
 * @param {string} startDate - 開始日 (YYYY-MM-DD形式)
 * @param {string} endDate - 終了日 (YYYY-MM-DD形式)
 * @param {number} limit - 取得する商品数の上限
 * @returns {Array} 商品別パフォーマンスデータの配列
 */
function getProductPerformance(startDate = null, endDate = null, limit = 10) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const ssId = properties.getProperty('MASTER_SPREADSHEET_ID');
    if (!ssId) throw new Error('マスタースプレッドシートが未作成です');
    
    const ss = SpreadsheetApp.openById(ssId);
    const salesSheet = ss.getSheetByName('販売管理');
    const productSheet = ss.getSheetByName('商品マスタ');
    const purchaseSheet = ss.getSheetByName('仕入管理');
    
    if (!salesSheet || !productSheet || !purchaseSheet) throw new Error('必要なシートが存在しません');
    
    // データ取得
    const salesData = salesSheet.getDataRange().getValues();
    const productPerformanceData = productSheet.getDataRange().getValues();
    const purchaseData = purchaseSheet.getDataRange().getValues();
    
    // ヘッダー行を除外
    const salesRows = salesData.slice(1);
    const productRows = productPerformanceData.slice(1);
    const purchaseRows = purchaseData.slice(1);
    
    const salesHeaders = salesData[0];
    const productHeaders = productPerformanceData[0];
    const purchaseHeaders = purchaseData[0];
    
    const saleDateIdx = salesHeaders.indexOf('販売日');
    const saleProductIdIdx = salesHeaders.indexOf('商品ID');
    const salePriceIdx = salesHeaders.indexOf('販売価格');
    const saleFeeIdx = salesHeaders.indexOf('販売手数料');
    const shippingFeeIdx = salesHeaders.indexOf('送料');
    const saleStatusIdx = salesHeaders.indexOf('取引ステータス');
    
    const productIdMasterIdx = productHeaders.indexOf('商品ID');
    const productNameIdx = productHeaders.indexOf('商品名');
    
    const purchaseProductIdIdx = purchaseHeaders.indexOf('商品ID');
    const purchasePriceIdx = purchaseHeaders.indexOf('仕入価格');
    const purchaseStatusIdx = purchaseHeaders.indexOf('ステータス');
    
    // 商品IDから商品名を取得するマップを作成
    const productNameMap = {};
    for (const row of productRows) {
      const productId = row[productIdMasterIdx];
      const productName = row[productNameIdx];
      productNameMap[productId] = productName;
    }
    
    // 商品IDごとの仕入れコストを計算するマップを作成
    const productCostMap = {};
    for (const row of purchaseRows) {
      if (row[purchaseStatusIdx] !== '完了') continue;
      
      const productId = row[purchaseProductIdIdx];
      const cost = Number(row[purchasePriceIdx]) || 0;
      
      if (!productCostMap[productId]) {
        productCostMap[productId] = {
          totalCost: 0,
          count: 0
        };
      }
      
      productCostMap[productId].totalCost += cost;
      productCostMap[productId].count++;
    }
    
    // 日付フィルタリングの準備
    const start = startDate ? new Date(startDate) : new Date(new Date().getFullYear(), 0, 1); // 今年の1月1日
    const end = endDate ? new Date(endDate) : new Date(); // 今日
    
    // 商品別データを格納するオブジェクト
    const productPerfData = {};
    
    for (const row of salesRows) {
      // 販売日の取得と検証
      const saleDate = row[saleDateIdx] instanceof Date ? row[saleDateIdx] : new Date(row[saleDateIdx]);
      
      if (isNaN(saleDate.getTime())) continue; // 無効な日付はスキップ
      
      // 日付フィルタリング
      if (saleDate < start || saleDate > end) continue;
      
      // 取引完了の販売のみ集計
      if (row[saleStatusIdx] !== '取引完了') continue;
      
      const productId = row[saleProductIdIdx];
      const productName = productNameMap[productId] || '不明な商品';
      
      // データ集計
      if (!productPerfData[productId]) {
        productPerfData[productId] = {
          productId: productId,
          productName: productName,
          sales: 0,
          fees: 0,
          shipping: 0,
          count: 0,
          cost: productCostMap[productId] ? productCostMap[productId].totalCost : 0
        };
      }
      
      productPerfData[productId].sales += Number(row[salePriceIdx]) || 0;
      productPerfData[productId].fees += Number(row[saleFeeIdx]) || 0;
      productPerfData[productId].shipping += Number(row[shippingFeeIdx]) || 0;
      productPerfData[productId].count++;
    }
    
    // 利益計算して配列に変換
    const result = Object.values(productPerfData).map(item => {
      const grossProfit = item.sales - item.fees - item.shipping;
      const netProfit = grossProfit - item.cost;
      const roi = item.cost > 0 ? (netProfit / item.cost) * 100 : 0;
      
      return {
        ...item,
        grossProfit: grossProfit,
        netProfit: netProfit,
        roi: roi.toFixed(2),
        averageSalePrice: item.count > 0 ? (item.sales / item.count).toFixed(2) : 0
      };
    });
    
    // 利益の高い順にソートして上位N件を返す
    return result
      .sort((a, b) => b.netProfit - a.netProfit)
      .slice(0, limit);
  } catch (e) {
    Logger.log('商品別パフォーマンスデータ取得エラー: ' + e.toString());
    throw new Error('商品別パフォーマンスデータの取得に失敗しました: ' + e.message);
  }
}

/**
 * 財務データをCSV形式で取得する
 * @param {string} reportType - レポートタイプ ('basic', 'period', 'category', 'product')
 * @param {string} startDate - 開始日 (YYYY-MM-DD形式)
 * @param {string} endDate - 終了日 (YYYY-MM-DD形式)
 * @returns {string} CSV形式のデータ
 */
function getFinanceDataAsCsv(reportType, startDate = null, endDate = null) {
  try {
    let data;
    let headers;
    
    switch (reportType) {
      case 'basic':
        data = [getFinanceReport()];
        headers = ['総売上', '手数料合計', '送料合計', '粗利益', '仕入コスト合計', '純利益', '利益率(%)', '完了販売数', '保留販売数', '仕入数'];
        return convertToCsv(data, [
          'totalSales', 'totalFees', 'totalShipping', 'grossProfit', 
          'totalPurchaseCost', 'netProfit', 'profitMargin', 
          'completedSales', 'pendingSales', 'totalPurchaseCount'
        ], headers);
        
      case 'period':
        data = getPeriodSales(startDate, endDate);
        headers = ['期間', '売上', '手数料', '送料', '販売数', '利益'];
        return convertToCsv(data, ['period', 'sales', 'fees', 'shipping', 'count', 'profit'], headers);
        
      case 'category':
        data = getCategorySales(startDate, endDate);
        headers = ['カテゴリ', '売上', '手数料', '送料', '販売数', '利益'];
        return convertToCsv(data, ['category', 'sales', 'fees', 'shipping', 'count', 'profit'], headers);
        
      case 'product':
        data = getProductPerformance(startDate, endDate, 100); // CSVでは上限を増やす
        headers = ['商品ID', '商品名', '売上', '手数料', '送料', '販売数', '仕入コスト', '粗利益', '純利益', 'ROI(%)', '平均販売価格'];
        return convertToCsv(data, [
          'productId', 'productName', 'sales', 'fees', 'shipping', 'count', 
          'cost', 'grossProfit', 'netProfit', 'roi', 'averageSalePrice'
        ], headers);
        
      default:
        throw new Error('不明なレポートタイプです');
    }
  } catch (e) {
    Logger.log('CSVデータ取得エラー: ' + e.toString());
    throw new Error('CSVデータの取得に失敗しました: ' + e.message);
  }
}

/**
 * オブジェクトの配列をCSV形式の文字列に変換する
 * @param {Array} data - オブジェクトの配列
 * @param {Array} fields - 出力するフィールド名の配列
 * @param {Array} headers - CSVヘッダー行の配列
 * @returns {string} CSV形式の文字列
 */
function convertToCsv(data, fields, headers) {
  // ヘッダー行
  let csv = headers.join(',') + '\n';
  
  // データ行
  for (const item of data) {
    const row = fields.map(field => {
      const value = item[field];
      // 数値はそのまま、文字列はダブルクォートで囲む
      return typeof value === 'string' ? `"${value}"` : value;
    });
    csv += row.join(',') + '\n';
  }
  
  return csv;
} 