/**
 * 実績データをIDから名称に変換し、F列(amount)を金額として集計。
 * 出力形式: | 日付 | 勘定科目 | 取引先 | 金額 |
 */
function aggregateActualsForBudget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- 1. シートの定義 ---
  const configSheet = ss.getSheetByName("Config");
  const accountMasterSheet = ss.getSheetByName("master_account_items");
  const partnerMasterSheet = ss.getSheetByName("master_partners");
  const outputSheet = ss.getSheetByName("実績分析_予算用");
  const sourceSheet = ss.getSheetByName("fact_journal_lines");
  
  if (!configSheet || !accountMasterSheet || !partnerMasterSheet || !outputSheet || !sourceSheet) {
    SpreadsheetApp.getUi().alert("必要なシートが見つかりません。");
    return;
  }

  // --- 2. マスタデータの読み込み（ID照合用） ---
  const getMasterMap = (sheet) => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return new Map();
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const map = new Map();
    data.forEach(row => {
      if (row[0] !== "") {
        const id = String(row[0]).trim();
        map.set(id, row[1]); // A列がID、B列が名称
      }
    });
    return map;
  };

  const accountMap = getMasterMap(accountMasterSheet);
  const partnerMap = getMasterMap(partnerMasterSheet);

  // --- 3. 設定の読み込み ---
  const method = configSheet.getRange("B1").getValue();

  // --- 4. 実績データの取得（fact_journal_lines） ---
  const sourceData = sourceSheet.getDataRange().getValues();
  sourceData.shift(); // ヘッダー除去

  /**
   * 列インデックス（0から開始）
   * D列: issue_date      => 3
   * E列: account_item_id => 4
   * F列: amount          => 5  <-- ここを集計
   * G列: partner_id      => 6
   */
  const COL = { DATE: 3, ACCOUNT_ID: 4, AMOUNT: 5, PARTNER_ID: 6 };

  const monthlyAggregated = {};
  const uniqueMonths = new Set();

  // --- 5. 集計処理 ---
  sourceData.forEach(row => {
    const rawAccountId = String(row[COL.ACCOUNT_ID] || "").trim();
    const rawPartnerId = String(row[COL.PARTNER_ID] || "").trim();
    
    // F列(index 5)を金額として取得
    const amount = parseFloat(row[COL.AMOUNT]) || 0;
    
    const accountName = accountMap.get(rawAccountId);
    const partnerName = partnerMap.get(rawPartnerId);

    // マスタに一致するIDがある場合のみ集計
    if (accountName && partnerName) {
      const date = new Date(row[COL.DATE]);
      if (!isNaN(date.getTime())) {
        const monthStr = Utilities.formatDate(date, "JST", "yyyy-MM");
        uniqueMonths.add(monthStr);
        
        const key = `${monthStr}|${accountName}|${partnerName}`;
        monthlyAggregated[key] = (monthlyAggregated[key] || 0) + amount;
      }
    }
  });

  // 対象期間の特定
  const sortedMonths = Array.from(uniqueMonths).sort();
  const periodStr = sortedMonths.length > 0 
    ? (sortedMonths.length === 1 ? sortedMonths[0] : `${sortedMonths[0]} ~ ${sortedMonths[sortedMonths.length-1]}`)
    : "期間不明";
  
  const monthCount = uniqueMonths.size || 1;

  // --- 6. 最終計算（月平均・最大・最小） ---
  const finalGrouping = {};
  for (let key in monthlyAggregated) {
    const [, accName, partName] = key.split("|");
    const groupKey = `${accName}|${partName}`;
    if (!finalGrouping[groupKey]) finalGrouping[groupKey] = [];
    finalGrouping[groupKey].push(monthlyAggregated[key]);
  }

  const results = [];
  for (let groupKey in finalGrouping) {
    const [accName, partName] = groupKey.split("|");
    const values = finalGrouping[groupKey];
    let finalAmount = 0;

    if (method === "月平均") {
      finalAmount = values.reduce((sum, v) => sum + v, 0) / monthCount;
    } else if (method === "最大値") {
      finalAmount = Math.max(...values);
    } else if (method === "最小値") {
      finalAmount = Math.min(...values);
    }
    
    results.push([periodStr, accName, partName, finalAmount]);
  }

  // --- 7. 出力 ---
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, 1, 4).setValues([["日付", "勘定科目", "取引先", "金額"]]);
  
  if (results.length > 0) {
    // 勘定科目 > 取引先 の順でソート
    results.sort((a, b) => a[1].localeCompare(b[1], 'ja') || a[2].localeCompare(b[2], 'ja'));
    outputSheet.getRange(2, 1, results.length, 4).setValues(results);
    
    // 見栄えの調整
    outputSheet.getRange(2, 4, results.length, 1).setNumberFormat("#,##0"); // 金額カンマ区切り
    outputSheet.getRange(2, 1, results.length, 1).setHorizontalAlignment("center"); // 日付を中央に
  }

  SpreadsheetApp.getUi().alert("集計完了！\n集計対象: F列(amount)\n手法: " + method);
}