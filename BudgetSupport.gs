// ==========================================
// 定数定義：シート名や列番号の設定
// ==========================================
const SS_CONFIG = {
  SHEET: {
    CONFIG: "config",
    ACCOUNT: "master_account_items",
    PARTNER: "master_partners",
    OUTPUT: "実績分析_予算用",
    SOURCE: "fact_journal_lines"
  },
  SOURCE_COL: {
    DATE: 3,       // D列
    ACCOUNT_ID: 4, // E列
    AMOUNT: 5,     // F列
    PARTNER_ID: 6  // G列
  }
};

/**
 * Configシートから現在の設定値を取得する
 */
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.SHEET.CONFIG);
  const values = sheet.getRange("B2:B6").getValues();
  return {
    companyId: values[0][0],
    fromDate:  values[1][0],
    toDate:    values[2][0],
    fiscalYear:values[3][0],
    method:    values[4][0] // B6セル
  };
}

/**
 * ボタン実行用のメイン関数
 */
function onExecuteMain() {
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings();

  if (!settings.method) {
    ui.alert("Methodが選択されていません。");
    return;
  }

  const confirmMsg = `集計を開始しますか？\n\n期間: ${Utilities.formatDate(new Date(settings.fromDate), "JST", "yyyy/MM/dd")} ～ ${Utilities.formatDate(new Date(settings.toDate), "JST", "yyyy/MM/dd")}\n手法: ${settings.method}`;
  if (ui.alert("実行確認", confirmMsg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  try {
    // データ取得処理（必要に応じて追加）
    // fetchFreeeData(settings); 

    // 実績集計の実行
    aggregateActualsForBudget(settings);

    ui.alert("すべての処理が正常に完了しました。");
  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
  }
}

/**
 * 集計コアロジック
 */
function aggregateActualsForBudget(settings = getSettings()) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 各シートとマスタMapの準備
  const accountMap = createMasterMap(ss.getSheetByName(SS_CONFIG.SHEET.ACCOUNT));
  const partnerMap = createMasterMap(ss.getSheetByName(SS_CONFIG.SHEET.PARTNER));
  const sourceSheet = ss.getSheetByName(SS_CONFIG.SHEET.SOURCE);
  const sourceData = sourceSheet.getDataRange().getValues();
  sourceData.shift(); // ヘッダー除去

  const monthlyAggregated = {};
  const uniqueMonths = new Set();
  const COL = SS_CONFIG.SOURCE_COL;

  // 2. 月別・IDペア別の一次集計
  sourceData.forEach(row => {
    const accId = String(row[COL.ACCOUNT_ID] || "").trim();
    const partId = String(row[COL.PARTNER_ID] || "").trim();
    const amount = parseFloat(row[COL.AMOUNT]) || 0;
    const date = new Date(row[COL.DATE]);
    
    if (isNaN(date.getTime()) || !accId || !partId) return;

    // 両方のマスタに存在するデータのみ対象
    if (accountMap.has(accId) && partnerMap.has(partId)) {
      const monthStr = Utilities.formatDate(date, "JST", "yyyy-MM");
      uniqueMonths.add(monthStr);
      const key = `${monthStr}|${accId}|${partId}`;
      monthlyAggregated[key] = (monthlyAggregated[key] || 0) + amount;
    }
  });

  // 3. 期間ラベルと月数の計算
  const sortedMonths = Array.from(uniqueMonths).sort();
  const periodStr = sortedMonths.length > 0 
    ? (sortedMonths.length === 1 ? sortedMonths[0] : `${sortedMonths[0]} ～ ${sortedMonths[sortedMonths.length-1]}`)
    : "期間不明";
  const monthCount = uniqueMonths.size || 1;

  // 4. 手法(Method)に基づく最終計算
  const results = calculateFinalValues(monthlyAggregated, accountMap, partnerMap, settings.method, monthCount, periodStr);

  // 5. 出力
  writeToOutputSheet(ss.getSheetByName(SS_CONFIG.SHEET.OUTPUT), results, settings.method);
}

/**
 * マスタシートからMap(ID -> 名称)を作成
 */
function createMasterMap(sheet) {
  const lastRow = sheet.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  data.forEach(row => {
    const id = String(row[0]).trim();
    if (id && id !== "null") map.set(id, String(row[1]));
  });
  return map;
}

/**
 * 集計データに対してMethod(平均等)を適用する
 */
function calculateFinalValues(monthlyAggregated, accountMap, partnerMap, method, monthCount, periodStr) {
  const grouping = {};
  for (let key in monthlyAggregated) {
    const [, accId, partId] = key.split("|");
    const idPairKey = `${accId}|${partId}`;
    if (!grouping[idPairKey]) grouping[idPairKey] = [];
    grouping[idPairKey].push(monthlyAggregated[key]);
  }

  return Object.keys(grouping).map(idPairKey => {
    const [accId, partId] = idPairKey.split("|");
    const values = grouping[idPairKey];
    let finalAmount = 0;

    if (method === "月平均") {
      finalAmount = values.reduce((sum, v) => sum + v, 0) / monthCount;
    } else if (method === "最大値") {
      finalAmount = Math.max(...values);
    } else if (method === "最小値") {
      finalAmount = Math.min(...values);
    }

    return [periodStr, accountMap.get(accId), partnerMap.get(partId), finalAmount];
  });
}

/**
 * 結果をシートに出力する
 */
function writeToOutputSheet(sheet, results, method) {
  sheet.clearContents();
  const header = [["日付", "勘定科目", "取引先", `金額（${method}）`]];
  sheet.getRange(1, 1, 1, 4).setValues(header);
  
  if (results.length > 0) {
    // 勘定科目(B列) > 取引先(C列) でソート
    results.sort((a, b) => a[1].localeCompare(b[1], 'ja') || a[2].localeCompare(b[2], 'ja'));
    sheet.getRange(2, 1, results.length, 4).setValues(results);
    sheet.getRange(2, 4, results.length, 1).setNumberFormat("#,##0");
  }
}