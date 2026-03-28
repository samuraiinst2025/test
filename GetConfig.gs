function getConfig_(key) {
const sh = SpreadsheetApp.getActive().getSheetByName('config');
const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
for (const [k, v] of values) if (k === key) return v;
return '';
}

// --- 新しく追加・修正するメイン処理のイメージ ---
function mainProcess() {
  // 1. 設定を読み込む
  const method = getConfig_('Method'); // 「月平均」「最大値」「最小値」のいずれかが入る
  const fromDate = getConfig_('from_date');
  const toDate = getConfig_('to_date');

  // 2. freeeなどからデータを取得する（ここは既存の取得処理を想定）
  // const rawData = fetchFreeeData(fromDate, toDate); 
  
  // 3. Methodに合わせて計算方法を変える
  let result;
  
  // ここで「Method」の中身を見て処理を分岐させます
  switch (method) {
    case '月平均':
      result = calculateAverage(rawData); // 平均を出す自作関数など
      break;
      
    case '最大値':
      result = calculateMax(rawData); // 最大値を出す自作関数など
      break;
      
    case '最小値':
      result = calculateMin(rawData); // 最小値を出す自作関数など
      break;
      
    default:
      console.log('Methodが正しく選択されていません');
      return;
  }

  // 4. 結果をシートに書き出す
  // outputToSheet(result);
}