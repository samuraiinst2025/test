/**
 * 指定されたフォルダ内の複数スプレッドシートから売上データを集計します。
 * 対象フォルダID: 1D_vd4qbIv-7pQx-OJo3sw48h93Fj7-KF
 * ファイル名規則: 「売上表_」で始まるもの
 */

/**
 * スプレッドシートを開いた時にメニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('売上集計')
    .addItem('集計を実行', 'aggregateSalesData')
    .addToUi();
}

function aggregateSalesData() {
  const folderId = '1D_vd4qbIv-7pQx-OJo3sw48h93Fj7-KF';
  const fileNamePrefix = '売上表_';
  
  // 1. ユーザー入力の取得
  const ui = SpreadsheetApp.getUi();
  const response = Browser.inputBox('集計期間の指定', '集計したい年月を入力してください（例: 2023/2, 2023/02）', Browser.Buttons.OK_CANCEL);
  
  if (response === 'cancel' || response === '') {
    return;
  }

  // 入力の正規化（yyyy/m または yyyy/mm）
  const dateParts = response.split('/');
  if (dateParts.length !== 2) {
    Browser.msgBox('形式が正しくありません。「yyyy/m」の形式で入力してください。');
    return;
  }
  
  const targetYear = parseInt(dateParts[0], 10);
  const targetMonth = parseInt(dateParts[1], 10);
  
  if (isNaN(targetYear) || isNaN(targetMonth)) {
    Browser.msgBox('数値が正しくありません。');
    return;
  }

  // 2. Google Driveフォルダからファイル取得
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const summaryData = [];
  let hasMatch = false;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // 「売上表_」で始まるスプレッドシートのみ対象
    if (fileName.indexOf(fileNamePrefix) === 0 && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      const storeName = fileName.replace(fileNamePrefix, '');
      const ss = SpreadsheetApp.open(file);
      const sheet = ss.getSheets()[0]; // 1番左のシート
      const lastRow = sheet.getLastRow();
      
      if (lastRow < 2) {
        summaryData.push([storeName, 0]);
        continue;
      }
      
      const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A列(1)からE列(5)まで取得
      let totalAmount = 0;
      let storeHasMatch = false;
      
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rawDate = row[0]; // A列: 日付
        const amount = row[4];  // E列: 金額
        
        let date;
        if (rawDate instanceof Date) {
          date = rawDate;
        } else if (typeof rawDate === 'string' && rawDate !== '') {
          date = new Date(rawDate);
        } else {
          continue;
        }
        
        // 年月が一致するか確認
        if (date.getFullYear() === targetYear && (date.getMonth() + 1) === targetMonth) {
          hasMatch = true;
          storeHasMatch = true;
          if (!isNaN(amount)) {
            totalAmount += Number(amount);
          }
        }
      }
      
      if (storeHasMatch) {
        summaryData.push([storeName, totalAmount]);
      }
    }
  }

  // 3. 結果の書き出し
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = activeSS.getSheets()[0]; // 1番左のシート
  
  outputSheet.clear(); // シートのクリア
  
  if (!hasMatch) {
    outputSheet.getRange(1, 1).setValue('集計対象がありません');
    Browser.msgBox('対象期間のデータが見つかりませんでした。');
    return;
  }

  // ヘッダーの書き込み
  const outputValues = [['店舗名', '金額'], ...summaryData];
  
  outputSheet.getRange(1, 1, outputValues.length, 2).setValues(outputValues);
  
  Browser.msgBox('集計が完了しました。');
}
