function debugAbsenceReportTimeError() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('貼付用'); // 不在抽出の元データ
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();

  Logger.log("=== 時間データの型エラー調査開始 ===");
  let errorCount = 0;

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const docName = row[0];
    const startTime = row[15]; // P列
    const endTime = row[19];   // T列

    // startTimeの検証
    if (startTime !== "" && startTime != null && typeof startTime !== 'string') {
      Logger.log(`行 ${i + 1} (${docName}): 開始時間 [${startTime}] は文字列ではありません。型: ${typeof startTime}`);
      errorCount++;
    }

    // endTimeの検証
    if (endTime !== "" && endTime != null && typeof endTime !== 'string') {
      Logger.log(`行 ${i + 1} (${docName}): 終了時間 [${endTime}] は文字列ではありません。型: ${typeof endTime}`);
      errorCount++;
    }
  }

  Logger.log(`=== 調査完了 (エラー原因セル: ${errorCount}件) ===`);
  if(errorCount > 0) {
    Logger.log("【結論】スプレッドシートから時間を取得した際、GASが自動的に「文字列」ではなく「Date型」や「数値」に変換してしまったセルが存在します。古い関数がこれを受け取ってクラッシュしています。");
  }
}