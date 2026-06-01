function convertDateFormat() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("未充足管理");
  if (!sheet) {
    Logger.log("エラー: シート「未充足管理」が見つかりません。処理を終了します。");
    return;
  }

  const lastRow = sheet.getLastRow(); // 最終行を取得
  // B列とT列のデータを取得（B6から最終行まで）
  const range = sheet.getRange(6, 2, lastRow - 5, 19); // B列からT列まで取得 (B=2, T=20 なので 20-2+1 = 19列)
  const values = range.getValues();

  Logger.log(`処理開始: B6:T${lastRow} までのデータを取得しました。`);

  for (let i = 0; i < values.length; i++) {
    let rowNumber = i + 6; // 実際の行番号
    let cellValueB = values[i][0]; // B列の値 (インデックス0)
    let cellValueT = values[i][18]; // T列の値 (インデックス18)

    // T列に既に値がある場合はスキップ
    if (cellValueT) {
      Logger.log(`スキップ: T${rowNumber} には既に値 (${cellValueT}) があります。`);
      continue;
    }

    if (!cellValueB) {
      Logger.log(`スキップ: B${rowNumber} の値が空です`);
      continue;
    }

    if (cellValueB instanceof Date) {
      // Google Sheets の日付データを yyyy/MM/dd 形式に変換
      // 年の判定ロジックは元のスクリプトのままですが、日付オブジェクトから年を取得する方がより堅牢です。
      // 必要に応じて Utilities.formatDate(cellValueB, Session.getScriptTimeZone(), "yyyy") で年を取得して比較してください。
      let year = (rowNumber >= 3589) ? 2025 : 2024; // このロジックで問題ないか確認してください
      let formattedDate = Utilities.formatDate(cellValueB, Session.getScriptTimeZone(), "yyyy/MM/dd");

      // 年のロジックを日付オブジェクトに基づいて行う場合 (推奨)
      // let dateObjectYear = cellValueB.getFullYear();
      // let formattedDate = Utilities.formatDate(cellValueB, Session.getScriptTimeZone(), "yyyy/MM/dd");
      // Logger.log(`デバッグ: B${rowNumber} の年は ${dateObjectYear} です。`);


      sheet.getRange(rowNumber, 20).setValue(formattedDate); // T列（20列目）に書き出し
      Logger.log(`成功: B${rowNumber} (日付データ ${cellValueB}) → T${rowNumber} (${formattedDate})`);
    } else if (typeof cellValueB === "string") {
      let match = cellValueB.match(/(\d+)月(\d+)日/);
      if (match) {
        let year = (rowNumber >= 3589) ? 2025 : 2024; // このロジックで問題ないか確認してください
        let month = match[1].padStart(2, '0');
        let day = match[2].padStart(2, '0');
        let formattedDate = `${year}/${month}/${day}`;
        sheet.getRange(rowNumber, 20).setValue(formattedDate);
        Logger.log(`成功: B${rowNumber} (${cellValueB}) → T${rowNumber} (${formattedDate})`);
      } else {
        Logger.log(`未変換: B${rowNumber} (${cellValueB}) は日付形式ではありません。`);
      }
    } else {
      Logger.log(`スキップ: B${rowNumber} の値が不明な形式 (${cellValueB})`);
    }
  }

  Logger.log("処理完了: すべての行をチェックしました。");
}