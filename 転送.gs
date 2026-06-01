/**
 * 「拠点移動ログ」「早出・延長」シートでチェックボックスがONの行を、
 * 別のスプレッドシート（管理台帳）に転記し、転記済みのマークを付ける。
 */
function transferCheckedLogsToMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubLogSheet = ss.getSheetByName("拠点移動ログ");
  const eoLogSheet = ss.getSheetByName("早出・延長");

  const destinationSpreadsheetId = "1Rjc87NrTImWJZ_odHbT7DNN6jtIUpbHdcNomcVk430U";
  const destinationSheetName = "拠点移動";
  
  const destSS = SpreadsheetApp.openById(destinationSpreadsheetId);
  const destSheet = destSS.getSheetByName(destinationSheetName);

  if (!hubLogSheet || !eoLogSheet || !destSheet) {
    console.error("エラー: 必須シートが見つかりません。");
    return;
  }

  const existingKeys = new Set();
  const lastDestRow = destSheet.getLastRow();
  if (lastDestRow > 1) {
    destSheet.getRange(2, 20, lastDestRow - 1, 1).getValues()
      .forEach(row => { if (row[0]) existingKeys.add(row[0]) });
  }

  const rowsToTransfer = [];
  const hubRowsToMark = [];
  const eoRowsToMark = [];

  // --- 2. 「拠点移動ログ」シートを処理 ---
  const hubLastRow = hubLogSheet.getLastRow();
  if (hubLastRow > 1) {
    const hubValues = hubLogSheet.getRange(2, 1, hubLastRow - 1, 12).getValues(); // A:L
    hubValues.forEach((row, index) => {
      const isChecked = row[0];
      const uniqueKey = row[11];
      if (isChecked && !existingKeys.has(uniqueKey)) {
        const newRow = new Array(20).fill('');
        newRow[1] = row[10]; // B列: ソート用日付 (元K)
        newRow[2] = row[2];  // C列: シフトカテゴリー (元C)
        newRow[6] = row[4];  // G列: 拠点 (元E)
        newRow[7] = row[5];  // H列: 開始時間 (元F)
        newRow[8] = row[6];  // I列: 終了時間 (元G)
        newRow[9] = row[7];  // J列: ★依頼した医師★ (元H)
        newRow[12] = row[3]; // M列: 進捗・ステータス (元D)
        newRow[17] = row[9]; // R列: 交渉結果 (元J)
        newRow[19] = uniqueKey; // T列: ユニークキー
        
        rowsToTransfer.push(newRow);
        existingKeys.add(uniqueKey);
        hubRowsToMark.push(index + 2); // 転記済みにする行番号を記録
      }
    });
  }

  // --- 3. 「早出・延長」シートを処理 ---
  const eoLastRow = eoLogSheet.getLastRow();
  if (eoLastRow > 1) {
    const eoValues = eoLogSheet.getRange(2, 1, eoLastRow - 1, 11).getValues(); // A:K
    eoValues.forEach((row, index) => {
      const isChecked = row[0];
      const uniqueKey = row[10];
      if (isChecked && !existingKeys.has(uniqueKey)) {
        const newRow = new Array(20).fill('');
        newRow[1] = row[9];  // B列: ソート用日付 (元J)
        newRow[2] = row[2];  // C列: シフトカテゴリー (元C)
        newRow[6] = row[4];  // G列: 拠点 (元E)
        newRow[7] = row[5];  // H列: 開始時間 (元F)
        newRow[8] = row[6];  // I列: 終了時間 (元G)
        newRow[9] = row[7];  // J列: 担当医師 (元H)
        newRow[12] = row[3]; // M列: 進捗・ステータス (元D)
        newRow[17] = row[8]; // R列: コメント (元I)
        newRow[19] = uniqueKey; // T列: ユニークキー
        
        rowsToTransfer.push(newRow);
        existingKeys.add(uniqueKey);
        eoRowsToMark.push(index + 2); // 転記済みにする行番号を記録
      }
    });
  }

  // --- 4. データを転記先に書き込む ---
  if (rowsToTransfer.length > 0) {
    const findLastRowByCol = (sheet, col) => {
      const colValues = sheet.getRange(1, col, sheet.getMaxRows(), 1).getValues();
      for (let i = colValues.length - 1; i >= 0; i--) {
        if (colValues[i][0] !== '') return i + 1;
      }
      return 0;
    };
    
    const startRow = findLastRowByCol(destSheet, 10) + 1;
    destSheet.getRange(startRow, 1, rowsToTransfer.length, 20).setValues(rowsToTransfer);
    console.log(`${rowsToTransfer.length} 件のデータを管理台帳に転記しました。`);

    // --- 5. 転記元のログに「済」を記録 ---
    const STATUS_COLUMN = 13; // M列
    hubRowsToMark.forEach(rowNum => {
      hubLogSheet.getRange(rowNum, STATUS_COLUMN).setValue("済");
    });
    eoRowsToMark.forEach(rowNum => {
      eoLogSheet.getRange(rowNum, STATUS_COLUMN).setValue("済");
    });
    console.log(`転記済みのマークを ${hubRowsToMark.length + eoRowsToMark.length} 件記録しました。`);
  }
}