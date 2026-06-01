/**
 * 「未充足管理」から「拠点移動」「早出」「延長」の依頼を、
 * それぞれの仕様に合わせてログシートに転記する。
 */
function logSpecialShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("未充足管理");
  const hubLogSheet = ss.getSheetByName("拠点移動ログ");
  const eoLogSheet = ss.getSheetByName("早出・延長");

  if (!sourceSheet || !hubLogSheet || !eoLogSheet) {
    console.error("エラー: 必須シートが見つかりません。");
    return;
  }

  // --- ヘッダー定義 ---
  const hubHeader = ["チェック", "日付", "シフトカテゴリー", "進捗・ステータス", "拠点", "開始時間", "終了時間", "依頼した医師", "最終担当医", "交渉結果", "ソート用日付", "ユニークキー"];
  const eoHeader = ["チェック", "日付", "シフトカテゴリー", "進捗・ステータス", "拠点", "開始時間", "終了時間", "担当医師", "コメント", "ソート用日付", "ユニークキー"];
  
  const SOURCE_START_ROW = 7;
  const TARGET_DATE = new Date("2025-06-01");
  TARGET_DATE.setHours(0, 0, 0, 0);

  const hubKeys = new Set();
  if (hubLogSheet.getLastRow() > 1) {
    hubLogSheet.getRange(2, 12, hubLogSheet.getLastRow() - 1, 1).getValues().forEach(row => { if (row[0]) hubKeys.add(row[0]) });
  }
  const eoKeys = new Set();
  if (eoLogSheet.getLastRow() > 1) {
    eoLogSheet.getRange(2, 11, eoLogSheet.getLastRow() - 1, 1).getValues().forEach(row => { if (row[0]) eoKeys.add(row[0]) });
  }

  const lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow < SOURCE_START_ROW) return;

  const sourceRange = sourceSheet.getRange(SOURCE_START_ROW, 2, lastSourceRow - SOURCE_START_ROW + 1, 19);
  const sourceValues = sourceRange.getValues();
  
  const hubChangesToTransfer = [];
  const eoChangesToTransfer = [];

  sourceValues.forEach(row => {
    const c_category = row[1];
    if (c_category !== "拠点移動" && c_category !== "早出" && c_category !== "延長") return;

    const g_location = row[5];
    const m_status = row[11];
    const r_comment = row[16];
    const t_date = row[18];
    if (!g_location || !(t_date instanceof Date) || t_date.getTime() < TARGET_DATE.getTime()) return;

    const h_start_raw = row[6];
    const i_end_raw = row[7];
    const j_doctor = row[8];
    
    const startTimeStr = (h_start_raw instanceof Date) ? Utilities.formatDate(h_start_raw, "JST", "HH:mm") : h_start_raw.toString();
    const endTimeStr = (i_end_raw instanceof Date) ? Utilities.formatDate(i_end_raw, "JST", "HH:mm") : i_end_raw.toString();
    const dateKey = Utilities.formatDate(t_date, "JST", "yyyy-MM-dd");
    const uniqueKey = `${dateKey}_${g_location}_${startTimeStr}_${endTimeStr}`;

    if (c_category === "拠点移動") {
      if (hubKeys.has(uniqueKey)) return;
      const newRowData = [ false, row[0], c_category, m_status, g_location, startTimeStr, endTimeStr, j_doctor, "", "", t_date, uniqueKey ];
      hubChangesToTransfer.push(newRowData);
      hubKeys.add(uniqueKey);
    } else { // 早出 or 延長
      if (eoKeys.has(uniqueKey)) return;
      const isRejected = (m_status === "不可");
      const newRowData = [ isRejected, row[0], c_category, m_status, g_location, startTimeStr, endTimeStr, j_doctor, r_comment, t_date, uniqueKey ];
      eoChangesToTransfer.push(newRowData);
      eoKeys.add(uniqueKey);
    }
  });
  
  if (hubLogSheet.getLastRow() < 1) {
    hubLogSheet.getRange(1, 1, 1, hubHeader.length).setValues([hubHeader]);
  }
  if (hubChangesToTransfer.length > 0) {
    const startRow = hubLogSheet.getLastRow() + 1;
    hubLogSheet.getRange(startRow, 1, hubChangesToTransfer.length, hubHeader.length).setValues(hubChangesToTransfer);
    hubLogSheet.getRange(startRow, 1, hubChangesToTransfer.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }

  if (eoLogSheet.getLastRow() < 1) {
    eoLogSheet.getRange(1, 1, 1, eoHeader.length).setValues([eoHeader]);
  }
  if (eoChangesToTransfer.length > 0) {
    const startRow = eoLogSheet.getLastRow() + 1;
    eoLogSheet.getRange(startRow, 1, eoChangesToTransfer.length, eoHeader.length).setValues(eoChangesToTransfer);
    eoLogSheet.getRange(startRow, 1, eoChangesToTransfer.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
}
/**
 * 「未充足管理」の最新状況を各ログシートに反映する。
 * - 交渉が成立・失敗した場合、チェックボックスを自動でONにする
 */
function updateNegotiationResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("未充足管理");
  
  const logSheets = [
    { name: "拠点移動ログ", isHub: true },
    { name: "早出・延長", isHub: false }
  ];

  const lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow < 7) return;
  
  const sourceRange = sourceSheet.getRange("G7:T" + lastSourceRow);
  const sourceDisplayValues = sourceRange.getDisplayValues();
  const statusMap = {};
  sourceDisplayValues.forEach(row => {
    const location = row[0], startTime = row[1], endTime = row[2], 
          finalDoctor = row[3], status = row[6], comment = row[11], dateStr = row[13];
    if (dateStr && location && startTime && endTime) {
      const dateKey = dateStr.replace(/\//g, '-');
      const uniqueKey = `${dateKey}_${location}_${startTime}_${endTime}`;
      statusMap[uniqueKey] = { status: status, finalDoctor: finalDoctor, comment: comment };
    }
  });

  logSheets.forEach(sheetInfo => {
    const logSheet = ss.getSheetByName(sheetInfo.name);
    if (!logSheet) return;

    const lastLogRaw = logSheet.getLastRow();
    if (lastLogRaw < 2) return;
    
    const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const keyCol = headers.indexOf("ユニークキー") + 1;
    const resultCol = headers.indexOf("交渉結果") + 1;
    const doctorCol = headers.indexOf("最終担当医") + 1;
    const statusCol = headers.indexOf("進捗・ステータス") + 1;
    const commentCol = headers.indexOf("コメント") + 1;
    const checkCol = headers.indexOf("チェック") + 1;
    if (!keyCol || !checkCol) return; // 必須列がなければスキップ

    const logRange = logSheet.getRange(2, 1, lastLogRaw - 1, Math.max(keyCol, commentCol, resultCol, doctorCol, statusCol));
    const logValues = logRange.getValues();
    const updates = {};

    logValues.forEach((row, index) => {
      const rowNum = index + 2;
      const uniqueKey = row[keyCol - 1];
      const isChecked = row[checkCol - 1];
      if (!uniqueKey || !statusMap[uniqueKey] || isChecked) return; // キーがない or 元データがない or 既にチェック済みならスキップ

      const currentData = statusMap[uniqueKey];
      if (!updates[rowNum]) updates[rowNum] = {};
      
      // ステータスを常に最新に更新
      if (statusCol && row[statusCol - 1] !== currentData.status) {
        updates[rowNum][statusCol] = currentData.status;
      }
      
      let negotiationConcluded = false;
      if (currentData.status === "確定" || currentData.status === "充足" || currentData.status === "応募有（充足）") {
        negotiationConcluded = true;
        if (sheetInfo.isHub && resultCol) updates[rowNum][resultCol] = "成立";
      } else if (currentData.status === "不可") {
        negotiationConcluded = true;
        if (sheetInfo.isHub && resultCol) updates[rowNum][resultCol] = "失敗";
      }

      // 交渉が妥結した場合の共通処理
      if (negotiationConcluded) {
        updates[rowNum][checkCol] = true; // チェックを入れる
        
        // シート別の詳細更新
        if (sheetInfo.isHub) {
          if (doctorCol) updates[rowNum][doctorCol] = currentData.finalDoctor;
        } else {
          if (commentCol) updates[rowNum][commentCol] = currentData.comment;
        }
      }
    });

    for (const rowNum in updates) {
      for (const colNum in updates[rowNum]) {
        logSheet.getRange(parseInt(rowNum), parseInt(colNum)).setValue(updates[rowNum][colNum]);
      }
    }
  });
}