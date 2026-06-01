/**
 * 「未充足管理」シートの既存データに対して、V列にユニークIDを初期設定する関数。
 * この関数は、新しい重複チェックロジック導入後、一度だけ実行してください。
 */
function initializeTargetSheetUniqueIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("未充足管理");

  if (!targetSheet) {
    Logger.log("エラー: シート「未充足管理」が見つかりません。");
    return;
  }

  const TARGET_DATA_START_ROW = 7;
  const UNIQUE_ID_COLUMN = 22; // V列

  const lastRow = targetSheet.getLastRow();
  if (lastRow < TARGET_DATA_START_ROW) {
    Logger.log("「未充足管理」シートに初期化対象のデータがありません。");
    return;
  }

  // B列(日付), G列(拠点), H列(開始時間), I列(終了時間), B列(ステータス) を読み込む
  const dataRange = targetSheet.getRange(TARGET_DATA_START_ROW, 2, lastRow - TARGET_DATA_START_ROW + 1, 9); // B列からI列まで
  const dataValues = dataRange.getDisplayValues(); // 表示されている値を取得

  const uniqueIdsToSet = [];

  dataValues.forEach(row => {
    const formattedDate = row[0]; // B列 (例: 6月22日(日))
    const currentStatus = row[1]; // B列 (画像ではB列が「新規応募待ち」などのステータスなので、ここから取得)
    const finalLocation = cleanLocation(row[5]); // G列
    const startTime = String(row[6]).trim(); // H列
    const endTime = String(row[7]).trim(); // I列

    // ユニークIDを生成（ステータスはB列の値を使用）
    const generatedUniqueId = `${formattedDate}_${finalLocation}_${startTime}_${endTime}_${currentStatus}`;
    uniqueIdsToSet.push([generatedUniqueId]);
  });

  if (uniqueIdsToSet.length > 0) {
    targetSheet.getRange(TARGET_DATA_START_ROW, UNIQUE_ID_COLUMN, uniqueIdsToSet.length, 1).setValues(uniqueIdsToSet);
    Logger.log(`「未充足管理」シートのV列に ${uniqueIdsToSet.length} 件のユニークIDを初期設定しました。`);
  } else {
    Logger.log("初期設定するユニークIDがありませんでした。");
  }
}