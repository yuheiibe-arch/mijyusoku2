/**
 * 医師不在拠点シートから未充足管理シートへデータを転記する（修正版 V3 - 重複チェックを未充足管理シートのV列で管理）
 * 変更点：
 * 1. 医師不在拠点シートのV列（処理済みフラグ）によるスキップ処理を削除。
 * このフラグは自動更新で消えるため、未充足管理シート側で重複管理を行う。
 * 2. 未充足管理シートのV列（22列目）を、転記済みデータのユニークID格納用として使用。
 * 3. ユニークIDは「日付_拠点_開始時間_終了時間_新規応募待ち」の形式で生成。
 * 4. 新しいデータを転記する際、このユニークIDが未充足管理シートに既に存在するかで重複を判断。
 * 5. 重複しない場合のみ転記し、未充足管理シートのV列にユニークIDを書き込む。
 * 6. タイムアウト対策は、この新しい重複チェックロジックに引き継がれる。
 */
function insertDoctorAbsenceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("医師不在拠点");
  const targetSheet = ss.getSheetByName("未充足管理");

  if (!sourceSheet || !targetSheet) {
    Logger.log("エラー: シート「医師不在拠点」または「未充足管理」が見つかりません。");
    return;
  }

  const lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow < 2) {
    Logger.log("ソースシート「医師不在拠点」に処理対象データがありません（ヘッダー行を除く）。");
    return;
  }
  // V列(22列目)まで読み込むように範囲を拡張 (現状維持、ただし今回はV列は読まなくてよい)
  const sourceDataRange = sourceSheet.getRange(2, 1, lastSourceRow - 1, 22);
  const sourceData = sourceDataRange.getValues();

  // ★★★ 修正点: 既存のユニークIDを格納するSetを準備 ★★★
  const existingUniqueIds = new Set();
  const TARGET_DATA_START_ROW = 7;
  const UNIQUE_ID_COLUMN = 22; // V列 (1-indexed)

  if (targetSheet.getLastRow() >= TARGET_DATA_START_ROW) {
    // 重複チェック用のデータ（B列～I列）
    const rangeForDuplicateCheck = targetSheet.getRange(TARGET_DATA_START_ROW, 2, targetSheet.getLastRow() - TARGET_DATA_START_ROW + 1, 8);
    const targetDisplayValues = rangeForDuplicateCheck.getDisplayValues(); // 表示されている値

    // ★★★ 修正点: ユニークIDが格納されているV列の値も読み込む ★★★
    const uniqueIdRange = targetSheet.getRange(TARGET_DATA_START_ROW, UNIQUE_ID_COLUMN, targetSheet.getLastRow() - TARGET_DATA_START_ROW + 1, 1);
    const existingUniqueIdValues = uniqueIdRange.getValues();

    for (let i = 0; i < existingUniqueIdValues.length; i++) {
      const uniqueId = existingUniqueIdValues[i][0];
      if (uniqueId && typeof uniqueId === 'string') {
        existingUniqueIds.add(uniqueId.trim()); // trimしてSetに追加
      }
    }
  }

  Logger.log(`医師不在データ処理開始: ${sourceData.length} 件のデータを処理します。`);
  Logger.log(`「未充足管理」シートに既存のユニークIDは ${existingUniqueIds.size} 件あります。`);
  const scriptTimeZone = Session.getScriptTimeZone();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const limitDate = new Date(today);
  limitDate.setDate(today.getDate() + 14);

  const logStartDate = new Date(today.getTime() + (24 * 60 * 60 * 1000));
  Logger.log(`処理対象の日付範囲 (目安): ${Utilities.formatDate(logStartDate, scriptTimeZone, "yyyy/MM/dd")} から ${Utilities.formatDate(limitDate, scriptTimeZone, "yyyy/MM/dd")} まで (今日から2週間以内)`);

  let addedRowCount = 0;

  // forEachのindexは0から始まるため、シートの行番号に変換する際は +2 が必要 (ヘッダーと0-indexedのため)
  sourceData.forEach((row, index) => {
    const date = row[0];         // A列のデータ
    const location = row[1];     // B列のデータ
    const timeRangeStr = row[2]; // C列のデータ
    // const status = row[21];   // V列のデータ (V2の処理済みフラグ、今回は使用しない)

    // ★★★ 修正点: 医師不在拠点シートのV列によるスキップは削除 ★★★
    // if (status === "処理済み") {
    //   return; // 次の行へ
    // }

    if (!date || !location || !timeRangeStr) {
      Logger.log(`スキップ (ソースデータ不完全): 日付=${date}, 拠点=${location}, 時間=${timeRangeStr}`);
      return;
    }
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      Logger.log(`スキップ (ソースの日付が無効なDateオブジェクト): ${date}`);
      return;
    }

    const sourceDateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());

    if (sourceDateOnly <= today) {
      return;
    }
    if (sourceDateOnly > limitDate) {
      return;
    }

    const formattedDate = formatDateToJapanese(date); // 例: 6月22日(日)
    const timeParts = String(timeRangeStr).split("-");
    if (timeParts.length !== 2) {
      Logger.log(`スキップ (時間範囲の形式が不正): "${timeRangeStr}"。`);
      return;
    }
    const startTime = timeParts[0].trim();
    const endTime = timeParts[1].trim();
    const finalLocation = cleanLocation(String(location));

    // ★★★ 修正点: ユニークIDを生成する ★★★
    // ステータスは「新規応募待ち」で固定
    const currentStatusForId = "新規応募待ち";
    const generatedUniqueId = `${formattedDate}_${finalLocation}_${startTime}_${endTime}_${currentStatusForId}`;

    // ★★★ 修正点: 生成したユニークIDで重複チェック ★★★
    if (existingUniqueIds.has(generatedUniqueId)) {
      Logger.log(`スキップ (重複検出 - ユニークID): ${generatedUniqueId}`);
      return;
    }

    // 重複チェックは上記で完結するため、以下の表示値による重複チェックは不要になるが、
    // 念のため、これまで通り`targetDataValues`を使うロジックは残しておく（必要なければ削除可）
    // const isDuplicate = targetDisplayValues.some(existingRowDisplayValues => {
    //   const targetBcolDisplayValue = existingRowDisplayValues[0];
    //   const targetGcolDisplayValue = existingRowDisplayValues[5];
    //   const targetHcolDisplayValue = existingRowDisplayValues[6];
    //   const targetIcolDisplayValue = existingRowDisplayValues[7];
    //   const dateMatches = (targetBcolDisplayValue === formattedDate);
    //   const locationMatches = (cleanLocation(targetGcolDisplayValue) === finalLocation);
    //   const startTimeMatches = (String(targetHcolDisplayValue).trim() === startTime);
    //   const endTimeMatches = (String(targetIcolDisplayValue).trim() === endTime);
    //   return dateMatches && locationMatches && startTimeMatches && endTimeMatches;
    // });

    // if (isDuplicate) {
    //   Logger.log(`スキップ (旧方式の重複検出): ${formattedDate}, ${finalLocation}, ${startTime}-${endTime}`);
    //   return;
    // }

    const dateForTcol = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    // findLastRowForDateはtargetDatesValues（T列の値）を使うが、今回は重複チェックでV列を使うため、
    // ここでtargetDatesValuesを読み込むのは冗長になる可能性あり。
    // targetDatesValues = rangeForSortHelper.getValues(); の部分も修正が必要かも
    // ただし、ソートのためにT列は必要なので、findLastRowForDateとtargetDatesValuesの準備はそのままにする。
    
    // findLastRowForDate はT列（日付）を元に挿入位置を見つけるヘルパー関数
    // そのためには targetDatesValues が必要
    let targetDatesValues = [];
    if (targetSheet.getLastRow() >= TARGET_DATA_START_ROW) {
        const rangeForSortHelper = targetSheet.getRange(TARGET_DATA_START_ROW, 20, targetSheet.getLastRow() - TARGET_DATA_START_ROW + 1, 1);
        targetDatesValues = rangeForSortHelper.getValues();
    }
    
    let lastRowForDateGroup = findLastRowForDate(targetDatesValues, dateForTcol, TARGET_DATA_START_ROW);
    let insertAfterRowNumber;

    if (lastRowForDateGroup === -1) {
      // 特定の日付のデータがまだない場合、末尾に挿入するか、データ開始行の直前に挿入
      insertAfterRowNumber = targetSheet.getLastRow() < TARGET_DATA_START_ROW ? TARGET_DATA_START_ROW - 1 : targetSheet.getLastRow();
    } else {
      // 特定の日付の最終行の次に挿入
      insertAfterRowNumber = lastRowForDateGroup;
    }
    
    let effectiveInsertAfterRow = Math.max(insertAfterRowNumber, TARGET_DATA_START_ROW - 1);

    targetSheet.insertRowAfter(effectiveInsertAfterRow);
    const newActualRowNumber = effectiveInsertAfterRow + 1;
    addedRowCount++;

    targetSheet.getRange(newActualRowNumber, 2).setValue(formattedDate);
    targetSheet.getRange(newActualRowNumber, 3).setValue(currentStatusForId); // B列に「新規応募待ち」
    targetSheet.getRange(newActualRowNumber, 7).setValue(finalLocation);
    targetSheet.getRange(newActualRowNumber, 8).setValue(startTime);
    targetSheet.getRange(newActualRowNumber, 9).setValue(endTime);
    targetSheet.getRange(newActualRowNumber, 13).setValue("未対応");
    targetSheet.getRange(newActualRowNumber, 20).setValue(dateForTcol);

    // ★★★ 修正点: 未充足管理シートのV列(22列目)にユニークIDを書き込む ★★★
    targetSheet.getRange(newActualRowNumber, UNIQUE_ID_COLUMN).setValue(generatedUniqueId);

    Logger.log(`追加: 行 ${newActualRowNumber}: B=${formattedDate}, G=${finalLocation} | ユニークID: ${generatedUniqueId}`);
  });

  Logger.log(`医師不在データの挿入処理ループ完了。 ${addedRowCount} 件のデータを追加対象としました。`);

  // ソート処理 (変更なし)
  if (addedRowCount > 0 || targetSheet.getLastRow() >= TARGET_DATA_START_ROW) {
    const dataRangeToSort = targetSheet.getRange(TARGET_DATA_START_ROW, 1, targetSheet.getLastRow() - TARGET_DATA_START_ROW + 1, targetSheet.getLastColumn());
    if (targetSheet.getLastRow() >= TARGET_DATA_START_ROW) {
      Logger.log(`「未充足管理」シートをT列 (20列目) 基準で昇順ソートします。範囲: ${dataRangeToSort.getA1Notation()}`);
      dataRangeToSort.sort({ column: 20, ascending: true });
      Logger.log("ソート処理完了。");
    }
  }

  Logger.log("引き続き、日付形式の変換処理 (convertDateFormat) を開始します。");
  convertDateFormat(); // この関数がプロジェクト内に定義されていることを確認してください。
}

// --- 以下、ヘルパー関数 (変更なし) ---

function formatDateToJapanese(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    Logger.log(`formatDateToJapanese: 無効な日付オブジェクトまたは値 ${dateObj}`);
    return "日付エラー";
  }
  const weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  const month = dateObj.getMonth() + 1;
  const day = dateObj.getDate();
  const weekDay = weekDays[dateObj.getDay()];
  return `${month}月${day}日(${weekDay})`;
}

function cleanLocation(locationStr) {
  if (typeof locationStr !== 'string') {
    Logger.log(`cleanLocation: 文字列でない入力値 ${locationStr}`);
    return "";
  }
  let cleaned = locationStr.replace(/　/g, " ").trim();
  cleaned = cleaned.replace(/\s*\(\s*/g, "(");
  cleaned = cleaned.replace(/\s*\)\s*/g, ")");
  cleaned = cleaned.replace(/\s+/g, " ");
  return cleaned;
}

function findLastRowForDate(targetDatesValues, dateToFind, dataStartRow) {
  let lastMatchingRowNumber = -1;
  if (!(dateToFind instanceof Date) || isNaN(dateToFind.getTime())) {
    Logger.log(`findLastRowForDate: 無効な dateToFind: ${dateToFind}`);
    return -1;
  }
  const targetDateEpoch = dateToFind.getTime();

  for (let i = 0; i < targetDatesValues.length; i++) {
    let cellValue = targetDatesValues[i][0];

    if (cellValue instanceof Date && !isNaN(cellValue.getTime())) {
      const cellDateForCompare = new Date(cellValue.getFullYear(), cellValue.getMonth(), cellValue.getDate());
      if (cellDateForCompare.getTime() === targetDateEpoch) {
        lastMatchingRowNumber = dataStartRow + i;
      }
    }
  }
  return lastMatchingRowNumber;
}

// 以下の convertDateFormat 関数の定義もプロジェクト内に必要です。
// function convertDateFormat() { ... }