/**
 * 医師の勤務データから、各拠点・シフトごとの未充足時間を抽出してシートに書き出すスクリプト
 * (構文再確認版)
 */
function extractDoctorAbsenceRevised() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("貼付用");
  const targetSheet = ss.getSheetByName("不在時間");

  // --- ログレベル設定 (trueで詳細ログ、falseで主要ログ) ---
  const DETAILED_LOGGING = true; 

  Logger.log("スクリプト開始");

  if (!sourceSheet) {
    Logger.log("エラー: 「貼付用」シートが見つかりません。");
    SpreadsheetApp.getUi().alert("エラー: 「貼付用」シートが見つかりません。");
    return;
  }
  if (!targetSheet) {
    Logger.log("エラー: 「不在時間」シートが見つかりません。");
    SpreadsheetApp.getUi().alert("エラー: 「不在時間」シートが見つかりません。");
    return;
  }

  targetSheet.clear();

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  Logger.log(`今日の日付 (処理基準日): ${Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy/MM/dd")}`);

  // shiftTimesMin の定義
  const shiftTimesMin = {
    "北葛西小児科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 20 * 60] },
    "北葛西内科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 20 * 60] },
    "亀有小児科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] },
    "亀有内科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] },
    "その他": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] }
  };
  
  if (DETAILED_LOGGING) Logger.log(`シフト定義(shiftTimesMin): ${JSON.stringify(shiftTimesMin)}`);

  // ▼▼▼ ここにカンマを追加して修正しました ▼▼▼
  const excludedLocations = [
    "有給", "欠勤", "院外勤務（小児科）", "院外勤務（内科）",
    "【関東】バックアップシフト", "医師会・嘱託医業務（小児科）",
    "医師会・嘱託医業務（内科）", // <-- ここにカンマを追加
    "医師会",
  ];
  // ▲▲▲ 修正ここまで ▲▲▲
  const excludedDepartments = [
    "小児科ワクチン専任(対象：小児～成人)", "内科ワクチン専任(対象：小児～成人)"
  ];

  const sourceData = sourceSheet.getDataRange().getValues();
  const processedWorkData = {};
  Logger.log(`貼付用シートから ${sourceData.length -1} 行のデータを読み込みました (ヘッダー除く)。`);

  let processedRowCount = 0;
  for (let i = 2; i < sourceData.length; i++) {
    const row = sourceData[i];
    const clinicName = row[12] ? String(row[12]).trim() : "";
    const department = row[13] ? String(row[13]).trim() : "";
    let dateRaw = row[14];
    let startTimeRaw = row[15];
    let endTimeRaw = row[19];

    if (!dateRaw || dateRaw === "" || !clinicName || clinicName === "") {
      if (DETAILED_LOGGING) Logger.log(`行 ${i + 1}: スキップ (勤務日または勤務拠点が空欄)`);
      continue;
    }
    if (excludedLocations.includes(clinicName) || excludedDepartments.includes(department)) {
      if (DETAILED_LOGGING) Logger.log(`行 ${i + 1}: スキップ (除外リスト該当 拠点: ${clinicName}, 診療科: ${department})`);
      continue;
    }

    if (typeof dateRaw === "number") {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const dateObjUTC = new Date(excelEpoch.getTime() + dateRaw * 24 * 60 * 60 * 1000);
      dateRaw = new Date(dateObjUTC.getUTCFullYear(), dateObjUTC.getUTCMonth(), dateObjUTC.getUTCDate());
    }
    const dateObj = parseDateToSafeDateObj(dateRaw);
    if (!dateObj) {
      if (DETAILED_LOGGING) Logger.log(`行 ${i + 1}: スキップ (日付パース失敗: [${dateRaw}])`);
      continue;
    }
    
    const dateKey = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd");

    const workStartMin = parseTimeToMinutes(startTimeRaw);
    const workEndMin = parseTimeToMinutes(endTimeRaw);

    if (isNaN(workStartMin) || isNaN(workEndMin) || workStartMin >= workEndMin) {
      if (DETAILED_LOGGING) Logger.log(`行 ${i + 1}: スキップ (時間パース失敗または不正な時間帯: 開始=[${startTimeRaw}], 終了=[${endTimeRaw}])`);
      continue;
    }

    const aggregationKey = `${clinicName}_${department || ""}`;
    if (DETAILED_LOGGING && processedRowCount < 5) { 
        Logger.log(`行 ${i+1}: 処理中データ - dateKey: ${dateKey}, aggregationKey: ${aggregationKey}, start: ${workStartMin}, end: ${workEndMin}`);
    }

    if (!processedWorkData[dateKey]) {
      processedWorkData[dateKey] = {};
    }
    if (!processedWorkData[dateKey][aggregationKey]) {
      processedWorkData[dateKey][aggregationKey] = [];
    }
    processedWorkData[dateKey][aggregationKey].push({ start: workStartMin, end: workEndMin });
    processedRowCount++;
  }
  Logger.log(`データ収集完了。${processedRowCount}件の有効な勤務データを処理しました。`);
  if (DETAILED_LOGGING && Object.keys(processedWorkData).length > 0) {
      Logger.log(`processedWorkData のキー (日付): ${Object.keys(processedWorkData).join(', ')}`);
      const firstDateKey = Object.keys(processedWorkData)[0];
      if (processedWorkData[firstDateKey]) { // キーが存在するか確認
        Logger.log(`processedWorkData の最初の日の集約キー: ${Object.keys(processedWorkData[firstDateKey]).join(', ')}`);
      }
  }

  const outputDataRows = [];
  const sortedDates = Object.keys(processedWorkData).sort((a, b) => new Date(a) - new Date(b));
  Logger.log(`計算対象の日付 (ソート済): ${sortedDates.join(', ') || "なし"}`);

  for (const dateKey of sortedDates) {
    const parts = dateKey.split('/');
    const currentDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));

    if (currentDateObj < today) {
      if (DETAILED_LOGGING) Logger.log(`日付 ${dateKey}: スキップ (過去日)`);
      continue;
    }
    Logger.log(`日付 ${dateKey}: 処理開始`);

    const dailyData = processedWorkData[dateKey];
    let absencesForDate = [];

    for (const aggregationKey in dailyData) {
      const [currentClinicName, currentDepartment] = aggregationKey.split('_');
      Logger.log(`  拠点/診療科キー (aggregationKey): ${aggregationKey} (拠点: ${currentClinicName}, 診療科: ${currentDepartment || "なし"})`);

      const workIntervalsForClinic = dailyData[aggregationKey];
      if (DETAILED_LOGGING) Logger.log(`    勤務区間 (集約前): ${JSON.stringify(workIntervalsForClinic)}`);

      let shiftLookupKey = "その他";
      const specificKeyWithDept = `${currentClinicName}${currentDepartment}`;
      const specificKeyNoDept = currentClinicName;

      if (shiftTimesMin.hasOwnProperty(specificKeyWithDept)) {
        shiftLookupKey = specificKeyWithDept;
      } else if (shiftTimesMin.hasOwnProperty(specificKeyNoDept)) {
        shiftLookupKey = specificKeyNoDept;
      }
      const shiftsDefinition = shiftTimesMin[shiftLookupKey] || shiftTimesMin["その他"];
      Logger.log(`    使用するシフト定義キー: ${shiftLookupKey}, 定義: ${JSON.stringify(shiftsDefinition)}`);

      for (const shiftType in shiftsDefinition) {
        const shiftStartTarget = shiftsDefinition[shiftType][0];
        const shiftEndTarget = shiftsDefinition[shiftType][1];
        if (DETAILED_LOGGING) Logger.log(`      シフト ${shiftType} (${formatMinutesToHHMM(shiftStartTarget)}-${formatMinutesToHHMM(shiftEndTarget)}): 計算開始`);

        const relevantIntervals = workIntervalsForClinic.map(interval => ({
          start: Math.max(interval.start, shiftStartTarget),
          end: Math.min(interval.end, shiftEndTarget)
        })).filter(interval => interval.start < interval.end);
        if (DETAILED_LOGGING) Logger.log(`        関連勤務区間 (relevantIntervals): ${JSON.stringify(relevantIntervals)}`);

        if (relevantIntervals.length === 0) {
          absencesForDate.push(formatAbsence(currentClinicName, currentDepartment, shiftStartTarget, shiftEndTarget));
          if (DETAILED_LOGGING) Logger.log(`        不在: ${formatAbsence(currentClinicName, currentDepartment, shiftStartTarget, shiftEndTarget)} (勤務なし)`);
          continue;
        }

        const mergedIntervals = mergeIntervals(relevantIntervals);
        if (DETAILED_LOGGING) Logger.log(`        マージ後勤務区間 (mergedIntervals): ${JSON.stringify(mergedIntervals)}`);

        let currentTimePointer = shiftStartTarget;
        for (const merged of mergedIntervals) {
          if (currentTimePointer < merged.start) {
            absencesForDate.push(formatAbsence(currentClinicName, currentDepartment, currentTimePointer, merged.start));
            if (DETAILED_LOGGING) Logger.log(`        不在: ${formatAbsence(currentClinicName, currentDepartment, currentTimePointer, merged.start)}`);
          }
          currentTimePointer = Math.max(currentTimePointer, merged.end);
        }
        if (currentTimePointer < shiftEndTarget) {
          absencesForDate.push(formatAbsence(currentClinicName, currentDepartment, currentTimePointer, shiftEndTarget));
          if (DETAILED_LOGGING) Logger.log(`        不在: ${formatAbsence(currentClinicName, currentDepartment, currentTimePointer, shiftEndTarget)}`);
        }
      }
    }
    
    if (absencesForDate.length > 0) {
      const uniqueAbsences = [...new Set(absencesForDate)];
      Logger.log(`  日付 ${dateKey} の不在時間 (ユニーク): ${uniqueAbsences.join('; ')}`);
      outputDataRows.push([dateKey, ...uniqueAbsences]);
    } else {
      Logger.log(`  日付 ${dateKey} の不在時間はありませんでした。`);
    }
  }
  Logger.log(`不在時間の計算完了。書き出し対象は ${outputDataRows.length} 行です。`);

  if (outputDataRows.length > 0) {
    let maxCols = 0;
    outputDataRows.forEach(row => {
      if (row.length > maxCols) maxCols = row.length;
    });

    const headerRow = ["日付"];
    for (let k = 1; k < maxCols; k++) {
      headerRow.push(`不在時間${k}`);
    }
    
    const finalOutput = [headerRow, ...outputDataRows];

    const outputForSheet = finalOutput.map(row => {
      const newRow = [...row];
      while (newRow.length < maxCols) {
        newRow.push("");
      }
      return newRow;
    });

    const rangeToWrite = targetSheet.getRange(1, 1, outputForSheet.length, maxCols);
    rangeToWrite.setValues(outputForSheet);
    rangeToWrite.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    Logger.log("医師未充足時間の抽出が完了し、「不在時間」シートに書き出しました。(今日以降のデータのみ、セル内折り返し有効)");
    SpreadsheetApp.getUi().alert("医師未充足時間の抽出が完了しました。(今日以降のデータのみ、セル内折り返し有効)");
  } else {
    targetSheet.getRange(1,1).setValue("日付");
    targetSheet.getRange(1,2).setValue("該当する不在時間はありませんでした。(今日以降)");
    Logger.log("該当する不在時間はありませんでした。(今日以降)");
    SpreadsheetApp.getUi().alert("該当する不在時間はありませんでした。(今日以降)");
  }
  Logger.log("スクリプト終了");
}

// --- ヘルパー関数 ---

function parseTimeToMinutes(timeInput) {
  if (timeInput instanceof Date) {
    if (isNaN(timeInput.getTime())) return NaN;
    return timeInput.getHours() * 60 + timeInput.getMinutes();
  }
  if (typeof timeInput === 'string') {
    const parts = timeInput.split(':');
    if (parts.length === 2) {
      const hours = parseInt(parts[0], 10);
      const minutes = parseInt(parts[1], 10);
      if (!isNaN(hours) && !isNaN(minutes) && hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60) {
        return hours * 60 + minutes;
      }
    }
  }
  if (typeof timeInput === 'number') {
      if (timeInput >= 0 && timeInput < 1) {
          const totalMinutesInDay = 24 * 60;
          return Math.round(timeInput * totalMinutesInDay);
      }
  }
  return NaN;
}

function formatMinutesToHHMM(totalMinutes) {
  if (isNaN(totalMinutes) || totalMinutes < 0) return "不明";
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

function formatAbsence(clinicName, department, startMin, endMin) {
  let departmentSuffix = "";
  if ((clinicName === "北葛西" || clinicName === "亀有") && (department === "小児科" || department === "内科")) {
    departmentSuffix = department;
  }
  return `【${clinicName}】${formatMinutesToHHMM(startMin)}~${formatMinutesToHHMM(endMin)}${departmentSuffix}`;
}

function mergeIntervals(intervals) {
  if (!intervals || intervals.length === 0) {
    return [];
  }
  intervals.sort((a, b) => a.start - b.start);

  const merged = [];
  let currentMerge = { ...intervals[0] };

  for (let i = 1; i < intervals.length; i++) {
    const nextInterval = intervals[i];
    if (nextInterval.start <= currentMerge.end) {
      currentMerge.end = Math.max(currentMerge.end, nextInterval.end);
    } else {
      merged.push(currentMerge);
      currentMerge = { ...nextInterval };
    }
  }
  merged.push(currentMerge);
  return merged;
}

function parseDateToSafeDateObj(dateRaw) {
  if (dateRaw instanceof Date && !isNaN(dateRaw.getTime())) {
    return dateRaw;
  }
  if (typeof dateRaw === 'string' || typeof dateRaw === 'number') {
    const parsedDate = new Date(dateRaw);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate;
    }
  }
  return null;
}