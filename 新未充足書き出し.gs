// ------------------------------------------------------------------------------------
// グローバル設定に近い定数
// ------------------------------------------------------------------------------------
const SCRIPT_LOGGING_LEVEL = true; // trueで詳細ログ、falseで主要ログ

const GLOBAL_STANDARD_SHIFT_ORDER = ["A", "B", "C"];
const GLOBAL_SHIFT_TIMES_MIN = {
  "北葛西小児科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 20 * 60] },
  "北葛西内科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 20 * 60] },
  "亀有小児科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] },
  "亀有内科": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] },
  "その他": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] }
};

// 除外リスト
const GLOBAL_EXCLUDED_LOCATIONS = [
  "有給", "欠勤", "院外勤務（小児科）", "院外勤務（内科）",
  "【関東】バックアップシフト", "医師会・嘱託医業務（小児科）",
  "医師会・嘱託医業務（内科）",
  "医師会",
  "医師会業務", 
  "嘱託医業務"
];

const GLOBAL_EXCLUDED_DEPARTMENTS = [
  "小児科ワクチン専任(対象：小児～成人)", "内科ワクチン専任(対象：小児～成人)"
];
const GLOBAL_TARGET_CLINICS_FOR_DEPT_INFO = ["北葛西", "亀有"];


// ------------------------------------------------------------------------------------
// ヘルパー関数群
// ------------------------------------------------------------------------------------

function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) {
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }
  if (typeof dateInput !== 'string' && typeof dateInput.toString !== 'function') {
    if (SCRIPT_LOGGING_LEVEL) Logger.log(`日付パース不可: 文字列でもDateオブジェクトでもない入力値 "[${dateInput}]"`);
    return null;
  }
  const dateStr = dateInput.toString();
  const cleanedDateStr = dateStr.replace(/\s*（.*?）/, '').replace(/-/g, '/');
  const parts = cleanedDateStr.split('/');
  if (parts.length === 3) {
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      return new Date(year, month, day);
    }
  }
  if (SCRIPT_LOGGING_LEVEL) Logger.log(`日付パース失敗: 入力値 "[${dateInput}]", クリーンアップ後 "[${cleanedDateStr}]"`);
  return null;
}

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

function mergeIntervals(intervals) {
  if (!intervals || intervals.length === 0) {
    return [];
  }
  intervals.sort((a, b) => a.start - b.start);
  const merged = [];
  if (intervals.length === 0) return merged;

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

function formatAbsenceForExtractSheet(clinicName, department, startMin, endMin) {
  let departmentSuffix = "";
  if (GLOBAL_TARGET_CLINICS_FOR_DEPT_INFO.includes(clinicName) && (department === "小児科" || department === "内科")) {
    departmentSuffix = department;
  }
  return `【${clinicName}】${formatMinutesToHHMM(startMin)}~${formatMinutesToHHMM(endMin)}${departmentSuffix}`;
}


// ------------------------------------------------------------------------------------
// 「不在時間」シートへの出力処理関数 (extractDoctorAbsenceRevised)
// ------------------------------------------------------------------------------------
function extractDoctorAbsenceRevised() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("貼付用");
  const targetSheet = ss.getSheetByName("不在時間");

  if (!sourceSheet || !targetSheet) return;

  targetSheet.clear();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const shiftTimesMin = GLOBAL_SHIFT_TIMES_MIN;
  const excludedLocations = GLOBAL_EXCLUDED_LOCATIONS;
  const excludedDepartments = GLOBAL_EXCLUDED_DEPARTMENTS;
  const standardShiftOrder = GLOBAL_STANDARD_SHIFT_ORDER;

  const sourceData = sourceSheet.getDataRange().getValues();
  const processedWorkData = {}; 

  for (let i = 2; i < sourceData.length; i++) {
    const row = sourceData[i];
    const clinicName = row[12] ? String(row[12]).trim() : "";
    const department = row[13] ? String(row[13]).trim() : "";
    if (!row[14] || !clinicName) continue;
    if (excludedLocations.includes(clinicName) || excludedDepartments.includes(department)) continue;

    const dateObj = parseDateToSafeDateObj(row[14]);
    if (!dateObj) continue;
    
    const dateKey = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const workStartMin = parseTimeToMinutes(row[15]);
    const workEndMin = parseTimeToMinutes(row[19]);

    if (isNaN(workStartMin) || isNaN(workEndMin) || workStartMin >= workEndMin) continue;

    const aggregationKey = `${clinicName}_${department || ""}`;
    
    if (!processedWorkData[dateKey]) processedWorkData[dateKey] = {};
    if (!processedWorkData[dateKey][aggregationKey]) processedWorkData[dateKey][aggregationKey] = [];
    processedWorkData[dateKey][aggregationKey].push({ start: workStartMin, end: workEndMin });
  }

  const outputDataRows = [];
  const sortedDates = Object.keys(processedWorkData).sort((a, b) => new Date(a) - new Date(b));

  for (const dateKey of sortedDates) {
    const currentDateObj = parseDateToSafeDateObj(dateKey);
    if (currentDateObj < today) continue;

    const dailyWorkDataByAggregationKey = processedWorkData[dateKey];
    let absencesForDateOutput = [];

    for (const aggregationKey in dailyWorkDataByAggregationKey) {
      const [currentClinicName, currentDepartment] = aggregationKey.split('_');
      const workIntervalsForClinic = dailyWorkDataByAggregationKey[aggregationKey];

      let shiftLookupKey = "その他";
      const specificKeyWithDept = `${currentClinicName}${currentDepartment}`;
      if (shiftTimesMin.hasOwnProperty(specificKeyWithDept)) shiftLookupKey = specificKeyWithDept;
      else if (shiftTimesMin.hasOwnProperty(currentClinicName)) shiftLookupKey = currentClinicName;
      
      const currentClinicShifts = shiftTimesMin[shiftLookupKey] || shiftTimesMin["その他"];

      for (const shiftKey of standardShiftOrder) {
        if (!currentClinicShifts[shiftKey]) continue;
        const [shiftStartTarget, shiftEndTarget] = currentClinicShifts[shiftKey];
        
        const relevantIntervals = workIntervalsForClinic
          .map(interval => ({
            start: Math.max(interval.start, shiftStartTarget),
            end: Math.min(interval.end, shiftEndTarget)
          }))
          .filter(interval => interval.start < interval.end);

        const mergedIntervals = mergeIntervals(relevantIntervals);
        
        let currentTimePointer = shiftStartTarget;
        for (const merged of mergedIntervals) {
          if (currentTimePointer < merged.start) {
            absencesForDateOutput.push(formatAbsenceForExtractSheet(currentClinicName, currentDepartment, currentTimePointer, merged.start));
          }
          currentTimePointer = Math.max(currentTimePointer, merged.end);
        }
        if (currentTimePointer < shiftEndTarget) {
          absencesForDateOutput.push(formatAbsenceForExtractSheet(currentClinicName, currentDepartment, currentTimePointer, shiftEndTarget));
        }
      }
    }
    if (absencesForDateOutput.length > 0) {
      outputDataRows.push([dateKey, ...absencesForDateOutput]);
    }
  }
  
  if (outputDataRows.length > 0) {
    let maxCols = Math.max(...outputDataRows.map(r => r.length));
    const headerRow = ["日付", ...Array.from({length: maxCols - 1}, (_, i) => `不在時間${i + 1}`)];
    const finalOutput = [headerRow, ...outputDataRows.map(r => r.concat(Array(maxCols - r.length).fill("")))];
    
    targetSheet.getRange(1, 1, finalOutput.length, maxCols).setValues(finalOutput).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  } else {
    targetSheet.getRange(1,1,1,2).setValues([["日付", "該当する不在時間はありませんでした。(今日以降)"]]);
  }
}

// ------------------------------------------------------------------------------------
// 「医師不在拠点」シートへの出力処理関数 (generateDoctorAbsenceReportWithContext) (修正版)
// ------------------------------------------------------------------------------------

function getAdjacentShiftDoctors(currentClinicWorkData, currentShiftKey, clinicShiftTimesForLookup, shiftOrder) {
  let prevDoctorsStr = "";
  let nextDoctorsStr = "";
  const currentShiftIndex = shiftOrder.indexOf(currentShiftKey);

  if (currentShiftIndex > 0) {
    const prevShiftKey = shiftOrder[currentShiftIndex - 1];
    if (clinicShiftTimesForLookup && clinicShiftTimesForLookup[prevShiftKey]) {
      const [prevStart, prevEnd] = clinicShiftTimesForLookup[prevShiftKey];
      const doctorsInPrev = [...new Set(currentClinicWorkData
        .filter(work => work.start < prevEnd && work.end > prevStart)
        .map(work => work.doctor))];
      prevDoctorsStr = doctorsInPrev.length > 0 ? `${doctorsInPrev.join(',')}：${formatMinutesToHHMM(prevStart)}-${formatMinutesToHHMM(prevEnd)}` : "未充足";
    }
  }

  if (currentShiftIndex < shiftOrder.length - 1) {
    const nextShiftKey = shiftOrder[currentShiftIndex + 1];
    if (clinicShiftTimesForLookup && clinicShiftTimesForLookup[nextShiftKey]) {
      const [nextStart, nextEnd] = clinicShiftTimesForLookup[nextShiftKey];
       const doctorsInNext = [...new Set(currentClinicWorkData
        .filter(work => work.start < nextEnd && work.end > nextStart)
        .map(work => work.doctor))];
      nextDoctorsStr = doctorsInNext.length > 0 ? `${doctorsInNext.join(',')}：${formatMinutesToHHMM(nextStart)}-${formatMinutesToHHMM(nextEnd)}` : "未充足";
    }
  }
  return { prevDoctorsStr, nextDoctorsStr };
}

function generateDoctorAbsenceReportWithContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("貼付用");
  const targetSheet = ss.getSheetByName("医師不在拠点");

  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert("エラー: 「貼付用」または「医師不在拠点」シートが見つかりません。");
    return;
  }

  targetSheet.clear();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const shiftTimesMin = GLOBAL_SHIFT_TIMES_MIN;
  const excludedLocations = GLOBAL_EXCLUDED_LOCATIONS;
  const excludedDepartments = GLOBAL_EXCLUDED_DEPARTMENTS;
  const targetClinicsForDeptInfo = GLOBAL_TARGET_CLINICS_FOR_DEPT_INFO;
  const standardShiftOrder = GLOBAL_STANDARD_SHIFT_ORDER;

  const sourceData = sourceSheet.getDataRange().getValues();
  const processedWorkData = {};

  for (let i = 2; i < sourceData.length; i++) {
    const row = sourceData[i];
    const doctorName = row[0] ? String(row[0]).trim() : "不明医師";
    const clinicName = row[12] ? String(row[12]).trim() : "";
    const department = row[13] ? String(row[13]).trim() : "";
    if (!row[14] || !clinicName) continue;
    if (excludedLocations.includes(clinicName) || excludedDepartments.includes(department)) continue;

    const dateObj = parseDateToSafeDateObj(row[14]);
    if (!dateObj) continue;
    
    const dateKey = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const workStartMin = parseTimeToMinutes(row[15]);
    const workEndMin = parseTimeToMinutes(row[19]);

    if (isNaN(workStartMin) || isNaN(workEndMin) || workStartMin >= workEndMin) continue;

    const aggregationKey = `${clinicName}_${department || ""}`;
    
    if (!processedWorkData[dateKey]) processedWorkData[dateKey] = {};
    if (!processedWorkData[dateKey][aggregationKey]) processedWorkData[dateKey][aggregationKey] = [];
    processedWorkData[dateKey][aggregationKey].push({ doctor: doctorName, start: workStartMin, end: workEndMin });
  }

  const unfulfilledShiftsList = []; 
  const sortedDates = Object.keys(processedWorkData).sort((a, b) => new Date(a) - new Date(b));

  for (const dateKey of sortedDates) {
    const currentDateObj = parseDateToSafeDateObj(dateKey);
    if (currentDateObj < today) continue;

    const dailyWorkDataByAggregationKey = processedWorkData[dateKey];
    for (const aggregationKey in dailyWorkDataByAggregationKey) {
      const [currentClinicName, currentDepartment] = aggregationKey.split('_');
      const workIntervalsForClinic = dailyWorkDataByAggregationKey[aggregationKey];

      let shiftLookupKey = "その他";
      const specificKeyWithDept = `${currentClinicName}${currentDepartment}`;
      if (shiftTimesMin.hasOwnProperty(specificKeyWithDept)) shiftLookupKey = specificKeyWithDept;
      else if (shiftTimesMin.hasOwnProperty(currentClinicName)) shiftLookupKey = currentClinicName;
      
      const currentClinicShifts = shiftTimesMin[shiftLookupKey] || shiftTimesMin["その他"];

      for (const shiftKey of standardShiftOrder) {
        if (!currentClinicShifts[shiftKey]) continue;
        const [shiftStartTarget, shiftEndTarget] = currentClinicShifts[shiftKey];
        
        // ★修正点1: 勤務時間をマージ
        const mergedIntervals = mergeIntervals(workIntervalsForClinic
          .map(interval => ({ start: Math.max(interval.start, shiftStartTarget), end: Math.min(interval.end, shiftEndTarget) }))
          .filter(interval => interval.start < interval.end));
        
        // ★修正点2: 正確な隙間時間（gaps）を計算
        let currentTimePointer = shiftStartTarget;
        let gaps = [];

        for (const merged of mergedIntervals) {
            if (currentTimePointer < merged.start) {
                // 勤務開始前の空白
                gaps.push(`${formatMinutesToHHMM(currentTimePointer)}-${formatMinutesToHHMM(merged.start)}`);
            }
            currentTimePointer = Math.max(currentTimePointer, merged.end);
        }
        // 勤務終了後の空白
        if (currentTimePointer < shiftEndTarget) {
            gaps.push(`${formatMinutesToHHMM(currentTimePointer)}-${formatMinutesToHHMM(shiftEndTarget)}`);
        }

        // 隙間がある場合のみリストに追加（定型枠ではなく、正確なgapsを保存）
        if (gaps.length > 0) {
          const preciseAbsenceStr = gaps.join(", ");
          unfulfilledShiftsList.push({ 
              dateKey, 
              aggregationKey, 
              standardShiftKey: shiftKey, 
              preciseAbsenceStr: preciseAbsenceStr // ★正確な時間をセット
          });
        }
      }
    }
  }

  const outputDataRows = [];
  for (const {dateKey, aggregationKey, standardShiftKey, preciseAbsenceStr} of unfulfilledShiftsList) {
    const [clinicName, department] = aggregationKey.split('_');
    let outputClinicName = targetClinicsForDeptInfo.includes(clinicName) && department ? `${clinicName} (${department})` : clinicName;
    
    let shiftLookupKey = "その他";
    const specificKeyWithDept = `${clinicName}${department}`;
    if (shiftTimesMin.hasOwnProperty(specificKeyWithDept)) shiftLookupKey = specificKeyWithDept;
    else if (shiftTimesMin.hasOwnProperty(clinicName)) shiftLookupKey = clinicName;
    const currentClinicShiftTimes = shiftTimesMin[shiftLookupKey] || shiftTimesMin["その他"];
    
    // ★修正点3: 保存しておいた正確な時間を使用
    const absenceTimeStr = preciseAbsenceStr;
    
    const { prevDoctorsStr, nextDoctorsStr } = getAdjacentShiftDoctors(
      processedWorkData[dateKey]?.[aggregationKey] || [],
      standardShiftKey,
      currentClinicShiftTimes,
      standardShiftOrder
    );

    outputDataRows.push([dateKey, outputClinicName, absenceTimeStr, prevDoctorsStr, nextDoctorsStr]);
  }
  
  const header = [["日付", "拠点名", "不在時間", "前の時間枠の医師", "後ろの時間枠の医師"]];
  targetSheet.getRange(1, 1, 1, header[0].length).setValues(header);

  if (outputDataRows.length > 0) {
    outputDataRows.sort((a, b) => new Date(a[0]) - new Date(b[0]) || a[1].localeCompare(b[1]));
    targetSheet.getRange(2, 1, outputDataRows.length, outputDataRows[0].length).setValues(outputDataRows).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  } else {
    targetSheet.getRange(2,1).setValue("該当する不在情報はありませんでした。(今日以降)");
  }
  
  Logger.log("続けて extractDoctorAbsenceRevised (「不在時間」シートへの出力) を実行します。");
  extractDoctorAbsenceRevised();

  Logger.log("スクリプト終了: generateDoctorAbsenceReportWithContext (全処理完了)");
  SpreadsheetApp.getUi().alert("全レポート処理完了", "「医師不在拠点」および「不在時間」シートの更新が完了しました。", SpreadsheetApp.getUi().ButtonSet.OK);
}