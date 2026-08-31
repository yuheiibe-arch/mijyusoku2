// メインの更新関数
function updateSheetRowAdjusted_CallingCellSpecificFormatting() {
  const ss = SpreadsheetApp.openById('1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA');
  const sourceSheet = ss.getSheetByName('貼付用');
  const targetSheet = ss.getSheetByName('確認用');
  const DETAILED_LOGGING = true;

  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert('貼付用 または 確認用 シートが見つかりません。');
    return;
  }

  // ▼▼▼ 修正点① ▼▼▼
  // この関数内でのみ使用する、特別ルールを適用した除外リストを作成します。
  // グローバルの除外リストから「【関東】バックアップシフト」だけを取り除きます。
  const localExcludedLocations = (typeof GLOBAL_EXCLUDED_LOCATIONS !== 'undefined') 
    ? GLOBAL_EXCLUDED_LOCATIONS.filter(item => item !== "【関東】バックアップシフト")
    : [];
  // ▲▲▲ 修正ここまで ▲▲▲

  const data = sourceSheet.getDataRange().getValues();
  const rows = data.slice(2);

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('貼付用シートの3行目以降にデータが見つかりません。');
    return;
  }

  const EXCLUDED_DEPARTMENTS = [
    "小児科ワクチン専任(対象：小児～成人)", "内科ワクチン専任(対象：小児～成人)"
  ];
  const TARGET_CLINICS_FOR_DEPT_SPLIT = ["北葛西", "亀有"];
  const results = {};
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];

  rows.forEach((row, index) => {
    const rowIndex = index + 3;
    const doctorName = row[0] ? row[0].toString().trim() : '未設定';
    const originalClinicName = row[12] ? row[12].toString().trim() : null;
    const originalDepartment = row[13] ? row[13].toString().trim() : null;
    const shiftDateRaw = row[14];
    
    // ▼▼▼ 修正箇所：参照列を BL, BM, BN に変更 (インデックス 63, 64, 65) ▼▼▼
    const aShiftValue = row[63]; // BL列
    const bShiftValue = row[64]; // BM列
    const cShiftValue = row[65]; // BN列
    // ▲▲▲ 修正ここまで ▲▲▲

    // ▼▼▼ 修正点② ▼▼▼
    // 除外判定を、グローバルリストの代わりに先ほど作成した「localExcludedLocations」で行います。
    if ((originalClinicName && localExcludedLocations.includes(originalClinicName)) ||
        (originalDepartment && EXCLUDED_DEPARTMENTS.includes(originalDepartment))) {
    // ▲▲▲ 修正ここまで ▲▲▲
      if (DETAILED_LOGGING) Logger.log(`行 ${rowIndex}: スキップ (理由: 除外項目該当) Clinic=${originalClinicName}, Dept=${originalDepartment}`);
      return;
    }
    
    if (!originalClinicName || !shiftDateRaw || !originalDepartment) {
      if (DETAILED_LOGGING) Logger.log(`行 ${rowIndex}: スキップ (必須項目不足) Clinic=${originalClinicName}, Date=${shiftDateRaw}, Dept=${originalDepartment}`);
      return;
    }
    let shiftDateObj, shiftDateKey;
    try {
      shiftDateObj = parseDateToSafeDateObj(shiftDateRaw); // ★強化版パースを利用
      if (!shiftDateObj || isNaN(shiftDateObj.getTime())) {
        Logger.log(`行 ${rowIndex}: スキップ (無効な日付形式: ${shiftDateRaw})`);
        return;
      }
      shiftDateKey = fastFormatDate(shiftDateObj).replace(/\//g, '-'); // ★爆速化
    } catch (e) {
      Logger.log(`行 ${rowIndex}: スキップ (日付処理エラー: ${e}, データ: ${shiftDateRaw})`);
      return;
    }

    let displayClinicName = originalClinicName;
    if (TARGET_CLINICS_FOR_DEPT_SPLIT.includes(originalClinicName) &&
        (originalDepartment === "小児科" || originalDepartment === "内科")) {
      displayClinicName = `${originalClinicName}（${originalDepartment}）`;
    }
    const key = `${shiftDateKey}-${displayClinicName}-${originalDepartment}`;

    if (!results[key]) {
      results[key] = {
        shiftDate: shiftDateObj, department: originalDepartment, clinicName: displayClinicName,
        aShiftSum: 0, bShiftSum: 0, cShiftSum: 0,
        doctorsA: [], doctorsB: [], doctorsC: []
      };
    }
    results[key].aShiftSum += Number(aShiftValue) || 0;
    results[key].bShiftSum += Number(bShiftValue) || 0;
    results[key].cShiftSum += Number(cShiftValue) || 0;
    if (Number(aShiftValue) == 1 && doctorName !== '未設定' && !results[key].doctorsA.includes(doctorName)) { results[key].doctorsA.push(doctorName); }
    if (Number(bShiftValue) == 1 && doctorName !== '未設定' && !results[key].doctorsB.includes(doctorName)) { results[key].doctorsB.push(doctorName); }
    if (Number(cShiftValue) == 1 && doctorName !== '未設定' && !results[key].doctorsC.includes(doctorName)) { results[key].doctorsC.push(doctorName); }
  });

  const outputHeader = ['拠点名', '勤務日', '診療科', '09:00~13:00', '15:00~18:00', '18:00~21:00', '', 'Aシフト医師 09:00-13:00', 'Bシフト医師 15:00-18:00', 'Cシフト医師 18:00-21:00'];
  const numOutputColumns = outputHeader.length;
  const outputRows = [];
  for (const key in results) {
    const record = results[key];
    let formattedShiftDate = '日付エラー';
    try {
      if (record.shiftDate instanceof Date && !isNaN(record.shiftDate.getTime())) {
        const yyyymmdd = fastFormatDate(record.shiftDate); // ★爆速化
        const weekday = weekdays[record.shiftDate.getDay()];
        formattedShiftDate = `${yyyymmdd}（${weekday}）`;
      } else { Logger.log(`不正な日付オブジェクトでフォーマット試行: key=${key}, shiftDate=${record.shiftDate}`); }
    } catch (e) { formattedShiftDate = 'フォーマットエラー'; Logger.log(`日付フォーマットエラー: ${e}, Date: ${record.shiftDate}`); }
    const outputRowData = [
      record.clinicName,
      formattedShiftDate,
      record.department,
      record.aShiftSum,
      record.bShiftSum,
      record.cShiftSum,
      '',
      record.doctorsA.join(', '),
      record.doctorsB.join(', '),
      record.doctorsC.join(', ')
    ];
    outputRows.push(outputRowData);
  }
  const sortedData = outputRows.sort((a, b) => {
    const clinicA = a[0], clinicB = b[0];
    const dateStrA = a[1].split('（')[0].trim(), dateStrB = b[1].split('（')[0].trim();
    let dateA = new Date("invalid"), dateB = new Date("invalid");
    try { dateA = parseDateToSafeDateObj(dateStrA) || new Date("invalid"); } catch(e){}
    try { dateB = parseDateToSafeDateObj(dateStrB) || new Date("invalid"); } catch(e){}
    if (clinicA < clinicB) return -1; if (clinicA > clinicB) return 1;
    if (!isNaN(dateA.getTime()) && !isNaN(dateB.getTime())) return dateA.getTime() - dateB.getTime();
    if (!isNaN(dateA.getTime())) return -1; if (!isNaN(dateB.getTime())) return 1;
    return a[1].localeCompare(b[1]);
  });
  const lastRowOutput = targetSheet.getLastRow();
  if (lastRowOutput >= 1) {
    targetSheet.getRange(1, 1, lastRowOutput, targetSheet.getMaxColumns()).clearContent();
  }
  targetSheet.getRange(1, 1, 1, numOutputColumns).setValues([outputHeader]);
  if (sortedData.length > 0) {
    targetSheet.getRange(2, 1, sortedData.length, numOutputColumns).setValues(sortedData);
  }
  SpreadsheetApp.flush();
  try {
    Logger.log('applyConditionalFormatting_CellSpecific() を呼び出します');
    applyConditionalFormatting_CellSpecific();
  } catch (e) {
    Logger.log(`セル別書式設定(applyConditionalFormatting_CellSpecific)の呼び出し中にエラー: ${e}`);
  }
  try {
    Logger.log('generateDoctorAbsenceReportWithContext() を呼び出します');
    if (typeof generateDoctorAbsenceReportWithContext === 'function') {
        generateDoctorAbsenceReportWithContext();
    }
  } catch (e) {
    Logger.log(`医師不在拠点書き出し(generateDoctorAbsenceReportWithContext)の呼び出し中にエラー: ${e}`);
    SpreadsheetApp.getUi().alert('エラー', `医師不在拠点シートへの書き出し中にエラーが発生しました: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  try {
    Logger.log('setupDateSelection() を呼び出します');
    setupDateSelection();
  } catch (e) {
    Logger.log(`setupDateSelection() の呼び出し中にエラー: ${e}`);
    SpreadsheetApp.getUi().alert('情報', `日付選択の設定処理(setupDateSelection)中にエラーが発生しました: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  try {
    Logger.log('updateUnfilledStatusWithClinicLogicAndReset_v3() を呼び出します');
    if (typeof updateUnfilledStatusWithClinicLogicAndReset_v3 === 'function') {
        updateUnfilledStatusWithClinicLogicAndReset_v3(ss);
    }
  } catch (e) {
    Logger.log(`「未充足管理」シート更新処理(updateUnfilledStatusWithClinicLogicAndReset_v3)の呼び出し中にエラー: ${e}`);
    SpreadsheetApp.getUi().alert('エラー', `「未充足管理」シートの更新処理中にエラーが発生しました: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  try {
    Logger.log('insertDoctorAbsenceData() を呼び出します');
    if (typeof insertDoctorAbsenceData === 'function') {
        insertDoctorAbsenceData(ss);
    }
  } catch (e) {
    Logger.log(`医師不在データの挿入処理(insertDoctorAbsenceData)の呼び出し中にエラー: ${e}`);
    SpreadsheetApp.getUi().alert('エラー', `医師不在データの挿入処理中にエラーが発生しました: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  Logger.log('スクリプト updateSheetRowAdjusted_CallingCellSpecificFormatting 完了');
}

// ------------------------------------------------------------------------------------
// 他の関数 (clearData, applyConditionalFormatting_CellSpecific, setupDateSelection)
// ------------------------------------------------------------------------------------
function clearData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
      'データの削除と書式リセット',
      'すべてのデータを削除しますか？',
      ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sourceSheet = ss.getSheetByName('貼付用');
      const targetSheet = ss.getSheetByName('確認用');

      let message = '';
      if (sourceSheet) {
        const lastRowSource = sourceSheet.getLastRow();
        if (lastRowSource >= 3) { 
          sourceSheet.getRange(3, 1, lastRowSource - 2, 49).clear();
          message += '「貼付用」シートのデータと書式を削除しました。\n';
        }
      }
      if (targetSheet) {
        const lastRowTarget = targetSheet.getLastRow();
        if (lastRowTarget >= 2) { 
          targetSheet.getRange(2, 1, lastRowTarget - 1, 10).clear();
          message += '「確認用」シートのデータと書式を削除しました。';
        }
      }
      ui.alert(message || '削除対象のデータがありませんでした。');
    } catch (e) {
      ui.alert(`データ削除中にエラーが発生しました: ${e.message}`);
    }
  }
}

function applyConditionalFormatting_CellSpecific() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName('確認用');
  if (!targetSheet) return;

  const startRow = 2;
  const lastRow = targetSheet.getLastRow();
  if (lastRow < startRow) return;

  const numColumns = 10;
  const dataRange = targetSheet.getRange(startRow, 1, lastRow - startRow + 1, numColumns);
  const values = dataRange.getValues();
  const backgroundColors = [];
  const redColor = '#FFCCCC';
  const whiteColor = '#FFFFFF';

  for (let i = 0; i < values.length; i++) {
    const rowBackgrounds = [];
    for (let j = 0; j < numColumns; j++) {
      let cellColor = whiteColor;
      const cellValue = values[i][j];
      switch (j) {
        case 3: // D
        case 4: // E
        case 5: // F
          if (cellValue === '' || cellValue == null || Number(cellValue) === 0) {
            cellColor = redColor;
          }
          break;
        case 7: // H
        case 8: // I
        case 9: // J
          if (cellValue === '' || cellValue == null) {
            cellColor = redColor;
          }
          break;
      }
      rowBackgrounds.push(cellColor);
    }
    backgroundColors.push(rowBackgrounds);
  }
  dataRange.setBackgrounds(backgroundColors);
}

function setupDateSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('文章自動作成');
  const sourceSheet = ss.getSheetByName('確認用');
  if (!sheet || !sourceSheet) return;

  const baseDate = new Date();
  if (baseDate.getHours() >= 15) {
    baseDate.setDate(baseDate.getDate() + 1);
  }

  const endDate = new Date(baseDate);
  endDate.setDate(baseDate.getDate() + 6); // 起点日から6日後（計7日間の1週間分）

  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"];
  const formattedStart = fastFormatDate(baseDate) + `（${weekdaysJP[baseDate.getDay()]}）`; // ★爆速化
  const formattedEnd = fastFormatDate(endDate) + `（${weekdaysJP[endDate.getDay()]}）`; // ★爆速化

  const lastRow = sourceSheet.getLastRow();
  let uniqueValues = [];

  if (lastRow >= 2) {
    const bColumnValues = sourceSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    uniqueValues = [...new Set(bColumnValues.filter(Boolean))].map(dateStr => {
      const dateObj = parseDateToSafeDateObj(dateStr); // ★強化版パース
      return !dateObj ? dateStr : fastFormatDate(dateObj) + `（${weekdaysJP[dateObj.getDay()]}）`;
    });
  }

  if (!uniqueValues.includes(formattedStart)) uniqueValues.unshift(formattedStart);
  if (!uniqueValues.includes(formattedEnd)) uniqueValues.push(formattedEnd);

  if (uniqueValues.length === 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueValues, true)
    .setAllowInvalid(true) 
    .build();

  // ★修正: 手動で選んだ日付を維持し、空欄の場合のみ自動セットする
  const cellB2 = sheet.getRange('B2');
  const cellB4 = sheet.getRange('B4');
  
  let currentB2 = cellB2.getValue();
  let currentB4 = cellB4.getValue();

  if (!currentB2) cellB2.setValue(formattedStart);
  if (!currentB4) cellB4.setValue(formattedEnd);
  
  cellB2.setDataValidation(rule);
  cellB4.setDataValidation(rule);

  Logger.log(`期間設定完了: ${formattedStart} 〜 ${formattedEnd}`);
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

    const dateObj = parseDateToSafeDateObj(row[14]); // ★強化版パース
    if (!dateObj) continue;
    
    const dateKey = fastFormatDate(dateObj); // ★爆速化
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
    const currentDateObj = parseDateToSafeDateObj(dateKey); // ★強化版パース
    if (!currentDateObj || currentDateObj < today) continue;

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
        
        const mergedIntervals = mergeIntervals(workIntervalsForClinic
          .map(interval => ({ start: Math.max(interval.start, shiftStartTarget), end: Math.min(interval.end, shiftEndTarget) }))
          .filter(interval => interval.start < interval.end));
        
        let currentTimePointer = shiftStartTarget;
        let gaps = [];

        for (const merged of mergedIntervals) {
            if (currentTimePointer < merged.start) {
                gaps.push(`${formatMinutesToHHMM(currentTimePointer)}-${formatMinutesToHHMM(merged.start)}`);
            }
            currentTimePointer = Math.max(currentTimePointer, merged.end);
        }
        if (currentTimePointer < shiftEndTarget) {
            gaps.push(`${formatMinutesToHHMM(currentTimePointer)}-${formatMinutesToHHMM(shiftEndTarget)}`);
        }

        if (gaps.length > 0) {
          const preciseAbsenceStr = gaps.join(", ");
          unfulfilledShiftsList.push({ 
              dateKey, 
              aggregationKey, 
              standardShiftKey: shiftKey, 
              preciseAbsenceStr: preciseAbsenceStr
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