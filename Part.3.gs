// ------------------------------------------------------------------------------------
// 不在枠の前後の医師を取得するヘルパー関数
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

// ------------------------------------------------------------------------------------
// 「医師不在拠点」シートへの出力メイン関数
// ------------------------------------------------------------------------------------
function generateDoctorAbsenceReportWithContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("貼付用");
  const targetSheet = ss.getSheetByName("医師不在拠点");

  if (!sourceSheet || !targetSheet) {
    ss.toast("エラー: 「貼付用」または「医師不在拠点」シートが見つかりません。", 'エラー', 5);
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
    
    const dateKey = fastFormatDate(dateObj); 
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
  
  ss.toast("「医師不在拠点」シートの更新が完了しました。", "完了", 3);
}