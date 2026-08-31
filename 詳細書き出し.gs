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
// ※ ヘルパー関数群（parseDateToSafeDateObj, parseTimeToMinutesなど）は、
// すでに Helpers.gs に最新版が定義されているため、ここからは完全に削除しました。
// これにより関数の重複エラーが解消されます。
// ------------------------------------------------------------------------------------

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
    
    // ★ Helpers.gsの高速フォーマットを利用
    const dateKey = fastFormatDate(dateObj);
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

        const mergedIntervals = mergeIntervalsForMetrics(relevantIntervals); // ★ Helpers.gsの統合関数を利用
        
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
// その他の旧フォーマット用メッセージ生成関数 (generateChatworkMessage2)
// ------------------------------------------------------------------------------------
function generateChatworkMessage2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('確認用');
  const targetSheet = ss.getSheetByName('文章自動作成');
  const mentionSheet = ss.getSheetByName('メンション先選択');

  if (!sourceSheet || !targetSheet || !mentionSheet) {
    Logger.log('「確認用」「文章自動作成」「メンション先選択」シートのいずれかが見つかりません。');
    return;
  }

  Logger.log('★★★ スクリプト実行開始 ★★★');
  targetSheet.getRange('A6').clearContent();
  Logger.log('A6 の内容をクリアしました');

  const startDateRaw = targetSheet.getRange('B2').getValue();
  const endDateRaw = targetSheet.getRange('B4').getValue();
  const startDate = parseDateToSafeDateObj(startDateRaw);
  const endDate = parseDateToSafeDateObj(endDateRaw);

  if (!startDate || !endDate) {
    Logger.log(`B2またはB4の日付が無効です。startDate: ${startDate}, endDate: ${endDate}`);
    return;
  }

  Logger.log(`パース後の日付範囲: startDate=[${startDate.toISOString().slice(0,10)}], endDate=[${endDate.toISOString().slice(0,10)}]`);

  const now = new Date();
  let hours = now.getHours();
  let minutes = now.getMinutes();
  if (minutes <= 19) minutes = "00";
  else if (minutes <= 49) minutes = "30";
  else {
    minutes = "00";
    hours = (hours + 1) % 24;
  }

  const formattedHours = hours < 10 ? '0' + hours : hours.toString();
  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"];
  const reportDateObj = new Date();
  reportDateObj.setDate(now.getDate() - 1);
  const formattedReportDate = `${reportDateObj.getMonth() + 1}月${reportDateObj.getDate()}日（${weekdaysJP[reportDateObj.getDay()]}）`;
  const reportTime = `【未充足報告】${formattedReportDate} ${formattedHours}:${minutes}時点\n\n`;

  Logger.log(`生成された reportTime: "${reportTime.replace(/\n/g, "\\n")}"`);

  // --- メンション情報取得 ---
  let mentionsArray = [];
  let ccArray = [];
  const mentionData = mentionSheet.getDataRange().getValues();

  for (let i = 1; i < mentionData.length; i++) {
    const toMention = mentionData[i][0];
    const ccMention = mentionData[i][1];
    if (toMention && toMention.toString().trim() !== "") {
      mentionsArray.push(toMention.toString().trim());
    }
    if (ccMention && ccMention.toString().trim() !== "") {
      ccArray.push(ccMention.toString().trim());
    }
  }

  let message = reportTime;

  // TOメンション整形
  if (mentionsArray.length > 0) {
    let toMentionsText = "";
    for (let i = 0; i < mentionsArray.length; i += 2) {
      toMentionsText += mentionsArray[i];
      if (i + 1 < mentionsArray.length) {
        toMentionsText += "\t\t" + mentionsArray[i + 1];
      }
      toMentionsText += "\n";
    }
    message += toMentionsText + "\n";
  }

  // CCメンション整形
  if (ccArray.length > 0) {
    let ccMentionsText = "CC:\n";
    for (let i = 0; i < ccArray.length; i += 2) {
      ccMentionsText += ccArray[i];
      if (i + 1 < ccArray.length) {
        ccMentionsText += "\t\t" + ccArray[i + 1];
      }
      ccMentionsText += "\n";
    }
    message += ccMentionsText + "\n";
  }

  Logger.log('★★★ 本文の生成を開始（データ集約フェーズ） ★★★');
  const sheetData = sourceSheet.getDataRange().getValues();
  const dailyReportData = {};
  const KANTO_BACKUP_SHIFT_NAME_CONST = "【関東】バックアップシフト";

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const clinicName = row[0] ? row[0].toString().trim() : "";
    if (!clinicName) continue;

    const shiftDateValueRaw = row[1];
    const shiftDateObj = parseDateToSafeDateObj(shiftDateValueRaw);
    if (!shiftDateObj) {
      Logger.log(`行 ${i + 1} (${clinicName}): 勤務日のパースに失敗。元値: "[${shiftDateValueRaw}]"`);
      continue;
    }
    if (shiftDateObj < startDate || shiftDateObj > endDate) continue;

    const dateKey = fastFormatDate(shiftDateObj);
    const formattedShiftDateForTitle = `${shiftDateObj.getMonth() + 1}月${shiftDateObj.getDate()}日（${weekdaysJP[shiftDateObj.getDay()]}）`;

    if (!dailyReportData[dateKey]) {
      dailyReportData[dateKey] = {
        titleDate: formattedShiftDateForTitle,
        backupText: "",
        unfilledList: [],
        hasNonBackupClinicsProcessed: false
      };
    }

    if (clinicName === KANTO_BACKUP_SHIFT_NAME_CONST) {
      const timeSlotsOriginal = ["09:00~13:00", "15:00~18:00", "18:00~21:00"];
      let backupDoctors = [];
      for (let j = 0; j < timeSlotsOriginal.length; j++) {
        const doctorNames = row[7 + j];
        if (doctorNames && doctorNames.toString().trim() !== "") {
          backupDoctors.push(`${timeSlotsOriginal[j]}：${doctorNames}先生（全拠点）`);
        }
      }
      if (backupDoctors.length > 0) {
        dailyReportData[dateKey].backupText = `【バックアップ】${backupDoctors.join('、')}`;
      }
    } else {
      dailyReportData[dateKey].hasNonBackupClinicsProcessed = true;

      let entryUnfilledThisTime = false;
      const timeSlotsDefinition = [
        { name: "09:00~13:00", conditionColIdx: 3, doctorColIdx: 7 },
        { name: "13:00~18:00", conditionColIdx: 4, doctorColIdx: 8 },
        { name: "18:00~21:00", conditionColIdx: 5, doctorColIdx: 9 }
      ];

      for (const slot of timeSlotsDefinition) {
        const numberOfPeopleCell = row[slot.conditionColIdx];
        const doctorNameCell = row[slot.doctorColIdx];
        const numberOfPeople = (typeof numberOfPeopleCell === 'number') ? numberOfPeopleCell : parseFloat(numberOfPeopleCell);
        const isDoctorSlotEmpty = (doctorNameCell === null || doctorNameCell === undefined || doctorNameCell.toString().trim() === "");

        if (numberOfPeople === 0 && isDoctorSlotEmpty) {
          let unfilledEntryText = `【${clinicName}】${slot.name}`;
          const department = row[2] ? row[2].toString().trim() : "";
          if ((clinicName === "北葛西" || clinicName === "亀有") && department !== "") {
            unfilledEntryText = `【${clinicName}】${slot.name} (${department})`;
          }

          if (!dailyReportData[dateKey].unfilledList.includes(unfilledEntryText)) {
            dailyReportData[dateKey].unfilledList.push(unfilledEntryText);
          }
          entryUnfilledThisTime = true;
        }
      }
    }
  }

  const sortedDateKeys = Object.keys(dailyReportData).sort();

  if (sortedDateKeys.length === 0) {
    message += "対象期間内に報告すべき未充足情報はありませんでした。\n";
  }

  for (const dateKey of sortedDateKeys) {
    const dayData = dailyReportData[dateKey];
    let entry = `[info][title]${dayData.titleDate}[/title]`;
    if (dayData.backupText) {
      entry += dayData.backupText;
    }
    entry += "[hr]";
    if (dayData.unfilledList.length > 0) {
      entry += dayData.unfilledList.join('\n') + '\n';
    } else if (dayData.hasNonBackupClinicsProcessed) {
      entry += "充足\n";
    }
    entry += "[/info]\n";
    message += entry;
  }

  targetSheet.getRange('A6').setValue(message);
  Logger.log('文章を A6 に書き出しました。スクリプト完了');
}