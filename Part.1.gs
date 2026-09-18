// メインの更新関数
function updateSheetRowAdjusted_CallingCellSpecificFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('貼付用');
  const targetSheet = ss.getSheetByName('確認用');
  
  // ★ 詳細ログを false にして空行の大量ログをストップし、処理を高速化
  const DETAILED_LOGGING = false; 

  if (!sourceSheet || !targetSheet) {
    ss.toast('貼付用 または 確認用 シートが見つかりません。', 'エラー', 5);
    return;
  }

  ss.toast('シフト集計処理を開始します...', '処理開始', 3);

  const localExcludedLocations = (typeof GLOBAL_EXCLUDED_LOCATIONS !== 'undefined') 
    ? GLOBAL_EXCLUDED_LOCATIONS.filter(item => item !== "【関東】バックアップシフト")
    : [];

  const data = sourceSheet.getDataRange().getValues();
  const rows = data.slice(2);

  if (rows.length === 0) {
    ss.toast('貼付用シートの3行目以降にデータが見つかりません。', 'エラー', 5);
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
    
    // 日付も拠点名も空っぽの「完全な空行」は静かにスキップ
    if (!row[0] && !row[12] && !row[14]) {
      return;
    }

    const doctorName = row[0] ? row[0].toString().trim() : '未設定';
    const originalClinicName = row[12] ? row[12].toString().trim() : null;
    const originalDepartment = row[13] ? row[13].toString().trim() : null;
    const shiftDateRaw = row[14];
    
    const aShiftValue = row[63]; // BL列
    const bShiftValue = row[64]; // BM列
    const cShiftValue = row[65]; // BN列

    if ((originalClinicName && localExcludedLocations.includes(originalClinicName)) ||
        (originalDepartment && EXCLUDED_DEPARTMENTS.includes(originalDepartment))) {
      if (DETAILED_LOGGING) Logger.log(`行 ${rowIndex}: スキップ (理由: 除外項目該当)`);
      return;
    }
    
    if (!originalClinicName || !shiftDateRaw || !originalDepartment) {
      if (DETAILED_LOGGING) Logger.log(`行 ${rowIndex}: スキップ (必須項目不足)`);
      return;
    }
    
    let shiftDateObj, shiftDateKey;
    try {
      shiftDateObj = parseDateToSafeDateObj(shiftDateRaw); 
      if (!shiftDateObj || isNaN(shiftDateObj.getTime())) return;
      shiftDateKey = fastFormatDate(shiftDateObj).replace(/\//g, '-'); 
    } catch (e) {
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
        const yyyymmdd = fastFormatDate(record.shiftDate);
        const weekday = weekdays[record.shiftDate.getDay()];
        formattedShiftDate = `${yyyymmdd}（${weekday}）`;
      }
    } catch (e) { formattedShiftDate = 'フォーマットエラー'; }
    
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
  
  // 別のファイルに分けた関数をここで呼び出し
  try { applyConditionalFormatting_CellSpecific(); } catch (e) {}
  try {
    if (typeof generateDoctorAbsenceReportWithContext === 'function') {
        generateDoctorAbsenceReportWithContext();
    }
  } catch (e) {
    ss.toast(`医師不在拠点シートへの書き出し中にエラーが発生しました: ${e.message}`, 'エラー', 5);
  }
  try { setupDateSelection(); } catch (e) {}
  
  ss.toast('本番シフト集計が完了しました。', '完了', 3);
}