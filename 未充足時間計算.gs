/**
 * 【新規追加】
 * 「貼付用」シートから正確な隙間時間（例: 17:00-18:00）を計算し、
 * 「不在書き出し」シートに上書き出力する関数
 */
function exportPreciseAbsenceToNewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("貼付用");
  let targetSheet = ss.getSheetByName("不在書き出し");

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("エラー: 「貼付用」シートが見つかりません。");
    return;
  }
  if (!targetSheet) {
    // シートがなければ作成
    targetSheet = ss.insertSheet("不在書き出し");
  }

  // シートをクリア（ヘッダー再設定）
  targetSheet.clear();
  const header = ["日付", "拠点名", "正確な不在時間", "備考"];
  targetSheet.getRange(1, 1, 1, header.length).setValues([header]);

  // --- シフト設定 (グローバル定数があればそれを利用、なければここで定義) ---
  const shiftTimesMin = (typeof GLOBAL_SHIFT_TIMES_MIN !== 'undefined') ? GLOBAL_SHIFT_TIMES_MIN : {
      "その他": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] }
  };
  const standardShiftOrder = ["A", "B", "C"];
  const excludedLocations = (typeof GLOBAL_EXCLUDED_LOCATIONS !== 'undefined') ? GLOBAL_EXCLUDED_LOCATIONS : [];
  const excludedDepartments = (typeof GLOBAL_EXCLUDED_DEPARTMENTS !== 'undefined') ? GLOBAL_EXCLUDED_DEPARTMENTS : [];

  // --- データ読み込み ---
  const sourceData = sourceSheet.getDataRange().getValues();
  const processedWorkData = {}; 

  for (let i = 2; i < sourceData.length; i++) { // 3行目から
    const row = sourceData[i];
    const clinicName = row[12] ? String(row[12]).trim() : "";
    const department = row[13] ? String(row[13]).trim() : "";
    const dateRaw = row[14];
    
    if (!dateRaw || !clinicName) continue;
    if (excludedLocations.includes(clinicName) || excludedDepartments.includes(department)) continue;

    const dateObj = parseDateToSafeDateObj(dateRaw); // 既存のヘルパー関数を利用
    if (!dateObj) continue;

    const dateKey = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const workStartMin = parseTimeToMinutes(row[15]); // 既存のヘルパー関数を利用
    const workEndMin = parseTimeToMinutes(row[19]);   // 既存のヘルパー関数を利用

    if (isNaN(workStartMin) || isNaN(workEndMin) || workStartMin >= workEndMin) continue;

    const aggKey = `${clinicName}_${department}`;

    if (!processedWorkData[dateKey]) processedWorkData[dateKey] = {};
    if (!processedWorkData[dateKey][aggKey]) processedWorkData[dateKey][aggKey] = [];
    processedWorkData[dateKey][aggKey].push({ start: workStartMin, end: workEndMin });
  }

  // --- 隙間時間の計算と出力 ---
  const outputRows = [];
  const sortedDates = Object.keys(processedWorkData).sort();
  const today = new Date();
  today.setHours(0,0,0,0);

  for (const dateKey of sortedDates) {
    const dateObj = new Date(dateKey);
    if (dateObj < today) continue; 

    const dailyData = processedWorkData[dateKey];
    
    for (const aggKey in dailyData) {
      const [clinicName, department] = aggKey.split('_');
      const workIntervals = dailyData[aggKey];

      // シフト時間を特定
      let shiftLookupKey = "その他";
      const specificKey = `${clinicName}${department}`;
      if (shiftTimesMin[specificKey]) shiftLookupKey = specificKey;
      else if (shiftTimesMin[clinicName]) shiftLookupKey = clinicName;
      
      const currentClinicShifts = shiftTimesMin[shiftLookupKey] || shiftTimesMin["その他"];

      // 各シフト枠についてチェック
      for (const shiftKey of standardShiftOrder) {
        if (!currentClinicShifts[shiftKey]) continue;
        const [slotStart, slotEnd] = currentClinicShifts[shiftKey];

        // この枠に関係する勤務だけ抽出
        const relevantIntervals = workIntervals
          .map(w => ({ start: Math.max(w.start, slotStart), end: Math.min(w.end, slotEnd) }))
          .filter(w => w.start < w.end);
        
        // 勤務時間を結合
        const merged = mergeIntervals(relevantIntervals); // 既存のヘルパー関数を利用

        // ★隙間を計算
        let pointer = slotStart;
        let gaps = [];

        for (const m of merged) {
          if (pointer < m.start) {
            gaps.push(`${formatMinutesToHHMM(pointer)}-${formatMinutesToHHMM(m.start)}`);
          }
          pointer = Math.max(pointer, m.end);
        }
        if (pointer < slotEnd) {
          gaps.push(`${formatMinutesToHHMM(pointer)}-${formatMinutesToHHMM(slotEnd)}`);
        }

        // 隙間があれば出力リストに追加
        if (gaps.length > 0) {
          let displayName = clinicName;
          // 必要に応じて科名を追加
          if ((clinicName === "北葛西" || clinicName === "亀有") && (department === "小児科" || department === "内科")) {
             displayName = `${clinicName}（${department}）`;
          }

          const preciseTimeStr = gaps.join(", ");
          outputRows.push([dateKey, displayName, preciseTimeStr, shiftKey]);
        }
      }
    }
  }

  // 書き出し
  if (outputRows.length > 0) {
    outputRows.sort((a, b) => {
        if (a[0] !== b[0]) return new Date(a[0]) - new Date(b[0]);
        return a[1].localeCompare(b[1]);
    });
    targetSheet.getRange(2, 1, outputRows.length, header.length).setValues(outputRows);
    Logger.log(`「不在書き出し」シートを更新しました (${outputRows.length}件)`);
  }
}