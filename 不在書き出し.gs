/**
 * 正確な不在時間（隙間時間）を計算し、「不在書き出し」シートに一括出力する関数
 * - 既存の「医師不在拠点」シートには影響を与えません。
 * - 実行のたびに「不在書き出し」シートはクリアされ、再作成されます。
 */
function exportPreciseAbsenceToNewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("貼付用");
  // ▼ 新しい出力先シート
  let targetSheet = ss.getSheetByName("不在書き出し");

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("エラー: 「貼付用」シートが見つかりません。");
    return;
  }
  if (!targetSheet) {
    // シートがなければ作成する安全策
    targetSheet = ss.insertSheet("不在書き出し");
  }

  // シートをクリア（ヘッダーも作り直すため全クリア）
  targetSheet.clear();
  
  // ヘッダー設定
  const header = ["日付", "拠点名", "正確な不在時間", "備考(元シフト枠)"];
  targetSheet.getRange(1, 1, 1, header.length).setValues([header]);

  // --- 設定読み込み ---
  // ※グローバル定数が定義されている前提です。未定義なら直書きしてください。
  const shiftTimesMin = (typeof GLOBAL_SHIFT_TIMES_MIN !== 'undefined') ? GLOBAL_SHIFT_TIMES_MIN : {
      "その他": { A: [9 * 60, 13 * 60], B: [15 * 60, 18 * 60], C: [18 * 60, 21 * 60] }
  };
  const standardShiftOrder = ["A", "B", "C"];
  
  // 除外リスト
  const excludedLocations = (typeof GLOBAL_EXCLUDED_LOCATIONS !== 'undefined') ? GLOBAL_EXCLUDED_LOCATIONS : [];
  const excludedDepartments = (typeof GLOBAL_EXCLUDED_DEPARTMENTS !== 'undefined') ? GLOBAL_EXCLUDED_DEPARTMENTS : [];

  // データ取得
  const sourceData = sourceSheet.getDataRange().getValues();
  const processedWorkData = {}; // { "yyyy/MM/dd": { "拠点_科": [ {start:900, end:1000}, ... ] } }

  // 1. 貼付用データを読み込んで整理
  for (let i = 2; i < sourceData.length; i++) { // 3行目から
    const row = sourceData[i];
    const clinicName = row[12] ? String(row[12]).trim() : "";
    const department = row[13] ? String(row[13]).trim() : "";
    const dateRaw = row[14];
    
    if (!dateRaw || !clinicName) continue;
    if (excludedLocations.includes(clinicName) || excludedDepartments.includes(department)) continue;

    const dateObj = parseDateToSafeDateObj(dateRaw);
    if (!dateObj) continue;

    const dateKey = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const workStartMin = parseTimeToMinutes(row[15]);
    const workEndMin = parseTimeToMinutes(row[19]);

    if (isNaN(workStartMin) || isNaN(workEndMin) || workStartMin >= workEndMin) continue;

    const aggKey = `${clinicName}_${department}`; // 集計キー

    if (!processedWorkData[dateKey]) processedWorkData[dateKey] = {};
    if (!processedWorkData[dateKey][aggKey]) processedWorkData[dateKey][aggKey] = [];
    
    processedWorkData[dateKey][aggKey].push({ start: workStartMin, end: workEndMin });
  }

  // 2. 隙間時間の計算と出力データの生成
  const outputRows = [];
  const sortedDates = Object.keys(processedWorkData).sort();
  const today = new Date();
  today.setHours(0,0,0,0);

  for (const dateKey of sortedDates) {
    const dateObj = new Date(dateKey);
    if (dateObj < today) continue; // 過去は無視

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

      // 各シフト枠（A, B, C）についてチェック
      for (const shiftKey of standardShiftOrder) {
        if (!currentClinicShifts[shiftKey]) continue;
        const [slotStart, slotEnd] = currentClinicShifts[shiftKey];

        // この枠に関係する勤務だけを抽出して結合
        // (枠からはみ出している部分はカットして考える)
        const relevantIntervals = workIntervals
          .map(w => ({ start: Math.max(w.start, slotStart), end: Math.min(w.end, slotEnd) }))
          .filter(w => w.start < w.end); // 有効なものだけ
        
        const merged = mergeIntervals(relevantIntervals);

        // 隙間を計算 (ここが肝)
        let pointer = slotStart;
        let gaps = [];

        for (const m of merged) {
          if (pointer < m.start) {
            // 勤務開始前の隙間
            gaps.push(`${formatMinutesToHHMM(pointer)}-${formatMinutesToHHMM(m.start)}`);
          }
          pointer = Math.max(pointer, m.end);
        }
        if (pointer < slotEnd) {
          // 勤務終了後の隙間
          gaps.push(`${formatMinutesToHHMM(pointer)}-${formatMinutesToHHMM(slotEnd)}`);
        }

        // 隙間があれば出力行に追加
        if (gaps.length > 0) {
          // 北葛西・亀有などで科を表示するかどうかの判定
          let displayName = clinicName;
          if ((clinicName === "北葛西" || clinicName === "亀有") && (department === "小児科" || department === "内科")) {
             displayName = `${clinicName}（${department}）`;
          }

          // 複数の隙間がある場合（中抜けなど）は結合
          const preciseTimeStr = gaps.join(", ");
          const slotNameStr = `${shiftKey}シフト(${formatMinutesToHHMM(slotStart)}~${formatMinutesToHHMM(slotEnd)})`;
          
          outputRows.push([dateKey, displayName, preciseTimeStr, slotNameStr]);
        }
      }
    }
  }

  // 3. 書き出し
  if (outputRows.length > 0) {
    // 日付順 > 拠点順 でソート
    outputRows.sort((a, b) => {
        if (a[0] !== b[0]) return new Date(a[0]) - new Date(b[0]);
        return a[1].localeCompare(b[1]);
    });
    targetSheet.getRange(2, 1, outputRows.length, header.length).setValues(outputRows);
    Logger.log(`「不在書き出し」シートに ${outputRows.length} 件出力しました。`);
  } else {
    Logger.log("「不在書き出し」: 該当する不在はありませんでした。");
  }
}