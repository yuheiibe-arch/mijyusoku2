// 関数名は前回提示したものをベースにしていますが、適宜実際の関数名に読み替えてください。
// 区別のため、関数名の末尾に "_v3" を追加しました。
function updateUnfilledStatusWithClinicLogicAndReset_v3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unfulfilledSheet = ss.getSheetByName("未充足管理");
  const anabukiSheet = ss.getSheetByName("貼付用");

  if (!unfulfilledSheet || !anabukiSheet) {
    Logger.log("エラー: 「未充足管理」シートまたは「貼付用」シートが見つかりません。");
    SpreadsheetApp.getUi().alert("エラー: 「未充足管理」シートまたは「貼付用」シートが見つかりません。");
    return;
  }

  const scriptTimeZone = Session.getScriptTimeZone();
  const currentTime = new Date(); // タイムスタンプ用

  // --- ヘルパー関数定義 ---
  function formatDateForCompare(dateObj) {
    if (dateObj instanceof Date && !isNaN(dateObj.getTime())) {
      return Utilities.formatDate(dateObj, scriptTimeZone, "yyyy-MM-dd");
    }
    return null;
  }

  function formatTimeForCompare(timeValue) {
    if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
      return Utilities.formatDate(timeValue, scriptTimeZone, "HH:mm");
    }
    if (typeof timeValue === 'string') {
        const trimmedTime = timeValue.trim();
        if (trimmedTime.match(/^\d{1,2}:\d{2}(:\d{2})?$/)) { // HH:mm or HH:mm:ss
            const parts = trimmedTime.split(":");
            return `${parts[0].padStart(2, '0')}:${parts[1]}`; // HH:mm に正規化
        }
    }
    if (typeof timeValue === 'number') { // スプレッドシートの時刻シリアル値の場合
        const dateFromSerial = new Date(1899, 11, 30 + timeValue);
        return Utilities.formatDate(dateFromSerial, scriptTimeZone, "HH:mm");
    }
    return null;
  }
  
  function cleanString(str) { 
    if (typeof str !== 'string' && typeof str !== 'number') return "";
    return String(str).trim().normalize("NFKC");
  }

  function parseLocationAndClinic(rawLocationStr) {
    let cleanedLoc = cleanString(rawLocationStr);
    let baseLocation = cleanedLoc;
    let clinic = "小児科"; 

    const clinicMatch = cleanedLoc.match(/(.+?)\s*（(小児科|内科)）$/);
    if (clinicMatch) {
      baseLocation = cleanString(clinicMatch[1]);
      clinic = cleanString(clinicMatch[2]); 
    } else {
      baseLocation = cleanedLoc; 
    }
    return { baseLocation: baseLocation, clinic: clinic };
  }
  
  function parseAnabukiDate(dateString) {
    if (!dateString) return null;
    const cleanedDateString = String(dateString).trim();

    const parts = cleanedDateString.match(/^(\d{2,4})[/\-](\d{1,2})[/\-](\d{1,2})$/);
    if (parts) {
      let year = parseInt(parts[1], 10);
      const month = parseInt(parts[2], 10) - 1; 
      const day = parseInt(parts[3], 10);
      if (year < 100 && year >= 0) { 
        year += 2000; 
      }
      if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
        return new Date(year, month, day);
      }
    }
    const directDate = new Date(cleanedDateString);
    if (directDate instanceof Date && !isNaN(directDate.getTime())) {
        return directDate;
    }
    return null;
  }
  // --- ヘルパー関数定義ここまで ---


  // --- 1. 「貼付用」シートのデータを処理しやすい形式に格納 ---
  const anabukiMap = new Map(); 
  const anabukiHeaderRows = 1; 
  const anabukiLastRow = anabukiSheet.getLastRow();
  
  if (anabukiLastRow > anabukiHeaderRows) {
    const anabukiData = anabukiSheet.getRange(anabukiHeaderRows + 1, 1, anabukiLastRow - anabukiHeaderRows, 20).getValues(); 

    anabukiData.forEach(row => {
      const doctorName = cleanString(row[0]);          
      const anabukiBaseLocRaw = row[12];               
      const anabukiClinicRaw = row[13];                
      const anabukiDateStr = row[14];                  
      const anabukiStartTimeRaw = row[15];             
      const anabukiEndTimeRaw = row[19];               

      if (doctorName && anabukiDateStr && anabukiBaseLocRaw && anabukiStartTimeRaw && anabukiEndTimeRaw) {
        const anabukiDateObj = parseAnabukiDate(anabukiDateStr);
        const formattedAnabukiDate = formatDateForCompare(anabukiDateObj);
        
        let cleanedAnabukiBaseLocation = cleanString(anabukiBaseLocRaw);
        let keyClinic;

        if (cleanedAnabukiBaseLocation === "北葛西" || cleanedAnabukiBaseLocation === "亀有") {
          keyClinic = cleanString(anabukiClinicRaw); 
          if (keyClinic !== "小児科" && keyClinic !== "内科") {
            return; 
          }
        } else {
          keyClinic = "小児科"; 
          const anabukiNColClinic = cleanString(anabukiClinicRaw);
          if (anabukiNColClinic !== "小児科" && anabukiNColClinic !== "") {
            return; 
          }
        }
        
        const formattedAnabukiStartTime = formatTimeForCompare(anabukiStartTimeRaw);
        const formattedAnabukiEndTime = formatTimeForCompare(anabukiEndTimeRaw);

        if (formattedAnabukiDate && cleanedAnabukiBaseLocation && keyClinic && formattedAnabukiStartTime && formattedAnabukiEndTime) {
          const key = `${formattedAnabukiDate}_${cleanedAnabukiBaseLocation}_${keyClinic}_${formattedAnabukiStartTime}_${formattedAnabukiEndTime}`;
          if (!anabukiMap.has(key)) { 
            anabukiMap.set(key, doctorName);
          }
        }
      }
    });
  }
  Logger.log(`貼付用データMap作成完了。${anabukiMap.size}件のマッチング用エントリ。`);

  // --- 2. 「未充足管理」シートのデータを処理 ---
  const UNFULFILLED_START_ROW = 6; 
  const unfulfilledLastRow = unfulfilledSheet.getLastRow();

  if (unfulfilledLastRow < UNFULFILLED_START_ROW) {
    Logger.log("「未充足管理」シートに処理対象データがありません。");
    return;
  }

  const unfulfilledDataRange = unfulfilledSheet.getRange(UNFULFILLED_START_ROW, 1, unfulfilledLastRow - UNFULFILLED_START_ROW + 1, 21); 
  const unfulfilledData = unfulfilledDataRange.getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0); 

  let updatedCount = 0;
  let hiddenCount = 0; // ★追加: 非表示にした行数をカウントする変数

  for (let i = 0; i < unfulfilledData.length; i++) {
    const currentRowInSheet = UNFULFILLED_START_ROW + i;
    const rowData = unfulfilledData[i];

    const unfulfilledDateValue = rowData[19]; // T列の日付
    const unfulfilledFullLocationRaw = rowData[6];  // G列
    const unfulfilledStartTimeRaw = rowData[7];    // H列
    const unfulfilledEndTimeRaw = rowData[8];      // I列
    
    const initialCstatus = cleanString(rowData[2]); // C列
    const initialJdoctor = cleanString(rowData[9]); // J列
    const initialMstatus = cleanString(rowData[12]);// M列

    // T列の日付が有効かチェック
    if (!(unfulfilledDateValue instanceof Date && !isNaN(unfulfilledDateValue.getTime()))) {
      Logger.log(`行 ${currentRowInSheet}: T列の日付が無効なためスキップします。`);
      continue; 
    }
    
    const unfulfilledDateObj = new Date(unfulfilledDateValue.getTime()); 
    unfulfilledDateObj.setHours(0,0,0,0); // 時刻部分をクリアして日付のみで比較

    // ▼▼▼ ★修正箇所: 非表示処理ブロック (以前の if (unfulfilledDateObj < today) を置き換え) ▼▼▼
    if (unfulfilledDateObj <= today) { // 条件: 本日以前
      unfulfilledSheet.hideRows(currentRowInSheet);
      hiddenCount++; // 非表示カウントを増やす
      Logger.log(`行 ${currentRowInSheet} を非表示にしました (T列日付: ${formatDateForCompare(unfulfilledDateObj)} が本日以前のため)。`);
      continue; // この行の以降の処理（ステータス更新など）は行わない
    }
    // ▲▲▲ ★修正箇所ここまで ▲▲▲
        
    // --- 日付が本日より後の場合の既存のステータス更新処理 ---
    const parsedUnfulfilledInfo = parseLocationAndClinic(unfulfilledFullLocationRaw);
    const baseUnfulfilledLocation = parsedUnfulfilledInfo.baseLocation;
    let clinicToMatch = parsedUnfulfilledInfo.clinic; 

    if (baseUnfulfilledLocation !== "北葛西" && baseUnfulfilledLocation !== "亀有") {
      clinicToMatch = "小児科";
    }

    const formattedUnfulfilledDate = formatDateForCompare(unfulfilledDateObj);
    const formattedUnfulfilledStartTime = formatTimeForCompare(unfulfilledStartTimeRaw);
    const formattedUnfulfilledEndTime = formatTimeForCompare(unfulfilledEndTimeRaw);

    let targetCstatus = initialCstatus; 
    let targetJdoctor = initialJdoctor; 
    let targetMstatus = initialMstatus; 
    let madeChange = false;

    if (formattedUnfulfilledDate && baseUnfulfilledLocation && clinicToMatch && formattedUnfulfilledStartTime && formattedUnfulfilledEndTime) {
      const keyToSearch = `${formattedUnfulfilledDate}_${baseUnfulfilledLocation}_${clinicToMatch}_${formattedUnfulfilledStartTime}_${formattedUnfulfilledEndTime}`;
      
      if (anabukiMap.has(keyToSearch)) { 
        const doctorToFill = anabukiMap.get(keyToSearch);
        targetCstatus = "充足";     
        targetJdoctor = doctorToFill; 
        targetMstatus = "応募有（充足）"; 
      } else { 
        if (initialCstatus === "充足") {
          targetCstatus = ""; 
          targetJdoctor = ""; 
          targetMstatus = ""; // ★★★ 変更: M列もクリア（空文字列に設定）★★★ (このコメントは元々ありました)
        }
      }

      if (initialCstatus !== targetCstatus) {
        unfulfilledSheet.getRange(currentRowInSheet, 3).setValue(targetCstatus); 
        madeChange = true;
      }
      if (initialJdoctor !== targetJdoctor) {
        unfulfilledSheet.getRange(currentRowInSheet, 10).setValue(targetJdoctor); 
        madeChange = true;
      }
      if (initialMstatus !== targetMstatus) {
        unfulfilledSheet.getRange(currentRowInSheet, 13).setValue(targetMstatus); 
        madeChange = true;
      }

      if (madeChange) {
        unfulfilledSheet.getRange(currentRowInSheet, 21).setValue(currentTime); 
        updatedCount++;
        Logger.log(`行 ${currentRowInSheet}: 更新。C="${targetCstatus}", J="${targetJdoctor}", M="${targetMstatus}". キー=${keyToSearch}`);
      }
    } else {
      // Logger.log(`行 ${currentRowInSheet}: 未充足データのキー生成に必要な情報が不足または不正です。`);
    }
  } // End of for loop

  // ★修正箇所: 完了メッセージに非表示件数を追加
  Logger.log(`処理完了。${updatedCount}行に更新を行い、${hiddenCount}行を非表示にしました。`);
  SpreadsheetApp.getUi().alert(`処理完了。${updatedCount}行に更新を行い、${hiddenCount}行を非表示にしました。`);
}