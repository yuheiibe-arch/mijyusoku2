/**
 * メイン処理: Chatwork用メッセージ生成
 */
function generateChatworkMessage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi(); 
  const scriptTimeZone = Session.getScriptTimeZone(); 
  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"]; 

  // --- 追加データソースの読み込みと事前グループ化（爆速化） ---
  const EXT_SS_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co';
  const extDataByDate = {}; 
  let isExtSsLoaded = false;
  const initExtDate = (dStr) => {
    if (!extDataByDate[dStr]) extDataByDate[dStr] = { paste: [], oubo: [], bosyu: [] };
  };

  try {
    const extSs = SpreadsheetApp.openById(EXT_SS_ID);
    isExtSsLoaded = true;
    
    const pasteSheet = ss.getSheetByName('貼付用');
    if (pasteSheet && pasteSheet.getLastRow() > 2) {
      const pasteData = pasteSheet.getDataRange().getValues();
      for (let i = 2; i < pasteData.length; i++) {
        const dStr = fastFormatDate(parseDateToSafeDateObj(pasteData[i][14]) || pasteData[i][14]);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].paste.push(pasteData[i]); }
      }
    }

    const ouboSheet = extSs.getSheetByName('応募シフト');
    if (ouboSheet && ouboSheet.getLastRow() > 1) {
      const ouboData = ouboSheet.getDataRange().getValues();
      for (let i = 1; i < ouboData.length; i++) {
        const dStr = fastFormatDate(ouboData[i][5]);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].oubo.push(ouboData[i]); }
      }
    }

    const bosyuSheet = extSs.getSheetByName('募集シフト');
    if (bosyuSheet && bosyuSheet.getLastRow() > 1) {
      const bosyuData = bosyuSheet.getDataRange().getValues();
      for (let i = 1; i < bosyuData.length; i++) {
        const dStr = fastFormatDate(bosyuData[i][3]);
        if (dStr) { initExtDate(dStr); extDataByDate[dStr].bosyu.push(bosyuData[i]); }
      }
    }
  } catch (e) {
    Logger.log('追加データソースの読み込みに失敗しました: ' + e.message);
  }

  // --- ★追加: COO室依頼２診要望データの読み込み ---
  const COO_SS_ID = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk';
  const cooDataByDate = {};
  try {
    const cooSS = SpreadsheetApp.openById(COO_SS_ID);
    const cooSheet = cooSS.getSheetByName("２診要望一覧");
    if (cooSheet && cooSheet.getLastRow() > 1) {
      const cooRaw = cooSheet.getDataRange().getDisplayValues().slice(1);
      cooRaw.forEach(row => {
        const clinic = String(row[0] || "").replace(/[（(]小児科[）)]/, "").trim();
        const dStrRaw = String(row[1] || "").trim().split(/[（(]/)[0].trim().replace(/-/g, '/');
        const parts = dStrRaw.split('/');
        if (parts.length >= 3) {
          const dObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          if (!isNaN(dObj.getTime())) {
            const dStr = fastFormatDate(dObj);
            if (!cooDataByDate[dStr]) cooDataByDate[dStr] = [];
            cooDataByDate[dStr].push(row);
          }
        }
      });
    }
  } catch (e) {
    Logger.log('COO要望一覧の読み込みに失敗しました: ' + e.message);
  }

  // --- シート取得 ---
  const sourceSheet = ss.getSheetByName('確認用');
  const targetSheet = ss.getSheetByName('文章自動作成');
  const mentionSheet = ss.getSheetByName('メンション先選択');
  const ishiFuzaiSheet = ss.getSheetByName('医師不在拠点');
  const closedSheet = ss.getSheetByName('休館日'); 
  const irregularSheet = ss.getSheetByName('変則営業');

  if (!sourceSheet || !targetSheet || !mentionSheet || !ishiFuzaiSheet) { 
    if (ui) ui.alert('エラー: 必要なシートが見つかりません。'); 
    return; 
  }

  // --- 🌟UI日付の取得とデフォルト設定ロジック ---
  const baseDate = new Date();
  if (baseDate.getHours() >= 15) baseDate.setDate(baseDate.getDate() + 1);
  const endDateCalc = new Date(baseDate);
  endDateCalc.setDate(baseDate.getDate() + 6); 

  const formattedStart = fastFormatDate(baseDate) + `（${weekdaysJP[baseDate.getDay()]}）`;
  const formattedEnd = fastFormatDate(endDateCalc) + `（${weekdaysJP[endDateCalc.getDay()]}）`;

  let uniqueValues = [];
  if (sourceSheet.getLastRow() >= 2) {
    const bColumnValues = sourceSheet.getRange(2, 2, sourceSheet.getLastRow() - 1, 1).getValues().flat();
    uniqueValues = [...new Set(bColumnValues.filter(Boolean))].map(dateStr => {
      const dateObj = parseDateToSafeDateObj(dateStr);
      return !dateObj ? dateStr : fastFormatDate(dateObj) + `（${weekdaysJP[dateObj.getDay()]}）`;
    });
  }
  if (!uniqueValues.includes(formattedStart)) uniqueValues.unshift(formattedStart);
  if (!uniqueValues.includes(formattedEnd)) uniqueValues.push(formattedEnd);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(uniqueValues, true).setAllowInvalid(true).build();
  
  const cellB2 = targetSheet.getRange('B2');
  const cellB4 = targetSheet.getRange('B4');
  
  let startDateRaw = cellB2.getValue();
  let endDateRaw = cellB4.getValue();
  
  if (!startDateRaw) {
    cellB2.setValue(formattedStart);
    startDateRaw = formattedStart;
  }
  if (!endDateRaw) {
    cellB4.setValue(formattedEnd);
    endDateRaw = formattedEnd;
  }
  
  cellB2.setDataValidation(rule);
  cellB4.setDataValidation(rule);

  const startDate = parseDateToSafeDateObj(startDateRaw);
  const endDate = parseDateToSafeDateObj(endDateRaw);

  if (!startDate || !endDate || startDate > endDate) {
    if (ui) ui.alert(`日付指定が無効です。\n開始: ${startDateRaw}\n終了: ${endDateRaw}`); 
    return; 
  }

  // --- 開院日マスタ・エリアマスタの読み込み ---
  const openDateMap = new Map();
  const areaMap = {}; 
  try {
    const masterSS = SpreadsheetApp.openById('14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs');
    const extData = masterSS.getSheetByName('拠点名').getDataRange().getValues();
    
    for (let i = 1; i < extData.length; i++) {
      const groupName = extData[i][5] ? String(extData[i][5]).trim() : "";
      let area = "その他";
      if (groupName.includes("関東第一") || groupName.includes("関東第二") || groupName.includes("東京第一") || groupName.includes("東京第二")) area = "東京";
      else if (groupName.includes("神奈川")) area = "神奈川";
      else if (groupName.includes("埼玉")) area = "埼玉";
      else if (groupName.includes("千葉")) area = "千葉";
      else if (groupName.includes("茨城")) area = "茨城";
      else if (groupName.includes("大阪") || groupName.includes("関西")) area = "大阪";

      const openDate = extData[i][7]; 
      const namesToRegister = [extData[i][0], extData[i][1], extData[i][2], extData[i][3], extData[i][4]];
      
      namesToRegister.forEach(name => {
        const cleanName = name ? String(name).trim() : "";
        if (cleanName !== "") {
          if (openDate instanceof Date) openDateMap.set(cleanName, openDate);
          areaMap[cleanName] = area;
        }
      });
    }
  } catch (e) {
    Logger.log('開院日マスタの読み込みに失敗: ' + e.message);
  }

  // --- 0. 変則営業データの読み込み ---
  const irregularMap = {}; 
  if (irregularSheet) {
    const iData = irregularSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      const row = iData[i];
      const dateObj = parseDateToSafeDateObj(row[0]); 
      const timeRangeStr = row[1];
      const clinicName = row[2] ? String(row[2]).trim() : "";

      if (dateObj && timeRangeStr && clinicName) {
        const dateKey = fastFormatDate(dateObj);
        const rangesStr = String(timeRangeStr).split(/[,/、]/);
        const validRanges = [];
        for (let rStr of rangesStr) {
          const parts = rStr.split('-');
          if (parts.length === 2) {
            const openMin = safeParseTime(parts[0]);
            const closeMin = safeParseTime(parts[1]);
            if (!isNaN(openMin) && !isNaN(closeMin) && openMin < closeMin) {
              validRanges.push({ open: openMin, close: closeMin });
            }
          }
        }
        
        if (validRanges.length > 0) {
           const normName = normalizeClinicName(clinicName);
           if (normName === "全拠点") {
               irregularMap[`${dateKey}_全拠点`] = validRanges;
           } else {
               irregularMap[`${dateKey}_${normName}`] = validRanges;
               irregularMap[`${dateKey}_${clinicName}`] = validRanges;
           }
        }
      }
    }
  }

  // --- 1. 休館日データの読み込み ---
  const closedDataMap = new Map();
  if (closedSheet) {
    const cData = closedSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      const row = cData[i];
      const dateObj = parseDateToSafeDateObj(row[0]);
      const cDept = row[2] ? String(row[2]).trim() : "";
      const cLoc = row[3] ? String(row[3]).trim() : "";
      const cTime = row[4] ? String(row[4]).trim() : "全日"; 
      
      if (dateObj && cLoc) {
        const dateKey = fastFormatDate(dateObj);
        if (cLoc === "全拠点") {
             closedDataMap.set(`${dateKey}_全拠点`, cTime);
        } else if (cDept === "内科") {
             closedDataMap.set(`${dateKey}_${normalizeClinicName(cLoc)}_内科`, cTime);
             closedDataMap.set(`${dateKey}_${cLoc}_内科`, cTime);
        } else {
             closedDataMap.set(`${dateKey}_${normalizeClinicName(cLoc)}`, cTime);
        }
      }
    }
  }

  // --- 2. データ準備 ---
  const EXCLUDED_KEYWORDS = ["有給", "欠勤", "院外勤務", "バックアップ", "医師会", "嘱託医", "出張インフルエンザワクチン"];
  const allClinics = new Set(); 
  const backupInfoMap = {}; 
  const rawNameMap = {}; 
  const workingClinics = {}; 
  
  const srcData = sourceSheet.getDataRange().getValues();
  for (let i = 1; i < srcData.length; i++) {
    const row = srcData[i];
    const name = row[0] ? String(row[0]).trim() : "";
    
    if (name === "【関東】バックアップシフト") {
        const dateObj = parseDateToSafeDateObj(row[1]);
        if (dateObj) {
            const dateKey = fastFormatDate(dateObj);
            const slots = ["09:00~13:00", "15:00~18:00", "18:00~21:00"];
            let backups = [];
            for (let j = 0; j < slots.length; j++) {
                const doc = row[7 + j]; 
                if (doc && String(doc).trim()) backups.push(`${slots[j]}：${doc}先生（全拠点）`);
            }
            if (backups.length > 0) backupInfoMap[dateKey] = `【バックアップ】${backups.join('、')}`;
        }
        continue; 
    }

    if (!name || EXCLUDED_KEYWORDS.some(ex => name.includes(ex))) continue;
    const normName = normalizeClinicName(name);
    allClinics.add(normName);
    rawNameMap[normName] = name; 

    const dateObj = parseDateToSafeDateObj(row[1]);
    if (dateObj) {
        const dateKey = fastFormatDate(dateObj);
        if (!workingClinics[dateKey]) workingClinics[dateKey] = new Set();
        workingClinics[dateKey].add(normName);
    }
  }
  
  const ishiFuzaiRawData = ishiFuzaiSheet.getDataRange().getValues();
  const fuzaiRecords = []; 

  for (let i = 1; i < ishiFuzaiRawData.length; i++) { 
    const row = ishiFuzaiRawData[i];
    const dateValue = row[0];        
    const name = row[1] ? row[1].toString().trim() : ""; 
    const time = row[2] ? row[2].toString().trim() : "";

    if (!dateValue || !name || !time) continue; 
    if (EXCLUDED_KEYWORDS.some(ex => name.includes(ex))) continue;

    const normName = normalizeClinicName(name);
    allClinics.add(normName); 
    rawNameMap[normName] = name; 

    const dObj = parseDateToSafeDateObj(dateValue);
    if (dObj) {
      fuzaiRecords.push({
        dateKey: fastFormatDate(dObj),
        name: name,      
        normName: normName, 
        timeRaw: time,
        isInternalMedicine: name.includes("内科")
      });
    }
  }

  const sortedClinicList = Array.from(allClinics).sort();

  // --- 3. メイン処理 ---
  targetSheet.getRange('A6').clearContent();

  const now = new Date();
  let hours = now.getHours();
  let minutes = now.getMinutes();
  if (minutes <= 19) minutes = "00"; else if (minutes <= 49) minutes = "30"; else { minutes = "00"; hours = (hours + 1) % 24; }
  const formattedHours = String(hours).padStart(2, '0');
  const formattedReportDate = `${now.getMonth() + 1}月${now.getDate()}日（${weekdaysJP[now.getDay()]}）`;
  
  let initialText = `【未充足報告】${formattedReportDate} ${formattedHours}:${minutes}時点\n\n`;

  let mentionsArray = [], ccArray = [];
  const mentionData = mentionSheet.getDataRange().getValues();
  for (let i = 1; i < mentionData.length; i++) {
    if (mentionData[i][0]) mentionsArray.push(String(mentionData[i][0]).trim());
    if (mentionData[i][1]) ccArray.push(String(mentionData[i][1]).trim());
  }
  if (mentionsArray.length > 0) initialText += mentionsArray.join('') + "\n";
  if (ccArray.length > 0) initialText += "CC:" + ccArray.join('') + "\n";
  initialText += "\n";

  let hasAnyContent = false;
  let dailyText = "";

  // ★ 週間・月間集計用変数
  let weeklyReq1st = 0, weeklyFilled1st = 0, weeklyGapMin = 0;
  let weeklyReq2nd = 0, weeklyFilled2nd = 0;
  let weeklyCooReqMin = 0, weeklyCooFilledMin = 0;
  let weeklyAbsenceClinics = []; 

  let monthlyReq1st = 0, monthlyFilled1st = 0, monthlyGapMin = 0;
  let monthlyReq2nd = 0, monthlyFilled2nd = 0;
  let monthlyCooReqMin = 0, monthlyCooFilledMin = 0;
  let monthlyAbsenceCount = 0;

  const loopStart = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const loopEnd = new Date(Math.max(endDate.getTime(), new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0).getTime()));

  for (let d = new Date(loopStart); d <= loopEnd; d.setDate(d.getDate() + 1)) {
    const isWeekly = (d.getTime() >= startDate.getTime() && d.getTime() <= endDate.getTime());
    const isMonthly = (d.getFullYear() === startDate.getFullYear() && d.getMonth() === startDate.getMonth());

    if (!isWeekly && !isMonthly) continue;

    const dateKey = fastFormatDate(d);
    const dateTitle = `${d.getMonth() + 1}月${d.getDate()}日（${weekdaysJP[d.getDay()]}）`;
    const isSpecialDay = (d.getMonth() + 1 === 12 && d.getDate() === 31) || (d.getMonth() + 1 === 1 && d.getDate() <= 3);

    let totalRequiredMinutes = 0;
    
    sortedClinicList.forEach(clinic => {
        if (openDateMap.has(clinic)) {
            const openDate = openDateMap.get(clinic);
            openDate.setHours(0, 0, 0, 0);
            const checkDate = new Date(d);
            checkDate.setHours(0, 0, 0, 0);
            if (checkDate < openDate) return; 
        }

        if (clinic.includes("内科")) return; 
        
        let closedTime = closedDataMap.get(`${dateKey}_${clinic}`) || closedDataMap.get(`${dateKey}_全拠点`);
        if (closedTime === "全日") return; 

        let reqMin = 0;
        if (isSpecialDay) {
            reqMin = 420; 
        } else {
            if (clinic.includes("北葛西")) reqMin = 540; 
            else reqMin = 600; 
        }

        if (closedTime === "午前") reqMin -= 4 * 60;
        else if (closedTime === "午後") reqMin -= 3 * 60;
        else if (closedTime === "夜間") reqMin -= 3 * 60;
        else if (closedTime === "午後夜間") reqMin -= 6 * 60;

        const irregularRule = irregularMap[`${dateKey}_${clinic}`] || irregularMap[`${dateKey}_全拠点`];
        if (irregularRule) {
            reqMin = 0;
            irregularRule.forEach(r => { reqMin += (r.close - r.open); });
        }

        totalRequiredMinutes += Math.max(0, reqMin);
    });

    let totalGapMinutes = 0;
    const dailyOutputLines = [];
    
    const dailyRecords = fuzaiRecords.filter(r => r.dateKey === dateKey);
    const groupedFuzai = {};
    dailyRecords.forEach(r => {
        if (!groupedFuzai[r.normName]) groupedFuzai[r.normName] = [];
        groupedFuzai[r.normName].push(r.timeRaw);
    });

    sortedClinicList.forEach(clinic => {
        if (openDateMap.has(clinic)) {
            const openDate = openDateMap.get(clinic);
            openDate.setHours(0, 0, 0, 0);
            const checkDate = new Date(d);
            checkDate.setHours(0, 0, 0, 0);
            if (checkDate < openDate) return;
        }

        if (clinic.includes("内科")) return; 

        let closedTime = closedDataMap.get(`${dateKey}_${clinic}`) || closedDataMap.get(`${dateKey}_全拠点`);
        if (closedTime === "全日") return;

        const hasFuzai = groupedFuzai[clinic] && groupedFuzai[clinic].length > 0;
        const isWorking = workingClinics[dateKey] && workingClinics[dateKey].has(clinic);

        if (!hasFuzai && !isWorking) {
            let pureAbsenceArr = isSpecialDay ? ["10:00-19:00"] : (clinic.includes("北葛西") ? ["09:00-13:00", "15:00-20:00"] : ["09:00-13:00", "15:00-21:00"]);
            groupedFuzai[clinic] = pureAbsenceArr;
        }
    });

    for (const normName in groupedFuzai) {
        if (openDateMap.has(normName)) {
            const openDate = openDateMap.get(normName);
            openDate.setHours(0, 0, 0, 0);
            const checkDate = new Date(d);
            checkDate.setHours(0, 0, 0, 0);
            if (checkDate < openDate) continue; 
        }

        let closedTime = null;
        if (normName.includes("内科")) {
             closedTime = closedDataMap.get(`${dateKey}_${normName}`) || closedDataMap.get(`${dateKey}_${normName}_内科`) || closedDataMap.get(`${dateKey}_全拠点`);
        } else {
             closedTime = closedDataMap.get(`${dateKey}_${normName}`) || closedDataMap.get(`${dateKey}_全拠点`);
        }
        if (closedTime === "全日") continue; 

        let mergedTimeStr = safeNormalizeAndMerge(groupedFuzai[normName], isSpecialDay);
        
        if (mergedTimeStr) {
            let finalTimeStr = mergedTimeStr;

            if (closedTime) finalTimeStr = safeRemoveClosed(finalTimeStr, closedTime);
            if (!finalTimeStr) continue;

            if (!isSpecialDay && normName.includes("北葛西")) {
                 finalTimeStr = finalTimeStr.replace("21:00", "20:00");
            }
            
            const irregularRule = irregularMap[`${dateKey}_${normName}`] || irregularMap[`${dateKey}_全拠点`];
            if (irregularRule) {
                 const ranges = finalTimeStr.split('/');
                 const validRanges = [];
                 for (const range of ranges) {
                     const parts = range.split('-');
                     if (parts.length === 2) {
                         const s = safeParseTime(parts[0]);
                         const e = safeParseTime(parts[1]);
                         for (const r of irregularRule) {
                             const adjustedS = Math.max(s, r.open);
                             const adjustedE = Math.min(e, r.close);
                             if (adjustedS < adjustedE) {
                                 validRanges.push(`${safeFormatTime(adjustedS)}-${safeFormatTime(adjustedE)}`);
                             }
                         }
                     }
                 }
                 if (validRanges.length === 0) continue; 
                 finalTimeStr = validRanges.join('/');
            }

            if (!normName.includes("内科")) {
                totalGapMinutes += safeCalcTotalMin(finalTimeStr);
            }

            let displayTimeStr = finalTimeStr;
            if (displayTimeStr.startsWith("09:00-13:00/15:00-")) {
                displayTimeStr = displayTimeStr.replace("09:00-13:00/15:00-", "09:00-");
            } else if (displayTimeStr.startsWith("10:00-13:00/15:00-")) { 
                displayTimeStr = displayTimeStr.replace("10:00-13:00/15:00-", "10:00-");
            }

            let record = dailyRecords.find(r => r.normName === normName);
            let displayName = record ? record.name : (rawNameMap[normName] || normName);
            
            const currentDepartment = record && record.isInternalMedicine ? "内科" : (displayName.includes("内科") ? "内科" : "小児科");
            
            if (isWeekly) {
                dailyOutputLines.push(`【${displayName}】${displayTimeStr}`);
                weeklyAbsenceClinics.push({ clinic: normName, dept: currentDepartment });
            }
            if (isMonthly) {
                monthlyAbsenceCount++;
            }
        }
    }

    const totalFilledMinutes = Math.max(0, totalRequiredMinutes - totalGapMinutes);
    const rate = totalRequiredMinutes > 0 ? Math.floor((totalFilledMinutes / totalRequiredMinutes) * 100) : 100;
    const requiredHours = Math.round(totalRequiredMinutes / 60);
    const filledHours = Math.round(totalFilledMinutes / 60);

    const dailyExtData = extDataByDate[dateKey] || { paste: [], oubo: [], bosyu: [] };
    const advMetrics = calculateAdvancedMetricsFast(dailyExtData.paste, dailyExtData.oubo, dailyExtData.bosyu);

    let dailyCooReqMin = 0;
    let dailyCooFilledMin = 0;

    if (cooDataByDate[dateKey]) {
      const covMap = {};
      const getCov = (c) => { if (!covMap[c]) covMap[c] = new Array(1440).fill(0); return covMap[c]; };

      dailyExtData.paste.forEach(row => {
        const doc = String(row[0] || "").trim();
        const clinic = String(row[12] || "").replace(/[（(]小児科[）)]/, "").trim();
        if (!doc || !clinic || doc.includes("バックアップ") || doc.includes("有給") || doc.includes("欠勤")) return;
        const startMin = safeParseTime(row[15]);
        const endMin = safeParseTime(row[19]);
        if (!isNaN(startMin) && !isNaN(endMin) && startMin < endMin) {
          const cov = getCov(clinic);
          for (let m = startMin; m < endMin; m++) cov[m]++;
        }
      });

      dailyExtData.oubo.forEach(row => {
        const doc = String(row[0] || "").trim();
        const clinic = String(row[3] || "").replace(/[（(]小児科[）)]/, "").trim();
        if (!doc || !clinic || clinic.includes("バックアップ")) return;
        if (doc.replace(/\s+/g, '') === "橋本浩") return;
        const startMin = safeParseTime(row[6]);
        const endMin = safeParseTime(row[7]);
        if (!isNaN(startMin) && !isNaN(endMin) && startMin < endMin) {
          const cov = getCov(clinic);
          for (let m = startMin; m < endMin; m++) cov[m]++;
        }
      });

      cooDataByDate[dateKey].forEach(row => {
        const clinic = String(row[0] || "").replace(/[（(]小児科[）)]/, "").trim();
        const startMin = safeParseTime(row[2]);
        const endMin = safeParseTime(row[3]);
        if (!isNaN(startMin) && !isNaN(endMin) && startMin < endMin) {
          dailyCooReqMin += (endMin - startMin);
          if (covMap[clinic]) {
            const cov = covMap[clinic];
            for (let m = startMin; m < endMin; m++) {
              if (cov[m] >= 2) dailyCooFilledMin++;
            }
          }
        }
      });
    }

    if (isWeekly) {
        weeklyReq1st += totalRequiredMinutes;
        weeklyFilled1st += totalFilledMinutes;
        weeklyGapMin += totalGapMinutes;
        weeklyReq2nd += advMetrics.secondReqMin;
        weeklyFilled2nd += advMetrics.secondActualMin;
        weeklyCooReqMin += dailyCooReqMin;
        weeklyCooFilledMin += dailyCooFilledMin;
        
        if (totalRequiredMinutes > 0) {
            let entry = `[info][title]${dateTitle}[/title]`;
            if (backupInfoMap[dateKey]) entry += backupInfoMap[dateKey];
            entry += `[hr]\n`;
            entry += `小児科１診目充足率：${rate}%（応募：${filledHours}h/募集：${requiredHours}h）\n`;
            
            if (isExtSsLoaded) {
                entry += `募集全体充足率：${advMetrics.overallRate}%（応募：${advMetrics.overallActualH}h/募集：${advMetrics.overallReqH}h）\n`;
                entry += `２診目充足率（全体）：${advMetrics.secondRate}%（応募：${advMetrics.secondActualH}h/募集：${advMetrics.secondReqH}h）\n`;
                if (dailyCooReqMin > 0) {
                    const dCooRate = Math.floor((dailyCooFilledMin / dailyCooReqMin) * 100);
                    const dCooReqH = Math.round(dailyCooReqMin / 60);
                    const dCooFilledH = Math.round(dailyCooFilledMin / 60);
                    entry += `└COO室依頼２診：${dCooRate}%（応募：${dCooFilledH}h/募集：${dCooReqH}h）\n`;
                }
            }
            entry += `\n`;
            if (dailyOutputLines.length > 0) {
                dailyOutputLines.sort();
                entry += `＜医師不在拠点＞\n` + dailyOutputLines.join('\n') + '\n';
            } else {
                entry += `充足\n`;
            }
            entry += `[/info]\n`;
            dailyText += entry;
            hasAnyContent = true;
        }
    }

    if (isMonthly) {
        monthlyReq1st += totalRequiredMinutes;
        monthlyFilled1st += totalFilledMinutes;
        monthlyGapMin += totalGapMinutes;
        monthlyReq2nd += advMetrics.secondReqMin;
        monthlyFilled2nd += advMetrics.secondActualMin;
        monthlyCooReqMin += dailyCooReqMin;
        monthlyCooFilledMin += dailyCooFilledMin;
    }
  }

  // --- ★ SummaryBuilder.gs にデータを渡してテキストを構築 ---
  const summaryParams = {
    startDate: startDate,
    endDate: endDate,
    isExtSsLoaded: isExtSsLoaded,
    hasCooData: (weeklyCooReqMin > 0 || Object.keys(cooDataByDate).length > 0),
    weekly: {
      req1st: weeklyReq1st, filled1st: weeklyFilled1st, gapMin: weeklyGapMin,
      req2nd: weeklyReq2nd, filled2nd: weeklyFilled2nd,
      cooReq: weeklyCooReqMin, cooFilled: weeklyCooFilledMin,
      absenceClinics: weeklyAbsenceClinics
    },
    monthly: {
      req1st: monthlyReq1st, filled1st: monthlyFilled1st, gapMin: monthlyGapMin,
      req2nd: monthlyReq2nd, filled2nd: monthlyFilled2nd,
      cooReq: monthlyCooReqMin, cooFilled: monthlyCooFilledMin,
      absenceCount: monthlyAbsenceCount
    },
    areaMap: areaMap,
    TARGET_SPLIT_CLINICS: ["北葛西", "西葛西"]
  };

  const summaryText = buildWeeklyMonthlySummaryText(summaryParams);

  if (!hasAnyContent) {
    dailyText = "対象期間内に報告すべきデータ（小児科）はありませんでした。\n";
  }

  targetSheet.getRange('A6').setValue(initialText + summaryText + dailyText);
  if (ui) ui.alert('文章自動作成が完了しました。'); 
}