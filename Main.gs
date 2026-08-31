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

  // --- シート取得 ---
  const sourceSheet = ss.getSheetByName('確認用');
  const targetSheet = ss.getSheetByName('文章自動作成');
  const mentionSheet = ss.getSheetByName('メンション先選択');
  const ishiFuzaiSheet = ss.getSheetByName('医師不在拠点');
  const closedSheet = ss.getSheetByName('休館日'); 
  const irregularSheet = ss.getSheetByName('変則営業');

  if (!sourceSheet || !targetSheet || !mentionSheet || !ishiFuzaiSheet) { 
    ui.alert('エラー: 必要なシートが見つかりません。'); 
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
    ui.alert(`日付指定が無効です。\n開始: ${startDateRaw}\n終了: ${endDateRaw}`); 
    return; 
  }

  // --- ★追加: 開院日マスタ・エリアマスタの読み込み ---
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

  // --- 0. 変則営業データの読み込み (★修正: カンマ区切り・全拠点対応) ---
  const irregularMap = {}; 
  if (irregularSheet) {
    const iData = irregularSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      const row = iData[i];
      const dateObj = parseDateToSafeDateObj(row[0]); 
      const timeRangeStr = row[1]; // 例: "09:00-12:30,15:00-21:00"
      const clinicName = row[2] ? String(row[2]).trim() : "";

      if (dateObj && timeRangeStr && clinicName) {
        const dateKey = fastFormatDate(dateObj);
        
        // カンマやスラッシュで複数区間に分割して配列化
        const rangesStr = String(timeRangeStr).split(/[,/、]/);
        const validRanges = [];
        for (let rStr of rangesStr) {
          const parts = rStr.split('-');
          if (parts.length === 2) {
            const openMin = parseTimeToMinutes(parts[0]);
            const closeMin = parseTimeToMinutes(parts[1]);
            if (!isNaN(openMin) && !isNaN(closeMin) && openMin < closeMin) {
              validRanges.push({ open: openMin, close: closeMin });
            }
          }
        }
        
        if (validRanges.length > 0) {
           const normName = normalizeClinicName(clinicName);
           // 全拠点を指定された場合、特別なキーで保存する
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
  
  // 冒頭メンションと挨拶ブロック
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

  // ★ 週間集計用変数
  let weeklyReq1st = 0, weeklyFilled1st = 0, weeklyGapMin = 0;
  let weeklyReq2nd = 0, weeklyFilled2nd = 0;
  let weeklyAbsenceClinics = []; 

  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const dateKey = fastFormatDate(d);
    const dateTitle = `${d.getMonth() + 1}月${d.getDate()}日（${weekdaysJP[d.getDay()]}）`;

    const month = d.getMonth() + 1;
    const day = d.getDate();
    const isSpecialDay = (month === 12 && day === 31) || (month === 1 && day <= 3);

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

        // ★修正: 変則営業の必要時間（分母）を配列から合算、全拠点も参照
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

        let mergedTimeStr = normalizeAndMergeTimes(groupedFuzai[normName], isSpecialDay);
        
        if (mergedTimeStr) {
            let finalTimeStr = mergedTimeStr;

            if (closedTime) finalTimeStr = removeClosedTime(finalTimeStr, closedTime);
            if (!finalTimeStr) continue;

            if (!isSpecialDay && normName.includes("北葛西")) {
                 finalTimeStr = finalTimeStr.replace("21:00", "20:00");
            }
            
            // ★修正: 変則営業の許可区間（分子）と重なる部分だけを不在として抽出
            const irregularRule = irregularMap[`${dateKey}_${normName}`] || irregularMap[`${dateKey}_全拠点`];
            if (irregularRule) {
                 const ranges = finalTimeStr.split('/');
                 const validRanges = [];
                 for (const range of ranges) {
                     const parts = range.split('-');
                     if (parts.length === 2) {
                         const s = parseTimeToMinutes(parts[0]);
                         const e = parseTimeToMinutes(parts[1]);
                         
                         // すべての許可区間との重なりを判定
                         for (const r of irregularRule) {
                             const adjustedS = Math.max(s, r.open);
                             const adjustedE = Math.min(e, r.close);
                             if (adjustedS < adjustedE) {
                                 validRanges.push(`${formatMinutesToHHMM(adjustedS)}-${formatMinutesToHHMM(adjustedE)}`);
                             }
                         }
                     }
                 }
                 if (validRanges.length === 0) continue; 
                 finalTimeStr = validRanges.join('/');
            }

            if (!normName.includes("内科")) {
                totalGapMinutes += calculateTotalMinutesFromStr(finalTimeStr);
            }

            let displayTimeStr = finalTimeStr;
            if (displayTimeStr.startsWith("09:00-13:00/15:00-")) {
                displayTimeStr = displayTimeStr.replace("09:00-13:00/15:00-", "09:00-");
            } else if (displayTimeStr.startsWith("10:00-13:00/15:00-")) { 
                displayTimeStr = displayTimeStr.replace("10:00-13:00/15:00-", "10:00-");
            }

            let record = dailyRecords.find(r => r.normName === normName);
            let displayName = record ? record.name : (rawNameMap[normName] || normName);
            
            dailyOutputLines.push(`【${displayName}】${displayTimeStr}`);
            
            // ★ 集計用配列にオブジェクトとして追加（科名も保持）
            const currentDepartment = record && record.isInternalMedicine ? "内科" : (displayName.includes("内科") ? "内科" : "小児科");
            weeklyAbsenceClinics.push({ clinic: normName, dept: currentDepartment });
        }
    }

    const totalFilledMinutes = Math.max(0, totalRequiredMinutes - totalGapMinutes);
    const rate = totalRequiredMinutes > 0 ? Math.floor((totalFilledMinutes / totalRequiredMinutes) * 100) : 100;
    const requiredHours = Math.round(totalRequiredMinutes / 60);
    const filledHours = Math.round(totalFilledMinutes / 60);

    // --- 日次データから全体充足率・２診目充足率の計算 ---
    const dailyExtData = extDataByDate[dateKey] || { paste: [], oubo: [], bosyu: [] };
    const advMetrics = calculateAdvancedMetricsFast(dailyExtData.paste, dailyExtData.oubo, dailyExtData.bosyu);

    // ★ 週間集計用データの蓄積
    weeklyReq1st += totalRequiredMinutes;
    weeklyFilled1st += totalFilledMinutes;
    weeklyGapMin += totalGapMinutes;
    weeklyReq2nd += advMetrics.secondReqMin;
    weeklyFilled2nd += advMetrics.secondActualMin;

    // --- 出力テキスト生成 ---
    if (totalRequiredMinutes > 0) {
        let entry = `[info][title]${dateTitle}[/title]`;
        if (backupInfoMap[dateKey]) entry += backupInfoMap[dateKey];

        entry += `[hr]\n`;
        entry += `小児科１診目充足率：${rate}%（応募：${filledHours}h/募集：${requiredHours}h）\n`;
        
        if (isExtSsLoaded) {
            entry += `募集全体充足率：${advMetrics.overallRate}%（応募：${advMetrics.overallActualH}h/募集：${advMetrics.overallReqH}h）\n`;
            entry += `２診目充足率：${advMetrics.secondRate}%（応募：${advMetrics.secondActualH}h/募集：${advMetrics.secondReqH}h）\n`;
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

  // --- 週間サマリーの生成と結合 ---
  const wRate1st = weeklyReq1st > 0 ? Math.floor((weeklyFilled1st / weeklyReq1st) * 100) : 100;
  const wReq1stH = Math.round(weeklyReq1st / 60);
  const wFilled1stH = Math.round(weeklyFilled1st / 60);

  const wRate2nd = weeklyReq2nd > 0 ? Math.floor((weeklyFilled2nd / weeklyReq2nd) * 100) : 100;
  const wReq2ndH = Math.round(weeklyReq2nd / 60);
  const wFilled2ndH = Math.round(weeklyFilled2nd / 60);

  const wGapHours = Math.round(weeklyGapMin / 60);
  
  // ★ エリアごとの集計ロジック
  const areaCount = { "東京": [], "神奈川": [], "埼玉": [], "千葉": [], "茨城": [], "大阪": [] };
  const TARGET_SPLIT_CLINICS = ["北葛西", "西葛西"];

  weeklyAbsenceClinics.forEach(record => {
    let displayName = record.clinic;
    if (TARGET_SPLIT_CLINICS.includes(record.clinic) && record.dept) {
      displayName = `${record.clinic}（${record.dept}）`;
    }
    const area = areaMap[record.clinic] || "その他";
    if (!areaCount[area]) areaCount[area] = [];
    areaCount[area].push(displayName);
  });

  let summaryText = `[info][title]週間医師充足数[/title]\n`;
  summaryText += `計測期間：${fastFormatDate(startDate)}~${fastFormatDate(endDate)}\n`;
  summaryText += `１診目充足率：${wRate1st}%（応募：${wFilled1stH}h/募集：${wReq1stH}h）\n`;
  
  if (isExtSsLoaded) {
    summaryText += `２診目充足率：${wRate2nd}%（応募：${wFilled2ndH}h/募集：${wReq2ndH}h）\n`;
  } else {
    summaryText += `２診目充足率：取得エラー\n`;
  }
  
  summaryText += `医師不在時間合計：${wGapHours}h\n`;
  summaryText += `医師不在拠点箇所（延べ数）：${weeklyAbsenceClinics.length}\n`;

  const kantouAreas = ["東京", "神奈川", "埼玉", "千葉", "茨城"];
  kantouAreas.forEach(area => {
    const clinics = areaCount[area] || [];
    if (clinics.length > 0) {
      const uniqueNames = [...new Set(clinics)];
      summaryText += `${area}：${uniqueNames.length}拠点（${uniqueNames.join('、')}）\n`;
    }
  });

  const osakaClinics = areaCount["大阪"] || [];
  if (osakaClinics.length > 0) {
    if (kantouAreas.some(a => (areaCount[a] || []).length > 0)) {
      summaryText += `[hr]\n`;
    }
    const uniqueNames = [...new Set(osakaClinics)];
    summaryText += `大阪：${uniqueNames.length}拠点（${uniqueNames.join('、')}）\n`;
  }

  summaryText += `[/info]\n\n`;

  if (!hasAnyContent) {
    dailyText = "対象期間内に報告すべきデータ（小児科）はありませんでした。\n";
  }

  targetSheet.getRange('A6').setValue(initialText + summaryText + dailyText);
  SpreadsheetApp.getActiveSpreadsheet().toast('文章自動作成が完了しました。');
}