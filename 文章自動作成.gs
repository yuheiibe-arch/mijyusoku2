// ==========================================
// 【最終確定版】本番用スクリプト (純粋不在・部分休館・昼休み連結表示 対応)
// ==========================================

/**
 * 日付パース関数
 */
function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) {
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }

  const dateStr = String(dateInput).trim();
  const jpMatch = dateStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (jpMatch) {
    const year = parseInt(jpMatch[1], 10);
    const month = parseInt(jpMatch[2], 10) - 1;
    const day = parseInt(jpMatch[3], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) return new Date(year, month, day);
  }

  const cleanedDateStr = dateStr.replace(/\s*（.*?）/, '').replace(/-/g, '/'); 
  const parts = cleanedDateStr.split('/');
  if (parts.length === 3) {
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) return new Date(year, month, day);
  }
  return null;
}

/**
 * 時間文字列("09:00") -> 分
 */
function parseTimeToMinutes(timeStr) {
  if (!timeStr || typeof timeStr !== 'string') return NaN;
  const parts = timeStr.trim().split(':');
  if (parts.length < 2) return NaN;
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

/**
 * 分 -> 時間文字列
 */
function formatMinutesToHHMM(totalMinutes) {
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

/**
 * 時間結合・補正ロジック
 */
function normalizeAndMergeTimes(timeStrings, isSpecialDay) {
  const OPEN = isSpecialDay ? 10 * 60 : 9 * 60;   
  const CLOSE = isSpecialDay ? 19 * 60 : 21 * 60; 

  let intervals = [];
  timeStrings.forEach(str => {
    const parts = str.split(/[-~]/);
    if (parts.length !== 2) return;
    let start = parseTimeToMinutes(parts[0]);
    let end = parseTimeToMinutes(parts[1]);

    if (isNaN(start) || isNaN(end)) return;
    if (end <= OPEN || start >= CLOSE) return; 

    start = Math.max(start, OPEN);
    end = Math.min(end, CLOSE);

    if (start < end) intervals.push({ start, end });
  });

  if (intervals.length === 0) return "";

  intervals.sort((a, b) => a.start - b.start);
  let merged = [];
  let current = intervals[0];
  for (let i = 1; i < intervals.length; i++) {
    if (current.end >= intervals[i].start) {
      current.end = Math.max(current.end, intervals[i].end);
    } else {
      merged.push(current);
      current = intervals[i];
    }
  }
  merged.push(current);
  return merged.map(m => `${formatMinutesToHHMM(m.start)}-${formatMinutesToHHMM(m.end)}`).join('/');
}

/**
 * 結合文字列から合計分数を算出
 */
function calculateTotalMinutesFromStr(mergedStr) {
    let total = 0;
    const parts = mergedStr.split('/');
    parts.forEach(p => {
        const rng = p.split('-');
        if (rng.length === 2) {
            const s = parseTimeToMinutes(rng[0]);
            const e = parseTimeToMinutes(rng[1]);
            if (!isNaN(s) && !isNaN(e)) total += (e - s);
        }
    });
    return total;
}

/**
 * 拠点名正規化
 */
function normalizeClinicName(name) {
    return name.replace(/[（(]小児科[）)]/, "").trim();
}

/**
 * 休館時間（プルダウン）をカットするヘルパー
 */
function removeClosedTime(mergedTimeStr, closedTime) {
    if (!closedTime || closedTime === "全日") return "";
    let closedRanges = [];
    if (closedTime === "午前") closedRanges.push({s: 9*60, e: 13*60});
    else if (closedTime === "午後") closedRanges.push({s: 15*60, e: 18*60});
    else if (closedTime === "夜間") closedRanges.push({s: 18*60, e: 21*60});
    else if (closedTime === "午後夜間") closedRanges.push({s: 15*60, e: 21*60});

    if (closedRanges.length === 0) return mergedTimeStr;

    const ranges = mergedTimeStr.split('/');
    const validRanges = [];

    for (const range of ranges) {
        const parts = range.split('-');
        if (parts.length !== 2) continue;
        let start = parseTimeToMinutes(parts[0]);
        let end = parseTimeToMinutes(parts[1]);

        for (const c of closedRanges) {
            if (start >= end) break;
            if (c.s <= start && c.e >= end) {
                start = end; 
            } else if (c.s <= start && c.e > start) {
                start = c.e; 
            } else if (c.s < end && c.e >= end) {
                end = c.s; 
            } else if (c.s > start && c.e < end) {
                validRanges.push(`${formatMinutesToHHMM(start)}-${formatMinutesToHHMM(c.s)}`);
                start = c.e; 
            }
        }
        if (start < end) {
            validRanges.push(`${formatMinutesToHHMM(start)}-${formatMinutesToHHMM(end)}`);
        }
    }
    return validRanges.join('/');
}


// ==========================================
// メイン処理関数
// ==========================================

function generateChatworkMessage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi(); 
  const scriptTimeZone = Session.getScriptTimeZone(); 
  const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"]; 

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
        const dateKey = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy/MM/dd");
        const parts = String(timeRangeStr).split('-');
        if (parts.length === 2) {
          const openMin = parseTimeToMinutes(parts[0]);
          const closeMin = parseTimeToMinutes(parts[1]);
          if (!isNaN(openMin) && !isNaN(closeMin)) {
             const normName = normalizeClinicName(clinicName);
             irregularMap[`${dateKey}_${normName}`] = { open: openMin, close: closeMin };
             irregularMap[`${dateKey}_${clinicName}`] = { open: openMin, close: closeMin };
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
        const dateKey = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy/MM/dd");
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
  
  // A. 確認用シート収集
  const srcData = sourceSheet.getDataRange().getValues();
  for (let i = 1; i < srcData.length; i++) {
    const row = srcData[i];
    const name = row[0] ? String(row[0]).trim() : "";
    
    if (name === "【関東】バックアップシフト") {
        const dateObj = parseDateToSafeDateObj(row[1]);
        if (dateObj) {
            const dateKey = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy/MM/dd");
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
        const dateKey = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy/MM/dd");
        if (!workingClinics[dateKey]) workingClinics[dateKey] = new Set();
        workingClinics[dateKey].add(normName);
    }
  }
  
  // B. 不在シート収集
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
        dateKey: Utilities.formatDate(dObj, scriptTimeZone, "yyyy/MM/dd"),
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

  const startDateRaw = targetSheet.getRange('B2').getValue(); 
  const endDateRaw = targetSheet.getRange('B4').getValue();   
  const startDate = parseDateToSafeDateObj(startDateRaw);
  const endDate = parseDateToSafeDateObj(endDateRaw);

  if (!startDate || !endDate || startDate > endDate) {
    ui.alert('日付指定が無効です。'); return;
  }

  const now = new Date();
  let hours = now.getHours();
  let minutes = now.getMinutes();
  if (minutes <= 19) minutes = "00"; else if (minutes <= 49) minutes = "30"; else { minutes = "00"; hours = (hours + 1) % 24; }
  const formattedHours = String(hours).padStart(2, '0');
  const formattedReportDate = Utilities.formatDate(now, scriptTimeZone, 'M月d日') + `（${weekdaysJP[now.getDay()]}）`;
  
  let message = `【未充足報告】${formattedReportDate} ${formattedHours}:${minutes}時点\n\n`;

  let mentionsArray = [], ccArray = [];
  const mentionData = mentionSheet.getDataRange().getValues();
  for (let i = 1; i < mentionData.length; i++) {
    if (mentionData[i][0]) mentionsArray.push(String(mentionData[i][0]).trim());
    if (mentionData[i][1]) ccArray.push(String(mentionData[i][1]).trim());
  }
  if (mentionsArray.length > 0) message += mentionsArray.join('') + "\n";
  if (ccArray.length > 0) message += "CC:" + ccArray.join('') + "\n";

  let hasAnyContent = false;

  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const dateKey = Utilities.formatDate(d, scriptTimeZone, "yyyy/MM/dd");
    const dateTitle = Utilities.formatDate(d, scriptTimeZone, 'M月d日') + `（${weekdaysJP[d.getDay()]}）`;

    const month = d.getMonth() + 1;
    const day = d.getDate();
    const isSpecialDay = (month === 12 && day === 31) || (month === 1 && day <= 3);

    // --- 分母 (募集時間) 計算 ---
    let totalRequiredMinutes = 0;
    
    sortedClinicList.forEach(clinic => {
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

        const irregularRule = irregularMap[`${dateKey}_${clinic}`];
        if (irregularRule) {
            reqMin = (irregularRule.close - irregularRule.open);
        }

        totalRequiredMinutes += Math.max(0, reqMin);
    });

    // --- 分子 (不在時間) 計算 & リスト作成 ---
    let totalGapMinutes = 0;
    const dailyOutputLines = [];
    
    const dailyRecords = fuzaiRecords.filter(r => r.dateKey === dateKey);
    const groupedFuzai = {};
    dailyRecords.forEach(r => {
        if (!groupedFuzai[r.normName]) groupedFuzai[r.normName] = [];
        groupedFuzai[r.normName].push(r.timeRaw);
    });

    // 純粋不在のあぶり出し注入
    sortedClinicList.forEach(clinic => {
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

            if (closedTime) {
                finalTimeStr = removeClosedTime(finalTimeStr, closedTime);
            }
            if (!finalTimeStr) continue;

            if (!isSpecialDay && normName.includes("北葛西")) {
                 finalTimeStr = finalTimeStr.replace("21:00", "20:00");
            }
            
            const irregularRule = irregularMap[`${dateKey}_${normName}`];
            if (irregularRule) {
                 const ranges = finalTimeStr.split('/');
                 const validRanges = [];
                 for (const range of ranges) {
                     const parts = range.split('-');
                     if (parts.length === 2) {
                         const s = parseTimeToMinutes(parts[0]);
                         const e = parseTimeToMinutes(parts[1]);
                         const adjustedS = Math.max(s, irregularRule.open);
                         const adjustedE = Math.min(e, irregularRule.close);
                         if (adjustedS < adjustedE) {
                             validRanges.push(`${formatMinutesToHHMM(adjustedS)}-${formatMinutesToHHMM(adjustedE)}`);
                         }
                     }
                 }
                 if (validRanges.length === 0) continue; 
                 finalTimeStr = validRanges.join('/');
            }

            // ★ 計算は「書き換え前」の正確な時間文字列で行う
            if (!normName.includes("内科")) {
                totalGapMinutes += calculateTotalMinutesFromStr(finalTimeStr);
            }

            // ★ 【追加】昼休み（13:00-15:00）を跨いでまるごと不在の場合、見た目だけを綺麗に連結する
            let displayTimeStr = finalTimeStr;
            // A枠・B枠を通しで不在なら、間の休憩時間を表記から消して繋げる
            if (displayTimeStr.startsWith("09:00-13:00/15:00-")) {
                displayTimeStr = displayTimeStr.replace("09:00-13:00/15:00-", "09:00-");
            } else if (displayTimeStr.startsWith("10:00-13:00/15:00-")) { // 年末年始などの10時出勤考慮
                displayTimeStr = displayTimeStr.replace("10:00-13:00/15:00-", "10:00-");
            }

            // リストに追加 (表示は繋げた文字列)
            let record = dailyRecords.find(r => r.normName === normName);
            let displayName = record ? record.name : (rawNameMap[normName] || normName);
            dailyOutputLines.push(`【${displayName}】${displayTimeStr}`);
        }
    }

    // --- 充足率計算 (小児科のみ) ---
    const totalFilledMinutes = Math.max(0, totalRequiredMinutes - totalGapMinutes);
    const rate = totalRequiredMinutes > 0 ? Math.floor((totalFilledMinutes / totalRequiredMinutes) * 100) : 100;
    const requiredHours = Math.round(totalRequiredMinutes / 60);
    const filledHours = Math.round(totalFilledMinutes / 60);

    // --- 出力テキスト生成 ---
    if (totalRequiredMinutes > 0) {
        let entry = `[info][title]${dateTitle}[/title]`;
        if (backupInfoMap[dateKey]) entry += backupInfoMap[dateKey];

        entry += `[hr]\n`;
        entry += `小児科１診目充足率：${rate}%（応募：${filledHours}h/募集：${requiredHours}h）\n`;

        if (dailyOutputLines.length > 0) {
            dailyOutputLines.sort();
            entry += `＜医師不在拠点＞\n` + dailyOutputLines.join('\n') + '\n';
        } else {
            entry += `充足\n`;
        }
        entry += `[/info]\n`;
        message += entry;
        hasAnyContent = true;
    }
  }

  if (!hasAnyContent) {
    message += "対象期間内に報告すべきデータ（小児科）はありませんでした。\n";
  }

  targetSheet.getRange('A6').setValue(message);
  SpreadsheetApp.getActiveSpreadsheet().toast('文章自動作成が完了しました。');
}