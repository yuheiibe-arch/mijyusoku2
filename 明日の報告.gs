// ==========================================
// 【最終確定・本番稼働用】reportDoctorAvailability
// ・検証済みの「年末年始タイムライン判定ロジック」を搭載
// ・休館日、変則営業シートの読み込み完備
// ==========================================

function reportDoctorAvailability() {
    Logger.log("★★★ スクリプト開始: reportDoctorAvailability (最終検証完了版) ★★★");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName("確認用");
    const targetSheet = ss.getSheetByName("文章自動作成");
    const mentionSheet = ss.getSheetByName("メンション先選択");
    const irregularSheet = ss.getSheetByName("変則営業");
    const closedSheet = ss.getSheetByName("休館日");

    if (!sourceSheet || !targetSheet || !mentionSheet || !irregularSheet || !closedSheet) {
        SpreadsheetApp.getUi().alert("エラー: 必要なシートが見つかりません。");
        return;
    }

    targetSheet.getRange("A6").clearContent();

    // --- 1. 日付設定 ---
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);
    tomorrow.setHours(0, 0, 0, 0);
    
    const tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const weekdaysJP = ["日", "月", "火", "水", "木", "金", "土"];
    const formattedTomorrowTitle = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), `M月d日（${weekdaysJP[tomorrow.getDay()]}）`);
    const formattedTodayReport = Utilities.formatDate(today, Session.getScriptTimeZone(), `M月d日（${weekdaysJP[today.getDay()]}）`);

    // --- 2. 休館日データの読み込み ---
    const closedDataMap = new Set();
    const cData = closedSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
        const row = cData[i];
        const d = parseDateToSafeDateObj(row[0]);
        // 日付が明日と一致し、拠点名がある場合
        if (d && Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd") === tomorrowStr && row[2]) {
            const name = String(row[2]).trim();
            closedDataMap.add(normalizeClinicName(name));
            closedDataMap.add(name);
        }
    }

    // --- 3. 変則営業データの読み込み ---
    const irregularMap = {}; 
    const iData = irregularSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
        const row = iData[i];
        const d = parseDateToSafeDateObj(row[0]);
        // 日付が明日と一致し、時間と拠点名がある場合
        if (d && Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd") === tomorrowStr && row[1] && row[2]) {
            const timeParts = String(row[1]).split('-');
            if (timeParts.length === 2) {
                const openMin = parseTimeToMinutes(timeParts[0]);
                const closeMin = parseTimeToMinutes(timeParts[1]);
                const name = String(row[2]).trim();
                const normName = normalizeClinicName(name);
                irregularMap[normName] = { open: openMin, close: closeMin };
                irregularMap[name] = { open: openMin, close: closeMin };
            }
        }
    }

    // --- 4. 年末年始判定 (12/31〜1/3) ---
    const tMonth = tomorrow.getMonth() + 1;
    const tDay = tomorrow.getDate();
    const isSpecialDay = (tMonth === 12 && tDay === 31) || (tMonth === 1 && tDay <= 3);

    // --- 5. 判定ループ ---
    const sourceData = sourceSheet.getDataRange().getValues();
    let unfilledList = [];
    const COL = { NAME: 0, DATE: 1, DEPT: 2, A: 3, B: 4, C: 5 };

    const isEmpty = (val) => val === 0 || val === "" || val === null || val === undefined;

    for (let i = 1; i < sourceData.length; i++) {
        const row = sourceData[i];
        const rawName = row[COL.NAME] ? String(row[COL.NAME]).trim() : "";
        
        // 基本除外チェック
        if (!rawName || rawName.includes("バックアップ") || rawName.includes("有給") || rawName.includes("欠勤")) continue;
        
        // 日付チェック
        const rowDate = parseDateToSafeDateObj(row[COL.DATE]);
        if (!rowDate || rowDate.getTime() !== tomorrow.getTime()) continue;

        const normName = normalizeClinicName(rawName);

        // 休館日チェック -> 休館ならスキップ
        if (closedDataMap.has(normName) || closedDataMap.has(rawName)) continue;

        // --- 不在の「時間の線」リストを作成 ---
        let absentTimeRanges = [];
        
        // A枠 (09:00-13:00)
        if (isEmpty(row[COL.A])) absentTimeRanges.push("09:00-13:00");
        
        // B枠判定 (通常15-18 / 特別日15-19)
        if (isEmpty(row[COL.B])) {
            if (isSpecialDay) {
                // 年末年始はB枠が19:00まで責任を持つ（延長）
                absentTimeRanges.push("15:00-19:00");
            } else {
                absentTimeRanges.push("15:00-18:00");
            }
        }
        
        // C枠判定 (18:00-21:00)
        // 年末年始(isSpecialDay)は、B枠延長でカバーするためC枠は見ない
        if (!isSpecialDay) {
            if (isEmpty(row[COL.C])) {
                if (rawName.includes("北葛西")) {
                    absentTimeRanges.push("18:00-20:00");
                } else {
                    absentTimeRanges.push("18:00-21:00");
                }
            }
        }

        if (absentTimeRanges.length === 0) continue;

        // --- ロジック適用1: 年末年始ルールで結合・カット ---
        // normalizeAndMergeTimesを使用 (年末年始は19:00でカット)
        const mergedTimeStr = normalizeAndMergeTimes(absentTimeRanges, isSpecialDay);

        if (mergedTimeStr) {
            let finalTimeStr = mergedTimeStr;

            // --- ロジック適用2: 変則営業シートでさらにカット ---
            const irregularRule = irregularMap[normName] || irregularMap[rawName];

            if (irregularRule) {
                const ranges = finalTimeStr.split('/');
                const validRanges = [];

                for (const range of ranges) {
                    const parts = range.split('-');
                    if (parts.length === 2) {
                        const s = parseTimeToMinutes(parts[0]);
                        const e = parseTimeToMinutes(parts[1]);

                        // 営業時間内にクリップ (交差判定)
                        const adjustedS = Math.max(s, irregularRule.open);
                        const adjustedE = Math.min(e, irregularRule.close);

                        // 有効な時間があれば追加
                        if (adjustedS < adjustedE) {
                            validRanges.push(`${formatMinutesToHHMM(adjustedS)}-${formatMinutesToHHMM(adjustedE)}`);
                        }
                    }
                }

                if (validRanges.length === 0) continue; // 全てカットされたら出力しない
                finalTimeStr = validRanges.join('/');
            }

            // リストに追加
            unfilledList.push(`【${rawName}】${finalTimeStr}`);
        }
    }

    // --- 6. 書き出しと送信 ---
    const now = new Date();
    let hours = now.getHours();
    let minutes = now.getMinutes();
    if (minutes <= 19) minutes = "00";
    else if (minutes <= 49) minutes = "30";
    else { minutes = "00"; hours = (hours + 1) % 24; }
    const formattedTime = `${String(hours).padStart(2,'0')}:${minutes}時点`;

    let message = `【未充足報告】${formattedTodayReport} ${formattedTime}\n\n`;

    // メンション
    const mentionData = mentionSheet.getDataRange().getValues();
    let mentions = [], ccs = [];
    for (let i = 1; i < mentionData.length; i++) {
        if (mentionData[i][0]) mentions.push(String(mentionData[i][0]).trim());
        if (mentionData[i][1]) ccs.push(String(mentionData[i][1]).trim());
    }
    if (mentions.length) message += mentions.join('') + "\n";
    if (ccs.length) message += "CC:" + ccs.join('') + "\n";
    message += "\n";

    message += "お疲れ様です。\n翌日の医師不在可能性大の拠点及び時間をご報告いたします。\nご確認よろしくお願いいたします。\n";
    message += `【翌日医師不在可能性大報告】${formattedTodayReport} ${formattedTime}\n`;
    message += `[info][title]${formattedTomorrowTitle}[/title]\n`;

    if (unfilledList.length > 0) {
        message += unfilledList.join("\n") + "\n";
    } else {
        message += "充足\n";
    }
    message += "[/info]\n";

    targetSheet.getRange("A6").setValue(message);
    
    // 送信処理
    if (typeof postToChatwork === 'function') {
        postToChatwork(message);
    } else {
        Logger.log("警告: postToChatwork 関数が見つかりません。");
    }

    Logger.log("★★★ 処理完了 ★★★");
}

// ==========================================
// 必須ヘルパー関数 (ロジックの要)
// ==========================================

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
    
    // ここで範囲外をカット（線を切る）
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

function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  const dateStr = String(dateInput).trim();
  const jpMatch = dateStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (jpMatch) return new Date(parseInt(jpMatch[1], 10), parseInt(jpMatch[2], 10) - 1, parseInt(jpMatch[3], 10));
  const cleanedDateStr = dateStr.replace(/\s*（.*?）/, '').replace(/-/g, '/'); 
  const parts = cleanedDateStr.split('/');
  if (parts.length === 3) return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  return null;
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr) return NaN;
  const parts = timeStr.trim().split(':');
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

function formatMinutesToHHMM(totalMinutes) {
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function normalizeClinicName(name) {
    return name.replace(/[（(]小児科[）)]/, "").trim();
}