// 日付文字列 "YYYY/MM/DD" or "YYYY-MM-DD" や Dateオブジェクトから、
// タイムゾーン問題を避けて年月日のみのDateオブジェクトを生成するヘルパー関数
function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;

  if (dateInput instanceof Date) {
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }

  if (typeof dateInput !== 'string' && typeof dateInput.toString !== 'function') {
    Logger.log(`日付パース不可: 文字列でもDateオブジェクトでもない入力値 "[${dateInput}]"`);
    return null;
  }

  const dateStr = dateInput.toString();
  const cleanedDateStr = dateStr.replace(/\s*（.*?）/, '').replace(/-/g, '/');
  const parts = cleanedDateStr.split('/');

  if (parts.length === 3) {
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      return new Date(year, month, day);
    }
  }

  Logger.log(`日付パース失敗: 入力値 "[${dateInput}]", クリーンアップ後 "[${cleanedDateStr}]"`);
  return null;
}

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
  const formattedReportDate = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'M月d日') + `（${weekdaysJP[reportDateObj.getDay()]}）`;
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

  Logger.log(`取得したTOメンション配列: ${mentionsArray.join('; ')}`);
  Logger.log(`取得したCCメンション配列: ${ccArray.join('; ')}`);

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
    Logger.log(`整形後のTOメンションテキスト:\n${toMentionsText}`);
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
    Logger.log(`整形後のCCメンションテキスト:\n${ccMentionsText}`);
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

    const dateKey = Utilities.formatDate(shiftDateObj, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const formattedShiftDateForTitle = Utilities.formatDate(shiftDateObj, Session.getScriptTimeZone(), 'M月d日') + `（${weekdaysJP[shiftDateObj.getDay()]}）`;

    if (!dailyReportData[dateKey]) {
      dailyReportData[dateKey] = {
        titleDate: formattedShiftDateForTitle,
        backupText: "",
        unfilledList: [],
        hasNonBackupClinicsProcessed: false
      };
      Logger.log(`新規日付キー作成: ${dateKey} (${formattedShiftDateForTitle})`);
    }

    Logger.log(`処理中: ${dateKey} - ${clinicName} (行 ${i + 1})`);

    if (clinicName === KANTO_BACKUP_SHIFT_NAME_CONST) {
      Logger.log(`  ${clinicName}: バックアップ情報処理`);
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
        Logger.log(`    バックアップ情報更新: ${dailyReportData[dateKey].backupText}`);
      }
    } else {
      dailyReportData[dateKey].hasNonBackupClinicsProcessed = true;
      Logger.log(`  ${clinicName}: 通常クリニック情報処理`);

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
            Logger.log(`      未充足として追加: ${unfilledEntryText}`);
          } else {
            Logger.log(`      未充足情報重複のためスキップ: ${unfilledEntryText}`);
          }
          entryUnfilledThisTime = true;
        }
      }

      if (!entryUnfilledThisTime) {
        Logger.log(`    ${clinicName} は全スロット充足または医師配置あり`);
      }
    }
  }

  Logger.log('★★★ メッセージ組み立てフェーズ ★★★');
  const sortedDateKeys = Object.keys(dailyReportData).sort();
  Logger.log(`集約された日付キー (ソート済): ${sortedDateKeys.join(', ')}`);

  if (sortedDateKeys.length === 0) {
    message += "対象期間内に報告すべき未充足情報はありませんでした。\n";
    Logger.log("対象期間内に処理対象データなし。");
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
      Logger.log(`${dayData.titleDate}: 未充足あり。リスト: ${dayData.unfilledList.join('; ')}`);
    } else if (dayData.hasNonBackupClinicsProcessed) {
      entry += "充足\n";
      Logger.log(`${dayData.titleDate}: 未充足なし (バックアップ以外のクリニック処理あり) -> 「充足」と表示`);
    }
    entry += "[/info]\n";
    message += entry;
    Logger.log(`生成されたentry (${dayData.titleDate}): ${entry.replace(/\n/g, "\\n")}`);
  }

  Logger.log('★★★ 本文の生成を終了 ★★★');
  targetSheet.getRange('A6').setValue(message);
  Logger.log('文章を A6 に書き出しました。スクリプト完了');
}