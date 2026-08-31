function debugCOORecruitment() {
  const startDate = new Date(2026, 8, 1);  // 2026/09/01
  const endDate = new Date(2026, 8, 30);   // 2026/09/30

  // 1. 各スプレッドシートの読み込み（★見た目通りの文字列として取得する getDisplayValues を使用）
  const mainSS = SpreadsheetApp.getActiveSpreadsheet();
  let pasteData = [];
  try {
    pasteData = mainSS.getSheetByName("貼付用").getDataRange().getDisplayValues().slice(2);
  } catch(e) { Logger.log("貼付用の読み込みに失敗"); return; }

  let ouboData = [];
  try {
    const extSS = SpreadsheetApp.openById('1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co');
    ouboData = extSS.getSheetByName("応募シフト").getDataRange().getDisplayValues().slice(1);
  } catch(e) { Logger.log("応募シフトの読み込みに失敗"); return; }

  let cooData = [];
  try {
    const cooSS = SpreadsheetApp.openById('1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk');
    cooData = cooSS.getSheetByName("２診要望一覧").getDataRange().getDisplayValues().slice(1);
  } catch(e) { Logger.log("COO要望一覧の読み込みに失敗"); return; }

  // 2. 日ごとの各拠点の「分単位の医師稼働数」を記録する配列（1440分 = 24時間）
  const dailyCoverage = {};
  function getCoverageArray(dateStr, clinic) {
    if (!dailyCoverage[dateStr]) dailyCoverage[dateStr] = {};
    if (!dailyCoverage[dateStr][clinic]) dailyCoverage[dateStr][clinic] = new Array(1440).fill(0);
    return dailyCoverage[dateStr][clinic];
  }

  // --- 貼付用（確定シフト）の処理 ---
  pasteData.forEach(row => {
    const doc = String(row[0] || "").trim();
    const clinic = String(row[12] || "").replace(/[（(]小児科[）)]/, "").trim();
    const dateObj = parseDateToSafeDateObjDebug(row[14]);
    
    if (!doc || !clinic || !dateObj || dateObj < startDate || dateObj > endDate) return;
    if (doc.includes("バックアップ") || doc.includes("有給") || doc.includes("欠勤")) return;
    
    const startMin = parseTimeToMinutesDebug(row[15]);
    const endMin = parseTimeToMinutesDebug(row[19]);
    if (isNaN(startMin) || isNaN(endMin) || startMin >= endMin) return;
    
    const dateStr = fastFormatDateDebug(dateObj);
    const cov = getCoverageArray(dateStr, clinic);
    
    // 勤務している分すべてに+1（重複すれば2, 3と増える）
    for (let m = startMin; m < endMin; m++) cov[m]++;
  });

  // --- 応募シフトの処理 ---
  ouboData.forEach(row => {
    const doc = String(row[0] || "").trim();
    const clinic = String(row[3] || "").replace(/[（(]小児科[）)]/, "").trim();
    const dateObj = parseDateToSafeDateObjDebug(row[5]);
    
    if (!doc || !clinic || !dateObj || dateObj < startDate || dateObj > endDate) return;
    if (clinic.includes("バックアップ")) return;
    
    // ★ 橋本浩医師は応募段階では未確定（空き）とする
    if (doc.replace(/\s+/g, '') === "橋本浩") return;

    const startMin = parseTimeToMinutesDebug(row[6]);
    const endMin = parseTimeToMinutesDebug(row[7]);
    if (isNaN(startMin) || isNaN(endMin) || startMin >= endMin) return;
    
    const dateStr = fastFormatDateDebug(dateObj);
    const cov = getCoverageArray(dateStr, clinic);
    for (let m = startMin; m < endMin; m++) cov[m]++;
  });

  // 3. COO要望一覧と突き合わせる
  let totalReqMin = 0;
  let totalFilledMin = 0;
  let sampleLogs = 0;

  cooData.forEach(row => {
    const clinic = String(row[0] || "").replace(/[（(]小児科[）)]/, "").trim();
    const dateObj = parseDateToSafeDateObjDebug(row[1]);
    
    if (!clinic || !dateObj || dateObj < startDate || dateObj > endDate) return;

    const startMin = parseTimeToMinutesDebug(row[2]);
    const endMin = parseTimeToMinutesDebug(row[3]);
    if (isNaN(startMin) || isNaN(endMin) || startMin >= endMin) return;

    const reqDuration = endMin - startMin;
    totalReqMin += reqDuration;

    let filledForThisReq = 0;
    const dateStr = fastFormatDateDebug(dateObj);
    
    if (dailyCoverage[dateStr] && dailyCoverage[dateStr][clinic]) {
      const cov = dailyCoverage[dateStr][clinic];
      // ★ その時間帯の「医師稼働数が2以上（2診目がいる）」分だけを充足分としてカウント
      for (let m = startMin; m < endMin; m++) {
        if (cov[m] >= 2) filledForThisReq++;
      }
    }
    totalFilledMin += filledForThisReq;

    // サンプルログを数件出力
    if (sampleLogs < 5) {
      Logger.log(`[検証] ${dateStr} 【${clinic}】 ${formatMinutesToHHMMDebug(startMin)}-${formatMinutesToHHMMDebug(endMin)}`);
      Logger.log(`  -> 要望: ${reqDuration}分 / 2診目稼働: ${filledForThisReq}分`);
      sampleLogs++;
    }
  });

  const rate = totalReqMin > 0 ? Math.floor((totalFilledMin / totalReqMin) * 100) : 0;
  const filledHours = Math.round(totalFilledMin / 60);
  const reqHours = Math.round(totalReqMin / 60);

  Logger.log("\n=================================");
  Logger.log(`9月 COO室依頼２診 集計結果`);
  Logger.log(`COO室依頼２診：${rate}%（応募：${filledHours}h/募集：${reqHours}h）`);
  Logger.log("=================================");

  // --- 内部ヘルパー関数 ---
  function fastFormatDateDebug(dateObj) {
    const y = dateObj.getFullYear();
    const m = String(dateObj.getMonth() + 1).padStart(2, '0');
    const d = String(dateObj.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  }
  function parseDateToSafeDateObjDebug(dateInput) {
    if (!dateInput) return null;
    let dateStr = String(dateInput).trim().split(/[（(]/)[0].trim().replace(/-/g, '/');
    const parts = dateStr.split('/');
    if (parts.length >= 3) return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    return null;
  }
  function parseTimeToMinutesDebug(timeInput) {
    if (typeof timeInput === 'string') {
      // 9:00 のような文字列を正しく分数に変換
      const p = timeInput.trim().split(':');
      if (p.length >= 2) return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
    }
    return NaN;
  }
  function formatMinutesToHHMMDebug(totalMin) {
    return `${String(Math.floor(totalMin/60)).padStart(2,'0')}:${String(totalMin%60).padStart(2,'0')}`;
  }
}