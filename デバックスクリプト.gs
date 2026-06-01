/**
 * 開院日マスタを読み込み、指定日が未開院であれば除外するロジックのテスト
 */
function debugTsukubaExclusion() {
  const EXTERNAL_SS_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  const TARGET_SHEET_NAME = '拠点名';
  
  // テスト用の日付と拠点（今回はユーザー様の例に合わせて2024/06/01とします）
  const testDate = new Date('2024/06/01');
  testDate.setHours(0, 0, 0, 0); // 時刻をリセット
  const targetClinic = 'つくば';

  try {
    Logger.log('1. 開院日マスタの構築を開始します...');
    const ss = SpreadsheetApp.openById(EXTERNAL_SS_ID);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    // 拠点名（正規記載）と開院日を紐付けるMapを作成
    const openDateMap = new Map();
    
    // 1行目（ヘッダー）を飛ばして2行目からループ
    for (let i = 1; i < data.length; i++) {
      const clinicName = data[i][0]; // インデックス0: 正規記載
      const openDateRaw = data[i][7]; // インデックス7: 開院日

      // 拠点名が存在し、開院日が日付データとして認識できる場合のみ登録
      if (clinicName && openDateRaw instanceof Date) {
        openDateMap.set(clinicName, openDateRaw);
      }
    }

    Logger.log(`✅ マスタ構築完了。登録拠点数: ${openDateMap.size}件`);
    if (openDateMap.has(targetClinic)) {
      const tsukubaOpenDate = openDateMap.get(targetClinic);
      Logger.log(`💡 「${targetClinic}」の開院日: ${Utilities.formatDate(tsukubaOpenDate, "JST", "yyyy/MM/dd")}`);
    } else {
      Logger.log(`⚠️ 警告: マスタに「${targetClinic}」が見つかりません。「正規記載」の列とテキストが一致しているか確認してください。`);
      return;
    }

    Logger.log('--------------------------------------------------');
    Logger.log(`2. 判定テスト: ${Utilities.formatDate(testDate, "JST", "yyyy/MM/dd")} 時点の「${targetClinic}」`);
    Logger.log('--------------------------------------------------');

    // ▼ 実装予定の除外ロジック ▼
    if (openDateMap.has(targetClinic)) {
      const openDate = openDateMap.get(targetClinic);
      openDate.setHours(0, 0, 0, 0); // 比較のために時刻をリセット

      if (testDate < openDate) {
        Logger.log(`🟢 結果: 【除外】`);
        Logger.log(`理由: ${Utilities.formatDate(testDate, "JST", "yyyy/MM/dd")} は、開院日（${Utilities.formatDate(openDate, "JST", "yyyy/MM/dd")}）より前のため、不在リストに出力しません。`);
      } else {
        Logger.log(`🔴 結果: 【表示（通常の不在判定）】`);
        Logger.log(`理由: 開院済みのため、通常の未充足計算ロジックに引き継がれます。`);
      }
    }

  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}`);
  }
}