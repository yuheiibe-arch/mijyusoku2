function debugOpeningDateMaster() {
  const EXTERNAL_SS_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  const TARGET_SHEET_NAME = '拠点名';
  // 調査したい拠点名
  const targetKeywords = ['あべの', '市川妙典']; 

  Logger.log('=== 開院日マスタ デバッグ開始（表記ゆれA〜E列対応版） ===');

  try {
    const extSS = SpreadsheetApp.openById(EXTERNAL_SS_ID);
    const extSheet = extSS.getSheetByName(TARGET_SHEET_NAME);

    if (!extSheet) {
      Logger.log(`❌ エラー: '${TARGET_SHEET_NAME}' という名前のシートが見つかりません。`);
      return;
    }

    const extData = extSheet.getDataRange().getValues();
    Logger.log(`✅ マスタシートへのアクセス成功。総行数: ${extData.length}`);

    let foundCount = 0;

    // 1行目（ヘッダー）を飛ばして2行目からループ
    for (let i = 1; i < extData.length; i++) {
      // A〜E列（インデックス0〜4）のデータを配列で取得
      const names = [extData[i][0], extData[i][1], extData[i][2], extData[i][3], extData[i][4]];
      const rawDate = extData[i][7]; // H列: 開院日

      // A〜E列のいずれかに検索キーワードが（部分一致でも）含まれているか
      const hasKeyword = names.some(name => name && targetKeywords.some(keyword => String(name).includes(keyword)));

      if (hasKeyword) {
        foundCount++;
        Logger.log(`\n--- スプレッドシートの ${i + 1} 行目 ---`);
        
        // 登録されている名前をすべて出力
        const registeredNames = names.filter(n => n && String(n).trim() !== "").map(n => `"${n}"`);
        Logger.log(`登録されている名称 (A〜E列): ${registeredNames.join(', ')}`);
        
        // 1. 名前チェック（システムが完全一致で拾えるものが1つでもあるか）
        const isExactMatch = names.some(name => name && targetKeywords.includes(String(name).trim()));
        if (isExactMatch) {
          Logger.log(`  └ 🟢 「${targetKeywords.join(' / ')}」と完全一致する名称がA〜E列に存在します。システムは正しく除外判定できます。`);
        } else {
          Logger.log(`  └ ❌ 部分一致のみです。システムが認識する【 】内の名称と「完全に一致」する表記がA〜E列のどこかに必要です。`);
        }

        // 2. 日付チェック
        Logger.log(`開院日 (H列) : ${rawDate}`);
        if (rawDate instanceof Date) {
          Logger.log(`  └ 🟢 日付データとして正しく認識されています。 (${Utilities.formatDate(rawDate, "JST", "yyyy/MM/dd")})`);
        } else {
          Logger.log(`  └ ❌ 日付オブジェクトではありません。（文字列として入力されているか、空欄です）`);
        }
      }
    }

    if (foundCount === 0) {
      Logger.log(`\n⚠️ 警告: A〜E列のどこにも「${targetKeywords.join(' / ')}」を含むデータが見つかりませんでした。`);
    }

  } catch (e) {
    Logger.log(`\n❌ 重大なエラー: マスタシートへのアクセスに失敗しました。`);
    Logger.log(`エラー内容: ${e.message}`);
  }

  Logger.log('\n=== デバッグ終了 ===');
}