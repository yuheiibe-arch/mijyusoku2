/**
 * 指定されたシートの範囲データをクリアし、書式もリセットします。
 * アラートなしで即時実行。保護列（AN列以降）は絶対に破壊しません。
 */
function clearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ss.toast('データをクリアしています...', '処理開始', 2);

    // 処理対象のシートリスト（isSource: true は入力用、false は出力用）
    const targetSheets = [
      { name: '貼付用', isSource: true },
      { name: 'テスト', isSource: true },
      { name: '確認用', isSource: false },
      { name: 'テスト確認用', isSource: false }
    ];

    targetSheets.forEach(sheetInfo => {
      const sheet = ss.getSheetByName(sheetInfo.name);
      if (!sheet) return; // シートがなければスキップ

      const lastRow = sheet.getLastRow();

      if (sheetInfo.isSource) {
        // --- 貼付用・テスト（3行目以降） ---
        // ⚠️ 保護列(AN〜AP)を守るため、A列〜AM列(39列)までのみを対象とする
        if (lastRow >= 3) {
          const range = sheet.getRange(3, 1, lastRow - 2, 39);
          range.clearContent();      // 値を消去
          range.setBackground(null); // 背景色を白（リセット）に戻す
          Logger.log(`[${sheetInfo.name}] 3行目以降、AM列までのデータと背景色をクリアしました。保護列は維持。`);
        }
      } else {
        // --- 確認用・テスト確認用（2行目以降） ---
        // A列〜J列(10列)をまるごとクリア（書式含む）
        if (lastRow >= 2) {
          sheet.getRange(2, 1, lastRow - 1, 10).clear(); 
          Logger.log(`[${sheetInfo.name}] 2行目以降、A〜J列を完全にクリアしました。`);
        }
      }
    });

    SpreadsheetApp.flush();
    ss.toast('すべてのデータと書式のリセットが完了しました。', '完了', 3);

  } catch (e) {
    Logger.log(`データ削除中にエラーが発生しました: ${e.message}`);
    ss.toast(`エラーが発生しました: ${e.message}`, 'エラー', 5);
  }
}