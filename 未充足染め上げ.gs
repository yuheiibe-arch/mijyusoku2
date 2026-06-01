/**
 * 確認用シートの特定のセルの背景色を設定します。
 * - D, E, F列のセルが「0または空白」の場合、そのセルを薄い赤にする。
 * - H, I, J列のセルが「空白」の場合、そのセルを薄い赤にする。
 * - 上記以外のセルは白に戻す。
 */
function applyConditionalFormatting_CellSpecific() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 現在アクティブなスプレッドシートを取得
  const targetSheet = ss.getSheetByName('確認用'); // 名前でシートを取得

  // シートが見つからない場合はエラーメッセージを表示して終了
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert('確認用シートが見つかりません。');
    return;
  }

  const startRow = 2; // データ処理を開始する行番号 (2行目から)
  const lastRow = targetSheet.getLastRow(); // シートの最終行番号を取得

  // データが2行目以降に存在しない場合は処理を終了
  if (lastRow < startRow) {
    Logger.log('確認用シートにデータ行（2行目以降）が存在しないため、書式設定をスキップします。');
    return;
  }

  const numColumns = 10; // 書式設定の対象列数 (A列からJ列まで)
  // 2行目から最終行まで、A列からJ列までのデータ範囲を取得
  const dataRange = targetSheet.getRange(startRow, 1, lastRow - startRow + 1, numColumns);
  const values = dataRange.getValues(); // その範囲の値を二次元配列として取得
  const backgroundColors = []; // 各セルの目標背景色を格納する二次元配列

  const redColor = '#FFCCCC'; // 薄い赤色のHEXコード
  const whiteColor = '#FFFFFF'; // 白色（デフォルト）のHEXコード

  // --- 各セルの色を決定 ---
  // 取得したデータを行ごと(i)にループ
  for (let i = 0; i < values.length; i++) {
    const rowBackgrounds = []; // 現在の行の背景色を格納する配列
    const rowValues = values[i]; // 現在の行の値の配列

    // 行の中のセルを列ごと(j)にループ
    for (let j = 0; j < numColumns; j++) {
      let cellColor = whiteColor; // デフォルトの色は白
      const cellValue = rowValues[j]; // 現在のセルの値

      // 列のインデックス(j)に基づいて、色を塗るか判定
      switch (j) {
        case 3: // D列 (インデックス 3)
        case 4: // E列 (インデックス 4)
        case 5: // F列 (インデックス 5)
          // 条件: セルの値が 0 または 空白か
          if (cellValue === '' || cellValue == null || Number(cellValue) === 0) {
            cellColor = redColor; // 条件に合えば赤
          }
          break; // 次の列へ

        case 7: // H列 (インデックス 7)
        case 8: // I列 (インデックス 8)
        case 9: // J列 (インデックス 9)
          // 条件: セルの値が 空白か
          if (cellValue === '' || cellValue == null) {
            cellColor = redColor; // 条件に合えば赤
          }
          break; // 次の列へ

        // default: // A, B, C, G列 (インデックス 0, 1, 2, 6) は何もしない (デフォルトの白のまま)
          // break;
      }
      rowBackgrounds.push(cellColor); // 決定したセルの色を行の配列に追加
    }
    backgroundColors.push(rowBackgrounds); // 行の色の配列を全体の配列に追加
  }

  // --- 背景色の適用 ---
  try {
    // 計算した背景色情報をシートの範囲にまとめて適用
    dataRange.setBackgrounds(backgroundColors);
    Logger.log('セルごとの条件付き書式を適用しました。');
    // 完了時のポップアップは表示しない
  } catch (e) {
    // エラーが発生した場合の処理
    Logger.log(`背景色の設定中にエラー: ${e}`);
    SpreadsheetApp.getUi().alert(`背景色の設定中にエラーが発生しました: ${e.message}`); // エラー時のみポップアップ
  }
}