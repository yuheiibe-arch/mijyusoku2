/**
 * ログ関連の3つのプロセスを正しい順序で実行する「まとめ役」関数。
 * この関数にトリガーを1つだけ設定します。
 */
function runAllLoggingProcesses() {
  console.log("--- 全ログ処理を開始 ---");
  
  try {
    console.log("ステップ1: 新しい依頼をログに記録します...");
    logSpecialShifts();
  } catch (e) {
    console.error("ステップ1 (logSpecialShifts)でエラーが発生しました: " + e.toString());
  }

  try {
    console.log("ステップ2: 交渉結果を自動更新します...");
    updateNegotiationResults();
  } catch (e) {
    console.error("ステップ2 (updateNegotiationResults)でエラーが発生しました: " + e.toString());
  }
  
  try {
    console.log("ステップ3: 完了済みログを管理台帳へ転記します...");
    transferCheckedLogsToMasterSheet();
  } catch (e) {
    console.error("ステップ3 (transferCheckedLogsToMasterSheet)でエラーが発生しました: " + e.toString());
  }

  console.log("--- 全ログ処理を完了 ---");
}