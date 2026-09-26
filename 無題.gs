function testDetectTimegai() {
  console.log("▼▼▼ 時間外着信 検知テスト 開始 ▼▼▼");

  // 1. 現在の失敗するクエリのテスト
  const oldQuery = 'subject:(("【みんなにでんわ転送】時間外応答がありました") OR ("【みんなにでんわ転送】不在着信がありました")) newer_than:1d';
  console.log("【1】 変更前のクエリ: " + oldQuery);
  const oldThreads = GmailApp.search(oldQuery);
  console.log(" → ヒット数: " + oldThreads.length + " 件 （※これが 0 になるためスルーされていました）");

  // 2. 修正後のクエリのテスト
  const fixedQuery = '(subject:"時間外応答がありました" OR subject:"不在着信がありました") newer_than:1d';
  console.log("\n【2】 修正後のクエリ: " + fixedQuery);
  const newThreads = GmailApp.search(fixedQuery);
  console.log(" → ヒット数: " + newThreads.length + " 件");

  if (newThreads.length > 0) {
    const latestMsg = newThreads[0].getMessages()[0];
    console.log("\n✅ 最も新しい検知メール:");
    console.log("・送信日時: " + latestMsg.getDate());
    console.log("・件名: " + latestMsg.getSubject());
    console.log("・送信元: " + latestMsg.getFrom());
    console.log("※このメールが正しく拾えていることが確認できました！");
  } else {
    console.log("❌ 修正後のクエリでもヒットしませんでした。対象メールが別のアカウントに届いている可能性があります。");
  }

  console.log("\n▼▼▼ デバッグ完了 ▼▼▼");
}