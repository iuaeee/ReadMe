function searchMails() {

  const query = 'from:(do-not-reply@gift-cards.amazon.co.jp) after:2025/2/18 before:2025/4/19'
  const threads = GmailApp.search(query);

  // 書き出しを行うスプレッドシートとそのシート
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1');

  threads.forEach(thread => {

    const messages = thread.getMessages();

    messages.forEach(message => {

      let fromData = message.getFrom(); // 送信元
      let subject = message.getSubject(); // 件名
      let body = message.getPlainBody(); // 本文
      let attachments = message.getAttachments(); // 添付ファイル群（配列）

      let attachmentList = []; // 添付ファイルのファイル名格納用の配列

      if(attachments.length > 0){
        attachments.forEach(attachment => {

          let name = attachment.getName();

          attachmentList.push(name);
        });
      }

      attachmentList = attachmentList.join(',');

      let data = [fromData, subject, body, attachmentList];

      // 書き出し（行追加）実行
      sheet.appendRow(data);

    });
  });
}
