const workerTemp = () => { workercheck('出勤'); };
const holiDayTemp = () => { workercheck(); }

const sendCondition = (event) => {
  // if (!event) {
  //   const formResponses = FormApp.openById("1S8rMG-CuGyWV1-bnRtGz4EM-LTh7VJonObL_UvhBJtM").getResponses();
  //   const lastResponses = formResponses.length - 1;
  //   var event = formResponses[lastResponses].getItemResponses();
  // }
  const name = event[0].getResponse();
  const temp = event[1].getResponse();
  const condition = event[2].getResponse();


  if (temp >= 37.5 || condition == '悪い' || condition == '非常に悪い') {
    const symptom = event[3].getResponse();
    let msg = '検温報告にて対象者の報告がありました。\nスケジュールを確認しましょう。\n';
    msg += `\n${name}さん :${temp}℃\n体調:${condition}\n自覚症状:${symptom}`;

    LINEWORKS.sendMsgRoom(setOptions(), accountId('room'), msg);

    const eMailAddress = slimstaffData([name], ['e-mail']);
    const sub = '【重要】追加報告が必要です。';
    let body = mailBody(name);

    GmailApp.sendEmail(eMailAddress, sub, body, { from: accountId('option') });
  }
}

const workercheck = (work = null) => {
  const tm = mainData('tm');
  const tms = tm.getSheetByName('Check');
  const tmsd = tms.getDataRange().getValues()
    .filter((values, index) => index > 0 && values[2] != true)
    .map(values => values = values.filter((value, index) => index <= 2));

  if (work == '出勤') {
    var trim_tmsd = tmsd.filter(values => values[1] == work);
  } else {
    var trim_tmsd = tmsd;
  }
  if (trim_tmsd[0] == null) { return; }
  const sub = '【至急】検温結果報告をお願いします。';
  let body = '下記URLから検温結果報告をお願いします。\n\n';
  body += '検温結果報告フォーム\n';
  body += 'https://docs.google.com/forms/d/e/1FAIpQLScMWgzo6FBP0DOtW5i45CzZayUs1PUvRAq7PWsubD9z8w_lfA/viewform\n';
  const names = trim_tmsd.map(values => values = values[0]).flat();
  const mails = slimstaffData(names, ['name', 'e-mail', 'スタッフ名']);
  mails.forEach(values => {
    GmailApp.sendEmail(values[1], sub, body, { from: accountId('options') });
  });
}

const mailBody = (e) => {
  //本文を定義
  //emailFromは回答者名
  let body = e + 'さん\n\n';
  body += '検温報告ありがとうございます。\n';
  body += '体調が優れないところ、大変申し訳ございませんが\n';
  body += '追加報告が必要となりますので\n\n';
  body += '下記URLより、Googleフォームへの回答と\n';
  body += '行動履歴のご提出をお願い致します。\n\n';
  body += '体調不良者報告用フォーム\n';
  body += 'https://forms.gle/89rBdVrBLtGKBtUS9';
  body += '\n\n直近の行動履歴フォーム\n';
  body += 'https://forms.gle/2uH743StyBjS4rPx7';

  return body;
}