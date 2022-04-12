const visitCheck_ = () => {
  const calendar = CalendarApp.getCalendarById('mg-dsg@mg-k.co.jp');
  const today_event = calendar.getEventsForDay(start_time);
  let body = '';
  if (today_event.length == 0) {
    body = '本日の来社予定はありません。'
  } else {
    body = '本日の来社予定は下記の通りです。\n';
    body += 'スタッフさんをお待たせしないよう、\n';
    body += '万全の準備をしてお待ちしましょう。\n';

    today_event.forEach(event => {
      const title = event.getTitle();
      const time = valueDate(event.getStartTime());
      const description = event.getDescription();

      body += `\n${title}\n${time}\n${description}\n`;
    })
  }
  LINEWORKS.sendMsgRoom(setOptions_(), accountId_('room'), body);
}