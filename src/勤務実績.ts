function workRecordCheck() {
  const date = new Date();
  let day = date.getDate()
  let sheetname
  if (day < 15) {
    day = 0
    sheetname = '月後半用'
  } else {
    day = 15
    sheetname = '月前半用'
  }
  date.setDate(day)
  const shift = new StaffWorkRecord(new AssignObject(date), date);
  const wr = mainData_('wr')
  const sheet = wr.getSheetByName(sheetname)
  sheet.getRange(3, 2, sheet.getLastRow() - 2, sheet.getLastColumn() - 1).clearContent()
  const staffs = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1).getValues().flat().filter(Boolean)
  const days = Object.keys(shift[staffs[0]]);
  days.splice(date.getDate())
  const data = days.map(day => staffs.flatMap(staff => returnWorkRecord_(shift[staff][day])))
  sheet.getRange(3, 2, data.length, data[0].length).setValues(data)
};
const returnWorkRecord_ = (obj) => {
  switch (obj.set_num) {
    case '希': return ['', '', '公休']
    case '有': return ['', '', '有休']
    case 'リ': return ['', '', 'リフレ']
    case 'M': return ['', '', 'ASB']
    case '当欠': return ['', '', '当欠']
    case '病欠': return ['', '', '病欠']
    case '忌引': return ['', '', '忌引']
    case '備': return ['9:00', '18:00', '準備日']
    case '研': return ['10:00', '18:00', '研修']
    default:
      if (obj.flag) {
        return [obj.meeting, obj.leave, '登壇']
      }
      return ['error', 'error', 'error']
  }
}
const addMenuWorkRecord_ = () => {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('追加メニュー');
  menu.addSubMenu(ui.createMenu('月前半用')
    .addItem('勤務実績反映', 'workRecordFirstHalf'));
  menu.addSubMenu(ui.createMenu('月後半用')
    .addItem('勤務実績反映', 'workRecordEoMonth')
    .addItem('勤務実績確認メール', 'eoMonthEmail'));
  menu.addToUi();
};
const eoMonthEmail_ = () => {
  const now = new Date();
  const wr = mainData_('wr');
  const wrs = wr.getSheetByName('月後半用');
  try {
    var write = wr.insertSheet(Utilities.formatDate(now, 'JST', 'yyyyMMdd'));
  }
  catch {
    var write = wr.getSheetByName(Utilities.formatDate(now, 'JST', 'yyyyMMdd'));
  }
  finally {
    write.getRange(1, 1, 1, 3).setValues([['e-mail', '件名', '本文']]);
  }
  const wrd = wrs.getDataRange().getDisplayValues();
  const label = wrd.filter((a, x) => x == 0).flat();
  const namelist = label.filter(a => a != '');
  const month = now.getMonth();
  const sub = month + '月勤務実績';
  const origindata = staffData_(['name', 'familyname', 'e-mail']);
  const names = origindata.map(a => a = a[0]);
  namelist.forEach(a => {
    var famName = origindata[names.indexOf(a)][1];
    var shift = wrd.map(b => b = b.filter((c, x) => x == label.indexOf(a) + 2 && c != '')).flat();
    var body = `${famName}さん\nお疲れ様です\n${month}月の勤務日数と休日の内訳をお送りします。\n\n`;
    body += infoBody_(shift);
    body += '\n\nご不明点等あれば、富樫・川手までご連絡ください。\n\n';
    body += '富樫:070-1486-2940\n川手:080-2553-7330';
    //GmailApp.sendEmail(origindata[names.indexOf(a)][2], sub, body)
    Logger.log(`${origindata[names.indexOf(a)][2]}\n${sub}\n${body}`);
    write.getRange(write.getLastRow() + 1, 1, 1, 3).setValues([[origindata[names.indexOf(a)][2], sub, body]]);
  });
};
const sheetClear_ = () => {
  var ss = mainData_('wr');
  var sh = ss.getSheetByName('月後半用');
  sh.clearFormats();
};
const infoBody_ = (values) => {
  var holiDayCount = 0;
  for (var i = 0; i < values.length; i++) {
    if (values[i] == '公休') {
      if (publicCount == null) {
        var publicCount = 0;
        var publicHoliday = i + 1 + '日';
      }
      else {
        publicHoliday += ',' + (i + 1) + '日';
      }
      ++publicCount;
      ++holiDayCount;
    }
    if (values[i] == '有休') {
      if (paidCount == null) {
        var paidCount = 0;
        var paidHoliday = i + 1 + '日';
      }
      else {
        paidHoliday += ',' + (i + 1) + '日';
      }
      ++paidCount;
    }
    if (values[i] == 'リフレッシュ') {
      if (refreshCount == null) {
        var refreshCount = 0;
        var refreshHoliday = i + 1 + '日';
      }
      else {
        refreshHoliday += ',' + (i + 1) + '日';
      }
      ++refreshCount;
    }
    if (values[i] == '病欠') {
      if (sickLieveCount == null) {
        var sickLieveCount = 0;
        var sickLieveHoliday = i + 1 + '日';
      }
      else {
        sickLieveHoliday += ',' + (i + 1) + '日';
      }
      ++sickLieveCount;
      ++holiDayCount;
    }
    if (values[i] == '当欠') {
      if (todayAbsentCount == null) {
        var todayAbsentCount = 0;
        var todayAbsentHoliday = i + 1 + '日';
      }
      else {
        todayAbsentHoliday += ',' + (i + 1) + '日';
      }
      ++todayAbsentCount;
      ++holiDayCount;
    }
  }
  var now = new Date();
  now.setDate(0);
  var day = now.getDate();
  var body = '勤務日数:' + (day - holiDayCount) + '日\n';
  body += '公休:' + publicCount + '日';
  if (paidCount != null) {
    body += '\n有休:' + paidCount + '日:該当日:' + paidHoliday;
  }
  if (refreshCount != null) {
    body += '\nリフレ:' + refreshCount + '日:該当日:' + refreshHoliday;
  }
  if (sickLieveCount != null) {
    body += '\n病欠:' + sickLieveCount + '日:該当日:' + sickLieveHoliday;
  }
  if (todayAbsentCount != null) {
    body += '\n当欠:' + todayAbsentCount + '日:該当日:' + todayAbsentHoliday;
  }
  return body;
};
