const workRecordCheck = (e) => {
  if (!e) { e = 'シート2' };
  const wr = mainData_('wr').getSheetByName(e);
  const label = wr.getRange(1, 1, 1, wr.getLastColumn()).getValues().flat()
    .filter((value, index, array) => index > 0 && array.indexOf(value) == index);

  const keys = ['MEETING', 'LEAVE', 'SET'];
  const date = new Date();
  const start = 1;
  let end = 0;
  switch (e) {
    case '月前半用':
      end = 15;
      break;
    case '月後半用':
      date.setDate(0);
      end = date.getDate();
  }
  const obj = shiftObjectCheck(date);
  const map = [];
  for (let i = start; i <= end; i++) {
    const dd = String(i).padStart(2, '0');
    const array = label.flatMap(staff => {
      const set = obj[dd][staff]['SET'];
      switch (set) {
        case '休':
        case '希': return ['', '', '公休'];
        case '有': return ['', '', '有休'];
        case '当欠': return ['', '', set];
        case '病欠': return ['', '', set];
        case 'リ': return ['', '', 'リフレ'];
        case '忌': return ['', '', '忌引'];
        case 'M': return ['', '', 'ASB'];
        case '備': return ['09:00', '18:00', '準備日'];
        case '研': return ['10:00', '18:00', '研修'];
        default: return [obj[dd][staff]['MEETING'], obj[dd][staff]['LEAVE'], '登壇'];
      };
    });
    map.push(array);
  }
  wr.getRange(3, 2, map.length, map[0].length).setValues(map);
}


const addMenuWorkRecord = () => {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('追加メニュー');
  menu.addSubMenu(ui.createMenu('月前半用')
    .addItem('勤務実績反映', 'workRecordFirstHalf'));
  menu.addSubMenu(ui.createMenu('月後半用')
    .addItem('勤務実績反映', 'workRecordEoMonth')
    .addItem('勤務実績確認メール', 'eoMonthEmail'));
  menu.addToUi();
}

const eoMonthEmail = () => {
  const now = new Date();
  const wr = mainData_('wr');
  const wrs = wr.getSheetByName('月後半用');
  try { var write = wr.insertSheet(Utilities.formatDate(now, 'JST', 'yyyyMMdd')); }
  catch { var write = wr.getSheetByName(Utilities.formatDate(now, 'JST', 'yyyyMMdd')); }
  finally { write.getRange(1, 1, 1, 3).setValues([['e-mail', '件名', '本文']]); }
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
    body += infoBody(shift);
    body += '\n\nご不明点等あれば、富樫・川手までご連絡ください。\n\n';
    body += '富樫:070-1486-2940\n川手:080-2553-7330';
    //GmailApp.sendEmail(origindata[names.indexOf(a)][2], sub, body);
    Logger.log(`${origindata[names.indexOf(a)][2]}\n${sub}\n${body}`);
    write.getRange(write.getLastRow() + 1, 1, 1, 3).setValues([[origindata[names.indexOf(a)][2], sub, body]]);
  });
}

const sheetClear = () => {
  var ss = mainData_('wr');
  var sh = ss.getSheetByName('月後半用');
  sh.clearFormats();
}

const infoBody = (values) => {
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
}

const isNaN_ = (value) => {
  return typeof value === 'number' && value !== value;
}