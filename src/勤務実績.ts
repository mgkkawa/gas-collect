// 勤務実績表をシフト表とアサインシートから作成

const workRecordFirstHalf = () => { workRecord('前半'); };

const workRecordEoMonth = () => { workRecord('後半'); };

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

const workRecord = (timing = '後半') => {
  //timing = 勤務実績の提示タイミング　'前半'or'後半'

  //必要なスプレッドシートを各種取得
  // sh=シフト表
  // wr=勤務実績表
  // vc=会場連絡シート
  const sh = mainData('sh');
  const wr = mainData('wr');
  const vc = mainData('vc');
  const now = new Date();//現在時刻の取得
  //引数timing の情報を確認
  if (timing == '前半') {
    //timing=前半なら
    // 取得シートは'月前半用'
    // 終了日は'15日'
    var sheetname = '月前半用';
    now.setDate(15);
    var last = Utilities.formatDate(now, 'JST', 'MM/dd');
  } else {
    // timing='後半'なら
    // 取得シートは'月後半用'
    // 終了日は前月最終日
    var sheetname = '月後半用';
    now.setDate(0);
    var last = Utilities.formatDate(now, 'JST', 'MM/dd');
  }
  now.setDate(1);//開始日にリセット
  const one = Utilities.formatDate(now, 'JST', 'MM/dd');//該当月を文字列で取得
  const vcag = vc.getSheetByName('集約');//会場連絡シート内、'集約'タブを取得
  const vcd = vcag.getDataRange().getValues();//'集約'タブからデータを取得
  const vclabel = vcd.filter(values => values.includes('日程')).flat();
  //[dc,ho,mc,sc,st,en,nm]
  const vckeys = ['開催\n可否', 'メイン\n講師', 'サポート講師', 'サポート2', 'サポート3', 'サポート4', 'サポート5', '開始', '終了', '通し番号']
    .map(key => vclabel.indexOf(key));
  let vcdays = vcd.map(values => values.filter((value, index) => index == vclabel.indexOf('日程')))
    .flat().map(value => valueDate(value));
  const trim_start = vcdays.indexOf(one);
  const trim_end = vcdays.lastIndexOf(last);
  const trim_vcd = vcd.map(values => vckeys.map(key => values[key]))
    .filter((values, index) => index >= trim_start && index <= trim_end);
  vcdays = vcdays.filter((value, index) => index >= trim_start && index <= trim_end);
  // '集約'タブのデータを必要な情報にトリミング
  // [日付, 開催可否, 開始時間, 終了時間, メイン講師, サポート講師, サポート2, サポート3, サポート4, サポート5, 通し番号]

  //datから通し番号のみを取得
  const datnum = trim_vcd.map(values => values = values[values.length - 1]).flat();
  //貼り付け用の勤務実績表を取得。
  const wrs = wr.getSheetByName(sheetname);
  const wrd = wrs.getDataRange().getValues().map(values => values = values.map((value, index) => {
    if (index == 0) {
      if (Object.prototype.toString.call(value) == "[object Date]") {
        return value = Utilities.formatDate(value, 'JST', 'MM/dd');
      } else { return value; }
    } else { return value; }
  }));
  //勤務実績表のスタッフ並び順をラベルで取得
  const wlabel = wrd[0].filter(Boolean);
  //勤務実績表の日付を取得
  const wday = wrd.map(values => values = values.filter((value, index) => index == 0 && value != '')).flat();
  // シフト表の該当月シフトを取得。
  const shs = sh.getSheetByName(Utilities.formatDate(now, 'JST', 'yyyy.MM'));
  const shd = shs.getDataRange().getValues().map((values, index) => {
    if (index == 0) { values = values.map(value => valueDate(value)); }
    return values;
  }
  );
  const slabel = shd.filter(values => values.includes(one)).flat();
  const sstart = slabel.indexOf(one);
  const send = slabel.indexOf(last);
  // シフト表のデータを[名前, 1日, 2日, 3日,...]の形に整形
  const sdat = shd.map(values => values = values.filter(
    (value, index) => index == 0 || (index >= sstart && index <= send)));
  const sname = sdat.map(values => values = values[0]);
  // 勤務実績表に貼り付ける形に整形
  const wshift = wday.map((values, index) => {
    let numlist = [];
    values = wlabel.map(value => {
      now.setDate(index + 1);
      let time = [];
      const date = Utilities.formatDate(now, 'JST', 'MM/dd');
      const ind = sname.indexOf(value);
      const num = String(sdat[ind][index + 1]);
      if (numlist.indexOf(date) == -1) {
        numlist.push(date);
      }
      switch (num) {
        case '休':
        case '希': return value = ['', '', '公休'];
        case '有': return value = ['', '', '有休'];
        case '当欠': return value = ['', '', num];
        case '病欠': return value = ['', '', num];
        case 'リ': return value = ['', '', 'リフレ'];
        case '忌': return value = ['', '', '忌引'];
        case 'M': return value = ['', '', 'ASB'];
        case '備': return value = ['09:00', '18:00', '準備日'];
        case '研': return value = ['10:00', '18:00', '研修'];
        default:
          for (let i = vcdays.indexOf(date); i <= vcdays.lastIndexOf(date); i++) {
            if (trim_vcd[i].includes(num)) {
              if (numlist.includes(num)) {
                const starttime = trim_vcd[i][1];
                starttime.setMinutes(starttime.getMinutes() - 90);
                const endtime = trim_vcd[i][2];
                endtime.setMinutes(endtime.getMinutes() + 60);
                time.push([starttime, endtime]);
                numlist.push(num);
              } else {
                const starttime = trim_vcd[i][1];
                const endtime = trim_vcd[i][2];
                time.push([starttime, endtime]);
              }
            }
          }
          time = time.sort((top, bottom) => top[1].getTime() - bottom[1].getTime());
          const setstart = Utilities.formatDate(time[0][0], 'JST', 'HH:mm');
          const setend = Utilities.formatDate(time[time.length - 1][1], 'JST', 'HH:mm');
          return value = [setstart, setend, '登壇'];
      }
    }).flat();
    return values;
  });
  console.log(wshift);
  wrs.getRange(3, 2, wshift.length, wshift[0].length).setValues(wshift);
}


const eoMonthEmail = () => {
  const now = new Date();
  const wr = mainData('wr');
  const wrs = wr.getSheetByName('月後半用');
  try { var write = wr.insertSheet(Utilities.formatDate(now, 'JST', 'yyyyMMdd')); }
  catch { var write = wr.getSheetByName(Utilities.formatDate(now, 'JST', 'yyyyMMdd')); }
  finally { write.getRange(1, 1, 1, 3).setValues([['e-mail', '件名', '本文']]); }
  const wrd = wrs.getDataRange().getDisplayValues();
  const label = wrd.filter((a, x) => x == 0).flat();
  const namelist = label.filter(a => a != '');
  const month = now.getMonth();
  const sub = month + '月勤務実績';
  const origindata = staffData(['name', 'familyname', 'e-mail']);
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
  var ss = mainData('wr');
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