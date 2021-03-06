// casting=
// [キャスティング]タブからキャスティング情報を取得
// 共有用アサインシートへメンバーの転記
// シフト表へ該当通し番号を転記
// シフト未作成のアラートや人数不足のアラートを検討
// logclock=
// [LOGCLOCK]タブの情報を[集約]タブへ転記
// アサイン変更があった場合には未反映をチェック
// 表示があればアラートを検討
// vencall=
// [会場連絡]タブの情報を[集約]タブへ転記
// 会場連絡が必要な会場を表示
// 会場連絡が必要な会場があればアラートを検討
function writeLogClock() {
  const vc = mainData_('vc');
  const vcc = vc.getSheetByName('LOGCLOCK');
  const origin_clock_data = vcc.getDataRange().getValues();
  const clock_label = origin_clock_data.filter(values => values.includes('日程')).flat();
  const keys = ['日程', '可否', '会場', '開始',
    'メイン', 'サポート1', 'サポート2', 'サポート3', 'サポート4', 'サポート5',
    'Check1', 'Check2', 'Check3']
    .map(key => clock_label.indexOf(key));
  const trim_clock_data = origin_clock_data.map(values => keys.map(key => values[key]))
    .filter(values => {
      let true_check = values.filter((value, index) => index >= 10)
        .some(value => value == true);
      return values[1] != '中止' && true_check;
    });
  if (trim_clock_data[0] == null) {
    return;
  }
  const vcpos = vc.getSheetByName('転記');
  const origin_vcpos_data = vcpos.getDataRange().getDisplayValues();
  const vcposlabel = origin_vcpos_data.filter(values => values.includes('日程')).flat();
  const vcpos_keys = ['日程', '会場\n名称', 'シフト開始', '現場登録', 'お仕事スケジュール', 'キャスティング']
    .map(key => vcposlabel.indexOf(key));
  const trim_vcpos_data = origin_vcpos_data.map(values => vcpos_keys.map(key => values[key]));
  const vcpos_days = origin_vcpos_data.map(values => values[vcpos_keys[0]]).flat();
  trim_clock_data.forEach(values => {
    const day = dateString(values[0]);
    let start = vcpos_days.indexOf(day);
    const end = vcpos_days.lastIndexOf(day);
    while (start <= end) {
      if (values[1] != '中止') {
        const venue_check = (trim_vcpos_data[start][1] == values[2]);
        const start_check = (trim_vcpos_data[start][2] == dateString(values[3], 'H:mm'));
        if (venue_check && start_check) {
          const no = keys.length;
          const main = values[4];
          const sup = values.filter((value, index) => index > no - 9 && index < no - 3 && value != '');
          const true_check = values.filter((value, index) => index >= no - 3)
            .map((value, index) => {
              value = Boolean(value);
              let val = (value != true);
              if (val && index < 3) {
                if (index == 0) {
                  return trim_vcpos_data[start][3];
                }
                else {
                  return trim_vcpos_data[start][4];
                }
              }
              else {
                return value;
              }
            });
          try {
            vcpos.getRange(start + 1, vcposlabel.indexOf('メイン\n講師') + 1).setValue(main);
            vcpos.getRange(start + 1, vcposlabel.indexOf('サポート講師') + 1, 1, sup.length).setValues([sup]);
          }
          catch (e) {
            true_check.splice(2, 1, false);
            Browser.msgBox('アサイン数を再度確認してください。');
          }
          finally {
            vcpos.getRange(start + 1, vcpos_keys[3] + 1, 1, true_check.length).insertCheckboxes().setValues([true_check]);
            break;
          }
        }
      }
      ++start;
    }
  });
  const lastRow = vcpos.getLastRow();
  const check1_col = `${NumToA1(clock_label.indexOf('Check1') + 1)}2:${NumToA1(clock_label.indexOf('Check2') + 1)}${lastRow}`;
  const check3_col = `${NumToA1(clock_label.indexOf('Check3') + 1)}2:${NumToA1(clock_label.indexOf('Check3') + 1)}${lastRow}`;
  vcc.getRangeList([check1_col, check3_col]).uncheck();
}
function writeVenCall() {
  const vc = mainData_('vc');
  const vcs = vc.getSheetByName('会場連絡');
  const origin_vc_data = vcs.getDataRange().getValues();
  const vclabel = origin_vc_data.filter(values => values.includes('日付')).flat();
  const vckeys = ['Check', '日付', '会場', '開始', '施設担当者', 'スクリーン', '前回入館', '前回引継ぎ', '人数', '施設担当者（今回）', 'スクリーン（今回）', '入館', '次回引継ぎ']
    .map(key => vclabel.indexOf(key));
  //Checkは消える。
  const vcsd = origin_vc_data.map(values => vckeys.map((key, index) => {
    if (index != 3) {
      return dateString(values[key]);
    }
    else {
      return dateString(values[key], 'H:mm');
    }
  })).filter(values => values[0] == true).map(values => values.filter((value, index) => index > 0));
  const vcpos = vc.getSheetByName('転記');
  const vcposd = vcpos.getDataRange().getValues();
  const vcposlabel = vcposd.filter(values => values.includes('日程')).flat();
  const vcposkeys = ['会場\n名称', '開始'].map(key => vcposlabel.indexOf(key));
  const vcpos_write_keys = ['会場連絡', '施設担当者'].map(key => vcposlabel.indexOf(key) + 1);
  const vcposdays = vcposd.map(values => values.filter((value, index) => index == vcposlabel.indexOf('日程')))
    .flat().map(value => dateString(value));
  const as = mainData_('as');
  const ass = as.getSheetByName(Utilities.formatDate(start_time, 'JST', 'yyyyMM'));
  const assd = ass.getDataRange().getValues();
  const asslabel = assd.filter(values => values.includes('日程')).flat();
  const asskeys = ['会場\n名称', '開始'].map(key => asslabel.indexOf(key));
  const assdays = assd.map(values => values.filter((value, index) => index == asslabel.indexOf('日程'))).flat().map(value => dateString(value));
  const ass_write_keys = ['参加予定人数', '確認日'].map(key => asslabel.indexOf(key) + 1);
  const trim_assd = assd.map(values => asskeys.map((key, index) => values[key]));
  vcsd.forEach(values => {
    const day = values[0];
    const venue = values[1];
    const start = values[2];
    let row = vcposdays.indexOf(day);
    let end = vcposdays.lastIndexOf(day);
    while (row <= end) {
      const venue_check = (vcposd[row][vcposkeys[0]] == venue);
      const start_check = (dateString(vcposd[row][vcposkeys[1]], 'H:mm') == start);
      if (venue_check && start_check) {
        const old_data = values.filter((value, index) => index >= 3 && index <= 6);
        const new_data = values.filter((value, index) => index >= 8);
        const set_data = new_data.map((value, index) => {
          if (value == '') {
            return old_data[index];
          }
          else {
            return value;
          }
        });
        vcpos.getRange(row + 1, vcpos_write_keys[0]).insertCheckboxes().check();
        vcpos.getRange(row + 1, vcpos_write_keys[1], 1, set_data.length).setValues([set_data]);
      }
      ++row;
    }
    row = assdays.indexOf(day);
    end = assdays.lastIndexOf(day);
    while (row <= end) {
      const as_venue_check = (trim_assd[row][0] == venue);
      const as_start_check = (dateString(trim_assd[row][1], 'H:mm') == start);
      if (as_venue_check && as_start_check) {
        const check_day = Utilities.formatDate(start_time, 'JST', 'M/d');
        ass.getRange(row + 1, ass_write_keys[0]).setValue(values[7]);
        ass.getRange(row + 1, ass_write_keys[1]).setValue(check_day);
      }
      ++row;
    }
  });
  vcs.getRange(2, vclabel.indexOf('Check') + 1, vcs.getLastRow() - 1).uncheck();
  vcs.getRange(2, vclabel.indexOf('Check') + 2, vcs.getLastRow() - 1, 5).clear();
}
function writeCasting() {
  const date = new Date();
  date.setDate(1);
  const vc = mainData_('vc');
  const casting = vc.getSheetByName('キャスティング');
  const cas_data = casting.getDataRange().getValues();
  const cas_obj = new Venuecall(cas_data, ['会場\n名称'], ['メイン\n講師', 'サポート講師']);
  if (Object.keys(cas_obj).filter(key => key != 'label').length == 0) {
    try {
      Browser.msgBox('対象のキャスティング情報がありませんでした。');
    }
    finally {
      Logger.log('対象のキャスティング情報がありませんでした。');
      return;
    }
  }
  const main_assign = new AssignObject();
  const main_table = new ShiftTable();
  const to_ = dateString(date, 'MM/');
  let sub_assign;
  let sub_table;
  if (cas_obj.check()) {
    date.setMonth(date.getMonth() + 1);
    sub_assign = new AssignObject(date);
    sub_table = new ShiftTable(date);
  }
  for (let row in cas_obj) {
    if (row == 'label') {
      continue;
    }
    let assign = main_assign;
    let table = main_table;
    const obj = cas_obj[row];
    const day = obj.date;
    if (!day.includes(to_)) {
      assign = sub_assign;
      table = sub_table;
    }
    const venue = obj.venue;
    const start = obj.start;
    const ascheck = assign.rowNum(day, venue, start);
    const main = `${NumToA1(assign.maincol + 1)}${ascheck[0]}`; //メイン講師の貼り付け範囲
    const supind = assign.supcol;
    if (!obj.mg_flag) {
      assign.setValue(main, obj.main)
    }
    let support;
    if (obj.support.length > 1) {
      support = `${NumToA1(supind + 1)}${ascheck[0]}:${NumToA1(supind + obj.support.length)}${ascheck[0]}`;
    } else {
      support = `${NumToA1(supind + 1)}${ascheck[0]}`;
    }
    assign.setValues(support, [obj.support])
    obj.support.push(obj.main);
    const range = [];
    obj.support.filter(Boolean).forEach(staff => {
      range.push(table.getCell(day, staff));
    });
    table.sheet.getRangeList(range).setValue(ascheck[1]);
    obj.support.filter(Boolean).forEach(staff => {
      Logger.log(`staff:${staff}\nset_num:${ascheck[1]}\nassign:${ascheck[0]}\ntable:${table.getCell(day, staff)}`);
    });
  }
}
function writeSupply() {
  const vc = mainData_('vc');
  const vcsu = vc.getSheetByName('備品お渡しリスト');
  const vcsud = vcsu.getDataRange().getDisplayValues();
  const members = [getName_(), vcsud[2][2]];
  const vcsu_set = [
    //['開催日', '会場名', '開始時間', '配備先', '配備予定日', 'コメント',
    // '準備物1', '配備数1', '準備物2', '配備数2', '準備物3', '配備数3', '準備物4', '配備数4', '準備物5', '配備数5', '準備物6', '配備数6', '準備物7', '配備数7']
    [vcsud[13][0], vcsud[13][1], vcsud[13][4], vcsud[16][1], vcsud[16][2], vcsud[19][1],
    vcsud[22][1], vcsud[22][3], vcsud[23][1], vcsud[23][3], vcsud[24][1], vcsud[24][3], vcsud[25][1], vcsud[25][3], vcsud[26][1], vcsud[26][3], vcsud[27][1], vcsud[27][3], vcsud[28][1], vcsud[28][3]],
    [vcsud[31][0], vcsud[31][1], vcsud[31][4], vcsud[34][1], vcsud[34][2], vcsud[37][1],
    vcsud[40][1], vcsud[40][3], vcsud[41][1], vcsud[41][3], vcsud[42][1], vcsud[42][3], vcsud[43][1], vcsud[43][3], vcsud[44][1], vcsud[44][3], vcsud[45][1], vcsud[45][3], vcsud[46][1], vcsud[46][3]],
    [vcsud[49][0], vcsud[49][1], vcsud[49][4], vcsud[52][1], vcsud[52][2], vcsud[55][1],
    vcsud[58][1], vcsud[58][3], vcsud[59][1], vcsud[59][3], vcsud[60][1], vcsud[60][3], vcsud[61][1], vcsud[61][3], vcsud[62][1], vcsud[62][3], vcsud[63][1], vcsud[63][3], vcsud[64][1], vcsud[64][3]],
    [vcsud[67][0], vcsud[67][1], vcsud[67][4], vcsud[70][1], vcsud[70][2], vcsud[73][1],
    vcsud[76][1], vcsud[76][3], vcsud[77][1], vcsud[77][3], vcsud[78][1], vcsud[78][3], vcsud[79][1], vcsud[79][3], vcsud[80][1], vcsud[80][3], vcsud[81][1], vcsud[81][3], vcsud[82][1], vcsud[82][3]]
  ].map(values => values.filter((value, index) => {
    if (index <= 5) {
      return true;
    }
    else {
      return value != '';
    }
  })).filter(values => values.length > 6);
  const vcpos = vc.getSheetByName('転記');
  const vcposd = vcpos.getDataRange().getValues();
  const vcpos_label = vcposd.filter(values => values.includes('日程')).flat();
  const vcpos_days = vcposd.map(values => [vcpos_label.indexOf('日程')].map(key => values[key])).flat().map(value => dateString(value));
  const vcpos_keys = ['会場\n名称', '開始'].map(key => vcpos_label.indexOf(key));
  const trim_vcposd = vcposd.map(values => vcpos_keys.map(key => values[key]));
  const col_list = ['お渡し', '配備担当', '配備先', '準備物1'].map(key => vcpos_label.indexOf(key) + 1);
  let error_count = 0;
  for (let i in vcsu_set) {
    const day = vcsu_set[i].shift();
    const venue = vcsu_set[i].shift();
    const time = vcsu_set[i].shift();
    if (/[\d]/.test(vcsu_set[i][0])) {
      try {
        Browser.msgBox(`会場${Number(i) + 1}の配備先が入力されていません。`);
      }
      finally {
        Logger.log(`会場${Number(i) + 1}の配備先が入力されていません。`);
      }
      ++error_count;
      continue;
    }
    const staff = vcsu_set[i].splice(0, 3);
    if (vcsu_set[i].some(value => /[\d]?[\d]/.test(String(value)))) { }
    else {
      try {
        Browser.msgBox(`会場${Number(i) + 1}の情報が不足しています。`);
      }
      finally {
        Logger.log(`会場${Number(i) + 1}の情報が不足しています。`);
      }
      ++error_count;
      continue;
    }
    let start = vcpos_days.indexOf(day);
    const end = vcpos_days.lastIndexOf(day);
    if (error_count > 0) {
      Browser.msgBox('エラー箇所を修正して再度実行してください。');
      Logger.log('エラー箇所を修正して再度実行してください。');
      return;
    }
    else {
      while (start <= end) {
        const venue_check = (trim_vcposd[start][0] == venue);
        const time_check = (dateString(trim_vcposd[start][1], 'H:mm') == time);
        if (venue_check && time_check) {
          const row = start + 1;
          col_list.forEach((value, index) => {
            switch (index) {
              case 0:
                vcpos.getRange(row, value).insertCheckboxes().check();
                break;
              case 1:
                vcpos.getRange(row, value, 1, members.length).setValues([members]);
                break;
              case 2:
                vcpos.getRange(row, value, 1, staff.length).setValues([staff]);
                break;
              case 3:
                vcpos.getRange(row, value, 1, vcsu_set[i].length).setValues([vcsu_set[i]]);
                break;
            }
          });
          break;
        }
        ++start;
      }
      Browser.msgBox('備品お渡し情報を登録しました。');
    }
  }
  vcsu.getRangeList(['C3:D10', 'E7:G10', 'D23:D29', 'D41:D47', 'D59:D65', 'D77:D83', 'B20:D20', 'B38:D38', 'B56:D56', 'B74:D74']).clearContent();
}
const addMaster = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 2;
  const shname = year + String(month).padStart(2, '0');
  const ass = mainData_('as').getSheetByName(shname);
  switch (true) {
    case ass.getRange(1, 1).isBlank():
      var firstRow = ass.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
      if (!firstRow) {
        var firstRow = 1;
      }
    case ass.getRange(firstRow + 1, 2).isBlank():
      var dataStartRow = ass.getRange(firstRow + 1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
      var dataStartA1 = ass.getRange(dataStartRow, 1).getA1Notation();
      var dataEndA1 = ass.getRange(1, ass.getLastColumn()).getA1Notation().replace(/\d/, '');
      if (!dataStartRow) {
        var dataStartRow = firstRow + 1;
      }
      if (!dataStartA1) {
        var dataStartA1 = ass.getRange(dataStartRow, 1).getA1Notation();
      }
      if (!dataEndA1) {
        var dataEndA1 = ass.getRange(1, ass.getLastColumn()).getA1Notation().replace(/\d/, '');
      }
    default:
      var label = ass.getRange(firstRow, 1, 1, ass.getLastColumn()).getValues().flat();
      var range = dataStartA1 + ':' + dataEndA1;
      var func = '=IMPORTRANGE("1m93CFX1uG67bO6c5xbSGoV5Bm0xNbfO0QAkE7nQqO5c","' + shname + '!' + range + '")';
      var serial = '通し番号';
      var col = NumToA1(label.indexOf(serial) + 1);
      col = col + '2:' + col;
      var numFunc = '=ARRAYFORMULA(IF(' + col + '<>"",TO_TEXT(' + col + '),""))';
  }
  const vc = mainData_('vc'); //[開発用]新会場連絡シート
  try {
    var vcadd = vc.insertSheet(shname);
  }
  catch (e) {
    var vcadd = vc.getSheetByName(shname);
  }
  finally {
    vcadd.hideSheet().getRange(1, 1, 1, label.length).setValues([label]);
    vcadd.getRange(2, 1).setValue(func);
    vcadd.getRange(1, label.length + 1, 2).setValues([[serial], [numFunc]]);
  }
  const vcag = vc.getSheetByName('集約');
  const vcagLastRow = vcag.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const vcaglabel = vcag.getRange(2, 1, 1, vcag.getLastColumn()).getValues().flat();
  vcaglabel.splice(vcaglabel.indexOf(serial) + 1);
  const colList = vcaglabel.map((value, x) => {
    switch (true) {
      case value == serial:
        return 'Col' + label.length;
      case value != '':
      case label.indexOf(value) != -1: return 'Col' + label.indexOf(value);
      default: return '';
    }
  });
  func = '=QUERY({\'' + shname + '\'!B2:' + NumToA1(label.length + 1) + '},"select ' + colList + ' where Col2 is not null")';
  vcag.getRange(vcagLastRow + 1, 1).setValue(func);
  addressUPDATE_(vcag);
};
const toDay_ = () => {
  const setFullYear = start_time.getFullYear();
  const setMonth = start_time.getMonth() + 1;
  const setDate = start_time.getDate();
  const vcag = mainData_('vc').getSheetByName('集約');
  const vcpos = mainData_('vc').getSheetByName('転記');
  vcag.getRange(1, 1, 1, 3).setValues([[setFullYear, setMonth, setDate]]);
  return Logger.log('toDay_:コンプリート');
};
const suiteCase_ = () => {
  const sc = mainData_('sc');
  const scs = sc.getSheetByName(dateString(start_time, 'yyyy.MM'));
  const scsd = scs.getDataRange().getValues();
  const scsd_label = scsd.filter(values => values.includes('所持')).flat()
    .map((value, index) => {
      if (index < 6) {
        return '';
      }
      else {
        return dateString(value);
      }
    });
  let date = start_time.getDate();
  if (date <= 10) {
    date = 1;
  }
  else if (date <= 20) {
    date -= 5;
  }
  start_time.setDate(date);
  const today = dateString(start_time);
  const today_ind = scsd_label.indexOf(today);
  for (let i in scsd) {
    if (scsd[i].includes('A')) {
      var trim_row = Number(i);
      break;
    }
  }
  const trim_data = scsd.filter((values, index) => index >= trim_row && index <= trim_row + 25)
    .map(values => values.filter((value, index) => index >= today_ind));
  const vc = mainData_('vc');
  const vcsuite = vc.getSheetByName('スーツケース②');
  vcsuite.getRange(3, 2, trim_data.length, trim_data[0].length).clearContent().setValues(trim_data);
  return Logger.log('suiteCase_:コンプリート');
};
const holdCheck_ = () => {
  const today = dateString(start_time);
  const get_date = start_time.getDate() + 7;
  start_time.setDate(get_date);
  const weekday = start_time.getDay();
  if (weekday == 0) {
    start_time.setDate(get_date + 1);
  }
  else if (weekday == 6) {
    start_time.setDate(get_date + 2);
  }
  const check_date = dateString(start_time);
  const weekdays = ['(日)', '(月)', '(火)', '(水)', '(木)', '(金)', '(土)'];
  const vcag = mainData_('vc').getSheetByName('集約');
  const vcagd = vcag.getDataRange().getValues();
  const vcag_label = vcagd.filter(values => values.includes('日程')).flat();
  const vcag_days = vcagd.map(values => values.filter((value, index) => index == vcag_label.indexOf('日程'))).flat().map(value => dateString(value));
  const vcag_keys = ['日程', '会場\n名称', '会場\n担当者名', '開始']
    .map(key => vcag_label.indexOf(key));
  const start_day = vcag_days.indexOf(today);
  const end_day = vcag_days.lastIndexOf(check_date);
  const vcag_hold = vcag_label.indexOf('開催\n可否');
  const trim_vcag_data = vcagd.filter((values, index) => index >= start_day && index <= end_day &&
    values[vcag_hold] == '' && values.includes('エムジー'))
    .map(values => vcag_keys.map(key => values[key]));
  const manegers = trim_vcag_data.map((values, index) => values[2]).flat()
    .filter((value, index, self) => self.indexOf(value) === index);
  manegers.forEach(value => {
    let body = '';
    if (value == '') {
      body += '担当者不明分\n';
    }
    else {
      body += `${value}様\n`;
    }
    body += '\nお世話になっております。\nエムジーの大山でございます。\n\n';
    const vens = trim_vcag_data.filter(values => values.includes(value));
    vens.forEach(values => {
      const day = dateString(new Date(values[0]), 'M/d');
      const week = weekdays[new Date(values[0]).getDay()];
      const venue = `・${values[1]}`;
      body += `${day} ${week}\n${venue}\n`;
    });
    body += '\nでのスマホ教室の開催可否はいかがでしょうか？\n';
    body += 'お忙しいところお手数ですが、\nご教示頂ければ幸いです。';
    LINEWORKS.sendMsg(setOptions_(), accountId_('大山夏美'), body);
  });
  // LINEWORKS.sendMsg(setOptions_(), accountId_(''), body)
};
const monthReset_ = (date = null) => {
  if (date == null) {
    date = new Date();
  }
  date.setMonth(date.getMonth() + 1);
  date.setDate(0);
  var nh = mainData_('nh');
  var nhs = nh.getSheetByName(Utilities.formatDate(date, 'JST', 'yyyy.MM'));
  const origin_nh_data = nhs.getDataRange().getValues();
  const nh_names = origin_nh_data.map(values => values = values[0]).flat();
  const nh_dat = origin_nh_data.filter((values, index) => index > 0 && index <= 50)
    .map(values => values.filter((value, index) => index > 0 && index <= date.getDate()));
  var shift_sheet = mainData_('sh').getSheetByName(Utilities.formatDate(date, 'JST', 'yyyy.MM'));
  const origin_sh_data = shift_sheet.getDataRange().getValues().map((values, index) => {
    if (index == 0) {
      return values.map(value => dateString(value));
    }
    else {
      return values;
    }
  });
  const sh_label = origin_sh_data.filter(values => values.includes(Utilities.formatDate(date, 'JST', 'MM/dd'))).flat();
  const sh_row = origin_sh_data.map(values => values = values[0]).flat().indexOf('スタッフ') + 2;
  date.setDate(1);
  shift_sheet.getRange(sh_row, sh_label.indexOf(Utilities.formatDate(date, 'JST', 'MM/dd')) + 1, nh_dat.length, nh_dat[0].length)
    .setValues(nh_dat);
};
const shiftSet_ = (date = null) => {
  if (date == null) {
    date = new Date();
  }
  var vc = mainData_('vc');
  var vcag = vc.getSheetByName('集約');
  date.setDate(1);
  const start = Utilities.formatDate(date, 'JST', 'MM/dd');
  date.setMonth(date.getMonth() + 1);
  date.setDate(0);
  const end = Utilities.formatDate(date, 'JST', 'MM/dd');
  const vcagd = vcag.getDataRange().getDisplayValues();
  const vclabel = vcagd.filter(values => values.includes('日程')).flat();
  const vcdays = vcagd.map(values => values[vclabel.indexOf('日程')]).flat();
  const vc_keys = ['日程', '通し番号', 'メイン\n講師', 'サポート講師', 'サポート2', 'サポート3', 'サポート4', 'サポート5',]
    .map((key => vclabel.indexOf(key))).flat();
  const start_row = vcdays.indexOf(start);
  const end_row = vcdays.lastIndexOf(end);
  const trim_vcagd = vcagd.map((values) => vc_keys.map(key => values[key]).filter(value => value != ''))
    .filter((values, index) => index >= start_row && index <= end_row);
  const trim_vcdays = trim_vcagd.map(values => values[0]).flat();
  const sh = mainData_('sh');
  const shs = sh.getSheetByName(Utilities.formatDate(date, 'JST', 'yyyy.MM'));
  const shsd = shs.getDataRange().getValues();
  const shdays = shsd.filter(values => values.includes('ス')).flat().map(value => dateString(value));
  for (let i in shsd) {
    if (shsd[i].includes('スタッフ')) {
      var staff_col = shsd[i].indexOf('スタッフ');
    }
  }
  const shstaffs = shsd.map(values => values[staff_col]).flat();
  trim_vcagd.forEach(values => {
    for (let i in values) {
      if (Number(i) > 1) {
        let row = shstaffs.indexOf(values[i]) + 1;
        let col = shdays.indexOf(values[0]) + 1;
        let no = values[1];
        shs.getRange(row, col).setValue(no);
      }
    }
  });
};
const addUi = () => {
  SpreadsheetApp.getUi()
    .createMenu('追加メニュー')
    .addSeparator()
    .addItem('LOGCLOCK', 'writeLogClock')
    .addSeparator()
    .addItem('会場連絡', 'writeVenCall')
    .addSeparator()
    .addItem('備品お渡しリスト', 'writeSupply')
    .addSeparator()
    .addItem('翌月分マスタ', 'assaignsheet');
};
