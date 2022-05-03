const classtest = () => {
  const date = new Date();
  const as = mainData_('as');
  const sheet = as.getSheetByName(dateString(date, 'yyyyMM'));
};
const shifTtest = () => {
  const nh = mainData_('nh');
  const sheet = nh.getSheetByName('2022.04');
  const data = sheet.getRange('AH2:AP51').getValues();
  const staffs = Object.keys(staffObject_());
  const month = data[0][1];
  data.forEach(values => values.splice(1, 4));
  const sheet2 = nh.getSheetByName('現在シフト');
  const data2 = sheet2.getDataRange().getValues();
  let row;
  data2.forEach((values, index) => {
    if (row) {
      return;
    }
    if (values[1] == '4月') {
      row = index + 1;
      return;
    }
  });
  console.log(row);
  const label = data2.splice(0, 1).flat();
  const set_data = staffs.map(staff => data.filter(values => values[0] == staff).flat().filter((value, index) => index > 0))
    .map(values => values.map(value => value.replace(/日 |日,/g, ',').replace(/,$|日$/, '')));
  sheet2.getRange(row, label.indexOf('ANAMTG') + 1, set_data.length, set_data[0].length).setValues(set_data);
};
const logtest = (date = new Date()) => {
  const yyyy = date.getFullYear();
  const M = date.getMonth();
  const check_month = String(date.getMonth() + 1) + '月';
  date.setMonth(date.getMonth() + 1);
  date.setDate(0);
  const check_date = date.getDate();
  const spread = mainData_('nh');
  const sheets = spread.getSheets();
  let sheet_data;
  let add_data;
  sheets.forEach((sheet, number) => {
    switch (number) {
      case 0:
        sheet_data = sheet.getDataRange().getValues().filter(values => values[2] == check_month)
          .map(values => [1, 3, 4, 5, 6].map(index => {
            switch (true) {
              case index >= 3 && index <= 5:
                return values[index].replace(/日/g, '');
              default: return values[index];
            }
          }));
        break;
      case 1: add_data = sheet.getDataRange();
    }
  });
  const answer_sheet = spread.getSheetByName('フォームの回答 1');
  // const sheet_data = answer_sheet.getDataRange().getValues().filter(values => values[2] == check_month)
  //   .map(values => [1, 3, 4, 5, 6].map(index => {
  //     switch (true) {
  //       case index >= 3 && index <= 5:
  //         return values[index].replace(/日/g, '')
  //       default: return values[index]
  //     }
  //   }))
  const sheet_staffs = sheet_data.flatMap(data => data[0]);
  const staffs = Object.keys(staffObject_());
  const add_sheet = spread.getSheetByName(dateString(date, 'yyyy.MM'));
  // const add_data = add_sheet.getRange('AH:AO').getValues()
  //   .filter((values, index) =>)
};
const shifttest = (date = new Date()) => {
  if (!date) {
    date = new Date();
  }
  const yyyy = date.getFullYear();
  const M = date.getMonth() + 1;
  const check_month = String(date.getMonth() + 1) + '月';
  date.setMonth(date.getMonth() + 1);
  date.setDate(0);
  const check_date = date.getDate();
  const spreadsheet = mainData_('nh');
  const sheet = spreadsheet.getSheetByName('フォームの回答 1');
  const sheet_data = sheet.getDataRange().getValues().filter(values => values[2] == check_month)
    .map(values => [1, 3, 4, 5, 6].map(index => {
      switch (true) {
        case index >= 3 && index <= 5:
          return values[index].replace(/日/g, '');
        default: return values[index];
      }
    }));
  const sheet_staffs = sheet_data.flatMap(values => values[0]);
  const all_hopes = sheet_data.map(values => JSON.parse(`[${String(values[1])}]`));
  const all_paids = sheet_data.map(values => JSON.parse(`[${String(values[2])}]`));
  const all_refreshs = sheet_data.map(values => JSON.parse(`[${String(values[3])}]`));
  const staffs = Object.keys(staffObject_());
  const obj = {};
  for (let d = 1; d <= check_date; d++) {
    const obj_ = {};
    for (let staff of staffs) {
      const index = sheet_staffs.indexOf(staff);
      const hopes = all_hopes[index];
      const paids = all_paids[index];
      const refreshs = all_refreshs[index];
      switch (true) {
        case hopes.includes(d):
          obj_[staff] = new Work(false, '希');
          break;
        case paids.includes(d):
          obj_[staff] = new Work(false, '有');
          break;
        case refreshs.includes(d):
          obj_[staff] = new Work(false, 'リ');
          break;
        default:
          obj_[staff] = new Work(true, '備');
      }
    }
    obj[String(d).padStart(2, '0')] = obj_;
  }
  const obj_ = shifttest2_(date, obj);
  console.log(obj_);
  return;
  const database = mainData_('db').getSheetByName('シフト');
  let set_obj = JSON.stringify(obj_);
  const set_strings = new Array(Math.ceil(set_obj.length / 50000));
  if (set_strings.length > 1) {
    for (let l = 0; l < set_strings.length; l++) {
      set_strings.splice(l, 1, set_obj.slice(l * 50000, (l + 1) * 50000));
    }
  }
  else {
    set_strings[set_strings.length - 1] = set_obj;
  }
  const db_data = database.getDataRange().getValues().filter((values, index) => index == 0 || values[0] || values[1]);
  const last_row = database.getLastRow() - 1;
  const last_col = database.getLastColumn();
  const db_label = db_data.splice(0, 1).flat();
  const range = [];
  const set_data = db_data.map((values, index) => {
    if (values[db_label.indexOf('Origin')] && values[db_label.indexOf('yyyy')] == yyyy && values[db_label.indexOf('M')] == M) {
      return [true, , new Date(), yyyy, M].concat(set_strings);
    }
    else {
      return values;
    }
  });
  database.getRange(2, 1, last_row, last_col).clear();
  set_data.forEach((values, index) => {
    if (values[0]) {
      range.push(`A${index + 2}`);
    }
    else if (values[1]) {
      range.push(`B${index + 2}`);
    }
    database.getRange(index + 2, 1, 1, values.length).setValues([values]);
  });
  database.getRangeList(range).insertCheckboxes().check();
  shiftObjectAddInfo(date, obj_);
};
const shifttest2_ = (date = new Date(), shift = originCheck_(date)) => {
  const nh = mainData_('nh');
  const nh_sheet = nh.getSheetByName(dateString(date, 'yyyy.MM'));
  const nh_label = nh_sheet.getRange(1, 1, 1, nh_sheet.getLastColumn()).getValues().flat();
  const name_ind = nh_label.indexOf('氏名');
  const nh_data = nh_sheet.getDataRange().getValues().filter(values => values[name_ind] != '氏名' && values[name_ind] != '')
    .map(values => ['氏名', 'MTG', '病欠', '当欠', '研修'].map(key => values[nh_label.indexOf(key)]));
  nh_data.forEach(values => {
    const staff = values[0];
    const mtg = values[1].replace(/日/g, '').padStart(2, '0');
    const sick = JSON.parse(`[${values[2].replace(/日/g, '')}]`).filter(Boolean);
    const absent = JSON.parse(`[${values[3].replace(/日/g, '')}]`).filter(Boolean);
    const training = JSON.parse(`[${values[4].replace(/日/g, '')}]`).filter(Boolean);
    switch (true) {
      case mtg.length != 0:
        if (shift[mtg][staff].flag) {
          shift[mtg][staff] = new Work(false, 'M');
        }
        else {
          try {
            Browser.msgBox(`MTGチェック${staff}さんの${mtg}日シフトは\n"${shift[mtg][staff].num}"です`);
          }
          catch {
            Logger.log(`MTGチェック${staff}さんの${mtg}日シフトは\n"${shift[mtg][staff].num}"です`);
          }
        }
      case sick.length > 0:
        sick.forEach(d => {
          const dd = String(d).padStart(2, '0');
          if (shift[dd][staff].flag) {
            shift[dd][staff] = new Work(false, '病欠');
          }
          else {
            try {
              Browser.msgBox(`病欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff].num}"です`);
            }
            catch {
              Logger.log(`病欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff].num}"です`);
            }
          }
        });
      case absent.length > 0:
        absent.forEach(d => {
          const dd = String(d).padStart(2, '0');
          if (shift[dd][staff].flag) {
            shift[dd][staff] = new Work(false, '当欠');
          }
          else {
            try {
              Browser.msgBox(`当欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff].num}"です`);
            }
            catch {
              Logger.log(`当欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff].num}"です`);
            }
          }
        });
      case training.length > 0:
        training.forEach(d => {
          const dd = String(d).padStart(2, '0');
          if (shift[dd][staff].flag) {
            shift[dd][staff] = new Work(true, '研');
          }
          else {
            try {
              Browser.msgBox(`研修チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff].num}"です`);
            }
            catch {
              Logger.log(`研修チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff].num}"です`);
            }
          }
        });
    }
  });
  return shift;
};
function cas() {
  const venue_call = mainData_('vc');
  const vc_casting = venue_call.getSheetByName('キャスティング');
  const shift = infoCheck_();
  const staff_obj = staffObject_();
  const staffs = Object.keys(staff_obj);
  const vc_data = vc_casting.getDataRange().getValues();
  const obj = returnCastingObject_(vc_data);
  obj['display'] = display_;
  const MMs = Object.keys(obj);
  for (let MM of MMs) {
    const dds = Object.keys(obj[MM]);
    Logger.log(obj[MM].keys());
    for (let dd of dds) {
      for (let staff of staffs) {
        Logger.log(`${staff}:${shift[dd][staff]['FLAG']}`);
      }
    }
  }
}
const returnCastingObject_ = (array) => {
  const label = array[0];
  const days = array.flatMap(values => dateString(values[label.indexOf('日付')]));
  const trim_days = days.filter((value, index, array) => index > 0 && array.indexOf(value) == index);
  const obj = {};
  const to_month = trim_days[0].slice(0, 2);
  const next_month = trim_days[trim_days.length - 1].slice(0, 2);
  const to_obj = {};
  const next_obj = {};
  for (let day of trim_days) {
    const obj_ = {};
    const dd = String(day.match(/[\d].$/));
    for (let [index, values] of array.entries()) {
      if (index >= days.indexOf(day) && index <= days.lastIndexOf(day)) {
        const ven_ = {};
        ven_['VENUE'] = values[label.indexOf('会場名')];
        ven_['SET'] = values[label.indexOf('通番')];
        ven_['START'] = dateString(values[label.indexOf('開始チェック')], 'H:mm');
        ven_['MAIN'] = values[label.indexOf('変更後メンバー')];
        ven_['SUPPORT'] = values.filter((value, ind) => ind > label.indexOf('変更後メンバー') && ind < label.indexOf('SPLIT1') && value != '');
        obj_[`index${index}`] = ven_;
      }
      else if (index > days.lastIndexOf(day)) {
        break;
      }
      else {
        continue;
      }
    }
    if (String(day.match(/^[\d]./)) == to_month) {
      to_obj[dd] = obj_;
    }
    else {
      next_obj[dd] = obj_;
    }
  }
  obj[to_month] = to_obj;
  if (to_month != next_month) {
    obj[next_month] = next_obj;
  }
  return obj;
};
//アサインシートをオブジェクト化（途中）
const assign_object = () => {
  const yyyy = '2021';
  const MM = '12';
  const as = mainData_('as');
  const as_sheet = as.getSheetByName(yyyy + MM);
  const as_data = as_sheet.getDataRange().getValues();
  let ind = 0;
  const as_label = as_data.filter((values, index) => {
    if (values.includes('日程')) {
      ind = index;
      return true;
    }
  }).flat();
  const keys = JSON.parse(properties('assign_label'));
  Logger.log(keys);
  const trim = as_data.filter((values, index) => typeof values[as_label.indexOf('日程')] == 'object' && index > ind &&
    values[as_label.indexOf('開催\n可否')] != '中止' && values[as_label.indexOf('会場\n名称')] != '')
    .map(values => keys.map(key => values[as_label.indexOf(key)]));
  const trim_days = trim.flatMap(values => values[keys.indexOf('日程')]);
  const filter_days = trim_days.filter((value, index, array) => array.indexOf(value) == index);
  // const trim_venues = trim.flatMap(values => values[keys.indexOf('会場\n名称')])
  const column = [
    '会場\n名称', 'コース', '開催\n可否', '集合時間', '開始',
    '終了', '退店時間', '講師', '定員\n(半角)', '参加予定人数',
    '実参加\n人数', '更新日', '必要キャリー数', 'アサイン数', '誘導先店舗',
    'SAD在籍状況', 'SADサポート'
  ];
  const index = [
    'VENUE', 'CORSE', 'HOLD', 'MEETING', 'START',
    'FINISH', 'LEAVE', 'FLAG', 'LIMIT', 'NOP_PLAN',
    'ACT_PEOPLE', 'UPDATE', 'CARRY', 'ASSIGN', 'STORE',
    'SAD', 'SAD_SUPPORT'
  ];
  const set_col = [
    '開催No.', '都道府県', '主催者TEL', '会場TEL', '会場担当者名'
  ];
  const set_index = [
    'SERIAL_NO', 'AREA', 'TEL1', 'TEL2', 'MANAGER'
  ];
  const supporter = ['サポート講師', 'サポート2', 'サポート3', 'サポート4', 'サポート5'];
  filter_days.forEach(day => {
    const dd = day.match(/[\d].$/);
    const ddobj = {};
    const ven = {};
    const ser = {};
    for (let i = trim_days.indexOf(day); i <= trim_days.lastIndexOf(day); i++) {
      const obj = {};
      const mem_obj = {};
      const serial = trim[i][keys.indexOf('開催No.')];
      const venue = trim[i][keys.indexOf('会場\n名称')];
      const start = trim[i][keys.indexOf('開始')];
      const finish = trim[i][keys.indexOf('終了')];
      const flag = trim[i][keys.indexOf('講師')];
      obj['serial'] = serial;
      obj['hold'] = trim[i][keys.indexOf('開催\n可否')];
      obj['course'] = trim[i][keys.indexOf('コース')];
      if (flag == 'エムジー') {
        obj['meeting'] = timeStartMain_(start);
        mem_obj['main'] = trim[i][keys].indexOf('メイン\n講師');
      }
      else {
        obj['meeting'] = timeStartSup_(start);
        mem_obj['main'] = null;
      }
      obj['start'] = dateString(start, 'H:mm');
      obj['finish'] = dateString(finish, 'H:mm');
      obj['leave'] = timeEnd_(finish);
      obj['limit'] = trim[i][keys.indexOf('定員\n(半角)')];
      obj['plan'] = trim[i][keys.indexOf('実参加\n人数')];
      obj['update'] = trim[i][keys.indexOf('更新日')];
    }
  });
};
const shiftkakunin = () => {
  const obj = infoCheck_();
  const staff = '西村佳苗';
  const dd = '11';
  Logger.log(obj[dd][staff]);
};
const testEcho = () => {
  console.log('consoleテスト成功!!');
  Logger.log('Loggerテスト成功!!');
  Browser.msgBox('Browser.msgBoxテスト成功!!');
};
const propertySet = () => {
  const scripts = PropertiesService.getScriptProperties();
};
const propertieCheck = () => {
  const prop = PropertiesService.getScriptProperties();
  const keys = prop.getKeys();
};
const propertieDeliete = () => {
  const prop = PropertiesService.getScriptProperties();
  const keys = prop.getKeys();
};
const assignkeys_ = (obj) => {
  return Object.keys(obj).filter(key => key != 'sheet' && key != 'label' && key != 'maincol' && key != 'supcol');
};
