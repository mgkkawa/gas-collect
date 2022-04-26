
const addShiftObject = (date = new Date()) => {
  const shift = shiftObjectCheck(date);

  const nh = mainData_('nh');
  const nh_sheet = nh.getSheetByName(valueDate(date, 'yyyy.MM'));
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

    if (mtg.length != 0) {
      const check = shift[mtg][staff]['FLAG']
      if (check) {
        shift[mtg][staff]['FLAG'] = false;
        shift[mtg][staff]['SET'] = 'M';
      } else {
        try { Browser.msgBox(`MTGチェック${staff}さんの${mtg}日シフトは\n"${shift[mtg][staff]['SET']}"です`) }
        catch { Logger.log(`MTGチェック${staff}さんの${mtg}日シフトは\n"${shift[mtg][staff]['SET']}"です`) };
      }
    }
    if (sick.length > 0) {
      sick.forEach(d => {
        const dd = String(d).padStart(2, '0');
        const check = shift[dd][staff]['FLAG'];
        if (check) {
          shift[dd][staff]['FLAG'] = false;
          shift[dd][staff]['SET'] = '病欠';
        } else {
          try { Browser.msgBox(`病欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
          catch { Logger.log(`病欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) };
        }
      });
    };
    if (absent.length > 0) {
      absent.forEach(d => {
        const dd = String(d).padStart(2, '0');
        const check = shift[dd][staff]['FLAG'];
        if (check) {
          shift[dd][staff]['FLAG'] = false;
          shift[dd][staff]['SET'] = '当欠';
        } else {
          try { Browser.msgBox(`当欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
          catch { Logger.log(`当欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) };
        }
      });
    };
    if (training.length > 0) {
      training.forEach(d => {
        const dd = String(d).padStart(2, '0');
        const check = shift[dd][staff]['FLAG'];
        if (check) {
          shift[dd][staff]['SET'] = '研';
        } else {
          try { Browser.msgBox(`研修チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
          catch { Logger.log(`研修チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) };
        }
      });
    };
  });
  shiftObjectUpdate(shift, date);
}

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
  const trim = as_data.filter((values, index) =>
    typeof values[as_label.indexOf('日程')] == 'object' && index > ind &&
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
  ]
  const set_col = [
    '開催No.', '都道府県', '主催者TEL', '会場TEL', '会場担当者名'
  ]
  const set_index = [
    'SERIAL_NO', 'AREA', 'TEL1', 'TEL2', 'MANAGER'
  ]

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
        obj['meeting'] = shiftMainStart(start);
        mem_obj['main'] = trim[i][keys].indexOf('メイン\n講師');
      } else {
        obj['meeting'] = shiftSupStart(start);
        mem_obj['main'] = null;
      }
      obj['start'] = valueDate(start, 'H:mm');
      obj['finish'] = valueDate(finish, 'H:mm');
      obj['leave'] = shiftEnd(finish);
      obj['limit'] = trim[i][keys.indexOf('定員\n(半角)')];
      obj['plan'] = trim[i][keys.indexOf('実参加\n人数')];
      obj['update'] = trim[i][keys.indexOf('更新日')];


    }
  })
}

const shiftkakunin = () => {
  const obj = shiftObjectCheck();

  const staff = '西村佳苗';
  const dd = '11';

  Logger.log(obj[dd][staff]);
}
const shiftCheck = (date = new Date()) => {
  const MM = valueDate(date, 'MM');
  const obj = shiftObjectCheck();

  const wr = mainData_('wr');
  const wr_sheet = wr.getSheetByName('月前半用');
  const wr_label = wr_sheet.getRange(1, 2, 1, wr_sheet.getLastColumn() - 1).getValues().flat().filter(Boolean);

  const array = [];
  for (let i = 1; i <= 15; i++) {
    const dd = String(i).padStart(2, '0');
    const push_array = wr_label.flatMap(staff => {
      const obj_staff = obj[MM][dd][staff];
      switch (obj_staff['set_value']) {
        case '備': return ['9:00', '18:00', '準備日'];
        case '休': return ['', '', '公休'];
        case '希': return ['', '', '公休'];
        case '有': return ['', '', '有休'];
        case 'リ': return ['', '', 'リフレ'];
        case 'M': return ['', '', 'ASB'];
        case '研': return ['10:00', '18:00', '研修'];
        case '当欠': return ['', '', '当欠'];
        case '病欠': return ['', '', '病欠'];
        default:
          const keys = Object.keys(obj_staff);
          if (keys.includes('info')) {
            const info_keys = Object.keys(obj_staff['info']);
            const start = obj_staff['info'][info_keys[0]]['meeting'];
            const end = obj_staff['info'][info_keys[info_keys.length - 1]]['leave'];
            return [start, end, '登壇'];
          }
      }
    })
    array.push(push_array);
  }
  wr.getSheetByName('シート2').getRange(3, 2, array.length, array[0].length)
    .setValues(array);
}
const shiftinfo = (date = new Date()) => {
  const yyyy = valueDate(date, 'yyyy');
  const MM = valueDate(date, 'MM');
  const shift = shiftObjectCheck();
  const staff_obj = staffObject_();
  const staffs = Object.keys(staff_obj);

  const as = mainData_('as');
  const as_sheet = as.getSheetByName(valueDate(date, 'yyyyMM'));
  const as_data = as_sheet.getDataRange().getValues();
  let ind = 0;
  const as_label = as_data.filter((values, index) => {
    if (values.includes('日程')) { ind = index + 2; };
    return values.includes('日程');
  }).flat();
  const as_days = as_data.flatMap(values => values[as_label.indexOf('日程')])
    .map(value => valueDate(value, 'dd'));
  const trim_days = as_days.filter((value, index, array) => String(value).match(/[\d]/) && array.indexOf(value) == index);

  trim_days.forEach(dd => {
    staffs.forEach(staff => {
      let count = 1;
      const work = shift[MM][dd][staff];
      if (work['judge'] == '出勤') {
        const info = {};
        for (let i = as_days.indexOf(dd); i <= as_days.lastIndexOf(dd); i++) {
          if (as_data[i].includes(staff)) {
            work['venue'] = as_data[i][as_label.indexOf('会場\n名称')];
            const time = {};
            const start = new Date(as_data[i][as_label.indexOf('開始')]);
            const end = new Date(as_data[i][as_label.indexOf('終了')]);
            const flag_check = (as_data[i][as_label.indexOf('講師')] == 'エムジー');
            if (flag_check) {
              const main_check = (as_data[i][as_label.indexOf('メイン\n講師')] == staff);
              if (main_check) {
                work['flag'] = 'メイン';
              } else {
                work['flag'] = 'サポート';
              }
              time['meeting'] = shiftMainStart(start);

            } else {
              work['flag'] = 'SB同行';
              time['meeting'] = shiftSupStart(start);
            }
            time['start'] = valueDate(start, 'H:mm');
            time['end'] = valueDate(end, 'H:mm');
            time['leave'] = shiftEnd(end);
            info[String(count).padStart(2, '0')] = time;
            work['info'] = info;
            ++count;
          }
        }
      }
    })
  });

  const database = mainData_('db').getSheetByName('DB');

  let set_obj = JSON.stringify(shift);
  const set_strings = new Array(Math.ceil(set_obj.length / 50000));
  if (set_strings.length > 1) {
    for (let l = 0; l < set_strings.length; l++) {
      set_strings.splice(l, 1, set_obj.slice(l * 50000, (l + 1) * 50000));
    }
  } else {
    set_strings[set_strings.length - 1] = set_obj;
  }
  database.getRange(database.getLastRow() + 1, 1, 1, set_strings.length + 2)
    .setValues([[yyyy, MM].concat(set_strings)]);
  return;
}

const obj_set = (obj) => {
  const date = new Date();
  const yyyy = valueDate(date, 'yyyy');
  const MM = valueDate(date, 'MM');
  const database = mainData_('db').getSheetByName('DB');
  let set_obj = JSON.stringify(obj);
  const set_strings = new Array(Math.ceil(set_obj.length / 50000));

  if (set_strings.length > 1) {
    for (let l = 0; l < set_strings.length; l++) {
      set_strings.splice(l, 1, set_obj.slice(l * 50000, (l + 1) * 50000));
    }
  } else {
    set_strings[set_strings.length - 1] = set_obj;
  }
  database.getRange(database.getLastRow() + 1, 1, 1, set_strings.length + 2)
    .setValues([[yyyy, MM].concat(set_strings)]);
}
const castingTest = () => {
  const vc = testData_('vc');
  const sh = testData_('sh');


}
const testEcho = () => {
  console.log('consoleテスト成功!!');
  Logger.log('Loggerテスト成功!!');
  Browser.msgBox('Browser.msgBoxテスト成功!!');
}
const propertySet = () => {
  const scripts = PropertiesService.getScriptProperties();
}

const propertieCheck = () => {
  const prop = PropertiesService.getScriptProperties();
  const keys = prop.getKeys();
}

const propertieDeliete = () => {
  const prop = PropertiesService.getScriptProperties();
  const keys = prop.getKeys();
}
const judge = (value) => {
  switch (value) {
    case '希': return '希望休';
    case '休': return '公休';
    case '有': return '有休';
    case 'リ': return 'リフレ';
    case 'M': return 'ASB';
    case '研': return '研修';
    default: return '出勤';
  }
}
  // {
    //   yyyy: {
    //     MM: {
    //       dd: {
    //         [venue='会場/n名称']: {
    //           number:'会場番号',
    //           area:'都道府県',
    //           TEL1:'主催者TEL',
    //           TEL2:'会場TEL',
    //           manager:'会場担当者名'
    //           [serial='開催No.']: {
    //             start: 'H:mm',
    //             finish: 'H:mm',
    //             meeting:'H:mm',
    //             leave:'H:mm',
    //             corse:'はじめての～',
    //             hold:['開催','中止',''],
    //             limit:'定員',
    //             nop_plan:'参加予定',
    //             flag:'講師',
    //             update:Date,
    //             carry:'必要キャリー数',
    //             assign:'アサイン数',
    //             store:'誘導先店舗',
    //             SAD:'SAD在籍状況',
    //             support:'SADサポート有なら名前',
    //             member:{
    //               main:'スタッフ名'※SB同行案件ならnull,
    //               support:['スタッフ名','スタッフ名','スタッフ名','スタッフ名','スタッフ名']
    //             }
    //           },
    //           [serial='開催No.']: {
    //             start: 'H:mm',
    //             finish: 'H:mm',
    //             meeting:'H:mm',
    //             leave:'H:mm',
    //             corse:'はじめての～',
    //             hold:['開催','中止',''],
    //             limit:'定員',
    //             nop_plan:'参加予定',
    //             flag:'講師',
    //             update:Date,
    //             carry:'必要キャリー数',
    //             assign:'アサイン数',
    //             store:'誘導先店舗',
    //             SAD:'SAD在籍状況',
    //             support:'SADサポート有なら名前',
    //             member:{
    //               main:'スタッフ名'※SB同行案件ならnull,
    //               support:['スタッフ名','スタッフ名','スタッフ名','スタッフ名','スタッフ名']
    //             }
    //           },
    //           [serial='開催No.']: {
    //             start: 'H:mm',
    //             finish: 'H:mm',
    //             meeting:'H:mm',
    //             leave:'H:mm',
    //             corse:'はじめての～',
    //             hold:['開催','中止',''],
    //             limit:'定員',
    //             nop_plan:'参加予定',
    //             flag:'講師',
    //             update:Date,
    //             carry:'必要キャリー数',
    //             assign:'アサイン数',
    //             store:'誘導先店舗',
    //             SAD:'SAD在籍状況',
    //             support:'SADサポート有なら名前',
    //             member:{
    //               main:'スタッフ名'※SB同行案件ならnull,
    //               support:['スタッフ名','スタッフ名','スタッフ名','スタッフ名','スタッフ名']
    //             }
    //           },
    //         }
    //       }
    //     }
    //   }
    // }
