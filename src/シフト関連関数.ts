const shiftObjectSet = (date = new Date()) => {
  const yyyy = valueDate(date, 'yyyy');
  const MM = valueDate(date, 'MM');
  const shift = shiftObjectCheck();
  const staff_obj = staffObject_();
  const staffs = Object.keys(staff_obj);
  const as = mainData_('as');
  const as_sheet = as.getSheetByName(valueDate(date, 'yyyyMM'));
  const as_data = as_sheet.getDataRange().getValues();
  const as_label = as_data.filter((values, index) => values.includes('日程')).flat();
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
            work['set_value'] = as_data[i][as_label.indexOf('通し番号')];
            const time = {};
            const start = new Date(as_data[i][as_label.indexOf('開始')]);
            const end = new Date(as_data[i][as_label.indexOf('終了')]);
            const flag_check = (as_data[i][as_label.indexOf('講師')] == 'エムジー');
            if (flag_check) {
              const main_check = (as_data[i][as_label.indexOf('メイン\n講師')] == staff);
              if (main_check) { work['flag'] = 'メイン'; }
              else { work['flag'] = 'サポート'; }
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
};
const shiftMainStart = (time) => {
  time = new Date(time);
  time.setMinutes(time.getMinutes() - 90);
  return valueDate(time, 'H:mm');
}
const shiftSupStart = (time) => {
  time = new Date(time);
  time.setMinutes(time.getMinutes() - 60);
  return valueDate(time, 'H:mm');
}
const shiftEnd = (time) => {
  time = new Date(time);
  time.setMinutes(time.getMinutes() + 60);
  return valueDate(time, 'H:mm');
}
const timesousa = (time, minutes) => {
  const type = Object.prototype.toString.call(time) == '[object Date]';
  if (type) {
    time = new Date(time);
    return time.setMinutes(time.getMinutes() + minutes);
  } else {
    try { new Date(time) }
    catch {
      if (String(time).match(/^[\d].:[\d].$/)) {
        const now = new Date();
        now.setHours(time.match(/^[\d]./), time.match(/[\d].$/));
        now.setMinutes(now.getMinutes() + minutes);
        return now;
      }
    }
    const now = new Date(time);
    now.setMinutes(now.getMinutes() + minutes);
    return now;
  }
}


// 翌月希望休申請フォームをオブジェクト化
// 希望休や有休、リフレの取得と出勤可否フラグの設定。
const shiftObjectCreate = (date = new Date()) => {
  if (!date) { date = new Date() }
  const yyyy = String(date.getFullYear());
  const MM = String(date.getMonth() + 1).padStart(2, '0');
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
  const month_obj = {};
  for (let d = 1; d <= check_date; d++) {
    const day_obj = {};
    for (let staff of staffs) {
      const staff_obj = {};
      const index = sheet_staffs.indexOf(staff);
      const hopes = all_hopes[index];
      const paids = all_paids[index];
      const refreshs = all_refreshs[index];
      if (hopes.includes(d) || paids.includes(d) || refreshs.includes(d)) {
        staff_obj['FLAG'] = false;
        switch (true) {
          case hopes.includes(d):
            staff_obj['SET'] = '希';
            break;
          case paids.includes(d):
            staff_obj['SET'] = '有';
            break;
          case refreshs.includes(d):
            staff_obj['SET'] = 'リ';
            break;
        }
      } else {
        staff_obj['FLAG'] = true;
        staff_obj['SET'] = '備';
      }
      day_obj[staff] = staff_obj;
    }
    month_obj[String(d).padStart(2, '0')] = day_obj;
  }
  const database = mainData_('db').getSheetByName('シフト');
  let set_obj = JSON.stringify(month_obj);
  const set_strings = new Array(Math.ceil(set_obj.length / 50000));
  if (set_strings.length > 1) {
    for (let l = 0; l < set_strings.length; l++) {
      set_strings.splice(l, 1, set_obj.slice(l * 50000, (l + 1) * 50000));
    }
  } else {
    set_strings[set_strings.length - 1] = set_obj;
  }
  const db_data = database.getDataRange().getValues();
  const last_row = db_data.length + 1;
  const db_label = db_data.splice(0, 1).flat();
  database.getRange(last_row, 2, 1, set_strings.length + 3)
    .setValues([[new Date(), yyyy, MM].concat(set_strings)]);
  database.getRange(last_row, 1).insertCheckboxes().check();
  const check_row = [];
  for (let [index, data] of db_data.entries()) {
    const db_yyyy = String(data[db_label.indexOf('yyyy')]);
    const db_MM = String(data[db_label.indexOf('M')]).padStart(2, '0');
    if (db_yyyy == yyyy && db_MM == MM) { check_row.push(`A${(index + 2)}`) }
  }
  if (check_row.length > 0) {
    try { database.getRangeList(check_row).uncheck(); }
    catch { database.getRange(check_row[0]).uncheck(); }
  }
  return;
};
// 最新のシフトオブジェクトを返す。
const shiftObjectCheck = (date = new Date()) => {
  const yyyy = date.getFullYear();
  const M = date.getMonth() + 1;
  const label = ['LOCK', 'TimeStamp', 'yyyy', 'M', 'Object'];
  const database = mainData_('db').getSheetByName('シフト');
  const data = database.getDataRange().getValues()
    .filter(values => values[0] == true && values[2] == yyyy && values[3] == M)
    .map(values => values.filter((value, index) => index >= label.indexOf('Object')));
  const to_string = String(data);
  return JSON.parse(to_string);
};
// 最新のシフトオブジェクトに対して、['MTG','病欠','当欠','研修']の情報を追加。
const shiftObjectUpdate = (obj, date = new Date()) => {
  const yyyy = date.getFullYear();
  const M = date.getMonth() + 1;
  const database = mainData_('db').getSheetByName('シフト');
  const db_data = database.getDataRange().getValues();
  const last_row = db_data.length + 1;
  const db_label = db_data.splice(0, 1).flat();
  let set_obj = JSON.stringify(obj);
  const set_strings = new Array(Math.ceil(set_obj.length / 50000));
  if (set_strings.length > 1) {
    for (let l = 0; l < set_strings.length; l++) {
      set_strings.splice(l, 1, set_obj.slice(l * 50000, (l + 1) * 50000));
    }
  } else { set_strings[set_strings.length - 1] = set_obj; }
  const check_row = [];
  for (let [index, data] of db_data.entries()) {
    const db_yyyy = data[db_label.indexOf('yyyy')];
    const db_M = data[db_label.indexOf('M')];
    if (db_yyyy == yyyy && db_M == M) { check_row.push(`A${(index + 2)}`) }
  }
  if (check_row.length > 0) {
    try { database.getRangeList(check_row).uncheck(); }
    catch { database.getRange(check_row[0]).uncheck(); }
  }
  database.getRange(last_row, 2, 1, set_strings.length + 3)
    .setValues([[new Date(), yyyy, M].concat(set_strings)]);
  database.getRange(last_row, 1).insertCheckboxes().check();
  return;
};