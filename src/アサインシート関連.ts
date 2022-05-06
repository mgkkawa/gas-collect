const diffCheck_ = () => {
  const date = new Date()
  const sheetname = dateString(date, 'yyyyMM')
  const vc = mainData_('vc')
  const as = mainData_('as')
  const put = new DiffSheet(vc.getSheetByName(sheetname))
  const origin = new DiffSheet(as.getSheetByName(sheetname), put.label)
  put.diffCheck(origin)

  date.setMonth(date.getMonth() + 1)
  const sheet = as.getSheetByName(dateString(date, 'yyyyMM'))
  new Promise(() => {
    return new DiffSheet(sheet, put.label)
  }).then((obj) => {
    const next_put = addMonthSheet_(vc, date)
    next_put.diffCheck(obj)
  }).catch(() => Logger.log('翌月分のアサインシートないよ'))
}
const assignObject = () => {
  const date = new Date();
  const as = mainData_('as');
  const as_sheet = as.getSheetByName(dateString(date, 'yyyyMM'));
  const as_data = as_sheet.getDataRange().getValues();
  let ind = 0;
  const as_label = as_data.filter((values, index) => {
    if (values.includes('日程')) {
      ind = index + 2;
      return true;
    }
  }).flat();
  const obj = {};
};
const assaignsheet = () => {
  const as = mainData_('as');
  const asag = as.getSheetByName('集約');
  const asag_label = asag.getRange(2, 1, 1, asag.getLastColumn()).getValues().flat();
  let ind = 0;
  const sheetname = '202109';
  const sheet = as.getSheetByName(sheetname);
  const sheet_data = sheet.getDataRange().getValues();
  const sheet_label = sheet_data.filter((values, index) => {
    if (values.includes('日程')) {
      ind = index + 2;
    }
    return values.includes('日程');
  }).flat();
  const sheet_indexs = fill(asag_label.map(key => sheet_label.indexOf(key) + 1));
  let Col_list = '';
  sheet_indexs.forEach((key, index) => {
    if (index == 0) {
      Col_list = `${NumToA1(key)}, `;
    }
    else if (index == sheet_indexs.length - 1) {
      Col_list += `${NumToA1(key)}`;
    }
    else {
      Col_list += `${NumToA1(key)}, `;
    }
  });
  const dayCol = NumToA1(sheet_label.indexOf('日程') + 1);
  const venCol = NumToA1(sheet_label.indexOf('会場\n名称') + 1);
  const lastCol = NumToA1(sheet_label.length);
  const initial = `'${sheetname}'!A${ind}:${lastCol}`;
  const func = `QUERY(${initial},"select ${Col_list} where ${dayCol} <>'' and ${venCol} <>''")`;
  const base_func = asag.getRange('F1').getValue();
  Logger.log(base_func);
  const write_func = base_func.replace(/\)}/, `\)${func}}`);
  asag.getRange('F1').setValue(write_func);
  asag.getRange('A3').setValue('=' + write_func);
  Logger.log(write_func);
};
const maxfill = (ary) => {
  return ary.map(value => {
    if (value != 0) {
      return value;
    }
    else {
      let i = 0;
      while (i <= ary.reduce((a, b) => Math.max(a, b))) {
        if (ary.indexOf(i) == -1) {
          return i;
        }
        ++i;
      }
    }
  });
};
const fill = (ary) => {
  const maxnum = ary.reduce((a, b) => Math.max(a, b));
  ary.forEach((num, index) => {
    if (num == 0) {
      let i = 0;
      while (i <= maxnum) {
        if (ary.indexOf(i) == -1) {
          ary.splice(index, 1, i);
        }
        ++i;
      }
    }
    else {
      return num;
    }
  });
  return ary;
};

class DiffSheet {
  label: any[];
  sheet: GoogleAppsScript.Spreadsheet.Sheet
  constructor(sheet, keys = undefined) {
    let data = sheet.getDataRange().getValues()
    let ind
    let label = data.filter((values, index) => {
      if (ind) { return false }
      if (values.includes('日程')) {
        ind = index
        return true
      }
    }).flat()
    if (Boolean(keys)) {
      data = data.map(values => keys.map((key, index) => {
        if (label.indexOf(key) == -1) {
          return ''
        }
        return values[label.indexOf(key)]
      }))
      label = keys
    }
    data.forEach((values, index) => {
      if (index > ind && (values[label.indexOf('日程')] != '' || values[label.indexOf('会場\n名称')] != '')) {
        this[index + 1] = trimValues_(values, label)
      }
    })
    this.label = label
    this.sheet = sheet
  }
  labelCreate() {
    return Object.keys(this).filter(key => key != 'label' && key != 'sheet')
  }
  reMap() {
    return this.labelCreate().map(key => this[key])
  }
  assignToPut(obj) {
    const sheet = obj.sheet
    let last_row = sheet.getLastRow() - 1
    if (last_row == 0) { last_row = 1 }
    const arr = []
    for (let key in this) {
      if (key == 'label' || key == 'sheet') { continue }
      arr.push(this[key])
    }
    sheet.getRange(2, 1, last_row, sheet.getLastColumn()).clearContent()
    sheet.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  diffCheck(obj) {
    const put = this.reMap()
    const out = obj.reMap()
    const serials_put = put.map(values => values[0])
    const serials_obj = out.map(values => values[0])
    if (JSON.stringify(put) == JSON.stringify(out)) {
      Logger.log('差分なし！')
      return
    }
    const map = serials_obj.map((serial, index) => {
      if (serial == serials_put[index]) {
        return diffMap(put[index], out[index])
      } else if (serials_put.includes(serial)) {
        return diffMap(put[serials_put.indexOf(serial)], out[index])
      } else {
        return out[index]//.concat(array)
      }
    })
    let last_row = this.sheet.getLastRow() - 1
    if (last_row == 0) { last_row = 1 }
    this.sheet.getRange(2, 1, last_row, this.sheet.getLastColumn()).clearContent()
    this.sheet.getRange(2, 1, map.length, map[0].length).setValues(map)
  }
}
const trimTel_ = (str) => {
  if (!Boolean(str)) { return str }
  const tel = str.replace(/[^\d]/g, '').match(/0[5789]0[\d]{8}|0[\d]{9}/)
  return String(tel)
}
const trimValues_ = (values, label) => {
  const start = values[label.indexOf('開始')]
  const finish = values[label.indexOf('終了')]
  const address = values[label.indexOf('会場\n住所')]
  const flag = values[label.indexOf('講師')] == 'エムジー'
  return values.map((value, index) => {
    switch (index) {
      case label.indexOf('参加予定人数'):
        if (!isNaN_(value) && value != '') {
          return Number(value.match(/[1-9]?[\d]/))
        }
        return value
      case label.indexOf('開始'):
      case label.indexOf('終了'): return new Times(value).Hmm()
      case label.indexOf('開催No.'):
      case label.indexOf('通し番号'): return String(value)
      case label.indexOf('LOG住所'): return new AddressWork(address).onlyAddress()
      case label.indexOf('LOG建物'): return new AddressWork(address).building()
      case label.indexOf('LOG主催者TEL'): return trimTel_(values[label.indexOf('主催者TEL')])
      case label.indexOf('LOG会場TEL'): return trimTel_(values[label.indexOf('会場TEL')])
      case label.indexOf('集合'): return new Times(start).meetingTime(flag)
      case label.indexOf('解散'): return new Times(finish).endTime()
    }
    return dateString(value)
  })
}