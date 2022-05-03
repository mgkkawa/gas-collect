class AddShift {
  constructor() {
    const sheet = mainData_('nh').getSheetByName('追加フォーム');
    const data = sheet.getDataRange().getValues();
    const label = data.splice(0, 1).flat();
    const target = data.flatMap(values => values[label.indexOf('対象月')])
      .filter((value, index, array) => array.indexOf(value) == index);
    target.forEach(month => {
      const target_data = data.filter(values => values[label.indexOf('対象月')] == month);
      target_data.forEach(values => {
        const obj_ = {};
        const staff = values[label.indexOf('対象者')];
        const days = values.filter((value, index) => index > 3).map(day => {
          day = new Date(day);
          return day.getDate();
        });
        obj_['flag'] = values[label.indexOf('追加or削除')];
        obj_['target'] = month;
        obj_['section'] = values[label.indexOf('区分')];
        obj_['days'] = days;
        this[staff] = obj_;
      });
    });
  }
  addDay() {
    const nh = mainData_('nh');
    const sheet = nh.getSheetByName('現在シフト');
    const data = sheet.getDataRange().getValues();
    const label = data[0];
    const staffs = Object.keys(this);
    for (let staff of staffs) {
      const staff_ = this[staff];
      for (let values of data) {
        if (values[label.indexOf('スタッフ名')] == staff &&
          values[label.indexOf('対象月')] == staff_.target) {
          if (staff_.flag == '追加') {
            values[label.indexOf(staff_.section)] =
              `${values[label.indexOf(staff_.section)]}, ${staff_.days.join(' ,')}`.replace(/^ , /, '');
          }
          else {
            staff_.days.forEach(d => {
              const reg = new RegExp(`^${d}, | ${d}$|^${d}$| ${d},`);
              values[label.indexOf(staff_.section)] =
                String(values[label.indexOf(staff_.section)]).replace(reg, '');
            });
          }
        }
      }
    }
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    const del = nh.getSheetByName('追加フォーム');
    del.getRange(2, 1, del.getLastRow() - 1, del.getLastColumn()).clearContent();
  }
}
class OriginShift {
  constructor(month) {
    const sheet = mainData_('nh').getSheetByName('現在シフト');
    const data = sheet.getDataRange().getValues().filter(values => values[1] == month);
    const staffs = Object.keys(staffObject_());
    staffs.forEach(staff => {
      data.forEach(values => {
        if (values[0] == staff) {
          this[staff] = new OriginStaffHope(values);
        }
      });
    });
  }
}
class OriginStaffHope {
  hope: any;
  paid: any;
  refresh: any;
  mtg: any;
  training: any;
  behind: any;
  absence: any;
  sick: any;
  constructor(values) {
    this.hope = JSON.parse(`[${values[2]}]`);
    this.paid = JSON.parse(`[${values[3]}]`);
    this.refresh = JSON.parse(`[${values[4]}]`);
    this.mtg = JSON.parse(`[${values[5]}]`);
    this.training = JSON.parse(`[${values[6]}]`);
    this.behind = JSON.parse(`[${values[7]}]`);
    this.absence = JSON.parse(`[${values[8]}]`);
    this.sick = JSON.parse(`[${values[9]}]`);
  }
}
class StaffWorkRecord {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  constructor(obj, date = new Date()) {
    const nh = mainData_('nh');
    const sheet = nh.getSheetByName('現在シフト');
    const month = date.getMonth() + 1;
    date.setMonth(date.getMonth() + 1, 0);
    this.sheet = sheet;
    let data = sheet.getDataRange().getValues();
    const staffs = Object.keys(staffObject_());
    staffs.forEach(staff => {
      let staff_array = data.filter(values => values[0] == staff && values[1] == month + '月').flat();
      staff_array.splice(0, 2);
      staff_array.unshift(date);
      staff_array = staff_array.map((value, index) => {
        if (index > 0) {
          return JSON.parse(`[${value}]`);
        }
        return value;
      });
      this[staff] = new dayObjectCreate(staff_array);
      this[staff].holiday = 0;
      for (let day in this[staff]) {
        const staff_ = this[staff][day];
        const flag = staff_.flag;
        if (!flag) {
          switch (staff_.set_num) {
            case '休':
            case '希':
              this[staff].holiday++;
              break;
            case '有':
              if (!this[staff].paid) {
                this[staff].paid = 0;
              }
              this[staff].paid++;
              break;
            case 'リ':
              if (!this[staff].paid) {
                this[staff].refresh = 0;
              }
              this[staff].refresh++;
              break;
          }
        }
        else {
          const keys = obj.returnkeys();
          let count = 1;
          const objs = {};
          for (let key of keys) {
            const obj_ = obj[key];
            if (obj_.date != day) {
              continue;
            }
            const members = obj_.member.support.concat(obj_.member.main);
            if (members.some(name => name == staff)) {
              objs[String(count)] = obj_;
              ++count;
            }
            if (Object.keys(objs).length > 0) {
              staff_.addInfo(objs);
            }
          }
        }
      }
    });
  }
  getShift(staff, day) {
    return this[staff][day];
  }
}
class dayObjectCreate {
  constructor(data) {
    const date = new Date(data.splice(0, 1));
    const month = date.getMonth() + 1;
    date.setMonth(month, 0);
    const end = date.getDate();
    for (let d = 1; d <= end; d++) {
      const day = `${String(month).padStart(2, '0')}/${String(d).padStart(2, '0')}`;
      this[day] = new Day(d, data);
    }
  }
}
class Day {
  flag: boolean;
  set_num: string;
  meeting: string;
  leave: string;
  venue: any;
  constructor(d, data) {
    switch (true) {
      case data[0].some(day => day == d):
        this.flag = false;
        this.set_num = '希';
        return this;
      case data[1].some(day => day == d):
        this.flag = false;
        this.set_num = '有';
        return this;
      case data[2].some(day => day == d):
        this.flag = false;
        this.set_num = 'リ';
        return this;
      case data[3].some(day => day == d):
        this.flag = false;
        this.set_num = 'M';
        return this;
      case data[4].some(day => day == d):
        this.flag = true;
        this.set_num = '研';
        this.meeting = '10:00';
        this.leave = '18:00';
        return this;
      case data[5].some(day => day == d):
        this.flag = false;
        this.set_num = '遅刻';
        return this;
      case data[6].some(day => day == d):
        this.flag = false;
        this.set_num = '当欠';
        return this;
      case data[7].some(day => day == d):
        this.flag = false;
        this.set_num = '病欠';
        return this;
      default:
        this.flag = true;
        this.set_num = '備';
        return this;
    }
  }
  addInfo(obj) {
    const keys = Object.keys(obj);
    if (obj.length == 1) {
      obj = obj[keys[0]];
      this.venue = obj.venue;
      this.set_num = obj.set;
      this['1'] = new DayInfo(obj);
      const flag = this['1'].mg_flag;
      const meeting = new Date(this['1'].start);
      if (flag) {
        meeting.setMinutes(meeting.getMinutes() - 90);
      }
      else {
        meeting.setMinutes(meeting.getMinutes() - 60);
      }
      this.meeting = dateString(meeting, 'H:mm');
      const leave = new Date(this['1'].finish);
      leave.setMinutes(leave.getMinutes() + 60);
      this.leave = dateString(leave, 'H:mm');
      return this;
    }
    const set_obj = obj[keys[0]];
    this.venue = set_obj.venue;
    this.set_num = set_obj.set;
    const flag = set_obj.mg_flag;
    let times = [];
    for (let key in obj) {
      this[String(key)] = new DayInfo(obj[key]);
      times.push([new Date(this[String(key)].start), new Date(this[String(key)].finish)]);
    }
    times = times.flat().sort((a, b) => a.getTime() - b.getTime());
    const meeting = times[0];
    const leave = times[times.length - 1];
    if (flag) {
      meeting.setMinutes(meeting.getMinutes() - 90);
    }
    else {
      meeting.setMinutes(meeting.getMinutes() - 60);
    }
    leave.setMinutes(leave.getMinutes() + 60);
    this.meeting = dateString(meeting, 'H:mm');
    this.leave = dateString(leave, 'H:mm');
    return this;
  }
  addFlyer() {
  }
  addSuiteCase() { }
}
class DayInfo {
  serial: any;
  corse: any;
  start: any;
  finish: any;
  limit: number;
  nop: number;
  carry: number;
  mg_flag: any;
  sad: any;
  update: any;
  checkday: any;
  constructor(obj) {
    this.serial = obj.serial;
    this.corse = obj.corse;
    this.start = obj.start;
    this.finish = obj.finish;
    this.limit = Number(obj.limit);
    this.nop = Number(obj.nop);
    let limit = this.limit;
    if (this.limit < this.nop) {
      limit = this.nop;
    }
    switch (true) {
      case limit < 11:
        this.carry = 1;
        break;
      case limit <= 21:
        this.carry = 2;
        break;
      default:
        this.carry = 3;
        break;
    }
    this.mg_flag = obj.mg_flag;
    this.sad = obj.sad;
    this.update = obj.update;
    this.checkday = obj.checkday;
  }
}
class AssignObject {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  label: any[];
  maincol: any;
  supcol: any;
  constructor(date = new Date()) {
    const as = mainData_('as');
    const sheet = as.getSheetByName(dateString(date, 'yyyyMM'));
    this.sheet = sheet;
    const data = sheet.getDataRange().getValues();
    let label_index;
    this.label = data.filter((values, index) => {
      if (values.includes('日程')) {
        label_index = index;
        return true;
      }
      ;
    }).flat();
    data.forEach((values, index) => {
      if (index > label_index && values[this.label.indexOf('日程')] != ''
        && values[this.label.indexOf('会場\n名称')] != '') {
        this[String(index + 1)] = new Venue(values, this.label);
      }
    });
    this.maincol = this.label.indexOf('メイン\n講師');
    this.supcol = this.label.indexOf('サポート講師');
  }
  returnkeys() {
    return Object.keys(this).filter(key => key != 'sheet' && key != 'label' && key != 'maincol' && key != 'supcol');
  }
  rowNum(date, venue, start) {
    for (let row in this) {
      if (this[row]['date'] == date && this[row]['venue'] == venue && this[row]['start'] == start) {
        return [Number(row), this[row]['set']];
      }
    }
  }
  getMainCell(date, venue, start) {
    for (let row in this) {
      if (this[row]['date'] == date && this[row]['venue'] == venue && this[row]['start'] == start) {
        return `${NumToA1(this.maincol + 1)}${row}`;
      }
    }
  }
  getSupCell(date, venue, start, length) {
    for (let row in this) {
      if (this[row]['date'] == date && this[row]['venue'] == venue && this[row]['start'] == start) {
        if (length == 0) {
          return `${NumToA1(this.supcol + 1)}${row}`;
        }
        else {
          return `${NumToA1(this.supcol + 1)}${row}:${NumToA1(this.supcol + length)}${row}`;
        }
      }
    }
  }
  getCells(date, venue, start, col, length) {
    for (let row in this) {
      if (this[row]['date'] == date && this[row]['venue'] == venue && this[row]['start'] == start) {
        return `${NumToA1(this.label.indexOf(col) + 1)}${row}:${NumToA1(this.label.indexOf(col) + length)}${row}`;
      }
    }
  }
  setValue(range, value) {
    this.sheet.getRange(range).setValue(value);
  }
  setValues(range, values) {
    this.sheet.getRange(range).setValues(values);
  }
  getValues(staff) {
    // const keys = 
  }
}
class Venue {
  serial: string;
  date: any;
  vennum: string;
  venue: any;
  area: any;
  hold: any;
  address: any;
  main_tel: string;
  sub_tel: string;
  corse: any;
  start: Date;
  finish: Date;
  limit: number;
  carry: number;
  assign: number;
  caution: any[];
  store: any;
  sad: any;
  update: any;
  maneger: any;
  nop: any;
  checkday: any;
  mg_flag: boolean;
  member: {};
  set: any;
  constructor(arg, label) {
    this.serial = String(arg[label.indexOf('開催No.')]);
    this.date = dateString(arg[label.indexOf('日程')]);
    this.vennum = String(arg[label.indexOf('会場\n番号')]);
    this.venue = arg[label.indexOf('会場\n名称')];
    this.area = arg[label.indexOf('地域')];
    this.hold = arg[label.indexOf('開催\n可否')];
    this.address = arg[label.indexOf('会場\n住所')];
    this.main_tel = String(arg[label.indexOf('主催者TEL')]);
    if (arg[label.indexOf('会場TEL')] != '') {
      this.sub_tel = String(arg[label.indexOf('会場TEL')]);
    }
    ;
    this.corse = arg[label.indexOf('コース')];
    this.start = new Date(arg[label.indexOf('開始')]);
    this.finish = new Date(arg[label.indexOf('終了')]);
    this.limit = Number(arg[label.indexOf('定員\n(半角)')]);
    this.carry = Number(arg[label.indexOf('必要キャリー数\n(半角)')]);
    this.assign = Number(arg[label.indexOf('アサイン数\n(半角)')]);
    this.caution = [arg[label.indexOf('会場運用上\n注意点')], arg[label.indexOf('カリキュラム\n補足')], arg[label.indexOf('連絡事項')]].filter(Boolean);
    this.store = arg[label.indexOf('誘導先店舗')];
    if (arg[label.indexOf('SAD在籍状況')] == '在籍') {
      this.sad = { flag: true };
    }
    else {
      this.sad = { flag: false };
    }
    ;
    this.sad.support = arg[label.indexOf('SADサポート有の場合\n(名前+店舗名)')];
    this.update = arg[label.indexOf('更新日')];
    this.maneger = arg[label.indexOf('会場\n担当者名')];
    this.nop = arg[label.indexOf('参加予定人数')];
    this.checkday = arg[label.indexOf('確認日')];
    if (arg[label.indexOf('講師')] == 'エムジー') {
      this.mg_flag = true;
    }
    else {
      this.mg_flag = false;
    }
    ;
    this.member = {};
    if (this.mg_flag == true) {
      this.member = { main: arg[label.indexOf('メイン\n講師')] };
    }
    else {
      this.member = {};
    }
    ;
    this.member.support = arg.filter((value, index) => index >= label.indexOf('サポート講師') && index < label.indexOf('通し番号') && value != '');
    this.set = arg[label.indexOf('通し番号')];
  }
}
