class Staffshift {
  constructor(date = new Date()) {
    const spread = mainData_('sh');
    const sheet = spread.getSheetByName(dateString(date, 'yyyy.MM'));
    let data = sheet.getDataRange().getValues();
    const label = data.splice(0, 1).flat();
    data = data.filter(values => values[label.indexOf('氏名')] != '')
      .map(values => values.filter((value, index) => index >= label.indexOf('氏名') && index != label.indexOf('申請するのは何月ですか？')))
      .map(values => values.map((value, index) => datereplace(value, index)));
    data.forEach(values => {
      const staff = values[0][0];
      this[staff] = new Monthryshift(values, date);
    });
  }
}
class Info {
  venue: any;
  area: any;
  vennum: any;
  update: any;
  meeting: any;
  leave: any;
  constructor(obj, day, staff) {
    // const label = obj.label
    const keys = Object.keys(obj).filter(key => key != 'label' && key != 'sheet' && key != 'maincol' && key != 'supcol');
    let count = 1;
    let times = [];
    for (let key of keys) {
      const obj_ = obj[key];
      if (obj_.date != day && (obj_.main != staff || obj.sup.every(member => member != staff))) {
        continue;
      }
      const obj_number = String(count).padStart(2, '0');
      const serial = obj_.serial;
      const start = obj_.start;
      const finish = obj_.finish;
      const corse = obj_.corse;
      const start_date = new Date(new Date().getFullYear(), Number(day.slice(0, 2)), Number(day.slice(3, 2)), Number(start.slice(0, 2)), Number(start.slice(3, 2)));
      const finish_date = new Date(new Date().getFullYear(), Number(day.slice(0, 2)), Number(day.slice(3, 2)), Number(finish.slice(0, 2)), Number(finish.slice(3, 2)));
      this[obj_number].serial = serial;
      this[obj_number].start = start;
      this[obj_number].finish = finish;
      this[obj_number].corse = corse;
      times.push([start_date, finish_date]);
      if (count > 1) {
        break;
      }
      const venue = obj_.venue;
      const area = obj_.area;
      const vennum = obj_.vennum;
      const update = obj_.update;
      this.venue = venue;
      this.area = area;
      this.vennum = vennum;
      this.update = update;
      ++count;
    }
    times = times.flat().sort((a, b) => a.getTime() - b.getTime());
    this.meeting = dateString(times[0]);
    this.leave = dateString(times[times.length - 1]);
  }
}
class Monthryshift {
  constructor(arg, date) {
    date.setMonth(date.getMonth() + 1, 0);
    const end = date.getDate();
    date.setDate(1);
    const month = dateString(date, 'MM/');
    const hopes = arg[1];
    const pays = arg[2];
    const refs = arg[3];
    const mtg = arg[4];
    const training = arg[5];
    const sicks = arg[6];
    const absence = arg[7];
    let day;
    let count = 1;
    while (count <= end) {
      day = `${month}${String(count).padStart(2, '0')}`;
      switch (true) {
        case hopes.some(key => key == count):
          this[day] = new Work(false, '希');
          break;
        case pays.some(key => key == count):
          this[day] = new Work(false, '有');
          break;
        case refs.some(key => key == count):
          this[day] = new Work(false, 'リ');
          break;
        case sicks.some(key => key == count):
          this[day] = new Work(false, '病欠');
          break;
        case absence.some(key => key == count):
          this[day] = new Work(false, '当欠');
          break;
        case mtg == count:
          this[day] = new Work(false, 'M');
          break;
        case training.some(key => key == count):
          this[day] = new Work(true, '研');
          break;
        default:
          this[day] = new Work(true, '備');
          break;
      }
      ++count;
    }
  }
}
const infoCheck = () => {
};
class Work {
  flag: any;
  number: string;
  constructor(flag, number) {
    this.flag = flag;
    this.number = String(number);
  }
  ;
}
const worktimecheck = (times) => {
  const date = new Date();
  times = times.map(time => {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), Number(time.slice(0, 2)), Number(time.slice(3, 2)));
  }).sort((a, b) => a.getTime() - b.getTime());
  const meeting = new Date(times[0]);
  meeting.setMinutes(meeting.getMinutes() - 90);
  const leave = new Date(times[times.length - 1]);
  leave.setMinutes(leave.getMinutes() + 60);
  return [dateString(meeting), dateString(leave)];
};
class Worktime {
  meeting: any;
  leave: any;
  constructor(times) {
    const date = new Date();
    times = times.map(time => {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate(), Number(time.slice(0, 2)), Number(time.slice(3, 2)));
    }).sort((a, b) => a.getTime() - b.getTime());
    const meeting = new Date(times[0]);
    meeting.setMinutes(meeting.getMinutes() - 90);
    this.meeting = dateString(meeting);
    const leave = new Date(times[times.length - 1]);
    leave.setMinutes(leave.getMinutes() + 60);
    this.leave = dateString(leave);
  }
}
class Venuecall {
  label: string[];
  constructor(arg, everys = undefined, somes = undefined) {
    //argにはシート全体の二次元配列。
    //特定indexのデータが空白なら取得しない。
    // everysは[ラベル名,...]内、全てが空白でない事をチェック。
    // somesは[ラベル名,...]内、空白が含まれていても何かデータが入っているかをチェック。
    this.label = labelCreate_(arg[0]);
    let echeck = true;
    let scheck = true;
    arg.forEach((values, index) => {
      if (!(everys == undefined)) {
        echeck = everys.map(key => this.label.indexOf(key)).every(col => values[col] != '');
      }
      if (!(somes == undefined)) {
        scheck = somes.map(key => this.label.indexOf(key)).some(col => values[col] != '');
      }
      if (echeck && scheck && index > 0) {
        index += 1;
        this[index] = {
          date: dateString(values[this.label.indexOf('日程')]),
          venue: values[this.label.indexOf('会場\n名称')],
          start: dateString(values[this.label.indexOf('開始')], 'H:mm'),
          serial: String(values[this.label.indexOf('開催No.')]),
          main: values[this.label.indexOf('メイン\n講師')],
          support: values.filter((value, ind) => ind >= this.label.indexOf('サポート講師') && ind <= this.label.indexOf('サポート5') && value != ''),
        };
        if (this.label.includes('スクリーン')) {
          ['人数', '施設担当者（今回）', 'スクリーン（今回）', '入館', '次回引継ぎ'];
          this[index].nop = values[this.label.indexOf('人数')];
          this[index].manager = values[this.label.indexOf('施設担当者（今回）')];
          this[index].screen = values[this.label.indexOf('スクリーン（今回）')];
          this[index].inside = values[this.label.indexOf('入館')];
          this[index].over = values[this.label.indexOf('次回引継ぎ')];
        }
        if (this.label.includes('Check1')) {
          this[index].field = values[this.label.indexOf('Check1')]
          this[index].work = values[this.label.indexOf('Check2')]
          this[index].casting = values[this.label.indexOf('Check3')]
        }
      }
      ;
    });
  }
  ;
  getCell(date, venue, start, col) {
    for (let row in this) {
      if (row == 'label') {
        continue;
      }
      ;
      if (this[row]['date'] == date && this[row]['venue'] == venue && this[row]['start'] == start) {
        return `${NumToA1(this.label.indexOf(col) + 1)}${row}`;
      }
      ;
    }
    ;
  }
  ;
  check() {
    const keys = Object.keys(this).filter(key => key != 'label');
    const to_month = this[keys[0]].date.slice(0, 3);
    return !(keys.every(key => String(this[key].date).includes(to_month)));
  }
  ;
}
class ShiftTable {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  constructor(date = new Date()) {
    const sheet = mainData_('sh').getSheetByName(dateString(date, 'yyyy.MM'));
    this.sheet = sheet;
    const keys = Object.keys(staffObject_());
    date.setDate(1);
    const start = dateString(date);
    date.setMonth(date.getMonth() + 1, 0);
    const last = dateString(date);
    const arg = sheet.getDataRange().getValues().map(values => values.map(value => dateString(value)));
    const days = arg.filter(values => values.includes(start)).flat();
    const staffs = arg.flatMap(values => values[days.indexOf('スタッフ')]);
    days.forEach((day, index) => {
      if (index >= days.indexOf(start) && index <= days.indexOf(last)) {
        this[day] = NumToA1(index + 1);
      }
    });
    staffs.forEach((staff, index) => {
      if (index >= staffs.indexOf(keys[0]) && index <= staffs.indexOf(keys[keys.length - 1])) {
        this[staff] = index + 1;
      }
    });
  }
  getCell(day, staff) { return `${this[day]}${this[staff]}`; }
  setValue(range, value) { this.sheet.getRange(range).setValue(value); }
  listSetValue(rangelist, value) { this.sheet.getRangeList(rangelist).setValue(value); }
}
;
class Assign {
  sheet: any;
  label: any;
  maincol: any;
  supcol: any;
  constructor(sheet) {
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
    this.member['support'] = arg.filter((value, index) => index >= label.indexOf('サポート講師') && index < label.indexOf('通し番号') && value != '');
    this.set = arg[label.indexOf('通し番号')];
  }
}

class AddressWork {
  value: any;
  constructor(str) {
    const zenkaku2hankaku = (str) => {
      return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => {
        return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
      });
    };
    this.value = str.replace(/[^\x01-\x7E\xA1-\xDF]/g, str => zenkaku2hankaku(str)).replace(/[\n\r]/g, '')
      .replace(/(?<=\d)[ーｰ－−-]|(丁目(?=\d)|番地の?(?=\d)|(?<=\d)番(?!([地 　]|$)))/g, '-')
      .replace(/[一二三四五六七八九十〇](?=-)|(?<=-)[一二三四五六七八九十〇]/g, str => kanji2num_(str))
      .replace(/[　]|(?<=\d)[\(（]|(番地|号|番(?!地))(?=[\(（])|(番地|番|号)([ 　]|$)/g, ' ')
      .replace(/!.*! |[\(\)（）]|[\s]{2,}|(?<!(\d|丁目))\s/g, '')
  }
  onlyAddress() {
    const address = this.value.trim().match(/^.*\d(丁目)?(?=(\s|$))/)
    if (Boolean(address)) {
      return address[0]
    }
    return ''
  }
  building() {
    const address = this.value.match(/(?<=\s).*$/)
    if (Boolean(address)) {
      return address[0]
    }
    return ''
  }
}
class TextNumbers {
  value: any;
  constructor(value) {
    value = String(value)
    this.value = value
  }
  onlyPhoneNumber() {
    if (!Boolean(this.value)) { return this.value }
    const tel = this.value.replace(/[^\d]/g, '').match(/0[5789]0[\d]{8}|0[\d]{9}/)
    if (!Boolean(tel)) { return '' }
    return String(tel)
  }
}