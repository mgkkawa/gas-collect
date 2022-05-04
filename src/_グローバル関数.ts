var start_time = new Date();
const triggerset = (t, time) => {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() == t) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(t).timeBased().at(time).create();
};
const properties = (str) => {
  return PropertiesService.getScriptProperties().getProperty(str);
};
const mainData_ = (s) => {
  switch (s) {
    case 'vc': return SpreadsheetApp.openById(properties('main_vencall')); //[移行先]
    case 'as': return SpreadsheetApp.openById(properties('main_assign')); //【ソフトバンク様】共有用アサインシート
    case 'sh': return SpreadsheetApp.openById(properties('main_shift')); //SBスマホ教室シフト ver.2
    case 'nh': return SpreadsheetApp.openById(properties('next_holiday')); //翌月希望休申請フォーム
    case 'ss': return SpreadsheetApp.openById(properties('detail_shift')); //編集用
    case 'mg': return SpreadsheetApp.openById(properties('mgshift')); //MGシフト
    case 'st': return SpreadsheetApp.openById(properties('stay_request')); //宿泊申請用フォーム(回答)
    case 'wr': return SpreadsheetApp.openById(properties('work_record')); //勤務実績表
    case 'cr': return SpreadsheetApp.openById(properties('folder_create')); //フォルダ作成用スプレッド
    case 'ex': return SpreadsheetApp.openById(properties('origin_exform')); //経費申請書
    case 'sc': return SpreadsheetApp.openById(properties('suit_case')); //新スーツケース管理表
    case 'tm': return SpreadsheetApp.openById(properties('temperature')); //検温結果報告フォーム(回答)
    case 'ef': return SpreadsheetApp.openById(properties('calc_request')); //経費申請希望フォーム（回答）
    case 'ec': return SpreadsheetApp.openById(properties('exform_calc')); //経費集計用
    case 'db': return SpreadsheetApp.openById(properties('shift_db')); //ObjectDB
  }
};
const mainForm_ = (s) => {
  switch (s) {
    case 'together': return FormApp.openById('13pB1ZKTIiMrS1FKa5hHGVIa8sgd4nLXYZfRDRwdzrtQ');
  }
};
const testData_ = (s) => {
  switch (s) {
    case 'vc': return SpreadsheetApp.openById(properties('test_vencall')); //[開発用]新会場連絡シート
    case 'sh': return SpreadsheetApp.openById(properties('test_shift')); //[開発用]SBスマホ教室シフト ver.2
    case 'as': return SpreadsheetApp.openById(properties('test_assign')); //[開発用]アサインシート
    //case 'nh': return
  }
};
const setOptions_ = () => {
  return JSON.parse(properties('setOption'));
};
const accountId_ = (name) => {
  const member = JSON.parse(properties('member_obj'));
  return member[name]['line'];
  switch (name) {
    case '大山夏美': return properties('line_oyama');
    case '山崎達也': return properties('line_yamazaki');
    case '富樫一世': return properties('line_togashi');
    case 'room': return properties('dsg_room');
    case 'dsg':
      const room = JSON.parse(properties('member_obj'));
      return room[name];
    case 'domainId': return properties('domainId');
    case 'options': return properties('options_mail');
    default: return properties('line_kawate');
  }
};
const callbackURL_ = () => {
  return properties('callbackURL');
};
const addressCheck_ = () => {
  var sheet = SpreadsheetApp.openById('1m93CFX1uG67bO6c5xbSGoV5Bm0xNbfO0QAkE7nQqO5c').getSheetByName('202204');
  //var sheet = SpreadsheetApp.openById('1aF-KKlYVWMNBO95Gc4B2d70cie7fPApz-G7m0PR2bVQ').getSheetByName('シート2')
  var dat = sheet.getDataRange().getDisplayValues();
  for (var i = 0; i < dat.length; i++) {
    if (dat[i].indexOf('会場\n住所') != -1) {
      var addCol = dat[i].indexOf('会場\n住所');
      var row = i + 2;
      break;
    }
  }
  var addDat = sheet.getRange(row, addCol + 1, sheet.getLastRow() - row).getDisplayValues().flat();
  var reg = /[一二三四五六七八九十〇](?=丁目|番地|号)|番(?=$|[0-9 ])/g;
  addDat = addDat.map(value => zen2han_(value)).map(value => value.replace(reg, s => { return kanji2num_(s); }));
  addDat = addDat.map(value => value.replace(/(?<=[0-9])(丁目|番地|番地の|[番のー－ｰ‐])(?=[0-9])/g, '-'));
  addDat = addDat.map(value => value.replace(/(?<=[0-9])(番地|[番号])(?!地|[0-9])[　| ]?|[　]|\n|\r\n|\r/g, ' '));
  var setAdd = addDat.map(address => [address.replace(/  /g, ' ')]);
  sheet.getRange(row, addCol + 1, setAdd.length).setValues(setAdd);
};
const NumToA1 = (num) => {
  const RADIX = 26;
  const A = 'A'.charCodeAt(0);
  var n = num;
  var s = "";
  while (n >= 1) {
    n--;
    s = String.fromCharCode(A + (n % RADIX)) + s;
    n = Math.floor(n / RADIX);
  }
  return s;
};
const datObject_ = (array) => {
  //受け取った配列を連想配列化して返す。
  var keys = array[0];
  array.shift();
  var obj = array.map(values => {
    var hash = {};
    values.map((value, x) => hash[keys[x]] = value);
    return hash;
  });
  return obj;
};
const objectCut_ = (obj, keys) => {
  //受け取った連想配列を受け取ったキーで取り出して二次元配列として返す。
  return obj.map(array => keys.map(key => array[key]));
};
const convertDate_ = (values, str) => {
  if (!str) {
    str = 'yyyy/MM/dd';
  }
  //date型をstringに変換
  for (var i = 0; i < values.length; i++) {
    var newValues = values[i].map(x => {
      var type = Object.prototype.toString.call(x);
      if (type == "[object Date]") {
        return x = Utilities.formatDate(x, 'JST', str);
      }
      else {
        return x;
      }
    });
    values[i] = newValues;
  }
};
const convertObj_ = (values) => {
  var reg = /^....\/..\/..$/;
  for (var i = 0; i < values.length; i++) {
    var newValues = values[i].map(x => {
      var regmatch = x.match(reg);
      if (regmatch != null) {
        return x = String(x.match(/(?!<\/)..\/..$/));
      }
      else {
        return x;
      }
    });
    values[i] = newValues;
  }
  return values;
};
const month_ = (value) => {
  switch (true) {
    case value >= 5: return value.match(/(?<=\/)[0-9][1-9](?=\/)/);
    default: return value.slice(0, 2);
  }
};
const dateString = (value, str = 'MM/dd') => {
  if (Object.prototype.toString.call(value) == "[object Date]") {
    return Utilities.formatDate(value, 'JST', str);
  }
  else {
    return value;
  }
};
const staffData_ = (name, keys = ['スタッフ名', '銀行名', '支店名', '口座番号']) => {
  const data = staffObject_();
  const array = keys.map(key => data[name][key]);
  return array;
};
const staffObject_ = () => {
  return JSON.parse(properties('STAFF_OBJ'))
};
const staffEmailAddress_ = (name) => {
  const staffs = staffData_(['name', 'e-mail']);
  for (let i in staffs) {
    if (staffs[i].includes(name)) {
      var eMail = staffs[i][1];
      break;
    }
  }
  return eMail;
};
const zenkana2Hankana = (str) => {
  const kanaMap = {
    "ガ": "ｶﾞ", "ギ": "ｷﾞ", "グ": "ｸﾞ", "ゲ": "ｹﾞ", "ゴ": "ｺﾞ",
    "ザ": "ｻﾞ", "ジ": "ｼﾞ", "ズ": "ｽﾞ", "ゼ": "ｾﾞ", "ゾ": "ｿﾞ",
    "ダ": "ﾀﾞ", "ヂ": "ﾁﾞ", "ヅ": "ﾂﾞ", "デ": "ﾃﾞ", "ド": "ﾄﾞ",
    "バ": "ﾊﾞ", "ビ": "ﾋﾞ", "ブ": "ﾌﾞ", "ベ": "ﾍﾞ", "ボ": "ﾎﾞ",
    "パ": "ﾊﾟ", "ピ": "ﾋﾟ", "プ": "ﾌﾟ", "ペ": "ﾍﾟ", "ポ": "ﾎﾟ",
    "ヴ": "ｳﾞ", "ヷ": "ﾜﾞ", "ヺ": "ｦﾞ",
    "ア": "ｱ", "イ": "ｲ", "ウ": "ｳ", "エ": "ｴ", "オ": "ｵ",
    "カ": "ｶ", "キ": "ｷ", "ク": "ｸ", "ケ": "ｹ", "コ": "ｺ",
    "サ": "ｻ", "シ": "ｼ", "ス": "ｽ", "セ": "ｾ", "ソ": "ｿ",
    "タ": "ﾀ", "チ": "ﾁ", "ツ": "ﾂ", "テ": "ﾃ", "ト": "ﾄ",
    "ナ": "ﾅ", "ニ": "ﾆ", "ヌ": "ﾇ", "ネ": "ﾈ", "ノ": "ﾉ",
    "ハ": "ﾊ", "ヒ": "ﾋ", "フ": "ﾌ", "ヘ": "ﾍ", "ホ": "ﾎ",
    "マ": "ﾏ", "ミ": "ﾐ", "ム": "ﾑ", "メ": "ﾒ", "モ": "ﾓ",
    "ヤ": "ﾔ", "ユ": "ﾕ", "ヨ": "ﾖ",
    "ラ": "ﾗ", "リ": "ﾘ", "ル": "ﾙ", "レ": "ﾚ", "ロ": "ﾛ",
    "ワ": "ﾜ", "ヲ": "ｦ", "ン": "ﾝ",
    "ァ": "ｧ", "ィ": "ｨ", "ゥ": "ｩ", "ェ": "ｪ", "ォ": "ｫ",
    "ッ": "ｯ", "ャ": "ｬ", "ュ": "ｭ", "ョ": "ｮ",
    "。": "｡", "、": "､", "ー": "ｰ", "「": "｢", "」": "｣", "・": "･"
  };
  const reg = new RegExp('(' + Object.keys(kanaMap).join('|') + ')', 'g');
  return str
    .replace(reg, match => {
      return kanaMap[match];
    })
    .replace(/゛/g, 'ﾞ')
    .replace(/゜/g, 'ﾟ');
};
const slimstaffData_ = (staffs, keys) => {
  const database = mainData_('sh')
    .getSheetByName('データベース').getDataRange().getDisplayValues();
  const label = database[0];
  const names = database.map(values => values[0]).flat();
  staffs = staffs.map(key => names.indexOf(key));
  keys = keys.map(key => label.indexOf(key));
  const slim = staffs.map(name => keys.map(key => database[name][key]));
  return slim;
};
const allStaffData = () => {
  return mainData_('sh').getSheetByName('データベース')
    .getDataRange().getValues();
};
const memberData_ = () => {
  return JSON.parse(properties('member_obj'));
  return mainData_('sh').getSheetByName('MGデータベース')
    .getDataRange().getDisplayValues();
  // const keys = database[0]
  // database.shift()
  // const object = database.map(values => {
  //   const obj = {}
  //   values.map((value, index) => {
  //     obj[keys[index]] = value
  //   })
  //   return obj
  // })
  // return object
};
const getName_ = () => {
  const database = mainData_('sh').getSheetByName('MGデータベース')
    .getDataRange().getDisplayValues();
  const label = database[0];
  const account = String(Session.getActiveUser());
  const name = database.filter(values => values.includes(account))
    .flat()[label.indexOf('name')];
  return name;
};
const kanji2num_ = (str) => {
  var reg;
  var kanjiNum = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '〇'];
  var num = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'];
  for (var i = 0; i < num.length; i++) {
    reg = new RegExp(kanjiNum[i], 'g'); // ex) reg = /三/g
    str = str.replace(reg, num[i]);
  }
  return str;
};
const zen2han_ = (str) => {
  return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
};
const address_trim_ = (value) => {
  return value
    .replace(/[^\x01-\x7E\xA1-\xDF]/g, str => zen2han_(str)).replace(/[\n\r]/g, '')
    .replace(/(?<=\d)[ーｰ－−-]|(丁目(?=\d)|番地の?(?=\d)|(?<=\d)番(?!([地 　]|$)))/g, '-')
    .replace(/[一二三四五六七八九十〇](?=-)|(?<=-)[一二三四五六七八九十〇]/g, str => kanji2num_(str))
    .replace(/[　]|(?<=\d)[\(（]|(番地|号|番(?!地))(?=[\(（])|(番地|番|号)([ 　]|$)/g, ' ')
    .replace(/!.*! |[\(\)（）]|[\s]{2,}|(?<!(\d|丁目))\s/g, '');
};
const split_a_ = (value) => {
  const address = address_trim_(value).match(/^.*\d(丁目)?(?=(\s|$))/)
  if (Boolean(address)) {
    return address[0]
  }
  return ''
};
const split_b_ = (value) => {
  const address = address_trim_(value).match(/(?<=\s).*$/)
  if (Boolean(address)) {
    return address[0]
  }
  return ''
};
const addressUPDATE_ = (sheet) => {
  if (!sheet) {
    sheet = mainData_('vc').getSheetByName('集約');
  }
  const sheet_dat = sheet.getDataRange().getValues();
  let label = sheet_dat.filter(values => values.includes('会場\n住所')).flat();
  if (!label) {
    label = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues().flat();
  }
  var array = sheet.getRange(3, label.indexOf('会場\n住所') + 1, sheet.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()).getDisplayValues().flat().map((value) => {
    if (value.match(/^.*(?<=[-][1-9]??)[0-9](?![\d])/) != null) {
      value = value.replace(/^[!].*[!][ ]?/, '');
      return value.match(/^.*(?<=[-][1-9]??)[0-9](?![\d])/);
    }
    else {
      return [value.replace(/^[!].*[!][ ]?/, '')];
    }
  });
  sheet.getRange(3, label.indexOf('住所') + 1, array.length).setValues(array);
};
function display_() {
  return Object.keys(this).filter(value => value != 'display');
}
const labelCreate = (arg) => {
  //新会場連絡シートのラベルを設定。
  //アサインシート側のラベルに名称を変更。
  const label = arg[0];
  return label.map(value => value
    .replace(/(^会場$|^会場名$)/, '会場\n名称')
    .replace(/^可否$/, '開催\n可否')
    .replace(/^日付$/, '日程')
    .replace(/(^メイン$|^変更後メンバー$)/, 'メイン\n講師')
    .replace(/(^サポート1$|^メンバー1$)/, 'サポート講師')
    .replace(/^メンバー/, 'サポート')
    .replace(/^会場特性$/, '会場運用上\n注意点')
    .replace(/^備考$/, 'カリキュラム\n補足')
    .replace(/^予定$/, '参加予定人数'));
};
const datereplace = (value, index) => {
  if (index == 0) {
    return [value];
  }
  return JSON.parse(`[${value.replace(/日 |日,/g, ',').replace(/,$|日$/, '')}]`);
};
const main_flag = (flag, str) => {
  switch (true) {
    case flag && str: return 'メイン';
    case flag: return 'サポート';
    default: return 'SB同行';
  }
};
const isNaN_ = (value) => {
  return typeof value === 'number'
};
const addMonthSheet_ = (spread, date) => {
  const set_date = new Date(date)
  let sheet
  try { sheet = spread.insertSheet(dateString(set_date, 'yyyyMM')) }
  catch { sheet = spread.getSheetByName(dateString(set_date, 'yyyyMM')) }
  set_date.setMonth(set_date.getMonth() - 1)
  const origin = spread.getSheetByName(dateString(set_date, 'yyyyMM'))
  const label = origin.getRange(1, 1, 1, origin.getLastColumn()).getValues().flat()
  const dayrange = [
    `${NumToA1(label.indexOf('日程') + 1)}2:${NumToA1(label.indexOf('日程') + 1)}999`,
    `${NumToA1(label.indexOf('更新日') + 1)}2:${NumToA1(label.indexOf('更新日') + 1)}999`,
    `${NumToA1(label.indexOf('確認日') + 1)}2:${NumToA1(label.indexOf('確認日') + 1)}999`,
    `${NumToA1(label.indexOf('配備日') + 1)}2:${NumToA1(label.indexOf('配備日') + 1)}999`,
  ]
  const timerange = [
    `${NumToA1(label.indexOf('開始') + 1)}2:${NumToA1(label.indexOf('開始') + 1)}999`,
    `${NumToA1(label.indexOf('終了') + 1)}2:${NumToA1(label.indexOf('終了') + 1)}999`,
  ]
  const numrange = [
    `${NumToA1(label.indexOf('定員\n(半角)') + 1)}2:${NumToA1(label.indexOf('定員\n(半角)') + 1)}999`,
    `${NumToA1(label.indexOf('参加予定人数') + 1)}2:${NumToA1(label.indexOf('参加予定人数') + 1)}999`,
    `${NumToA1(label.indexOf('必要キャリー数\n(半角)') + 1)}2:${NumToA1(label.indexOf('必要キャリー数\n(半角)') + 1)}999`,
    `${NumToA1(label.indexOf('アサイン数\n(半角)') + 1)}2:${NumToA1(label.indexOf('アサイン数\n(半角)') + 1)}999`,
    `${NumToA1(label.indexOf('配備数1') + 1)}2:${NumToA1(label.indexOf('配備数1') + 1)}999`,
    `${NumToA1(label.indexOf('配備数2') + 1)}2:${NumToA1(label.indexOf('配備数2') + 1)}999`,
    `${NumToA1(label.indexOf('配備数3') + 1)}2:${NumToA1(label.indexOf('配備数3') + 1)}999`,
    `${NumToA1(label.indexOf('配備数4') + 1)}2:${NumToA1(label.indexOf('配備数4') + 1)}999`,
    `${NumToA1(label.indexOf('配備数5') + 1)}2:${NumToA1(label.indexOf('配備数5') + 1)}999`,
    `${NumToA1(label.indexOf('配備数6') + 1)}2:${NumToA1(label.indexOf('配備数6') + 1)}999`,
    `${NumToA1(label.indexOf('配備数7') + 1)}2:${NumToA1(label.indexOf('配備数7') + 1)}999`,
    `${NumToA1(label.indexOf('配備数8') + 1)}2:${NumToA1(label.indexOf('配備数8') + 1)}999`,
    `${NumToA1(label.indexOf('配備数9') + 1)}2:${NumToA1(label.indexOf('配備数1') + 1)}999`,
    `${NumToA1(label.indexOf('配備数10') + 1)}2:${NumToA1(label.indexOf('配備数1') + 1)}999`,
  ]
  const textrange = [
    `${NumToA1(label.indexOf('開催No.') + 1)}2:${NumToA1(label.indexOf('開催No.') + 1)}999`,
    `${NumToA1(label.indexOf('会場\n番号') + 1)}2:${NumToA1(label.indexOf('会場\n番号') + 1)}999`,
    `${NumToA1(label.indexOf('LOG主催者TEL') + 1)}2:${NumToA1(label.indexOf('LOG主催者TEL') + 1)}999`,
    `${NumToA1(label.indexOf('LOG会場TEL') + 1)}2:${NumToA1(label.indexOf('LOG会場TEL') + 1)}999`,
    `${NumToA1(label.indexOf('通し番号') + 1)}2:${NumToA1(label.indexOf('通し番号') + 1)}999`,
  ]
  sheet.setFrozenRows(1)
  sheet.getRange(1, 1, 1, label.length).setValues([label])
  sheet.getRangeList(dayrange).setNumberFormat('MM/dd')
  sheet.getRangeList(timerange).setNumberFormat('H:mm')
  sheet.getRangeList(numrange).setNumberFormat('0')
  sheet.getRangeList(textrange).setNumberFormat('@')
  return sheet
}
const diffMap = (a, b) => {
  return a.map((value, index) => {
    if (value != b[index]) {
      return b[index]
    }
    return value
  })
}