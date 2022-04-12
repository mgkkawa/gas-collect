var start_time = new Date();

function doGet() {
  zeroOclock();
  nineOclock();
  nineHirfOclock();
  fifteenOclock();
}

const triggerset = (t: string, time: Date) => {
  if (!t) {
    t = 'check';
  }
  ;
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() == t) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(t).timeBased().at(time).create();
};

const mainData = (s) => {
  switch (s) {
    case 'vc': return SpreadsheetApp.openById('12px9xnwlW5W4lkcYkT3ifQY_WbroAz4Hc-MFp6H0-GM'); //[移行先]
    case 'as': return SpreadsheetApp.openById('1m93CFX1uG67bO6c5xbSGoV5Bm0xNbfO0QAkE7nQqO5c'); //【ソフトバンク様】共有用アサインシート
    case 'sh': return SpreadsheetApp.openById('14KJJ0cDL_iwIyYOFpHoutgBa1IhFz-C0bGLrru-V6Vw'); //SBスマホ教室シフト ver.2
    case 'nh': return SpreadsheetApp.openById('16Sauir48G8L5nYlZn6PXRA_4dYbrg-K_oM5A2yN7828'); //翌月希望休申請フォーム
    case 'ss': return SpreadsheetApp.openById('1QUgp80m71kV3Z-tZF5kd66MZPgBkaAJzNUzEkvumKoI'); //編集用
    case 'mg': return SpreadsheetApp.openById('1Li-BWteJg-Nn4nWjv3Ha1XXl_2gz-IGLZcpZMV8WiJo'); //MGシフト
    case 'st': return SpreadsheetApp.openById('16r_yPRELNL57_kq03CE4OsIkZS6WiEmAsUxXqkvPGlU'); //宿泊申請用フォーム(回答)
    case 'wr': return SpreadsheetApp.openById('1vsEs4HxI9tyQaduT-XLxUfKZQcbZHdGGaPlklYuu524'); //勤務実績表
    case 'cr': return SpreadsheetApp.openById('1xaqLyOfVIZDiy7KuYP3_aOjjxk7Pc8JtFdvzlCA_ZdQ'); //フォルダ作成用スプレッド
    case 'ex': return SpreadsheetApp.openById('1D1bUKQviM7mOkZozknLRk2g8_oQ7t6w0EgQc_4l6Vnk'); //経費申請書
    case 'sc': return SpreadsheetApp.openById('1l-2c0yjfQTTcfnuhvOnVkf_aU6NuEXMXveG32nxs_w8'); //新スーツケース管理表
    case 'tm': return SpreadsheetApp.openById('1ZjRt8x-PZ8QUIukTATts1v3n-89JH6R9Hz3wnA6DKC4'); //検温結果報告フォーム(回答)
  }
}

const testData = (s) => {
  switch (s) {
    case 'vc': return SpreadsheetApp.openById('1vj_jYh177A1LxSZ2xxLBNKnXxV-gGGfanrJpjqR-A8g');//[開発用]新会場連絡シート
    case 'sh': return SpreadsheetApp.openById('1WWgrfYK06KbQXaLgHQGSTHf4DMmZOFixElsTlqnbHHs');//[開発用]SBスマホ教室シフト ver.2
    case 'as': return SpreadsheetApp.openById('1W-1ktAbv71eBG8yoK1pdBKVBZdIcbCFyUhFBdyIxcKc');//[開発用]アサインシート
    //case 'nh': return;
  }
}

const setOptions = () => {
  return {
    "apiId": "jp2ZfyeuwcreZ",
    "consumerKey": "WCuVQsBn1LN1IlfD5_nx",
    "serverId": "0258d47c34e54a6db60f0374120c71bd",
    "privateKey": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC9D+Sp6ANTG8T9sbdhJFsipcwBa7HUEQZfj2t2XOtkiQhDOF9h7oC3w+wditc0bnOxUbgFy2mEaWZnhQ7wjgUY2elvrE/zSk3Dg+czKe+wYzvCQkNzBoAmrG/0qssQOv31ainTJgsEuvYuOnaOaf/0sQUWFbDBFK51bSSX0HWsZrPoyKQMu5aOBdxeO/TN/4Ks1c2115V/DR4ijsAQPXy9+0ajQHA/HpqJEXiRM1ZQbxZb8dXI1Z589vLHAWl6RheBHfsbCHKvYB/dRbEaul3N1M6+S8B7YG8WbUxZVwYtORbvRZ5pOwj2dGMUmYShNal+bylxTYADD9TXAfXiFWpZAgMBAAECggEAYEWjjtFSQBO37+d7FcBJmA8NHvwUBYTV1ftWIWOXig4tYu1lxJyKdwkRRsnYZB6KUxTlvC2kgYSaXMRoox3ugoUUVYVNAPopNxIHvQnxv8QIPhc3+W6p+wd7yv7dgFpJz5pLyfVpTvNVQJ0MmeBoMdWiXWiWJPu/CpSVOakxAqQ4uCC0P6mU+8l2jJHCVJlxG655Hc6efW+Tttj3/HgjrZt41OZaKDcpFbdNthKCdIqdFJUjtxnS7/BhRXmTDOdDrz8tDZ/g0jI1V2B+IJxpS/S2bhQ7xpkUMHl8gndDt4PUwPiazUvh1rVQ51a2V/I+XSH0MW9rOQt08NDnspNQAQKBgQDwDoheeU7eHGMU95o9XTM9O2dWZO1kjmLjCV1AmxYA/SigAqK3eajirqmHC2RQTR9oEz0+nVRB3MP2cdwviq8C3EEK9Uk5qo5hIiptEdB/C5vBR4c4pjKMLhmmnUxJOA7wTLuSDs7S/gV6TJnNjjayl63Ho9fE8oAaN8sspfwx2QKBgQDJnlgcVU4vPlARByO0QxR1ZOpdFClmFcb5K3IBZSLReNgAf8pAh7ZzydVeUayXbwFcm7hsfHBgsXrO/1xU25jLYYzZdy0dkjCYFklXml5s8okPKS1VlKS0p3pNXSMF/vTZnGRWv5GPdPzbvD/suIioicd6Yq5efH0WnJd3/IcsgQKBgFFYm11oPhGGDQ1N2jZxHqvhNWNRUCCDH39P7Qm3g/1RzDJssjb6QwCNr2TTt2p+t+XMnzRYm/S1mrIoYzS5ChUuDdxowieIZDcGHNpY7w4eIOmeE3Re3UIjXQ3Yv8EP/f/wAJ+95NgavEB1NKkzR9W9EGZWXweeJUq9HVxD4vIBAoGAaf9khWjX1QCmqX0eaUwG+n0Wtd6+sRvr4t3TbyZTutsUfJayp2ByLMfE9HRuD6NdjOmVePunPT6xuRgJ29cjqnFb0ozUXEKaZ6r33iWbVR3hruXQHrkqKowOQZQgsKtM4QdWvUCz7z7rHT8IQJT3y9u70v0EfUpnnTawb3WFP4ECgYEA4/uIzOXreoJOiMPR0KiZQxciXknd+U8EEDBhDs4nognqzxpAt4JdeIHVG++vrVBYvA0y9Y5Paggn3nHEBbx3Rj23U/cbpKInbLuDexzC8zr+G6Z/jr1R0O/oppeuTe3fMTuKg6sFbHnMuNpAuXGQCP603gW1axmQKboSYTzJ4wc=\n-----END PRIVATE KEY-----",
    "botNo": 1416739
  };
}

const accountId = (name) => {
  switch (name) {
    case '大山夏美': return 'mg.06210@mgcorporation';
    case '山崎達也': return 'mg.24637@mgcorporation';
    case '富樫一世': return 'mg.95657@mgcorporation';
    case 'room': return '88547072';
    case 'domainId': return '12013748';
    case 'options': return 'info-dsg@mg-k.co.jp';
    default: return 'k.kawate@mgcorporation';
  }
}

const callbackURL = () => {
  return "https://script.google.com/macros/s/AKfycbxkGkqNEqqA9QYvySiS4TUEpjd7poF7DLUYDL5G6ghZLgtOkmc/exec";
}

const addressCheck = () => {
  var sheet = SpreadsheetApp.openById('1m93CFX1uG67bO6c5xbSGoV5Bm0xNbfO0QAkE7nQqO5c').getSheetByName('202204');
  //var sheet = SpreadsheetApp.openById('1aF-KKlYVWMNBO95Gc4B2d70cie7fPApz-G7m0PR2bVQ').getSheetByName('シート2');
  var dat = sheet.getDataRange().getDisplayValues();
  for (var i = 0; i < dat.length; i++) {
    if (dat[i].indexOf('会場\n住所') != -1) {
      var addCol = dat[i].indexOf('会場\n住所');
      var row = i + 2;
      break;
    }
  }
  var addDat = sheet.getRange(row, addCol + 1, sheet.getLastRow() - row).getDisplayValues().flat();
  function check(str) {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
  }
  function kanji2num(str) { // 漢数字を半角数字に
    var reg;
    var kanjiNum = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '〇'];
    var num = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'];
    for (var i = 0; i < num.length; i++) {
      reg = new RegExp(kanjiNum[i], 'g'); // ex) reg = /三/g
      str = str.replace(reg, num[i]);
    }
    return str;
  }
  var reg = /[一二三四五六七八九十〇](?=丁目|番地|号)|番(?=$|[0-9 ])/g
  addDat = addDat.map(value => check(value)).map(value => value.replace(reg, function (s) { return kanji2num(s); }));
  addDat = addDat.map(value => value.replace(/(?<=[0-9])(丁目|番地|番地の|[番のー－ｰ‐])(?=[0-9])/g, '-'));
  addDat = addDat.map(value => value.replace(/(?<=[0-9])(番地|[番号])(?!地|[0-9])[　| ]?|[　]|\n|\r\n|\r/g, ' '));
  var setAdd = addDat.map(address => [address.replace(/  /g, ' ')]);
  sheet.getRange(row, addCol + 1, setAdd.length).setValues(setAdd);
}

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
}

const datObject = (array) => {
  //受け取った配列を連想配列化して返す。
  var keys = array[0];
  array.shift();
  var obj = array.map(values => {
    var hash = {};
    values.map((value, x) => hash[keys[x]] = value)
    return hash;
  })
  return obj;
}

const objectCut = (obj, keys) => {
  //受け取った連想配列を受け取ったキーで取り出して二次元配列として返す。
  return obj.map(array => keys.map(key => array[key]));
}

const convertDate = (values, str) => {
  if (!str) { str = 'yyyy/MM/dd' };
  //date型をstringに変換
  for (var i = 0; i < values.length; i++) {
    var newValues = values[i].map(
      function (x) {
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
}

const convertObj = (values) => {
  var reg = /^....\/..\/..$/;
  for (var i = 0; i < values.length; i++) {
    var newValues = values[i].map(
      function (x) {
        var regmatch = x.match(reg);
        if (regmatch != null) {
          return x = String(x.match(/(?!<\/)..\/..$/));
        } else {
          return x;
        }
      });
    values[i] = newValues;
  }
  return values;
}

const month = (value) => {
  switch (true) {
    case value >= 5: return value.match(/(?<=\/)[0-9][1-9](?=\/)/);
    default: return value.slice(0, 2);
  }
}

const valueDate = (value, str = 'MM/dd') => {
  if (Object.prototype.toString.call(value) == "[object Date]") {
    return Utilities.formatDate(value, 'JST', str);
  } else {
    return value;
  }
};

const staffData = (keys = ['スタッフ名', '銀行名', '支店名', '口座番号']) => {
  const data = staffObject();
  const array = data.map(function (staff) { return keys.map(function (key) { return staff[key]; }); });
  return array;
};

const staffObject = () => {
  const database = SpreadsheetApp.openById('14KJJ0cDL_iwIyYOFpHoutgBa1IhFz-C0bGLrru-V6Vw').getSheetByName('データベース').getDataRange().getDisplayValues();
  const keys = database[0];
  database.shift();
  const array = database.map(function (values) {
    let hash = {};
    values.map(function (value, index) {
      hash[keys[index]] = value;
    });
    return hash;
  });
  return array;
};

const staffEmailAddress = (name) => {
  const staffs = staffData(['name', 'e-mail']);
  for (let i in staffs) {
    if (staffs[i].includes(name)) {
      var eMail = staffs[i][1];
      break;
    }
  }
  return eMail;
}

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
    .replace(reg, function (match) {
      return kanaMap[match];
    })
    .replace(/゛/g, 'ﾞ')
    .replace(/゜/g, 'ﾟ');
}
const slimstaffData = (staffs, keys) => {
  const database = mainData('sh')
    .getSheetByName('データベース').getDataRange().getDisplayValues();
  const label = database[0];
  const names = database.map(values => values[0]).flat();
  staffs = staffs.map(key => names.indexOf(key));
  keys = keys.map(key => label.indexOf(key));
  const slim = staffs.map(name => keys.map(key => database[name][key]));
  return slim;
};

const allStaffData = () => {
  return mainData('sh').getSheetByName('データベース')
    .getDataRange().getValues();
}


const memberData = () => {
  return mainData('sh').getSheetByName('MGデータベース')
    .getDataRange().getDisplayValues();
  // const keys = database[0];
  // database.shift();
  // const object = database.map(values => {
  //   const obj = {};
  //   values.map((value, index) => {
  //     obj[keys[index]] = value;
  //   })
  //   return obj;
  // })
  // return object;
}

const getName = () => {
  const database = mainData('sh').getSheetByName('MGデータベース')
    .getDataRange().getDisplayValues();
  const label = database[0];

  const account = String(Session.getActiveUser());

  const name = database.filter(values => values.includes(account))
    .flat()[label.indexOf('name')];
  return name;
}
