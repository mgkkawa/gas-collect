const calculationEoMonth = () => { return calculation('月末分'); }
const calculationFirstHalf = () => { return calculation('15日〆分'); }

const advanceBorrowing = () => { return calculation('前借分'); }

const calculation = (timing) => {
  const exag = SpreadsheetApp.openById('1PwQn8OqKnFu2aC4BvdcJr43_tDUL90Mtd7ZfCLfCZF8');
  let url = exag.getUrl();
  try {
    var exsheet = exag.insertSheet(Utilities.formatDate(new Date(), 'JST', 'yyyy.M.d'));
  }
  catch (_a) {
    var exsheet = exag.getSheetByName(Utilities.formatDate(new Date(), 'JST', 'yyyy.M.d'));
  }
  finally {
    var shid = exsheet.getIndex();
  }
  url += "#gid=" + shid;
  let body = url + '\n\n' + timing + ' 経費計算を開始しました。\n計算完了までしばらくお待ちください。';
  try {
    LINEWORKS.sendMsgRoom(setOptions_(), '130629262', body);
  }
  catch (_c) {
    var addbody = '経費申請用グループへの投稿にエラーが発生しました。';
    addbody += '\n\n送信予定の本文は以下の通りです。\n';
    body = addbody + body;
    LINEWORKS.sendMsg(setOptions_(), accountId_('山崎達也'), body);
  }
  //LINEWORKS.sendMsg(setOptions_(), accountId_(), body);
  const formlabel = [['スタッフ名', '口座', '', '', '', '合計金額', '減算額合計', '当月振込額合計'], ['', '銀行', '支店', '口座番号', '口座名義', '', '', '']];
  exsheet.getRange(1, 1, formlabel.length, formlabel[0].length).setValues(formlabel)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9).setFontWeight('bold').setBackground('#D3D3D3');
  ;
  exsheet.getRange(1, 1, 2).merge();
  exsheet.getRange(1, 2, 1, 4).merge();
  exsheet.getRange(1, 6, 2, 3).mergeVertically();
  exsheet.getRange(exsheet.getLastRow() + 1, formlabel[1].indexOf('口座番号') + 1, 50).setNumberFormat('@');
  exsheet.getRange(exsheet.getLastRow() + 1, formlabel[0].indexOf('合計金額') + 1, 50, 3).setNumberFormat('[$¥-411]#,##0');
  const sheets = SpreadsheetApp.openById('1kn_kB0bTaKflVtvlhMXTODbUo63gMbTW7j-m8hNt85s').getSheets();
  const sheet = sheets[0];
  let dat = sheet.getDataRange().getValues();
  const keys = dat[0];
  dat = objectCut_(datObject_(dat), keys).map(function (array) {
    array = array.map(function (x) {
      var type = Object.prototype.toString.call(x);
      if (type == "[object Date]") {
        return x = x.getTime();
      }
      else {
        return x;
      }
    });
    return array;
  });
  const check = new Date(2022, new Date().getMonth(), 1, 0, 0, 0, 0);
  //Logger.log(Utilities.formatDate(check, 'JST', 'yyyy/MM/dd'))
  const staffData = staffData_(['name', '銀行名', '支店名', '口座番号', 'スタッフ名']);
  let checkname = staffData.map(function (array) { return array[0]; });
  let paylist = [];
  let allsum = 0;
  let allpaid = 0;
  let time = check.getTime();
  switch (timing) {
    case '月末分':
      check.setMonth(check.getMonth() - 1);
      break;
    case '15日〆分':
      let time = check.getTime();
    default:
      dat = dat.filter(function (array) { return array[0] >= time && array[2] == timing; })
        .map(function (array) {
          array[3] = Number(array[3].match(/[1-9]?[\d]$/));
          return array;
        });
      checkname = dat.map(function (array) { return array[1]; });
  }
  Logger.log(dat);
  Logger.log(checkname);
  const root = DriveApp.getFolderById('1UT1mgpweki9sixQ3ZCteV1Oh_p49JvYq');
  const rootfolder = root.getFoldersByName(Utilities.formatDate(check, 'JST', 'yyyy.MM')).next();
  const folders = rootfolder.getFolders();
  const mimeType = 'application/vnd.google-apps.spreadsheet';
  var _loop_1 = function () {
    var folder = folders.next();
    var staff = String(folder.getName().match(/(?<=[.]).*$/));
    if (checkname.indexOf(staff) == -1) {
      return "continue";
    }
    var id = folder.getFilesByType(mimeType).next().getId();
    var exform = SpreadsheetApp.openById(id).getSheets();
    var sum = 0;
    var paid = 0;
    var pay = exform[1].getDataRange().getValues().filter(function (array) { return Number(array[0]) >= 1 && array[6] != ''; });
    var ex2 = exform[2].getDataRange().getValues().filter(function (array) { return Number(array[0]) >= 1 && array[6] != ''; });
    ex2.forEach(function (array) { return pay.push(array); });
    if (timing == '前借分') {
      pay = pay.filter(function (array) {
        return array[3] >= dat[checkname.indexOf(staff)][3]
          && array[3] <= dat[checkname.indexOf(staff)][3] + dat[checkname.indexOf(staff)][4];
      });
    }
    pay = pay.map(function (array) {
      array[0] = staff;
      sum += array[6];
      return array;
    });
    try {
      exform[4].getDataRange().getValues().forEach(function (a, x) {
        if (a[0] != '' && x >= 1) {
          paid += a[0];
          sum -= a[0];
        }
      });
      allpaid += paid;
    }
    catch (_b) { }
    ;
    var set = staffData.filter(function (array) { return array.indexOf(staff) != -1; }).concat([sum, paid, (sum + paid)]).flat();
    paylist.push(set);
    allsum += sum;
    Logger.log(pay);
  };
  while (folders.hasNext()) {
    _loop_1();
  }
  paylist = paylist.sort(function (a, b) {
    switch (true) {
      case a[4] > b[4]: return 1;
      case a[4] < b[4]: return -1;
      default: return 0;
    }
  });
  paylist = paylist.map(function (array) {
    return array.map(function (value, x) {
      if (x == 3) {
        return String(value).padStart(7, '0');
      }
      if (x == 4) {
        return zenkana2Hankana(value);
      }
      else {
        return value;
      }
    });
  });
  Logger.log(paylist);
  exsheet.getRange(exsheet.getLastRow() + 1, 1, paylist.length, paylist[0].length)
    .setValues(paylist);
  exsheet.getRange(exsheet.getLastRow() + 1, 1, 2, 5).setValue('総計')
    .merge().setHorizontalAlignment('right').setVerticalAlignment('middle')
    .setFontSize(9).setFontWeight('bold').setBackground('#D3D3D3');
  exsheet.getRange(exsheet.getLastRow(), formlabel[0].indexOf('合計金額') + 1, 2, 3)
    .setValues([['今回支払総額', '減算総額', '該当月支払総額'], [allsum, allpaid, allsum + allpaid]]).setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setFontSize(9).setFontWeight('bold').setNumberFormat('[$¥-411]#,##0');
};
