// 毎月の経費申請フォルダを自動作成。
// 経費申請書も各スタッフ用にコピー
const addMenuExFormFoldar = () => {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('追加機能');
  //menu.addItem('経費申請書作成');
  menu.addSeparator();
  menu.addItem('フォルダ作成', 'newExform');
  menu.addToUi();
}

const newExform = () => {
  /*フォルダ作成用スプレッドシートを参照
    何年何月の経費申請書、フォルダーを作成したいのか
    スプレッドシートを確認*/
  const sheetnamelast = ['', '', '(2)', '※領収書のみ'];
  const cr = mainData('cr');
  const crs = cr.getSheetByName('経費申請書フォルダ作成用');
  const check_Year_Month = crs.getRange(1, 1, 2, 1).getDisplayValues();
  const year = check_Year_Month.shift();
  const month = check_Year_Month.shift();
  const sheetname = `${year}年${month}月`;
  const foldername = `${year}.${String(month).padStart(2, '0')}`;
  sheetnamelast.forEach(function (s) { return Logger.log(sheetname + s); });
  const names = staffData(['name']).flat();
  /*経費申請書原本のスプレッドシートを参照
    各種成型しコピー、各スタッフフォルダへコピー*/
  const ex = mainData('ex');
  const exs = ex.getSheets();
  exs.forEach(function (sheet, index) {
    if (index != 1) { sheet.hideSheet(); }
    else { sheet.setName(sheetname).getRange(1, 2, 1, 2).setValues([[sheetname]]); }
    if (index >= 1 && index <= 3) { sheet.setName(sheetname + sheetnamelast[index]); }
  });
  var origin = DriveApp.getFileById('1D1bUKQviM7mOkZozknLRk2g8_oQ7t6w0EgQc_4l6Vnk');
  var formatName = origin.getName();
  //スタッフ共有用フォルダを参照
  var folder = DriveApp.getFolderById('1UT1mgpweki9sixQ3ZCteV1Oh_p49JvYq');
  var id = folder.createFolder(foldername).getId();
  var mfolder = DriveApp.getFolderById(id);
  names.forEach(function (name, i) {
    var sfolder = mfolder.createFolder(String(i + 1).padStart(2, '0') + '.' + name);
    var fileId = origin.makeCopy(formatName, sfolder).getId();
    var sheets = SpreadsheetApp.openById(fileId).addEditors(['k.kawate@mg-k.co.jp', 't.yamazaki@mg-k.co.jp', 'misano@mg-k.co.jp']).getSheets();
    sheets.forEach(function (sheet, x) {
      if (x != 1) {
        sheet.hideSheet();
      }
      else {
        sheet.getRange(3, 7).setValue(name);
      }
    });
  });
}
//sheet(スプレッドシートのタブ)とname(タブ名)を指定すればシート名変更
function nameSet(sheet, name) {
  sheet.setName(name);
}
function firstHalfShow() {
  var base_folder = DriveApp.getFolderById('1UT1mgpweki9sixQ3ZCteV1Oh_p49JvYq');
  var folderName = Utilities.formatDate(new Date(), 'JST', 'yyyy.MM');
  var folders = base_folder.getFoldersByName(folderName);
  while (folders.hasNext()) {
    var folder = folders.next();
    var targetFolderId = folder.getId();
  }
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  var staff_Folders = targetFolder.getFolders();
  while (staff_Folders.hasNext()) {
    var staff_folder = staff_Folders.next();
    var files = staff_folder.searchFiles("title contains '経費申請'");
    while (files.hasNext()) {
      var fileId = files.next().getId();
      var staffspread = SpreadsheetApp.openById(fileId);
      staffspread.addEditors(['t.yamazaki@mg-k.co.jp', 'misano@mg-k.co.jp']);
      var sheets = staffspread.getSheets();
      sheets[1].showSheet();
      sheets[2].showSheet();
      sheets[3].showSheet();
    }
  }
  var setTime = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 16, 1, 0, 0);
  var triggers = ScriptApp.getProjectTriggers();
  for (var _i = 0, triggers_1 = triggers; _i < triggers_1.length; _i++) {
    var trigger = triggers_1[_i];
    if (trigger.getHandlerFunction() == 'firstHalfShow') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger('firstHalfShow').timeBased().at(setTime).create();
}
function eoMonthShow() {
  var month = new Date(new Date().getFullYear(), new Date().getMonth() - 1);
  var base_folder = DriveApp.getFolderById('1UT1mgpweki9sixQ3ZCteV1Oh_p49JvYq');
  var folderName = Utilities.formatDate(month, 'JST', 'yyyy.MM');
  var folders = base_folder.getFoldersByName(folderName);
  while (folders.hasNext()) {
    var folder = folders.next();
    var targetFolderId = folder.getId();
  }
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  var staff_Folders = targetFolder.getFolders();
  while (staff_Folders.hasNext()) {
    var staff_folder = staff_Folders.next();
    var files = staff_folder.searchFiles("title contains '経費申請'");
    while (files.hasNext()) {
      var fileId = files.next().getId();
      var sheets = SpreadsheetApp.openById(fileId).getSheets();
      sheets[1].showSheet();
      sheets[2].showSheet();
      sheets[3].showSheet();
    }
  }
  var setTime = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1, 1, 0, 0);
  var triggers = ScriptApp.getProjectTriggers();
  for (var _i = 0, triggers_1 = triggers; _i < triggers_1.length; _i++) {
    var trigger = triggers_1[_i];
    if (trigger.getHandlerFunction() == 'eoMonthShow') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger('eoMonthShow').timeBased().at(setTime).create();
}
function firstHalfHide() {
  var base_folder = DriveApp.getFolderById('1UT1mgpweki9sixQ3ZCteV1Oh_p49JvYq');
  var folderName = Utilities.formatDate(new Date(), 'JST', 'yyyy.MM');
  var folders = base_folder.getFoldersByName(folderName);
  while (folders.hasNext()) {
    var folder = folders.next();
    var targetFolderId = folder.getId();
  }
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  var staff_Folders = targetFolder.getFolders();
  while (staff_Folders.hasNext()) {
    var staff_folder = staff_Folders.next();
    var files = staff_folder.searchFiles("title contains '経費'");
    while (files.hasNext()) {
      var fileId = files.next().getId();
      var sheets = SpreadsheetApp.openById(fileId).getSheets();
      sheets[2].hideSheet();
      sheets[3].hideSheet();
    }
  }
}
function exForm() {
}
