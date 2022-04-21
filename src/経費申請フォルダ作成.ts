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
  const members = memberData_();
  const editors = ['川手健人', '山崎達也', '伊佐野美奈'].map(key => members[key]['e-mail']);

  const date = new Date();
  date.setMonth(date.getMonth() + 1);
  const year = valueDate(date, 'yyyy年');
  const month = valueDate(date, 'M月');

  const sheetname = year + month;
  const foldername = valueDate(date, 'yyyy.MM');
  const sheetnamelast = ['', '※領収書のみ'];
  const staff_obj = staffObject_();
  const names = Object.keys(staff_obj);


  const ex = mainData_('ex');
  const exs = ex.getSheets();
  sheetnamelast.forEach((value, index) => {
    exs[index].setName(sheetname + value);
    if (index == 0) { exs[index].getRange(1, 2, 1, 2).setValues([[year, month]]) };
  });
  const origin = DriveApp.getFileById(properties('origin_exform'));
  const formatName = origin.getName();
  //スタッフ共有用フォルダを参照
  const folder = DriveApp.getFolderById(properties('staff_exform_folder'));
  const id = folder.createFolder(foldername).getId();
  const mfolder = DriveApp.getFolderById(id);
  names.forEach((name, index) => {
    const sfolder = mfolder.createFolder(String(index + 1).padStart(2, '0') + '.' + name);
    const fileId = origin.makeCopy(formatName, sfolder).getId();
    const sheets = SpreadsheetApp.openById(fileId).addEditors(editors).getSheets();
    sheets.forEach((sheet, x) => {
      if (x != 0) {
        sheet.hideSheet();
        if (x == 1) { sheet.getRange('J6:J45').insertCheckboxes() };
      }
      else { sheet.getRange(3, 7).setValue(name); }
    });
  });
}
//sheet(スプレッドシートのタブ)とname(タブ名)を指定すればシート名変更
const nameSet = (sheet, name) => {
  sheet.setName(name);
}

const firstHalfShow = () => {
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
const eoMonthShow = () => {
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
const firstHalfHide = () => {
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
const exForm = () => {
}
