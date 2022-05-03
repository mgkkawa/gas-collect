// Compiled using gas_collect 1.2.0 (TypeScript 4.6.4)
const folderCreate_ = () => {
  start_time = new Date();
  start_time.setHours(0, 0, 0, 0);
  const start_stamp = start_time.getTime();
  const set_date = new Date(start_time);
  start_time.setDate(start_time.getDate() + 1);
  const end_stamp = start_time.getTime();
  const root = DriveApp.getFolderById('1WXWYUMIiw_U5zLgUtIYdtU1mEYiCfe8A');
  try {
    root.getFoldersByName(dateString(set_date, 'yyyy.MM')).next();
  }
  catch {
    root.createFolder(dateString(set_date, 'yyyy.MM'));
  }
  const month_folder = root.getFoldersByName(dateString(set_date, 'yyyy.MM')).next();
  try {
    month_folder.getFoldersByName(dateString(set_date, 'MM/dd')).next();
  }
  catch {
    month_folder.createFolder(dateString(set_date, 'MM/dd'));
  }
  const day_folder = month_folder.getFoldersByName(dateString(set_date, 'MM/dd')).next();
  const as_sheet = mainData_('as').getSheetByName(dateString(set_date, 'yyyyMM'));
  const as_data = as_sheet.getDataRange().getValues();
  const as_label = as_data.filter(values => values.includes('日程')).flat();
  const as_indexs = ['会場\n名称', 'メイン\n講師', 'サポート講師', 'サポート2', 'サポート3', 'サポート4', 'サポート5', '講師']
    .map(key => as_label.indexOf(key));
  const today_data = as_data.filter(values => {
    const stamp = new Date(values[as_label.indexOf('日程')]).getTime();
    const hold = (values[as_label.indexOf('開催\n可否')] != '中止');
    return (stamp >= start_stamp && stamp < end_stamp && hold);
  }).map(values => as_indexs.map(key => values[key]).filter(value => value != ''))
    .filter((array, index, obj) => {
      return obj.findIndex(values => values.includes(array[0])) == index;
    });
  today_data.forEach(values => {
    try {
      day_folder.getFoldersByName(values[0]).next();
    }
    catch {
      day_folder.createFolder(values[0]);
    }
    const ven_folder = day_folder.getFoldersByName(values[0]).next();
    if (values[values.length - 1] == 'エムジー') {
      const report_folder = ven_folder.createFolder('03.会場報告用');
      report_folder.createFolder('登壇風景');
      report_folder.createFolder('来店予約書き込みシート');
      report_folder.createFolder('座席表');
    }
    const assign_folder = ven_folder.createFolder('01.集合写真報告');
    const pic_folder = ven_folder.createFolder('02.アピアランス写真報告');
    Logger.log(values);
    values.filter((value, index) => index > 0 && index < values.length - 1)
      .forEach(staff => {
        Logger.log(staff);
        try {
          assign_folder.getFoldersByName(staff).next();
          pic_folder.getFoldersByName(staff).next();
        }
        catch {
          assign_folder.createFolder(staff);
          pic_folder.createFolder(staff);
        }
      });
  });
};
