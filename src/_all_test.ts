const testestes = () => {
  const date = new Date();
  date.setDate(1);
  const vc = mainData_('vc');
  const casting = vc.getSheetByName('キャスティング');
  const cas_data = casting.getDataRange().getValues();
  const cas_obj = new Venuecall(cas_data, ['会場\n名称'], ['メイン\n講師', 'サポート講師']);
  if (Object.keys(cas_obj).filter(key => key != 'label').length == 0) {
    try {
      Browser.msgBox('対象のキャスティング情報がありませんでした。');
    }
    finally {
      Logger.log('対象のキャスティング情報がありませんでした。');
      return;
    }
  }
  return
  const as = mainData_('as');
  const to_assheet = as.getSheetByName(dateString(date, 'yyyyMM'));
  const main_assign = new Assign(to_assheet);
  const sh = mainData_('sh');
  const to_shsheet = sh.getSheetByName(dateString(date, 'yyyy.MM'));
  const main_table = new ShiftTable();
  const to_ = dateString(date, 'MM/');
  let sub_as;
  let sub_assign;
  let sub_sh;
  let sub_table;
  if (cas_obj.check()) {
    date.setMonth(date.getMonth() + 1);
    sub_as = as.getSheetByName(dateString(date, 'yyyyMM'));
    sub_assign = new Assign(sub_as);
    sub_sh = sh.getSheetByName(dateString(date, 'yyyy.MM'));
    sub_table = new ShiftTable(date);
  }
  for (let row in cas_obj) {
    let as_sheet = to_assheet;
    let sh_sheet = to_shsheet;
    let assign = main_assign;
    let table = main_table;
    if (row == 'label') {
      continue;
    }
    const obj = cas_obj[row];
    const day = obj.date;
    if (!day.includes(to_)) {
      as_sheet = sub_as;
      sh_sheet = sub_sh;
      assign = sub_assign;
      table = sub_table;
    }
    const venue = obj.venue;
    const start = obj.start;
    const ascheck = assign.rowNum(day, venue, start);
    const main = `${NumToA1(assign.maincol + 1)}${ascheck[0]}`; //メイン講師の貼り付け範囲
    const supind = assign.supcol;
    let support;
    if (obj.support.length > 1) {
      support = `${NumToA1(supind + 1)}${ascheck[0]}:${NumToA1(supind + obj.support.length)}${ascheck[0]}`;
    }
    else {
      support = `${NumToA1(supind + 1)}${ascheck[0]}`;
    }
    as_sheet.getRange(main).setValue(obj.main);
    as_sheet.getRange(support).setValues([obj.support]);
    obj.support.push(obj.main);
    const range = [];
    obj.support.filter(Boolean).forEach(staff => {
      range.push(table.getCell(day, staff));
    });
    sh_sheet.getRangeList(range).setValue(ascheck[1]);
    obj.support.filter(Boolean).forEach(staff => {
      Logger.log(`staff:${staff}\nset_num:${ascheck[1]}\nassign:${ascheck[0]}\ntable:${table.getCell(day, staff)}`);
    });
  }
}