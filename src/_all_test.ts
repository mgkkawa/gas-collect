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
  const main_assign = new AssignObject();
  const main_table = new ShiftTable();
  const to_ = dateString(date, 'MM/');
  let sub_assign;
  let sub_table;
  if (cas_obj.check()) {
    date.setMonth(date.getMonth() + 1);
    sub_assign = new Assign(date);
    sub_table = new ShiftTable(date);
  }
  for (let row in cas_obj) {
    if (row == 'label') {
      continue;
    }
    let assign = main_assign;
    let table = main_table;
    const obj = cas_obj[row];
    const day = obj.date;
    if (!day.includes(to_)) {
      assign = sub_assign;
      table = sub_table;
    }
    const venue = obj.venue;
    const start = obj.start;
    const ascheck = assign.rowNum(day, venue, start);
    const main = `${NumToA1(assign.maincol + 1)}${ascheck[0]}`; //メイン講師の貼り付け範囲
    const supind = assign.supcol;
    if (!obj.mg_flag) {
      assign.setValue(main, obj.main)
    }
    let support;
    if (obj.support.length > 1) {
      support = `${NumToA1(supind + 1)}${ascheck[0]}:${NumToA1(supind + obj.support.length)}${ascheck[0]}`;
    } else {
      support = `${NumToA1(supind + 1)}${ascheck[0]}`;
    }
    assign.setValues(support, [obj.support])
    obj.support.push(obj.main);
    const range = [];
    obj.support.filter(Boolean).forEach(staff => {
      range.push(table.getCell(day, staff));
    });
    table.sheet.getRangeList(range).setValue(ascheck[1]);
    obj.support.filter(Boolean).forEach(staff => {
      Logger.log(`staff:${staff}\nset_num:${ascheck[1]}\nassign:${ascheck[0]}\ntable:${table.getCell(day, staff)}`);
    });
  }
}