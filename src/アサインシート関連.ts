
const assaignsheet = () => {

  const as = mainData_('as');
  const asag = as.getSheetByName('集約');
  const asag_label = asag.getRange(2, 1, 1, asag.getLastColumn()).getValues().flat();

  let ind = 0;
  const sheetname = '202109';
  const sheet = as.getSheetByName(sheetname);
  const sheet_data = sheet.getDataRange().getValues();
  const sheet_label = sheet_data.filter((values, index) => {
    if (values.includes('日程')) { ind = index + 2; };
    return values.includes('日程');
  }).flat();

  const sheet_indexs = fill(asag_label.map(key => sheet_label.indexOf(key) + 1));


  let Col_list = '';
  sheet_indexs.forEach((key, index) => {
    if (index == 0) { Col_list = `${NumToA1(key)}, `; }
    else if (index == sheet_indexs.length - 1) { Col_list += `${NumToA1(key)}`; }
    else { Col_list += `${NumToA1(key)}, `; }
  });

  const dayCol = NumToA1(sheet_label.indexOf('日程') + 1);
  const venCol = NumToA1(sheet_label.indexOf('会場\n名称') + 1);
  const lastCol = NumToA1(sheet_label.length);
  const initial = `'${sheetname}'!A${ind}:${lastCol}`;
  const func = `QUERY(${initial},"select ${Col_list} where ${dayCol} <>'' and ${venCol} <>''")`;

  const base_func = asag.getRange('F1').getValue();
  Logger.log(base_func);
  const write_func = base_func.replace(/\)}/, `\);${func}}`);
  asag.getRange('F1').setValue(write_func);
  asag.getRange('A3').setValue('=' + write_func);

  Logger.log(write_func);
};

const maxfill = (ary) => {
  return ary.map(value => {
    if (value != 0) {
      return value;
    } else {
      let i = 0;
      while (i <= ary.reduce((a, b) => Math.max(a, b))) {
        if (ary.indexOf(i) == -1) { return i; }
        ++i;
      }
    }
  })
};

const fill = (ary) => {
  const maxnum = ary.reduce((a, b) => Math.max(a, b));
  ary.forEach((num, index) => {
    if (num == 0) {
      let i = 0;
      while (i <= maxnum) {
        if (ary.indexOf(i) == -1) {
          ary.splice(index, 1, i);
        }
        ++i;
      }
    } else { return num; }
  })
  return ary;
};
