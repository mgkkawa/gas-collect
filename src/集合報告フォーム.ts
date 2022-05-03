// Compiled using gas_collect 1.2.0 (TypeScript 4.6.4)
const writeForm_ = () => {
  const start_time = new Date();
  start_time.setHours(0, 0, 0, 0);
  const get_time = start_time.getTime();
  start_time.setDate(start_time.getDate() + 1);
  const end_time = start_time.getTime();
  const sheetname = dateString(start_time, 'yyyyMM');
  triggerset('writeForm', start_time);
  const form = mainForm_('together');
  const items = form.getItems();
  const question_1 = items[0];
  const question_2 = items[1];
  const ss = mainData_('as');
  const sheet = ss.getSheetByName(sheetname);
  const data = sheet.getDataRange().getValues();
  let ind = 0;
  const label = data.filter((values, index) => {
    if (values.includes('日程') && ind == 0) {
      ind = index;
    }
    return values.includes('日程');
  }).flat();
  data.splice(0, ind + 1);
  const keys = ['日程', '会場\n名称', 'メイン\n講師', 'サポート講師', 'サポート2', 'サポート3', 'サポート4', 'サポート5']
    .map(key => label.indexOf(key));
  let name_list = [];
  let venue_list = [];
  const trim_data = data.map(values => keys.map(key => values[key]).filter(value => value != ''))
    .filter(values => new Date(values[0]).getTime() >= get_time && new Date(values[0]).getTime() < end_time)
    .forEach(values => {
      venue_list.push(values[1]);
      for (let i = 2; i < keys.length; i++) {
        name_list.push(values[i]);
      }
    });
  name_list = name_list.filter((v, i, a) => v != null && a.indexOf(v) == i)
    .map(value => [value]);
  venue_list = venue_list.filter((v, i, a) => v != null && a.indexOf(v) == i);
  Logger.log(name_list);
  Logger.log(venue_list);
  question_1.asListItem().setChoiceValues(venue_list);
  question_2.asListItem().setChoiceValues(name_list);
};
