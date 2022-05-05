// 集約したトリガーを格納
function doGet() {
  const date = new Date()
  const day = date.getDate();
  const hour = date.getHours();
  const minutes = date.getMinutes();
  diffTrigger()
  if (hour >= 15) {
    date.setDate(day + 1);
  }
  date.setHours(15, 0, 0, 0);
  triggerset('fifteenOclock', date);
  if (hour >= 12 && hour < 15) {
    date.setDate(day + 1);
  }
  date.setHours(12);
  triggerset('twelveOclock', date);
  if (hour >= 9 && minutes >= 30 && hour < 12) {
    date.setDate(day + 1);
  }
  date.setHours(9, 30, 0, 0);
  triggerset('nineHirfOclock', date);
  if (hour >= 9 && hour < 12) {
    date.setDate(day + 1);
  }
  date.setMinutes(0);
  triggerset('nineOclock', date);
  if (hour < 9) {
    date.setDate(day + 1);
  }
  date.setHours(0);
  triggerset('zeroOclock', date);
}
const zeroOclock = () => {
  try { toDay_() } catch (e) { console.log(`toDay_()は失敗しました。\n${e}`) }
  try { writeForm_() } catch (e) { console.log(`writeForm_()は失敗しました。\n${e}`) }
  try { folderCreate_() } catch (e) { console.log(`folderCreate_()は失敗しました。\n${e}`) }
  try { labelCheck_() } catch (e) { console.log(`labelCheck_()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(0, 0, 0, 0);
  triggerset('zeroOclock', start_time);
};
const nineOclock = () => {
  try { workerTemp(); } catch (e) { console.log(`workerTemp()は失敗しました。\n${e}`) }
  try { tatsuyacheck_() } catch (e) { console.log(`tatsuyacheck_()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(9, 0, 0, 0);
  triggerset('nineOclock', start_time);
};
const nineHirfOclock = () => {
  try { visitCheck_() } catch (e) { console.log(`visitCheck_()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(9, 30, 0, 0);
  triggerset('nineHirfOclock', start_time);
};
const twelveOclock = () => {
  try { holiDayTemp() } catch (e) { console.log(`holiDayTemp()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(12, 0, 0, 0);
  triggerset('twelveOclock', start_time);
};
const fifteenOclock = () => {
  try { holiDayTemp() } catch (e) { console.log(`holiDayTemp()は失敗しました。\n${e}`) }
  try { tatsuyacheck_() } catch (e) { console.log(`tatsuyacheck_()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(15, 0, 0, 0);
  triggerset('fifteenOclock', start_time);
};
const diffTrigger = () => {
  try { diffCheck_() } catch (e) { console.log(`diffCheck_()は失敗しました。\n${e}`) }
  const date = new Date()
  let hour = date.getHours()
  if (hour >= 20) {
    hour = 8
    date.setDate(date.getDate() + 1)
  } else if (hour % 2 == 0) {
    hour += 2
  } else {
    hour += 1
  }
  date.setHours(hour, 0, 0, 0)
  triggerset('diffTrigger', date)
}
const labelCheck_ = () => {
  const vc = mainData_('vc')
  const sheet = vc.getSheetByName('転記')
  let label = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues().flat()
  label = label.filter((key, index) => index <= label.indexOf('通し番号'))
  const stringify = JSON.stringify(label)
  const prop = properties('sheet_label')
  if (stringify != prop) {
    PropertiesService.getScriptProperties().setProperty('sheet_label', stringify)
    return console.log('sheet_labelを上書きしました。')
  }
  return
}