// 集約したトリガーを格納
const zeroOclock = () => {
  try { toDay_() } catch (e) { console.log(`toDay_()は失敗しました。\n${e}`) }
  try { writeForm_() } catch (e) { console.log(`writeForm_()は失敗しました。\n${e}`) }
  try { diffCheck_() } catch (e) { console.log(`diffCheck_()は失敗しました。\n${e}`) }
  try { folderCreate_() } catch (e) { console.log(`folderCreate_()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(0, 0, 0, 0);
  triggerset('zeroOclock', start_time);
};
const nineOclock = () => {
  try { workerTemp(); } catch (e) { console.log(`workerTemp()は失敗しました。\n${e}`) }
  try { tatsuyacheck_() } catch (e) { console.log(`tatsuyacheck_()は失敗しました。\n${e}`) }
  try { diffCheck_() } catch (e) { console.log(`diffCheck_()は失敗しました。\n${e}`) }
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
  try { diffCheck_() } catch (e) { console.log(`diffCheck_()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(12, 0, 0, 0);
  triggerset('twelveOclock', start_time);
};
const fifteenOclock = () => {
  try { holiDayTemp() } catch (e) { console.log(`holiDayTemp()は失敗しました。\n${e}`) }
  try { tatsuyacheck_() } catch (e) { console.log(`tatsuyacheck_()は失敗しました。\n${e}`) }
  try { diffCheck_() } catch (e) { console.log(`diffCheck_()は失敗しました。\n${e}`) }
  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(15, 0, 0, 0);
  triggerset('fifteenOclock', start_time);
};
