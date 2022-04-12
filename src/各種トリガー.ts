// 集約したトリガーを格納

const zeroOclock = () => {
  toDay_();
  suiteCase_();

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(0);
  start_time.setMinutes(0);
  start_time.setSeconds(0);
  start_time.setMilliseconds(0);

  triggerset('zeroOclock', start_time);
}

const nineOclock = () => {
  workerTemp();
  tatsuyacheck_();

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(9);
  start_time.setMinutes(0);
  start_time.setSeconds(0);
  start_time.setMilliseconds(0);

  triggerset('nineOclock', start_time);
}

const nineHirfOclock = () => {
  visitCheck_();

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(9);
  start_time.setMinutes(0);
  start_time.setSeconds(0);
  start_time.setMilliseconds(0);

  triggerset('nineHirfOclock', start_time);
}

const fifteenOclock = () => {
  holiDayTemp();
  tatsuyacheck_();

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(15);
  start_time.setMinutes(0);
  start_time.setSeconds(0);
  start_time.setMilliseconds(0);


  triggerset('fifteenOclock', start_time);
}
