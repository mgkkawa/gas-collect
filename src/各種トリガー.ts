// 集約したトリガーを格納

const zeroOclock = () => {
  toDay();
  suiteCase();

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
  tatsuyacheck();

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(9);
  start_time.setMinutes(0);
  start_time.setSeconds(0);
  start_time.setMilliseconds(0);

  triggerset('nineOclock', start_time);
}

const nineHirfOclock = () => {
  visitCheck();

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
  tatsuyacheck();

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(15);
  start_time.setMinutes(0);
  start_time.setSeconds(0);
  start_time.setMilliseconds(0);


  triggerset('fifteenOclock', start_time);
}
