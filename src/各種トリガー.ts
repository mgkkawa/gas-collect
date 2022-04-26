// 集約したトリガーを格納

const zeroOclock = () => {
  try {
    toDay_();
    writeForm_();
    diffCheck_();
    folderCreate_();
  } finally { };

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(0, 0, 0, 0);

  triggerset('zeroOclock', start_time);
}

const nineOclock = () => {
  try {
    workerTemp();
    tatsuyacheck_();
    diffCheck_();
  }
  finally { };

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(9, 0, 0, 0);

  triggerset('nineOclock', start_time);
}

const nineHirfOclock = () => {
  try { visitCheck_(); }
  finally { };

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(9, 30, 0, 0);

  triggerset('nineHirfOclock', start_time);
}

const twelveOclock = () => {
  try {
    holiDayTemp();
    diffCheck_();
  }
  finally { };

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(12, 0, 0, 0);

  triggerset('twelveOclock', start_time);
}

const fifteenOclock = () => {
  try {
    holiDayTemp();
    tatsuyacheck_();
    diffCheck_();
  }
  finally { };

  start_time = new Date();
  start_time.setDate(start_time.getDate() + 1);
  start_time.setHours(15, 0, 0, 0);


  triggerset('fifteenOclock', start_time);
}
