const tatsuyacheck_ = () => {
  const mgshift = mainData_('mg').getSheetByName(Utilities.formatDate(start_time, 'JST', 'yyyy.M'));
  const shiftDat = mgshift.getDataRange().getValues().map(values => {
    return values.filter((values, index) => index > 0);
  });
  let tatsuyaShift = shiftDat.filter(values => values.includes('山崎'));
  tatsuyaShift.splice(1);
  tatsuyaShift = tatsuyaShift.flat();
  const today = start_time.getDate();
  const tatsuya_today = String(tatsuyaShift[today]);
  const tatsuya_check = (tatsuya_today != '休' && tatsuya_today != '休業');
  if (tatsuya_check) {
    stayCheck_();
  }
};
const stayCheck_ = () => {
  let body = '未回答の宿泊申請リストは以下の通りです。\n';
  const st = mainData_('st').getSheetByName('フォームの回答 1');
  const stdat = st.getDataRange().getValues().filter((values, index) => index > 0 && values[values.length - 1] != true);
  if (stdat[0] == null) {
    return;
  }
  const vcag = mainData_('vc').getSheetByName('集約');
  const vcagdat = vcag.getDataRange().getValues();
  const vcaglabel = vcagdat.filter(values => values.includes('日程')).flat();
  const keys = ['会場\n名称', '会場\n住所', 'シフト開始', 'シフト終了']
    .map(key => vcaglabel.indexOf(key));
  const member_keys = ['メイン\n講師', 'サポート講師', 'サポート2', 'サポート3', 'サポート4', 'サポート5']
    .map(key => vcaglabel.indexOf(key));
  const vcagdays = vcagdat.map(values => values = dateString(values[vcaglabel.indexOf('日程')])).flat();
  const trim_vcag = vcagdat.map(values => keys.map(key => values[key]));
  const trim_member = vcagdat.map(values => member_keys.map(key => values[key]));
  const staffs = staffData_(['name']).flat();
  const address = staffData_(['住所']).flat();
  const set_time = [];
  stdat.forEach(values => {
    let date_time = [];
    const name = values[1];
    const date = values[2].replace("-", "/");
    const day = Number(values[2].match(/[\d].$/));
    const app = values[3];
    let start_date = vcagdays.indexOf(date);
    const end_date = vcagdays.lastIndexOf(date);
    while (start_date <= end_date) {
      const member_check = trim_member[start_date].some(value => value == name);
      if (member_check) {
        var venue = trim_vcag[start_date][0];
        var venue_address = String(trim_vcag[start_date][1]);
        date_time.push([trim_vcag[start_date][2], trim_vcag[start_date][3]]);
        break;
      }
      ++start_date;
    }
    if (app == '中泊') {
      body += `\n${name}:${app}\n【日程】 ${date}\n【会場】 ${venue}\n`;
      return;
    }
    if (date_time.length > 0) {
      if (date_time.length > 1) {
        date_time = date_time.sort((a, b) => a[0].getTime() - b[0].getTime());
      }
      if (app == '前泊') {
        var time = new Date(date_time[0][0]);
        var origin = String(address[staffs.indexOf(name)]);
        var destination = venue_address;
      }
      else if (app == '後泊') {
        var time = new Date(date_time[date_time.length - 1][1]);
        var origin = venue_address;
        var destination = String(address[staffs.indexOf(name)]);
      }
      else {
        return;
      }
      body += `\n${name}:${app}\n【日程】 ${date}\n【会場】 ${venue}\n`;
      time.setFullYear(start_time.getFullYear());
      time.setMonth(start_time.getMonth());
      time.setDate(day);
      const transit_object = Maps.newDirectionFinder().setOrigin(origin).setDestination(destination)
        .setArrive(time).setMode(Maps.DirectionFinder.Mode.TRANSIT).setLanguage('ja').getDirections();
      try {
        var transit_route = transit_object.routes[0].legs[0];
      }
      catch {
        body += '検索エラー\n';
        return;
      }
      const duration_time = transit_route.duration.text;
      const departure_time = transit_route.departure_time.text;
      const arrival_time = transit_route.arrival_time.text;
      body += `【出発】 ${arrival_time}\n【到着】 ${departure_time}\n【移動時間】 ${duration_time}\n`;
    }
  });
  console.log(body)
};
