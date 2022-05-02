const calculationEoMonth = () => { return calculation('月末分') }
const calculationFirstHalf = () => { return calculation('15日〆分') }
const advanceBorrowing = () => { return calculation('前借分') }

function calculation(e) {
  const staff_obj = staffObject_()//スタッフデータまとめ
  const indexs = Object.keys(staff_obj)
  //↓経費関連→スタッフ共有用フォルダ
  const root = DriveApp.getFolderById('1UT1mgpweki9sixQ3ZCteV1Oh_p49JvYq')
  const date = new Date()//現在時刻の取得
  const sheetname = dateString(date, 'yyyy.M.d')
  const mimeType = 'application/vnd.google-apps.spreadsheet'//SpreadSheetファイルを指定する。

  // チェックタイミングに応じて
  // 対象スタッフの絞り込み
  let name_list = []
  if (e == '月末分') {
    name_list = indexs
    date.setDate(0)
    date.setHours(0, 0, 0, 0)
  }
  else {
    if (e == '15日〆分') { date.setDate(16) }
    else { date.setDate(date.getDate() + 1) }
    date.setHours(0, 0, 0, 0)
  }
  const exform = mainData_('ef')
  const sheets = exform.getSheets()
  const sheet = sheets[0]
  const end_time = new Date(date)
  if (e != '月末分') {
    if (e == '15日〆分') { date.setDate(1) }
    else { date.setDate(date.getDate() - 5) }
    var sheet_data = sheet.getDataRange().getValues().filter(values =>
      new Date(values[0]).getTime() >= date.getTime() &&
      new Date(values[0]).getTime() < end_time.getTime() &&
      values.includes(e)
    )
    name_list = sheet_data.flatMap(values => values[1])
  }

  let write_data = []
  name_list.forEach(staff => {
    const num = String(indexs.indexOf(staff) + 1).padStart(2, '0').padEnd(3, '.')
    const array = ['銀行名', '支店名', '口座番号', '口座名義']
      .map(key => staff_obj[staff][key])
    const folder = root.getFoldersByName(dateString(date, 'yyyy.MM')).next()
    const staff_folder = folder.getFoldersByName(num + staff).next()
    const id = staff_folder.getFilesByType(mimeType).next().getId()
    const exform = SpreadsheetApp.openById(id)
    const ex_sheet1 = exform.getSheetByName(dateString(date, 'yyyy年M月'))
    const ex_sheet2 = exform.getSheetByName('当月精算履歴')
    const sheet1_data = ex_sheet1.getDataRange().getValues()
      .filter(values => {
        const type = (Object.prototype.toString.call(values[0]) == '[object Number]')
        switch (e) {
          case '月末分': return type && values[7] != ''
          case '15日〆分': return type && values[7] != '' && values[3] <= 15
          default:
            const staff_data = sheet_data.filter(values => values.includes(staff)).flat()
            let day = Number(staff_data[3].match(/[\d].$/))
            let add_day = day
            if (staff_data[4] != '') { add_day += Number(staff_data[4]) }
            return type && values[7] != '' && values[3] >= day && values[3] <= add_day
        }
      }).flatMap(values => values[6])
    const sum = sheet1_data.reduce((a, b) => a + b)
    const dsum = ex_sheet2.getRange('C1').getValue()


    array.unshift(staff)
    if (e != '前借分') {
      array.push(sum - dsum)
      array.push(dsum)
    }
    else {
      array.push(sum)
      array.push(0)
    }
    array.push(sum + dsum)

    if (e != '月末分') {
      ex_sheet2.getRange(ex_sheet2.getLastRow() + 1, 1, 1, 2)
        .setValues([[sum, dateString(new Date(), 'yyyy/M/d')]])
    }

    write_data.push(array)
  })
  write_data = write_data.sort((a, b) => {
    switch (true) {
      case a[4] > b[4]: return 1
      case a[4] < b[4]: return -1
      default: return 0
    }
  })

  const all_sum = write_data.flatMap(values => values[5]).reduce((a, b) => a + b)
  const all_dsum = write_data.flatMap(values => values[6]).reduce((a, b) => a + b)
  const all_month = write_data.flatMap(values => values[7]).reduce((a, b) => a + b)

  //経費集計用にフォーマット貼り付け
  const excheck = mainData_('ec')//経費集計用スプレッドシート

  try { excheck.insertSheet(sheetname) }
  catch (_a) { }
  const new_sheet = excheck.getSheetByName(sheetname)
  const origin_sheet = excheck.getSheetByName('原本')

  let lastRow = 3
  new_sheet.getRange(lastRow, 6, write_data.length, 3).setNumberFormat('[$¥-411]#,##0')
  new_sheet.getRange(lastRow, 1, write_data.length, write_data[0].length).setValues(write_data).setBorder(false, true, true, true, true, true)
  lastRow += write_data.length
  const set_range = new_sheet.getRange(lastRow, 1, 2, 8)
  origin_sheet.getRange('A1:H2').copyTo(new_sheet.getRange('A1:H2'))
  origin_sheet.getRange('A4:H5').copyTo(set_range)
  new_sheet.getRange(lastRow + 1, 6, 1, 3).setValues([[all_sum, all_dsum, all_month]])

  const url = `${excheck.getUrl()}#gid=${new_sheet.getSheetId()}`
  let body = `${url}\n\n${e}経費計算結果を表示します。`
  try {
    LINEWORKS.sendMsgRoom(setOptions_(), '130629262', body)
  }
  catch (_b) {
    let addbody = '経費申請用グループへの投稿にエラーが発生しました。'
    addbody += '\n\n送信予定の本文は以下の通りです。\n'
    body = `${addbody}―――――――――――――――――――――――――――${body}`
    LINEWORKS.sendMsg(setOptions_(), accountId_('山崎達也'), body)
  }
}
