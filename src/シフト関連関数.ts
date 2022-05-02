function addShiftData() {
  new AddShift().addDay()
}
// shiftObjectCreate()→shiftObjectAddValue()→shiftObjectAddInfo()
// 翌月希望休申請フォームをオブジェクト化
// 希望休や有休、リフレの取得と出勤可否フラグの設定。
const shiftObjectCreate = (date = new Date()) => {
  if (!date) { date = new Date() }
  const yyyy = date.getFullYear()
  const M = date.getMonth() + 1
  const check_month = String(date.getMonth() + 1) + '月'
  date.setMonth(date.getMonth() + 1)
  date.setDate(0)
  const check_date = date.getDate()
  const spreadsheet = mainData_('nh')
  const sheet = spreadsheet.getSheetByName('フォームの回答 1')
  const sheet_data = sheet.getDataRange().getValues().filter(values => values[2] == check_month)
    .map(values => [1, 3, 4, 5, 6].map(index => {
      switch (true) {
        case index >= 3 && index <= 5:
          return values[index].replace(/日/g, '')
        default: return values[index]
      }
    }))
  const sheet_staffs = sheet_data.flatMap(values => values[0])
  const all_hopes = sheet_data.map(values => JSON.parse(`[${String(values[1])}]`))
  const all_paids = sheet_data.map(values => JSON.parse(`[${String(values[2])}]`))
  const all_refreshs = sheet_data.map(values => JSON.parse(`[${String(values[3])}]`))
  const staffs = Object.keys(staffObject_())
  const obj = {}
  for (let d = 1; d <= check_date; d++) {
    const obj_ = {}
    for (let staff of staffs) {
      const staff_obj = {}
      const index = sheet_staffs.indexOf(staff)
      const hopes = all_hopes[index]
      const paids = all_paids[index]
      const refreshs = all_refreshs[index]
      if (hopes.includes(d) || paids.includes(d) || refreshs.includes(d)) {
        staff_obj['FLAG'] = false
        switch (true) {
          case hopes.includes(d):
            staff_obj['SET'] = '希'
            break
          case paids.includes(d):
            staff_obj['SET'] = '有'
            break
          case refreshs.includes(d):
            staff_obj['SET'] = 'リ'
            break
        }
      } else {
        staff_obj['FLAG'] = true
        staff_obj['SET'] = '備'
      }
      obj_[staff] = staff_obj
    }
    obj[String(d).padStart(2, '0')] = obj_
  }
  const obj_ = shiftObjectAddValue_(date, obj)
  const database = mainData_('db').getSheetByName('シフト')
  let set_obj = JSON.stringify(obj_)
  const set_strings = new Array(Math.ceil(set_obj.length / 50000))
  if (set_strings.length > 1) {
    for (let l = 0; l < set_strings.length; l++) {
      set_strings.splice(l, 1, set_obj.slice(l * 50000, (l + 1) * 50000))
    }
  } else { set_strings[set_strings.length - 1] = set_obj }
  const db_data = database.getDataRange().getValues().filter((values, index) => index == 0 || values[0] || values[1])
  const last_row = database.getLastRow() - 1
  const last_col = database.getLastColumn()
  const db_label = db_data.splice(0, 1).flat()
  const range = []
  const set_data = db_data.map((values, index) => {
    if (values[db_label.indexOf('Origin')] && values[db_label.indexOf('yyyy')] == yyyy && values[db_label.indexOf('M')] == M) {
      return [true, , new Date(), yyyy, M].concat(set_strings)
    } else { return values }
  })
  database.getRange(2, 1, last_row, last_col).clear()
  set_data.forEach((values, index) => {
    if (values[0]) { range.push(`A${index + 2}`) }
    else if (values[1]) { range.push(`B${index + 2}`) }
    database.getRange(index + 2, 1, 1, values.length).setValues([values])
  })
  database.getRangeList(range).insertCheckboxes().check()
  shiftObjectAddInfo(date, obj_)
}
// 最新のシフトオブジェクトに対して
// ['MTG','当欠','病欠','研修']の情報を追加
const shiftObjectAddValue_ = (date = new Date(), shift = originCheck_(date)) => {
  const nh = mainData_('nh')
  const nh_sheet = nh.getSheetByName(dateString(date, 'yyyy.MM'))
  const nh_label = nh_sheet.getRange(1, 1, 1, nh_sheet.getLastColumn()).getValues().flat()
  const name_ind = nh_label.indexOf('氏名')
  const nh_data = nh_sheet.getDataRange().getValues().filter(values => values[name_ind] != '氏名' && values[name_ind] != '')
    .map(values => ['氏名', 'MTG', '病欠', '当欠', '研修'].map(key => values[nh_label.indexOf(key)]))
  nh_data.forEach(values => {
    const staff = values[0]
    const mtg = values[1].replace(/日/g, '').padStart(2, '0')
    const sick = JSON.parse(`[${values[2].replace(/日/g, '')}]`).filter(Boolean)
    const absent = JSON.parse(`[${values[3].replace(/日/g, '')}]`).filter(Boolean)
    const training = JSON.parse(`[${values[4].replace(/日/g, '')}]`).filter(Boolean)
    if (mtg.length != 0) {
      const check = shift[mtg][staff]['FLAG']
      if (check) {
        shift[mtg][staff]['FLAG'] = false
        shift[mtg][staff]['SET'] = 'M'
      } else {
        try { Browser.msgBox(`MTGチェック${staff}さんの${mtg}日シフトは\n"${shift[mtg][staff]['SET']}"です`) }
        catch { Logger.log(`MTGチェック${staff}さんの${mtg}日シフトは\n"${shift[mtg][staff]['SET']}"です`) }
      }
    }
    if (sick.length > 0) {
      sick.forEach(d => {
        const dd = String(d).padStart(2, '0')
        const check = shift[dd][staff]['FLAG']
        if (check) {
          shift[dd][staff]['FLAG'] = false
          shift[dd][staff]['SET'] = '病欠'
        } else {
          try { Browser.msgBox(`病欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
          catch { Logger.log(`病欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
        }
      })
    }
    if (absent.length > 0) {
      absent.forEach(d => {
        const dd = String(d).padStart(2, '0')
        const check = shift[dd][staff]['FLAG']
        if (check) {
          shift[dd][staff]['FLAG'] = false
          shift[dd][staff]['SET'] = '当欠'
        } else {
          try { Browser.msgBox(`当欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
          catch { Logger.log(`当欠チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
        }
      })
    }
    if (training.length > 0) {
      training.forEach(d => {
        const dd = String(d).padStart(2, '0')
        const check = shift[dd][staff]['FLAG']
        if (check) {
          shift[dd][staff]['SET'] = '研'
        } else {
          try { Browser.msgBox(`研修チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
          catch { Logger.log(`研修チェック${staff}さんの${dd}日シフトは\n"${shift[dd][staff]['SET']}"です`) }
        }
      })
    }
  })
  return shift
}
const shiftObjectAddInfo = (date = new Date(), obj = originCheck_(date)) => {
  const yyyy = date.getFullYear()
  const MM = date.getMonth() + 1
  const as = mainData_('as').getSheetByName(yyyy + String(MM).padStart(2, '0'))
  const origin = as.getDataRange().getValues()
  const as_label = origin.filter(values => values.includes('日程')).flat()
  const check = ['日程', '開催\n可否'].map(key => as_label.indexOf(key))
  const times = ['開始', '終了'].map(key => as_label.indexOf(key))
  const keys = ['会場\n名称', '集合', '解散', '通し番号', '地域']
  const indexs = keys.map(key => {
    if (key == '集合' || key == '解散') { return key }
    else { return as_label.indexOf(key) }
  })
  const keys_obj =
  {
    '会場\n名称': 'VENUE',
    '集合': 'MEETING',
    '解散': 'LEAVE',
    '通し番号': 'SET',
    '地域': 'AREA'
  }
  const as_data = origin.filter((values) =>
    Object.prototype.toString.call(values[check[0]]) == '[object Date]' &&
    values[check[1]] != '中止' && values[indexs[0]] != '')
    .map(values => values.map((value, index) => {
      switch (index) {
        case check[0]: return dateString(value, 'dd')
        default: return value
      }
    }))
  const days = as_data.flatMap(values => values[check[0]])
  const trim_days = days.filter((value, index, array) => array.indexOf(value) == index)
  const staff_obj = staffObject_()
  const staffs = Object.keys(staff_obj)
  trim_days.forEach(dd => {
    for (let staff of staffs) {
      if (obj[dd][staff]['FLAG']) {
        const dd_map = as_data.filter((values, index) =>
          index >= days.indexOf(dd) && index <= days.lastIndexOf(dd) && values.includes(staff)
        )
        const flag_check = dd_map.flat().includes('エムジー')
        let time = []
        switch (dd_map.length) {
          case 0: continue
          case 1:
            time = dd_map.flat().filter((value, index) => index == times[0] || index == times[1])
            break
          default:
            time = dd_map.flatMap(values => values.filter((value, index) => index == times[0] || index == times[1]))
        }
        const obj_ = ddObject_(dd_map, as_label, staff)
        time.sort((a, b) => a.getTime() - b.getTime())
        keys.forEach(key => {
          if (key != '集合') {
            switch (key) {
              case '開始':
                obj[dd][staff][keys_obj[key]] = dateString(time[0], 'H:mm')
                break
              case '終了':
                obj[dd][staff][keys_obj[key]] = dateString(time[time.length - 1], 'H:mm')
                break
              case '解散':
                obj[dd][staff][keys_obj[key]] = timeEnd_(time[time.length - 1])
                break
              default:
                obj[dd][staff][keys_obj[key]] = dd_map[0][as_label.indexOf(key)]
            }
          }
          else if (flag_check) {
            obj[dd][staff][keys_obj[key]] = timeStartMain_(time[0])
          } else {
            obj[dd][staff][keys_obj[key]] = timeStartSup_(time[0])
          }
        })
        obj[dd][staff]['INFO'] = obj_
      } else { continue }
    }
  })
  shiftObjectUpdate_(obj)
}
function ddObject_(array, label, staff) {
  const keys = ['開始', '終了', '講師', 'コース']
  const obj_ = {
    '開始': 'START',
    '終了': 'FINISH',
    '講師': 'MAIN_FLAG',
    'コース': 'CORSE'
  }
  const obj = {}
  array.forEach((values, index) => {
    const dd_obj = {}
    const index_ = String(index + 1).padStart(2, '0')
    keys.forEach(key => {
      switch (key) {
        case '開始': dd_obj[obj_[key]] = dateString(values[label.indexOf(key)], 'H:mm')
          break
        case '終了': dd_obj[obj_[key]] = dateString(values[label.indexOf(key)], 'H:mm')
          break
        case '講師':
          const flag_ = (values[label.indexOf(key)] == 'エムジー')
          const main = (values[label.indexOf('メイン\n講師')] == staff)
          const sup = (values[label.indexOf('メイン\n講師')] != staff)
          if (flag_ && main) { dd_obj[obj_[key]] = 'メイン' }
          else if (flag_ && sup) { dd_obj[obj_[key]] = 'サポート' }
          else { dd_obj[obj_[key]] = 'SB同行' }
          break
        default: dd_obj[obj_[key]] = values[label.indexOf(key)]
      }
    })
    obj[index_] = dd_obj
  })
  return obj
}
// 最新のOriginシフトオブジェクトを返す。
const originCheck_ = (date = new Date()) => {
  const yyyy = date.getFullYear()
  const M = date.getMonth() + 1
  const database = mainData_('db').getSheetByName('シフト')
  const data = database.getDataRange().getValues()
  const label = data.filter(values => values.includes('Origin') && values.includes('Object')).flat()
  const filter = data.filter(values => values[0] && values[label.indexOf('yyyy')] == yyyy && values[label.indexOf('M')] == M)
    .flatMap(values => values.filter((value, index) => index >= label.indexOf('Object') && value != ''))
  const to_string = filter.join('')
  return JSON.parse(to_string)
}

function infoCheck_(date = new Date()) {
  const yyyy = date.getFullYear()
  const M = date.getMonth() + 1
  const database = mainData_('db').getSheetByName('シフト')
  const data = database.getDataRange().getValues()
  const label = data.filter(values => values.includes('INFO') && values.includes('Object')).flat()
  const filter = data.filter(values => values[1] && values[label.indexOf('yyyy')] == yyyy && values[label.indexOf('M')] == M)
    .flatMap(values => values.filter((value, index) => index >= label.indexOf('Object') && value != ''))
  const to_string = filter.join('')
  return JSON.parse(to_string)
}
// シフトオブジェクトを保存する。
const shiftObjectUpdate_ = (obj, date = new Date()) => {
  const yyyy = date.getFullYear()
  const M = date.getMonth() + 1
  const database = mainData_('db').getSheetByName('シフト')
  const db_data = database.getDataRange().getValues()
  const last_row = db_data.length + 1
  const db_label = db_data.splice(0, 1).flat()
  let set_obj = JSON.stringify(obj)
  const set_strings = new Array(Math.ceil(set_obj.length / 50000))
  if (set_strings.length > 1) {
    for (let l = 0; l < set_strings.length; l++) {
      set_strings.splice(l, 1, set_obj.slice(l * 50000, (l + 1) * 50000))
    }
  } else { set_strings[set_strings.length - 1] = set_obj }
  const check_row = []
  for (let [index, data] of db_data.entries()) {
    const db_yyyy = data[db_label.indexOf('yyyy')]
    const db_M = data[db_label.indexOf('M')]
    if (db_yyyy == yyyy && db_M == M) { check_row.push(`B${(index + 2)}`) }
  }
  if (check_row.length > 0) {
    try { database.getRangeList(check_row).uncheck() }
    catch { database.getRange(check_row[0]).uncheck() }
  }
  database.getRange(last_row, db_label.indexOf('TimeStamp') + 1, 1, set_strings.length + 3)
    .setValues([[new Date(), yyyy, M].concat(set_strings)])
  database.getRange(last_row, db_label.indexOf('INFO') + 1).insertCheckboxes().check()
  return
}
const timeStartMain_ = (time) => {
  time = new Date(time)
  time.setMinutes(time.getMinutes() - 90)
  return dateString(time, 'H:mm')
}
const timeStartSup_ = (time) => {
  time = new Date(time)
  time.setMinutes(time.getMinutes() - 60)
  return dateString(time, 'H:mm')
}
const timeEnd_ = (time) => {
  time = new Date(time)
  time.setMinutes(time.getMinutes() + 60)
  return dateString(time, 'H:mm')
}
const timesousa_ = (time, minutes) => {
  const type = Object.prototype.toString.call(time) == '[object Date]'
  if (type) {
    time = new Date(time)
    return time.setMinutes(time.getMinutes() + minutes)
  } else {
    try { new Date(time) }
    catch {
      if (String(time).match(/^[\d].:[\d].$/)) {
        const now = new Date()
        now.setHours(time.match(/^[\d]./), time.match(/[\d].$/))
        now.setMinutes(now.getMinutes() + minutes)
        return now
      }
    }
    const now = new Date(time)
    now.setMinutes(now.getMinutes() + minutes)
    return now
  }
}