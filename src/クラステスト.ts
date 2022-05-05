class WritingSheets {
  vencallspread: GoogleAppsScript.Spreadsheet.Spreadsheet
  assign: GoogleAppsScript.Spreadsheet.Sheet
  vencall: any
  shift: GoogleAppsScript.Spreadsheet.Sheet
  constructor(times = new Times()) {
    this.vencallspread = mainData_('vc')
    this.assign = mainData_('as').getSheetByName(times.yyyyMM())
    this.vencall = this.vencallspread.getSheetByName(times.yyyyMM())
    this.shift = mainData_('sh').getSheetByName(times.yyyy_MM())
  }
}
class Logclock extends WritingSheets {
  sheet: any
  label: any
  constructor(times = new Times()) {
    super(times)
    const sheet = this.vencallspread.getSheetByName('LOGCLOCK')

    this.sheet = sheet
  }
  fieldCheck(row) {
    new Promise(() => {
      row = row.map(value => `${NumToA1(this.label.indexOf('Check1') + 1)}${value}`)
      this.sheet.getRangeList(row).insertCheckBoxes().check()
    }).catch(() => {
      this.sheet.getRange(`${NumToA1(this.label.indexOf('Check1') + 1)}${row}`).insertCheckBoxes().check()
    })
    return this
  }
  workCheck(row) {
    new Promise(() => {
      row = row.map(value => `${NumToA1(this.label.indexOf('Check1') + 1)}${value}`)
      this.sheet.getRangeList(row).insertCheckBoxes().check()
    }).catch(() => {
      this.sheet.getRange(`${NumToA1(this.label.indexOf('Check1') + 1)}${row}`).insertCheckBoxes().check()
    })
    return this
  }
  castingCheck(row) {
    new Promise(() => {
      row = row.map(value => `${NumToA1(this.label.indexOf('Check1') + 1)}${value}`)
      this.sheet.getRangeList(row).insertCheckBoxes().check()
    }).catch(() => {
      this.sheet.getRange(`${NumToA1(this.label.indexOf('Check1') + 1)}${row}`).insertCheckBoxes().check()
    })
    return this
  }
  castingUncheck(row) {
    new Promise(() => {
      row = row.map(value => `${NumToA1(this.label.indexOf('Check1') + 1)}${value}`)
      this.sheet.getRangeList(row).insertCheckBoxes().uncheck()
    }).catch(() => {
      this.sheet.getRange(`${NumToA1(this.label.indexOf('Check1') + 1)}${row}`).insertCheckBoxes().uncheck()
    })
    return this
  }
  allCheck(row) {
    new Promise(() => {
      row = row.flatMap(value => [`${NumToA1(this.label.indexOf('Check1') + 1)}${value}:${NumToA1(this.label.indexOf('Check2') + 1)}${value}`,
      `${NumToA1(this.label.indexOf('Check3') + 1)}${value}`])
      this.sheet.getRangeList(row).insertCheckBoxes().check()
    }).catch(() => {
      this.sheet.getRangeList([`${NumToA1(this.label.indexOf('Check1') + 1)}${row}:${NumToA1(this.label.indexOf('Check2') + 1)}${row}`,
      `${NumToA1(this.label.indexOf('Check3') + 1)}${row}`]).insertCheckBoxes().check()
    })
    return this
  }
}
class LogclockCheck extends Logclock {
  constructor(times = new Times()) {
    super(times)
    let data = this.sheet.getDataRange().getValues()
    const label = labelCreate(data)
    const checks = ['現場チェック', 'シフトチェック', 'お仕事チェック', 'キャスティングチェック'].map(key => label.indexOf(key))
    data.forEach((values, index) => {
      if (index < 1) { return }
      if (checks.some(key => values[key])) {
        const obj = {}
        obj['checks'] = new LogChecks(values, label)
        obj['log'] = new LogData(values, label)
        this[index + 1] = obj
      }
    })
  }
}
class LogChecks {
  day: string
  venue: any
  serial: string
  constructor(values, label) {
    this.day = new Times(values[label.indexOf('日程')]).MMdd()
    this.venue = values[label.indexOf('会場\n名称')]
    this.serial = String(values[label.indexOf('開催No.')])
  }
}
class LogData {
  spot: any
  shift: any
  schedule: any
  casting: any
  constructor(values, label) {
    const flagCheck = (a, b) => {
      if (a) {
        return a
      } else {
        return b
      }
    }
    const start = label.indexOf('メイン\n講師')
    const finish = label.indexOf('サポート5')
    const member = values.filter((value, index) => index >= start && index <= finish)
    this.spot = flagCheck(values[label.indexOf('現場チェック')], values[label.indexOf('現場登録')])
    this.shift = flagCheck(values[label.indexOf('シフトチェック')], values[label.indexOf('シフト登録')])
    this.schedule = flagCheck(values[label.indexOf('お仕事チェック')], values[label.indexOf('お仕事スケジュール')])
    this.casting = flagCheck(values[label.indexOf('キャスティングチェック')], values[label.indexOf('キャスティング')])

  }
}
class Times {
  value: Date
  fullYear: number
  month: number
  date: number
  hours: number
  minutes: number
  constructor(value = new Date()) {
    if (Object.prototype.toString.call(value) != '[object Date]') {
      value = new Date(value)
    }
    this.value = value
    this.fullYear = value.getFullYear()
    this.month = value.getMonth()
    this.date = value.getDate()
    this.hours = value.getHours()
    this.minutes = value.getMinutes()
    // return value
  }
  /**@return yyyyMM形式の文字列*/
  yyyyMM() { return `${this.fullYear}${String(this.month + 1).padStart(2, '0')}` }
  yyyy_MM() { return `${this.fullYear}.${String(this.month + 1).padStart(2, '0')}` }
  /**@return MM/dd形式の文字列 */
  MMdd() { return `${String(this.month + 1).padStart(2, '0')}/${String(this.date).padStart(2, '0')}` }
  /**@returns H:mm形式の文字列 */
  Hmm() { return this.hours + String(this.minutes).padStart(2, '0') }
  /** @returns 引数に与えたbool値に応じて集合時間を返す。true:(-90) false:(-60)*/
  meetingTime(bool) {
    if (bool) {
      this.value.setMinutes(this.minutes - 90)
      return Utilities.formatDate(this.value, 'JST', 'H:mm')
    }
    this.value.setMinutes(this.minutes - 60)
    return Utilities.formatDate(this.value, 'JST', 'H:mm')
  }
  endTime() {
    this.value.setMinutes(this.minutes + 60)
    return Utilities.formatDate(this.value, 'JST', 'H:mm')
  }

}