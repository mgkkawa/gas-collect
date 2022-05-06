class WritingSheets {
  vencallspread: GoogleAppsScript.Spreadsheet.Spreadsheet
  assign: GoogleAppsScript.Spreadsheet.Sheet
  vencall: GoogleAppsScript.Spreadsheet.Sheet
  staffshift: GoogleAppsScript.Spreadsheet.Sheet
  returnObj: () => {}
  keys: () => string[]
  getMap: () => any[][]
  constructor(times = new Times()) {
    this.vencallspread = mainData_('vc')
    this.assign = mainData_('as').getSheetByName(times.yyyyMM())
    this.vencall = this.vencallspread.getSheetByName(times.yyyyMM())
    this.staffshift = mainData_('sh').getSheetByName(times.yyyy_MM())
    this.returnObj = () => {
      const obj = {}
      for (let row in this) {
        if (String(row).match(/^\d*$/) != null) {
          obj[String(row)] = this[row]
        }
      }
      return obj
    }
    this.keys = () => {
      return Object.keys(this).filter(key => String(key).match(/^\d*$/) != null)
    }
    this.getMap = () => {
      return this.keys().map(key => this[key])
    }
  }
}
class MonthSheet extends WritingSheets {
  label: string[]
  sheet: GoogleAppsScript.Spreadsheet.Sheet
  serials: () => any[]
  fieldCheck: (row: any) => this
  shiftCheck: (row: any) => this
  workCheck: (row: any) => this
  castingCheck: (row: any) => this
  castingUncheck: (row: any) => this
  allCheck: (row: any) => this

  constructor(date = undefined) {
    super()
    const possheet = this.vencallspread.getSheetByName('転記')
    const data = possheet.getDataRange().getValues()
    let start
    data.forEach((values, index) => {
      if (values.includes('日程')) {
        start ??= index
        this.label ??= values
      }
      if (index > start && (values[this.label.indexOf('日程')] != '' || values[this.label.indexOf('会場\n名称')])) {
        this[index + 1] = trimValues_(values, this.label)
      }
    })
    this.serials = () => {
      return this.keys().map(key => this[key][this.label.indexOf('開催No.')])
    }
    this.fieldCheck = (row) => {
      new Promise(() => {
        const range = row.map(value => `${NumToA1(this.label.indexOf('現場チェック') + 1)}${value}`)
        this.sheet.getRangeList(range).insertCheckboxes().check()
      }).catch(() => {
        this.sheet.getRange(`${NumToA1(this.label.indexOf('現場チェック') + 1)}${row}`).insertCheckboxes().check()
      })
      return this
    }
    this.shiftCheck = (row) => {
      new Promise(() => {
        const range = row.map(value => `${NumToA1(this.label.indexOf('シフトチェック') + 1)}${value}`)
        this.sheet.getRangeList(range).insertCheckboxes().check()
      }).catch(() => {
        this.sheet.getRange(`${NumToA1(this.label.indexOf('シフトチェック') + 1)}${row}`).insertCheckboxes().check()
      })
      return this
    }
    this.workCheck = (row) => {
      new Promise(() => {
        const range = row.map(value => `${NumToA1(this.label.indexOf('お仕事チェック') + 1)}${value}`)
        this.sheet.getRangeList(range).insertCheckboxes().check()
      }).catch(() => {
        this.sheet.getRange(`${NumToA1(this.label.indexOf('お仕事チェック') + 1)}${row}`).insertCheckboxes().check()
      })
      return this
    }
    this.castingCheck = (row) => {
      new Promise(() => {
        const range = row.map(value => `${NumToA1(this.label.indexOf('キャスティングチェック') + 1)}${value}`)
        this.sheet.getRangeList(range).insertCheckboxes().check()
      }).catch(() => {
        this.sheet.getRange(`${NumToA1(this.label.indexOf('キャスティングチェック') + 1)}${row}`).insertCheckboxes().check()
      })
      return this
    }
    this.castingUncheck = (row) => {
      new Promise(() => {
        const range = row.map(value => `${NumToA1(this.label.indexOf('キャスティングチェック') + 1)}${value}`)
        this.sheet.getRangeList(range).insertCheckboxes().uncheck()
        this.sheet.getRangeList(row).clearContent()
      }).catch(() => {
        this.sheet.getRange(`${NumToA1(this.label.indexOf('キャスティングチェック') + 1)}${row}`).insertCheckboxes().uncheck()

      })
      return this
    }
    this.allCheck = (row) => {
      new Promise(() => {
        const range = row.flatMap(value => [`${NumToA1(this.label.indexOf('現場チェック') + 1)}${value}:${NumToA1(this.label.indexOf('お仕事チェック') + 1)}${value}`,
        `${NumToA1(this.label.indexOf('キャスティングチェック') + 1)}${value}`])
        this.sheet.getRangeList(range).insertCheckboxes().check()
      }).catch(() => {
        this.sheet.getRangeList([`${NumToA1(this.label.indexOf('現場チェック') + 1)}${row}:${NumToA1(this.label.indexOf('お仕事チェック') + 1)}${row}`,
        `${NumToA1(this.label.indexOf('キャスティングチェック') + 1)}${row}`]).insertCheckboxes().check()
      })
      return this
    }
  }
}
class Logclock extends WritingSheets {
  sheet: GoogleAppsScript.Spreadsheet.Sheet
  label: string[]
  constructor(times = new Times()) {
    super(times)
    const sheet = this.vencallspread.getSheetByName('LOGCLOCK')
    this.sheet = sheet
  }

}
class LogclockCheck extends Logclock {
  constructor(times = new Times()) {
    super(times)
    let data = this.sheet.getDataRange().getValues()
    this.label = JSON.parse(properties('LOG_label'))
    const checks = ['現場チェック', 'シフトチェック', 'お仕事チェック', 'キャスティングチェック'].map(key => this.label.indexOf(key))
    data.forEach((values, index) => {
      if (index < 1) { return }
      if (checks.some(key => values[key])) {
        const obj =
        {
          checks: {},
          logs: {}
        }
        obj.checks = new LogChecks(values, this.label)
        obj.logs = new LogData(values, this.label)
        this[index + 1] = obj
        console.log(this[index + 1])
      }
    })
  }
  data() {
    const obj = {}
    for (let key of this.keys()) {
      obj[key] = this[key]
    }
    return obj
  }
}
class LogChecks {
  day: string
  venue: string
  serial: string
  constructor(values, label) {
    this.day = new Times(values[label.indexOf('日程')]).MMdd()
    this.venue = values[label.indexOf('会場\n名称')]
    this.serial = String(values[label.indexOf('開催No.')])
  }
}
class LogData extends LogclockCheck {
  spot: boolean
  shift: boolean
  schedule: boolean
  casting: boolean
  member: string[]
  check: () => boolean[]
  constructor(values, label) {
    super()
    const flagCheck = (a, b) => {
      if (Boolean(a)) {
        return Boolean(a)
      } else {
        return Boolean(b)
      }
    }
    const start = label.indexOf('メイン\n講師')
    const finish = label.indexOf('サポート5')
    const member = values.filter((value, index) => index >= start && index <= finish)
    this.spot = flagCheck(values[label.indexOf('現場チェック')], values[label.indexOf('現場登録')])
    this.shift = flagCheck(values[label.indexOf('シフトチェック')], values[label.indexOf('シフト登録')])
    this.schedule = flagCheck(values[label.indexOf('お仕事チェック')], values[label.indexOf('お仕事スケジュール')])
    this.casting = flagCheck(values[label.indexOf('キャスティングチェック')], values[label.indexOf('キャスティング')])
    this.member = member
    this.check = () => {
      return [this.spot, this.shift, this.schedule, this.casting]
    }
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
  Hmm() { return `${this.hours}:${String(this.minutes).padStart(2, '0')}` }
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
  nextDay() {
    this.value.setDate(this.value.getDate() + 1)
    return this.value
  }
}