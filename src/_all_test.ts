const difftes = () => {
  const log = new LogclockCheck()
  console.log(log.data())
}
const tes = () => {
  return new MonthSheet().serials()
}
class AssignSheet {
  sheet: GoogleAppsScript.Spreadsheet.Sheet
  constructor(times = new Times()) {
    this.sheet = mainData_('as').getSheetByName(times.yyyyMM())
  }
  setValue(range: string, value: any) {
    this.sheet.getRange(range).setValue(value)
    console.log(`${this.sheet.getName()}:${range}へ${value}を書き込みました。`)
  }
  setValues(range: string, values: any[][]) {
    this.sheet.getRange(range).setValues(values)
    console.log(`${this.sheet.getName()}:${range}へ${values}を書き込みました。`)
  }
  setRangeValue(range: string[], value: any) {
    this.sheet.getRangeList(range).setValue(value)
    console.log(`${this.sheet.getName()}:${range}へ${value}を書き込みました。`)
  }
}
class AssignData extends AssignSheet {
  rows: () => string[]
  constructor(times = undefined) {
    times ? super() : super(times)
    const data = this.sheet.getDataRange().getValues()
    let start
    const label = data.filter((values, index) => {
      if (values.includes('日程')) {
        start = index
        return true
      }
    }).flat()
    const day = label.indexOf('日程')
    const venue = label.indexOf('会場\n名称')
    data.forEach((values, index) => {
      if (index > start && (values[day] != '' || values[venue])) {
        this[index + 1] = new ArrayValues(values, label)
      }
    })
    this.rows = () => {
      return Object.keys(this).filter(key => key.match(/^\d*$/) != null)
    }
  }
  getCell(serial, colname) {
    for (let row of this.rows()) {
      if (this[row][0] == serial) {
        const label = JSON.parse(properties('sheet_label'))
        return `${NumToA1(label.indexOf(colname) + 1)}${row}`
      }
    }
  }
  getCells(serial, colStart, colEnd) {
    for (let row of this.rows()) {
      if (this[row][0] == serial) {
        const label = JSON.parse(properties('sheet_label'))
        return `${NumToA1(label.indexOf(colStart) + 1)}${row}:${NumToA1(label.indexOf(colEnd) + 1)}${row}`
      }
    }

  }
}
class ArrayValues {
  values: any[]
  constructor(values, label) {
    const _label = JSON.parse(properties('sheet_label'))
    const start = new Times(values[label.indexOf('開始')])
    const finish = new Times(values[label.indexOf('終了')])
    const address = new AddressWork(values[label.indexOf('会場\n住所')])
    const flag = (values[label.indexOf('講師')] == 'エムジー')
    this.values = _label.map(key => {
      const value = values[label.indexOf(key)]
      switch (key) {
        case '日程':
        case '更新日':
        case '確認日':
          return new Times(value).MMdd()
        case '開始': return start.Hmm()
        case '終了': return finish.Hmm()
        case '集合': return start.meetingTime(flag)
        case '解散': return finish.endTime()
        case 'LOG住所': return address.onlyAddress()
        case 'LOG建物': return address.building()
        case 'LOG主催者TEL': return new TextNumbers(values[label.indexOf('主催者TEL')]).onlyPhoneNumber()
        case 'LOG会場TEL': return new TextNumbers(values[label.indexOf('会場TEL')]).onlyPhoneNumber()
        default: return value
      }
    })
  }
}