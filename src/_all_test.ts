const testestes = () => {
  const date = new Date()
  const a_obj = new AssignObject(mainData_('as').getSheetByName(dateString(date, 'yyyyMM')))
  const s_obj = new StaffWorkRecord(a_obj, mainData_('nh').getSheetByName('現在シフト'), date)
  for (let staff in s_obj) {
    for (let day in s_obj[staff]) {
      if (day == 'holiday') { break }
      console.log(`staff:${staff}\nday:${day}`)
      console.log(s_obj[staff][day])
    }
  }
}
