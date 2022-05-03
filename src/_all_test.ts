const testestes = () => {
  const date = new Date();
  const shift = new StaffWorkRecord(new AssignObject());
  const staffs = Object.keys(staffObject_());
  const days = Object.keys(shift[staffs[0]]);
  console.log(days);
};
class ShiftObject {
  constructor(date = new Date()) {
    const nh = mainData_('nh');
    const sheet = nh.getSheetByName('現在シフト');
    let data = sheet.getDataRange().getValues();
    const label = data.splice(0, 1).flat();
    const month = date.getMonth() + 1;
    data = data.filter(values => values[1] == month + '月');
    console.log(data);
  }
}
