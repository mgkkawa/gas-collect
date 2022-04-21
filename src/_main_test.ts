const shiftCheck = (date) => {
  // const shift_sheet = mainData_('sh');
  if (!date) { date = new Date() };
  const assign_sheet = mainData_('as');
  const target_sheet = assign_sheet.getSheetByName(valueDate(date, 'yyyyMM'));
  const sheet_data = target_sheet.getDataRange().getValues();
  const sheet_label = sheet_data.filter(values => values.includes('日程')).flat();
  const target_data = sheet_data.filter(values =>
    Object.prototype.toString.call(values[sheet_label.indexOf('日程')]) == "[object Date]"
    && values[sheet_label.indexOf('開催\n可否')] != '中止');



}
const castingTest = () => {
  const vc = testData_('vc');
  const sh = testData_('sh');


}
const testEcho = () => {
  console.log('consoleテスト成功!!');
  Logger.log('Loggerテスト成功!!');
  Browser.msgBox('Browser.msgBoxテスト成功!!');
}
const propertySet = () => {
  const scripts = PropertiesService.getScriptProperties();
}

const propertieCheck = () => {
  const prop = PropertiesService.getScriptProperties();
  const keys = prop.getKeys();
}

const propertieDeliete = () => {
  const prop = PropertiesService.getScriptProperties();
  const keys = prop.getKeys();
}