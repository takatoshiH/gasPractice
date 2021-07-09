function myFunction() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let targetDay = new Date;
  let dayStartOfMonth = new Date(targetDay.getFullYear(), targetDay.getMonth() - 1, 1);
  let dayEndOfMonth = new Date(targetDay.getFullYear(), targetDay.getMonth() , 0);

  writeMetaData(activeSheet, dayStartOfMonth, dayEndOfMonth);

 let numberOfWorkingDays = 0;
  
  while(dayStartOfMonth <= dayEndOfMonth) {
    const offset = 6;
    let record = dayStartOfMonth.getDate();
    activeSheet.getRange("A" + (record + offset)).setValue((dayStartOfMonth.getMonth() + 1) + '月' + dayStartOfMonth.getDate() + '日');
    activeSheet.getRange("B" + (record + offset)).setValue(giveJapaneseWeekName(dayStartOfMonth));

    if(!isHoliday(dayStartOfMonth)) {
      writeDailyWorkingTimes(activeSheet, dayStartOfMonth, offset);
      numberOfWorkingDays++;
    }

    dayStartOfMonth.setDate(dayStartOfMonth.getDate() + 1);
  }

  activeSheet.getRange("I2").setValue(numberOfWorkingDays + '日');
  activeSheet.getRange("I3").setValue(numberOfWorkingDays * 8 + '時間');

  sendMonthlyInformationMail();
}

function isHoliday(targetDay) {
  return CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com").getEventsForDay(targetDay).length > 0　|| targetDay.getDay() == 0 || targetDay.getDay() == 6;
}

function giveJapaneseWeekName(targetDay) {
  const weeks = ['日', '月', '火', '水', '木', '金', '土'];
  return weeks[targetDay.getDay()];
}

function writeMetaData(activeSheet, dayStartOfMonth, dayEndOfMonth) {
  activeSheet.getRange("D3").setValue((dayStartOfMonth.getMonth() + 1) + '月' + dayStartOfMonth.getDate() + '日');
  activeSheet.getRange("G3").setValue((dayEndOfMonth.getMonth() +1) + '月' + dayEndOfMonth.getDate() + '日');
  activeSheet.getRange("B2").setValue('テスト太郎');
  activeSheet.getRange("D1").setValue('令和3年' + (dayStartOfMonth.getMonth() + 1) + '月');
}


function writeDailyWorkingTimes(activeSheet, targetDay, offset) {
      let record = targetDay.getDate() + offset;
      activeSheet.getRange("C" + record).setValue('10:00');
      activeSheet.getRange("D" + record).setValue('19:00');
      activeSheet.getRange("E" + record).setValue('1:00');
      activeSheet.getRange("F" + record).setValue('8:00');
}

function sendMonthlyInformationMail() {
  let title = 'タイトル';
  let sentence = 'お疲れ様です。勤怠管理表を提出します。' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
  GmailApp.sendEmail('test@sample.com', title, sentence , {
    from:'test@sample.com',
    name:'てすとたろう',
    });
}
