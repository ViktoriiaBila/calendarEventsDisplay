const students = ['Поліна', 'Валерія ПТ', 'Валерія ЧТ', 'Валерія ВТ', 'Максим', 'Нікіта'];

function getCalendarData() {
  let cal = CalendarApp.getCalendarById('vika.bila97@gmail.com');
  let table = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let today = new Date();
  let startDate = table.getRange(1,2).getValue();
  let endDate = table.getRange(1,3).getValue();

  let events = cal.getEvents(startDate, endDate);

  // let dateTest = events[1].getStartTime();
  // table.getRange(3,1).setValue(dateTest);
  // let titleTest = events[1].getTitle();
  // table.getRange(3,2).setValue(titleTest);
  // let colorTest = events[1].getColor();
  // console.log(colorTest);
  // table.getRange(3,2).setBackground(colorTest);

  let j=3;
  for(let i=0; i<events.length; i++) {
    let name = events[i].getTitle();
    if(!students.includes(name)) {
      continue;
    }

    let date = events[i].getStartTime();
    let day = date.getDate();
    let hours = date.getHours();
    let minutes = date.getMinutes();
    if (minutes === 0) {
      minutes = '00';
    }

    if(day === today.getDate()+1) {
      last = table.getRange(j,4);
      last.setValue(`=SUM(D3:D${j-1})`);
      return;
    }
    
    table.getRange(i+3,1).setValue(day);
    table.getRange(i+3,2).setValue(`${hours}:${minutes}`);
    table.getRange(i+3,3).setValue(name);

    if(name === 'Максим') {
      table.getRange(i+3,4).setValue(500);
    }

    j++;
  }
}