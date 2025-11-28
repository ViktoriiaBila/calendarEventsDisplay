const students = ['Поліна', 'Валерія ПТ', 'Валерія ЧТ', 'Валерія ВТ', 'Максим', 'Нікіта'];
const VALIRIIA = 'Валерія';
const MAXIM = 'Максим';
const dayColumn = 1;
const timeColumn = 2;
const nameColumn = 3;
const costColumn = 4;

const getEndDate = (month) => {
  let today = new Date();
  let currentMonth = today.getMonth();

  if(currentMonth === month) {
    return today;
  } else {
    // need to set the last date of month to endDate
    new Data(today.getFullYear(), month, );
  }
};

const setSumToElementOfTable = (table, row, column) => {
  table.getRange(row,column).setValue(`=SUM(D3:D${row-1})`);
};

const getCalendarData = () => {
  let cal = CalendarApp.getCalendarById('vika.bila97@gmail.com');
  let table = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let startDate = table.getRange(1,2).getValue();
  let tableMonth = startDate.getMonth();
  let endDate = getCalendarData(tableMonth);

  let events = cal.getEvents(startDate, endDate);

  // let dateTest = events[1].getStartTime();
  // table.getRange(3,1).setValue(dateTest);
  // let titleTest = events[1].getTitle();
  // table.getRange(3,2).setValue(titleTest);
  // let colorTest = events[1].getColor();
  // console.log(colorTest);
  // table.getRange(3,2).setBackground(colorTest);
  
  let currentRow = 2;
  let lastRow = 4;
  for(let i = 0; i<events.length; i++) {
    let name = events[i].getTitle();
    if(!students.includes(name)) {
      continue;
    }
    currentRow++;

    if(name.length > 7 && name.slice(0, 7) === VALIRIIA) {
      name = VALIRIIA;
    }
    table.getRange(currentRow,nameColumn).setValue(name);

    let date = events[i].getStartTime();
    let day = date.getDate();
    table.getRange(currentRow,dayColumn).setValue(day);

    let hours = date.getHours();
    let minutes = date.getMinutes();
    if (minutes === 0) {
      minutes = '00';
    }
    table.getRange(currentRow,timeColumn).setValue(`${hours}:${minutes}`);
    
    table.getRange(currentRow,costColumn).setValue(500);

    lastRow++;
  }

  setSumToElementOfTable(table, lastRow, costColumn);
};