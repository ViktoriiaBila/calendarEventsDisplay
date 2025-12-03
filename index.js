const VALIRIIA = 'Валерія';
const students = ['Поліна', 'Валерія ПТ', 'Валерія ЧТ', 'Валерія ВТ', 'Максим', 'Нікіта', 'Сергій', VALIRIIA];
const dayColumn = 1;
const timeColumn = 2;
const nameColumn = 3;
const costColumn = 4;

const getStartDate = (month) => {
  const result = new Date();
  result.setMonth(month, 1);
  result.setHours(0);

  return result;
};

const getEndDate = (month) => {
  const result = new Date();
  if(result.getMonth() !== month) {
    result.setMonth(month+1, 0);
    result.setHours(23);
  }

  return result;
};

const setSumToElementOfTable = (table, row, column) => {
  table.getRange(row,column).setValue(`=SUM(D3:D${row-1})`);
};

const getCalendarData = () => {
  const cal = CalendarApp.getCalendarById('vika.bila97@gmail.com');
  const table = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const tableMonth = table.getRange(1,1).getValue()-1;

  const startDate = getStartDate(tableMonth);
  const endDate = getEndDate(tableMonth);

  const events = cal.getEvents(startDate, endDate);

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

    const date = events[i].getStartTime();
    const day = date.getDate();
    table.getRange(currentRow,dayColumn).setValue(day);

    const hours = date.getHours();
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