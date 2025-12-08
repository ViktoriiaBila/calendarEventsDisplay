const VALIRIIA = 'Валерія';
const students = ['Поліна', 'Валерія ПТ', 'Валерія ЧТ', 'Валерія ВТ', 'Максим', 'Нікіта', 'Сергій', VALIRIIA];
const titles = ['День', 'Час', 'Учень', 'Оплата'];
const columns = {day: 1, time: 2, name: 3, cost: 4};
const cost = 500;
const titleColor = '#b7e1cd';
const cellColor = '#cccccc';
const columnWidths = [100, 80, 150, 100];

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

// need to unite setTitles and fillCells
const setTitles = (table) => {
  for(let i = 0; i < titles.length; i++) {
    table.getRange(1, i+1).setValue(titles[i]);
    formatCell(table, 1, i+1, titleColor, true, true);
  }
};

const fillCells = (table, row, values) => {
  for(let c in columns) {
    table.getRange(row, columns[c]).setValue(values[c]);
    formatCell(table, row, columns[c], cellColor, true);
  }
};

const formatCell = (table, row, column, color, centerFlag, boldFlag = false) => {
  const cell = table.getRange(row, column);

  cell.setBackground(color);

  if(centerFlag) {
    cell.setHorizontalAlignment(CardService.HorizontalAlignment.CENTER);
  }
  
  if(boldFlag) {
    cell.setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build());
  }
};

const setSumToElementOfTable = (table, row, column) => {
  table.getRange(row,column).setValue(`=SUM(D2:D${row-1})`);
  formatCell(table, row, column, titleColor, true, true);
};

const getCalendarData = () => {
  const cal = CalendarApp.getCalendarById('vika.bila97@gmail.com');
  const table = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  table.clear();
  for(let i = 0; i < columnWidths.length; i++) {
    table.setColumnWidth(i+1, columnWidths[i]);
  }

  const tableMonth = Number(table.getName())-1;

  const startDate = getStartDate(tableMonth);
  const endDate = getEndDate(tableMonth);

  const events = cal.getEvents(startDate, endDate);

  setTitles(table);

  let currentRow = 1;
  let lastRow = 2;
  for(let i = 0; i < events.length; i++) {
    let name = events[i].getTitle();
    if(!students.includes(name)) {
      continue;
    }
    
    if(name.length > 7 && name.slice(0, 7) === VALIRIIA) {
      name = VALIRIIA;
    }

    const date = events[i].getStartTime();
    const day = date.getDate();
    const hours = date.getHours();
    let minutes = date.getMinutes();
    if (minutes === 0) {
      minutes = '00';
    }
    const time = `${hours}:${minutes}`;

    currentRow++;
    lastRow++;

    fillCells(table, currentRow, {day, time, name, cost});
  }

  setSumToElementOfTable(table, lastRow, columns.cost);
};