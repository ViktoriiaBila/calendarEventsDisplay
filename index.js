const id = 'vika.bila97@gmail.com';
const VALIRIIA = 'Валерія';
const students = ['Поліна', 'Валерія ПТ', 'Валерія ЧТ', 'Валерія ВТ', 'Максим', 'Нікіта', 'Сергій', VALIRIIA];
const titlesArray = ['День', 'Час', 'Учень', 'Оплата'];
const titles = {day: 'День', time: 'Час', name: 'Учень', cost: 'Оплата'};
const columns = {day: 1, time: 2, name: 3, cost: 4};
const cost = 500;
const titleColor = '#b7e1cd';
const cellColor = '#cccccc';
const columnWidths = [100, 80, 150, 100];

const main = () => {
  const table = getTable();
  const events = getCalendarData(id, table);
  
  table.clear();
  for(let i = 0; i < columnWidths.length; i++) {
    table.setColumnWidth(i+1, columnWidths[i]);
  }

  fillRow(table, 1, titles);
  for(let title in titles) {
    formatCell(table, 1, columns[title], titleColor, true, true);
  }

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

    fillRow(table, currentRow, {day, time, name, cost});
    formatRow(table, currentRow, Object.values(columns), cellColor, true);
  }

  setSumToCell(table, lastRow, columns.cost);
};

const getTable = () => {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
};

const getCalendarData = (id, table) => {
  const tableMonth = Number(table.getName())-1;
  const startDate = getStartDate(tableMonth);
  const endDate = getEndDate(tableMonth);
  const calendar = CalendarApp.getCalendarById(id);
  const events = calendar.getEvents(startDate, endDate);

  return events;
};

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

const fillRow = (table, row, values) => {
  for(let c in columns) {
    table.getRange(row, columns[c]).setValue(values[c]);
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

/*
columns - array of columns numbers - [Number]
*/
const formatRow = (table, row, columns, color, centerFlag, boldFlag = false) => {
  columns.forEach((column) => formatCell(table, row, column, color, centerFlag, boldFlag));
};

const setSumToCell = (table, row, column) => {
  table.getRange(row,column).setValue(`=SUM(D2:D${row-1})`);
  formatCell(table, row, column, titleColor, true, true);
};