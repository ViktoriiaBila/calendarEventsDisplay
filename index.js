const id = 'vika.bila97@gmail.com';
const VALIRIIA = 'Валерія';
const students = ['Поліна', 'Валерія ПТ', 'Валерія ЧТ', 'Валерія ВТ', 'Максим', 'Нікіта', 'Сергій', VALIRIIA];
const columns = {
  day: {
    title: 'День',
    number: 1,
    alfabetCharacter: 'A'
  },
  time: {
    title: 'Час',
    number: 2,
    alfabetCharacter: 'B'
  },
  name: {
    title: 'Учень',
    number: 3,
    alfabetCharacter: 'C'
  },
  cost: {
    title: 'Оплата',
    number: 4,
    alfabetCharacter: 'D'
  }
};
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

  fillRowWithTitles(table);
  formatRowWithTitles(table);

  let currentRow = 1;
  let lastRow = 2;
  let rowCount = 1;
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

    fillRow(table, currentRow, {time, name, cost});
    formatRow(table, currentRow, cellColor, true);

    let nextDay = 0;
    if(i !== events.length - 1) {
      nextDay = events[i+1].getStartTime().getDate();
    }
    
    if(day === nextDay) {
      rowCount++;
    } else if(rowCount === 1) {
      fillCell(table, currentRow, columns.day.number, day);
    } else {
      table
        .getRange(`${columns.day.alfabetCharacter}${currentRow+1-rowCount}:${columns.day.alfabetCharacter}${currentRow}`)
        .merge()
        .setValue(day);
      
      rowCount = 1;
    }

  }

  setSumToCell(table, lastRow, columns.cost.number);
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

const fillCell = (table, row, column, value) => {
  table.getRange(row, column).setValue(value);
};

const fillRow = (table, row, values) => {
  for(let value in values) {
    fillCell(table, row, columns[value].number, values[value]);
  }
};

const fillRowWithTitles = (table) => {
  Object.values(columns).forEach((element) =>
    fillCell(table, 1, element.number, element.title));
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

const formatRow = (table, row, color, centerFlag, boldFlag = false) => {
  Object.values(columns).forEach((element) => 
    formatCell(table, row, element.number, color, centerFlag, boldFlag));
};

const formatRowWithTitles = (table) => {
  Object.values(columns).forEach((element) =>
    formatCell(table, 1, element.number, titleColor, true, true));
};

const setSumToCell = (table, row, column) => {
  table.getRange(row,column).setValue(`=SUM(D2:D${row-1})`);
  formatCell(table, row, column, titleColor, true, true);
};