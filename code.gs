/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopene
 */
function onOpen() {
  const spread = SpreadsheetApp.getActive();

  // Скрываем служебные листы, если они существуют
  const sheetNames = ['Отчет по внесенной информации по продукции в Битрикс24', 'Отчёт формулы', 'Отчёт 2023-08-04', 'Отчёт 2023-08-11', 'Отчёт 2023-08-18', 'Отчёт 2023-08-25', 'Описание этой гуглтаблицы', 'Looker'];
  sheetNames.forEach(name => {
    const sheet = spread.getSheetByName(name);
    if (sheet) {
      // sheet.hideSheet();
    }
  });

  // Фильтруем сводную таблицу
  spread.toast('Фильтрую сводную ..');
  pivotFilterColumnRowsHide_RUN();

  spread.toast('Подготовка завершена.');
}

/** Заполнить таблицу товарами 
*/
function aMain_Bitrix24_Products_2_Sheet() {
  // очистить лист под заголовками
  // взять массив заголовков - по ним создавать таблицу для размещения на листе
  // запросить Bitrix24 batch

  let sheet_Name = 'Отчет по внесенной информации по продукции в Битрикс24';
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_Name);
  sheet.getRange("A2:H").clear();

  let arr_Heads = sheet.getRange("1:1").getValues();
  // REST_KEY на https://domain.bitrix24.ru/devops/edit/in-hook/477/
  let url = 'https://domain.bitrix24.ru/rest/853/REST_KEY/crm.product.list.json?select=id';

  const result = UrlFetchApp.fetch(url);
  // const unj = JSON.parse(result);
  // const id = unj['result'][0];
  const arr_IDs = JSON.parse(result)['result'];

}

/** пакетный зарос данных от Bitrix24 */
function bitrix24_Batch(url) {
  const result = UrlFetchApp.fetch(url);
  const arr_IDs = JSON.parse(result)['result'];


}

/** Trigger every 1 hours */
function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('aMain_Bitrix24_Products_2_Sheet')
    .timeBased()
    .everyHours(1)
    .create();
}

function sheetUpdateFromSheetRun() {
  sheetclearContent(SpreadsheetApp.getActiveSpreadsheet(), "Looker")
  sheetUpdateFromSheet("Отчет по внесенной информации по продукции в Битрикс24", "A:H", "Looker", "A1");
}

/** 
 * sheetSour имя листа источника
 * rangeSour диапазон источника "A:H"
 * sheetDest имя листа назначения
 * cellDest адрес ячейки куда вставлять массив
 * Пример использования:
 * sheetUpdateFromSheet("Лист1", "A1:H10", "Лист2", "A1");
 */
function sheetUpdateFromSheet(sheetSour, rangeSour, sheetDest, cellDest) {
  // копирует данные с одного листа на другой
  var sour = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetSour);
  var dest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetDest);
  var sourData = sour.getRange(rangeSour).getValues();
  dest.getRange(cellDest).offset(0, 0, sourData.length, sourData[0].length).setValues(sourData);
}

/** 
 * очистить содержимое листа
 * вызов
 * sheetclearContent(SpreadsheetApp.getActiveSpreadsheet(), "Лист1")
 */
function sheetclearContent(spread, sheetName) {
  var sheet = spread.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  sheet.getRange(1, 1, lastRow, lastColumn).clearContent();
}

function updatepivot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pivotSheetName = "Сводная"; // Замените на имя своего листа с сводной таблицей
  var pivotSheetId = ss.getSheetByName(pivotSheetName).getSheetId();

  var fields = "sheets(properties.sheetId,data.rowData.values.pivot)";
  var sheets = Sheets.Spreadsheets.get(ss.getId(), { fields: fields }).sheets;
  for (var i in sheets) {
    if (sheets[i].properties.sheetId == pivotSheetId) {
      var pivotParams = sheets[i].data[0].rowData[0].values[0].pivot;
      break;
    }
  }

  // Установите новые условия фильтрации, чтобы скрыть строки с "06.02" в столбце "Название"
  pivotParams.criteria = { 4: { "visibleValues": ["06.02"], "visibleByDefault": false } };

  // Отправьте обновленные параметры обратно
  var request = {
    "updateCells": {
      "rows": {
        "values": [{
          "pivot": pivotParams
        }]
      },
      "start": {
        "sheetId": pivotSheetId
      },
      "fields": "pivot",
    }
  };

  Sheets.Spreadsheets.batchUpdate({ 'requests': [request] }, ss.getId());
}

/**
 * Сводную фильтровать
 */
function pivotFilterColumnRowsHide_RUN() {

  // очистить все фильтры сводной
  SpreadsheetApp.getActive().getSheetByName('сводная').
    getPivotTables()[0].getFilters().
    forEach(filter => filter.remove());

  //Пустые строки скрываются в pivotFilterColumnRowHide
  let hide = [
    ['01.03. '],
    ['02.07.'],
    ['06.02'],
    ['07.01.'], ['07.02.'], ['07.03.'], ['07.04.'],
    ['08. Фи'],
    ['10.01. '], ['10.02. '], ['10.03. '],
    ['13.01. '],
    ['15. Прочая '], ['15.01. '], ['15.02. '], ['15.08. '], ['15.09. '],
    ['16. Това'],
    ['16.01. '], ['16.04. '], ['16.10. '],
    ['17.03'],
    ['20.01. '], ['20.02. '], ['20.03. '],
    ['99. Проч'],
    ['Импортированные товары'],
    ['Межкооперация']
  ];

  pivotFilterColumnRowHide('Сводная', 0, 'Раздел (уровень 2)', hide);

  // Скрыть пустые элементы в поле "Артикул"
  pivotHideEmptyValues('Сводная', 0, 'Артикул');
}

/**
 * В сводной таблице скрыть пустые значения в указанном столбце.
 */
function pivotHideEmptyValues(sheetName, pivotIndex, columnHeader) {

  // Получить массив значений из указанного столбца
  const column = pivotColumnValues(sheetName, pivotIndex, columnHeader);

  // Преобразовать двумерный массив в одномерный массив
  const columnValues = column.flat()
    .filter((item, index, self) => self.indexOf(item) === index) // Удалить дубликаты
    .filter(item => item !== ''); // Удалить пустые элементы

  const pivot = SpreadsheetApp.getActive().getSheetByName(sheetName).getPivotTables()[pivotIndex];

  // Номер заголовка столбца на листе источнике
  const sourceDataColumn = pivot.getSourceDataRange().getValues()[0].indexOf(columnHeader) + 1;

  if (sourceDataColumn < 1)
    throw new Error('Не найден заголовок ' + columnHeader);

  const criteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(columnValues)
    .build();

  pivot.addFilter(sourceDataColumn, criteria);
}

/** 
 * В сводной фильтровать столбец.
 * В гуглсводной фильтр только через создание массива одномерного значений.
 * Поэтому беру данные с листа источника и ихто и "фильтрую" - отсекая ненужное.
 */
function pivotFilterColumnRowHide(sheetName, pivotIndex, columnHeader, hide) {

  const column = pivotColumnValues(sheetName, pivotIndex, columnHeader);

  const columnValues = column
    .filter(row => !hide.some(term => row[0].includes(term)) && row[0] !== '')
    .flatMap(row => row) // в одномерный массив
    .filter((item, index, self) => self.indexOf(item) === index); // уникальные

  const pivot = SpreadsheetApp.getActive().getSheetByName(sheetName).getPivotTables()[pivotIndex];

  // номер заголовка столбца на листе источнике
  const sourceDataColumn = SpreadsheetApp.getActive().getSheetByName(sheetName).getPivotTables()[pivotIndex].getSourceDataRange().getValues()[0].indexOf(columnHeader) + 1;

  if (sourceDataColumn < 1)
    throw new Error('Не найден заголовок ' + эcolumnHeader);

  const criteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(columnValues)
    .build();

  pivot.addFilter(sourceDataColumn, criteria);
}

function pivotColumnValues_Test() {

  let array = pivotColumnValues('Сводная (копия)', 0, 'Раздел (уровень 2)');

  if (array.length < 1) throw new Error('Ошибка');

  Logger.log(array);
}

/** 
 * получить массив столбца сводной из источника
 */
function pivotColumnValues(sheetName, pivotIndex, columnName) {
  return SpreadsheetApp.getActive().getSheetByName(sheetName).getPivotTables()[pivotIndex].getSourceDataRange().getValues().arrayColumnbyHeader(columnName);
}

/**
 * Функция, которой можно обработать массив
 * Вызов массив.arrayColumnbyHeader(columnName)
 */
Array.prototype.arrayColumnbyHeader = function (columnHeader) {

  const columnIndex = this[0].indexOf(columnHeader);

  return this.map(row => [row[columnIndex]]);
};

/**
 * Заполнить формулами соседние ячейки
 * Формулы должны быть уже заполнены
 * https://domain.bitrix24.ru/company/personal/user/853/tasks/task/view/46578/?any=user%2F853%2Ftasks%2Ftask%2Fview%2F46578%2F
 */

function formulasFill() {
  const spread = SpreadsheetApp.getActive();
  const sheet = spread.getSheetByName("Отчёт формулы");
  const column_Start = 31;
  const column_Stop = 59;
  const suffix = " 2023-09-20";

  for (let col = column_Start; col < column_Stop; col += 2) {

    setFormulaForRow(sheet, 1, col);

    setFormulaForRow(sheet, 2, col);

    cellValueBasedOnAnotherCell(sheet, 1, col, suffix);

    formulaWithChangedColumn(sheet, 2, col, 1);
  }
}

function formulaWithChangedColumn(sheet, row, col, change) {
  sheet.getRange(row, col + 3).setFormula(
    formulaChangeComplex(sheet.getRange(row, col + 1).getFormula(), change));
}

function cellValueBasedOnAnotherCell(sheet, row, col, suffix) {
  sheet.getRange(row, col + 3).setValue(
    sheet.getRange(row, col + 2).getValue() + suffix);
}

function setFormulaForRow(sheet, row, col) {
  let cellSour = sheet.getRange(row, col);
  let formula = cellSour.getFormula();
  let cellDest = cellSour.offset(0, 2);
  cellDest.setFormula(addressColumnchange(formula, 1));
}

function formulaChangeComplex_Test() {
  let formula = '=ЕСЛИ( ДЛСТР($B2) <> 12;""; ЕСЛИОШИБКА( ИНДЕКС(\'2023-09-20\'!R:R; ПОИСКПОЗ($B2; \'2023-09-20\'!$B:$B; 0); 1);""))';
  let wanted = '=ЕСЛИ( ДЛСТР($B2) <> 12;""; ЕСЛИОШИБКА( ИНДЕКС(\'2023-09-20\'!S:S; ПОИСКПОЗ($B2; \'2023-09-20\'!$B:$B; 0); 1);""))';

  let result = formulaChangeComplex(formula, 1)

  loggerLog_If_Diff(result, wanted);

  formula = '=ЕСЛИ( ДЛСТР($B2) <> 12;""; ЕСЛИОШИБКА( ИНДЕКС(\'2023-09-20\'!AE:AE; ПОИСКПОЗ($B2; \'2023-09-20\'!$B:$B; 0); 1);""))';
  wanted = '=ЕСЛИ( ДЛСТР($B2) <> 12;""; ЕСЛИОШИБКА( ИНДЕКС(\'2023-09-20\'!AG:AG; ПОИСКПОЗ($B2; \'2023-09-20\'!$B:$B; 0); 1);""))';

  result = formulaChangeComplex(formula, 2)

  loggerLog_If_Diff(result, wanted);
}

function loggerLog_If_Diff(result, wanted) {
  if (result !== wanted) {
    Logger.log("ОШибка" + "\n");
    Logger.log('result = ' + result + "\n");
    Logger.log('wanted = ' + wanted + "\n");
  }
}

/**
 * изменит столбцы относительных адресов по diff, 
 * абсолютные адреса без изменений
 */
function formulaChangeComplex(formula, diff) {
  // Регулярное выражение для поиска адресов в формуле в формате R:R или AE:AE
  const addressRegex = /([A-Z]+:[A-Z]+)/g;

  // Функция, которая будет вызвана для каждого найденного адреса
  function replaceAddress(match) {
    // Разбиваем адрес на две части, разделенные ":"
    const parts = match.split(':');
    if (parts.length === 2) {
      // Вычисляем новые адреса для каждой части и объединяем их снова в формате R:R или AE:AE
      const newAddress = columnToLetter(letterToColumn(parts[0]) + diff) + ':' + columnToLetter(letterToColumn(parts[1]) + diff);
      return newAddress;
    } else {
      // Если формат адреса не соответствует ожидаемому, возвращаем его без изменений
      return match;
    }
  }

  // Применяем функцию замены к формуле
  formula = formula.replace(addressRegex, replaceAddress);

  return formula;
}

/**
 * преобразует буквенное обозначение столбца в числовое
 */
function letterToColumn(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
  }
  return column;
}

/**
 * преобразует числовое обозначение столбца в буквенное
 */
function columnToLetter(column) {
  let letter = '';
  while (column > 0) {
    let remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - remainder - 1) / 26); // изменено
  }
  return letter;
}

function addressColumnchange_Test() {
  let formula = "='Лист'!R1";
  let change = 1;
  let result = addressColumnchange(formula, change);
  let wanted = "='Лист'!S1";
  if (result !== wanted) Logger.log('Ошибка, вернулось: ' + result + ' вместо ' + wanted);

  formula = "='Лист'!BI23";
  change = -2;
  result = addressColumnchange(formula, change);
  wanted = "='Лист'!BG23";
  if (result !== wanted) Logger.log('Ошибка, вернулось: ' + result + ' вместо ' + wanted);
}

/** 
 * В формуле изменить буквы столбца на колво change
 */
function addressColumnchange(formula, change) {
  // Регулярное выражение для поиска буквенной части столбца
  let regex = /[A-Z]+/;
  let match = formula.match(regex);
  let column = match[0];

  // Преобразуем буквенную часть столбца в числовое значение
  let columnNumber = 0;
  for (let i = 0; i < column.length; i++) {
    columnNumber = columnNumber * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }

  // Изменяем номер столбца на указанное значение
  columnNumber += change;

  // Преобразуем числовое значение обратно в буквенную часть столбца
  let newColumn = '';
  while (columnNumber > 0) {
    let remainder = (columnNumber - 1) % 26;
    newColumn = String.fromCharCode('A'.charCodeAt(0) + remainder) + newColumn;
    columnNumber = Math.floor((columnNumber - remainder) / 26);
  }

  // Заменяем старую буквенную часть столбца на новую в формуле
  return formula.replace(regex, newColumn);
}

function arraySuffix_Test() {
  let inputArray = ['1', '2'];
  let suffix = ' q';
  let expected = ['1', '1 q', '2', '2 q'];
  let result = arraySuffix(inputArray, suffix);

  // Проверяем, совпадает ли результат с ожидаемым результатом
  assertEquals(expected, result);
}

function assertEquals(expected, actual) {
  if (JSON.stringify(expected) !== JSON.stringify(actual)) {
    Logger.log('Тест не пройден: результаты не совпадают.');
  }
}

function arraySuffix(inputArray, suffix) {
  let resultArray = [];
  for (let i = 0; i < inputArray.length; i++) {
    resultArray.push(inputArray[i]);
    resultArray.push(inputArray[i] + suffix);
  }
  return resultArray;
}

/**
 * промежуточный вариант
 */
function pivotFieldCalculatedAdd_Test() {
  let pivot = SpreadsheetApp.getActive().getSheetByName('сводная').getPivotTables()[0];
  let fieldName = 'Прод. группа';
  let formula = "=(СЧЁТЗ('Прод. группа') - СЧИТАТЬПУСТОТЫ('Прод. группа')) / СЧЁТЗ('Название')";

  pivot.addCalculatedPivotValue(fieldName, formula).summarizeBy(SpreadsheetApp.PivotTableSummarizeFunction.CUSTOM);
}

/**
 * Добавить столбцы в сводную 
 */
function pivotFieldsCalculatedAdd_RUN() {

  const pivot = SpreadsheetApp.getActive().getSheetByName('сводная').getPivotTables()[0];

  const column_Start = 31;

  const numerators = rangeToLastNonEmptyCell(
    SpreadsheetApp.getActive().getSheetByName('Отчёт формулы').getRange(1, column_Start)).getValues()[0];

  pivotFieldsCalculatedAdd(pivot, numerators, 'Название')
}

function pivotFieldsCalculatedAdd(pivot, numerators, deNominator) {
  // let pivot = SpreadsheetApp.getActive().getSheetByName('сводная').getPivotTables()[0];
  // let fieldName = 'Прод. группа';

  for (i = 0; i < numerators.length; i++) {

    let numerator = numerators[i];

    // let formula = "=(СЧЁТЗ('" + numerator + "') - СЧИТАТЬПУСТОТЫ('" + numerator + "')) / СЧЁТЗ('" + deNominator + "')";

    // непонятно, но иногда СЧЁТЗ, то учитывает формулы, то не учитывает
    let formula = "=ЕСЛИ((СЧЁТЗ('" + numerator + "') - СЧИТАТЬПУСТОТЫ('" + numerator + "')) < 0; СЧЁТЗ('" + numerator + "'); СЧЁТЗ('" + numerator + "') - СЧИТАТЬПУСТОТЫ('" + numerator + "')) / СЧЁТЗ('" + deNominator + "')";

    pivot.addCalculatedPivotValue(numerator, formula).summarizeBy(SpreadsheetApp.PivotTableSummarizeFunction.CUSTOM);
  }
}

function rangeToLastNonEmptyCell_Test() {
  let arr = rangeToLastNonEmptyCell(
    SpreadsheetApp.getActive().getSheetByName('Отчёт формулы').getRange(1, 31)).getValues()[0];
  Logger.log(arr.length);
}

/**
 * принимает ячейку, а возвращает диапазон вправо по строке по последнюю непустую ячейку,  без цикла.
 */
function rangeToLastNonEmptyCell(cell) {
  var sheet = cell.getSheet();
  var row = cell.getRow();
  var column = cell.getColumn();

  var values = sheet.getRange(row, column, 1, sheet.getLastColumn() - column + 1).getValues()[0];
  var lastNonEmptyIndex = values.filter(String).length;

  return sheet.getRange(row, column, 1, lastNonEmptyIndex);
}

function pivotRangeA1Notation_Test() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('сводная');
  let a1NOtation = pivotRangeA1Notation(sheet, 0);
  if (a1NOtation === undefined) throw new Error('pivotRange_Test sheet,1) === undefined');
  a1NOtation = pivotRangeA1Notation(sheet, 1);
  if (a1NOtation !== undefined) throw new Error('pivotRange_Test sheet,1) !== undefined');
}

function pivotRangeA1Notation(sheet, index) {
  const pivot = sheet.getPivotTables()[index];

  if (pivot) {
    return pivot.getSourceDataRange().getA1Notation();
  }
}
