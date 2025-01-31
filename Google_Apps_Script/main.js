/*
Все права защищены (c) 2024
Данный скрипт предназначен для работы с Google Таблицами:
- копирование листов,
- настройка вида и форматов,
- поиск дублей,
- замена ИНН и удаление ссылок,
- назначение ответственных и прочее.

Не подлежит свободному использованию или копированию без согласия автора.
*/

var rowsToCheck = [];
var iAmDead = [];
var really = [];

function переносНовыхТаблиц() {
  var sourceSpreadsheetId = '';
  var targetSpreadsheetId = '';
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  var sourceSheets = sourceSpreadsheet.getSheets();
  var targetSheets = targetSpreadsheet.getSheets().map(sheet => sheet.getName());
  for (var i = 0; i < sourceSheets.length; i++) {
    var sourceSheet = sourceSheets[i];
    var sourceSheetName = sourceSheet.getName();
    if (!targetSheets.includes(sourceSheetName)) {
      var copiedSheet = sourceSheet.copyTo(targetSpreadsheet);
      copiedSheet.setName(sourceSheetName);
      copiedSheet.getRange('Z2').setValue('новый');     // Флаг для "заменитьИННВоВсехЛистах"
      copiedSheet.getRange('Z3').setValue('проверка');  // Флаг для "вид"
      copiedSheet.getRange('Z4').setValue('честно');    // Флаг для "сверкаДубликатовВНовыхЛистахСНашимиИННизБитрикса"
      copiedSheet.getRange('Z5').setValue('вдруг');     // Флаг для "ответственный"
      copiedSheet.getRange('Z6').setValue('пример');    // Флаг для "убратьОбъединения"
      copiedSheet.getRange('Z7').setValue('арбуз');     // Флаг для "дублиВнутриЛиста"
    }
  }
}

function заменитьИННВоВсехЛистах() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  sheets.forEach(function(sheet) {
    var isNewFlag = sheet.getRange('Z2').getValue() === 'новый';
    if (isNewFlag) {
      rowsToCheck.forEach(function(row) {
        var cellValue = sheet.getRange('B' + row).getValue();
        if (typeof cellValue === 'string' && cellValue.startsWith("ИНН: ")) {
          sheet.getRange('B' + row).setValue(cellValue.replace("ИНН: ", ""));
        }
      });
      sheet.getRange('Z2').setValue('');
    }
  });
}

function вид() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  sheets.forEach(function(sheet) {
    var isNewFlag = (sheet.getRange('Z3').getValue() === 'проверка');
    if (isNewFlag) {
      sheet.setColumnWidth(1, 75);
      sheet.setColumnWidth(3, 500);
      sheet.getRange('A:A').setFontWeight('bold');
      sheet.getRange('1:1')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      sheet.getRange('C1').setValue('Комментарий менеджера');
      sheet.getRange('D1').setValue('Дубль');
      sheet.getRange('E1').setValue('ИHH');
      sheet.getRange('F1').setValue('Ответственный');
      sheet.getRange('A1:F1').setBorder(true, true, true, true, true, true);
      sheet.getRange('Z3').setValue('');
    }
  });
}

function сверкаДубликатовВНовыхЛистахСНашимиИННизБитрикса() {
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSpreadsheetId = '';
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  var innSheet = targetSpreadsheet.getSheetByName("ИНН");
  var innData = innSheet.getRange("A:B").getValues();
  var sourceSheets = sourceSpreadsheet.getSheets();
  sourceSheets.forEach(sheet => {
    if (sheet.getRange('Z4').getValue() === 'честно') {
      rowsToCheck.forEach(row => {
        var cellValue = sheet.getRange('B' + row).getValue();
        innData.forEach(([innValue, relatedValue]) => {
          if (cellValue === innValue) {
            var cellRange = sheet.getRange('E' + row);
            cellRange.setValue(`Дубль: ${innValue}, ${relatedValue}`);
            sheet.getRange('B' + row).setBackground('#ffc7ce');
          }
        });
      });
      sheet.getRange('Z4').setValue('');
    }
  });
}

function убратьОбъединения() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  sheets.forEach(function(sheet) {
    var flagCell = sheet.getRange('Z6');
    if (flagCell.getValue() === 'пример') {
      var range = sheet.getDataRange();
      var mergedRanges = range.getMergedRanges();
      mergedRanges.forEach(function(mergedRange) {
        mergedRange.breakApart();
      });
      flagCell.setValue('');
      SpreadsheetApp.flush();
    }
  });
}

function ответственный() {
  var targetSpreadsheetId = '';
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  var innSheet = targetSpreadsheet.getSheetByName("Компании");
  var innData = innSheet.getRange("A:B").getValues();
  var backgroundColors = innSheet.getRange("B1:B" + innSheet.getLastRow()).getBackgrounds();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    var flagCell = sheet.getRange('Z5');
    if (flagCell.getValue() === 'вдруг') {
      rowsToCheck.forEach(row => {
        var cellValue = sheet.getRange('B' + row).getValue();
        innData.forEach(([innValue], index) => {
          if (cellValue === innValue && index < backgroundColors.length) {
            var backgroundColor = backgroundColors[index][0];
            if (backgroundColor) {
              var startRow = row - 1;
              var numRows = 8;
              startRow = Math.max(startRow, 1);
              var lastRow = Math.min(sheet.getLastRow(), startRow + numRows - 1);
              sheet.getRange('F' + startRow + ':F' + lastRow).setBackground(backgroundColor);
            }
          }
        });
      });
      flagCell.setValue('');
      SpreadsheetApp.flush();
    }
  });
}

function дублиВнутриЛиста() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  sheets.forEach(function(sheet) {
    var flagCell = sheet.getRange('Z7');
    if (flagCell.getValue() === 'арбуз') {
      var valuesToCheck = {};
      rowsToCheck.forEach(function(row) {
        var value = sheet.getRange("B" + row).getValue();
        if (value) {
          if (valuesToCheck.hasOwnProperty(value)) {
            var message = "Нашелся дубль: " + value + ", строка " + valuesToCheck[value].row;
            sheet.getRange("A" + row).setValue(message).setBackground("#000").setFontColor("#FFF");
          } else {
            valuesToCheck[value] = { row: row };
          }
        }
      });
      flagCell.setValue('');
      SpreadsheetApp.flush();
    }
  });
}

function сверкаДубликатов() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  sheets.sort((a, b) => b.getIndex() - a.getIndex());
  sheets.forEach((sheet, sIndex) => {
    var sheetValues = rowsToCheck.map(row => sheet.getRange('B' + row).getValue()).filter(val => val);
    sheets.slice(sIndex + 1).forEach(compareSheet => {
      var compareSheetValues = rowsToCheck.map(row => compareSheet.getRange('B' + row).getValue()).filter(val => val);
      sheetValues.forEach((value, index) => {
        if (compareSheetValues.includes(value)) {
          var rowIndex = rowsToCheck[index];
          var compareRowIndex = rowsToCheck[compareSheetValues.indexOf(value)];
          var cellAddress = 'D' + rowIndex;
          var compareCellAddress = 'D' + compareRowIndex;
          sheet.getRange(cellAddress).setValue(`Дубль на листе ${compareSheet.getName()} (${compareCellAddress})`);
          compareSheet.getRange(compareCellAddress).setValue(`Дубль на листе ${sheet.getName()} (${cellAddress})`);
        }
      });
    });
  });
}

function удалениеСсылок() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  sheets.forEach(function(sheet) {
    if (sheet.getRange('Z8').getValue() === 'ссылка') {
      var range = sheet.getRange('A:A');
      var values = range.getValues();
      var regex = /<a\s+(?:[^>]*?\s+)?href=(["'])(.*?)\1/ig;
      for (var i = 0; i < values.length; i++) {
        var cellValue = values[i][0];
        if (typeof cellValue === 'string') {
          var newValue = cellValue.replace(regex, '');
          range.getCell(i + 1, 1).setValue(newValue);
        }
      }
      sheet.getRange('Z8').setValue('');
    }
  });
}
