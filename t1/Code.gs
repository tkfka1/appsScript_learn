// 스프레드시트의 모든 시트 이름을 반환하는 함수
function getSheetNames() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(sheet => sheet.getName());
}

// doGet 함수를 사용하여 HTML 파일을 서브
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

// 특정 시트의 특정 행 데이터를 반환하는 함수
function getRowData(sheetName, rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return [];
  }
  var data = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  // 날짜 데이터를 문자열로 변환
  return data.map(function(value) {
    if (value instanceof Date) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    return value;
  });
}

// 특정 시트의 특정 행에 데이터를 업데이트하는 함수
function updateRowData(sheetName, rowIndex, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return;
  }
  var range = sheet.getRange(rowIndex, 1, 1, rowData.length);
  range.setValues([rowData]);
}