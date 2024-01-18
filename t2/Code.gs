function myFunction() {
  sendSpreadsheetAsExcel();
}

// 직접 데이터 보내기
function sendEmailFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A1").getValue(); // A1 셀의 데이터를 읽습니다.
  var recipient = "jeonghk@mz.co.kr"; // 이메일 받는 사람 주소
  var subject = "오늘의 이메일 제목"; // 이메일 제목
  var body = "오늘의 데이터: " + data; // 이메일 내용
  
  GmailApp.sendEmail(recipient, subject, body);
}

// 파일로 보내기
function sendSpreadsheetAsExcel() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = spreadsheet.getId();
  var url = 'https://docs.google.com/feeds/download/spreadsheets/Export?key=' + fileId + '&exportFormat=xlsx';
  var params = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, params);
  var blobs = [response.getBlob().setName(spreadsheet.getName() + ".xlsx")];

  // 현재 날짜를 문자열로 포맷팅
  var currentDate = new Date();
  var formattedDate = currentDate.getFullYear() + "-" +
                      (currentDate.getMonth() + 1) + "-" +
                      currentDate.getDate();

  MailApp.sendEmail({
    to: 'dsa@das.co.kr',
    subject: '['+formattedDate + '] 엑셀 관련 건',
    body: '퇴근 ㄱㄱ',
    attachments: blobs
  });

}

// 변한 부분 체크
function onEdit(e) {
  // e 객체에는 편집된 셀의 정보가 포함됩니다.
  var range = e.range; // 변경된 셀의 범위
  var sheet = range.getSheet();
  var editedValue = range.getValue(); // 변경된 셀의 값
  var recipient = "받는사람의이메일@example.com"; // 이메일 받는 사람 주소
  var subject = sheet.getName() + "에서 변경 발생"; // 이메일 제목
  var body = "변경된 셀: " + range.getA1Notation() + "\n변경된 값: " + editedValue; // 이메일 내용

  // 이메일 보내기
  GmailApp.sendEmail(recipient, subject, body);
}