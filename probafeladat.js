var url = 'https://jsonplaceholder.typicode.com/comments'
var response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
var json = response.getContentText();
var jsonData = JSON.parse(json);

function calculateDiffAndAddEmail() {
  var sheet = SpreadsheetApp.getActive()
  var data = sheet.getDataRange().getValues()

  for (i in data) {
    var start = new Date(data[i][0])
    var end = new Date(data[i][1])
    var differenceInMilis = end - start
    var days = Math.floor(differenceInMilis / (1000 * 60 * 60 * 24)) + 1
    var currentCell = parseInt(i) + 1
    sheet.getRange('C' + currentCell).setValue(days)
    sheet.getRange('D' + currentCell).setValue(jsonData[i].email)
  }
}