function sendXcel() {
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheet.getId();
  var sheets = spreadsheet.getSheets();
  var keepSheet = 'November Automated Output';

  //Code to hide all the sheets that don't need to be sent.
  for(var i=0; i<sheets.length; i++){ 
    Logger.log(i); 
    if(sheets[i].getName()!=keepSheet){ 
      sheets[i].hideSheet(); } }
  
    //Code to remove all blank rows from sheet
  for (var s in sheets){
  var sheet=sheets[s]
  var maxRows = sheet.getMaxRows(); 
  var lastRow = sheet.getLastRow();
  try{
    if (maxRows-lastRow != 0){sheet.deleteRows(lastRow+1, maxRows-lastRow);}
  }
  catch(e){}
}
  
  //Code to save excel attachment.
  var url = 'https://docs.google.com/spreadsheets/d/'+spreadsheetId+'/export?format=xlsx';
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, { headers: { 'Authorization' : 'Bearer ' + token } } );
  var fileName = (spreadsheet.getName()) + '.xlsx';
  var blobs = [response.getBlob().setName(fileName)];

 //Code to send the email with attachment.
  var mailTo = 'smit1090@gmail.com',
        subject = 'Hello, This is my Report',
        body = 'Please check the attached PDF, it contains the report for Employees Total Sales, Total Return, Total Net Sales, and Total Customer Complaints.'
  MailApp.sendEmail(mailTo, subject, body, {attachments: blobs});
  
  //Code to restore the other sheet in the spreadsheet once email is sent.
     sheets.forEach(function(s) {s.showSheet();})
}
