function sendPdf() {

  var date = new Date();
  var mt = date.getMonth();
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var currentD = months[mt] +" Customer Relations Report For Human Resources";  
  
  
    var mailTo = 'smit1090@gmail.com',
        subject = currentD,
        body = 'Please check the attached PDF, it contains the report for Employees Total Sales, Total Return, Total Net Sales, and Total Customer Complaints.',
        sheetNum = 4, // first sheet(tab) is zero, second sheet is 1, etc..

        source = SpreadsheetApp.getActiveSpreadsheet(),
        sheets = source.getSheets(),
        url, token, response;
    sheets.forEach(function(s, i) {
        if (i !== sheetNum) s.hideSheet();
    });
    url = Drive.Files.get(source.getId())
        .exportLinks['application/pdf'];
    url = url + '&size=letter' + //paper size
        '&portrait=true' + //orientation, false for landscape
        '&fitw=true' + //fit to width, false for actual size
        '&sheetnames=false&printtitle=false&pagenumbers=false' + //hide optional
        '&gridlines=true' + //false = hide gridlines
        '&fzr=false'; //do not repeat row headers (frozen rows) on each page
    token = ScriptApp.getOAuthToken();
    response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });

    MailApp.sendEmail(mailTo, subject, body, {
        attachments: [response.getBlob()]
    });
    sheets.forEach(function(s) {
        s.showSheet();
    })

}
