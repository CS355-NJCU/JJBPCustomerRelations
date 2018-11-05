function automateOutput() {
  
  var date = new Date();
  var mt = date.getMonth();
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var currentD = months[mt] +" Report For Human Resources";
  
  if(mt+1==10){
  
  var output = SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName(months[mt]+" Automated Output");
    
  output.getRange('A2').setFormula('=UNIQUE(SSalespersonName)').setFontWeight('bold');
  output.getRange('B2').setFormula('=UNIQUE(SSalespersonID)').setFontWeight('bold');
  
  output.getRange('C2').setValue('Total Sales').setFontWeight('bold');
  output.getRange('C3').setFormula('=SUM(IFERROR(FILTER(STotalSold,MONTH(SDatePurchased)=10,SSalespersonID=325291)))');
  output.getRange('C4').setFormula('=SUM(IFERROR(FILTER(STotalSold,MONTH(SDatePurchased)=10,SSalespersonID=348471)))');
  output.getRange('C5').setFormula('=SUM(IFERROR(FILTER(STotalSold,MONTH(SDatePurchased)=10,SSalespersonID=379409)))');
  output.getRange('C6').setFormula('=SUM(IFERROR(FILTER(STotalSold,MONTH(SDatePurchased)=10,SSalespersonID=345059)))');
  output.getRange('C:C').setNumberFormat('"$"#,##0.00');
  
  output.getRange('D2').setValue('Total Return').setFontWeight('bold');
  output.getRange('D3').setFormula('=SUM(IFERROR(FILTER(FTotalReturn,Month(FDatePurchased)=10,FSalespersonID=325291)))');
  output.getRange('D4').setFormula('=SUM(IFERROR(FILTER(FTotalReturn,Month(FDatePurchased)=10,FSalespersonID=348471)))');
  output.getRange('D5').setFormula('=SUM(IFERROR(FILTER(FTotalReturn,Month(FDatePurchased)=10,FSalespersonID=379409)))');
  output.getRange('D6').setFormula('=SUM(IFERROR(FILTER(FTotalReturn,Month(FDatePurchased)=10,FSalespersonID=345059)))');
  output.getRange('D:D').setNumberFormat('"$"#,##0.00');
  
  output.getRange('E2').setValue('Net Sales').setFontWeight('bold');
  output.getRange('E3').setFormula('=C3-D3');
  output.getRange('E4').setFormula('=C4-D4');
  output.getRange('E5').setFormula('=C5-D5');
  output.getRange('E6').setFormula('=C5-D5');
  output.getRange('E:E').setNumberFormat('"$"#,##0.00');
  
  output.getRange('F2').setValue('Number of Complaints').setFontWeight('bold');
  output.getRange('F3').setFormula('=COUNT(IFERROR(FILTER(HSalespersonID,MONTH(HDatePurchased)=10,HSalespersonID=325291, HComplaint="YES")))');
  output.getRange('F4').setFormula('=COUNT(IFERROR(FILTER(HSalespersonID,MONTH(HDatePurchased)=10,HSalespersonID=348471, HComplaint="YES")))');
  output.getRange('F5').setFormula('=COUNT(IFERROR(FILTER(HSalespersonID,MONTH(HDatePurchased)=10,HSalespersonID=379409, HComplaint="YES")))');
  output.getRange('F6').setFormula('=COUNT(IFERROR(FILTER(HSalespersonID,MONTH(HDatePurchased)=10,HSalespersonID=345059, HComplaint="YES")))');
  
 // output.getRange('G2').setValue('Average Salesperson Rating: ').setFontWeight('bold');
 // output.getRange('G3').setFormula('=AVERAGE(IFERROR(FILTER(HRating,MONTH(HDatePurchased)=10,HSalespersonID=325291))');
 // output.getRange('G4').setFormula('=AVERAGE(IFERROR(FILTER(HRating,MONTH(HDatePurchased)=10,HSalespersonID=348471))');
 // output.getRange('G5').setFormula('=AVERAGE(IFERROR(FILTER(HRating,MONTH(HDatePurchased)=10,HSalespersonID=379409))');
 // output.getRange('G6').setFormula('=AVERAGE(IFERROR(FILTER(HRating,MONTH(HDatePurchased)=10,HSalespersonID=345059))');
  
  output.getRange('A1:F1').activate()
  .mergeAcross().setValue(currentD);
  output.getActiveRangeList().setBackground('#b7e1cd')
  .setFontSize(14)
  .setFontWeight('bold')
  .setHorizontalAlignment('center');
  
 
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G:Z').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
}
}
