function AutomateOutput() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var output = activeSpreadsheet.insertSheet().setName("Automated Output");
  
  output.getRange('A2').setFormula('=UNIQUE(SalespersonNameSales)').setFontWeight('bold');
  output.getRange('B2').setFormula('=UNIQUE(SalespersonIDSales)').setFontWeight('bold');
  
  output.getRange('C2').setValue('Total Sales: ').setFontWeight('bold');
  output.getRange('C3').setFormula('=SUM(IFERROR(FILTER(TotalSoldSales,MONTH(DatePurchasedSales)=10,SalespersonIDSales=325291)))');
  output.getRange('C4').setFormula('=SUM(IFERROR(FILTER(TotalSoldSales,MONTH(DatePurchasedSales)=10,SalespersonIDSales=348471)))');
  output.getRange('C5').setFormula('=SUM(IFERROR(FILTER(TotalSoldSales,MONTH(DatePurchasedSales)=10,SalespersonIDSales=379409)))');
  output.getRange('C6').setFormula('=SUM(IFERROR(FILTER(TotalSoldSales,MONTH(DatePurchasedSales)=10,SalespersonIDSales=345059)))');
  output.getRange('C:C').setNumberFormat('"$"#,##0.00');
  
  output.getRange('D2').setValue('Total Return: ').setFontWeight('bold');
  output.getRange('D3').setFormula('=SUM(IFERROR(FILTER(TotalReturnFinance,Month(DatePurchasedFinance)=10,SalespersonIDFinance=325291)))');
  output.getRange('D4').setFormula('=SUM(IFERROR(FILTER(TotalReturnFinance,Month(DatePurchasedFinance)=10,SalespersonIDFinance=348471)))');
  output.getRange('D5').setFormula('=SUM(IFERROR(FILTER(TotalReturnFinance,Month(DatePurchasedFinance)=10,SalespersonIDFinance=379409)))');
  output.getRange('D6').setFormula('=SUM(IFERROR(FILTER(TotalReturnFinance,Month(DatePurchasedFinance)=10,SalespersonIDFinance=345059)))');
  output.getRange('D:D').setNumberFormat('"$"#,##0.00');
  
  output.getRange('E2').setValue('Net Sales (Difference): ').setFontWeight('bold');
  output.getRange('E3').setFormula('=C3-D3');
  output.getRange('E4').setFormula('=C4-D4');
  output.getRange('E5').setFormula('=C5-D5');
  output.getRange('E6').setFormula('=C5-D5');
  output.getRange('E:E').setNumberFormat('"$"#,##0.00');
  
  output.getRange('F2').setValue('Number of Complaints: ').setFontWeight('bold');
  output.getRange('F3').setFormula('=COUNT(IFERROR(FILTER(SalespersonIDHD,MONTH(DatePurchasedHD)=10,SalespersonIDHD=325291, ComplaintSalesperson="YES")))');
  output.getRange('F4').setFormula('=COUNT(IFERROR(FILTER(SalespersonIDHD,MONTH(DatePurchasedHD)=10,SalespersonIDHD=348471, ComplaintSalesperson="YES")))');
  output.getRange('F5').setFormula('=COUNT(IFERROR(FILTER(SalespersonIDHD,MONTH(DatePurchasedHD)=10,SalespersonIDHD=379409, ComplaintSalesperson="YES")))');
  output.getRange('F6').setFormula('=COUNT(IFERROR(FILTER(SalespersonIDHD,MONTH(DatePurchasedHD)=10,SalespersonIDHD=345059, ComplaintSalesperson="YES")))');
  
  var date = new Date();
  var mt = date.getMonth();
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var currentD = months[mt] + " Report For Human Resources";  
  
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
