// This script will run on the 1st of every month from 12 AM - 1 AM GMT. 
function automateOutput() {
  
  var date = new Date();
  var mt = date.getMonth();
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var currentD = months[mt-2] +" Report For Human Resources";
  var monthReport=mt-1;
  
  //Deletes the current Output Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Output"));
  ss.deleteActiveSheet();     

  //Creates a new Output Sheet
  var output = SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName("Output");
  ss.moveActiveSheet(ss.getNumSheets()); //moves the sheet at the end for simplicity sake.
  
  //Input all of the Data into the Output Sheet.
  output.getRange('A2').setFormula('=UNIQUE(SSalespersonName)').setFontWeight('bold');
  output.getRange('B2').setFormula('=UNIQUE(SSalespersonID)').setFontWeight('bold');
  
  //Gets the Range for SalespersonID from the Sales Sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales");
  var myStringArray = sheet.getRange('C2:C').getValues();
  
  //Converts the 2D SalespersonID into 1D array.
  var newArr = [];
  for(var i = 0; i < myStringArray.length; i++){
    newArr = newArr.concat(myStringArray[i]);}
  
 //Use the filter method to get the unique values from the SalespersonID array.
  Array.prototype.unique = function() {
  return this.filter(function (value, index, self) { 
    return self.indexOf(value) === index;});}
  var x = newArr.unique();

  //Create a New Array using the Map function with only the unique values
  var map = x.map(function (el) {
        return [el];
    });
  
  //Input Sales Data
  output.getRange('C2').setValue('Total Sales').setFontWeight('bold');
  output.getRange(3, 3, map.length-1).setFormulaR1C1('=SUM(IFERROR(FILTER(STotalSold,MONTH(SDatePurchased)=prevMonth(),SSalespersonID=R[0]C[-1])))');
  output.getRange('C3:D').setNumberFormat('"$"#,##0.00');
  
  //Input Return Data
  output.getRange('D2').setValue('Total Return').setFontWeight('bold');
  output.getRange(3, 4, map.length-1).setFormulaR1C1('=SUM(IFERROR(FILTER(FTotalReturn,Month(FDatePurchased)=prevMonth(),FSalespersonID=R[0]C[-2])))');
  output.getRange('D3:D').setNumberFormat('"$"#,##0.00');
  
  //Input Net Sales
  output.getRange('E2').setValue('Net Sales').setFontWeight('bold');
  output.getRange(3, 5, map.length-1).setFormulaR1C1('=R[0]C[-2]-R[0]C[-1]');
  output.getRange('E3:E').setNumberFormat('"$"#,##0.00');
  
  //Input Complaint
  output.getRange('F2').setValue('Number of Complaints').setFontWeight('bold');
  output.getRange(3, 6, map.length-1).setFormulaR1C1('=COUNT(IFERROR(FILTER(HSalespersonID,MONTH(HDatePurchased)=prevMonth(),HSalespersonID=R[0]C[-4], HComplaint="YES")))');
  
  
   //Data Validation for Complaints
   output.getRange(3, 6, map.length)
  .setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireNumberLessThanOrEqualTo(0)
  .build());
  
  //Stylize the Output Sheet.
  output.getRange('A1:F1').activate()
  .mergeAcross().setValue(currentD);
  output.getActiveRangeList().setBackground('#b7e1cd')
  .setFontSize(14)
  .setFontWeight('bold')
  .setHorizontalAlignment('center');
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G:Z').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getActiveSheet().autoResizeColumns(1, 6);
}

function prevMonth() {
  var date = new Date();
  var mt = date.getMonth();
  var monthReport=mt-1;
  return monthReport;
}
