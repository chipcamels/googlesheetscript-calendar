var SPREADSHEET_APP_ID = 'app-id';
var SPREADSHEET_NAME = 'sheet-name';
var CALENDAR_ID = 'cal-id';

function runMe() {
  
//  DeleteAllRows(); // Reset Sheet
  var startTime= new Date();

  var sheet = SpreadsheetApp.openById(SPREADSHEET_APP_ID).getSheetByName(SPREADSHEET_NAME);
  
  var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event","Tag Name","Week Number","Project","Root Tag","Next Tag"]]
  var range = sheet.getRange(1,1,1,19);
  
  range.setValues(header);
  
  myFunction(sheet, CALENDAR_ID, new Date('Jan 1, 2019 00:00:00 EST'), new Date('Jan 31, 2019 23:59:59'));
//  myFunction(sheet, CALENDAR_ID, new Date('Feb 1, 2019 00:00:00 EST'), new Date());
  
//  while (rowNumber < numRows && !isTimeUp(startTime)) {
//    Logger.log(rowNumber);
//    myFunction(CALENDAR_ID, sheet, events[rowNumber]);
//    rowNumber++;
//  }

}
function myFunction(sheet, mycal, start, end) {
  var cal = CalendarApp.getCalendarById(mycal);
  var events = cal.getEvents(start, end);
//  var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event","Tag Name","Week Number","Project","Root Tag","Next Tag"]]
//  var range = sheet.getRange(1,1,1,19);
//  range.setValues(header);
  
  for (var i = 0; i < events.length; i ++) {

    var myformula_placeholder = '';
    // Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
    // NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
    var details=[mycal, events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent()];
    //var range=sheet.getRange(row,1,1,14);
    sheet.appendRow(details);
    var row = sheet.getLastRow();
    // Writing formulas from scripts requires that you write the formulas separate from non-formulas
    // Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
    var cell=sheet.getRange(row,7);
    cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))+IFERROR(DATEDIF(E' +row+ ',F' +row+ ',"d")*24)');
    cell.setNumberFormat('.00');
    
    cell=sheet.getRange(row,15);
    cell.setFormula('=trim(if(REGEXMATCH(B' +row+ ',"^Project: "),"Project", if(iserror(REGEXEXTRACT(B' +row+ ',".*\\|\\|+\\s([\\w\\s:]+)") = ""), if(iserror(REGEXEXTRACT(C' +row+ ',".*\\|\\|+\\s([\\w\\s:]+)")),iferror(VLOOKUP(B' +row+ ',Mapping!$A$2:$B$50,2,FALSE),""),REGEXEXTRACT(C' +row+ ',".*\\|\\|+\\s([\\w\\s:]+)")), REGEXEXTRACT(B' +row+ ',".*\\|\\|+\\s([\\w\\s:]+)"))))');
    cell=sheet.getRange(row,16);
    cell.setFormula('=TEXT(E'+row+'-WEEKDAY(E'+row+')+1,"yyyy-mm-dd")');
    cell=sheet.getRange(row,17);
    cell.setFormula('=if(iserror(REGEXEXTRACT(B'+row+',"^Project: (.*)")),"",REGEXEXTRACT(B'+row+',"^Project: (.*)"))')
    cell=sheet.getRange(row,18);
    cell.setFormula('=IFERROR(INDEX(SPLIT(O'+row+',":"),1,1),"")')
    cell=sheet.getRange(row,19);
    cell.setFormula('=IFERROR(JOIN(":",R'+row+',INDEX(SPLIT(O'+row+',":"),1,2)),R'+row+')')
  }
  
 
}


function getLastRow(){
  var sheet = SpreadsheetApp.openById(SPREADSHEET_APP_ID).getSheetByName(SPREADSHEET_NAME);
  var lastRow = sheet.getLastRow();
  Logger.log(lastRow);
  
  var cal = CalendarApp.getCalendarById(CALENDAR_ID);
  var events = cal.getEvents(new Date('Jan 1, 2019 00:00:00 EST'), new Date('Jan 31, 2019 23:59:59'));
  var len = events.length;
  Logger.log(len);
}
function isTimeUp(start) {
  var now = new Date();
  return now.getTime() - start.getTime() > 240000;
}
function DeleteAllRows() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_APP_ID).getSheetByName(SPREADSHEET_NAME);
  
  sheet.clearContents();
}
