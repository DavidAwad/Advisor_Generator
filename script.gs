/* 
@Author - David Awad

Some Template Info
==================
Times start at D23

D15,E15 invoice start, end dates (End date is static, filled in by you)

C23 Shift Invoice Dates Start

To add this script, open up the invoice template

tools > 

*/

function getcalendars(){
  Logger.info("getting calendars")
  var calendars  = CalendarApp.getAllCalendars();
  return calendars
}
 
function getss(){ 
  Logger.info("getting spreadsheet")
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss
}

/*
NUM_HOURS Function, simply parses a shift and returns the number of hours

Applied to D23 and written to E23
*/
function NUM_HOURS(input) {
  if(input==="8AM-10AM" || input === "10PM-12AM"){
    return 2
  }
  else if (input==="10AM-2PM" || input === "2PM-6PM" || input === "6PM-10PM"){
    return 4
  }
  else{
    return "invalid date? double check script"
  }
}

/*
subtract number of days from a date

int d - number of day ro substract and date = start date
Date date - date object to subtract d days from.
*/
function subDaysFromDate(date,d){
  var result = new Date(date.getTime() - d*(24*3600*1000));
  return result
}

/*
print the calendar names so you know which entry is your Codecademy calendar
*/
function log_calendars(){
  for (var i = 0; i < calendars.length; i++){
    Logger.info(calendars[i].getName())
  }
}
 

/*
converts tracksmart representation of shift to invoice rep.
*/
function conv_shift(shift){
  switch(shift) {
    case "8-10":
      return "8AM-10AM";
    case "10-2":
        return "10AM-2PM";
    case "2-6":
      return "2PM-6PM";
    case "6-10":
      return "6PM-10PM";
    case "10P-12A":
      return "10PM-12AM";
    default: 
      return "error?"
  }
}


/* 
iterate through CC calendar and write shifts
int index - an integer to specify which shift to return
*/
function WRITE_SHIFT_TIME_INDEX(){
  
  // get CC calendar
  var calendars  = CalendarApp.getAllCalendars();
  var CC = calendars[3]; // my codecademy schedule in the array of calendars I'm subscribed to
    
  // access spreadsheet app
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  
  Logger.info(ss.getName());
  // Logger.info(ss.getName()); FIXME, grab end date from sheet name
  
  var sheet = ss.getSheetByName("Invoice"); // get the sheet this script is attached to 
  var range = sheet.getRange(15,5);  // get end date (static value, written in cell E15)
  var end_date = range.getValue();

  // TODO watch inclusive bounds?? 
  // find date from 2 weeks before, for the proper date range
  var start_date = subDaysFromDate(end_date, 14);
  
  // get all events in this date range
  var events = CC.getEvents(start_date, end_date); // grab the range of the last two weeks
  
  
  for(var i = 0; i < events.length; i++){  
    var curr = events[i];
    // get event title, remove the `(HS)` characters
    var title = curr.getTitle().substring(0, curr.getTitle().length - 5);
    // Logger.info(title);
    Logger.info(conv_shift(title))
    
    range = sheet.getRange(23 + i, 4);
    range.setValue(conv_shift(title));
  }
}


function WRITE_SHIFT_DATE_INDEX(){
  // get CC calendar
  var calendars  = getcalendars()
  var CC = calendars[3]; // my codecademy schedule in the array of calendars I'm subscribed to
    
  // access spreadsheet app
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
   
  var sheet = ss.getSheetByName("Invoice"); // get the sheet this script is attached to 
  var range = sheet.getRange(15,5);  // get end date (static value, written in cell E15)
  var end_date = range.getValue();
  
  // find date from 2 weeks before, for the proper date range
  var start_date = subDaysFromDate(end_date, 14);
  
  // get all events in this date range
  var events = CC.getEvents(start_date, end_date ); // grab the range of the last two weeks
  
  for(var i = 0 ; i < events.length; i++){
    var curr = events[i].getStartTime(); 
    var ret =  (curr.getMonth() + 1) + "/" + curr.getDate() + "/" + curr.getFullYear() ;
    Logger.info(ret)
    range = sheet.getRange(23 + i, 3);
    range.setValue(ret);
  }
}
