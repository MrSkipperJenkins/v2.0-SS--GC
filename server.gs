// LIVE SERVER SCRIPT VERSION 1.0 - inclusion of Google Sheet database and compartmental code execution (i.e, background process of schedule parsing + user triggered calendar updates)   
// "Director Schedule Spreadsheet to Google Calendar" written by Benjamin Bauer
// started: February 28, 2017


function myServerFunction() {
  var scheduleSheetId = PropertiesService.getScriptProperties().getProperty('databaseID'); //id of "Schedule Event List for each user"
  var ss = SpreadsheetApp.openById(scheduleSheetId);
  var sheets = getScheduleSheets();
  
  //var liveSchedule = parseSchedule(sheets.thisWeek).concat(parseSchedule(sheets.nextWeek), parseSchedule(sheets.thirdWeek));

  writeAllToNamedDB(parseSchedule(sheets.thisWeek).concat(parseSchedule(sheets.nextWeek), parseSchedule(sheets.thirdWeek)), ss);
  
}


//object definition of "eventFromSchedule"
function eventFromSchedule(user, start, end, title, shift, eventId) { //object containing event info to be created for each day
    this.user = user;  
    this.start = start; 
    this.end = end; 
    this.title = title;
    this.shift = shift;
    this.eventId = eventId;
}


//clears values of sheets in supplied spreadsheet
function clearAllSheets(ss) {
  var sheet = ss.getSheets();
  for (var a = 0; a < sheet.length; a++) {
    sheet[a].clearContents();
  }
}


//writes all user's schedule events to the supplied spreadsheet
function writeAllToNamedDB(sched, ss) {
  
  //convert schedule object/array to array/array
  var convrt = [];
  
  for (var a = 0; a < sched.length; a++) {
    convrt[a] = [];
    
    for (var prop in sched[a]) {
      convrt[a].push(sched[a][prop]); 
    }
  }
  
  //sort convrt array by name
  convrt.sort();
  
  //clear all sheets
  clearAllSheets(ss);
  
  //go through each line and write to db sheet
  for (var b = 0; b < convrt.length; b++) {
    var name = convrt[b][0];
    var nextName = name;
    var currentNamedSet = [];
    
    do {
      
      currentNamedSet.push(convrt[b]);
            
      if (convrt[(b+1)]) { //check if there is another name to get
        nextName = convrt[(b+1)][0]; 
      
        Logger.log(convrt[b][0] + convrt[b][1] + b);
        
      } else { nextName = ''}; 
      
      if (name == nextName) { b++ };
      
    } while (name == nextName);
        
    var sheet = ss.getSheetByName(name); //grab stored sheet from spreadsheet passed
    
    
    var range = sheet.getRange(1,1,currentNamedSet.length,6);
    range.setValues(currentNamedSet); //writes all values
    range.sort(2);                    //sort by start time
  }
}


function parseSchedule(sht) {
  
  var schedule = [];
  
  //get sheet for each day of week
  var spreadsheet = SpreadsheetApp.open(sht);
  var DaysOfWeek = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"];
  
  //Assign current work day
  for (var x = 0; x < 7; x++) {
    var sheet = spreadsheet.getSheetByName(DaysOfWeek[x]); //grab sheet for that day              
    Logger.log(DaysOfWeek[x]);
    
    //spreadsheet variables
    var lastRow = 60; //sheet.getLastRow()+1; //TOO SLOW 
    var lastColumn = 30;//sheet.getLastColumn(); //TOO SLOW 
    var range = sheet.getRange(1,1,lastRow,lastColumn);
    var valueAt = range.getValues();
 
    for (var Worker = 3; Worker < lastColumn; Worker++) {
      var firstlastName = valueAt[2][Worker];
                  
      if (firstlastName != false) {
        if (firstlastName.trim()) { //is field blank or false? then skip everything go to next column
          
          if (valueAt[4][Worker]) { //get shift time and reformat
            var shift = getShiftTime(valueAt[4][Worker]);
          }  
          
          for (var i = 5; i < lastRow; i++) { //get all jobs for each worker
            var ECount = 1; 
     
            //if job value found this create/modify events, if not delete calendar events if found, ignore white space // check if PTO or OPH event
            if ((valueAt[i][Worker].trim()) && (valueAt[i][Worker].search('PTO') == -1) && (valueAt[i][Worker].search('OPH') == -1)) {  
             
              for (var u = i; u < lastRow; u++) { //find show names       
                var Job = valueAt[u][Worker].trim();
                var NextJob = valueAt[u+1][Worker].trim();        
                
                //find show name
                if (valueAt[u][1]) {         
                  var Show = valueAt[u][1].trim();
                }
                else {
                  for (var j = u-1; j > 4; j--) {
                    if (valueAt[j][1]) {
                      var Show = valueAt[j][1].trim();
                      j = 4;
                    }
                  }
                }  
                
                //find NEXT show name
                var NextShow = valueAt[u+1][1] ? valueAt[u+1][1] : Show;
                
                //decide whether to combine into one event or start a new one
                if (((Job==NextJob) && (Show==NextShow)) || (((Job ==  'PRE-PRO') || (Job ==  'PREPRO')) && (Job==NextJob))) {
                  ECount += 1;}
                else {
                  var CurrentRow = u;
                  u = lastRow;
                }  
              }
              
              if ((Job == 'BWTM') || (Job ==  'PRE-PRO') || (Job ==  'PREPRO')) {
                var newEventTitle = 'On-Air | ' + Job;
              } else {
                var newEventTitle = 'On-Air | ' + Show + ": " + Job;
              }
              
              var newEventStart = findShowStartTime(valueAt, i);
              
              var newEventEnd = new Date(newEventStart);
              newEventEnd.setMinutes(newEventEnd.getMinutes()+(ECount*30)); //multiply by the amount of extra blocks of :30
              var eventshift = 'In time: ' + shift.inTime + '\nOut time: ' + shift.outTime;
              
              
              //put into object array
              schedule.push(new eventFromSchedule(getEmailFromScheduleName(firstlastName), newEventStart, newEventEnd, newEventTitle, eventshift, '')); 
              Logger.log(firstlastName + newEventStart + newEventEnd + newEventTitle + eventshift);
              
              i = CurrentRow; //update to latest row
            }
          }
        } else { Logger.log('No worker found on this day.')} 
      } else { Logger.log('No worker found on this day.')}
    }
  }
  
  return schedule;
}


function getEmailFromScheduleName(name) {
  switch (name) {
    case "Ben Bauer":
      return "benjamin.bauer@weathergroup.com";
      break;
     
    case "Ryan Bowles":
      return "ryan.bowles@weathergroup.com";
      break;
      
    case "Brian Bruck":
      return "brian.bruck@weathergroup.com";
      break;
  
    case "Edward Bruno-Gaston":
      return "edward.bruno-gaston@weathergroup.com";
      break;
    
    case "Fatinah Chen":
      return "fatinah.chen@weathergroup.com";
      break;
    
    case "Pierce Gossett":
      return "pierce.gossett@weathergroup.com";
      break;  
      
    case "Chip Hirzel":
      return "conrad.hirzel@weathergroup.com";
      break;  
      
    case "AJ Keane":
      return "aj.keane@weathergroup.com";
      break;  
    
    case "Chris McDaniel":
      return "chris.mcdaniel@weathergroup.com";
      break;
      
    case "Matthew Miller":
      return "matthew.miller@weathergroup.com";
      break;
      
    case "Kevin Sheridan":
      return "kevin.sheridan@weathergroup.com";
      break;  
      
    case "Robb Williams":
      return "robb.williams@weathergroup.com";
      break;  
      
    case "Tom Williams":
      return "tom.williams@weathergroup.com";
      break;  
      
    case "Matt Wooley":
      return "matthew.wooley@weathergroup.com";
      break;  
  }
}


//outputs formatted shift time info from a given spreadsheet value
function getShiftTime(val) {
  //var value = '630A-230P';
  var shift = val.trim().toUpperCase().split('-'); //remove possible whitespace / split into two different strings at the "-"
  
  var inAMPM = shift[0].match(/(A|P)$/); //reg exp for finding letters A or P in any case
  
  if (inAMPM) {
    var inAMPM = inAMPM[0];
    var inTime = shift[0].slice(0,shift[0].length-1);
  } else {
    var inAMPM = 'A';
    var inTime = shift[0];
  }
  
  if (shift[1]) { //check to see if there even is a hyphen, if not return blanks of in/out times
    var outAMPM = shift[1].match(/(A|P)$/);
  } else {return {inTime: '', outTime: ''};} 

  if (outAMPM) {
    var outAMPM = outAMPM[0];
    var outTime = shift[1].slice(0,shift[1].length-1);
  } else {
    var outAMPM = 'A';
    var outTime = shift[1];
  }
  
  if (inTime.length > 2) {
    var inMins = inTime.slice(-2);
    if (inTime.length == 3) {
      var inHrs = inTime.slice(0,1);
    } else {
      var inHrs = inTime.slice(0,2);
    }    
  } else {
    var inHrs = inTime;
    var inMins = 0;
  }
  
  if (outTime.length > 2) {
    var outMins = outTime.slice(-2);
    
    if (outTime.length == 3) {
      var outHrs = outTime.slice(0,1);
    } else {
      var outHrs = outTime.slice(0,2);
    }    
  } else {
    var outHrs = outTime;
    var outMins = 0;
  }
  
  return {inTime: (inTime + inAMPM), outTime: (outTime + outAMPM)};
}


//finds show time from a spreadsheet range and given row  
function findShowStartTime(val, i) {

  //var Today = ; //from same place in every sheet
  var start = new Date(val[1][5]);
  var jobTime = val[i][0]; //value for Time
  var res = jobTime.split(':'); //split into two different strings at the ":"
  var hrs = parseInt(res[0], 10); //+3 to make spreadsheet correct, take out +3 to make calender correct
  var mins = parseInt(res[1].slice(0,res[1].length-1), 10); //strip "A" or "P" off end AND parse into integer
  
  //check to see if AM/PM, change hours accordingly 
  if ((jobTime.search('A') == -1) && (hrs < 12)) { //set PM hours to 24
    hrs += 12; //PM offset +3 for some reason.. GMT to EST?
  } 
  
  if ((jobTime.search('A') != -1) && (i > 45)) { //set 12A onward down sheet to next day
    start.setDate(start.getDate()+1);
    hrs = hrs == 12 ? 0 : hrs; 
  }
  
  start.setHours(hrs);
  start.setMinutes(mins);

  return start;
}


//gets this and next week's director's schedule spreadsheets
function getScheduleSheets() {
  
// Find parent 2017 Schedule Folder
//  var AllFolders = DriveApp.searchFolders('title contains "2017 Production Control Schedules"');
//  while (AllFolders.hasNext()) {
//    var PCSfolder = AllFolders.next();
//    Logger.log(PCSfolder.getName());
//    Logger.log(PCSfolder.getId()); // 0B6BXSE7lNq7FNXdxaGRVQzZ1VjQ
//  }
//  
//  
//  //Find child Director Schedule Folder
//  var dirFolder = PCSfolder.searchFolders('title contains "DIRECTOR SCHEDULES"');
//  while (dirFolder.hasNext()) {
//    var Dfolder = DirFolder.next();
//    Logger.log(Dfolder.getName());
//    Logger.log(Dfolder.getId()); // 0B6BXSE7lNq7FX1dsM3RLRHBaRTQ
//  }
  
  //skip all that use exact id for Director Schedules folder
  var dirFolderID = PropertiesService.getScriptProperties().getProperty('directorFolderID'); //id of "DIRECTOR SCHEDULES" folder on shared google drive
  var dirFolder = DriveApp.getFolderById(dirFolderID);
  
  //Find child month folder
  var Today = new Date();
  
  //Get current month in string form
  var MonthOfYear = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
  var Month = MonthOfYear[Today.getMonth()]; //Get string form of month
  var LastMonth = (Month == 0) ? 11 : MonthOfYear[Today.getMonth()-1];
  var NextMonth = MonthOfYear[Today.getMonth()+1]; // THIS NEEDS UPDATING FOR YEAR TO YEAR FUNCTIONALITY 
  Logger.log('This month: ' + Month)
  Logger.log('Last month: ' + LastMonth);
  Logger.log('Next month: ' + NextMonth);

  var thisMonthFolder = dirFolder.searchFolders('title contains "'+ Month +'"');
  while (thisMonthFolder.hasNext()) {
    var folder = thisMonthFolder.next();
    Logger.log(folder.getName());
  }
  
    //Find this and next week's sheets
  var Day = Today.getDate(); //returns day of month int betw 1 - 31
  var folderSheets = folder.getFiles();
  
  while (folderSheets.hasNext()) {
    var Sheets = folderSheets.next();
    var WeekOf = Sheets.getName().toUpperCase();
    var W = WeekOf.search(Month) + 1;
    var WeekOfNum = WeekOf.slice(W+Month.length);
    
    Logger.log('Week of Number:' + WeekOfNum);
    
    if (((Day-WeekOfNum) < 7) && ((Day-WeekOfNum) >= 0)) {
      var ThisWeekSheet = Sheets;    
      Logger.log('This weeks calender: ' + ThisWeekSheet);
    }
    if ((Math.abs(Day-WeekOfNum) <= 7) && ((Day-WeekOfNum) < 0)) {
      var NextWeekSheet = Sheets; 
      Logger.log('Next weeks calender: ' + NextWeekSheet);
    }
    if ((Math.abs(Day-WeekOfNum) > 7) && (Math.abs(Day-WeekOfNum) <= 14) && ((Day-WeekOfNum) < 0)) {
      var ThirdWeekSheet = Sheets; 
      Logger.log('Third weeks calender: ' + ThirdWeekSheet);
    }
  }  

  //find this week's sheet if we need to go into last month
  if (!ThisWeekSheet) {
    Logger.log('This week not found, searching last months folder..');
    var lastMonthFolder = dirFolder.searchFolders('title contains "'+ LastMonth +'"');
    while (lastMonthFolder.hasNext()) {
      var folder = lastMonthFolder.next();
      Logger.log(folder.getName());
    }
    
    var folderSheets = folder.getFiles();
    var WeekOfNumCheck;
    
    while (folderSheets.hasNext()) {
      var Sheets = folderSheets.next();
      var WeekOf = Sheets.getName().toUpperCase();
      var W = WeekOf.search(LastMonth) + 1;
      var WeekOfNum = WeekOf.slice(W+LastMonth.length);
      
      if (!WeekOfNumCheck) {WeekOfNumCheck = parseInt(WeekOfNum, 10)}; //grabs first WeekOf number 
      
      if (WeekOfNumCheck <= WeekOfNum) {ThisWeekSheet = Sheets}; //checks against all others and returns largest
    }
    
    Logger.log('This is the last week of last month: ' + ThisWeekSheet);      
  }
  
  //find next and third week's sheet if we need to go into next month 
  if (!NextWeekSheet) { 
    Logger.log('Next week not found, searching next months folder..');
    var nextMonthFolder = dirFolder.searchFolders('title contains "'+ NextMonth +'"');
    while (nextMonthFolder.hasNext()) {
      var folder = nextMonthFolder.next();
      Logger.log(folder.getName());
    }  

    var folderSheets = folder.getFiles();
    
    while (folderSheets.hasNext()) {
      var Sheets = folderSheets.next();
      var WeekOf = Sheets.getName().toUpperCase();
      var W = WeekOf.search(NextMonth) + 1;
      var WeekOfNum = WeekOf.slice(W+NextMonth.length);
      Logger.log(WeekOfNum);
      if (WeekOfNum <= 7) {
        var NextWeekSheet = Sheets;    
        Logger.log('This is the first week of next month for second sheet: ' + NextWeekSheet);
      }
      if ((WeekOfNum > 7) && (WeekOfNum <= 14)) {
        var ThirdWeekSheet = Sheets;    
        Logger.log('This is the second week of next month for third sheet: ' + ThirdWeekSheet);
      }
    }      
  }  
  
  //find third week's sheet if we need to go into next month 
  if (!ThirdWeekSheet) {
    Logger.log('Third week not found, searching next months folder..');
    var thirdWeekFolder = dirFolder.searchFolders('title contains "'+ NextMonth +'"');
    while (thirdWeekFolder.hasNext()) {
      var folder = thirdWeekFolder.next();
      Logger.log(folder.getName());
    }  

    var folderSheets = folder.getFiles();
    
    while (folderSheets.hasNext()) {
      var Sheets = folderSheets.next();
      var WeekOf = Sheets.getName().toUpperCase();
      var W = WeekOf.search(NextMonth) + 1;
      var WeekOfNum = WeekOf.slice(W+NextMonth.length);
      
      if (WeekOfNum <= 7) {
        var ThirdWeekSheet = Sheets;    
        Logger.log('This is the first week of next month for the third sheet: ' + ThirdWeekSheet);
      }
    }      
  }
  
  function calendarStart(sht) {
    var sheet = SpreadsheetApp.open(sht);
    var calStart = new Date(sheet.getSheetByName("MONDAY").getRange(2,6,1,1).getValues());
    
    calStart.setHours(0);
    calStart.setMinutes(0)
    
    return calStart;  
  }

  function calendarEnd(sht) {
    var sheet = SpreadsheetApp.open(sht);
    var calEnd = new Date(sheet.getSheetByName("SUNDAY").getRange(2,6,1,1).getValues());
    
    calEnd.setHours(23);
    calEnd.setMinutes(59);
    
    return calEnd; 
  }
  
  return {thisWeek: ThisWeekSheet, nextWeek: NextWeekSheet, thirdWeek: ThirdWeekSheet, calendarStart: calendarStart(ThisWeekSheet), calendarEnd: calendarEnd(ThirdWeekSheet)};
}



