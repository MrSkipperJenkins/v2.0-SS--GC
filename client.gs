// LIVE CLIENT MODULE SCRIPT VERSION 2.0 - inclusion of stored database and 
// compartmental code execution (i.e, background process schedule parsing on state change script + user triggered calendar updates script )   
// Director Schedule Spreadsheet to Google Calendar code written by Benjamin Bauer
// started: February 28, 2017
// released: 9/15/2017
// https://script.google.com/a/weathergroup.com/macros/s/AKfycbwwM_aKRxgPvxr-Pa3sIf_LrGG-VHo3pUFuPVgGY3YE7-i7IAhj/exec
// shortened URL: https://goo.gl/Yta73z
//
// https://script.google.com/a/weathergroup.com/macros/s/AKfycbyo3aVhuOHUZtuXNM-zfyFAMYGtoUAiqlZypwnmL4_9/dev

function myClientFunction() {
  var email = Session.getActiveUser().getEmail();
  var dates = scheduleDates(); //find start and end times to check calendar events
  var scheduleSheetId = PropertiesService.getScriptProperties().getProperty('databaseID'); //id of "Schedule Event List for each user"
  //var userCalendarEvents = grabUserCalendarEvents(dates.start, dates.end, email);
  
  //make sure database is on current week and grab it
//  if (!dbUpToDate(scheduleSheetId, dates.start)) {
//    Logger.log("Database not on current week, running myServerFunction()");
//    myServerFunction(); //run an update from SERVERModule library
//  } else {
//    Logger.log("Database is on current week."); //database is up to date, do nothing
//  }
    
  var ss = SpreadsheetApp.openById(scheduleSheetId);
  
  Logger.log("Reading schedule starting on: " + dates.start + " to: " + dates.end); 
  
  editUserCalendar(objChangeInspector(readUserSheet(ss, email), grabUserCalendarEvents(dates.start, dates.end, email))); //adds/removes events in user's calendar based on change list
  
  Logger.log("All done!");

  removeDuplicates(dates.start, dates.end, email);
                   
  return checkTrigger();
}


function removeDuplicates(start, end, email) {
  Logger.log("Checking for calendar event duplicates...");

  start.setDate((start.getDate() - 14));
  
  var events = grabUserCalendarEvents(start, end, email);
  
  var matches = [];
  var nonmatches = [];
   
  //routine 1 - start with calendar and look for any mismatches, return list of mismatches and duplicates
  for (var a = 0; a < (events.length - 1); a++) {
    var noMatchFound = true;

    for (var b = (a + 1); b < events.length; b++) {
      var match = true;
      
      for (var prop in events[a]) {
        
        if (prop != 'eventId') {

          if (!((events[a][prop] == events[b][prop]) && (match))) {
            match = false;
          }
        }
        
      }
    
      if (match) {
        Logger.log("Calendar event at " + a + " exactly matched event at " + b + ". I'll mark it for deletion.");
        Logger.log(events[a].title);
        matches.push(events[a]);
        noMatchFound = false;
        b = events.length; //escape the for loop, you have a match
      }      
    }  
    
    if (noMatchFound) { 
      Logger.log("Calendar event at " + a + " did not match anything. I'll leave it alone.");
      nonmatches.push(events[a]);
    } 
  }   
  
  //now take that list of matches and delete the events from calendar
  var calendar = CalendarApp.getDefaultCalendar();
  
  for (var taskIndex = 0; taskIndex < matches.length; taskIndex++) {
      Logger.log("Deleting event: " + matches[taskIndex].title + " from " + matches[taskIndex].start + " to " + matches[taskIndex].end);
      
      try {
        calendar.getEventSeriesById(matches[taskIndex].eventId).deleteEventSeries(); 
      } catch(e) {
        Logger.log('Could not delete task ' + taskIndex + '. ' + e);
        
        if (e == 'Exception: You have been creating or deleting too many calendars or calendar events in a short time. Please try again later.') {
          Logger.log('Waiting for 1 second..');
          Utilities.sleep(1000);
        }
        
      }
  }
}


//checks if schedule database has been updated to the current week (sometimes right after the week change it has not)
function dbUpToDate(id, start) {
  var db = DriveApp.getFileById(id) ;
  var lastUpdateDate = db.getLastUpdated().getDate();

  Logger.log("Database last updated on " + db.getLastUpdated());
  
  return (lastUpdateDate >= start.getDate());
}


//removes all 'On-Air |' events from default calendar within a given time range
function removeUserCalendarEvents() {
  var dates = scheduleDates(); //find start and end times to check calendar events
  var calendarEvents = CalendarApp.getDefaultCalendar().getEvents(dates.start, dates.end);
  var events = [];
  
  for (var a = 0; a < calendarEvents.length; a++) {
    if (calendarEvents[a].getTitle().search("On-­Air \\|") != -1) {  // uses two backslashes to escape the '|' character
      Logger.log((a+1) + ": " + calendarEvents[a].getTitle() + " event deleted.");
      calendarEvents[a].deleteEvent();    
    }
  }             
}


//grabs all 'On-Air |' events from default calendar within a given time range
function grabUserCalendarEvents(start, end, email) {
  var calendarEvents = CalendarApp.getDefaultCalendar().getEvents(start, end);
  var events = [];
  
  for (var a = 0; a < calendarEvents.length; a++) {
    if (calendarEvents[a].getTitle().search("On-­Air \\|") != -1) {  // uses two backslashes to escape the '|' character
      events.push(new eventFromSchedule(email, calendarEvents[a].getStartTime().valueOf(), calendarEvents[a].getEndTime().valueOf(), 
                                        calendarEvents[a].getTitle(), calendarEvents[a].getDescription(), 
                                        calendarEvents[a].getId()));
    }
  }             
  
  return events;
}


//returns dates starting with first sunday of current week and last saturday of three weeks ahead 
function scheduleDates() {
  var calStart = new Date();
  if (calStart.getDay() == 0) {
    calStart.setDate((calStart.getDate() - 6));
  } else {
    calStart.setDate((calStart.getDate() - calStart.getDay()) + 1);
  }

  calStart.setHours(0);
  calStart.setMinutes(0);
  
  var calEnd = new Date();
  calEnd.setDate((calEnd.getDate() + (21 + (7 - calEnd.getDay())))); //two weeks plus how many days left in week
  calEnd.setHours(23);
  calEnd.setMinutes(59);  
        
  return {start : calStart, end : calEnd}; 
}


//reads a specific sheet from a given spreadsheet into an eventFromSchedule object
function readUserSheet(ss, email) {
  var sheet = ss.getSheetByName(email); //grab user's sheet from passed spreadsheet
  var lastRow = sheet.getLastRow(); //TOO SLOW 
  
  var schedule = [];

  if (lastRow) { //if sheet not blank..
    var range = sheet.getRange(1,1,lastRow,6);
    var valueAt = range.getValues();
      
    for (var row = 0; row < lastRow; row++) {
      schedule.push(new eventFromSchedule(valueAt[row][0], valueAt[row][1].valueOf(), valueAt[row][2].valueOf(), valueAt[row][3], valueAt[row][4], valueAt[row][5], valueAt[row][6]));
    }
  } 
  
  Logger.log(email + ' schedule sheet read. ');
  
  return schedule;
}


//outputs events that do not match between two eventFromSchedule object sets
function objChangeInspector(database, calendar) {
  var matches = [];
  var nonmatches = [];
   
  //routine 1 - start with calendar and look for any mismatches, return list of mismatches and duplicates
  for (var a = 0; a < calendar.length; a++) {
    var noMatchFound = true;

    for (var b = 0; b < database.length; b++) {
      var match = true;
      
      for (var prop in calendar[a]) {
        
        if (prop != 'eventId') {

          if (!((calendar[a][prop] == database[b][prop]) && (match))) {
            match = false;
          }
        }
        
      }
    
      if (match) {
        Logger.log("Calendar event " + (a+1) + " matched a schedule event. I'll leave it alone.");
        matches.push(calendar[a]);
        noMatchFound = false;
      }      
    }  
    
    if (noMatchFound) { 
      Logger.log("Calendar event " + (a+1) + " did not match a schedule event. I'll add it to the list for deletion.");
      nonmatches.push(calendar[a]);
    } 
    
  }        
    
  //routine 2 - start with databse and look for any mismatches, return list of mismatches
  for (var a = 0; a < database.length; a++) {
    var noMatchFound = true;
    
    for (var b = 0; b < calendar.length; b++) {
      var match = true;
      
      for (var prop in database[a]) {
        if (prop != 'eventId') {
          if (!((database[a][prop] == calendar[b][prop]) && (match))) {
            match = false; 
          }
        }
      }
      
      if (match && (!noMatchFound)) {
        Logger.log("Duplicate calendar event detected at database event " + (a+1) + ". I'll add it to the list for deletion.");
        nonmatches.push(calendar[b]); //add this calendar event to return list 
      }
      
      if (match && noMatchFound) {
        Logger.log("Schedule event " + (a+1) + " matched a calendar event. I'll leave it alone.");
        matches.push(database[a]);
        noMatchFound = false;
      }
      
       
    }
    
    if (noMatchFound) { 
      Logger.log("Schedule event " + (a+1) + " did not match a calendar event. I'll it add to your calendar.");
      nonmatches.push(database[a]);
    } 
    
  }        
  
  return nonmatches;
}


//adds/removes events on user's default calendar according to given list of changes
function editUserCalendar(list) {
  var calendar = CalendarApp.getDefaultCalendar();
  
  for (var taskIndex = 0; taskIndex < list.length; taskIndex++) {
    Logger.log('Event id ' + taskIndex + ' is ' + list[taskIndex].eventId);
    
    if (list[taskIndex].eventId) {
      
      Logger.log("Event to be deleted: " + list[taskIndex].title + " from " + list[taskIndex].start + " to " + list[taskIndex].end);
      
      try {
        calendar.getEventSeriesById(list[taskIndex].eventId).deleteEventSeries(); 
      } catch(e) {
        Logger.log('Could not delete task ' + taskIndex + '. ' + e);
      }
       
    } else {
      var now = new Date();
      
      calendar.createEvent(list[taskIndex].title,
                           new Date(list[taskIndex].start),
                           new Date(list[taskIndex].end),
                           {description: list[taskIndex].shift});
      
      Logger.log("Event created: " + list[taskIndex].title + " from " + new Date(list[taskIndex].start) + " to " + new Date(list[taskIndex].end));
    }
  
    // if (list[taskIndex].start
  
  }
}