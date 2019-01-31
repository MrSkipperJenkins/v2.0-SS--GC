//checks to see if user has specfically installed a trigger for the myClientFunction()
function checkTrigger() {
  try {
    var triggers = ScriptApp.getProjectTriggers()
  }  
  catch(e) {
    Logger.log("Error thrown. Will try again next time around..");
    return; // if error, then try again in a few minutes when checkTrigger runs again
  }
    
  Logger.log("I've found a total of " + triggers.length + " trigger(s) currently installed.");
  
  if (triggers[0]) {
    for (var i = 0; i < triggers.length; i++) {
      var triggerFunction = triggers[i].getHandlerFunction();
      Logger.log("Trigger " + (i+1) + ": " + triggerFunction);
      
      if (triggerFunction == "myClientFunction") {
        Logger.log("myClientFunction script trigger is installed")
        return true;
      }
    }
    
    Logger.log("myClientFunction script trigger not installed");
    return false;
  
  } else {
    Logger.log("myClientFunction script trigger not installed");
    return false;
  }
}


//removes all myClientFunction triggers and adds a single time based trigger
function refreshTrigger() {
  removeTrigger();
  addTimeTrigger();
}


// Deletes all triggers in the current project.
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


function setInstalledFlag(flg) {
  var scriptProperties = PropertiesService.getScriptProperties();  
  var email = Session.getActiveUser().getEmail();
  
  if (flg) {
    scriptProperties.setProperty(email, 'Installed.');
    Logger.log(email + " property set to Installed.");
  } else {
    scriptProperties.setProperty(email, 'Not installed.');
    Logger.log(email + " property set to Not installed.");
  }
}


//adds a time based trigger to the function myClientFunction()
function addTimeTrigger() {
  ScriptApp.newTrigger("myClientFunction")
  .timeBased()
  .everyMinutes(10)
  .create();
  
  Logger.log("myClientFunction() trigger added");
  
  setInstalledFlag(true);
  
  return checkTrigger();
}


//removes user's myClientFunction() trigger
function removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers() 
  Logger.log("I've found a total of " + triggers.length + " trigger(s) currently installed.");

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "myClientFunction") {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("myClientFunction() trigger found and removed.");
    }
    
    if (triggers[i].getHandlerFunction() == "myFunction") {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("myFunction() trigger found and removed.");
    }
  }

  setInstalledFlag(false);
  return false;
}


//grabs user's email
function getEmail() {
  return Session.getActiveUser().getEmail();
}


//required to render html from js/gs
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function isAuthorized() {
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  Logger.log(authInfo.getAuthorizationUrl());
  
  if (authInfo.getAuthorizationStatus() == ScriptApp.AuthorizationStatus.REQUIRED) {
    Logger.log("User not authorized.");
    return false;
  } else {
    Logger.log("User is authorized.");
    return true;
  }
}