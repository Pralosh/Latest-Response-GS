// Function to clear the Response Collection Sheet at a specified time

function tcs() {
  
  //delete the old triggers
  let scriptID = PropertiesService.getScriptProperties().getProperty("scriptID");
  if(scriptID != null) {
    let triggers = ScriptApp.getProjectTriggers();
    for(i = 0; i < triggers.length; i++) {
      if(triggers[i].getUniqueId() == scriptID) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  }

  //run the script daily at the given time (5:01PM)
  let addDays = 1;
  let thisDay = new Date();
  thisDay.setHours(17,1,0,0);
  var nextRunTime = futureRunTime(thisDay, addDays);

  console.log("Current Date:",new Date(), "Scheduled Date:", nextRunTime);

  //Need to set the last runtime previous day to run the function for the first time.
  //PropertiesService.getScriptProperties().setProperty('lastRunDate', new Date(2024,7,22).toDateString());

  var lastRunTime = PropertiesService.getScriptProperties().getProperty('lastRunDate'); //Format: Day Month Date Year

  console.log('Last Script Run Date:', lastRunTime);

  if(lastRunTime != new Date().toDateString()) {
    //set the current date as lastRuntime Format: Day Month Date Year
    PropertiesService.getScriptProperties().setProperty('lastRunDate', new Date().toDateString());

    console.log('New Script Run Date:', PropertiesService.getScriptProperties().getProperty('lastRunDate'));
    
    //create the time based trigger and set scriptID  to run the script
    let newTrigger = ScriptApp.newTrigger('tcs')
    .timeBased()
    .at(nextRunTime)
    .create()
    .getUniqueId();
    PropertiesService.getScriptProperties().setProperty("scriptID", newTrigger);

    console.log('New Time Based Trigger Created.');
    
    // get the sheet that needs to be cleared
    var sheetName = 'Latest Response Sheet';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    // clear sheet
    sheet.clear();

    //confirmation message
    console.log('Sheet cleared !');
  } else {
    console.log('Script Ran Once On:', lastRunTime);
    console.log('No New Time Based Trigger Created.');
  }
}

//returns the future runtime date
function futureRunTime(selectedDate, daysToAdd) {
  let nowTime = selectedDate.getTime();
  let add = 1000 * 60 * 60 * 24 * daysToAdd; //milliseconds * seconds * minutes * hours * days to add 
  let newTime = add + nowTime;
  let futureDate = new Date(newTime);
  return futureDate;
}
