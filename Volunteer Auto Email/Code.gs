// Create UI element
function onOpen() { 
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Script")
    .addItem("Run Now","AutoEmailer")
    .addItem("Create Project Triggers","MakeTrigger")
    .addItem("Clear Project Triggers","ClearTriggers")
    .addToUi();
}


// Weekly Trigger on Thursday at 4:00 p.m.
function MakeTrigger() {
  ScriptApp.newTrigger('AutoEmailer')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .atHour(16)
      .create();
}


function ClearTriggers() {
  Logger.log('Current project has ' + ScriptApp.getProjectTriggers().length + ' triggers.');
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


function AutoEmailer() {
   // Get data from the volunteer email spreadsheet
   var sheet = SpreadsheetApp.getActiveSheet();
   var data = sheet.getDataRange().getValues();
   
   // Get data from the volunteer & DFS appointment calendars
   var volCal = CalendarApp.getCalendarById('DOMAIN.org_ID@group.calendar.google.com');
   var aptCal = CalendarApp.getCalendarById('DOMAIN.org_ID@group.calendar.google.com');
   var now = new Date();
   var wkStart = new Date(now.getTime() + (86400000 * 1));
   var wkEnd = new Date(now.getTime() + (86400000 * 17));
   var volEvents = volCal.getEvents(wkStart, wkEnd);
   var aptEvents = aptCal.getEvents(wkStart, wkEnd);
   
   // Create arrays for storing data
   var volArray = getVols(data);
   var volSkdArr = getVolSkd(volEvents);
   var aptSkdArr = getAptSkd(aptEvents);
   
   
   // Build Email 
   for (var i = 1; i < data.length; i++) {
      var tempName = volArray[i].Name;
      var tempSkdName = volArray[i].ScheduleName;
      var tempEmail = volArray[i].Email;
      var count = 0;
      var apt = false;
      var subject = "DFS Schedule " + wkStart.toLocaleDateString() + " - " + wkEnd.toLocaleDateString();
      var toLine = ("Hello " + tempName + ",");
      // Template & Variables for HTML Body of Template_Apt
      var template_apt = HtmlService.createTemplateFromFile('Template_Apt');
      template_apt.toLine = toLine;
      template_apt.tempName = tempName;
      template_apt.tempSkdName = tempSkdName;
      template_apt.tempEmail = tempEmail;
      // Template & Variables for HTML Body of Template_noApt
      var template_noApt = HtmlService.createTemplateFromFile('Template_NoApt');
      template_noApt.toLine = toLine;
      template_noApt.tempName = tempName;
      template_noApt.tempSkdName = tempSkdName;
      template_noApt.tempEmail = tempEmail;
      
      // Check to see if someone is suiting in the next two weeks and sets variables so that they get the right email
      for (w = 0; w < volSkdArr.length; w++) {
        if (volSkdArr[w].EventTitle == tempSkdName) {
          count++
        }
        else {
          count += 0
        }
      }
      if (count >= 1) {
        apt = true;
      }
       
      
      if (apt == true) {
         var message = template_apt.evaluate();
         MailApp.sendEmail(tempEmail, 
                           subject, 
                           message.getContent().replace(/ CDT/g, ""), {
                             htmlBody: message.getContent().replace(/ CDT/g, ""),
                             replyTo: 'NAME@DOMAIN.org, NAME@DOMAIN.org, NAME@DOMAIN.org'
                           });
      }
      else {
         var message_alt = template_noApt.evaluate();
         MailApp.sendEmail(tempEmail, 
                           subject, 
                           message_alt.getContent(), {
                             htmlBody: message_alt.getContent(),
                             replyTo: 'NAME@DOMAIN.org, NAME@DOMAIN.org, NAME@DOMAIN.org'
                           });
      }
      
   }
}

function getVols(data) { // Fill volunteer array
  var volArray = [];
  for (i = 0; i < data.length; i++) { 
     var id = data[i][0];
     var name = data[i][1];
     var email = data[i][2];
     var skdName = data[i][3];
     var vol = {
       ID: id,
       Name: name,
       Email: email,
       ScheduleName: skdName,
     };
     volArray[i] =  vol;  
   }
   return volArray;
}

function getVolSkd(volEvents) { // Fill volunteer schedule array
  var volSkdArr = [];
  for (var i = 0; i < volEvents.length; i++) { 
     var volEventTitle = volEvents[i].getTitle();
     var volEventId = volEvents[i].getId();
     var volEventStart = volEvents[i].getStartTime();
     var volEventEnd = volEvents[i].getEndTime();
     var volInst = {
       EventTitle: volEventTitle,
       EventID: volEventId,
       EventStart: volEventStart,
       EventEnd: volEventEnd
     };
     volSkdArr[i] = volInst;
   }
   return volSkdArr;
}

function getAptSkd(aptEvents) { // Fill DFS appointment schedule array
  var aptSkdArr = [];
  for (i = 0; i < aptEvents.length; i++) { 
     var aptEventTitle = aptEvents[i].getTitle();
     var aptEventId = aptEvents[i].getId();
     var aptEventStart = aptEvents[i].getStartTime();
     var aptEventEnd = aptEvents[i].getEndTime();
     var aptInst = {
       EventTitle: aptEventTitle,
       EventID: aptEventId,
       EventStart: aptEventStart,
       EventEnd: aptEventEnd
     };
     aptSkdArr[i] = aptInst;
   }
   return aptSkdArr;
}
