
function onSubmit(e) {
  
  var errorFlag = 0;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('Could not obtain lock after 10 seconds.');
    console.log('Could not obtain lock after 10 seconds.');
  }
  
  //var formDestId = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById("1OyywoBuYZbjRP-Or6WdddVw3gI4XNfyzP3XyAeYdrws");
  
  SpreadsheetApp.setActiveSpreadsheet(ss);
  shFormResponses = ss.getSheetByName('Form Responses 1');
  
  var responses = FormApp.getActiveForm().getResponses();
  
  var length = responses.length;
  
  var lastResponse = responses[length-1];
  var formValues = lastResponse.getItemResponses();
  var length = lastResponse.getItemResponses().length;
  var formFindValue = formValues[0].getResponse();
  var priority = 0;
  
  Logger.log("length = " + length + ". lastResponse = " + lastResponse + ". formValues = " + formValues + ". formFindValue = " + formFindValue);
  
  
  
  Logger.log("Form find value: " + formFindValue);
  //Logger.log("Form destination id : " + FormApp.getActiveForm().getDestinationId());
  
  var currentRow, tktSubmitterEmail;
  
  if(formFindValue != "AM-SE") {
    var values = shFormResponses.getDataRange().getValues();
    //Logger.log("Form find for total values: " + values);
    for(var j=0, jLen=values.length; j<jLen; j++) {
      for(var k=0, kLen=values[0].length; k<kLen; k++) {
        var find = values[j][k];
        //Logger.log(find);
        if(find == formFindValue) {
          //Logger.log([find, "row "+(j+1)+"; "+"col "+(k+1)]);
          currentRow = j+1;
          break;           
        }
      }
    }
  }
  
  ss.setActiveSheet(shFormResponses);
  
  var ticketNumber;
  var ticketCounter;
  var row;
  var etrolControlsServiceEmail1, etrolControlsServiceEmail2, etrolControlsServiceEmail3 = null, timestamp, ticketNumberLocation;
  var editFlag = 0;
  var numberAsNumber2;
  var caseType;
  var meetingDate;
  var meetingTime;
  var meetingFinish;
  var duplicateCaseNum;
  var missingDuplicateFlag = 0;
  var months = {
      '1': "January", 
      '2': "February", 
      '3': "March", 
      '4': "April", 
      '5': "May", 
      '6': "June", 
      '7': "July", 
      '8': "August", 
      '9': "September", 
      '10': "October", 
      '11': "November", 
      '12': "December"
  };
    
  //var responses = FormApp.getActiveForm().getResponses();
  //var responseId = responses[0].getRespondentEmail();
  
  if(currentRow) {
    
    tktSubmitterEmail = shFormResponses.getRange(currentRow, 2).getValues();
    
    Logger.log("1. Submitter email is " + tktSubmitterEmail + " currentRow " + currentRow);
  }
  
  /* Dealing with invalid entries */
  if(!currentRow && (formFindValue != "AM-SE")) /* We couldn't find your case so don't let them submit */
  {
    tktSubmitterEmail = e.response.getRespondentEmail();
    
    
    Logger.log("Invalid entry " + formFindValue + "getEditors returned" + tktSubmitterEmail);
    errorFlag = 2;
  }
  
  if(currentRow && (formFindValue != "AM-SE")) /* Email mismatch between submitter and Account manager */
  {
    tktSubmitterEmail = e.response.getRespondentEmail();
    Logger.log("Ticket submitter email is: " + tktSubmitterEmail + " and the original submitter email is: " + shFormResponses.getRange(currentRow, 1).getValues().toString());
    
    var tktSEemail = null;
    tktSEemail = shFormResponses.getRange(currentRow, 7).getValues().toString();
    
    if (tktSEemail) {   /* Checking if SE is changing tkt status */
      if (tktSubmitterEmail != 'jay.bhaskar@gmail.com' && tktSubmitterEmail != 'alarkin@fastly.com' && tktSubmitterEmail != tktSEemail && tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      //if (tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      {
        Logger.log("Not your ticket to edit " + formFindValue);
        errorFlag = 1;
      }
    } else {
      if (tktSubmitterEmail != 'jay.bhaskar@gmail.com' && tktSubmitterEmail != 'alarkin@fastly.com' && tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      //if (tktSubmitterEmail != shFormResponses.getRange(currentRow, 1).getValues().toString())
      {
        Logger.log("Not your ticket to edit " + formFindValue);
        errorFlag = 1;
      }
    }
  }
  
  if((currentRow === undefined || currentRow === null) && formFindValue == "AM-SE") { /* This is a new entry */
    /* Checking if you're allowed to pick the A/c manager */
    tktSubmitterEmail = e.response.getRespondentEmail();
    if (tktSubmitterEmail != 'jay.bhaskar@gmail.com' && tktSubmitterEmail != 'alarkin@fastly.com' && tktSubmitterEmail != formValues[3].getResponse()) // The index for formValues will change depending on whether
    //if (tktSubmitterEmail != 'jay.bhaskar@gmail.com' && tktSubmitterEmail != formValues[3].getResponse()) // The index for formValues will change depending on whether
    {
      errorFlag = 3; 
    }
  }  
  
  
  if(currentRow && (formFindValue != "AM-SE")) /* Extract portion after AM-SE and validate value -- Not implemented yet*/
  {
    var localValue = formFindValue[0];
  }
  
  /* Dealing with invalid entries ends */
  
  if(errorFlag == 0) {
    
    if(currentRow && (formFindValue != "AM-SE") ) { /* 0. Someone is editing an existing response so get ticket number */
      
      Logger.log("formFindValue" + formFindValue);
      ticketNumber = shFormResponses.getRange(currentRow, 2).getValues();
      row = currentRow;
      ticketCounter = currentRow;
      editFlag = 1; /* Not being used as of now */
      
    }
    
    var tktSheet = ss.getSheetByName('Maintenance');
    
    
    //if(currentRow && (formFindValue == "AM-SE")){
    if((currentRow === undefined || currentRow === null) && formFindValue == "AM-SE"){ /* This is a new entry */
      
      //Logger.log("Looking to get the ticket counter from sheet");
      ticketCounter = tktSheet.getRange(2, 1).getValue();
      
    }
    
    //Logger.log("Ticket counter at:" + ticketCounter); /* 2. Just finding out if ticket counter was initialized and is rolling */
    
    if (ticketCounter === undefined || ticketCounter === null || ticketCounter === NaN || ticketCounter == 0) { /* 3. This is the first time ticket counter is being initialized */
            
      ticketCounter = '1';
      
      tktSheet.getRange(2, 1).setValue(ticketCounter);
    }
    
    /* {  5. Ticket number has been defined but possible that this is the new ticket */
    
    //Logger.log("This is the form find value" + formFindValue + "currentRow" + currentRow);
    if(currentRow && (formFindValue == "AM-SE")) { /* this is the new ticket */
      
      numberAsNumber2 = Number(ticketCounter);
      
      if (row === undefined || row === null)

        row = parseInt(numberAsNumber2+1);
      
      ticketCounter = (numberAsNumber2+1).toString();
            
      ticketNumber = "AM-SE" + ticketCounter;
      
      
    } else if (currentRow === undefined) {
      numberAsNumber2 = Number(ticketCounter);
      ticketCounter = numberAsNumber2+1;
      row = ticketCounter;
      ticketNumber = "AM-SE" + ticketCounter;
    }
    
    etrolControlsServiceEmail1 = "jay.bhaskar@gmail.com",
      etrolControlsServiceEmail2 = "jay.bhaskar@gmail.com",
        tktSubmitterEmail = e.response.getRespondentEmail(),
          //timestamp = shFormResponses.getRange(row, 1).getValues(),
          timestamp = e.response.getTimestamp(),
            ticketNumberLocation = shFormResponses.getRange(row, 3);
    
    
    ticketNumberLocation.setValue(ticketNumber);
    
    /* {  Let's set the other values */
    
    var sheetIndex = 1;
    
    if (tktSubmitterEmail == 'jay.bhaskar@gmail.com' || tktSubmitterEmail != 'alarkin@fastly.com') /* We are allowed to log tkts for someone else */
      shFormResponses.getRange(row, sheetIndex++).setValue(formValues[3].getResponse());
    else
      shFormResponses.getRange(row, sheetIndex++).setValue(tktSubmitterEmail);
    
    shFormResponses.getRange(row, sheetIndex++).setValue(ticketNumber);
    Logger.log("resp length =" + length);
    for (var resp = 1; resp < length; resp++, sheetIndex++) {
      //Logger.log(formValues[resp].getResponse() + " Sheet Index is at " + sheetIndex);
      if(formValues[resp-1].getResponse() == 'No' && formValues[resp].getItem().getTitle() == "Is this an opportunity?") { /* We will follow discipline of getting y/n answers for items right before they're entered */
        //respFlag = 1;
        //Logger.log("Flag is now set");
        //Logger.log("Will continue to next form entry +2 " + formValues[resp].getItem().getTitle());
        sheetIndex+=2;
      } else if(formValues[resp-1].getResponse() == 'No') { /* We will follow discipline of getting y/n answers for items right before they're entered */
        //respFlag = 1;
        //Logger.log("Flag is now set");
        //Logger.log("Will continue to next form entry +1 " + formValues[resp].getItem().getTitle());
        sheetIndex++;
      }
      if(formValues[resp].getItem().getTitle() == "Case Type")
      {
        caseType = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "Due date")
      {
        meetingDate = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "Assigned SE Email address")
      {
        etrolControlsServiceEmail3 = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "If meeting from what time?")
      {
        meetingTime = formValues[resp].getResponse();
      }
      if(formValues[resp].getItem().getTitle() == "If meeting for how long?")
      {
        meetingFinish = formValues[resp].getResponse();
      }
      if(formValues[resp-1].getResponse() == 'Yes' && formValues[resp].getItem().getTitle() == 'Duplicate or related case number')
      {
        duplicateCaseNum = formValues[resp].getResponse();
        Logger.log("Duplicate case number " + duplicateCaseNum);
        var localValues = shFormResponses.getDataRange().getValues();
        var localCurrentRow = null;
        for(var jLocal=0, jLocalLen=localValues.length; jLocal<jLocalLen; jLocal++) {
          for(var kLocal=0, kLocalLen=localValues[0].length; kLocal<kLocalLen; kLocal++) {
            var find = localValues[jLocal][kLocal];
            if(find == duplicateCaseNum) {
              localCurrentRow = jLocal+1; // If ticket number not found then throw an error saying duplicate number missing
              break;           
            }
          }
        }
        if (localCurrentRow == null)
          missingDuplicateFlag = 1;
      }
      
      Logger.log("Setting value " + formValues[resp].getResponse() + " at sheetIndex " + sheetIndex + " row " + row);
      
      if (formValues[resp].getItem().getTitle() == "Priority") {
        Logger.log("Priority is " +  formValues[resp].getResponse());
        priority = formValues[resp].getResponse();
      }
      if (formValues[resp].getItem().getTitle() == 'Duplicate or related case number') { //We don't want to set case numbers not present
        if (!missingDuplicateFlag)
          shFormResponses.getRange(row, sheetIndex).setValue(formValues[resp].getResponse());
      } //Else business as usual
      else
        shFormResponses.getRange(row, sheetIndex).setValue(formValues[resp].getResponse());
    }  
    
    /* } */
    
    //ticketCounter = (numberAsNumber2).toString();
    tktSheet = ss.getSheetByName('Maintenance');
    if (!editFlag)
      tktSheet.getRange(2, 1).setValue(ticketCounter);
    
    /* } 5 */
    
    //var reportedBy = formValues[2].getResponse() == 'No' ? formValues[3].getResponse() : formValues[4].getResponse()
    //var reportedBy = shFormResponses.getRange(row, 3).getValues();
    var reportedBy = e.response.getRespondentEmail();
    //var amEmail = shFormResponses.getRange(row, 5).getValues();
    var amEmail = shFormResponses.getRange(row, 5).getValue();
    //var priority = shFormResponses.getRange(row, 10).getValues();
    var customerName = shFormResponses.getRange(row, 4).getValues();
    var subject;
    var emailBody;
    if (formFindValue == "AM-SE") {
      
      subject =  "An issue for \""+ customerName + "\" has been reported on " + 
        timestamp + " " + "with ticket Number " + ticketNumber;
      
      emailBody = "To: Team Member AM-SE Cases " + 
        /*"\nRE: Issue reported by " + reportedBy + "." + */
        "\n\nAn issue has been reported for " + amEmail + 
          ". Please see the details below:" + "\nTicket Number: " + 
            ticketNumber + "\nCase Type: " + 
              caseType + "\nReported By: " + 
                reportedBy + "\nPriority Level: " + priority + "\nAccount: " + 
                  customerName + "\nYou may view/edit this case using this URL - " + e.response.getEditResponseUrl()+"&entry.425022159="+ticketNumber;
    } else {
      
      subject =  "An issue for \""+ customerName + "\" has been updated on " + 
        timestamp + " " + "with ticket Number " + ticketNumber;
      
      emailBody = "To: Team Member AM-SE Cases " + 
        /*"\nRE: Issue updated by " + reportedBy + "." + */
        "\n\nAn issue has been updated for " + amEmail + 
          ". Please see the details below:" + "\nTicket Number: " + 
            ticketNumber + "\n Case Type: " + 
              caseType + "\nReported By: " + 
                reportedBy + "\nPriority Level: " + priority + "\nAccount: " + 
                  customerName + "\nYou may view/edit this case using this URL - " + e.response.getEditResponseUrl()+"&entry.425022159="+ticketNumber; /* We are prefilling the response */
      //Logger.log("Query string is " + e.response.getEditResponseUrl());
    }
    
    if (meetingTime != NaN && meetingTime != null && meetingTime != "" && meetingFinish != NaN && meetingFinish != null && meetingFinish != "")
    {
      if (+String(meetingFinish).substring(0, 2) == "0" && +String(meetingFinish).substring(3, 5) == "0")
      {
        ;
      } else {
        var title = "Meeting with Account " + formValues[2].getResponse() + " for " + formValues[1].getResponse();
        //var calendar = CalendarApp.getCalendarById('fastly.com_92l2lnefqj26b9ucc8boc61onc@group.calendar.google.com');
        var calendar = CalendarApp.getCalendarById('jay.bhaskar@gmail.com');
        Logger.log("Meeting date is " + meetingDate + "Meeting time is " + meetingTime + " & duration is " + meetingFinish);
        
        /*var tempDate = new Date(String(meetingDate));
        Logger.log("Meeting date in Date obj format is " + tempDate.SetDate(1,1,2020));*/
        
        var year = +String(meetingDate).substring(0, 4);
        var month = +String(meetingDate).substring(5, 7);
        var day = +String(meetingDate).substring(8, 10);
        var mTimeHrs = +String(meetingTime).substring(0, 2);
        
        var mTimeMin = +String(meetingTime).substring(3, 5);
        var timeZone = Session.getScriptTimeZone();
        //var timeZone = Session.getTimeZone();
        
        var remainderDays = null;
        var remainderHrs = null;
        var remainderMins = null;
        var fTimeHrs = +String(meetingFinish).substring(0, 2);
        fTimeHrs = +String(Number(mTimeHrs)+Number(fTimeHrs));
        
        /*remainderDays = +String(ParseInt(fTimeHrs)/24);
        if(remainderDays) {
          day = +String(Number(day)+Number(remainderDays));
          fTimeHrs = +String(Number(fTimeHrs)%24);
        }*/
        
        var fTimeMin = +String(meetingFinish).substring(3, 5);
        fTimeMin = +String(Number(mTimeMin)+Number(fTimeMin));
        var fTimeNum = Number(fTimeMin);
        var remainderHrsNum = 0;
        for(;fTimeNum>=60; fTimeNum-=60) {
          remainderHrsNum+=1
        }
       
        if(remainderHrsNum) {
          fTimeHrs = +String(Number(fTimeHrs)+remainderHrsNum);
          fTimeMin = +String(fTimeNum);
        }
        
        
        Logger.log("Year Month Day Hrs Min TimeZone " + year + "," + month + "," + day + "," + mTimeHrs + "," + mTimeMin + "," + timeZone);
        Logger.log("Finish Time Hrs Min " + fTimeHrs + "," + fTimeMin);
        
        //var meetingTimeMath = new Date('February 17, 2018 13:00:00 -0500');
        
        //var meetingTimeMs = meetingTimeMath.getTime();
        
        //Logger.log("Meeting time math is " + meetingTimeMath + "Meeting time in ms is " + meetingTimeMs + "Meeting time given is" + String(meetingTime));
        //var event = calendar.createEvent(title, new Date(meetingTime), new Date(meetingTime)+(Number(meetingLength)* 1000 * 60 * 60) );
        //var event = calendar.createEvent(title, new Date('January 25 2020 13:00:00 -0500'), new Date('January 25 2020 14:00:00 -0500'));
        //var event = calendar.createEvent(title, new Date(meetingTime), new Date(meetingTimeMs + Number(meetingLength)* 1000 * 60 * 60) );
        //var startTime = String(months[Number(month)])+" "+String(day)+", "+String(year) +" "+mTimeHrs+":"+mTimeMin+" "+timeZone;
        var startTime = String(months[Number(month)])+" "+String(day)+", "+String(year) +" "+mTimeHrs+":"+mTimeMin;
        //var finishTime = String(months[Number(month)])+" "+String(day)+", "+String(year) +" "+fTimeHrs+":"+fTimeMin+" "+timeZone;
        var finishTime = String(months[Number(month)])+" "+String(day)+", "+String(year) +" "+fTimeHrs+":"+fTimeMin;
        Logger.log("Meeting start time is " + startTime);
        Logger.log("Meeting finish time is " + finishTime);
        var startDate = new Date(String(startTime));
        Logger.log("Meeting start date is " + startDate);
        var startDate = new Date(startTime);
        Logger.log("Meeting start date is " + startDate);
        
        
        if (formFindValue == "AM-SE") { /* New tkt */
          
          var advancedArgs = {description: title};
          
          if (etrolControlsServiceEmail3)
           advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+amEmail+','+etrolControlsServiceEmail3, sendInvites:true};
          else
            advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+amEmail, sendInvites:true};
          
          var event = calendar.createEvent(title, new Date(String(startTime)), new Date(String(finishTime)), advancedArgs);
          //event.addGuest(etrolControlsServiceEmail1);
/*          
          var eventResource = {
            "kind": "calendar#event",
              "description": title,
                "start": {
                  "date": meetingDate,
                    "dateTime": mTimeHrs+":"+mTimeMin,
                      "timeZone": timeZone
                },
                  "end": {
                    "date": meetingDate,
                      "dateTime": fTimeHrs+":"+fTimeMin,
                        "timeZone": timeZone
                  },
                    "anyoneCanAddSelf": true,
                      "guestsCanInviteOthers": true,
                        "guestsCanModify": true,
                          "guestsCanSeeOtherGuests": true,
          }
          
          Calendar.Events.patch(eventResource, calendar.getId(), event.getId(), {sendUpdates: "all"}); // Using the options Resources -> Advanced Google services 
*/
          emailBody += "\n\nMeeting event created in your calendar for " + meetingDate + meetingTime + " " + timeZone;
                
          event.addGuest(amEmail);
          if (etrolControlsServiceEmail3 != null)
            event.addGuest(etrolControlsServiceEmail3);
          event.setDescription(customerName+caseType);
          shFormResponses.getRange(row, sheetIndex).setValue(event.getId());
        } else {
          //var events = calendar.getEventsForDay(new Date(meetingDate));
          var eventId = shFormResponses.getRange(row, sheetIndex).getValue();
          var event = calendar.getEventById(eventId);
          var advancedArgs = {description: title};
          Logger.log("Deleting event with id" + eventId);
          
          if (event != undefined && event != null)
            event.deleteEvent();
          
          if (etrolControlsServiceEmail3)
           advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+amEmail+','+etrolControlsServiceEmail3, sendInvites:true};
          else
            advancedArgs = {description: title, location: 'here', guests:tktSubmitterEmail+','+amEmail, sendInvites:true};
          
          event = calendar.createEvent(title, new Date(String(startTime)), new Date(String(finishTime)), advancedArgs);
          //event.addGuest(etrolControlsServiceEmail1);
          emailBody += "\n\nMeeting event created in your calendar for " + meetingDate + meetingTime + " " + timeZone;
          
          event.addGuest(amEmail);
          if (etrolControlsServiceEmail3 != null)
            event.addGuest(etrolControlsServiceEmail3);
          event.setDescription(customerName+caseType);
          shFormResponses.getRange(row, sheetIndex).setValue(event.getId());
        }
        Logger.log("email " + etrolControlsServiceEmail1 + " customer " + customerName + "case type" + caseType);
        Logger.log("Event end time " + event.getEndTime() + "Event start time " + new Date(startTime));
      }
    }
    
    //MailApp.sendEmail(etrolControlsServiceEmail1, subject, emailBody);
    //MailApp.sendEmail(etrolControlsServiceEmail2, subject, emailBody);
    MailApp.sendEmail(tktSubmitterEmail, subject, emailBody);
    if (etrolControlsServiceEmail3 != null)
      MailApp.sendEmail(etrolControlsServiceEmail3, subject, emailBody);
    if (missingDuplicateFlag)
      MailApp.sendEmail(tktSubmitterEmail, "Couldn't find duplicate/related ticket", emailBody+"\n\nDuplicate/Related SE Ticket Submitted " + duplicateCaseNum + " Not found");
      Logger.log("2. AM email is " + amEmail);
    MailApp.sendEmail(amEmail, subject, emailBody);
    //Logger.log("2. Submitter email is " + tktSubmitterEmail);
    
  }
  else if (errorFlag == 1){
    
    var errEmailBody = "You can't submit this ticket no. " + formFindValue + " as you don't own it, sorry.";
    MailApp.sendEmail(tktSubmitterEmail,"Not your ticket to edit", errEmailBody);
  }
  else if(errorFlag == 2) {
    var errEmailBody = "This ticket no. wasn't found " + formFindValue + ". Please try again.";
    MailApp.sendEmail(tktSubmitterEmail,"Ticket not found", errEmailBody);
  }
  else if(errorFlag == 3) {
    var errEmailBody = "You can't pick the AM. Please try again.";
    MailApp.sendEmail(tktSubmitterEmail,"Not your ticket to edit", errEmailBody);
  }
  
  lock.releaseLock();
}

function onOpen(e)
{
}

/*
function onEdit(e) {

var formDestId = FormApp.getActiveForm().getDestinationId();
var ss = SpreadsheetApp.openById(formDestId);
SpreadsheetApp.setActiveSpreadsheet(ss);
shFormResponses = ss.getSheetByName('Maintenance');

var responses = FormApp.getActiveForm().getResponses();
var users = e.user;

var length = responses.length;
var lastResponse = responses[length-1];
var formValues = lastResponse.getItemResponses();
var formFindValue = formValues[0].getResponse();

Logger.log(formFindValue);

var currentRow, tktSubmitterEmail;

{
var values = shFormResponses.getDataRange().getValues();

for(var j=0, jLen=values.length; j<jLen; j++) {
for(var k=0, kLen=values[0].length; k<kLen; k++) {
var find = values[j][k];

if(find == formFindValue) {
Logger.log([find, "row "+(j+1)+"; "+"col "+(k+1)]);
currentRow = j+1;
break;           
}
}
}
}

ss.setActiveSheet(shFormResponses);

var ticketNumber;
var ticketCounter;
var row;

if(currentRow && (formFindValue != "AM-SE"))
{
tktSubmitterEmail = shFormResponses.getRange(currentRow, 2).getValues();

if (tktSubmitterEmail != user.getEmail()) {

Browser.msgBox("Not your ticket to edit");

}
}

}
*/