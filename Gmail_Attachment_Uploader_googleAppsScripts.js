//help from: https://stackoverflow.com/questions/46492616/using-google-script-to-get-attachment-from-gmail

/**
*Run runEveryDay to trigger uploadSpreadsheetToDrive every day between 1am and 2am, replacing your old data with current data.
*uploadDataToDrive gets a specified number of spreadsheets from your emails, and uploads them to your specified root Drive folder 
*based on the search returned. Once the data is in Google Drive, you can access visualize your data with tools like Splunk and Google 
*Data Studio. 
*Need to set the variables before running. 
**/

//Upload Gmail Attachment to Google Drive

//run from this once i get it working
function runEveryDay() {
  // Trigger every 24 hours. run at 1:00am each day
  ScriptApp.newTrigger('uploadDataToDrive')
      .timeBased()
      .atHour(1)
      .everyDays(1)
      .create();
  
}

// variables to set
var mySpreadsheetId = ""; // string ID of file, found in URL
var emailSearchTerm = '""'; //search term for your email results
var attachmentName = ".csv"; //name of your attachment in email
var myEmail = "@gmail.com"; // email to send to in event of error
var header = [""]; // string array header
var numEmails = 1; //  the number of emails you want to limit your search to (sorted by most recent)
var numAttachmentsWanted = 1; // maximum total number of attachments desired from email search to upload to Drive
var doAppendDate = true; //boolean to append date to last column in CSV. If true then must set "Date" as last column in header

function uploadDataToDrive(){
  Logger.clear();
  Logger.log("Uploading .CSVs to Google Drive");
  //need to create file within folder
  var dApp = DriveApp;
  var driveSpreadsheet = getExistingSheet();
  
  if(driveSpreadsheet == null){
    
    //if problem finding desired sheet, create a new one and run anyways but send me an email
    var errorEmail = GmailApp.createDraft(myEmail, "Spreadsheet error: creating new spreadsheet instead of adding to old. Fix it!");
    errorEmail.send();
    
    Logger.log("Filename does not exist yet. Creating file");
    driveSpreadsheet = SpreadsheetApp.create("new_Spreadsheet");
  }
  writeEmailLinesToCsv(driveSpreadsheet); 
}

//Retrieves already existing data.csv in Google Spreadsheets and clears it for new data
function getExistingSheet(spreadsheetId){
  Logger.log("Getting existing sheet from drive");
  spreadsheetId = mySpreadsheetId; 
  var currentSheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  if (currentSheet != null) {
    currentSheet.clear();
  }
  return currentSheet
}
  

function writeEmailLinesToCsv(csvFile){
  Logger.log("Writing csv lines to Drive file");
  csvFile.appendRow(header);
  Logger.log("Appended header");
  var csvFileArray = getSpreadsheetFromEmail();
  for(var i=0;i<csvFileArray.length;i++){
    //get array of csv data from each email attachment
    var dataArray = csvFileArray[i];
    //start at 1 because don't want headers ever.
    for(var j = 1; j<dataArray.length;j++){
      dataRow = dataArray[j];
      csvFile.appendRow(dataRow);
    }
  }
  Logger.log("Returning completed csv file");
  return csvFile;
}

function getSpreadsheetFromEmail() {
  var searchTerm= GmailApp.search(emailSearchTerm, 0, numEmails);
  var messages = GmailApp.getMessagesForThreads(searchTerm);
  Logger.log("Messages length: " + messages.length);
  
  //take first attachment from first message
  var attachments = [];
  if(messages.length < numEmails) numEmails = messages.length; 
  for(var i = 0; i<numEmails;i++){
    //if we already have numAttachmentsWanted, break from loop to return attachments
    if(attachments.length > numAttachmentsWanted) break;
    Logger.log("Attachments length: " + attachments.length);
    //the [0] gets the most recent email in the thread
    var message = messages[i][0];
    //[0] gets first attachment in email. Can change if different attachment is desired
    var attachment = message.getAttachments()[0];
    
    //if there are no attachments, then continue to next message in loop
    if(undefined === attachment) continue;
    
    //if the email has an attachment and it has the correct name, then we'll append. Otherwise loop to next email
    if(attachment.getName() === attachmentName){
      var date = message.getDate();
      var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");
      if(doAppendDate == true){
        var csvData = appendDateToCsv(csvData, date);
      }
      
      attachments.push(csvData);
    }
  }
  Logger.log("Returning attachments");
  return attachments;
}


function appendDateToCsv(csvData, date){
  for(var i = 1; i<csvData.length;i++){
    var pacingDataLine = csvData[i];
    //Logger.log("line before: " + pacingDataLine);
    //0 for insert for splice method. do not want to replace any data
    var insertOrReplace = 0
    //splice into the end of the array BUT before the ",," line ending. 
    pacingDataLine.splice(pacingDataLine.length,0,date);
    //Logger.log("line after: " + pacingDataLine);
  }
  return csvData
}
