 /* The script is deployed as a web app and renders the form */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('sender.html');
}

function send(form) {     
  var docHtml = form.htmlFile.getDataAsString();
  var emails = getEmailIds(form.sheet, form.from, form.column, (form.to - form.from) + 1); //search results
  var unsubscribeList = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3, -1); //unsubscribe 
  var report = '';
  
  for(var i= 0; i<emails.length;i++){
     if(unsubscribeList.indexOf(emails[i])==-1){
       Logger.log(emails[i]);
       try{
         MailApp.sendEmail({
           to: emails[i], 
           replyTo: "contact@ruralindiaonline.org",
           subject: form.subject,
           htmlBody: docHtml,
           name: "P. Sainath, PARI"
         });
         report+=addToReport(report, emails[i]+" SENT");
       }catch(e){
         report+=addToReport(report, emails[i]+" "+e.message);
       }
     }else{
       report+=addToReport(report, emails[i]+" HAS UNSUBSCRIBED");
     }
   }
  return report;
}

function addToReport(report, message){
   Logger.log(message);
   return report+=message+"<br>";
}

function testGetEmail(){
  var emails = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 1, 2);
  Logger.log(emails.length);
  Logger.log(emails[0]);
  Logger.log(emails[1]);
}
  
function getEmailIds(sheet, startRow, startColumn, numRows){
     //get email addresses from sheet
     var ss = SpreadsheetApp.openByUrl(sheet);
     SpreadsheetApp.setActiveSpreadsheet(ss);    
     ss = SpreadsheetApp.getActiveSpreadsheet();
     SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
     Logger.log("opening successful")
     var contents;
     if(numRows!=-1){
       contents = ss.getActiveSheet().getSheetValues(startRow, startColumn, numRows, 1);
     }else{
       contents =  ss.getActiveSheet().getSheetValues(startRow, startColumn, ss.getActiveSheet().getLastRow() - startRow + 1, 1)
     }
  return removeDuplicates(contents);
}

function removeDuplicates(contents){
    //remove duplicates
     var emails = new Array();
     var j=0;
     for(var i=0;i<contents.length;i++){
       if(contents[i][0] !='' && emails.indexOf(contents[i][0])==-1){
         emails[j++] = contents[i][0]; 
       }
     }
     Logger.log((contents.length-emails.length)+" nulls-duplicates removed");
     return emails;  
}

function getGoogleDocumentAsHTML(id){
  var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+id+"&exportFormat=html";
  var param = {
    method      : "get",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions:true,
  };
  return UrlFetchApp.fetch(url,param).getContentText();
}