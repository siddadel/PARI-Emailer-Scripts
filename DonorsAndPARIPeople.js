 /* The script is deployed as a web app and renders the form */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('OnlyNL.html');
}

function send(form) {  
  var report = '';
  
  if(form.test=="on"){
    report=addToReport(report, "This is a test. You had checked the test checkbox");
  }else{
    report=addToReport(report, "This is not a test. You did not check the test checkbox");
  }
 
  if(form.subject==""){
    report=addToReport(report, "You have not entered a subject. No emails were sent");
    return report;
  }else{
    report=addToReport(report, "Subject: "+form.subject);
  }
  
  var docHtml = form.htmlFile.getDataAsString();
  
  if(docHtml==''){
    report=addToReport(report, "Please upload a newsletter html file. You did not upload an html file. No emails were sent");
  }
  else {
    var donors = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3); // test1
    var paripeople = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3); //test2 
    
    if(form.test!="on"){
      donors = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3); // donors Donor list for Newsletter ZL/SID
      paripeople = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3); //pari people  
    }
    
    var emails = concat(donors, paripeople);
    var unsubscribeList = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3); //unsubscribe 
    
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
          report=addToReport(report, emails[i]+" SENT");
        }catch(e){
          report=addToReport(report, emails[i]+" "+e.message);
        }
      }else{
        report5tr=addToReport(report, emails[i]+" HAS UNSUBSCRIBED");
      }
    }
  }
  
  return report;
}

function concat(donors, paripeople){
  var emails = new Array();
  for(var i=0; i<donors.length; i++){
    emails[i] = donors[i];
  }
  
  var k=0;
  for(var j=0;j<paripeople.length;j++){
    if(emails.indexOf(paripeople[j])!=-1){
      emails[k++] = paripeople[j]
    }
  }
  
  return emails;
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
  
function getEmailIds(sheet, startRow, startColumn){
     //get email addresses from sheet
     var ss = SpreadsheetApp.openByUrl(sheet);
     SpreadsheetApp.setActiveSpreadsheet(ss);    
     ss = SpreadsheetApp.getActiveSpreadsheet();
     SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
     var contents = ss.getActiveSheet().getSheetValues(startRow, startColumn, ss.getActiveSheet().getLastRow(), 1);
     
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