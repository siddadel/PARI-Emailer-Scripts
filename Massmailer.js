function sendEmails() {    
  //email ids should be in 3rd column
     var emails = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3); //search results
     var docHtml =  HtmlService.createHtmlOutputFromFile('letter.html').getContent();
   
  var unsubscribeList = getEmailIds("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit#gid=0", 2, 3); //unsubscribe
   
   Logger.log(emails.length);
   var c = 1;
   var onetime= 500;
   var start = c*onetime;
   var end = start+onetime;
    Logger.log(start+" "+end)
   for(var i= start;i<end && i<emails.length;i++){
     if(unsubscribeList.indexOf(emails[i])==-1){
       Logger.log(emails[i]);
       try{
         MailApp.sendEmail({
           to: emails[i], 
           replyTo: "contact@ruralindiaonline.org",
           subject: "Sturdy cattle that sustain fragile communities; Lilabai's ceaseless loop of labour; Into the precarious world of India's weavers",
           htmlBody: docHtml,
           name: "P. Sainath, PARI"
         });
       }catch(e){
         Logger.log(emails[i]+" "+e.message)
       }
     }else{
       Logger.log(emails[i]+" HAS UNSUBSCRIBED")
     }
   }
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