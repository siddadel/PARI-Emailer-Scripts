 
/* The script is deployed as a web app and renders the form */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('form.html');
}

function send(form) { 
  try{
   var report = ''; 
   var blob = form.myFile;
   var attachmentId = uploadFile(blob);
   var doc = DocumentApp.openByUrl(form.content);
   
   var docHtml = getGoogleDocumentAsHTML(doc.getId());
   var emails = getEntries(form.list, 2 , 2);   
   var size = emails.length;   

    
   var list = new Array(); 
    
   for(var i=0;i<size;i++){
     try{
       var htmlbody = docHtml.replace('pari_reader', emails[i][0]);
         if(attachmentId ==null){
           MailApp.sendEmail({
             to: emails[i][1], 
             subject: form.subject,
             htmlBody: htmlbody,
             name: form.fromName 
           });
         } else{
           var file = DriveApp.getFileById(attachmentId);
           MailApp.sendEmail({
             to: emails[i][1],
             subject: form.subject,
             htmlBody: htmlbody,
             name: form.fromName,
             attachments: [file.getAs(file.getMimeType())]
           });
           
         }
       
     
       report = report + emails[i][0] +" "+emails[i][1] + "<br>";
     } 
     catch(error){
       report = report + "ERROR "+ emails[i][0] + " "+emails[i][1] +"<br>"+error.message+'<br>';
     }
   }
   return "<h1><a href=\"https://script.google.com/a/ruralindiaonline.org/macros/s/AKfycbwPbSyRjqN12zen_cRzZgbHVjwp2IASrhMPERLoO07SB87TbP8/exec\">Click here to go Back</a></h1><br>Emails have been sent to the following <br>"+ report;
  } catch(error){
    return "<h1><a href=\"https://script.google.com/a/ruralindiaonline.org/macros/s/AKfycbwPbSyRjqN12zen_cRzZgbHVjwp2IASrhMPERLoO07SB87TbP8/exec\">Click here to go Back</a></h1><br>There has been an error. Report sent to only the following<br>"+report+error.message;
  }
}

function contains(email, list){
   var size = list.length;    
  for(var i=0;i<size;i++){
    if(list[i]==email){
      return true;
    }
  }
  list[list.length] =email;
  return false;
}

function uploadFile(blob){
  if(blob.getDataAsString()==''){
    return null;
  }
  var folder = DriveApp.getFoldersByName("Email Attachments");  
  var file = folder.next().createFile(blob); 
  return file.getId();
}

  
function getEntries(sheet, startRow, startColumn){
     //get email addresses from sheet
     var ss = SpreadsheetApp.openByUrl(sheet);
     SpreadsheetApp.setActiveSpreadsheet(ss);    
     ss = SpreadsheetApp.getActiveSpreadsheet();
     SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
     var contents = ss.getActiveSheet().getSheetValues(startRow, startColumn, ss.getActiveSheet().getLastRow(), 2);
     
  return contents;
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