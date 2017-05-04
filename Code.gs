function onOpen() {
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('PDF Report')
      .addItem('Save PDF and send email', 'showPrompt')
      .addToUi();
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Let\'s get to know each other!',
      'Please enter your name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK && text != "")  {
    // User clicked "OK".
    SpreadsheetToPDF(text)
    ui.alert('Report salvato in  https://drive.google.com/drive/u/0/folders/0B8_ub-Gf21e-c0J6cmVlWWFNVUU');
    
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('Operazione annullata');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}

function SpreadsheetToPDF(email) {
  var spreadsheetId = '1XuJdpezUNOuzjKpWMz1cDBKRuKcC1_-uVcXi4YDt9lc';
  var file = Drive.Files.get(spreadsheetId);
  var url = file.exportLinks['application/pdf'];
  var url_ext = '&size=letter'                                           // paper size
              + '&portrait=true'                                         // orientation, false for landscape
              + '&fitw=true'                                             // fit to width, false for actual size
              + '&sheetnames=false&printtitle=false&pagenumbers=false'   // hide optional headers and footers
              + '&gridlines=false'                                       // hide gridlines
              + '&fzr=false';                                            // do not repeat row headers (frozen rows) on each page
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url + url_ext, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
 
  

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();

var folderId = "0B8_ub-Gf21e-c0J6cmVlWWFNVUU";  //report clienti - folder
var outputFilename_long = sheets[1].getRange("C5").getValue();
  var outputFilename = "Report: " +outputFilename_long.replace(/ /g,"_");  // replace spaces with _
  
sheets[0].hideSheet();  // hide sheet for pdf

 var pdf = response.getBlob();
  
DriveApp.getFolderById(folderId)
    .createFile(UrlFetchApp.fetch(
      url+url_ext,
      {
        method: "GET",
        headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true
      })
    .getBlob())
    .setName(outputFilename);  



  
  var emailToDriver = email + ',cristiano.dalfarra@bakeca.it,alice.cestari@bakeca.it,adele.toma@bakeca.it';
  var subjectDriver = 'Report: '+ outputFilename_long ;
  var messageDriver = "In allegato il report richiesto. Grazie.";
  var attach = pdf;
  
  
    
  var files = DriveApp.getFilesByName(outputFilename);
  if(files.hasNext()){
  var file = files.next();
  MailApp.sendEmail(emailToDriver, subjectDriver, messageDriver, {attachments:file});
}
    
 // MailApp.sendEmail(emailToDriver, subjectDriver, messageDriver, {attachments:[attach]});   // Se  
  
  
  
sheets[0].showSheet();
sheets[1].showSheet();  
  
  return pdf;
}

