function onOpen() {
    SpreadsheetApp.getUi().createMenu('RemoteAccess')
        .addItem('Authorize', 'createSpreadsheetEditTrigger')
        .addItem('Logout Anywhere', 'logoutAnywhere')
        .addItem('Open #remote-access channel', 'openRemoteAccessChannel')
        .addItem('About', 'About')
        .addToUi()
}


function createSpreadsheetEditTrigger() {
    checkTriggers();
    ScriptApp.newTrigger('runOnEdit')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
    var user = Session.getActiveUser().getEmail();
    SpreadsheetApp.getUi().alert('Hi, ' + user + '. The script has been authorized.');
}


function runOnEdit(e) {
//    if (e.source.getActiveSheet().getName() !== 'ras' 
//        //|| e.range.columnStart !== 6 || !e.value
//       ) return;
//    e.range.offset(0, 1, 1, 2).setValues([
//        [Session.getActiveUser().getEmail().split("@")[0], new Date()]
//    ])
}

function checkTriggers() {
    var allTriggers = ScriptApp.getProjectTriggers();
    if (allTriggers.length > 0) {
        for (var i = 0; i < allTriggers.length; i++) {
            ScriptApp.deleteTrigger(allTriggers[i]);
        }
    }
}

/*CyberHacktivist.com*/

function onEdit(event)
{ 


  var timezone = "GMT+3";
  var timestamp_format = "dd.MM.yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "IsBusy";
  var updateColName2 = "PreviousUser";
  var updateColNameCurrentUser = "CurrentUser";
  var timeStampColName = "Timestamp";
  var sheet = event.source.getSheetByName('RA'); //Name of the sheet where you want to run this script.

  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  var updateCol2 = headers[0].indexOf(updateColName2); updateCol2 = updateCol2+1;
  var updateColNameCurrentUser = headers[0].indexOf(updateColNameCurrentUser); updateColNameCurrentUser = updateColNameCurrentUser+1;

  if (dateCol > -1 && index > 1 && editColumn == updateCol)// && indexOf(updateColName) >= 1)// && editColumn == updateCol) 
  { 
    // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    //   Browser.msgBox(event.value);
    //    Browser.msgBox(event.oldValue);
    
    // Place current user to the "previousUser" column (if exist)
    if (sheet.getRange(index, updateColNameCurrentUser).getValue() != "")
    {
      var cell = sheet.getRange(index, updateCol2);
      cell.setValue(sheet.getRange(index, updateColNameCurrentUser).getValue());    
    }
    
    if (event.value == "TRUE")
    {
      var cell = sheet.getRange(index, dateCol + 1);
      var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    
      var email = Session.getActiveUser().getEmail();
      cell.setValue(date);
      //cell.setValue(date).setComment(email);
    
      var email = Session.getActiveUser().getEmail();
      sheet.getRange(event.source.getActiveRange().getRowIndex(), event.source.getActiveRange().getColumnIndex()+1).setValue(email);
    }
    else
    {
      var cell = sheet.getRange(index, dateCol + 1);
      cell.setValue("");
      sheet.getRange(event.source.getActiveRange().getRowIndex(), event.source.getActiveRange().getColumnIndex()+1).setValue("");
    }
  }  
}

function logoutAnywhere() {

}


