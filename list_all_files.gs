
var IT_FOLDER_PATH = "IT Tutorial"
var APP = SpreadsheetApp.getActive();
var ACT_SHEET = SpreadsheetApp.getActiveSheet();
var START_ROW = 3;
var NOT_DOCS_SHEET = ["CALENDAR"]

function deleteAllList(sheet){
  while (sheet.getLastRow()> START_ROW){
    sheet.deleteRow(START_ROW+1);
  }
  sheet.insertRows(START_ROW+1);
}

function getFolderPath(sheet){
  return sheet.getRange(1, 2).getValue().toString();
}
function doListAllCalendar(){
  var all_cal = CalendarApp.getAllCalendars();
  var sh_out = APP.getSheetByName("CALENDAR");
  
  var COL_NAME = 1;
  var COL_ID = 2;
  var COL_URL = 3;
  
  var curr_cal;
  var row_id = START_ROW;
  var now = new Date();
  
  //var date_start =  new Date(now.getYear(), now.getMonth(), now.getDay()+1, hours, minutes, seconds, milliseconds);
  for (var i=0;i < all_cal.length ; i++){
    row_id++;
    curr_cal = all_cal[i]; 
    sh_out.insertRowsBefore(row_id,1);
    sh_out.getRange(row_id, COL_NAME).setValue(curr_cal.getName());
    sh_out.getRange(row_id, COL_ID).setValue(curr_cal.getId());   
  }
}
function doListAllLinks(){
  var all_sheet = APP.getSheets();
  
  var sheet ;
  for (var i =0 ; i <all_sheet.length ; i++ ) {
    sheet = all_sheet[i];
    if (sheet.getName() in NOT_DOCS_SHEET) {
      continue;
    }
    deleteAllList(sheet);
    const COL_ID= 3, COL_NAME = 2, COL_MIME = 4, COL_URL = 1;
    
    var iter = DriveApp.getFoldersByName(getFolderPath(sheet));
    if (!iter.hasNext()){
      return 0;
    }
    var all_it = iter.next();
    iter = all_it.getFiles();
    // iter = DriveApp.getFiles();
    var each_file;
    var row_id = START_ROW;
    
    var curr_document;
    var curr_range;
    while (iter.hasNext()){
      each_file = iter.next();
      if (each_file.getMimeType() != "application/vnd.google-apps.document") {
        continue;
      }
      row_id++;
      sheet.insertRowsBefore(row_id,1);
      sheet.getRange(row_id, COL_ID).setValue(each_file.getId());
      curr_range = sheet.getRange(row_id, COL_NAME);
      curr_range.setValue(each_file.getName());
      curr_range.setFontWeight("bold");
      
      sheet.getRange(row_id, COL_MIME).setValue(each_file.getMimeType());
      curr_document = DocumentApp.openById(each_file.getId());
      
      sheet.getRange(row_id, COL_URL).setValue(curr_document.getUrl());
    }
  }
}
function onOpen(){
  ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('TutorialTools');
  menu.addItem('List All Tutorial', 'doListAllLinks')
      .addItem('List All Calendar', 'doListAllCalendar')
      .addSeparator();
  
  menu.addToUi();
  
}
