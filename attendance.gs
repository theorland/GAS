var moment = Moment.load();
var SELECTED_SHEET = SpreadsheetApp.getActive().getSheetByName("MAIN");
var ALL_DATA ;
var RESULT, 
    RESULT_SHEET_NAME ;
var FORMAT_RAW = "DD-MMM-YY hh:mm A",
    FORMAT_DATE = "DD-MMM-YYYY",
    FORMAT_DATE_COMPLETE = "DD MMMM YYYY",
    FORMAT_DATE_VALUE = "YYYY-MM-DD",
    FORMAT_DATE_DEBUG = "YYYY-MM-DD hh:mm",
    FORMAT_DAY = "dddd",
    FORMAT_TIME  =  "HH:mm";

var XY = {
  x:1,y:1
};
var S_XY = JSON.stringify(XY);

var NAME_XY = JSON.parse(S_XY), 
    CHECKED_XY = JSON.parse(S_XY),
    APPROVED_XY = JSON.parse(S_XY),
    DATE_XY = JSON.parse(S_XY);
    
var NEW_RECORD = {
  acc_no : "",
  name :"",
  date : "",
  day : "",
  min : "",
  max : "",
};
var S_NEW_RECORD = JSON.stringify(NEW_RECORD);

var IDENTITY = {
  dept: "",
  name: "",
  checked:"",
  approved:"",
};
var S_IDENTITY = JSON.stringify(IDENTITY), ALL_IDENTITY = {};

function getAllData(){
  // FOR DEBUGGING
  //SELECTED_SHEET = SpreadsheetApp.getActive().getSheetByName("2017_08_16");
  var last_row = SELECTED_SHEET.getLastRow();
  var max_col = SELECTED_SHEET.getLastColumn();
  var all_values = SELECTED_SHEET.getRange(1, 1,last_row,max_col).getValues();
  ALL_DATA = all_values;
};

function gatherData(){
  var COL_ACC = 1 , COL_NAME = 3 , COL_TIME = 4;
  COL_ACC = COL_ACC -1;
  COL_NAME =  COL_NAME - 1;
  COL_TIME = COL_TIME - 1 ;
  var i_row = 0, m_row =ALL_DATA.length;
  
  var temp_acc = "";
  var temp_name = "";
  var temp_raw_data = "";
  var temp_date = "";
  var temp_time ="";
  var new_record = "";
  RESULT = {};
  
  for (i_row =1;i_row < m_row; i_row++) {
    temp_acc = ALL_DATA[i_row][COL_ACC];
    temp_name = ALL_DATA[i_row][COL_NAME];
    temp_raw_data = moment(ALL_DATA[i_row][COL_TIME],FORMAT_RAW);
    
    if (RESULT[temp_name] ==undefined){
      RESULT[temp_name] = {};
    }
    temp_date = temp_raw_data.format(FORMAT_DATE_VALUE);
    
    if (RESULT[temp_name][temp_date]==undefined){
      new_record = JSON.parse(S_NEW_RECORD);
      new_record.acc_no = temp_acc;
      new_record.name = temp_name;
      new_record.date = temp_raw_data.format(FORMAT_DATE);
      new_record.day = temp_raw_data.format(FORMAT_DAY);
      new_record.min = temp_raw_data.format(FORMAT_TIME);
      new_record.max =  new_record.min;
      RESULT[temp_name][temp_date] = new_record;
      
    }else {
      RESULT[temp_name][temp_date].max = temp_raw_data.format(FORMAT_TIME);
    }
    
  } 
}

function importHeader(sheet,curr_row){
  var ref_sheet = SpreadsheetApp.getActive().getSheetByName("HEADER");
  
  var max_row = ref_sheet.getLastRow();
  var max_col = ref_sheet.getLastColumn();
  var range_ref = ref_sheet.getRange(1, 1,max_row,max_col);
  var range_dest = sheet.getRange(curr_row, 1,1,max_col-1);
  sheet.insertRows(curr_row);
  range_ref.copyTo(range_dest);
  return range_ref.getNumRows();
}

function importFooter(sheet,curr_row,identity){
  var ref_sheet = SpreadsheetApp.getActive().getSheetByName("FOOTER");
  
  var max_row = ref_sheet.getLastRow();
  var max_col = ref_sheet.getLastColumn();
  var range_ref = ref_sheet.getRange(1, 1,max_row,max_col);
  var range_dest = sheet.getRange(curr_row, 1,1,max_col-1);
  var backup_values = range_ref.getValues();
  var changed_values = range_ref.getValues() ;
  
  changed_values[DATE_XY.y-1][DATE_XY.x-1] = moment().format(FORMAT_DATE_COMPLETE);
  
  if (identity!=null){
    changed_values[CHECKED_XY.y-1][CHECKED_XY.x-1] = "(" + identity.checked + ")";  
    changed_values[APPROVED_XY.y-1][APPROVED_XY.x-1] = "(" + identity.approved + ")";  
    changed_values[NAME_XY.y-1][NAME_XY.x-1] = "(" + identity.name + ")";  
  }
  range_ref.setValues(changed_values);
  SpreadsheetApp.flush();
  sheet.insertRows(curr_row);
  range_ref.copyTo(range_dest);
  range_ref.setValues(backup_values);
  SpreadsheetApp.flush();
  return range_ref.getNumRows();
}

function putDataToResult(){
  SELECTED_SHEET = SpreadsheetApp.getActive().getSheetByName(RESULT_SHEET_NAME);
  var curr_row = 1;
  var curr_staff ,curr_data ; 
  var i_name, i_date, i_no;
  var COL_NO = 2, COL_NO_AKUN = 3, COL_NAMA = 4,
      COL_DEPT = 5, COL_DATE = 6, COL_MIN = 7, COL_MAX = 8,
      NUM_COL = 8;
  var all_record, new_record, curr_range, name_dept, curr_identity;
  for (i_name in RESULT) { 
    curr_staff=RESULT[i_name];
    
    all_record =[];
    i_no =0;
    curr_identity = null;
    name_dept = "No Dept";
    if (ALL_IDENTITY[i_name]!=null){
      name_dept = ALL_IDENTITY[i_name].dept;
      curr_identity = ALL_IDENTITY[i_name];
    }
    
    for (i_date in curr_staff){
       curr_data = curr_staff[i_date];
       all_record.push([
         ++i_no,
         curr_data.acc_no,
         curr_data.name,
         name_dept,
         curr_data.date,
         curr_data.day,
         curr_data.min,
         curr_data.max]);
    }
    if (all_record.length<=0) continue;
    
    curr_row+=importHeader(SELECTED_SHEET,curr_row);
    SELECTED_SHEET.insertRows(curr_row,all_record.length);
    curr_range=SELECTED_SHEET.getRange(curr_row, 2,all_record.length,NUM_COL);
    curr_range.setValues(all_record);
    curr_row +=all_record.length;
    curr_row+=importFooter(SELECTED_SHEET,curr_row,curr_identity);
    curr_range.setBorder(true, true, true, true, true, true,"black", SpreadsheetApp.BorderStyle.SOLID);
    SpreadsheetApp.flush();
  }
}


function reportGenerate(){
   var sheet_name = Browser.inputBox("Input the sheet name");
   //var sheet_name = "2017_08_16";
   loadAllConfig();
   var SS = SpreadsheetApp.getActive();
   SELECTED_SHEET = SS.getSheetByName(sheet_name);
   if (SELECTED_SHEET == null) {
     Browser.msgBox("The sheet is not found");
     return 0;
   }
   getAllData();
   gatherData();
   RESULT_SHEET_NAME = "RESULT - " + sheet_name;
   sheet_name = RESULT_SHEET_NAME;
   SELECTED_SHEET = SS.getSheetByName(sheet_name);
   if (SELECTED_SHEET != null) {
     SS.deleteSheet(SELECTED_SHEET);
   }
   SELECTED_SHEET = SS.insertSheet();
   SELECTED_SHEET.setName(sheet_name);
   SpreadsheetApp.flush();
   SELECTED_SHEET = SS.getSheetByName(sheet_name);
   //MailApp.sendEmail("theo@is-indonesia.com", "debugging",JSON.stringify(RESULT));
   putDataToResult();
}

function doTest(){
     //var sheet_name = Browser.inputBox("Input the sheet name");
   var sheet_name = "2017_08_16";
   loadAllConfig();
   var SS = SpreadsheetApp.getActive();
   
   RESULT_SHEET_NAME = "RESULT - " + sheet_name;
   sheet_name = RESULT_SHEET_NAME;
   SELECTED_SHEET = SS.getSheetByName(sheet_name);
   if (SELECTED_SHEET != null) {
     SS.deleteSheet(SELECTED_SHEET);
   }
   
   SELECTED_SHEET = SS.insertSheet();
   SELECTED_SHEET.setName(sheet_name);
   SpreadsheetApp.flush();
  
   SELECTED_SHEET = SS.getSheetByName(sheet_name);
   importFooter(SELECTED_SHEET,1,IDENTITY["Marc"]);
}

function loadAllConfig(){
  var SS = SpreadsheetApp.getActive();
  var act_sheet = SS.getSheetByName("CONFIG");
  var COL_CONFIG = 2;
  var ROW_RAW = 2,
      ROW_DATE = 3,
      ROW_TIME = 4,
      ROW_DAY = 5,
      ROW_COMPLETE = 6;
  var config_values = act_sheet.getRange(ROW_RAW, COL_CONFIG,5,1).setNumberFormat("@").getValues();
  FORMAT_RAW =  config_values[ROW_RAW-ROW_RAW][0].toString();
  FORMAT_DATE = config_values[ROW_DATE-ROW_RAW][0].toString();
  FORMAT_TIME = config_values[ROW_TIME-ROW_RAW][0].toString();
  FORMAT_DAY = config_values[ROW_DAY-ROW_RAW][0].toString();
  FORMAT_DATE_COMPLETE = config_values[ROW_COMPLETE-ROW_RAW][0].toString();
  
  var ROW_XY = 7;
  var LIST_CONFIG = [DATE_XY, NAME_XY,CHECKED_XY,APPROVED_XY];
  
  config_values  = act_sheet.getRange(ROW_XY,COL_CONFIG,4,2).setNumberFormat("@").getValues();
  var i,i_max;
  for (i=0, i_max = config_values.length; i<i_max ; i++){
     LIST_CONFIG[i].y = parseInt(config_values[i][0]);
     LIST_CONFIG[i].x = parseInt(config_values[i][1]);
  }
  
  ALL_IDENTITY = {};
  act_sheet = SS.getSheetByName("NAME_LIST");
  var act_max = act_sheet.getLastRow() - 2;
  config_values = act_sheet.getRange(2,1,act_max,5).setNumberFormat("@").getValues();
  var new_record, curr_record;
  for (i=0,i_max = config_values.length; i<i_max ; i++){
    curr_record = config_values[i];
    new_record = JSON.parse(S_IDENTITY);
    new_record.dept=curr_record[1].toString();
    new_record.name=curr_record[2].toString();
    new_record.checked=curr_record[3].toString();
    new_record.approved=curr_record[4].toString();
    ALL_IDENTITY[curr_record[0].toString()] = new_record;
  }
  
}

function configTest(){
  var SS = SpreadsheetApp.getActive();
  var act_sheet = SS.getSheetByName("CONFIG");
  var COL_SAMPLE = 4;
  var ROW_RAW = 2,
      ROW_DATE = 3,
      ROW_TIME = 4,
      ROW_DAY = 5,
      ROW_COMPLETE = 6;
  
  loadAllConfig();
  
  act_sheet.getRange(ROW_RAW, COL_SAMPLE).setNumberFormat("@").setValue(moment().format(FORMAT_RAW));
  act_sheet.getRange(ROW_DATE,COL_SAMPLE).setNumberFormat("@").setValue(moment().format(FORMAT_DATE));
  act_sheet.getRange(ROW_TIME,COL_SAMPLE).setNumberFormat("@").setValue(moment().format(FORMAT_TIME));
  act_sheet.getRange(ROW_DAY,COL_SAMPLE).setNumberFormat("@").setValue(moment().format(FORMAT_DAY));
  act_sheet.getRange(ROW_COMPLETE,COL_SAMPLE).setNumberFormat("@").setValue(moment().format(FORMAT_DATE_COMPLETE));
  
  var ROW_XY = 7;
  var LIST_CONFIG = [DATE_XY, NAME_XY,CHECKED_XY,APPROVED_XY];
  var i, i_max;
  var ref = SS.getSheetByName("FOOTER");
  for (i = 0, i_max = LIST_CONFIG.length ; i<i_max ; i++){
    act_sheet.getRange(ROW_XY+i,COL_SAMPLE).setNumberFormat("@").setValue(ref.getRange(LIST_CONFIG[i].y,LIST_CONFIG[i].x).getValue().toString());
  }
  
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  var menu = ui.createMenu('Report');
  menu.addItem('Report Creation', 'reportGenerate')
      .addSeparator()
  // FOR TESTING
      .addItem("Test Configuration","configTest")
  //  .addItem('test','doTest')
  // END FOR TESTING
      .addToUi();
  
}
