var moment = Moment.load();

var SHEET_NAME = "2017";
var email_to_send = ["operation@is-indonesia.com","it@is-indonesia.com"];
moment.lang("en-US");
var NORMAL_COLOR = "#000000";
var NORMAL_STYLE = SpreadsheetApp.BorderStyle.SOLID;
var SEPARATOR_COLOR = "#FF0000";
var SEPARATOR_STYLE = SpreadsheetApp.BorderStyle.SOLID_THICK;
var HELPER = {
  SS :  null,
  DEBUG : false,
  CURR_YEAR : 2017,
  formatDateRaw : "YYYY-MM-DD HH:mm:ss",
  formatDateFile : "YYYY_MM_DD-HH_mm_ss",
  formatDateStr : "DD-MMM",
  nowRaw : function(){
    return moment().format(this.formatDateRaw);
  },
  nowStr : function(){
    return moment().format(this.formatDateStr);
  },
  /**
  * @param {string} before argument substractor
  * @param {string} after argument which substracted
  * @return true if pass 24 hour from before to after
  */
  isDiffPass24Hours : function(before,after){

    var m_before = moment(before,this.formatDateStr);
    var m_after;
    if (after==null){
      m_after = moment();
    }else {
      m_after = moment(after,this.formatDateStr);
    }
    
    m_before = m_before.add(24,"h");
    DoLog.append("(" + before + "," + after +") " + m_before.format() + "<=" + m_after.format() + " = " + 
      (m_before.format()<=m_after.format()).toString());
    return m_before.format()<=m_after.format();
  },
  isDiffPass3Month: function(m_before,m_after){
    
    m_before = m_before.add(3,"M");
    DoLog.append( m_before.format(this.formatDateRaw) + "<=" + m_after.format(this.formatDateRaw) + " = " + 
      (m_before.format()<=m_after.format()).toString());
    return m_before.format()<=m_after.format();
  },
 
  toCellAddr : function(row,col){
    DoLog.append(String.fromCharCode(64+col) + row.toString());
    return String.fromCharCode(64+col) + row.toString();
  },
  viewArchiveSheet : function (name){
    var APP = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_name = "ARCHIVE - " + name;
    var current = APP.getSheetByName(sheet_name);
    var reference = APP.getSheetByName(name);
    if ((current == null) && 
        (reference!=null)){
      APP.insertSheet(
        sheet_name,1);
      SpreadsheetApp.flush();
      current = APP.getSheetByName(sheet_name);
      
      this.duplicateHeaderSheet(reference,current);
    }
    return current;
  },
  HEADER_MAX_ROW : 2,
  duplicateHeaderSheet : function(ref_sheet,new_sheet){
    var max_col_ref = ref_sheet.getMaxColumns(),
        max_col_new = new_sheet.getMaxColumns();
    if (max_col_new> max_col_ref){
      new_sheet.deleteColumns(max_col_ref+1,max_col_new-max_col_ref);
    }else if (max_col_new<max_col_ref) {
      new_sheet.insertColumns(1, max_col_ref-max_col_new)
    }
    var max_row_ref = ref_sheet.getMaxRows();
    var max_row_to_copy = this.HEADER_MAX_ROW;
    if (max_row_ref<max_row_to_copy){
      new_sheet.insertRows(1,max_row_to_copy);
    }
    
    var ref_range = ref_sheet.getRange(1,1,max_row_to_copy,max_col_ref);
    var new_range = new_sheet.getRange(1,1,max_row_to_copy,max_col_ref);
    ref_range.copyTo(new_range);
    new_range = new_sheet.getRange(this.HEADER_MAX_ROW,1,1,max_col_ref);
    new_range.setValue("");
  },
  lastRowDate :null,
  isRowPass3Month : function(sheet,row){
    
    var COL_DATE = [ 1, 2 ] ;
    if (sheet.getName().indexOf("ARCHIVE")!=-1){
      return false;
    }
    var curr_date = moment();
    var m_value,value_date, c_row;
    var curr_range,pass = true;
    var num_empty = 0;
    for (curr in COL_DATE){
      c_col = COL_DATE[curr];
      curr_range = sheet.getRange(row,c_col);
      curr_range.setNumberFormat("@");

      value_date = curr_range.getValue().toString();
      if (value_date == "") {
        num_empty++;
        continue;
      }
      if (this.lastRowDate == value_date){
        continue;
      }
      
      //DoLog.append("\\n" + value_date+" ::::");
      m_value = moment(value_date,this.formatDateStr);
      
      DoLog.append(m_value.format(this.formatDateRaw)+ " -->");
      m_value.year(this.CURR_YEAR);
      DoLog.append(m_value.format(this.formatDateRaw)+ "\\n");
      if (!this.isDiffPass3Month(m_value,curr_date)){
        pass = false;
      }else if (curr == 0 ){
        this.lastRowDate = value_date;
      }
      
    }
    DoLog.show();
    if (num_empty>=2){
      return false;
    }
    return pass;
      
  },
  isArrRowPass3Month : function (values){
    var COL_DATE = [ 1, 2 ] ;
   
    var m_value,value_date, c_row;
    var curr_range,pass = true;
    var num_empty = 0;
    var curr_date = moment();
    for (curr in COL_DATE){
      c_col = COL_DATE[curr]-1;
      
      value_date = values[c_col];
      if (value_date == "") {
        num_empty++;
        continue;
      }
      if (this.lastRowDate == value_date){
        continue;
      }
      
      //DoLog.append("\\n" + value_date+" ::::");
      m_value = moment(value_date,this.formatDateStr);
      
      DoLog.append(m_value.format(this.formatDateRaw)+ " -->");
      m_value.year(this.CURR_YEAR);
      DoLog.append(m_value.format(this.formatDateRaw)+ "\\n");
      if (!this.isDiffPass3Month(m_value,curr_date)){
        pass = false;
      }else if (curr == 0 ){
        this.lastRowDate = value_date;
      }
      
    }
    if (!pass) DoLog.show();
    if (num_empty>=2){
      return false;
    }
    return pass;
  },
  isArrRowExists : function (values){
    var COL_DATE_1 = 1, COL_DATE_2 = 2, COL_CLIENT = 5, COL_AGENT = 6;  
    var COL_LIST = [ COL_DATE_1, COL_DATE_2, COL_CLIENT, COL_AGENT];
    var col_no;
    var pass = false;
    for (i_col in COL_LIST) {
      col_no = COL_LIST[i_col]-1;
      result = values[col_no]
      if (result !="") {
        pass = true;
        break;
      }
    }
    return pass;
  },
  isRowExists : function(sheet,row){
  
    if (sheet.getMaxRows()<row)  return false;
    var pass = false;
    var col_no;
    for (i_col in COL_LIST) {
      col_no = COL_LIST[i_col];
      result = sheet.getRange(row,col_no).getValue().toString();
      if (result !="") {
        pass = true;
        break;
      }
    }
    
    return pass;
    
  },
  backupSendEmail : function(email){
    
    var file = Drive.Files.get(HELPER.SS.getId());
    var url = file.exportLinks[MimeType.MICROSOFT_EXCEL];
    
    var response = UrlFetchApp.fetch(url, {
       headers: {
         Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
       }
     });
    
    var name = HELPER.SS.getName();
    var blobs = [response.getBlob().setName(name + ".xlsx")];
    blobs = Utilities.zip(blobs).setName(name+ " " + moment().format(this.formatDateFile)+".zip");

   
    MailApp.sendEmail(email, 
       HELPER.SS.getName()+ " Backup " + moment().format(this.formatDateRaw), 
      "Here the backup created by automatic script",{
        attachments : blobs
      });
  },
  fill1DArray : function(i,ref,new_value){
    if (new_value.length){
      var ii = 0,imax = new_value.length;
      for (ii=0;ii<imax;ii++){
        ref[i][ii] = new_value[ii];
      }
    }else {
      ref[i] = new_value;
    }
  },
  push2DArray : function(ref,new_value){
    var i = 0, max = ref.length;
    for (i=0;i<max-1;i++){
      this.fill1DArray(i,ref,ref[i+1]);
    }
    this.fill1DArray(max-1,ref,new_value);
  }
}

HELPER.SS = SpreadsheetApp.getActiveSpreadsheet();
HELPER.SS.getId()

function doTest(){
  HELPER.viewArchiveSheet("2017");
  /*var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2017");
  DoLog.append("5 = exists " + HELPER.isRowExists(sheet,5));
  HELPER.isRowPass3Month(sheet,5);
  HELPER.isRowPass3Month(sheet,20);
  HELPER.isRowPass3Month(sheet,30);
  HELPER.isRowPass3Month(sheet,40);
  
  DoLog.show();*/
  /*var curr_xy = sheet.getActiveCell();
  
  //HELPER.checkRowArchievable(sheet,13);
  /*HELPER.DEBUG = true;
  HELPER.checkRowArchievable(sheet,parseInt(Browser.inputBox("Please insert row no (" + curr_xy.getRow() +")"),10));
  HELPER.DEBUG = false;*/
}

var DoLog =  {
  type: 1,
  enabled:0,
  s_logs:"",
  append: function(new_log){
    if (this.enabled == 0 ) return 0;
    switch(this.type){
      case 1:
        this.s_logs += new_log;
        this.s_logs += "\n";
      break;
      case 2:
        Logger.log(new_log);
      break;
    }

  },
  show : function(){
    if (this.enabled == 0 ) return 0;
    var all_logs;
    switch(this.type){
      case 1:
        all_logs =  this.s_logs;
        this.s_log = "";
      break;
      case 2:
        all_logs = Logger.getLog();
        Logger.clear();
      break;
    }
    Browser.msgBox(all_logs);
  },
  simpleShow : function(message){
    if (this.enabled == 0 ) return 0;
    Browser.msgBox(message);
  }
};
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  
  if (Session.getActiveUser()== null) 
    return 0;
  var email =  Session.getActiveUser().getEmail(); 
  if (email != "reservation.isindonesia@gmail.com" && 
      email != "gd.wirawan@gmail.com"){
    return 0;
  }
  var menu = ui.createMenu('Daily-Helper');
  menu.addItem('Add 50 Rows', 'addMoreRows')
      .addSeparator();
  menu.addItem('Archive Selected', 'archiveSelectedRange')
      .addItem('Archive Past 3 Month','archive3Monthv2')
      .addSeparator();
  menu.addItem("Move Selected to Top", "moveSelectedRangeToTop")
      .addSeparator();
  menu.addItem("Create Monthly Separator","createMonthSeparator")
  // FOR TESTING
      //.addItem('test','doTest')
  // END FOR TESTING
      .addToUi();

}

function addMoreRows(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var REF_ROW = 5;
  
  var END_COL = sheet.getMaxColumns();
  var END_ROW = sheet.getMaxRows();
  
  sheet.insertRowsAfter(END_ROW, 50);
  var range_ref = sheet.getRange(REF_ROW,1, 1,END_COL);
  
  var C_ROW = END_ROW+1;
  var range_new = sheet.getRange(C_ROW,1, 1,END_COL);
  range_ref.copyTo(range_new);
  
  range_new.setBorder(false, true, false, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  range_new.setValue("");
  
  for (var i = 1 ;i<50;i++){
    C_ROW++;
    
    range_new = sheet.getRange(C_ROW,1,1,END_COL);
    range_ref.copyTo(range_new);
   
  } 
}
function archive3Monthv2(){
  var APP = SpreadsheetApp.getActiveSpreadsheet();
  HELPER.backupSendEmail(email_to_send);
  var current_sheet = APP.getSheetByName(SHEET_NAME);
  var archive_sheet = HELPER.viewArchiveSheet(SHEET_NAME);
  
  if (archive_sheet ==null){
    return 0;
  }
  var COL_NUM = current_sheet.getMaxColumns(), 
      START_ROW = HELPER.HEADER_MAX_ROW,
      COL_UPDATE,
      CURR_MAX_COL = 6,
      CURR_MAX_ROW = current_sheet.getMaxRows(),
      MAX_NUM_ROW = 1000;
  
 
  var num_row = 0,i_row_start,
      i_value, max_value,
      i_row_num ,rng_copy,
      result_num_row;

  var range_value,curr_value;
  i_row_start = START_ROW;
  result_row = 0;
  while (true){
    i_row_num = MAX_NUM_ROW;
    if (i_row_start>=CURR_MAX_ROW){
      break;
    }
    if (i_row_start+i_row_num>CURR_MAX_ROW) {
      i_row_num = CURR_MAX_ROW - i_row_start+1;
    }
   
    Logger.log("Mulai while " + i_row_start + " num " + i_row_num + " of " + CURR_MAX_ROW);
    
    current_range = current_sheet.getRange(i_row_start, 1, i_row_num, CURR_MAX_COL);
    current_range.setNumberFormat("@");
    
    range_value =  current_range.getValues();
    i_value = -1;
    max_value = range_value.length;
    curr_value = range_value[i_value+1];
    while (HELPER.isArrRowExists(curr_value) &&
           HELPER.isArrRowPass3Month(curr_value) ){
      i_value++;
      if (i_value+1>=max_value) break;
      curr_value = range_value[i_value+1];
    }
    Logger.log(i_value + ":" + curr_value.join(",") + " exists "+HELPER.isArrRowExists(curr_value) +
      " pass 3 month " +   HELPER.isArrRowPass3Month(curr_value));
    result_num_row = i_row_start -START_ROW+1+ i_value;
    if (i_value+1<max_value){
      break;
    }
    if (i_row_num<=MAX_NUM_ROW){
      break;
    }
    i_row_start = i_row_start+ MAX_NUM_ROW;
  }
  if (result_num_row<=0){
    Browser.msgBox("Nothing to Move");
    return false;
  }
   rng_copy = current_sheet.getRange(START_ROW, 1,result_num_row, COL_NUM);
  var log_start = current_sheet.getRange(START_ROW,1,1,2).getValues()[0].join("/"),
      log_end = current_sheet.getRange(rng_copy.getRow() + rng_copy.getNumRows()-1,1,1,2).getValues()[0].join("/");
      
  
  Browser.msgBox("Number of Row Moved :" + result_num_row + "\\n From " + 
       log_start + " s/d " + log_end);
  
  var ar_row = archive_sheet.getLastRow();
  archive_sheet.insertRowAfter(ar_row);
  ar_row++;
  SpreadsheetApp.flush();
  var range_dst = archive_sheet.getRange(ar_row,1,1,COL_NUM);
  rng_copy.copyTo(range_dst);
  

  current_sheet.deleteRows(START_ROW, result_num_row-1);
 
  SpreadsheetApp.flush();
}
function archiveSelectedRange(){
  var APP = SpreadsheetApp.getActiveSpreadsheet();
  var c_sheet = APP.getActiveSheet();
  var c_range = APP.getActiveRange();
  var START_ROW = c_range.getRow();
  var archive_sheet = HELPER.viewArchiveSheet(SHEET_NAME);
  var COL_NUM = c_sheet.getMaxColumns();
  var result_num_row = c_range.getNumRows();
  if (c_range.getNumColumns()<COL_NUM){
    c_range = c_sheet.getRange(START_ROW,1,result_num_row,COL_NUM);
  }
  
  var ar_row = archive_sheet.getLastRow();
  archive_sheet.insertRowAfter(ar_row);
  SpreadsheetApp.flush();
  ar_row++;
  var range_dst = archive_sheet.getRange(ar_row,1,1,COL_NUM);
  
  c_range.copyTo(range_dst);
  c_sheet.deleteRows(START_ROW, result_num_row);
  SpreadsheetApp.flush();
}
function moveSelectedRangeToTop(){
  var APP = SpreadsheetApp.getActiveSpreadsheet();
  var c_sheet = APP.getActiveSheet();
  var c_range = APP.getActiveRange();
  var START_ROW = c_range.getRow();
 
  var target_sheet = APP.getSheetByName(SHEET_NAME);
  var COL_NUM = c_sheet.getMaxColumns();
  var result_num_row = c_range.getNumRows();
  if (c_range.getNumColumns()<COL_NUM){
    c_range = c_sheet.getRange(START_ROW,1,result_num_row,COL_NUM);
  }
  
  var ar_row =HELPER.HEADER_MAX_ROW;
  target_sheet.insertRowBefore(HELPER.HEADER_MAX_ROW);
  var range_dst = target_sheet.getRange(ar_row,1,1,COL_NUM);
  
  c_range.copyTo(range_dst);
  c_sheet.deleteRows(START_ROW, result_num_row);
  SpreadsheetApp.flush();
}
function createMonthSeparator(){
  
   var APP = SpreadsheetApp.getActiveSpreadsheet();
   var c_range = APP.getActiveRange();
   var c_sheet = APP.getActiveSheet();
   if (c_range.getNumRows()!=2) {
      return false;
   }
  var MAX_COL = c_sheet.getMaxColumns();
   c_range = c_sheet.getRange(c_range.getRow(),1, 1,MAX_COL);
   c_range.setBorder(null,null,true,null,null,null,SEPARATOR_COLOR,SEPARATOR_STYLE);
   c_range.setBorder(true,true,null,true,true,true,NORMAL_COLOR,NORMAL_STYLE);
   c_range = c_sheet.getRange(c_range.getRow()+1,1,1,MAX_COL);
   c_range.setBorder(true,null,null,null,null,null,SEPARATOR_COLOR,SEPARATOR_STYLE);
   c_range.setBorder(null,true,true,true,true,true,NORMAL_COLOR,NORMAL_STYLE);
}
