var moment = Moment.load();

var HELPER = {
  DEBUG : false,
  formatDateRaw : "YYYY-MM-DD HH:mm:ss",
  formatDateStr : "DD/MM/YYYY HH:mm:ss",
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
  isDiffPass2Week: function(before,after){
    var m_before = moment(before,this.formatDateStr);
    var m_after;
    if (after==null){
      m_after = moment();
    }else {
      m_after = moment(after,this.formatDateStr);
    }
    
    m_before = m_before.add(14,"d");
    DoLog.append("(" + before + "," + after +") " + m_before.format() + "<=" + m_after.format() + " = " + 
      (m_before.format()<=m_after.format()).toString());
    return m_before.format()<=m_after.format();
  },
 
  toCellAddr : function(row,col){
    DoLog.append(String.fromCharCode(64+col) + row.toString());
    return String.fromCharCode(64+col) + row.toString();
  },
  checkRowArchievable : function(sheet,row){
    
    var COL_DATE = 1, COL_STATUS = 7;
    if (sheet.getName()=="ARCHIVE"){
      return false;
    }
    
    var value_date = sheet.getRange(row,COL_DATE).getValue().toString();
    var value_status = sheet.getRange(row,COL_STATUS).getValue().toString();
    
    var result = (value_status!="") && 
      (value_status=="Complete")
    &&
      (value_date=="" || 
      this.isDiffPass2Week(value_date));
    
    if (this.DEBUG == true){
      Browser.msgBox(" "+ value_status + " || " + value_date  +"   " +  
                     ((value_status!="") && (value_status=="Complete")) + "//" +
        (value_date=="" || this.isDiffPass2Week(value_date)) + " === " + result); 
        }
    return result;  
  },
  isRowExists : function(sheet,row){
  
     var COL_FIRST = 1, COL_UPDATE = 9, COL_AGENT = 3;  
    
     var result = 
       (sheet.getRange(row,COL_FIRST)!=null && 
        sheet.getRange(row, COL_FIRST).getValue().toString()!="") ||
       (sheet.getRange(row,COL_UPDATE)!=null && 
        sheet.getRange(row, COL_UPDATE).getValue().toString()!="") ||
       (sheet.getRange(row,COL_AGENT)!=null && 
        sheet.getRange(row, COL_AGENT).getValue().toString()!="")
       
       if (this.DEBUG){
         Browser.msgBox(sheet.getRange(row, COL_FIRST).getValue().toString() + " == " + 
                        sheet.getRange(row, COL_UPDATE).getValue().toString() + " == " + 
           sheet.getRange(row, COL_SOURCE).getValue().toString() + " :: " +  result
           );
       }
     return result;
  },
  isRowPass2Week: function(sheet,row){
    var COL_DATE = 1;
    if (sheet.getName()=="ARCHIVE"){
      COL_DATE = 2;
    }
    var value_date = sheet.getRange(row,COL_DATE).getValue().toString();
    var result = (value_date=="" || 
      this.isDiffPass2Week(value_date));
    
    return result;
  },
}
function doTest(){
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  var curr_xy = sheet.getActiveCell();
  
  //HELPER.checkRowArchievable(sheet,13);
  HELPER.DEBUG = true;
  HELPER.checkRowArchievable(sheet,parseInt(Browser.inputBox("Please insert row no (" + curr_xy.getRow() +")"),10));
  HELPER.DEBUG = false;
}

var DoLog =  {
  type: 1,
  enabled: 0,
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
/**
 * Checking one line 
 *
 * @param {Spreadsheet} current active spreadsheed
 * @param {number} row number
 * @return none
 */
function calcOneLine(curr_sheet,row_no){
  //curr_sheet = SpreadsheetApp.getActive();
  var COL_FIRST_DATE = 1,
      COL_AGENT = 3,
      COL_TYPE = 4, 
      COL_REF = 6,
      COL_STATUS = 7, 
      COL_LAST_UPDATE = 9;
  var today_s = HELPER.nowStr();
  
  // CELL STATUS
  
  var curr_cell_agent = curr_sheet.getRange(HELPER.toCellAddr(row_no,COL_AGENT));
  var curr_cell_status = curr_sheet.getRange(HELPER.toCellAddr(row_no,COL_STATUS));
  var curr_status = curr_cell_status.getValue().toString();

  curr_cell_agent.setFontWeight("bold");
  switch(curr_status){
    case "New":
      curr_cell_agent.setFontColor("black");
      break;
    case "Complete":
      curr_cell_agent.setFontColor("green");
      break;
    default:
      curr_cell_agent.setFontColor("black");
      break;
  }
  
  // CELL FIRST EMAIL COLORING
  
  var curr_cell_date = curr_sheet.getRange(HELPER.toCellAddr(row_no, COL_FIRST_DATE));
  var curr_value = curr_cell_date.getValue().toString();

  if (curr_status=="Complete"){
    curr_cell_date.setFontColor("green");
    curr_cell_date.setFontWeight("bold");
  } else {
    if (HELPER.isDiffPass24Hours(curr_value,today_s)){
      curr_cell_date.setFontColor("red");
      curr_cell_date.setFontWeight("bold");
    }
  }
} 

function onOpen(){
  
  var ui = SpreadsheetApp.getUi();
  
  var menu = ui.createMenu('IB-Workflow');
      menu.addItem('Add 50 Rows', 'addMoreRows')
      .addSeparator();
  
  if (Session.getActiveUser().getEmail() != "booking.is.indonesia@gmail.com"){
    return 0;
  }
  
  menu.addItem('Archive Past 2 Weeks','archive2Week')
      .addItem('Recoloring','checkAllSheet')
  // FOR TESTING
     // .addItem('test','doTest')
  // END FOR TESTING
      .addToUi();
  
  checkAllSheet();
}

function addMoreRows(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var REF_ROW = 4;
  
  var END_ROW = sheet.getMaxRows();
  
  sheet.insertRowsAfter(END_ROW, 50);
  var range_ref = sheet.getRange(HELPER.toCellAddr(REF_ROW,1) + ':' + HELPER.toCellAddr(REF_ROW,10));
  
  var C_ROW = END_ROW+1;
  var range_new = sheet.getRange(HELPER.toCellAddr(C_ROW,1) + ':' + HELPER.toCellAddr(C_ROW,10));
  range_ref.copyTo(range_new);
  
  
  range_new.getCell(1,1).setValue("").setNumberFormat("@").setFontColor("black"); // First email
  range_new.getCell(1,2).setValue("☐").setFontColor("black"); // Send Acknowledge
  range_new.getCell(1,3).setValue("").setFontColor("black"); // Agent
  range_new.getCell(1,4).setValue("").setFontColor("black"); // Type
  range_new.getCell(1,5).setValue("").setFontColor("black"); // Source
  range_new.getCell(1,6).setValue("").setFontColor("black"); // Ref
  range_new.getCell(1,7).setValue("").setFontColor("black"); // Status
  range_new.getCell(1,8).setValue("☐").setFontColor("black"); // Inform Client
  range_new.getCell(1,9).setValue("").setFontColor("black"); // Last Update
  range_new.getCell(1,10).setValue("").setFontColor("black"); // Comment
  
  range_new.setBorder(false, true, false, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  range_ref = range_new; 
  
  for (var i = 1 ;i<50;i++){
    C_ROW++;
    
    range_new = sheet.getRange(HELPER.toCellAddr(C_ROW,1) + ':' + HELPER.toCellAddr(C_ROW,9));
    range_ref.copyTo(range_new);
   
  } 
}
  
function archive2Week(){
  var SSheet = SpreadsheetApp.getActiveSpreadsheet();
  var archive_sheet = SSheet.getSheetByName("ARCHIVE");
  
  if (archive_sheet ==null){
    return 0;
  }
  var COL_NUM = 10, 
      START_ROW = 4,
      COL_UPDATE = 9;
  
  var all_sheets = SSheet.getSheets();
  
  var range_dst = archive_sheet.getRange(START_ROW,2,1,COL_NUM+1);
  
  var name_dst = archive_sheet.getRange(START_ROW,1);
  
  var ctr_sheet, user,i_r,rng_copy;
  for (var sh_i=0,sh_len = all_sheets.length; sh_i < sh_len ; sh_i++ ){
    ctr_sheet = all_sheets[sh_i];
    user = ctr_sheet.getName();
    if (ctr_sheet.getName()=="ARCHIVE") 
      continue;
    
    i_r = START_ROW; 
    HELPER.is
    while (HELPER.isRowExists(ctr_sheet,i_r) && 
           HELPER.isRowPass2Week(ctr_sheet,i_r)){
      //Browser.msgBox(user +  " : " + i_r);
      if (HELPER.checkRowArchievable(ctr_sheet,i_r)){
        rng_copy = ctr_sheet.getRange(i_r,1,1,COL_NUM);
        
        archive_sheet.insertRowBefore(START_ROW);
        name_dst.setValue(user);
        name_dst.setNumberFormat("@")
        rng_copy.copyTo(range_dst);
        ctr_sheet.deleteRow(i_r);
        
      }else {
        i_r++;
      }
    }
  }
  
}
  
function checkAllSheet(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var curr_sheet;
  var all_sheet = spreadsheet.getSheets();
   
  var COL_DATE = 1, 
      COL_CHECK = 2,
      COL_STATUS = 7,
      START_ROW = 4;
  
  for (var i= 0;i<all_sheet.length; i++){
    curr_sheet = all_sheet[i];
    
    if (curr_sheet.getName()=="ARCHIVE") 
      continue;
    
    var curr_r = START_ROW;
    var curr_cell_date = curr_sheet.getRange(HELPER.toCellAddr(curr_r, COL_DATE));
    var curr_value = curr_cell_date.getValue().toString();
  
    while (HELPER.isRowExists(curr_sheet,curr_r)){
      //curr_sheet = sheet.getSheetByName(curr_sheet.getName());
      calcOneLine(curr_sheet,curr_r);
      
      curr_r ++;
      curr_cell_date = curr_sheet.getRange(HELPER.toCellAddr(curr_r, COL_DATE));
      curr_value = curr_cell_date.getValue().toString();
    }
  
  }
}

function onEdit(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var r = sheet.getActiveCell();
  var curr_c = r.getColumn(); 
  var curr_r = r.getRow(); 
  var curr_value = r.getValue().toString();
  var START_ROW = 4;
  
  
  if (sheet.getName()=="ARCHIVE"){
    return 0;
  }
  if (curr_r <START_ROW){
    return 0;
  }
  
  var COL_DATE = 1, 
      COL_AGENT = 3, 
      COL_TYPE = 4, 
      COL_REF = 6,
      COL_STATUS = 7,
      COL_COMMENT = 10,
      COL_UPDATE = 9;

  var time = HELPER.nowStr();
  var r_color, r_date,r_curr = r;
  switch (curr_c){
    case COL_TYPE:
      r_color = sheet.getRange(HELPER.toCellAddr(curr_r,COL_DATE));
      r_color.setValue(time);
      r_color.setNumberFormat("@");
      r_color.setFontColor("black");
      r_color.setFontWeight("bold");
      
      r.setFontWeight("bold")
      switch (curr_value){
        case "New Request":
          r.setFontColor("green");
          break;
        case "Follow Up":
          r.setFontColor("black");
          break;
      }
    break;
    case COL_STATUS:
      r_date = sheet.getRange(HELPER.toCellAddr(curr_r,COL_UPDATE));
      r_date.setValue(time);
      r_date.setNumberFormat("@");
      r_date.setFontWeight("bold");
      var cell_reason = sheet.getRange(HELPER.toCellAddr(curr_r,COL_COMMENT));
      switch(curr_value){
        case "On Progress":
          r_date.setFontColor("black");
          SpreadsheetApp.getActive().setActiveSelection(HELPER.toCellAddr(curr_r,COL_COMMENT));
          
          cell_reason.setValue(reason);
          cell_reason.setNumberFormat("@");
          
        break;
        case "Complete":
          r_date.setFontColor("green");
        break;
      }
    break;
    case COL_REF:
      if (curr_value == "") return 0
      var value = '=HYPERLINK("is-intl.net/bookingonline_v2/booking_edit.php?id='+curr_value+'";"'+curr_value+'")';
      r_curr.setValue(value);
      r_curr.setFontLine("underline");
      r_curr.setFontColor("blue");
      r_curr.setNumberFormat("@");
    break;
  }
  
  if (Session.getActiveUser().getEmail()=="booking.is.indonesia@gmail.com"){
    calcOneLine(sheet,curr_r);
  }
}
