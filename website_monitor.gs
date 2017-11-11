var moment = Moment.load();
moment().zone(8);
var FORMAT_DATE_TIME = "YYYY-MM-DD HH:mm:ss";
var FORMAT_FILE = "YYYY-MM-DD_HH_mm_ss";
var EMAIL_TO_SEND = ["it@is-indonesia.com","it.ics.ind@gmail.com"];
var MAX_LOG = 120;
var SHEET_SUMMARY_NAME = "(SUMMARY)";

function testURL(url){
  var OPTIONS = {
    validateHttpsCertificates : false,
  };
  var before = new Date();
  var delay = before.getTime();
  var response_code = 0;
  try {
    var response = UrlFetchApp.fetch(url, OPTIONS );
    response_code = response.getResponseCode();
  }catch (e){
    response_code = 404;
  }
  var delay = new Date().getTime() - delay;
  return [moment().format(FORMAT_DATE_TIME),
          delay,
          response_code]
}
function addConclusion(result,criteria){
  var conclusion = "BAD";
  if (result[2]==200){
    if (result[1]<=criteria[0]){
      conclusion= "GREAT";
    }else if (result[1]<=criteria[1]){
      conclusion = "GOOD";
    } else if (result[1]<=criteria[2]){
      conclusion = "AVERAGE";
    }
  }
   result.push(conclusion)
  return result;
}
function doTest(){
  /*var file_active = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  var name = file_active.getName();
  var folder = file_active.getParents().next().getId();
  var new_file = file_active.makeCopy(folder);
  new_file.setName(name + " " + moment().format(FORMAT_FILE)); */
  var properties = PropertiesService.getScriptProperties()
  Logger.log(properties.getProperty("notify_start"))
  Logger.log(properties.getProperty("notify_end"))
  Logger.log(moment().format("HH"));
}

function summarize_health(range_values){
  var summarize = {};
  for (var i in range_values){
    each_range =  range_values[i][0];
    
    var value = each_range.toString();
    if (!summarize[value]){
      summarize[value] = 0;
    }
    summarize[value]++;
  }
  return summarize;
  
}


function healthChecking(){
   var START_LOG_ROW = 4, COL_CHECK = 4, NUM_CHECK = 10,
      LAST_STATUS_ROW = 2, LAST_STATUS_COL = 1;
   var ALL_URGENT = [], MISS_STATUS = [];
   var APP = SpreadsheetApp.getActive();
  
   var sheets = APP.getSheets();
   var statuses, curr_sheet, result;
   var warning_status = ["AVERAGE","BAD"];
   var urgent_count = 0, url;
   var last_update, today = moment();
   var s_today = today.format(FORMAT_DATE_TIME)
  
   
   var diff_time, sum_result, sum_color;
   var sum_sheet = APP.getSheetByName(SHEET_SUMMARY_NAME);
   var sum_row = 3;
  
  
   for (var i=0,i_max = sheets.length;i<i_max ; i++){
     curr_sheet = sheets[i];
     if (curr_sheet.getName() == SHEET_SUMMARY_NAME) 
       continue
     
     urgent_count = 0;
     
     statuses = curr_sheet.getRange(START_LOG_ROW, COL_CHECK,NUM_CHECK,1).getValues();
     
     last_update = curr_sheet.getRange(LAST_STATUS_ROW,LAST_STATUS_COL).setNumberFormat("@").getValue();
     last_update = moment(last_update,FORMAT_DATE_TIME)
     
     
     sum_color = "black";
     diff_time=  today.diff(last_update,"minutes");
     if (diff_time>30){
       MISS_STATUS.push(curr_sheet.getSheetName());
       sum_color ="red";
     }
     
     
     var status_value;
     for (var i_statuses= 0 , i_statuses_max = statuses.length; 
          i_statuses< i_statuses_max; 
          i_statuses++){
        status_value = statuses[i_statuses][0].toString(); 
       
       if (status_value == warning_status[0] || 
           status_value == warning_status[1]){
         urgent_count++;
       }else {
         break;
       }
     }
     
     Logger.log(curr_sheet.getSheetName() + " with urgent " + urgent_count);
     if (urgent_count>=3){
       ALL_URGENT.push(curr_sheet.getSheetName());
       sum_color ="#B0B000";
     }
     
     result = JSON.stringify(summarize_health(statuses));
     
     sum_sheet.getRange(sum_row++, 1,1,3).setValues([
       [curr_sheet.getName(),result,diff_time]]).setFontColor(sum_color)
     
   }
  
  
  var properties = PropertiesService.getScriptProperties();
  var last_message =  properties.getProperty("miss_message").trim();
  if (ALL_URGENT.length>0 || MISS_STATUS.length>0){
    var message = [];
    if (ALL_URGENT.length>0) {
      message.push("This sites has issue: " + ALL_URGENT.join(", ") + " \n\n" );
    }
    if (MISS_STATUS.length>0){
      message.push("And We have missing sites " + MISS_STATUS.join(", ") + " \n\n" );
    }
    
    var num_send = parseInt(properties.getProperty("send_email"));
    num_send = parseInt(num_send) + 0 ;
    
   
    var today_hour = moment().format("HH");
     
    if (num_send<=10 ){
          
      if (properties.getProperty("notify_start")<=today_hour && 
          properties.getProperty("notify_end")>=today_hour) {
        message.push("\n");
        message.push("Please see url for more detail: " + APP.getUrl() );
        message.unshift("\n");
        
        if (last_message!="") {
          message.unshift(last_message);
          message.unshift("=========================");
          message.unshift("RECAP PASS END OF THE DAY");
        }
        
        MailApp.sendEmail(EMAIL_TO_SEND, 
                          "IT ICS : " + ALL_URGENT.length + 
                          " CRITICAL, " + MISS_STATUS.length +
                          " MISSING ", message.join("\n"));

        properties.setProperty("miss_message","")
        num_send++
      }else {
        message.unshift("======================================================");
        message.unshift("On " + moment().format(FORMAT_DATE_TIME)+ ALL_URGENT.length + 
                          " CRITICAL, " + MISS_STATUS.length +
                          " MISSING ");
        
        last_message += "\n" + message.join("\n");
        properties.setProperty("miss_message", last_message)
      }
    }
    properties.setProperty("send_email", num_send);
  }else if (last_message!="" && 
            properties.getProperty("notify_start")<=today_hour && 
            properties.getProperty("notify_end")>=today_hour) {
   
    message.unshift(last_message);
    message.unshift("=========================");
    message.unshift("RECAP PASS END OF THE DAY");
              
    MailApp.sendEmail(EMAIL_TO_SEND, message.join("\n"));

    properties.setProperty("miss_message","");
    num_send++
  } else {
    properties.setProperty("send_email", 0);
  }
}
  
function sendEmail(email){
    
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
       " Website Urgent ISSUE " + moment().format(this.formatDateRaw), 
      "Here the backup created by automatic script",{
        attachments : blobs
      });
}

function doTesting() {
   var START_LOG_ROW = 4;
   var LAST_UPDATE = 2;
   var APP = SpreadsheetApp.getActive();
   var sheets = APP.getSheets();
   var url, curr_sheet, result,criteria;
   for (var i=0,i_max = sheets.length;i<i_max ; i++){
     curr_sheet =  sheets[i];
     url = curr_sheet.getRange(1, 1).getValue().toString();
     if (url=="") 
       continue;
     if (curr_sheet.getName() == "(SUMMARY)") 
       continue;
     curr_sheet.getRange(LAST_UPDATE, 1).setValue(moment().format(FORMAT_DATE_TIME));
     criteria = curr_sheet.getRange(1,2,1,3).getValues()[0]
     Logger.log(criteria.join(", "));
     result = testURL(url);
     result = addConclusion(result,criteria);
     curr_sheet.insertRows(START_LOG_ROW);
     curr_sheet.getRange(START_LOG_ROW, 1,1,result.length).setValues([result]);
     if (curr_sheet.getMaxRows()>=MAX_LOG){
       curr_sheet.deleteRow(MAX_LOG);
     }
   }
  healthChecking();
}
function email_testing(){
  Browser.msgBox(moment().duration().asMinute());
}

function onOpen(){
  var UI = SpreadsheetApp.getUi();
  UI.createMenu("Running")
    .addItem("Create Testing", "doTesting")
  .addToUi();
  
}
