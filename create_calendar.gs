var SS;
var moment = Moment.load();
var CONFIG_EMAIL = []
var CONFIG_ATT_ROW = 1, CONFIG_ATT_NUM = 2, CONFIG_ATT_VALUE = 3;
var RESULT_IDX_TIMESTAMP = 0, RESULT_IDX_DATE = 1, RESULT_IDX_REASON =2, RESULT_IDX_CHARGE = 3;
var SHEET_RESPONSE = "Result";
var DATE_TIME_FORMAT = "YYYY-MM-DD HH:mm:ss";
var DATE_FORMAT = "YYYY-MM-DD";
var WEEKDAY_FORMAT ="e";
var CALENDAR_ID = "44kobgclad183kd3gl6gasmd08@group.calendar.google.com";
var TITLE_PREFIX = "Morning :";
var CONFIG_COLOR = {
  "Theo" : CalendarApp.EventColor.BLUE,
  "Haryata" : CalendarApp.EventColor.GREEN};
var EMAIL_NOTIF = "theo@is-indonesia.com";
var APP = {
  SS : null,
  c_sheet : null,
  r_sheet : null,
  task_cal : null,
  getEmailAttendance: function(){
    if (CONFIG_EMAIL.length <=0){
      var c_sht = APP.c_sheet;
      var email = c_sht.getRange(CONFIG_ATT_ROW, CONFIG_ATT_NUM).getValue();
      
      var num = parseInt(email);
      var email_result = [];
      var col = CONFIG_ATT_VALUE;
      
      for (var i =0; i<num ; i++ ){
        email = c_sht.getRange(CONFIG_ATT_ROW, col).getValue();
        email_result.unshift(email);
        col++;
      }
      CONFIG_EMAIL = email_result;
    }
    return CONFIG_EMAIL
  },
  
  setScheduleWeek : function(start_date,in_charge){
    var num_weekends = 0 ;
    var start_date = moment(start_date);
    var c_date = start_date;
    var n_in_charge = in_charge.length, c_in_charge = 0;
    var week_ke = "";
    var cal = APP.task_cal;
    var who_in_charge, new_event;
    while (num_weekends<2){
      week_ke = parseInt(c_date.format(WEEKDAY_FORMAT));
      if (week_ke == 0){
        num_weekends++;
      }
      if (week_ke>=1 && week_ke<=5) {
        who_in_charge = in_charge[c_in_charge++];
        if (c_in_charge>=n_in_charge){
          c_in_charge = 0;
        }
        new_event = cal.createAllDayEvent(TITLE_PREFIX+ who_in_charge, c_date.toDate());
        new_event.setColor(CONFIG_COLOR[who_in_charge]);
      }
      c_date = c_date.add("days",1);
    }  
  },
  getLastSchedule : function(){
    var last_row = APP.r_sheet.getLastRow(); 
    var value = APP.r_sheet.getRange(last_row, 1,1,5).getValues();
    value = value[0];
    return value;
  },
  getListCharge : function(input){
    email = [];
    if (input == "Theo") {
      email = ["Theo","Haryata"];
    }else if (input == "Haryata") {
      email = ["Haryata","Theo"];
    }
    return email;
  },
  deleteAllSchedule : function(since,much_week){
    var cal = APP.task_cal;
    var end = moment(since);
    end = end.add("weeks",much_week);

    var all_events = cal.getEvents(since, end.toDate());
    var c, event,title,week_no;
    var title_num = TITLE_PREFIX.length;
    for (c in all_events){
      event = all_events[c];
      week_no = parseInt(moment(event.getAllDayStartDate()).format(WEEKDAY_FORMAT));
      if (week_no>=1 && week_no<=5){
        title = event.getTitle().toString(); 
        if (title.substr(0,title_num) == TITLE_PREFIX){
          event.deleteEvent();
        }
      }
    }
  },
  sendNotification : function(last_record){
    var result ;
    var report ="";
    result = last_record[RESULT_IDX_TIMESTAMP];
    report += moment(result).format(DATE_TIME_FORMAT) + "\n" + "\n";
    
    result = last_record[RESULT_IDX_DATE];
    report +=  "Changes for Start Date : " + moment(result).format(DATE_FORMAT) + "\n" + "\n";
    
    result = last_record[RESULT_IDX_REASON];
    report +=  "Reason : " + result + "\n" + "\n";
    
    result = last_record[RESULT_IDX_CHARGE];
    report +=  "In Charge : " + result + "\n" + "\n";
    
    MailApp.sendEmail("theo@is-indonesia.com", "Morning in Charge Issued",report);
  }
}
APP.SS = SpreadsheetApp.getActive();
APP.r_sheet = APP.SS.getSheetByName("Result");
APP.c_sheet = APP.SS.getSheetByName("CONFIG");
APP.task_cal = CalendarApp.getCalendarById(CALENDAR_ID);

function onFormSubmit(e){
  MailApp.sendEmail("theo@is-indonesia.com","Testing something" ,JSON.stringify(e.namedValues));
  return false;
  var last_record = APP.getLastSchedule();
  var input_date = last_record[RESULT_IDX_DATE];
  var input_charge = last_record[RESULT_IDX_CHARGE];
  
  input_date = moment(input_date).toDate();
  
  
  var list_in_charge = APP.getListCharge(input_charge);
 
  
  APP.sendNotification(last_record);
  APP.deleteAllSchedule(input_date,2);
  APP.setScheduleWeek(input_date,list_in_charge);
}
function MENU_scheduleMorning(){
  var last_record = APP.getLastSchedule();
  var input_date = last_record[RESULT_IDX_DATE];
  var input_charge = last_record[RESULT_IDX_CHARGE];
  APP.deleteAllSchedule(input_date,2);
  APP.setScheduleWeek(input_date,input_charge);
}
function MENU_reattach(){
  var SS = SpreadsheetApp.getActive();
  var sheet = SS.getSheetByName("Morning Responsibility");
  var form = FormApp.create("Morning Responsibility Form");
  
  form.setDestination(FormApp.DestinationType.SPREADSHEET, SS.getId());
  form.setTitle("Morning Responsibility");
  form.setDescription("It will create morning in charge IT for every 2 days, shifting, for 2 weeks.");
  form.addDateItem()
      .setTitle("Start Date")
      .setHelpText("Please assign the start date (must in the work day)")
      .setRequired(true)
      .setIncludesYear(false);
  form.addMultipleChoiceItem()
      .setTitle("Who is starting")
      .setHelpText("IT Staff List")
      .setChoiceValues(["Theo","Haryata"])
      .setRequired(true);
  var validation = FormApp.createParagraphTextValidation().requireTextLengthGreaterThanOrEqualTo(5)
      .setHelpText("Please fill comment minimal 5 character")
      ;
  
  form.addParagraphTextItem()
      .setTitle("Comment")
      .setHelpText("Information why it is submitted")
      .setValidation(validation);
  ScriptApp.newTrigger("onFormSubmit").forSpreadsheet(SS.getId()).onFormSubmit();
}

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  
  var menu = ui.createMenu('Task');
      menu.addItem('Schedule Morning', 'MENU_scheduleMorning')
      menu.addItem('ReInstall Event', 'MENU_reattach');
  
  menu.addToUi();
}
