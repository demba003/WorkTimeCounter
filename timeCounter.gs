function generateTimeReport() {
  if (!initSettings()) {
   // Polish translation
  const DATE = "Data"
  const FROM = "Od"
  const TO = "Do"
  const WORKING_HOURS = "Godzin do zrobienia:"
  const HOURS_COUNT = "Ilość godzin"
  const HOURS_TOTAL = "Godzin w sumie:"
  const HOURS_LEFT = "Pozostało:"
  const CONFIG = "Ustawienia"
  const calendarID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG).getRange(1, 2) .getValue()
  
  const mySheet = SpreadsheetApp.getActiveSheet();
  const myCalendar = CalendarApp.getCalendarById(calendarID)
  const timeToCount = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName().split(" ");
  const currentMonth = fromRoman(timeToCount[0]);
  const currentYear = timeToCount[1];
  
  var index = 1
  
  mySheet.clear()
  mySheet.appendRow([DATE, FROM,  TO, HOURS_COUNT]);
  
  for (index = 1; index <= daysInMonth(currentMonth, currentYear); index++) {
    var outputDate = index + "." + currentMonth + "." + currentYear
    var day = new Date(currentYear, currentMonth - 1, index)
    var events = myCalendar.getEventsForDay(day)
    
    var from = ""
    var to = ""
    var time = ""
    
    if (day.getDay() == 6 || day.getDay() == 0) {
      mySheet.getRange("A" + (index + 1) + ":D" + (index + 1) ).setBackground("#ea5b5b")
    }
    
    if (Array.isArray(events) && events.length >= 1) {
      from = events[0].getStartTime()
      to =  events[0].getEndTime()
      time = events[0].getEndTime().getTime() - events[0].getStartTime().getTime()
      
      mySheet.appendRow([outputDate, formatOutputTime(from), formatOutputTime(to), msToTime(time)])
      mySheet.getRange("D" + (index + 1)).setNumberFormat("hh:mm")
    } else {
      mySheet.appendRow([outputDate])
    }
  }
  
  mySheet.appendRow([" "])
  mySheet.appendRow([WORKING_HOURS, "='" + CONFIG + "'!C" + (4 + currentMonth), HOURS_TOTAL, "=SUM(D2:D" + index + ")", HOURS_LEFT, "=B" + (index + 2) + "-SUM(D2:D" + index + ")"])
  //mySheet.getRange("B" + (index + 2)).setNumberFormat("[hh]:mm")
  //mySheet.getRange("D" + (index + 2)).setNumberFormat("[hh]:mm")
  //mySheet.getRange("F" + (index + 2)).setNumberFormat("[hh]:mm")
  }
}

function formatOutputTime(date) {
  var hours = date.getHours()
  var minutes = date.getMinutes()
  if (minutes == "0") 
    minutes = "00"
    return hours + ":" + minutes
}

function daysInMonth (month, year) {
    return new Date(year, month, 0).getDate();
}

function msToTime(duration) {
  var minutes = parseInt((duration / (1000 * 60)) % 60)
  var hours = parseInt((duration / (1000 * 60 * 60)) % 24)

  hours = (hours < 10) ? "0" + hours : hours;
  minutes = (minutes < 10) ? "0" + minutes : minutes;

  return hours + ":" + minutes
}

function fromRoman(str) {  
  var result = 0;
  var decimal = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  var roman = ["M", "CM","D","CD","C", "XC", "L", "XL", "X","IX","V","IV","I"];
  for (var i = 0; i <= decimal.length; i++) {
    while (str.indexOf(roman[i]) === 0){
      result += decimal[i];
      str = str.replace(roman[i], "");
    }
  }
  return result;
}

function initSettings() {
  const CALENDAR_ID = "ID kalendarza"
  const PART_OF_WORK = "Część etatu"
  const MONTH = "Miesiąc"
  const HOURS = "Godziny"
  const MY_HOURS = "Moje godziny"
  const CONFIG = "Ustawienia"
  
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG) == null) {
    var configSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG, 0);
    configSheet.appendRow([CALENDAR_ID, "insert_id_here@group.calendar.google.com"]);
    configSheet.appendRow([PART_OF_WORK, "0,6"]);
    configSheet.appendRow([" "]);
    configSheet.appendRow([PART_OF_WORK, HOURS]);
    configSheet.appendRow(["I", "176", "=B5*$B$2"]);
    configSheet.appendRow(["II", "160", "=B6*$B$2"]);
    configSheet.appendRow(["III", "168", "=B7*$B$2"]);
    configSheet.appendRow(["IV", "168", "=B8*$B$2"]);
    configSheet.appendRow(["V", "168", "=B9*$B$2"]);
    configSheet.appendRow(["VI", "152", "=B10*$B$2"]);
    configSheet.appendRow(["VII", "184", "=B11*$B$2"]);
    configSheet.appendRow(["VIII", "168", "=B12*$B$2"]);
    configSheet.appendRow(["IX", "168", "=B13*$B$2"]);
    configSheet.appendRow(["X", "184", "=B14*$B$2"]);
    configSheet.appendRow(["XI", "152", "=B15*$B$2"]);
    configSheet.appendRow(["XII", "160", "=B16*$B$2"]);
    return true
  }
  return false
}
