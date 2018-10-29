function generateTimeReport() {
  // Change to your calendar ID
  const calendarID = "xxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com"
  // Change to how many hour and minutes per month you work
  const hoursToWork = "110:24"
  
  // Polish translation
  const DATE = "Data"
  const FROM = "Od"
  const TO = "Do"
  const WORKING_HOURS = "Godzin do zrobienia:"
  const HOURS_COUNT = "Ilość godzin"
  const HOURS_TOTAL = "Godzin w sumie:"
  const HOURS_LEFT = "Pozostało:"
  
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
  mySheet.appendRow([WORKING_HOURS, hoursToWork, HOURS_TOTAL, "=SUM(D2:D" + index + ")", HOURS_LEFT, "=\"" + hoursToWork + "\"- SUM(D2:D" + index + ")"])
  mySheet.getRange("B" + (index + 2)).setNumberFormat("[hh]:mm")
  mySheet.getRange("D" + (index + 2)).setNumberFormat("[hh]:mm")
  mySheet.getRange("F" + (index + 2)).setNumberFormat("[hh]:mm")
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
