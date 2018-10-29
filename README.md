# Work Time Counter

If you have ever worked part time you probably had to plan when you want to come to work. Maybe today and tomorrow 8 hours and some other day you wanted to leave earlier to do something at home. But how to keep on track if there are enough working hours already?

To solve this issue myself I've written this little script that allows you to create an event in dedicated Google Calendar (so you can hide it to not mess with other stuff) every time you work and then after only a couple of clicks make from these events a Google Spreadsheet with a summary of everyday work and how many hours it is in a whole month. If you define your monthly goal it will show you how many hours are left to be planned or how much time you've exceeded your plan.

## Tutorial
1. Go to Google Calendar settings (gear icon)
2. Add new calendar

![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/newCalendar.png)

3. Click on your new calendar and choose Integrate calendar. Here's your calendar ID - keep it somewhere

![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/integrate.png)
![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/calendarId.png)

4. Open new Google Spreadsheet and create sheet named like "XI 2018" for November 2018 or "III 2019" for March 2019 (you can create many sheets for different months)

![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/sheet.png)

5. Then open Script editor and paste the script from this repository

![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/scriptEditorOpening.png)

6. Adjust the first line of code with your own calendar Id and amount of hours you want to work this month

![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/enterCalendarId.png)

7. Save project with any name you like
8. Import macro to the spreadsheet

![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/importMacro.png)

9. Choose Tools > Macros > generateTimeReport

![](https://raw.githubusercontent.com/demba003/WorkTimeCounter/master/images/generate.png)

10. When asked for permissions, give access to everything listed
11. If you add some events to this calendar they will be shown in the summary in this sheet. That's all
