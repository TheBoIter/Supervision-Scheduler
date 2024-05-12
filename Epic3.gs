/*
  Description:
  These contributions focus mainly on accessibility and formatting, namely creating a new sheet which is organized alphabetically by teachers and allowing users to import duties into their personal Google Calendar
*/

let ss = SpreadsheetApp.getActiveSpreadsheet();

/*
  This function allows administrators to import their duties into Google Calendar

  Fulfills Epic 3 User Story 3

  Note: Teachers cannot access the custom 'Admin' menu to import their duties into their Google Calendar due to being a viewer of the spreadsheet. The thought process is that only a few teachers will use this feature. Thus, those teachers that want this feature could make a copy of the spreadsheet with its code and access the menu from their copy.
*/
function importCal() {
  let ui = SpreadsheetApp.getUi();

  // Get an array of all the teachers
  let unfilteredTeachers = sheet.getRange('D4:H23').getValues();
  let teachers = [];

  for (i = 0; i < unfilteredTeachers.length; i++) {
    for (j = 0; j < unfilteredTeachers[i].length; j++) {
      if (unfilteredTeachers[i][j] != "") {
        teachers.push(unfilteredTeachers[i][j]); // Add teachers to a 1D array
      }
    }
  }

  // Get a valid input from the user
  let teacher = ui.prompt('Enter Your Last Name:', 'Input is case-sensitive', ui.ButtonSet.OK).getResponseText();

  while (!teachers.includes(teacher)) {
    ui.alert("Invalid Input!", "Enter a valid last name that is on the calendar.", ui.ButtonSet.OK);
    teacher = ui.prompt('Enter Your Last Name:', 'Input is case-sensitive', ui.ButtonSet.OK).getResponseText();
  }

  ui.alert("Notice", "Importing your duties into your personal calendar should not take longer than 20 seconds. Click 'OK' to resume the import.", ui.ButtonSet.OK);

  // Duty variables
  let teacherDuties = findTeacher(teacher); // Get the duties for the inputted teacher
  let dutyRow;
  let dutyColumn;
  let duty;
  let column = 1;

  // Time variables
  let time;
  let startTime;
  let endTime;

  // Variables to help get the date
  let allDays;
  let temp;
  let char;
  let validDay;

  // Date variables
  let month = sheet.getRange("A1").getValues();
  let year = sheet.getRange("B1").getValues();
  let days = [];
  let start;
  let end;

  for (i = 0; i < teacherDuties.length; i++) {
    dutyRow = teacherDuties[i].getRow();
    dutyColumn = teacherDuties[i].getColumn();

    // Get duties + times from the calendar
    if (dutyRow % 2 == 0) {
      duty = sheet.getRange(dutyRow, column).getValues() + "";
      allDays = sheet.getRange(1, dutyColumn).getValues() + "";
      time = sheet.getRange(dutyRow, column + 1).getValues() + "";
    }
    else {
      duty = sheet.getRange(dutyRow - 1, column).getValues() + "";
      allDays = sheet.getRange(2, dutyColumn).getValues() + "";
      time = sheet.getRange(dutyRow - 1, column + 1).getValues() + "";
    }

    // Get starting time
    for (j = 0; j < time.length; j++) {
      if (time.charAt(j) == ' ' || time.charAt(j) == '-') {
        startTime = time.substring(0, j);
        break;
      }
    }

    // Get ending time
    for (k = time.length - 1; k >= 0; k--) {
      if (time.charAt(k) == ' ' || time.charAt(k) == '-') {
        endTime = time.substring(k + 1);
        break;
      }
    }

    validDay = false;
    temp = -1;

    // Get the date of the duty
    for (l = 0; l < allDays.length; l++) {
      char = allDays.charAt(l);

      // Check if loop is on the last character
      if (l == allDays.length - 1 && validDay) {
        days.push(allDays.substring(temp, l + 1));
      }
      // Check when days separated by commas
      else if (char == ',') {
        // Add day to the days array
        if (validDay) {
          days.push(allDays.substring(temp, l));
        }

        validDay = true;
        temp = -1;
      }
      // Start checking for days after the colon
      else if (char == ":") {
        validDay = true;
      }
      // If the next day in the string is a holiday, do not bother getting the day
      else if (isNaN(char) && validDay) {
        validDay = false;
      }
      // Get the start (first digit) of the day 
      else if (!isNaN(char) && char != " " && validDay && temp == -1) {
        temp = l;
      }
    }

    // Get the start and end dates for the duties
    for (m = 0; m < days.length; m++) {
      start = days[m] + " " + month + " " + year + " " + startTime;
      end = days[m] + " " + month + " " + year + " " + endTime;

      createEvent(duty, start, end); // Create the calendar event
    }

    days = [];
  }

  ui.alert("Done!", "You can find all your scheduled duties in your personal Google Calendar.", ui.ButtonSet.OK);
}

// This function creates a calendar event in the user's default calendar
function createEvent(title, startTime, endTime) {
  let calendar = CalendarApp.getDefaultCalendar();
  let startDate = new Date(startTime);
  let endDate = new Date(endTime);

  // Create and format event
  let event = calendar.createEvent(title, startDate, endDate);
  event.setDescription("You are scheduled for the " + title + " duty")
  event.setLocation("Earl Of March Secondary School");
  event.setColor(CalendarApp.EventColor.BLUE);
  event.addPopupReminder(15);
}

/*  
  This function takes all of the teachers in the calendar and organizes them alphabetically in a new sheet instead of by each duty. This allows teachers to easily view what days they have duties and what days their colleagues have duties

  Fulfills Epic 2 User Story 4 and Epic 3 User Story 1
*/
function formatTeachers() {
  SpreadsheetApp.getUi().alert("Notice", "Generating the sheet of teachers should not take longer than 30 seconds. Click 'OK' to resume generation.", SpreadsheetApp.getUi().ButtonSet.OK);

  // Delete any old sheets of data if needed
  let sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().includes("Teachers - ")) {
      ss.deleteSheet(sheets[i]);
    }
  }

  // Create a new sheet
  let month = sheet.getRange("A1").getValues();
  let year = sheet.getRange("B1").getValues();
  let name = "Teachers - " + month + " " + year;
  let teacherSheet = ss.insertSheet();
  teacherSheet.setName(name);

  // Protect the sheet so only administrators can edit it
  protection = teacherSheet.protect();
  protection.addEditor(Session.getActiveUser()); 
  protection.addEditors(["vton1@ocdsb.ca", "pxu3@ocdsb.ca", "ltu1@ocdsb.ca"]);

  // Format sheet
  let teacherRange = teacherSheet.getRange("A:Z");
  teacherRange.setFontWeight('normal');
  teacherRange.setFontFamily('Calibri');
  teacherRange.setHorizontalAlignment('center');
  teacherRange.setVerticalAlignment('middle');
  teacherRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  teacherRange.setFontSize(14);

  // Display month and year from the calendar
  teacherSheet.getRange("A1").setValue(month + " " + year);
  teacherSheet.getRange("A1").setFontSize(20);
  teacherSheet.getRange("A1").setFontWeight('bold');
  teacherSheet.getRange("A1").setHorizontalAlignment('left');
  teacherSheet.getRange("A1").setVerticalAlignment('bottom');
  teacherSheet.setColumnWidth(1, 220);

  // Display days 1 and 2 from the duties schedule
  let days = sheet.getRange("D1:H2").getValues();
  let row = 1;
  let column = 2;

  for (i = 0; i < days.length; i++) {
    for (j = 0; j < days[i].length; j++) {
      // Format each cell
      teacherSheet.setColumnWidth(column, 160);
      teacherSheet.setColumnWidth(column + 1, 40);
      teacherSheet.getRange(row, column, 1, 2).mergeAcross();
      teacherSheet.getRange(row, column).setValue(days[i][j]);
      teacherSheet.getRange(row, column).setFontWeight('bold');
      teacherSheet.getRange(row, column).setVerticalAlignment('bottom');

      // Merge adjacent cells one row below that will soon have data added
      if (i == days.length - 1) {
        teacherSheet.getRange(row + 1, column, 1, 2).mergeAcross();
      }
      
      column += 2;
    }

    row++;
    column = 2;
  }

  const firstRow = row + 2; // First row where new, different data is added each time sheet is generated
  const firstColumn = 1; // First column of data
  const lastColumn = days[0].length * 2 + 1; // Last column of data (excluding legend): 5 days multiplied by 2 for the merged cells plus the first column

  teacherSheet.getRange(1, 1, days.length, lastColumn).setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID); // Add borders

  // Format the header for the data table
  teacherSheet.getRange(row, firstColumn, 2, 1).mergeVertically();

  let tableHeader = teacherSheet.getRange(row, firstColumn, 2, lastColumn);
  tableHeader.setBackground("#6d9eeb");
  tableHeader.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
  
  teacherSheet.getRange(row, firstColumn).setValue("Teacher (alphabetical)");
  teacherSheet.getRange(row, firstColumn).setFontWeight('bold');
  column = 4;

  for (i = column; i < column + days[0].length; i++) {
    // The column in the teacherSheet is the same as the column in the calendar sheet multiplied by 2 for the merged cells and subtracted by 6 as the days start on column 2 for teacherSheet and column 4 for calendar sheet
    teacherSheet.getRange(row, i * 2 - 6).setValue(sheet.getRange(row, i).getValues());
    teacherSheet.getRange(row, i * 2 - 6).setFontWeight('bold');
    teacherSheet.getRange(row + 1, i * 2 - 6).setValue("Duty");
    teacherSheet.getRange(row + 1, i * 2 - 5).setValue("Day");
  }
  
  // Add the legend to the sheet
  row = 1;
  column += days[0].length;
  let legend = [];

  // Add the legend to a 2D array
  while (sheet.getRange(row, column).getValues() != "") {
    legend.push([sheet.getRange(row, column).getValues()]); 
    row++;
  }

  row--;
  teacherSheet.setColumnWidth(lastColumn + 1, 160);
  teacherSheet.getRange(1, lastColumn + 1, row, 1).setValues(legend); // Display values from the array in the sheet

  // Format the legend
  teacherSheet.getRange(1, lastColumn + 1).setFontWeight('bold'); // Make legend title bold
  teacherSheet.getRange(1, lastColumn + 1).setVerticalAlignment('bottom'); // Set vertical alignment of legend title to bottom

  let legendTable = teacherSheet.getRange(1, lastColumn + 1, row, 1);
  legendTable.setFontSize(12);
  legendTable.setBackground(sheet.getRange(row, column).getBackground());
  legendTable.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

  // Get teachers from duties schedule
  let unsortedTeachers = sheet.getRange('D4:H23').getValues();
  let teachers = [];
  let teacher;
  let isDuplicate;

  for (i = 0; i < unsortedTeachers.length; i++) {
    for (j = 0; j < unsortedTeachers[i].length; j++) {
      teacher = unsortedTeachers[i][j];
      isDuplicate = teachers.includes(teacher);

      // Ensure the same teacher is not added to the array twice
      if (!isDuplicate) {
        if (teacher != "") {
          teachers.push(teacher);
        }
      }
    }
  }

  // Sort teachers alphabetically
  let temp;

  for (i = 0; i < teachers.length - 1; i++) {
    for (j = i + 1; j < teachers.length; j++) {
      if (teachers[i] > teachers[j]) {
        temp = teachers[j];
        teachers[j] = teachers[i];
        teachers[i] = temp;
      }
    }
  }

  // Create a 2D array to store data
  let lastRow = teachers.length + firstRow;
  let data = [];

  for (i = 0 ; i < teachers.length; i++) {
    data.push([teachers[i]]);

    for (j = 0; j < days[0].length * 2; j++) {
      data[i].push("");
    }
  }

  // Add teachers and their duties to the data array
  column = firstColumn;
  let teacherDuties;
  let dutyRow;
  let dutyColumn;
  let duty = "";
  let day = "";

  for (i = 0; i < data.length; i++) {
    teacher = data[i][0];
    teacherDuties = findTeacher(teacher); // Call the findTeacher function

    // Add all duties for each teacher to the data array
    for (j = 0; j < teacherDuties.length; j++) {
      dutyRow = teacherDuties[j].getRow();
      dutyColumn = teacherDuties[j].getColumn();

      // If a teacher has multiple duties on the same day, ensure it is formatted properly
      if (j > 0 && data[i][dutyColumn * 2 - 7] != "") {
        duty += "\n";
        day += "\n\n";
      }

      // Get duties + times from the calendar
      if (dutyRow % 2 == 0) {
        duty += sheet.getRange(dutyRow, column).getValues() + "\n" + sheet.getRange(dutyRow, column + 1).getValues();
        day += 1;
      }
      else {
        duty += sheet.getRange(dutyRow - 1, column).getValues() + "\n" + sheet.getRange(dutyRow - 1, column + 1).getValues();
        day += 2;
      }

      // Add duties + times to the array
      data[i][dutyColumn * 2 - 7] += duty; // - 7 instead of "- 6" because array starts counting at 0
      data[i][dutyColumn * 2 - 6] += day; // - 6 instead of "- 5" because array starts counting at 0

      duty = "";
      day = "";
    }
  }

  // Display and format data
  let dataTable = teacherSheet.getRange(firstRow, firstColumn, lastRow - firstRow, lastColumn);
  dataTable.setValues(data);
  dataTable.setFontSize(12);
  teacherSheet.getRange(firstRow, firstColumn, lastRow - firstRow, 1).setFontSize(14); // Make sure teachers in the first column have a larger font size

  // Find the starting row for the teachers who do not have an assigned duty (or have their own duties)
  column = 2;
  let specialTeacherRow = 1;

  while (!sheet.getRange(specialTeacherRow, column).isPartOfMerge()) {
    specialTeacherRow++;
  }

  // Format all the rows of special teachers
  while (sheet.getRange(specialTeacherRow, column).isPartOfMerge()) {
    teacherSheet.getRange(lastRow, column, 1, lastColumn - 1).mergeAcross();
    teacherSheet.getRange(lastRow, firstColumn).setValue(sheet.getRange(specialTeacherRow, firstColumn).getValues());
    teacherSheet.getRange(lastRow, firstColumn).setHorizontalAlignment('left');
    teacherSheet.getRange(lastRow, column).setValue(sheet.getRange(specialTeacherRow, column).getValues());

    lastRow++;
    specialTeacherRow++;
  }

  // Format the whole table with all the data (including special teachers)
  let fullTable = teacherSheet.getRange(firstRow, firstColumn, lastRow - firstRow, lastColumn);
  fullTable.setBackground("#c9daf8");
  fullTable.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.getUi().alert("Done!", "The teacher schedule has finished generating.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// This function finds teachers in the calendar and returns an array of ranges
function findTeacher(teacher) {
  return sheet
    .getRange('D4:H23')
    .createTextFinder(teacher)
    .matchEntireCell(true)
    .matchCase(true)
    .matchFormulaText(false)
    .ignoreDiacritics(false)
    .findAll();
}
