/*

Handles all date generating, including various holidays, exam days, and PD days.
Ensures that there are no typos or schedule conflicts (in case of hand-made edits).
Completes User stories in Epic 1.
While this code is lengthy, much of it is repetition.

*/
let sheet = mainSheet;
let test = sheet.getRange("J1");

// This function calculates the date of Good Friday on a given year.
// Algorithm source: https://en.wikipedia.org/wiki/Date_of_Easter#Anonymous_Gregorian_algorithm
function calcEaster(y){
  let a = y%19, b = Math.floor(y/100), c = y%100;
  let d = Math.floor(b/4), e = b%4, f = Math.floor((b+8)/25), i = Math.floor(c/4), k = c%4;
  let g = Math.floor((b-f+1)/3);
  let h = (19*a + b - d - g + 15)%30;
  let l = (32 + 2*e + 2*i - h - k)%7;
  let m = Math.floor((a+11*h+22*l)/451);
  let n = Math.floor((h+l-7*m+114)/31);
  let o = (h+l-7*m+114)%31;
  return [n, o-1]; // n is month, o is day
}

// This function validates everything involving supervision assignments
// Depending on the level of warning required, it colours the background of the errored cells to alert the administrator
function validate() {
  let cols = "DEFGH"; // index of columns with content
  let ulim = 4, llim = 23; // index of rows with content, ulim is upper limit, llim is lower limit
  let num_tchrs = 120;
  let ans = 0; // count of number of invalid entries
  let namesArr = [];
  let nameSheet = prepSheet.getRange("A2:A"+(num_tchrs+1)).getValues();
  
  for(let i = 0; i < num_tchrs; i++){
    namesArr.push(nameSheet[i].toString());
  }
  const names = new Set(namesArr); // Gets the list of teacher names, and puts them into a set
  let testBlock = sheet.getRange(cols[0]+ulim+":"+cols[cols.length-1]+llim); // entries to validate
  let testBlockValues = testBlock.getValues();

  let colourArr = []; // maintains the list of colours for all the cells, in order to avoid accessing the sheet too much and slowing the program down
  for(let i = 0; i <= llim-ulim; i++){
    let colourRow = [];
    for(let j = 0; j < cols.length; j++){
      colourRow.push("white");
    }
    colourArr.push(colourRow);
  }

  for(let i = 0; i < cols.length; i++){
    for(let j = 0; j <= llim-ulim; j++){
      if(!names.has(testBlockValues[j][i])){ // if the name is not on the list, alert
        ans++;
        colourArr[j][i] = "red";
      }
      else {
        colourArr[j][i] = "white"; // otherwise, clear the alert
      }
    }
  }
  
  let ans2 = 0; // count of number of conflicting entries
  let preps = prepSheet.getRange("B2:B"+(num_tchrs+1)).getValues();
  for(let i = 0; i < preps.length; i++){
    preps[i] = preps[i].toString();
  }
  let prepArr = [];
  for(let i = 0; i < num_tchrs; i++){
    prepArr.push(parseInt(preps[i]));
  }
  let indMap = new Map(); // this maps the names to their indices (and thereby prep periods), allows for efficient searching
  for(let i = 0; i < namesArr.length; i++){
    indMap.set(namesArr[i], i);
  }

  for(let i = 0; i < cols.length; i++){
    for(let j = 0; j <= llim-ulim; j++){
      if(colourArr[j][i] != 'white') continue; // if the name is invalid, don't override the colour!
      if(j%2 == 0){ // Day 1
        if(j <= 3){ // meaning this is a period 1 library supervision
          if(prepArr[indMap.get(testBlockValues[j][i])] != 1){
            ans2++;
            colourArr[j][i] = "orange"; // if the prep is not in period 1, alert
          }
          else {
            colourArr[j][i] = "white"; // otherwise clear the alert
          }
        }
        else { // meaning this is a lunch supervision
          if(prepArr[indMap.get(testBlockValues[j][i])] != 2 && prepArr[indMap.get(testBlockValues[j][i])] != 3){
            ans2++;
            colourArr[j][i] = "orange";
          }
          else {
            colourArr[j][i] = "white";
          }
        }
      }
      else { // Day 2
        if(j <= 3){
          if(prepArr[indMap.get(testBlockValues[j][i])] != 2){ // flipped, due to day 2
            ans2++;
            colourArr[j][i] = "orange";
          }
          else {
            colourArr[j][i] = "white";
          }
        }
        else {
          if(prepArr[indMap.get(testBlockValues[j][i])] != 1 && prepArr[indMap.get(testBlockValues[j][i])] != 4){
            ans2++;
            colourArr[j][i] = "orange";
          }
          else {
            colourArr[j][i] = "white";
          }
        }
      }
    }
  }
  let ans3 = 0; // counts the amount of unpreferred assignments
  let unpref = prepSheet.getRange("E2:E"+(num_tchrs+1)).getValues(); // column with unpreferred duties
  for(let i = 0; i < unpref.length; i++){
    unpref[i] = unpref[i].toString();
  }
  let unprefArr = [];
  for(let i = 0; i < num_tchrs; i++){
    unprefArr.push(unpref[i]);
  }
  let updays = prepSheet.getRange("F2:F"+(num_tchrs+1)).getValues(); // column with unpreferred days
  for(let i = 0; i < updays.length; i++){
    updays[i] = parseInt(updays[i].toString());
  }
  let updaysArr = [];
  for(let i = 0; i < num_tchrs; i++){
    updaysArr.push(updays[i]);
  }

  for(let i = 0; i < cols.length; i++){
    for(let j = 0; j <= llim-ulim; j++){
      if(colourArr[j][i] != 'white') continue; // #ffffff: white, so this ensures the cell is not alerted, and does not override another alert that is more important
      if(unprefArr[indMap.get(testBlockValues[j][i])] === mainSheetValues[Math.floor(j/2)*2+3][0]){ // the math equation at the end finds the row in the heading corresponding to the duty name
        ans3++;
        colourArr[j][i] = "yellow"; // if the duty is unpreferred, alert
      }
      else {
        colourArr[j][i] = "white"; // else clear
      }
      if(colourArr[j][i] != 'white') continue;
      if(updaysArr[indMap.get(testBlockValues[j][i])] === i+1){ // i+1 is the day (1 for Monday, etc.)
        ans3++;
        colourArr[j][i] = "yellow"; // if the day is unpreferred, alert
      }
      else {
        colourArr[j][i] = "white"; // else clear
      }
    }
  }

  sheet.getRange(cols[0]+ulim+":"+cols[cols.length-1]+llim).setBackgrounds(colourArr);
  SpreadsheetApp.getUi().alert(ans+" invalid name(s) found (highlighted in red). "+ans2+" schedule conflicts found (highlighted in orange). "+ans3+" unfavourable schedules found (highlighted in yellow)."); // print to the administrator
}

// This function generates the heading for the calendar
function genCal() {
  let ui = SpreadsheetApp.getUi();
  let month = ui.prompt("Enter the month: ").getResponseText(); // user input for month and year to generate calendar
  let year = ui.prompt("Enter the year: ").getResponseText();
  
  while(true){ // Error trapping
    if(!isNaN(parseInt(month)) && !isNaN(parseInt(year)) && parseInt(year) > 2020 && parseInt(month) < 13){
      break;
    }
    if(isNaN(parseInt(month))){
      month = ui.prompt("The month must be a whole number between 1 and 12. Try again: ").getResponseText();
      continue;
    }
    if(isNaN(parseInt(year))){
      year = ui.prompt("The year must be a whole number. Try again: ").getResponseText();
      continue;
    }
    if(year <= 2020){
      year = ui.prompt("Generating a schedule that far back will not help. Try again: ").getResponseText();
      continue;
    }
    if(month > 12){
      month = ui.prompt("There are only 12 months in a year. Try again: ").getResponseText();
    }
  }

  sheet.getRange("B1").setValue(year);
  let seed = 0; // number of days since 1/1/1973, a Monday, therefore this number mod 7 gives the day of the week
  for (let i = 1973; i < year; i++) {
    if (i % 4 == 0 && i % 100 != 0) { // leap years
      seed += 366;
    }
    else if (i % 400 == 0) {
      seed += 366;
    }
    else {
      seed += 365;
    }
  }

  if (month >= 1) seed += 0;
  if (month >= 2) seed += 31;
  if (month >= 3) {
    if (year % 4 == 0 && year % 100 != 0) seed += 29;
    else if (year % 400 == 0) seed += 29;
    else seed += 28;
  }
  if (month >= 4) seed += 31;
  if (month >= 5) seed += 30;
  if (month >= 6) seed += 31;
  if (month >= 7) seed += 30;
  if (month >= 8) seed += 31;
  if (month >= 9) seed += 31;
  if (month >= 10) seed += 30;
  if (month >= 11) seed += 31;
  if (month >= 12) seed += 30;

  let monthEl = sheet.getRange("A1"); // set the title
  if (month == 1) monthEl.setValue("January");
  if (month == 2) monthEl.setValue("February");
  if (month == 3) monthEl.setValue("March");
  if (month == 4) monthEl.setValue("April");
  if (month == 5) monthEl.setValue("May");
  if (month == 6) monthEl.setValue("June");
  if (month == 7) monthEl.setValue("July");
  if (month == 8) monthEl.setValue("August");
  if (month == 9) monthEl.setValue("September");
  if (month == 10) monthEl.setValue("October");
  if (month == 11) monthEl.setValue("November");
  if (month == 12) monthEl.setValue("December");
  // Finding the days of week corresponding to every day in the month
  let mon = [], tues = [], wed = [], thurs = [], fri = [];
  let base = seed;
  if (month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12) { // 31 days
    for (let i = 1; i <= 31; i++) {
      seed++;
      if (seed % 7 == 0 || seed % 7 == 6) continue; // Saturday or Sunday, skip
      if (seed % 7 == 1) mon.push(i);
      if (seed % 7 == 2) tues.push(i);
      if (seed % 7 == 3) wed.push(i);
      if (seed % 7 == 4) thurs.push(i);
      if (seed % 7 == 5) fri.push(i);
    }
  }
  if (month == 4 || month == 6 || month == 9 || month == 11) { // 30 days
    for (let i = 1; i <= 30; i++) {
      seed++;
      if (seed % 7 == 0 || seed % 7 == 6) continue;
      if (seed % 7 == 1) mon.push(i);
      if (seed % 7 == 2) tues.push(i);
      if (seed % 7 == 3) wed.push(i);
      if (seed % 7 == 4) thurs.push(i);
      if (seed % 7 == 5) fri.push(i);
    }
  }
  if (month == 2) {
    if (year % 4 == 0 && year % 100 != 0) { // 28/29 days
      for (let i = 1; i <= 29; i++) {
        seed++;
        if (seed % 7 == 0 || seed % 7 == 6) continue;
        if (seed % 7 == 1) mon.push(i);
        if (seed % 7 == 2) tues.push(i);
        if (seed % 7 == 3) wed.push(i);
        if (seed % 7 == 4) thurs.push(i);
        if (seed % 7 == 5) fri.push(i);
      }
    }
    else if (year % 400 == 0) {
      for (let i = 1; i <= 29; i++) {
        seed++;
        if (seed % 7 == 0 || seed % 7 == 6) continue;
        if (seed % 7 == 1) mon.push(i);
        if (seed % 7 == 2) tues.push(i);
        if (seed % 7 == 3) wed.push(i);
        if (seed % 7 == 4) thurs.push(i);
        if (seed % 7 == 5) fri.push(i);
      }
    }
    else {
      for (let i = 1; i <= 28; i++) {
        seed++;
        if (seed % 7 == 0 || seed % 7 == 6) continue;
        if (seed % 7 == 1) mon.push(i);
        if (seed % 7 == 2) tues.push(i);
        if (seed % 7 == 3) wed.push(i);
        if (seed % 7 == 4) thurs.push(i);
        if (seed % 7 == 5) fri.push(i);
      }
    }
  }
  
  // Checking for the various holidays, PA days, and Exam days
  if(month == 9){
    // labour day is 1st Monday, anything before is summer break/PD day, the next Tuesday is 1st day of school
    if(tues[0] < mon[0]) tues[0] = "SB"+tues[0];
    if(wed[0] < mon[0]) wed[0] = "SB"+wed[0];
    if(thurs[0] < mon[0]) thurs[0] = "SB"+thurs[0];
    if(fri[0] < mon[0]) fri[0] = "PA"+fri[0];
    mon[0] = "LD"+mon[0];
  }
  if(month == 10){
    // second Monday is thanksgiving, the Friday before that is PD day
    if(fri[0] == mon[1]-3) fri[0] = "PA"+fri[0];
    else fri[1] = "PA"+fri[1];
    mon[1] = "FD"+mon[1];
  }
  if(month == 11){
    // the Friday of the last full week is PD day
    if(fri[fri.length-1] != 30) fri[fri.length-1] = "PA"+fri[fri.length-1];
    else fri[fri.length-2] = "PA"+fri[fri.length-2];
  }
  if(month == 12){
    /*
    Rules for holidays:
    - if Christmas is on a weekday, then all that week, as well as the week after, are off
    - else, the 2 weeks after Christmas are off
    - this may spillover into the next year
    */
    seed -= 6; // the seed of the day of Christmas
    mon[mon.length-1] = "HB"+mon[mon.length-1]; // making use of a fact: no matter what, the last week of the year is always off
    tues[tues.length-1] = "HB"+tues[tues.length-1];
    wed[wed.length-1] = "HB"+wed[wed.length-1];
    thurs[thurs.length-1] = "HB"+thurs[thurs.length-1];
    fri[fri.length-1] = "HB"+fri[fri.length-1];
    if(seed%7 == 2){ // accounting for the possible days of Christmas
      mon[mon.length-2] = "HB"+mon[mon.length-2];
    }
    if(seed%7 == 3){
      mon[mon.length-2] = "HB"+mon[mon.length-2];
      tues[tues.length-2] = "HB"+tues[tues.length-2];
    }
    if(seed%7 == 4){
      mon[mon.length-2] = "HB"+mon[mon.length-2];
      tues[tues.length-2] = "HB"+tues[tues.length-2];
      wed[wed.length-2] = "HB"+wed[wed.length-2];
    }
    if(seed%7 == 5){
      mon[mon.length-2] = "HB"+mon[mon.length-2];
      tues[tues.length-2] = "HB"+tues[tues.length-2];
      wed[wed.length-2] = "HB"+wed[wed.length-2];
      thurs[thurs.length-2] = "HB"+thurs[thurs.length-2];
    }
  }
  if(month == 1){
    /*
    Jan exams:
    - if Jan 31 is a weekday, then that day is a PD day, and the 5 school days preceeding it are exam days
    - else, the last Friday before Jan 31 is a PD day, and the 5 school days preceeding that are exam days
    */
    seed -= 37; // again, day of Christmas
    if(seed%7 == 5){ // accounting for the spillover from December
      fri[0] = "HB"+fri[0];
    }
    if(seed%7 == 4){
      thurs[0] = "HB"+thurs[0];
      fri[0] = "HB"+fri[0];
    }
    if(seed%7 == 3){
      wed[0] = "HB"+wed[0];
      thurs[0] = "HB"+thurs[0];
      fri[0] = "HB"+fri[0];
    }
    if(seed%7 == 2){
      tues[0] = "HB"+tues[0];
      wed[0] = "HB"+wed[0];
      thurs[0] = "HB"+thurs[0];
      fri[0] = "HB"+fri[0];
    }
    if(seed%7 == 0 || seed%7 == 1 || seed%7 == 6){
      mon[0] = "HB"+mon[0];
      tues[0] = "HB"+tues[0];
      wed[0] = "HB"+wed[0];
      thurs[0] = "HB"+thurs[0];
      fri[0] = "HB"+fri[0];
    }
    seed += 37; // seed of Jan 31
    while(seed%7 == 6 || seed%7 == 0){
      seed--;
    }
    if(seed%7 == 1){
      mon[mon.length-1] = "PA"+mon[mon.length-1];
      fri[fri.length-1] = "ED"+fri[fri.length-1];
      thurs[thurs.length-1] = "ED"+thurs[thurs.length-1];
      wed[wed.length-1] = "ED"+wed[wed.length-1];
      tues[tues.length-1] = "ED"+tues[tues.length-1];
      mon[mon.length-2] = "ED"+mon[mon.length-2];
    }
    if(seed%7 == 2){
      mon[mon.length-1] = "ED"+mon[mon.length-1];
      fri[fri.length-1] = "ED"+fri[fri.length-1];
      thurs[thurs.length-1] = "ED"+thurs[thurs.length-1];
      wed[wed.length-1] = "ED"+wed[wed.length-1];
      tues[tues.length-1] = "PA"+tues[tues.length-1];
      tues[tues.length-2] = "ED"+tues[tues.length-2];
    }
    if(seed%7 == 3){
      mon[mon.length-1] = "ED"+mon[mon.length-1];
      fri[fri.length-1] = "ED"+fri[fri.length-1];
      thurs[thurs.length-1] = "ED"+thurs[thurs.length-1];
      wed[wed.length-1] = "PA"+wed[wed.length-1];
      tues[tues.length-1] = "ED"+tues[tues.length-1];
      wed[wed.length-2] = "ED"+wed[wed.length-2];
    }
    if(seed%7 == 4){
      mon[mon.length-1] = "ED"+mon[mon.length-1];
      fri[fri.length-1] = "ED"+fri[fri.length-1];
      thurs[thurs.length-1] = "PA"+thurs[thurs.length-1];
      wed[wed.length-1] = "ED"+wed[wed.length-1];
      tues[tues.length-1] = "ED"+tues[tues.length-1];
      thurs[thurs.length-2] = "ED"+thurs[thurs.length-2];
    }
    if(seed%7 == 5){
      mon[mon.length-1] = "ED"+mon[mon.length-1];
      fri[fri.length-1] = "PA"+fri[fri.length-1];
      thurs[thurs.length-1] = "ED"+thurs[thurs.length-1];
      wed[wed.length-1] = "ED"+wed[wed.length-1];
      tues[tues.length-1] = "ED"+tues[tues.length-1];
      fri[fri.length-2] = "ED"+fri[fri.length-2];
    }
  }

  if(month == 2){
    // 2nd last Monday is family day, the Friday before it is PD day,
    if(fri[fri.length-2] == mon[mon.length-2]-3) fri[fri.length-2] = "PA"+fri[fri.length-2]; 
    else fri[fri.length-3] = "PA"+fri[fri.length-3];
    mon[mon.length-2] = "FD"+mon[mon.length-2];
  }
  if(month == 3){
    // March break: 2nd full week of March
    let mb = mon[1];
    mon[1] = "MB"+mon[1]
    if(tues[1] == mb+1) tues[1] = "MB"+tues[1];
    else tues[2] = "MB"+tues[2];
    if(wed[1] == mb+2) wed[1] = "MB"+wed[1];
    else wed[2] = "MB"+wed[2];
    if(thurs[1] == mb+3) thurs[1] = "MB"+thurs[1];
    else thurs[2] = "MB"+thurs[2];
    if(fri[1] == mb+4) fri[1] = "MB"+fri[1];
    else fri[2] = "MB"+fri[2];
    seed++; // seed of April 1
    let m = calcEaster(year)[0], d = calcEaster(year)[1]; // Easter can be in March, so it much be accounted for
    Logger.log(m);
    if(m == 3){
      if(fri[fri.length-1] == d) fri[fri.length-1] = "GF"+fri[fri.length-1];
      else fri[fri.length-2] = "GF"+fri[fri.length-2];
    }
    if(m == 3 && d <= 28){ // if Good Friday is before March 28, Easter Monday is also in March
      if(mon[mon.length-1] == d+3) mon[mon.length-1] = "EM"+mon[mon.length-1];
      else mon[mon.length-2] = "EM"+mon[mon.length-2];
    }
  }
  if(month == 4){
    // Again, accounting for Easter
    let m = calcEaster(year)[0], d = calcEaster(year)[1];
    if(m == 4){
      if(fri[0] == d) fri[0] = "GF"+fri[0];
      else if(fri[1] == d) fri[1] = "GF"+fri[1];
      else fri[2] = "GF"+fri[2];
      if(mon[0] == d+3) mon[0] = "EM"+mon[0]; // Easter can be very late in April...
      else if(mon[1] == d+3) mon[1] = "EM"+mon[1];
      else mon[2] = "EM"+mon[2];
    }
    else {
      if(d > 28){
        mon[0] = 'EM'+mon[0];
      }
    }
  }
  if(month == 5){
    // 2nd last Monday is Victoria Day
    mon[mon.length-2] = "VD"+mon[mon.length-2];
  }
  if(month == 6){
    /*
    June exams:
    - the Wednesday of the last full week of June is the first day of Summer (for students)
    - this Wednesday and the Thursday after it are PD days
    - the 5 days preceeding the Wednesday are exam days
    */
    let lwed = wed[wed.length-1];
    if(lwed+2 <= 30){
      wed[wed.length-2] = "ED"+wed[wed.length-2];
      thurs[thurs.length-2] = "ED"+thurs[thurs.length-2];
      fri[fri.length-2] = "ED"+fri[fri.length-2];
      if(tues[tues.length-1] < lwed){
        tues[tues.length-1] = "ED"+tues[tues.length-1];
      }
      else {
        tues[tues.length-2] = "ED"+tues[tues.length-2];
        tues[tues.length-1] = "SB"+tues[tues.length-1];
      }
      if(mon[mon.length-1] < lwed){
        mon[mon.length-1] = "ED"+mon[mon.length-1];
      }
      else {
        mon[mon.length-2] = "ED"+mon[mon.length-2];
        mon[mon.length-1] = "SB"+mon[mon.length-1];
      }
      wed[wed.length-1] = "PA"+wed[wed.length-1];
      thurs[thurs.length-1] = "PA"+thurs[thurs.length-1];
      fri[fri.length-1] = "SB"+fri[fri.length-1];
    }
    else {   
      if(thurs[thurs.length-2] == wed[wed.length-2]+1) thurs[thurs.length-2] = "PA"+thurs[thurs.length-2];
      else thurs[thurs.length-1] = "PA"+thurs[thurs.length-1];
      if(fri[fri.length-2] == wed[wed.length-2]+2) fri[fri.length-2] = "SB"+fri[fri.length-2];
      else fri[fri.length-1] = "SB"+fri[fri.length-1];
      if(thurs[thurs.length-2] == wed[wed.length-2]+1) thurs[thurs.length-3] = "ED"+thurs[thurs.length-3];
      else thurs[thurs.length-2] = "ED"+thurs[thurs.length-2];
      if(fri[fri.length-2] == wed[wed.length-2]+2) fri[fri.length-3] = "ED"+fri[fri.length-3];
      else fri[fri.length-2] = "ED"+fri[fri.length-2];
      lwed = wed[wed.length-2];
      wed[wed.length-2] = "PA"+wed[wed.length-2];
      wed[wed.length-3] = "ED"+wed[wed.length-3];
      if(tues[tues.length-1] < lwed){
        tues[tues.length-1] = "ED"+tues[tues.length-1];
      }
      else {
        tues[tues.length-2] = "ED"+tues[tues.length-2];
        tues[tues.length-1] = "SB"+tues[tues.length-1];
      }
      if(mon[mon.length-1] < lwed){
        mon[mon.length-1] = "ED"+mon[mon.length-1];
      }
      else {
        mon[mon.length-2] = "ED"+mon[mon.length-2];
        mon[mon.length-1] = "SB"+mon[mon.length-1];
      }
    }
  }
  let ans1 = "DAY 1: " // print the headings, one by one (Monday)
  let ans2 = "DAY 2: "
  for(let i = 0; i < mon.length; i++){
    if(isNaN(parseInt(mon[i]))){
      if(parseInt(mon[i][mon[i].length-1])%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += mon[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += mon[i];
      }
    }
    else {
      if(mon[i]%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += mon[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += mon[i];
      }
    }
  }
  sheet.getRange("D1").setValue(ans1);
  sheet.getRange("D2").setValue(ans2);

  ans1 = "DAY 1: " // Tuesday
  ans2 = "DAY 2: "
  for(let i = 0; i < tues.length; i++){
    if(isNaN(parseInt(tues[i]))){
      if(parseInt(tues[i][tues[i].length-1])%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += tues[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += tues[i];
      }
    }
    else {
      if(tues[i]%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += tues[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += tues[i];
      }
    }
  }
  sheet.getRange("E1").setValue(ans1);
  sheet.getRange("E2").setValue(ans2);

  ans1 = "DAY 1: " // Wednesday
  ans2 = "DAY 2: "
  for(let i = 0; i < wed.length; i++){
    if(isNaN(parseInt(wed[i]))){
      if(parseInt(wed[i][wed[i].length-1])%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += wed[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += wed[i];
      }
    }
    else {
      if(wed[i]%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += wed[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += wed[i];
      }
    }
  }
  sheet.getRange("F1").setValue(ans1);
  sheet.getRange("F2").setValue(ans2);

  ans1 = "DAY 1: " // Thursday
  ans2 = "DAY 2: "
  for(let i = 0; i < thurs.length; i++){
    if(isNaN(parseInt(thurs[i]))){
      if(parseInt(thurs[i][thurs[i].length-1])%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += thurs[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += thurs[i];
      }
    }
    else {
      if(thurs[i]%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += thurs[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += thurs[i];
      }
    }
  }
  sheet.getRange("G1").setValue(ans1);
  sheet.getRange("G2").setValue(ans2);

  ans1 = "DAY 1: " // Friday
  ans2 = "DAY 2: "
  for(let i = 0; i < fri.length; i++){
    if(isNaN(parseInt(fri[i]))){
      if(parseInt(fri[i][fri[i].length-1])%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += fri[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += fri[i];
      }
    }
    else {
      if(fri[i]%2 == 0){
        if(ans2 != "DAY 2: ") ans2 += ", ";
        ans2 += fri[i];
      }
      else {
        if(ans1 != "DAY 1: ") ans1 += ", ";
        ans1 += fri[i];
      }
    }
  }
  sheet.getRange("H1").setValue(ans1);
  sheet.getRange("H2").setValue(ans2);
}
