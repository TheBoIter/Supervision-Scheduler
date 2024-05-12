/*
  Date Updated: 2024/03/26

  User Stories completed in this section: all of Epic 2 except Story #4:
  To specify:
    - User Story 1: As a VP or administrator, I would like to automatically generate the schedule to save time and reduce errors.
    - User Story 2: As a VP or administrator, I would like to change the supervisors as needed. 
    - User Story 3: As an administrator, I want to make sure everyone gets an equal amount of time supervising so that I get the least amount of complaints.
    - User Story 5: As a teacher, I want to make sure lunch supervisions occur directly before or after my prep, so that I have time to eat.
    - User Story 6: As a teacher, I want to ensure that library supervisions only occur when I have prep during Block 1.
    - Additional Story: As a teacher, I want to ensure that the gym is only supervised by gym teachers.
*/

/*
* This section below gets the mainSheet (Sheet object), mainSheetValues (Sheet converted to 2D array), prepSheet (Sheet object), prepSheetValues (Sheet converted to 2D array) values.
*/
let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
let mainSheet, mainSheetValues, prepSheet, prepSheetValues;
mainSheet = -1;
prepSheet = -1;
for(let i = 0; i < sheets.length; i++){
  if(sheets[i].getName() == "Supervision Calendar"){
    mainSheet = sheets[i];
    mainSheetValues = mainSheet.getDataRange().getValues();
  }
  else if(sheets[i].getName() == "Teacher Profiles"){
    prepSheet = sheets[i];
    prepSheetValues = prepSheet.getDataRange().getValues();
  }
}
if(mainSheet == -1 || prepSheet == -1){
  throw new Error('Cannot find spreadsheets "Supervision Calendar" and "Teacher Profiles".');
}

const SCHEDULE_HORIZONTAL = 5; // Use the horizontal width of the section that will be filled with teacher names
const SCHEDULE_VERTICAL = 20; // Use the vertical width of the section that will be filled with teacher names
const TRIALS = 50000; // Number of times to generate a random schedule

// Function: Creates the custom drop-down menu named "Admin"
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin')
    .addItem('Generate Calendar Dates', 'genCal')
    .addItem("Generate Teacher Schedule", "generateSchedule")
    .addItem('Validate Teacher Schedule', 'validate')
    .addItem("Change Teacher Schedule", "changeSchedule")
    .addItem("Make Schedule by Teachers", 'formatTeachers')
    .addItem('Export to Google Calendar', 'importCal')
    // Test functionality - ONLY FOR TESTING .addItem('Run 100 Test Schedule Generations', 'testf')
    .addToUi();
}

// Function: Changes the queue and slots supervised in "Teacher Profiles" sheet according to teachers added and deleted on an edit
function onEdit(e){
  let range = e.range;

  // Warns user before editng parts of the schedule
  if(e.source.getActiveSheet().getName() === "Supervision Calendar" && range.getColumn()){
    if(range.getColumn() != range.getLastColumn() || range.getRow() != range.getLastRow()){
      SpreadsheetApp.getUi().alert("Warning! You edited multiple cells at once on the schedule! When changing the time slots, change only one at the time to ensure that times supervised is recorded correctly.");
    }
    else if(range.getColumn() >= 4+SCHEDULE_HORIZONTAL || range.getColumn() < 4 || range.getRow() >= 4+SCHEDULE_VERTICAL || range.getRow() < 4){
      SpreadsheetApp.getUi().alert("Warning! You edited the schedule format!");
    }
    else{
      let oldTeacher = e.oldValue;
      let newTeacher = e.value;

      addToPrepSheetQueue(oldTeacher, 1);
      addToPrepSheetSlotsSupervised(oldTeacher, -1);
      addToPrepSheetSlotsSupervised(newTeacher, 1);
      let newTeacherInQueue = addToPrepSheetQueue(newTeacher, -1);
      prepSheet.getDataRange().setValues(prepSheetValues);

      if(!newTeacherInQueue){
        SpreadsheetApp.getUi().alert('Warning! The new teacher does not have a teacher profile!');
      }
    }
  }
}

// Helper function for: onEdit()
// Description: Changes the "slots supervised" values in Teacher Profiles to match the change in duty assignment. Returns boolean of whether teacher is in prep sheet or not.
function addToPrepSheetSlotsSupervised(teacher, x){
  for(let i = 0; i < prepSheetValues.length; i++){
    if(prepSheetValues[i][0] == teacher){
      if(prepSheetValues[i][8] === ""){
        prepSheetValues[i][8] = x;
      }
      else{
        prepSheetValues[i][8] += x;
      }
      return true;
    }
  }
  return false;
}

// Function: Adds the functionality behind the button "Change Schedule Slot" within the drop-down menu. Returns nothing.
function changeSchedule() {
  let input = SpreadsheetApp.getUi().prompt("Enter the teacher to be switched, and the teacher to be switched with. Follow the format \"[teacher to be switched],[teacher to be switched with]\", without square brackets:").getResponseText(); // user input
  let t1 = "", t2 = "";
  let ok = 0; // if the input is valid
  for(let i = 0; i < input.length; i++){
    if(input[i] == ',' && i != input.length-1){
      ok = 1;
    }
    // splits the input into two names
    else {
      if(ok == 0) t1 += input[i];
      else t2 += input[i];
    }
  }
  t1 = t1.trim();
  t2 = t2.trim();

  if(ok == 0){
    throw new Error("Invalid format!");
  }

  let t1_duty = [], t2_duty = []; // arrays for the cells corresponding to teachers' duties
  let ok1 = 0, ok2 = 0; // check if the teachers exist
  for(let i = 3; i < 8; i++){
    for(let j = 3; j < 23; j++){
      if(mainSheetValues[j][i] == t1){
        t1_duty.push(String.fromCharCode(i+65)+(j+1));
        ok1 = 1;
      }
      if(mainSheetValues[j][i] == t2){
        t2_duty.push(String.fromCharCode(i+65)+(j+1));
        ok2 = 1;
      }
    }
  }
  if(ok1 == 0 || ok2 == 0){
    throw new Error("One of the teachers do not exist!"); 
  }

  let c1 = 0; // counts the number of entries changed
  for(duty of t1_duty){ // assigns teacher 1's duties to teacher 2
    let x = parseInt(duty.charCodeAt(0)-65);
    let y = parseInt(duty.substring(1))-1;
    mainSheetValues[y][x] = t2;
    c1++;
  }
  let c2 = 0;
  for(duty of t2_duty){ // vise versa
    let x = parseInt(duty.charCodeAt(0)-65);
    let y = parseInt(duty.substring(1))-1;
    mainSheetValues[y][x] = t1;
    c2++;
  }
  mainSheet.getDataRange().setValues(mainSheetValues); // finalizes the changes
  addToPrepSheetQueue(t1, c2-c1);
  addToPrepSheetQueue(t2, c1-c2); // fixes the queue (since the number of duties need to be changed as well)
  prepSheet.getDataRange().setValues(prepSheetValues);
}

// Function: Adds the functionality behind the button "Generate Schedule" within the drop-down menu
/* Description: The function uses the helper function generateRandomSchedule() to generate thousands of different schedules. The best schedule (according to its compatibility with each teacher's preferences) will be chosen. Returns nothing. Each random schedule already accounts for most other requirements:
*  1. Day 1's and Day 2's;
*  2. No duplicate teachers in one day; 
*  3. Duties in the first period require prep in the first period; 
*  4. Duties at lunch require prep in one of the neighboring periods; 
*  5. Duty assignments are according to contractual status (done through the use of a "queue" that 
*      tracks the number of times a teacher still needs to supervise to complete a "supervision cycle". More 
*      explanation is available in the overall helper functions description below this function).
*  6. Only gym teachers can supervise in the gym.
*/

function generateSchedule() {
  let bestSchedule = {schedule: [], score: {preferred: 0, unpreferred: 100}, conflicts: [], queues: []};
  let prefMap = makePreferenceMapByCell(); // Makes a 2D array that has each cell's preferred and unpreferred teachers
  let schedule;
  SpreadsheetApp.getUi().alert("Generating schedule should not take more than 20 seconds... Ok to start generation.");

  // For loop to generate TRIALS number of schedules and find the best schedule
  for(let _ = 0; _ < TRIALS; _++){ 
    // Each variable:
    // schedule: 2D array with only teachers in accordance with the sheet's schedule's format
    // score: Keeps track of score to find the schedule with the best score
    // conflicts: All conflicts with preference in the schedule
    schedule = {schedule: generateRandomSchedule(), score: {preferred: 0, unpreferred: 100}, conflicts: []};

    // For loop helps determine the score of a random schedule
    for(let i = 0; i < schedule.schedule.length; i++){
      for(let j = 0; j < schedule.schedule[0].length; j++){
        if(prefMap[i][j].preferred.includes(schedule.schedule[i][j])){
          schedule.score.preferred++;
        }
        if(prefMap[i][j].unpreferred.includes(schedule.schedule[i][j])){
          schedule.score.unpreferred--;
          schedule.conflicts.push([i, j]);
        }
      }
    }

    // Checks if this score beats the best score (prioritizes having fewest conflicts with unfavored duties, then most matches with favorable duties)
    if(schedule.score.unpreferred > bestSchedule.score.unpreferred || bestSchedule.score.unpreferred == 100){
      bestSchedule.score = schedule.score;
      bestSchedule.schedule = schedule.schedule;
      bestSchedule.queues = queuesByPrep;
      bestSchedule.conflicts = schedule.conflicts;
    }
    else if(schedule.score.unpreferred == bestSchedule.score.unpreferred && schedule.score.preferred > bestSchedule.score.preferred){
      bestSchedule.score = schedule.score;
      bestSchedule.schedule = schedule.schedule;
      bestSchedule.queues = queuesByPrep;
      bestSchedule.conflicts = schedule.conflicts;
    }
  }
  
  // Updates Time Supervised column
  for(let i = 0; i < SCHEDULE_VERTICAL; i++){
    for(let j = 0; j < SCHEDULE_HORIZONTAL; j++){
      for(let k = 1; k < prepSheetValues.length; k++){
        if(bestSchedule.schedule[i][j] == prepSheetValues[k][0]){
          if(prepSheetValues[k][8] != ""){
            prepSheetValues[k][8]++;
          }
          else{
            prepSheetValues[k][8] = 1;
          }
        }
      }
    }
  }
  prepSheet.getDataRange().setValues(prepSheetValues);

  // Make the change in the H Column ("Supervisions needed to reach next cycle for queue") in the "Teacher Profiles" sheet
  saveQueue(bestSchedule.queues.flat(1)); 

  // For debugging purposes of schedule generation on the sheets, uncomment these print statements
  //Logger.log(bestSchedule.schedule);
  //Logger.log(bestSchedule.score);
  //Logger.log(bestSchedule.queues);

  mainSheet.getRange("D4:"+String.fromCharCode(67+SCHEDULE_HORIZONTAL)+String(3+SCHEDULE_VERTICAL)).setValues(bestSchedule.schedule); // Copies the schedule to the sheet

  // Mark conflicts with unfavored duties
  mainSheet.getRange(4, 4, SCHEDULE_VERTICAL, SCHEDULE_HORIZONTAL).setBackgroundColor("white");
  for(let i = 0; i < bestSchedule.conflicts.length; i++){
    mainSheet.getRange(4+parseInt(bestSchedule.conflicts[i][0]), 4+parseInt(bestSchedule.conflicts[i][1])).setBackgroundColor("yellow");
  }
  SpreadsheetApp.getUi().alert(String(bestSchedule.conflicts.length) + " conflict(s) with teacher's unfavorable duty assignments found.")
}

// Test function: ONLY USE ON A COPIED SPREADSHEET
// To run, make TRIAL 100 and comment out the alerts in generateSchedule(). 
function testf(){
  SpreadsheetApp.getUi().alert("Generating 100 schedules for testing.");
  for(let i = 0; i < 10; i++){
    generateSchedule();
  }
}

// Helper function for: generateSchedule()
/* Description: 
  Summary - Generates a random schedule with the given restrictions to what slots teacher could take. However, the function does not account for favorable/unfavorable time slots. Returns a 2D array with randomized teachers which fulfilled the requirements.

  "Queue": The "queue" below tracks the number of times a teacher still needs to supervise to complete the supervision cycle (In a supervision cycle, a contract status of 100 means 3 supervisions needed initially; 66 means 2; and 33 means 1. When all supervisions are depleted, we move on to the next cycle). 
    Format: Each element in a "queue" array is [teacher name, number of supervisions until next cycle].
  
  "Preference map": The "preference map" below is used to keep track of unfavorable and favorable teachers for each given cell. The map will be used to determine out of randomly generated schedules which one is the best in terms of "favoribility score".
    Format: The "preference map" is a 2D array representing the cells to be filled in by teachers in the schedule. At each index is an object: {preferred: array of teachers, unpreferred: array of teachers}.
*/
var queuesByPrep; // Global variable that could be reached by generateSchedule() to be stored in bestQueue (only for the best queue out of the given number of trials). 
function generateRandomSchedule(){
  // Initializing schedule
  let schedule = [];
  for(let i = 0; i < SCHEDULE_VERTICAL; i++){
    schedule.push([]);
  }

  queuesByPrep = getQueuesByPrepFromSheet(); // Creates a 2D array with 4 queues by prep periods
  // Randomize to ensure randomness when picking a teacher
  for(let i = 0; i < 4; i++){
    queuesByPrep[i].sort(() => Math.random() - 0.5);
  }

  /** Fill the gym slots first with the available gym teachers **/
  // Find all gym teachers
  let gymTeachers = [];
  for(let i = 1; i < prepSheetValues.length; i++){
    if(prepSheetValues[i][6] == "Yes"){
      gymTeachers.push(prepSheetValues[i][0]);
    }
  }
  // Find all gym teacher's indices in the queuesByPrep variable
  let gymIndicesByPrep = [[], [], [], []];
  for(let i = 0; i < 4; i++){
    for(let j = 0; j < queuesByPrep[i].length; j++){
      if(gymTeachers.includes(queuesByPrep[i][j][0])){
        gymIndicesByPrep[i].push(j);
      }
    }
  }      
  // Main code to find a random gym teacher for each gym slot while fulfilling the initial requirements
  for(let i = SCHEDULE_VERTICAL-2; i < SCHEDULE_VERTICAL; i++){
    for(let j = 0; j < SCHEDULE_HORIZONTAL; j++){
      // If a given queue by prep period is empty make a completely new queue according to their contractual status (99 = 3, 66 = 2, 33 = 1)

      if(isQueueCleared(queuesByPrep[0])) {
        queuesByPrep[0] = makeQueueByPrepPeriod(1);
      }
      if(isQueueCleared(queuesByPrep[1])) {
        queuesByPrep[1] = makeQueueByPrepPeriod(2);
      }
      if(isQueueCleared(queuesByPrep[2])) {
        queuesByPrep[2] = makeQueueByPrepPeriod(3);
      }
      if(isQueueCleared(queuesByPrep[3])) {
        queuesByPrep[3] = makeQueueByPrepPeriod(4);
      }

      // Find the possible preps for a given cell
      if(i % 2 == 0){
        preps = [3, 2];
      }
      else if(i % 2 == 1){
        preps = [4, 1];
      }

      // Sum up the possible queues by prep numbers. Find random number between 0 and the sum.
      let count = 0;
      for(let ind of gymIndicesByPrep[preps[0]-1]){
        count += queuesByPrep[preps[0]-1][ind][1];
      }
      for(let ind of gymIndicesByPrep[preps[1]-1]){
        count += queuesByPrep[preps[1]-1][ind][1];
      }
      if(count != 0) count = Math.floor(Math.random()*count)+1;

      // Finds where exactly does this random number lies given the sum of the queue's numbers. The teacher at the index will be used in that cell of the schedule.
      let k = 0;
      let prep = preps[0];
      while(count > 0 && k < gymIndicesByPrep[preps[0]-1].length){
        count -= queuesByPrep[preps[0]-1][gymIndicesByPrep[preps[0]-1][k]][1];
        k++;
      }
      if(count > 0){
        k = 0;
        if(prep == preps[0]) prep = preps[1];
        while(count > 0){
          count -= queuesByPrep[preps[1]-1][gymIndicesByPrep[preps[1]-1][k]][1];
          k++;
        }
      }
      k--;
      if(k < 0) k = 0; if(k >= gymIndicesByPrep[prep-1].length) k = gymIndicesByPrep[prep-1].length-1;
      
      schedule[i].push(queuesByPrep[prep-1][gymIndicesByPrep[prep-1][k]][0]);
      queuesByPrep[prep-1][gymIndicesByPrep[prep-1][k]][1]--;
    }
  }

  /** Other slots other than gym **/
  // Main code to find a random teacher for each non-gym slot while fulfilling the initial requirements
  for(let j = 0; j < SCHEDULE_HORIZONTAL; j++){ 
    let restriction = schedule.flat(1);
    //let restriction = [schedule[SCHEDULE_VERTICAL-2][j], schedule[SCHEDULE_VERTICAL-1][j]];
    for(let i = 0; i < SCHEDULE_VERTICAL-2; i++){
      // If a given queue by prep period is empty make a completely new queue according to their contractual status (99 = 3, 66 = 2, 33 = 1)

      if(isQueueCleared(queuesByPrep[0])) {
        queuesByPrep[0] = makeQueueByPrepPeriod(1);
      }
      if(isQueueCleared(queuesByPrep[1])) {
        queuesByPrep[1] = makeQueueByPrepPeriod(2);
      }
      if(isQueueCleared(queuesByPrep[2])) {
        queuesByPrep[2] = makeQueueByPrepPeriod(3);
      }
      if(isQueueCleared(queuesByPrep[3])) {
        queuesByPrep[3] = makeQueueByPrepPeriod(4);
      }

      // Find the possible preps for a given cell
      let preps = [];
      if(0 <= i && i <= 3){
        if(i % 2 == 0){
          preps = [1];
        }
        else if(i % 2 == 1){
          preps = [2];
        }
      }
      else if(4 <= i && i <= SCHEDULE_VERTICAL-3){
        if(i % 2 == 0){
          preps = [3, 2];
        }
        else if(i % 2 == 1){
          preps = [4, 1];
        }
      }

      // If only 1 prep period is possible
      if(preps.length == 1){
        let prep = preps[0];
        let count;
        if(totalQueueScore(queuesByPrep[prep-1]) == 0) {
          count = 0;
        }
        else {
          count = Math.floor(Math.random()*totalQueueScore(queuesByPrep[prep-1]))+1; // Finds a random number within the sum of the queue
        }

        // Finds the queue index at the given random number
        let k = 0;
        while(count > 0){
          count -= queuesByPrep[prep-1][k][1];
          k++;
        }
        k--;
        if(k < 0) k = 0; if(k >= queuesByPrep[prep-1].length) k = queuesByPrep[prep-1].length-1;

        // Because the teacher at index k may already have a duty on the same day. We check for that and only go the nearest teacher in the indices that have not had a duty yet.
        let i1 = k;
        let i2 = k+1;
        let ansIndex = -1;
        while(ansIndex == -1){
          if(i1 < 0 && i2 >= queuesByPrep[prep-1].length){
            ansIndex = k;
            break;
          }
          if(i1 >= 0){
            if(!restriction.includes(queuesByPrep[prep-1][i1][0]) && queuesByPrep[prep-1][i1][1] > 0){
              ansIndex = i1;
              break;
            }
            i1--;
          }
          if(i2 < queuesByPrep[prep-1].length){
            if(!restriction.includes(queuesByPrep[prep-1][i2][0]) && queuesByPrep[prep-1][i2][1] > 0){
              ansIndex = i2;
              break;
            }
            i2++;
          }
        }

        // Setting new variables
        schedule[i].push(queuesByPrep[prep-1][ansIndex][0]); // Add the new teacher to the schedule
        restriction.push(queuesByPrep[prep-1][ansIndex][0]); // Add the new teacher to the restrictions as teachers can't have 2 supervisions on the same day
        queuesByPrep[prep-1][ansIndex][1]--; // Decrease the teacher's queue number by 1
      }

      // If 2 prep periods are possible
      else if(preps.length == 2){
        let combinedQueue = queuesByPrep[preps[0]-1].concat(queuesByPrep[preps[1]-1]); // Combine the queues by the 2 prep periods into 1 big one
        let count;
        if(totalQueueScore(combinedQueue) == 0) {
          count = 0;
        }
        else {
          count = Math.floor(Math.random()*totalQueueScore(combinedQueue))+1; // Finds a random number within the sum of the queue
        }

        // Finds the queue index at the given random number
        let k = 0;
        while(count > 0){
          count -= combinedQueue[k][1];
          k++;
        }
        k--;
        if(k < 0) k = 0; if(k >= combinedQueue.length) k = combinedQueue.length-1;

        // Because the teacher at index k may already have a duty on the same day. We check for that and only go the nearest teacher in the indices that have not had a duty yet.
        let i1 = k;
        let i2 = k+1;
        let ansIndex = -1;
        while(ansIndex == -1){
          if(i1 < 0 && i2 >= combinedQueue.length){
            ansIndex = k;
            break;
          }
          if(i1 >= 0){
            if(!restriction.includes(combinedQueue[i1][0]) && combinedQueue[i1][1] > 0){
              ansIndex = i1;
              break;
            }
            i1--;
          }
          if(i2 < combinedQueue.length){
            if(!restriction.includes(combinedQueue[i2][0]) && combinedQueue[i2][1] > 0){
              ansIndex = i2;
              break;
            }
            i2++;
          }
        }

        // Setting new variables
        schedule[i].push(combinedQueue[ansIndex][0]); // Add the new teacher to the schedule
        restriction.push(combinedQueue[ansIndex][0]); // Add the new teacher to the restrictions for today
        // Decrease the teacher's queue number by 1
        if(ansIndex < queuesByPrep[preps[0]-1].length){
          queuesByPrep[preps[0]-1][ansIndex][1]--;
        }
        else{
          queuesByPrep[preps[1]-1][ansIndex-queuesByPrep[preps[0]-1].length][1]--;
        }
      }
    }    
  }
  return schedule;
}

/*
*   Below are helper functions for: generateRandomSchedule()
*   Overall Description:
*     1. 
        The "queue" below tracks the number of times a teacher still needs to supervise to complete the supervision cycle (In a supervision cycle, a contract status of 100 means 3 supervisions needed initially; 66 means 2; and 33 means 1. When all supervisions are depleted, we move on to the next cycle). 
        Format: Each element in a "queue" array is [teacher name, number of supervisions until next cycle].
*     2.
        The "preference map" below is used to keep track of unfavorable and favorable teachers for each given cell. The map will be used to determine out of randomly generated schedules which one is the best in terms of "favoribility score".
        Format: The "preference map" is a 2D array representing the cells to be filled in by teachers in the schedule. At each index is an object: {preferred: array of teachers, unpreferred: array of teachers}.
*/

// Helper function for: changeSchedule()
// Description: Changes the queue values in the spreadsheet column "Supervisions needed to reach next cycle for queue" to match the change in duty assignment. Returns boolean of whether teacher is in prep sheet or not.
function addToPrepSheetQueue(teacher, x){
  for(let i = 0; i < prepSheetValues.length; i++){
    if(prepSheetValues[i][0] == teacher){
      if(prepSheetValues[i][7] === ""){
        prepSheetValues[i][7] = x;
      }
      else{
        prepSheetValues[i][7] += x;
      }
      return true;
    }
  }
  return false;
}

// Helper function for: generateSchedule()
// Description: Gets the "queues" by prep periods from the Prep Sheet's Column H. Returns a 3D array with each index representing its prep period number minus 1. At each index is a 2D "queue" array.
function getQueuesByPrepFromSheet(){
  let queues = [[], [], [], []];
  for(let i = 1; i < prepSheetValues.length; i++){
    let prepIndex = prepSheetValues[i][1];
    if(prepSheetValues[i][7] === "" || prepSheetValues[i][7] <= 0){
      queues[prepIndex-1].push([prepSheetValues[i][0], Math.floor(prepSheetValues[i][2]/33)]);
    }
    else{
      queues[prepIndex-1].push([prepSheetValues[i][0], prepSheetValues[i][7]]);
    }
  }
  return queues;
}

// Helper function for: generateSchedule()
// Description: Makes a completely new "queue" for teachers in a given prep period. Returns a 2D array of queue.
function makeQueueByPrepPeriod(prepPeriod){
  let queue = [];
  for(let i = 1; i < prepSheetValues.length; i++){
    if(prepSheetValues[i][1] == prepPeriod){
      if(Math.floor(prepSheetValues[i][2]/33) == 3){
        queue.push([prepSheetValues[i][0], 3]);
      }
      else if(Math.floor(prepSheetValues[i][2]/33) == 2){
        queue.push([prepSheetValues[i][0], 2]);
      }
      else if(Math.floor(prepSheetValues[i][2]/33) == 1){
        queue.push([prepSheetValues[i][0], 1]);
      }
    }
  }

  queue.sort(() => Math.random() - 0.5); // Randomize to ensure randomness when picking a teacher
  return queue;
}

// Helper function for: generateSchedule()
// Description: Tallies up the number of supervisions needed until the next cycle for each teacher in the "queue". Returns an integer representing the sum.
function totalQueueScore(queue){
  let count = 0;
  for(let i = 0; i < queue.length; i++){
    count += queue[i][1];
  }
  return count;
}

// Helper function for: generateSchedule()
// Description: Sees if the queue is completely cleared. Returns a boolean.
function isQueueCleared(queue){
  if(totalQueueScore(queue) <= 0){
    return true;
  }
  else{
    return false;
  }
}

// Helper function for: generateSchedule()
// Description: Saves the queue back to the Prep Sheet. Returns nothing.
function saveQueue(queue){
  queue = new Map(queue);
  for(let i = 1; i < prepSheetValues.length; i++){
    let val = queue.get(prepSheetValues[i][0]);
    prepSheetValues[i][7] = val;
  }
  prepSheet.getDataRange().setValues(prepSheetValues);
}

// Helper function for: generateSchedule()
// Description: Makes a 2D array for each cell. In each slot of the 2D array, favorable and unfavorable teachers are listed in an object variable. Returns a 2D "preference map". The "preference map" is a 2D array representing the cells to be filled in by teachers in the schedule. At each index is an object: {preferred: array of teachers, unpreferred: array of teachers}.
function makePreferenceMapByCell(){
  let prefMap = [];

  for(let i = 0; i < SCHEDULE_VERTICAL; i++){
    prefMap.push([]);
    for(let j = 0; j < SCHEDULE_HORIZONTAL; j++){
      prefMap[i].push({preferred: [], unpreferred: []});
    }
  }

  for(let i = 1; i < prepSheetValues.length; i++){
    //Unpreferred Days
    if(prepSheetValues[i][5] != ""){
      for(let j = 0; j < SCHEDULE_VERTICAL; j++){
        if(!prefMap[j][prepSheetValues[i][5]-1].unpreferred.includes(prepSheetValues[i][0])){
          prefMap[j][prepSheetValues[i][5]-1].unpreferred.push(prepSheetValues[i][0]);
        }
      }
    }
    //Unpreferred Duties
    if(prepSheetValues[i][4] != ""){
      for(let j = 0; j < SCHEDULE_VERTICAL; j+=2){
        if(mainSheetValues[3+j][0].toLowerCase() === prepSheetValues[i][4].toLowerCase()){
          for(let k = 0; k < SCHEDULE_HORIZONTAL; k++){
            if(!prefMap[j][k].unpreferred.includes(prepSheetValues[i][0])){
              prefMap[j][k].unpreferred.push(prepSheetValues[i][0]);
            }
            if(!prefMap[j+1][k].unpreferred.includes(prepSheetValues[i][0])){
              prefMap[j+1][k].unpreferred.push(prepSheetValues[i][0]);
            }
          }
        }
      }
    }
    //Preferred Duties
    if(prepSheetValues[i][3] != ""){
      for(let j = 0; j < SCHEDULE_VERTICAL; j+=2){
        if(mainSheetValues[3+j][0].toLowerCase() === prepSheetValues[i][3].toLowerCase()){
          for(let k = 0; k < SCHEDULE_HORIZONTAL; k++){
            if(!prefMap[j][k].preferred.includes(prepSheetValues[i][0])){
              prefMap[j][k].preferred.push(prepSheetValues[i][0]);
            }
            if(!prefMap[j+1][k].preferred.includes(prepSheetValues[i][0])){
              prefMap[j+1][k].preferred.push(prepSheetValues[i][0]);
            }
          }
        }
      }
    }
  }
  return prefMap;
}
