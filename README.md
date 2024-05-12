# Supervision-Scheduler

A generator made in Spring 2024 for Ottawa-Carleton District School Board. Organizes supervisions for teachers automatically, in the best way possible, avoiding schedule conflicts and multiple supervisions on one day.

Runs in Google Sheets, using Google App Script.

# Using the project

This project will automatically generate supervision schedules, given a list (google sheet) of teachers, with their prep periods, and preferences regarding supervision. It will print its output in the form of a google sheet, under the section titled "Supervising Calendar". If doing so is possible, it will always ensure that every teacher's prep period fits their duty, and teachers are never scheduled to their unpreferred duties and days. If possible, it will also try to split the work as evenly as possible, and put teachers with their preferred duties.  

Of course, if needed, the administrator can always make manual modifications to the calendar and prep list, such as adding and removing teachers and their duties, or switching a teacher's duties with another, using the project's built-in functions. If they need to make a direct modification to the spreadsheet, there is also a validate feature to make sure that their edits do not cause any conflicts.

Instructions for every feature:
Features with an asterisk (*) do not require input. 

Generate Teacher Schedule*: 
Generates the supervision schedule. The data is read from the Prep List ("Teacher Profiles"). The generated calendar will meet the restrictions as described in the description, to the best of its ability.

Change teacher schedule:
Allows for the switching of a teacher's duties with another.

Input 1: The two teachers to be switched
Error "Invalid Format!":
The format must be as follows: [teacher1],[teacher2].
Acceptable ones include: "Torres,Jin", "Smith,JohnSmith"
Unacceptable ones include: "Torres, Jin" (extra space), "Torres,Jin " (extra space at the end), "Torres Jin" (no comma)
Error "One of the teachers do not exist!": One of the teacher's names cannot be found in the prep list. 
All duties of teacher 1 will now belong to teacher 2, and vice versa.


Generate Calendar*:
Generates the heading of the spreadsheet. It automatically generates the days of the week, the holidays and PA days, and if there are any inconsistencies, they can always be fixed manually, without impacting other aspects of the project.
Input 1: the month. (1-12) for January to December.
Input 2: the year.
The legend, corresponding to the abbrivations of all of the holidays, can be found to the right of the spreadsheet.

Validate Calendar*:
Checks the schedule for conflicts, and counts the number of conflicts found.
Cells with a red background color: The name in this cell is invalid (i.e. the teacher's name is not on the prep list)
Cells with a orange background color: The name is valid, but the period conflicts with their prep period indicated in the prep list. (e.g.: it is Day 1, the teacher has a prep period in Period 1, but they are assigned a lunch duty).
Cells with a yellow background color: The teacher has stated that this period/day is their unpreferrable period/day, so it is not optimal (e.g.: a teacher who does not want to be assigned a duty on Friday is assigned one this Friday).

Make Schedule by Teachers*:
Re-formats the calendar into one which is more convinient for teachers to check. Creates a new spreadsheet named "Teachers - (followed by the month and year)".

Import Duties:
Input 1: The teacher's name.
Error "Invalid Input!": The name does not exist on the calendar. The input is case-sensitive.
Imports all duties of the given teacher into the user's personal Google Calendar.
