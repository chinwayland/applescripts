use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Numbers"
	set listOfDocuments to name of documents
	set chosenDocument to choose from list listOfDocuments
	set chosenDocument to item 1 of chosenDocument
	tell document chosenDocument
		tell sheet 2
			tell table 1
				# Find table of room numbers and teachers and put it into a list
				set roomIndex to first cell whose value is "Room"
				set columnNumber to address of column of roomIndex
				set rowNumber to address of row of roomIndex
				set rowCount to count of cells in column columnNumber
				set roomNumbers to {}
				repeat with i from rowNumber + 1 to rowCount
					set end of roomNumbers to {(value of cell i of column (columnNumber - 1)), value of cell i of column columnNumber}
				end repeat
				
				set yearName to value of cell 2 of column "M"
				set fullClassName to {}
				repeat with i from 4 to 21
					set className to value of cell i of column "M"
					set teacherName to value of cell i of column "O"
					set end of fullClassName to yearName & " " & className & " " & teacherName
				end repeat
				set yearName to value of cell 23 of column "M"
				repeat with i from 25 to 41
					set className to value of cell i of column "M"
					set teacherName to value of cell i of column "O"
					set end of fullClassName to yearName & " " & className & " " & teacherName
				end repeat
				
				
				set lessons to {}
				repeat with i from 2 to 11
					repeat with j from 4 to 35
						if value of cell j of column i is not missing value then
							set end of lessons to {value of cell 2 of column i, value of cell 3 of column i, value of cell j of column 1, value of cell j of column i}
						end if
					end repeat
				end repeat
			end tell
		end tell
	end tell
end tell

set lessons2 to {}
repeat with lesson in lessons
	set end of lessons2 to {do shell script "echo " & quoted form of item 1 of lesson & " | tr -dc '[:alnum:]' | tr '[:upper:]' '[:lower:]'", item 2 of lesson, item 3 of lesson, item 4 of lesson}
end repeat

set lessons3 to {}
repeat with lesson in lessons2
	if item 2 of lesson contains "大一" then
		set end of lessons3 to {item 1 of lesson, "Year 1", item 3 of lesson, item 4 of lesson}
	else
		set end of lessons3 to {item 1 of lesson, "Year 2", item 3 of lesson, item 4 of lesson}
	end if
end repeat

set lessons4 to {}
repeat with lesson in lessons3
	if item 3 of lesson contains "8:00" then
		set end of lessons4 to {item 1 of lesson, item 2 of lesson, "8:00:00 AM", item 4 of lesson}
	else if item 3 of lesson contains "9:50" then
		set end of lessons4 to {item 1 of lesson, item 2 of lesson, "9:50:00 AM", item 4 of lesson}
	else if item 3 of lesson contains "11:30" then
		set end of lessons4 to {item 1 of lesson, item 2 of lesson, "11:30:00 AM", item 4 of lesson}
	else if item 3 of lesson contains "14:00" then
		set end of lessons4 to {item 1 of lesson, item 2 of lesson, "2:00:00 PM", item 4 of lesson}
	else if item 3 of lesson contains "15:40" then
		set end of lessons4 to {item 1 of lesson, item 2 of lesson, "3:40 PM", item 4 of lesson}
	else if item 3 of lesson contains "18:30" then
		set end of lessons4 to {item 1 of lesson, item 2 of lesson, "6:30 PM", item 4 of lesson}
	end if
end repeat

set lessons5 to {}
repeat with lesson in lessons4
	set UpperFirstCharString to do shell script "echo " & character 1 of item 1 of lesson & " | tr [:lower:] [:upper:]"
	set theString to UpperFirstCharString & characters 2 through -1 of item 1 of lesson
	set end of lessons5 to {item 2 of lesson & " " & item 4 of lesson, theString, item 3 of lesson}
end repeat

set lessons6 to {}
repeat with lesson in lessons5
	if item 2 of lesson is "Monday" then
		set end of lessons6 to {item 1 of lesson, item 2 of lesson & ", March 1, 2021 at " & item 3 of lesson}
	else if item 2 of lesson is "Tuesday" then
		set end of lessons6 to {item 1 of lesson, item 2 of lesson & ", March 2, 2021 at " & item 3 of lesson}
	else if item 2 of lesson is "Wednesday" then
		set end of lessons6 to {item 1 of lesson, item 2 of lesson & ", March 3, 2021 at " & item 3 of lesson}
	else if item 2 of lesson is "Thursday" then
		set end of lessons6 to {item 1 of lesson, item 2 of lesson & ", March 4, 2021 at " & item 3 of lesson}
	else if item 2 of lesson is "Friday" then
		set end of lessons6 to {item 1 of lesson, item 2 of lesson & ", March 5, 2021 at " & item 3 of lesson}
	end if
end repeat

set lessons7 to {}
repeat with lesson in lessons6
	set dateString to item 2 of lesson
	set theDate to date dateString
	set end of lessons7 to {item 1 of lesson, theDate}
end repeat

set lessons8 to {}
repeat with lesson in lessons7
	if item 1 of lesson contains "Alain" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 604-1"}
	else if item 1 of lesson contains "Edward" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 601-1"}
	else if item 1 of lesson contains "Susan" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 602-1"}
	else if item 1 of lesson contains "Tina" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 603-1"}
	else if item 1 of lesson contains "Steven" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 603-2"}
	else if item 1 of lesson contains "Bruce" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 608-1"}
	else if item 1 of lesson contains "David" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 607-1"}
	else if item 1 of lesson contains "Jim" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 609"}
	else if item 1 of lesson contains "Iohann" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 605-1"}
	else if item 1 of lesson contains "Wayland" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "East Building 607-2"}
	end if
	
end repeat


tell application "Calendar"
	activate
	repeat with i from 1 to count of fullClassName
		if not (exists (calendar (item i of fullClassName))) then
			create calendar with name item i of fullClassName
		end if
		
	end repeat
	
	repeat with lesson in lessons8
		tell calendar (item 1 of lesson)
			make new event at end with properties {summary:item 1 of lesson, location:item 3 of lesson, start date:item 2 of lesson, end date:(item 2 of lesson) + (minutes * 80), recurrence:"FREQ=WEEKLY"}
		end tell
	end repeat
end tell