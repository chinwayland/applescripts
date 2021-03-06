--  This script will:
--  Close all workbooks in Excel
--  Ask you where the Easy-LMS spreadsheet is
--  Ask you wehre the gradebook is
--  Open Easy-LMS Spreadsheet first
--  Then open the gradebook second
--  This will do most of the heavy lifting
--  Aftr this script is done, you can manually inspect the gradebook for missing scores

--	Created by: Wayland Chin
--	Created on: 6/20/20
--
--	Copyright © 2020 Wintec, All Rights Reserved

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Microsoft Excel"
	close workbooks
end tell

display dialog "Please tell me where the easy lms file is."

set fileEasyLMS to choose file with prompt "Please select the Easy LMS spreadsheet to process:"

display dialog "Please tell me where the grade book is."

set fileGradebook to choose file with prompt "Please select the gradebook"

tell application "Microsoft Excel" to open fileEasyLMS

tell application "Microsoft Excel" to open fileGradebook

tell application "Microsoft Excel"
	set students to {}
	tell workbook 1
		tell worksheet 1
			repeat with i from 2 to count of rows of used range
				if the (count of ((value of cell ("P" & i)) as string)) is 15 then #Make sure student ID is 15 digits
					if the value of cell ("J" & i) as integer is greater than 41 then #Make sure at least 42 questions were answered
						set end of students to {value of cell ("P" & i), value of cell ("J" & i), value of cell ("K" & i), value of cell ("D" & i)}
					end if
				end if
			end repeat
		end tell
	end tell
end tell

#find duplicates scores and put it into the list named duplicates
set noDuplicates to {}
set duplicates to {}
repeat with student in students
	if item 1 of student is not in noDuplicates then
		set end of noDuplicates to item 1 of student
	else
		if item 1 of student is not in duplicates then
			set end of duplicates to item 1 of student
		end if
	end if
end repeat

#create a list of students with multiple exam attempts
set duplicatesWithData to {}
repeat with i from 1 to count of duplicates
	set end of duplicatesWithData to {}
	repeat with j from 1 to count of students
		if item i of duplicates is in item 1 of item j of students then
			set end of item i of duplicatesWithData to item j of students
		end if
	end repeat
end repeat

#convert date time string to date object and add it to the list
repeat with duplicates in duplicatesWithData
	repeat with duplicat in duplicates
		set end of duplicat to convertDate(item 4 of duplicat)
	end repeat
end repeat

#get earliest time
set earliest to current date
set earliestData to {}
repeat with i from 1 to count of duplicatesWithData
	set earliest to current date
	repeat with j from 1 to count of item i of duplicatesWithData
		if item 5 of item j of item i of duplicatesWithData comes before earliest then
			set earliest to item 5 of item j of item i of duplicatesWithData
			set end of earliestData to {¬
				item 1 of item j of item i of duplicatesWithData, ¬
				item 2 of item j of item i of duplicatesWithData, ¬
				item 3 of item j of item i of duplicatesWithData}
		end if
	end repeat
end repeat

#add earliest data from duplicate data to final student list to be used to add data to excel

repeat with earliest in earliestData
	set end of students to {¬
		item 1 of earliest, ¬
		item 2 of earliest, ¬
		item 3 of earliest}
end repeat

tell application "Microsoft Excel"
	# Enter scores of all students, will include the latest attempt for duplicates - which will be fixed later.
	repeat with student in students
		tell workbook 2
			repeat with i from 1 to count of worksheets
				tell worksheet i
					try
						set targetRow to first row index of (find used range what (item 1 of student))
						set value of cell ("G" & targetRow) to item 2 of student
						set value of cell ("H" & targetRow) to item 3 of student
						exit repeat
					end try
				end tell
			end repeat
		end tell
	end repeat
end tell

-- Convert date function. Call with string in YYYY-MM-DD HH:MM:SS format (time part optional)
to convertDate(textDate)
	
	set resultDate to the current date
	set the month of resultDate to (1 as integer)
	set the day of resultDate to (1 as integer)
	
	set the year of resultDate to (text 1 thru 4 of textDate)
	set the month of resultDate to (text 6 thru 7 of textDate)
	set the day of resultDate to (text 9 thru 10 of textDate)
	set the time of resultDate to 0
	
	if (length of textDate) > 10 then
		set the hours of resultDate to (text 12 thru 13 of textDate)
		set the minutes of resultDate to (text 15 thru 16 of textDate)
		
		if (length of textDate) > 16 then
			set the seconds of resultDate to (text 18 thru 19 of textDate)
		end if
	end if
	
	return resultDate
end convertDate
