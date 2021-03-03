use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This script converts an Excel timetable to Apple Calendar. Must open the excel timetable in Numbers first, then run this script

display dialog "This script will first delete any calendars that contain the word 'Year' in the calendar name."

repeat
	set theCurrentDate to current date
	display dialog "When is the first day of the semester? (Should be on a Monday, you can enter a delayed start day later)" default answer (date string of theCurrentDate & space & time string of theCurrentDate)
	set theText to text returned of result
	try
		if theText is not "" then
			set firstDay to date theText -- a date object
			exit repeat
		end if
	on error
		beep
	end try
end repeat

display dialog "The first day of the semester is: " default answer (date string of firstDay & space & time string of firstDay)


set chosenCalendarApp to choose from list {"Apple Calendar", "Microsoft Outlook"} with prompt "Which calendar app do you want the calendars created in?"
set chosenCalendarApp to item 1 of chosenCalendarApp

repeat
	set theCurrentDate to current date
	display dialog "When is the last day of the semester?" default answer (date string of theCurrentDate & space & time string of theCurrentDate)
	set theText to text returned of result
	try
		if theText is not "" then
			set theDate to date theText -- a date object
			exit repeat
		end if
	on error
		beep
	end try
end repeat

display dialog "The last day of the semester is: " default answer (date string of theDate & space & time string of theDate)


set theDate to theDate + (days * 1)

set a to (year of theDate)
set b to (month of theDate) as number
set c to day of theDate
#set d to "T000000Z"
set e to "FREQ=WEEKLY;INTERVAL=1;BYDAY=;UNTIL="

set endDate to theDate # this variable represents the last day of the semester for Microsoft Outlook

set theList to {a, b, c}

set {TID, text item delimiters} to {text item delimiters, "0"}
set theListAsString to theList as text
set text item delimiters to TID

set theList to {e, theListAsString}

# the variable endOfRecurrence represents the last day of classes for Apple Calendar
set {TID, text item delimiters} to {text item delimiters, ""}
set endOfRecurrence to theList as text
set text item delimiters to TID

tell application "Numbers"
	activate
	set listOfDocuments to name of documents
	set chosenDocument to choose from list listOfDocuments
	set chosenDocument to item 1 of chosenDocument
	
	tell document chosenDocument
		set listOfSheets to name of sheets
		set chosenSheet to choose from list listOfSheets
		set chosenSheet to item 1 of chosenSheet
		tell sheet chosenSheet
			tell table 1
				# Find table of room numbers and teachers and put it into a list
				set roomIndex to first cell whose value is "Room"
				set columnNumber to address of column of roomIndex
				set rowNumber to address of row of roomIndex
				set rowCount to count of cells in column columnNumber
				set roomNumbers to {}
				repeat with i from rowNumber + 1 to rowCount
					
					set end of roomNumbers to {(value of cell i of column (columnNumber - 1)), value of cell i of column columnNumber as text}
				end repeat
				
				# Create name for all calendars Year X / Class / Teacher Name
				
				set columnNumber to address of column of cell 1 where value of it contains "New Teacher"
				set rowNumber to address of row of cell 1 where value of it contains "New Teacher"
				set className to {}
				
				# Name Year 1 classes
				repeat with i from rowNumber + 1 to count of rows of column columnNumber
					if value of cell i of column columnNumber is not missing value then
						set end of className to "Year 1 " & value of cell i of column (columnNumber - 2) & " " & value of cell i of column columnNumber
					else
						exit repeat
					end if
				end repeat
				
				set columnNumber to address of column of cell 2 where value of it contains "New Teacher"
				set rowNumber to address of row of cell 2 where value of it contains "New Teacher"
				
				# Name Year 2 classes
				repeat with i from rowNumber + 1 to count of rows of column columnNumber
					if value of cell i of column columnNumber is not missing value then
						set end of className to "Year 2 " & value of cell i of column (columnNumber - 2) & " " & value of cell i of column columnNumber
					else
						exit repeat
					end if
				end repeat
				
				# Grab data from spreadsheet
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

# convert year 1 or year 2 from chinese to english
set lessons3 to {}
repeat with lesson in lessons
	if item 2 of lesson contains "大一" then
		set end of lessons3 to {item 1 of lesson, "Year 1", item 3 of lesson, item 4 of lesson}
	else
		set end of lessons3 to {item 1 of lesson, "Year 2", item 3 of lesson, item 4 of lesson}
	end if
end repeat

# add time to each event
set lessons4 to {}
repeat with lesson in lessons3
	set inputTime to item 3 of lesson
	set theTime to word 4 of inputTime & ":" & word 5 of inputTime
	if word 1 of the theTime as number > 12 then
		set theTime2 to (word 1 of theTime) - 12 & ":" & word 2 of theTime & "PM" as text
	else
		set theTime2 to word 1 of theTime & ":" & word 2 of theTime & "AM" as text
	end if
	set end of lessons4 to {item 1 of lesson, item 2 of lesson, theTime2, item 4 of lesson}
end repeat

# Combine Year, class, and teacher name
set lessons5 to {}
repeat with lesson in lessons4
	set end of lessons5 to {item 2 of lesson & " " & item 4 of lesson, item 1 of lesson, item 3 of lesson}
end repeat

# add date to each item in the list
set lessons6 to {}
repeat with lesson in lessons5
	if item 2 of lesson contains "Monday" then
		set end of lessons6 to {item 1 of lesson, date string of firstDay & " at " & item 3 of lesson}
	else if item 2 of lesson contains "Tuesday" then
		set end of lessons6 to {item 1 of lesson, date string of (firstDay + (days * 1)) & " at " & item 3 of lesson}
	else if item 2 of lesson contains "Wednesday" then
		set end of lessons6 to {item 1 of lesson, date string of (firstDay + (days * 2)) & " at " & item 3 of lesson}
	else if item 2 of lesson contains "Thursday" then
		set end of lessons6 to {item 1 of lesson, date string of (firstDay + (days * 3)) & " at " & item 3 of lesson}
	else if item 2 of lesson contains "Friday" then
		set end of lessons6 to {item 1 of lesson, date string of (firstDay + (days * 4)) & " at " & item 3 of lesson}
	end if
end repeat

# convert each item to a date object
set lessons7 to {}
repeat with lesson in lessons6
	set dateString to item 2 of lesson
	set theDate to date dateString
	set end of lessons7 to {item 1 of lesson, theDate}
end repeat

# Add room numbers to the list
set lessons8 to {}
repeat with lesson in lessons7
	repeat with room in roomNumbers
		if item 1 of lesson contains item 1 of room then
			set end of lessons8 to {item 1 of lesson, item 2 of lesson, item 2 of room}
		end if
	end repeat
end repeat


# Creates the calendars in the chosen app
if chosenCalendarApp is "Apple Calendar" then
	-- Deletes previous work calendars
	tell application "Calendar"
		activate
		tell (calendars whose name contains "Year")
			delete
		end tell
	end tell
	
	
	# add calendars and events to the Calendar app
	tell application "Calendar"
		activate
		repeat with i from 1 to count of className
			if not (exists (calendar (item i of className))) then
				create calendar with name item i of className
			end if
		end repeat
		
		repeat with lesson in lessons8
			if calendar (item 1 of lesson) exists then
				tell calendar (item 1 of lesson)
					make new event with properties {summary:item 1 of lesson, location:item 3 of lesson, start date:item 2 of lesson, end date:(item 2 of lesson) + (minutes * 80), recurrence:endOfRecurrence}
				end tell
			else
				set problem to words of item 1 of lesson
				set newEventSummary to item 1 of problem & " " & item 2 of problem & " " & item 3 of problem & " " & item 4 of problem
				tell calendar newEventSummary
					make new event with properties {summary:item 1 of lesson, location:item 3 of lesson, start date:item 2 of lesson, end date:(item 2 of lesson) + (minutes * 80), recurrence:endOfRecurrence}
				end tell
			end if
		end repeat
	end tell
else
	# Outlook --> Incomplete code, needs repeat function
	# Creates all the events on the first day but does not repeat
	tell application "Microsoft Outlook"
		activate
		
		delete (calendars whose name contains "year")
		
		set wintecAccount to first calendar whose email address of accounts contains "wintec"
		tell wintecAccount
			
			repeat with i from 1 to count of className
				if not (exists (calendar (item i of className))) then
					make new calendar with properties {name:item i of className}
				end if
			end repeat
			
			repeat with lesson in lessons8
				if calendar (item 1 of lesson) exists then
					tell calendar (item 1 of lesson)
						make new calendar event with properties {subject:item 1 of lesson, location:item 3 of lesson, start time:item 2 of lesson, end time:(item 2 of lesson) + (minutes * 80), has reminder:false}
					end tell
				else
					set problem to words of item 1 of lesson
					set newEventSummary to item 1 of problem & " " & item 2 of problem & " " & item 3 of problem & " " & item 4 of problem
					tell calendar newEventSummary
						make new calendar event with properties {subject:item 1 of lesson, location:item 3 of lesson, start time:item 2 of lesson, end time:(item 2 of lesson) + (minutes * 80), has reminder:false}
					end tell
				end if
			end repeat
			
		end tell
	end tell
	
	tell application "Microsoft Outlook"
		set allEvents to {}
		set wintecAccount to first calendar whose email address of accounts contains "wintec"
		tell wintecAccount
			set wintecCalendars to calendars whose name contains "year"
			repeat with aCalendar in wintecCalendars
				set theEvents to calendar events of aCalendar
				repeat with anEvent in theEvents
					set end of allEvents to anEvent
				end repeat
			end repeat
		end tell
		repeat with anEvent in allEvents
			subject of anEvent
			set startDay to start time of anEvent
			set startDay to weekday of startDay as text
			
			# next edit -- change start date into a variable instead of a hard date
			if startDay is "Monday" then
				set recurrence of anEvent to {days of week:{monday:true}, end date:{end type:end date type, data:endDate}, start date:date "Monday, March 1, 2021 at 12:00:00 AM", occurrence interval:1, recurrence type:weekly}
			else if startDay is "Tuesday" then
				set recurrence of anEvent to {days of week:{tuesday:true}, end date:{end type:end date type, data:endDate}, start date:date "Tuesday, March 2, 2021 at 12:00:00 AM", occurrence interval:1, recurrence type:weekly}
			else if startDay is "Wednesday" then
				set recurrence of anEvent to {days of week:{wednesday:true}, end date:{end type:end date type, data:endDate}, start date:date "Wednesday, March 3, 2021 at 12:00:00 AM", occurrence interval:1, recurrence type:weekly}
			else if startDay is "Thursday" then
				set recurrence of anEvent to {days of week:{thursday:true}, end date:{end type:end date type, data:endDate}, start date:date "Thursday, March 4, 2021 at 12:00:00 AM", occurrence interval:1, recurrence type:weekly}
			else if startDay is "Friday" then
				set recurrence of anEvent to {days of week:{friday:true}, end date:{end type:end date type, data:endDate}, start date:date "Friday, March 5, 2021 at 12:00:00 AM", occurrence interval:1, recurrence type:weekly}
			end if
			
		end repeat
	end tell
	
	
end if