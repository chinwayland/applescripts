-- This script converts an Excel timetable to Apple Calendar. Must open the excel timetable in Numbers first, then run this script

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "CalendarLib EC" version "1.1.5" # This external library is for changing time zones of events

say "Please open the calendar in Apple Numbers first. This script will first delete any calendars that contain the word 'Year' in the calendar name. Please confirm"
display dialog "This script will first delete any calendars that contain the word 'Year' in the calendar name."

set chosenCalendarApp to "Apple Calendar"

say "When is the first day of the semester? (Should be on a Monday, you can enter a delayed start day later)"
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

say "The first day of the semester is: " & firstDay & "Please Confirm, by pressing OK"

display dialog "The first day of the semester is: " default answer (date string of firstDay & space & time string of firstDay)

say "When is the last day of the semester?"
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

say "The last day of the semester is: " & theDate & "Please Confirm, by pressing OK"
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

set classNames to {}
tell application "Numbers"
	activate
	set listOfDocuments to name of documents
	tell me to say "Choose which Numbers document has the timetable and class names."
	set chosenDocument to choose from list listOfDocuments
	set chosenDocument to item 1 of chosenDocument
	
	
	tell document chosenDocument
		tell me to say "Choose the sheet with the timetable."
		set listOfSheets to name of sheets
		set chosenSheet to choose from list listOfSheets
		set chosenSheet to item 1 of chosenSheet
		
		tell me to say "Getting class names."
		
		tell sheet chosenSheet
			tell table 1
				repeat with columnIndex from 2 to 15
					repeat 1 times
						set yearValue to value of cell 3 of column columnIndex
						if yearValue contains "大一" then
							set yearValue to "Year 1"
						else if yearValue contains "大二" then
							set yearValue to "Year 2"
						else
							exit repeat
						end if
					end repeat
					
					repeat with rowIndex from 4 to 37
						if value of cell rowIndex of column columnIndex is not missing value then
							if yearValue & " " & value of cell rowIndex of column columnIndex is not in classNames then
								set end of classNames to yearValue & " " & value of cell rowIndex of column columnIndex
							end if
							
						end if
						
					end repeat
					
				end repeat
			end tell
		end tell
		tell me to say "Grabbing data from the numbers spreadsheet."
		tell sheet chosenSheet
			tell table 1
				# Find table of room numbers and teachers and put it into a list
				set roomIndex to first cell whose value is "Room"
				set columnNumber to address of column of roomIndex
				set rowNumber to address of row of roomIndex
				set rowCount to count of cells in column columnNumber
				set roomNumbers to {}
				repeat with i from rowNumber + 1 to rowCount
					if value of cell i of column (columnNumber - 1) is not missing value then
						set end of roomNumbers to {(value of cell i of column (columnNumber - 1)), value of cell i of column columnNumber as text}
					end if
					
				end repeat
				# Grab data from spreadsheet
				set lessons to {}
				repeat with i from 2 to 15
					repeat with j from 4 to 34
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
	if item 1 of lesson contains "/" then
		set end of lessons8 to {item 1 of lesson, item 2 of lesson, "Multiple"}
	else
		repeat with room in roomNumbers
			if item 1 of lesson contains item 1 of room then
				set end of lessons8 to {item 1 of lesson, item 2 of lesson, item 2 of room}
			end if
		end repeat
	end if
end repeat

# Creates the calendars in the chosen app
if chosenCalendarApp is "Apple Calendar" then
	-- Deletes previous work calendars
	tell application "Calendar"
		activate
		tell me to say "Deleting previous Calendars with the word \"Year\" in it."
		
		tell (calendars whose name contains "Year")
			delete
		end tell
	end tell
	
	# add calendars and events to the Calendar app
	say "Creating Calendars"
	tell application "Calendar"
		activate
		repeat with i from 1 to count of classNames
			if not (exists (calendar (item i of classNames))) then
				create calendar with name item i of classNames
			end if
		end repeat
		tell me to say "Creating Events"
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
	
	-- Sets the time zone
	tell me to say "Changing the time zone to the Chinese time zone"
	
	tell application "Calendar"
		set theCalendars to name of calendars whose name contains "Year"
	end tell
	
	set theStore to fetch store
	
	set startDate to firstDay
	set endDate to startDate + days * 5
	
	set theCalendars to fetch calendars theCalendars cal type list cal local event store theStore
	
	set theEvents to fetch events starting date startDate ending date endDate searching cals theCalendars event store theStore
	
	repeat with i from 1 to count of theEvents
		set theNewEvent to modify zone event item i of theEvents time zone "Asia/Shanghai"
		store event event theNewEvent event store theStore
	end repeat
	
	# export calendars to the desktop
	
	say "Exporting the calendars to the Desktop"
	set calendarNames to {}
	tell application "Calendar" to activate
	tell application "System Events"
		tell process "Calendar"
			keystroke "f" using command down # search bar
			#		delay 1
			keystroke tab # moves focus to calendar list
			#		delay 1
			key code 126 using {option down} # up arrow --> moves focus to first calendar / option up arrow
			#		delay 1
			set countOfRows to count of rows in outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
			repeat with i from 1 to countOfRows
				try
					set calendarName to value of text field 1 of UI element 1 of row i of outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
					if calendarName contains "Year" then
						set end of calendarNames to calendarName
						key code 120 using {control down}
						key code 124 #right arrow
						key code 124 #right arrow
						repeat 7 times
							key code 125 # down arrow
						end repeat
						key code 124 # right arrow
						keystroke return
						delay 1
						keystroke "D" using {command down, shift down} # change to desktop
						keystroke return # save to desktop
						delay 1
						if button 1 of sheet 1 of sheet 1 of window 1 exists then
							keystroke tab
							keystroke space
							delay 1
						end if
						key code 125 # down arrow
						keystroke "f" using command down # search bar
						keystroke tab # moves focus to calendar list
					else
						key code 125 # down arrow
					end if
				end try
			end repeat
		end tell
	end tell
	
	tell me to say "Calendars have been exported to your desktop"
	
	# This turns off and on the iCloud Calendar in preferences so the calendars gets uploaded to the cloud
	say "Moving Calendars from local to iCloud"
	tell application "System Preferences"
		activate
		delay 3
		quit
		delay 3
		activate
		delay 3
	end tell
	
	tell application "System Events"
		tell process "System Preferences"
			tell window 1
				tell button 1
					click
					delay 3
				end tell
			end tell
		end tell
	end tell
	
	tell application "System Events"
		tell process "System Preferences"
			tell window 1
				tell group 1
					tell scroll area 1
						tell table 1
							tell row 8
								tell UI element 1
									tell UI element 1
										tell me to say "Unchecking iCloud Calendars."
										click # turn off icloud for calendar
										delay 5
										tell me to say "Checking iCloud Calendars. Please wait 180 seconds. Moving calendars to iCloud."
										click # turn on icloud for calendar
										delay 180
									end tell
								end tell
							end tell
						end tell
					end tell
				end tell
			end tell
		end tell
	end tell
else
	# Outlook --> Incomplete code, needs repeat function
	# Creates all the events on the first day but does not repeat
	tell application "Microsoft Outlook"
		activate
		
		delete (calendars whose name contains "Year")
		
		set wintecAccount to first calendar whose email address of accounts contains "wintec"
		tell wintecAccount
			
			repeat with i from 1 to count of classNames
				if not (exists (calendar (item i of classNames))) then
					make new calendar with properties {name:item i of classNames}
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
			set wintecCalendars to calendars whose name contains "Year"
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

say "Finished."

on simple_sort(my_list)
	set the index_list to {}
	set the sorted_list to {}
	repeat (the number of items in my_list) times
		set the low_item to ""
		repeat with i from 1 to (number of items in my_list)
			if i is not in the index_list then
				set this_item to item i of my_list as text
				if the low_item is "" then
					set the low_item to this_item
					set the low_item_index to i
				else if this_item comes before the low_item then
					set the low_item to this_item
					set the low_item_index to i
				end if
			end if
		end repeat
		set the end of sorted_list to the low_item
		set the end of the index_list to the low_item_index
	end repeat
	return the sorted_list
end simple_sort

