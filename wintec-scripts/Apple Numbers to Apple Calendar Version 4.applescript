use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "CalendarLib EC" version "1.1.5" -- put this at the top of your scripts

-- This script converts an Excel timetable to Apple Calendar. Must open the excel timetable in Numbers first, then run this script

say "Please open the calendar in Apple Numbers first. This script will first delete any calendars that contain the word 'Year' in the calendar name. Please confirm"
display dialog "This script will first delete any calendars that contain the word 'Year' in the calendar name."

#say "Which calendar app do you want to create the events in? Apple Calendar or Microsoft Outlook?"
#set chosenCalendarApp to choose from list {"Apple Calendar", "Microsoft Outlook"} with prompt "Which calendar app do you want the calendars created in?"
#set chosenCalendarApp to item 1 of chosenCalendarApp

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

(*
-- Ask for delayed start
say "Is the first day of the semester delayed?"
set delayedYesOrNo to choose from list {"Yes", "No"} with prompt "Is the first day of the semester delayed?"
set delayedYesOrNo to item 1 of delayedYesOrNo

if delayedYesOrNo is "Yes" then
	say "If the first day is delayed and not on a Monday, when is the delayed first day?"
	repeat
		set theCurrentDate to current date
		display dialog "If the first day is delayed and not on a Monday, when is the delayed first day?" default answer (date string of theCurrentDate & space & time string of theCurrentDate)
		set theText to text returned of result
		try
			if theText is not "" then
				set delayedFirstDay to date theText -- a date object
				exit repeat
			end if
		on error
			beep
		end try
	end repeat
end if
*)

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
		(*
		tell me to say "Choose the sheet named \"Classes\"."
		set listOfSheets to name of sheets
		set chosenSheet to choose from list listOfSheets
		set chosenSheet to item 1 of chosenSheet
		tell sheet chosenSheet
			tell table 1
				# get year one class names
				set yearOneBegin to first cell whose value of it contains "Class"
				set yearOneColumnAddress to address of column of yearOneBegin
				set yearOneColumnCount to count of cells of column yearOneColumnAddress
				set yearOneBegin to (address of row of yearOneBegin) + 1
				
				repeat with i from yearOneBegin to yearOneColumnCount
					if value of cell i of column yearOneColumnAddress is missing value then
						exit repeat
					else
						#						set end of classNames to "Year 1 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))
						set end of classNames to value of cell i of column (yearOneColumnAddress + 4) & " Year 1 " & (value of cell i of column yearOneColumnAddress)
					end if
					#					value of cell i of column yearOneColumnAddress
				end repeat
				
				# get year two class names
				set yearOneBegin to second cell whose value of it contains "Class"
				set yearOneColumnAddress to address of column of yearOneBegin
				set yearOneColumnCount to count of cells of column yearOneColumnAddress
				set yearOneBegin to (address of row of yearOneBegin) + 1
				
				repeat with i from yearOneBegin to yearOneColumnCount
					if value of cell i of column yearOneColumnAddress is missing value then
						exit repeat
					else
						#						set end of classNames to "Year 2 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))
						set end of classNames to value of cell i of column (yearOneColumnAddress + 4) & " Year 2 " & (value of cell i of column yearOneColumnAddress)
					end if
					#					value of cell i of column yearOneColumnAddress
				end repeat
				
				set sortedList to my simple_sort(the classNames)
				set classNames to sortedList
				
			end tell
			
		end tell
		*)
		
		
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
				(*
				# Create name for all calendars Year X / Class / Teacher Name
				
				set columnNumber to address of column of cell 1 where value of it contains "New Teacher"
				set rowNumber to address of row of cell 1 where value of it contains "New Teacher"
				set className to {}
				
				# Name Year 1 classes
				repeat with i from rowNumber + 1 to count of rows of column columnNumber
					if value of cell i of column columnNumber is not missing value then
						set end of className to "Bear 1 " & value of cell i of column (columnNumber - 2) & " " & value of cell i of column columnNumber
					else
						exit repeat
					end if
				end repeat
				
				set columnNumber to address of column of cell 2 where value of it contains "New Teacher"
				set rowNumber to address of row of cell 2 where value of it contains "New Teacher"
				
				# Name Year 2 classes
				repeat with i from rowNumber + 1 to count of rows of column columnNumber
					if value of cell i of column columnNumber is not missing value then
						set end of className to "Bear 2 " & value of cell i of column (columnNumber - 2) & " " & value of cell i of column columnNumber
					else
						exit repeat
					end if
				end repeat
				*)
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

(*
set classNames to {}
repeat with lesson in lessons8
	set end of classNames to item 1 of lesson
end repeat

set classNames2 to addr(classNames)
*)

(*
set lessons9 to {}
#set classNamesTeacherFirst to {}
repeat with lesson in lessons8
	if item 1 of lesson contains "/" then
		set teacherFirst to word 4 of item 1 of lesson
		set teacherSecond to word 5 of item 1 of lesson
		set yearWord to word 1 of item 1 of lesson
		set yearNumber to word 2 of item 1 of lesson
		set className to word 3 of item 1 of lesson
		set eventName to teacherFirst & "/" & teacherSecond & " " & yearWord & " " & yearNumber & " " & className
	else
		set teacher to word 4 of item 1 of lesson
		set yearWord to word 1 of item 1 of lesson
		set yearNumber to word 2 of item 1 of lesson
		set className to word 3 of item 1 of lesson
		set eventName to teacher & " " & yearWord & " " & yearNumber & " " & className
	end if
	set end of lessons9 to {eventName, item 2 of lesson, item 3 of lesson}
end repeat
*)


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
	
	(*
	-- remove days because of delayed start
	if delayedYesOrNo is "Yes" then
		set theStore to fetch store
		
		set theCalendars to fetch calendars {} cal type list cal cloud event store theStore
		set startingDate to firstDay
		set endingDate to delayedFirstDay - (days * 1)
		set theEvents to fetch events starting date startingDate ending date endingDate searching cals theCalendars event store theStore
		
		repeat with theEvent in theEvents
			remove event event theEvent event store theStore without future events
		end repeat
		
	end if
	*)
	
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
	(*
	tell application "System Events"
		tell process "System Preferences"
			tell window 1
				tell sheet 1
					tell button 2
						click #click the merge button
						delay 90
					end tell
				end tell
			end tell
		end tell
	end tell
	*)
	
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
						#					delay 1
						key code 124 #right arrow
						#					delay 1
						key code 124 #right arrow
						#					delay 1
						keystroke return
						#					delay 1
						repeat 5 times
							key code 125 # down arrow
							#						delay 1
						end repeat
						key code 124 # right arrow
						#					delay 1
						keystroke return
						delay 1
						keystroke "D" using {command down, shift down} # change to desktop
						#					delay 1
						keystroke return # save to desktop
						delay 1
						if button 1 of sheet 1 of sheet 1 of window 1 exists then
							keystroke tab
							#						delay 1
							keystroke space
							delay 1
						end if
						key code 125 # down arrow
						#					delay 1
						keystroke "f" using command down # search bar
						#					delay 1
						keystroke tab # moves focus to calendar list
						#					delay 1
					else
						key code 125 # down arrow
					end if
				end try
			end repeat
		end tell
	end tell
	
	tell me to say "check out your desktop"
	(*
	tell application "Calendar"
		activate
		delay 1
	end tell
	
	say "Publishing the calendars by making them public"
	tell application "System Events"
		tell process "Calendar"
			set calendarNames to {}
			set calendarURLs to {}
			
			#keystroke "f" using command down # search bar
			#delay 1
			keystroke tab # moves focus to calendar list
			delay 1
			key code 126 using option down # go to first one
			delay 1
			
			set countOfRows to count of rows in outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
			repeat with i from 1 to countOfRows
				try
					set calendarNameText to value of text field 1 of UI element 1 of row i of outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
					if calendarNameText contains "Year" then
						set end of calendarNames to calendarNameText
						tell menu "Edit" of menu bar item "Edit" of menu bar 1
							click # click Edit Menu
						end tell
						tell UI element 15 of menu "Edit" of menu bar item "Edit" of menu bar 1
							click # click Share Calendar under Edit Menu
						end tell
						#delay 1
						tell checkbox 1 of pop over 1 of row i of outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
							click # click the share calendar checkbox
						end tell
						#delay 1
						tell button "Done" of pop over 1 of row i of outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
							click
						end tell
						
						keystroke "I" using command down
						delay 1
						tell window 1
							tell sheet 1
								tell scroll area 1
									UI elements
									tell text area 1
										set theURL to value
									end tell
								end tell
							end tell
						end tell
						set end of calendarURLs to {calendarNameText, theURL}
						
						
						#delay 1
						keystroke tab
						delay 1
						key code 125 # down arrow
						delay 1
					else
						key code 125 # down arrow
					end if
				end try
			end repeat
		end tell
	end tell
	
	
	say "Pasting public calendar links into a Numbers spreadsheet"
	tell application "Numbers"
		activate
		make new document
		tell document 1
			tell sheet 1
				tell table 1
					tell row 1
						set value of cell 1 to "Calendar Name"
						set value of cell 2 to "URL"
					end tell
					
					repeat with i from 1 to count of calendarURLs
						set calendarName to item 1 of item i of calendarURLs
						set theURL to item 2 of item i of calendarURLs
						if not (exists row (i + 1)) then
							make new row
						end if
						tell row (i + 1)
							set value of cell 1 to calendarName
							set value of cell 2 to theURL
						end tell
						
					end repeat
				end tell
			end tell
		end tell
	end tell
	*)
	
	
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

