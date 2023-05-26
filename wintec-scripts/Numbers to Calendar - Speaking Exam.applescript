use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "CalendarLib EC" version "1.1.5" # This external library is for changing time zones of events

tell application "Numbers"
	tell document 1
		tell sheet 1
			tell table 1
				#				set tableData to {}
				set testInfo to {}
				set columnStart to 2
				set rowStart to 4
				set firstDay to date "Thursday, June 1, 2023 at 12:00:00 AM"
				repeat with columnIndex from columnStart to 15
					if columnIndex < 4 then
						set testDate to date "Thursday, June 1, 2023 at 12:00:00 AM"
					else if columnIndex < 7 then
						set testDate to date "Friday, June 2, 2023 at 12:00:00 AM"
					else if columnIndex < 10 then
						set testDate to date "Monday, June 5, 2023 at 12:00:00 AM"
					else if columnIndex < 13 then
						set testDate to date "Tuesday, June 6, 2023 at 12:00:00 AM"
					else
						set testDate to date "Wednesday, June 7, 2023 at 12:00:00 AM"
					end if
					
					repeat with rowIndex from rowStart to count of rows
						if rowIndex < 13 then
							set eventTime to testDate + (hours * 8) + (minutes * 10)
						else if rowIndex < 21 then
							set eventTime to testDate + (hours * 10) + (minutes * 10)
						else if rowIndex < 31 then
							set eventTime to testDate + (hours * 14) + (minutes * 10)
						else
							set eventTime to testDate + (hours * 15) + (minutes * 50)
						end if
						if value of cell rowIndex of column columnIndex is not missing value then
							set end of testInfo to {eventTime, value of cell rowIndex of column columnIndex}
						end if
					end repeat
				end repeat
			end tell
		end tell
	end tell
end tell


tell application "Calendar"
	activate
	switch view to week view
	view calendar at testDate
	
	# Deletes previous work calendars
	tell application "Calendar"
		activate
		tell me to say "Deleting previous Calendars with the word \"Speaking\" in it."
		tell (calendars whose name contains "Speaking")
			delete
		end tell
	end tell
	
	create calendar with name "Final Exams - Speaking"
	
	tell me to say "Creating Events"
	
	repeat with info in testInfo
		tell calendar "Final Exams - Speaking"
			make new event with properties {summary:item 2 of info, start date:item 1 of info, end date:(item 1 of info) + (minutes * 80)}
		end tell
	end repeat
	
end tell


-- Sets the time zone
tell me to say "Changing the time zone to the Chinese time zone"
tell application "Calendar"
	set theCalendars to name of calendars whose name contains "Speaking"
end tell

set theStore to fetch store
set startDate to firstDay
set endDate to startDate + (days * 7)
set theCalendar to item 1 of theCalendars

set theCalendar2 to fetch calendar theCalendar cal type cal local event store theStore

set theEvents to fetch events starting date startDate ending date endDate searching cals theCalendar2 event store theStore
repeat with i from 1 to count of theEvents
	set theNewEvent to modify zone event item i of theEvents time zone "Asia/Shanghai"
	store event event theNewEvent event store theStore
end repeat

set allInfo to ""
repeat with examsInfo in testInfo
	repeat with examInfo in examsInfo
		set allInfo to (allInfo & examInfo as string) & " "
	end repeat
	set allInfo to allInfo & return & return
end repeat