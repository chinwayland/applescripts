use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "CalendarLib EC" version "1.1.5" # This external library is for changing time zones of events


tell application "Numbers"
	tell document 1
		tell sheet 1
			tell table 1
				set columnStart to 3
				set tableData to {}
				set rowStart to 5
				set testInfo to {}
				set firstDay to value of cell 2 of column 3
				set theDate to words of firstDay
				set theDate to item 1 of theDate
				set theDay to character 1 of theDate & character 2 of theDate
				set testDate to "12/" & theDay & "/22"
				tell me to set testDate to date testDate
				set firstDay to testDate
				repeat with columnIndex from columnStart to 13 by 5
					set testDate to value of cell 2 of column columnIndex
					set theDate to words of testDate
					set theDate to item 1 of theDate
					set theDay to character 1 of theDate & character 2 of theDate
					set testDate to "12/" & theDay & "/22"
					tell me to set testDate to date testDate
					repeat with rowIndex from rowStart to rowStart + 14
						if rowIndex < 9 then
							set eventTime to testDate + (hours * 8) + (minutes * 10)
						else if rowIndex < 13 then
							set eventTime to testDate + (hours * 10) + (minutes * 10)
						else if rowIndex < 17 then
							set eventTime to testDate + (hours * 13) + (minutes * 40)
						else
							set eventTime to testDate + (hours * 15) + (minutes * 20)
						end if
						if value of cell rowIndex of column columnIndex is not missing value then
							if value of cell rowIndex of column columnIndex is not "x" then
								set end of testInfo to {eventTime, value of cell rowIndex of column columnIndex, value of cell rowIndex of column (columnIndex + 1), value of cell rowIndex of column (columnIndex + 2), value of cell rowIndex of column (columnIndex + 3)}
							end if
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
		tell me to say "Deleting previous Calendars with the word \"Written\" in it."
		tell (calendars whose name contains "Written")
			delete
		end tell
	end tell
	
	create calendar with name "Final Exams - Written"
	tell me to say "Creating Events"
	
	repeat with info in testInfo
		tell calendar "Final Exams - Written"
			make new event with properties {summary:item 2 of info & " " & item 3 of info & " " & item 5 of info, start date:item 1 of info, end date:(item 1 of info) + (minutes * 80), location:item 4 of info}
		end tell
	end repeat
	
	
end tell


-- Sets the time zone
tell me to say "Changing the time zone to the Chinese time zone"
tell application "Calendar"
	set theCalendars to name of calendars whose name contains "Exams"
end tell

set theStore to fetch store
set startDate to firstDay
set endDate to startDate + (days * 3)
#set theCalendar to item 1 of theCalendars

set theCalendars2 to fetch calendars theCalendars cal type list cal local event store theStore

set theEvents to fetch events starting date startDate ending date endDate searching cals theCalendars2 event store theStore

repeat with i from 1 to count of theEvents
	set theNewEvent to modify zone event item i of theEvents time zone "Asia/Shanghai"
	store event event theNewEvent event store theStore
end repeat
