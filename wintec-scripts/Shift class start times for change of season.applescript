use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- next action is to put this in a nested loop, big loop five days of the week, small loop periods 1 to 4

set startDate to date "Friday, February 17, 2023 at 2:40:00 PM"
set theEvents to {}
tell application "Calendar"
	activate
	set theCalendars to calendars
	repeat with theCalendar in theCalendars
		set theEvents to theEvents & (every event in theCalendar ¬
			whose start date = startDate)
	end repeat
	
	repeat with theEvent in theEvents
		set theStartDate to start date of theEvent
		set theNewStartDate to theStartDate - (20 * minutes)
		set theEndDate to end date of theEvent
		set theNewEndDate to theEndDate - (20 * minutes)
		set start date of theEvent to theNewStartDate
		set end date of theEvent to theNewEndDate
	end repeat
end tell