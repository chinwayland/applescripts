-- This script turns off all calendar alerts for a calendar

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Calendar"
	set calendarList to name of calendars
	set calendarName to choose from list calendarList
	set calendarName to item 1 of calendarName
end tell


tell application "Calendar" to tell calendar calendarName
	set theEvents to get events
	set countEvents to count theEvents
end tell
if countEvents = 0 then return

set progress total steps to countEvents

repeat with i from 1 to countEvents
	set progress description to "Process Event " & i & " of " & countEvents
	
	tell application "Calendar" to tell calendar calendarName
		
		set anEvent to item i of theEvents
		tell anEvent
			set |allday event| to allday event
			set |start date| to start date
			set |end date| to end date
			set |url| to url
			set |location| to location
			set |recurrence| to recurrence
			set |description| to description
			set |excluded dates| to excluded dates
			set |stamp date| to stamp date
			set |summary| to summary
		end tell
		
		delete anEvent
		
		set newEvent to (make new event at end of events with properties {allday event:|allday event|, start date:|start date|, end date:|end date|})
		tell newEvent
			if not (|url| is missing value) then set url to |url|
			if not (|location| is missing value) then set location to |location|
			if not (|recurrence| is missing value) then set recurrence to |recurrence|
			if not (|description| is missing value) then set description to |description|
			if not (|summary| is missing value) then set summary to |summary|
			if not (|excluded dates| is missing value) then set excluded dates to |excluded dates|
			if not (|stamp date| is missing value) then set stamp date to |stamp date|
		end tell
		
	end tell
	
	set progress completed steps to i
end repeat