-- This script deletes all school calendars in Microsoft Outlook

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Microsoft Outlook"
	set calendarsHolder to calendars of calendar 2
	repeat with i from 1 to count of calendarsHolder
		if name of item i of calendarsHolder contains "Year" then
			delete item i of calendarsHolder
		end if
	end repeat
end tell