-- This script deletes all school calendars in the Calendar app

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Calendar"
	activate
	tell (calendars whose name contains "Year")
		delete
	end tell
end tell