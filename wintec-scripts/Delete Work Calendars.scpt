use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Calendar"
	tell (calendars whose name contains "Year")
		delete
	end tell
end tell