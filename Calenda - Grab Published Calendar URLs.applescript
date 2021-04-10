use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Calendar"
	activate
end tell

set calendarURLs to {}
tell application "System Events"
	tell process "Calendar"
		keystroke "f" using command down
		delay 2
		keystroke tab
		delay 2
		key code 126 using {option down}
		delay 2
		
		repeat with i from 1 to 35
			key code 125
			delay 2
			keystroke "i" using command down
			delay 2
			tell window 1
				tell sheet 1
					set calendarName to value of text field 1
					tell scroll area 1
						tell text area 1
							set theURL to value
						end tell
					end tell
				end tell
			end tell
			keystroke return
			delay 2
			
			
			set end of calendarURLs to {calendarName, theURL}
		end repeat
	end tell
end tell
