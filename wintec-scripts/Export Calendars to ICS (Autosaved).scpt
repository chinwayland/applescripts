-- This script exports multiple calendars to ICS one by one using GUI Scripting

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

set numberOfCalendars to display dialog "How many calendars?" default answer ""
set numberOfCalendars to text returned of numberOfCalendars
set numberOfCalendars to numberOfCalendars as integer
set numberOfCalendars to numberOfCalendars + 1

tell application "Calendar"
	activate
end tell

tell application "System Events"
	tell process "Calendar"
		repeat with i from 2 to numberOfCalendars
			
			# Select each calendar
			tell row i of outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
				set selected to true
				delay 1
			end tell
			
			# Click Export from Menu
			tell menu item 1 of menu of menu item "Export" of menu "File" of menu bar 1
				click
				delay 3
			end tell
			
			# Click Export Button from pop up window
			tell button "Export" of sheet 1 of window 1
				click
				delay 3
			end tell
			
			
			if button "Replace" of sheet 1 of sheet 1 of window 1 exists then
				tell button "Replace" of sheet 1 of sheet 1 of window 1
					click
				end tell
				delay 3
			end if
		end repeat
	end tell
end tell