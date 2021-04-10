use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- Publishes the Calendars as Public

tell application "Calendar"
	activate
end tell

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
			tell menu bar 1
				tell menu bar item "Edit"
					tell menu "Edit"
						click
						delay 2
						tell UI element 15
							click
							delay 2
						end tell
					end tell
				end tell
			end tell
			
			tell window 1
				tell splitter group 1
					tell splitter group 1
						tell scroll area 1
							tell outline 1
								tell row (i + 2)
									tell pop over 1
										set shareCalendar to checkbox 1
										set doneButton to button "Done"
										click shareCalendar
										delay 2
										click doneButton
										delay 2
										keystroke tab
										delay 2
									end tell
								end tell
							end tell
						end tell
					end tell
				end tell
			end tell
			
			
		end repeat
	end tell
end tell

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

-- Grabs the published calendars and writes the URLs to a numbers document
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