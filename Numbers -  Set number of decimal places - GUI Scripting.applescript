-- This script sets the number of decimal places in a cell in Numbers


tell application "Numbers"
	activate
	tell document 1
		tell active sheet
			set the selectedTable to ¬
				(the first table whose class of selection range is range)
		end tell
		tell selectedTable
			tell column "J"
				set the format of cell 3 to number
			end tell
		end tell
	end tell
end tell

tell application "System Events"
	tell process "Numbers"
		tell radio group 1 of window 1
			if value of radio button "Cell" is 0 then
				click radio button "Cell"
				repeat until value of radio button "Cell" is 1
					delay 0.2
				end repeat
			end if
		end tell
		tell text field 1 of scroll area 4 of window 1
			set value of attribute "AXFocused" to true
			set value to "2"
			keystroke return
		end tell
	end tell
end tell