use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Numbers"
	tell document 1
		repeat with sheetNumber from 1 to 4
			tell sheet sheetNumber
				tell table 1
					tell column "F"
						repeat with cellNumber from 2 to count of cells
							get cell cellNumber
						end repeat
					end tell
				end tell
			end tell
		end repeat
	end tell
end tell