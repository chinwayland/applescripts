-- This script scrapes all the Specializations from Coursera

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

set specializations to {}

tell application "Safari"
	activate
	set listOfDocs to name of documents
	set chosenWindow to choose from list listOfDocs with prompt "Which Safari window?"
	set chosenWindow to item 1 of chosenWindow
	set theResponse to display dialog "How many pages?" default answer "" with icon note buttons {"Cancel", "Continue"} default button "Continue"
	set numOfPages to text returned of theResponse as number
	tell document chosenWindow
		repeat with i from 1 to numOfPages
			do JavaScript "document.getElementById('pagination_number_box_" & i & "').click()"
			tell me to say "page" & i
			delay 10
			set countOfClass to do JavaScript "document.getElementsByClassName('color-primary-text card-title headline-1-text').length"
			repeat with j from 0 to countOfClass - 1
				set partner to do JavaScript "document.getElementsByClassName('card-info vertical-box')[" & j & "].childNodes[1].textContent"
				set subject to do JavaScript "document.getElementsByClassName('color-primary-text card-title headline-1-text')[" & j & "].textContent"
				set difficulty to do JavaScript "document.getElementsByClassName('difficulty')[" & j & "].textContent"
				set productType to do JavaScript "document.getElementsByClassName('_jen3vs _1d8rgfy3')[" & j & "].textContent"
				set end of specializations to {partner, subject, difficulty, productType}
			end repeat
		end repeat
	end tell
end tell

tell application "Numbers"
	activate
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				repeat with i from 1 to count of specializations
					if not (exists cell (i + 1) of column 1) then
						make new row
					end if
					set value of cell (i + 1) of column 1 to item 1 of item i of specializations
					set value of cell (i + 1) of column 2 to item 2 of item i of specializations
					set value of cell (i + 1) of column 3 to item 3 of item i of specializations
					set value of cell (i + 1) of column 4 to item 4 of item i of specializations
				end repeat
			end tell
		end tell
	end tell
end tell