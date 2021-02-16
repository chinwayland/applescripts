-- This script scrapes all the college degrees from Coursera.

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Safari"
	activate
	tell document 1
		set numOfSchools to do JavaScript "document.getElementsByClassName('font-md font-weight-bold d-block').length"
		set university to {}
		repeat with i from 0 to numOfSchools - 1
			set degreeType to do JavaScript "document.getElementsByClassName('font-md font-weight-bold d-block')[" & i & "].textContent"
			set universityName to do JavaScript "document.getElementsByClassName('font-sm m-b-0 d-block')[" & i & "].textContent"
			set duration to do JavaScript "document.getElementsByClassName('program-details font-weight-bold m-t-1s m-b-0')[" & i & "].textContent"
			set ranking to do JavaScript "document.getElementsByClassName('ranking-description font-sm d-block m-t-1s')[" & i & "].textContent"
			set end of university to {degreeType, universityName, duration, ranking}
		end repeat
	end tell
end tell

tell application "Numbers"
	activate
	tell document 1
		tell sheet 1
			tell table 1
				repeat with i from 1 to count of university
					set value of cell (i + 1) of column 1 to item 1 of item i of university
					set value of cell (i + 1) of column 2 to item 2 of item i of university
					set value of cell (i + 1) of column 3 to item 3 of item i of university
					set value of cell (i + 1) of column 4 to item 4 of item i of university
				end repeat
			end tell
		end tell
	end tell
end tell
