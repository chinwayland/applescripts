use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

# This script grabs all the data about degree from coursera and puts them into a Numbers spreadsheet

set degreeName to {}
set partnerName to {}
set tuition to {}
set lengthOfCourse to {}
set hoursPerWeek to {}
set numberOfCourses to {}

tell application "Safari"
	make new document
	tell document 1
		set URL to "https://www.coursera.org/degrees/"
	end tell
end tell

tell application "Numbers"
	make new document
end tell

display dialog "Page loaded yet?"

repeat with i from 0 to 20
	tell application "Safari"
		activate
		tell document 1
			do JavaScript "document.getElementsByClassName('_ftgvbcw')[" & i & "].firstChild.firstChild.click()"
			delay 5
			set end of degreeName to do JavaScript "document.getElementsByClassName('page-title')[0].textContent"
			set end of partnerName to do JavaScript "document.getElementsByClassName('partner-name')[0].textContent"
			set end of tuition to do JavaScript "document.getElementsByClassName('feature-title')[0].textContent"
			set end of lengthOfCourse to do JavaScript "document.getElementsByClassName('feature-title')[1].textContent"
			set end of hoursPerWeek to do JavaScript "document.getElementsByClassName('feature-description')[1].childNodes[1].firstChild.textContent"
			set end of numberOfCourses to do JavaScript "document.getElementsByClassName('feature-title')[2].textContent"
			say "Collected information from school " & i + 1
			do JavaScript "window.history.back()"
			delay 5
		end tell
	end tell
	
	tell application "Numbers"
		activate
		tell document 1
			tell sheet 1
				tell table 1
					tell row (i + 2)
						set value of cell 1 to item (i + 1) of degreeName
						set value of cell 2 to item (i + 1) of partnerName
						set value of cell 3 to item (i + 1) of tuition
						set value of cell 4 to item (i + 1) of lengthOfCourse
						set value of cell 5 to item (i + 1) of hoursPerWeek
						set value of cell 6 to item (i + 1) of numberOfCourses
					end tell
				end tell
			end tell
		end tell
	end tell
end repeat
