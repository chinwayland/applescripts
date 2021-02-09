use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Safari"
	activate
	tell document 1
		set numOfCerts to do JavaScript "document.getElementsByClassName('m-y-2').length"
		set certificate to {}
		set j to 0
		repeat with i from 0 to numOfCerts - 1
			set certificateName to do JavaScript "document.getElementsByClassName('m-y-2')[" & i & "].firstChild.textContent"
			set universityName to do JavaScript "document.getElementsByClassName('font-sm m-b-0 d-block')[" & i & "].firstChild.textContent"
			set ranking to do JavaScript "document.getElementsByClassName('text-uppercase font-sm d-block')[" & i & "].nextSibling.textContent"
			set cost to do JavaScript "document.getElementsByClassName('_xliqh9g')[" & i + 1 + j & "].firstChild.childNodes[1].textContent"
			set duration to do JavaScript "document.getElementsByClassName('_xliqh9g')[" & i + 1 + j & "].childNodes[1].childNodes[1].textContent"
			set requiredCourses to do JavaScript "document.getElementsByClassName('_xliqh9g')[" & i + 1 + j & "].childNodes[2].childNodes[1].textContent"
			# set realWorldProjects to do JavaScript "document.getElementsByClassName('_xliqh9g')[" & i + 1 + j & "].childNodes[3].childNodes[1].textContent"
			set j to j + 1
			set end of certificate to {certificateName, universityName, ranking, cost, duration, requiredCourses}
		end repeat
	end tell
end tell

tell application "Numbers"
	activate
	tell document 1
		tell sheet 2
			tell table 1
				repeat with i from 1 to count of certificate
					if not (exists cell (i + 1) of column 1) then
						make row
					end if
					set value of cell (i + 1) of column 1 to item 1 of item i of certificate
					set value of cell (i + 1) of column 2 to item 2 of item i of certificate
					set value of cell (i + 1) of column 3 to item 3 of item i of certificate
					set value of cell (i + 1) of column 4 to item 4 of item i of certificate
					set value of cell (i + 1) of column 5 to item 5 of item i of certificate
					set value of cell (i + 1) of column 6 to item 6 of item i of certificate
				end repeat
			end tell
		end tell
	end tell
end tell