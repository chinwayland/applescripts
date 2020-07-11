use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Safari"
	tell document 1
		
		(*
		repeat with i from 1 to 4
			*)
		
		repeat with j from 0 to 9
			do JavaScript "document.getElementsByClassName('img-for-audio img-responsive unveil-loaded')[" & j & "].click()" # Song Art
			delay 30
			do JavaScript "document.getElementsByClassName('btn btn-white btn-orange button-download')[0].click()" # first download button
			delay 30
			do JavaScript "document.getElementsByClassName('btn btn-primary')[1].click()" # second download button
			delay 30
			do JavaScript "document.getElementsByClassName('btn btn-orange btn-block')[0].click()" # third download button
			delay 30
			do JavaScript "window.history.back()"
			delay 30
		end repeat
		(*
			if i is not 4 then
				do JavaScript "document.getElementsByClassName('active')[0].nextSibling.firstChild.click()"
				delay 30
			end if

		end repeat
		*)
	end tell
end tell