use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions


set theResponse to display dialog "How many videos are there in total?" default answer "" with icon note buttons {"Cancel", "Continue"} default button "Continue"

set theResponse2 to display dialog "How many second to wait in between clicks? (Depend on speed of your internet connection)" default answer "" with icon note buttons {"Cancel", "Continue"} default button "Continue"

set secondsToWait to (text returned of theResponse2)
set numberOfPages to (text returned of theResponse) / 15
set numberOfVideosOnLastPage to 15 - (text returned of theResponse) mod 2
set numberOfTimesToLoop to round of numberOfPages rounding up

repeat with i from 1 to numberOfTimesToLoop
	if i is not numberOfTimesToLoop then
		set numberOfVideos to 15 - 1
	else
		set numberOfVideos to numberOfVideosOnLastPage - 1
	end if
	tell application "Safari"
		activate
		tell document 1
			repeat with j from 0 to 14
				do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist')[" & j & "].click()" #share
				delay secondsToWait
				do JavaScript "document.getElementsByClassName('zm-switch__core')[0].click()" #toggle
				delay secondsToWait
				do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #done
				delay secondsToWait
			end repeat
		end tell
	end tell
	
end repeat

(*
tell application "Safari"
	activate
	tell document 1
		repeat with i from 0 to 14
			do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist')[" & i & "].click()" #share
			delay 2
			do JavaScript "document.getElementsByClassName('zm-switch__core')[0].click()" #toggle
			delay 2
			do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #done
			delay 2
		end repeat
	end tell
end tell
*)