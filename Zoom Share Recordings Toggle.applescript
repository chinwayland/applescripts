use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

# This script will toggle sharing of video recordings in Zoom
# First, Paste this script into Script Editor on a Mac
# Second, open up the list of video recordings in Zoom (Safari)
# Third, Press Play to run the script (Command-R)
# There is a bug where it counts
display dialog "Please make sure you have first opened up the list of video recordings in Zoom (Safari). Leave that as the frontmost window in Safari."

tell application "Safari"
	tell document 1
		set totalRecords to (do JavaScript "document.getElementsByName('totalRecords')[0].textContent") as integer
	end tell
end tell

set theResponse to display dialog "How many second to wait in between clicks? (Depends on speed of your internet connection and computer. Two seconds is a good start to try.)" default answer "2" with icon note buttons {"Cancel", "Continue"} default button "Continue"

set secondsToWait to (text returned of theResponse)
set numberOfPages to totalRecords / 15
set numberOfVideosOnLastPage to totalRecords mod 15
set numberOfTimesToLoop to round of numberOfPages rounding up
set videoCount to 0


repeat with pageCount from 1 to numberOfTimesToLoop
	if pageCount is not numberOfTimesToLoop then
		set numberOfVideosThisPage to 15 - 1
	else
		set numberOfVideosThisPage to numberOfVideosOnLastPage - 1
	end if
	tell application "Safari"
		activate
		tell document 1
			repeat with j from 0 to numberOfVideosThisPage
				set videoCount to videoCount + 1
				do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist')[" & j & "].click()" #share
				delay secondsToWait
				do JavaScript "document.getElementsByClassName('zm-switch__core')[0].click()" #toggle
				delay secondsToWait
				do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #done
				say "Video " & (videoCount as text) & " toggled"
			end repeat
		end tell
	end tell
	
	tell application "Safari" # go to next page
		tell document 1
			set nextButtonStatus to do JavaScript "document.getElementsByClassName('next disabled').length"
			if nextButtonStatus is 0 then
				do JavaScript "document.getElementsByClassName('next')[0].firstChild.click()"
			end if
		end tell
	end tell
	say "Going to the next page, which is page " & (pageCount + 1 as text)
end repeat

say "All videos have been processed."
say (videoCount as text) & "videos were toggled."