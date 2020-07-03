use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

set degreeInfo to {}
tell application "Safari"
	activate
	tell document 1
		set elementCount to do JavaScript "document.getElementsByClassName('_lriijlm').length - 1"
		repeat with i from 0 to elementCount
			do JavaScript "document.getElementsByClassName('_lriijlm')[" & i & "].click()"
			delay 2
			set elementCount2 to do JavaScript "document.getElementsByClassName('_abk2w2g headerBox').length - 1"
			repeat with j from 0 to elementCount2
				set end of degreeInfo to ¬
					{do JavaScript "document.getElementsByClassName('_abk2w2g headerBox')[" & j & "].textContent", ¬
						do JavaScript "document.getElementsByClassName('font-sm m-b-2')[].textContent"
			end repeat
			do JavaScript "window.history.back()"
			delay 2
		end repeat
	end tell
end tell