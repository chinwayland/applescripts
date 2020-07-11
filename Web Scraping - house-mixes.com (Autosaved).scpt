use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Safari"
	activate
	set downloadCount to 0
	tell document 1
		#get number of pages
		repeat with i from 0 to do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childElementCount"
			if (do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childNodes[" & i & "].textContent") contains "→" then
				set pageCount to do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childNodes[7].previousElementSibling.textContent"
			end if
		end repeat
		
		repeat with i from 3 to pageCount # Loop through the pages
			# get number of songs on this page
			set numberOfSongsOnThisPage to do JavaScript "document.getElementsByClassName('audio-image-link').length"
			repeat with j from 0 to numberOfSongsOnThisPage
				do JavaScript "document.getElementsByClassName('audio-image-link')[" & j & "].click()" # click on song art
				
				repeat until source contains "FAQ" #wait until page is loaded
					delay 1
				end repeat
				
				do JavaScript "document.getElementsByClassName('btn btn-orange button-download')[0].click()" # first download button
				delay 2
				do JavaScript "document.getElementsByClassName('btn btn-primary')[1].click()" # second download button
				
				repeat until source contains "FAQ" #wait until page is loaded
					delay 1
				end repeat
				
				do JavaScript "document.getElementsByClassName('btn btn-orange btn-block')[0].click()" # third and final download button
				
				delay 2
				
				set downloadCount to downloadCount + 1
				tell me to say "Clicked to download page" & i & "song" & j + 1
				do JavaScript "history.back()"
				
				repeat until source contains "FAQ" #wait until page is loaded
					delay 1
				end repeat
				
			end repeat
			
			if i is not pageCount then # go to the next page
				do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childNodes[7].firstChild.click()" # the click
				repeat until source contains "FAQ"
					delay 1
				end repeat
				tell me to say "Going to the next page, which is page" & i
			end if
		end repeat
	end tell
end tell