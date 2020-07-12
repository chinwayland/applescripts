use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

# this script is buggy because the website is buggy
tell application "Safari"
	activate
	set downloadCount to 0
	tell document 1
		set URL to "https://www.house-mixes.com/profile/bartoszt/mediabywidget/1/latest"
		
		delay 10
		
		
		
		#get number of pages
		repeat with i from 0 to do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childElementCount"
			if (do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childNodes[" & i & "].textContent") contains "→" then
				set pageCount to do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childNodes[7].previousElementSibling.textContent"
			end if
		end repeat
		
		repeat with i from 1 to pageCount # Loop through the pages
			# get number of songs on this page
			set numberOfSongsOnThisPage to do JavaScript "document.getElementsByClassName('audio-image-link').length"
			
			repeat with j from 0 to numberOfSongsOnThisPage
				do JavaScript "document.getElementsByClassName('audio-image-link')[" & j & "].click()" # click on song art
				delay 10
				
				do JavaScript "document.getElementsByClassName('btn btn-orange button-download')[0].click()" # first download button
				delay 2
				do JavaScript "document.getElementsByClassName('btn btn-primary')[1].click()" # second download button
				
				delay 10
				
				do JavaScript "document.getElementById('btnDownload').click()" # third and final download button
				
				delay 2
				
				set downloadCount to downloadCount + 1
				
				tell me to say "Clicked to download page" & i & "song" & (j + 1)
				do JavaScript "history.back()"
				
				delay 10
				
			end repeat
			
			# go to the next page
			if i is less than pageCount then
				set the URL to "https://www.house-mixes.com/profile/bartoszt/mediabywidget/1/latest" & "/" & i + 1
			end if
			
			delay 10
			
			if i is less than pageCount then
				tell me to say "Going to the next page, which is page" & i + 1
			end if
			
		end repeat
	end tell
end tell