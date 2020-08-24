use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This Script will download all the MP3s from a specifix user from house-mixes.com

tell application "Safari"
	activate
	make new document
	set downloadCount to 0
	tell document 1
		set URL to "https://www.house-mixes.com/profile/bartoszt/mediabywidget/1/latest"
		delay 120
		
		#get number of pages
		repeat with i from 0 to do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childElementCount"
			if (do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childNodes[" & i & "].textContent") contains "→" then
				set pageCount to do JavaScript "document.getElementsByClassName('pagination pagination-centered')[0].childNodes[7].previousElementSibling.textContent"
			end if
		end repeat
		
		repeat with i from 1 to pageCount # Loop through the pages
			# get number of songs on this page
			set numberOfSongsOnThisPage to do JavaScript "document.getElementsByClassName('audio-image-link').length"
			
			repeat with j from 0 to numberOfSongsOnThisPage - 1
				try
					set songName to words of (do JavaScript "document.getElementsByClassName('audio-image-link')[" & j & "].innerHTML")
				end try
				
				try
					do JavaScript "document.getElementsByClassName('audio-image-link')[" & j & "].click()" # click on song art				
				end try
				
				delay 3
				
				try
					do JavaScript "document.getElementsByClassName('btn btn-orange button-download')[0].click()" # first download button					
				end try
				
				delay 3
				
				try
					do JavaScript "document.getElementsByClassName('btn btn-primary')[1].click()" # second download button					
				end try
				
				delay 3
				
				try
					do JavaScript "document.getElementById('btnDownload').click()" # third and final download button					
					delay 3
				end try
				
				set downloadCount to downloadCount + 1
				
				repeat with k from 1 to count of songName
					if item k of songName is "alt" then
						set startRead to k
						exit repeat
					end if
				end repeat
				
				tell me to say "Clicked to download"
				
				repeat with k from startRead + 2 to (count of songName) - 1
					tell me to say item (k) of songName
				end repeat
				
				tell me to say "which is  page" & i & "song" & (j + 1)
				tell me to say "which is song number " & downloadCount
				do JavaScript "history.back()"
				
				delay 3
				
				tell application "System Events" to key code 125 using command down
				
			end repeat
			
			if i is less than pageCount then
				tell me to say "Going to the next page, which is page" & i + 1
			end if
			
			if i is less than pageCount then
				set the URL to "https://www.house-mixes.com/profile/bartoszt/mediabywidget/1/latest" & "/" & i + 1
			end if
			
			delay 5
			
			tell application "System Events" to key code 125 using command down
			
			delay 5
			
		end repeat
	end tell
end tell