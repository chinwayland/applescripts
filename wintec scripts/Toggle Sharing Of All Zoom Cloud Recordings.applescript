--This script will open the list of Zoom cloud recordings in Safari and then turn them all on or off, depending on your choice

set videoShareGoalList to {"On", "Off"}
set videoShareGoal to choose from list videoShareGoalList with prompt "Do you want the videos On (Shared) or Off (Unshared)?" default items {"Off"}
set videoShareGoal to item 1 of videoShareGoal

set theResponse to display dialog "How many second to wait in between clicks? (Depends on speed of your internet connection and computer. Two seconds is a good start to try.)" default answer "2" with icon note buttons {"Cancel", "Continue"} default button "Continue"
set secondsToWait to (text returned of theResponse) as number

set videoCount to 0

# opens new window and goes to zoom recordings
tell application "Safari"
	activate
	make new document
	set URL of document 1 to "https://zoom.us/recording"
	delay secondsToWait
end tell

# User checks if zoom is logged in
display dialog "Switch to Safari and check if Zoom is logged in and if your Zoom Recordings page has loaded."

# Check for the total number of records/videos
tell application "Safari"
	tell document 1
		set totalRecords to (do JavaScript "document.getElementsByName('totalRecords')[0].textContent") as integer
	end tell
end tell

# Calculate number of pages
set numberOfPages to totalRecords / 15

# Calculate number of pages to loop
set numberOfTimesToLoop to round of numberOfPages rounding up

repeat with pageCount from 1 to numberOfTimesToLoop
	
	#check for  number of videos on this page
	tell application "Safari"
		activate
		tell document 1
			set numberOfVideosOnThisPage to do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist').length"
		end tell
	end tell
	
	# Loop through each video displayed on this page
	repeat with i from 0 to numberOfVideosOnThisPage - 1
		tell application "Safari" # check if share is on or off
			tell document 1
				do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist')[" & i & "].click()" #click share button
				delay secondsToWait
				set videoCount to videoCount + 1
				if (do JavaScript "document.getElementsByClassName('zm-switch').length") is greater than 1 then #check if share toggle is on or off. Greater than 1 = On
					# Share Toggle is On
					if videoShareGoal is "On" then
						# Do Nothing and click Done
						do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #click done button
						say "Video " & videoCount & " is " & videoShareGoal
					else
						if videoShareGoal is "Off" then
							# Click the toggle button and click done
							do JavaScript "document.getElementsByClassName('zm-switch__core')[0].click()" #click On/Off Button
							delay secondsToWait
							do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #click done button
							say "Video " & videoCount & " is " & videoShareGoal
						end if
					end if
				else
					# Share Toggle is Off, Less than 5 = Off
					if videoShareGoal is "On" then
						# Click the toggle button and click done
						do JavaScript "document.getElementsByClassName('zm-switch__core')[0].click()" #click On/Off Button
						delay secondsToWait
						do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #click done button
						say "Video " & videoCount & " is " & videoShareGoal
					else
						if videoShareGoal is "Off" then
							# Do Nothing and click Done
							do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #click done button
							say "Video " & videoCount & " is " & videoShareGoal
						end if
					end if
				end if
			end tell
		end tell
	end repeat
	
	tell application "Safari" #check if next button is clickable, if clickable, then click it.
		tell document 1
			set nextButtonStatus to do JavaScript "document.getElementsByClassName('next disabled').length"
			if nextButtonStatus is 0 then
				do JavaScript "document.getElementsByClassName('next')[0].firstChild.click()" # the click
				if pageCount is not numberOfTimesToLoop then
					say "Going to the next page, which is page " & pageCount + 1
				end if
			end if
		end tell
	end tell
end repeat

say "All videos have been processed"
say (videoCount as text) & " videos were turned " & videoShareGoal
