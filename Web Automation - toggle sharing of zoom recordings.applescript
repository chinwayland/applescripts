# This script does not work, it's got bugs, use the other one
# First, Paste this script into Script Editor on a Mac
# Second, open up the list of video recordings in Zoom (Safari)
# Third, Press Play to run the script (Command-R)

set videoShareGoalList to {"On", "Off"}
set videoShareGoal to choose from list videoShareGoalList with prompt "Do you want the videos On or Off?" default items {"Off"}

set theResponse to display dialog "How many second to wait in between clicks? (Depends on speed of your internet connection and computer. Two seconds is a good start to try.)" default answer "2" with icon note buttons {"Cancel", "Continue"} default button "Continue"
set secondsToWait to (text returned of theResponse)

set videoCount to 0

tell application "Safari" #check if next button is clickable, 0 means clickable
	activate
	tell document 1
		set nextButtonStatus to do JavaScript "document.getElementsByClassName('next disabled').length"
	end tell
end tell

repeat while nextButtonStatus is 0 # repeat while next button is clickable
	
	tell application "Safari" #check for  number of videos
		tell document 1
			set numberOfVideosOnThisPage to do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist').length"
			set nextButtonStatus to 1
		end tell
	end tell
	
	repeat with i from 0 to numberOfVideosOnThisPage - 1 # repeat with number of videos on page
		
		tell application "Safari" # check if share is on or off
			tell document 1
				do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist')[" & i & "].click()" #click share button
				delay secondsToWait
				if (do JavaScript "document.getElementsByClassName('zm-switch').length") is 1 then #check if share button is on or off
					set shareButton to false
				else
					set shareButton to true
				end if
				do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #click done button
				delay secondsToWait
				
				if videoShareGoal is "On" then
					if shareButton is true then
						do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist')[" & i & "].click()" #click share button
						delay secondsToWait
						do JavaScript "document.getElementsByClassName('zm-switch__core')[0].click()" #click On/Off Button
						delay secondsToWait
						do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #click done button
						set videoCount to videoCount + 1
						say "Video " & videoCount & videoShareGoal
						
					else
						if shareButton is false then
							# do nothing
							
						end if
					end if
				else
					if videoShareGoal is "Off" then
						if shareButton is true then
							#do nothing
							
						else
							if shareButton is false then
								do JavaScript "document.getElementsByClassName('btn btn-default sharemeet_from_myrecordinglist')[" & i & "].click()" #click share button
								delay secondsToWait
								do JavaScript "document.getElementsByClassName('zm-switch__core')[0].click()" #click On/Off Button
								delay secondsToWait
								do JavaScript "document.getElementsByClassName('zm-button__slot')[0].click()" #click done button
								set videoCount to videoCount + 1
								say "Video " & videoCount & videoShareGoal
								
							end if
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
				do JavaScript "document.getElementsByClassName('next')[0].firstChild.click()"
				delay secondsToWait
			end if
		end tell
	end tell
	
end repeat

say "All videos have been processed"
say (videoCount as text) & "videos were turned" & videoShareGoal