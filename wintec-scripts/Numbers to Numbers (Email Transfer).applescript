-- This script transfers emails from Moodle export to gradebook.
-- To do - Transfer MEL username

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Numbers"
	activate
	set listOfDocs to name of documents
	set sourceDoc to choose from list listOfDocs with prompt "Choose Source Doc"
	set targetDoc to choose from list listOfDocs with prompt "Choose Target Doc"
	set sourceDoc to item 1 of sourceDoc
	set targetDoc to item 1 of targetDoc
	
	tell document sourceDoc
		tell sheet 1
			tell table 1
				set emailsSource to column "C"
				set usernamesSource to column "D"
			end tell
		end tell
	end tell
	
	tell document targetDoc
		tell sheet 1
			tell table 1
				set emailsTarget to column "I"
				set usernamesTarget to column "H"
			end tell
		end tell
	end tell
	
	repeat with i from 2 to count of cells of usernamesSource
		set usernameSource to value of cell i of usernamesSource
		set email to value of cell i of emailsSource
		repeat with j from 2 to count of cells of usernamesTarget
			set usernameTarget to value of cell j of usernamesTarget
			if usernameSource is equal to usernameTarget then
				set value of cell j of emailsTarget to email
			end if
		end repeat
	end repeat
	
end tell