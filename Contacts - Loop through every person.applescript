# This script loops through every person in the Contacts apps, it then loops through every phone of each person.

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

set defaultPhoneLabels to {"home", "work", "school", "iPhone", "mobile", "main", "home fax", "work fax", "pager", "other"}
set unwantedPhoneLabels to {}
set failedPersons to {}
set failedPhones to {}
tell application "Contacts"
	repeat with i from 1 to count of people
		tell person i
			try
				repeat with j from 1 to count of phones
					tell phone j
						try
							if label is not in defaultPhoneLabels then
								set end of unwantedPhoneLabels to {i, j, label}
							end if
						on error
							set end of failedPhones to {i, j}
						end try
					end tell
				end repeat
			on error
				set end of failedPersons to name
			end try
		end tell
	end repeat
end tell