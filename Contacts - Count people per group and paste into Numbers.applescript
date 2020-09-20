-- This script counts the number of people per group and pastes the data in a Numbers spreadsheet

set groupNameAndCount to {}
tell application "Contacts"
	repeat with i from 1 to count of groups
		tell group i
			set end of groupNameAndCount to {name, count of people}
		end tell
	end repeat
end tell

tell application "Numbers"
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				repeat with i from 1 to count of groupNameAndCount
					if not (exists row (i + 1)) then
						add row below last row
					end if
					tell row (i + 1)
						set value of cell 1 to item 1 of item i of groupNameAndCount
						set value of cell 2 to item 2 of item i of groupNameAndCount
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell