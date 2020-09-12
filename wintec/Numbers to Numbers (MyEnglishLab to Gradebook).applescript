-- This script will check the exported gradebook from MyEnglishLav and check if the student has done the assignments. They will get a checkmark for each assignment done. The MyEnglishLab Username has to be in the gradebook for this to work.
tell application "Numbers"
	
	set sourceFile to id of (open (choose file with prompt "Source File"))
	set targetFile to id of (open (choose file with prompt "Target File"))
	
	tell document id sourceFile
		tell me to say "grabbing data"
		tell sheet 3
			tell table 1
				set unitName to value of cell "C2"
				set reducedUnitName to characters 1 through (offset of ":" in unitName) of unitName
				set reducedUnitName to reducedUnitName as string
				set scores to {}
				set skip to {"---"}
				set activity to {}
				repeat with i from 3 to (((count of columns) - 4)) by 3 # loop through every third column
					set end of activity to reducedUnitName & " " & value of cell 4 of column i # add column name to variable
					repeat with j from 6 to (count of cells of column 1) - 1 # loop through every row in each column
						if value of cell i of row j is not in skip then # if score not blank then add activity name, username, and score to list
							set end of scores to {reducedUnitName & " " & value of cell 4 of column i, value of cell 2 of row j, value of cell i of row j}
						end if
					end repeat
				end repeat
			end tell
		end tell
	end tell
	
	tell document id targetFile
		tell me to say "pasting data"
		tell sheet 1
			tell table 1
				repeat with i from 1 to count of activity
					add column after the last column
					set the value of the first cell of the last column to item i of activity # set column name
				end repeat
				repeat with i from 1 to count of scores
					tell column 11
						set rowAddress to address of row of (first cell whose value of it contains item 2 of item i of scores) # get row name of username
					end tell
					tell row 1
						set columnAddress to address of column of (first cell whose value of it contains item 1 of item i of scores) # get column name of activity
					end tell
					tell column columnAddress
						tell cell rowAddress
							set value to true
						end tell
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell