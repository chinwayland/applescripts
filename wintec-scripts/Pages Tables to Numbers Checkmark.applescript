-- This script checks a table in Pages if a form was filled out and then gives a checkmark in the corresponding Numbers spreadsheet

tell application "Pages"
	tell document 1
		set columnName to name
		set match to {}
		repeat with i from 1 to count of tables
			tell table i
				if the value of the last cell of the last column is not missing value then
					set end of match to value of fourth cell of the last column
				end if
			end tell
		end repeat
	end tell
end tell

tell application "Numbers"
	tell document 1
		tell sheet 1
			tell table 1
				add column after last column
				set the value of the first cell of the last column to columnName
				repeat with i from 1 to count of rows
					if value of cell i of column "Student Number" is in match then
						set the value of cell i of the last column to true
					end if
				end repeat
			end tell
		end tell
	end tell
end tell