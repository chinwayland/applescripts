-- This script checks if a number from a Numbers spreadsheet exists in another spreadsheet and makes a checkmark in a column of the second spreadsheet.

set input to display dialog "What is the new column name?" default answer ""
set columnName to input's text returned

#get row data into memory, row data needs to be split
tell application "Numbers"
	tell document 1
		tell sheet 1
			tell table 1
				set columnData to column 1's cell's value
			end tell
		end tell
	end tell
end tell

#split row data into individual items, 4 items total, item 3 and item 4 are useful
set splitData to {}
repeat with rowData in columnData
	set end of splitData to split(rowData, "!")
end repeat


#Search Numbers document 2 for Student ID Number and add the score to the last column

tell application "Numbers"
	tell document 2
		repeat with i from 1 to 4
			tell sheet i
				tell table 1
					add column after last column
					set last column's format to number
					set last column's first cell's value to columnName
					tell column "F"
						repeat with j in splitData
							if (first cell whose value is j's item 3 as number) exists then
								set studentID to (first cell whose value is j's item 3 as number)
								set row (get address of row of studentID)'s last cell's value to "=" & j's item 4
								set answer to row (get address of row of studentID)'s last cell's value
								set answer to answer * 100
								set row (get address of row of studentID)'s last cell's value to answer
							end if
						end repeat
					end tell
				end tell
			end tell
		end repeat
	end tell
end tell

# set missing scores to zero
tell application "Numbers"
	tell document 2
		repeat with i from 1 to 4
			tell sheet i
				tell table 1
					tell the last column
						repeat with j from 1 to count of cells
							tell cell j
								if value is missing value then
									set value to 0
								end if
							end tell
						end repeat
					end tell
				end tell
			end tell
		end repeat
	end tell
end tell







on convertListToString(theList, theDelimiter)
	set AppleScript's text item delimiters to theDelimiter
	set theString to theList as string
	set AppleScript's text item delimiters to ""
	return theString
end convertListToString

on getPositionOfItemInList(theItem, theList)
	repeat with a from 1 to count of theList
		if item a of theList is theItem then return a
	end repeat
	return 0
end getPositionOfItemInList

on getPositionsOfItemInList(theItem, theList, listFirstPositionOnly)
	set thePositions to {}
	repeat with a from 1 to length of theList
		if item a of theList is theItem then
			if listFirstPositionOnly = true then return a
			set end of thePositions to a
		end if
	end repeat
	if listFirstPositionOnly is true and thePositions = {} then return 0
	return thePositions
end getPositionsOfItemInList

to split(someText, delimiter)
	set AppleScript's text item delimiters to delimiter
	set someText to someText's text items
	set AppleScript's text item delimiters to {""} --> restore delimiters to default value
	return someText
end split

on trimText(theText, theCharactersToTrim, theTrimDirection)
	set theTrimLength to length of theCharactersToTrim
	if theTrimDirection is in {"beginning", "both"} then
		repeat while theText begins with theCharactersToTrim
			try
				set theText to characters (theTrimLength + 1) thru -1 of theText as string
			on error
				-- text contains nothing but trim characters
				return ""
			end try
		end repeat
	end if
	if theTrimDirection is in {"end", "both"} then
		repeat while theText ends with theCharactersToTrim
			try
				set theText to characters 1 thru -(theTrimLength + 1) of theText as string
			on error
				-- text contains nothing but trim characters
				return ""
			end try
		end repeat
	end if
	return theText
end trimText

