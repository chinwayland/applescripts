use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

set dictionary to {} #create dictionary for distance
set j to 0
repeat with i from 0 to 19
	set end of dictionary to 50 + j
	set j to j + 5
end repeat


repeat with i from 0 to 19
	
	set distance to getInputByClass("ng-binding", item (i + 1) of dictionary)
	tell application "Safari"
		tell document 1
			activate
			do JavaScript "document.getElementsByTagName('img')[" & i & "].click()" # click on girl photo
			delay 2
			set theURL to URL
		end tell
	end tell
	set workerName to getInputByClass("ng-binding", 39)
	set price to getInputByClass("ng-binding", 43)
	set price to findAndReplaceInText(price, "&nbsp;k", "")
	set phone to getInputByClass("ng-binding", 44)
	set addressText to getInputByClass("ng-binding", 46)
	set summary to getInputByClass("ng-binding ng-scope", 1)
	set yearBorn to getInputByClass("ng-binding ng-scope", 3)
	try
		set age to (year of (current date)) - yearBorn as number
	end try
	set workerheight to getInputByClass("ng-binding ng-scope", 4)
	set weight to getInputByClass("ng-binding ng-scope", 5)
	set measurement1 to getInputByClass("ng-binding ng-scope", 6)
	set measurement2 to getInputByClass("ng-binding ng-scope", 7)
	set measurement3 to getInputByClass("ng-binding ng-scope", 8)
	set origin to getInputByClass("ng-binding ng-scope", 9)
	set hair to getInputByClass("ng-binding ng-scope", 10)
	set face to getInputByClass("ng-binding ng-scope", 11)
	set extraData to {}
	repeat with j from 12 to 22
		set extraData to extraData & getInputByClass("ng-binding ng-scope", j)
	end repeat
	
	tell application "Safari"
		activate
		do JavaScript "history.back()" in document 1
	end tell
	
	tell application "Numbers"
		activate
		tell document 1
			tell sheet 1
				tell table 1
					set value of cell 1 of row (i + 2) to workerName
					set value of cell 2 of row (i + 2) to distance
					set value of cell 3 of row (i + 2) to addressText
					set value of cell 4 of row (i + 2) to phone
					set value of cell 5 of row (i + 2) to price
					set value of cell 6 of row (i + 2) to age
					set value of cell 7 of row (i + 2) to theURL
					set value of cell 8 of row (i + 2) to summary
					try
						repeat with j from 1 to 12 #loop through extraData
							set value of cell (9) of row (i + 2) to value of cell (9) of row (i + 2) & item j of extraData as string
						end repeat
					end try
					set value of cell 10 of row (i + 2) to workerheight
					set value of cell 11 of row (i + 2) to weight
					set value of cell 12 of row (i + 2) to measurement1
					set value of cell 13 of row (i + 2) to measurement2
					set value of cell 14 of row (i + 2) to measurement3
					set value of cell 15 of row (i + 2) to origin
					set value of cell 16 of row (i + 2) to hair
					set value of cell 17 of row (i + 2) to face
				end tell
			end tell
		end tell
		
	end tell
end repeat







to getInputByClass(theClass, num) -- defines a function with two inputs, theClass and num
	tell application "Safari" --tells AS that we are going to use Safari
		set input to do JavaScript "document.getElementsByClassName('" & theClass & "')[" & num & ¬
			"].innerHTML;" in document 1 -- uses JavaScript to set the variable input to the information we want
	end tell
	return input --tells the function to return the value of the variable input
end getInputByClass

on findAndReplaceInText(theText, theSearchString, theReplacementString)
	set AppleScript's text item delimiters to theSearchString
	set theTextItems to every text item of theText
	set AppleScript's text item delimiters to theReplacementString
	set theText to theTextItems as string
	set AppleScript's text item delimiters to ""
	return theText
end findAndReplaceInText

to getInputByTag(theTag, num) -- defines a function with two inputs, theTag and num
	tell application "Safari" --tells AS that we are going to use Safari
		set input to do JavaScript "document.getElementsByTagName('" & theTag & "')[" & num & ¬
			"].innerHTML;" in document 1
	end tell
	return input
end getInputByTag