use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

#set theString to getInputByClass("control-label", 0) # get total number of workers
#set numberOfWorkers to the last item of returnNumbersInString(theString)

#clickClassName("btn btn-default btn-block btn-default ng-binding", 0) # 20 more workers button

set theResult to {}
tell application "Safari"
	tell document 1
		repeat with i from 0 to 19
			set the end of theResult to do JavaScript "document.getElementsByClassName('thumbnail')[" & i & "].childNodes[2].outerHTML"
		end repeat
	end tell
end tell

set theURL to {}
repeat with result in theResult
	set the end of theURL to item 4 of splitText(result, "\"")
end repeat

set girlURL to {}
repeat with i in theURL
	set the end of girlURL to "https://www.gaito.vip" & i
end repeat



(*
set theURL to {}
repeat with result in theResult
	if result contains 32370 then
		set the end of theURL to result
	end if
end repeat
*)

to getInputByClass(theClass, num) -- defines a function with two inputs, theClass and num
	tell application "Safari" --tells AS that we are going to use Safari
		set input to do JavaScript "document.getElementsByClassName('" & theClass & "')[" & num & ¬
			"].innerHTML;" in document 1 -- uses JavaScript to set the variable input to the information we want
	end tell
	return input --tells the function to return the value of the variable input
end getInputByClass

to getInputByTag(theTag, num) -- defines a function with two inputs, theTag and num
	tell application "Safari" --tells AS that we are going to use Safari
		set input to do JavaScript "document.getElementsByTagName('" & theTag & "')[" & num & ¬
			"].innerHTML;" in document 1
	end tell
	return input
end getInputByTag

on returnNumbersInString(inputString)
	set s to quoted form of inputString
	do shell script "sed s/[a-zA-Z\\']//g <<< " & s
	set dx to the result
	set numlist to {}
	repeat with i from 1 to count of words in dx
		set this_item to word i of dx
		try
			set this_item to this_item as number
			set the end of numlist to this_item
		end try
	end repeat
	return numlist
end returnNumbersInString

to clickTagName(thetagName, elementnum)
	tell application "Safari"
		do JavaScript "document.getElementsByTagName('" & thetagName & "')[" & elementnum & "].click();" in document 1
	end tell
end clickTagName

to clickClassName(theClassName, elementnum)
	tell application "Safari"
		do JavaScript "document.getElementsByClassName('" & theClassName & "')[" & elementnum & "].click();" in document 1
	end tell
end clickClassName

on splitText(theText, theDelimiter)
	set AppleScript's text item delimiters to theDelimiter
	set theTextItems to every text item of theText
	set AppleScript's text item delimiters to ""
	return theTextItems
end splitText