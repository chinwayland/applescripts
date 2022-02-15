use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This script converts the official new york election calendar found here into events in Apple Calendar. 
-- The original PDF file is here: https://www.elections.ny.gov/NYSBOE/law/2022PoliticalCalendar.pdf
-- I used Adobe Acrobat to convert it to an excel spreadsheet, then opened it in Apple Numbers.
-- Then I used AppleScript to conver it to Apple Calendar
-- Once in Apple Calendar, you can export as an ics file or publish as a feed


set theMonths to {"Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"}
set theTable to {}
set theDates to {}

tell application "Numbers"
	set listOfDocuments to name of documents
	#	tell me to say "Choose which Numbers document has the election calendar"
	set chosenDocument to choose from list listOfDocuments
	set chosenDocument to item 1 of chosenDocument
	tell document chosenDocument
		tell sheet 1
			tell table 1
				set theTable to value of cells
			end tell
		end tell
	end tell
end tell

set theEvents to {}

repeat with i from 1 to count theMonths
	set theMonth to item i of theMonths
	repeat with j from 1 to count of theTable
		set theCell to item j of theTable
		if theCell is not missing value then
			if (count of theCell) < 13 then
				if theCell contains theMonth then
					set end of theEvents to {theCell, item (j + 1) of theTable}
				end if
			end if
		end if
	end repeat
end repeat

set theEvents2 to {}

repeat with theEvent in theEvents
	set theDate to item 1 of theEvent
	set theSummary to item 2 of theEvent
	if theDate is not equal to theSummary then
		set end of theEvents2 to {theDate, theSummary}
	end if
end repeat

set theEvents3 to {}

repeat with i from 1 to count of theEvents2
	set theEvent to item i of theEvents2
	try
		set nextEvent to item (i + 1) of theEvents2
	end try
	
	if theEvent is not equal to nextEvent then
		if item 1 of theEvent does not contain "January 18" then
			set end of theEvents3 to theEvent
		end if
	end if
end repeat

# Add year to the date string
set theEvents4 to {}

repeat with i from 1 to count of theEvents3
	set theEvent to item i of theEvents3
	set end of theEvents4 to {item 1 of theEvent & " 2022", item 2 of theEvent}
end repeat

set the theEvents5 to {}

repeat with i from 1 to count of theEvents4
	set theEvent to item i of theEvents4
	set theSummary to item 2 of theEvent
	set theDate to item 1 of theEvent
	if theDate contains "–" then
		set theDate to words of theDate
		set startDate to item 1 of theDate & " " & item 2 of theDate & " " & item 4 of theDate
		set startDate to date startDate
		set endDate to item 1 of theDate & space & item 3 of theDate & space & item 4 of theDate
		set endDate to date endDate
		set end of theEvents5 to {theSummary, startDate, endDate}
	else if theDate contains "th" then
		set theNumber to ""
		set theDay to second word of theDate
		set theCharacters to characters of theDay
		repeat with i from 1 to count of theCharacters
			set aCharacter to item i of theCharacters
			if aCharacter is not "t" then
				if aCharacter is not "h" then
					set theNumber to theNumber & aCharacter
				end if
			end if
		end repeat
		set theDate to word 1 of theDate & space & theNumber & space & word 3 of theDate
		set end of theEvents5 to {theSummary, date theDate, date theDate}
	else if theDate contains "Sept" then
		set theDate to replace_chars(theDate, "t", "")
		set startDate to theDate
		set endDate to theDate
		set end of theEvents5 to {theSummary, date startDate, date endDate}
		
	else
		set startDate to item 1 of theEvent
		set endDate to item 1 of theEvent
		set end of theEvents5 to {theSummary, date startDate, date endDate}
	end if
end repeat

set theEvents6 to {}
set theSummaryList to {}

repeat with i from 1 to count of items in theEvents5
	set theEvent to item i of theEvents5
	set theSummary to item 1 of theEvent
	if theSummaryList does not contain theSummary then
		set end of theEvents6 to theEvent
	end if
	set end of theSummaryList to theSummary
end repeat

tell application "Calendar"
	activate
	if not (exists (calendar "New York Elections")) then
		create calendar with name "New York Elections"
	end if
	tell calendar "New York Elections"
		repeat with aData in theEvents6
			set theSummary to item 1 of aData
			set startDate to item 2 of aData
			set endDate to item 3 of aData
			make new event with properties {summary:theSummary, start date:startDate, end date:endDate, allday event:true}
		end repeat
	end tell
end tell

on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars