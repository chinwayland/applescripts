-- This script will extract the first names from a padlet page

tell application "Safari"
	tell document 1
		set numberOfPosts to do JavaScript "document.getElementsByClassName('accessibility oob').length"
		set names to {}
		repeat with i from 0 to numberOfPosts - 1
			set end of names to do JavaScript "document.getElementsByClassName('accessibility oob')[" & i & "].textContent"
		end repeat
		set pageURL to URL
	end tell
end tell

set tid to AppleScript's text item delimiters
set AppleScript's text item delimiters to space
set L to names as text
set AppleScript's text item delimiters to tid
set FL to words of L
(*
repeat with i from 1 to count of FL
	if item i of FL contains "and" then
		replace_chars(item i of FL, "and", " ")
	end if
end repeat
*)

set splitNames to {}
repeat with i from 1 to count of FL
	if item i of FL is not "and" then
		set end of splitNames to item i of FL
	end if
end repeat

repeat with i from 1 to count of splitNames
	if item i of splitNames contains "and" then
		set end of splitNames to replace_chars(item i of splitNames, "and", " ")
	end if
end repeat

set splitNames2 to {}
repeat with i from 1 to count of splitNames
	if item i of splitNames does not contain "and" then
		set end of splitNames2 to item i of splitNames
	end if
end repeat

set tid to AppleScript's text item delimiters
set AppleScript's text item delimiters to space
set L to splitNames2 as text
set AppleScript's text item delimiters to tid
set splitNames3 to words of L

tell application "Numbers"
	activate
	tell document 1
		tell sheet 1
			tell table 1
				add column after last column
				repeat with i from 2 to count of cells of column 1
					if value of cell i of column 1 is in splitNames3 then
						set value of cell i of the last column to true
					end if
				end repeat
				set value of cell 1 of the last column to pageURL
			end tell
		end tell
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