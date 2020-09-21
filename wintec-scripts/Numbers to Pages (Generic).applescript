-- This script takes a spreadsheet and converts each row into a separate page in a Word Processing Document. (Uses Numbers and Pages to do this)

tell application "Numbers"
	set sourceFile to id of (open (choose file with prompt "Source File"))
	tell document id sourceFile
		tell sheet 1
			tell table 1
				set numberOfPages to (count of cells of column 1) - 2
				tell row 1
					set labels to {}
					repeat with i from 1 to count of cells
						if value of cell i is not missing value then
							set end of labels to value of cell i
						end if
					end repeat
					
				end tell
				set classColumnName to choose from list labels with prompt "Please tell me which field is the class name."
				set classColumnName to item 1 of classColumnName
				tell row 1
					repeat with i from 1 to count of columns
						if value of cell i contains classColumnName then
							set classColumn to column of cell 10
						end if
					end repeat
				end tell
				tell classColumn
					set classesWithDuplicates to value of cells
				end tell
				set classNames to my addr(classesWithDuplicates)
				set chosenClass to choose from list classNames with prompt "Please tell me which class you want to process."
				set chosenClass to item 1 of chosenClass
				set rowsToProcess to {}
				repeat with i from 1 to (count of cells of classColumn)
					if value of cell i of classColumn is chosenClass then
						set end of rowsToProcess to value of cells of row i
					end if
				end repeat
				
				
			end tell
		end tell
	end tell
end tell



tell application "Pages"
	activate
	make document
	tell document 1
		repeat with i from 1 to count of rowsToProcess
			if not (exists page i) then
				make page
			end if
			tell page i
				make table with properties {row count:count of labels, column count:2}
				tell table 1
					tell column 1
						repeat with j from 1 to count of labels
							set value of cell j to item j of labels
						end repeat
					end tell
					tell column 2
						repeat with j from 1 to count of labels
							set value of cell j to item j of item i of rowsToProcess
						end repeat
					end tell
				end tell
			end tell
		end repeat
	end tell
end tell





to addr(l)
	script foo
		property foo2 : l
		property okAddresses : {}
	end script
	
	considering case
		repeat with i from 1 to count (foo's foo2)
			set x to ((foo's foo2)'s item i)
			if x is in foo's okAddresses then
			else
				set end of foo's okAddresses to x
			end if
		end repeat
	end considering
	foo's okAddresses
end addr