use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Numbers"
	activate
	set listOfDocuments to name of documents
	set chosenDocument to choose from list listOfDocuments
	set chosenDocument to item 1 of chosenDocument
	tell document chosenDocument
		#		set listOfSheets to name of sheets
		set listOfSheets to name of sheets
		set chosenSheet to choose from list listOfSheets
		set chosenSheet to item 1 of chosenSheet
		tell sheet chosenSheet
			tell table 1
				set listOfColumns to name of columns
				set chosenColumn to choose from list listOfColumns
				set chosenColumn to item 1 of chosenColumn
				tell column chosenColumn
					set cellValues to {}
					repeat with i from 2 to count of cells
						set end of cellValues to value of cell i
					end repeat
				end tell
			end tell
		end tell
	end tell
end tell

tell application "Microsoft Excel"
	activate
	set listOfDocuments to name of workbooks
	set chosenDocument to choose from list listOfDocuments
	set chosenDocument to item 1 of chosenDocument
	tell workbook chosenDocument
		set listOfSheets to name of sheets
		set chosenSheet to choose from list listOfSheets
		set chosenSheet to item 1 of chosenSheet
		tell sheet chosenSheet
			set theResponse to display dialog "which is the first cell?" default answer "" with icon note buttons {"Cancel", "Continue"} default button "Continue"
			--> {button returned:"Continue", text returned:"Jen"}
			display dialog "You entered: " & (text returned of theResponse) & "."
			#column of cell (text returned of theResponse)
			set columnIndex to first column index of cell (text returned of theResponse)
			set rowIndex to first row index of cell (text returned of theResponse)
			tell column columnIndex
				repeat with i from 1 to count of cellValues
					if item i of cellValues is true then
						set value of cell (rowIndex - 1 + i) to "P"
					else if item i of cellValues is false then
						set value of cell (rowIndex - 1 + i) to "A"
					else if item i of cellValues is "✅" then
						set value of cell (rowIndex - 1 + i) to "P"
					else if item i of cellValues is "📝" then
						set value of cell (rowIndex - 1 + i) to "E"
					else if item i of cellValues is "⏳" then
						set value of cell (rowIndex - 1 + i) to "L"
					else if item i of cellValues is "❌" then
						set value of cell (rowIndex - 1 + i) to "A"
					end if
				end repeat
			end tell
		end tell
	end tell
end tell