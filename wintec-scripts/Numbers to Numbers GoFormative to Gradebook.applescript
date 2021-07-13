use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Numbers"
	set documentList to name of documents
	choose from list documentList with prompt "What is the Source Document?"
	set sourceDocument to item 1 of result
	set assignmentTitle to word 3 of sourceDocument & " " & word 4 of sourceDocument & " " & word 5 of sourceDocument & " " & word 6 of sourceDocument
	choose from list documentList with prompt "What is the Target Document?"
	set targetDocument to item 1 of result
	
	tell document targetDocument
		set sheetList to name of sheets
		choose from list sheetList with prompt "Where will the writing grades go?"
		set writingSheet to item 1 of result
	end tell
	
	tell document sourceDocument
		tell sheet 1
			tell table 1
				tell column "C"
					set format to text
				end tell
			end tell
		end tell
	end tell
	
	add column after last column of table 1 of sheet writingSheet of document targetDocument
	set value of cell 1 of last column of table 1 of sheet writingSheet of document targetDocument to assignmentTitle
	
	
	repeat with i from 2 to count of cells of column "A" of table 1 of sheet writingSheet of document targetDocument
		set studentIDTarget to value of cell i of column "A" of table 1 of sheet writingSheet of document targetDocument
		repeat with j from 2 to count of cells of column "C" of table 1 of sheet 1 of document sourceDocument
			#			set format of cell j of column "C" of table 1 of sheet 1 of document sourceDocument to text
			set studentIDSource to value of cell j of column "C" of table 1 of sheet 1 of document sourceDocument
			if studentIDSource is equal to studentIDTarget then
				if cell j of column "H" of table 1 of sheet 1 of document sourceDocument is not missing value then
					set value of cell i of the last column of table 1 of sheet writingSheet of document targetDocument to true
					exit repeat
				end if
			end if
		end repeat
	end repeat
end tell