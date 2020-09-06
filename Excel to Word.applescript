-- This script is not yet finished.

#Grab information from Excel
tell application "Microsoft Excel"
	open (choose file with prompt "Choose File:")
	set rowData to {}
	tell workbook 1
		tell worksheet 1
			tell used range
				repeat with i from 2 to count of rows
					tell row i
						set end of rowData to {value of cell 6, value of cell 7, value of cell 8, value of cell 9, value of cell 10, value of cell 11, value of cell 12, value of cell 13, value of cell 14}
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell

#Paste data into Word inside tables (not complete)
tell application "Microsoft Word"
	make new document
	
	repeat ((count of rowData) - 1) times
		set oDoc to active document
		set end of oTable to make new table at oDoc with properties Â
			{text object:(create range oDoc start 0 end 0), number of rows:8, number of columns:2}
		tell document 1
			set enable borders of border options of table 1 to true
		end tell
		move end of range (text object of document 1) by a character item count -1
		tell document 1
			insert break at text object
		end tell
	end repeat
	
end tell