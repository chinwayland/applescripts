-- This script transfers a score from a Numbers Document to an Excel Workbook

set tests to {"Progress Test #1", "Progress Test #2", "Progress Test #3", "Participation (out of 10)", "Speaking Exam", "Written Exam", "Final Exam Score"}
set chosenTest to choose from list (tests) with prompt "Please choose the test you are transfering data for"
set chosenTest to item 1 of chosenTest
set scoreData to {}
tell application "Numbers"
	set openDocuments to name of documents
	set sourceFile to choose from list openDocuments with prompt "What is the source file?"
	set sourceFile to item 1 of sourceFile
	#set sourceFile to id of (open (choose file with prompt "Source File"))
	tell me to say "Grabbing data"
	tell document sourceFile
		set excelSheetName to name
		tell sheet 1
			tell table 1
				repeat with i from 1 to count of columns
					if value of cell 1 of column i is "Student ID" then
						set studentIDColumn to column i
						exit repeat
					end if
				end repeat
				repeat with i from 1 to count of columns
					if value of cell 1 of column i contains chosenTest then
						set testScoreColumn to column i
					end if
				end repeat
				repeat with i from 2 to (count of cells of studentIDColumn) - 1
					set end of scoreData to {value of cell i of studentIDColumn, value of cell i of testScoreColumn}
				end repeat
			end tell
		end tell
	end tell
end tell

tell application "Microsoft Excel"
	activate
	set openDocuments to name of documents
	set targetFile to choose from list openDocuments with prompt "Choose target file."
	set targetFile to item 1 of targetFile
	tell me to say "pasting data"
	#set targetFile to id of (open (choose file with prompt "Target File"))
	tell workbook targetFile
		repeat with i from 1 to count of worksheets
			try
				tell worksheet i
					set targetSheetNumber to i
					set classTitleCell to (find used range what excelSheetName)
					exit repeat
				end tell
			end try
		end repeat
		tell worksheet targetSheetNumber
			set targetColumn to first column index of (find used range what chosenTest)
			repeat with i from 1 to count of scoreData
				set targetRow to first row index of (find used range what item 1 of item i of scoreData)
				set value of cell targetRow of column targetColumn to item 2 of item i of scoreData
			end repeat
		end tell
	end tell
end tell

say "Finished"