-- This script will check the exported gradebook from MyEnglishLav and check if the student has done the assignments. They will get a checkmark for each assignment done. The MyEnglishLab Username has to be in the gradebook for this to work.
tell application "Numbers"
	tell me to say "Where is the master gradebook report export from MyEnglishLab?"
	set sourceFile to id of (open (choose file with prompt "Source File"))
	tell me to say "Where is the Class gradebook that you want to update?"
	set targetFile to id of (open (choose file with prompt "Target File"))
	
	tell document id sourceFile
		set unitGrades to name of sheets
		set reducedUnitGrades to {}
		repeat with i from 1 to count of unitGrades
			if item i of unitGrades contains "Assignments" then
				set end of reducedUnitGrades to item i of unitGrades
			end if
		end repeat
		tell me to say "Which Unit's grades do you want to import?"
		set chosenUnit to choose from list (reducedUnitGrades) with prompt "Please choose the Unit you want to import grades from (Usually, one of the 'Assignments' or 'Video Activities'." # ask user to choose which sheet
		set chosenUnit to item 1 of chosenUnit # flattens the list to a string
		tell me to say "grabbing data"
		tell sheet chosenUnit
			tell table 1
				if chosenUnit is "Assignments VIDEO ACTIVITIES" then
					set unitName to value of cell "C3" # Grab Unit Name
				else
					set unitName to value of cell "C2" # Grab Unit Name					
				end if
				
				set reducedUnitName to text 1 through (offset of ":" in unitName) of unitName --> "UNIT 1:" # truncate to just unit and number and colon
				set scores to {}
				set skip to {"---", missing value}
				set activity to {}
				repeat with i from 3 to (((count of columns) - 4)) by 3 # loop through every third column, skipping first two and last two columns
					if chosenUnit is "Assignments VIDEO ACTIVITIES" then
						set end of activity to value of cell 3 of column i & " " & value of cell 4 of column i # add column name to variable
					else
						set end of activity to value of cell 2 of column i & " " & value of cell 4 of column i # add column name to variable
					end if
					repeat with j from 6 to (count of cells of column 1) - 1 # loop through every row in each column
						if value of cell i of row j is not in skip then # if score not blank then add activity name, username, and score to list
							if chosenUnit is "Assignments VIDEO ACTIVITIES" then
								set end of scores to {value of cell 3 of column i & " " & value of cell 4 of column i, value of cell 2 of row j, value of cell i of row j}
							else
								set end of scores to {value of cell 2 of column i & " " & value of cell 4 of column i, value of cell 2 of row j, value of cell i of row j}
								
							end if
							
						end if
					end repeat
				end repeat
			end tell
		end tell
	end tell
	
	
	tell document id targetFile
		tell me to say "pasting data"
		if chosenUnit is "Assignments VIDEO ACTIVITIES" then
			set sheetToUpdate to sheet 2
		else
			set sheetToUpdate to sheet 1
		end if
		tell sheetToUpdate
			tell table 1
				try # check if checkmarks already exist for this unit
					if (first cell of row 1 whose value of it contains item 1 of activity) is not missing value then
						tell me to say "updating cells"
					end if
				on error
					repeat with i from 1 to count of activity # If not previous checked, add new columns
						add column after the last column
						set the format of the last column to checkbox
						set the value of the first cell of the last column to item i of activity # set column name
						
						(*
						repeat with j from 2 to (count of cells in the last column) - 1
							tell the last column
								set the format of cell j to checkbox
							end tell
						end repeat
						*)
					end repeat
				end try
				
				set newUsernames to {}
				
				# find column with username
				tell row 1
					repeat with i from 1 to count of cells
						if value of cell i contains "MyEnglishLab Username" then
							set usernameColumn to column of cell i
						end if
					end repeat
				end tell
				
				
				
				repeat with i from 1 to count of scores
					
					tell usernameColumn
						
						try
							set rowAddress to address of row of (first cell whose value of it contains item 2 of item i of scores) # get row name of username							
						on error
							# tell me to say "encountered new username, add and run script again"
							set end of newUsernames to item 2 of item i of scores
							repeat 1 times # fake loop to simulate continue
								exit repeat
							end repeat
						end try
						
					end tell
					
					tell row 1
						set columnName to column of (first cell whose value of it contains item 1 of item i of scores) # get column name of activity
					end tell
					tell columnName
						tell cell rowAddress
							set value to true
						end tell
					end tell
				end repeat
			end tell
		end tell
	end tell
	
	
end tell

say "Finished"