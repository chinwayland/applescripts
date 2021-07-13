-- This script will check the exported gradebook or multiple gradebook export from MyEnglishLab and check if the student has done the assignments. They will get a checkmark for each assignment done. The MyEnglishLab Username has to be in the gradebook for this to work. does not yet work for unit progress tests and video activities

set skip to {"---", "0%"}
tell application "Numbers"
	(*
	tell me to say "Where is the gradebook report from MyEnglishLab?"
	set openDocuments to documents
	set sourceFile to id of (open (choose file with prompt "Source File"))
	tell me to say "Where is the class gradebook that you want to update?"
	set targetFile to id of (open (choose file with prompt "Target File"))
	*)
	
	set documentNames to name of documents
	
	tell me to say "Where is the gradebook report from MyEnglishLab?"
	set sourceFile to choose from list documentNames with prompt "Source File:"
	set sourceFile to item 1 of sourceFile
	
	tell me to say "Where is the class gradebook that you want to update?"
	set targetFile to choose from list documentNames with prompt "Target File:"
	set targetFile to item 1 of targetFile
	
	set targetSheets to name of sheets of document targetFile
	tell me to say "Which sheet in the target file do you want to update?"
	set targetSheet to choose from list targetSheets with prompt "Target Sheet:"
	set targetSheet to item 1 of targetSheet
	
	
	tell document sourceFile
		set gradesFromEachUnit to {}
		repeat with i from 1 to count of sheets
			if name of sheet i contains "Assignments" then
				set end of gradesFromEachUnit to name of sheet i
			end if
		end repeat
		tell me to say "Which Unit's grades do you want to import?"
		# ask user to choose which sheet / unit to import
		set chosenUnit to choose from list (gradesFromEachUnit) with prompt "Please choose the Unit you want to import grades from"
		set chosenUnit to item 1 of chosenUnit # flattens the list to a string
		if chosenUnit is "Assignments VIDEO ACTIVITIES" then
			tell me to say "How many video units do you want to import?"
			# ask user to choose how many video units
			set videoUnitList to {}
			repeat with i from 1 to 14
				set end of videoUnitList to i
			end repeat
			set chosenVideoUnits to choose from list videoUnitList with prompt "Choose how many video actibvity units to import"
			set chosenVideoUnits to item 1 of chosenVideoUnits # flatten list
		end if
		
		tell sheet chosenUnit
			tell table 1
				
				tell column 1
					set classNames to {}
					repeat with i from 1 to count of cells
						if value of cell i contains "(ID: " then
							set end of classNames to value of cell i
						end if
					end repeat
				end tell
				
				# Ask user to choose the class to import grades from
				tell me to say "Which class' grades do you want to import?"
				set chosenClass to choose from list classNames with prompt "Which class do you want to import grades from?"
				set chosenClass to item 1 of chosenClass # Flatten to string
				tell me to say "Grabbing Data"
			end tell
		end tell
	end tell
end tell
(*
tell application "Script Debugger"
	activate
	set event log visible of documents to true
end tell
*)

tell application "Numbers"
	tell document sourceFile
		tell sheet chosenUnit
			tell table 1
				# Find the starting row of the the class
				set startOfClassInfoRowAddress to address of row of first cell where value of it contains chosenClass
				# Find the ending row of the class
				repeat with i from startOfClassInfoRowAddress to count of rows
					if value of cell i of column 1 is "Summary" then
						set endOfClassInfoRowAddress to address of row of cell i of column 1
						exit repeat
					end if
				end repeat
				set scoreColumnAddress to address of column of first cell where value of it is "Score"
				set scoreRowAddress to address of row of first cell where value of it is "Score"
				set columnHeaderNames to {}
				set scores to {}
				set loopCounter to 0
				repeat with i from 1 to count of columns
					#get and create column header names
					if value of cell scoreRowAddress of column i contains "Score" then
						
						if chosenUnit is "Assignments VIDEO ACTIVITIES" then
							# set unitName to value of cell 2 of column i
							set unitName to "VIDEO"
							set unitName2 to value of cell 3 of column i
							
							if item 2 of words of unitName2 > chosenVideoUnits then
								exit repeat
							end if
							
							set end of columnHeaderNames to "MEL " & unitName & " " & text 1 through ((offset of ":" in unitName2) + 1) of unitName2 & value of cell 4 of column i
							
						else if chosenUnit contains "test" then
							set unitName to value of cell 2 of column i
							set end of columnHeaderNames to "MEL " & unitName
							
						else
							set unitName to value of cell 2 of column i
							set end of columnHeaderNames to "MEL " & text 1 through ((offset of ":" in unitName) + 1) of unitName & " " & value of cell 4 of column i
						end if
						# set end of columnHeaderNames to value of cell 2 of column i & " " & value of cell 3 of column i & " " & value of cell 4 of column i
						# set end of columnHeaderNames to value of cell 2 of column i & " " & value of cell 4 of column i
						
						set loopCounter to loopCounter + 1
						repeat with j from (startOfClassInfoRowAddress + 1) to (endOfClassInfoRowAddress - 1)
							if value of cell j of column i is not in skip then
								set end of scores to {value of cell j of column 2, item loopCounter of columnHeaderNames, value of cell j of column i}
							end if
						end repeat
					end if
				end repeat
			end tell
		end tell
	end tell
end tell

tell me to say "pasting data"
tell application "Numbers"
	activate
	tell document targetFile
		tell sheet targetSheet
			tell table 1
				# Check if Unit already has been imported, if not add columns
				repeat with h from 1 to count of columnHeaderNames
					if value of cells of row 1 contains item h of columnHeaderNames then
						# Update Cells
					else
						add column after last column
						set the format of the last column to checkbox
						set the value of the first cell of the last column to item h of columnHeaderNames
					end if
				end repeat
				
				
				# get column where the usernames are				
				tell row 1
					set usernameColumnAddress to address of column of first cell whose value of it contains "MEL Username"
				end tell
				
				(*
				tell row 1
					repeat with i from 1 to count of cells
						if value of cell i contains "MyEnglishLab Username" then
							set usernameColumn to column of cell i
						end if
					end repeat
				end tell
				
				*)
				(*
				repeat with i from 1 to count of scores # loop through scores import
					repeat with j from 1 to count of columns # loop through columns to find matches
						if value of cell 1 of column j contains item 2 of item i of scores then # if the column matches
							repeat with k from 1 to count of rows
								if value of cell k of usernameColumn contains item 1 of item i of scores then
									set value of cell k of column j to true
									exit repeat
								end if
							end repeat
						end if
					end repeat
				end repeat
				*)
				set changedGradeCounter to 0
				repeat with i from 1 to count of scores # loop through scores import
					tell column usernameColumnAddress
						set rowCoordinate to address of row of (first cell whose value of it is item 1 of item i of scores)
					end tell
					tell row 1
						set columnCoordinate to address of column of (first cell whose value of it is item 2 of item i of scores)
					end tell
					tell column columnCoordinate
						tell cell rowCoordinate
							if chosenUnit contains "TEST" then
								set format to percent
								set value to item 3 of item i of scores
								set changedGradeCounter to changedGradeCounter + 1
							else
								if value is false then
									set value to true
									set changedGradeCounter to changedGradeCounter + 1
								end if
							end if
						end tell
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell
say "Finished"
say "checked " & changedGradeCounter & " checkboxes"