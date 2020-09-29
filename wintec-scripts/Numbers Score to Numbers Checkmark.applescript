-- This script checks if a number from a Numbers spreadsheet exists in another spreadsheet and makes a checkmark in a column of the second spreadsheet.

# Declare variables

set listOfAllEmails to {{{}, {}, {}}, {{}, {}, {}}, {{}, {}, {}}, {{}, {}, {}}}
set listOfSourceEmails to {}
set mysteryEmails to {}
set combinedReferenceSheets to {}


# Ask user for Numbers column letter of the source keywords

set input to display dialog "What is the column letter of the email address or Student ID#?" default answer ""
set columnLetter to input's text returned

# Ask user for the new column name in the target Numbers document

set input to display dialog "What is the new column name?" default answer ""
set columnName to input's text returned


tell application "Numbers"
	
	# Get all emails from the source document
	
	tell document 1
		tell sheet 1
			tell table 1
				tell column columnLetter
					repeat with rowNumber from 2 to the count of the rows
						set the end of listOfSourceEmails to the value of the cell rowNumber
					end repeat
				end tell
			end tell
		end tell
	end tell
	
	# Get all the emails from the target Numbers document, 
	
	tell document 2
		repeat with sheetNumber from 1 to 4
			tell sheet sheetNumber
				tell table 1
					repeat with columnNumber from 7 to 9 # Get all the emails from the target Numbers document
						tell column columnNumber
							set listOfAllEmails's item sheetNumber's item (columnNumber - 6) to the value of the cells
						end tell
					end repeat
				end tell
			end tell
		end repeat
		
		# Add New Columns
		repeat with sheetNumber from 1 to 4
			tell sheet sheetNumber
				tell table 1
					add column after the last column # Add Column to the right side of sheets 1 to 4
				end tell
			end tell
		end repeat
		
		# Set New Column Names
		
		repeat with sheetNumber from 1 to 4
			tell sheet sheetNumber
				tell table 1
					set the value of row 1's last cell to columnName # Set the title of the new column headers
				end tell
			end tell
		end repeat
		
		# format the cells to checkbox
		
		repeat with sheetNumber from 1 to 4
			tell sheet sheetNumber
				tell table 1
					tell the last column
						repeat with rowCounter from 2 to count of column
							set the format of the cells to checkbox # Set the format of the rightmost column to checkbox
						end repeat
					end tell
				end tell
			end tell
		end repeat
		
	end tell
	
	# Check if sourceEmail is in listOfSourceEmails then check the box
	
	repeat with sourceEmail in listOfSourceEmails
		
		set documentSheetNumber to 0
		repeat with sheetNumber in listOfAllEmails
			set documentSheetNumber to documentSheetNumber + 1
			repeat with columnNumber in sheetNumber
				set emailRowNumber to 0
				repeat with email in columnNumber
					set emailRowNumber to emailRowNumber + 1
					if (sourceEmail as string) is equal to (email as string) then
						tell document 2
							tell sheet documentSheetNumber
								tell table 1
									set the value of row emailRowNumber's last cell to true
								end tell
							end tell
						end tell
					end if
				end repeat
			end repeat
		end repeat
	end repeat
	
end tell

# Get all 4 sheets data into memory
tell application "Numbers"
	tell document 2
		repeat with sheetNumber from 1 to 4
			tell sheet sheetNumber
				tell table 1
					set referenceSheets's item sheetNumber to value of cells
				end tell
			end tell
		end repeat
	end tell
end tell

#Check for mystery Emails
repeat with referenceSheet in referenceSheets
	set combinedReferenceSheets to combinedReferenceSheets & referenceSheet
end repeat

repeat with sourceEmail in sourceEmails
	if sourceEmail is not in combinedReferenceSheets then
		set end of mysteryEmails to sourceEmail
	end if
end repeat
