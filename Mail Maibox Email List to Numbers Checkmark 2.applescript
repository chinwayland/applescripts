# Declare variables Needed

set roughList to {}
set duplicateEmailList to {}
set listOfAllEmails to {{{}, {}, {}}, {{}, {}, {}}, {{}, {}, {}}, {{}, {}, {}}}
set sourceEmails to {}
set combinedReferenceSheets to {}

set input to display dialog "What is the new column name?" default answer ""
set columnName to input's text returned

# Extract list of email addresses from currently selected mailbox in Mail

tell application "Mail"
	repeat with senderComplex in (get (get selection)'s item 1's mailbox's messages's sender)
		set roughList's end to extract address from senderComplex
	end repeat
end tell

# Remove duplicate email addresses from above list

repeat with i from 1 to count of items of roughList
	if item i of roughList is not in sourceEmails then
		set sourceEmails to sourceEmails & item i of roughList
	else
		set end of duplicateEmailList to item i of roughList
	end if
end repeat

tell application "Numbers"
	
	tell document 1
		
		# Get all the emails from the target Numbers document, 
		
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
		
		# Add New Columns and set the format to checkbox
		repeat with sheetNumber from 1 to 4
			tell sheet sheetNumber
				tell table 1
					add column after the last column # Add Column to the right side of sheets 1 to 4
					set the format of the last column to checkbox
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
	
	repeat with sourceEmail in sourceEmails
		set documentSheetNumber to 0
		repeat with sheetNumber in listOfAllEmails
			set documentSheetNumber to documentSheetNumber + 1
			repeat with columnNumber in sheetNumber
				set emailRowNumber to 0
				repeat with email in columnNumber
					set emailRowNumber to emailRowNumber + 1
					if (sourceEmail as string) is equal to (email as string) then
						tell document 1
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
set referenceSheets to {{}, {}, {}, {}}
tell application "Numbers"
	tell document 1
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
set mysteryEmails to {}

repeat with referenceSheet in referenceSheets
	set combinedReferenceSheets to combinedReferenceSheets & referenceSheet
end repeat

repeat with sourceEmail in sourceEmails
	if sourceEmail is not in combinedReferenceSheets then
		set end of mysteryEmails to sourceEmail
		log sourceEmail
	end if
end repeat