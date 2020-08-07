use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This Script grabs records from Numbers and puts them into tables in Pages, one table per page
# Prompt user to choose excel file and grab data
set rowData to {}
tell application "Numbers"
	open (choose file with prompt "Choose file:")
	tell document 1
		tell sheet 1
			tell table 1
				set recordCount to (count of rows) - 2
				repeat with i from 1 to count of rows
					tell row i
						set end of rowData to {value of cell 6, value of cell 7, value of cell 8, value of cell 9, value of cell 10, value of cell 11, value of cell 12, value of cell 13}
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell

#Create new pages document and create the number of pages and tables
tell application "Pages"
	make new document
	tell document 1
		repeat with i from 1 to recordCount - 1
			make page
		end repeat
		repeat with i from 1 to count of pages
			tell page i
				make table with properties {column count:2, row count:count of items of item 1 of rowData}
				tell table 1
					set position to {25, 25 + (792 * (i - 1))}
					set width of first column to 100
					set width of second column to 465
				end tell
				
			end tell
		end repeat
	end tell
end tell

# Set the Field labels
tell application "Pages"
	tell document 1
		repeat with i from 1 to count of pages
			tell page i
				tell table 1
					set value of cell 1 of row 1 to "Activity Title"
					set value of cell 1 of row 2 to "Level"
					set value of cell 1 of row 3 to "Cutting Edge Module"
					set value of cell 1 of row 4 to "Language Focus"
					set value of cell 1 of row 5 to "Activity Duration"
					set value of cell 1 of row 6 to "Overview"
					set value of cell 1 of row 7 to "Instructions"
					set value of cell 1 of row 8 to "Adaptations"
				end tell
			end tell
		end repeat
	end tell
end tell


# Fill in the data for each table
tell application "Pages"
	tell document 1
		repeat with i from 1 to count of pages
			tell page i
				tell table 1
					repeat with j from 1 to count of rows
						tell row j
							set value of cell 2 to item j of item (i + 1) of rowData
						end tell
					end repeat
				end tell
			end tell
		end repeat
	end tell
end tell