use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This Script grabs records from Numbers and puts them into tables in Pages, one table per page
# Prompt user to choose excel file and grab data
set paddingLeft to display dialog "Choose the width of the padding on the left margin. (default: 25)" default answer 25 with icon note buttons {"Cancel", "Continue"} default button "Continue"
set paddingTop to display dialog "Choose the width of the padding on the top margin. (default: 25)" default answer 25 with icon note buttons {"Cancel", "Continue"} default button "Continue"
set paddingLeftRightMatchChoice to {true, false}
set paddingLeftRightMatch to choose from list paddingLeftRightMatchChoice with prompt "Do you want the padding for the left and right margin to match?" default items {false}


set leftPadding to text returned of paddingLeft
set topPadding to text returned of paddingTop

#set logo to (choose file with prompt "Please select logo file:")
set rowData to {}
tell application "Numbers"
	display dialog "Please choose the input file (Excel)"
	open (choose file with prompt "Please choose the input file:")
	tell document 1
		tell sheet 1
			tell table 1
				set recordCount to (count of rows) - 2
				repeat with i from 1 to count of rows
					tell row i
						set end of rowData to {value of cell 6, value of cell 7, value of cell 8, value of cell 9, value of cell 10, value of cell 11, value of cell 12, value of cell 13, value of cell 5, value of cell 4}
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell



#Create new pages document and create the number of pages and tables
tell application "Pages"
	activate
	make new document
	tell document 1
		repeat with i from 1 to recordCount - 1
			make page
		end repeat
		repeat with i from 1 to count of pages
			tell page i
				make table with properties {column count:2, row count:(count of items of item 1 of rowData) + 1}
				tell table 1
					set position to {leftPadding, topPadding + (792 * (i - 1))}
					set width of first column to 100
					
					if item 1 of paddingLeftRightMatch is true then
						set width of second column to (465 - leftPadding)
					else
						set width of second column to (465 - (leftPadding * 0.75))
					end if
					
					tell cell 1
						set value to "Jinhua English Teachers' Online Activities/Resources Form"
						set alignment to center
					end tell
				end tell
			end tell
		end repeat
		
		#merge the first row / Title
		repeat with i from 1 to count of tables
			tell table i
				tell row 1
					merge
				end tell
			end tell
		end repeat
	end tell
end tell


# Set the Field labels
tell application "Pages"
	activate
	tell document 1
		set fieldLabels to {"Activity Title", "Level", "Cutting Edge Module", "Language Focus", "Activity Duration", "Overview", "Instructions", "Adaptations", "Contributor Name", "Contributor Email"}
		repeat with i from 1 to count of pages
			tell page i
				tell table 1
					repeat with j from 1 to count of fieldLabels
						set value of cell 1 of row (j + 1) to item j of fieldLabels
					end repeat
					(* 
					set value of cell 1 of row 1 to "Activity Title"
					set value of cell 1 of row 2 to "Level"
					set value of cell 1 of row 3 to "Cutting Edge Module"
					set value of cell 1 of row 4 to "Language Focus"
					set value of cell 1 of row 5 to "Activity Duration"
					set value of cell 1 of row 6 to "Overview"
					set value of cell 1 of row 7 to "Instructions"
					set value of cell 1 of row 8 to "Adaptations"
					set value of cell 1 of row 9 to "Contributor Name"
					set value of cell 1 of row 10 to "Contributor Email"
					*)
				end tell
			end tell
		end repeat
	end tell
end tell


# Fill in the data for each table
tell application "Pages"
	activate
	tell document 1
		repeat with i from 1 to count of pages
			tell page i
				tell table 1
					repeat with j from 1 to ((count of rows) - 1)
						tell row (j + 1)
							set value of cell 2 to item j of item (i + 1) of rowData
						end tell
					end repeat
				end tell
			end tell
		end repeat
	end tell
end tell