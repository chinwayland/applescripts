-- This Script grabs records from Numbers and puts them into tables in Pages, one table per page

# Prompt user to choose excel file and grab data
#say "Please choose the left margin size"
set marginLeft to display dialog "Choose the width of the margin on the left. (default: 25, minimum: 1)" default answer 25 with icon note buttons {"Cancel", "Continue"} default button "Continue"
#say "Please choose the top margin size"
set marginTop to display dialog "Choose the width of the margin on the top. (default: 70,minimum: 1)" default answer 70 with icon note buttons {"Cancel", "Continue"} default button "Continue"
#say "Please choose the right margin size"
set marginRight to display dialog "Choose the width of the margin on the right. (default: 25, minimum: 1)" default answer 25 with icon note buttons {"Cancel", "Continue"} default button "Continue"

#set marginLeftRightMatchChoice to {true, false}
#say "Do you want the left and right margin sizes to be the same?"
#set marginLeftRightMatch to choose from list marginLeftRightMatchChoice with prompt "Do you want the margin for the left side and right side to match?" default items {false}
-- promput user for images
#say "Please tell me where the wintec logo is"
display dialog ("Select the Wintec logo to import")
set logoWintec to (choose file with prompt "Select the Wintec Logo to import:")
#say "Please tell me where the jinhua logo is"
display dialog ("Select the Jinhua Polytechnic logo to import")
set logoJinhua to (choose file with prompt "Select the Jinhua Logo to import:")

set leftMargin to text returned of marginLeft
set topMargin to text returned of marginTop
set rightMargin to text returned of marginRight

#set logo to (choose file with prompt "Please select logo file:")
set rowData to {}
tell application "Numbers"
	#	tell me to say "Please tell me where the excel file is"
	display dialog "Please choose the input file (Excel)"
	open (choose file with prompt "Please choose the input file:")
	tell document 1
		tell sheet 1
			tell table 1
				set recordCount to (count of rows) - 2
				repeat with i from 1 to count of rows
					tell row i
						set end of rowData to {value of cell 6, value of cell 7, value of cell 8, value of cell 9, value of cell 10, value of cell 11, value of cell 12, value of cell 13, value of cell 5, value of cell 4, value of cell 14}
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell



#Create new pages document and create the number of pages and tables
set maxPageWidth to 611
tell application "Pages"
	activate
	make new document
	tell document 1
		repeat with i from 1 to recordCount - 1
			make new page
		end repeat
		
		repeat with i from 1 to count of pages
			tell page i
				-- import the image
				make new image with properties {file:logoWintec, height:40, position:{leftMargin, topMargin - 50}}
				make new image with properties {file:logoJinhua, height:40, position:{maxPageWidth - 40 - rightMargin, topMargin - 50}}
			end tell
		end repeat
		
		repeat with i from 1 to count of pages
			tell page i
				make table with properties {column count:2, row count:(count of items of item 1 of rowData)}
				tell table 1
					set position to {leftMargin, topMargin + (792 * (i - 1))}
					set firstColumnWidth to 95
					set width of first column to firstColumnWidth
					set width of second column to (maxPageWidth - leftMargin - firstColumnWidth - rightMargin)
				end tell
			end tell
		end repeat
		
		#merge the first row / Title
		(*
		repeat with i from 1 to count of tables
			tell table i
				tell row 1
					merge
				end tell
			end tell
		end repeat
		*)
		
	end tell
end tell

# Add a Title to each page

set tableWidth to maxPageWidth - leftMargin - rightMargin

tell application "Pages"
	tell document 1
		repeat with i from 1 to count of pages
			tell page i
				make text item with properties {position:{leftMargin + tableWidth * 0.11, topMargin - 40}, height:25, width:400}
				tell text item 1
					set object text to "Jinhua English Teachers' Online Activities/Resources Form"
					set size of object text to 16
				end tell
			end tell
		end repeat
	end tell
end tell

# Set the Field labels
tell application "Pages"
	tell document 1
		set fieldLabels to {"Activity Title", "Level", "Cutting Edge Unit", "Language Focus", "Activity Duration", "Overview", "Instructions", "Adaptations", "Contributor Name", "Contributor Email", "Linked File"}
		repeat with i from 1 to count of pages
			tell page i
				tell table 1
					repeat with j from 1 to count of fieldLabels
						set value of cell 1 of row j to item j of fieldLabels
					end repeat
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
					repeat with j from 1 to ((count of rows))
						tell row j
							set value of cell 2 to item j of item (i + 1) of rowData
							set alignment to left
						end tell
					end repeat
				end tell
			end tell
		end repeat
	end tell
end tell
