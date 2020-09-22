-- This script is not yet finished. Need to figure out how to refer to a specific cell in a table

#Grab information from Excel
tell application "Microsoft Excel"
	open (choose file with prompt "Source File")
	set sourceFile to active workbook
end tell

tell application "Script Debugger"
	activate
end tell

tell application "Microsoft Excel"
	tell me to say "grabbing data"
	set rowData to {}
	tell sourceFile
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
(*
tell application "Microsoft Word"
	activate
	tell me to say "pasting data"
	make new document
	
	repeat ((count of rowData) - 1) times
		set oDoc to active document
		set end of oTable to make new table at oDoc with properties ¬
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
*)

set maxPages to number of items in rowData
set tblRows to number of items in item 1 of rowData
set tblColumns to 2

tell application "Microsoft Word"
	activate
	make document
	set doc to active document
	(*
	tell doc # create the necessary number of pages
		repeat maxPages - 1 times
			insert break at text object break type page break
		end repeat
	end tell
	*)
	tell doc # create the necessary number of pages and tables
		repeat with i from 1 to maxPages
			# new table
			set tbl to make new table at doc with properties {text object:(create range doc start 0 end 0), number of rows:tblRows, number of columns:tblColumns}
			# new page
			set rng to collapse range (text object of tbl) direction collapse start
			if i < maxPages then
				insert break at rng break type page break
			end if
		end repeat
		
		repeat with i from 1 to number of tables
			tell border options of table i
				set inside line style to line style single
				set inside line width to line width100 point
				set outside line style to line style single
				set outside line width to line width100 point
			end tell
			
			tell table i
				
				-- Something to do with the text object of a table
				
				set table gridlines to true
				tell column 1
					repeat with j from 1 to tblRows
						#			set value of cell j to item j of item i of rowData
					end repeat
				end tell
			end tell
		end repeat
		#		set table gridlines of view of active window to true
	end tell
	
	
	
end tell