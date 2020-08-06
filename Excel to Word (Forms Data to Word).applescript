use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This script is unfinished. Can grab data data from Excel

tell application "Microsoft Excel"
	set theDocument to choose file with prompt "Please select a document to process:"
	open theDocument
	tell active sheet
		tell used range
			set rowCount to count of rows
		end tell
	end tell
end tell

tell application "Microsoft Word"
	make new document
	
end tell

repeat with i from 2 to rowCount
	tell application "Microsoft Excel"
		tell active sheet
			tell row i
				set activityTitle to value of cell 6
				set level to value of cell 7
				set cuttingEdgeModule to value of cell 8
				set languageFocus to value of cell 9
				set duration to value of cell 10
				set overview to value of cell 11
				set instructions to value of cell 12
				set adapations to value of cell 13
				set supplementaryMaterials to value of cell 14
			end tell
		end tell
	end tell
	
	tell application "Microsoft Word"
		tell active document
			insert text activityTitle at end of text object
		end tell
	end tell
end repeat


