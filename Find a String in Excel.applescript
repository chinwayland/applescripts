use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Microsoft Excel"
	tell workbook 1
		tell worksheet 1
			try
				find used range what "Student"
			end try
		end tell
	end tell
	
end tell
