use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

--This script create tables in Microsoft Word. One table per page

set maxPages to 25
set tblRows to 9
set tblColumns to 2

tell application "Microsoft Word"
	make new document
	set doc to active document
	
	repeat with i from 1 to maxPages
		
		set tbl to make new table at doc with properties {text object:(create range doc start 0 end 0), number of rows:tblRows, number of columns:tblColumns}
		set rng to collapse range (text object of tbl) direction collapse start
		if i < maxPages then
			insert break at rng break type page break
		end if
		
		tell border options of tbl
			set inside line style to line style single
			set inside line width to line width100 point
			set outside line style to line style single
			set outside line width to line width100 point
		end tell
		
	end repeat
	set table gridlines of view of active window to true
	
end tell
