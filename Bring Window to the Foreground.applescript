use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

raiseWindow2 of "Microsoft Excel" for "Book2"

to raiseWindow2 of theApplicationName for theName
	tell the application named theApplicationName
		activate
		set theWindow to the first item of ¬
			(get the windows whose name is theName)
		if the index of theWindow is not 1 then
			set the index of theWindow to 2
			tell application "System Events" to ¬
				tell application process theApplicationName to ¬
					keystroke "`" using command down
		end if
	end tell
end raiseWindow2