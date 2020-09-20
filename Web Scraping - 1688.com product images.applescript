-- This script extracts image URLs from a 1688.com product page and downloads the photos to the desktop

set photoURL to {}
tell application "Safari"
	tell document 1
		repeat with i from 0 to 28 by 3
			set end of photoURL to do JavaScript "document.getElementById('desc-lazyload-container').childNodes[3].childNodes[" & i & "].src"
		end repeat
	end tell
end tell

tell application "Safari"
	activate
	make new document
	tell document 1
		repeat with photo in photoURL
			set URL to photo
			delay 2
			tell application "System Events" to keystroke "s" using command down
			delay 2
			tell application "System Events" to keystroke return
			delay 2
		end repeat
	end tell
end tell