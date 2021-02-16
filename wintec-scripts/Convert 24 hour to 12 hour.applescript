-- This script converts times from the 24 hour format to the 12 hour format

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

set theDate to "March 1, 2021 at 9:00"

set theTime to word 5 of theDate & ":" & word 6 of theDate

if word 1 of the theTime > 12 then
	set theTime2 to (word 1 of theTime) - 12 & ":" & word 2 of theTime & "PM" as text
else
	set theTime2 to word 1 of theTime & ":" & word 2 of theTime & "AM" as text
end if
