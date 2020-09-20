use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

(*
tell application "System Events"
	tell process "WeChat"
		tell window 1
			tell splitter group 1
				tell splitter group 1
					tell scroll area 2
						value of text area 1
					end tell
				end tell
			end tell
		end tell
	end tell
end tell
*)

tell application "System Events"
	tell process "WeChat"
		tell window 1
			tell splitter group 1
				tell scroll area 1
					tell table 1
						tell row 1
							tell UI element 1
								click
							end tell
						end tell
					end tell
				end tell
			end tell
		end tell
	end tell
end tell