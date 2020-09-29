-- This script searches for a specific conversation in WeChat and then sends a designated message to that conversation

tell application "System Events"
	tell process "WeChat"
		tell window 1
			tell splitter group 1
				tell scroll area 1
					tell table 1
						repeat with i from 1 to count of rows
							tell row i
								tell UI element 1
									if the value of static text 1 is "2020 Foreign Teachers" then
										set targetRow to i
										exit repeat
									end if
								end tell
							end tell
						end repeat
					end tell
				end tell
			end tell
		end tell
	end tell
end tell


set message to my messageWithRowIndex:targetRow
select the message -- You don't need to, but you can
tell application "WeChat"
	activate
end tell
tell application "System Events" # Grabs the text from the text entry box
	tell process "WeChat"
		tell window 1
			tell splitter group 1
				tell splitter group 1
					tell scroll area 2
						set value of text area 1 to "1"
						keystroke return
					end tell
				end tell
			end tell
		end tell
	end tell
end tell

on messageWithRowIndex:i
	tell application id "com.apple.systemevents"
		tell (the first process whose bundle identifier is ¬
			"com.tencent.xinWeChat") to tell window 1 ¬
			to tell splitter group 1's scroll area 1's ¬
			table 1's row (i) to if it exists then ¬
			return a reference to it
	end tell
end messageWithRowIndex:
