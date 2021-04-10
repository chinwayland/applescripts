tell application "System Preferences"
	activate
	delay 3
end tell

tell application "System Events"
	tell process "System Preferences"
		tell window 1
			tell button 1
				click
			end tell
		end tell
	end tell
end tell

tell application "System Events"
	tell process "System Preferences"
		tell window 1
			tell group 1
				tell scroll area 1
					tell table 1
						tell row 5
							tell UI element 1
								tell UI element 1
									click
									delay 3
									click
								end tell
							end tell
						end tell
					end tell
				end tell
			end tell
		end tell
	end tell
end tell