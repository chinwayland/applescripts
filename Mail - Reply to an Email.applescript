use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- Yvan KOENIG (VALLAURIS, France) 14 juin 2020 16:41:04

my Germaine()

on Germaine()
	tell application "Mail"
		-- Grab the localized spelling only once
		set x to "CTM-M5-1Qy.title"
		if (system attribute "sys2") as integer < 15 then
			set signature_loc to localized string x from table "ComposeViewController"
		else
			set signature_loc to localized string x from table "ComposeViewControllerWK2"
		end if
		if signature_loc is x then set signature_loc to "Signature:"
		set _sel to selection
	end tell -- Mail
	-- Edit the instruction below to fit your needs
	set myAnswer to "Thank you." & return & "The text gets inserted into the first line in the Mail reply message, but the signature and quoted text are deleted." & return & "Is there a way to create a reply message, maintain the signature and quoted text, while inserting next text above it?"
	
	repeat with _msg in _sel
		tell application "Mail" to tell _msg
			set replyMessage to reply
		end tell
		tell application "System Events" to tell process "Mail"
			set frontmost to true
			tell window 1
				-- class of UI elements --> {static text, text field, button, static text, text field, button, static text, text field, button, static text, text field, button, static text, text field, scroll area, pop up button, static text, pop up button, static text, button, button, button, toolbar, group}
				-- name of pop up buttons --> {"De :", "Signature :"}
				-- value of pop up buttons --> {"Yvan KOENIG – koenigyvan<at>mac<dot>com", "Aucune"}
				if not (exists (pop up button 1 whose name is signature_loc)) then error "The “" & signature_loc & "” pop up is unavailable" number -128
				tell scroll area 1
					tell UI element 1
						tell group 3
							tell static text 1
								keystroke myAnswer & return
							end tell -- static text 1
						end tell -- group 3
					end tell -- UI element 1
				end tell -- scroll area 1
			end tell -- window 1
		end tell -- System Events
		
		tell application "Mail" to tell replyMessage
			-- send -- send the message -- ENABLE it after testing
		end tell
	end repeat
end Germaine

#=====
