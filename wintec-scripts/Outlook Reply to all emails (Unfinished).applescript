-- This script replies to all emails in an Outlook Inbox (not complete)

set senders to {}
set emails to {}
set cellHit to {}
set senderNames to {}

tell application "Microsoft Outlook"
	set currentMailFolder to folder of (get selection)
	tell currentMailFolder
		repeat with messageCounter from 1 to count of messages in currentMailFolder
			set replyToMessage to message messageCounter
			set OldContent to content of replyToMessage
			set end of senders to sender of replyToMessage
			
			tell app "Numbers"
			
			
			
			
			set replyMessage to reply to replyToMessage without opening window
			if has html of replyMessage then
				log ("HTML!")
				set the content of replyMessage to "Thank you " & "br><br><br><hr>" & OldContent
			else
				log ("PLAIN TEXT!")
				set the content of replyMessage to "Hello World<br> This is a plain text test <br><br><br><hr>" & OldContent
			end if
			open replyMessage
		end repeat
	end tell
	repeat with sender in senders
		set end of emails to address of sender
	end repeat
end tell


repeat with email in emails
	tell application "Numbers"
		tell document 1
			repeat with sheetNumber from 1 to 4
				tell sheet sheetNumber
					tell table 1
						if (every cell whose value is email) is not {} then
							set cellHit to (every cell whose value is email)
							set sheetIndex to sheetNumber
						end if
					end tell
				end tell
			end repeat
			set rowAddress to address of row of item 1 of cellHit
			tell sheet sheetIndex
				tell table 1
					set end of senderNames to value of first cell of row rowAddress
				end tell
			end tell
		end tell
	end tell
end repeat
