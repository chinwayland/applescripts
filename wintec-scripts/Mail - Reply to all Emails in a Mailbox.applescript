-- This script replies to all emails in an inbox

# Loop through each message in currently selected mailbox to grab sender's info and put into variables
tell application "Mail"
	set senderContents to {}
	set senderInfos to {}
	set senderAddresses to {}
	set senderSubjects to {}
	set senderAttachmentFileNames to {}
	set mailboxName to name of mailbox of item 1 of (get selection)
	set accountName to name of account of mailbox of item 1 of (get selection)
	
	tell account accountName
		tell mailbox mailboxName
			tell messages
				repeat with messageNumber from (count) to 1 by -1
					tell item messageNumber
						set end of senderContents to content
						set end of senderInfos to sender
						set end of senderAddresses to extract address from senderInfos as string
						(*
						set end of senderSubjects to subject
						if mail attachments exists then
							set end of senderAttachmentFileNames to name of item 1 of mail attachments
						else
							set end of senderAttachmentFileNames to "No Attachments"
						end if
						*)
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell

# Get English first names from Numbers document and put into list senderNames
set cellHit to {}
set senderNames to {}
repeat with senderAddress in senderAddresses
	tell application "Numbers"
		tell document 1
			repeat with sheetNumber from 1 to 4
				tell sheet sheetNumber
					tell table 1
						if (every cell whose value is senderAddress) is not {} then
							set cellHit to (every cell whose value is senderAddress)
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

#send a message to each sender from above steps
repeat with messageNumber from 1 to count of senderAddresses
	tell application "Mail"
		set signatureContent to content of signature 1
		set newOutGoingMessage to make new outgoing message
		tell newOutGoingMessage
			set visible to true
			make new to recipient at end of to recipients with properties {address:item messageNumber of senderAddresses}
			set subject to "Re: " & item messageNumber of senderSubjects
			set content to "Thank you " & item messageNumber of senderNames & "!


" & signatureContent & "	
	You wrote: 
	" & item messageNumber of senderContents & "
"
			#& item messageNumber of senderAttachmentFileNames
			send
		end tell
	end tell
end repeat
