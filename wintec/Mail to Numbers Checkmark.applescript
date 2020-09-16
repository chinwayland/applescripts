use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions


-- This script checks a mailbox for messages and email addresses and then checks a box in the corresponding cell in a Numbers document

tell application "Mail"
	(*
	set emailAccounts to name of accounts # grab list of accounts
	set chosenAccount to choose from list emailAccounts with prompt "Please choose the mail account" # choose an account
	set chosenAccount to item 1 of chosenAccount # flatten list to string
	*)
	
	set accountMailboxes to name of mailboxes of account "Wintec" #chosenAccount # grab list of mailboxes
	set chosenMailbox to choose from list accountMailboxes with prompt "Please select the mailbox with the emails you want to process." # choose a mailbox
	set chosenMailbox to item 1 of chosenMailbox # flatten list to string
	set chosenMailboxAndAccount to mailbox chosenMailbox of account "Wintec" #chosenAccount # combine mailbox and account into one variable
	set senders to sender of messages of chosenMailboxAndAccount # grab list of senders (Name, Email)
	set listOfEmails to {} # variable for emails only
	repeat with i in senders # extract the email addresses from the list of senders
		set end of listOfEmails to extract address from i
	end repeat
end tell

tell application "Numbers"
	activate
	set targetFile to id of (open (choose file with prompt "Target File"))
	tell document id targetFile
		tell sheet 1
			tell table 1
				add column after last column
				set value of first cell of last column to chosenMailbox
				tell row 1
					set columnAddress to address of column of first cell whose value contains "email"
				end tell
				set gradebookEmails to value of cells of column columnAddress
				set emailsThatDoNotMatch to {}
				repeat with email in listOfEmails
					if email is in gradebookEmails then
						set the value of the cell (address of row of (first cell whose value is email)) of the last column to true
					else
						set end of emailsThatDoNotMatch to email
					end if
				end repeat
			end tell
		end tell
	end tell
end tell