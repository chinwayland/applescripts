-- This script checks a mailbox for messages and email addresses and then checks a box in the corresponding cell in a Numbers document

tell application "Mail"
	activate
	(*
	set emailAccounts to name of accounts # grab list of accounts
	tell me to say "Select the account with the assignments you want to retrieve."
	set chosenAccount to choose from list emailAccounts with prompt "Please choose the mail account" # choose an account
	set chosenAccount to item 1 of chosenAccount # flatten list to string
	*)
	set chosenAccount to "Wintec"
	set accountMailboxes to name of mailboxes of account chosenAccount # grab list of mailboxes
	tell me to say "Select the mailbox"
	set chosenMailbox to choose from list accountMailboxes with prompt "Please select the mailbox" # choose a mailbox
	set chosenMailbox to item 1 of chosenMailbox # flatten list to string
	set chosenMailboxAndAccount to mailbox chosenMailbox of account "Wintec" #chosenAccount # combine mailbox and account into one variable
	tell me to say "Grabbing Data"
	set senders to sender of messages of chosenMailboxAndAccount # grab list of senders (Name, Email)
	set listOfEmails to {} # variable for emails only
	repeat with i in senders # extract the email addresses from the list of senders
		set end of listOfEmails to extract address from i
	end repeat
end tell

tell application "Numbers"
	activate
	tell me to say "Select the gradebook"
	set openDocuments to name of documents
	set targetFile to choose from list openDocuments with prompt "Select Target File."
	set targetFile to item 1 of targetFile
	#	set targetFile to id of (open (choose file with prompt "Target File"))
	tell me to say "transfering data"
	tell document targetFile
		tell sheet 1
			tell table 1
				set rowData to value of cells of row 1
				if chosenMailbox is not in rowData then
					add column after last column
					set value of first cell of last column to chosenMailbox
					(*
					tell row 1 # find email column
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
*)
				end if
				tell row 1 # find email column
					set columnAddress to address of column of first cell whose value contains "email"
				end tell
				
				set gradebookEmails to value of cells of column columnAddress
				# find column to update
				repeat with i from 1 to count of columns
					tell row 1
						set columnToUpdate to column of (first cell whose value is chosenMailbox)
					end tell
				end repeat
				set emailsThatDoNotMatch to {}
				set gradedLoopCounter to 0
				repeat with email in listOfEmails
					if email is in gradebookEmails then
						if the value of the cell (address of row of (first cell whose value is email)) of columnToUpdate is false then
							set the value of the cell (address of row of (first cell whose value is email)) of columnToUpdate to true
							set gradedLoopCounter to gradedLoopCounter + 1
						end if
					else
						set end of emailsThatDoNotMatch to email
					end if
				end repeat
			end tell
		end tell
	end tell
end tell

say "finshed and checked " & gradedLoopCounter & " checkboxes"