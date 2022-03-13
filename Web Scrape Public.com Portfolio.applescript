use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Safari"
	activate
	tell document 1
		set buttonStatus to do JavaScript "document.getElementsByClassName('css-t33jsd')[0].textContent"
		if buttonStatus is "Show all" then
			do JavaScript "document.getElementsByClassName('css-t33jsd')[0].click()"
		end if
		set numberOfRows to do JavaScript "document.body.getElementsByClassName('css-t46qbh').length"
		
		set stockDataList to {}
		repeat with i from 0 to numberOfRows - 1
			set stockData to do JavaScript "document.body.getElementsByClassName('css-t46qbh')[" & i & "].innerText" #Grab stock data from row
			#do JavaScript "document.body.getElementsByClassName('css-t46qbh')[" & i & "].click()" # click each row
			#delay 2
			#set positionData to do JavaScript "document.body.getElementsByClassName('css-w8sedg')[0].innerText" # grab position data
			#do JavaScript "history.back()"
			#delay 2
			#set end of stockDataList to paragraphs 2 thru 8 of stockData & paragraph 2 of positionData & paragraph 6 of positionData & paragraph 8 of positionData
			set end of stockDataList to paragraphs 2 thru 7 of stockData & 0.01 * (word 1 of paragraph 8 of stockData)
		end repeat
	end tell
end tell

tell application "Numbers"
	activate
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				#				make 3 new columns
				tell row 1
					set value of cell 1 to "Stock Name"
					set value of cell 2 to "Stock Ticker"
					set value of cell 3 to "Last Price"
					set value of cell 4 to "Daily Change Amount"
					set value of cell 5 to "Daily Change Percentage"
					set value of cell 6 to "Holdings Amount"
					set value of cell 7 to "Holdings Percentage"
					#set value of cell 8 to "All Time Gains"
					#set value of cell 9 to "Shares"
					#set value of cell 10 to "Average Price Paid"
				end tell
				
				repeat with i from 1 to count of stockDataList
					if not (exists row (i + 1)) then
						make row
					end if
					tell row (i + 1)
						set value of cell 1 to item 1 of item i of stockDataList
						set value of cell 1 to item 1 of item i of stockDataList
						set value of cell 2 to item 2 of item i of stockDataList
						set value of cell 3 to item 3 of item i of stockDataList
						set value of cell 4 to item 4 of item i of stockDataList
						set value of cell 5 to item 5 of item i of stockDataList
						set value of cell 6 to item 6 of item i of stockDataList
						set value of cell 7 to item 7 of item i of stockDataList
						#	set value of cell 8 to item 8 of item i of stockDataList
						#	set value of cell 9 to item 9 of item i of stockDataList
						#	set value of cell 10 to item 10 of item i of stockDataList
					end tell
				end repeat
				set format of column 7 to percent
			end tell
		end tell
	end tell
end tell