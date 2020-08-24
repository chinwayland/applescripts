use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This script scrapes jobs from craigslist and puts them into a Numbers document

set columnHeaders to {"Job Title", "Location", "Compensation", "Employment Type", "Internship", "Non-Profit", "Telecommuting", "Disability", "Relocation", "jobURL"}

tell application "Numbers"
	activate
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				tell row 1
					repeat with i from 1 to count of columnHeaders
						if (not (exists cell i)) then
							add column after the last column
						end if
						set the value of cell i to item i of columnHeaders
					end repeat
				end tell
				repeat with i from 5 to 9
					set the format of column i to checkbox
				end repeat
			end tell
		end tell
	end tell
end tell

tell application "Safari"
	activate
	make new document
	tell document 1
		set URL to "https://vietnam.craigslist.org/search/edu?cc=us&lang=en"
		repeat until source contains "feedback"
			delay 5
		end repeat
	end tell
	
end tell

tell application "Safari"
	tell document 1
		set totalCount to do JavaScript "document.getElementsByClassName('totalcount')[0].textContent"
		set numberOfPages to round totalCount / 120 rounding up
	end tell
end tell

repeat with pageNumber from 1 to numberOfPages
	repeat with i from 0 to 119
		tell application "Safari"
			tell document 1
				do JavaScript "document.getElementsByClassName('result-title hdrlnk')[" & i & "].click()"
				tell me to say "Clicked on job number " & i + 1
				repeat until source contains "post id"
					delay 5
				end repeat
				
				set jobTitle to do JavaScript "document.getElementById('titletextonly').textContent"
				set location to do JavaScript "document.getElementById('titletextonly').nextSibling.textContent"
				set postingBody to do JavaScript "document.getElementById('postingbody').textContent"
				set numberOfProperties to do JavaScript "document.getElementsByClassName('attrgroup')[0].children.length"
				set jobProperties to do JavaScript "document.getElementsByClassName('attrgroup')[0].textContent"
				set rowData to {}
				repeat with j from 0 to (numberOfProperties * 2) - 1
					
					try
						
						if (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.textContent") contains "compensation" then
							set compensation to (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.nextSibling.textContent")
							
							
						else if (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.textContent") contains "employment" then
							set employmentType to (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.nextSibling.textContent")
							
							
						else if (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.textContent") contains "internship" then
							set internship to true
							
						else if (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.nextSibling.textContent") contains "non-profit" then
							set nonProfit to true
							
							
						else if (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.textContent") contains "telecommuting" then
							set telecommuting to true
							
							
						else if (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.textContent") contains "disability" then
							set disability to true
							
						else if (do JavaScript "document.getElementsByClassName('attrgroup')[0].childNodes[" & j & "].firstChild.textContent") contains "relocation" then
							set relocation to true
						end if
					end try
				end repeat
				set jobURL to URL
			end tell
		end tell
		
		tell application "Numbers"
			tell document 1
				tell sheet 1
					tell table 1
						if (not (exists ((row (i + 2 + (120 * (pageNumber - 1))))))) then
							add row below last row
						end if
						tell application "System Events" to key code 125 using command down
						tell application "System Events" to key code 123 using command down
						
						tell row (i + 2 + (120 * (pageNumber - 1)))
							
							try
								set value of cell 1 to jobTitle
								set jobTitle to missing value
							end try
							
							try
								set value of cell 2 to location
								set location to missing value
							end try
							try
								set value of cell 3 to compensation
								set compensation to missing value
							end try
							
							try
								set value of cell 4 to employmentType
								set employmentType to missing value
							end try
							try
								set value of cell 5 to internship
								set format of cell 5 to checkbox
								set internship to missing value
							end try
							try
								set value of cell 6 to nonProfit
								set nonProfit to missing value
							end try
							try
								set value of cell 7 to telecommuting
								set telecommuting to missing value
							end try
							try
								set value of cell 8 to disability
								set disability to missing value
							end try
							try
								set value of cell 9 to relocation
								set relocation to missing value
							end try
							try
								set value of cell 10 to jobURL
								set jobURL to missing value
							end try
							(*
							try
								set value of cell 9 to jobProperties
								set jobProperties to missing value
							end try
							try
								set value of cell 10 to postingBody
								set postingBody to missing value
							end try
							*)
							
						end tell
					end tell
				end tell
			end tell
		end tell
		
		tell application "Safari"
			tell document 1
				do JavaScript "history.back()"
				delay 3
			end tell
		end tell
		
	end repeat
	
	tell application "Safari"
		tell document 1
			if pageNumber is less than numberOfPages then
				do JavaScript "document.getElementsByClassName('button next')[0].click()"
				repeat until source contains "feedback"
					delay 5
				end repeat
				tell me to say "Going to the next page"
			end if
		end tell
	end tell
end repeat