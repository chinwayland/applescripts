use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- Scrape jobs from vietnamteachingjobs.com

#repeated click load button until there's nothing more to load

tell application "Safari"
	activate
	make new document
	tell document 1
		set URL to "https://vietnamteachingjobs.com"
		delay 3
		set clickCount to 0
		set loadButton to "Load 22 more"
		repeat until (do JavaScript "document.getElementsByClassName('wpjb-ls-load-more').length") is 0
			if do JavaScript "document.getElementsByClassName('wpjb-ls-load-more').length" is greater than 0 then
				do JavaScript "document.getElementsByClassName('wpjb-ls-load-more')[0].firstChild.firstChild.click()"
				delay 1
				set clickCount to clickCount + 1
				# tell application "System Events" to key code 125 using command down
			end if
		end repeat
	end tell
end tell

#grab strings
set theString to {}
tell application "Safari"
	tell document 1
		set URLCount to do JavaScript "document.getElementsByClassName('wpjb-column-title').length"
		repeat with i from 0 to URLCount - 1
			set end of theString to do JavaScript "document.getElementsByClassName('wpjb-column-title')[" & i & "].innerHTML"
		end repeat
	end tell
end tell

#More processing of the strings to get the list of URLs
set theString2 to {}
repeat with i in theString
	set end of theString2 to paragraph 2 of i
end repeat

set theString3 to {}
repeat with i in theString2
	set end of theString3 to decoupe(i, {"\"", "\""})
end repeat

set theURLs to {}
repeat with i in theString3
	try
		set end of theURLs to item 2 of i
	end try
end repeat

tell application "Numbers"
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				repeat 5 times
					add column after last column
				end repeat
				tell row 1
					set value of cell 1 to "Job Title"
					set value of cell 2 to "Company Name"
					set value of cell 3 to "Location"
					set value of cell 4 to "Start Date"
					set value of cell 5 to "Category"
					set value of cell 6 to "Job Type"
					set value of cell 7 to "Nationality of Teacher"
					set value of cell 8 to "Teaching Experience"
					set value of cell 9 to "Candidate Requirements"
					set value of cell 10 to "Salary"
					set value of cell 11 to "URL"
					set value of cell 12 to "Job Order"
				end tell
			end tell
		end tell
	end tell
end tell

repeat with i from 1 to count of theURLs
	tell application "Safari"
		tell document 1
			set URL to item i of theURLs
			delay 3
			set theTitle to do JavaScript "document.getElementsByClassName('entry-title')[0].textContent"
			set theCompany to do JavaScript "document.getElementsByClassName('wpjb-job-company')[0].text"
			repeat with j from 0 to 9
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Date Posted" then
						set datePosted to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.textContent.trim()"
					end if
				end try
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Category" then
						set category to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.innerText.trim()"
					end if
				end try
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Job Type" then
						set jobType to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.childNodes[3].firstElementChild.textContent"
					end if
				end try
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Nationality Of teacher" then
						set nationalityOfTeacher to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.textContent.trim()"
					end if
				end try
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Teaching Experience" then
						set teachingExperience to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.textContent.trim()"
					end if
				end try
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Candidate Requirements" then
						set candidateRequirements to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.textContent.trim()"
					end if
				end try
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Where is the school" then
						set location to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.textContent.trim()"
					end if
				end try
				try
					if (do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].textContent.trim()") contains "Salary:" then
						set salary to do JavaScript "document.getElementsByClassName('wpjb-info-label')[" & j & "].nextElementSibling.textContent.trim()"
					end if
				end try
				
			end repeat
			
			set theURL to URL
		end tell
	end tell
	
	tell application "Numbers"
		tell document 1
			tell sheet 1
				tell table 1
					if (not (exists row (i + 1))) then
						add row below last row
					end if
					tell application "System Events" to key code 125 using command down
					tell row (i + 1)
						try
							set value of cell 1 to theTitle
						end try
						try
							set value of cell 2 to theCompany
						end try
						try
							set value of cell 3 to location
						end try
						try
							set value of cell 4 to datePosted
						end try
						try
							set value of cell 5 to category
						end try
						try
							set value of cell 6 to jobType
						end try
						try
							set value of cell 7 to nationalityOfTeacher
						end try
						try
							set value of cell 8 to teachingExperience
						end try
						try
							set value of cell 9 to candidateRequirements
						end try
						try
							set value of cell 10 to salary
						end try
						try
							set value of cell 11 to theURL
						end try
						try
							set value of cell 12 to i
						end
					end tell
				end tell
			end tell
		end tell
	end tell
end repeat

on decoupe(t, d)
	local oTIDs, l
	set {oTIDs, AppleScript's text item delimiters} to {AppleScript's text item delimiters, d}
	set l to text items of t
	set AppleScript's text item delimiters to oTIDs
	return l
end decoupe