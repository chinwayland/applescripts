-- This script scrapes jobs from tealit.com and puts them into a Numbers spresadsheet
set clickCounter to 0
set jobCount to 0

tell application "Safari" # open tealit website in safari
	make new document
	tell document 1
		set URL to "https://www.tealit.com/?view=section&sid=29"
		delay 5
	end tell
end tell

# find number of pages
tell application "Safari"
	tell document 1
		set totalPages to do JavaScript "document.getElementsByClassName('page-link')[4].parentElement.previousSibling.firstChild.textContent"
	end tell
end tell

tell application "Numbers" # create a new numbers spreadsheet, add columns and name the headers
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				repeat until (count of columns) is 14
					add column after last column
				end repeat
				
				tell row 1
					set value of cell 1 to "Listing Number"
					set value of cell 2 to "Company Name"
					set value of cell 3 to "Email"
					set value of cell 4 to "Phone"
					set value of cell 5 to "Available When"
					set value of cell 6 to "Education Required"
					set value of cell 7 to "Student Age"
					set value of cell 8 to "Students Per Class"
					set value of cell 9 to "Contact Person"
					set value of cell 10 to "Location"
					set value of cell 11 to "District"
					set value of cell 12 to "Street Address"
					set value of cell 13 to "Company URL"
					set value of cell 14 to "Tealit URL"
				end tell
			end tell
		end tell
	end tell
end tell

repeat with pageNumber from 1 to totalPages # Loop through all the pages
	
	tell application "Safari" # get the active page
		tell document 1
			set activePage to do JavaScript "document.getElementsByClassName('page-item active')[0].firstChild.textContent"
		end tell
	end tell
	say "on Page" & activePage
	say "page count is " & pageNumber
	
	tell application "Safari" # get the number of jobs on this page.
		tell document 1
			set pageJobCount to do JavaScript "document.getElementsByClassName('title headline-font').length"
		end tell
	end tell
	
	repeat with i from 0 to pageJobCount - 1 #Go through each job per page and scrape data
		tell application "Safari"
			tell document 1
				set listingNumber to do JavaScript "document.getElementsByClassName('listing-num')[" & i & "].textContent"
				do JavaScript "document.getElementsByClassName('title headline-font')[" & i & "].firstChild.click()"
				delay 3
				set jobCount to jobCount + 1
				say "job count" & jobCount
				
				set companyName to (do JavaScript "document.getElementsByClassName('title company-name')[0].textContent")
				try
					say companyName
				end try
				set email to do JavaScript "document.getElementsByClassName('fa fa-envelope')[0].nextSibling.nextSibling.textContent"
				set phoneNumber to do JavaScript "document.getElementsByClassName('fa fa-phone')[0].nextSibling.textContent"
				set tealitURL to URL
				-- Position of data can change, so loop through them to find titles, then grab data
				set numberOfDataItems to do JavaScript "document.getElementsByClassName('inline-title').length"
				repeat with j from 0 to numberOfDataItems - 1
					set dataHolder to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].textContent"
					if dataHolder is "When Job is Available" then
						set jobAvailableWhen to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "Degree Required" then
						set educationRequired to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "Age Level Of Class" then
						set studentAge to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "Average Number of Students in Class" then
						set averageNumberOfStudentsPerClass to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "Contact Person" then
						set contactPerson to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "Location" then
						set location to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "District" then
						set district to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "Street Address" then
						set streetAddress to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.textContent"
					else if dataHolder is "Web Site URL" then
						set schoolURL to do JavaScript "document.getElementsByClassName('inline-title')[" & j & "].nextElementSibling.nextElementSibling.nextElementSibling.textContent"
					end if
				end repeat
				
				do JavaScript "history.back()"
				delay 3
			end tell
		end tell
		
		tell application "Numbers" # paste data into Numbers spreadsheet
			tell document 1
				tell sheet 1
					tell table 1
						set rowNumber to ((pageNumber - 1) * 30) + (i + 2)
						if not (exists row rowNumber) then
							add row below last row
						end if
						tell row rowNumber
							try
								set value of cell 1 to listingNumber
							end try
							try
								set value of cell 2 to companyName
							end try
							try
								set value of cell 3 to email
							end try
							try
								set value of cell 4 to phoneNumber
							end try
							try
								set value of cell 5 to jobAvailableWhen
							end try
							try
								set value of cell 6 to educationRequired
							end try
							try
								set value of cell 7 to studentAge
							end try
							try
								set value of cell 8 to averageNumberOfStudentsPerClass
							end try
							try
								set value of cell 9 to contactPerson
							end try
							try
								set value of cell 10 to location
							end try
							try
								set value of cell 11 to district
							end try
							try
								set value of cell 12 to streetAddress
							end try
							try
								set value of cell 13 to schoolURL
							end try
							try
								set value of cell 14 to tealitURL
							end try
						end tell
					end tell
				end tell
			end tell
		end tell
	end repeat
	
	repeat with j from 0 to 7 #Click on the next page
		tell application "Safari"
			tell document 1
				try
					if (do JavaScript "document.getElementsByClassName('page-link')[" & j & "].textContent") is "Ý" then
						do JavaScript "document.getElementsByClassName('page-link')[" & j & "].click()"
						delay 3
						say "clicked to go to the next page"
						set clickCounter to clickCounter + 1
						exit repeat
					end if
				end try
			end tell
		end tell
	end repeat
end repeat