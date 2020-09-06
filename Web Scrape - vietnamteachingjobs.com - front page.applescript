-- Scrape jobs from vietnamteachingjobs.com -- front page only

#repeated click load button until there's nothing more to load

tell application "Safari"
	activate
	make new document
	tell document 1
		set URL to "https://vietnamteachingjobs.com"
		delay 10
		set clickCount to 0
		repeat until (do JavaScript "document.getElementsByClassName('wpjb-ls-load-more').length") is 0
			if do JavaScript "document.getElementsByClassName('wpjb-ls-load-more').length" is greater than 0 then
				try
					do JavaScript "document.getElementsByClassName('wpjb-ls-load-more')[0].firstChild.firstChild.click()"
				end try
				delay 1
				set clickCount to clickCount + 1
				log clickCount
			end if
		end repeat
		set jobCount to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childElementCount"
	end tell
end tell

-- Create new Numbers document
tell application "Numbers"
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				add column after last column
				tell row 1
					set value of cell 1 to "Job Title"
					set value of cell 2 to "Company Name"
					set value of cell 3 to "Location"
					set value of cell 4 to "Part or Full Time"
					set value of cell 5 to "Salary"
					set value of cell 6 to "Date Posted"
					set value of cell 7 to "Logo URL"
					set value of cell 8 to "URL"
				end tell
			end tell
		end tell
	end tell
end tell

-- Scrape data from each job on the front page
repeat with i from 0 to (jobCount - 1)
	tell application "Safari"
		tell document 1
			set title to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].childNodes[3].firstElementChild.text"
			set company to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].childNodes[3].childNodes[3].textContent"
			set location to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].childNodes[5].childNodes[2].textContent.trim()"
			set partTimeOrFullTime to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].childNodes[5].childNodes[3].textContent.trim()"
			set salary to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].childNodes[7].textContent.trim()"
			set datePosted to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].childNodes[9].childNodes[0].textContent.trim()"
			set logoURL to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].firstElementChild.firstElementChild.src"
			set jobURL to do JavaScript "document.getElementById('wpjb-job-list').firstElementChild.childNodes[" & i & "].childNodes[3].firstElementChild.href"
		end tell
	end tell
	
	tell application "Numbers"
		tell document 1
			tell sheet 1
				tell table 1
					
					if (not (exists row (i + 2))) then
						add row below last row
					end if
					
					tell row (i + 2)
						
						set value of cell 1 to title
						
						
						set value of cell 2 to company
						
						
						set value of cell 3 to location
						
						
						set value of cell 4 to partTimeOrFullTime
						
						
						set value of cell 5 to salary
						
						
						set value of cell 6 to datePosted
						
						
						set value of cell 7 to logoURL
						
						
						set value of cell 8 to jobURL
						
					end tell
				end tell
			end tell
		end tell
	end tell
end repeat

