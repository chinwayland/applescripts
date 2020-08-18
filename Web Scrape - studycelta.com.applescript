use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- This script scrapes data from studycelta.com

set timeZone to {"GMT -7 to -3 Americas ", "GMT 0 to + 2 Europe-Africa", "GMT +3 to + 6 Central Asia", "GMT + 7 to + 10 East Asia-Pacific"}


tell application "Safari"
	set centerData to {}
	tell window 1
		
		# if the internet is fast, the code below cycles through all the pages
		(*
		set URL to "https://www.studycelta.com/tefl-courses-online/online-celta-courses/"
		delay 3
		set urls to {}
		repeat with i from 0 to ((do JavaScript "document.getElementsByClassName('ow-button-hover').length") - 1)
			set end of urls to do JavaScript "document.getElementsByClassName('ow-button-hover')[" & i & "].href"
		end repeat
				
		repeat with i from 1 to count of urls
			set URL to item i of urls
			delay 3
		end repeat
				
		*)
		repeat with safariTab from 1 to count of tabs
			tell tab safariTab
				# i = class="celta_course_table" = coursetype 		
				# j = class="celta_course_location_table"
				# i loops through the course types or blue table sections
				set numberOfCourseTypes to (do JavaScript "document.getElementsByClassName('celta_course_table').length")
				repeat with i from 0 to (numberOfCourseTypes - 1)
					set courseType to do JavaScript "document.getElementsByClassName('celta_course_table')[" & i & "].firstChild.firstChild.firstChild.textContent"
					# j loops through each individual center
					set numberOfCenters to (do JavaScript "document.getElementsByClassName('celta_course_table')[" & i & "].childNodes[1].firstChild.firstChild.childNodes.length")
					repeat with j from 0 to (numberOfCenters - 1)
						
						# k loops through each row or data item of each center
						set numberOfRows to (do JavaScript "document.getElementsByClassName('celta_course_table')[" & i & "].childNodes[1].firstChild.firstChild.childNodes[" & j & "].firstChild.firstChild.childElementCount")
						set centerCity to do JavaScript "document.getElementsByClassName('celta_course_table')[" & i & "].childNodes[1].firstChild.firstChild.childNodes[" & j & "].firstChild.firstChild.childNodes[0].firstChild.firstChild.textContent"
						set centerPrice to do JavaScript "document.getElementsByClassName('celta_course_table')[" & i & "].childNodes[1].firstChild.firstChild.childNodes[" & j & "].firstChild.firstChild.childNodes[" & (numberOfRows - 1) & "].firstChild.firstChild.textContent"
						set centerDetails to do JavaScript "document.getElementsByClassName('celta_course_table')[" & i & "].childNodes[1].firstChild.firstChild.childNodes[" & j & "].firstChild.firstChild.childNodes[" & (numberOfRows - 1) & "].firstChild.childNodes[3].textContent"
						
						repeat with k from 1 to (numberOfRows - 2)
							set centerDate to (do JavaScript "document.getElementsByClassName('celta_course_table')[" & i & "].childNodes[1].firstChild.firstChild.childNodes[" & j & "].firstChild.firstChild.childNodes[" & k & "].firstChild.firstChild.textContent")
							set end of centerData to {item safariTab of timeZone, courseType, centerCity, centerPrice, centerDate}
						end repeat
					end repeat
				end repeat
			end tell
			
		end repeat
	end tell
end tell

tell application "Numbers"
	activate
	make new document
	tell document 1
		tell sheet 1
			tell table 1
				tell row 1
					set value of cell 1 to "Time Zone"
					set value of cell 2 to "Course Type"
					set value of cell 3 to "Course City"
					set value of cell 4 to "Course Price"
					set value of cell 5 to "Course Date Period"
					#set value of cell 6 to "Course Details"
				end tell
				
				repeat with i from 1 to count of centerData
					if (not (exists row (i + 1))) then
						add row below last row
					end if
					tell row (i + 1)
						repeat with j from 1 to count of item i of centerData
							set value of cell j to item j of item i of centerData
						end repeat
						
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell