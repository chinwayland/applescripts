use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

tell application "Numbers"
	tell document 1
		tell sheet 1
			tell table 1
				set numberOfStudents to (count of cells in column 1) - 2
				tell row 1
					set labels to {value of cell 1, value of cell 3, value of cell 4, value of cell 9}
				end tell
				set studentData to {}
				repeat with i from 2 to numberOfStudents + 1
					tell row i
						set end of studentData to {value of cell 1, value of cell 3, value of cell 4, value of cell 9}
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell

tell application "Pages"
	activate
	make new document
	tell document 1
		repeat with i from 1 to numberOfStudents
			if not (exists page i) then
				make page
			end if
			tell page i
				make table with properties {column count:2, row count:18}
				tell table 1
					set position to {25, 25 + (792 * (i - 1))}
					tell column 1
						set width to 611 - 500 - 25
						repeat with j from 1 to count of labels
							set value of cell j to item j of labels
						end repeat
						set value of cell 5 to "Question 1"
						set value of cell 6 to "Answer (One Sentence)"
						set the text color of cell 6 to "red"
						set value of cell 7 to "Question 2"
						set value of cell 8 to "Answer (One Sentence)"
						set the text color of cell 8 to "red"
						set value of cell 9 to "Question 3"
						set value of cell 10 to "Answer (One Sentence)"
						set the text color of cell 10 to "red"
						set value of cell 11 to "Question 4 (Choose from Questions 4 - 12 on page 12)"
						set value of cell 12 to "Answer (Two Sentences)"
						set the text color of cell 12 to "red"
						set value of cell 13 to "Question 5 (Choose from Questions 4 - 12 on page 12)"
						set value of cell 14 to "Answer (Two Sentences)"
						set the text color of cell 14 to "red"
						set value of cell 15 to "Question 6 (Write your own question)"
						set value of cell 16 to "Answer (Two Sentences)"
						set the text color of cell 16 to "red"
						set value of cell 17 to "Question 7 (Write your own question)"
						set value of cell 18 to "Answer (Two Sentences)"
						set the text color of cell 18 to "red"
					end tell
					tell column 2
						set width to 611 - 111 - 25
						repeat with j from 1 to count of labels
							set value of cell j to item j of item i of studentData
						end repeat
						set alignment of cell 4 to left
						set value of cell 5 to "What's your full name? (Partner)"
						set value of cell 7 to "Have you got a nickname? (Partner)"
						set value of cell 9 to "Where and when were you born? (Partner)"
					end tell
				end tell
			end tell
		end repeat
	end tell
end tell