set columnList to {}
set rowList to {}

# Get entire table into a two dimensional list by column
tell application "Numbers"
	tell document 1
		tell sheet 1
			tell table 1
				repeat with columnNumber from 1 to the count of columns
					set end of columnList to {}
					tell column columnNumber
						repeat with rowNumber from 1 to the count of rows
							set the end of columnList's item columnNumber to the value of cell rowNumber
						end repeat
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell

# Get entire table into a two dimensional list by row
tell application "Numbers"
	tell document 1
		tell sheet 1
			tell table 1
				repeat with rowNumber from 1 to the count of rows
					set end of rowList to {}
					tell row rowNumber
						repeat with columnNumber from 1 to the count of columns
							set the end of rowList's item rowNumber to the value of cell columnNumber
						end repeat
					end tell
				end repeat
			end tell
		end tell
	end tell
end tell