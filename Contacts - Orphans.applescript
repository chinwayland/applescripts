property orphan : "Orphans 2"
tell application "Contacts"
	try
		if group orphan exists then
			repeat with this_person in every person of group orphan
				remove this_person from group orphan
			end repeat
		else
			make new group at the end of groups with properties {name:orphan}
		end if
		save
		repeat with this_person in every person
			if number of groups of this_person = 0 then
				add this_person to group orphan
			end if
		end repeat
		save
	end try
end tell