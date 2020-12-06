-- This script scrapes flight data from Google Travel Flights


tell application "Safari"
	tell document 1
		#set numberOfDestinations to do JavaScript "document.getElementsByClassName('tsAU4e').length"
		set destinationData to {}
		set numberOfMystery to do JavaScript "document.getElementsByClassName('Xq1DAb').length"
		(*
		set duration to {}
		repeat with i from 1 to numberOfMystery - 1
			set end of duration to do JavaScript "document.getElementsByClassName('Xq1DAb')[" & i & "].textContent"
		end repeat
		*)
		repeat with i from 3 to numberOfMystery + 1
			set end of destinationData to {¬
				do JavaScript "document.getElementsByClassName('tsAU4e')[" & i & "].firstChild.firstChild.textContent", ¬
				do JavaScript "document.getElementsByClassName('tsAU4e ')[" & i & "].childNodes[1].firstChild.firstChild.firstChild.textContent", ¬
				do JavaScript "document.getElementsByClassName('Xq1DAb')[" & i - 2 & "].textContent"}
		end repeat
		
		#		do JavaScript "document.getElementsByClassName('tsAU4e ')[3].childNodes[1].firstChild.firstChild.firstChild.textContent"
	end tell
end tell