-- This script changes the time zone on an event

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "CalendarLib EC" version "1.1.5" -- put this at the top of your scripts

tell application "Calendar"
	set theCalendars to name of calendars whose name contains "year"
end tell

set theStore to fetch store

set startDate to date "Monday, March 1, 2021 at 12:00:00 AM"
set endDate to date "Saturday, March 6, 2021 at 12:00:00 AM"

set theCalendars to fetch calendars theCalendars cal type list cal local event store theStore

set theEvents to fetch events starting date startDate ending date endDate searching cals theCalendars event store theStore

repeat with i from 1 to count of theEvents
	set theNewEvent to modify zone event item i of theEvents time zone "Asia/Shanghai"
	store event event theNewEvent event store theStore
end repeat


