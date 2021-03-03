-- This script changes the time zone on an event

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "CalendarLib EC" version "1.1.5" -- put this at the top of your scripts

set startDate to date "Monday, March 1, 2021 at 12:00:00 AM"
set endDate to date "Wednesday, July 7, 2021 at 12:00:00 AM"

set theStore to fetch store
set theEvents to fetch events starting date startDate ending date endDate searching cals {} event store theStore

set theEvents to modify zone event item 1 of theEvents time zone "Asia/Bangkok"

store event event theEvents event store theStore