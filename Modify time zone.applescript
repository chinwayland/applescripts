-- This script changes the time zone on an event

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "CalendarLib EC"

set startDate to (current date) - (4 * hours)
set endDate to startDate + (hours * 23)

set theStore to fetch store
set theEvents to fetch events starting date startDate ending date endDate searching cals {} event store theStore

set theEvents to modify zone event item 1 of theEvents time zone "Asia/Bangkok"

store event event theEvents event store theStore