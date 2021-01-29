# ChangeLog: ADDS (Active Directory Domain Services) Tool


## Features Heading
- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Fixed` for any bug fixes.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Deprecated` for soon-to-be removed features.

[//]: # (Copy paste pallette)
[//]: # (#### Added)
[//]: # (#### Changed)
[//]: # (#### Fixed)
[//]: # (#### Removed)
[//]: # (#### Security)
[//]: # (#### Deprecated)

---

## Version 0.10.0 Build: 2021-01-29
#### Changed
- Search by Group name updated with template
- Search by Group description Search  updated with template
- Search by Group DisplayName updated with template
- Search by Group multi attribute updated with template
- Search by Computer name updated with template

#### Removed
- mode con:lines outside of primary

---


## Version 0.9.0 Build: 2021-01-28
#### Changed
- Search by User global inactive updated with template
- Search by User global stalepassword updated with template
- Search by User global disabbled updated with template


## Version 0.8.0 Build: 2021-01-27
#### Changed
- Search by User displayName updated with template
- Search by User Custom Attributes updated with template

## Version 0.7.0 Build: 2021-01-26
#### Changed
- Search by User First and last name updated with template


## Version 0.6.0 Build: 2021-01-22
#### Added
- sAMaccount to Universal attribute list

#### Changed
- Search by User UPN updated with template
- search variables with "Last" prefix, dropped prefix. 

#### Removed
- Last Search type from Search main menu


## Version 0.5.0 Build: 2021-01-15
#### Added
- SHA256 file for commandlet

#### Changed
- Close last search log is now a subroutine
- Search by User name updated with template
- changed location for documents to /docs

#### Fixed
- formatting 


## Version 0.4.0 Build: 2021-01-08
#### Added
- Visual processing in Universal search

#### Changed
- Banner formatting
- User searchs navigation menu go back to user search
- Menu Logs now opens directory to logs 

#### Fixed
- Setting custom domain credentials


## Version 0.3.0 Build: 2021-01-04
#### Added
- Suppress verbose output. Off by default.

#### Changed
- Universal search will be a template for all other searches
- Code style for blocks
- use of [TimeSpan] instead of [dateTime] for stopwatch
- Log output header is now subroutine
- more powershell color text

#### Fixed
- Stop watch for searches
- User UPN search


## Version 0.2.0 Build: 2020-12-30
#### Added
- PID to variable debug
- User search

#### Changed
- echo processing... to powershell processing...
- Variable debug uses PID in file name


## Version 0.1.0 Build: 2020-12-24
#### Added
- Computer search
- Variable debug function

#### Changed
- color output by type
- $Counter to $COUNTER_SEARCH
- UTC gathering method
- Version in menu

#### Fixed
- local user query for Group

#### Removed
- timeout after no search results


## Version 0.0.0 Build: 2020-12-17
#### Added
- End search time
- Server (DC) search
- ISO8601 date stamp
- recall custom OU in search settings
- Total lapse time for search(s)
- Session log includes PID
- UTC to logs

#### Fixed
- Group search using description that had spaces in DN

#### Removed
- Universal search using "*" wildcard


## Version 0.0.0 Build: 2020-12-15
#### Added
- Search Universal attribute
- Group search by name, description, displayname

#### Changed
- Universal output to be consistent
- OU search and output to be consistent

#### Fixed
- log name typo
- Forestroot search for OU


## Version 0.0.0 Build: 2020-12-08
#### Added
- Under development

#### Fixed
- Crash's due to `GoTo:EOF` with no `Call`