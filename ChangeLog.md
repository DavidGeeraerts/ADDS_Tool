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

## Version 0.0.0 Build: 2020-12-
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


---

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