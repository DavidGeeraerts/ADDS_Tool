# Active Directory Domain Services Tool (ADDS_Tool)

![Main Banner](./images/ADDS_T_Main_Banner.png)

<h2 align="center"> :bangbang:  :construction:  :bangbang: UNDER DEVELOPMENT :bangbang:  :construction:  :bangbang: </h2> <br>


*Weekly build release*


## Table of Contents

- [Images](#Images)
- [Introduction](#introduction)
- [Dependencies](#Dependencies)
- [Features](#features)
- [Changelog](#Changelog)
- [License](#License)

## Images

![Main Menu](./images/Main_Menu_Local.png)


## Introduction

Windows Command shell program that is a wrapper for ADDS toolset:
	- `DSQUERY`
	- `DSGET`
	- `DSADD`
	- `DSMOD`
	- `DSMOVE`
	
The advantage of using ADDS Tool over something like [Active Directory Explorer](https://docs.microsoft.com/en-us/sysinternals/downloads/adexplorer):
 - Command shell is faster.
 - Navigating shell menu is quicker.
 - Every search is saved to log file; text files are easy to extract information from, and stores a historical search record.
 - ADDS Tool allows setting parameters, most important is the `-limit` parameter, which by default is set to `-limit 0`.

*Why not PowerShell?*
I like the windows command shell. It does most of what's needed. When the shell is lacking, PowerShell can be leveraged, which the program does.


## Dependencies

ADDS_Tool requires (RSAT) [Remote Server Administrative Tools](https://docs.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools). 
The program will ask to install if not detected, using [DISM](https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/what-is-dism). 
Must be runnig with administartive privilege in order to install RSAT. 

[PowerShell](https://docs.microsoft.com/en-us/powershell/scripting/overview): Leveraged for some functions.


## Features

 What's Working

*Currently only searching is working.*

- [X] Main Menu
- [X] Settings Menu
- [X] Logs
- [X] Search Universal
- [ ] Search User
- [X] Search Group
- [ ] Search Computer
- [X] Search Server
- [X] Search OU

## Changelog

See [changelog.md](changelog.md)


## License

[GPL](LICENSE) © David Geeraerts


:us:
