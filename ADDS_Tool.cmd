:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Author:		David Geeraerts
:: Location:	Olympia, Washington USA
:: E-Mail:		dgeeraerts.evergreen@gmail.com
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Copyleft License(s)
:: GNU GPL (General Public License)
:: https://www.gnu.org/licenses/gpl-3.0.en.html
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::::::::::::::::::::::::::::::::
:: VERSIONING INFORMATION		::
::  Semantic Versioning used	::
::   http://semver.org/			::
::	Major.Minor.Revision		::
::::::::::::::::::::::::::::::::::

::#############################################################################
::							#DESCRIPTION#
::
::	SCRIPT STYLE: Interactive
::	Program is a wrapper for ADDS (Active Directory Domain Services)
::	Active Directory search's
::#############################################################################

@Echo Off
@SETLOCAL enableextensions
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
@SET $START_LOAD_TIME=%TIME%
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


SET $PROGRAM_NAME=Active_Directory_Domain_Services_Tool
SET $Version=0.2.0
SET $BUILD=2020-12-30 10:00
Title %$PROGRAM_NAME%
Prompt ADT$G
color 8F
mode con:cols=80
mode con:lines=50

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Declare Global variables
:: All User variables are set within here.
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Defaults
::	uses user profile location for logs
SET "$LOGPATH=%APPDATA%\ADDS"
SET $SESSION_LOG=ADDS_Tool_Active_Session.log
SET $SEARCH_SESSION_LOG=ADDS_Tool_Session_Search.log
SET $LAST_SEARCH_LOG=ADDS_Tool_Last_Search.log
SET $ARCHIVE_LOG=ADDS_Tool_Session_Archive.log
SET $ARCHIVE_SEARCH_LOG=ADDS_Tool_Search_Archive.log

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Advanced Settings
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::	Suppress_Console_Threshold
::	too many results isn't useful to display
SET $SUPPRESS_CONSOLE_THRESHOLD=3

:: Sort --the search results
:: {0 [No] , 1 [Yes]}
SET $SORTED=1

::	Keep all logs
::	{Yes, No}
SET $KPLOG=Yes

::	Keep Session Settings
::	{Yes, No}
SET $SAVE_SETTINGS=Yes

::	Load Settings --from file
:: {0 [Off/No] , 1 [On/Yes]}
SET $LOAD_SETTINGS=Yes

:: DEBUG
:: {0 [Off/No] , 1 [On/Yes]}
SET $DEGUB_MODE=1
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::##### Everything below here is 'hard-coded' [DO NOT MODIFY] #####
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:: Program Variables
::	Defaults
SET $COUNTER_SEARCH=0
SET $sLimit=0
::	Domain User status
::	0 - Local User , 1 - Domain User
SET $DU=1
SET $DC=%USERDOMAIN%
SET $SESSION_USER=%USERNAME%
SET $DOMAIN_USER=NA
SET $cUSERNAME=
SET $adgroup.n=NA
SET $LAST_SEARCH_TYPE=NA
SET $LAST_SEARCH_KEY=NA
SET $LAST_SEARCH_COUNT=NA
SET $DOMAIN=%USERDNSDOMAIN%
SET $DSITE=Default
IF NOT DEFINED $DOMAIN SET $DOMAIN=NA
REM Doesn't like On Off words
IF %$SORTED% EQU 1 (SET $SORTED_N=Yes) ELSE (SET $SORTED_N=No)
SET $SUPPRESS_CONSOLE_THRESHOLD=3
SET $SEARCH_SETTINGS_CHECK=0
SET $LAST_SEARCH_ATTRIBUTE=name
:: Defaults
SET $AD_BASE=domainroot
SET $AD_SCOPE=subtree
SET "$AD_SERVER_SEARCH=-s %$DC%"
:: Dependency Checks
::	assumes ready to go
SET $PREREQUISITE_STATUS=1
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:CD
	:: Launched from directory
	SET "$PROGRAM_PATH=%~dp0"
	::	Setup logging
	IF NOT EXIST "%$LOGPATH%\var" MD "%$LOGPATH%\var"
	cd /D "%$LOGPATH%"
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:PID
	:: Program information including PID
	tasklist /FI "WINDOWTITLE eq %$PROGRAM_NAME%*" > "%$LogPath%\var\var_TaskInfo_PID.txt"
	for /F "skip=3 tokens=2 delims= " %%P IN ('tasklist /FI "WINDOWTITLE eq %$PROGRAM_NAME%*"') DO echo %%P> "%$LogPath%\var\var_$PID.txt"
	SET /P $PID= < "%$LogPath%\var\var_$PID.txt"
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:fISO8601
	:: Function to ensure ISO 8601 Date format yyyy-mmm-dd
	:: Easiest way to get ISO date
	@powershell Get-Date -format "yyyy-MM-dd" > "%$LogPath%\var\var_ISO8601_Date.txt"
	SET /P $ISO_DATE= < "%$LogPath%\var\var_ISO8601_Date.txt"
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:UTC
	:: Universal Time Coordinate
	IF EXIST "%$LogPath%\var\var_$UTC.txt" SET /P $UTC= < "%$LogPath%\var\var_$UTC.txt"
	IF NOT DEFINED $UTC FOR /F "tokens=1 delims=()" %%P IN ('wmic timezone get Description ^| findstr /C:"UTC" /I') DO ECHO %%P > "%$LogPath%\var\var_$UTC.txt"
	IF NOT DEFINED $UTC SET /P $UTC= < "%$LogPath%\var\var_$UTC.txt"
	IF EXIST "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt" SET /P $UTC_STANDARD_NAME= < "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt"
	IF NOT DEFINED $UTC_STANDARD_NAME FOR /F "tokens=2 delims==" %%P IN ('wmic timezone get StandardName /value ^| findstr /C:"=" /I') DO ECHO %%P > "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt"
	IF NOT DEFINED $UTC_STANDARD_NAME SET /P $UTC_STANDARD_NAME= < "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt"
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:wLog
	:: Start session and write to log
	Echo Start Session %DATE% %TIME% > "%$LogPath%\%$SESSION_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Program Name: %$PROGRAM_NAME% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Program Version: %$Version% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Program Build: %$BUILD% >> "%$LogPath%\%$SESSION_LOG%"
	Echo PC: %COMPUTERNAME% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Session User: %USERNAME% >> "%$LogPath%\%$SESSION_LOG%"
	echo PID: %$PID% >> "%$LogPath%\%$SESSION_LOG%"
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:DUC
	::	Check for Domain computer
	::	If value is 1 domain, is 0 workgroup
	SET $DOMAIN_PC=1
	wmic computersystem get DomainRole /value | (FIND "0") && (SET $DOMAIN_PC=0)
	echo Domain_PC: %$DOMAIN_PC% >> "%$LogPath%\%$SESSION_LOG%"
	if %$DOMAIN_PC% EQU 0 SET $DOMAIN=%COMPUTERNAME%
	if %$DOMAIN_PC% EQU 1 SET $DOMAIN=%USERDNSDOMAIN%
	:: Can be local or domain
	IF %$DOMAIN_PC% EQU 0 SET $SESSION_USER_STATUS=local
	IF %$DOMAIN_PC% EQU 0 GoTo skipDUC

	:: Is domain user or local user?
	whoami /UPN 2> nul || FOR /F "tokens=1-2 delims=\" %%P IN ('whoami') Do SET $DOMAIN=%%P && SET $DU=0
	IF %$DU% EQU 0 SET $SESSION_USER_STATUS=local
	IF %$DU% EQU 0 GoTo skipDUC

	:: If domain user use UPN to set domain instead of USERDNSDOMAIN
	FOR /F "tokens=2 delims=^@" %%P IN ('whoami /UPN') Do SET $DOMAIN=%%P
	::	Default credentials for Active Directory is current logged on user.
	::	can be changed in console program
	SET $DC=%LOGONSERVER:~2%
	SET $DOMAIN_USER=%USERNAME%
	SET $SESSION_USER_STATUS=domain
:skipDUC

	:: Friendly name
	if %$DOMAIN_PC% EQU 0 SET $DOMAIN_PC_N=workgroup
	if %$DOMAIN_PC% EQU 1 SET $DOMAIN_PC_N=domain
	echo workgroup else echo domain)

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::	Administrator Privilege Check
:subA
	openfiles.exe 1> "%$LOGPATH%\var\var_$Admin_Status_M.txt" 2> "%$LOGPATH%\var\var_$Admin_Status_E.txt"
	SET $ADMIN_STATUS=0
	FIND "ERROR:" "%$LOGPATH%\var\var_$Admin_Status_E.txt" && (SET $ADMIN_STATUS=1)
	IF %$ADMIN_STATUS% EQU 0 (SET "$ADMIN_STATUS_N=Yes") ELSE (SET "$ADMIN_STATUS_N=No")
	echo %$ADMIN_STATUS_N%> "%$LOGPATH%\var\var_$Admin_Status_N.txt" 
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:CheckRSAT
	::	Check RSAT-Remote Server Administration Tools
	dsquery /? 1> nul 2> nul
	SET $RSAT_STATUS=%ERRORLEVEL%
	IF %$RSAT_STATUS% EQU 0 GoTo Start
	:: Admin privileges required
	IF %$ADMIN_STATUS% NEQ 0 GoTo err10
	::	Remote Server Administrator Tools Message
	::	Requires RSAT-Remote Server Administration Tools
	mode con:cols=80 lines=40
	cls
	color 8B
	SET PREREQUISITE_STATUS=0
	Echo.
	Echo It appears this computer [%COMPUTERNAME%] doesn't have:
	Echo [RSAT] Remote Server Administration Tools installed!
	Echo.
	Echo Try installing "Remote Server Admin Tool --> AD DS and AD LDS Tools".
	Echo Reference for doing this:
	Echo https://support.microsoft.com/en-us/help/2693643/remote-server-administration-tools-rsat-for-windows-operating-systems
	Echo.
	echo Install RSAT?
	set /p $INSTALL_DEPENDENCY=[Y]es or [N]o:
	echo.
	echo Selected: %$INSTALL_DEPENDENCY%
	IF NOT DEFINED $INSTALL_DEPENDENCY GoTo skipRSAT
	IF /I "%$INSTALL_DEPENDENCY%"=="N" GoTo skipRSAT
	GoTo RSAT
	Echo.
:skipRSAT
Timeout /t 10
GoTo end
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:RSAT
	:: Requires RSAT-Remote Server Administration Tools
	::	Active Directory Domain Services (AD DS) tools and Active Directory Lightweight Directory Services (AD LDS) tools
	::	Requires administrative privileges!
	::	Rsat.ActiveDirectory.DS-LDS.Tools
	FOR /F "tokens=2 delims=:" %%P IN ('DISM /Online /Get-Capabilities ^| find /I "Rsat.ActiveDirectory.DS-LDS.Tools"') DO SET $RSAT_ADDS_FULL=%%P
	::	remove leading space
	FOR /F "tokens=1 delims= " %%P IN ("%$RSAT_ADDS_FULL%") DO SET $RSAT_ADDS_FULL=%%P
	echo %$RSAT_ADDS_FULL%
	::	RSAT Installation
	ECHO Going to try to install RSAT (Remote Server Administration Tools)...
	echo.
	DISM /Online /add-capability /CapabilityName:%$RSAT_ADDS_FULL%
	SET $DISM_ERR=%ERRORLEVEL%
	echo $DISM_ERR:%DISM_STATUS%
	::after install check
	DISM /Online /Get-CapabilityInfo /CapabilityName:%$RSAT_ADDS_FULL% > "%$LOGPATH%\var\var_DISM_ADDS.txt"
	FIND /I "State : Installed" "%$LOGPATH%\var\var_DISM_ADDS.txt"
	type "%$LOGPATH%\var\var_DISM_ADDS.txt" >> "%$LogPath%\%$SESSION_LOG%"
	SET $RSAT_STATUS=%ERRORLEVEL%
	IF %$RSAT_STATUS% NEQ 0 SET $PREREQUISITE_STATUS=%$RSAT_STATUS%
	IF %$RSAT_STATUS% NEQ 0 GoTo err20
	IF %$PREREQUISITE_STATUS% EQU 0 GoTo Start
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:Start
:: Capture program load time
	@PowerShell.exe -c "$span=([datetime]'%Time%' - [datetime]'%$START_LOAD_TIME%'); '{0:00}:{1:00}:{2:00}' -f $span.Hours, $span.Minutes, $span.Seconds" > "%$LogPath%\var\var_Load_Time.txt"
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:Menu
	Color 0F
	mode con:cols=58 lines=40
	Cls
	ECHO *********************************************************
	ECHO		%$PROGRAM_NAME%
	echo			Version: %$Version%
	IF %$DEGUB_MODE% EQU 1 Echo			Build: %$BUILD%
	echo.
	echo		 	%DATE% %TIME%
	ECHO.
	Echo		Location: Main Menu   
	ECHO *********************************************************
	Echo.
	ECHO Session Information:
	Echo ------------------------
	echo  Session User: %USERNAME%
	echo  Session Admin Privilege: %$ADMIN_STATUS_N%
	echo  Session User Status: %$SESSION_USER_STATUS%
	echo  Session PC: %COMPUTERNAME%
	echo  Session PC Role: %$DOMAIN_PC_N%
	echo.
	echo Log Settings:
	Echo ------------------------
	Echo  Log File Path: %$LogPath%
	Echo  Log File Name: %$SESSION_LOG%
	Echo  Keep Log at End: %$kpLog%
	Echo.
	Echo Current Domain settings:
	Echo ------------------------
	Echo  Domain Account: %$DOMAIN_USER%
	Echo  Domain: %$DOMAIN%
	Echo  Domain Controller: %$DC%
	echo  Domain Site: %$DSITE%
	ECHO *********************************************************
	Echo.
	Echo Choose an action to perform from the list:
	Echo.
	Echo [1] Search
	Echo [2] Settings
	Echo [3] Logs
	Echo [4] Exit
	Echo.
	Choice /c 1234
	Echo.
	::
	If ERRORLevel 4 GoTo End
	If ERRORLevel 3 GoTo Logs
	If ERRORLevel 2 GoTo Uset
	If ERRORLevel 1 GoTo Search
	Echo.
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SMB
	::	Search Menu Banner
	Color 0A
	mode con:cols=55 lines=40
	Cls
	ECHO ******************************************************
	ECHO		%$PROGRAM_NAME%
	echo.
	echo		 	%DATE% %TIME%
	ECHO.
	Echo		Location: Search Menu     
	echo.
	Echo ******************************************************
	Echo.
	Echo Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo  Query limit: %$sLimit%
	echo  Sorted: %$SORTED_N%
	Echo  Last Search Type: %$LAST_SEARCH_TYPE%
	Echo  Search count: %$COUNTER_SEARCH%
	Echo ******************************************************
	Echo.
	GoTo:EOF
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:Search
	Color 0A
	mode con:lines=40
	:: Trap: Domain User check
	echo %COMPUTERNAME% | (FIND /I "%$DOMAIN%") && (GoTo :subDomain)
	IF %$DU% EQU 0 call :subDA
	IF /I "%$DC%"=="%COMPUTERNAME%" call :subDC
	call :SMB
	Echo Choose search type from the list:
	Echo.
	Echo [1] Universal
	Echo [2] User
	Echo [3] Group
	ECho [4] Computer
	echo [5] Server ^(DC's^)
	echo [6] OU
	echo [7] Settings Menu
	echo [8] Main Menu
	Echo [9] Exit
	Echo.
	Choice /c 123456789
	Echo.
	If ERRORLevel 9 GoTo end
	If ERRORLevel 8 GoTo Menu
	If ERRORLevel 7 GoTo Uset
	If ERRORLevel 6 GoTo sOU
	If ERRORLevel 5 GoTo sServer
	If ERRORLevel 4 GoTo sComputer
	If ERRORLevel 3 GoTo sGroup
	If ERRORLevel 2 GoTo sUser
	If ERRORLevel 1 GoTo sUniversal
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SM
	:: Search Menu banner
	cls
	ECHO ******************************************************
	ECHO		%$PROGRAM_NAME%
	echo.
	echo		 	%DATE% %TIME%
	ECHO.
	Echo		Location: %$LAST_SEARCH_TYPE% Search
	Echo ******************************************************
	echo.
	Echo Domain Settings
	Echo ------------------------
	Echo  Domain Running Account: %$DOMAIN_USER%
	Echo  Domain Controller: %$DC%
	echo  Domain Site: %$DSITE%
	Echo  Domain: %$domain%
	echo.
	Echo Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo  Query limit: %$sLimit%
	echo  Sorted: %$SORTED_N%
	echo.
	echo Search HUD
	Echo ------------------------
	Echo  Search Type: %$LAST_SEARCH_TYPE%
	echo  Search Attribute: %$LAST_SEARCH_ATTRIBUTE%
	echo  Search Key: %$LAST_SEARCH_KEY%
	echo  Search Results: %$LAST_SEARCH_COUNT%
	Echo  Search count: %$COUNTER_SEARCH%
	Echo ******************************************************
	echo.
	GoTo:EOF
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subSK

:: Sub-routin for Search Key
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subSET
	:: Start Elapse Time
	SET $START_TIME=%TIME%
	GoTo:EOF
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subTLT
	:: Total Lapse Time
	@PowerShell.exe -c "$span=([datetime]'%Time%' - [datetime]'%$START_TIME%'); '{0:00}:{1:00}:{2:00}' -f $span.Hours, $span.Minutes, $span.Seconds" > "%$LogPath%\var\var_Total_Lapsed_Time.txt"
	SET /P $TOTAL_LAPSE_TIME= < "%$LogPath%\var\var_Total_Lapsed_Time.txt"
	GoTo:EOF
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:sUniversal
	:: Search Universal
	SET $LAST_SEARCH_TYPE=Universal
	call :SM
	SET $SEARCH_KEY=
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
:SUC
	echo Choose attribute to search against:
	echo ^(default is name^; leave blank for default^)
	SET /P $LAST_SEARCH_ATTRIBUTE=Attribute:
	IF NOT DEFINED $LAST_SEARCH_ATTRIBUTE SET $LAST_SEARCH_ATTRIBUTE=name
	call :SM
	echo use "*" wildcard, e.g. Key*, *key*
	echo If left blank, will abort.
	SET $SEARCH_KEY=
	SET /P $SEARCH_KEY=Choose a search key:
	IF NOT DEFINED $SEARCH_KEY GoTo Search
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY%
	call :SM
	echo Selected {%$SEARCH_KEY%} as search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Start search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Search Term: %$SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: No point in sorting since it won't match the details from attr *
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$LAST_SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" %$AD_SERVER_SEARCH% -attr name distinguishedName > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$LAST_SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$LAST_SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" %$AD_SERVER_SEARCH% -attr distinguishedName > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$LAST_SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)		
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%K > "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"	
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"	
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUC
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUL
	FOR /F "USEBACKQ skip=1 tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr distinguishedName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
GoTo SkipSUL	
	:jumpSUL
	:: Session user is a local user
	FOR /F "USEBACKQ skip=1 tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr name %$AD_SERVER_SEARCH%  -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr description %$AD_SERVER_SEARCH%  -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr displayName %$AD_SERVER_SEARCH%  -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr distinguishedName %$AD_SERVER_SEARCH%  -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%N)" -attr * %$AD_SERVER_SEARCH%  -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:SkipSUL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
:jumpSUC	
	echo Search Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Menu
	IF %ERRORLEVEL% EQU 1 GoTo sUniversal
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:sUser
	:: Search User
	SET $LAST_SEARCH_TYPE=User
	SET $LAST_SEARCH_ATTRIBUTE=
	call :SM
	SET $SEARCH_KEY=
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul

	echo User search using:
	echo.
	echo [1] Name
	echo [2] UPN
	echo [3] First and Last name
	echo [4] Display name
	echo [5] Custom attribute search 
	echo [6] Global
	echo [7] Abort
	echo.
	Choice /c 1234567
	echo.
	If ERRORLevel 7 GoTo Search
	If ERRORLevel 6 GoTo SUG
	If ERRORLevel 5 GoTo SUCA
	If ERRORLevel 4 GoTo SUDN
	If ERRORLevel 3 GoTo SUFL
	If ERRORLevel 2 GoTo SUU
	If ERRORLevel 1 GoTo SUN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SUN
	:: Search User Name
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	call :SM
	@powershell Write-Host "Name" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort." -ForegroundColor Red
	SET $LAST_SEARCH_ATTRIBUTE=name
	SET $SEARCH_KEY_USER_NAME=
	SET /P $SEARCH_KEY_USER_NAME=User name search key:
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_USER_NAME%
	IF NOT DEFINED $SEARCH_KEY_USER_NAME GoTo sUser
	call :SM
	echo Selected {%$SEARCH_KEY_USER_NAME%} as name search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key Name: %$SEARCH_KEY_USER_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_NAME%"-limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSUNS
:jumpSUNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSUNS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUNO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	

	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUNL
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUNO
:jumpSUNL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUNO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUN
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"


:skipSUN
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo SUN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	

:SUU
	:: Search User UPN
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	call :SM
	@powershell Write-Host "UPN" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort." -ForegroundColor Red
	SET $LAST_SEARCH_ATTRIBUTE=UPN
	SET $SEARCH_KEY_USER_UPN=
	SET /P $SEARCH_KEY_USER_UPN=User upn search key:
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_USER_NAME%
	IF NOT DEFINED $SEARCH_KEY_USER_UPN GoTo sUser
	call :SM
	echo Selected {%$SEARCH_KEY_USER_UPN%} as upn search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key UPN: %$SEARCH_KEY_USER_UPN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUUS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_UPN%" -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_UPN%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_UPN%"-limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_UPN%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSUUS
:jumpSUUS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_UPN%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_USER_UPN%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_UPN%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_USER_UPN%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSUUS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUUO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	

	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUUL
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUUO
:jumpSUUL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUUO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUU
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"

:skipSUU
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo SUU	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SUFL
	:: Search User first and last name
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	SET $LAST_SEARCH_ATTRIBUTE=FirstLast
	call :SM
	SET $SEARCH_KEY_USER_FIRST=
	@powershell Write-Host "givenName - First" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to wildcard *" -ForegroundColor Magenta
	SET /P $SEARCH_KEY_USER_FIRST=User FirstName search key:
	IF NOT DEFINED $SEARCH_KEY_USER_FIRST SET $SEARCH_KEY_USER_FIRST=*
	call :SM
	SET $SEARCH_KEY_USER_LAST=
	@powershell Write-Host "sn - surename - Last" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to wildcard *" -ForegroundColor Magenta
	SET /P $SEARCH_KEY_USER_LAST=User LastName search key:
	IF NOT DEFINED $SEARCH_KEY_USER_LAST SET $SEARCH_KEY_USER_LAST=*
	call :SM	
	echo Selected {%$SEARCH_KEY_USER_FIRST%} as FirstName search key.
	echo Selected {%$SEARCH_KEY_USER_LAST%} as LastName search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key FirstName: %$SEARCH_KEY_USER_FIRST% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key LastName: %$SEARCH_KEY_USER_LAST% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUFLS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name distinguishedName %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr distinguishedName %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
	)	
	if %$SORTED% NEQ 1 GoTo skipSUFLS
:jumpSUFLS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name distinguishedName %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr distinguishedName %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
	)	

:skipSUFLS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUFLO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	
	
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"	
	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	

	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUFLL
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUFLO
:jumpSUFLL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUFLO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUFL
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"

:skipSUFL
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo SUFL
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SUDN
	:: Search User display name
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	SET $LAST_SEARCH_ATTRIBUTE=DisplayName
	call :SM
	SET $SEARCH_KEY_USER_DIPLAYNAME=	
	@powershell Write-Host "DisplayName" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Red
	SET /P $SEARCH_KEY_USER_DIPLAYNAME=User DisplayName search key:
	IF NOT DEFINED $SEARCH_KEY_USER_DIPLAYNAME GoTo SUDN
	call :SM
	echo Selected {%$SEARCH_KEY_USER_DIPLAYNAME%} as DisplayName search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key %$LAST_SEARCH_ATTRIBUTE%: %$SEARCH_KEY_USER_DIPLAYNAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUDNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr displayName name distinguishedName %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr displayName name distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr name %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
	)	
	if %$SORTED% NEQ 1 GoTo skipSUDNS
:jumpSUDNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr displayName name distinguishedName %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr displayName name distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr name %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
	)	

:skipSUDNS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUDNO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo DisplayName				Name			DistinguishedName^(DN^): >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUDNL
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUDNO
:jumpSUDNL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUDNO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUDN
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"

:skipSUDN
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo SUDN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SUCA
	:: Search User Custom Attributes
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2> nul 1> nul
	
	SET $SEARCH_TYPE=user
	SET $LAST_SEARCH_ATTRIBUTE=Custom Attribe
	SET "$ATTRIBUTES_USER=cn department description directReports displayName email extensionAttribute# givenName l mail mailNickname manager memberOf name physicalDeliveryOfficeName postalCode userPrincipalName sn st title telephoneNumber"
	call :SM
	@powershell Write-Host "Custom-Attribute:" -ForegroundColor Gray
	@powershell Write-Host "cn department description directReports displayName email extensionAttribute# givenName l mail mailNickname manager memberOf name physicalDeliveryOfficeName postalCode userPrincipalName sn st title telephoneNumber" -ForegroundColor Blue
	echo ----------------------------------------
	echo.
	:: Choose custom user attribute
	@powershell Write-Host "Choose one user attribute to search against:" -ForegroundColor Gray
	@powershell Write-Host "name will automatically be used" -ForegroundColor Gray
	SET $FILTER=^(objectClass=%$SEARCH_TYPE%^)
	SET /P $ATTRIBUTES_USER=User attribute:
	echo %$ATTRIBUTES_USER% | FIND /I "%$ATTRIBUTES_USER%"
	SET $USER_ATT_VALID=%ERRORLEVEL%
	IF %$USER_ATT_VALID% NEQ 0 (ECHO NOT VALID!) & (timeout /t 10) & (GoTo SUCA)
	:: Choose custom attribute search key
	call :SM
	echo Custom User Attribute: %$ATTRIBUTES_USER%
	@powershell Write-Host "Choose a search key:" -ForegroundColor Gray
	@powershell Write-Host "Leave blank for wildcard *" -ForegroundColor Magenta
	SET $USER_ATTR_CUSTOM_SEARCH_KEY=*
	SET /P $USER_ATTR_CUSTOM_SEARCH_KEY=User %$ATTRIBUTES_USER% search key:
	:: Choose name attribute search key*
	call :SM
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Choose a search key:" -ForegroundColor Gray
	@powershell Write-Host "Leave blank for wildcard *" -ForegroundColor Magenta
	SET $USER_ATTR_NAME_SEARCH_KEY=*
	SET /P $USER_ATTR_NAME_SEARCH_KEY=User name search key:
	:: Console Display		
	call :SM	
	@powershell Write-Host "Search parameters:" -ForegroundColor Gray
	echo Attribute		Operator  Search Key
	echo Name			[=]	%$USER_ATTR_NAME_SEARCH_KEY%
	echo %$ATTRIBUTES_USER%	[=]	%$USER_ATTR_CUSTOM_SEARCH_KEY%
	:: Make filter
	SET $FILTER=%$FILTER%^(name=%$USER_ATTR_NAME_SEARCH_KEY%^)^(%$ATTRIBUTES_USER%=%$USER_ATTR_CUSTOM_SEARCH_KEY%^)
	echo.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo  name Attribute: [=]	%$USER_ATTR_NAME_SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo  %$ATTRIBUTES_USER% Attribute: [=]	%$USER_ATTR_CUSTOM_SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUCAS
	:: Search
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr %$ATTRIBUTES_USER% name displayName -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (	
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr %$ATTRIBUTES_USER% name displayName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD%  > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)

	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (	
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD%  > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSUCAS
		
:jumpSUCAS

	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr %$ATTRIBUTES_USER% name displayName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (	
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr %$ATTRIBUTES_USER% name displayName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)

	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (	
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSUCAS

	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUCAL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute Summary: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$ATTRIBUTES_USER%	Name		displayName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "Name" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	:: Check User session
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUCAO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUCAO
:jumpSUCAL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUCAO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUCA
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"

:skipSUCA
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUCA
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SUG
	:: Search User Global
	call :SM
	echo User global search using:
	echo.
	echo [1] Inactive
	echo [2] StalePassword
	echo [3] Disabled
	echo [4] Abort
	echo.
	Choice /c 1234
	echo.
	If ERRORLevel 4 GoTo Search
	If ERRORLevel 3 GoTo SUGD
	If ERRORLevel 2 GoTo SUGS
	If ERRORLevel 1 GoTo SUGI
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SUGI
	:: Search User Global Inactive
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	SET $LAST_SEARCH_ATTRIBUTE=Inactive
	call :SM
	@powershell Write-Host "Inactive" -ForegroundColor Blue	
	@powershell Write-Host "Choose n number of weeks" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to 0" -ForegroundColor Magenta
	SET /P $SEARCH_INACTIVE_KEY=%$LAST_SEARCH_ATTRIBUTE% n weeks:
	IF NOT DEFINED $SEARCH_INACTIVE_KEY=0
	call :SM
	@powershell Write-Host "Attribute: Name" -ForegroundColor Blue	
	@powershell Write-Host "Can use wildcard *" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to *" -ForegroundColor Magenta
	SET $SEARCH_NAME_KEY=*
	SET /P $SEARCH_NAME_KEY=name search key:
	call :SM
	echo Selected {%$SEARCH_INACTIVE_KEY%} as inactive search key.
	echo Selected {%$SEARCH_NAME_KEY%} as name search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute %$LAST_SEARCH_ATTRIBUTE% key: %$SEARCH_INACTIVE_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute Name key: %$SEARCH_NAME_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUGIS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSUGIS
:jumpSUGIS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSUGIS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUGIO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	

	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUGIL
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUGIO
:jumpSUGIL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUGIO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUGI
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"


:skipSUGI
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo SUGI
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:SUGS
	:: Search User Global StalePassword
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	SET $LAST_SEARCH_ATTRIBUTE=StalePassword
	call :SM
	@powershell Write-Host "Attribute: StalePassword" -ForegroundColor Blue	
	@powershell Write-Host "Choose n number of days" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to 0" -ForegroundColor Magenta
	SET /P $SEARCH_STALEPWD_KEY=%$LAST_SEARCH_ATTRIBUTE% n days:
	IF NOT DEFINED $SEARCH_STALEPWD_KEY=0
	call :SM
	@powershell Write-Host "Attribute: Name" -ForegroundColor Blue	
	@powershell Write-Host "Can use wildcard *" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to *" -ForegroundColor Magenta
	SET $SEARCH_NAME_KEY=*
	SET /P $SEARCH_NAME_KEY=name search key:
	call :SM
	echo Selected {%$SEARCH_STALEPWD_KEY%} as stalepassword search key.
	echo Selected {%$SEARCH_NAME_KEY%} as name search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute %$LAST_SEARCH_ATTRIBUTE% key: %$SEARCH_STALEPWD_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute Name key: %$SEARCH_NAME_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUGSS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSUGSS
:jumpSUGSS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSUGSS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUGSO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	

	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUGSL
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUGSO
:jumpSUGSL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUGSO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUGS
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"


:skipSUGS
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo SUGS
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	

	
:SUGD
	:: Search User Global Disabled
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	SET $LAST_SEARCH_ATTRIBUTE=Disabled
	call :SM
	@powershell Write-Host "Attribute: Name" -ForegroundColor Blue	
	@powershell Write-Host "Can use wildcard *" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to *" -ForegroundColor Magenta
	SET $SEARCH_NAME_KEY=*
	SET /P $SEARCH_NAME_KEY=name search key:
	call :SM
	echo Selected {%$SEARCH_NAME_KEY%} as name search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute Name key: %$SEARCH_NAME_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSUGDS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSUGDS
:jumpSUGDS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSUGDS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSUGDO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	

	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSUGDL
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSUGDO
:jumpSUGDL
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (		
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)	

:jumpSUGDO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSUGD
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"


:skipSUGD
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo SUGD	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:sGroup
	:: Search Group
	SET $LAST_SEARCH_TYPE=Group
	call :SM
	SET $SEARCH_KEY=
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
	Echo Group search using:
	Echo.
	Echo [1] Name
	Echo [2] Description
	Echo [3] DisplayName
	echo [4] Abort
	Echo.
	Choice /c 1234
	Echo.
	If ERRORLevel 4 GoTo Search
	If ERRORLevel 3 GoTo sGDN
	If ERRORLevel 2 GoTo sGD
	If ERRORLevel 1 GoTo sGN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:sGN
	:: Group Search using Name attribute
	SET $LAST_SEARCH_ATTRIBUTE=name
	CALL :SM
	echo ^(Can use "*" wildcard for search.^)
	echo If left blank, will abort.
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY_LAST=NA) ELSE (SET $SEARCH_KEY_LAST=%$SEARCH_KEY%)
	SET $SEARCH_KEY=
	SET /P $SEARCH_KEY=Choose a search key ^(word^):
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY=%$SEARCH_KEY_LAST%)
	IF /I "%$SEARCH_KEY%"=="NA" GoTo skipSGO
	call :SM
	echo Selected {%$SEARCH_KEY%} as search key.
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY%
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Term: %$SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	
	if %$SORTED% EQU 1 GoTo jumpSGC 
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSGS
:jumpSGC
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSGS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSGL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSGMO 
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -dn %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -members %$AD_SERVER_SEARCH% 2> nul | DSGET USER -upn -fn -mi -ln -display -email 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
GoTo jumpSGL
	:jumpSGMO 
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -dn %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -members %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2> nul | DSGET USER -upn -fn -mi -ln -display -email 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSGL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSGO
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSGO
	echo Search Group Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sGroup	

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:sGD
	:: Group Search using Description attribute
	SET $LAST_SEARCH_ATTRIBUTE=description
	CALL :SM
	echo ^(Can use "*" wildcard for search.^)
	echo If left blank, will abort.
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY_LAST=NA) ELSE (SET $SEARCH_KEY_LAST=%$SEARCH_KEY%)
	SET $SEARCH_KEY=
	SET /P $SEARCH_KEY=Choose a search key ^(word^):
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY=%$SEARCH_KEY_LAST%)
	IF /I "%$SEARCH_KEY%"=="NA" GoTo skipSGDA
	call :SM
	echo Selected {%$SEARCH_KEY%} as search key.
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY%
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Term: %$SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	
	if %$SORTED% EQU 1 GoTo jumpSGDS 
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSGDS
:jumpSGDS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o rdn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY GROUP %$AD_BASE% -scope %$AD_SCOPE% -o dn -desc "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSGDS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSGDL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSGDO 
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -dn %$AD_SERVER_SEARCH% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -members %$AD_SERVER_SEARCH% 2> nul | DSGET USER -upn -fn -mi -ln -display -email 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
GoTo jumpSGDL
	:jumpSGDO 
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr displayName %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -dn %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP %%N -members %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2> nul | DSGET USER -upn -fn -mi -ln -display -email 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSGDL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSGDA
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSGDA
	echo Search Group Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sGroup	

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:sGDN
	:: Group Search using DisplayName attribute
	SET $LAST_SEARCH_ATTRIBUTE=description
	CALL :SM
	echo ^(Can use "*" wildcard for search.^)
	echo If left blank, will abort.
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY_LAST=NA) ELSE (SET $SEARCH_KEY_LAST=%$SEARCH_KEY%)
	SET $SEARCH_KEY=
	SET /P $SEARCH_KEY=Choose a search key ^(word^):
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY=%$SEARCH_KEY_LAST%)
	IF /I "%$SEARCH_KEY%"=="NA" GoTo skipSGDNA
	call :SM
	echo Selected {%$SEARCH_KEY%} as search key.
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY%
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Term: %$SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	
	if %$SORTED% EQU 1 GoTo jumpSGDNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSGDNS
:jumpSGDNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSGDNS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"	
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSGDNL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSGDNO
	:: Session user is a domain user
	
	FOR /F "USEBACKQ skip=1 tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP "%%N" -dn %$AD_SERVER_SEARCH% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP "%%N" -members %$AD_SERVER_SEARCH% 2> nul | DSGET USER -upn -fn -mi -ln -display -email 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
GoTo jumpSGDNL
:jumpSGDNO
	:: Session user is a local user
	FOR /F "USEBACKQ skip=1 tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr displayName %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP "%%N" -dn %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSGET GROUP "%%N" -members %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2> nul | DSGET USER -upn -fn -mi -ln -display -email 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSGDNL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSGDNA
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSGDNA
	echo Search Group Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sGroup	

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:sComputer
	:: Search Computer
	SET $LAST_SEARCH_TYPE=Computer
	SET $LAST_SEARCH_ATTRIBUTE=name
	call :SM
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul

	echo Computer search using:
	echo.
	Echo [1] Name
	Echo [2] Advanced
	Echo [3] Abort
	Echo.
	Choice /c 123
	Echo.
	If ERRORLevel 3 GoTo Search
	If ERRORLevel 2 GoTo sCA
	If ERRORLevel 1 GoTo sCN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	
	
:sCN
	:: Search Computer Name
	call :SM
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort." -ForegroundColor Red
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY_LAST=NA) ELSE (SET $SEARCH_KEY_LAST=%$SEARCH_KEY%)
	SET $SEARCH_KEY_PC_NAME=
	SET /P $SEARCH_KEY_PC_NAME=Computer name search key:
	IF NOT DEFINED $SEARCH_KEY_PC_NAME (SET $SEARCH_KEY_PC_NAME=NA)
	IF /I "%$SEARCH_KEY_PC_NAME%"=="NA" GoTo sComputer
	call :SM
	echo Selected {%$SEARCH_KEY_PC_NAME%} as name search key.
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_PC_NAME%
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key Name: %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"		
	
	if %$SORTED% EQU 1 GoTo jumpSCNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCNS
:jumpSCNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCNS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCNL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCMO 
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCNL
	:jumpSCMO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCNL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSCN
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSCN
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sComputer	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	

:sCA
	:: Search Computer Advanced
	SET $LAST_SEARCH_TYPE=Computer
	SET $LAST_SEARCH_ATTRIBUTE=name
	mode con:lines=42
	Color 0A
	call :SM
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2> nul 1> nul

	::Selection
	echo Computer Advanced search using:
	echo [1] Disabled
	echo [2] Inactive
	echo [3] StalePWD
	echo [4] Operating System ^(attributes^)
	echo [5] Time Series ^(attributes^)
	echo [6] LogonCount
	echo [7] Multiple Attribute search
	echo [8] Abort
	Echo.
	Choice /c 12345678
	Echo.
	If ERRORLevel 8 GoTo Search
	If ERRORLevel 7 GoTo sCMA
	If ERRORLevel 6 GoTo sCLC
	If ERRORLevel 5 GoTo sCTS
	If ERRORLevel 4 GoTo sCOS
	If ERRORLevel 3 GoTo sCS
	If ERRORLevel 2 GoTo sCI
	If ERRORLevel 1 GoTo sCD
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	
:sCD
	:: Search computer disabled
	SET $LAST_SEARCH_ATTRIBUTE=Disabled
	call :SM
	echo ^(Can use "*" wildcard for search.^)
	echo If left blank, will default to "*".
	SET /P $SEARCH_KEY_PC_NAME=Choose name search key:
	IF NOT DEFINED $SEARCH_KEY_PC_NAME SET $SEARCH_KEY_PC_NAME=*
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_PC_NAME%
	call :SM
	echo Selected {%$SEARCH_KEY_PC_NAME%} as name search key.	
	echo Search for all disabled computers with name key: {%$SEARCH_KEY_PC_NAME%}
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow	
	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key Name: %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
		if %$SORTED% EQU 1 GoTo jumpSCDS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCDS
:jumpSCDS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCDS

	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCDL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCDO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCDL
	:jumpSCDO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCDL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipsCD
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSCD
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sComputer	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

	
:sCI
	:: Computers inactive search
	SET $LAST_SEARCH_ATTRIBUTE=Inactive
	call :SM
	echo ^(Can use "*" wildcard for search.^)
	echo If left blank, will default to "*".
	SET /P $SEARCH_KEY_PC_NAME=Choose name search key:
	IF NOT DEFINED $SEARCH_KEY_PC_NAME SET $SEARCH_KEY_PC_NAME=*
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_PC_NAME%
	Echo Inactive for ^<n^> Weeks:
	echo ^(If left blank, will abort!^)
	SET /P $SEARCH_INACTIVE=Inactive number of weeks:
	IF NOT DEFINED $SEARCH_INACTIVE GoTo skipSCI
	call :SM
	echo Selected {%$SEARCH_KEY_PC_NAME%} as name search key.	
	echo Search for all inactive computers with name key: {%$SEARCH_KEY_PC_NAME%}
	echo ...for the last {%$SEARCH_INACTIVE%} weeks...
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow	
	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute Parameter: %$SEARCH_INACTIVE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key Name: %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	if %$SORTED% EQU 1 GoTo jumpSCIS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCIS
:jumpSCIS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_INACTIVE% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCIS

	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCIL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCIO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCIL
:jumpSCIO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCIL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipsCI
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSCI
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sComputer	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:sCS
	:: Computers with stale passwords search
	SET $LAST_SEARCH_ATTRIBUTE=StalePassword
	call :SM
	echo ^(Can use "*" wildcard for search.^)
	echo If left blank, will default to "*".
	SET /P $SEARCH_KEY_PC_NAME=Choose name search key:	
	IF NOT DEFINED $SEARCH_KEY_PC_NAME SET $SEARCH_KEY_PC_NAME=*
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_PC_NAME%	
	Echo Stale password for ^<n^> days:
	echo ^(If left blank, will abort!^)	
	SET /P $SEARCH_STALEPWD=Stale password number of days:
	IF NOT DEFINED $SEARCH_STALEPWD GoTo skipSCS
	call :SM	
	echo Selected {%$SEARCH_KEY_PC_NAME%} as name search key.	
	echo Search for all inactive computers with name key: {%$SEARCH_KEY_PC_NAME%}
	echo ...for the last {%$SEARCH_STALEPWD%} days...
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow	
	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute Parameter: %$SEARCH_STALEPWD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key Name: %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	if %$SORTED% EQU 1 GoTo jumpSCSS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCSS
:jumpSCSS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_stalepwd% -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCSS

	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCSL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCSO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCSL
:jumpSCSO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCSL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSCS
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSCS
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sComputer	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	
	
:SCOS
	:: Search computer Operating system andOr version
	SET $LAST_SEARCH_ATTRIBUTE=OperatingSystem
	call :SM
	@powershell Write-Host "Define the following search parameters:" -ForegroundColor DarkYellow
	@powershell Write-Host "operatingSystem" -ForegroundColor Blue
	@powershell Write-Host "operatingSystemVersion" -ForegroundColor Blue
	@powershell Write-Host "operatingSystemServicePack" -ForegroundColor Blue
	echo ^(Can use "*" wildcard for search.^)
	echo If left blank, will default to "*".
	SET /P $SEARCH_KEY_PC_NAME=Choose computer name search key:
	IF NOT DEFINED $SEARCH_KEY_PC_NAME SET $SEARCH_KEY_PC_NAME=*
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_PC_NAME%		
	@powershell Write-Host "Operating System:" -ForegroundColor Blue
	echo ^(If left blank, will abort!^)	
	SET /P $SEARCH_OS=Operating System:
	IF NOT DEFINED $SEARCH_OS GoTo skipSCOS
	echo If left blank, will default to "*".
	@powershell Write-Host "Operating System Version:" -ForegroundColor Blue
	SET /P $SEARCH_OSV=Operating System Version:
	IF NOT DEFINED $SEARCH_OSV SET $SEARCH_OSV=*
	echo If left blank, will default to "*".
	@powershell Write-Host "Operating System Service Pack:" -ForegroundColor Blue 
	SET /P $SEARCH_OSSP=Operating System Service Pack:
	IF NOT DEFINED $SEARCH_OSSP SET $SEARCH_OSSP=*
	call :SM		
	echo Selected {%$SEARCH_KEY_PC_NAME%} as computer name search key.
	echo Selected {%$SEARCH_OS%} as Operating System search key.
	echo Selected {%$SEARCH_OSV%} as Operating System Version search key.
	echo Selected {%$SEARCH_OSSP%} as Operating System Service Pack search key.	
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Computer Name: %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Operating System: %$SEARCH_OS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Operating System Version: %$SEARCH_OSV% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Operating System Service pack: %$SEARCH_OSSP% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	if %$SORTED% EQU 1 GoTo jumpSCOSS
	:: Search
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCOSS
:jumpSCOSS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCOSS

	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCOSL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Names and DN returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	:: Check User session
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCOSO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCOSL
:jumpSCOSO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCOSL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSCOS
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
:skipSCOS
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sComputer	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	

:SCTS
	:: Search Computer Time Series attributes
	SET $LAST_SEARCH_ATTRIBUTE=TimeSeries
	SET $SEARCH_KEY_PC_NAME=
	SET $SEARCH_WHENCREATED=
	SET $SEARCH_WHENCHANGED=
	SET $SEARCH_LASTLOGONTIMESTAMP=
	call :SM
	@powershell Write-Host "Timeseries search parameters:" -ForegroundColor Gray
	@powershell Write-Host "whenCreated" -ForegroundColor Blue
	@powershell Write-Host "whenChanged" -ForegroundColor Blue
	@powershell Write-Host "lastLogonTimestamp" -ForegroundColor Blue
	@powershell Write-Host "Use "*" wildcard for search." -ForegroundColor Magenta
	@powershell Write-Host "If left blank, will default to * wildcard" -ForegroundColor Red
	echo ----------------------------------------
	:: Computer Name search key
	@powershell Write-Host "Search key Computer Name:" -ForegroundColor Cyan
	@powershell Write-Host "If left blank, will abort!" -ForegroundColor Red
	SET /P $SEARCH_KEY_PC_NAME=Choose computer name search key:
	IF NOT DEFINED $SEARCH_KEY_PC_NAME GoTo skipSCTS
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_PC_NAME%
	call :SM
	:: whenCreated
	@powershell Write-Host "whenCreated operator:" -ForegroundColor Blue
	echo [1] Equal [=]
	echo [2] Approximately equal to [^~=]
	echo [3] Less [^<=]
	echo [4] Greater [^>=]
	Choice /c 1234
	If ERRORLevel 4 (SET "$SEARCH_WHENCREATED_OPERATOR=>=") & (SET "$SEARCH_WHENCREATED_OPERATOR_DISPLAY=^>^=")
	If ERRORLevel 3 (SET "$SEARCH_WHENCREATED_OPERATOR=<=") & (SET "$SEARCH_WHENCREATED_OPERATOR_DISPLAY=^<^=")
	If ERRORLevel 2 (SET "$SEARCH_WHENCREATED_OPERATOR=~=") & (SET "$SEARCH_WHENCREATED_OPERATOR_DISPLAY=^~^=")
	If ERRORLevel 1 (SET "$SEARCH_WHENCREATED_OPERATOR==") & (SET "$SEARCH_WHENCREATED_OPERATOR_DISPLAY=^=")
	@powershell Write-Host "whenCreated:" -ForegroundColor Blue
	@powershell Write-Host "[YYYY MM DD HH mm ss.s Z] i.g. 20200101120000.0Z" -ForegroundColor Cyan
	SET /P $SEARCH_WHENCREATED=whenCreated search key:
	IF NOT DEFINED $SEARCH_WHENCREATED SET $SEARCH_WHENCREATED=*	
	call :SM	
	:: whenChanged
	@powershell Write-Host "whenChanged operator:" -ForegroundColor Blue
	echo [1] Equal [=]
	echo [2] Approximately equal to [^~=]
	echo [3] Less [^<=]
	echo [4] Greater [^>=]
	Choice /c 1234
	If ERRORLevel 4 (SET "$SEARCH_WHENCHANGED_OPERATOR=>=") & (SET "$SEARCH_WHENCHANGED_OPERATOR_DISPLAY=^>^=")
	If ERRORLevel 3 (SET "$SEARCH_WHENCHANGED_OPERATOR=<=") & (SET "$SEARCH_WHENCHANGED_OPERATOR_DISPLAY=^<^=")
	If ERRORLevel 2 (SET "$SEARCH_WHENCHANGED_OPERATOR=~=") & (SET "$SEARCH_WHENCHANGED_OPERATOR_DISPLAY=^~^=")
	If ERRORLevel 1 (SET "$SEARCH_WHENCHANGED_OPERATOR==") & (SET "$SEARCH_WHENCHANGED_OPERATOR_DISPLAY=^=")
	@powershell Write-Host "whenChanged:" -ForegroundColor Blue
	@powershell Write-Host "[YYYY MM DD HH mm ss.s Z] i.g. 20200101120000.0Z" -ForegroundColor Cyan
	SET /P $SEARCH_whenChanged=whenChanged search key:
	IF NOT DEFINED $SEARCH_WHENCHANGED SET $SEARCH_WHENCHANGED=*
	call :SM
	:: lastLogonTimestamp
	@powershell Write-Host "lastLogonTimestamp operator:" -ForegroundColor Blue
	echo [1] Equal [=]
	echo [2] Approximately equal to [^~=]
	echo [3] Less [^<=]
	echo [4] Greater [^>=]
	Choice /c 1234
	If ERRORLevel 4 (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR=>=") & (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY=^>^=")
	If ERRORLevel 3 (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR=<=") & (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY=^<^=")
	If ERRORLevel 2 (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR=~=") & (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY=^~^=")
	If ERRORLevel 1 (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR==") & (SET "$SEARCH_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY=^=")
	@powershell Write-Host "lastLogonTimestamp search key:" -ForegroundColor Blue
	@powershell Write-Host "[NT Time] e.g. 132530551699076595" -ForegroundColor Cyan
	SET /P $SEARCH_LASTLOGONTIMESTAMP=lastLogonTimestamp:
	IF NOT DEFINED $SEARCH_LASTLOGONTIMESTAMP SET $SEARCH_LASTLOGONTIMESTAMP=*
	IF "%$SEARCH_LASTLOGONTIMESTAMP%"=="*" SET $NT_TIME_CONVERTED=*
	IF "%$SEARCH_LASTLOGONTIMESTAMP%"=="*" GoTo skipLLTSC
	FOR /F "tokens=2 delims=-" %%P IN ('w32tm.exe /ntte %$SEARCH_LASTLOGONTIMESTAMP%') DO echo %%P> "%$LogPath%\var\var_$NT_TIME_CONVERTED.txt"
	SET /P $NT_TIME_CONVERTED= < "%$LogPath%\var\var_$NT_TIME_CONVERTED.txt"
	:skipLLTSC
	call :SM
	:: Console Display	
	echo Search parameters:
	echo Attribute		Operator  Search Key
	echo  name			[^=]	%$SEARCH_KEY_PC_NAME%
	echo  whenCreated		[%$SEARCH_WHENCREATED_OPERATOR_DISPLAY%]	%$SEARCH_WHENCREATED%
	echo  whenChanged		[%$SEARCH_WHENCHANGED_OPERATOR_DISPLAY%]	%$SEARCH_WHENCHANGED%
	echo  lastLogonTimestamp	[%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%]	%$NT_TIME_CONVERTED%
	echo.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Computer Name: [^=] %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search whenCreated: [%$SEARCH_WHENCREATED_OPERATOR_DISPLAY%] %$SEARCH_WHENCREATED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search whenChanged: [%$SEARCH_WHENCHANGED_OPERATOR_DISPLAY%] %$SEARCH_WHENCHANGED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search	lastLogonTimestamp: [%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%] %$SEARCH_LASTLOGONTIMESTAMP% %$NT_TIME_CONVERTED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	if %$SORTED% EQU 1 GoTo jumpSCTSS
	:: Search
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE%-filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCTSS
:jumpSCTSS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$SEARCH_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$SEARCH_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCTSS

	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCTSL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Name	distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	:: Check User session
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCTSO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCTSL
:jumpSCTSO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCTSL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSCTS
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"

:skipSCTS
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo SCTS	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	

:sCLC
	:: Search Computer LogonCount
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2> nul 1> nul
	SET $LAST_SEARCH_ATTRIBUTE=LogonCount
	SET $SEARCH_KEY_PC_NAME=
	SET $SEARCH_LOGONCOUNT=
	call :SM
	@powershell Write-Host "LogonCount" -ForegroundColor Blue
	:: Computer Name search key
	@powershell Write-Host "Search key Computer Name:" -ForegroundColor Cyan
	@powershell Write-Host "If left blank, will abort!" -ForegroundColor Red
	SET /P $SEARCH_KEY_PC_NAME=Choose computer name search key:
	IF NOT DEFINED $SEARCH_KEY_PC_NAME GoTo skipSCLC
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY_PC_NAME%
	call :SM
	:: LogonCount
	@powershell Write-Host "logonCount operator:" -ForegroundColor Blue
	echo [1] Equal [=]
	echo [2] Approximately equal to [^~=]
	echo [3] Less [^<=]
	echo [4] Greater [^>=]
	Choice /c 1234
	If ERRORLevel 4 (SET "$SEARCH_LOGONCOUNT_OPERATOR=>=") & (SET "$SEARCH_LOGONCOUNT_OPERATOR_DISPLAY=^>^=")
	If ERRORLevel 3 (SET "$SEARCH_LOGONCOUNT_OPERATOR=<=") & (SET "$SEARCH_LOGONCOUNT_OPERATOR_DISPLAY=^<^=")
	If ERRORLevel 2 (SET "$SEARCH_LOGONCOUNT_OPERATOR=~=") & (SET "$SEARCH_LOGONCOUNT_OPERATOR_DISPLAY=^~^=")
	If ERRORLevel 1 (SET "$SEARCH_LOGONCOUNT_OPERATOR==") & (SET "$SEARCH_LOGONCOUNT_OPERATOR_DISPLAY=^=")
	@powershell Write-Host "LOGONCOUNT:" -ForegroundColor Blue
	@powershell Write-Host "Use * wildcard for search." -ForegroundColor Magenta
	@powershell Write-Host "If left blank, will abort!" -ForegroundColor Red
	SET /P $SEARCH_LOGONCOUNT=LogonCount search key:
	IF NOT DEFINED $SEARCH_LOGONCOUNT GoTo skipSCLC
	call :SM		
	:: Console Display	
	echo Search parameters:
	echo Attribute		Operator  Search Key
	echo  name			[^=]	%$SEARCH_KEY_PC_NAME%
	echo  LogonCount		[%$SEARCH_LOGONCOUNT_OPERATOR_DISPLAY%]	%$SEARCH_LOGONCOUNT%	
	echo.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow

	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Computer Name: [^=] %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: logonCount: [%$SEARCH_LOGONCOUNT_OPERATOR_DISPLAY%] %$SEARCH_LOGONCOUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	if %$SORTED% EQU 1 GoTo jumpSCLCS
	:: Search
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE%-filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCLCS
:jumpSCLCS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$SEARCH_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCLCS

	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCLCL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Name	distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	:: Check User session
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCLCO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr logonCount %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCLCL
:jumpSCLCO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr logonCount %$AD_SERVER_SEARCH%  -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCLCL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSCLC
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	
:skipSCLC
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo sCLC	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:SCMA

	:: Search Computer Multiple Attributes
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2> nul 1> nul
	SET $LAST_SEARCH_ATTRIBUTE=Multiple
	SET $SEARCH_TYPE=COMPUTER
	SET "$ATTRIBUTES_COMPUTER=name cn description displayName whenCreated whenChanged logonCount lastLogonTimestamp objectSid dNSHostName operatingSystem operatingSystemVersion operatingSystemServicePack managedBy"
	call :SM
	@powershell Write-Host "Multiple-meta:" -ForegroundColor Gray
	@powershell Write-Host "name cn description displayName whenCreated whenChanged logonCount lastLogonTimestamp objectSid dNSHostName operatingSystem operatingSystemVersion operatingSystemServicePack managedBy" -ForegroundColor Blue
	echo ----------------------------------------
	echo.
	timeout /T 10
	SET $COUNTER=1
	SET $COUNTER_MAX=15

:subSCMA
	FOR /F "tokens=%$COUNTER%" %%P IN ("%$ATTRIBUTES_COMPUTER%") DO (
		(call :SM) & (
		IF %$COUNTER% EQU %$COUNTER_MAX% GoTo eSubSCMA) & (
		echo Attribute: %%P) & (
		call :subOperator %%P) & (
		@powershell Write-Host "Leave blank for wildcard *" -ForegroundColor Magenta) & (
		SET $ATTR_%%P=*) & (
		SET /P $ATTR_%%P=Computer %%P search key: ) & (
		SET /a $COUNTER+=1) & (
		GoTo subSCMA)
	)
:eSubSCMA	
	
	call :SM		
	:: Console Display	
	@powershell Write-Host "Search parameters:" -ForegroundColor Gray
	echo Attribute			Operator  Search Key
	echo  name				[%$ATTR_NAME_OPERATOR_DISPLAY%]	%$ATTR_NAME%
	echo  cn				[%$ATTR_CN_OPERATOR_DISPLAY%]	%$ATTR_CN%
	echo  description			[%$ATTR_DESCRIPTION_OPERATOR_DISPLAY%]	%$ATTR_DESCRIPTION%
	echo  displayname			[%$ATTR_DISPLAYNAME_OPERATOR_DISPLAY%]	%$ATTR_DISPLAYNAME%
	echo  WHENCREATED			[%$ATTR_WHENCREATED_OPERATOR_DISPLAY%]	%$ATTR_WHENCREATED%
	echo  WHENCHANGED			[%$ATTR_WHENCHANGED_OPERATOR_DISPLAY%]	%$ATTR_WHENCHANGED%
	echo  LOGONCOUNT			[%$ATTR_LOGONCOUNT_OPERATOR_DISPLAY%]	%$ATTR_LOGONCOUNT%
	echo  LASTLOGONTIMESTAMP		[%$ATTR_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%]	%$ATTR_LASTLOGONTIMESTAMP%
	echo  OBJECTSID			[%$ATTR_OBJECTSID_OPERATOR_DISPLAY%]	%$ATTR_OBJECTSID%
	echo  DNSHOSTNAME			[%$ATTR_DNSHOSTNAME_OPERATOR_DISPLAY%]	%$ATTR_DNSHOSTNAME%
	echo  OPERATINGSYSTEM		[%$ATTR_OPERATINGSYSTEM_OPERATOR_DISPLAY%]	%$ATTR_OPERATINGSYSTEM%
	echo  OPERATINGSYSTEMVERSION		[%$ATTR_OPERATINGSYSTEMVERSION_OPERATOR_DISPLAY%]	%$ATTR_OPERATINGSYSTEMVERSION%
	echo  OPERATINGSYSTEMSERVICEPACK	[%$ATTR_OPERATINGSYSTEMSERVICEPACK_OPERATOR_DISPLAY%]	%$ATTR_OPERATINGSYSTEMSERVICEPACK%
	echo  MANAGEDBY			[%$ATTR_MANAGEDBY_OPERATOR_DISPLAY%]	%$ATTR_MANAGEDBY%
	echo.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET	
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Name Attribute: [%$ATTR_NAME_OPERATOR_DISPLAY%] %$ATTR_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo cn Attribute: [%$ATTR_CN_OPERATOR_DISPLAY%] %$ATTR_CN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo description Attribute: [%$ATTR_DESCRIPTION_OPERATOR_DISPLAY%] %$ATTR_DESCRIPTION% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo displayName Attribute: [%$ATTR_DISPLAYNAME_OPERATOR_DISPLAY%] %$ATTR_DISPLAYNAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo whenCreated Attribute: [%$ATTR_WHENCREATED_OPERATOR_DISPLAY%] %$ATTR_WHENCREATED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo whenChanged Attribute: [%$ATTR_WHENCHANGED_OPERATOR_DISPLAY%] %$ATTR_WHENCHANGED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo logonCount Attribute: [%$ATTR_LOGONCOUNT_OPERATOR_DISPLAY%] %$ATTR_LOGONCOUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo lastLogonTimestamp Attribute: [%$ATTR_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%] %$ATTR_LASTLOGONTIMESTAMP% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo objectSid Attribute: [%$ATTR_OBJECTSID_OPERATOR_DISPLAY%] %$ATTR_OBJECTSID% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo dNSHostName Attribute: [%$ATTR_DNSHOSTNAME_OPERATOR_DISPLAY%] %$ATTR_DNSHOSTNAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo operatingSystem Attribute: [%$ATTR_OPERATINGSYSTEM_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEM% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo operatingSystemVersion Attribute: [%$ATTR_OPERATINGSYSTEMVERSION_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEMVERSION% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo operatingSystemServicePack Attribute: [%$ATTR_OPERATINGSYSTEMSERVICEPACK_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEMSERVICEPACK% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo managedBy Attribute: [%$ATTR_MANAGEDBY_OPERATOR_DISPLAY%] %$ATTR_MANAGEDBY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	if %$SORTED% EQU 1 GoTo jumpSCMAS
	:: Search
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH%  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE%-filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSCMAS
:jumpSCMAS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr name -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)(displayname%$ATTR_displayname_OPERATOR%%$ATTR_displayname%)(whenCreated%$ATTR_whenCreated_OPERATOR%%$ATTR_whenCreated%)(whenChanged%$ATTR_whenChanged_OPERATOR%%$ATTR_whenChanged%)(logonCount%$ATTR_logonCount_OPERATOR%%$ATTR_logonCount%)(lastLogonTimestamp%$ATTR_lastLogonTimestamp_OPERATOR%%$ATTR_lastLogonTimestamp%)(objectSid%$ATTR_objectSid_OPERATOR%%$ATTR_objectSid%)(dNSHostName%$ATTR_dNSHostName_OPERATOR%%$ATTR_dNSHostName%)(operatingSystem%$ATTR_operatingSystem_OPERATOR%%$ATTR_operatingSystem%)(operatingSystemVersion%$ATTR_operatingSystemVersion_OPERATOR%%$ATTR_operatingSystemVersion%)(operatingSystemServicePack%$ATTR_operatingSystemServicePack_OPERATOR%%$ATTR_operatingSystemServicePack%)(managedBy%$ATTR_managedBy_OPERATOR%%$ATTR_managedBy%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSCMAS
	:: Main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%	
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo jumpSCMAL
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Name	distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	:: Check User session
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpSCMAO
	:: Session user is a domain user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr logonCount %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
	GoTo jumpSCMAL
:jumpSCMAO
	:: Session user is a local user
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	echo Computer DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (	
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr logonCount %$AD_SERVER_SEARCH%  -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_$lastLogonTimestamp.txt") & (
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -disabled -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	dsget computer "%%N" -loc -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSGET computer "%%N" -memberof -s %$DC%.%$DOMAIN% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% 2>nul >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSCMAL
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSCMA
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"

:skipSCMA
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo SCMA

	
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:sServer
	:: Search Server
	SET $LAST_SEARCH_TYPE=Server
	SET $LAST_SEARCH_ATTRIBUTE=name
	call :SM
	SET $SEARCH_KEY=
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2> nul 1> nul

	echo Use wildcard "*"; if "*" is used alone, will search for all domain controllers. 
	echo ^(If left blank, will abort.^)
	SET /P $SEARCH_KEY=Choose a search key:
	IF NOT DEFINED $SEARCH_KEY GoTo skipsServer
	call :SM
	echo Selected {%$SEARCH_KEY%} as search key.
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY%
	echo %$SEARCH_KEY%> "%$LogPath%\var\var_$SEARCH_KEY.txt"

	:: Check on Wildcard *
	IF "%$SEARCH_KEY%"=="*" (SET $SERVER_SEARCH_GLOBAL=0) ELSE (SET $SERVER_SEARCH_GLOBAL=1)
	echo %$SERVER_SEARCH_GLOBAL% > "%$LogPath%\var\var_$SERVER_SEARCH_GLOBAL.txt"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"	
	:: Search Servers

	IF %$SERVER_SEARCH_GLOBAL% EQU 0 SET $AD_BASE=forestroot

	call :SM
	:: Log output
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$LAST_SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Search Term: %$SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"

	echo Selected {%$SEARCH_KEY%} as search key.

	:: Search type? Domain or Global
	IF /I "%$AD_BASE%"=="forestroot" GoTo sServerG
	
:sServerN
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Search for servers using search key
	:: Unsorted
	if %$SORTED% EQU 1 GoTo sServerNS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -domain %$DOMAIN% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% > "%$LogPath%\var\var_Last_Search_N.txt") else (
		DSQUERY SERVER -domain %$DOMAIN% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -domain %$DOMAIN% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY SERVER -domain %$DOMAIN% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	
GoTo sServerO
	
:sServerNS
	:: Search for servers using search key
	:: sorted
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -domain %$DOMAIN% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% | sort > "%$LogPath%\var\var_Last_Search_N.txt") else (
		DSQUERY SERVER -domain %$DOMAIN% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% | sort -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -domain %$DOMAIN% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY SERVER -domain %$DOMAIN% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)		
	
GoTo sServerO

:sServerG
	echo Global Server search...
	:: Global Server search
	:: unsorted
	if %$SORTED% EQU 1 GoTo sServerGS
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -forest -o rdn -limit %$sLimit% > "%$LogPath%\var\var_Last_Search_N.txt") else (
		DSQUERY SERVER -forest -o rdn -gc -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -forest -o dn -limit %$sLimit% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY SERVER -forest -o dn -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD%> "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo sServerO
	
:sServerGS
	:: Global Server search
	:: sorted
	
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -forest -o rdn -limit %$sLimit% | sort > "%$LogPath%\var\var_Last_Search_N.txt") else (
		DSQUERY SERVER -forest -o rdn -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -forest -o dn -limit %$sLimit% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY SERVER -forest -o dn -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)

:sServerO
	:: Server search main output
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipSO
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Server ^(DC's^) returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Server ^(DC's^) Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	
	:: Check for local or domain user
	if NOT "%$SESSION_USER%"=="%$DOMAIN_USER%" GoTo jumpsSL
	:: Domain User
	echo Forest and Domain Functional Level: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * domainroot -scope base -domain %$DOMAIN% -attr msDS-Behavior-Version >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Global Catalog Domain Controllers: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -isgc -limit %$sLimit% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Schema master of the forest: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo schema -limit %$sLimit% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain naming master of the forest: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo name -limit %$sLimit% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Infrastructure master of the domain: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo infr -limit %$sLimit% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Primary domain controller ^(PDC^) role owner: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo pdc -limit %$sLimit% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Relative identifier master ^(RID master^): >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo rid -limit %$sLimit% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output	
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr name >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr description >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr displayName >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr * >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)

GoTo skipsSL

:jumpsSL
	:: Session user is a local user
	echo Forest and Domain Functional Level: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * domainroot -scope base -domain %$DOMAIN% -attr msDS-Behavior-Version -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Global Catalog Domain Controllers: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -isgc -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Schema master of the forest: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo schema -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain naming master of the forest: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo name -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Infrastructure master of the domain: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo infr -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Primary domain controller ^(PDC^) role owner: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo pdc -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Relative identifier master ^(RID master^): >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY SERVER -forest -o rdn -hasfsmo rid -limit %$sLimit% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output	
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr name -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr description -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr displayName -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * forestroot -filter "(distinguishedName=%%~N)" -attr * -u %$DOMAIN_USER% -p %$cUSERPASSWORD% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:skipsSL

:skipSO
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: skip server output
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 @powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipsServer
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	
:skipsServer
	echo Search Again?
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sServer
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:sOU
	:: Search OU
	SET "$LAST_SEARCH_TYPE=OrganizationalUnit^(OU^)"
	call :SM
	IF NOT DEFINED $SEARCH_KEY (SET $SEARCH_KEY_LAST=NA) ELSE (SET $SEARCH_KEY_LAST=%$SEARCH_KEY%)
	SET $SEARCH_KEY=
	::	Close previous Windows
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2> nul 1> nul
	Echo Use wildcard "*"
	echo If left blank, will abort.
	SET /P $SEARCH_KEY=Choose a search key ^(word^):
	IF NOT DEFINED $SEARCH_KEY GoTo Menu
	IF /I "%$SEARCH_KEY%"=="NA" GoTo jumpsOU
	IF "%$SEARCH_KEY%"=="""" GoTo jumpsOU
	call :SM
	echo Selected {%$SEARCH_KEY%} as search key.
	SET $LAST_SEARCH_KEY=%$SEARCH_KEY%
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	call :subSET
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Start search %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Search Type: %$LAST_SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Search Term: %$SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%" 
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	
	if %$SORTED% EQU 1 GoTo jumpSOU 
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
	if %$SORTED% NEQ 1 GoTo skipSOU
:jumpSOU
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort  > "%$LogPath%\var\var_Last_Search_N.txt") ELSE (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
		)
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% | sort > "%$LogPath%\var\var_Last_Search_DN.txt") ELSE (
		DSQUERY OU %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
		)	
:skipSOU
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"	
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	IF %$LAST_SEARCH_COUNT% EQU 0 (Echo Nothing found! Try again with broader wildcard.)
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo skipOUO
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%" 
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 0 (Echo Nothing found! Try again with broader wildcard.) & (echo.) & (timeout /t 10)
	IF %$LAST_SEARCH_COUNT% EQU 0 GoTo sOU
	:: Search counter increment
	Call :fSC
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$LAST_SEARCH_COUNT% EQU 1 (SET $OU_USE_OPT=Y) ELSE (SET $OU_USE_OPT=N)
	IF /I "%$OU_USE_OPT%"=="Y" (Echo Single OU found!) ELSE (GoTo skipOUO)
	echo Set the OU as search base?
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo skipOUO
	IF %ERRORLEVEL% EQU 1 for /F "skip=1 delims=" %%S IN ('FIND /I "=" "%$LogPath%\%$LAST_SEARCH_LOG%"') DO ECHO %%S> "%$LogPath%\var\var_OU_Base.txt"
	SET /P $AD_BASE= < "%$LogPath%\var\var_OU_Base.txt"
	echo AD Base: %$AD_BASE%
:skipOUO	
	echo Search Again?
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo Menu
	IF %ERRORLEVEL% EQU 1 GoTo sOU
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::




:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:Uset
	::	User Settings
	Color 8E
	mode con:cols=60 lines=40
	cls
	ECHO ************************************************************
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 	%DATE% %TIME%
	echo.
	Echo		Location: Settings    
	ECHO ************************************************************
	Echo.
	Echo Current Log Settings
	Echo ------------------------
	Echo  Log File Path: %$LogPath%
	Echo  Log File Name: %$SESSION_LOG%
	Echo  Keep Log at End: %$kpLog%
	Echo.
	Echo Current Domain Settings
	Echo ------------------------
	Echo  Domain Running Account: %$DOMAIN_USER%
	Echo  Domain Controller: %$DC%
	echo  Domain Site: %$DSITE%
	Echo  Domain: %$domain%
	Echo.
	Echo Current Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo  Query limit: %$sLimit%
	ECHO  Sorted: %$SORTED%
	ECHO  Suppress Console Threshold: %$SUPPRESS_CONSOLE_THRESHOLD%
	ECHO ************************************************************
	Echo.
	Echo Choose an action from the list:
	Echo.
	Echo [1] Change Log Settings
	Echo [2] Change Domain Settings
	Echo [3] Change Search Settings
	echo [4] Search Menu
	Echo [5] Main menu
	Echo.
	Choice /c 12345
	Echo.
	If ERRORLevel 5 GoTo Menu
	If ERRORLevel 4 GoTo Search
	If ERRORLevel 3 GoTo uSetS
	If ERRORLevel 2 GoTo uSetDC
	If ERRORLevel 1 GoTo uSetL
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:UsetL
	::	Log Settings
	mode con:cols=60 lines=40
	cls
	ECHO ************************************************************
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 	%DATE% %TIME%
	echo.
	Echo		Location: Log Settings
	echo.
	ECHO ************************************************************
	Echo.
	Echo Current Log Settings
	Echo ------------------------
	Echo  Log File Path: %$LogPath%
	Echo  Log File Name: %$SESSION_LOG%
	Echo  Keep Log at End: %$kpLog%
	Echo.
	Echo  Instructions
	Echo ------------------------
	Echo.
	Echo If no change is desired,
	Echo just hit enter and leave blank.
	Echo.
	echo %$LOGPATH%> "%$LOGPATH%\var\var_$LOGPATH.txt"
	echo %$SESSION_LOG%> "%$LOGPATH%\var\var_$SESSION_LOG.txt"
	echo %$kpLog%> "%$LOGPATH%\var\var_$kpLog.txt"
	SET /p $LOGPATH=Log Path:
	echo.
	Echo ^("Yes" or "No"^)
	SET /P $kpLog=Keep Logs:
	echo %$kpLog% | FIND /I "Y" && SET $kpLog=Yes
	echo %$kpLog% | FIND /I "N" && SET $kpLog=No
	IF /I NOT "%$kpLog%"=="Yes" SET /A $CHECK_KPLOG+=1
	IF /I NOT "%$kpLog%"=="No" SET /A $CHECK_KPLOG+=1
	IF %$CHECK_KPLOG% EQU 2 SET /P $kpLog= < "%$LOGPATH%\var\var_$kpLog.txt"
	:: ERROR CHECKING
	IF NOT EXIST %$LogPath% mkdir %$LogPath% || Echo Log path not valid and/or file name not valid. Back to default!
	IF NOT EXIST %$LogPath% SET /P $LogPath= < "%$LOGPATH%\var\var_$LOGPATH.txt"
	Echo Close ALL open logs?
	choice /c YN /m "[y]es, [n]o?"
	If ERRORLevel 2 GoTo skipCL	
	If ERRORLevel 1 taskkill /F /IM notepad.exe 2> nul 1> nul
:skipCL	
	GoTo uSet
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

SET "$DC_TAG=DS Settings"
:bannerDS
	mode con:cols=60 lines=40
	cls
	ECHO ************************************************************
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 	%DATE% %TIME%
	echo.
	Echo		Location: Domain Settings [%$DC_TAG%]
	ECHO ************************************************************
	echo.
	Echo Current Domain Settings
	Echo ------------------------
	Echo  Domain Running Account: %$DOMAIN_USER%
	Echo  Domain Controller: %$DC%
	echo  Domain Site: %$DSITE%
	Echo  Domain: %$domain%
	ECHO ************************************************************
	echo.
	GoTo:EOF
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:uSetDC
	SET "$DC_TAG=Domain Settings"
	CALL :bannerDS
	Echo Choose an action from the list:
	Echo.
	Echo [1] Change Domain Running Account
	Echo [2] Change Domain Controller
	Echo [3] Change Domain
	echo [4] Chanage Domain Site
	echo [5] Back Settings
	Echo [6] Main menu
	Echo.
	Choice /c 123456
	Echo.
	If ERRORLevel 6 GoTo Menu
	If ERRORLevel 5 GoTo Uset
	If ERRORLevel 4 GoTo subDS
	If ERRORLevel 3 GoTo subDomain
	If ERRORLevel 2 GoTo subDC
	If ERRORLevel 1 GoTo subDA
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subDA
	::trap Domain must be set first
	IF /I "%$DOMAIN%"=="%COMPUTERNAME%" GoTo subDomain
	IF /I "%$DOMAIN_USER%"=="NA" GoTo jumpDA
	IF %$DU% EQU 0 GoTo jumpDA
	IF /I "%$DC%"=="%COMPUTERNAME%" GoTo subDC
:jumpDA
	::	sub-routin Domain Account
	SET "$DC_TAG=Domain Account"
	CALL :bannerDS
	Echo  Instructions
	Echo ------------------------
	Echo If no change is desired,
	Echo just hit enter and leave blank.
	echo.
	echo %$DOMAIN_USER%> "%$LOGPATH%\var\var_$DOMAIN_USER.txt"
	echo Provide Credentials ^(searches Name ^& UPN^)
	SET /P $DOMAIN_USER=UserName:
	SET /P $cUSERPASSWORD=Password:
	:: Using name search
	dsquery user forestroot -o rdn -scope subtree -domain %$domain% -name "%$DOMAIN_USER%" -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc 2> nul 1> "%$LOGPATH%\var\var_Custom_User_Domain_Authentication.txt"
	::	mainly to capture authentication faliure
	SET $DA_QUERY_RESULT=%ERRORLEVEL%
	IF NOT DEFINED $DA_QUERY_RESULT SET $DA_QUERY_RESULT=0
	echo %$DA_QUERY_RESULT% > "%$LOGPATH%\var\var_$DA_QUERY_RESULT.txt"
	::	Athentication error -2147023570
	IF %$DA_QUERY_RESULT% EQU -2147023570 (
		SET /P $DOMAIN_USER= < "%$LOGPATH%\var\var_$DOMAIN_USER.txt") & (
		ECHO Authentication failed!) & (
		echo.) & (
		Timeout /t 10) & (
		GoTo subDA
		)
	echo.
	SET /P $CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION= < "%$LOGPATH%\var\var_Custom_User_Domain_Authentication.txt"
	::Will contain double quotes to remove
	FOR /F "usebackq delims=" %%P IN ('%$CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION%') DO SET $CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION=%%~P
	:: Using UPN search
	IF NOT DEFINED $CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION dsquery user forestroot -o rdn -scope subtree -domain %$domain% -UPN "%$DOMAIN_USER%*" -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc 2> nul 1> "%$LOGPATH%\var\var_Custom_User_Domain_Authentication.txt"
	SET /P $CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION= < "%$LOGPATH%\var\var_Custom_User_Domain_Authentication.txt"
	::Will contain double quotes to remove
	FOR /F "usebackq delims=" %%P IN ('%$CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION%') DO SET $CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION=%%~P
	SET $DA_VALID=0
	IF NOT DEFINED $CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION SET $DA_VALID=1
	IF %$DA_VALID% EQU 0 SET $DU=1
	IF %$DA_VALID% EQU 1 GoTo subDA
	echo Domain User Name: %$DOMAIN_USER% ^(%$CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION%^) >> "%$LOGPATH%\ADDS_Tool_Active_Session.log"
	echo Success!
	timeout /t 10
	IF /I "%COMPUTERNAME%"=="%$DC%" GoTo subDC
	GoTo uSetDC

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:subDC
	::traps
	IF /I "%$DOMAIN_USER%"=="NA" call :subDA
	::	sub-routine Domain Controller
	SET "$DC_TAG=Domain Controller"
	CALL :bannerDS
	Echo  Instructions
	Echo ------------------------
	Echo If no change is desired,
	Echo just hit enter and leave blank.
	echo %$DC%> "%$LOGPATH%\var\var_$DC.txt"
	echo.
	IF /I "%$SESSION_USER_STATUS%"=="local" IF NOT DEFINED $cUSERPASSWORD GoTo err30
	echo Pick a Domain Controller to connect to
	IF "%$SESSION_USER_STATUS%"=="domain" (dsquery server -o rdn -forest -domain %$domain% -name "*" -limit %$sLimit% -uc 2> nul) ELSE (
		dsquery server -o rdn -forest -domain %$domain% -name "*" -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc 2> nul
		) > "%$LOGPATH%\var\var_Domain_Controller_List.txt"
	type "%$LOGPATH%\var\var_Domain_Controller_List.txt"
	SET /P $DC=Domain Controller:
	:: Validate input
	FIND /I "%$DC%" "%$LOGPATH%\var\var_Domain_Controller_List.txt" 1> nul 2> nul
	SET $DC_INPUT_CHECK=%ERRORLEVEL%
	IF %$DC_INPUT_CHECK% NEQ 0 (
		@powershell Write-Host "Not a valid DC!" -ForegroundColor Red) & (
		echo.) & (
		timeout /t 10) & (
		GoTo subDC
		)
	CALL :bannerDS
	echo Checking...
	:: Validate connection to DC
	if "%$SESSION_USER%"=="%$DOMAIN_USER%" (
		DSQUERY SERVER -forest -o rdn -name %$DC% -s %$DC%) else (
		DSQUERY SERVER -forest -o rdn -name %$DC% -s %$DC% -u %$DOMAIN_USER% -p %$cUSERPASSWORD%
		)
	SET $DC_CHECK=%ERRORLEVEL%
	IF %$DC_CHECK% NEQ 0 SET $DC_CHECK=1
	IF %$DC_CHECK% EQU 1 (
		SET /P $DC= < "%$LOGPATH%\var\var_$DC.txt") & (
		echo DC is not responding! Choose another DC) & (
		timeout /t 10) & (
		GoTo subDC)
	echo.
	echo Domain Controller: %$DC% >> "%$LOGPATH%\ADDS_Tool_Active_Session.log"
	echo Perform PATHPING?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo uSetDC
	IF %ERRORLEVEL% EQU 1 GoTo checkDC
	:checkDC
	CALL :bannerDS
	pathping %$DC%.%$domain%
	echo Change Domain Controller?
	Choice /c yn /m "[y]es or [n]o":
	:: mark may not work
	IF %ERRORLEVEL% EQU 2 GoTo uSetDC
	IF %ERRORLEVEL% EQU 1 GoTo subDC

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subDomain
	::	sub-routine Domain
	SET "$DC_TAG=Domain"
	CALL :bannerDS
	Echo  Instructions
	Echo ------------------------
	Echo If no change is desired,
	Echo just hit enter and leave blank.
	echo %$DOMAIN%> "%$LOGPATH%\var\var_$DOMAIN.txt"
	echo.
	SET /P $DOMAIN=Domain:
	(nslookup %$DOMAIN% 2> nul) | (FIND /I "NAME:")
	SET $CHECK_DOMAIN=%ERRORLEVEL%
	IF %$CHECK_DOMAIN% EQU 1 (SET /P $DOMAIN= < "%$LOGPATH%\var\var_$DOMAIN.txt")
	IF %$CHECK_DOMAIN% EQU 1 (Echo Domain not found!) & (timeout /t 10) & (GoTo subDomain)
	Echo Domain configured: %$DOMAIN%
	echo Domain: %$DOMAIN% >> "%$LOGPATH%\ADDS_Tool_Active_Session.log"
	timeout /t 10
	IF %$DU% EQU 1 GoTo subDA
	GoTo uSetDC
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subDS
	:: Domain Sites
	::traps
	echo %COMPUTERNAME% | (FIND /I "%$DOMAIN%") && (GoTo subDomain)
	IF %$DU% EQU 0 call :subDA
	IF /I "%$DC%"=="%COMPUTERNAME%" call :subDC
	:: sub-routine Domain Site
	SET "$DC_TAG=Domain Site"
	CALL :bannerDS
	Echo  Instructions
	Echo ------------------------
	Echo If no change is desired,
	Echo just hit enter and leave blank.
	Echo Can use "*" at end of string.
	echo %$DSITE%> "%$LOGPATH%\var\var_$DSITE.txt"
	echo.
	IF %$DU% EQU 0 IF NOT DEFINED $cUSERPASSWORD GoTo err30
	echo Choose a site
	Echo --------------
	IF "%$SESSION_USER_STATUS%"=="domain" (dsquery site -o rdn -domain %$domain% -limit %$sLimit% -uc) ELSE (
		dsquery site -o rdn -domain %$domain% -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc 2> nul
		)
	SET /P $DSITE=Domain Site:
	SET $DS_CHECK=0
	IF "%$SESSION_USER_STATUS%"=="domain" (
		dsquery site -o rdn -domain %$domain% -name "%$DSITE%" -limit %$sLimit% -uc | FIND /I """") ELSE (
		dsquery site -o rdn -domain %$domain% -name "%$DSITE%" -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc | FIND /I """"
		)
	SET $DS_CHECK=%ERRORLEVEL%
	IF %$DS_CHECK% NEQ 0 SET /P $DSITE= < "%$LOGPATH%\var\var_$DSITE.txt"
	IF %$DS_CHECK% NEQ 0 Echo Not a valid site, reverted back to previous setting.
	IF %$DS_CHECK% EQU 0 GoTo jumpDSS
	echo Try to change Domain Site again?
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo uSetDC
	IF %ERRORLEVEL% EQU 1 GoTo subDS
	:jumpDSS
	IF "%$SESSION_USER_STATUS%"=="domain" (
		dsquery site -o rdn -domain %$domain% -name "%$DSITE%" -limit %$sLimit% -uc > "%$LOGPATH%\var\var_$DSITE_List.txt") ELSE (
		dsquery site -o rdn -domain %$domain% -name "%$DSITE%" -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc > "%$LOGPATH%\var\var_$DSITE_List.txt"
		)
	SET /P $DSITE_N= < "%$LOGPATH%\var\var_$DSITE_List.txt"
	::Will contain double quotes to remove
	echo %$DSITE% | (FIND /I "*" 1> nul 2> nul) && ( 
		FOR /F "usebackq delims=" %%P IN ('%$DSITE_N%') DO SET $DSITE_N=%%~P)
	SET $DSITE=%$DSITE_N%
	echo %$DSITE%> "%$LOGPATH%\var\var_$DSITE.txt"
	Echo Success!
	timeout /t 10 
	GoTo uSetDC
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:uSetS
:: Search Settings
	mode con:cols=60 lines=40
	cls
	ECHO ************************************************************
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 	%DATE% %TIME%
	echo.
	Echo		Location: Search Settings
	echo.
	ECHO ************************************************************
	Echo.
	IF %$SEARCH_SETTINGS_CHECK% EQU 1 GoTo jumpSSC
	Echo Current Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo  Suppress Console Threshold: %$SUPPRESS_CONSOLE_THRESHOLD%
	Echo  Sort: %$SORTED_N%
	Echo ------------------------
	echo.
	Echo %$AD_BASE%> "%$LOGPATH%\var\var_AD_Base.txt"
	Echo %$AD_SCOPE%> "%$LOGPATH%\var\var_AD_Scope.txt"
	Echo Select AD Base:
	echo [1] domainroot
	echo [2] forestroot
	echo [3] custom OU
	echo [4] load from previous
	Echo.
	Choice /c 1234
	Echo.
	If ERRORLevel 4 GoTo subADB
	If ERRORLevel 3 GoTo sOU
	If ERRORLevel 2 SET $AD_BASE=forestroot
	If ERRORLevel 1 SET $AD_BASE=domainroot
	
:subADbase
	REM if string already contains double-quotes, string comparison in quotes will crash;
	REM e.g. "%$AD_BASE%"=="forestroot" --^> ""string""=="forestroot"
	echo %$AD_BASE% | (FIND /I "=" 2> nul) & SET $AD_BASE_CUSTOM=%ERRORLEVEL%
	echo %$AD_BASE_CUSTOM% > "%$LOGPATH%\var\var_$AD_BASE_CUSTOM.txt"
	if %$AD_BASE_CUSTOM% EQU 0 GoTo skipADC
	:: Not a custom OU base
	if /I "%$AD_BASE%"=="forestroot" (SET $AD_SCOPE=subtree) & (GoTo skipASS)

:skipADC
	echo Select AD Scope:
	echo [1] subtree
	echo [2] onelevel
	echo [3] base
	echo.
	Choice /c 123
	Echo.	
	If ERRORLevel 3 SET $AD_SCOPE=base
	If ERRORLevel 2 SET $AD_SCOPE=onelevel
	If ERRORLevel 1 SET $AD_SCOPE=subtree
:skipASS


	
	echo %$SUPPRESS_CONSOLE_THRESHOLD% > "%$LOGPATH%\var\var_Suppress_Console_Threshold.txt"
:jumpSCT	
	echo Provide a number
	SET /P $SUPPRESS_CONSOLE_THRESHOLD=Suppress Console Threshold:
	echo %$SORTED% > "%$LOGPATH%\var\var_$SORTED.txt"
	echo Provide {0 [No] , 1 [Yes]}
	SET /P $SORTED=Sorted:
	IF %$SORTED% LEQ 1 GoTo skipSC
	echo %$SORTED% | (find /I "Y")
	IF %ERRORLEVEL% EQU 0 (SET $SORTED=1) ELSE (SET $SORTED=0)
:skipSC
	IF %$SORTED% EQU 1 (SET $SORTED_N=Yes) ELSE (SET $SORTED_N=No)
	SET $SEARCH_SETTINGS_CHECK=1
	GoTo uSetS
:jumpSSC
	echo New Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo  Suppress Console Threshold: %$SUPPRESS_CONSOLE_THRESHOLD%
	Echo  Sort: %$SORTED_N%
	echo.
	echo Change Search settings?
	SET $SEARCH_SETTINGS_CHECK=0
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo Uset
	IF %ERRORLEVEL% EQU 1 GoTo uSetS
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subADB
	:: Subroutine to set AD Base with OU
	IF NOT EXIST "%$LOGPATH%\var\var_OU_Base.txt" GoTo skipADB
	SET /P $AD_BASE= < "%$LOGPATH%\var\var_OU_Base.txt"
:skipADB
	IF NOT DEFINED $AD_BASE GoTo uSetS
	IF NOT EXIST "%$LOGPATH%\var\var_OU_Base.txt" GoTo uSetS
	GoTo subADbase
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:Logs
	IF EXIST "%$LOGPATH%\ADDS_Tool_Active_Session.log" @explorer "%$LOGPATH%\ADDS_Tool_Active_Session.log"
	IF EXIST "%$LOGPATH%\ADDS_Search_Session.log" @explorer "%$LOGPATH%\ADDS_Search_Session.log"
	IF EXIST "%$LOGPATH%\ADDS_Tool_Last_Search.log" @explorer "%$LOGPATH%\ADDS_Tool_Last_Search.log"
	IF EXIST "%$LOGPATH%\var\var_Last_Search_N_DN.txt" @explorer "%$LOGPATH%\var\var_Last_Search_N_DN.txt"
	GoTo menu
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:::: FUNCTIONS ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:fSC
	::	Search Counter
	SET /A $COUNTER_SEARCH+=1
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subOperator
	:: Subroutine for search operators
	SET $SEARCH_TYPE_ATTR_%1=%1
	@powershell Write-Host " search operator:" -ForegroundColor Blue
	echo [1] Equal [=]
	echo [2] Approximately equal to [^~=]
	echo [3] Less [^<=]
	echo [4] Greater [^>=]
	echo.
	Choice /c 1234
	If ERRORLevel 4 (SET "$ATTR_%1_OPERATOR=>=") & (SET "$ATTR_%1_OPERATOR_DISPLAY=^>^=")
	If ERRORLevel 3 (SET "$ATTR_%1_OPERATOR=<=") & (SET "$ATTR_%1_OPERATOR_DISPLAY=^<^=")
	If ERRORLevel 2 (SET "$ATTR_%1_OPERATOR=~=") & (SET "$ATTR_%1_OPERATOR_DISPLAY=^~^=")
	If ERRORLevel 1 (SET "$ATTR_%1_OPERATOR==") & (SET "$ATTR_%1_OPERATOR_DISPLAY=^=")
GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:fVarD
	:: Function Variable Debug
	set | FINDSTR /B /C:"$" > "%$LOGPATH%\var\Variable_Debug_%$PID%.txt"
GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::











::	jump error section 
GoTo end

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
:: ERROR SECTION
::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:ErrBann
:: Banner
	cls
	color 4E
	mode con:cols=80
	mode con:lines=40
	ECHO   ***************************************************************************
	ECHO. 
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 %DATE% %TIME%
	ECHO.
	ECHO   ***************************************************************************
	ECHO   ***************************************************************************
	echo.
	Echo		!!ERROR!! !!ERROR!! !!ERROR!! !!ERROR!! !!ERROR!! !!ERROR!!
	echo.
	ECHO   ***************************************************************************
	echo.
	echo.
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:err10
	:: error Administrative Privilege
	call :ErrBann
	echo Administrative Privilege Error
	Echo.
	echo Current user doesn't have administrative privilege!
	Echo.
	echo This is likely a fatal error!
	echo Try running the program as an administrator.
	echo ^(An action is required (likely installing dependencies),
	echo  which requires administrative privilege.^)
	echo.
	echo Aborting!
	echo.
	Timeout/t 120
	GoTo end
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:err20
	::error RSAT
	call :ErrBann
	echo RSAT-Remote Server Administration Tools ERROR
	echo.
	echo Something went wrong trying to install RSAT!
	echo RSAT is a core dependency for this program.
	echo This is a fatal error!
	echo.
	echo Aborting!
	echo.
	Timeout/t 120
	GoTo end
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:err30
	:: Error PW Cache
	:: 
	call :ErrBann
	echo Logged in User is not a domain user, and no PW cached!
	echo Custom domain user requires that the password be cached.
	echo (Password is cached in memory)
	Echo Jumping to allow setting custom user and password...
	echo.
	timeout /t 60
	GoTo subDA
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:err40
	:: Error Under Development
	call :ErrBann
	echo	UNDER CONSTRUCTION
	echo.
	echo	Feature: %$LAST_SEARCH_TYPE% search
	echo.
	timeout /t 60
GoTo Search
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:end
	:: End session
	IF EXIST "%$LOGPATH%\%$SESSION_LOG%" Echo End Session %DATE% %TIME%. >> "%$LOGPATH%\%$SESSION_LOG%"
	IF EXIST "%$LOGPATH%\%$SESSION_LOG%" Echo. >> "%$LOGPATH%\%$SESSION_LOG%"
	IF EXIST "%$LogPath%\var\var_$PID.txt" del /q "%$LogPath%\var\var_$PID.txt"
	:: [FUTURE FEATURE]
	::	Save Session Settings
	:: IF /I NOT "%$SAVE_SETTINGS%"=="Yes" GoTo skipSSS
	:: IF NOT EXIST "%$LOGPATH%\Settings" mkdir "%$LOGPATH%\Settings"
	:: :skipSSS	
	
	::	Check for debug mode
	IF %$DEGUB_MODE% EQU 1 GoTo skipCL
	:: Last Search files
	IF EXIST "%$LOGPATH%\%$LAST_SEARCH_LOG%" Del /q "%$LOGPATH%\%$LAST_SEARCH_LOG%"
	IF EXIST "%$LOGPATH%\var\var_Last_Search_N_DN.txt" Del /q "%$LOGPATH%\var\var_Last_Search_N_DN.txt"
	IF EXIST "%$LOGPATH%\var" RD /S /Q "%$LOGPATH%\var"
:skipCL
	:: Archive session
	Type "%$LOGPATH%\%$SESSION_LOG%" >> "%$LOGPATH%\%$ARCHIVE_LOG%"
	Del /q "%$LOGPATH%\%$SESSION_LOG%"
	Type "%$LOGPATH%\%$SEARCH_SESSION_LOG%" >> "%$LOGPATH%\%$ARCHIVE_SEARCH_LOG%"
	Del /q "%$LOGPATH%\%$SEARCH_SESSION_LOG%"
	IF %$DEGUB_MODE% EQU 1 GoTo skipLC
	:: Keep logs check
	IF /I %$KPLOG%==Yes IF EXIST "%$LOGPATH%\ReadMe.txt" Del /q "%$LOGPATH%\ReadMe.txt"
	IF /I %$KPLOG%==Yes GoTo skipLC
	::	Delete all logs
		:: Close any open files
	taskkill /F /FI "WINDOWTITLE eq ADDS*"
	IF EXIST "%$LOGPATH%" RD /S /Q "%$LOGPATH%"
	echo %DATE% %TIME% > "%$LOGPATH%\ReadMe.txt"
	echo Directory was nuked! >> "%$LOGPATH%\ReadMe.txt"
:skipLC
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:credits
	::	Credits
	cls
	mode con:cols=55 lines=25
	COLOR 0B
	Echo.
	ECHO Developed by:
	ECHO David Geeraerts {dgeeraerts.evergreen@gmail.com}
	ECHO GitHub: https://github.com/DavidGeeraerts/ADDS_Tool
	ECHO.
	Echo.
	ECHO Contributors:
	ECHO.
	Echo.
	Echo.
	ECHO.
	ECHO.
	ECHO Copyleft License
	ECHO GNU GPL (General Public License)
	ECHO https://www.gnu.org/licenses/gpl-3.0.en.html
	Echo.
	Timeout /T 30
	ENDLOCAL
Exit
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::