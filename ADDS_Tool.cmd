:::: Active Directory Domain Services Tool [ADDS] :::::::::::::::::::::::::::::

::#############################################################################
::							#DESCRIPTION#
::
::	SCRIPT STYLE: Interactive
::	Program is a wrapper for ADDS (Active Directory Domain Services)
::	Active Directory search's
::#############################################################################

:::: Developer ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Author:		David Geeraerts
:: Location:	Olympia, Washington USA
:: E-Mail:		dgeeraerts.evergreen@gmail.com
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: GitHub :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::	https://github.com/DavidGeeraerts/ADDS_Tool
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: License ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Copyleft License(s)
:: GNU GPL v3 (General Public License)
:: https://www.gnu.org/licenses/gpl-3.0.en.html
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Versioning Schema ::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::		VERSIONING INFORMATION												 ::
::		Semantic Versioning used											 ::
::		http://semver.org/													 ::
::		Major.Minor.Revision												 ::
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Stopwatch start ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
@SET $START_LOAD_TIME=%TIME%
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Command shell ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
@Echo Off
@SETLOCAL enableextensions
SET $PROGRAM_NAME=Active_Directory_Domain_Services_Tool
SET $Version=0.15.0
SET $BUILD=2021-04-09 10:00
Title %$PROGRAM_NAME%
Prompt ADT$G
color 8F
mode con:cols=80 lines=56
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::: Configuration File ::::::::::::::::::::::::::::::::::::::::::::::::::::::::
SET $CONFIG_FILE=ADDS_Tool.config
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Configuration - Basic ::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Declare Global variables
:: All User variables are set within here.
:: Defaults
::	uses user profile location for logs
SET "$LOGPATH=%APPDATA%\ADDS"
SET $SESSION_LOG=ADDS_Tool_Active_Session.log
SET $SEARCH_SESSION_LOG=ADDS_Tool_Session_Search.log
SET $LAST_SEARCH_LOG=ADDS_Tool_Last_Search.log
SET $ARCHIVE_LOG=ADDS_Tool_Session_Archive.log
SET $ARCHIVE_SEARCH_LOG=ADDS_Tool_Search_Archive.log
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Configuration - Advanced :::::::::::::::::::::::::::::::::::::::::::::::::
:: Advanced Settings

:: Suppress_Verbose Output on searches
::	0=Off {no}; 1=On {yes}
SET $SUPPRESS_VERBOSE=0

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
SET $DEGUB_MODE=0
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::#############################################################################
::	!!!!	Everything below here is 'hard-coded' [DO NOT MODIFY]	!!!!
::#############################################################################

:::: Default Program Variables ::::::::::::::::::::::::::::::::::::::::::::::::
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
SET $ADGROUP.N=NA
SET $SEARCH_TYPE=NA
SET $SEARCH_KEY=NA
SET $LAST_SEARCH_COUNT=NA
SET $DOMAIN=%USERDNSDOMAIN%
SET $DSITE=Default
IF NOT DEFINED $DOMAIN SET $DOMAIN=NA
REM Doesn't like On Off words
IF %$SORTED% EQU 1 (SET $SORTED_N=Yes) ELSE (SET $SORTED_N=No)
IF %$SUPPRESS_VERBOSE% EQU 0 (SET $SUPPRESS_VERBOSE_N=No) ELSE (SET $SUPPRESS_VERBOSE_N=Yes)
SET $SEARCH_SETTINGS_CHECK=0
SET $SEARCH_ATTRIBUTE=name
:: Defaults
SET $AD_BASE=domainroot
SET $AD_SCOPE=subtree
SET "$AD_SERVER_SEARCH=-s %$DC%"
:: Dependency Checks
::	assumes ready to go
SET $PREREQUISITE_STATUS=1
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::###########################################################################::
:: CONFIGURATION FILE OVERRIDE
::###########################################################################::

IF NOT EXIST "%~dp0\%$CONFIG_FILE%" Goto skipCF

:: FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"<$VARIABLE>" "%~dp0\%$CONFIG_FILE%"') DO SET "<$VARIABLE>=%%V"

::   ADDS_TOOL_CONFIG_SCHEMA_VERSION
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$CONFIG_SCHEMA_VERSION" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_SCHEMA_VERSION=%%V"
::	Logging
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$LOGPATH" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_LOGPATH=%%V"
IF DEFINED $CONFIG_LOGPATH SET "$LOGPATH=%$CONFIG_LOGPATH%"
FOR /F %%R IN ('ECHO %$LOGPATH%') DO SET $LOGPATH=%%R
::	Session log
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$SESSION_LOG" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_SESSION_LOG=%%V"
IF DEFINED $CONFIG_SESSION_LOG SET "$SESSION_LOG=%$CONFIG_SESSION_LOG%"
::	Session Search log
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$SEARCH_SESSION_LOG" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_SEARCH_SESSION_LOG=%%V"
IF DEFINED $CONFIG_SEARCH_SESSION_LOG SET "$SEARCH_SESSION_LOG=%$CONFIG_SEARCH_SESSION_LOG%"
::	Last search log
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$LAST_SEARCH_LOG" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_LAST_SEARCH_LOG=%%V"
IF DEFINED $CONFIG_LAST_SEARCH_LOG SET "$LAST_SEARCH_LOG=%$CONFIG_LAST_SEARCH_LOG%"
::	Archive Session log
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$ARCHIVE_LOG" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_ARCHIVE_LOG=%%V"
IF DEFINED $CONFIG_ARCHIVE_LOG SET "$ARCHIVE_LOG=%$CONFIG_ARCHIVE_LOG%"
::	Archive Search log
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$ARCHIVE_SEARCH_LOG" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_ARCHIVE_SEARCH_LOG=%%V"
IF DEFINED $CONFIG_ARCHIVE_SEARCH_LOG SET "$ARCHIVE_SEARCH_LOG=%$CONFIG_ARCHIVE_SEARCH_LOG%"

:: Search Defaults
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$sLimit" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_sLimit=%%V"
IF DEFINED $CONFIG_sLimit SET "$sLimit=%$CONFIG_sLimit%"
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$AD_BASE" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_AD_BASE=%%V"
IF DEFINED $CONFIG_AD_BASE SET "$AD_BASE=%$CONFIG_AD_BASE%"
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$AD_SCOPE" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_AD_SCOPE=%%V"
IF DEFINED $CONFIG_AD_SCOPE SET "$AD_SCOPE=%$CONFIG_AD_SCOPE%"

::	Credentials
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$DOMAIN_USER" "%~dp0\%$CONFIG_FILE%"') DO SET "$DOMAIN_USER=%%V"
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$DOMAIN_USER_PASSWORD" "%~dp0\%$CONFIG_FILE%"') DO SET "$DOMAIN_USER_PASSWORD=%%V"
REM Friendly name to variable name
IF DEFINED $DOMAIN_USER_PASSWORD SET $cUSERPASSWORD=%$DOMAIN_USER_PASSWORD%

:: Advanced Settings
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$DEGUB_MODE" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_DEGUB_MODE=%%V"
IF DEFINED $CONFIG_DEGUB_MODE SET "$DEGUB_MODE=%$CONFIG_DEGUB_MODE%"
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$SUPPRESS_VERBOSE" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_SUPPRESS_VERBOSE=%%V"
IF DEFINED $CONFIG_SUPPRESS_VERBOSE SET "$SUPPRESS_VERBOSE=%$CONFIG_SUPPRESS_VERBOSE%"
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$SORTED" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_SORTED=%%V"
IF DEFINED $CONFIG_SORTED SET "$SORTED=%$CONFIG_SORTED%"
FOR /F "tokens=2 delims=^=" %%V IN ('FINDSTR /BC:"$KPLOG" "%~dp0\%$CONFIG_FILE%"') DO SET "$CONFIG_KPLOG=%%V"
IF DEFINED $CONFIG_KPLOG SET "$KPLOG=%$CONFIG_KPLOG%"

REM variable name to Friendly name
IF %$SORTED% EQU 1 (SET $SORTED_N=Yes) ELSE (SET $SORTED_N=No)
IF %$SUPPRESS_VERBOSE% EQU 0 (SET $SUPPRESS_VERBOSE_N=No) ELSE (SET $SUPPRESS_VERBOSE_N=Yes)

:skipCF


:::: Directory ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:CD
	:: Launched from directory
	SET "$PROGRAM_PATH=%~dp0"
	::	Setup logging
	IF NOT EXIST "%$LOGPATH%\var" MD "%$LOGPATH%\var"
	cd /D "%$LOGPATH%"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: PID ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:PID
	:: Program information including PID
	tasklist /FI "WINDOWTITLE eq %$PROGRAM_NAME%*" > "%$LogPath%\var\var_TaskInfo_PID.txt"
	for /F "skip=3 tokens=2 delims= " %%P IN ('tasklist /FI "WINDOWTITLE eq %$PROGRAM_NAME%*"') DO echo %%P> "%$LogPath%\var\var_$PID.txt"
	SET /P $PID= < "%$LogPath%\var\var_$PID.txt"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: fISO8601 :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:fISO8601
	:: Function to ensure ISO 8601 Date format yyyy-mmm-dd
	:: Easiest way to get ISO date
	@powershell Get-Date -format "yyyy-MM-dd" > "%$LogPath%\var\var_ISO8601_Date.txt"
	SET /P $ISO_DATE= < "%$LogPath%\var\var_ISO8601_Date.txt"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: UTC ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:UTC
	:: Universal Time Coordinate
	IF EXIST "%$LogPath%\var\var_$UTC.txt" SET /P $UTC= < "%$LogPath%\var\var_$UTC.txt"
	IF NOT DEFINED $UTC FOR /F "tokens=1 delims=()" %%P IN ('wmic timezone get Description ^| findstr /C:"UTC" /I') DO ECHO %%P > "%$LogPath%\var\var_$UTC.txt"
	IF NOT DEFINED $UTC SET /P $UTC= < "%$LogPath%\var\var_$UTC.txt"
	IF EXIST "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt" SET /P $UTC_STANDARD_NAME= < "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt"
	IF NOT DEFINED $UTC_STANDARD_NAME FOR /F "tokens=2 delims==" %%P IN ('wmic timezone get StandardName /value ^| findstr /C:"=" /I') DO ECHO %%P > "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt"
	IF NOT DEFINED $UTC_STANDARD_NAME SET /P $UTC_STANDARD_NAME= < "%$LogPath%\var\var_$UTC_STANDARD_NAME.txt"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Session Log Header :::::::::::::::::::::::::::::::::::::::::::::::::::::::
:wLog
	:: Start session and write to log
	Echo Start Session %DATE% %TIME% > "%$LogPath%\%$SESSION_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Program Name: %$PROGRAM_NAME% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Program Version: %$Version% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Program Build: %$BUILD% >> "%$LogPath%\%$SESSION_LOG%"
	IF DEFINED $CONFIG_SCHEMA_VERSION echo Program config schema: %$CONFIG_SCHEMA_VERSION% >> "%$LogPath%\%$SESSION_LOG%"
	echo Program Path: %$PROGRAM_PATH% >> "%$LogPath%\%$SESSION_LOG%"
	Echo PC: %COMPUTERNAME% >> "%$LogPath%\%$SESSION_LOG%"
	echo PC Domain: %USERDOMAIN% >> "%$LogPath%\%$SESSION_LOG%"
	Echo Session User: %USERNAME% >> "%$LogPath%\%$SESSION_LOG%"
	echo PID: %$PID% >> "%$LogPath%\%$SESSION_LOG%"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Computer Domain ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:DUC
	::	Check for Domain computer
	::	If value is 1 domain, is 0 workgroup
	SET $DOMAIN_PC=1
	wmic computersystem get DomainRole /value | (FIND "0") && (SET $DOMAIN_PC=0)
	echo %$DOMAIN_PC% > "%$LogPath%\var\var_$DOMAIN_PC.txt"
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
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Administrator Privilege Check ::::::::::::::::::::::::::::::::::::::::::::
:subA
	openfiles.exe 1> "%$LOGPATH%\var\var_$Admin_Status_M.txt" 2> "%$LOGPATH%\var\var_$Admin_Status_E.txt"
	SET $ADMIN_STATUS=0
	FIND "ERROR:" "%$LOGPATH%\var\var_$Admin_Status_E.txt" 2> nul > nul  && (SET $ADMIN_STATUS=1)
	IF %$ADMIN_STATUS% EQU 0 (SET "$ADMIN_STATUS_N=Yes") ELSE (SET "$ADMIN_STATUS_N=No")
	echo %$ADMIN_STATUS_N%> "%$LOGPATH%\var\var_$Admin_Status_N.txt"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Check RSAT-Remote Server Administration Tools ::::::::::::::::::::::::::::
:CheckRSAT
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
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: RSAT Installer :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
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
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Stopwatch - Start ::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:Start
:: Capture program load time
	@PowerShell.exe -c "$span=([TimeSpan]'%Time%' - [TimeSpan]'%$START_LOAD_TIME%'); '{0:00}:{1:00}:{2:00}.{3:00}' -f $span.Hours, $span.Minutes, $span.Seconds, $span.Milliseconds" > "%$LogPath%\var\var_Load_Time.txt"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Parameters :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:Param
	:: Parameter #1 Search Type
	SET $PARAMETER1=%~1
	IF NOT DEFINED $PARAMETER1 GoTo skipParam
	set $SEARCH_TYPE=%$PARAMETER1%
	echo %$SEARCH_TYPE%> "%$LOGPATH%\var\var_$SEARCH_TYPE.txt"
	:: Parameter #2 Attribute
	SET $PARAMETER2=%~2
	IF NOT DEFINED $PARAMETER2 GoTo skipParam
	set $SEARCH_ATTRIBUTE=%$PARAMETER2%
	echo %$SEARCH_ATTRIBUTE%> "%$LOGPATH%\var\var_$SEARCH_ATTRIBUTE.txt"
	:: Parameter #3 Search Key
	SET $PARAMETER3=%~3
	IF NOT DEFINED $PARAMETER3 GoTo skipParam
	set $SEARCH_KEY=%$PARAMETER3%
	echo %$SEARCH_KEY%> "%$LOGPATH%\var\var_$SEARCH_KEY.txt"
	:: Automated Search
	GoTo SAUTO
:skipParam
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Main Menu ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
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
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Menu ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:Search
	Color 0A
	mode con:cols=80 lines=52
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
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::	##	START SUBROUTINES	##	:::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Menu Banner :::::::::::::::::::::::::::::::::::::::::::::::::::::::
:SMB
	Color 0A
	mode con:cols=55 lines=40
	Cls
	ECHO ******************************************************
	ECHO		%$PROGRAM_NAME%
	echo.
	echo		 	%DATE% %TIME%
	ECHO.
	Echo		Location: Search Menu
	Echo ******************************************************
	Echo.
	Echo Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo  Query limit: %$sLimit%
	echo  Sorted: %$SORTED_N%
	echo  Suppress Verbose: %$Suppress_Verbose_N%
	Echo  Search count: %$COUNTER_SEARCH%
	Echo ******************************************************
	Echo.
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Menu Dashboard ::::::::::::::::::::::::::::::::::::::::::::::::::::
:SM
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
	echo  Suppress Verbose: %$Suppress_Verbose_N%
	echo.
	echo Search HUD
	Echo ------------------------
	Echo  Search Type: %$SEARCH_TYPE%
	echo  Search Attribute: %$SEARCH_ATTRIBUTE%
	echo  Search Key: %$SEARCH_KEY%
	echo  Search Results: %$LAST_SEARCH_COUNT%
	Echo  Search count: %$COUNTER_SEARCH%
	Echo ******************************************************
	echo.
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::: Sub-routin for Search Key :::::::::::::::::::::::::::::::::::::::::::::::
:subSK

:: # Future development
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Start Elapse Time ::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:subSET
	SET $START_TIME=%TIME%
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Total Lapse Time :::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:subTLT
	@PowerShell.exe -c "$span=([TimeSpan]'%Time%' - [TimeSpan]'%$START_TIME%'); '{0:00}:{1:00}:{2:00}.{3:00}' -f $span.Hours, $span.Minutes, $span.Seconds, $span.Milliseconds" > "%$LogPath%\var\var_Total_Lapsed_Time.txt"
	SET /P $TOTAL_LAPSE_TIME= < "%$LogPath%\var\var_Total_Lapsed_Time.txt"
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Log Header ::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:sHeader
	IF EXIST "%$LogPath%\%$LAST_SEARCH_LOG%" DEL /Q "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Start search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo UTC: %$UTC% %$UTC_STANDARD_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Search Type: %$SEARCH_TYPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Attribute: %$SEARCH_ATTRIBUTE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key: %$SEARCH_KEY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Sorted: %$SORTED_N% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Root: %$AD_BASE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search AD Scope: %$AD_SCOPE% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Domain Controller: %$DC% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	Echo Domain: %$DOMAIN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Last Search Log Close - Notepad ::::::::::::::::::::::::::::::::::::::::::
:LSLCN
	taskkill /F /FI "WINDOWTITLE eq %$LAST_SEARCH_LOG% - Notepad" 2>nul 1>nul
GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::	##	END SUBROUTINES	##	:: ::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Universal :::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:sUniversal
	:: Search Universal
	:: 	Reset search variables for HUD
	SET $SEARCH_TYPE=Universal
	SET $SEARCH_ATTRIBUTE=
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	::	Close previous Windows
	::	Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	SET "$ATTRIBUTES_Universal=name cn displayName distinguishedName extensionAttribute<#> objectSid sAMAccountName"
	@powershell Write-Host "Universal Attributes:" -ForegroundColor Gray
	@powershell Write-Host "%$ATTRIBUTES_Universal%" -ForegroundColor Blue
	:: User input

	@powershell Write-Host "Choose attribute to search against:" -ForegroundColor Gray
	@powershell Write-Host "default is name, leave blank for default" -ForegroundColor Magenta
	SET /P $SEARCH_ATTRIBUTE=Attribute:
	IF NOT DEFINED $SEARCH_ATTRIBUTE SET $SEARCH_ATTRIBUTE=name
	call :SM
	@powershell Write-Host "Attribute: %$SEARCH_ATTRIBUTE%" -ForegroundColor Blue
	@powershell Write-Host "can use wildcard * :  Key* *Key *key*" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort!" -ForegroundColor Red

	SET /P $SEARCH_KEY=Choose a search key:
	IF NOT DEFINED $SEARCH_KEY GoTo Search
	call :SM
	echo Selected {%$SEARCH_KEY%} as search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Start Elapse Time
	call :subSET
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUS
	::	unsorted
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" -attr name displayName distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUS
:jumpSUS
	::	sorted
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" -attr name displayName distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=*)(%$SEARCH_ATTRIBUTE%=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSUS

	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%K > "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUC
	)

	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	:: Output search results
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo Name					displayName			distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	REM powershell write output adds slight overhead.
	REM	e.g. 293 sarch result:
	REM powershell	Total Search Time: 00:07:52.610
	REM echo		Total Search Time: 00:05:34.370
	FOR /F "tokens=* delims=" %%N IN (%$LogPath%\var\var_Last_Search_DN.txt) DO (
		(call :SM) & (
		@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow) & (
		@powershell Write-Host '%%N' -ForegroundColor DarkGray) & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr name %$AD_SERVER_SEARCH%  %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr description %$AD_SERVER_SEARCH%  %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr displayName %$AD_SERVER_SEARCH%  %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%N)" -attr distinguishedName %$AD_SERVER_SEARCH%  %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		DSQUERY * -filter "(distinguishedName=%%N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%") & (
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%")
	)
:jumpSUO
	call :subTLT
	:: Search counter increment
	Call :fSC
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
:skipSUC
	@powershell Write-Host "Search Again?" -ForegroundColor Gray
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sUniversal
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search User Menu :::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:sUser
	::	Close previous Windows
	call :LSLCN
	SET $SEARCH_TYPE=User
	SET $SEARCH_ATTRIBUTE=
	SET $SEARCH_KEY=
	call :SM
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

::::: Search User Name ::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:SUN
	:: Reset Variables
	SET $SEARCH_ATTRIBUTE=name
	SET $SEARCH_KEY=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: Name" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort." -ForegroundColor Magenta

	SET /P $SEARCH_KEY=User name search key:
	IF NOT DEFINED $SEARCH_KEY GoTo sUser
	call :SM
	echo Selected {%$SEARCH_KEY%} as name search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUNS
	:: Unsorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUNS

:jumpSUNS
	:: Sorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSUNS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUN
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUNO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		call :SM
		@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
		@powershell Write-Host '%%N' -ForegroundColor DarkGray
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name userPrincipalName displayName %$AD_SERVER_SEARCH%  %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSUNO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUN
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::: Search User UPN :::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:SUU
	SET $SEARCH_ATTRIBUTE=UPN
	SET $SEARCH_KEY=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: UPN" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort." -ForegroundColor Magenta
	SET /P $SEARCH_KEY=User upn search key:
	IF NOT DEFINED $SEARCH_KEY GoTo sUser
	call :SM
	echo Selected {%$SEARCH_KEY%} as upn search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUUS
	:: Unsorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -upn "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -upn "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"

	GoTo skipSUUS
:jumpSUUS
	:: Sorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -upn "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -upn "%$SEARCH_KEY%" -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSUUS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUU
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUUO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName displayName %$AD_SERVER_SEARCH%  %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSUUO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUU
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUU
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::: Search User first and last name :::::::::::::::::::::::::::::::::::::::::
:SUFL
	SET $SEARCH_TYPE=user
	SET $SEARCH_ATTRIBUTE=FirstLast-Names
	SET $SEARCH_KEY=
	SET $SEARCH_KEY_USER_FIRST=
	SET $SEARCH_KEY_USER_LAST=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: sn givenName" -ForegroundColor Blue
	@powershell Write-Host "Attribute: sn:Surname LastName" -ForegroundColor Blue
	@powershell Write-Host "Attribute: givenName FirstName" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort." -ForegroundColor Magenta
	SET /P $SEARCH_KEY_USER_FIRST=User FirstName search key:
	IF NOT DEFINED $SEARCH_KEY_USER_FIRST GoTo sUser
	SET /P $SEARCH_KEY_USER_LAST=User LastName search key:
	IF NOT DEFINED $SEARCH_KEY_USER_LAST GoTo sUser
	SET $SEARCH_KEY=%$SEARCH_KEY_USER_FIRST% %$SEARCH_KEY_USER_LAST%
	call :SM
	echo Selected {%$SEARCH_KEY_USER_FIRST%} as FirstName search key.
	echo Selected {%$SEARCH_KEY_USER_LAST%} as LastName search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	echo Search Key FirstName: %$SEARCH_KEY_USER_FIRST% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Key LastName: %$SEARCH_KEY_USER_LAST% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo.
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo	jumpSUFLS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUFLS
:jumpSUFLS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(givenName=%$SEARCH_KEY_USER_FIRST%)(sn=%$SEARCH_KEY_USER_LAST%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSUFLS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUFL
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Name					User DN >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUFLO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn userPrincipalName displayName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
)

:jumpSUFLO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUFL
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUFL
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search User display name :::::::::::::::::::::::::::::::::::::::::::::::::
:SUDN
	SET $SEARCH_ATTRIBUTE=DisplayName
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: DisplayName" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	@powershell Write-Host "DisplayName by default is: Last`, First" -ForegroundColor Gray
	SET /P $SEARCH_KEY_USER_DIPLAYNAME=User DisplayName search key:
	IF NOT DEFINED $SEARCH_KEY_USER_DIPLAYNAME GoTo sUser
	SET $SEARCH_KEY=%$SEARCH_KEY_USER_DIPLAYNAME%
	call :SM
	echo Selected {%$SEARCH_KEY_USER_DIPLAYNAME%} as DisplayName search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUDNS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr displayName name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUDNS

:jumpSUDNS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr displayName name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(displayName=%$SEARCH_KEY_USER_DIPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSUDNS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUDN
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DisplayName			Name				User DN >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUDNO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr displayName userPrincipalName name cn %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSUDNO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUDN
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUDN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search User Custom Attributes ::::::::::::::::::::::::::::::::::::::::::::
:SUCA
	SET $SEARCH_ATTRIBUTE=
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	SET $FILTER=^(objectClass=%$SEARCH_TYPE%^)
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	SET "$ATTRIBUTES_USER=cn department description directReports displayName email extensionAttribute# givenName l mail mailNickname manager memberOf name physicalDeliveryOfficeName postalCode userPrincipalName sAMAccountName sn st title telephoneNumber"
	@powershell Write-Host "%$ATTRIBUTES_USER%" -ForegroundColor Blue
	echo ----------------------------------------
	@powershell Write-Host "Choose one user attribute to search against:" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	SET /P $SEARCH_ATTRIBUTE=User attribute:
	IF NOT DEFINED $SEARCH_ATTRIBUTE GoTo sUser
	echo %$ATTRIBUTES_USER% | FIND /I "%$SEARCH_ATTRIBUTE%"
	SET $USER_ATT_VALID=%ERRORLEVEL%
	IF %$USER_ATT_VALID% NEQ 0 (ECHO NOT VALID!) & (timeout /t 10) & (GoTo SUCA)
	:: Choose custom attribute search key
	call :SM
	@powershell Write-Host "Custom User Attribute: %$SEARCH_ATTRIBUTE%" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Choose a search key:" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	SET /P $SEARCH_KEY=Attribute %$SEARCH_ATTRIBUTE% search key:
	IF NOT DEFINED $SEARCH_KEY GoTo sUser
	call :SM
	:: Console Display
	@powershell Write-Host "Attribute: %$SEARCH_ATTRIBUTE%" -ForegroundColor Blue
	@powershell Write-Host "Search key: %$SEARCH_KEY%" -ForegroundColor Gray
	echo.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Make filter
	SET $FILTER=%$FILTER%^(%$SEARCH_ATTRIBUTE%=%$SEARCH_KEY%^)
	echo LDAP Filter: %$FILTER% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo.
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUCAS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr %$SEARCH_ATTRIBUTE% name displayName distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUCAS
:jumpSUCAS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr %$SEARCH_ATTRIBUTE% name displayName distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&%$FILTER%)" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSUCAS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUCA
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_ATTRIBUTE%		Name	DisplayName			distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUCAO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn displayName userPrincipalName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)
:jumpSUCAO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUCA
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUCA
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search User Global :::::::::::::::::::::::::::::::::::::::::::::::::::::::
:SUG
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

:::: Search User Global Inactive ::::::::::::::::::::::::::::::::::::::::::::::
:SUGI
	SET $SEARCH_ATTRIBUTE=Inactive
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: Inactive" -ForegroundColor Blue
	@powershell Write-Host "Choose n number of weeks" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to 0" -ForegroundColor Magenta
	SET /P $SEARCH_INACTIVE_KEY=%$LAST_SEARCH_ATTRIBUTE% n weeks:
	IF NOT DEFINED $SEARCH_INACTIVE_KEY=0
	SET $SEARCH_KEY=%$SEARCH_INACTIVE_KEY%
	call :SM
	@powershell Write-Host "Attribute: Name" -ForegroundColor Blue
	@powershell Write-Host "Can use wildcard *" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard" -ForegroundColor Magenta
	SET $SEARCH_NAME_KEY=*
	SET /P $SEARCH_NAME_KEY=name search key:
	call :SM
	echo Selected {%$SEARCH_INACTIVE_KEY%} as inactive search key.
	echo Selected {%$SEARCH_NAME_KEY%} as name search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUGIS
	:: Unsorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUGIS

:jumpSUGIS
	:: Sorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -inactive %$SEARCH_INACTIVE_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSUGIS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUGI
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUGIO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn displayName userPrincipalName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSUGIO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUGI
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUGI
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search User Global StalePassword :::::::::::::::::::::::::::::::::::::::::
:SUGS
	SET $SEARCH_ATTRIBUTE=StalePassword
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: StalePassword" -ForegroundColor Blue
	@powershell Write-Host "Choose n number of days" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to 0" -ForegroundColor Magenta
	SET $SEARCH_STALEPWD_KEY=0
	SET /P $SEARCH_STALEPWD_KEY=%$LAST_SEARCH_ATTRIBUTE% n days:
	SET $SEARCH_KEY=%$SEARCH_STALEPWD_KEY%
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
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUGSS
	:: Unsorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUGSS

:jumpSUGSS
	:: Sorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -stalepwd %$SEARCH_STALEPWD_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"


:skipSUGSS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUGS
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUGSO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn displayName userPrincipalName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSUGSO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUGS
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUGS
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search User Global Disabled ::::::::::::::::::::::::::::::::::::::::::::::
:SUGD
	SET $SEARCH_ATTRIBUTE=Disabled
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: Name" -ForegroundColor Blue
	@powershell Write-Host "Can use wildcard *" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to *" -ForegroundColor Magenta
	SET $SEARCH_NAME_KEY=*
	SET /P $SEARCH_NAME_KEY=name search key:
	SET $SEARCH_KEY=%$SEARCH_NAME_KEY%
	call :SM
	echo Selected {%$SEARCH_NAME_KEY%} as name search key.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSUGDS
	:: Unsorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSUGDS
:jumpSUGDS
	:: Sorted
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY USER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_NAME_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSUGDS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSUGD
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Names returned: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User Distinguisged Names: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSUGDO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr name cn displayName userPrincipalName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr sn givenName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo User DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Organization Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr company %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr department %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr title %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr manager %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Contact Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr mail %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr telephoneNumber %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Location Information: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr physicalDeliveryOfficeName streetAddress l st postalCode co %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=user)(distinguishedName=%%~N))" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSUGDO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSUGD
	echo Search User Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sUser
	IF %ERRORLEVEL% EQU 1 GoTo SUGD
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Group :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:sGroup
	SET $SEARCH_TYPE=Group
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	call :SM
	::	Close previous Windows
	call :LSLCN
	Echo Group search using:
	Echo.
	Echo [1] Name
	Echo [2] Description
	Echo [3] DisplayName
	echo [4] Multi-Attribute
	echo [5] Abort
	Echo.
	Choice /c 12345
	If ERRORLevel 5 GoTo Search
	If ERRORLevel 4 GoTo sGM
	If ERRORLevel 3 GoTo sGDN
	If ERRORLevel 2 GoTo sGD
	If ERRORLevel 1 GoTo sGN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Group Search: Name attribute ::::::::::::::::::::::::::::::::::::::::
:sGN
	SET $SEARCH_ATTRIBUTE=name
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	SET /P $SEARCH_KEY=Group name search key:
	IF NOT DEFINED $SEARCH_KEY GoTo skipSGN
	call :SM
	@powershell Write-Host "Selected %$SEARCH_KEY% as search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSGNS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSGS
:jumpSGNS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSGS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSGN
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Name					DN >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSGNO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name description displayName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		DSGET GROUP %%N -members %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% 2> nul | DSGET USER -upn -samid -fn -mi -ln -display -email %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSGNO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSGN
	echo Search Group Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sGroup
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Group Search: Description attribute ::::::::::::::::::::::::::::::::::::::
:sGD
	SET $SEARCH_ATTRIBUTE=description
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	SET /P $SEARCH_KEY=Group %$SEARCH_ATTRIBUT% search key:
	IF NOT DEFINED $SEARCH_KEY GoTo skipSGD
	call :SM
	@powershell Write-Host "Selected %$SEARCH_KEY% as search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo	jumpSGDS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSGDS
:jumpSGDS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(description=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSGDS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSGD
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Name					DN >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSGDO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name description displayName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET GROUP %%N -members %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% 2> nul | DSGET USER -upn -samid -fn -mi -ln -display -email %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSGDO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSGD
	echo Search Group Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sGroup
	IF %ERRORLEVEL% EQU 1 GoTo sGD
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Group Search: DisplayName Attribute attribute ::::::::::::::::::::::::::::
:sGDN
	SET $SEARCH_ATTRIBUTE=displayName
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: %$SEARCH_ATTRIBUTE%" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	SET /P $SEARCH_KEY=Group %$SEARCH_ATTRIBUT% search key:
	IF NOT DEFINED $SEARCH_KEY GoTo skipSGDN
	call :SM
	@powershell Write-Host "Selected %$SEARCH_KEY% as search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo	jumpSGDNS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(displayName=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(displayName=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSGDNS
:jumpSGDNS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(displayName=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(displayName=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSGDNS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSGDN
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group Name					DN >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSGDNO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name description displayName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET GROUP %%N -members %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% 2> nul | DSGET USER -upn -samid -fn -mi -ln -display -email %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSGDNO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
:skipSGDN
	echo Search Group Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sGroup
	IF %ERRORLEVEL% EQU 1 GoTo sGDN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Group Search: Multi attribute ::::::::::::::::::::::::::::::::::::::::::::
:sGM
::	mode con:cols=80
::	mode con:lines=50
	SET $SEARCH_ATTRIBUTE=name-description-displayName
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attributes: name description displayName" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, defaults to * wildcard" -ForegroundColor Magenta
	SET $SEARCH_KEY_GROUP_NAME=*
	SET /P $SEARCH_KEY_GROUP_NAME=Group name search key:
	call :SM
	@powershell Write-Host "Attributes: name description displayName" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Attribute: description" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, defaults to * wildcard" -ForegroundColor Magenta
	SET $SEARCH_KEY_GROUP_DESCRIPTION=*
	SET /P $SEARCH_KEY_GROUP_DESCRIPTION=Group description search key:
	call :SM
	@powershell Write-Host "Attributes: name description displayName" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Attribute: displayName" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, defaults to * wildcard" -ForegroundColor Magenta
	SET $SEARCH_KEY_GROUP_DISPLAYNAME=*
	SET /P $SEARCH_KEY_GROUP_DESCRIPTION=Group displayName search key:
	call :SM
	SET $SEARCH_KEY=%$SEARCH_KEY_GROUP_NAME%-%$SEARCH_KEY_GROUP_DESCRIPTION%-%$SEARCH_KEY_GROUP_DISPLAYNAME%
	@powershell Write-Host "Selected %$SEARCH_KEY_GROUP_NAME% as name search key." -ForegroundColor Gray
	@powershell Write-Host "Selected %$SEARCH_KEY_GROUP_DESCRIPTION% as description search key." -ForegroundColor Gray
	@powershell Write-Host "Selected %$SEARCH_KEY_GROUP_DISPLAYNAME% as displayName search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSGMS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY_GROUP_NAME%)(description=%$SEARCH_KEY_GROUP_DESCRIPTION%)(displayName=%$SEARCH_KEY_GROUP_DISPLAYNAME%))" -attr name description displayName distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY_GROUP_NAME%)(description=%$SEARCH_KEY_GROUP_DESCRIPTION%)(displayName=%$SEARCH_KEY_GROUP_DISPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSGMS

:jumpSGMS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY_GROUP_NAME%)(description=%$SEARCH_KEY_GROUP_DESCRIPTION%)(displayName=%$SEARCH_KEY_GROUP_DISPLAYNAME%))" -attr name description displayName distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=group)(name=%$SEARCH_KEY_GROUP_NAME%)(description=%$SEARCH_KEY_GROUP_DESCRIPTION%)(displayName=%$SEARCH_KEY_GROUP_DISPLAYNAME%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSGMS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSGM
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Group: Name				description				displayName						DN >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSGMO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name description displayName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% Members: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET GROUP %%N -members %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% 2> nul | DSGET USER -upn -samid -fn -mi -ln -display -email %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSGMO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSGM
	echo Search Group Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sGroup
	IF %ERRORLEVEL% EQU 1 GoTo sGM
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Computer Search ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:sComputer
	SET $SEARCH_TYPE=Computer
	SET $SEARCH_ATTRIBUTE=
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	echo Computer search using:
	echo.
	Echo [1] Name
	Echo [2] Advanced
	Echo [3] Back to Search Menu
	Echo.
	Choice /c 123
	Echo.
	If ERRORLevel 3 GoTo Search
	If ERRORLevel 2 GoTo sCA
	If ERRORLevel 1 GoTo sCN
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Computer Search: name ::::::::::::::::::::::::::::::::::::::::::::::::::::
:sCN
	SET $SEARCH_ATTRIBUTE=name
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use "*" wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	SET /P $SEARCH_KEY=%$SEARCH_TYPE% %$SEARCH_ATTRIBUTE% search key:
	IF NOT DEFINED $SEARCH_KEY GoTo	skipSCN
	call :SM
	@powershell Write-Host "Selected %$SEARCH_KEY% as search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCNS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSCNS
:jumpSCNS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSCNS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCN
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE%: %$SEARCH_ATTRIBUTE%					DN >> "%$LogPath%\%$LAST_SEARCH_LOG%"
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
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCNO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name cn %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCNO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCN
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sComputer
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Computer Advanced :::::::::::::::::::::::::::::::::::::::::::::::::
:sCA
	SET $SEARCH_TYPE=Computer
	SET $SEARCH_ATTRIBUTE=
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	Color 0A
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM

	::Selection
	echo Computer Advanced search using:
	echo [1] Disabled
	echo [2] Inactive
	echo [3] StalePWD
	echo [4] Operating System ^(attributes^)
	echo [5] Time Series ^(attributes^)
	echo [6] LogonCount
	echo [7] Multiple Attribute search
	echo [8] Back to Computer search
	Echo.
	Choice /c 12345678
	Echo.
	If ERRORLevel 8 GoTo sComputer
	If ERRORLevel 7 GoTo sCMA
	If ERRORLevel 6 GoTo sCLC
	If ERRORLevel 5 GoTo sCTS
	If ERRORLevel 4 GoTo sCOS
	If ERRORLevel 3 GoTo sCS
	If ERRORLevel 2 GoTo sCI
	If ERRORLevel 1 GoTo sCD
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search computer disabled :::::::::::::::::::::::::::::::::::::::::::::::::
:sCD
	SET $SEARCH_ATTRIBUTE=disabled name
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: name disabled" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	SET /P $SEARCH_KEY=%$SEARCH_TYPE% %$SEARCH_ATTRIBUTE% search key:
	IF NOT DEFINED $SEARCH_KEY GoTo skipSCD
	call :SM
	@powershell Write-Host "Selected %$SEARCH_KEY% as %$SEARCH_ATTRIBUTE% search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCDS
	:: Unsorted
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSCDS
:jumpSCDS
	:: Sorted
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY%" -disabled -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSCDS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCD
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Name			distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN.txt" DEL /F /Q "%$LogPath%\var\var_Last_Search_N_DN.txt"
	FOR /F " USEBACKQ tokens=* delims=" %%R IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~R)" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	FINDSTR /V /C:"name" "%$LogPath%\var\var_Last_Search_N_DN.txt" > "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCDO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCDO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCD
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo sCA
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Computers inactive search	:::::::::::::::::::::::::::::::::::::::::::::::
:sCI
	SET $SEARCH_ATTRIBUTE=inactive
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: inactive name" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Attribute: %$SEARCH_ATTRIBUTE%" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	set /P $SEARCH_KEY=%$SEARCH_TYPE% %$SEARCH_ATTRIBUTE% number of weeks:
	IF NOT DEFINED $SEARCH_KEY GoTo skipSCI
	call :SM
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, defaults to * wildcard" -ForegroundColor Magenta
	set $SEARCH_KEY_PC_NAME=*
	set /P $SEARCH_KEY_PC_NAME=Choose name search key:
	call :SM
	@powershell Write-Host "Search for all inactive [%$SEARCH_KEY% weeks] computers with name key: %$SEARCH_KEY_PC_NAME%" -ForegroundColor Gray
	@powershell Write-Host "Selected %$SEARCH_KEY_PC_NAME% as name search key" -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCIS
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSCIS
:jumpSCIS
	:: Sorted
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -inactive %$SEARCH_KEY% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSCIS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCI
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Name			distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN.txt" DEL /F /Q "%$LogPath%\var\var_Last_Search_N_DN.txt"
	FOR /F " USEBACKQ tokens=* delims=" %%R IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~R)" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	FINDSTR /V /C:"name" "%$LogPath%\var\var_Last_Search_N_DN.txt" > "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCIO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCIO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCI
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo SCI
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Computers with stale passwords search ::::::::::::::::::::::::::::::::::::
:sCS
	SET $SEARCH_ATTRIBUTE=StalePassword
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: StalePassword name" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Attribute: %$SEARCH_ATTRIBUTE%" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort" -ForegroundColor Magenta
	@powershell Write-Host "Stale password for n days" -ForegroundColor Gray
	SET /P $SEARCH_STALEPWD=Stale password number of days:
	IF NOT DEFINED $SEARCH_STALEPWD GoTo skipSCS
	call :SM
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	SET $SEARCH_KEY_PC_NAME=*
	SET /P $SEARCH_KEY_PC_NAME=Choose name search key:
	call :SM
	SET $SEARCH_KEY=%$SEARCH_KEY_PC_NAME% %$SEARCH_STALEPWD%
	@powershell Write-Host "Search for all %$SEARCH_KEY_PC_NAME% computers with stalepassword %$SEARCH_STALEPWD% days..." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCSS
	:: Unsorted
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_STALEPWD% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_STALEPWD% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSCSS
:jumpSCSS
	:: Sorted
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o rdn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_STALEPWD% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N.txt"
	DSQUERY COMPUTER %$AD_BASE% -scope %$AD_SCOPE% -o dn -name "%$SEARCH_KEY_PC_NAME%" -stalepwd %$SEARCH_STALEPWD% -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSCSS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCS
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Computer Name			distinguishedName >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN.txt" DEL /F /Q "%$LogPath%\var\var_Last_Search_N_DN.txt"
	FOR /F " USEBACKQ tokens=* delims=" %%R IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
		DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~R)" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\var\var_Last_Search_N_DN.txt"
		)
	FINDSTR /V /C:"name" "%$LogPath%\var\var_Last_Search_N_DN.txt" > "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCSO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCSO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCS
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo sCS
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search computer Operating system andOr version :::::::::::::::::::::::::::
:SCOS
	SET $SEARCH_ATTRIBUTE=name OperatingSystem operatingSystemVersion operatingSystemServicePack
	SET $SEARCH_OS=
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: OperatingSystem operatingSystemVersion operatingSystemServicePack name" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Attribute: operatingSystem" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort!" -ForegroundColor Magenta
	set /P $SEARCH_OS=Operating System:
	IF NOT DEFINED $SEARCH_OS GoTo skipSCOS
	call :SM
	@powershell Write-Host "Attribute: operatingSystemVersion" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	set $SEARCH_OSV=*
	set /P $SEARCH_OSV=Operating System Version:
	call :SM
	@powershell Write-Host "Attribute: operatingSystemServicePack" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	set $SEARCH_OSSP=*
	set /P $SEARCH_OSSP=Operating System Service Pack:
	call :SM
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	set $SEARCH_KEY_PC_NAME=*
	set /P $SEARCH_KEY_PC_NAME=Choose computer name search key:
	SET "$SEARCH_KEY=%$SEARCH_KEY_PC_NAME% %$SEARCH_OS% %$SEARCH_OSV% %$SEARCH_OSSP%"
	call :SM
	@powershell Write-Host "Computer name search key: %$SEARCH_KEY_PC_NAME%" -ForegroundColor Gray
	@powershell Write-Host "Operating System search key: %$SEARCH_OS%" -ForegroundColor Gray
	@powershell Write-Host "operatingSystemVersion search key: %$SEARCH_OSV%" -ForegroundColor Gray
	@powershell Write-Host "operatingSystemServicePack search key: %$SEARCH_OSSP%" -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	echo Search Computer Name: %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Operating System: %$SEARCH_OS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Operating System Version: %$SEARCH_OSV% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Search Operating System Service pack: %$SEARCH_OSSP% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCOSS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSCOSS
:jumpSCOSS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(operatingSystem=%$SEARCH_OS%)(operatingSystemVersion=%$SEARCH_OSV%)(operatingSystemServicePack=%$SEARCH_OSSP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSCOSS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCOS
	)
	:: Main output
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCOSO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCOSO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCOS
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo SCOS
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Computer Time Series attributes :::::::::::::::::::::::::::::::::::
:SCTS
	SET $SEARCH_ATTRIBUTE=TimeSeries
	SET $SEARCH_KEY_PC_NAME=
	SET $SEARCH_WHENCREATED=
	SET $SEARCH_WHENCHANGED=
	SET $SEARCH_LASTLOGONTIMESTAMP=
	SET $LAST_SEARCH_COUNT=
	SET $SEARCH_KEY=multi-key
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attributes: name lastLogonTimestamp whenCreated whenChanged" -ForegroundColor Blue
	echo.
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	SET $SEARCH_KEY_PC_NAME= *
	SET /P $SEARCH_KEY_PC_NAME=name search key:
	call :SM
	call :subOperator "LASTLOGONTIMESTAMP"
	call :SM
	@powershell Write-Host "Attribute: lastLogonTimestamp" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	@powershell Write-Host "[NT Time] e.g. 132530551699076595" -ForegroundColor Cyan
	SET $SEARCH_LASTLOGONTIMESTAMP=*
	SET /P $SEARCH_LASTLOGONTIMESTAMP=lastLogonTimestamp search key:
	call :SM
	call :subOperator "WHENCREATED"
	call :SM
	@powershell Write-Host "Attribute: whenCreated" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	@powershell Write-Host "[YYYY MM DD HH mm ss.s Z] e.g. 20200101120000.0Z" -ForegroundColor Cyan
	set $SEARCH_WHENCREATED=*
	SET /P $SEARCH_WHENCREATED=whenCreated search key:
	call :SM
	call :subOperator "WHENCHANGED"
	call :SM
	@powershell Write-Host "whenChanged" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	@powershell Write-Host "[YYYY MM DD HH mm ss.s Z] e.g. 20200101120000.0Z" -ForegroundColor Cyan
	set $SEARCH_WHENCHANGED=*
	SET /P $SEARCH_WHENCHANGED=whenChanged search key:
	call :SM
	:: Console Display
		:: Convert NT time
	IF "%$SEARCH_LASTLOGONTIMESTAMP%"=="*" SET $NT_TIME_CONVERTED=*
	IF "%$NT_TIME_CONVERTED%"=="*" GoTo skipLLTSC
	FOR /F "tokens=2 delims=-" %%P IN ('w32tm.exe /ntte %$SEARCH_LASTLOGONTIMESTAMP%') DO echo %%P> "%$LogPath%\var\var_$NT_TIME_CONVERTED.txt"
	SET /P $NT_TIME_CONVERTED= < "%$LogPath%\var\var_$NT_TIME_CONVERTED.txt"
:skipLLTSC
	@powershell Write-Host "Search parameters:" -ForegroundColor Gray
	echo Attribute		Operator  Search Key
	echo  name			[^=]	%$SEARCH_KEY_PC_NAME%
	echo  whenCreated		[%$SEARCH_WHENCREATED_OPERATOR_DISPLAY%]	%$SEARCH_WHENCREATED%
	echo  whenChanged		[%$SEARCH_WHENCHANGED_OPERATOR_DISPLAY%]	%$SEARCH_WHENCHANGED%
	echo  lastLogonTimestamp	[%$SEARCH_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%]	%$NT_TIME_CONVERTED%
	echo.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	echo ---- Search parameters ---- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute		Operator  Search Key >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo name:		[^=] %$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo whenCreated:			[%$ATTR_WHENCREATED_OPERATOR_DISPLAY%] %$SEARCH_WHENCREATED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo whenChanged:			[%$ATTR_WHENCHANGED_OPERATOR_DISPLAY%] %$SEARCH_WHENCHANGED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo lastLogonTimestamp:	[%$ATTR_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%] %$SEARCH_LASTLOGONTIMESTAMP% %$NT_TIME_CONVERTED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCTSS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$ATTR_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$ATTR_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$ATTR_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"

	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$ATTR_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$ATTR_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$ATTR_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
GoTo skipSCTSS
:jumpSCTSS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$ATTR_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$ATTR_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$ATTR_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"

	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(whenCreated%$ATTR_WHENCREATED_OPERATOR%%$SEARCH_WHENCREATED%)(whenChanged%$ATTR_WHENCHANGED_OPERATOR%%$SEARCH_WHENCHANGED%)(lastLogonTimestamp%$ATTR_LASTLOGONTIMESTAMP_OPERATOR%%$SEARCH_LASTLOGONTIMESTAMP%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSCTSS

	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCTS
	)
	:: Main output
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCTSO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCTSO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCTS
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo SCTS
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Computer LogonCount :::::::::::::::::::::::::::::::::::::::::::::::
:sCLC
	SET $SEARCH_ATTRIBUTE=logonCount
	SET $SEARCH_KEY=
	SET $SEARCH_KEY_PC_NAME=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attributes: name LogonCount" -ForegroundColor Blue
	:: Computer Name search key
	echo.
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	SET $SEARCH_KEY_PC_NAME=*
	SET /P $SEARCH_KEY_PC_NAME=name search key:
	call :SM
	:: LogonCount search operator
	call :subOperator "LOGONCOUNT"
	call :SM
	@powershell Write-Host "Attribute: logonCount:" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard!" -ForegroundColor Magenta
	SET $SEARCH_LOGONCOUNT=*
	SET /P $SEARCH_LOGONCOUNT=LogonCount search key:
	set $SEARCH_KEY=%$SEARCH_KEY_PC_NAME% %$SEARCH_LOGONCOUNT%
	call :SM
	:: Console Display
	@powershell Write-Host "---- Search parameters ----" -ForegroundColor Gray
	echo Attribute		Operator  Search Key
	echo  name			[^=]	%$SEARCH_KEY_PC_NAME%
	echo logonCount		[%$SEARCH_LOGONCOUNT_OPERATOR_DISPLAY%]	%$SEARCH_LOGONCOUNT%
	echo.
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	echo ---- Search parameters ---- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Attribute		Operator  Search Key >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Name			[^=]		%$SEARCH_KEY_PC_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo logonCount		[%$ATTR_LOGONCOUNT_OPERATOR%]		%$SEARCH_LOGONCOUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Search
		:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCLCS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$ATTR_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"

	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$ATTR_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"

	GoTo skipSCLCS

 :jumpSCLCS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$ATTR_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"

	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "(&(objectClass=computer)(name=%$SEARCH_KEY_PC_NAME%)(logonCount%$ATTR_LOGONCOUNT_OPERATOR%%$SEARCH_LOGONCOUNT%))" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSCLCS

	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCLC
	)
	:: Main output
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCLCO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCLCO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCLC
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo SCLC
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Computer Multiple Attributes ::::::::::::::::::::::::::::::::::::::
:SCMA
	SET $SEARCH_TYPE=Computer
	SET $SEARCH_ATTRIBUTE=Multiple
	SET $SEARCH_KEY=Multi-key
	SET $LAST_SEARCH_COUNT=
	SET "$ATTRIBUTES_COMPUTER_LIST=name cn description displayName distinguishedName whenCreated whenChanged logonCount lastLogonTimestamp objectSid dNSHostName operatingSystem operatingSystemVersion operatingSystemServicePack managedBy"
	set $ATTRIBUTES_COMPUTER=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Multiple-Attributes: %$ATTRIBUTES_COMPUTER_LIST% " -ForegroundColor Blue
	echo ----------------------------------------
	@powershell Write-Host "Type attributes seperated by spaces" -ForegroundColor Gray
	@powershell Write-Host "If left blank, will abort!" -ForegroundColor Magenta
	set /P $ATTRIBUTES_COMPUTER=Multi-Attribute List:
	if not defined $ATTRIBUTES_COMPUTER GoTo skipSCMA
	SET $COUNTER=0
	SET $COUNTER_MAX=
	:: Sub-Routine Count Multi Attributes
:subCMA
	set /A $COUNTER+=1
	for /F "tokens=%$COUNTER% delims= " %%P IN ("%$ATTRIBUTES_COMPUTER%") Do (
	if "%%P"=="""" GoTo esubCMA
	set /A $COUNTER_MAX+=1
	GoTo subCMA
	)
:esubCMA

	:: Process Multi Attributes
	set $COUNTER=1
	set /A $COUNTER_MAX+=1
:subSCMA
	FOR /F "tokens=%$COUNTER%" %%P IN ("%$ATTRIBUTES_COMPUTER%") DO (
		call :SM
		IF %$COUNTER% EQU %$COUNTER_MAX% GoTo eSubSCMA
		@powershell Write-Host "Attribute: %%P" -ForegroundColor Blue
		call :subOperator %%P
		@powershell Write-Host "Leave blank for wildcard *" -ForegroundColor Magenta
		SET $ATTR_%%P=*
		SET /P $ATTR_%%P=Computer %%P search key:
		SET /a $COUNTER+=1
		GoTo subSCMA
	)
:eSubSCMA

	call :SM
	:: Console Display
	@powershell Write-Host "---- Search parameters ----" -ForegroundColor Gray
	@powershell Write-Host "Attribute `| Operator `| Search Key" -ForegroundColor Cyan

::#############################################################################
:: Parameters within a variable don't work (?)
::#############################################################################
::	set $COUNTER=1
:::subSCMAD
	::	Display multi attribute settings
::	setlocal enabledelayedexpansion
::	for /F "tokens=%$COUNTER% delims= " %%P IN ("%$ATTRIBUTES_COMPUTER%") Do (
::		IF %$COUNTER% EQU %$COUNTER_MAX% GoTo esubSCMAD
::		echo %%P	[%$ATTR_%%P_OPERATOR_DISPLAY%]	%$ATTR_%%P%
::		SET /a $COUNTER+=1
::		GoTo subSCMAD
::	)
::	echo.
::	setlocal disabledelayedexpansion
:esubSCMAD
::#############################################################################

:: Workaround
	IF DEFINED $ATTR_NAME echo name	[%$ATTR_NAME_OPERATOR_DISPLAY%] %$ATTR_NAME%
	IF DEFINED $ATTR_CN echo cn	[%$ATTR_CN_OPERATOR_DISPLAY%] %$ATTR_CN%
	IF DEFINED $ATTR_DESCRIPTION echo description [%$ATTR_DESCRIPTION_OPERATOR_DISPLAY%] %$ATTR_DESCRIPTION%
	IF DEFINED $ATTR_DISPLAYNAME echo displayName [%$ATTR_DISPLAYNAME_OPERATOR_DISPLAY%] %$ATTR_DISPLAYNAME%
	IF DEFINED $ATTR_WHENCREATED echo whenCreated [%$ATTR_WHENCREATED_OPERATOR_DISPLAY%] %$ATTR_WHENCREATED%
	IF DEFINED $ATTR_WHENCHANGED echo whenChanged [%$ATTR_WHENCHANGED_OPERATOR_DISPLAY%] %$ATTR_WHENCHANGED%
	IF DEFINED $ATTR_LOGONCOUNT echo logonCount [%$ATTR_LOGONCOUNT_OPERATOR_DISPLAY%] %$ATTR_LOGONCOUNT%
	IF DEFINED $ATTR_LASTLOGONTIMESTAMP echo lastLogonTimestamp [%$ATTR_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%] %$ATTR_LASTLOGONTIMESTAMP%
	IF DEFINED $ATTR_OBJECTSID echo objectSid [%$ATTR_OBJECTSID_OPERATOR_DISPLAY%] %$ATTR_OBJECTSID%
	IF DEFINED $ATTR_DNSHOSTNAME echo dNSHostName [%$ATTR_DNSHOSTNAME_OPERATOR_DISPLAY%] %$ATTR_DNSHOSTNAME%
	IF DEFINED $ATTR_OPERATINGSYSTEM echo operatingSystem [%$ATTR_OPERATINGSYSTEM_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEM%
	IF DEFINED $ATTR_OPERATINGSYSTEMVERSION echo operatingSystemVersion [%$ATTR_OPERATINGSYSTEMVERSION_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEMVERSION%
	IF DEFINED $ATTR_OPERATINGSYSTEMSERVICEPACK echo operatingSystemServicePack [%$ATTR_OPERATINGSYSTEMSERVICEPACK_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEMSERVICEPACK%
	IF DEFINED $ATTR_MANAGEDBY echo managedBy [%$ATTR_MANAGEDBY_OPERATOR_DISPLAY%] %$ATTR_MANAGEDBY%
	echo.
	:: Create filter string
	SET "$FILTER=(objectClass=computer)"
	IF DEFINED $ATTR_NAME SET "$FILTER=%$FILTER%(name%$ATTR_NAME_OPERATOR%%$ATTR_NAME%)"
	IF DEFINED $ATTR_CN	set "$FILTER=%$FILTER%(cn%$ATTR_CN_OPERATOR%%$ATTR_CN%)"
	IF DEFINED $ATTR_DESCRIPTION set "$FILTER=%$FILTER%(description%$ATTR_DESCRIPTION_OPERATOR%%$ATTR_DESCRIPTION%)"
	IF DEFINED $ATTR_DISPLAYNAME set "$FILTER=%$FILTER%(displayName%$ATTR_DISPLAYNAME_OPERATOR%%$ATTR_DISPLAYNAME%)"
	IF DEFINED $ATTR_WHENCREATED set "$FILTER=%$FILTER%(whenCreated%$ATTR_WHENCREATED_OPERATOR%%$ATTR_WHENCREATED%)"
	IF DEFINED $ATTR_WHENCHANGED set "$FILTER=%$FILTER%(whenChanged%$ATTR_WHENCHANGED_OPERATOR%%$ATTR_WHENCHANGED%)"
	IF DEFINED $ATTR_LOGONCOUNT	set "$FILTER=%$FILTER%(logonCount%$ATTR_LOGONCOUNT_OPERATOR%%$ATTR_LOGONCOUNT%)"
	IF DEFINED $ATTR_LASTLOGONTIMESTAMP	set "$FILTER=%$FILTER%(lastLogonTimestamp%$ATTR_LASTLOGONTIMESTAMP_OPERATOR%%$ATTR_LASTLOGONTIMESTAMP%)"
	IF DEFINED $ATTR_OBJECTSID	set "$FILTER=%$FILTER%(objectSid%$ATTR_OBJECTSID_OPERATOR%%$ATTR_OBJECTSID%)"
	IF DEFINED $ATTR_DNSHOSTNAME set "$FILTER=%$FILTER%(dNSHostName%$ATTR_DNSHOSTNAME_OPERATOR%%$ATTR_DNSHOSTNAME%)"
	IF DEFINED $ATTR_OPERATINGSYSTEM set "$FILTER=%$FILTER%(operatingSystem%$ATTR_OPERATINGSYSTEM_OPERATOR%%$ATTR_OPERATINGSYSTEM%)"
	IF DEFINED $ATTR_OPERATINGSYSTEMVERSION	set "$FILTER=%$FILTER%(operatingSystemVersion%$ATTR_OPERATINGSYSTEMVERSION_OPERATOR%%$ATTR_OPERATINGSYSTEMVERSION%)"
	IF DEFINED $ATTR_OPERATINGSYSTEMSERVICEPACK	set "$FILTER=%$FILTER%(operatingSystemServicePack%$ATTR_OPERATINGSYSTEMSERVICEPACK_OPERATOR%%$ATTR_OPERATINGSYSTEMSERVICEPACK%)"
	IF DEFINED $ATTR_MANAGEDBY set "$FILTER=%$FILTER%(managedBy%$ATTR_MANAGEDBY_OPERATOR%%$ATTR_MANAGEDBY%)"
	SET "$FILTER=(&%$FILTER%)"
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	echo ---- Search parameters ---- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_NAME echo name	[%$ATTR_NAME_OPERATOR_DISPLAY%] %$ATTR_NAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_CN echo cn	[%$ATTR_CN_OPERATOR_DISPLAY%] %$ATTR_CN% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_DESCRIPTION echo description [%$ATTR_DESCRIPTION_OPERATOR_DISPLAY%] %$ATTR_DESCRIPTION% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_DISPLAYNAME echo displayName [%$ATTR_DISPLAYNAME_OPERATOR_DISPLAY%] %$ATTR_DISPLAYNAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_WHENCREATED echo whenCreated [%$ATTR_WHENCREATED_OPERATOR_DISPLAY%] %$ATTR_WHENCREATED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_WHENCHANGED echo whenChanged [%$ATTR_WHENCHANGED_OPERATOR_DISPLAY%] %$ATTR_WHENCHANGED% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_LOGONCOUNT echo logonCount [%$ATTR_LOGONCOUNT_OPERATOR_DISPLAY%] %$ATTR_LOGONCOUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_LASTLOGONTIMESTAMP echo lastLogonTimestamp [%$ATTR_LASTLOGONTIMESTAMP_OPERATOR_DISPLAY%] %$ATTR_LASTLOGONTIMESTAMP% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_OBJECTSID echo objectSid [%$ATTR_OBJECTSID_OPERATOR_DISPLAY%] %$ATTR_OBJECTSID% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_DNSHOSTNAME echo dNSHostName [%$ATTR_DNSHOSTNAME_OPERATOR_DISPLAY%] %$ATTR_DNSHOSTNAME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_OPERATINGSYSTEM echo operatingSystem [%$ATTR_OPERATINGSYSTEM_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEM% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_OPERATINGSYSTEMVERSION echo operatingSystemVersion [%$ATTR_OPERATINGSYSTEMVERSION_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEMVERSION% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_OPERATINGSYSTEMSERVICEPACK echo operatingSystemServicePack [%$ATTR_OPERATINGSYSTEMSERVICEPACK_OPERATOR_DISPLAY%] %$ATTR_OPERATINGSYSTEMSERVICEPACK% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF DEFINED $ATTR_MANAGEDBY echo managedBy [%$ATTR_MANAGEDBY_OPERATOR_DISPLAY%] %$ATTR_MANAGEDBY% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Query Filter string: "%$FILTER%" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Search
		:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSCMAS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
GoTo skipSCMAS
:jumpSCMAS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"
:skipSCMAS

	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSCMA
	)
	:: Main output
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSCMAO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo LastLogonTimestamp: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr lastLogonTimestamp %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_$lastLogonTimestamp.txt"
	FOR /F "skip=1 delims=" %%P IN (%$LogPath%\var\var_$lastLogonTimestamp.txt) DO w32tm.exe /ntte %%P >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -disabled -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	dsget computer %%N -loc -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo MemberOf: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET computer %%N -memberof -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSCMAO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipSCMA
	echo Search Computer Again?
	Choice /c YN /m "[Y]es or [N]o":
	IF %ERRORLEVEL% EQU 2 GoTo sComputer
	IF %ERRORLEVEL% EQU 1 GoTo SCMA
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Server ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:sServer
	
	SET $SEARCH_TYPE=Server
	SET $SEARCH_ATTRIBUTE=name
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard" -ForegroundColor Magenta
	set $SEARCH_KEY=*
	SET /P $SEARCH_KEY=%$SEARCH_TYPE% %$SEARCH_ATTRIBUTE% search key:
	call :SM
	@powershell Write-Host "Selected [%$SEARCH_KEY%] as search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Check on Wildcard *
	IF "%$SEARCH_KEY%"=="*" (SET $SERVER_SEARCH_GLOBAL=1) ELSE (SET $SERVER_SEARCH_GLOBAL=0)
	echo %$SERVER_SEARCH_GLOBAL% > "%$LogPath%\var\var_$SERVER_SEARCH_GLOBAL.txt"
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	IF %$SERVER_SEARCH_GLOBAL% EQU 1 SET $AD_BASE=forestroot	
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Search Servers
		:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSSS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=server)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=server)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
GoTo skipSSS
:jumpSSS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=server)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=server)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSSS

	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipsServer
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSSO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name displayName description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSGET server "%%N" -site -s %$DC%.%$DOMAIN% %$DOMAIN_CREDENTIALS% 2> nul >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSSO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"

:skipsServer
	echo Search Again?
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo Search
	IF %ERRORLEVEL% EQU 1 GoTo sServer
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search OU ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:sOU
	SET "$SEARCH_TYPE=OrganizationalUnit^(OU^)"
	SET $SEARCH_ATTRIBUTE=name
	SET $SEARCH_KEY=
	SET $LAST_SEARCH_COUNT=
	:: Last Search Log Close - Notepad
	call :LSLCN
	call :SM
	:: Console output
	@powershell Write-Host "Attribute: name" -ForegroundColor Blue
	@powershell Write-Host "Can use * wildcard for search." -ForegroundColor Gray
	@powershell Write-Host "If left blank, will default to * wildcard" -ForegroundColor Magenta
	@powershell Write-Host "Tpye abort to exit" -ForegroundColor Magenta
	set $SEARCH_KEY=*
	SET /P $SEARCH_KEY=%$SEARCH_TYPE% %$SEARCH_ATTRIBUTE% search key:
	IF /I %$SEARCH_KEY%==abort GoTo Search
	call :SM
	@powershell Write-Host "Selected [%$SEARCH_KEY%] as search key." -ForegroundColor Gray
	@powershell Write-Host "Searching..." -ForegroundColor DarkYellow
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Check on Wildcard *
	IF "%$SEARCH_KEY%"=="*" (SET $SERVER_SEARCH_GLOBAL=1) ELSE (SET $SERVER_SEARCH_GLOBAL=0)
	echo %$SERVER_SEARCH_GLOBAL% > "%$LogPath%\var\var_$SERVER_SEARCH_GLOBAL.txt"
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	IF %$SERVER_SEARCH_GLOBAL% EQU 1 SET $AD_BASE=forestroot	
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Search Servers
		:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSOUS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=organizationalUnit)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=organizationalUnit)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"
	GoTo skipSOUS
:jumpSOUS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=organizationalUnit)(name=%$SEARCH_KEY%))" -attr name distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(&(objectClass=organizationalUnit)(name=%$SEARCH_KEY%))" -attr distinguishedName %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"

:skipSOUS

	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	call :SM
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		@powershell Write-Host "Nothing found! Try again with broader wildcard" -ForegroundColor Red
		GoTo skipSOU
	)
	:: Main output
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT%
	@powershell Write-Host "Processing..." -ForegroundColor DarkYellow
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSOUO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	call :SM
	@powershell Write-Host "Processing verbose..." -ForegroundColor DarkYellow
	@powershell Write-Host '%%N' -ForegroundColor DarkGray
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name description %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"	
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)
	
:jumpSOUO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	:: Refresh HUD
	call :SM
	echo Total Search Time: %$TOTAL_LAPSE_TIME%
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"
	:: Open log files
	@explorer "%$LogPath%\%$LAST_SEARCH_LOG%"
	
:skipSOU
	echo What to do next?
	echo	[1] Set OU as AD Base search?
	echo	[2] Search OU again?
	echo	[3] Go back to search menu?
	echo.
	Choice /c 123 /m "Select:"
	IF %ERRORLEVEL% EQU 3 GoTo Search
	IF %ERRORLEVEL% EQU 2 GoTo SOU
	IF %ERRORLEVEL% EQU 1 GoTo subADB
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:://///////////////////////////////////////////////////////////////////////////
:::: User Settings ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:Uset
	
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
	ECHO  Suppress Verbose Output: %$SUPPRESS_VERBOSE_N%
	ECHO ************************************************************
	Echo.
	Echo Choose an action from the list:
	Echo.
	Echo [1] Change Log Settings
	Echo [2] Change Domain Settings
	Echo [3] Change AD Settings
	echo [4] Search Parameters
	echo [5] Search Menu
	Echo [6] Main menu
	Echo.
	Choice /c 123456
	Echo.
	If ERRORLevel 6 GoTo Menu
	If ERRORLevel 5 GoTo Search
	If ERRORLevel 4 GoTo uSetSP
	If ERRORLevel 3 GoTo uSetADS
	If ERRORLevel 2 GoTo uSetDC
	If ERRORLevel 1 GoTo uSetL
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Log Settings :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:UsetL
	
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
	echo [5] Settings
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
	IF DEFINED $cUSERPASSWORD echo %$cUSERPASSWORD%> "%$LOGPATH%\var\var_$cUSERPASSWORD.txt"
	echo Provide Credentials ^(searches Name ^& UPN^)
	@powershell Write-Host "Leave password blank to abort!" -ForegroundColor Cyan
	SET $cUSERPASSWORD=
	SET /P $DOMAIN_USER=UserName:
	SET /P $cUSERPASSWORD=Password:
	IF NOT DEFINED $cUSERPASSWORD (
		SET /P $cUSERPASSWORD= <  "%$LOGPATH%\var\var_$cUSERPASSWORD.txt"
		DEL /F /Q "%$LOGPATH%\var\var_$cUSERPASSWORD.txt"
		GoTo uSetDC
	)
	:: Using name search
	dsquery user forestroot -o rdn -scope subtree -domain %$domain% -name "%$DOMAIN_USER%" -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc 2> nul 1> "%$LOGPATH%\var\var_Custom_User_Domain_Authentication.txt"
	::	mainly to capture authentication faliure
	SET $DA_QUERY_RESULT=%ERRORLEVEL%
	IF NOT DEFINED $DA_QUERY_RESULT SET $DA_QUERY_RESULT=0
	echo %$DA_QUERY_RESULT% > "%$LOGPATH%\var\var_$DA_QUERY_RESULT.txt"
	:: skip if name search succeded
	IF %$DA_QUERY_RESULT% EQU 0 GoTo skipDAupn
	:: Check UPN search
	IF %$DA_QUERY_RESULT% NEQ 0 dsquery user forestroot -o rdn -scope subtree -domain %$domain% -UPN "%$DOMAIN_USER%*" -u %$DOMAIN_USER% -p %$cUSERPASSWORD% -limit %$sLimit% -uc 2> nul 1> "%$LOGPATH%\var\var_Custom_User_Domain_Authentication.txt"
	SET $DA_QUERY_RESULT=%ERRORLEVEL%
	IF NOT DEFINED $DA_QUERY_RESULT SET $DA_QUERY_RESULT=0
	echo %$DA_QUERY_RESULT% > "%$LOGPATH%\var\var_$DA_QUERY_RESULT.txt"
:skipDAupn

	::	Athentication error -2147023570
	IF %$DA_QUERY_RESULT% EQU -2147023570 (
		SET /P $DOMAIN_USER= < "%$LOGPATH%\var\var_$DOMAIN_USER.txt"
		DEL /F /Q "%$LOGPATH%\var\var_$cUSERPASSWORD.txt"
		@powershell Write-Host "Authentication failed!" -ForegroundColor Red
		echo.
		@powershell Write-Host "Try again?" -ForegroundColor white
	)
	IF %$DA_QUERY_RESULT% NEQ -2147023570 GoTo skipDACh
	REM choice didn't work in the IF satement, likely needs DELAYEDEXPANSION, which I don't want to do.
	Choice /c yn /m "[y]es or [n]o":
		IF %ERRORLEVEL% EQU 2 GoTo uSetDC
		IF %ERRORLEVEL% EQU 1 GoTo subDA
:skipDACh
	echo.
	echo Domain User Name: %$DOMAIN_USER% ^(%$CHECK_CUSTOM_USER_DOMAIN_ATHENTICATION%^) >> "%$LOGPATH%\ADDS_Tool_Active_Session.log"
	@powershell Write-Host "Success!" -ForegroundColor Green
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

:uSetADS
:: Search AD Settings
	mode con:cols=60 lines=40
	cls
	ECHO ************************************************************
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 	%DATE% %TIME%
	echo.
	Echo		Location: AD Search Settings
	echo.
	ECHO ************************************************************
	Echo.
	Echo Current AD Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
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
	If ERRORLevel 4 GoTo err40
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

	echo  New Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo ------------------------
	echo.
	echo Change AD Search settings?
	Choice /c 12 /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo Uset
	IF %ERRORLEVEL% EQU 1 GoTo uSetADS
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:subADB
:: Search AD Settings
	mode con:cols=60 lines=40
	cls
	ECHO ************************************************************
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 	%DATE% %TIME%
	echo.
	Echo		Location: AD Search Settings
	echo.
	ECHO ************************************************************
	Echo.
	Echo Current AD Search Settings
	Echo ------------------------
	Echo  AD Base: %$AD_BASE%
	Echo  AD Scope: %$AD_SCOPE%
	Echo ------------------------
	echo.
	echo Change AD Base to OU:
	echo format: OU=OU,DC=domain,DC=domainRoot
	SET /P $OU_Base=
	echo "%$OU_Base%">"%$LOGPATH%\var\var_OU_Base.txt"
	echo Checking OU...
	DSQUERY OU "%$OU_Base%" 2> nul
	SET $OU_BASE_ERROR=%ERRORLEVEL%
	IF %$OU_BASE_ERROR% NEQ 0 (
		SET $OU_Base=
		echo Not a valid OU!
		GoTo skipADB
		)
	SET /P $AD_BASE= < "%$LOGPATH%\var\var_OU_Base.txt"
:skipADB
	IF DEFINED $OU_Base GoTo skipOUB
	Echo Try again to set OU?
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo skipOUB
	IF %ERRORLEVEL% EQU 1 GoTo subADB
:skipOUB
	echo AD_BASE changed to: %$OU_Base%
	echo.
	timeout /t 20
	GoTo Uset
	
	
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:uSetSP
:: Search parametrs Settings
	mode con:cols=60 lines=40
	cls
	ECHO ************************************************************
	ECHO		%$PROGRAM_NAME% %$VERSION%
	echo.
	echo		 	%DATE% %TIME%
	echo.
	Echo		Location: Search Parameters
	echo.
	ECHO ************************************************************
	echo.
	Echo ------------------------
	Echo  Sort: %$SORTED_N%
	Echo  Suppress Verbose Output: %$SUPPRESS_VERBOSE_N%
	Echo ------------------------
	echo.
	::	Sorted
	echo %$SORTED% > "%$LOGPATH%\var\var_$SORTED.txt"
	echo Sorted search results?
	echo.
	echo [1] 1=On	{yes}
	echo [2] 0=Off	{no}
	Choice /c 12
	If ERRORLevel 2 (SET $SORTED=0) & (SET $SORTED_N=No)
	If ERRORLevel 1 (SET $SORTED=1) & (SET $SORTED_N=Yes)
	echo.
	::	Suppress Verbose Output
	::	Capture current setting
	echo %$SUPPRESS_VERBOSE% > "%$LOGPATH%\var\var_$SUPPRESS_VERBOSE.txt"
	echo Suppress Verbose output on search results:
	echo.
	echo [1] 1=On	{yes}
	echo [2] 0=Off	{no}
	Choice /c 12
	If ERRORLevel 2 (SET $SUPPRESS_VERBOSE=0) & (SET $SUPPRESS_VERBOSE_N=No)
	If ERRORLevel 1 (SET $SUPPRESS_VERBOSE=1) & (SET $SUPPRESS_VERBOSE_N=Yes)
	echo.
	echo  New Search parameters
	Echo ------------------------
	Echo  Suppress Verbose Output: %$SUPPRESS_VERBOSE_N%
	Echo  Sort: %$SORTED_N%
	Echo ------------------------
	echo.
	echo Change Search parameters?
	Choice /c yn /m "[y]es or [n]o":
	IF %ERRORLEVEL% EQU 2 GoTo Uset
	IF %ERRORLEVEL% EQU 1 GoTo uSetSP
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::Logs:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:Logs
	IF EXIST "%$LOGPATH%\" @explorer "%$LOGPATH%\"
	GoTo menu
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


:::: FUNCTIONS ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Counter :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:fSC
	::	Search Counter
	SET /A $COUNTER_SEARCH+=1
	GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Subroutine for search operators ::::::::::::::::::::::::::::::::::::::::::
:subOperator
	SET $SEARCH_ATTR_%~1=%~1
	@powershell Write-Host "%~1 search operator:" -ForegroundColor Gray
	echo [1] Equal [=]
	echo [2] Approximately equal to [^~=]
	echo [3] Less [^<=]
	echo [4] Greater [^>=]
	echo.
	Choice /c 1234
	If ERRORLevel 4 (SET "$ATTR_%~1_OPERATOR=>=") & (SET "$ATTR_%~1_OPERATOR_DISPLAY=^>^=")
	If ERRORLevel 3 (SET "$ATTR_%~1_OPERATOR=<=") & (SET "$ATTR_%~1_OPERATOR_DISPLAY=^<^=")
	If ERRORLevel 2 (SET "$ATTR_%~1_OPERATOR=~=") & (SET "$ATTR_%~1_OPERATOR_DISPLAY=^~^=")
	If ERRORLevel 1 (SET "$ATTR_%~1_OPERATOR==") & (SET "$ATTR_%~1_OPERATOR_DISPLAY=^=")
GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Function Variable Debug ::::::::::::::::::::::::::::::::::::::::::::::::::
:fVarD
	
	set | FINDSTR /B /C:"$" > "%$LOGPATH%\var\Variable_Debug_%$PID%.txt"
GoTo:EOF
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:::: Search Automatic :::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:SAUTO
	:: Parse parameter 1
	IF /I %$SEARCH_TYPE%==OU set $SEARCH_TYPE=OrganizationalUnit
	:: Create filter
	SET "$FILTER=(objectClass=%$SEARCH_TYPE%)(%$SEARCH_ATTRIBUTE%=%$SEARCH_KEY%)"
::	SET "$FILTER_N=^(objectClass^=%$SEARCH_TYPE%^)^(%$SEARCH_ATTRIBUT%^=%$SEARCH_KEY%^)"
	SET "$FILTER=(&%$FILTER%)"
::	SET "$FILTER_N=^(^&%$FILTER%^)"
::	echo %$FILTER%> "%$LogPath%\var\var_$FILTER.txt"
	:: Start Elapse Time
	call :subSET
	:: Write log headers
	call :sHeader
	:: Credentials
	:: Domain credentials default to blank
	SET $DOMAIN_CREDENTIALS=
	if not "%$SESSION_USER%"=="%$DOMAIN_USER%" SET "$DOMAIN_CREDENTIALS=-u %$DOMAIN_USER% -p %$cUSERPASSWORD%"
	REM If Forestroot, then searches GC Global Catalog.
	set "$AD_SERVER_SEARCH=-s %$DC%.%$DOMAIN%"
	if %$AD_BASE%==forestroot SET "$AD_SERVER_SEARCH=-gc"
	:: Debug
	if %$DEGUB_MODE% EQU 1 CALL :fVarD
	:: Search
		:: Check for sorted
	if %$SORTED% EQU 1 GoTo jumpSAUTOS
	:: Unsorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% > "%$LogPath%\var\var_Last_Search_DN.txt"	
	GoTo skipSAUTOS
:jumpSAUTOS
	:: Sorted
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr name distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -filter "%$FILTER%" -attr distinguishedName -limit %$sLimit% %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% | sort > "%$LogPath%\var\var_Last_Search_DN.txt"	
:skipSAUTOS
	:: Check results
	FOR /F "tokens=3 delims=:" %%K IN ('FIND /I /C "=" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%K> "%$LogPath%\var\var_Last_Search_Count.txt"
	:: remove leading space
	FOR /F "tokens=1 delims= " %%P IN (%$LogPath%\var\var_Last_Search_Count.txt) DO echo %%P> "%$LogPath%\var\var_Last_Search_Count.txt"
	SET /P $LAST_SEARCH_COUNT= < "%$LogPath%\var\var_Last_Search_Count.txt"
	IF %$LAST_SEARCH_COUNT% EQU 0 (
		echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
		TYPE "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$SEARCH_SESSION_LOG%"
		GoTo skipSAUTO
		)
	:: Main output
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Number of search results: %$LAST_SEARCH_COUNT% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_DN.txt"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Munge N_DN file
	IF EXIST "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" DEL /Q /F "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	FOR /F "skip=2 delims=" %%M IN ('FIND /I /V "distinguishedName" "%$LogPath%\var\var_Last_Search_N_DN.txt"') DO echo %%M >> "%$LogPath%\var\var_Last_Search_N_DN_munge.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN_munge.txt" > "%$LogPath%\var\var_Last_Search_N_DN.txt"
	type "%$LogPath%\var\var_Last_Search_N_DN.txt" >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	IF %$SUPPRESS_VERBOSE% EQU 1 GoTo jumpSAUTOO
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Verbose Output: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	:: Detailed Output
	FOR /F "USEBACKQ tokens=* delims=" %%N IN ("%$LogPath%\var\var_Last_Search_DN.txt") DO (
	DSQUERY * %$AD_BASE% -scope %$AD_SCOPE% -limit %$sLimit% -filter "(distinguishedName=%%~N)" -attr name %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo %$SEARCH_TYPE% DN: %%N >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo Details: >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	DSQUERY * -filter "(distinguishedName=%%~N)" -attr * %$AD_SERVER_SEARCH% %$DOMAIN_CREDENTIALS% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo ---------------------------------------------------------------------- >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	)

:jumpSAUTOO
	:: Search counter increment
	Call :fSC
	:: Total Lapse TIme
	call :subTLT
	echo Total Search Time: %$TOTAL_LAPSE_TIME% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo End search: %DATE% %Time% >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	echo. >> "%$LogPath%\%$LAST_SEARCH_LOG%"
	type "%$LogPath%\%$LAST_SEARCH_LOG%" >> "%$LogPath%\%$SEARCH_SESSION_LOG%"

:skipSAUTO
	GoTo end

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::



:::: jump error section :::::::::::::::::::::::::::::::::::::::::::::::::::::::
GoTo end

::!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
::#############################################################################
:: ERROR SECTION
::#############################################################################

:::: Banner :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:ErrBann
	cls
	color 4E
	mode con:cols=80 lines=40
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

:::: error Administrative Privilege :::::::::::::::::::::::::::::::::::::::::::
:err10
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

:::: error RSAT :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:err20

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

:::: Error PW Cache :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:err30
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

:::: Error Under Development ::::::::::::::::::::::::::::::::::::::::::::::::::
:err40
	call :ErrBann
	echo	UNDER CONSTRUCTION
	echo.
	echo	Feature: %$LAST_SEARCH_TYPE% search
	echo.
	timeout /t 60
GoTo Search
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


:::: End session ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:end
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

:::: Credits ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:credits
	:: Exit if run as auto
	if defined $PARAMETER1 exit /B
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