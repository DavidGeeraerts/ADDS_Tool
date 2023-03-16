::::SHA256 Generator:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::#############################################################################
::							#DESCRIPTION#
::
::	SCRIPT STYLE: Interactive
::	Script generates a SHA256 file for a commandlet
::
::#############################################################################

::::Developer::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Author:		David Geeraerts
:: Location:	Olympia, Washington USA
:: E-Mail:		dgeeraerts.evergreen@gmail.com
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::License::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Copyleft License(s)
:: GNU GPL v3 (General Public License)
:: https://www.gnu.org/licenses/gpl-3.0.en.html
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::Versioning Schema::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::		VERSIONING INFORMATION												 ::		
::		Semantic Versioning used											 ::
::		http://semver.org/													 ::
::		Major.Minor.Revision												 ::
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::Command shell::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
@Echo Off
@SETLOCAL enableextensions
SET $PROGRAM_NAME=SHA256-Generator
SET $Version=1.0.0
SET $BUILD=2021-01-15 07:00
Title %$PROGRAM_NAME%
Prompt $G
color 0B
mode con:cols=80
mode con:lines=20
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::User Variables ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Declare variables
::	All User variables are set within here.
::		(configure variables)
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

SET $CMD=ADDS_Tool.cmd

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
::##### Everything below here is 'hard-coded' [DO NOT MODIFY] #####
::
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::Directory::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:CD
	:: Launched from directory
	SET "$PROGRAM_PATH=%~dp0"
	cd /D "%$PROGRAM_PATH%"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::Commandlet:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

	SET $PARAMETER1=%~1
	IF DEFINED $PARAMETER1 SET $CMD=%$PARAMETER1%
:CMD
	IF NOT DEFINED $CMD (
		echo Generate SHA256 for Commandlet...
		echo.
		dir /A:-D /B | FIND /I ".cmd"
		echo.
		SET /P $CMD=Commandlet Name:
	)
	DIR /B /A:-D | FIND /I "%$CMD%" || (SET $CMD=)
	IF NOT DEFINED $CMD (
		Echo Not Found!
		GoTo CMD
	)
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::	
	
	
:::: Generate SHA256 ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:Generator
	IF EXIST "SHA256.txt" DEL /Q /F SHA256.txt
	FOR /F "skip=1 tokens=1" %%P IN ('certUtil -hashfile "%$CMD%" SHA256') DO ECHO %%P>> SHA256.txt
	SET /P $SHA256= < SHA256.txt
	ECHO %$SHA256%> "SHA256.txt"
	echo SHA256 generated:
	ECHO %$SHA256%
	echo.
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

timeout /t 20
EXIT