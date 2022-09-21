@ECHO OFF
@REM Initialize Variables 
SET HelpFlag=True

@REM Validate Parameters
IF /i "%~1"=="-help" GOTO :ValidParameter
IF /i "%~1"=="-h" GOTO :ValidParameter
IF /i "%~1"=="-Background" GOTO :ValidParameter
IF /i "%~1"=="" GOTO :ValidParameter
GOTO :InvalidParameter

:ValidParameter
GOTO :Main

:InvalidParameter
ECHO Invalid Parameter: "%~1"
ECHO.
ECHO PARAMETERS
ECHO -Background ^[^<SwitchParameter^>^]
ECHO     Force closes all Office 365 applications and performs a background repair.
ECHO.
ECHO -help / h ^[^<SwitchParameter^>^]
ECHO     Opens help menu.
GOTO :CleanUp

@REM Main body of script.
:Main
@REM Set variable to value of argument
IF NOT "%~1"=="" SET arg1=-ArgumentList %~1

@REM If a Help flag is set then run without "-WindowStyle Hidden" parameter
IF /i NOT "%arg1%"=="-ArgumentList -help" IF /i NOT "%arg1%"=="-ArgumentList -h" SET HelpFlag=False
IF "%HelpFlag%"=="False" SET WindowHide=-WindowStyle Hidden 

PowerShell -ExecutionPolicy Bypass %WindowHide%-File "%~dp0OfficeRepair.ps1" %arg1%
GOTO :CleanUp

@REM Reset Variables 
@REM that way if ran multiple times in the same terminal session previous runs
@REM wont leak environment variables into this run.
:CleanUp
SET "arg1="
SET "HelpFlag="
SET "WindowHide="
GOTO :EOF