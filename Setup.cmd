:: =================================================================================================
::                                  Program start code for batch file.
:: =================================================================================================

@echo off
@setlocal enabledelayedexpansion
set "dir=%~dp0"
set "bin=!dir!bin\"
set "conf[ath="
set "v=1.0.2"
set "prod="
set "prodtype="
set "prodname="
set "arch="
cd /d "%~dp0"
mode con cols=70 lines=22
color 02
title Microsoft O365 And Office Installation Script

goto :admincheck

:: =================================================================================================
:: =================================================================================================

:: =================================================================================================
::                                             UAC prompt.
:: =================================================================================================

:admincheck
net session > nul 2>&1
if %errorlevel% neq 0 (
    call :banner
    echo                 Requesting Administrator privileges...
    echo.
    echo                 This script requires Administrator privileges....
    goto :UACPrompt
)

if %errorlevel% == 0 (
    call :top
)

:: Creating VBS file for UAC.
:UACPrompt
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
echo UAC.ShellExecute "%~0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
"%temp%\getadmin.vbs"
exit

:: ================================================================================================
:: ================================================================================================

:: =================================================================================================
::                                          Main code.
:: =================================================================================================
:top
cls
call :banner
call :setarch
goto :eof

:banner
cls
color 02
echo ______________________________________________________________________
echo.
echo              Microsoft O365 And Office Installation Script
echo.
echo                           Created by: S4NDM4N
echo                      Home page: harboroftech[.]com
echo.
echo ______________________________________________________________________
echo.
echo.
goto :eof

:setarch
:: Check OS architecture
echo              Finding system architecture...
echo.
for /f "tokens=3" %%a in ('systeminfo ^| findstr /C:"System Type"') do (
    set arch=%%a
    set arch=!arch:-based=!
)

set "confpath=!dir!configs\!arch!\"

if defined arch goto :startsetup

:startsetup
call :banner
echo                      [1] Office 2016.
echo                      [2] Office 2019.
echo                      [3] Office 2021.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Exit the script.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234x /N /M ".                      Enter your choice:"

if /i %errorlevel% equ 1 (set "yy=2016")
if /i %errorlevel% equ 2 (set "yy=2019")
if /i %errorlevel% equ 3 (set "yy=2021")
if /i %errorlevel% equ 4 (goto :activate)
if /i %errorlevel% equ 5 (exit)

if defined yy goto :setup

:setup
call :banner
echo                      [1] Install Office ProPlus !arch! !yy!.
echo                      [2] Install Office Pro !arch! !yy!.
echo                      [3] Install Project Pro !arch! !yy!.
echo                      [4] Install Visio Pro !arch! !yy!.
echo                      [5] Activate Office.
echo           _____________________________________________
echo                      [x] Go to main menu.
echo.
echo ______________________________________________________________________
echo.
echo.
choice /C:12345x /N /M ".                     Enter your choice:"

if /i %errorlevel% equ 1 (set "prodtype=proplus")
if /i %errorlevel% equ 2 (set "prodtype=pro")
if /i %errorlevel% equ 3 (set "prodtype=project")
if /i %errorlevel% equ 4 (set "prodtype=visiopro")
if /i %errorlevel% equ 5 (goto :activate)
if /i %errorlevel% equ 6 (goto :startsetup)

if defined prodtype goto :cmdinstall

:cmdinstall
if !prodtype! EQU proplus set "prodname=Office Professional Plus !yy!"
if !prodtype! EQU pro set "prodname=Office Professional !yy!"
if !prodtype! EQU project set "prodname=Project Professional !yy!"
if !prodtype! EQU visiopro set "prodname=Visio Professional !yy!"

call :banner
echo              Downloading !prodname! Setup...
start /wait /b "" "!bin!setup.exe" /download "!confpath!config_!prodtype!!yy!.xml"
call :banner
echo              Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo              Starting the installation...
start /wait /b "" "!bin!setup.exe" /configure "!confpath!\config_!prodtype!!yy!.xml"
call :banner
echo              Installation completed...
timeout /t 10 /nobreak > nul
call :complete
goto :eof

:complete
call :banner
echo.
echo              Installation completed...
echo.
echo ______________________________________________________________________
echo.
choice /C:yn /N /M ".         Do you wish to activate !prodname!? (y/n)"
goto :end_%errorlevel%

:end_1
goto :activate

:end_2
goto :startsetup

:default
call :banner
echo                 Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :startsetup
goto :eof

:activate
call :banner
echo.
echo              Activating !prodname!...
call "!bin!MAS.cmd"
call :banner
echo              !prodname! activation completed...
timeout /t 10 /nobreak > nul
call :finish
goto :eof

:finish
call :banner
echo.
echo              !prodname! installation completed...
echo.
echo ______________________________________________________________________
echo.
choice /C:yn /N /M ".              Would you like to exit the script? (y/n)"
goto :end_%errorlevel%

:end_1
rd /S /Q "!dir!Office"
exit

:end_2
goto :top
goto :eof
endlocal
:: ================================================================================================
::                                  Program end code for batch file.
:: ================================================================================================