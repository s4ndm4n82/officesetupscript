:: =================================================================================================
::                                  Program start code for batch file.
:: =================================================================================================

@echo off
@setlocal enabledelayedexpansion
set "dir=%~dp0"
set "bin=!dir!bin\"
set "conf[ath="
set "v=1.2.2"
set "yy="
set "mo="
set "mp="
set "mv="
set "prod="
set "prodtype="
set "prodname="
set "arch="
cd /d "%~dp0"
mode con cols=70 lines=22
color 02
title OSS v !v!

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
:: Checking if the previous "OFFICE" folder exitst.
:: If it's there it will be removed.
if exist "OFFICE" (
    rd /s /q "OFFICE"
)
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
echo                      [4] Microsoft 365.
echo                      [5] Microsoft Project.
echo                      [6] Microsoft Visio.
echo                      [7] Activate Office.
echo           _____________________________________________
echo                      [x] Exit the script.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234567x /N /M ".                      Enter your choice:"

if /i %errorlevel% equ 1 (set "yy=2016")
if /i %errorlevel% equ 2 (set "yy=2019")
if /i %errorlevel% equ 3 (set "yy=2021")
if /i %errorlevel% equ 4 (set "mo=365")
if /i %errorlevel% equ 5 (set "mp=365")
if /i %errorlevel% equ 6 (set "mv=365")
if /i %errorlevel% equ 7 (goto :activate)
if /i %errorlevel% equ 8 (exit)

if defined yy goto :oselect
if defined mo goto :365select
if defined mp goto :pselect
if defined mv goto :vselect

:365select
set "yy=365"
call :banner
echo                      [1] Microsoft 365.
echo                      [2] Microsoft 365 Small Business Premium.
echo                      [3] Microsoft 365 Family ^& Personal.
echo                      [4] Microsoft 365 Apps for Business.
echo                      [5] Microsoft 365 Apps for Enterprise.
echo                      [6] Activate Office.
echo           _____________________________________________
echo                      [x] Go to main menu.
echo.
echo ______________________________________________________________________
echo.
echo.
choice /C:123456x /N /M ".                     Enter your choice:"

if /i %errorlevel% equ 1 (set "prodtype=o")
if /i %errorlevel% equ 2 (set "prodtype=osb")
if /i %errorlevel% equ 3 (set "prodtype=ohp")
if /i %errorlevel% equ 4 (set "prodtype=ob")
if /i %errorlevel% equ 5 (set "prodtype=oe")
if /i %errorlevel% equ 6 (goto :activate)
if /i %errorlevel% equ 7 (goto :startsetup)
if defined prodtype goto :cmdinstall

:oselect
call :banner
echo                      [1] Office Professional Plus !arch! !yy!.
echo                      [2] Office Professional !arch! !yy!.
echo                      [3] Office Home and Business  !arch! !yy!.
echo                      [4] Office Home and Student !arch! !yy!.
echo                      [5] Office Personal !arch! !yy!.
echo                      [6] Office Standard !arch! !yy!.
echo                      [7] Activatse Office.
echo           _____________________________________________
echo                      [x] Go to main menu.
echo.
echo ______________________________________________________________________
echo.
echo.
choice /C:1234567x /N /M ".                     Enter your choice:"

if /i %errorlevel% equ 1 (set "prodtype=proplus")
if /i %errorlevel% equ 2 (set "prodtype=pr")
if /i %errorlevel% equ 3 (set "prodtype=hb")
if /i %errorlevel% equ 4 (set "prodtype=hs")
if /i %errorlevel% equ 5 (set "prodtype=p")
if /i %errorlevel% equ 6 (set "prodtype=s")
if /i %errorlevel% equ 7 (goto :activate)
if /i %errorlevel% equ 8 (goto :startsetup)

if defined prodtype goto :cmdinstall

:pselect
set "prodtype=project"
call :banner
echo                      [1] Project Professional 2016 !arch!.
echo                      [2] Project Professional 2019 !arch!.
echo                      [3] Project Professional 2021 !arch!.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Go to main menu.
echo.
echo ______________________________________________________________________
echo.
echo.
choice /C:1234x /N /M ".                     Enter your choice:"

if /i %errorlevel% equ 1 (set "yy=2016")
if /i %errorlevel% equ 2 (set "yy=2019")
if /i %errorlevel% equ 3 (set "yy=2021")
if /i %errorlevel% equ 4 (goto :activate)
if /i %errorlevel% equ 5 (goto :startsetup)

if defined yy goto :cmdinstall

:vselect
set "prodtype=visiopro"
call :banner
echo                      [1] Visio Professional 2016 !arch!.
echo                      [2] Visio Professional 2019 !arch!.
echo                      [3] Visio Professional 2021 !arch!.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Go to main menu.
echo.
echo ______________________________________________________________________
echo.
echo.
choice /C:1234x /N /M ".                     Enter your choice:"

if /i %errorlevel% equ 1 (set "yy=2016")
if /i %errorlevel% equ 2 (set "yy=2019")
if /i %errorlevel% equ 3 (set "yy=2021")
if /i %errorlevel% equ 4 (goto :activate)
if /i %errorlevel% equ 5 (goto :startsetup)

if defined yy goto :cmdinstall

:cmdinstall
if !prodtype! EQU o set "prodname=Microsoft !yy!"
if !prodtype! EQU osb set "prodname=Microsoft !yy! Small Business Premium"
if !prodtype! EQU ohp set "prodname=Microsoft !yy! Family & Personal"
if !prodtype! EQU ob set "prodname=Microsoft !yy! Apps for Business"
if !prodtype! EQU oe set "prodname=Microsoft !yy! Apps for Enterprise"
if !prodtype! EQU proplus set "prodname=Office Professional Plus !yy!"
if !prodtype! EQU pro set "prodname=Office Professional !yy!"
if !prodtype! EQU project set "prodname=Project Professional !yy!"
if !prodtype! EQU visiopro set "prodname=Visio Professional !yy!"
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