:: =================================================================================================
::                                  Program start code for batch file.
:: =================================================================================================

@echo off
@setlocal enabledelayedexpansion
set "dir=%~dp0"
cd /d "%~dp0"
mode con cols=70 lines=30
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
call :options
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

:options
echo                      [1] Install 64 bit packages.
echo                      [2] Install 32 bit packages.
echo           _____________________________________________
echo                      [x] Exit the script.
echo.
echo ______________________________________________________________________
echo.
choice /C:12x /N /M ".                      Enter your choice:"
goto :choice_%errorlevel%

:choice_1
goto :x64

:choice_2
goto :x32

:choice_3
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :top
goto :eof

:x64
call :banner
echo                      [1] Office 2016.
echo                      [2] Office 2019.
echo                      [3] Office 2021.
echo                      [4] Activate Office.
echo                      [b] Go back.
echo           _____________________________________________
echo                      [x] Exit the script.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234bx /N /M ".                      Enter your choice:"
goto :choice64_%errorlevel%

:choice64_1
goto :pro2016_x64

:choice64_2
goto :pro2019_x64

:choice64_3
goto :pro2021_x64

:choice64_4
goto :activate

:choice64_5
goto :top

:choice64_6
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :x64
goto :eof

:pro2016_x64
call :banner
echo                      [1] Install Office ProPlus 2016.
echo                      [2] Install Project Pro 2016.
echo                      [3] Install Visio Pro 2016.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Go to main menu.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234x /N /M ".                     Enter your choice:"
goto :choice2016_%errorlevel%

:choice2016_1
goto :cmd2016_x64_proplus

:choice2016_2
goto :cmd2016_x64_project

:choice2016_3
goto :cmd2016_x64_visio

:choice2016_4
goto :activate

:choice2016_5
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :pro2016_x64
goto :eof

:pro2019_x64
call :banner
echo                      [1] Install Office ProPlus 2019.
echo                      [2] Install Project Pro 2019.
echo                      [3] Install Visio Pro 2019.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Got to main menu.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234x /N /M ".                     Enter your choice:"
goto :choice2019_%errorlevel%

:choice2019_1
goto :cmd2019_x64_proplus

:choice2019_2
goto :cmd2019_x64_project

:choice2019_3
goto :cmd2019_x64_visio

:choice2019_4
goto :activate

:choice2019_5
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :pro2019_x64
goto :eof

:pro2021_x64
call :banner
echo                      [1] Install Office ProPlus 2021.
echo                      [2] Install Project Pro 2021.
echo                      [3] Install Visio Pro 2021.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Go to main menu.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234x /N /M ".                     Enter your choice:"
goto :choice2021_%errorlevel%

:choice2021_1
goto :cmd2021_x64_proplus

:choice2021_2
goto :cmd2021_x64_project

:choice2021_3
goto :cmd2021_x64_visio

:choice2021_4
goto :activate

:choice2021_5
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :pro2021_x64
goto :eof

:cmd2016_x64_proplus
set "prod=Office Professional Plus 2016"
call :banner
echo                      Downloading Office ProPlus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_proplus2016.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_proplus2016.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2016_x64_project
set "prod=Project Professional 2016"
call :banner
echo                      Downloading Project Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_project2016.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_project2016.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2016_x64_visio
set "prod=Visio Professional 2016"
call :banner
echo                      Downloading Visio Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_visiopro2016.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_visiopro2016.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2019_x64_proplus
set "prod=Office Professional Plus 2019"
call :banner
echo                      Downloading Visio Professional Plus 2019 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_proplus2019.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_proplus2019.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2019_x64_project
set "prod=Project Professional 2019"
call :banner
echo                      Downloading Project Professional Plus 2019 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_project2019.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_project2019.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2019_x64_visio
set "prod=Visio Professional 2019"
call :banner
echo                      Downloading Visio Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_visiopro2019.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_visiopro2019.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2021_x64_proplus
set "prod=Office Professional Plus 2021"
call :banner
echo                      Downloading Visio Professional Plus 2021 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_proplus2021.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_proplus2021.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2021_x64_project
set "prod=Project Professional 2021"
call :banner
echo                      Downloading Project Professional Plus 2021 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_project2021.xml"
call :banner
echo Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_project2021.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:cmd2021_x64_visio
set "prod=Visio Professional 2021"
call :banner
echo                      Downloading Visio Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config64pro\config_visiopro2021.xml"
call :banner
echo                      Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                      Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config64pro\config_visiopro2021.xml"
call :banner
echo                      Installation completed...
timeout /t 10 /nobreak > nul
call :sendx64
goto :eof

:x32
call :banner
echo                      [1] Office 2016.
echo                      [2] Office 2019.
echo                      [3] Office 2021.
echo                      [4] Activate Office.
echo                      [b] Go back.
echo           _____________________________________________
echo                      [x] Exit the script.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234bx /N /M ".                      Enter your choice:"
goto :choice32_%errorlevel%

:choice32_1
goto :pro2016_x32

:choice32_2
goto :pro2019_x32

:choice32_3
goto :pro2021_x32

:choice32_4
goto :activate

:choice32_5
goto :top

:choice32_6
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :x32
goto :eof

:pro2016_x32
call :banner
echo                      [1] Install Office ProPlus 2016.
echo                      [2] Install Project Pro 2016.
echo                      [3] Install Visio Pro 2016.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Got to main menu.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234x /N /M ".                      Enter your choice:"
goto :choice2016_%errorlevel%

:choice2016_1
goto :cmd2016_x32_proplus

:choice2016_2
goto :cmd2016_x32_project

:choice2016_3
goto :cmd2016_x32_visio

:choice2016_4
goto :activate

:choice2016_5
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :pro2016_x32
goto :eof

:pro2019_x32
call :banner
echo                      [1] Install Office ProPlus 2019.
echo                      [2] Install Project Pro 2019.
echo                      [3] Install Visio Pro 2019.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Got to main menu.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234x /N /M ".                      Enter your choice:"
goto :choice2019_%Choice2019%

:choice2019_1
goto :cmd2019_x32_proplus

:choice2019_2
goto :cmd2019_x32_project

:choice2019_3
goto :cmd2019_x32_visio

:choice2019_4
goto :activate

:choice2019_5
exit

:default
call :banner
echo                      Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :pro2019_x32
goto :eof

:pro2021_x32
call :banner
echo                      [1] Install Office ProPlus 2021 setup.
echo                      [2] Install Project Pro 2021.
echo                      [3] Install Visio Pro 2021.
echo                      [4] Activate Office.
echo           _____________________________________________
echo                      [x] Got to main menu.
echo.
echo ______________________________________________________________________
echo.
choice /C:1234x /N /M ".                      Enter your choice:"
goto :choice2021_%errorlevel%

:choice2021_1
goto :cmd2021_x32_proplus

:choice2021_2
goto :cmd2021_x32_project

:choice2021_3
goto :cmd2021_x32_visio

:choice2021_4
goto :activate

:choice2021_x
:choice2021_X
exit

:default
call :banner
echo                 Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :pro2021_x32
goto :eof

:cmd2016_x32_proplus
set "prod=Office Professional Plus 2016"
call :banner
echo                 Downloading Office ProPlus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_proplus2016.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_proplus2016.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2016_x32_project
set "prod=Project Professional 2016"
call :banner
echo                 Downloading Project Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_project2016.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_project2016.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2016_x32_visio
set "prod=Visio Professional 2016"
call :banner
echo                 Downloading Visio Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_visiopro2016.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_visiopro2016.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2019_x32_proplus
set "prod=Office Professional Plus 2019"
call :banner
echo                 Downloading Visio Professional Plus 2019 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_proplus2019.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_proplus2019.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2019_x32_project
set "prod=Project Professional 2019"
call :banner
echo                 Downloading Project Professional Plus 2019 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_project2019.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_project2019.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2019_x32_visio
set "prod=Visio Professional 2019"
call :banner
echo                 Downloading Visio Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_visiopro2019.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_visiopro2019.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2021_x32_proplus
set "prod=Office Professional Plus 2021"
call :banner
echo                 Downloading Visio Professional Plus 2021 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_proplus2021.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_proplus2021.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2021_x32_project
set "prod=Project Professional 2021"
call :banner
echo                 Downloading Project Professional Plus 2021 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_project2021.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_project2021.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:cmd2021_x32_visio
set "prod=Visio Professional 2021"
call :banner
echo                 Downloading Visio Professional Plus 2016 Setup...
start /wait /b "" "!dir!setup.exe" /download "!dir!Config32pro\config_visiopro2021.xml"
call :banner
echo                 Downloading finished...
timeout /t 10 /nobreak > nul
call :banner
echo                 Starting the installation...
start /wait /b "" "!dir!setup.exe" /configure "!dir!Config32pro\config_visiopro2021.xml"
call :banner
echo                 Installation completed...
timeout /t 10 /nobreak > nul
call :sendx32
goto :eof

:activate
call :banner
echo.
echo              Activating !prod!...
call "!dir!Activator\MAS.cmd"
call :banner
echo              !prod! activation completed...
timeout /t 10 /nobreak > nul
call :finish
goto :eof

:sendx64
call :banner
echo.
echo              Installation completed...
echo.
echo ______________________________________________________________________
echo.
choice /C:yn /N /M ".         Do you wish to activate %prod%? (y/n)"
goto :endx64_%errorlevel%

:endx64_1
goto :activate

:endx64_2
goto :x64

:default
call :banner
echo                 Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :x64
goto :eof

:sendx32
call :banner
echo.
echo              Installation completed...
echo.
echo ______________________________________________________________________
echo.
choice /C:yn /N /M ".         Do you wish to activate %prod%? (y/n)"
goto :endx32_%errorlevel%

:endx32_1
goto :activate

:endx32_2
goto :x64

:default
call :banner
echo                 Not a valid choice, try again.
timeout /t 5 /nobreak > nul
goto :x64
goto :eof

:finish
call :banner
echo.
echo              !prod! installation completed...
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