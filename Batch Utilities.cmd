@echo off

@rem Checking & Prompting for run as administrator
@rem --------------------------------------------------------------------------
@rem Checking for permissions...

if "%PROCESSOR_ARCHITECTURE%" equ "amd64" ( 
    "%SYSTEMROOT%\SysWOW64\cacls.exe" "%SYSTEMROOT%\SysWOW64\config\system">nul 2>&1 
) else ( 
    "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system">nul 2>&1 
)

if '%errorlevel%' neq '0' (
    @rem Administrator permissions not found

    echo Requesting administrative privileges...
    goto UACPrompt
) else ( 
    @rem Administrator permissions found
    
    goto gotAdmin
)

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params= %*
    echo UAC.ShellExecute "cmd.exe", "/c ""%~s0"" %params:"=""%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"
@rem --------------------------------------------------------------------------  

title Batch Utilies


@REM 
@REM  Varriables
@REM 

set wantsRestartNow=n
set deviceNeedsRestart=false

call :mainMenu 1 1

exit /b

:mainMenu <1=current-row> <2=current-col>
    setlocal
    title Batch Utilities
    cls

    @rem Set Varriables
    set /a minRows=1
    set /a maxRows=5
    set /a minCols=1
    set /a maxCols=2
    set /a currentRow=%1
    set /a currentCol=%2

    if %currentRow% gtr %maxRows% set /a currentRow=%maxRows%
    if %currentRow% lss %minRows% set /a currentRow=%minRows%
    if %currentCol% gtr %maxCols% set /a currentCol=%maxCols%
    if %currentCol% lss %minCols% set /a currentCol=%minCols%
    
    set /a nextRow=%currentRow%
    set /a nextCol=%currentCol%

    set optionChar_about=[ ]
    if %currentRow% equ 1 if %currentCol% equ 1 set optionChar_about=[X]

    set optionChar_changelog=[ ]
    if %currentRow% equ 1 if %currentCol% equ 2 set optionChar_changelog=[X]

    set optionChar_deviceInfo=[ ]
    if %currentRow% equ 2 if %currentCol% equ 1 set optionChar_deviceInfo=[X]

    set optionChar_grantLocalAdmin=[ ]
    if %currentRow% equ 2 if %currentCol% equ 2 set optionChar_grantLocalAdmin=[X]

    set optionChar_openAppsInFolder=[ ]
    if %currentRow% equ 3 if %currentCol% equ 1 set optionChar_openAppsInFolder=[X]

    set optionChar_revealWiFiPasswords=[ ]
    if %currentRow% equ 3 if %currentCol% equ 2 set optionChar_revealWiFiPasswords=[X]

    set optionChar_getWindowsActivationKey=[ ]
    if %currentRow% equ 4 if %currentCol% equ 1 set optionChar_getWindowsActivationKey=[X]

    set optionChar_fixUnquotedServicePaths=[ ]
    if %currentRow% equ 4 if %currentCol% equ 2 set optionChar_fixUnquotedServicePaths=[X]

    set optionChar_getBitLockerKey=[ ]
    if %currentRow% equ 5 if %currentCol% equ 1 set optionChar_getBitLockerKey=[X]

    set optionChar_exit=[ ]
    if %currentRow% equ 5 if %currentCol% equ 2 set optionChar_exit=[X]
   
    @REM echo %currentRow% %currentCol%
    echo    +--------------------------------------------------------------+
    echo    ^|                        BATCH UTILITES                        ^|
    echo    ^|                         Version 1.1.0                        ^|
    echo    +--------------------------------------------------------------+
    echo            Use the arrow keys on the keyboard to navigate.
    echo                Use the enter key to make a selection.
    echo               Use the escape key to close the utility.
    echo:
    echo          %optionChar_about% About                        %optionChar_changelog% Changelog
    echo          %optionChar_deviceInfo% Device Information           %optionChar_grantLocalAdmin% Grant Local Admin Permissions 
    echo          %optionChar_openAppsInFolder% Open Applications In Folder  %optionChar_revealWiFiPasswords% Reveal All Saved Wi-Fi Passwords
    echo          %optionChar_getWindowsActivationKey% Get Windows Activation Key   %optionChar_fixUnquotedServicePaths% Quote Unquoted Service Paths                         
    echo          %optionChar_getBitLockerKey% Get BitLocker Key            %optionChar_exit% Exit                         
    echo:

    for /F %%k in ('PowerShell Write-Host $Host.UI.RawUI.ReadKey(\"NoEcho,IncludeKeyDown\"^).VirtualKeyCode') do set "key=%%k"

    @rem Enter key
    if %key% equ 13 if %currentRow% equ 1 if %currentCol% equ 1 endlocal & goto about
    if %key% equ 13 if %currentRow% equ 1 if %currentCol% equ 2 endlocal & goto changelog
    if %key% equ 13 if %currentRow% equ 2 if %currentCol% equ 1 endlocal & goto deviceInfo
    if %key% equ 13 if %currentRow% equ 2 if %currentCol% equ 2 endlocal & goto grantLocalAdminPermissions
    if %key% equ 13 if %currentRow% equ 3 if %currentCol% equ 1 endlocal & goto openAppsInFolder
    if %key% equ 13 if %currentRow% equ 3 if %currentCol% equ 2 endlocal & goto revealWiFiPasswords
    if %key% equ 13 if %currentRow% equ 4 if %currentCol% equ 1 endlocal & goto getWindowsActivationKey
    if %key% equ 13 if %currentRow% equ 4 if %currentCol% equ 2 endlocal & goto quoteUnquotedServicePaths
    if %key% equ 13 if %currentRow% equ 5 if %currentCol% equ 1 endlocal & goto getBitLockerKey
    if %key% equ 13 if %currentRow% equ 5 if %currentCol% equ 2 endlocal & goto exitUtility

    @rem Escape key
    if %key% equ 27 goto exitUtility

    @rem Left key
    if %key% equ 37 set /a nextCol=%currentCol% - 1

    @rem Up key
    if %key% equ 38 set /a nextRow=%currentRow% - 1
    
    @rem Right key
    if %key% equ 39 set /a nextCol=%currentCol% + 1
    
    @rem Down key
    if %key% equ 40 set /a nextRow=%currentRow% + 1

    @REM echo %nextRow% %nextCol%
    @REM pause
    call :mainMenu %nextRow% %nextCol%
    exit /b


@REM 
@REM  Feature Functions
@REM 

:about
    title Batch Utilities - About
    cls

    echo About Batch Utilities
    echo:
    echo Version: 1.1.0
    echo Author: Jacob Paulin
    echo GitHub: TBD
    echo:

    echo [Press any key to return to the main menu]
    pause >nul
    call :mainMenu 1 1
    

:changelog
    title Batch Utilities - Changelog
    cls

    echo [------------------------------]
    echo Version 1.1.0
    echo Released 2022-11-16
    echo:
    echo Updated the main menu to be interactive.
    echo Added a feature to find the BitLocker recovery key.
    echo Added a feature to find the Windows activation key.
    echo Added a feature to quote all unquoted service paths.
    echo:

    echo [------------------------------]
    echo Version 1.0.0
    echo Released 2022-11-16
    echo:
    echo First release of the utility.

    echo:
    echo:
    echo [Press any key to return to the main menu]
    pause >nul
    call :mainMenu 1 2


:deviceInfo
    title Batch Utilities - Device Information
    cls

    echo Gathering device information...

    for /f "tokens=2 delims==" %%i in ('"WMIC cpu get name /format:list"') do set sysInfo_cpu_name=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC cpu get threadcount /format:list"') do set sysInfo_cpu_threadCount=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC cpu get numberofcores /format:list"') do set sysInfo_cpu_numberOfCores=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC bios get SerialNumber /format:list"') do set sysInfo_bios_serialNumber=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC bios get Version /format:list"') do set sysInfo_bios_version=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC computersystem get name /format:list"') do set sysInfo_computerSystem_name=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC computersystem get SystemFamily /format:list"') do set sysInfo_computerSystem_systemFamily=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC computersystem get SystemType /format:list"') do set sysInfo_computerSystem_systemType=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC os get Caption /format:list"') do set sysInfo_os_caption=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC os get BuildNumber /format:list"') do set sysInfo_os_buildNumber=%%i
    for /f "tokens=2 delims==" %%i in ('"WMIC os get Version /format:list"') do set sysInfo_os_version=%%i
    for /f "usebackq tokens=1 delims=" %%i in (`PowerShell -command "(GWMI Win32_PhysicalMemory|Measure Capacity -Sum).Sum/1MB"`) do set sysInfo_ram_totalMB=%%i MB
    for /f "tokens=2 delims==" %%i in ('"WMIC memorychip get Speed /format:list"') do set sysInfo_ram_clockSpeed=%%i MHz
    for /f "tokens=1-3" %%n in ('"WMIC /node:"localhost" logicaldisk get Name,Size,FreeSpace | find /i "C:""') do ( 
        @SETLOCAL ENABLEEXTENSIONS
        @SETLOCAL ENABLEDELAYEDEXPANSION
        
        set FreeBytes=%%n & set TotalBytes=%%p
        set TotalGB=0
        set FreeGB=0
        set gbConvertVal=1074

        set /a TotalSpace=!TotalBytes:~0,-6! / !gbConvertVal!
        set /a FreeSpace=!FreeBytes:~0,-7! / !gbConvertVal!

        set TotalGB=!TotalSpace!
        set FreeGB=!FreeSpace!

        set /A TotalUsed=!TotalSpace! - !FreeSpace!
        set /A MULTIUSED=!TotalUsed!*100
        set /A PERCENTUSED=!MULTIUSED!/!TotalGB!
        
        set sysInfo_disk_totalSpace=!TotalGB! GB
        set sysInfo_disk_freeSpace=!FreeGB! GB
        set sysInfo_disk_percentUsed=!PERCENTUSED!%%
    )


    @rem Find longest Line
    set /a longestLineLength=0

    @rem Test Line 2
        set /a testNumber=0
        call :strlen testNumberLen sysInfo_computerSystem_name
        set /a testNumber=%testNumberLen% + 37
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 4
        set /a testNumber=0
        call :strlen testNumber sysInfo_computerSystem_name
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 5
        set /a testNumber=0
        call :strlen testNumber sysInfo_computerSystem_systemFamily
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 6
        set /a testNumber=0
        call :strlen testNumber sysInfo_bios_serialNumber
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 7
        set /a testNumber=0
        call :strlen testNumber sysInfo_bios_version
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 8
        set /a testNumber=0
        call :strlen testNumber sysInfo_os_caption
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 9
        set /a testNumber=0
        call :strlen testNumber sysInfo_os_buildNumber
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 10
        set /a testNumber=0
        call :strlen testNumber sysInfo_os_version
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 11
        set /a testNumber=0
        call :strlen testNumber sysInfo_computerSystem_systemType
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 12
        call set "sysInfo_cpu_name=%%sysInfo_cpu_name:(=^(%%"
        call set "sysInfo_cpu_name=%%sysInfo_cpu_name:)=^)%%"
        set /a testNumber=0
        call :strlen testNumber sysInfo_cpu_name
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 13
        set /a testNumber=0
        call :strlen testNumber sysInfo_cpu_threadCount
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 14
        set /a testNumber=0
        call :strlen testNumber sysInfo_cpu_numberOfCores
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 15
        set /a testNumber=0
        call :strlen testNumber sysInfo_ram_totalMB
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 16
        set /a testNumber=0
        call :strlen testNumber sysInfo_ram_clockSpeed
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 17
        set /a testNumber=0
        call :strlen testNumber sysInfo_disk_totalSpace
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 18
        set /a testNumber=0
        call :strlen testNumber sysInfo_disk_freeSpace
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    @rem Test Line 19
        set /a testNumber=0
        call :strlen testNumber sysInfo_disk_percentUsed
        set /a testNumber=%testNumber% + 25
        if %testNumber% gtr %longestLineLength% set /a longestLineLength=%testNumber%

    cls

    @rem LINE 1
        set printLine_1=----------------------------------------------------------------------------------------------------------------------------------------------------
        call set printLine_1=%%printLine_1:~0,%longestLineLength%%%
        echo +-%printLine_1%-+

    @rem LINE 2
        @rem Min 37, Max 52 (before space and line on each end)
        set printLine_2=System Info ^For Device %sysInfo_computerSystem_name%                                                                                                  
        call set printLine_2=%%printLine_2:~0,%longestLineLength%%%
        echo ^| %printLine_2% ^|

    @rem LINE 3
        set printLine_3=----------------------------------------------------------------------------------------------------------------------------------------------------
        call set printLine_3=%%printLine_3:~0,%longestLineLength%%%
        echo +-%printLine_3%-+

    @rem LINE 4
        set printLine_4=Device Name:             %sysInfo_computerSystem_name%                                                                                                  
        call set printLine_4=%%printLine_4:~0,%longestLineLength%%%
        echo ^| %printLine_4% ^|

    @rem LINE 5
        set printLine_5=Device Model:            %sysInfo_computerSystem_systemFamily%                                                                                                  
        call set printLine_5=%%printLine_5:~0,%longestLineLength%%%
        echo ^| %printLine_5% ^|

    @rem LINE 6
        set printLine_6=Device Serial Number:    %sysInfo_bios_serialNumber%                                                                                                  
        call set printLine_6=%%printLine_6:~0,%longestLineLength%%%
        echo ^| %printLine_6% ^|

    @rem LINE 7
        set printLine_7=BIOS Version:            %sysInfo_bios_version%                                                                                                  
        call set printLine_7=%%printLine_7:~0,%longestLineLength%%%
        echo ^| %printLine_7% ^|

    @rem LINE 8
        set printLine_8=Windows Edition:         %sysInfo_os_caption%                                                                                                  
        call set printLine_8=%%printLine_8:~0,%longestLineLength%%%
        echo ^| %printLine_8% ^|

    @rem LINE 9
        set printLine_9=Windows Build:           %sysInfo_os_buildNumber%                                                                                                  
        call set printLine_9=%%printLine_9:~0,%longestLineLength%%%
        echo ^| %printLine_9% ^|

    @rem LINE 10
        set printLine_10=Windows Version:         %sysInfo_os_version%                                                                                                  
        call set printLine_10=%%printLine_10:~0,%longestLineLength%%%
        echo ^| %printLine_10% ^|

    @rem LINE 11
        set printLine_11=System Architecture:     %sysInfo_computerSystem_systemType%                                                                                                  
        call set printLine_11=%%printLine_11:~0,%longestLineLength%%%
        echo ^| %printLine_11% ^|

    @rem LINE 12
        call set "sysInfo_cpu_name=%%sysInfo_cpu_name:^(=(%%"
        call set "sysInfo_cpu_name=%%sysInfo_cpu_name:^)=)%%"
        set printLine_12=CPU Model:               %sysInfo_cpu_name%                                                                                                  
        call set printLine_12=%%printLine_12:~0,%longestLineLength%%%
        echo ^| %printLine_12% ^|

    @rem LINE 13
        set printLine_13=CPU Threads:             %sysInfo_cpu_threadCount%                                                                                                  
        call set printLine_13=%%printLine_13:~0,%longestLineLength%%%
        echo ^| %printLine_13% ^|

    @rem LINE 14
        set printLine_14=CPU Cores:               %sysInfo_cpu_numberOfCores%                                                                                                  
        call set printLine_14=%%printLine_14:~0,%longestLineLength%%%
        echo ^| %printLine_14% ^|

    @rem LINE 15
        set printLine_15=Total RAM:               %sysInfo_ram_totalMB%                                                                                                  
        call set printLine_15=%%printLine_15:~0,%longestLineLength%%%
        echo ^| %printLine_15% ^|

    @rem LINE 16
        set printLine_16=RAM Clock Speed:         %sysInfo_ram_clockSpeed%                                                                                                  
        call set printLine_16=%%printLine_16:~0,%longestLineLength%%%
        echo ^| %printLine_16% ^|

    @rem LINE 17
        set printLine_17=Total Disk Space ^(C:^):   %sysInfo_disk_totalSpace%                                                                                                  
        call set printLine_17=%%printLine_17:~0,%longestLineLength%%%
        echo ^| %printLine_17% ^|

    @rem LINE 18
        set printLine_18=Free Disk Space ^(C:^):    %sysInfo_disk_freeSpace%                                                                                                  
        call set printLine_18=%%printLine_18:~0,%longestLineLength%%%
        echo ^| %printLine_18% ^|

    @rem LINE 19
        set printLine_19=Used Disk Space ^(C:^):    %sysInfo_disk_percentUsed%                                                                                                  
        call set printLine_19=%%printLine_19:~0,%longestLineLength%%%
        echo ^| %printLine_19% ^|

    @rem LINE 20
        set printLine_20=----------------------------------------------------------------------------------------------------------------------------------------------------
        call set printLine_20=%%printLine_20:~0,%longestLineLength%%%
        echo +-%printLine_20%-+

    echo:

    echo [Press any key to return to the main menu]
    pause >nul
    call :mainMenu 2 1


:grantLocalAdminPermissions
    title Batch Utilities - Grant Local Admin Permissions
    cls

    echo If you are trying to give a domain user local admin permissions on their device
    echo don't forget to prefix their username with the domain ^(DOAMIN\USER^)
    echo:
    set /p username=Enter username of user you wish to grand local administrator permissions: 
    
    echo Attempting to grant %username% local administrative permissions...
    net localgroup administrators %username% /add
    echo:
    echo Operation completed.
    echo:

    echo [Press any key to return to the main menu]
    pause >nul
    call :mainMenu 2 2


:openAppsInFolder
    title Batch Utilities - Open Applications In Folder
    cls

    echo Open Modes:
    echo 1 - Opens the files one by one, waiting for the first application to close before opening the next.
    echo 2 - Opens all the files at once.
    echo:
    set /p openMode=Please select a open mode: 

    set /p targetDir=Please enter a directory path: 

    set firstChar=%targetDir:~0,1%
    set lastChar=%targetDir:~-1%

    set firstChar=%firstChar:"=+%
    set lastChar=%lastChar:"=+%

    if not "%firstChar%"=="+" set targetDir=^"%targetDir%
    if not "%lastChar%"=="+" set targetDir=%targetDir%"

    if exist %targetDir% (
        pushd %targetDir% 

        echo:
        echo Moved into directory %targetDir%
        echo:
    ) else (
        echo That directory doesn't exist! Make sure you entered the path correctly.
        echo:
    )

    if exist %targetDir% if %openMode% equ 1 FOR %%F IN (*.*) DO timeout 2 > nul & echo Opening file %%F & start "" /wait "%%F"

    if exist %targetDir% if %openMode% equ 2 FOR %%F IN (*.*) DO timeout 2 > nul & echo Opening file %%F & start "" "%%F"

    if exist %targetDir% echo: & echo Finished opening all files. & echo: & popd

    echo [Press any key to return to the main menu]
    pause >nul
    call :mainMenu 3 1


:revealWiFiPasswords
    title Batch Utilities - Reveal Wi-Fi Passwords
    cls
    setlocal enabledelayedexpansion

    set KEYCONTENT=Key Content
    set ALL_PROFILES=All User Profile

    echo Searching for Wi-Fi passwords...
    echo:

    call :getWiFiProfiles results

    :nextWiFiProfile
        for /f "tokens=1* delims=," %%a in ("%results%") do (
            call :getWiFiProfileKey "%%a" key
            if "!key!" NEQ "" (
                echo WiFi Name: [%%a] Password: [!key!]
            )
            set results=%%b
        )
        if "%results%" NEQ "" goto nextWiFiProfile

    endlocal
    echo:
    echo [Press any key to return to the main menu]
    pause >nul
    call :mainMenu 3 2


:getBitLockerKey
    title Batch Utilities - Get BitLocker Key
    cls
    
    set tempFile=%TEMP%\BitLockerKey.txt

    if exist %tempFile% del %tempFile%

    PowerShell -command "Get-BitLockerVolume -MountPoint C:| Select-Object -ExpandProperty KeyProtector | Where-Object KeyProtectorType -eq RecoveryPassword | Select-Object KeyProtectorID, RecoveryPassword | Format-List" > %tempFile%

    for /f "usebackq skip=1 tokens=2 delims=: " %%i in (%tempFile%) do set key=%%i

    echo:
    echo Your BitLocker Key Is: %key%
    echo:

    set /p doCopy="Would you like copy the key to your clipboard? (y/n) "

    if %doCopy%==y (
        echo | set /p=%key%| clip
        echo Key successfully copied!
    )

    if exist %tempFile% del %tempFile%

    echo:
    echo [Press any key to close this window]
    pause >nul
    call :mainMenu 5 1


:getWindowsActivationKey
    title Batch Utilities - Get Windows Activation Key
    cls
    
    for /f "tokens=2 delims==" %%i in ('"wmic path softwarelicensingservice get oa3xoriginalproductkey /format:list"') do set key=%%i

    echo:
    echo Your Windows Licensing Key Is: %key%
    echo:

    set /p doCopy="Would you like copy the key to your clipboard? (y/n) "

    if %doCopy%==y (
        echo | set /p=%key%| clip
        echo Key successfully copied!
        echo:
    )

    echo:
    echo [Press any key to close this window]
    pause >nul
    call :mainMenu 4 1


:quoteUnquotedServicePaths
    title Batch Utilities - Quote Unquoted Service Paths
    cls
    
    set tempFile=%TEMP%\UnquottedPaths.txt

    if exist %tempFile% del %tempFile%

    wmic service get name,pathname,startmode,displayname /format:csv | findstr /i auto | findstr /i /v "C:\Windows\\" | findstr /i /v ^""" > %tempFile%

    for /f "tokens=2,4 delims=," %%a in (%tempFile%) do (
        reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\%%a" /v ImagePath /t REG_EXPAND_SZ /d "\"%%b"\" /f >NUL 2>&1
        echo Fixed path for service %%a,
        echo quoted the following path: %%b
        echo:
    )

    if exist %tempFile% del %tempFile%

    echo Successfully fixed unquotted paths!

    echo:
    echo [Press any key to close this window]
    pause >nul
    call :mainMenu 4 2


@REM 
@REM  Utility Functions
@REM 

:exitUtility
    title Batch Utilites - ^Exit Utility
    cls

    if %deviceNeedsRestart%==true (
        echo Your device needs to restart in order for some changes to take effect.
        set /p wantsRestartNow=Would you like to restart now? ^(y/n^) 
    )
    if /I %wantsRestartNow%==y goto restartNow
    exit


:restartNow
    title Image Utilites - Restart now
    cls

    echo Your device will retart in 5 seconds...
    shutdown.exe /r /t 5
    pause >nul


:getWiFiProfileKey <1=profile-name> <2=out-profile-key>
    setlocal

    set result=

    for /f "usebackq tokens=2 delims=:" %%a in (`netsh wlan show profile name^="%~1" key^=clear ^| findstr /C:"%KEYCONTENT%"`) do (
        set result=%%a
        set result=!result:~1!
    )
    (
        endlocal
        set %2=%result%
    )

    goto :eof


:getWiFiProfiles <1=result-variable>
    setlocal

    set result=
   
    for /f "usebackq tokens=2 delims=:" %%a in (`netsh wlan show profiles ^| findstr /C:"%ALL_PROFILES%"`) do (
        set val=%%a
        set val=!val:~1!

        set result=%!val!,!result!
    )
    (
        endlocal
        set %1=%result:~0,-1%
    )

    goto :eof


:toUppercase
    if not defined %~1 exit /b
    for %%a in ("a=A" "b=B" "c=C" "d=D" "e=E" "f=F" "g=G" "h=H" "i=I" "j=J" "k=K" "l=L" "m=M" "n=N" "o=O" "p=P" "q=Q" "r=R" "s=S" "t=T" "u=U" "v=V" "w=W" "x=X" "y=Y" "z=Z" "ä=Ä" "ö=Ö" "ü=Ü") do (
        call set %~1=%%%~1:%%~a%%
    )
    goto :eof


:strlen <resultVar> <stringVar>
    (   
        setlocal EnableDelayedExpansion
        (set^ tmp=!%~2!)
        if defined tmp (
            set "len=1"
            for %%P in (4096 2048 1024 512 256 128 64 32 16 8 4 2 1) do (
                if "!tmp:~%%P,1!" NEQ "" ( 
                    set /a "len+=%%P"
                    set "tmp=!tmp:~%%P!"
                )
            )
        ) ELSE (
            set len=0
        )
    )
    ( 
        endlocal
        set "%~1=%len%"
        exit /b
    )
