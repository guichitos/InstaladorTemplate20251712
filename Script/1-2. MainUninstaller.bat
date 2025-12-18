@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem ===========================================================
rem === UNIVERSAL OFFICE TEMPLATE UNINSTALLER (v1.2) ==========
rem ===========================================================


rem === Mode and logging configuration ========================
rem true  = verbose mode with console messages, logging, and final pause.
rem false = silent mode (no console output or pause).
set "IsDesignModeEnabled=true"

if /I not "%IsDesignModeEnabled%"=="true" (
    title TEMPLATE INSTALLER
    echo Removing custom templates and restoring the Microsoft Office default settings
    echo Executing...
)

rem If wrapper passed the launcher directory (payload), use it.
if not "%~1"=="" (
    set "LauncherDirectory=%~1"
) else (
    rem Fallback: assume current directory is the launcher/payload location
    set "LauncherDirectory=%CD%"
)

rem ScriptDirectory = real location of this uninstaller (in AppData)
set "ScriptDirectory=%~dp0"

if /I "%IsDesignModeEnabled%"=="true" (
    call :DebugTrace "[INFO] Script directory (uninstaller) resolved to: %ScriptDirectory%"
    call :DebugTrace "[INFO] Launcher/payload directory resolved to: %LauncherDirectory%"
)
call :DebugTrace "[FLAG] Script initialization started."

set "UserLaunchDirectory=%CD%"

rem Usamos la carpeta del launcher para resolver la payload real
call :ResolveBaseDirectory "%LauncherDirectory%" BaseDirectoryPath
call :ResolveBaseDirectory "%UserLaunchDirectory%" LaunchDirectoryPath

set "BaseHasPayload=0"
set "LaunchHasPayload=0"

call :HasTemplatePayload "%BaseDirectoryPath%" BaseHasPayload
if /I not "%LaunchDirectoryPath%"=="%BaseDirectoryPath%" call :HasTemplatePayload "%LaunchDirectoryPath%" LaunchHasPayload

if "!BaseHasPayload!"=="0" if "!LaunchHasPayload!"=="1" (
    set "BaseDirectoryPath=!LaunchDirectoryPath!"
    if /I "%IsDesignModeEnabled%"=="true" call :DebugTrace "[INFO] No payload found at primary path; using launch directory payload location instead."
)

rem OJO: aquí ya volvemos a usar ScriptDirectory (AppData) para libs y logs
set "LibraryDirectoryPath=%ScriptDirectory%lib"
set "LogsDirectoryPath=%ScriptDirectory%logs"
set "LogFilePath=%LogsDirectoryPath%\uninstall_log.txt"

call :LocateLibrary "1-2. ResolveAppProperties.bat" OfficeTemplateLib "%ScriptDirectory%" "%BaseDirectoryPath%" "%LauncherDirectory%" "%LaunchDirectoryPath%"
call :LocateLibrary "1-2. TemplatePathResolver.bat" TemplatePathLib "%ScriptDirectory%" "%BaseDirectoryPath%" "%LauncherDirectory%" "%LaunchDirectoryPath%"

if not defined OfficeTemplateLib (
    echo [ERROR] Shared library not found: "1-2. ResolveAppProperties.bat"
    exit /b 1
)
if not defined TemplatePathLib (
    echo [ERROR] Shared library not found: "1-2. TemplatePathResolver.bat"
    exit /b 1
)
if not exist "%TemplatePathLib%" (
    echo [ERROR] Shared library not found: "%TemplatePathLib%"
    exit /b 1
)

if /I "%IsDesignModeEnabled%"=="true" (
    if not exist "%LogsDirectoryPath%" mkdir "%LogsDirectoryPath%"
    echo [%DATE% %TIME%] --- START UNINSTALL --- > "%LogFilePath%"
    title OFFICE TEMPLATE UNINSTALLER - DEBUG MODE
    echo [DEBUG] Running from payload base: %BaseDirectoryPath%
)

call :DebugTrace "[FLAG] Target paths and logging configured."


call "%TemplatePathLib%" :ResolveDefaultTemplatePaths "%IsDesignModeEnabled%"
set "OPEN_WORD_FLAG=0"
set "OPEN_PPT_FLAG=0"
set "OPEN_EXCEL_FLAG=0"
set "OPENED_TEMPLATE_FOLDERS=;"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Closing Office applications before uninstall...
    call :CloseOfficeApps
) else (
    call :CloseOfficeApps >nul 2>&1
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [TARGET CLEANUP PATHS]
    echo ----------------------------
    echo WORD PATH:       %WORD_PATH%
    echo POWERPOINT PATH: %PPT_PATH%
    echo EXCEL PATH:      %EXCEL_PATH%
    echo THEMES PATH:     !THEME_PATH!
    echo ----------------------------
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo [INFO] --- TARGET CLEANUP PATHS --- >> "%LogFilePath%"
    echo Word path: %WORD_PATH% >> "%LogFilePath%"
    echo PowerPoint path: %PPT_PATH% >> "%LogFilePath%"
    echo Excel path: %EXCEL_PATH% >> "%LogFilePath%"
    echo Themes path: !THEME_PATH! >> "%LogFilePath%"
    echo ---------------------------- >> "%LogFilePath%"
)

rem === Detect custom template folders for optional cleanup ===
set "WORD_CUSTOM_TEMPLATE_PATH="
set "PPT_CUSTOM_TEMPLATE_PATH="
set "EXCEL_CUSTOM_TEMPLATE_PATH="
set "DEFAULT_CUSTOM_TEMPLATE_DIR="
call "%TemplatePathLib%" :DetectCustomTemplatePaths "%LogFilePath%" "%IsDesignModeEnabled%"

if /I "%IsDesignModeEnabled%"=="true" (
    call :DebugTrace "[DEBUG] Custom template cleanup targets:"
    if defined WORD_CUSTOM_TEMPLATE_PATH (
        call :DebugTrace "        Word: !WORD_CUSTOM_TEMPLATE_PATH!"
    ) else (
        call :DebugTrace "        Word: <not detected>"
    )
    if defined PPT_CUSTOM_TEMPLATE_PATH (
        call :DebugTrace "        PowerPoint: !PPT_CUSTOM_TEMPLATE_PATH!"
    ) else (
        call :DebugTrace "        PowerPoint: <not detected>"
    )
    if defined EXCEL_CUSTOM_TEMPLATE_PATH (
        call :DebugTrace "        Excel: !EXCEL_CUSTOM_TEMPLATE_PATH!"
    ) else (
        call :DebugTrace "        Excel: <not detected>"
    )
)

call :DebugTrace "[FLAG] Built-in template definitions resolved."

rem === Define files ==========================================
set "WordFile=%WORD_PATH%\Normal.dotx"
set "WordMacroFile=%WORD_PATH%\Normal.dotm"
set "WordEmailFile=%WORD_PATH%\NormalEmail.dotx"
set "WordEmailMacroFile=%WORD_PATH%\NormalEmail.dotm"

set "PptFile=%PPT_PATH%\Blank.potx"
set "PptMacroFile=%PPT_PATH%\Blank.potm"

set "ExcelBookFile=%EXCEL_PATH%\Book.xltx"
set "ExcelBookMacroFile=%EXCEL_PATH%\Book.xltm"

set "ExcelSheetFile=%EXCEL_PATH%\Sheet.xltx"
set "ExcelSheetMacroFile=%EXCEL_PATH%\Sheet.xltm"

rem === Helper routine: delete templates =======================
if exist "%WordFile%" set "OPEN_WORD_FLAG=1"
call :ProcessFile "Word (.dotx)" "%WordFile%" "%LogFilePath%"
if exist "%WordMacroFile%" set "OPEN_WORD_FLAG=1"
call :ProcessFile "Word (.dotm)" "%WordMacroFile%" "%LogFilePath%"
if exist "%WordEmailFile%" set "OPEN_WORD_FLAG=1"
call :ProcessFile "Word Email (.dotx)" "%WordEmailFile%" "%LogFilePath%"
if exist "%WordEmailMacroFile%" set "OPEN_WORD_FLAG=1"
call :ProcessFile "Word Email (.dotm)" "%WordEmailMacroFile%" "%LogFilePath%"
if exist "%PptFile%" set "OPEN_PPT_FLAG=1"
call :ProcessFile "PowerPoint (.potx)" "%PptFile%" "%LogFilePath%"
if exist "%PptMacroFile%" set "OPEN_PPT_FLAG=1"
call :ProcessFile "PowerPoint (.potm)" "%PptMacroFile%" "%LogFilePath%"
if exist "%ExcelBookFile%" set "OPEN_EXCEL_FLAG=1"
call :ProcessFile "Excel Book (.xltx)" "%ExcelBookFile%" "%LogFilePath%"
if exist "%ExcelBookMacroFile%" set "OPEN_EXCEL_FLAG=1"
call :ProcessFile "Excel Book (.xltm)" "%ExcelBookMacroFile%" "%LogFilePath%"
if exist "%ExcelSheetFile%" set "OPEN_EXCEL_FLAG=1"
call :ProcessFile "Excel Sheet (.xltx)" "%ExcelSheetFile%" "%LogFilePath%"
if exist "%ExcelSheetMacroFile%" set "OPEN_EXCEL_FLAG=1"
call :ProcessFile "Excel Sheet (.xltm)" "%ExcelSheetMacroFile%" "%LogFilePath%"

set "THEME_PAYLOAD_TRACK="
if defined THEME_PATH (
    for %%F in ("%BaseDirectoryPath%*.thmx") do (
        if exist "%%~fF" set "THEME_PAYLOAD_TRACK=!THEME_PAYLOAD_TRACK!;%%~nxF;"
    )
)

rem Clean Document Themes by comparing against installer payloads and only delete matches
if defined THEME_PATH if exist "!THEME_PATH!" (
    for /f "delims=" %%T in ('dir /A-D /B "!THEME_PATH!\*.thmx" 2^>nul') do (
        set "THEME_HAS_PAYLOAD=0"
        if defined THEME_PAYLOAD_TRACK (
            echo !THEME_PAYLOAD_TRACK! | find /I ";%%~nT%%~xT;" >nul && set "THEME_HAS_PAYLOAD=1"
        )

        if "!THEME_HAS_PAYLOAD!"=="1" (
            set "CurrentThemeFile=!THEME_PATH!\%%~nxT"
            call :ProcessFile "Office Theme (%%~nxT)" "!CurrentThemeFile!" "%LogFilePath%"
        ) else (
            if /I "%IsDesignModeEnabled%"=="true" call :DebugTrace "        [SKIP] Preserved Office Theme (%%~nxT) with no installer match."
        )
    )
)

call :DebugTrace "[FLAG] Starting custom template cleanup."

call :RemoveCustomTemplates "%BaseDirectoryPath%" "%LogFilePath%" "%IsDesignModeEnabled%" "!WORD_CUSTOM_TEMPLATE_PATH!" "!PPT_CUSTOM_TEMPLATE_PATH!" "!EXCEL_CUSTOM_TEMPLATE_PATH!"

echo.
call :DebugTrace "[FLAG] Repairing template MRU entries via helper script."

call "%ScriptDirectory%1-2. Repair Office template MRU.bat"

call :DebugTrace "[FLAG] Opening affected folders and relaunching Office apps."
call :DebugTrace "[DEBUG] Launch flags before opening folders -> Word:%OPEN_WORD_FLAG% PPT:%OPEN_PPT_FLAG% Excel:%OPEN_EXCEL_FLAG%"
call :OpenTemplateFolder "%WORD_PATH%" "%IsDesignModeEnabled%" "Word template folder" ""
call :OpenTemplateFolder "%PPT_PATH%" "%IsDesignModeEnabled%" "PowerPoint template folder" ""
call :OpenTemplateFolder "%EXCEL_PATH%" "%IsDesignModeEnabled%" "Excel template folder" ""
if defined THEME_PATH call :OpenTemplateFolder "!THEME_PATH!" "%IsDesignModeEnabled%" "Document Themes folder" ""
if defined WORD_CUSTOM_TEMPLATE_PATH call :OpenTemplateFolder "!WORD_CUSTOM_TEMPLATE_PATH!" "%IsDesignModeEnabled%" "Custom Word templates" ""
if defined PPT_CUSTOM_TEMPLATE_PATH call :OpenTemplateFolder "!PPT_CUSTOM_TEMPLATE_PATH!" "%IsDesignModeEnabled%" "Custom PowerPoint templates" ""
if defined EXCEL_CUSTOM_TEMPLATE_PATH call :OpenTemplateFolder "!EXCEL_CUSTOM_TEMPLATE_PATH!" "%IsDesignModeEnabled%" "Custom Excel templates" ""

call :DebugTrace "[DEBUG] Launch flags after cleanup -> Word:%OPEN_WORD_FLAG% PPT:%OPEN_PPT_FLAG% Excel:%OPEN_EXCEL_FLAG%"

call :DebugTrace "[FLAG] Finalizing uninstaller."

call :Finalize "%LogFilePath%"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    pause
)

endlocal
exit /b

rem Base dir resolver keeps template source tied to the unpacked executable location
:ResolveBaseDirectory
setlocal
set "RBD_INPUT=%~1"
set "RBD_OUTPUT_VAR=%~2"

if "%RBD_INPUT:~-1%" NEQ "\\" set "RBD_INPUT=%RBD_INPUT%\\"

set "RBD_FOUND="
for %%D in ("%RBD_INPUT%" "%RBD_INPUT%payload\\" "%RBD_INPUT%templates\\" "%RBD_INPUT%extracted\\") do (
    for %%F in ("%%~D*.dot*" "%%~D*.pot*" "%%~D*.xlt*" "%%~D*.thmx") do (
        if exist "%%~fF" set "RBD_FOUND=%%~D"
    )
    if defined RBD_FOUND goto :RBD_Found
)

:RBD_Found
if not defined RBD_FOUND set "RBD_FOUND=%RBD_INPUT%"

endlocal & set "%RBD_OUTPUT_VAR%=%RBD_FOUND%"
exit /b 0

:HasTemplatePayload
setlocal enabledelayedexpansion
set "HP_PATH=%~1"
set "HP_OUT=%~2"
set "HP_FOUND=0"

if not defined HP_PATH goto :HasTemplatePayloadEnd
if "!HP_PATH:~-1!" NEQ "\\" set "HP_PATH=!HP_PATH!\\"

for %%F in ("!HP_PATH!*.dot*" "!HP_PATH!*.pot*" "!HP_PATH!*.xlt*" "!HP_PATH!*.thmx") do (
    if exist "%%~fF" set "HP_FOUND=1"
)

:HasTemplatePayloadEnd
set "HP_RESULT=!HP_FOUND!"
endlocal & if not "%HP_OUT%"=="" set "%HP_OUT%=%HP_RESULT%"
exit /b 0

:LocateLibrary
setlocal
set "LL_NAME=%~1"
set "LL_OUT_VAR=%~2"
set "LL_OUT_PATH="

for %%D in ("%~3" "%~4" "%~5" "%~6") do (
    set "LL_DIR=%%~D"
    if defined LL_DIR (
        call :NormalizePath LL_DIR
        for %%P in ("!LL_DIR!" "!LL_DIR!Script\") do (
            if exist "%%~fP" (if "%%~aP" GEQ "d" (
                if exist "%%~fP%LL_NAME%" set "LL_OUT_PATH=%%~fP%LL_NAME%"
            ))
        )
    )
    if defined LL_OUT_PATH goto :LocateLibraryFound
)

:LocateLibraryFound
endlocal & if defined LL_OUT_PATH set "%LL_OUT_VAR%=%LL_OUT_PATH%"
exit /b 0


:RemoveCustomTemplates
setlocal enabledelayedexpansion
set "BASE_DIR=%~1"
set "LOG_FILE=%~2"
set "DESIGN_MODE=%~3"
set "WORD_DIR=%~4"
set "PPT_DIR=%~5"
set "EXCEL_DIR=%~6"

set "RCT_WORD_FLAG=!OPEN_WORD_FLAG!"
set "RCT_PPT_FLAG=!OPEN_PPT_FLAG!"
set "RCT_EXCEL_FLAG=!OPEN_EXCEL_FLAG!"

if not defined BASE_DIR exit /b 0
if "!BASE_DIR:~-1!" NEQ "\" set "BASE_DIR=!BASE_DIR!\"

if /I "!DESIGN_MODE!"=="true" (
    call :DebugTrace "        [DEBUG] RemoveCustomTemplates invoked with:"
    call :DebugTrace "        Base dir: !BASE_DIR!"
    call :DebugTrace "        Word dir: !WORD_DIR!"
    call :DebugTrace "        PPT dir: !PPT_DIR!"
    call :DebugTrace "        Excel dir: !EXCEL_DIR!"
)

set /a CUSTOM_REMOVED_COUNT=0
set /a CUSTOM_SKIP_COUNT=0
set /a CUSTOM_ERROR_COUNT=0
set /a CUSTOM_TOTAL_CANDIDATES=0
set "CUSTOM_GENERIC_SKIP_LIST=Normal.dotx NormalEmail.dotx Blank.potx Book.xltx Normal.dotm NormalEmail.dotm Blank.potm Book.xltm Sheet.xltx Sheet.xltm"

call :CleanCustomTemplateFiles "!WORD_DIR!" ".dotx .dotm" "!BASE_DIR!" "%LOG_FILE%" "!DESIGN_MODE!" "Word custom templates" "WORD"
call :CleanCustomTemplateFiles "!PPT_DIR!" ".potx .potm" "!BASE_DIR!" "%LOG_FILE%" "!DESIGN_MODE!" "PowerPoint custom templates" "PPT"
call :CleanCustomTemplateFiles "!EXCEL_DIR!" ".xltx .xltm" "!BASE_DIR!" "%LOG_FILE%" "!DESIGN_MODE!" "Excel custom templates" "EXCEL"

if /I "!DESIGN_MODE!"=="true" (
    call :DebugTrace "[INFO] Custom template cleanup summary: Removed !CUSTOM_REMOVED_COUNT!, skipped !CUSTOM_SKIP_COUNT!, errors !CUSTOM_ERROR_COUNT!."
)

endlocal & (
    set "OPEN_WORD_FLAG=%RCT_WORD_FLAG%"
    set "OPEN_PPT_FLAG=%RCT_PPT_FLAG%"
    set "OPEN_EXCEL_FLAG=%RCT_EXCEL_FLAG%"
)
exit /b 0

:CleanCustomTemplateFiles
set "CCF_TARGET_DIR=%~1"
set "CCF_EXT_LIST=%~2"
set "CCF_BASE_DIR=%~3"
call :NormalizePath CCF_BASE_DIR
set "CCF_LOG_FILE=%~4"
set "CCF_DESIGN_MODE=%~5"
set "CCF_LABEL=%~6"
set "CCF_APP_KEY=%~7"

if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[FLAG] CleanCustomTemplateFiles invoked with parameters %*"

if not defined CCF_TARGET_DIR exit /b 0
if "!CCF_TARGET_DIR!"=="" exit /b 0
if not exist "!CCF_TARGET_DIR!" (
    if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[INFO] !CCF_LABEL! not found at '!CCF_TARGET_DIR!' - skipping."
    exit /b 0
)

set "CCF_TOP_LEVEL_COUNT=0"
set "CCF_RECURSIVE_COUNT=0"

for /f %%C in ('dir /A /B "!CCF_TARGET_DIR!" 2^>nul ^| find /C /V ""') do set "CCF_TOP_LEVEL_COUNT=%%C"
for /f %%C in ('dir /A /B /S "!CCF_TARGET_DIR!" 2^>nul ^| find /C /V ""') do set "CCF_RECURSIVE_COUNT=%%C"

    set "CCF_DIR_FILE_COUNT=0"
    set "CCF_DIR_REMOVED=0"
    set "CCF_DIR_SKIPPED=0"
    set "CCF_DIR_ERRORS=0"

    for %%E in (!CCF_EXT_LIST!) do (
        set "CCF_EXT_COUNT=0"
        set "CCF_EXT_REMOVED=0"
        set "CCF_EXT_SKIPPED=0"
        set "CCF_EXT_ERRORS=0"
        for /f %%C in ('dir /A-D /B /S "!CCF_TARGET_DIR!\*%%~E" 2^>nul ^| find /C /V ""') do set "CCF_EXT_COUNT=%%C"

        for /f "delims=" %%F in ('dir /A-D /B /S "!CCF_TARGET_DIR!\*%%~E" 2^>nul') do (
            if exist "%%~fF" (
                set "CCF_FILE=%%~nxF"
                set /a CUSTOM_TOTAL_CANDIDATES+=1
                set /a CCF_DIR_FILE_COUNT+=1
                set "CCF_SKIP_GENERIC=0"
                for %%G in (!CUSTOM_GENERIC_SKIP_LIST!) do (
                    if /I "!CCF_FILE!"=="%%~G" set "CCF_SKIP_GENERIC=1"
                )

                if "!CCF_SKIP_GENERIC!"=="1" (
                    rem === Preserve generic system templates ===
                    set /a CUSTOM_SKIP_COUNT+=1
                    set /a CCF_DIR_SKIPPED+=1
                    set /a CCF_EXT_SKIPPED+=1
                    if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[SKIP] Preserved generic template !CCF_FILE! in !CCF_LABEL!."
                ) else (
                    set "CCF_INSTALLER_FILE=!CCF_BASE_DIR!!CCF_FILE!"

                    rem === Files that ARE part of installer payload MUST be deleted ===
                    if exist "!CCF_INSTALLER_FILE!" (
                        set "CCF_DELETE_REASON=installer payload match"
                        del /F /Q "%%~fF" >nul 2>&1
                        if exist "%%~fF" (
                            set /a CUSTOM_ERROR_COUNT+=1
                            set /a CCF_DIR_ERRORS+=1
                            set /a CCF_EXT_ERRORS+=1
                            if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[ERROR] Could not delete !CCF_FILE! from !CCF_LABEL!."
                        ) else (
                            set /a CUSTOM_REMOVED_COUNT+=1
                            set /a CCF_DIR_REMOVED+=1
                            set /a CCF_EXT_REMOVED+=1
                            call :MarkAppAsInvolved "!CCF_APP_KEY!"
                            if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[OK] Deleted !CCF_FILE! from !CCF_LABEL! (!CCF_DELETE_REASON!)."
                        )
                    ) else (
                        rem === Files NOT in installer payload MUST be PRESERVED ===
                        set /a CUSTOM_SKIP_COUNT+=1
                        set /a CCF_DIR_SKIPPED+=1
                        set /a CCF_EXT_SKIPPED+=1
                        if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[SKIP] Preserved !CCF_FILE! in !CCF_LABEL! (user/custom file)."
                    )
                )

            ) else (
                if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[WARN] Candidate vanished before delete: %%~fF"
            )
        )
    )

    set "CCF_REMOVED_DIRS=0"
    for /f "delims=" %%D in ('dir /AD /B /S "!CCF_TARGET_DIR!" ^| sort /R') do (
        rd "%%~fD" 2>nul && set /a CCF_REMOVED_DIRS+=1
    )

    if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[INFO] !CCF_LABEL!: removed !CCF_REMOVED_DIRS! empty directories."

set "CCF_TARGET_DIR="
set "CCF_EXT_LIST="
set "CCF_BASE_DIR="
set "CCF_LOG_FILE="
set "CCF_DESIGN_MODE="
set "CCF_LABEL="
set "CCF_FILE="
set "CCF_INSTALLER_FILE="
set "CCF_SKIP_GENERIC="
set "CCF_TOP_LEVEL_COUNT="
set "CCF_RECURSIVE_COUNT="
set "CCF_EXT_COUNT="
set "CCF_DIR_REMOVED="
set "CCF_DIR_SKIPPED="
set "CCF_DIR_ERRORS="
set "CCF_EXT_REMOVED="
set "CCF_EXT_SKIPPED="
set "CCF_EXT_ERRORS="
set "CCF_REMOVED_DIRS="
endlocal & (
    set "OPEN_WORD_FLAG=%RCT_WORD_FLAG%"
    set "OPEN_PPT_FLAG=%RCT_PPT_FLAG%"
    set "OPEN_EXCEL_FLAG=%RCT_EXCEL_FLAG%"
)
exit /b 0

:MarkAppAsInvolved
set "MAI_APP=%~1"
if /I "%MAI_APP%"=="WORD" set "RCT_WORD_FLAG=1"
if /I "%MAI_APP%"=="PPT" set "RCT_PPT_FLAG=1"
if /I "%MAI_APP%"=="EXCEL" set "RCT_EXCEL_FLAG=1"
exit /b 0


:OpenTemplateFolder
set "TARGET_PATH=%~1"
set "DESIGN_MODE=%~2"
set "FOLDER_LABEL=%~3"
set "SELECT_PATH=%~4"

if "%TARGET_PATH%"=="" exit /b
if not exist "%TARGET_PATH%" exit /b
if "%FOLDER_LABEL%"=="" set "FOLDER_LABEL=template folder"
if not defined OPENED_TEMPLATE_FOLDERS set "OPENED_TEMPLATE_FOLDERS=;"
set "TOKEN=;%TARGET_PATH%;"
if "!OPENED_TEMPLATE_FOLDERS:%TOKEN%=!"=="!OPENED_TEMPLATE_FOLDERS!" (
    if /I "%DESIGN_MODE%"=="true" (
        if defined SELECT_PATH (
            echo [ACTION] Opening !FOLDER_LABEL! and selecting: !SELECT_PATH!
        ) else (
            echo [ACTION] Opening !FOLDER_LABEL!: !TARGET_PATH!
        )
    )
    if defined SELECT_PATH (
        if exist "%SELECT_PATH%" (
            start "" explorer /select,"!SELECT_PATH!"
        ) else (
            start "" explorer "!TARGET_PATH!"
        )
    ) else (
        start "" explorer "!TARGET_PATH!"
    )
    set "OPENED_TEMPLATE_FOLDERS=!OPENED_TEMPLATE_FOLDERS!!TOKEN!"
)
exit /b

:ProcessFile
rem ===========================================================
rem Args: AppName, TargetFile, LogFile
rem ===========================================================
setlocal enabledelayedexpansion
set "AppName=%~1"
set "TargetFile=%~2"
set "LogFile=%~3"

rem === Step 1: Always delete current template (factory reset) ===
if exist "%TargetFile%" (
    del /F /Q "%TargetFile%" >nul 2>&1
    if exist "%TargetFile%" (
        set "Message=[%AppName%] [ERROR] Could not delete %TargetFile%. File may be locked."
    ) else (
        set "Message=[%AppName%] [OK] Deleted %TargetFile%"
    )
) else (
    set "Message=[%AppName%] [INFO] %TargetFile% not found."
)

rem === Step 2: Emit verbose trace if enabled ===
if /I "%IsDesignModeEnabled%"=="true" (
    call :DebugTrace "        !Message!"
    if defined LogFile (>>"%LogFile%" echo [%DATE% %TIME%] !Message!)
)

endlocal
exit /b 0

:Finalize
setlocal enabledelayedexpansion
if /I not "%IsDesignModeEnabled%"=="true" (
    endlocal
    exit /b 0
)

set "ResolvedLogPath=%~1"

>>"%~1" echo [%DATE% %TIME%] --- UNINSTALL COMPLETED ---

endlocal
exit /b 0

:DebugTrace
if /I not "%IsDesignModeEnabled%"=="true" exit /b 0
setlocal enabledelayedexpansion
set "DebugMessage=%~1"
if defined DebugMessage (
    echo !DebugMessage!
) else (
    echo.
)
endlocal
exit /b 0

:Log
call "%OfficeTemplateLib%" :Log %*
exit /b %errorlevel%

:NormalizePath
setlocal
set "NP_VAR=%~1"
set "NP_VAL=!%NP_VAR%!"

rem Remove trailing backslashes
:NP_LOOP
if "!NP_VAL!"=="\" goto NP_END
if "!NP_VAL:~-1!"=="\" (
    set "NP_VAL=!NP_VAL:~0,-1!"
    goto NP_LOOP
)

:NP_END
rem Add exactly one backslash
set "NP_VAL=!NP_VAL!\"
endlocal & set "%~1=%NP_VAL%"
exit /b 0

:CloseOfficeApps
call :DebugTrace "[DEBUG] Entering Closing Office applications with args: %*"
taskkill /IM WINWORD.EXE /F >nul 2>&1
taskkill /IM POWERPNT.EXE /F >nul 2>&1
taskkill /IM EXCEL.EXE /F >nul 2>&1
taskkill /IM OUTLOOK.EXE /F >nul 2>&1
call :DebugTrace "[DEBUG] Exiting Closing Office applications..."
exit /b 0