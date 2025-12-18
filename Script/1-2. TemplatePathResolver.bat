@echo off
rem ============================================================
rem ===      1-2. TemplatePathResolver.bat (biblioteca)      ===
rem ===   Resolución de rutas de plantillas y temas Office   ===
rem ===  Uso: call "%~dp0...TemplatePathResolver.bat" :Label ... ===
rem ============================================================

set "__TPR_ENTRY=%~1"
if not defined __TPR_ENTRY goto :EOF

rem Permitir sintaxis con ":Label"
if "%__TPR_ENTRY:~0,1%"==":" set "__TPR_ENTRY=%__TPR_ENTRY:~1%"
shift

if not defined __TPR_ENTRY goto :EOF

goto :%__TPR_ENTRY%

:ResolveDefaultTemplatePaths
rem Args: [DesignModeFlag]
setlocal EnableDelayedExpansion
set "TPR_DESIGN_MODE=%~1"
if not defined TPR_DESIGN_MODE set "TPR_DESIGN_MODE=%IsDesignModeEnabled%"

set "WORD_PATH=%APPDATA%\Microsoft\Templates"
set "PPT_PATH=%APPDATA%\Microsoft\Templates"
set "EXCEL_PATH=%APPDATA%\Microsoft\Excel\XLSTART"

set "APPDATA_EXPANDED="
for /f "delims=" %%T in ('powershell -NoLogo -Command "$app=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name AppData -ErrorAction SilentlyContinue).AppData; if ($app) {[Environment]::ExpandEnvironmentVariables($app)}"') do set "APPDATA_EXPANDED=%%T"
if not defined APPDATA_EXPANDED set "APPDATA_EXPANDED=%APPDATA%"
if defined APPDATA_EXPANDED set "THEME_PATH=!APPDATA_EXPANDED!\Microsoft\Templates\Document Themes"

endlocal & (
    set "WORD_PATH=%WORD_PATH%"
    set "PPT_PATH=%PPT_PATH%"
    set "EXCEL_PATH=%EXCEL_PATH%"
    if defined THEME_PATH set "THEME_PATH=%THEME_PATH%"
)
exit /b 0

:DetectCustomTemplatePaths
rem Args: LogFile DesignModeFlag
setlocal EnableDelayedExpansion
set "DCTP_LOG_FILE=%~1"
set "DCTP_DESIGN_MODE=%~2"
if not defined DCTP_DESIGN_MODE set "DCTP_DESIGN_MODE=%IsDesignModeEnabled%"

set "WORD_CUSTOM_TEMPLATE_PATH="
set "PPT_CUSTOM_TEMPLATE_PATH="
set "EXCEL_CUSTOM_TEMPLATE_PATH="
set "DEFAULT_CUSTOM_TEMPLATE_DIR="
set "DEFAULT_CUSTOM_DIR_STATUS=unknown"
set "DCTP_DOCUMENTS_PATH="
set "DCTP_OFFICE_VERSIONS=16.0 15.0 14.0 12.0"

for /f "delims=" %%D in ('powershell -NoLogo -Command "$path=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name Personal -ErrorAction SilentlyContinue).Personal; if ($path) {[Environment]::ExpandEnvironmentVariables($path)}"') do set "DCTP_DOCUMENTS_PATH=%%D"

if defined DCTP_DOCUMENTS_PATH (
    if "!DCTP_DOCUMENTS_PATH:~-1!"=="\\" set "DCTP_DOCUMENTS_PATH=!DCTP_DOCUMENTS_PATH:~0,-1!"
    set "DEFAULT_CUSTOM_TEMPLATE_DIR=!DCTP_DOCUMENTS_PATH!\Custom Templates"
) else (
    set "DEFAULT_CUSTOM_TEMPLATE_DIR=%USERPROFILE%\Documents\Custom Templates"
)

if not defined DEFAULT_CUSTOM_TEMPLATE_DIR set "DEFAULT_CUSTOM_TEMPLATE_DIR=%USERPROFILE%\Documents\Custom Templates"

if defined DEFAULT_CUSTOM_TEMPLATE_DIR (
    if exist "!DEFAULT_CUSTOM_TEMPLATE_DIR!" (
        set "DEFAULT_CUSTOM_DIR_STATUS=exists"
    ) else (
        mkdir "!DEFAULT_CUSTOM_TEMPLATE_DIR!" >nul 2>&1
        if not errorlevel 1 (
            set "DEFAULT_CUSTOM_DIR_STATUS=created"
        ) else (
            set "DEFAULT_CUSTOM_DIR_STATUS=create_failed"
        )
    )
)

for %%V in (!DCTP_OFFICE_VERSIONS!) do (
    if not defined WORD_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "WORD_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined PPT_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "PPT_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined EXCEL_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Excel\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "EXCEL_CUSTOM_TEMPLATE_PATH=%%C"
    )
)

for %%V in (!DCTP_OFFICE_VERSIONS!) do (
    if not defined WORD_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "WORD_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined PPT_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "PPT_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined EXCEL_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "EXCEL_CUSTOM_TEMPLATE_PATH=%%C"
    )
)

if not defined WORD_CUSTOM_TEMPLATE_PATH set "WORD_CUSTOM_TEMPLATE_PATH=!DEFAULT_CUSTOM_TEMPLATE_DIR!"
if not defined PPT_CUSTOM_TEMPLATE_PATH set "PPT_CUSTOM_TEMPLATE_PATH=!DEFAULT_CUSTOM_TEMPLATE_DIR!"
if not defined EXCEL_CUSTOM_TEMPLATE_PATH set "EXCEL_CUSTOM_TEMPLATE_PATH=!DEFAULT_CUSTOM_TEMPLATE_DIR!"

call :CleanPath WORD_CUSTOM_TEMPLATE_PATH
call :CleanPath PPT_CUSTOM_TEMPLATE_PATH
call :CleanPath EXCEL_CUSTOM_TEMPLATE_PATH

endlocal & (
    if defined DCTP_LOG_FILE set "DCTP_LOG_FILE=%DCTP_LOG_FILE%"
    if defined WORD_CUSTOM_TEMPLATE_PATH set "WORD_CUSTOM_TEMPLATE_PATH=%WORD_CUSTOM_TEMPLATE_PATH%"
    if defined PPT_CUSTOM_TEMPLATE_PATH set "PPT_CUSTOM_TEMPLATE_PATH=%PPT_CUSTOM_TEMPLATE_PATH%"
    if defined EXCEL_CUSTOM_TEMPLATE_PATH set "EXCEL_CUSTOM_TEMPLATE_PATH=%EXCEL_CUSTOM_TEMPLATE_PATH%"
    if defined DEFAULT_CUSTOM_TEMPLATE_DIR set "DEFAULT_CUSTOM_TEMPLATE_DIR=%DEFAULT_CUSTOM_TEMPLATE_DIR%"
)
exit /b 0

:CleanPath
if defined OfficeTemplateLib (
    call "%OfficeTemplateLib%" :CleanPath %*
    exit /b %errorlevel%
)
setlocal EnableDelayedExpansion
set "CP_VAR=%~1"
set "CP_VAL=!%CP_VAR%!"
if defined CP_VAL (
    if "!CP_VAL:~-1!"=="\\" set "CP_VAL=!CP_VAL:~0,-1!"
    if "!CP_VAL:~0,1!"=="\"" set "CP_VAL=!CP_VAL:~1!"
    if "!CP_VAL:~-1!"=="\"" set "CP_VAL=!CP_VAL:~0,-1!"
)
endlocal & set "%CP_VAR%=%CP_VAL%"
exit /b 0
