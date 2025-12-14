@echo off
setlocal enabledelayedexpansion
mode con: cols=140 lines=42
powershell -NoProfile -Command "$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(140, 1000); $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(140, 42)"

REM ========================================
REM File Launch Sequencer by landn.thrn
REM ========================================

REM Initialize variables
set "TEMP_DIR=%TEMP%\FileLaunchSequencer"
if not exist "!TEMP_DIR!" mkdir "!TEMP_DIR!"

REM Main Menu Entry Point
:main
cls
echo ========================================
echo   File Launch Sequencer by landn.thrn
echo ========================================
echo.
echo Choose an option:
echo.
echo   1. Folder-Based Protocol
echo   2. Manual File Path Protocol
echo   3. Presets From Folder-Based Protocol
echo   4. Presets From Manual File Path Protocol
echo.
echo.
set /p "menu_choice=Enter Command: "

REM Normalize input (trim spaces)
set "menu_choice=!menu_choice: =!"

REM Handle exit command (case-insensitive)
if /i "!menu_choice!"=="exit" (
    exit /b 0
)

REM Route to appropriate option
if "!menu_choice!"=="1" (
    call :option1
    set "menu_choice="
    goto :main
)
if "!menu_choice!"=="2" (
    call :option2
    set "menu_choice="
    goto :main
)
if "!menu_choice!"=="3" (
    call :presets_option1
    REM Check if preset should be run FIRST (before edit check)
    REM This ensures that if both are set (from edit flow), we run instead of editing again
    if defined run_preset_num (
        call :run_preset_handler
        set "run_preset_num="
        REM Clear edit_preset_num if it was set (from edit flow) so it doesn't trigger edit again
        set "edit_preset_num="
    )
    REM Check if preset should be edited (only if run_preset_num wasn't set)
    if defined edit_preset_num (
        call :edit_preset_selection_option1
        set "edit_preset_num="
    )
    goto :main
)
if "!menu_choice!"=="4" (
    call :presets_option2
    goto :main
)

REM Invalid input
echo.
echo Invalid. Try again.
timeout /t 1 /nobreak >nul
goto :main

REM ========================================
REM Option 1: Folder-Based Auto-Open Protocol
REM ========================================
:option1
cls
echo ========================================
echo          Folder-Based Protocol
echo ========================================
echo.

REM Step 1: File Format Input
:ask_format
echo Enter the file format for the files you want to auto open
echo Examples: prproj  ^|  .prproj  ^|  .txt   ^|   ...etc...
echo.
echo M - Back to Menu
echo.
set /p "file_ext=File Format: "

REM Handle menu return
if /i "!file_ext!"=="M" goto :option1_end

REM Normalize input (trim spaces)
set "file_ext=!file_ext: =!"

REM Empty input returns to menu
if "!file_ext!"=="" goto :option1_end

REM Normalize format (add dot if missing)
if "!file_ext:~0,1!" NEQ "." set "file_ext=.!file_ext!"

REM Step 2: Target Folder Path Input
:ask_scan_path
echo.
echo Enter the path of the folder where all !file_ext! files will be opened
echo.
:retry_scan_path
set /p "scan_path=Target Folder Path: "

REM Handle menu return
if /i "!scan_path!"=="M" goto :eof

REM Normalize input (remove surrounding quotes if present)
set "scan_path=!scan_path:"=!"

REM Trim leading/trailing spaces
for /f "tokens=* delims= " %%A in ("!scan_path!") do set "scan_path=%%A"

REM Validate folder exists
if exist "!scan_path!\" (
    set "scan_path=!scan_path!\"
) else if exist "!scan_path!" (
    set "scan_path=!scan_path!\"
) else (
    echo Invalid. Try again.
    goto :retry_scan_path
)

REM Step 2.1: Folder Scan Options
:ask_folder_scan_option
set "folder_scan_option="
echo.
echo What folder options would you like to scan/open by? 
echo Target files in:
echo.
echo 1 - All Subfolders in Target Folder
echo 2 - First Forefront Subfolders in Target Folder
echo 3 - Only Target Folder
echo.
echo M - Back to Menu
echo.
:retry_folder_scan_option
set /p "folder_scan_option=Enter Command: "

REM Handle menu return
if /i "!folder_scan_option!"=="M" goto :eof

REM Normalize input - remove spaces
set "folder_scan_option=!folder_scan_option: =!"

REM Validate input
if "!folder_scan_option!"=="1" goto :after_folder_scan_option
if "!folder_scan_option!"=="2" goto :after_folder_scan_option
if "!folder_scan_option!"=="3" goto :after_folder_scan_option

REM Invalid input
echo Invalid. Enter 1, 2, 3, or M.
goto :retry_folder_scan_option

:after_folder_scan_option

REM Step 3: Exclude Options
:ask_exclude_options
set "exclude_paths="
set "exclude_names="
set "exclude_keywords="
echo.
echo Do you want to use exclude options?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_exclude_options
set /p "exclude_options_choice=Enter Command: "

REM Handle menu return
if /i "!exclude_options_choice!"=="M" goto :eof

REM Normalize input
set "exclude_options_choice=!exclude_options_choice: =!"

REM Check for Y
if /i "!exclude_options_choice!"=="Y" goto :show_exclude_options

REM Check for S
if /i "!exclude_options_choice!"=="S" goto :skip_exclude_options

REM Invalid input
echo Invalid. Enter Y, S, or M.
goto :retry_exclude_options

:show_exclude_options
REM Check if option 3 (Only Target Folder) is selected - skip folder excludes if so
if "!folder_scan_option!"=="3" goto :skip_folder_excludes

REM Step 3a: Exclude Folders by Path
:ask_exclude_paths
set "exclude_paths="
echo.
echo Would you like to exclude folder(s) in your target path?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_exclude_paths
set /p "exclude_paths_choice=Enter Command: "

REM Handle menu return
if /i "!exclude_paths_choice!"=="M" goto :eof

REM Normalize input
set "exclude_paths_choice=!exclude_paths_choice: =!"

REM Check for Y
if /i "!exclude_paths_choice!"=="Y" goto :get_exclude_paths

REM Check for S
if /i "!exclude_paths_choice!"=="S" goto :skip_exclude_paths

REM Invalid input
echo Invalid. Enter Y, S, or M.
goto :retry_exclude_paths

:get_exclude_paths
echo.
echo Enter the folder path(s) you want excluded (comma-separated if multiple)
:retry_exclude_paths_input
set /p "input_exclude_paths=Folder Path(s): "
if /i "!exclude_paths!"=="M" goto :eof
REM Trim leading/trailing spaces only (preserve internal spaces)
for /f "tokens=* delims= " %%A in ("!input_exclude_paths!") do set "exclude_paths=%%A"
for /f "tokens=* delims= " %%A in ("!exclude_paths!") do set "exclude_paths=%%A"
if "!exclude_paths!"=="" (
    echo Invalid. Try again.
    goto :retry_exclude_paths_input
)

REM Validate paths exist (supports comma-separated multiple paths)
echo !exclude_paths!> "%TEMP%\validate_paths.txt"
echo !scan_path!> "%TEMP%\validate_base.txt"
for /f %%A in ('powershell -NoProfile -Command "$pathsStr = Get-Content (Join-Path $env:TEMP 'validate_paths.txt') -Raw; $basePath = Get-Content (Join-Path $env:TEMP 'validate_base.txt') -Raw; $pathsStr = $pathsStr.Trim(); $basePath = $basePath.Trim().TrimEnd('\'); $paths = $pathsStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }; $allValid = $true; foreach ($path in $paths) { if ($path) { $fullPath = $path; if (-not [System.IO.Path]::IsPathRooted($path)) { $fullPath = Join-Path $basePath $path }; try { $normalized = [System.IO.Path]::GetFullPath($fullPath); if (-not (Test-Path -LiteralPath $normalized -PathType Container)) { $allValid = $false; break } } catch { $allValid = $false; break } } }; if ($allValid) { Write-Output 'VALID' } else { Write-Output 'INVALID' }"') do set "path_validation=%%A"
del "%TEMP%\validate_paths.txt" "%TEMP%\validate_base.txt" >nul 2>&1

if "!path_validation!"=="INVALID" (
    echo Invalid. One or more paths do not exist. Try again.
    goto :retry_exclude_paths_input
)

goto :after_exclude_paths

:skip_exclude_paths
set "exclude_paths="

:after_exclude_paths

REM Step 3b: Exclude Folders by Name
:ask_exclude_names
set "exclude_names="
echo.
echo Would you like to exclude folder(s) in your target path by foldername?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_exclude_names
set /p "exclude_names_choice=Enter Command: "

REM Handle menu return
if /i "!exclude_names_choice!"=="M" goto :eof

REM Normalize input - remove spaces
set "exclude_names_choice=!exclude_names_choice: =!"

REM Check for Y
if /i "!exclude_names_choice!"=="Y" goto :get_exclude_names

REM Check for S
if /i "!exclude_names_choice!"=="S" goto :skip_exclude_names

REM Invalid input
echo Invalid. Enter Y, S, or M.
goto :retry_exclude_names

:get_exclude_names
echo.
echo Enter folder name(s) to exclude (comma-separated if multiple)
echo Example: Adobe Premiere Pro Auto-Save, Backup, Temp
echo.
:retry_exclude_names_input
set /p "input_exclude_names=Folder Name(s): "
if /i "!exclude_names!"=="M" goto :eof

REM Trim leading/trailing spaces only (preserve internal spaces)
if defined input_exclude_names (
    for /f "tokens=* delims= " %%A in ("!input_exclude_names!") do set "exclude_names=%%A"
    for /f "tokens=* delims= " %%A in ("!exclude_names!") do set "exclude_names=%%A"
)

if "!exclude_names!"=="" (
    echo Invalid. Try again.
    goto :retry_exclude_names_input
)
goto :after_exclude_names

:skip_exclude_names
set "exclude_names="

:after_exclude_names

:skip_folder_excludes
REM Step 3c: Exclude Files by Keyword
:ask_exclude_keywords
set "exclude_keywords="
echo.
echo Would you like to exclude files that have specific keyword(s) in the filename?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_exclude_keywords
set /p "exclude_keywords_choice=Enter Command: "

REM Handle menu return
if /i "!exclude_keywords_choice!"=="M" goto :eof

REM Normalize input - remove spaces
set "exclude_keywords_choice=!exclude_keywords_choice: =!"

REM Check for Y
if /i "!exclude_keywords_choice!"=="Y" goto :get_exclude_keywords

REM Check for S
if /i "!exclude_keywords_choice!"=="S" goto :skip_exclude_keywords

REM Invalid input
echo Invalid. Enter Y, S, or M.
goto :retry_exclude_keywords

:get_exclude_keywords
echo.
echo Enter keyword(s) to exclude files that have it in their name (comma-separated if multiple)
echo Example: cracked, backup, temp
echo.
:retry_exclude_keywords_input
set /p "input_exclude_keywords=Keyword(s): "
if /i "!exclude_keywords!"=="M" goto :eof

REM Trim leading/trailing spaces only (preserve internal spaces)
if defined input_exclude_keywords (
    for /f "tokens=* delims= " %%A in ("!input_exclude_keywords!") do set "exclude_keywords=%%A"
    for /f "tokens=* delims= " %%A in ("!exclude_keywords!") do set "exclude_keywords=%%A"
)

if "!exclude_keywords!"=="" (
    echo Invalid. Try again.
    goto :retry_exclude_keywords_input
)
goto :after_exclude_keywords

:skip_exclude_keywords
set "exclude_keywords="

:after_exclude_keywords
goto :after_all_exclude_options

:skip_exclude_options
set "exclude_paths="
set "exclude_names="
set "exclude_keywords="

:after_all_exclude_options

REM Step 4: Initial Delay Configuration
:ask_initial_delay
echo.
echo Set initial delay after first file is opened?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
echo.
:retry_initial_delay
set /p "use_initial=Enter Command: "

REM Handle menu return
if /i "!use_initial!"=="M" goto :eof

REM Normalize input
set "use_initial=!use_initial: =!"
set "use_initial=!use_initial!"

if /i "!use_initial!"=="Y" goto :get_initial_delay
if /i "!use_initial!"=="S" goto :skip_initial_delay
echo Invalid. Enter Y, S, or M.
goto :retry_initial_delay

:get_initial_delay
echo.
echo.
:retry_initial_delay_input
set /p "input_initial_delay=Initial Delay (seconds): "
if /i "!input_initial_delay!"=="M" goto :eof
set "initial_delay=!input_initial_delay: =!"
if "!initial_delay!"=="" set "initial_delay=0"
REM Validate it's a number
set "delay_validation="
echo !initial_delay!> "%TEMP%\validate_delay.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$delay = Get-Content (Join-Path $env:TEMP 'validate_delay.txt') -Raw; $delay = $delay.Trim(); if ($delay -eq '') { Write-Output 'VALID' } else { try { [double]$delay | Out-Null; Write-Output 'VALID' } catch { Write-Output 'INVALID' } }"') do set "delay_validation=%%A"
del "%TEMP%\validate_delay.txt" >nul 2>&1
if "!delay_validation!"=="INVALID" (
    echo Invalid. Enter a number.
    goto :retry_initial_delay_input
)
goto :after_initial_delay

:skip_initial_delay
set "use_initial=S"
set "initial_delay=0"

:after_initial_delay

REM Step 5: Delay Between Files Configuration
:ask_between_delay
echo.
echo Set delay between files?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
echo.
:retry_between_delay
set /p "use_between=Enter Command: "

REM Handle menu return
if /i "!use_between!"=="M" goto :eof

REM Normalize input
set "use_between=!use_between: =!"
set "use_between=!use_between!"

if /i "!use_between!"=="Y" goto :get_between_delay
if /i "!use_between!"=="S" goto :skip_between_delay
echo Invalid. Enter Y, S, or M.
goto :retry_between_delay

:get_between_delay
echo.
echo.
:retry_between_delay_input
set /p "input_between_delay=Delay Between Files (seconds): "
if /i "!input_between_delay!"=="M" goto :eof
set "between_delay=!input_between_delay: =!"
if "!between_delay!"=="" set "between_delay=0"
REM Validate it's a number
set "delay_validation="
echo !between_delay!> "%TEMP%\validate_delay.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$delay = Get-Content (Join-Path $env:TEMP 'validate_delay.txt') -Raw; $delay = $delay.Trim(); if ($delay -eq '') { Write-Output 'VALID' } else { try { [double]$delay | Out-Null; Write-Output 'VALID' } catch { Write-Output 'INVALID' } }"') do set "delay_validation=%%A"
del "%TEMP%\validate_delay.txt" >nul 2>&1
if "!delay_validation!"=="INVALID" (
    echo Invalid. Enter a number.
    goto :retry_between_delay_input
)
goto :after_between_delay

:skip_between_delay
set "use_between=S"
set "between_delay=0"

:after_between_delay

REM Step 6: File Scanning
call :scan_files
if errorlevel 1 goto :eof

REM Step 7: File Display
call :display_files
if errorlevel 1 goto :eof

REM Step 8: File Selection
call :ask_selection
if errorlevel 1 goto :eof

REM Step 9: Save as Preset
call :save_preset_option1
if errorlevel 1 goto :eof

REM Step 10: Confirmation
:option1_after_delays
REM Check if we're in auto-run mode (preset loaded)
if defined AUTO_RUN_PRESET (
    set "AUTO_RUN_PRESET="
    REM Skip confirmation and go directly to file opening
    goto :option1_open_files
)

call :confirm_open
if errorlevel 1 goto :eof

:option1_open_files

REM Step 11: File Opening
call :open_selected_files

REM Step 12: Completion Summary
call :show_completion

pause
if defined AUTO_RUN_PRESET (
    set "AUTO_RUN_PRESET="
    goto :main
)
goto :eof

:option1_end
exit /b 0

REM ========================================
REM File Scanning Subroutine
REM ========================================
:scan_files
if not defined SILENT_SCAN (
    echo.
    echo [SCAN] Scanning folder: !scan_path!
    echo [SCAN] Looking for: *!file_ext!
)

REM Initialize file list
set "file_count=0"
set "LAST_FILES=!TEMP_DIR!\FileLaunchSequencer_last.txt"
if exist "!LAST_FILES!" del /f /q "!LAST_FILES!"

REM Scan recursively and collect files using PowerShell (handles apostrophes and special chars)
REM Write parameters to temp file to avoid escaping issues
echo !file_ext!> "%TEMP%\scan_ext.txt"
echo !scan_path!> "%TEMP%\scan_path.txt"
echo !LAST_FILES!> "%TEMP%\output_file.txt"
if not defined folder_scan_option set "folder_scan_option=1"
echo !folder_scan_option!> "%TEMP%\folder_scan_option.txt"
if defined exclude_paths (
    echo !exclude_paths!> "%TEMP%\exclude_paths.txt"
) else (
    echo.> "%TEMP%\exclude_paths.txt"
)
if defined exclude_names (
    echo !exclude_names!> "%TEMP%\exclude_names.txt"
) else (
    echo.> "%TEMP%\exclude_names.txt"
)
if defined exclude_keywords (
    echo !exclude_keywords!> "%TEMP%\exclude_keywords.txt"
) else (
    echo.> "%TEMP%\exclude_keywords.txt"
)

powershell -NoProfile -Command "$ext = Get-Content (Join-Path $env:TEMP 'scan_ext.txt') -Raw; $scanPath = Get-Content (Join-Path $env:TEMP 'scan_path.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'output_file.txt') -Raw; $folderScanOption = Get-Content (Join-Path $env:TEMP 'folder_scan_option.txt') -Raw; $excludePathsStr = Get-Content (Join-Path $env:TEMP 'exclude_paths.txt') -Raw; $excludeNamesStr = Get-Content (Join-Path $env:TEMP 'exclude_names.txt') -Raw; $excludeKeywordsStr = Get-Content (Join-Path $env:TEMP 'exclude_keywords.txt') -Raw; $ext = $ext.Trim(); $scanPath = $scanPath.Trim().TrimEnd('\'); $outFile = $outFile.Trim(); $folderScanOption = if ($folderScanOption) { $folderScanOption.Trim() } else { '1' }; $excludePathsStr = $excludePathsStr.Trim(); $excludeNamesStr = $excludeNamesStr.Trim(); $excludeKeywordsStr = $excludeKeywordsStr.Trim(); $scanPathNormalized = [System.IO.Path]::GetFullPath($scanPath); $scanDepth = ($scanPathNormalized.Split('\') | Where-Object { $_ }).Count; if ($folderScanOption -eq '3') { $files = Get-ChildItem -LiteralPath $scanPath -Filter ('*' + $ext) -File -ErrorAction SilentlyContinue } elseif ($folderScanOption -eq '2') { $allFiles = Get-ChildItem -LiteralPath $scanPath -Filter ('*' + $ext) -Recurse -File -ErrorAction SilentlyContinue; $files = $allFiles | Where-Object { $fileDir = Split-Path -Parent $_.FullName; $fileDirNormalized = [System.IO.Path]::GetFullPath($fileDir); $fileDepth = ($fileDirNormalized.Split('\') | Where-Object { $_ }).Count; $fileDepth -le ($scanDepth + 1) } } else { $files = Get-ChildItem -LiteralPath $scanPath -Filter ('*' + $ext) -Recurse -File -ErrorAction SilentlyContinue }; if ($files) { $allFiles = $files; $filteredFiles = $files; if ($excludePathsStr -and $excludePathsStr.Length -gt 0 -and $excludePathsStr -ne '') { $excludePaths = @(); $excludePathsStr -split ',' | ForEach-Object { $p = $_.Trim(); if ($p -and $p.Length -gt 0) { if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $scanPath $p }; try { $normalized = [System.IO.Path]::GetFullPath($p).TrimEnd('\'); $excludePaths += $normalized } catch { } } }; if ($excludePaths.Count -gt 0) { $filteredFiles = $filteredFiles | Where-Object { $file = $_; $fileDir = Split-Path -Parent $file.FullName; $excluded = $false; foreach ($excludePath in $excludePaths) { if ($fileDir -eq $excludePath) { $excluded = $true; break } elseif ($fileDir.Length -gt $excludePath.Length) { $nextChar = $fileDir[$excludePath.Length]; if ($nextChar -eq '\' -and $fileDir.StartsWith($excludePath, [System.StringComparison]::OrdinalIgnoreCase)) { $excluded = $true; break } } }; -not $excluded } } }; if ($excludeNamesStr -and $excludeNamesStr.Length -gt 0 -and $excludeNamesStr -ne '') { $excludeNames = $excludeNamesStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }; if ($excludeNames.Count -gt 0) { $filteredFiles = $filteredFiles | Where-Object { $file = $_; $fileDir = Split-Path -Parent $file.FullName; $excluded = $false; $dirParts = $fileDir.Split('\'); foreach ($dirPart in $dirParts) { foreach ($excludeName in $excludeNames) { if ($dirPart -eq $excludeName) { $excluded = $true; break } } if ($excluded) { break } }; -not $excluded } } }; if ($excludeKeywordsStr -and $excludeKeywordsStr.Length -gt 0 -and $excludeKeywordsStr -ne '') { $excludeKeywords = $excludeKeywordsStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }; if ($excludeKeywords.Count -gt 0) { $filteredFiles = $filteredFiles | Where-Object { $file = $_; $fileName = $file.Name; $excluded = $false; foreach ($keyword in $excludeKeywords) { if ($fileName -like ('*' + $keyword + '*')) { $excluded = $true; break } }; -not $excluded } } }; $excludedCount = $allFiles.Count - $filteredFiles.Count; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, ($filteredFiles | ForEach-Object { $_.FullName }), $utf8NoBom); Write-Output ($filteredFiles.Count.ToString() + '|' + $excludedCount.ToString()) } else { Write-Output '0|0' }" > "%TEMP%\file_count.txt"

REM Cleanup temp files
del "%TEMP%\scan_ext.txt" "%TEMP%\scan_path.txt" "%TEMP%\folder_scan_option.txt" "%TEMP%\output_file.txt" "%TEMP%\exclude_paths.txt" "%TEMP%\exclude_names.txt" >nul 2>&1

REM Read file count and excluded count from PowerShell output
set "excluded_count=0"
for /f "tokens=1,2 delims=|" %%A in ('type "%TEMP%\file_count.txt"') do (
    set "file_count=%%A"
    set "excluded_count=%%B"
)
del "%TEMP%\file_count.txt" >nul 2>&1
if "!excluded_count!"=="" set "excluded_count=0"

REM If PowerShell failed or returned empty, try fallback method with batch
if "!file_count!"=="" set "file_count=0"
if !file_count! EQU 0 (
    REM Fallback: Use batch for loop (disable delayed expansion for pattern matching)
    set "file_count=0"
    setlocal DisableDelayedExpansion
    for /r "%scan_path%" %%F in (*%file_ext%) do (
        endlocal
        set /a file_count+=1
        echo %%F>>"!LAST_FILES!"
        setlocal DisableDelayedExpansion
    )
    endlocal
)

REM Check if files found
if !file_count! LEQ 0 (
    echo.
    echo [SCAN] No files found.
    pause
    exit /b 1
)

REM Sort files using Windows Explorer natural sort (PowerShell)
set "SORTED_FILES=!TEMP_DIR!\FileLaunchSequencer_sorted.txt"
REM Write file paths to temp variable file to avoid apostrophe issues
echo !LAST_FILES!> "%TEMP%\sort_input.txt"
echo !SORTED_FILES!> "%TEMP%\sort_output.txt"

powershell -NoProfile -Command "$inFile = Get-Content (Join-Path $env:TEMP 'sort_input.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'sort_output.txt') -Raw; $inFile = $inFile.Trim(); $outFile = $outFile.Trim(); $fileArray = Get-Content -LiteralPath $inFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($fileArray) { $sorted = $fileArray | Sort-Object { [regex]::Replace($_, '\d+', { $args[0].Value.PadLeft(20) }) }, { $_.ToLower() }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $sorted, $utf8NoBom) }"

REM Cleanup temp files
del "%TEMP%\sort_input.txt" "%TEMP%\sort_output.txt" >nul 2>&1

if exist "!SORTED_FILES!" (
    move /y "!SORTED_FILES!" "!LAST_FILES!" >nul
)

exit /b 0

REM ========================================
REM File Display Subroutine
REM ========================================
:display_files
echo.
echo FILE SCAN:
echo.

REM Verify file list exists
if not exist "!LAST_FILES!" (
    echo [ERROR] File list not found.
    exit /b 1
)

REM Display files with numbering using PowerShell (more reliable)
set "DISPLAY_FILE=!TEMP_DIR!\FileLaunchSequencer_display.txt"
if exist "!DISPLAY_FILE!" del /f /q "!DISPLAY_FILE!"

REM Write scan path to temp file for PowerShell
echo !scan_path!> "%TEMP%\display_base.txt"
echo !LAST_FILES!> "%TEMP%\display_list.txt"
echo !DISPLAY_FILE!> "%TEMP%\display_output.txt"

REM Use PowerShell to format and display files
powershell -NoProfile -Command "$base = Get-Content (Join-Path $env:TEMP 'display_base.txt') -Raw; $listFile = Get-Content (Join-Path $env:TEMP 'display_list.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'display_output.txt') -Raw; $base = $base.Trim().TrimEnd('\'); $listFile = $listFile.Trim(); $outFile = $outFile.Trim(); $fileArray = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($fileArray) { $counter = 0; $baseName = Split-Path -Leaf $base; $formattedLines = @(); foreach ($file in $fileArray) { $counter++; $file = $file.Trim(); $rel = $file.Replace($base, ''); if ($rel -match '^\\') { $rel = $rel.Substring(1) }; $formattedLines += ($counter.ToString() + '. \' + $baseName + '\' + $rel) }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $formattedLines, $utf8NoBom) }"

REM Cleanup temp files
del "%TEMP%\display_base.txt" "%TEMP%\display_list.txt" "%TEMP%\display_output.txt" >nul 2>&1

REM Display the formatted file list
if exist "!DISPLAY_FILE!" (
    type "!DISPLAY_FILE!"
    del /f /q "!DISPLAY_FILE!" >nul 2>&1
) else (
    echo [ERROR] Could not format file list.
)

echo.
echo Found !file_count! !file_ext! Files
echo Found !excluded_count! !file_ext! Files to Exclude
echo.

exit /b 0

REM ========================================
REM File Selection Subroutine
REM ========================================
:ask_selection
echo.
echo How would you like to open these?
echo.
echo File Selection Examples: (all, 1-11, 8-22, 1,4,20,11,14,15,)
echo If you use 'all' you can exclude like (all,-13,-24,-30-40)
echo.
:retry_selection
set /p "selection=File Selection: "

REM Handle menu return
if /i "!selection!"=="M" goto :after_selection_edit

REM Normalize input (trim spaces but preserve structure)
set "selection=!selection: =!"

REM Check for "all" command (with or without exclusions)
set "is_all=0"
set "selection_lower=!selection!"
if /i "!selection!"=="all" (
    set "is_all=1"
    set "has_exclusions=0"
) else (
    REM Check if starts with "all" (case-insensitive)
    for /f "delims=" %%A in ('powershell -NoProfile -Command "Write-Output ('!selection!'.ToLower().StartsWith('all'))"') do set "starts_all=%%A"
    if "!starts_all!"=="True" (
        set "is_all=1"
        set "has_exclusions=1"
    )
)

REM If "all" without exclusions, use full list
if !is_all! EQU 1 (
    if !has_exclusions! EQU 0 (
        set "SELECTED_FILES=!LAST_FILES!"
        exit /b 0
    )
)

REM Parse selection using PowerShell (handles "all" with exclusions and regular selections)
set "SELECTED_FILES=!TEMP_DIR!\FileLaunchSequencer_selected.txt"
if exist "!SELECTED_FILES!" del /f /q "!SELECTED_FILES!"

REM Write parameters to temp files to avoid apostrophe issues
echo !selection!> "%TEMP%\sel_input.txt"
echo !LAST_FILES!> "%TEMP%\sel_list.txt"
echo !SELECTED_FILES!> "%TEMP%\sel_output.txt"

REM Process selection using PowerShell (handles "all" with exclusions and regular selections)
powershell -NoProfile -Command "$selection = Get-Content (Join-Path $env:TEMP 'sel_input.txt') -Raw; $listFile = Get-Content (Join-Path $env:TEMP 'sel_list.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'sel_output.txt') -Raw; $selection = $selection.Trim(); $listFile = $listFile.Trim(); $outFile = $outFile.Trim(); $files = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if (-not $files -or $files.Count -eq 0) { exit 1 }; $selected = @(); $isAll = $selection -match '^all'; if ($isAll) { $selected = $files; $parts = $selection -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^-' }; foreach ($part in $parts) { $part = $part.TrimStart('-'); if ($part -match '^(\d+)-(\d+)$') { $start = [int]$matches[1]; $end = [int]$matches[2]; if ($start -ge 1 -and $end -le $files.Count -and $start -le $end) { for ($i = $start-1; $i -le $end-1; $i++) { $selected = $selected | Where-Object { $_ -ne $files[$i] } } } } elseif ($part -match '^\d+$') { $num = [int]$part; if ($num -ge 1 -and $num -le $files.Count) { $selected = $selected | Where-Object { $_ -ne $files[$num-1] } } } } } else { $parts = $selection -split ',' | ForEach-Object { $_.Trim() }; foreach ($part in $parts) { if ($part -match '^(\d+)-(\d+)$') { $start = [int]$matches[1]; $end = [int]$matches[2]; if ($start -ge 1 -and $end -le $files.Count -and $start -le $end) { $selected += $files[($start-1)..($end-1)] } } elseif ($part -match '^\d+$') { $num = [int]$part; if ($num -ge 1 -and $num -le $files.Count) { $selected += $files[$num-1] } } }; $selected = $selected | Select-Object -Unique }; if ($selected.Count -eq 0) { exit 1 }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $selected, $utf8NoBom)"

del "%TEMP%\sel_input.txt" "%TEMP%\sel_list.txt" "%TEMP%\sel_output.txt" >nul 2>&1

if not exist "!SELECTED_FILES!" (
    echo Invalid. Try again.
    goto :retry_selection
)

REM Check if selection is empty
for /f %%A in ('powershell -NoProfile -Command "(Get-Content '!SELECTED_FILES!' -Raw -Encoding UTF8).Length"') do set "file_size=%%A"
if "!file_size!"=="0" (
    echo Invalid. Try again.
    goto :retry_selection
)

exit /b 0

REM ========================================
REM Confirmation Subroutine
REM ========================================
:confirm_open
echo.
echo Ready to open !file_ext! files?
echo   Y - Yes
echo   M - Back to Menu
echo.
echo.
:retry_confirm
set /p "confirm=Enter Command: "

REM Handle menu return
if /i "!confirm!"=="M" exit /b 1

REM Normalize input
set "confirm=!confirm: =!"

if /i "!confirm!"=="Y" (
    exit /b 0
) else (
    echo Invalid. Enter Y or M.
    goto :retry_confirm
)

REM ========================================
REM File Opening Subroutine
REM ========================================
:open_selected_files
echo.

REM Count total files to open first
set "total_to_open=0"
echo !SELECTED_FILES!> "%TEMP%\count_file.txt"
for /f %%A in ('powershell -NoProfile -Command "$filePath = Get-Content (Join-Path $env:TEMP 'count_file.txt') -Raw; $filePath = $filePath.Trim(); $files = Get-Content -LiteralPath $filePath -Encoding UTF8 | Where-Object { $_.Trim() }; Write-Output $files.Count"') do set "total_to_open=%%A"
del "%TEMP%\count_file.txt" >nul 2>&1

echo [SCAN] Opening !total_to_open! !file_ext! files from the scanned list
echo.

set "opened_count=0"
set "failed_count=0"
set "file_index=0"
set "FAILED_FILES=!TEMP_DIR!\FileLaunchSequencer_failed.txt"
if exist "!FAILED_FILES!" del /f /q "!FAILED_FILES!" >nul 2>&1

for /f "usebackq delims=" %%F in ("!SELECTED_FILES!") do (
    set /a file_index+=1
    set "current_file=%%F"
    
    REM Display opening message with shortened path (handle apostrophes)
    echo !scan_path!> "%TEMP%\open_base.txt"
    echo %%F> "%TEMP%\open_full.txt"
    for /f "delims=" %%P in ('powershell -NoProfile -Command "$base = Get-Content (Join-Path $env:TEMP 'open_base.txt') -Raw; $full = Get-Content (Join-Path $env:TEMP 'open_full.txt') -Raw; $base = $base.Trim().TrimEnd('\'); $full = $full.Trim(); $rel = $full.Replace($base, ''); if ($rel -match '^\\') { $rel = $rel.Substring(1) }; $baseName = Split-Path -Leaf $base; Write-Output ('\' + $baseName + '\' + $rel)"') do (
        echo [OPEN] %%P
    )
    del "%TEMP%\open_base.txt" "%TEMP%\open_full.txt" >nul 2>&1
    
    REM Open file and check for errors
    start "" "%%F" 2>nul
    if errorlevel 1 (
        echo [ERROR] Failed to open: %%F
        echo %%F>>"!FAILED_FILES!"
        set /a failed_count+=1
    ) else (
        set /a opened_count+=1
    )

    REM Apply initial delay after first file is opened
    if !file_index! EQU 1 (
        if /i "!use_initial!"=="Y" (
            if !initial_delay! GTR 0 (
                echo [WAIT] Initial !initial_delay! seconds...
                timeout /t !initial_delay! /nobreak >nul
            )
        )
    )

    REM Apply between delay after each file (except first file and last file)
    if !file_index! GTR 1 if !file_index! LSS !total_to_open! (
        if /i "!use_between!"=="Y" (
            if !between_delay! GTR 0 (
                echo [WAIT] !between_delay! seconds...
                timeout /t !between_delay! /nobreak >nul
            )
        )
    )
)

exit /b 0

REM ========================================
REM Completion Summary Subroutine
REM ========================================
:show_completion
echo.
echo Completed!
echo.
echo Opened !opened_count! !file_ext! files
echo Excluded !excluded_count! !file_ext! files
if !failed_count! GTR 0 (
    echo !failed_count! Files Failed to Open
    if exist "!FAILED_FILES!" (
        echo.
        echo Failed files:
        type "!FAILED_FILES!"
    )
) else (
    echo 0 Files Failed to Open
)
echo.

REM Cleanup
if exist "!SELECTED_FILES!" if not "!SELECTED_FILES!"=="!LAST_FILES!" del /f /q "!SELECTED_FILES!" >nul 2>&1
if exist "!FAILED_FILES!" del /f /q "!FAILED_FILES!" >nul 2>&1

exit /b 0

REM ========================================
REM Option 2: Manual File Path Protocol
REM ========================================
:option2
cls
echo ========================================
echo        Manual File Path Protocol
echo ========================================
echo.

REM Step 1: File Paths Input
:ask_file_paths
echo Enter file paths to open (comma-separated if multiple)
echo Example: C:\Folder\file1.txt, C:\Folder\file2.txt, C:\Folder\file3.txt
echo.
echo M - Back to Menu
echo.
:retry_file_paths
set /p "file_paths=File Paths: "

REM Handle menu return
if /i "!file_paths!"=="M" goto :option2_end

REM Normalize input (remove surrounding quotes if present)
set "file_paths=!file_paths:"=!"

REM Trim leading/trailing spaces
for /f "tokens=* delims= " %%A in ("!file_paths!") do set "file_paths=%%A"

REM Validate input not empty
if "!file_paths!"=="" (
    echo Invalid. Try again.
    goto :retry_file_paths
)

REM Validate all paths exist (supports comma-separated multiple paths)
echo !file_paths!> "%TEMP%\validate_file_paths.txt"
for /f %%A in ('powershell -NoProfile -Command "$pathsStr = Get-Content (Join-Path $env:TEMP 'validate_file_paths.txt') -Raw; $pathsStr = $pathsStr.Trim(); $paths = $pathsStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }; $allValid = $true; foreach ($path in $paths) { if ($path) { try { $normalized = [System.IO.Path]::GetFullPath($path); if (-not (Test-Path -LiteralPath $normalized -PathType Leaf)) { $allValid = $false; break } } catch { $allValid = $false; break } } }; if ($allValid) { Write-Output 'VALID' } else { Write-Output 'INVALID' }"') do set "path_validation=%%A"
del "%TEMP%\validate_file_paths.txt" >nul 2>&1

if "!path_validation!"=="INVALID" (
    echo Invalid. One or more file paths do not exist. Try again.
    goto :retry_file_paths
)

REM Store file paths in list (parse and write to temp file)
set "MANUAL_FILES=!TEMP_DIR!\FileLaunchSequencer_manual.txt"
if exist "!MANUAL_FILES!" del /f /q "!MANUAL_FILES!"

REM Parse paths and write to file using PowerShell
echo !file_paths!> "%TEMP%\parse_paths.txt"
echo !MANUAL_FILES!> "%TEMP%\parse_output.txt"
powershell -NoProfile -Command "$pathsStr = Get-Content (Join-Path $env:TEMP 'parse_paths.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'parse_output.txt') -Raw; $pathsStr = $pathsStr.Trim(); $outFile = $outFile.Trim(); $paths = $pathsStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }; $fullPaths = @(); foreach ($path in $paths) { try { $normalized = [System.IO.Path]::GetFullPath($path); $fullPaths += $normalized } catch { } }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $fullPaths, $utf8NoBom)"
del "%TEMP%\parse_paths.txt" "%TEMP%\parse_output.txt" >nul 2>&1

REM Count files
set "file_count=0"
for /f %%A in ('powershell -NoProfile -Command "$filePath = '!MANUAL_FILES!'; $files = Get-Content -LiteralPath $filePath -Encoding UTF8 | Where-Object { $_.Trim() }; Write-Output $files.Count"') do set "file_count=%%A"

if "!file_count!"=="0" (
    echo Invalid. No valid file paths found. Try again.
    goto :retry_file_paths
)

REM Step 2: Initial Delay Configuration
:ask_initial_delay_option2
echo.
echo Set initial delay after first file is opened?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
echo.
:retry_initial_delay_option2
set /p "use_initial=Enter Command: "

REM Handle menu return
if /i "!use_initial!"=="M" goto :option2_end

REM Normalize input
set "use_initial=!use_initial: =!"
set "use_initial=!use_initial!"

REM Check for empty input
if "!use_initial!"=="" (
    echo Invalid. Enter Y, S, or M.
    goto :retry_initial_delay_option2
)

if /i "!use_initial!"=="Y" goto :get_initial_delay_option2
if /i "!use_initial!"=="S" goto :skip_initial_delay_option2
echo Invalid. Enter Y, S, or M.
goto :retry_initial_delay_option2

:get_initial_delay_option2
echo.
echo.
:retry_initial_delay_input_option2
set /p "initial_delay=Initial Delay (seconds): "
if /i "!initial_delay!"=="M" goto :option2_end
set "initial_delay=!initial_delay: =!"
if "!initial_delay!"=="" set "initial_delay=0"
REM Validate it's a number
set "delay_validation="
echo !initial_delay!> "%TEMP%\validate_delay.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$delay = Get-Content (Join-Path $env:TEMP 'validate_delay.txt') -Raw; $delay = $delay.Trim(); if ($delay -eq '') { Write-Output 'VALID' } else { try { [double]$delay | Out-Null; Write-Output 'VALID' } catch { Write-Output 'INVALID' } }"') do set "delay_validation=%%A"
del "%TEMP%\validate_delay.txt" >nul 2>&1
if "!delay_validation!"=="INVALID" (
    echo Invalid. Enter a number.
    goto :retry_initial_delay_input_option2
)
goto :after_initial_delay_option2

:skip_initial_delay_option2
set "use_initial=S"
set "initial_delay=0"
goto :after_initial_delay_option2

:after_initial_delay_option2

REM Step 3: Delay Between Files Configuration
:ask_between_delay_option2
echo.
echo Set delay between files?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
echo.
:retry_between_delay_option2
set /p "use_between=Enter Command: "

REM Handle menu return
if /i "!use_between!"=="M" goto :option2_end

REM Normalize input
set "use_between=!use_between: =!"
set "use_between=!use_between!"

REM Check for empty input
if "!use_between!"=="" (
    echo Invalid. Enter Y, S, or M.
    goto :retry_between_delay_option2
)

if /i "!use_between!"=="Y" goto :get_between_delay_option2
if /i "!use_between!"=="S" goto :skip_between_delay_option2
echo Invalid. Enter Y, S, or M.
goto :retry_between_delay_option2

:get_between_delay_option2
echo.
echo.
:retry_between_delay_input_option2
set /p "between_delay=Delay Between Files (seconds): "
if /i "!between_delay!"=="M" goto :option2_end
set "between_delay=!between_delay: =!"
if "!between_delay!"=="" set "between_delay=0"
REM Validate it's a number
set "delay_validation="
echo !between_delay!> "%TEMP%\validate_delay.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$delay = Get-Content (Join-Path $env:TEMP 'validate_delay.txt') -Raw; $delay = $delay.Trim(); if ($delay -eq '') { Write-Output 'VALID' } else { try { [double]$delay | Out-Null; Write-Output 'VALID' } catch { Write-Output 'INVALID' } }"') do set "delay_validation=%%A"
del "%TEMP%\validate_delay.txt" >nul 2>&1
if "!delay_validation!"=="INVALID" (
    echo Invalid. Enter a number.
    goto :retry_between_delay_input_option2
)
goto :after_between_delay_option2

:skip_between_delay_option2
set "use_between=S"
set "between_delay=0"
goto :after_between_delay_option2

:after_between_delay_option2

REM Ensure all delay variables are set
if not defined use_initial set "use_initial=S"
if not defined initial_delay set "initial_delay=0"
if not defined use_between set "use_between=S"
if not defined between_delay set "between_delay=0"

REM Step 4: Save as Preset
echo.
echo Save this protocol as a preset?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_save_preset_option2_inline
set "save_preset_choice="
set /p "save_preset_choice=Enter Command: "

REM Handle menu return
if /i "!save_preset_choice!"=="M" goto :option2_end

REM Normalize input
set "save_preset_choice=!save_preset_choice: =!"

if /i "!save_preset_choice!"=="Y" goto :get_preset_name_option2_inline
if /i "!save_preset_choice!"=="S" goto :after_save_preset_option2
echo Invalid. Enter Y, S, or M.
goto :retry_save_preset_option2_inline

:get_preset_name_option2_inline
echo.
echo Enter preset name
:retry_preset_name_option2_inline
set "preset_name="
set /p "preset_name=Preset Name: "
if /i "!preset_name!"=="M" goto :option2_end

REM Basic validation
if not defined preset_name (
    echo Invalid. Try again.
    goto :retry_preset_name_option2_inline
)
if "!preset_name!"=="" (
    echo Invalid. Try again.
    goto :retry_preset_name_option2_inline
)

REM Create preset directory if it doesn't exist
set "PRESET_DIR=%~dp0Presets\Option2"
if not exist "!PRESET_DIR!" mkdir "!PRESET_DIR!" >nul 2>&1

REM Sanitize preset name for filename (simple batch method)
REM Replace invalid filename characters one by one
set "safe_name=!preset_name!"
set "safe_name=!safe_name:\=_!"
set "safe_name=!safe_name:/=_!"
set "safe_name=!safe_name::=_!"
set "safe_name=!safe_name:?=_!"
set "safe_name=!safe_name:"=_!"
set "safe_name=!safe_name:<=_!"
set "safe_name=!safe_name:>=_!"
set "safe_name=!safe_name:|=_!"
REM For * character, replace with underscore (batch can't do this directly, so skip it for now)
REM If preset name contains *, it will cause issues, but that's rare
set "PRESET_FILE=!PRESET_DIR!\!safe_name!.json"

REM Initialize variables if not set
if not defined use_initial set "use_initial="
if not defined initial_delay set "initial_delay=0"
if not defined use_between set "use_between="
if not defined between_delay set "between_delay=0"

REM Create PowerShell script to save preset
set "PS_SCRIPT=!TEMP_DIR!\save_preset_option2_script.ps1"
(
    echo try {
    echo     $filePaths = Get-Content (Join-Path $env:TEMP 'preset_file_paths.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $useInitial = Get-Content (Join-Path $env:TEMP 'preset_use_initial.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $initialDelay = Get-Content (Join-Path $env:TEMP 'preset_initial_delay.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $useBetween = Get-Content (Join-Path $env:TEMP 'preset_use_between.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $betweenDelay = Get-Content (Join-Path $env:TEMP 'preset_between_delay.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $outFile = Get-Content (Join-Path $env:TEMP 'preset_output.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $displayName = Get-Content (Join-Path $env:TEMP 'preset_display_name.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     if ($outFile^) { $outFile = $outFile.Trim^(^) } else { $outFile = '' }
    echo     if ($displayName^) { $displayName = $displayName.Trim^(^) } else { $displayName = '' }
    echo     if (-not $outFile -or $outFile -eq '' -or -not $displayName -or $displayName -eq ''^) { Write-Output 'ERROR'; exit }
    echo     if ($filePaths^) { $filePaths = $filePaths.Trim^(^) } else { $filePaths = '' }
    echo     if ($useInitial^) { $useInitial = $useInitial.Trim^(^) } else { $useInitial = '' }
    echo     if ($initialDelay^) { $initialDelay = $initialDelay.Trim^(^) } else { $initialDelay = '0' }
    echo     if ($useBetween^) { $useBetween = $useBetween.Trim^(^) } else { $useBetween = '' }
    echo     if ($betweenDelay^) { $betweenDelay = $betweenDelay.Trim^(^) } else { $betweenDelay = '0' }
    echo     $preset = @{ 'preset_name' = $displayName; 'file_paths' = $filePaths; 'use_initial' = $useInitial; 'initial_delay' = $initialDelay; 'use_between' = $useBetween; 'between_delay' = $betweenDelay; 'created' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'^) }
    echo     $json = $preset ^| ConvertTo-Json
    echo     $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    echo     [System.IO.File]::WriteAllText($outFile, $json, $utf8NoBom^)
    echo     Write-Output 'SUCCESS'
    echo } catch {
    echo     Write-Output 'ERROR'
    echo }
) > "!PS_SCRIPT!" 2>nul
if not exist "!PS_SCRIPT!" (
    echo.
    echo [ERROR] Failed to create preset script.
    goto :after_save_preset_option2
)

REM Write all settings to temp files for PowerShell
if not defined file_paths set "file_paths="
> "%TEMP%\preset_file_paths.txt" echo !file_paths!
> "%TEMP%\preset_use_initial.txt" echo !use_initial!
> "%TEMP%\preset_initial_delay.txt" echo !initial_delay!
> "%TEMP%\preset_use_between.txt" echo !use_between!
> "%TEMP%\preset_between_delay.txt" echo !between_delay!
> "%TEMP%\preset_output.txt" echo !PRESET_FILE!
> "%TEMP%\preset_display_name.txt" echo !preset_name!


REM Verify critical files exist
if not exist "%TEMP%\preset_output.txt" (
    echo.
    echo [ERROR] Failed to create output file path.
    goto :after_save_preset_option2
)
if not exist "%TEMP%\preset_display_name.txt" (
    echo.
    echo [ERROR] Failed to create display name file.
    goto :after_save_preset_option2
)

REM Run PowerShell script to save preset
set "save_result=ERROR"
if exist "!PS_SCRIPT!" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "!PS_SCRIPT!" > "%TEMP%\preset_save_result.txt" 2> "%TEMP%\preset_save_error.txt"
    if exist "%TEMP%\preset_save_result.txt" (
        for /f "delims=" %%A in ('type "%TEMP%\preset_save_result.txt"') do set "save_result=%%A"
    )
    if exist "%TEMP%\preset_save_error.txt" (
        for %%F in ("%TEMP%\preset_save_error.txt") do if %%~zF gtr 0 (
            del "%TEMP%\preset_save_error.txt" >nul 2>&1
        )
    )
    del "!PS_SCRIPT!" >nul 2>&1
) else (
    echo.
    echo [ERROR] Preset script file not found.
    goto :after_save_preset_option2
)

REM Cleanup temp files
del "%TEMP%\preset_file_paths.txt" "%TEMP%\preset_use_initial.txt" "%TEMP%\preset_initial_delay.txt" "%TEMP%\preset_use_between.txt" "%TEMP%\preset_between_delay.txt" "%TEMP%\preset_output.txt" "%TEMP%\preset_display_name.txt" "%TEMP%\preset_save_result.txt" >nul 2>&1

if "!save_result!"=="SUCCESS" (
    if exist "!PRESET_FILE!" (
        echo.
        echo Preset saved: !preset_name!
    ) else (
        echo.
        echo [ERROR] Failed to save preset.
    )
) else (
    echo.
    echo [ERROR] Failed to save preset.
)

:after_save_preset_option2

REM Step 5: Confirmation
call :confirm_open_option2
if errorlevel 1 goto :eof

REM Step 6: File Opening
call :open_manual_files

REM Step 7: Completion Summary
call :show_completion_option2

pause
goto :option2_end

REM ========================================
REM Option 2: Confirmation Subroutine
REM ========================================
:confirm_open_option2
echo.
echo Ready to open files?
echo   Y - Yes
echo   M - Back to Menu
echo.
echo.
:retry_confirm_option2
set /p "confirm=Enter Command: "

REM Handle menu return
if /i "!confirm!"=="M" exit /b 1

REM Normalize input
set "confirm=!confirm: =!"

if /i "!confirm!"=="Y" (
    exit /b 0
) else (
    echo Invalid. Enter Y or M.
    goto :retry_confirm_option2
)

REM ========================================
REM Option 2: File Opening Subroutine
REM ========================================
:open_manual_files
echo.

set "opened_count=0"
set "failed_count=0"
set "file_index=0"
set "FAILED_FILES=!TEMP_DIR!\FileLaunchSequencer_failed.txt"
if exist "!FAILED_FILES!" del /f /q "!FAILED_FILES!" >nul 2>&1

REM Count total files
set "total_to_open=0"
echo !MANUAL_FILES!> "%TEMP%\count_file.txt"
for /f %%A in ('powershell -NoProfile -Command "$filePath = Get-Content (Join-Path $env:TEMP 'count_file.txt') -Raw; $filePath = $filePath.Trim(); $files = Get-Content -LiteralPath $filePath -Encoding UTF8 | Where-Object { $_.Trim() }; Write-Output $files.Count"') do set "total_to_open=%%A"
del "%TEMP%\count_file.txt" >nul 2>&1

REM Get file extension from first file for display
set "file_ext_display="
if exist "!MANUAL_FILES!" (
    for /f "usebackq delims=" %%F in ("!MANUAL_FILES!") do (
        for %%E in ("%%F") do set "file_ext_display=%%~xE"
        goto :got_ext
    )
)
:got_ext
if not defined file_ext_display set "file_ext_display="
if "!file_ext_display!"=="" (
    echo [SCAN] Opening !total_to_open! files
) else (
    echo [SCAN] Opening !total_to_open! !file_ext_display! files
)
echo.

for /f "usebackq delims=" %%F in ("!MANUAL_FILES!") do (
    set /a file_index+=1
    set "current_file=%%F"
    
    REM Display opening message with file name
    echo [OPEN] %%F
    
    REM Open file and check for errors
    start "" "%%F" 2>nul
    if errorlevel 1 (
        echo [ERROR] Failed to open: %%F
        echo %%F>>"!FAILED_FILES!"
        set /a failed_count+=1
    ) else (
        set /a opened_count+=1
    )

    REM Apply initial delay after first file is opened
    if !file_index! EQU 1 (
        if /i "!use_initial!"=="Y" (
            if !initial_delay! GTR 0 (
                echo [WAIT] Initial !initial_delay! seconds...
                timeout /t !initial_delay! /nobreak >nul
            )
        )
    )

    REM Apply between delay after each file (except first file and last file)
    if !file_index! GTR 1 if !file_index! LSS !total_to_open! (
        if /i "!use_between!"=="Y" (
            if !between_delay! GTR 0 (
                echo [WAIT] !between_delay! seconds...
                timeout /t !between_delay! /nobreak >nul
            )
        )
    )
)

:option2_end
set "menu_choice="
goto :main

REM ========================================
REM Option 2: Completion Summary Subroutine
REM ========================================
:show_completion_option2
echo.
echo Completed!
echo.
echo Opened !opened_count! files
if !failed_count! GTR 0 (
    echo !failed_count! Files Failed to Open
    if exist "!FAILED_FILES!" (
        echo.
        echo Failed files:
        type "!FAILED_FILES!"
    )
) else (
    echo 0 Files Failed to Open
)
echo.

REM Cleanup
if exist "!MANUAL_FILES!" del /f /q "!MANUAL_FILES!" >nul 2>&1
if exist "!FAILED_FILES!" del /f /q "!FAILED_FILES!" >nul 2>&1

exit /b 0

REM ========================================
REM Save Preset Subroutine (Option 2)
REM ========================================
:save_preset_option2
echo [TEST] Subroutine save_preset_option2 called successfully!
echo.
echo Save this protocol as a preset?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_save_preset_option2
set "save_preset_choice="
set /p "save_preset_choice=Enter Command: "

REM Handle menu return
if /i "!save_preset_choice!"=="M" exit /b 1

REM Normalize input
set "save_preset_choice=!save_preset_choice: =!"

if /i "!save_preset_choice!"=="Y" goto :get_preset_name_option2
if /i "!save_preset_choice!"=="S" exit /b 0
echo Invalid. Enter Y, S, or M.
goto :retry_save_preset_option2

:get_preset_name_option2
echo.
echo Enter preset name
:retry_preset_name_option2
set "preset_name="
set /p "preset_name=Preset Name: "
if /i "!preset_name!"=="M" exit /b 1

REM Basic validation
if not defined preset_name (
    echo Invalid. Try again.
    goto :retry_preset_name_option2
)
if "!preset_name!"=="" (
    echo Invalid. Try again.
    goto :retry_preset_name_option2
)

REM Create preset directory if it doesn't exist
set "PRESET_DIR=%~dp0Presets\Option2"
if not exist "!PRESET_DIR!" mkdir "!PRESET_DIR!" >nul 2>&1

REM Sanitize preset name for filename using PowerShell (after user enters name)
set "PRESET_FILE=!PRESET_DIR!\!preset_name!.json"
echo !preset_name!> "%TEMP%\preset_name_temp.txt"
echo !PRESET_FILE!> "%TEMP%\preset_file_temp.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "try { $name = Get-Content (Join-Path $env:TEMP 'preset_name_temp.txt') -Raw -ErrorAction Stop; $file = Get-Content (Join-Path $env:TEMP 'preset_file_temp.txt') -Raw -ErrorAction Stop; $name = $name.Trim(); $file = $file.Trim(); $dir = Split-Path $file -Parent; $sanitized = $name -replace '[\\\\/:*?\"<>|]', '_'; $finalFile = Join-Path $dir ($sanitized + '.json'); Write-Output $finalFile } catch { Write-Output '' }"') do set "PRESET_FILE=%%A"
del "%TEMP%\preset_name_temp.txt" "%TEMP%\preset_file_temp.txt" >nul 2>&1
if "!PRESET_FILE!"=="" (
    echo.
    echo [ERROR] Failed to sanitize preset name.
    goto :retry_preset_name_option2
)

REM Initialize variables if not set
if not defined use_initial set "use_initial="
if not defined initial_delay set "initial_delay=0"
if not defined use_between set "use_between="
if not defined between_delay set "between_delay=0"

REM Create PowerShell script to save preset
set "PS_SCRIPT=!TEMP_DIR!\save_preset_option2_script.ps1"
(
    echo try {
    echo     $filePaths = Get-Content (Join-Path $env:TEMP 'preset_file_paths.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $useInitial = Get-Content (Join-Path $env:TEMP 'preset_use_initial.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $initialDelay = Get-Content (Join-Path $env:TEMP 'preset_initial_delay.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $useBetween = Get-Content (Join-Path $env:TEMP 'preset_use_between.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $betweenDelay = Get-Content (Join-Path $env:TEMP 'preset_between_delay.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $outFile = Get-Content (Join-Path $env:TEMP 'preset_output.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     $displayName = Get-Content (Join-Path $env:TEMP 'preset_display_name.txt'^) -Raw -ErrorAction SilentlyContinue
    echo     if ($outFile^) { $outFile = $outFile.Trim^(^) } else { $outFile = '' }
    echo     if ($displayName^) { $displayName = $displayName.Trim^(^) } else { $displayName = '' }
    echo     if (-not $outFile -or $outFile -eq '' -or -not $displayName -or $displayName -eq ''^) { Write-Output 'ERROR'; exit }
    echo     if ($filePaths^) { $filePaths = $filePaths.Trim^(^) } else { $filePaths = '' }
    echo     if ($useInitial^) { $useInitial = $useInitial.Trim^(^) } else { $useInitial = '' }
    echo     if ($initialDelay^) { $initialDelay = $initialDelay.Trim^(^) } else { $initialDelay = '0' }
    echo     if ($useBetween^) { $useBetween = $useBetween.Trim^(^) } else { $useBetween = '' }
    echo     if ($betweenDelay^) { $betweenDelay = $betweenDelay.Trim^(^) } else { $betweenDelay = '0' }
    echo     $preset = @{ 'preset_name' = $displayName; 'file_paths' = $filePaths; 'use_initial' = $useInitial; 'initial_delay' = $initialDelay; 'use_between' = $useBetween; 'between_delay' = $betweenDelay; 'created' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'^) }
    echo     $json = $preset ^| ConvertTo-Json
    echo     $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    echo     [System.IO.File]::WriteAllText($outFile, $json, $utf8NoBom^)
    echo     Write-Output 'SUCCESS'
    echo } catch {
    echo     Write-Output 'ERROR'
    echo }
) > "!PS_SCRIPT!" 2>nul
if not exist "!PS_SCRIPT!" (
    echo.
    echo [ERROR] Failed to create preset script.
    exit /b 0
)

REM Write all settings to temp files for PowerShell
REM Note: file_paths variable contains the original comma-separated input
if not defined file_paths set "file_paths="
> "%TEMP%\preset_file_paths.txt" echo !file_paths!
> "%TEMP%\preset_use_initial.txt" echo !use_initial!
> "%TEMP%\preset_initial_delay.txt" echo !initial_delay!
> "%TEMP%\preset_use_between.txt" echo !use_between!
> "%TEMP%\preset_between_delay.txt" echo !between_delay!
> "%TEMP%\preset_output.txt" echo !PRESET_FILE!
> "%TEMP%\preset_display_name.txt" echo !preset_name!

REM Verify critical files exist
if not exist "%TEMP%\preset_output.txt" (
    echo.
    echo [ERROR] Failed to create output file path.
    exit /b 0
)
if not exist "%TEMP%\preset_display_name.txt" (
    echo.
    echo [ERROR] Failed to create display name file.
    exit /b 0
)

REM Run PowerShell script to save preset
set "save_result=ERROR"
if exist "!PS_SCRIPT!" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "!PS_SCRIPT!" > "%TEMP%\preset_save_result.txt" 2> "%TEMP%\preset_save_error.txt"
    if exist "%TEMP%\preset_save_result.txt" (
        for /f "delims=" %%A in ('type "%TEMP%\preset_save_result.txt"') do set "save_result=%%A"
    )
    if exist "%TEMP%\preset_save_error.txt" (
        for %%F in ("%TEMP%\preset_save_error.txt") do if %%~zF gtr 0 (
            del "%TEMP%\preset_save_error.txt" >nul 2>&1
        )
    )
    del "!PS_SCRIPT!" >nul 2>&1
) else (
    echo.
    echo [ERROR] Preset script file not found.
    exit /b 0
)

REM Cleanup temp files
del "%TEMP%\preset_file_paths.txt" "%TEMP%\preset_use_initial.txt" "%TEMP%\preset_initial_delay.txt" "%TEMP%\preset_use_between.txt" "%TEMP%\preset_between_delay.txt" "%TEMP%\preset_output.txt" "%TEMP%\preset_display_name.txt" "%TEMP%\preset_save_result.txt" >nul 2>&1

if "!save_result!"=="SUCCESS" (
    if exist "!PRESET_FILE!" (
        echo.
        echo Preset saved: !preset_name!
        exit /b 0
    ) else (
        echo.
        echo [ERROR] Failed to save preset.
        exit /b 0
    )
) else (
    echo.
    echo [ERROR] Failed to save preset.
    exit /b 0
)

REM ========================================
REM Save Preset Subroutine (Option 1)
REM ========================================
:save_preset_option1
echo.
echo Save this protocol as a preset?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
echo.
:retry_save_preset
set "save_preset_choice="
set /p "save_preset_choice=Enter Command: "

REM Handle menu return
if /i "!save_preset_choice!"=="M" exit /b 1

REM Normalize input
set "save_preset_choice=!save_preset_choice: =!"

if /i "!save_preset_choice!"=="Y" goto :get_preset_name
if /i "!save_preset_choice!"=="S" exit /b 0
echo Invalid. Enter Y, S, or M.
goto :retry_save_preset

:get_preset_name
echo.
echo Enter preset name
:retry_preset_name
set "preset_name="
set /p "preset_name=Preset Name: "
if /i "!preset_name!"=="M" exit /b 1

REM Basic validation - check if empty (simple check)
if not defined preset_name (
    echo Invalid. Try again.
    goto :retry_preset_name
)
if "!preset_name!"=="" (
    echo Invalid. Try again.
    goto :retry_preset_name
)

REM Validate no invalid filename characters (skip validation for now to avoid crashes)
REM Temporarily disabled - will sanitize filename instead
REM echo !preset_name! | findstr /r /c:"[\\/:*?\"<>|]" >nul 2>&1
REM if not errorlevel 1 (
REM     echo Invalid. Preset name cannot contain: \ / : * ? \" < > ^|
REM     goto :retry_preset_name
REM )

REM Create preset directory if it doesn't exist
set "PRESET_DIR=%~dp0Presets\Option1"
if not exist "!PRESET_DIR!" mkdir "!PRESET_DIR!" >nul 2>&1

REM Sanitize preset name for filename (simple batch method, avoid * character)
set "preset_filename=!preset_name!"
if defined preset_filename (
    set "preset_filename=!preset_filename:\=_!"
    set "preset_filename=!preset_filename:/=_!"
    set "preset_filename=!preset_filename::=_!"
    set "preset_filename=!preset_filename:?=_!"
    set "preset_filename=!preset_filename:"=_!"
    set "preset_filename=!preset_filename:<=_!"
    set "preset_filename=!preset_filename:>=_!"
    set "preset_filename=!preset_filename:|=_!"
    REM Skip * replacement - it causes issues in batch, will handle in PowerShell script
) else (
    set "preset_filename=preset"
)

REM Save preset to JSON using PowerShell
set "PRESET_FILE=!PRESET_DIR!\!preset_filename!.json"

REM Initialize variables if not set
if not defined folder_scan_option set "folder_scan_option=1"
if not defined exclude_options_choice set "exclude_options_choice="
if not defined exclude_paths set "exclude_paths="
if not defined exclude_names set "exclude_names="
if not defined exclude_keywords set "exclude_keywords="
if not defined use_initial set "use_initial="
if not defined initial_delay set "initial_delay=0"
if not defined use_between set "use_between="
if not defined between_delay set "between_delay=0"


REM Save preset directly using batch commands
set "save_result=ERROR"

REM Create JSON using batch echo commands with safe escaping
set "scan_path_safe=!scan_path!"
if defined scan_path set "scan_path_safe=!scan_path:\=\\!"

set "exclude_paths_safe=!exclude_paths!"
if defined exclude_paths if not "!exclude_paths!"=="" set "exclude_paths_safe=!exclude_paths:\=\\!"

REM Create JSON using batch echo commands
(
    echo {
    echo     "preset_name": "!preset_name!",
    echo     "file_ext": "!file_ext!",
    echo     "scan_path": "!scan_path_safe!",
    echo     "folder_scan_option": "!folder_scan_option!",
    echo     "exclude_options_choice": "!exclude_options_choice!",
    echo     "exclude_paths": "!exclude_paths_safe!",
    echo     "exclude_names": "!exclude_names!",
    echo     "exclude_keywords": "!exclude_keywords!",
    echo     "use_initial": "!use_initial!",
    echo     "initial_delay": "!initial_delay!",
    echo     "use_between": "!use_between!",
    echo     "between_delay": "!between_delay!",
    echo     "file_selection": "!selection!",
    echo     "created": "2025-12-12T!time!"
    echo }
) > "!PRESET_FILE!"

if exist "!PRESET_FILE!" (
    set "save_result=SUCCESS"
) else (
    set "save_result=ERROR"
)

REM Cleanup temp files (keep result file for now)
del "%TEMP%\preset_file_ext.txt" "%TEMP%\preset_scan_path.txt" "%TEMP%\preset_folder_scan_option.txt" "%TEMP%\preset_exclude_options.txt" "%TEMP%\preset_exclude_paths.txt" "%TEMP%\preset_exclude_names.txt" "%TEMP%\preset_exclude_keywords.txt" "%TEMP%\preset_use_initial.txt" "%TEMP%\preset_initial_delay.txt" "%TEMP%\preset_use_between.txt" "%TEMP%\preset_between_delay.txt" "%TEMP%\preset_file_selection.txt" "%TEMP%\preset_output.txt" "%TEMP%\preset_display_name.txt" "%TEMP%\preset_save_result.txt" >nul 2>&1

if "!save_result!"=="SUCCESS" (
    if exist "!PRESET_FILE!" (
        echo.
        echo Preset saved: !preset_name!
        exit /b 0
    ) else (
        echo.
        echo [ERROR] Failed to save preset.
        exit /b 0
    )
) else (
    echo.
    echo [ERROR] Failed to save preset.
    exit /b 0
)

REM ========================================
REM Presets for Folder-Based Protocol
REM ========================================
:presets_option1
cls
echo ========================================
echo    Presets for Folder-Based Protocol
echo ========================================
echo.

REM Create preset directory if it doesn't exist
set "PRESET_DIR=%~dp0Presets\Option1"
if not exist "!PRESET_DIR!" mkdir "!PRESET_DIR!" >nul 2>&1

REM List all presets
set "preset_count=0"
for %%F in ("!PRESET_DIR!\*.json") do set /a "preset_count+=1"

if !preset_count! EQU 0 (
    echo No presets saved yet.
    echo.
    pause
    goto :eof
)

REM Display presets using PowerShell
echo !PRESET_DIR!> "%TEMP%\preset_list_dir.txt"
set "PRESET_LIST_FILE=!TEMP_DIR!\preset_list_display.txt"
if exist "!PRESET_LIST_FILE!" del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
echo !PRESET_LIST_FILE!> "%TEMP%\preset_list_output.txt"
powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'preset_list_dir.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'preset_list_output.txt') -Raw; $dir = $dir.Trim(); $outFile = $outFile.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; $counter = 0; $lines = @(); foreach ($preset in $presets) { $counter++; try { $json = Get-Content -LiteralPath $preset.FullName -Raw -Encoding UTF8 | ConvertFrom-Json; $name = $json.preset_name; if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($preset.Name) }; $lines += ('p' + $counter.ToString() + ' - ' + $name) } catch { $lines += ('p' + $counter.ToString() + ' - ' + [System.IO.Path]::GetFileNameWithoutExtension($preset.Name)) } }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $lines, $utf8NoBom)"
del "%TEMP%\preset_list_dir.txt" "%TEMP%\preset_list_output.txt" >nul 2>&1

if exist "!PRESET_LIST_FILE!" (
    type "!PRESET_LIST_FILE!"
    del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
)

echo.
echo p#   - Run a Preset
echo p#p  - Preview a Preset
echo d#   - Delete a Preset
echo c#   - Create a .txt with the path's of a preset scan
echo M    - Back to Menu
echo.
:retry_preset_select
REM Refresh preset count to ensure it's current (important after deletions)
set "PRESET_DIR=%~dp0Presets\Option1"
set "preset_count=0"
for %%F in ("!PRESET_DIR!\*.json") do set /a "preset_count+=1"

set "preset_select="
set "preset_num="
set "preset_preview="
set "preset_delete="
set /p "preset_select=Enter Command: "

REM Normalize input
set "preset_select=!preset_select: =!"
set "preset_select=!preset_select!"

REM Handle menu return FIRST - must exit before parsing
if /i "!preset_select!"=="M" (
    REM Return to main menu - use goto :eof to match Option 2 behavior
    goto :eof
)

REM Parse command using PowerShell - get command type and number
echo !preset_select!> "%TEMP%\preset_cmd_input.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$input = Get-Content (Join-Path $env:TEMP 'preset_cmd_input.txt') -Raw; $input = $input.Trim().ToLower(); $result = ''; if ($input -eq 'm') { $result = 'MENU' } elseif ($input -match '^d(\d+)$') { $result = 'DELETE|' + $matches[1] } elseif ($input -match '^c(\d+)$') { $result = 'CREATE_TXT|' + $matches[1] } elseif ($input -match '^p(\d+)p$') { $result = 'PREVIEW|' + $matches[1] } elseif ($input -match '^p(\d+)$') { $result = 'RUN|' + $matches[1] } else { $result = 'INVALID' }; Write-Output $result"') do set "cmd_parsed=%%A"
del "%TEMP%\preset_cmd_input.txt" >nul 2>&1

REM Split command type and number
for /f "tokens=1,2 delims=|" %%A in ("!cmd_parsed!") do (
    set "cmd_type=%%A"
    set "cmd_num=%%B"
)

REM Handle MENU command
if "!cmd_type!"=="MENU" goto :eof

REM Handle DELETE command
if "!cmd_type!"=="DELETE" (
    REM Refresh preset count right before validation to ensure it's current
    set "PRESET_DIR=%~dp0Presets\Option1"
    set "preset_count=0"
    for %%F in ("!PRESET_DIR!\*.json") do set /a "preset_count+=1"
    
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter d1-d!preset_count! to delete, or M.
        goto :retry_preset_select
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter d1-d!preset_count! to delete, or M.
        goto :retry_preset_select
    )
    call :delete_preset_option1 !cmd_num!
    if errorlevel 2 (
        REM N pressed - cancelled
        goto :refresh_and_retry
    )
    if errorlevel 1 (
        REM M pressed - exit to main menu
        goto :eof
    )
    REM Delete succeeded
    echo Preset deleted successfully.
    timeout /t 1 /nobreak >nul
    goto :refresh_and_retry
)

REM Handle CREATE_TXT command
if "!cmd_type!"=="CREATE_TXT" (
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter c1-c!preset_count! to create .txt file, or M.
        goto :retry_preset_select
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter c1-c!preset_count! to create .txt file, or M.
        goto :retry_preset_select
    )
    call :create_preset_paths_txt_option1 !cmd_num!
    if errorlevel 1 (
        echo [ERROR] Failed to create .txt file.
        pause
        goto :retry_preset_select
    )
    echo.
    echo .txt file created successfully.
    timeout /t 2 /nobreak >nul
    goto :refresh_and_retry
)

REM Handle PREVIEW command
if "!cmd_type!"=="PREVIEW" (
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter p1p-p!preset_count!p to preview, or M.
        goto :retry_preset_select
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter p1p-p!preset_count!p to preview, or M.
        goto :retry_preset_select
    )
    call :preview_preset_option1 !cmd_num!
    if errorlevel 3 (
        REM E pressed - edit selection
        REM Clear run_preset_num to ensure edit takes priority over any lingering run request
        set "run_preset_num="
        set "edit_preset_num=!cmd_num!"
        exit /b 0
    )
if errorlevel 2 (
    REM S pressed - skip
    timeout /t 1 /nobreak >nul
    goto :refresh_and_retry
)
    if errorlevel 1 (
        REM M pressed - exit to main menu
        goto :eof
    )
    REM Y pressed - run_preset_num was set, now run it
    if defined run_preset_num (
        REM Exit presets_option1 first, then run handler
        exit /b 0
    )
    REM Fallback
    goto :retry_preset_select
)

REM Handle RUN command
if "!cmd_type!"=="RUN" (
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter p1-p!preset_count! to run, or M.
        goto :retry_preset_select
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter p1-p!preset_count! to run, or M.
        goto :retry_preset_select
    )
    REM Load and run preset - exit presets_option1 first, then run handler
    set "run_preset_num=!cmd_num!"
    exit /b 0
)

REM Handle INVALID command
if "!cmd_type!"=="INVALID" (
    echo Invalid. Enter p1-p!preset_count! to run, p1p-p!preset_count!p to preview, d1-d!preset_count! to delete, c1-c!preset_count! to create .txt, or M.
    goto :retry_preset_select
)

REM Fallback - shouldn't reach here
echo Invalid. Enter p1-p!preset_count! to run, p1p-p!preset_count!p to preview, d1-d!preset_count! to delete, c1-c!preset_count! to create .txt, or M.
goto :retry_preset_select

REM ========================================
REM Refresh Menu and Retry (Option 3)
REM ========================================
:refresh_and_retry
cls
echo ========================================
echo    Presets for Folder-Based Protocol
echo ========================================
echo.
REM Re-display preset list
set "PRESET_DIR=%~dp0Presets\Option1"
REM Count presets using batch command for consistency
set "preset_count=0"
for %%F in ("!PRESET_DIR!\*.json") do set /a "preset_count+=1"
if !preset_count! EQU 0 (
    echo No presets saved yet.
    echo.
    pause
    goto :main
)
echo !PRESET_DIR!> "%TEMP%\preset_list_dir.txt"
set "PRESET_LIST_FILE=!TEMP_DIR!\preset_list_display.txt"
if exist "!PRESET_LIST_FILE!" del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
echo !PRESET_LIST_FILE!> "%TEMP%\preset_list_output.txt"
powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'preset_list_dir.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'preset_list_output.txt') -Raw; $dir = $dir.Trim(); $outFile = $outFile.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; $counter = 0; $lines = @(); foreach ($preset in $presets) { $counter++; try { $json = Get-Content -LiteralPath $preset.FullName -Raw -Encoding UTF8 | ConvertFrom-Json; $name = $json.preset_name; if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($preset.Name) }; $lines += ('p' + $counter.ToString() + ' - ' + $name) } catch { $lines += ('p' + $counter.ToString() + ' - ' + [System.IO.Path]::GetFileNameWithoutExtension($preset.Name)) } }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $lines, $utf8NoBom)"
del "%TEMP%\preset_list_dir.txt" "%TEMP%\preset_list_output.txt" >nul 2>&1
if exist "!PRESET_LIST_FILE!" (
    type "!PRESET_LIST_FILE!"
    del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
)
echo.
echo p#   - Run a Preset
echo p#p  - Preview a Preset
echo d#   - Delete a Preset
echo c#   - Create a .txt with the path's of a preset scan
echo M    - Back to Menu
echo.
goto :retry_preset_select

REM ========================================
REM Run Preset Handler (Option 3)
REM ========================================
:run_preset_handler
REM Preserve edited values if using edited selection
if defined USE_EDITED_SELECTION (
    set "EDITED_file_ext=!file_ext!"
    set "EDITED_scan_path=!scan_path!"
    set "EDITED_folder_scan_option=!folder_scan_option!"
    set "EDITED_exclude_options_choice=!exclude_options_choice!"
    set "EDITED_exclude_paths=!exclude_paths!"
    set "EDITED_exclude_names=!exclude_names!"
    set "EDITED_exclude_keywords=!exclude_keywords!"
    set "EDITED_use_initial=!use_initial!"
    set "EDITED_initial_delay=!initial_delay!"
    set "EDITED_use_between=!use_between!"
    set "EDITED_between_delay=!between_delay!"
)

call :load_preset_option1 !run_preset_num!
if errorlevel 1 (
    pause
    goto :eof
)

REM Restore edited values if using edited selection
if defined USE_EDITED_SELECTION (
    set "file_ext=!EDITED_file_ext!"
    set "scan_path=!EDITED_scan_path!"
    set "folder_scan_option=!EDITED_folder_scan_option!"
    set "exclude_options_choice=!EDITED_exclude_options_choice!"
    set "exclude_paths=!EDITED_exclude_paths!"
    set "exclude_names=!EDITED_exclude_names!"
    set "exclude_keywords=!EDITED_exclude_keywords!"
    set "use_initial=!EDITED_use_initial!"
    set "initial_delay=!EDITED_initial_delay!"
    set "use_between=!EDITED_use_between!"
    set "between_delay=!EDITED_between_delay!"
)

set "AUTO_RUN_PRESET=1"
echo.
echo [SCAN] Loading preset and running protocol...
echo.
call :scan_files
if errorlevel 1 (
    echo [ERROR] Failed to scan files.
    pause
    goto :eof
)
REM Check if we should use edited selection (before display)
if defined USE_EDITED_SELECTION (
    if defined EDITED_SELECTED_FILES (
        set "SELECTED_FILES=!EDITED_SELECTED_FILES!"
        REM Display selected files instead of all files
        if exist "!SELECTED_FILES!" (
            REM Count selected files using temp file
            set "selected_count=0"
            echo !SELECTED_FILES!> "%TEMP%\count_selected.txt"
            for /f %%A in ('powershell -NoProfile -Command "$listFile = Get-Content (Join-Path $env:TEMP 'count_selected.txt') -Raw; $listFile = $listFile.Trim(); $files = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($files) { Write-Output $files.Count } else { Write-Output 0 }"') do set "selected_count=%%A"
            del "%TEMP%\count_selected.txt" >nul 2>&1
            
            REM Default to 0 if empty
            if not defined selected_count set "selected_count=0"
            if "!selected_count!"=="" set "selected_count=0"
            
            REM Count excluded files (total - selected)
            set "excluded_count=0"
            if exist "!LAST_FILES!" (
                set "total_count=0"
                echo !LAST_FILES!> "%TEMP%\count_total.txt"
                for /f %%A in ('powershell -NoProfile -Command "$listFile = Get-Content (Join-Path $env:TEMP 'count_total.txt') -Raw; $listFile = $listFile.Trim(); $files = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($files) { Write-Output $files.Count } else { Write-Output 0 }"') do set "total_count=%%A"
                del "%TEMP%\count_total.txt" >nul 2>&1
                
                REM Default to 0 if empty
                if not defined total_count set "total_count=0"
                if "!total_count!"=="" set "total_count=0"
                
                REM Calculate excluded count
                if defined total_count if defined selected_count (
                    set /a "excluded_count=!total_count! - !selected_count!" 2>nul
                    if errorlevel 1 set "excluded_count=0"
                )
            )
            
            REM Display selected files with numbering using PowerShell
            set "DISPLAY_FILE=!TEMP_DIR!\FileLaunchSequencer_selected_display.txt"
            if exist "!DISPLAY_FILE!" del /f /q "!DISPLAY_FILE!"
            
            REM Write scan path to temp file for PowerShell
            echo !scan_path!> "%TEMP%\selected_display_base.txt"
            echo !SELECTED_FILES!> "%TEMP%\selected_display_list.txt"
            echo !DISPLAY_FILE!> "%TEMP%\selected_display_output.txt"
            
            REM Use PowerShell to format and display selected files
            powershell -NoProfile -Command "$base = Get-Content (Join-Path $env:TEMP 'selected_display_base.txt') -Raw; $listFile = Get-Content (Join-Path $env:TEMP 'selected_display_list.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'selected_display_output.txt') -Raw; $base = $base.Trim().TrimEnd('\'); $listFile = $listFile.Trim(); $outFile = $outFile.Trim(); $fileArray = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($fileArray) { $counter = 0; $baseName = Split-Path -Leaf $base; $formattedLines = @(); foreach ($file in $fileArray) { $counter++; $file = $file.Trim(); $rel = $file.Replace($base, ''); if ($rel -match '^\\') { $rel = $rel.Substring(1) }; $formattedLines += ($counter.ToString() + '. \' + $baseName + '\' + $rel) }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $formattedLines, $utf8NoBom) }"
            
            REM Cleanup temp files
            del "%TEMP%\selected_display_base.txt" "%TEMP%\selected_display_list.txt" "%TEMP%\selected_display_output.txt" >nul 2>&1
            
            REM Display the formatted file list
            echo.
            echo FILE SCAN:
            echo.
            if exist "!DISPLAY_FILE!" (
                type "!DISPLAY_FILE!"
                del /f /q "!DISPLAY_FILE!" >nul 2>&1
            )
            
            echo.
            echo Found !selected_count! !file_ext! Files Found
            echo Found !excluded_count! !file_ext! Files to Exclude
            echo.
        ) else (
            REM Fallback to normal display if selected files not found
            call :display_files
            if errorlevel 1 (
                echo [ERROR] Failed to display files.
                pause
                goto :eof
            )
            set "SELECTED_FILES=!LAST_FILES!"
            echo.
            echo [SCAN] Auto-selecting all files for preset...
        )
    ) else (
        REM No edited selection file, use normal display
        call :display_files
        if errorlevel 1 (
            echo [ERROR] Failed to display files.
            pause
            goto :eof
        )
        set "SELECTED_FILES=!LAST_FILES!"
        echo.
        echo [SCAN] Auto-selecting all files for preset...
    )
) else (
    REM Check if preset has a saved selection
    call :load_preset_selection_option1 !run_preset_num!
    if errorlevel 0 (
        if defined PRESET_SELECTION (
            REM Use saved selection from preset
            echo !PRESET_SELECTION!> "%TEMP%\preset_sel_input.txt"
            echo !LAST_FILES!> "%TEMP%\preset_sel_list.txt"
            set "SELECTED_FILES=!TEMP_DIR!\FileLaunchSequencer_preset_selected.txt"
            if exist "!SELECTED_FILES!" del /f /q "!SELECTED_FILES!"
            echo !SELECTED_FILES!> "%TEMP%\preset_sel_output.txt"
            powershell -NoProfile -Command "$selection = Get-Content (Join-Path $env:TEMP 'preset_sel_input.txt') -Raw; $listFile = Get-Content (Join-Path $env:TEMP 'preset_sel_list.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'preset_sel_output.txt') -Raw; $selection = $selection.Trim(); $listFile = $listFile.Trim(); $outFile = $outFile.Trim(); $files = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if (-not $files -or $files.Count -eq 0) { exit 1 }; $selected = @(); $isAll = $selection -match '^all'; if ($isAll) { $selected = $files; $parts = $selection -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^-' }; foreach ($part in $parts) { $part = $part.TrimStart('-'); if ($part -match '^(\d+)-(\d+)$') { $start = [int]$matches[1]; $end = [int]$matches[2]; if ($start -ge 1 -and $end -le $files.Count -and $start -le $end) { for ($i = $start-1; $i -le $end-1; $i++) { $selected = $selected | Where-Object { $_ -ne $files[$i] } } } } elseif ($part -match '^\d+$') { $num = [int]$part; if ($num -ge 1 -and $num -le $files.Count) { $selected = $selected | Where-Object { $_ -ne $files[$num-1] } } } } } else { $parts = $selection -split ',' | ForEach-Object { $_.Trim() }; foreach ($part in $parts) { if ($part -match '^(\d+)-(\d+)$') { $start = [int]$matches[1]; $end = [int]$matches[2]; if ($start -ge 1 -and $end -le $files.Count -and $start -le $end) { $selected += $files[($start-1)..($end-1)] } } elseif ($part -match '^\d+$') { $num = [int]$part; if ($num -ge 1 -and $num -le $files.Count) { $selected += $files[$num-1] } } }; $selected = $selected | Select-Object -Unique }; if ($selected.Count -eq 0) { exit 1 }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $selected, $utf8NoBom)"
            del "%TEMP%\preset_sel_input.txt" "%TEMP%\preset_sel_list.txt" "%TEMP%\preset_sel_output.txt" >nul 2>&1
            if exist "!SELECTED_FILES!" (
                echo.
                echo [SCAN] Using saved selection from preset...
            ) else (
                set "SELECTED_FILES=!LAST_FILES!"
                echo.
                echo [SCAN] Auto-selecting all files for preset...
            )
        ) else (
            set "SELECTED_FILES=!LAST_FILES!"
            echo.
            echo [SCAN] Auto-selecting all files for preset...
        )
    ) else (
        REM Not using edited selection - display all files normally
        call :display_files
        if errorlevel 1 (
            echo [ERROR] Failed to display files.
            pause
            goto :eof
        )
        set "SELECTED_FILES=!LAST_FILES!"
        echo.
        echo [SCAN] Auto-selecting all files for preset...
    )
)
goto :option1_open_files

REM ========================================
REM Load Preset Subroutine (Option 1)
REM ========================================
:load_preset_option1
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option1"

REM Get preset file by index (sorted by CreationTime to match display)
set "preset_file="
echo !PRESET_DIR!> "%TEMP%\load_preset_dir.txt"
echo !preset_index!> "%TEMP%\load_preset_index.txt"
for /f "delims=" %%F in ('powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'load_preset_dir.txt') -Raw; $index = Get-Content (Join-Path $env:TEMP 'load_preset_index.txt') -Raw; $dir = $dir.Trim(); $index = [int]$index.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; if ($presets -and $index -le $presets.Count) { Write-Output $presets[$index-1].FullName }"') do set "preset_file=%%F"
del "%TEMP%\load_preset_dir.txt" "%TEMP%\load_preset_index.txt" >nul 2>&1

if not defined preset_file (
    echo [ERROR] Preset not found.
    exit /b 1
)

REM Load preset using PowerShell
echo !preset_file!> "%TEMP%\load_preset_file.txt"
powershell -NoProfile -Command "try { $presetFile = Get-Content (Join-Path $env:TEMP 'load_preset_file.txt') -Raw; $presetFile = $presetFile.Trim(); $json = Get-Content -LiteralPath $presetFile -Raw -Encoding UTF8 | ConvertFrom-Json; $presetName = if ($json.preset_name) { $json.preset_name } else { '' }; $fileExt = if ($json.file_ext) { $json.file_ext } else { '.txt' }; $scanPath = if ($json.scan_path) { $json.scan_path } else { 'C:\' }; $folderScanOption = if ($json.folder_scan_option) { $json.folder_scan_option } else { '1' }; $excludeOptions = if ($json.exclude_options_choice) { $json.exclude_options_choice } else { '' }; $excludePaths = if ($json.exclude_paths) { $json.exclude_paths } else { '' }; $excludeNames = if ($json.exclude_names) { $json.exclude_names } else { '' }; $excludeKeywords = if ($json.exclude_keywords) { $json.exclude_keywords } else { '' }; $useInitial = if ($json.use_initial) { $json.use_initial } else { '' }; $initialDelay = if ($json.initial_delay) { $json.initial_delay } else { '0' }; $useBetween = if ($json.use_between) { $json.use_between } else { '' }; $betweenDelay = if ($json.between_delay) { $json.between_delay } else { '0' }; Write-Output ($presetName + '|' + $fileExt + '|' + $scanPath + '|' + $folderScanOption + '|' + $excludeOptions + '|' + $excludePaths + '|' + $excludeNames + '|' + $excludeKeywords + '|' + $useInitial + '|' + $initialDelay + '|' + $useBetween + '|' + $betweenDelay) } catch { Write-Output 'Error|Error|C:\|1|||||||0||0' }" > "%TEMP%\load_preset_data.txt"

REM Parse loaded data
set /p "load_line="<"%TEMP%\load_preset_data.txt"

REM Split the line by pipe delimiter
set "token_num=0"
for %%T in ("!load_line:|= " "!") do (
    set /a "token_num+=1"
    set "token!token_num!=%%~T"
)

REM Assign to variables and trim trailing spaces
set "preset_name=!token1!"
set "file_ext=!token2!"
set "scan_path=!token3!"
set "folder_scan_option=!token4!"
set "exclude_options_choice=!token5!"
set "exclude_paths=!token6!"
set "exclude_names=!token7!"
set "exclude_keywords=!token8!"
set "use_initial=!token9!"
set "initial_delay=!token10!"
set "use_between=!token11!"
set "between_delay=!token12!"

REM Trim trailing spaces from key variables
set "temp_use_initial=!use_initial!"
:trim_use_initial
if "!temp_use_initial:~-1!"==" " (
    set "temp_use_initial=!temp_use_initial:~0,-1!"
    goto :trim_use_initial
)
set "use_initial=!temp_use_initial!"

set "temp_initial_delay=!initial_delay!"
:trim_initial_delay
if "!temp_initial_delay:~-1!"==" " (
    set "temp_initial_delay=!temp_initial_delay:~0,-1!"
    goto :trim_initial_delay
)
set "initial_delay=!temp_initial_delay!"

set "temp_use_between=!use_between!"
:trim_use_between
if "!temp_use_between:~-1!"==" " (
    set "temp_use_between=!temp_use_between:~0,-1!"
    goto :trim_use_between
)
set "use_between=!temp_use_between!"

set "temp_between_delay=!between_delay!"
:trim_between_delay
if "!temp_between_delay:~-1!"==" " (
    set "temp_between_delay=!temp_between_delay:~0,-1!"
    goto :trim_between_delay
)
set "between_delay=!temp_between_delay!"

REM Handle empty values - PowerShell returns empty strings, batch treats them as spaces sometimes
if "!exclude_paths!"=="" set "exclude_paths="
if "!exclude_names!"=="" set "exclude_names="
if "!exclude_keywords!"=="" set "exclude_keywords="

del "%TEMP%\load_preset_file.txt" "%TEMP%\load_preset_data.txt" >nul 2>&1

REM Validate target folder exists
if not exist "!scan_path!\" (
    if not exist "!scan_path!" (
        echo.
        echo [ERROR] Target folder no longer exists: !scan_path!
        pause
        exit /b 1
    )
)

REM Ensure all variables are properly set
if not defined file_ext set "file_ext="
if not defined scan_path set "scan_path="
if not defined folder_scan_option set "folder_scan_option=1"
if not defined exclude_options_choice set "exclude_options_choice="
if not defined exclude_paths set "exclude_paths="
if not defined exclude_names set "exclude_names="
if not defined exclude_keywords set "exclude_keywords="
if not defined use_initial set "use_initial="
if not defined initial_delay set "initial_delay=0"
if not defined use_between set "use_between="
if not defined between_delay set "between_delay=0"

exit /b 0

REM ========================================
REM Preview Preset Subroutine (Option 1)
REM ========================================
:preview_preset_option1
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option1"

REM Get preset file by index (sorted by CreationTime to match display)
set "preset_file="
echo !PRESET_DIR!> "%TEMP%\preview_preset_dir.txt"
echo !preset_index!> "%TEMP%\preview_preset_index.txt"
for /f "delims=" %%F in ('powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'preview_preset_dir.txt') -Raw; $index = Get-Content (Join-Path $env:TEMP 'preview_preset_index.txt') -Raw; $dir = $dir.Trim(); $index = [int]$index.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; if ($presets -and $index -le $presets.Count) { Write-Output $presets[$index-1].FullName }"') do set "preset_file=%%F"
del "%TEMP%\preview_preset_dir.txt" "%TEMP%\preview_preset_index.txt" >nul 2>&1

if not defined preset_file (
    echo [ERROR] Preset not found.
    pause
    exit /b 1
)

REM Load preset settings first
call :load_preset_option1 !preset_index!
if errorlevel 1 (
    echo [ERROR] Failed to load preset.
    pause
    exit /b 1
)

REM Display preset summary first using PowerShell script file
set "PREVIEW_SCRIPT=!TEMP_DIR!\preview_preset.ps1"
echo !preset_file!> "%TEMP%\preview_preset_file.txt"
(
    echo $presetFile = Get-Content (Join-Path $env:TEMP 'preview_preset_file.txt'^) -Raw
    echo $presetFile = $presetFile.Trim^(^)
    echo $json = Get-Content -LiteralPath $presetFile -Raw -Encoding UTF8 ^| ConvertFrom-Json
    echo Write-Output ''
    echo Write-Output '================'
    echo Write-Output ''
    echo $created = if ($json.created^) { $json.created } else { 'Unknown' }
    echo if ($created -match 'T'^) {
    echo     try {
    echo         $dateObj = [DateTime]::Parse^($created.Replace^('T', ' '^)^)
    echo         $created = $dateObj.ToString^('yyyy-MM-dd h:mmtt'^)
    echo     } catch { }
    echo }
    echo Write-Output ('Created: ' + $created^)
    echo Write-Output ''
    echo $presetName = if ($json.preset_name^) { $json.preset_name } else { 'N/A' }
    echo Write-Output ('PRESET NAME: ' + $presetName^)
    echo $fileExt = if ($json.file_ext^) { $json.file_ext } else { 'N/A' }
    echo Write-Output ('FILE FORMAT: ' + $fileExt^)
    echo $scanPath = if ($json.scan_path^) { $json.scan_path } else { 'N/A' }
    echo Write-Output ('TARGET FOLDER: ' + $scanPath^)
    echo $folderScanOption = if ($json.folder_scan_option^) { $json.folder_scan_option } else { '1' }
    echo $folderScanText = if ($folderScanOption -eq '1'^) { 'All Subfolders' } elseif ($folderScanOption -eq '2'^) { 'First-Level Subfolders' } else { 'Target Folder Only' }
    echo Write-Output ('FOLDER SCAN: ' + $folderScanText^)
    echo Write-Output ''
    echo Write-Output 'EXCLUSION SETTINGS:'
    echo $excludeOptions = if ($json.exclude_options_choice -eq 'y'^) { 'Yes' } elseif ($json.exclude_options_choice -eq 's'^) { 'Skipped' } else { 'No' }
    echo Write-Output ('  Use Exclusions: ' + $excludeOptions^)
    echo $excludePaths = if ($json.exclude_paths -and $json.exclude_paths.Trim^(^) -ne '' -and $json.exclude_paths.Trim^(^) -ne 'ECHO is off.'^) { $json.exclude_paths } else { 'Skipped' }
    echo Write-Output ('  Exclude Paths: ' + $excludePaths^)
    echo $excludeNames = if ($json.exclude_names -and $json.exclude_names.Trim^(^) -ne '' -and $json.exclude_names.Trim^(^) -ne 'ECHO is off.'^) { $json.exclude_names } else { 'Skipped' }
    echo Write-Output ('  Exclude Names: ' + $excludeNames^)
    echo $excludeKeywords = if ($json.exclude_keywords -and $json.exclude_keywords.Trim^(^) -ne '' -and $json.exclude_keywords.Trim^(^) -ne 'ECHO is off.'^) { $json.exclude_keywords } else { 'Skipped' }
    echo Write-Output ('  Exclude Keywords: ' + $excludeKeywords^)
    echo Write-Output ''
    echo Write-Output 'DELAY SETTINGS:'
    echo $initialDelay = if ($json.use_initial -eq 'y'^) { $json.initial_delay + ' seconds' } elseif ($json.use_initial -eq 's'^) { 'Skipped' } else { 'Disabled' }
    echo Write-Output ('  Initial Delay: ' + $initialDelay^)
    echo $betweenDelay = if ($json.use_between -eq 'y'^) { $json.between_delay + ' seconds' } elseif ($json.use_between -eq 's'^) { 'Skipped' } else { 'Disabled' }
    echo Write-Output ('  Between Files Delay: ' + $betweenDelay^)
    echo Write-Output ''
) > "!PREVIEW_SCRIPT!"

powershell -NoProfile -ExecutionPolicy Bypass -File "!PREVIEW_SCRIPT!"
del "!PREVIEW_SCRIPT!" "%TEMP%\preview_preset_file.txt" >nul 2>&1

REM Scan files to get scan totals
call :scan_files
if errorlevel 1 (
    echo [ERROR] Failed to scan files.
    pause
    exit /b 1
)

REM Display files (shows scan totals)
call :display_files
if errorlevel 1 (
    echo [ERROR] Failed to display files.
    pause
    exit /b 1
)

REM Load saved selection if it exists
set "PRESET_SELECTION="
call :load_preset_selection_option1 !preset_index! 2>nul
if errorlevel 0 (
    if defined PRESET_SELECTION (
        echo Selection: !PRESET_SELECTION!
    )
)

echo.
echo ================

echo.
echo Do you want to run this preset?
echo   Y - Yes
echo   E - Edit Selection
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_preview_action
set "preview_action="
set /p "preview_action=Enter Command: "

REM Handle menu return
if /i "!preview_action!"=="M" exit /b 1

REM Normalize input
set "preview_action=!preview_action: =!"

if /i "!preview_action!"=="Y" (
    REM Run the preset - set variable and jump to handler (preset_index is already set from function parameter)
    set "run_preset_num=!preset_index!"
    exit /b 0
)
if /i "!preview_action!"=="E" (
    REM Edit selection - set variable to trigger edit flow
    set "edit_preset_num=!preset_index!"
    exit /b 3
)
if /i "!preview_action!"=="S" (
    REM Skip - return to preset menu (return code 2 to distinguish from other returns)
    exit /b 2
)
echo Invalid. Enter Y, E, S, or M.
goto :retry_preview_action

REM ========================================
REM Create .txt File with Preset Paths (Option 1)
REM ========================================
:create_preset_paths_txt_option1
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option1"

REM Find the preset file by counting files
set "counter=0"
for %%F in ("!PRESET_DIR!\*.json") do (
    set /a "counter+=1"
    if !counter! equ !preset_index! (
        set "preset_file=%%F"
        goto :found_preset
    )
)
:found_preset

if not defined preset_file (
    echo [ERROR] Preset not found.
    exit /b 1
)

REM Load the preset to get settings
call :load_preset_option1 !preset_index!
if errorlevel 1 (
    echo [ERROR] Failed to load preset.
    exit /b 1
)

REM Scan files silently
set "SILENT_SCAN=1"
call :scan_files
set "SILENT_SCAN="
if errorlevel 1 (
    echo [ERROR] Failed to scan files.
    exit /b 1
)

REM Get preset name safely
set "preset_name="
for /f "tokens=2 delims=:" %%i in ('findstr /c:"preset_name" "!preset_file!"') do (
    set "line=%%i"
    set "line=!line:"=!"
    set "line=!line:,=!"
    for /f "tokens=* delims= " %%j in ("!line!") do set "preset_name=%%j"
)
if "!preset_name!"=="" set "preset_name=preset"

REM Sanitize filename (simple batch approach)
set "safe_name=!preset_name!"
REM Replace characters that are safe to replace
set "safe_name=!safe_name: =_!"
set "safe_name=!safe_name:\=_!"
set "safe_name=!safe_name:/=_!"
set "safe_name=!safe_name::=_!"
set "safe_name=!safe_name:?=_!"
set "safe_name=!safe_name:"=_!"
set "safe_name=!safe_name:<=_!"
set "safe_name=!safe_name:>=_!"
set "safe_name=!safe_name:|=_!"
REM Skip asterisk replacement to avoid batch wildcard issues

REM Get timestamp
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'h_mmtt-dd.MM.yyyy'"') do set "timestamp=%%i"

REM Create folder and file
if not exist "%~dp0Exported-Paths" mkdir "%~dp0Exported-Paths"
set "output_file=%~dp0Exported-Paths\!safe_name!-paths.-!timestamp!.txt"

REM Create the content
echo ================ > "!output_file!"
echo. >> "!output_file!"
echo Created: !date! !time! >> "!output_file!"
echo. >> "!output_file!"
echo PRESET NAME: !preset_name! >> "!output_file!"
echo FILE FORMAT: !file_ext! >> "!output_file!"
echo TARGET FOLDER: !scan_path! >> "!output_file!"
echo. >> "!output_file!"
echo FILE SCAN: >> "!output_file!"
echo. >> "!output_file!"

REM Add file paths
if exist "!LAST_FILES!" (
    type "!LAST_FILES!" >> "!output_file!"
) else (
    echo No files found. >> "!output_file!"
)

echo. >> "!output_file!"
echo ================ >> "!output_file!"

if exist "!output_file!" (
    echo.
    echo .txt file created: !output_file!
    exit /b 0
) else (
    echo [ERROR] Failed to create .txt file.
    exit /b 1
)

:preview_menu_return
exit /b 0

REM ========================================
REM Edit Preset Selection (Option 1)
REM ========================================
:edit_preset_selection_option1
set "preset_index=!edit_preset_num!"

REM Load the preset
call :load_preset_option1 !preset_index!
if errorlevel 1 (
    pause
    exit /b 1
)

REM Store the preset file path for updating
set "preset_file_path=!preset_file!"

REM Store original values for comparison
set "original_file_ext=!file_ext!"
set "original_scan_path=!scan_path!"
set "original_folder_scan_option=!folder_scan_option!"
set "original_exclude_options_choice=!exclude_options_choice!"
set "original_exclude_paths=!exclude_paths!"
set "original_exclude_names=!exclude_names!"
set "original_exclude_keywords=!exclude_keywords!"
set "original_use_initial=!use_initial!"
set "original_initial_delay=!initial_delay!"
set "original_use_between=!use_between!"
set "original_between_delay=!between_delay!"

REM Ask about changing target folder
echo.
echo Do you want to change the target folder?
echo Currently: !scan_path!
echo   Y - Yes
echo   S - Skip
echo.
:retry_change_folder
set "change_folder="
set /p "change_folder=Enter Command: "

if /i "!change_folder!"=="M" goto :preview_menu_return
set "change_folder=!change_folder: =!"

if /i "!change_folder!"=="Y" (
    goto :get_new_folder
)
if /i "!change_folder!"=="S" (
    goto :after_folder_change
)
echo Invalid. Enter Y, S, or M.
goto :retry_change_folder

:get_new_folder
echo.
echo Enter the path of the folder where all !file_ext! files will be opened
echo.
:retry_new_folder
set /p "input_new_scan_path=New Target Folder Path: "

REM Handle menu return
if /i "!input_new_scan_path!"=="M" goto :preview_menu_return

REM Normalize input (remove surrounding quotes if present)
set "new_scan_path=!input_new_scan_path:"=!"

REM Trim leading/trailing spaces
for /f "tokens=* delims= " %%A in ("!new_scan_path!") do set "new_scan_path=%%A"

REM Validate folder exists
if exist "!new_scan_path!\" (
    set "new_scan_path=!new_scan_path!\"
) else if exist "!new_scan_path!" (
    set "new_scan_path=!new_scan_path!\"
) else (
    echo Invalid. Try again.
    goto :retry_new_folder
)

REM Update scan_path variable
set "scan_path=!new_scan_path!"
goto :after_folder_change

:after_folder_change

:after_folder_change

REM Ask about changing file format
echo.
echo Do you want to change the file format to open?
echo Currently: !file_ext!
echo   Y - Yes
echo   S - Skip
echo.
:retry_change_format
set "change_format="
set /p "change_format=Enter Command: "

if /i "!change_format!"=="M" goto :preview_menu_return
set "change_format=!change_format: =!"

if /i "!change_format!"=="Y" (
    goto :get_new_format
)
if /i "!change_format!"=="S" (
    goto :after_format_change
)
echo Invalid. Enter Y, S, or M.
goto :retry_change_format

:get_new_format
echo.
echo Enter the new file format for the files you want to auto open
echo Examples: prproj  ^|  .prproj  ^|  .txt   ^|   ...etc...
echo.
echo M - Back to Menu
echo.
set /p "input_new_file_ext=New File Format: "

REM Handle menu return
if /i "!input_new_file_ext!"=="M" goto :preview_menu_return

REM Normalize input (trim spaces)
set "new_file_ext=!input_new_file_ext: =!"

REM Empty input returns to menu
if "!new_file_ext!"=="" goto :get_new_format

REM Normalize format (add dot if missing)
if "!new_file_ext:~0,1!" NEQ "." set "new_file_ext=.!new_file_ext!"

REM Update file_ext variable
set "file_ext=!new_file_ext!"
goto :after_format_change

:after_format_change

REM Check if anything changed that requires a new scan
set "needs_rescan=0"
if not "!file_ext!"=="!original_file_ext!" set "needs_rescan=1"
if not "!scan_path!"=="!original_scan_path!" set "needs_rescan=1"

if !needs_rescan! EQU 1 (
    REM Do a new scan with updated settings
    echo.
    echo [SCAN] Scanning with updated settings...
    set "SILENT_SCAN=1"
    call :scan_files
    set "SILENT_SCAN="
    if errorlevel 1 (
        echo.
        echo No files found with that format. Please try a different file extension.
        echo.
        goto :get_new_format
    )

    REM Display the new scan results
    call :display_files
    if errorlevel 1 (
        echo [ERROR] Failed to display files.
        pause
        exit /b 1
    )
) else (
    REM No changes, use original scan
    set "SILENT_SCAN=1"
    call :scan_files
    set "SILENT_SCAN="
    if errorlevel 1 (
        echo [ERROR] Failed to scan files.
        pause
        exit /b 1
    )
)

REM Allow editing the selection
echo.
echo Do you want to edit the open sequence of these?
if defined file_selection (
    echo Currently: !file_selection!
) else (
    echo Currently: all
)
echo.
echo Y - Yes
echo S - Skip
echo M - Back to Menu
echo.
:retry_edit_selection
set "edit_selection_choice="
set /p "edit_selection_choice=Enter Command: "

if /i "!edit_selection_choice!"=="M" goto :preview_menu_return
set "edit_selection_choice=!edit_selection_choice: =!"

if /i "!edit_selection_choice!"=="Y" (
    REM Yes - proceed to selection input
    goto :do_selection_edit
)
if /i "!edit_selection_choice!"=="S" (
    REM Skip editing - use existing selection or "all"
    if defined file_selection (
        set "selection=!file_selection!"
    ) else (
        set "selection=all"
    )
    REM Process the selection so SELECTED_FILES exists for preview
    goto :process_selection
)
echo Invalid. Enter Y, S, or M.
goto :retry_edit_selection

:do_selection_edit
echo.
echo File Selection Examples: (all, 1-11, 8-22, 1,4,20,11,14,15,)
echo If you use 'all' you can exclude like (all,-13,-24,-30-40)
echo.
set /p "input_selection=New Selection: "

REM Handle menu return
if /i "!input_selection!"=="M" goto :after_selection_edit

REM Trim spaces and assign to selection variable
set "selection=!input_selection: =!"

REM If empty, default to "all"
if "!selection!"=="" set "selection=all"

REM Process the selection using the same logic as :ask_selection
goto :process_selection

:process_selection

REM Check for "all" command (with or without exclusions)
set "is_all=0"
set "selection_lower=!selection!"
if /i "!selection!"=="all" (
    set "is_all=1"
    set "has_exclusions=0"
) else (
    REM Check if starts with "all" (case-insensitive)
    for /f "delims=" %%A in ('powershell -NoProfile -Command "Write-Output ('!selection!'.ToLower().StartsWith('all'))"') do set "starts_all=%%A"
    if "!starts_all!"=="True" (
        set "is_all=1"
        set "has_exclusions=1"
    )
)

REM If "all" without exclusions, use full list
if !is_all! EQU 1 (
    if !has_exclusions! EQU 0 (
        set "SELECTED_FILES=!LAST_FILES!"
        goto :selection_processed
    )
)

REM Simplified selection validation for edit flow
set "selection=!input_selection!"

REM Basic validation - accept common patterns
if "!selection!"=="" (
    echo Invalid selection. Try again.
    goto :do_selection_edit
)

REM Check for "all" (exact match)
if /i "!selection!"=="all" goto :selection_valid

REM Check for single numbers (basic check)
echo !selection! | findstr /r /c:"^[0-9]" >nul 2>&1
if not errorlevel 1 (
    REM It's a number, check if it's just a single number or a range
    echo !selection! | findstr /r /c:"^-[0-9]" >nul 2>&1
    if errorlevel 1 goto :selection_valid
)

REM Check for ranges like "1-4" (contains hyphen)
echo !selection! | findstr /r /c:"-" >nul 2>&1
if not errorlevel 1 goto :selection_valid

REM Check for exclusions (contains comma)
echo !selection! | findstr /r /c:"," >nul 2>&1
if not errorlevel 1 goto :selection_valid

REM If we get here, invalid syntax
echo Invalid selection. Try again.
goto :do_selection_edit

:selection_valid
REM Selection syntax is valid

REM Process the selection to create SELECTED_FILES
REM Set up temp file for selected files
set "SELECTED_FILES=!TEMP_DIR!\FileLaunchSequencer_edited_selected.txt"
if exist "!SELECTED_FILES!" del /f /q "!SELECTED_FILES!"

REM Write parameters to temp files to avoid apostrophe issues
echo !selection!> "%TEMP%\edit_sel_input.txt"
echo !LAST_FILES!> "%TEMP%\edit_sel_list.txt"
echo !SELECTED_FILES!> "%TEMP%\edit_sel_output.txt"

REM Process selection using PowerShell (handles "all" with exclusions and regular selections)
powershell -NoProfile -Command "$selection = Get-Content (Join-Path $env:TEMP 'edit_sel_input.txt') -Raw; $listFile = Get-Content (Join-Path $env:TEMP 'edit_sel_list.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'edit_sel_output.txt') -Raw; $selection = $selection.Trim(); $listFile = $listFile.Trim(); $outFile = $outFile.Trim(); $files = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if (-not $files -or $files.Count -eq 0) { exit 1 }; $selected = @(); $isAll = $selection -match '^all'; if ($isAll) { $selected = $files; $parts = $selection -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^-' }; foreach ($part in $parts) { $part = $part.TrimStart('-'); if ($part -match '^(\d+)-(\d+)$') { $start = [int]$matches[1]; $end = [int]$matches[2]; if ($start -ge 1 -and $end -le $files.Count -and $start -le $end) { for ($i = $start-1; $i -le $end-1; $i++) { $selected = $selected | Where-Object { $_ -ne $files[$i] } } } } elseif ($part -match '^\d+$') { $num = [int]$part; if ($num -ge 1 -and $num -le $files.Count) { $selected = $selected | Where-Object { $_ -ne $files[$num-1] } } } } } else { $parts = $selection -split ',' | ForEach-Object { $_.Trim() }; foreach ($part in $parts) { if ($part -match '^(\d+)-(\d+)$') { $start = [int]$matches[1]; $end = [int]$matches[2]; if ($start -ge 1 -and $end -le $files.Count -and $start -le $end) { $selected += $files[($start-1)..($end-1)] } } elseif ($part -match '^\d+$') { $num = [int]$part; if ($num -ge 1 -and $num -le $files.Count) { $selected += $files[$num-1] } } }; $selected = $selected | Select-Object -Unique }; if ($selected.Count -eq 0) { exit 1 }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $selected, $utf8NoBom)"

del "%TEMP%\edit_sel_input.txt" "%TEMP%\edit_sel_list.txt" "%TEMP%\edit_sel_output.txt" >nul 2>&1

REM Check if selection processing succeeded
if not exist "!SELECTED_FILES!" (
    echo Invalid selection. Try again.
    goto :do_selection_edit
)

REM Check if selection is empty
for /f %%A in ('powershell -NoProfile -Command "(Get-Content '!SELECTED_FILES!' -Raw -Encoding UTF8).Length"') do set "file_size=%%A"
if "!file_size!"=="0" (
    echo Invalid selection. Try again.
    goto :do_selection_edit
)

:selection_processed

:after_selection_edit
set "edited_selection=!selection!"
set "EDITED_SELECTED_FILES=!SELECTED_FILES!"

REM Display preview of selected files
if exist "!SELECTED_FILES!" (
    REM Count selected files using temp file
    set "selected_count=0"
    echo !SELECTED_FILES!> "%TEMP%\count_selected.txt"
    for /f %%A in ('powershell -NoProfile -Command "$listFile = Get-Content (Join-Path $env:TEMP 'count_selected.txt') -Raw; $listFile = $listFile.Trim(); $files = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($files) { Write-Output $files.Count } else { Write-Output 0 }"') do set "selected_count=%%A"
    del "%TEMP%\count_selected.txt" >nul 2>&1
    
    REM Default to 0 if empty
    if not defined selected_count set "selected_count=0"
    if "!selected_count!"=="" set "selected_count=0"
    
    REM Count excluded files (total - selected)
    set "excluded_count=0"
    if exist "!LAST_FILES!" (
        set "total_count=0"
        echo !LAST_FILES!> "%TEMP%\count_total.txt"
        for /f %%A in ('powershell -NoProfile -Command "$listFile = Get-Content (Join-Path $env:TEMP 'count_total.txt') -Raw; $listFile = $listFile.Trim(); $files = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($files) { Write-Output $files.Count } else { Write-Output 0 }"') do set "total_count=%%A"
        del "%TEMP%\count_total.txt" >nul 2>&1
        
        REM Default to 0 if empty
        if not defined total_count set "total_count=0"
        if "!total_count!"=="" set "total_count=0"
        
        REM Calculate excluded count
        if defined total_count if defined selected_count (
            set /a "excluded_count=!total_count! - !selected_count!" 2>nul
            if errorlevel 1 set "excluded_count=0"
        )
    )
    
    REM Display selected files with numbering using PowerShell
    set "DISPLAY_FILE=!TEMP_DIR!\FileLaunchSequencer_selected_display.txt"
    if exist "!DISPLAY_FILE!" del /f /q "!DISPLAY_FILE!"
    
    REM Write scan path to temp file for PowerShell
    echo !scan_path!> "%TEMP%\selected_display_base.txt"
    echo !SELECTED_FILES!> "%TEMP%\selected_display_list.txt"
    echo !DISPLAY_FILE!> "%TEMP%\selected_display_output.txt"
    
    REM Use PowerShell to format and display selected files
    powershell -NoProfile -Command "$base = Get-Content (Join-Path $env:TEMP 'selected_display_base.txt') -Raw; $listFile = Get-Content (Join-Path $env:TEMP 'selected_display_list.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'selected_display_output.txt') -Raw; $base = $base.Trim().TrimEnd('\'); $listFile = $listFile.Trim(); $outFile = $outFile.Trim(); $fileArray = Get-Content -LiteralPath $listFile -Encoding UTF8 | Where-Object { $_.Trim() }; if ($fileArray) { $counter = 0; $baseName = Split-Path -Leaf $base; $formattedLines = @(); foreach ($file in $fileArray) { $counter++; $file = $file.Trim(); $rel = $file.Replace($base, ''); if ($rel -match '^\\') { $rel = $rel.Substring(1) }; $formattedLines += ($counter.ToString() + '. \' + $baseName + '\' + $rel) }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $formattedLines, $utf8NoBom) }"
    
    REM Cleanup temp files
    del "%TEMP%\selected_display_base.txt" "%TEMP%\selected_display_list.txt" "%TEMP%\selected_display_output.txt" >nul 2>&1
    
    REM Display the formatted file list
    echo.
    echo EDITED FILE SCAN:
    echo.
    if exist "!DISPLAY_FILE!" (
        type "!DISPLAY_FILE!"
        del /f /q "!DISPLAY_FILE!" >nul 2>&1
    )
    
    echo.
    echo Found !selected_count! !file_ext! Files Selected
    echo Found !excluded_count! !file_ext! Files to Exclude
    echo.
) else (
    echo [ERROR] Selected files list not found.
    echo.
)

REM Ask if they want to save changes
echo.
echo Do you want to save changes?
echo   1 - Save changes to preset
echo   2 - Save as new preset
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_save_changes
set "save_changes="
set /p "save_changes=Enter Command: "

REM Handle menu return
if /i "!save_changes!"=="M" goto :after_save_changes

REM Normalize input
set "save_changes=!save_changes: =!"

if "!save_changes!"=="1" (
    REM Save the selection to the existing preset
    call :update_preset_selection_option1
    if errorlevel 1 (
        echo [ERROR] Failed to save changes to preset.
        pause
        exit /b 1
    )
    echo.
    echo Changes saved to preset.
    goto :after_save_changes
)
if "!save_changes!"=="2" (
    REM Save as new preset
    echo.
    echo Enter preset name
    :retry_new_preset_name
    set "input_new_preset_name="
    set /p "input_new_preset_name=Preset Name: "
    if /i "!input_new_preset_name!"=="M" goto :after_save_changes
    
    REM Basic validation
    if not defined input_new_preset_name (
        echo Invalid. Try again.
        goto :retry_new_preset_name
    )
    if "!input_new_preset_name!"=="" (
        echo Invalid. Try again.
        goto :retry_new_preset_name
    )
    
    REM Use the new preset name
    set "new_preset_name=!input_new_preset_name!"
    
    REM Create preset directory if it doesn't exist
    set "PRESET_DIR=%~dp0Presets\Option1"
    if not exist "!PRESET_DIR!" mkdir "!PRESET_DIR!" >nul 2>&1
    
    REM Sanitize preset name for filename
    set "preset_filename=!new_preset_name!"
    if defined preset_filename (
        set "preset_filename=!preset_filename:\=_!"
        set "preset_filename=!preset_filename:/=_!"
        set "preset_filename=!preset_filename::=_!"
        set "preset_filename=!preset_filename:?=_!"
        set "preset_filename=!preset_filename:"=_!"
        set "preset_filename=!preset_filename:<=_!"
        set "preset_filename=!preset_filename:>=_!"
        set "preset_filename=!preset_filename:|=_!"
    ) else (
        set "preset_filename=preset"
    )
    
    REM Save preset to JSON using current edited values
    set "NEW_PRESET_FILE=!PRESET_DIR!\!preset_filename!.json"
    
    REM Ensure all variables are defined with safe defaults
    if not defined file_ext set "file_ext=.txt"
    if not defined scan_path set "scan_path=C:\"
    if not defined folder_scan_option set "folder_scan_option=1"
    if not defined exclude_options_choice set "exclude_options_choice="
    if not defined exclude_paths set "exclude_paths="
    if not defined exclude_names set "exclude_names="
    if not defined exclude_keywords set "exclude_keywords="
    if not defined use_initial set "use_initial="
    if not defined initial_delay set "initial_delay=0"
    if not defined use_between set "use_between="
    if not defined between_delay set "between_delay=0"
    if not defined edited_selection set "edited_selection=all"
    
    REM Use edited_selection if available, otherwise use selection
    if defined edited_selection (
        set "save_selection=!edited_selection!"
    ) else if defined selection (
        set "save_selection=!selection!"
    ) else (
        set "save_selection=all"
    )
    
    REM Escape backslashes for JSON
    set "scan_path_safe=!scan_path!"
    if defined scan_path set "scan_path_safe=!scan_path:\=\\!"
    
    set "exclude_paths_safe=!exclude_paths!"
    if defined exclude_paths if not "!exclude_paths!"=="" set "exclude_paths_safe=!exclude_paths:\=\\\\!"
    
    REM Get current timestamp in 12-hour format
    set "timestamp="
    for /f "tokens=*" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd h:mmtt'"') do set "timestamp=%%i"
    
    REM Create the JSON file with current edited values
    (
        echo {
        echo     "preset_name": "!new_preset_name!",
        echo     "file_ext": "!file_ext!",
        echo     "scan_path": "!scan_path_safe!",
        echo     "folder_scan_option": "!folder_scan_option!",
        echo     "exclude_options_choice": "!exclude_options_choice!",
        echo     "exclude_paths": "!exclude_paths_safe!",
        echo     "exclude_names": "!exclude_names!",
        echo     "exclude_keywords": "!exclude_keywords!",
        echo     "use_initial": "!use_initial!",
        echo     "initial_delay": "!initial_delay!",
        echo     "use_between": "!use_between!",
        echo     "between_delay": "!between_delay!",
        echo     "file_selection": "!save_selection!",
        echo     "created": "!timestamp!"
        echo }
    ) > "!NEW_PRESET_FILE!"
    
    if exist "!NEW_PRESET_FILE!" (
        echo.
        echo Preset saved: !new_preset_name!
        goto :after_save_changes
    ) else (
        echo.
        echo [ERROR] Failed to save new preset.
        pause
        exit /b 1
    )
)
if /i "!save_changes!"=="S" (
    REM Skip saving - continue
    goto :after_save_changes
)
echo Invalid. Enter 1, 2, S, or M.
goto :retry_save_changes

:after_save_changes

REM Ask if they want to run it
echo.
echo Ready to open files with edited selection?
echo   Y - Yes
echo   M - Back to Menu
echo.
:retry_run_edited
set "run_edited="
set /p "run_edited=Enter Command: "

REM Handle menu return
if /i "!run_edited!"=="M" goto :preview_menu_return

REM Normalize input
set "run_edited=!run_edited: =!"

if /i "!run_edited!"=="Y" (
    REM Run with edited selection - preserve edited selection variables
    set "run_preset_num=!preset_index!"
    set "USE_EDITED_SELECTION=1"
    REM EDITED_SELECTED_FILES is already set above
    call :run_preset_handler
    set "USE_EDITED_SELECTION="
    set "EDITED_SELECTED_FILES="
    set "edited_selection="
    REM Clear edit_preset_num and run_preset_num so they don't trigger again on next visit
    set "edit_preset_num="
    set "run_preset_num="
    exit /b 0
)
echo Invalid. Enter Y or M.
goto :retry_run_edited

REM ========================================
REM Update Preset (Option 1)
REM ========================================
:update_preset_selection_option1
set "preset_file=!preset_file_path!"

if not defined preset_file (
    echo [ERROR] Preset file path not set.
    goto :update_preset_selection_option1_return
)

if not exist "!preset_file!" (
    echo [ERROR] Preset file does not exist: !preset_file!
    goto :update_preset_selection_option1_return
)

REM Ensure all variables are defined with safe defaults
if not defined preset_name set "preset_name=Unknown"
if not defined file_ext set "file_ext=.txt"
if not defined scan_path set "scan_path=C:\"
if not defined folder_scan_option set "folder_scan_option=1"
if not defined exclude_options_choice set "exclude_options_choice="
if not defined exclude_paths set "exclude_paths="
if not defined exclude_names set "exclude_names="
if not defined exclude_keywords set "exclude_keywords="
if not defined use_initial set "use_initial="
if not defined initial_delay set "initial_delay=0"
if not defined use_between set "use_between="
if not defined between_delay set "between_delay=0"
if not defined edited_selection set "edited_selection=all"

REM Escape backslashes for JSON
set "scan_path_escaped=!scan_path:\=\\!"
set "exclude_paths_escaped=!exclude_paths:\=\\!"

REM Get current timestamp in 12-hour format
set "timestamp="
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd h:mmtt'"') do set "timestamp=%%i"

REM Recreate the JSON file with updated values
(
    echo {
    echo     "preset_name": "!preset_name!",
    echo     "file_ext": "!file_ext!",
    echo     "scan_path": "!scan_path_escaped!",
    echo     "folder_scan_option": "!folder_scan_option!",
    echo     "exclude_options_choice": "!exclude_options_choice!",
    echo     "exclude_paths": "!exclude_paths_escaped!",
    echo     "exclude_names": "!exclude_names!",
    echo     "exclude_keywords": "!exclude_keywords!",
    echo     "use_initial": "!use_initial!",
    echo     "initial_delay": "!initial_delay!",
    echo     "use_between": "!use_between!",
    echo     "between_delay": "!between_delay!",
    echo     "file_selection": "!edited_selection!",
    echo     "created": "!timestamp!"
    echo }
) > "!preset_file!"

REM Success
exit /b 0

REM ========================================
REM Load Preset Selection (Option 1)
REM ========================================
:load_preset_selection_option1
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option1"

REM Get preset file by index (sorted by CreationTime to match display)
set "preset_file="
echo !PRESET_DIR!> "%TEMP%\load_selection_dir.txt"
echo !preset_index!> "%TEMP%\load_selection_index.txt"
for /f "delims=" %%F in ('powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'load_selection_dir.txt') -Raw; $index = Get-Content (Join-Path $env:TEMP 'load_selection_index.txt') -Raw; $dir = $dir.Trim(); $index = [int]$index.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; if ($presets -and $index -le $presets.Count) { Write-Output $presets[$index-1].FullName }"') do set "preset_file=%%F"
del "%TEMP%\load_selection_dir.txt" "%TEMP%\load_selection_index.txt" >nul 2>&1

if not defined preset_file (
    exit /b 1
)

REM Load selection from preset JSON using PowerShell
echo !preset_file!> "%TEMP%\load_selection_file.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$presetFile = Get-Content (Join-Path $env:TEMP 'load_selection_file.txt') -Raw; $presetFile = $presetFile.Trim(); $json = Get-Content -LiteralPath $presetFile -Raw -Encoding UTF8 | ConvertFrom-Json; if ($json.file_selection) { Write-Output $json.file_selection } else { exit 1 }"') do set "PRESET_SELECTION=%%A"
del "%TEMP%\load_selection_file.txt" "%TEMP%\load_selection_dir.txt" "%TEMP%\load_selection_index.txt" >nul 2>&1

if defined PRESET_SELECTION (
    exit /b 0
) else (
    exit /b 1
)

REM ========================================
REM Refresh Preset Menu (Option 1)
REM ========================================
:refresh_preset_menu_option1
cls
echo ========================================
echo    Presets for Folder-Based Protocol
echo ========================================
echo.
REM Re-display preset list
set "PRESET_DIR=%~dp0Presets\Option1"
set "preset_count=0"
for %%F in ("!PRESET_DIR!\*.json") do (
    set /a preset_count+=1
)
if !preset_count! EQU 0 (
    echo No presets saved yet.
    echo.
    pause
    goto :main
)
echo !PRESET_DIR!> "%TEMP%\preset_list_dir.txt"
set "PRESET_LIST_FILE=!TEMP_DIR!\preset_list_display.txt"
if exist "!PRESET_LIST_FILE!" del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
echo !PRESET_LIST_FILE!> "%TEMP%\preset_list_output.txt"
powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'preset_list_dir.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'preset_list_output.txt') -Raw; $dir = $dir.Trim(); $outFile = $outFile.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; $counter = 0; $lines = @(); foreach ($preset in $presets) { $counter++; try { $json = Get-Content -LiteralPath $preset.FullName -Raw -Encoding UTF8 | ConvertFrom-Json; $name = $json.preset_name; if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($preset.Name) }; $lines += ('p' + $counter.ToString() + ' - ' + $name) } catch { $lines += ('p' + $counter.ToString() + ' - ' + [System.IO.Path]::GetFileNameWithoutExtension($preset.Name)) } }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $lines, $utf8NoBom)"
del "%TEMP%\preset_list_dir.txt" "%TEMP%\preset_list_output.txt" >nul 2>&1
if exist "!PRESET_LIST_FILE!" (
    type "!PRESET_LIST_FILE!"
    del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
)
echo.
echo p#   - Run a Preset
echo p#p  - Preview a Preset
echo d#   - Delete a Preset
echo M    - Back to Menu
echo.
exit /b 0

REM ========================================
REM Delete Preset (Option 1)
REM ========================================
:delete_preset_option1
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option1"

REM Get preset file by index (sorted by CreationTime to match display)
set "preset_file="
echo !PRESET_DIR!> "%TEMP%\delete_preset_dir.txt"
echo !preset_index!> "%TEMP%\delete_preset_index.txt"
for /f "delims=" %%F in ('powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'delete_preset_dir.txt') -Raw; $index = Get-Content (Join-Path $env:TEMP 'delete_preset_index.txt') -Raw; $dir = $dir.Trim(); $index = [int]$index.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; if ($presets -and $index -le $presets.Count) { Write-Output $presets[$index-1].FullName }"') do set "preset_file=%%F"
del "%TEMP%\delete_preset_dir.txt" "%TEMP%\delete_preset_index.txt" >nul 2>&1

if not defined preset_file (
    echo [ERROR] Preset not found.
    pause
    exit /b 1
)

REM Get preset name for display
set "preset_name="
echo !preset_file!> "%TEMP%\delete_preset_file.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$presetFile = Get-Content (Join-Path $env:TEMP 'delete_preset_file.txt') -Raw; $presetFile = $presetFile.Trim(); try { $json = Get-Content -LiteralPath $presetFile -Raw -Encoding UTF8 | ConvertFrom-Json; $name = $json.preset_name; if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($presetFile) }; Write-Output $name } catch { Write-Output ([System.IO.Path]::GetFileNameWithoutExtension($presetFile)) }"') do set "preset_name=%%A"
del "%TEMP%\delete_preset_file.txt" >nul 2>&1

if not defined preset_name (
    for %%F in ("!preset_file!") do set "preset_name=%%~nF"
)

echo.
echo Are you sure you want to delete preset: !preset_name!
echo   Y - Yes
echo   N - No
echo   M - Back to Menu
echo.
:retry_delete_confirm
set "delete_confirm="
set /p "delete_confirm=Enter Command: "

REM Handle menu return
if /i "!delete_confirm!"=="M" exit /b 1

REM Normalize input
set "delete_confirm=!delete_confirm: =!"

if /i "!delete_confirm!"=="Y" (
    REM Delete the preset file
    if exist "!preset_file!" (
        del /f /q "!preset_file!" >nul 2>&1
        if errorlevel 1 (
            echo [ERROR] Failed to delete preset.
            pause
            exit /b 1
        )
        echo [LOG] Preset deleted successfully.
        exit /b 0
    ) else (
        echo [ERROR] Preset file not found.
        pause
        exit /b 1
    )
)
if /i "!delete_confirm!"=="N" (
    REM Cancelled - return to preset menu (errorlevel 2 to distinguish from M)
    exit /b 2
)
echo Invalid. Enter Y, N, or M.
goto :retry_delete_confirm

REM ========================================
REM Presets for Manual File Path Protocol
REM ========================================
:presets_option2
cls
echo ========================================
echo  Presets for Manual File Path Protocol
echo ========================================
echo.

REM Create preset directory if it doesn't exist
set "PRESET_DIR=%~dp0Presets\Option2"
if not exist "!PRESET_DIR!" mkdir "!PRESET_DIR!" >nul 2>&1

REM List all presets
set "preset_count=0"
for %%F in ("!PRESET_DIR!\*.json") do set /a "preset_count+=1"

if !preset_count! EQU 0 (
    echo No presets saved yet.
    echo.
    pause
    goto :eof
)

REM Display presets using PowerShell
echo !PRESET_DIR!> "%TEMP%\preset_list_dir.txt"
set "PRESET_LIST_FILE=!TEMP_DIR!\preset_list_display.txt"
if exist "!PRESET_LIST_FILE!" del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
echo !PRESET_LIST_FILE!> "%TEMP%\preset_list_output.txt"
powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'preset_list_dir.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'preset_list_output.txt') -Raw; $dir = $dir.Trim(); $outFile = $outFile.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; $counter = 0; $lines = @(); foreach ($preset in $presets) { $counter++; try { $json = Get-Content -LiteralPath $preset.FullName -Raw -Encoding UTF8 | ConvertFrom-Json; $name = $json.preset_name; if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($preset.Name) }; $lines += ('p' + $counter.ToString() + ' - ' + $name) } catch { $lines += ('p' + $counter.ToString() + ' - ' + [System.IO.Path]::GetFileNameWithoutExtension($preset.Name)) } }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $lines, $utf8NoBom)"
del "%TEMP%\preset_list_dir.txt" "%TEMP%\preset_list_output.txt" >nul 2>&1

if exist "!PRESET_LIST_FILE!" (
    type "!PRESET_LIST_FILE!"
    del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
)

echo.
echo p#   - Run a Preset
echo p#p  - Preview a Preset
echo d#   - Delete a Preset
echo c#   - Create a .txt with the path's of a preset scan
echo M    - Back to Menu
echo.
:retry_preset_select_option2
set "preset_select="
set /p "preset_select=Enter Command: "

REM Normalize input
set "preset_select=!preset_select: =!"
set "preset_select=!preset_select!"

REM Handle menu return FIRST
if /i "!preset_select!"=="M" goto :eof

REM Parse command using PowerShell - get command type and number
echo !preset_select!> "%TEMP%\preset_cmd_input.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$input = Get-Content (Join-Path $env:TEMP 'preset_cmd_input.txt') -Raw; $input = $input.Trim().ToLower(); $result = ''; if ($input -eq 'm') { $result = 'MENU' } elseif ($input -match '^d(\d+)$') { $result = 'DELETE|' + $matches[1] } elseif ($input -match '^c(\d+)$') { $result = 'CREATE_TXT|' + $matches[1] } elseif ($input -match '^p(\d+)p$') { $result = 'PREVIEW|' + $matches[1] } elseif ($input -match '^p(\d+)$') { $result = 'RUN|' + $matches[1] } else { $result = 'INVALID' }; Write-Output $result"') do set "cmd_parsed=%%A"
del "%TEMP%\preset_cmd_input.txt" >nul 2>&1

REM Split command type and number
for /f "tokens=1,2 delims=|" %%A in ("!cmd_parsed!") do (
    set "cmd_type=%%A"
    set "cmd_num=%%B"
)

REM Handle MENU command
if "!cmd_type!"=="MENU" goto :eof

REM Handle DELETE command
if "!cmd_type!"=="DELETE" (
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter d1-d!preset_count! to delete, or M.
        goto :retry_preset_select_option2
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter d1-d!preset_count! to delete, or M.
        goto :retry_preset_select_option2
    )
    call :delete_preset_option2 !cmd_num!
    if errorlevel 2 (
        REM N pressed - cancelled
        goto :refresh_and_retry_option2
    )
    if errorlevel 1 (
        REM M pressed - exit to main menu
        goto :eof
    )
    REM Delete succeeded
    echo Preset deleted successfully.
    timeout /t 1 /nobreak >nul
    goto :refresh_and_retry_option2
)

REM Handle CREATE_TXT command
if "!cmd_type!"=="CREATE_TXT" (
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter c1-c!preset_count! to create .txt file, or M.
        goto :retry_preset_select_option2
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter c1-c!preset_count! to create .txt file, or M.
        goto :retry_preset_select_option2
    )
    call :create_preset_paths_txt_option2 !cmd_num!
    if errorlevel 1 (
        echo [ERROR] Failed to create .txt file.
        pause
        goto :retry_preset_select_option2
    )
    echo.
    echo .txt file created successfully.
    timeout /t 2 /nobreak >nul
    goto :refresh_and_retry_option2
)

REM Handle PREVIEW command
if "!cmd_type!"=="PREVIEW" (
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter p1p-p!preset_count!p to preview, or M.
        goto :retry_preset_select_option2
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter p1p-p!preset_count!p to preview, or M.
        goto :retry_preset_select_option2
    )
    call :preview_preset_option2 !cmd_num!
    if errorlevel 2 (
        REM S pressed - skip
        timeout /t 1 /nobreak >nul
        goto :refresh_and_retry_option2
    )
    if errorlevel 1 (
        REM M pressed - exit to main menu
        goto :eof
    )
    REM Y pressed - run_preset_num was set, now run it
    if defined run_preset_num (
        goto :run_preset_handler_option2
    )
    REM Fallback
    goto :retry_preset_select_option2
)

REM Handle RUN command
if "!cmd_type!"=="RUN" (
    set /a cmd_num_val=!cmd_num! 2>nul
    REM Validate preset count is valid
    if "!preset_count!"=="" set "preset_count=0"
    if !cmd_num_val! LSS 1 (
        echo Invalid. Enter p1-p!preset_count! to run, or M.
        goto :retry_preset_select_option2
    )
    if !cmd_num_val! GTR !preset_count! (
        echo Invalid. Enter p1-p!preset_count! to run, or M.
        goto :retry_preset_select_option2
    )
    REM Load and run preset - use goto to exit if block first
    set "run_preset_num=!cmd_num!"
    goto :run_preset_handler_option2
)

REM Handle INVALID command
if "!cmd_type!"=="INVALID" (
    echo Invalid. Enter p1-p!preset_count! to run, p1p-p!preset_count!p to preview, d1-d!preset_count! to delete, c1-c!preset_count! to create .txt, or M.
    goto :retry_preset_select_option2
)

REM Fallback - shouldn't reach here
echo Invalid. Enter p1-p!preset_count! to run, p1p-p!preset_count!p to preview, d1-d!preset_count! to delete, c1-c!preset_count! to create .txt, or M.
goto :retry_preset_select_option2

REM ========================================
REM Refresh Menu and Retry (Option 4)
REM ========================================
:refresh_and_retry_option2
cls
echo ========================================
echo  Presets for Manual File Path Protocol
echo ========================================
echo.
REM Re-display preset list
set "PRESET_DIR=%~dp0Presets\Option2"
REM Count presets using batch command for consistency
set "preset_count=0"
for %%F in ("!PRESET_DIR!\*.json") do set /a "preset_count+=1"
if !preset_count! EQU 0 (
    echo No presets saved yet.
    echo.
    pause
    goto :main
)
echo !PRESET_DIR!> "%TEMP%\preset_list_dir.txt"
set "PRESET_LIST_FILE=!TEMP_DIR!\preset_list_display.txt"
if exist "!PRESET_LIST_FILE!" del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
echo !PRESET_LIST_FILE!> "%TEMP%\preset_list_output.txt"
powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'preset_list_dir.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'preset_list_output.txt') -Raw; $dir = $dir.Trim(); $outFile = $outFile.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; $counter = 0; $lines = @(); foreach ($preset in $presets) { $counter++; try { $json = Get-Content -LiteralPath $preset.FullName -Raw -Encoding UTF8 | ConvertFrom-Json; $name = $json.preset_name; if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($preset.Name) }; $lines += ('p' + $counter.ToString() + ' - ' + $name) } catch { $lines += ('p' + $counter.ToString() + ' - ' + [System.IO.Path]::GetFileNameWithoutExtension($preset.Name)) } }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $lines, $utf8NoBom)"
del "%TEMP%\preset_list_dir.txt" "%TEMP%\preset_list_output.txt" >nul 2>&1
if exist "!PRESET_LIST_FILE!" (
    type "!PRESET_LIST_FILE!"
    del /f /q "!PRESET_LIST_FILE!" >nul 2>&1
)
echo.
echo p#   - Run a Preset
echo p#p  - Preview a Preset
echo d#   - Delete a Preset
echo c#   - Create a .txt with the path's of a preset scan
echo M    - Back to Menu
echo.
goto :retry_preset_select_option2

REM ========================================
REM Run Preset Handler (Option 4)
REM ========================================
:run_preset_handler_option2
call :load_preset_option2 !run_preset_num!
if errorlevel 1 (
    pause
    goto :eof
)
echo.
echo [SCAN] Loading preset and running protocol...
echo.
REM Parse file paths and write to MANUAL_FILES temp file
set "MANUAL_FILES=!TEMP_DIR!\FileLaunchSequencer_manual.txt"
if exist "!MANUAL_FILES!" del /f /q "!MANUAL_FILES!"
echo !file_paths!> "%TEMP%\parse_paths.txt"
echo !MANUAL_FILES!> "%TEMP%\parse_output.txt"
powershell -NoProfile -Command "$pathsStr = Get-Content (Join-Path $env:TEMP 'parse_paths.txt') -Raw; $outFile = Get-Content (Join-Path $env:TEMP 'parse_output.txt') -Raw; $pathsStr = $pathsStr.Trim(); $outFile = $outFile.Trim(); $paths = $pathsStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }; $fullPaths = @(); foreach ($path in $paths) { try { $normalized = [System.IO.Path]::GetFullPath($path); $fullPaths += $normalized } catch { } }; $utf8NoBom = New-Object System.Text.UTF8Encoding $false; [System.IO.File]::WriteAllLines($outFile, $fullPaths, $utf8NoBom)"
del "%TEMP%\parse_paths.txt" "%TEMP%\parse_output.txt" >nul 2>&1
set "file_count=0"
for /f %%A in ('powershell -NoProfile -Command "$filePath = '!MANUAL_FILES!'; $files = Get-Content -LiteralPath $filePath -Encoding UTF8 | Where-Object { $_.Trim() }; Write-Output $files.Count"') do set "file_count=%%A"
if "!file_count!"=="0" (
    echo [ERROR] No valid file paths found in preset.
    pause
    goto :eof
)
REM Skip confirmation for presets - open files directly
call :open_manual_files
echo.
echo Preset execution completed.
pause
goto :main

REM ========================================
REM Load Preset Subroutine (Option 2)
REM ========================================
:load_preset_option2
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option2"

REM Get preset file by index (sorted by CreationTime to match display)
set "preset_file="
echo !PRESET_DIR!> "%TEMP%\load_preset_dir.txt"
echo !preset_index!> "%TEMP%\load_preset_index.txt"
for /f "delims=" %%F in ('powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'load_preset_dir.txt') -Raw; $index = Get-Content (Join-Path $env:TEMP 'load_preset_index.txt') -Raw; $dir = $dir.Trim(); $index = [int]$index.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; if ($presets -and $index -le $presets.Count) { Write-Output $presets[$index-1].FullName }"') do set "preset_file=%%F"
del "%TEMP%\load_preset_dir.txt" "%TEMP%\load_preset_index.txt" >nul 2>&1

if not defined preset_file (
    echo [ERROR] Preset not found.
    exit /b 1
)

REM Load preset using PowerShell
echo !preset_file!> "%TEMP%\load_preset_file.txt"
powershell -NoProfile -Command "$presetFile = Get-Content (Join-Path $env:TEMP 'load_preset_file.txt') -Raw; $presetFile = $presetFile.Trim(); $json = Get-Content -LiteralPath $presetFile -Raw -Encoding UTF8 | ConvertFrom-Json; $filePaths = if ($json.file_paths) { $json.file_paths } else { '' }; $useInitial = if ($json.use_initial) { $json.use_initial } else { '' }; $initialDelay = if ($json.initial_delay) { $json.initial_delay } else { '' }; $useBetween = if ($json.use_between) { $json.use_between } else { '' }; $betweenDelay = if ($json.between_delay) { $json.between_delay } else { '' }; Write-Output ($filePaths + '|' + $useInitial + '|' + $initialDelay + '|' + $useBetween + '|' + $betweenDelay)" > "%TEMP%\load_preset_data.txt"

REM Parse loaded data
set /p "load_line="<"%TEMP%\load_preset_data.txt"

REM Split the line by pipe delimiter
for /f "tokens=1,2,3,4,5 delims=|" %%A in ("!load_line!") do (
    set "file_paths=%%A"
    set "use_initial=%%B"
    set "initial_delay=%%C"
    set "use_between=%%D"
    set "between_delay=%%E"
)

del "%TEMP%\load_preset_file.txt" "%TEMP%\load_preset_data.txt" >nul 2>&1

REM Ensure all variables are properly set
if not defined file_paths set "file_paths="
if not defined use_initial set "use_initial="
if not defined initial_delay set "initial_delay=0"
if not defined use_between set "use_between="
if not defined between_delay set "between_delay=0"

exit /b 0

REM ========================================
REM Preview Preset Subroutine (Option 2)
REM ========================================
:preview_preset_option2
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option2"

REM Get preset file by index (sorted by CreationTime to match display)
set "preset_file="
echo !PRESET_DIR!> "%TEMP%\preview_preset_dir.txt"
echo !preset_index!> "%TEMP%\preview_preset_index.txt"
for /f "delims=" %%F in ('powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'preview_preset_dir.txt') -Raw; $index = Get-Content (Join-Path $env:TEMP 'preview_preset_index.txt') -Raw; $dir = $dir.Trim(); $index = [int]$index.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; if ($presets -and $index -le $presets.Count) { Write-Output $presets[$index-1].FullName }"') do set "preset_file=%%F"
del "%TEMP%\preview_preset_dir.txt" "%TEMP%\preview_preset_index.txt" >nul 2>&1

if not defined preset_file (
    echo [ERROR] Preset not found.
    pause
    exit /b 1
)

REM Display preset summary using inline PowerShell (avoids script file creation issues)
echo !preset_file!> "%TEMP%\preview_preset_file.txt"
powershell -NoProfile -Command "$presetFile = Get-Content (Join-Path $env:TEMP 'preview_preset_file.txt') -Raw; $presetFile = $presetFile.Trim(); $json = Get-Content -LiteralPath $presetFile -Raw -Encoding UTF8 | ConvertFrom-Json; Write-Output ''; Write-Output '================'; Write-Output ''; $created = if ($json.created) { $json.created } else { 'Unknown' }; if ($created -match 'T') { try { $dateObj = [DateTime]::Parse($created.Replace('T', ' ')); $created = $dateObj.ToString('yyyy-MM-dd h:mmtt') } catch { } }; Write-Output ('Created: ' + $created); Write-Output ''; $presetName = if ($json.preset_name) { $json.preset_name } else { 'N/A' }; Write-Output ('PRESET NAME: ' + $presetName); Write-Output ''; Write-Output 'Files:'; $filePaths = if ($json.file_paths) { $json.file_paths } else { '' }; if ($filePaths -and $filePaths.Trim() -ne '') { $paths = $filePaths -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }; foreach ($path in $paths) { try { $fullPath = [System.IO.Path]::GetFullPath($path); $dirName = [System.IO.Path]::GetDirectoryName($fullPath); $fileName = [System.IO.Path]::GetFileName($fullPath); $parentDir = Split-Path -Parent $dirName; $targetFolder = Split-Path -Leaf $parentDir; $subFolder = Split-Path -Leaf $dirName; if ($subFolder -ne $targetFolder) { Write-Output ('\' + $targetFolder + '\' + $subFolder + '\' + $fileName) } else { Write-Output ('\' + $targetFolder + '\' + $fileName) } } catch { Write-Output $path } } }; Write-Output ''; Write-Output 'Delay Settings:'; $useInitial = if ($json.use_initial) { $json.use_initial } else { 'No' }; if ($useInitial -eq 'Y') { $initialDelay = if ($json.initial_delay) { $json.initial_delay } else { '0' }; Write-Output ('  Initial Delay: ' + $initialDelay + ' seconds') } else { Write-Output '  Initial Delay: Skipped' }; $useBetween = if ($json.use_between) { $json.use_between } else { 'No' }; if ($useBetween -eq 'Y') { $betweenDelay = if ($json.between_delay) { $json.between_delay } else { '0' }; Write-Output ('  Between Delay: ' + $betweenDelay + ' seconds') } else { Write-Output '  Between Delay: Skipped' }; Write-Output ''; Write-Output '================'; Write-Output ''"

del "%TEMP%\preview_preset_file.txt" >nul 2>&1

echo.
echo Do you want to run this preset?
echo   Y - Yes
echo   S - Skip
echo   M - Back to Menu
echo.
:retry_preview_action_option2
set "preview_action="
set /p "preview_action=Enter Command: "

REM Handle menu return
if /i "!preview_action!"=="M" exit /b 1

REM Normalize input
set "preview_action=!preview_action: =!"

if /i "!preview_action!"=="Y" (
    REM Run the preset - set variable and jump to handler
    set "run_preset_num=!preset_index!"
    exit /b 0
)
if /i "!preview_action!"=="S" (
    REM Skip - return to preset menu (return code 2 to distinguish from other returns)
    exit /b 2
)
echo Invalid. Enter Y, S, or M.
goto :retry_preview_action_option2

REM ========================================
REM Delete Preset Subroutine (Option 2)
REM ========================================
:delete_preset_option2
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option2"

REM Get preset file by index (sorted by CreationTime to match display)
set "preset_file="
echo !PRESET_DIR!> "%TEMP%\delete_preset_dir.txt"
echo !preset_index!> "%TEMP%\delete_preset_index.txt"
for /f "delims=" %%F in ('powershell -NoProfile -Command "$dir = Get-Content (Join-Path $env:TEMP 'delete_preset_dir.txt') -Raw; $index = Get-Content (Join-Path $env:TEMP 'delete_preset_index.txt') -Raw; $dir = $dir.Trim(); $index = [int]$index.Trim(); $presets = Get-ChildItem -LiteralPath $dir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object CreationTime; if ($presets -and $index -le $presets.Count) { Write-Output $presets[$index-1].FullName }"') do set "preset_file=%%F"
del "%TEMP%\delete_preset_dir.txt" "%TEMP%\delete_preset_index.txt" >nul 2>&1

if not defined preset_file (
    echo [ERROR] Preset not found.
    pause
    exit /b 1
)

REM Get preset name for display
set "preset_name="
echo !preset_file!> "%TEMP%\delete_preset_file.txt"
for /f "delims=" %%A in ('powershell -NoProfile -Command "$presetFile = Get-Content (Join-Path $env:TEMP 'delete_preset_file.txt') -Raw; $presetFile = $presetFile.Trim(); try { $json = Get-Content -LiteralPath $presetFile -Raw -Encoding UTF8 | ConvertFrom-Json; $name = $json.preset_name; if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($presetFile) }; Write-Output $name } catch { Write-Output ([System.IO.Path]::GetFileNameWithoutExtension($presetFile)) }"') do set "preset_name=%%A"
del "%TEMP%\delete_preset_file.txt" >nul 2>&1

if not defined preset_name (
    for %%F in ("!preset_file!") do set "preset_name=%%~nF"
)

echo.
echo Are you sure you want to delete preset: !preset_name!
echo   Y - Yes
echo   N - No
echo   M - Back to Menu
echo.
:retry_delete_confirm_option2
set "delete_confirm="
set /p "delete_confirm=Enter Command: "

REM Handle menu return
if /i "!delete_confirm!"=="M" exit /b 1

REM Normalize input
set "delete_confirm=!delete_confirm: =!"

if /i "!delete_confirm!"=="Y" (
    REM Delete the preset file
    if exist "!preset_file!" (
        del /f /q "!preset_file!" >nul 2>&1
        if errorlevel 1 (
            echo [ERROR] Failed to delete preset.
            pause
            exit /b 1
        )
        echo [LOG] Preset deleted successfully.
        exit /b 0
    ) else (
        echo [ERROR] Preset file not found.
        pause
        exit /b 1
    )
)
if /i "!delete_confirm!"=="N" (
    REM Cancelled - return to preset menu (errorlevel 2 to distinguish from M)
    exit /b 2
)
echo Invalid. Enter Y, N, or M.
goto :retry_delete_confirm_option2

REM ========================================
REM Create .txt File with Preset Paths (Option 2)
REM ========================================
:create_preset_paths_txt_option2
set "preset_index=%~1"
set "PRESET_DIR=%~dp0Presets\Option2"

REM Find the preset file by counting files
set "counter=0"
for %%F in ("!PRESET_DIR!\*.json") do (
    set /a "counter+=1"
    if !counter! equ !preset_index! (
        set "preset_file=%%F"
        goto :found_preset_2
    )
)
:found_preset_2

if not defined preset_file (
    echo [ERROR] Preset not found.
    exit /b 1
)

REM Get preset name safely
set "preset_name="
for /f "tokens=2 delims=:" %%i in ('findstr /c:"preset_name" "!preset_file!"') do (
    set "line=%%i"
    set "line=!line:"=!"
    set "line=!line:,=!"
    for /f "tokens=* delims= " %%j in ("!line!") do set "preset_name=%%j"
)
if "!preset_name!"=="" set "preset_name=preset"

REM Sanitize filename (simple batch approach)
set "safe_name=!preset_name!"
REM Replace characters that are safe to replace
set "safe_name=!safe_name: =_!"
set "safe_name=!safe_name:\=_!"
set "safe_name=!safe_name:/=_!"
set "safe_name=!safe_name::=_!"
set "safe_name=!safe_name:?=_!"
set "safe_name=!safe_name:"=_!"
set "safe_name=!safe_name:<=_!"
set "safe_name=!safe_name:>=_!"
set "safe_name=!safe_name:|=_!"
REM Skip asterisk replacement to avoid batch wildcard issues

REM Get timestamp
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'h_mmtt-dd.MM.yyyy'"') do set "timestamp=%%i"

REM Create folder and file
if not exist "%~dp0Exported-Paths" mkdir "%~dp0Exported-Paths"
set "output_file=%~dp0Exported-Paths\!safe_name!-paths.-!timestamp!.txt"

REM Create the content
echo ================ > "!output_file!"
echo. >> "!output_file!"
echo Created: !date! !time! >> "!output_file!"
echo. >> "!output_file!"
echo PRESET NAME: !preset_name! >> "!output_file!"
echo. >> "!output_file!"
echo FILE PATHS: >> "!output_file!"
echo. >> "!output_file!"

REM Get file paths from JSON using pure batch
copy "!preset_file!" "%TEMP%\safe_preset.json" >nul 2>&1

REM Extract the file_paths line
for /f "tokens=*" %%l in ('findstr "file_paths" "%TEMP%\safe_preset.json"') do (
    set "line=%%l"
    goto :parse_line
)
:parse_line

REM Extract just the value between the quotes
REM The JSON has: "file_paths":  "value"  (note the two spaces after colon)
set "line=!line:*"file_paths":  =!"
REM Remove the trailing quote and comma
set "line=!line:",=!"
set "line=!line:"=!"

REM Now split by comma and process each path
setlocal enabledelayedexpansion
set "remaining=!line!"
:split_loop
for /f "tokens=1* delims=," %%a in ("!remaining!") do (
    set "current_path=%%a"
    set "remaining=%%b"

    REM Trim spaces and unescape backslashes
    for /f "tokens=* delims= " %%s in ("!current_path!") do (
        set "clean_path=%%s"
        REM Unescape backslashes (replace \\ with \)
        set "clean_path=!clean_path:\\=\!"
        if "!clean_path!" neq "" echo !clean_path! >> "!output_file!"
    )

    if defined remaining goto :split_loop
)
endlocal

del "%TEMP%\safe_preset.json" 2>nul

echo. >> "!output_file!"
echo ================ >> "!output_file!"

if exist "!output_file!" (
    echo.
    echo .txt file created: !output_file!
    exit /b 0
) else (
    echo [ERROR] Failed to create .txt file.
    exit /b 1
)

if exist "!OUTPUT_TXT_FILE!" (
    echo.
    echo .txt file created: !OUTPUT_TXT_FILE!
    exit /b 0
) else (
    echo [ERROR] Failed to create .txt file.
    exit /b 1
)

