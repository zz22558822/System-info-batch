@echo off
@Title Windows 系統資訊
setlocal enabledelayedexpansion
Color 0B & Mode 57,3
Set Systeminfo_TXT=%TEMP%\系統資訊_%ComputerName%.txt
Set Systeminfo_HTML=%~dp0系統資訊_%ComputerName%.html
echo(
systeminfo>"%Systeminfo_TXT%"


::REM -----CPU資訊-----
set "CPU_AddressWidth="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET AddressWidth') do (
    if not "%%A"=="" (
        set "CPU_AddressWidth=%%A"
        goto :CPU_AddressWidth_done
    )
)
:CPU_AddressWidth_done
call :TrimSpaces "!CPU_AddressWidth!" CPU_AddressWidth

set "CPU_MaxClockSpeed="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET MaxClockSpeed') do (
    if not "%%A"=="" (
        set "CPU_MaxClockSpeed=%%A"
        goto :CPU_MaxClockSpeed_done
    )
)
:CPU_MaxClockSpeed_done
call :TrimSpaces "!CPU_MaxClockSpeed!" CPU_MaxClockSpeed

set "CPU_L2CacheSize="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET L2CacheSize') do (
    if not "%%A"=="" (
        set "CPU_L2CacheSize=%%A"
        goto :CPU_L2CacheSize_done
    )
)
:CPU_L2CacheSize_done
call :TrimSpaces "!CPU_L2CacheSize!" CPU_L2CacheSize

set "CPU_L3CacheSize="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET L3CacheSize') do (
    if not "%%A"=="" (
        set "CPU_L3CacheSize=%%A"
        goto :CPU_L3CacheSize_done
    )
)
:CPU_L3CacheSize_done
call :TrimSpaces "!CPU_L3CacheSize!" CPU_L3CacheSize

set "CPU_Manufacturer="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET Manufacturer') do (
    if not "%%A"=="" (
        set "CPU_Manufacturer=%%A"
        goto :CPU_Manufacturer_done
    )
)
:CPU_Manufacturer_done
call :TrimSpaces "!CPU_Manufacturer!" CPU_Manufacturer

set "CPU_Name="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET Name') do (
    if not "%%A"=="" (
        set "CPU_Name=%%A"
        goto :CPU_Name_done
    )
)
:CPU_Name_done
call :TrimSpaces "!CPU_Name!" CPU_Name

set "CPU_NumberOfCores="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET NumberOfCores') do (
    if not "%%A"=="" (
        set "CPU_NumberOfCores=%%A"
        goto :CPU_NumberOfCores_done
    )
)
:CPU_NumberOfCores_done
call :TrimSpaces "!CPU_NumberOfCores!" CPU_NumberOfCores

set "CPU_NumberOfLogicalProcessors="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET NumberOfLogicalProcessors') do (
    if not "%%A"=="" (
        set "CPU_NumberOfLogicalProcessors=%%A"
        goto :CPU_NumberOfLogicalProcessors_done
    )
)
:CPU_NumberOfLogicalProcessors_done
call :TrimSpaces "!CPU_NumberOfLogicalProcessors!" CPU_NumberOfLogicalProcessors

set "CPU_ProcessorID="
for /f "skip=1 tokens=*" %%A in ('wmic CPU GET ProcessorID') do (
    if not "%%A"=="" (
        set "CPU_ProcessorID=%%A"
        goto :CPU_ProcessorID_done
    )
)
:CPU_ProcessorID_done
call :TrimSpaces "!CPU_ProcessorID!" CPU_ProcessorID


::REM -----主機板資訊-----
set "MB_Manufacturer="
for /f "skip=1 tokens=*" %%A in ('wmic BASEBOARD GET Manufacturer') do (
    if not "%%A"=="" (
        set "MB_Manufacturer=%%A"
        goto :MB_Manufacturer_done
    )
)
:MB_Manufacturer_done
call :TrimSpaces "!MB_Manufacturer!" MB_Manufacturer

set "MB_Product="
for /f "skip=1 tokens=*" %%A in ('wmic BASEBOARD GET Product') do (
    if not "%%A"=="" (
        set "MB_Product=%%A"
        goto :MB_Product_done
    )
)
:MB_Product_done
call :TrimSpaces "!MB_Product!" MB_Product

set "MB_SerialNumber="
for /f "skip=1 tokens=*" %%A in ('wmic BASEBOARD GET SerialNumber') do (
    if not "%%A"=="" (
        set "MB_SerialNumber=%%A"
        goto :MB_SerialNumber_done
    )
)
:MB_SerialNumber_done
call :TrimSpaces "!MB_SerialNumber!" MB_SerialNumber

set "MB_Version="
for /f "skip=1 tokens=*" %%A in ('wmic BASEBOARD GET Version') do (
    if not "%%A"=="" (
        set "MB_Version=%%A"
        goto :MB_Version_done
    )
)
:MB_Version_done
call :TrimSpaces "!MB_Version!" MB_Version

set "MB_CreationClassName="
for /f "skip=1 tokens=*" %%A in ('wmic BASEBOARD GET CreationClassName') do (
    if not "%%A"=="" (
        set "MB_CreationClassName=%%A"
        goto :MB_CreationClassName_done
    )
)
:MB_CreationClassName_done
call :TrimSpaces "!MB_CreationClassName!" MB_CreationClassName


::REM -----記憶體資訊-----
set "RAM_Manufacturer="
for /f "skip=1 tokens=*" %%A in ('wmic memorychip GET Manufacturer') do (
    if not "%%A"=="" (
        set "RAM_Manufacturer=%%A"
        goto :RAM_Manufacturer_done
    )
)
:RAM_Manufacturer_done
call :TrimSpaces "!RAM_Manufacturer!" RAM_Manufacturer

set "RAM_Capacity="
for /f "skip=1 tokens=*" %%A in ('wmic memorychip GET Capacity') do (
    if not "%%A"=="" (
        set "RAM_Capacity=%%A"
        goto :RAM_Capacity_done
    )
)
:RAM_Capacity_done
call :TrimSpaces "!RAM_Capacity!" RAM_Capacity

for /f %%B in ('powershell "[Math]::Round(%RAM_Capacity% / 1GB, 2)"') do (
    set "RAM_Capacity_GB=%%B"
)

set "RAM_SerialNumber="
for /f "skip=1 tokens=*" %%A in ('wmic memorychip GET SerialNumber') do (
    if not "%%A"=="" (
        set "RAM_SerialNumber=%%A"
        goto :RAM_SerialNumber_done
    )
)
:RAM_SerialNumber_done
call :TrimSpaces "!RAM_SerialNumber!" RAM_SerialNumber

set "RAM_ConfiguredClockSpeed="
for /f "skip=1 tokens=*" %%A in ('wmic memorychip GET ConfiguredClockSpeed') do (
    if not "%%A"=="" (
        set "RAM_ConfiguredClockSpeed=%%A"
        goto :RAM_ConfiguredClockSpeed_done
    )
)
:RAM_ConfiguredClockSpeed_done
call :TrimSpaces "!RAM_ConfiguredClockSpeed!" RAM_ConfiguredClockSpeed

set "RAM_MemoryType="
for /f "skip=1 tokens=*" %%A in ('wmic memorychip GET MemoryType') do (
    if not "%%A"=="" (
        set "RAM_MemoryType=%%A"
        goto :RAM_MemoryType_done
    )
)
:RAM_MemoryType_done
call :TrimSpaces "!RAM_MemoryType!" RAM_MemoryType
if "!RAM_MemoryType!"=="0" (
    set "MemoryTypeDesc=桌上型"
) else if "!RAM_MemoryType!"=="8" (
    set "MemoryTypeDesc=桌上型"
) else if "!RAM_MemoryType!"=="12" (
    set "MemoryTypeDesc=筆記型"
) else (
    set "MemoryTypeDesc=未知類型"
)

set "RAM_MemoryDevices="
for /f "skip=1 tokens=*" %%A in ('wmic Memphysical get MemoryDevices') do (
    if not "%%A"=="" (
        set "RAM_MemoryDevices=%%A"
        goto :RAM_MemoryDevices_done
    )
)
:RAM_MemoryDevices_done
call :TrimSpaces "!RAM_MemoryDevices!" RAM_MemoryDevices

echo ----------:-------------------------------------------------- >>"%Systeminfo_TXT%"
echo CPU 名稱: %CPU_Name% >>"%Systeminfo_TXT%"
echo CPU 位元: %CPU_AddressWidth% >>"%Systeminfo_TXT%"
echo CPU 時脈: %CPU_MaxClockSpeed% >>"%Systeminfo_TXT%"
echo CPU L2: %CPU_L2CacheSize% >>"%Systeminfo_TXT%"
echo CPU L3: %CPU_L3CacheSize% >>"%Systeminfo_TXT%"
echo CPU 製造商: %CPU_Manufacturer% >>"%Systeminfo_TXT%"
echo CPU 唯一碼: %CPU_ProcessorID% >>"%Systeminfo_TXT%"
echo CPU 核心數: %CPU_NumberOfCores% >>"%Systeminfo_TXT%"
echo CPU 執行緒: %CPU_NumberOfLogicalProcessors% >>"%Systeminfo_TXT%"
echo 主機板 名稱: %MB_Product% >>"%Systeminfo_TXT%"
echo 主機板 製造商: %MB_Manufacturer% >>"%Systeminfo_TXT%"
echo 主機板 版本: %MB_Version% >>"%Systeminfo_TXT%"
echo 主機板 序列號: %MB_SerialNumber% >>"%Systeminfo_TXT%"
echo 主機板 類別: %MB_CreationClassName% >>"%Systeminfo_TXT%"
echo RAM 製造商: %RAM_Manufacturer% ※四位代碼為製造商代號 >>"%Systeminfo_TXT%"
echo RAM 容量: %RAM_Capacity_GB% GB >>"%Systeminfo_TXT%"
echo RAM 時脈: %RAM_ConfiguredClockSpeed% MHz >>"%Systeminfo_TXT%"
echo RAM 類別: %MemoryTypeDesc% >>"%Systeminfo_TXT%"
echo RAM 序列號: %RAM_SerialNumber% >>"%Systeminfo_TXT%"
echo RAM 最大槽數: %RAM_MemoryDevices% >>"%Systeminfo_TXT%"


call :forHTML "%Systeminfo_TXT%" "%Systeminfo_HTML%"
Start "" "%Systeminfo_HTML%"
Exit /b
::------------------------------------------------------------------------------------------------------------------------------------------------------
:forHTML <inputfile> <outputfile>
>"%~2" ( 
    echo ^<!DOCTYPE html^>
    echo ^<html lang="zh-Hant-TW"^>
    echo ^<head^>
    echo ^<meta HTTP-EQUIV="Content-Type" 
    echo CONTENT="text/html; charset=Big5"^>
	echo ^<meta name="viewport" content="initial-scale=1.0; maximum-scale=1.0; width=device-width;"^>
	echo ^<title^>%ComputerName% 系統資訊^</title^>
    echo ^</head^>
    echo ^<body^>
    echo ^<style type="text/css"^>
    echo body {background-color: #3a444e;font-family: "微軟正黑體", helvetica, arial, sans-serif;font-size: 16px;text-rendering: optimizeLegibility;}
    echo .wrap .table-title {display: block;margin: auto;max-width: 1024px;padding:5px;width: 100%%;}
    echo .table thead {position: sticky;top: 0px;}
	echo thead th {text-align: center;}
    echo .table-title h3 {color: #fafafa;font-size: 30px;margin: 10px 0;font-style:normal;font-family: "微軟正黑體", helvetica, arial, sans-serif;text-transform:uppercase;text-align: center;}
    echo .table {background: white;border-radius:3px;border-collapse: collapse;height: 320px;margin: auto;max-width: 1024px;padding:5px;width: 100%%;animation: float 5s infinite;}
    echo th {color:#D5DDE5;;background:#1b1e24;border-bottom:4px solid #9ea7af;font-size:23px;padding:15px;text-align:left;vertical-align:middle;}
    echo th:first-child {border-top-left-radius:3px;}
    echo th:last-child {border-top-right-radius:3px;border-right:none;}
    echo tr {border-top: 1px solid #C1C3D1;border-bottom-: 1px solid #C1C3D1;color:#666B85;font-size:16px;}
    echo tr:hover td {background:#4E5066;color:#FFFFFF;border-top: 1px solid #22262e;}
    echo tr:first-child {border-top:none;}
    echo tr:last-child {border-bottom:none;}
    echo tr:nth-child^(odd^) td {background:#EBEBEB;}
    echo tr:nth-child^(odd^):hover td {background:#4E5066;}
    echo tr:last-child td:first-child {border-bottom-left-radius:3px;}
    echo tr:last-child td:last-child {border-bottom-right-radius:3px;}
    echo td {background:#FFFFFF;padding:10px 20px;text-align:left;vertical-align:middle;font-size:18px;border-right: 1px solid #C1C3D1;}
    echo td:last-child {border-right: 0px;}
    echo th {text-align: left;}
    echo td {text-align: left;font-weight: bold;}
	echo .popup { position: fixed;top: 50px;right: 20px;opacity: 0;visibility: hidden;background-color: #666;color: #fff;text-align: center;border-radius: 6px;padding: 8px;transition: all 0.3s;z-index: 999;user-select: none; }
	echo .popup.show { opacity: 1;visibility: visible; }
	echo .popup .popup-content { display: block;margin: 10px;font-size: 36px; }
	echo .popup .progress-bar { height: 5px;background-color: #fff;width: 100%;margin-bottom: 0;border-radius: 2px; }
	echo .popup: hover { transform: scale^(1.05^);cursor: pointer; }
    echo ^</style^>
    echo ^<div class="wrap"^>^<div class="table-title"^>^<h3^>%ComputerName% 系統資訊^</h3^>^</div^>^<table class="table"^>
	echo ^<thead^>^<tr^>^<th class="text-left"^>種類^</th^>^<th class="text-left"^>資料^</th^>^</tr^>^</thead^>
)
@for /f "tokens=1,* delims==:" %%a in ('Type "%~1"') do (
    >>"%~2" echo ^<tr^>^<td^>%%a^</td^>^<td^>%%b^</td^>^</tr^>
)
>>"%~2" (
    echo ^</div^>
    echo ^</table^>
	echo ^<div id="myPopup" class="popup"^>^<span class="popup-content"^>複製到剪貼簿^</span^>^<div class="progress-bar"^>^</div^>^</div^>
	echo ^<script^>
	echo document.addEventListener^('DOMContentLoaded', function^(^) {
	echo     const tableDataCells = document.querySelectorAll^('.table td'^);
	echo     tableDataCells.forEach^(td =^> {
	echo         td.addEventListener^('click', function^(^) {
	echo             const tempInput = document.createElement^('textarea'^);
	echo             tempInput.value = td.textContent.trim^(^);
	echo             document.body.appendChild^(tempInput^);
	echo             tempInput.select^(^);
	echo             document.execCommand^('copy'^);
	echo             document.body.removeChild^(tempInput^);
	echo             console.log^(`複製成功: ^${td.textContent.trim^(^)}`^);
	echo             triggerPopupAndProgress^(^);
	echo         }^);
	echo     }^);
	echo }^);
    echo ^</script^>
	echo ^<script^>
	echo function triggerPopupAndProgress^(^) {
	echo     let timeoutId;
	echo     let progressInterval;
	echo     let progressBar = document.querySelector^(".progress-bar"^);
	echo     const popup = document.getElementById^("myPopup"^);
	echo     popup.addEventListener^("click", function^(^) {
	echo         popup.classList.remove^("show"^);
	echo         progressBar.style.width = "0%%";
	echo     }^);
	echo     clearTimeout^(timeoutId^);
	echo     clearInterval^(progressInterval^);
	echo     progressBar.style.width = "100%%";
	echo     progressInterval = setInterval^(function^(^) {
	echo         let width = parseInt^(progressBar.style.width^);
	echo         width -= 1;
	echo         progressBar.style.width = width + "%%";
	echo         if ^(width ^<= 0^) {
	echo             clearInterval^(progressInterval^);
	echo             popup.classList.remove^("show"^);
	echo         }
	echo     }, 10^);
	echo     popup.classList.add^("show"^);
	echo     timeoutId = setTimeout^(function^(^) {
	echo         popup.classList.remove^("show"^);
	echo     }, 1000^);
	echo }
	echo ^</script^>
	echo ^<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.7.0/jquery.min.js"^>^</script^>
    echo ^</body^>
    echo ^</html^>
)

del /F /S /Q %~dp0系統資訊_%ComputerName%.txt >nul 2>&1
del /F /S /Q %TEMP%\系統資訊_%ComputerName%.txt >nul 2>&1
Exit /B



::REM ----- 去除空白副程式 -----
:TrimSpaces
setlocal
set "str=%~1"
if not defined str (
    endlocal
    set "%~2="
    exit /b
)
:TrimSpacesLoop
if "%str:~-1%"==" " (
    set "str=%str:~0,-1%"
    goto TrimSpacesLoop
)
endlocal & set "%~2=%str%"
exit /b