@echo off
title Activator for Microsoft Office 2019 Pro Plus - The Mode&cls&echo ----------------------------------------------------------------------------&echo.&echo Copyright - The Mode.&echo.&echo ----------------------------------------------------------------------------&echo.&echo # Supported products: Microsoft Office 2019 Pro Plus [ x86 or x64 ].&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ----------------------------------------------------------------------------&echo.&echo Activating Microsoft Office 2019...&echo.&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&set i=1&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul||cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul||goto notsupported
:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=s8.uk.to
if %i% EQU 3 set KMS=s9.us.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo ----------------------------------------------------------------------------&echo.&cscript //nologo ospp.vbs /act | find /i "Product Activation Successful" && (echo.&echo ----------------------------------------------------------------------------&echo.&echo # Join To My Discord Server: Dsc.gg/TheMode&echo.&echo # Visit My YouTube Channel: YouTube.com/TheModeUniverse&echo.&echo # Special Thanks To: MsGuides.com&echo.&echo ----------------------------------------------------------------------------&echo.&choice /n /c YN /m "Would You Like To Be A Valuable Member Of My YouTube Channel [ Y, N ]?" & if errorlevel 2 exit) || (echo The connection to msguides.com first KMS server failed! & echo Trying to connect to another one... & echo Please wait... & echo. & set /a i+=1 & goto skms)
explorer "https://YouTube.com/TheModeUniverse/Join"&goto halt
:notsupported
echo ----------------------------------------------------------------------------&echo.&echo Sorry, your version is not supported.&echo.&goto halt
:busy
echo ----------------------------------------------------------------------------&echo.&echo Sorry, the server is busy and can't respond to your request. Please try again.&echo.
:halt
pause >nul