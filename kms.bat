@title 中国科学技术大学正版软件非官方激活脚本
@echo.========================================================================
@echo by cutecutecat
@echo.========================================================================
@echo.
@echo off
@echo off

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (

echo 请求管理员权限...

goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt

echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"

echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

"%temp%\getadmin.vbs"

exit /B

:gotAdmin

ping /n 1 /w 500 kms.ustc.edu.cn >nul && (goto ok) || (goto error)

:ok
@echo.
@echo.==================================================
echo 尝试连接KMS服务器，服务器连接正常
@echo.==================================================
@echo.
pause
@echo.
@echo 1. Windows 7/8/10 Professional专业版 激活 请按 [1]
@echo.
@echo 2. Office 2016 Pro Plus   专业增强版 激活 请按 [2]
@echo.
@echo 3. Office 2013 Pro Plus   专业增强版 激活 请按 [3]
@echo.
@echo 4. Office 2010 Pro Plus   专业增强版 激活 请按 [4]
@echo.
@echo off
set /p choice="请选择对应版本按 [1] 或 [2] 或 [3] 或 [4] 回车"
if %choice%==1 goto a1
if %choice%==2 goto a2
if %choice%==3 goto a3
if %choice%==4 goto a4

:a1
rem Windows 7/8/10 Professional
@echo.
@echo.==================================================================
@echo 注意：如果安装光盘不是由网络信息中心提供的，可能会导致激活失败。
@echo.==================================================================
@echo.
pause
@echo. 
c:
cd\
cd windows
cd system32
@cscript "%windir%\system32\slmgr.vbs" /skms kms.ustc.edu.cn
@cscript "%windir%\system32\slmgr.vbs" /ato
@echo.
pause
exit

:a2
rem Office 2016 Pro Plus
@echo.=======================================================================
@echo 开始检测Office 2016
@echo.=======================================================================
pause
@echo.
FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )
set file=OSPP.VBS
set vbs=%appdir%%file%
@echo %appdir%
@echo %vbs%
if exist "%vbs%" goto b
@echo.
@echo.=======================================================================
echo 未检测到Office 2016
@echo.=======================================================================
@echo.
pause
exit

:a3
rem Office 2013 Pro Plus
@echo.=======================================================================
@echo 开始检测Office 2013
@echo.=======================================================================
pause
@echo.
FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\13.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )
set file=OSPP.VBS
set vbs=%appdir%%file%
@echo %appdir%
@echo %vbs%
if exist "%vbs%" goto b
@echo.
@echo.=======================================================================
echo 未检测到Office 2013
@echo.=======================================================================
@echo.
pause
exit

:a4
rem Office Office 2010 Pro Plus
@echo.=======================================================================
@echo 开始检测Office 2010
@echo.=======================================================================
pause
@echo.
FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\10.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )
set file=OSPP.VBS
set vbs=%appdir%%file%
@echo %appdir%
@echo %vbs%
if exist "%vbs%" goto b
@echo.
@echo.=======================================================================
echo 未检测到Office 2010
@echo.=======================================================================
@echo.
pause
exit

:b
c:
cd\
cd "%appdir%"
cscript ospp.vbs /sethst:kms.ustc.edu.cn
cscript ospp.vbs /act
pause
exit

:error
@echo.
@echo.===================================
echo 无法与KMS服务器连接，请联系管理员
@echo.===================================
pause
@echo.
exit
