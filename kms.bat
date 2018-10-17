@title 中国科学技术大学正版软件激活脚本
@echo.========================================================================
@echo 中国科学技术大学正版软件激活脚本
@echo modified by starkind
@echo.========================================================================
@echo.
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
set file=OSPP.VBS
FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )

if "%appdir:~-1%"==" " (set appdir=%appdir:~0,-1%)
set vbs=%appdir%%file%
if exist "%vbs%" goto b

FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )
echo %appdir:~,-1%
if "%appdir:~-1%"==" " (set appdir=%appdir:~0,-1%)
set vbs=%appdir%%file%
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
set file=OSPP.VBS
FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )

if "%appdir:~-1%"==" " (set appdir=%appdir:~0,-1%)
set vbs=%appdir%%file%
if exist "%vbs%" goto b

FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )
echo %appdir:~,-1%
if "%appdir:~-1%"==" " (set appdir=%appdir:~0,-1%)
set vbs=%appdir%%file%
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
set file=OSPP.VBS
FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )

if "%appdir:~-1%"==" " (set appdir=%appdir:~0,-1%)
set vbs=%appdir%%file%
if exist "%vbs%" goto b

FOR /F "usebackq tokens=3*" %%A IN (`REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Word\InstallRoot" /v Path`) DO (
    set appdir=%%A %%B
    )
echo %appdir:~,-1%
if "%appdir:~-1%"==" " (set appdir=%appdir:~0,-1%)
set vbs=%appdir%%file%
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
echo 无法与KMS服务器连接，请检查网络
@echo.===================================
pause
@echo.
exit
