@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

REM ##------------------------------------------------------------------------------##
REM ##  3S Pocketnet Technology Inc.                                                ##
REM ##                                                                              ##
REM ##  http://www.3spocketnet.com/                                                 ##
REM ##  7F., No.5, Lane 16, Sec. 2, Sichuan Rd. Banqiao Dist.,                      ##
REM ##  New Taipei City 22061, Taiwan (R.O.C.)                                      ##
REM ##  Tel: +886.2.8967.2909                                                       ##
REM ##  Fax: +886.2.8967.2779                                                       ##
REM ##  Author: peter.yang@3spocketnet.com.tw                                       ##
REM ##------------------------------------------------------------------------------##
REM 
REM The test tools is developed for IPCammera product of 3S Pocketnet Technology Inc.
REM
REM How to use:
REM 1. To make sure cscript/wscript is in your window system.
REM 2. Open main.bat, and change to your ipcamera ip
REM 3. Run main.bat

set CurrPath=%~dp0
set HTTP_IP=192.168.20.49
set PortNo=80
set ProjPath="X:\hisi3511\release\"
set FileName="N8072_V1.09_STD-1_20140604-101509.pkg"

start cmd /c "cscript FW_Upgrade.vbs !HTTP_IP! !PortNo! !ProjPath! !FileName!"
::start cmd /c "cscript FW_Upgrade.vbs !HTTP_IP! !PortNo! !ProjPath!"

