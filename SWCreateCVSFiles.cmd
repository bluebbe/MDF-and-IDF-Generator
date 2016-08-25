@echo off

title Create CVS Files

:MainMenu
cls
Echo Which File(s) are you creating
Echo.
Echo 1) MDF, IDF, Verify files 
Echo 2) Time Zone Reboot files
Echo.
Set /p File=Enter your choice: 

if /i "%File%" EQU "1" Goto :SubMain1
if /i "%File%" EQU "2" goto :SubMain2
goto :MainMenu


:SubMain1

IF Exist MDF.csv del MDF.csv
If Exist IDF*.csv del IDF*.csv
If Exist Verify.csv del Verify.csv
If Exist MDF_IDF_VerifyLog.txt del MDF_IDF_VerifyLog.txt

notepad Storelist.txt


cscript CreateCVS.vbs Storelist.txt //Nologo
cls 
type MDF_IDF_VerifyLog.txt
pause
goto :eof

:SubMain2
IF Exist *Reboot.csv del *Reboot.csv
IF Exist *RebootLog.txt del *RebootLog.txt
echo txxxx>  Reboot.txt
notepad Reboot.txt


:Choice
cls
Echo Which Reboot file are you creating
Echo.
Echo E)astern Zone
Echo C)entral Zone
Echo M)ountain Zone
Echo P)acfic Zone
Echo.
Set /p TimeZone=Enter your choice: 
if /i "%TimeZone%" EQU "E" Goto :ValidChoice
if /i "%TimeZone%" EQU "C" goto :ValidChoice
if /i "%TimeZone%" EQU "M" goto :ValidChoice
if /i "%TimeZone%" EQU "P" goto :ValidChoice
goto :Choice


:ValidChoice

cscript CreateCVS.vbs Reboot.txt %TimeZone% //Nologo
cls
type *RebootLog.txt
pause
:eof