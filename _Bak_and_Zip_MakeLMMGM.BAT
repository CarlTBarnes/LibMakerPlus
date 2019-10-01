@echo off
for /f %%I in ('wmic os get localdatetime ^|find "20"') do set dt=%%I
REM dt format is now YYYYMMDDhhmmss...
REM dt format is now 01234567890123
Set DATE_TIME=%dt:~0,8%_%dt:~8,6%
echo %DATE_TIME% %dt%
REM ------------- Backup Folder Settings -----------------
SET SrcFld=.
SET BakFld=_Bak_\_Bak_Last
SET BakZip=_Bak_\LibMkrMgMns_%DATE_TIME%.ZIP

echo Backup    %SrcFld% 
echo Backup to %BakFld%  
echo Zip into  %BakZip%
REM ------------------------------------------------------
rem pause
@echo off
MD %BakFld%
@IF NOT EXIST %BakFld% THEN PAUSE

ROBOcopy %SrcFld% %BakFld% ^
  *.clw *.inc  *.app *.ap~ *.bpp *.dct  ^
  *.cwproj *.sln *.manifest *.red *.lib *.pr ^
  *.cur *.ico *.png *.gif *.jpg ^
  *.dctx *.txa *.txd  *.manifest        ^
  *.bat *.ini *.msbuild *.txt *.h       ^
  /MT /R:2 /W:5  /PURGE          ^
  /TEE /LOG+:%BakFld%\_RoboCopy.LOG

REM /MT=Multi Thread,  /Retry /W=Wait Error, /TEE=Counsel and Log File,  /LOG=Log File 
REM  /PURGE :: delete dest files/dirs that no longer exist in source.

echo+
echo Finished Copy to %BakFld%
echo --------------------------------
echo powershell Compress-Archive -Path %BakFld%\* -DestinationPath %BakZip%
rem pause
powershell Compress-Archive -Path %BakFld%\* -DestinationPath %BakZip%
@rem pause
DEL %0.bak
DEL Bak_*.bat.bak