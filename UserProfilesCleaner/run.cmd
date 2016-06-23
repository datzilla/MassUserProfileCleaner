@ECHO OFF
ECHO.>> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
ECHO ---------------------------------------------------------------------------------------------------------->> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
ECHO                      RUN AT %TIME% %DATE% >> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
ECHO -------------------------------------------------------------------->> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
@RD /S /Q %SYSTEMDRIVE%\Users\Default\AppData\Roaming\Microsoft\Templates\LiveContent>> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
@CSCRIPT //NoLogo "%~DP0UserProfilesCleaner.vbs" exclude bloombergtraining removeall cleartemp defrag reboot>> c:\%COMPUTERNAME%.runlog
ECHO ---------------------------------------------------------------------------------------------------------->> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
ECHO                   COMPLETED AT %TIME% %DATE% >> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
ECHO ---------------------------------------------------------------------------------------------------------->> %SYSTEMDRIVE%\%COMPUTERNAME%.runlog
@RD /S /Q "%~DP0"
EXIT /B %ERRORLEVEL%