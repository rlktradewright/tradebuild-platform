@echo off
setlocal enableextensions enabledelayedexpansion

::=============================================================================+
::                                                                             +
::   This command file runs the TradeSkil Demo program, which is a manual      +
::   trading client that uses the TradeBuild Platform. It is mainly intended   +
::   as a sample program that demonstrates houw to use the various components  +
::   of the plarform to create a realistic (though limited) trading client.    +
::   testing and auto-execution of trading strategies.                         +
::                                                                             +
::   You may wish to change some of the settings below, but it should work     +
::   well without changes if TWS is running on this computer, and the          +
::   TWS/Gateway API port is configured to 7497 (see the PORT setting below).  +
::                                                                             +
::                                                                             +
::   The following lines, beginning with 'set', are the only ones you may      +
::   need to change.                                                           +
::                                                                             +
::   The notes below give further information on why you might need to         +
::   change the settings.                                                      +
::                                                                             +
::=============================================================================+




set TOPDIR=%~dp0

set CONFIGFILE=%APPDATA%\TradeWright\TradeSkil Demo Edition\settings.xml

set LOG=%TOPDIR%\Log\TradeSkil.log
set LOGLEVEL=D
set LOGOVERWRITE=yes
set LOGBACKUP=yes

set APIMESSAGELOGGING=NNN

set INSTALLFOLDER=


::              PLEASE DON'T CHANGE ANYTHING BELOW THIS LINE !!
::==============================================================================
::
::   Notes:
::
::
::
:: TOPDIR	
::   Set this to the root folder for use of the TradeSkil program.
::
::
:: CONFIGFILE
::   Set this to the path and filename where the program's configuration
::   settings are to be stored. The settings are stored in XML format.
::
::
:: LOG
::   Set this to the program's log filename. This file records details of
::   program operation, which can be helpful in identifying the reason for
::   program failures. If no value is specified, the filename will be
::   %APPDATA%\TradeWright\plord\plord.log
::
:: LOGLEVEL
::   Set this to control the level of detail in the program logfile.
::   Permitted values, in increasing level of detail, are N, D, M and H. For
::   normal operation N is recommended. Note that the higher the level of
::   detail, the larger the logfile.
::
:: LOGOVERWRITE
::   Indicates whether to overwrite the previous logfile. Permitted values are
::   YES and NO. If set to NO, the previous logfile is appended to. Note that
::   if LOGBACKUP is set to YES, this setting is ignored, because a new logfile
::   is always created.
::
:: LOGBACKUP
::   Indicates whether to retain the previous logfile by appending a suffix to
::   its filename. The suffix is of the form .bak[n], where n is incremented
::   for each run.
::
::
:: APIMESSAGELOGGING
::   This setting controls the logging of API messages and API message 
::   statistics. It should be left at the default value of 'NNN' except under
::   advice from the program developer. API message logging can cause 
::   logfiles to become very large: the default setting turns off all API
::   message logging except for those related to orders, API connections and
::   API errors/notfications.
::
::
:: INSTALLFOLDER
::   Set this to the folder where you installed the TradeBuild Platform.
::   If you installed TradeBuild Platform using the .msi installer with the 
::   default installation location, you do not need to set this value.
::
::
::
::   End of Notes
::==============================================================================

if "%INSTALLFOLDER%" NEQ "" (
	set "SCRIPTS=%INSTALLFOLDER%\Scripts"
) else if defined PROGRAMFILES^(X86^) (
	set "SCRIPTS=%PROGRAMFILES(X86)%\TradeWright Software Systems\TradeBuild Platform 2.7\Scripts"
) else (
	set "SCRIPTS=%PROGRAMFILES%\TradeWright Software Systems\TradeBuild Platform 2.7\Scripts"
)
"%SCRIPTS%\TradeSkil.bat"
