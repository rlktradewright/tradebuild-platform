@echo off
setlocal enableextensions enabledelayedexpansion

::=============================================================================+
::                                                                             +
::   This command file runs the gsd27.exe program for retrieving market scan   +
::   data from IBKR's Trader Workstation or Gateway.                           + 
::                                                                             +
::   If the first argument is /I, the program is run in interactive mode,      +
::   meaning that you must type commands manually. Any other arguments are     +
::   ignored.                                                                  +
::                                                                             +
::   If the first argument is a filename, the program reads the file from      +
::   the directory specified by the INPUTFILESDIR setting (see below) and      +
::   actions it. Any other arguments are ignored.                              +
::                                                                             +
::   If there are no arguments, the program monitors the INPUTFILESDIR         +
::   directory and actions any commandfiles that are subsequently placed in    +
::   it.                                                                       +
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

set TWSSERVER=
set PORT=7497
set CLIENTID=

set LOG=%TOPDIR%\Log\gsd27.log
set LOGLEVEL=N
set APIMESSAGELOGGING=NNN

set FILEFILTER=gsd*.txt
set INPUTFILESDIR=%TOPDIR%\InputFiles
set ARCHIVEDIR=%TOPDIR%\Archive
set OUTPUTDIR=%TOPDIR%\ScanData\Scan-{$scancode}\scan {$today}.txt

set INSTALLFOLDER=

set PIPELINE=



::              PLEASE DON'T CHANGE ANYTHING BELOW THIS LINE !!
::==============================================================================
::
::   Notes:
::
::
::
:: TOPDIR	
::   Set this to the root folder for use of the GetScanData program.
::
::
:: TWSSERVER
::   Set this to the name of the computer that is running TWS or Gateway. If 
::   that's this computer, just leave it blank.
::
:: PORT
::   Set this to the API port number for TWS/Gateway. The default value is 7496.
::
::   You can configure the value to use in the API settings section of 
::   TWS/Gateway's Global Configuration dialog. Note that if you run live and
::   paper-trading instances of TWS/Gateway, you'll need to make sure that they
::   use different values for the API port number. A common convention is as
::   follows:
::
::      TWS (live)        7496
::      TWS (paper)       7497
::      Gateway (live)    4000
::      Gateway (paper)   4001
::
:: CLIENTID
::   [NB: there is no need to set this value unless you need to run two or more
::   instances of this program simultaneously: if the value is not set, the
::   program will use its built-in client id.]
::   Set this to a value that is unique to this program (ie it mustn't be the
::   same clientID that's used by any other API program). Note that if you need
::   to run two or more instances of this program at the same time, they must
::   have different clientIDs. The value must be a positive integer between 1 and
::   999999999.
::
::
:: LOG
::   Set this to the program's log filename. This file records details of
::   program operation, which can be helpful in identifying the reason for
::   program failures. If no value is specified, the filename will be
::   %APPDATA%\TradeWright\gsd\gsd.log
::
:: LOGLEVEL
::   Set this to control the level of detail in the program logfile.
::   Permitted values, in increasing level of detail, are N, D, M and H. For
::   normal operation N is recommended. Note that the higher the level of
::   detail, the larger the logfile.
::
:: APIMESSAGELOGGING
::   This setting controls the logging of API messages and API message 
::   statistics. It should be left at the default value of 'NNN' except under
::   advice from the program developer. API message logging can cause 
::   logfiles to become very large: the default setting turns off all API
::   message logging except for those related to API connections and
::   API errors/notfications.
::
::
:: FILEFILTER
::   Set this to specify the filenames that will contain input commands.
::   For example gsd*.txt or *.gsd. Files whose names do not pass the filter
::   are ignored.
::
:: INPUTFILESDIR
::   Set this to the folder into which input files must be placed for their
::   commands to be actioned.
::
:: ARCHIVEDIR
::   Set this to the folder where input files will be placed after they have
::   been processed. Note that if you move or copy a file from this folder
::   to the input files folder, it will be processed again. When a file has
::   been processed, if the archive directory already contains a file with
::   the same name, it will be overwritten by the new one: thus the archive
::   is not necessarily a complete record of all order files processed.
::
:: OUTPUTDIR
::   Set this to the path for the folder where scan results files are to be
::   stored (unless otherwise specified by commands). You can include a
::   filename and both the path and filename can include substitution
::   variables.
::
:: INSTALLFOLDER
::   Set this to the folder where you installed the TradeBuild Platform.
::   If you installed TradeBuild Platform using the .msi installer with the 
::   default installation location, you do not need to set this value.
::
::
:: PIPELINE
::   Set this to a command into which will be piped any output from the GSD27
::   program resulting from use of a SCANECHO command. Typical use is to pipe
::   the output to the GBD27 program to download historical bars for the
::   contracts returned by the scan(s).
::
::   Note that the ECHO command in GSD72 can be used to run commands in the
::   target program: in the case of the target program being GBD27, the ECHO
::   command could be used to set the required timefrae, start and end times,
::   number of bars, and so on.
::
::   Note also that the ECHORESULTFORMAT command in GSD27 can be used to control
::   the format of the output from SCANECHO commands, such that this output
::   forms one or more valid commands for the target program.
::
::   For example, a typical value of this setting for piping into the GBD27
::   program could be:
::
::
::   set PIPELINE=GBD27 -fromtws:"%TWSSERVER%,%PORT%" -log:"%TOPDIR%\Log\gbd27.log"
::
::
::   This would generate historical bar data for all the contracts listed in
::   the scan output. 'ECHO' commands to GSD could be used to control the ouput
;;   from the GBD27 program.
::
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
"%SCRIPTS%\GetScanData.bat" %~1
