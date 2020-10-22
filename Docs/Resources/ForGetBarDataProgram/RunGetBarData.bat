@echo off
setlocal enableextensions enabledelayedexpansion

::=============================================================================+
::                                                                             +
::   This command file runs the gbd27.exe program for downloadig historical    +
::   data from IBKR's Trader Workstation or Gateway.                           + 
::                                                                             +
::   If the first argument is /I, the program is run in interactive mode,      +
::   meaning that you must type commnds manually. Any other arguments are      +
::   ignored.                                                                  +
::                                                                             +
::   If the first argument is a filename, the program reads the file from      +
::   the INPUTFILES directory and actions it. Any other arguments are          +
::   ignored.                                                                  +
::                                                                             +
::   If there are no arguments, the program monitors the INPUTFILES            +
::   directory and actions any files that are subsequently placed in it.       +
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

set TOPDIR=%SYSTEMDRIVE%\GetBarData

set TWSSERVER=
set PORT=7497
set CLIENTID=

set LOG=%TOPDIR%\Log\gbd27.log
set LOGLEVEL=N
set APIMESSAGELOGGING=NNN

set FILEFILTER=gbd*.txt
set INPUTFILESDIR=%TOPDIR%\InputFiles
set ARCHIVEDIR=%TOPDIR%\Archive
set OUTPUTDIR=%TOPDIR%\BarData\{$contract}\Bars({$fromdatetime}-{$todatetime}).txt

set BIN=


::              PLEASE DON'T CHANGE ANYTHING BELOW THIS LINE !!
::==============================================================================
::
::   Notes:
::
::
::
:: TOPDIR	
::   Set this to the root folder for use of the GetBarData program.
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
::   %APPDATA%\TradeWright\gbd\gbd.log
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
::   For example gbd*.txt or *.gbd. Files whose names do not pass the filter
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
::   Set this to the path for the folder where downloaded historical data files
::   are to be stored (unless otherwise specified by commands). You can include
::   a filename and both the path and filename can include substitution
::   variables.
::
:: BIN
::   Set this to the folder that contains the TradeBuild Platform programs.
::   If you installed TradeBuild Platform using the .msi installer using the 
::   default installation location, there should be no reason to set this value.
::
::
::
::
::   End of Notes
::==============================================================================

Scripts\GetBarData.bat %~1
