@echo off
setlocal enableextensions enabledelayedexpansion

::=============================================================================+
::                                                                             +
::   This command file runs the plord27.exe program for placing orders to      +
::   IBKR's Trader Workstation or Gateway.                                     + 
::                                                                             +
::   If the first argument is /I, the program is run in interactive mode,      +
::   meaning that you must type order specifications manually. Any other       +
::   arguments are ignored.                                                    +
::                                                                             +
::   If the first argument is a filename, the program reads the file from      +
::   the orderfiles directory and actions it. Any other arguments are          +
::   ignored.                                                                  +
::                                                                             +
::   If there are no arguments, the program monitors the orderfiles directory  +
::   and actions any order files that are subsequently laced in it.            +
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

set TOPDIR=%SYSTEMDRIVE%\PlaceOrders

set MONITOR=yes

set TWSSERVER=
set PORT=7497
set CLIENTID=27236
set CONNECTIONRETRYINTERVAL=

set LOG=%TOPDIR%\Log\plord27.log
set LOGLEVEL=N
set APIMESSAGELOGGING=NNN

set FILEFILTER=Orders*.txt
set ORDERFILESDIR=%TOPDIR%\OrderFiles
set ARCHIVEDIR=%TOPDIR%\Archive
set RESULTSDIR=%TOPDIR%\Results

set BATCHORDERS=no
set STAGEORDERS=no

set SIMULATEORDERS=no

set SCOPENAME=%CLIENTID%
set RECOVERYFILEDIR=%TOPDIR%\Recovery


::              PLEASE DON'T CHANGE ANYTHING BELOW THIS LINE !!
::==============================================================================
::
::   Notes:
::
::
::
:: TOPDIR	
::   Set this to the location where the installation zip was extracted.
::
::
:: MONITOR (NB: this setting is under review, and may be changed in future)
::   Set this to 'yes' if you want the program to create files that contain
::   information about how trades were executed. If you set it to 'no',
::   the only way to find out about what happened is via the TWS Trade Log.
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
::   Set this to a value that is unique to this program (ie it mustn't be the
::   same clientID that's used by any other API program). Note that if you need
::   to run two instances of this program at the same time, they must have
::   different clientIDs. The value must be a positive integer between 1 and
::   999999999. Do not change this setting if you have created positions using
::   this program that have not yet been closed out: if you do, you will have
::   to close them out by other means (for example using TWS).
::
:: CONNECTIONRETRYINTERVAL
::   This specifies how frequently the program will attempt to reconnection to
::   TWS/Gateway after failing to connect, or losing the connection. The value
::   is a number of seconds, with a default of 60.
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
:: APIMESSAGELOGGING
::   This setting controls the logging of API messages and API message 
::   statistics. It should be left at the default value of 'NNN' except under
::   advice from the program developer. API message logging can cause 
::   logfiles to become very large: the default setting turns off all API
::   message logging except for those related to orders, API connections and
::   API errors/notfications.
::
::
:: FILEFILTER
::   Set this to specify the filenames that will contain order
::   specifications. For example TSOrder*.txt or *.ord. Files whose names do
::   not pass the filter are ignored.
::
:: ORDERFILESDIR
::   Set this to the folder into which order files must be placed for their
::   order specifications to be actioned.
::
:: ARCHIVEDIR
::   Set this to the folder where order files will be placed after they have
::   been processed. Note that if you move or copy a file from this folder
::   to the order files folder, it will be processed again. When a file has
::   been processed, if the archive directory already contains a file with
::   the same name, it will be overwritten by the new one: thus the archive
::   is not necessarily a complete record of all order files processed.
::
:: RESULTSDIR
::   Set this to the folder where output files are to be stored (these files
::   record details of orders placed, and the outcome of completed bracket
::   orders).
::
::
:: BATCHORDERS
::   Set this to 'yes', if you want orders to be submitted only when an
::   ENDORDERS command is input. If you set it to 'no', or no value, each 
::   order will be submitted as soon as its definition is complete (so
::   ENDORDERS commands are not needed).
::
:: STAGEORDERS
::   Set this to 'yes', if you want orders to be held in TWS for manual
::   placement by the user. Set it to 'no', or no value, for orders to be
::   actioned immediately by TWS. It is unwise to set this to 'yes' if using
::   the Gateway rather than TWS, since the Gateway provides no means for the
::   user to action staged orders.
::
::
:: SIMULATEORDERS
::   Set this to 'yes', for orders to be simulated rather than being passed to
::   TWS. Set it to 'no', or no value, to use real orders.
::
::
:: SCOPENAME
::   Set this to a name that identifies the set of orders placed via this
::   instance of the program. By default, the ClientId is used, but you can
::   set this to any name you like. Note that each instance of the program
::   must have a different value for this setting.
::
:: RECOVERYFILEDIR
::   Set this to the directory where order recovery information will be stored.
::   If no value is supplied, %APPDATA%\TradeWright\plord will be used. Do not 
::   change this setting if you have created positions using this program, 
::   until they have all been closed out.
::
::
::
::   End of Notes
::==============================================================================

%TOPDIR%\CommandFiles\Scripts\PlaceOrders.bat %~1
