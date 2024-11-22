@echo off
setlocal enableextensions enabledelayedexpansion

::=============================================================================+
::                                                                             +
::   This command file runs the StrategyHost program for use in development,   +
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

set CONTRACT=FUT:ES(0[1d])@CME

set STRATEGY=Strategies27.MACDStrategy21
set STOPLOSSSTRATEGY=Strategies27.StopStrategyFactory5
set TARGETSTRATEGY=

set SIMULATEORDERS=no

set RUN=yes

set RESULTSDIR=%TOPDIR%\Results

set TWSSERVER=
set PORT=7497
set CLIENTID=
set CONNECTIONRETRYINTERVAL=

set LOG=%TOPDIR%\Log\StrategyHost.log
set LOGLEVEL=N
set LOGOVERWRITE=yes
set LOGBACKUP=yes

set INSTALLFOLDER=


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
:: CONTRACT
::   Specifies the contract against which the specified strategy is to be run.
::   Any TradeBuild contract specification format can be used, for example
::   "FUT:ES(0,1d)@CME" is the current ES future trading on CME, switching to
::   the next contract 1 day before expiry.
::
:: STRATEGY
::   The prog id for the trading strategy to be executed. This is of the form:
::
::      <dll>.<classname>
::
::   where 
::      <dll> is the filename of the dll that contains the trading strategy
::      <classname> is the name of the class within the dll that implements the
::      strategy
::
:: STOPLOSSSTRATEGY
::   The prog id for the stop-loss management strategy to be executed, if any. 
::   This is of the form:
::
::      <dll>.<classname>
::
::   where 
::      <dll> 	is the filename of the dll that contains the stop-loss management
::      	strategy
::      <classname> is the name of the class within the dll that implements the
::      	stop-loss management strategy
::
:: TARGETSTRATEGY
::   The prog id for the target management strategy to be executed, if any. 
::   This is of the form:
::
::      <dll>.<classname>
::
::   where 
::      <dll> 	is the filename of the dll that contains the target management
::      	strategy
::      <classname> is the name of the class within the dll that implements the
::      	target management strategy
::
::
:: SIMULATEORDERS
::   Set this to 'yes', for orders to be simulated rather than being passed to
::   TWS. Set it to 'no', or no value, to use real orders.
::
::
:: RUN
::  Indicates whether the program is to immediately commence execution of the 
::  specified trading strategy. Permitted values are 'yes' and 'no'. The
::  default is 'no'.
::
::
:: RESULTSDIR
::   Set this to the folder where output files are to be stored (these files
::   record details of orders placed, and the outcome of completed bracket
::   orders).
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
::   999999999. 
::
::   If you don't supply a value, 723 will be used.
::
:: CONNECTIONRETRYINTERVAL
::   This specifies how frequently the program will attempt to reconnect to
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
:: LOGLEVEL
::   Set this to control the level of detail in the program logfile.
::   Permitted values, in increasing level of detail, are N, D, M and H. For
::   normal operation N is recommended. Note that the higher the level of
::   detail, the larger the logfile.
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
"%SCRIPTS%\StrategyHost.bat"
