# GETBARDATA COMMAND LINE PROGRAM

## 1. Introduction

The GetBarData program  is a Windows command-line utility that retrieves historical bar data either from TWS,
or from the TradeBuild historical database, or from files.

## 2. How to Install

Go to the Release page at https://github.com/rlktradewright/tradebuild-platform/releases, and download and run
the TradeBuild-2.7.235.msi installer file. This will take you through a straightforward install process, which
will result in the whole TradeBuild platform being installed on your computer, at
C:\Program Files (x86)\TradeWright Software Systems\TradeBuild Platform 2.7.

That folder contains a number of files that you don't need to do anything with unless you intend to develop
your own programs that use TradeBuild, and that is way out of scope of this email. In particular, note that
it is not necessary to do any dll registration.

Installation adds a new folder to your Start menu, called TradeBuild Platform 2.7.  This lists four programs,
none of which are relevant to this topic.

Should you want to uninstall TradeBuild at some point, just go to Control Panel > Programs and Features in
the usual way and find the `TradeBuild Platform 2.7` entry: then right click and select Uninstall.


## 2. How to Run the Historical Data Downloader program

The folder that contains the program is the Bin subfolder (which contains all the TradeBuild programs), and
the program file is called gbd27.exe. Once you've started the program, it will prompt you for input, which
must be in the form of the commands detailed in section 4 below.

To run the program, start a Command Prompt (or Powershell) session in this folder (or you can add this folder
to your path in the usual way).

The program can take historical data from three sources: TWS, text files (where the data has been collected
by the TradeBuild DataCollector program) or the TradeBuild Database (also containing data collected by the
DataCollector).

Note that if you enter this command:

`gbd27 /?`

the program displays a summary help page, and then exits. This information can be useful if you've forgotten
something, but is fairly cryptic.

### 2.1 How to GetHistorical Data from TWS

To get data from TWS, we use a command like this:

`gbd27 /fromTWS`

This assumes that TWS is on the local computer and uses port 7496 for API connections. It uses a default
clientID of 205644991. If your TWS is running on another computer called TWSPC, uses port 7497 for API
conections, and you want to specify a ClientID of 99, then the command would be like this:

`gbd27 /fromTWS:TWSPC,7497,99`

Note that if you need to include spaces between or in the values after the colon, the the whole string must
be enclosed in double quotes. For example:

`gbd27 /fromTWS:"My TWS PC, 7497, 99"`

And note that if you want to run more than one instance of the program simultaneously, you must use a unique
ClientID for each one.

### 2.2. How to GetHistorical Data from the TradeBuild Database

To get data from the TradeBuild database, we use a command like this:

`gbd27 /fromDB:DBServer,SqlServer,Trading`

Here `DBServer` is the computer that hosts the database; `SqlServer` indicates the database software in use (
at present the only other option is `MySql`); and `Trading` is the name of the database.

Historical data is collected to the TradeBuild database using the DataCollector27.exe program.

### 2.3. How to GetHistorical Data from a File

To get data from a TradeBuild tickfile, we use a command like this:

`gbd27 /fromfile:C:\Data\Tickfiles\ESH0\ESH0-Week-20200217.tck`

Here the value of the /fromfile argument is the path and filename of the file that contains the tick data from
which the required bars are to be constructed.

Note that the file must be a TradeBuild tick data file.

Tick data is collected to files using the DataCollector27.exe program.

## 3. Log Files

The program writes information to a log file that can be indispensable when diagnosing problems. If the
program suddenly terminates unexpectedly or outputs a stack trace to the console, then you should access
the log file and attach it to a new Issue on GitHub.

To find the log file, open File Explorer and type or paste the following into the address field and press
return:

`%APPDATA%\TradeWright\gbd`

The log file is called gbd.log.


## 4. Commands

Once the program has started, you'll see a prompt character, which is simply a colon ':'. There are then
several different commands that you can enter. Note that all input is case-insensitive.

Similar to a normal command prompt, you can use the up and down arrow keys at the prompt to cycle through
previous commands which you can then amend and resubmit

Right-clicking at the prompt pastes input from the clipboard.

The commands are:

|  Command      |  Purpose                                                     |
| ------------- | ------------------------------------------------------------ |
| contract      | Specifies the contract that you want data for                |
| from          | Specifies the start date-time for the data                   |
| to            | Specifies the end date-time for the data                     |
| number        | Specifies the number of bars you want to retrieve            |
| timeframe     | Specifies the timeframe of the bars you want                 |
| dateonly      | Specifies whether daily, weekly, monthly and yearly bars contain only the date as timestamp |
| nonsess       | Specifies that you want bars outside the main trading session     |
| sess          | Specifies that you only want bars during the main trading session |
| sessiononly   | Specifies whether you only want bars during the main trading session |
| sessionstarttime | Allows you to specify the session start... and end times       |
| sessionendtime   | ...and end times                                               |
| millisecs        | Specifies that milliseconds are to be included in the bar timestamps         |
| nomillisecs      | Specifies that milliseconds are not to be included in the bar timestamps     |
| start            | Starts retrieving historical data as currently defined by the other commands |
| stop             | Stops the historical data retrieval                                          |
| Ctrl-z           | Exits the program                                                            |

Note that you can enter these commands in any order (except that `start` must obviously be after the
others), and if you make a mistake you can just repeat the command. The latest value you supply will be
used when you enter start.

Here is an example session that connects to TWS on a computer called essy, which uses port 7497 for
the TWS API, and using the default client id:

```
gbd27 -fromtws:essy,7497
Connected to TWS: server=essy port=7497 client Id=205644991
:contract esm0
:timeframe 5 m
:to latest
:number 5
:start
Fetch started for contract ESM0@GLOBEX
Data retrieved from source
2020-04-01 06:05:00,2481.75,2483.75,2479.50,2480.00,2123,1705,0
2020-04-01 06:10:00,2479.75,2484.50,2478.00,2484.50,4448,2586,0
2020-04-01 06:15:00,2484.50,2486.50,2482.00,2484.25,2409,1983,0
2020-04-01 06:20:00,2484.25,2487.50,2480.50,2484.00,2472,2031,0
2020-04-01 06:25:00,2484.25,2485.25,2484.25,2484.50,70,68,0
Fetch completed for contract: ESM0@GLOBEX
Number of bars output:  5
:
```

The following gives more details of each command.

### 4.1 Contract Command

This command specifies the contract that you want data for.

Note that the contract command is not needed if a .tck file is used as the historical data source, since
the file includes the relevant contract definition.

There are several ways of specifying the contract. The simplest is just to use IB's local symbol and (if
necessary) the exchange.

#### Example 1

So for example, to get the March 2020 ES futures contract, use this command:

`contract ESH0`

No exchange is needed here because futures are only traded on a single exchange

#### Example 2

Here's an example future where the exchange is supplied (though it doesn't need to be):

`contract FDAX MAR 20@DTB`

In this case it's the March 2020 DAX future on the DTB exchange. Note that `FDAX MAR 20` is IB's local
symbol for that future.

#### Example 3

If a contract can be smart routed, or if a stock is traded on several exchanges, you'll need to specify
the exchange. For example

`contract MSFT@SMARTUS`

Note that SMARTUS is TradeBuild's way of specifying that you want the US-based SMART-routed contract
(some contracts can be SMART-routed in other continents, for example for Europe use SMARTEUR; for UK
use SMARTUK).

#### Example 4

You can also specify SMART-routing by giving the primary exchange, like this:

`contract MSFT@SMART/ISLAND`

However this is entirely equivalent to SMARTUS.

#### Example 5

You can also specify the contract by explicitly spelling out its attributes. This is most useful for
options:

`contract /SYMBOL:MSFT /SECTYPE:OPT /EXCHANGE:CBOE /EXPIRY:20200124 /STRIKE:150 /RIGHT:C`

but can also be used for any sort of contract. Note that for contracts that expire (futures, options and
futures options), you can specify the expiry without knowing the actual expiry date: just use 0 to mean the
current contract, 1 for the next one, 2 for the one after that and so on (note that negative numbers for
expired contracts are not currently supported). For example:

`contract /SYMBOL:ES /SECTYPE:FUT /EXCHANGE:GLOBEX /EXPIRY:1`

will use the ES futures contract for June 2020 (at the time of writing, the current contract is
March 2020, so it will use the next one).

#### Example 6

Finally, you can specify the contract details in a fixed order separated by commas. They must be in this
order:

localsymbol,sectype,exchange,symbol,currency,expiry,multiplier,strike,right

For example:

`contract ,FUT,GLOBEX,ES,,1`

Other ways are much simpler! But this format could be useful if another program is piping its output into
this one.


### 4.2 From Command

This command specifies the start date-time for the data. You must supply either this start date or a non-zero
number of required bars, or both. If you use this command without a parameter, the from date is reset to
'no date'.

The time supplied must be in the timezone of the relevant exchange.

For example:

`from 2019/12/15 19:25`

Pretty much any common date format can be used.


### 4.3 To Command

This command specifies the end date-time for the data. You can use the special parameter `LATEST` to return
data right up to the current time, and this is also the initial setting if you don't use this command. If
you use this command without a parameter, the to date is reset to 'no date'.

The time supplied must be in the timezone of the relevant exchange.

For example:

`to 2020/01/15`

`to latest`

Again all common date formats can be used.


### 4.4 Number  Command

This command specifies the number of bars you want to retrieve. You can either specify an actual number like
100, or use the special value `all` which means all the bars implied by the from and to dates. 0 is not a
valid value.

If the number of bars you specify is greater than the number implied by the to and from dates, you will only
get the bars between those dates. If it's less than the number implied by the to and from dates, you will
only get the number specified.

For example:

`number 100`

`number ALL`


### 4.5 Timeframe Command

This command specifies the timeframe of the bars you want. For example:

`timeframe 5 m`

means you want 5-minute bars.

`timeframe 13 m`

means you want 13-minute bars.

The number you give as the first parameter can be any positive number (though you'd be unlike to want,
say, 657-minute bars).

The second parameter must be one of the following:

s   seconds
m   minutes (default)
h   hours
d   days
w   weeks
mm   months
v   volume (constant volume bars)
tv  tick volume (constant tick volume bars)
tm   ticks movement (constant range bars)


### 4.6 DateOnly Command

If you are requesting daily, weekly, monthly or yearly bars, you probably only want the bars' timestamp to
include the relevant date, and the not the time of day that the trading session actually started. This command
specifies whether only the date is included. It has no effect on bars of other timeframes, which always include
a time part.

The command has a parameter which must be one of the following, with the obvious meanings:

yes
true
on

no
false
off

If the parameter is not included, 'yes' is assumed.

When the program starts, `dateonly on` is automatically set.

Note that for contracts where the trading session spans midnight, such as the E-Mini futures on Globex, the date
included when `dateonly off` is set depends on whether only bars during the main session are requested (see the
SesionOnly, Sess, and NonSess commands): if so, the date will be for the day that includes the main session;
if not, the date will be for the previous day (when the overall session started) and the timestamp will include
the time that the overall session started.

### 4.7 Nonsess Command

This command specifies that you want to include bars outside the main trading session. It has no parameters.
When the program starts, this is assumed.

Note that for contracts that have expired, TWS doesn't give any information about session start and end times,
so you will need to use the `sessionstarttime` and `sessionendtime` commands to specify them if these boundaries
are important to you.

Note that `nonsess` is exactly equivalent to `sessiononly off`.


### 4.8 Sess Command

This command specifies that you only want bars during the main trading session. It has no parameters. When
the program starts this is assumed not to be the case.

Note that for contracts that have expired, TWS doesn't give any information about session start and end times,
so you will need to use the sessionstarttime and sessionendtime commands to specify them if these boundaries
are important to you.

Note that `sess` is exactly equivalent to `sessiononly on`.


### 4.9 SessionOnly Command

This command specifies whether you only want bars during the main trading session.

The command has a parameter which must be one of the following, with the obvious meanings:

yes
true
on

no
false
off

If the parameter is not included, 'yes' is assumed.

When the program starts, `sessiononly on` is automatically set.

Note that `sessiononly on` is exactly equivalent to `sess`, and `sessiononly off` is exactly equivalent to
`nonsess`.


### 4.10 SessionStartTime and SessionEndTime Commands

These command allows you to specify the session start and end times. You can use these command in situations
where TWS provides no session start information, or to override that information. For example, TWS indicates
that the liquid trading hours for FTSE 100 Futures contracts is from 01:00 to 21;00, but in practice
significant trading only happens between 08:00 and 17:30.

The commands take a single parameter which must be a time of day;

For example:

`sessionstarttime 08:00`
`sessionendtime 17:30`


### 4.11 Millisecs Command

This command specifies that milliseconds are to be included in the bar timestamps. When the program starts,
this is assumed to be false. In practice this is never useful when information is being sourced from TWS,
as TWS historical data is never timestamped at the sub-second level.

The command has a parameter which must be one of the following, with the obvious meanings:

yes
true
on

no
false
off

If the parameter is not included, 'no' is assumed.

When the program starts, `millisecs off` is automatically set.

Note that `millisecs off` is exactly equivalent to `nomillisecs`.


### 4.12 NoMillisecs Command

This command specifies that milliseconds are not to be included in the bar timestamps. When the program starts,
this is assumed to be true. In practice this is never useful when information is being source from TWS,
as TWS historical data is never timestamped at the sub-second level.


### 4.13 Start Command

This command starts the historical data retrieval process as currently defined by the other commands.

The program takes account of the historical data pacing rules enforced by IB. This can mean that large
retrievals can take quite a long time, and the program may appear to be sitting doing nothing, when it is
actually spacing out the historical data requests to stay within the rules.

Note that it doesn't use a crude mechanism of simply spacing the requests every 11 seconds or whatever. It
actually records which requests are outstanding and when they were made, and then uses the rules to determine
when the next one can be submitted.

It should also be noted that if there is more than one instance of the program running, or there are other
TWS API programs that are also making historical data requests, these programs are unaware of each other and
will implement the rules individually – however TWS applies the rules collectively across all API clients, so
this may lead to problems. Therefore it's suggested that you only run one instance at a time, unless the
requests are for relatively small numbers of bars for different contracts, but it's difficult to give hard and
fast criteria for this.


### 4.14 Stop Command

Stops the historical data retrieval process. You can use this command at any time after the start command,
even while the bar data is being output. It can be useful if you've inadvertently requested a large amount of
data (for example a year's worth of 15-minute bars).


## 5. How To Exit the Program

To end your session with the program, press Ctrl-Z at the prompt and then press return.


## 6. How To Output Data to a File

The program output can be redirected to a file using the standard command-line output redirection operators
`>` and `>>`. For example:

`gbd27 /fromTWS:TWSPC,7497,99 > C:\BarData\bars.txt`

creates or overwrites C:\BarData\bars.txt.

`gbd27 /fromTWS:TWSPC,7497,99 >> C:\BarData\bars.txt`

creates or appends to C:\BarData\bars.txt.

Note that such redirection applies to all output during a run of the program (other than error messages and
informational comments). It is intended to provide additional commands to specify an output file for a
particular retrieval operation.


## 7. How To Input Commands from a File or Another Program

The standard command-line input redirection operator `<` can be used to read commands from a file.

For example, create a file called `C:\Qaz\getbars.txt` containing the following commands:

```
contract /symbol:ES /sectype:fut /expiry:0 /exchange:globex
timeframe 5 mins
from 2020/02/17
to 2020/02/18
start
```

Now run a command like this:

`gbd27 -fromTWS <  C:\Qaz\getbars.txt > C:\Qaz\Bars.txt`

The commands will be read from the `getbars.txt` file, and the output will be sent to `Bars.txt`.

Similarly the standard pipe operator '|' can be used to pass output from another program as input to gbd27.
For example:

`type E:\qaz\GetBars.txt | gbd27 -fromTWS > C:\Qaz\Bars.txt`






