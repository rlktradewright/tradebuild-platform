# **PLACEORDERS COMMAND LINE PROGRAM**

## 1. Introduction

The PlaceOrders program is a Windows command-line utility that provides a means 
for submitting orders to Interactive Brokers' Trader Workstation (TWS) or 
Gateway via IBKR's API, without the need for programming.

It can be used interactively, where the user manually types in the order 
specifications. Alternatively, order specifications can be written to files 
that are dropped into a specified folder - when a new file is placed there, 
the program immediately execute the order(s) as specified.

There is also a sample Excel worksheet that shows how to use the program to 
submit orders defined within Excel.

Additionally, the program can accept input from any other program that can write 
the correct syntax to StdOut.

Finally, the code which provides the command syntax analysis and order 
management is available as part of the TradeBuild API, enabling it to be 
incorporated into other programs (but not tonight, Josephine...).

###  Program Overview

The program is command driven. The same set of commands applies regardless of 
where the input is coming from. A command is a single line of input with a 
defined syntax. Some commands can be used at any time; others can only be used 
in specific contexts, dependent on what commands have been used previously.

Some commands result in immediate output, which appears on the screen (though it 
may be piped to another program or redirected to a file, using standard command 
line mechanisms). Other commands produce no output. Output may also occur 
'asynchronously', typically showing the progress of orders between placement and 
final execution.

###  Running the PlaceOrders Program

The program has a number of command line arguments which control various aspects 
of its operation. A command file called `RunPlaceOrders.bat` is provided which 
greatly simplifies the task of getting it started. 

blah blah blah

### # Interactive Usage

blah blah blah

### # FileReader Program Usage

blah blah blah

### # Excel Usage

blah blah blah

### # API Usage

blah blah blah

## 2. Commands

A command is a single line of text with a specified syntax. All commands have 
the same general form, consisting of a commmand code, optionally followed by a 
list of arguments. The command code is separated from the arguments by one or 
more spaces, as are the arguments themselves. 

A special case is the comment command, which has a command code of '#' and does
not have arguments: anything after the # on the same line is simply ignored.

Arguments can be either positional or tagged:

* A positional argument consists only of a value. If the value includes any 
  spaces, it must be surrounded by double quotes (which do not form part of 
  the value).

  Positional arguments must occur in a specific order: a positional 
  argument's meaning is defined by its position relative to other 
  positional arguments.

* A tagged argument consists of a tag and an optional value, separated by a 
  colon ':'. The tag starts with a solidus '/' and is followed by an 
  identifier that does not contain spaces. If the value is present and 
  contains any spaces, it must be surrounded by double quotes (which do not 
  form part of the value). If the value is not present, the colon separator 
  can also be omitted.

  Tagged arguments may appear anywhere in the list of arguments, and in any 
  order. The tag identifies the meaning of the value. 

  Tagged arguments may also be referred to as 'attributes'.

  If a given taggged argument appears more than once in the argument list, 
  only the first instance is relevant.

  If a tagged argument is included that is not part of the command's 
  defined syntax, it may be ignored or an error may result. In the long 
  term it is intended that all such occurrences will be treated as errors.

These rules mean that the following sample commands are equivalent:

BUY 1 STPLMT 3114.25 3114.00 /TIF:DAY /IGNORERTH

BUY 1 STPLMT /TIF:DAY /IGNORERTH 3114.25 3114.00

BUY /TIF:DAY 1 STPLMT 3114.25 /IGNORERTH 3114.00

###  List of commands

The following is a summary of all the commands currently defined:

|  Command      |  Purpose                                                     |                   
| ------------- | ------------------------------------------------------------ |
|  #            | Starts a comment line                                        |            
|  ?            | Outputs a list of commands that are currently valid          |
|  BATCHORDERS  | Specifies whether bracket orders should be accumulated and only submitted when an ENDORDERS command is received                                                  |
|  BRACKET      | Starts a bracket order specification                         |
|  BUY          | Specifies a single buy order                                 |
|  B            | Repeats the previous BUY command                             |
|  CLOSEOUT     | Closes all positions in one or all groups                    |
|  CONTRACT     | Specifies a contract, which becomes the current contract in the current group |
|  ENDBRACKET   | Ends a bracket order specification                           |
|  ENDORDERS    | Submits a batch of bracket orders for execution              |
|  ENTRY        | Specifies the entry order of a bracket order                 |
|  EXIT         | Ends the session and terminates the program                  |
|  GROUP        | Defines a new group or switches to an existing group         |
|  HELP         | Outputs the syntax summary                                   |
|  LIST         | List current groups, positions or trades                     |
|  PURGE        | Removes all knowledge of a group, without affecting any orders that have been defined and submitted in that group                                                    |
|  QUIT         | Aborts the current bracket order definition                  |
|  QUOTE        | Displays bid, ask and last prices and sizes for the current or a specified contract |
|  RESET        | Cancels any bracket orders that have not yet been submitted, and any bracket order specifications that have not been completed                                    |
|  SELL         | Specifies a single sell order                                |
|  S            | Repeats the previous SELL command                            |
|  STAGEORDERS  | Specifies that orders are to be sent to TWS but not transmitted to the broker for execution (manual intervention in TWS is required)                             |
|  STOPLOSS     | Specifies the stop-loss order of a bracket order             |
|  TARGET       | Specifies the target order of a bracket order                |

## 3. Contracts

A contract specification defines a specific security for which you can place 
orders.

You use the CONTRACT command to specify a contract. This can take two forms:

* Full contract specification: this contains a number of tagged arguments that 
  define various characteristics of the desired contract

* Abbreviated contract specification: this uses just the 'local symbol' of the 
  contract and (if need be) the relevant exchange

When a contract is specified in an interactive session, its local name and 
exchange are subsequently included in the command prompt.

A CONTRACT command only applies to the current group (see [Groups](Groups)). 
If the same contract is to be used in another group, another CONTRACT command 
must be issued after switching to that group (or in the GROUP command itself).

###  Full Contract Specification

The full contract specification can contain the tagged arguments listed in 
the [CONTRACT command specification](contract-command). Note that it is only 
necessary to specify sufficient attributes to uniquely identify the contract.

Examples:

CONTRACT /SYMBOL:ES /SECTYPE:FUT /EXPIRY:201912 /EXCHANGE:GLOBEX

CONTRACT /SYMBOL:ES /SECTYPE:FOP /EXPIRY:20191220 /STRIKE:3100 /RIGHT:C

CONTRACT /SYMBOL:MSFT /EXCHANGE:SMARTUS

CONTRACT /SYMBOL:DAX /SECTYPE:FUT /EXPIRY:201912 /MULTIPLIER:5

CONTRACT /SYMBOL:DAX /SECTYPE:FUT /EXPIRY:201912 /MULTIPLIER:25

CONTRACT /SYMBOL:MSFT /SECTYPE:OPT /EXCHANGE:CBOE /EXPIRY:1 /STRIKE:160 /RIGHT:C

###  Abbreviated Contract Specification

The abbreviated contract specification uses IBKR's local symbol for the 
contract, and where necessary the exchange. The two items are separated by an 
'@' character, which is omitted if the exchange is not supplied.

* For futures and futures options contracts, the local symbol alone is 
  sufficient to identify the contract, so the exchange need not be supplied.

* For stocks that are traded on multiple exchanges, the exchange is always 
  required.

* For options, the exchange is always required.

Note that local symbols can be found from the Symbol column in TWS ticker 
pages. Alternatively use the 
[Contract Inspector](https://github.com/tradewright/ibapi-tools/blob/master/ContractInspector/readme.md).

Examples (these are equivalent to the corresponding full contract 
specification examples above):

CONTRACT ESZ9

CONTRACT "ESZ9 C3100"

CONTRACT MSFT@SMARTUS

CONTRACT "FDXM DEC 19"

CONTRACT "FDAX DEC 19"

CONTRACT "MSFT  191213C00150000@CBOE"

## 4. Groups

Groups provide a way of organizing related orders. For example if your trading 
strategy involves placing orders for more than one contract, you can define a 
group for these orders, and then subsequently use the group name for reporting 
on profit/loss, or executions, and if need be for closing out all the positions 
in the group.

At any particular time, there is a current group, and there may be a current 
contract within that group. 

There is a default group with the special name '$'. This is the current group 
when the program starts, until another group is specified. You can switch back 
to the default group at any time, just like any other group.

A new group is defined using the GROUP command. A group name must start with a 
letter or digit, and this can be followed by any number of letters, digits, 
underscores '_' or hyphens '-'.

The GROUP command is also used to switch between groups: thus, after a GROUP 
command that group becomes the current group, and any contract commands or order 
commands relate to that group only, until another group is made current.

Group names are case-insensitive, but the program will always display a group 
name exactly as you first specified it.

In interactive usage, the current group name is included in the command prompt, 
along with the current contract (if any).

Note that in a GROUP command, you can follow the group name with a contract 
specification. For example:

GROUP MyGlobexFutures ESZ9

In this example, 'MyGlobexFutures' becomes the current group and 'ESZ9' becomes 
the current contract in that group. In interactive mode, the command prompt will 
now appear like this:

MyGlobexFutures!ESZ9@GLOBEX:

You can obtain a list of all groups known to the program using the LIST GROUPS 
command. The list includes the current contract (if any) for each group, and the 
current group is indicated.

The PURGE command can be used to make the program discard all information 
relating to one or all groups. This has no impact on any outstanding orders 
associated with the group(s), and it is up to the user to take any necessaary 
actions (for example using TWS directly) such as closing out unwanted positions. 
The PURGE command is intended for use where something appears to have gone 
wrong: for example information output by the program differs from information 
shown by TWS.

## 5. Positions

A position is the overall number of shares, options, futures or currency within 
a particular group relating to actual executions for a particular contract. Note 
that within a group you can place any number of orders for the same contract at 
different times: the position at any time is the total number taking account of 
all fills so far for all those orders.

A pending position is the overall number of shares, options, futures or currency 
within a particular group relating to orders for a particular contract that have 
not yet fully filled. This includes for example limit orders waiting at the 
exchange to be matched, or stop orders waiting for the trigger price to be hit, 
as well as orders that have partially filled.

The LIST POSITIONS command shows the current positions and pending positions for 
each group, including the current profit/loss for each position. 

Note that information about groups and positions is stored between sessions. So 
LIST POSITIONS will show information about groups and positions from previous 
sessions, unless they were flat with no pending positions. Positions that have 
become flat (with no pending positions) while the program has been running will 
be included in LIST POSITIONS in the next run, but will be forgotten in 
subsequent runs.

## 6. Orders

An order is an instruction to buy or sell a defined quantity of a particular 
contract. There are two types of order recognised by the program:

* Single orders: these consist of a simple buy or sell order with no associated  
  orders. When filled they increase or decrease the current position for the 
  relevant contract in the group via which the order was placed. 

* Bracket orders: these consist of an entry order that is used to increase or 
  decrease the current position for the relevant contract in the group via which 
  the order was placed, together with optional stop-loss and target orders that 
  (when fully filled) will reduce the overall position by the same amount. The 
  stop-loss and target orders are only activated when the entry order receives 
  its first fill. Once either of the stop-loss or target orders is filled, the 
  other is cancelled.

Note that as a matter of implementation convenience, single orders are actually 
implemented as bracket orders with an entry order but without the associated 
stop-loss or target orders. However the single orders have a more succinct 
syntax (though the full bracket order syntax can also be used to define a 
single order).

The LIST TRADES command lists all executions that have occurred for all 
currently defined groups.

###  Order Types for Entry Orders

Entry orders (and single orders) may use the following order types:

Limit
Limit If Touched
Limit On Close
Limit On Open
Market
Market If Touched
Market On Close
Market On Open
Market To Limit
Stop
Stop Limit
Trail
Trail Limit

###  Order Types for Stop-Loss Orders

Stop-loss orders may use the following order types:

Stop
Stop Limit
Trail
Trail Limit

###  Order Types for Target Orders

Target orders may use the following order types:

Limit
Limit If Touched
Limit On Close
Limit On Open
Market If Touched
Market On Close
Market On Open

## 7. Order Pricing

Price specifications have two parts: a base price and an optional offset. The 
offset adds or subtracts an amount from the base price to yield the actual price 
with which the order is submitted.

Note that a price is considered to be 'more aggressive' than another price if it 
has a greater likelihood of being filled, and 'less aggressive' if it has a 
lesser likelihood of being filled. 

Price specifications are not processed until immediately before an order is 
actually placed with the broker.

Calculating a price may depend on information that is not available at the time 
the price specification is defined (for example the last traded price might not 
be available for a period after market open): in these circumstances, the 
order will be held until the required information becomes available.

Where the price yielded by a price specification is not an exact multiple of a 
tick, it is rounded to the nearest more aggressive price for a buy, or to the 
nearest less aggressive price for a sell.

###  Base Price

The base price can be either a numeric value, or one of the following 
identifiers which have special meanings:

* BID - use the current bid price

* ASK - use the current ask price

* MID - use the price that's midway between the current bid and ask prices, 
  rounded if need be to the nearest more aggresssive tick boundary

* BIDASK - use the current bid if the order is a buy, or the current ask if it 
  is a sell

* LAST - use the most recent execution price

* ENTRY - use the price at which the entry order for this bracket order was 
  filled (this only applies to Stop-loss and Target orders). Note that where the 
  entry order has multiple fills, the price from the first fill is used

###  Offset

An offset is expressed as a numeric value (the multiplier) followed by an 
optional qualifier character, which determines the meaning of the numeric value. 
The whole is enclosed in square brackets '[' and ']' and immediately follows the 
base price. If the multiplier is positive, then the offset is added to the 
base price; if negative, it is subtracted from the base price. 

EXCEPTION: note, however, that in the CLOSEOUT command, where it may not be 
known whether the position being closed is long or short, or where the base 
price is BIDASK, the meaning of the qualifier is altered: a positive multiplier 
means make the price more aggressive, and a negative multiplier means make the 
price less aggressive. 

Permitted qualifiers are:

* T - the offset is a multiple of the contract's ticksize (ie minimum price 
  variation)

* % - the offset is a percentage of the base price

* S - the offset is a percentage of the bid/ask spread

If no qualifier is specified, the offset is simply the numeric value.

###  Examples

1. Price specifier: 3114.25

    The price is specified directly. There is no offset.

2. Price specifier: BID

    The price to be used is the current bid price.

3. Price specifier: BID[1T]

    The price to be used is 1 tick above the current bid price.

4. Price specifier: BID[40S]

    The price to be used is 40% of the current bid/ask spread above the current 
    bid price.

5. Price specifier: MID[-2T]

    The price to be used is 2 ticks below the current mid-price of the bid/ask 
    spread,

6. Price specifier: ENTRY[-2%] 

    The price to be used is 2% below the entry price obtained for this bracket 
    order. This might for example be used as the trigger price for a trailing 
    stop order used as a stop-loss.

## 8. Command Reference

This section provides the detailed syntax and effect of every command.

</br></br>
### # Command

Starts a comment line. 

This is a special command that does not have the normal command syntax. 
Anything after the initial '#' is ignored, except that the complete command
is output to the console and recorded in the various program logs.

</br></br>
### ? Command

Outputs a list of commands that are currently valid.

Positional aruments: None

Tagged arguments: None

</br></br>
### BATCHORDERS Command

Specifies whether bracket orders should be accumulated and only submitted when 
an ENDORDERS command is entered.

Batching orders in this way is useful where a number of related orders must 
either all be submitted or none must be submitted. If an error occurs during 
one of the order specifications, then none of them will be submitted when 
ENDORDERS is entered.

Positional aruments:

| Position | Permitted&nbsp;values | Meaning                                      |
| -------- | ---------------- | -------------------------------------------- |
| 0        | YES              | Bracket orders are not submitted until and ENDORDERS command is entered |
|          | NO               | Bracket orders are submitted as soon as the ENDBRACKET command is entered |

Tagged arguments: None

</br></br>
### BRACKET Command

Starts a bracket order specification.

The bracket order specification is terminated with an [ENDBRACKET](ENDBRACKET) command, and 
there must be at least an [ENTRY](ENTRY) command before the ENDBRACKET command.

The bracket order specification may also include [STOPLOSS](STOPLOSS) and/or [TARGET](TARGET)
commands.

Syntax summary:

BRACKET \<buyorsell\> \<quantity\> [\<attribute\>...]

Positional aruments:

| Position | Permitted&nbsp;values | Meaning                                      |
| -------- | ---------------- | -------------------------------------------- |
| 0        | BUY              | This bracket order will buy on entry         |
|          | SELL             | This bracket order will sell on entry        |
| 1        | integer > 0      | The quantity to be bought or sold            |

Tagged arguments: 

| Tag            | Permitted&nbsp;values | Meaning                                      |
| -------------- | ---------------- | -------------------------------------------- |
| /cancelafter   | integer > 0     | The time in seconds after which this bracket order will be automatically cancelled. |
| /cancelprice   | \<price\>       | The traded market price at which the order will be cancelled: |
|                |                 | If entry order is LMT, MTL or MIT and BUY then cancel when market rises above \<price\> |                                  
|                |                 | If entry order is LMT, MTL or MIT and SELL then cancel when market falls below \<price\> |                                  
|                |                 | If entry order is STP or STPLMT and BUY then cancel when market falls below \<price\> |                                  
|                |                 | If entry order is STP or STPLMT and SELL then cancel when market rises above \<price\> |                                  
| /description   | \<string\>      | The specified string is included in log entries relating to this bracket order |
| /goodaftertime | \<datetime\>    | The bracket order is not to be submitted until the specified date and time |
| /goodtilldate  | \<datetime\>    | The bracket order will be cancelled if the entry order is still unfilled at the specified date and time |
| /timezone      | \<timezonename\> | Specifies the timezone for the date/times in the /goodaftertime and /goodtilldate attributes |

</br></br>
### BUY Command

Specifies a single buy order. The order specification is processed immediately, regardless of whether bracket order batching has been set ON with the [BATCHORDERS](BATCHORDERS]) command.

This command has two forms: 

* the first form includes an abbreviated contract specification as the first (positional) argument - this contract then becomes the current contract for the current group 

* the second omits this, using the current contract instead.

Syntax summary:

Form 1:

BUY \<contractspec\> \<quantity\> \<ordertype\> [\<pricespec1\>  [\<pricespec2\>]] [\<attribute\>...]

Form 2:

BUY \<quantity\> \<ordertype\> [\<pricespec1\> [\<pricespec2\>]] [\<attribute\>...]


Positional aruments (form 1):

| Position | Permitted&nbsp;values | Meaning                                      |
| -------- | ---------------- | -------------------------------------------- |
| 0        | \<contractspec\> | The contract to be bought or sold. This must be an [abbreviated contract specification](abbreviated-contract-specification). This contract becomes the current contract for the current group
| 1        | integer > 0      | The quantity to be bought or sold            |
| 2        | LMT              | The type of order to use.                    |
|          | LIT              |                                              |
|          | LOC              |                                              |
|          | LOO              |                                              |
|          | MKT              |                                              |
|          | MIT              |                                              |
|          | MOC              |                                              |
|          | MOO              |                                              |
|          | MTL              |                                              |
|          | STP              |                                              |
|          | STPLMT           |                                              |
|          | TRAIL            |                                              |
|          | TRAILLMT         |                                              |
| 3        | \<pricespec\>    | Limit price, if one is required; otherwise trigger price, if one is required. This applies to the following order types |
|          |                  | LMT                                          |  
|          |                  | LIT                                          |
|          |                  | LOC                                          |
|          |                  | LOO                                          |
|          |                  | MIT                                          |
|          |                  | STP                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAIL                                        |
|          |                  | TRAILLMT                                     |
| 4        | \<pricespec\>    | Trigger price, if one is required and was not specified in positional argument 3. This applies to the following order types |
|          |                  | LIT                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAILLMT                                     |

Positional aruments (form 2):

| Position | Permitted&nbsp;values | Meaning                                      |
| -------- | ---------------- | -------------------------------------------- |
| 0        | integer > 0      | The quantity to be bought or sold            |
| 1        | LMT              | The type of order to use.                    |
|          | LIT              |                                              |
|          | LOC              |                                              |
|          | LOO              |                                              |
|          | MKT              |                                              |
|          | MIT              |                                              |
|          | MOC              |                                              |
|          | MOO              |                                              |
|          | MTL              |                                              |
|          | STP              |                                              |
|          | STPLMT           |                                              |
|          | TRAIL            |                                              |
|          | TRAILLMT         |                                              |
| 2        | \<pricespec\>    | Limit price specifier, if one is required; otherwise trigger price, if one is required. This applies to the following order types |
|          |                  | LMT                                          |  
|          |                  | LIT                                          |
|          |                  | LOC                                          |
|          |                  | LOO                                          |
|          |                  | MIT                                          |
|          |                  | STP                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAIL                                        |
|          |                  | TRAILLMT                                     |
| 3        | \<pricespec\>    | Trigger price specifier, if one is required and was not specified in positional argument 2. This applies to the following order types |
|          |                  | LIT                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAILLMT                                     |

Tagged arguments: 

| Tag           | Permitted&nbsp;values | Meaning                                      |
| ------------- | ---------------- | -------------------------------------------- |
| /cancelafter  | hh:mm:ss         | The time after which this order will be automatically canceeled |
| /cancelprice   | \<price\>       | The traded market price at which the order will be cancelled: |
|                |                 | * if order is LMT, MTL or MIT and BUY then cancel when market rises above \<price\> |                                  
|                |                 | * if order is LMT, MTL or MIT and SELL then cancel when market falls below \<price\> |                                  
|                |                 | * if order is STP or STPLMT and BUY then cancel when market falls below \<price\> |                                  
|                |                 | * if order is STP or STPLMT and SELL then cancel when market rises above \<price\> |                                  
| /description   | \<string\>      | The specified string is included in log entries relating to this bracket order |
| /goodaftertime | \<datetime\>    | The bracket order is not to be submitted until the specified date and time |
| /goodtilldate  | \<datetime\>    | The order will be cancelled if the entry order is still unfilled at the specified date and time |
| /ignorerth     | n/a             | The order will be actioned if placed outside Regular Trading Hours. Otherwise it will be held up at the IBKR servers until market open |
| /timezone      | \<timezonename\> | Specifies the timezone for the date/times in the /goodaftertime and /goodtilldate attributes |

</br></br>
### B Command

Repeats the previous BUY command. This command is intended for interactive use 
during scalping, for example to rapidly repeat a buy order 1 tick above the current bid price.

Positional aruments: None

Tagged arguments: None
</br></br>

### CLOSEOUT Command

Closes all positions and pending positions in one or all groups.

This command has two forms: 

* the first form closes out the current group 

* the second closes out the specified group, or all groups if 'ALL' is supplied 
  in the first positional argument

Syntax summary:

Form 1:

CLOSEOUT [\<ordertype\> [\<pricespec1\> [\<pricespec2\>]]] [\<attribute\>...]

Form 2:

CLOSEOUT \<groupname\> | ALL [\<ordertype\> [\<pricespec1\> [\<pricespec2\>]]] [\<attribute\>...]

Positional aruments (form 1):

| Position | Permitted&nbsp;values | Meaning                                      |
| -------- | ---------------- | -------------------------------------------- |
| 0        | MKT              | The type of order to use.                    |
|          | LIT              |                                              |
|          | LOC              |                                              |
|          | LOO              |                                              |
|          | MKT              |                                              |
|          | MIT              |                                              |
|          | MOC              |                                              |
|          | MOO              |                                              |
|          | MTL              |                                              |
|          | STP              |                                              |
|          | STPLMT           |                                              |
|          | TRAIL            |                                              |
|          | TRAILLMT         |                                              |
| 1        | \<pricespec\>    | Limit price, if one is required; otherwise trigger price, if one is required. This applies to the following order types |
|          |                  | LMT                                          |  
|          |                  | LIT                                          |
|          |                  | LOC                                          |
|          |                  | LOO                                          |
|          |                  | MIT                                          |
|          |                  | STP                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAIL                                        |
|          |                  | TRAILLMT                                     |
| 2        | \<pricespec\>    | Trigger price, if one is required and was not specified in positional argument 3. This applies to the following order types |
|          |                  | LIT                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAILLMT                                     |

Positional aruments (form 2):

| Position | Permitted&nbsp;values | Meaning                                      |
| -------- | ---------------- | -------------------------------------------- |
| 0        | ALL              | Specifies that all groups are to be closed out |
|          | \<groupname\>    | Specifies the name of a group to closeout    |
| 1        | \<ordertype>     | The type of order to use.                    |
|          |                  | LMT                                          |  
|          |                  | LIT                                          |
|          |                  | LOC                                          |
|          |                  | LOO                                          |
|          |                  | MKT                                          |
|          |                  | MIT                                          |
|          |                  | MOC                                          |
|          |                  | MOO                                          |
|          |                  | MTL                                          |
|          |                  | STP                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAIL                                        |
|          |                  | TRAILLMT                                     |
| 2        | \<pricespec\>    | Limit price, if one is required; otherwise trigger price, if one is required. This applies to the following order types |
|          |                  | LMT                                          |  
|          |                  | LIT                                          |
|          |                  | LOC                                          |
|          |                  | LOO                                          |
|          |                  | MIT                                          |
|          |                  | STP                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAIL                                        |
|          |                  | TRAILLMT                                     |
| 3        | \<pricespec\>    | Trigger price, if one is required and was not specified in positional argument 2. This applies to the following order types |
|          |                  | LIT                                          |
|          |                  | STPLMT                                       |
|          |                  | TRAILLMT                                     |

Tagged arguments (both forms): None

<br/><br/>
### CONTRACT Command

Specifies a contract, which becomes the current contract in the current group.

This command has two forms: 

* full contract specification: this form explicitly specifies the 
  characteristics of the desired contract

* abbreviated contract specification: this form makes use of broker-dependent
  contract names, where necessary together with an exchange name, that 
  uniquely identify a contract

Syntax summary:

Form 1:

CONTRACT \<attribute\> [\<attribute\>...]

Form 2:

CONTRACT \<contractspec\>]

Positional aruments (form 1): None

Tagged arguments (form 1): 

| Tag           | Permitted&nbsp;values | Meaning                                 |
| ------------- | ---------------- | -------------------------------------------- |
| /curr[ency]   | \<currency\>     | The currency in which the contract is traded |
| /exch[ange]   | \<exchange\>     | The exchange at which the contract is traded |
| /exp[iry]     | yyyymm           | The expiry date for the contract: see [Contract Expiry](Contract-Expiry) | 
|               | yymmdd           |                                              |
|               | INTEGER[0..10]   |                                              |
| /local[symbol]| IDENTIFIER       | The broker's name for the contract           |
| /mult[iplier] | INTEGER[0..1000] | The factor used to convert prices into monetary values The factor used to convert prices into monetary values |
| /right        | C \| P           | Call or put                                  |
| /sec[type]    | CASH             | The security type                            |
|               | FOP              |                                              |
|               | FUT              |                                              |
|               | OPT              |                                              |
|               | STK              |                                              |
| /str[ike]     | DOUBLE           | Strike price                                 |
| /symb[ol]     | IDENTIFIER       | The underlying symbol for the contract       |

Positional arguments (form 2): 

| Position | Permitted&nbsp;values | Meaning                                 |
| -------- | ---------------- | -------------------------------------------- |
| 0        | \<contractspec\> | Specifies the contract by means of the broker's local symbol and, if need be, the exchange where it is traded |

Tagged arguments (form 2): None

<br/><br/>
### ENDBRACKET Command

Ends a bracket order specification. If there have been no errors during the 
bracket order specification, and bracket order batching is not in force, 
the bracket order is immediately processed for submission.

Syntax summary:

ENDBRACKET

Positional arguments: None

Tagged arguments: None

<br/><br/>
### ENDORDERS Command

Processes a batch of bracket orders for submission.

Syntax summary:

ENDORDERS

Positional arguments: None

Tagged arguments: None

<br/><br/>
### ENTRY Command

Specifies the entry order of a bracket order.

Syntax summary:

ENTRY \<ordertype\> [\<attribute\>...]

Positional arguments:

| Position | Permitted&nbsp;values | Meaning                                 |
| -------- | ---------------- | -------------------------------------------- |
| 0        | MKT              | The type of order to use.                    |
|          | LIT              |                                              |
|          | LOC              |                                              |
|          | LOO              |                                              |
|          | MKT              |                                              |
|          | MIT              |                                              |
|          | MOC              |                                              |
|          | MOO              |                                              |
|          | MTL              |                                              |
|          | STP              |                                              |
|          | STPLMT           |                                              |
|          | TRAIL            |                                              |
|          | TRAILLMT         |                                              |

Tagged arguments: 

| Tag           | Permitted&nbsp;values | Meaning                                 |
| ------------- | ---------------- | -------------------------------------------- |
| /ignorerth    | n/a                |                                            |
| /price        | \<pricespec\>    | The price specifier for the limit price      |
| /reason       |
| /tif          | DAY                | Time in force                              |
|               | GTC                |                                            |
|               | IOC                |                                            |
| /trigger[price] | \<pricespec\>    | The price specifier for the trigger price  |

<br/><br/>
### EXIT Command

Ends the session and terminates the program.

Syntax summary:

EXIT

Positional arguments: None

Tagged arguments: None

<br/><br/>
### GROUP Command

Defines a new group and makes it current, or makes an existing group 
the current group. Optionally a contract may be specified, which 
becomes the current contract for the group.

Syntax summary:

GROUP \<groupname\> [\<contractspec\>]

<br/><br/>
### HELP Command

Outputs the syntax summary.

<br/><br/>
### LIST Command

List current groups, positions or trades.

<br/><br/>
### PURGE

Removes all knowledge of a group, without affecting any orders that have been 
defined and submitted in that group.

<br/><br/>
### QUIT Command

Aborts the current bracket order definition.

<br/><br/>
### QUOTE Command

Displays bid, ask and last prices and sizes for the current or a specified 
contract.

<br/><br/>
### RESET Command

Cancels any bracket orders that have not yet been submitted, and any bracket 
order specifications that have not been completed.

<br/><br/>
### SELL Command

Specifies a single sell order.

<br/><br/>
### S Command

Repeats the previous SELL command.

<br/><br/>
### STAGEORDERS Command

Specifies that orders are to be sent to TWS but not transmitted to the broker 
for execution (manual intervention in TWS is required to actually transmit 
the orders for execution).

<br/><br/>
### STOPLOSS Command

Specifies the stop-loss order of a bracket order.

<br/><br/>
### TARGET Command

Specifies the target order of a bracket order.

<br/><br/>
## 9. Common Syntax Elements

### Contract Expiry

<br/><br/>
## 10. Detailed Syntax Specification



