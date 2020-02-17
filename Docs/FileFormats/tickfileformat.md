# TradeBuild Tickfile Format

A tickfile is a text file with a .tck file extension. Each line is a separate record. It contains the following records: 

* Header record
* Contract details record
* Tick records
* Comment records

Comment records may occur anywhere after the header record. Blank lines are ignored anywhere after the header record.

Note that all timestamps are in the timezone specified in the contract details record (unless otherwise indicated).

## Header record

This must be the first record. It contains comma separated fields. The fields are:

* declarer: indicates that this is a tickfile, it contains the string "tickfile".
* version: indicates the version of the tickfile format. The current version is 5.
* exchange: the name of the exchange for the contract for which this tickfile contains data.
symbol: the contract symbol.
* expiry date: the contract expiry date (for futures or options - blank for stocks).
* first tick time: the timestamp of the first tick record (see the 'timestamp' bullet under Tick Record below for details of the format).
* first tick time (readable): same as first tick time but formatted as a date & time string (in dd/mm/yy hh:mm:ss format)

Note that of these fields, only declarer, version and first tick time are actually used by the TradeBuild software. The other fields are documentary, and are included to make it easy to check at a glance what the tickfile contains.

Note that due to a bug, some tickfiles may contain the first tick time fields in the local timezone of the computer that recorded the file, rather than the timezone specified in the contract details record.

## Contract Details Record

This starts with the string "contractdetails=". The remainder of the record is an XML string containing the contract details. I won't define the XML syntax here, though if you're familiar with XML it should be pretty obvious by inspection. If anyone wants a detailed definition, contact me.
Note that the timezone specified in the contract details record is generally expected to be the official timezone of the relevant exchange. However this is not necessarily the case. What is important is that all tick records will be timestamped in the timezone specified in the contract details record.

## Tick Record

Contains comma separated fields. The fields are:
timestamp: the date and time at which the tick occurred. This is a VB datetime value expressed as a double (ie the integer part is the number of days since 30 Dec 1899, and the fractional part represents the fraction of the day since midnight).
readable time: the time part of the timestamp, formatted as hhmmss.ddd (where ddd is milliseconds).

* tick type: identifies what type of tick this record is. Possible values are:

  | Value | Meaning            |
  | ----- | ----------------   |
  | B     | Bid                |
  | A     | Ask                |
  | T     | Trade              |
  | H     | Session high       |
  | L     | Session low        |
  | C     | Previous session close |
  | V     | Volume             |
  | D     | Market depth       |
  | R     | Reset market depth (ie clear the market depth display) |

  The remaining fields are dependent on the tick type.

  For tick types B, A, T:

* tick price: the price of the bid, ask or trade
* tick size: the size of the bid, ask or trade

  For tick types H, L, C:

* tick price: the high or low price or the session close price

  For tick type V:

* volume: the current total volume for the session

  For tick type D:

* MDposition: the position in the relevant side of the market depth display
* MDMarketMaker: the market maker name (NASDAQ Level II data only)
* MDOperation: how to process this record: 0=insert, 1=update, 2=delete
* MDSide: which side of the book: 0=ask, 1=bid
* MDPrice: the price
* MDSize: the size
  For tick type R there are no further fields

## Comment Record

The comment record starts with //. It is entirely ignored by TradeBuild.
  