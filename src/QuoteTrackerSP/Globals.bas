Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ServiceProviderName As String = "QuoteTrackerSP"

Public Const TickfileFormatQuoteTracker As String = "urn:tradewright.com:names.tickfileformats.quotetrackerstreaming"

Public Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF
Public Const TIME_ZONE_ID_UNKNOWN  As Long = 0
Public Const TIME_ZONE_ID_STANDARD As Long = 1
Public Const TIME_ZONE_ID_DAYLIGHT As Long = 2

'================================================================================
' Enums
'================================================================================

Public Enum ErrorCodes
    LongJump = vbObjectError + 512
End Enum

Public Enum FieldTypes
    Symbol
    BidPrice = 2
    AskPrice = 3
    LastPrice = 4
    Change
    volume = 6
    openPrice
    PrevClose = 8
    BidSize = 9
    AskSize = 10
    LastSize = 11
    High = 12
    Low = 13
    Tick
    timestamp = 15
    HighPrice52
    LowPrice52
    CompanyName
    prevVolume
    NumberOfTrades
    AverageTradeSize
    AverageVolume
    Dividend
    Earnings
    Exchange
    MarketCap
    PE
    OpenInterest
    UPC
    Yield
End Enum

Public Enum SPErrorCodes
    CantConnectToQuoteTracker = 700
    LogonResponseCannotBeParsed
    FieldDescriptionCannotBeParsed
    ErrorCannotBeParsed
    PasswordInvalid
    UnexpectedError
    ErrorFromQuoteTracker
End Enum

Public Enum TickTypes
    None
    Bid
    Ask
    Last
    volume
    PrevClose
    High
    Low
    OpenInterest
End Enum

'================================================================================
' Types
'================================================================================

Public Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Public Type TIME_ZONE_INFORMATION
    Bias            As Long
    StandardName    As String * 64
    StandardDate    As SYSTEMTIME
    StandardBias    As Long
    DaylightName    As String * 64
    DaylightDate    As SYSTEMTIME
    DaylightBias    As Long
End Type

'================================================================================
' Declares
'================================================================================

Public Declare Sub GetSystemTime Lib "kernel32" ( _
                            lpSystemTime As SYSTEMTIME)

Public Declare Function GetTimeZoneInformation Lib "kernel32" ( _
                            TimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" ( _
                            lpTimeZone As TIME_ZONE_INFORMATION, _
                            lpUniversalTime As SYSTEMTIME, _
                            lpLocalTime As SYSTEMTIME) As Long

Public Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" ( _
                            lpTimeZone As TIME_ZONE_INFORMATION, _
                            lpLocalTime As SYSTEMTIME, _
                            lpUniversalTime As SYSTEMTIME) As Long

'================================================================================
' Variables
'================================================================================

Public gName As String

'================================================================================
' Procedures
'================================================================================

Public Function gCapabilities() As Long
gCapabilities = _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.PositionExact
End Function

Public Function gConvertLocalTimeToEST(ByVal timestamp As Date) As Date
Dim currTZ As TIME_ZONE_INFORMATION
Dim estTZ As TIME_ZONE_INFORMATION
Dim inLocalTime As SYSTEMTIME
Dim inSystime As SYSTEMTIME
Dim inESTTime As SYSTEMTIME

estTZ.Bias = 300
estTZ.DaylightBias = -60
estTZ.DaylightDate.wDayOfWeek = 0   ' Sunday
estTZ.DaylightDate.wDay = 1         ' first
estTZ.DaylightDate.wMonth = 4
estTZ.DaylightDate.wHour = 2
estTZ.StandardBias = 0
estTZ.StandardDate.wDayOfWeek = 0
estTZ.StandardDate.wDay = 5         ' last
estTZ.StandardDate.wMonth = 10
estTZ.StandardDate.wHour = 2

inLocalTime.wYear = Year(timestamp)
inLocalTime.wMonth = Month(timestamp)
inLocalTime.wDay = Day(timestamp)
inLocalTime.wHour = Hour(timestamp)
inLocalTime.wMinute = Minute(timestamp)
inLocalTime.wSecond = Second(timestamp)

GetTimeZoneInformation currTZ

TzSpecificLocalTimeToSystemTime currTZ, inLocalTime, inSystime

SystemTimeToTzSpecificLocalTime estTZ, inSystime, inESTTime

gConvertLocalTimeToEST = DateSerial(inESTTime.wYear, _
                                inESTTime.wMonth, _
                                inESTTime.wDay) + _
                        TimeSerial(inESTTime.wHour, _
                                inESTTime.wMinute, _
                                inESTTime.wSecond)
End Function

Public Function gSupports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
Select Case FormatIdentifier
Case TickfileFormatQuoteTracker, ""
    gSupports = (gCapabilities And Capabilities)
End Select

End Function

