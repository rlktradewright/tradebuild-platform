Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                As String = "QTSP26"

Public Const TickfileFormatQuoteTracker As String = "urn:tradewright.com:names.tickfileformats.quotetrackerstreaming"

Public Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF
Public Const TIME_ZONE_ID_UNKNOWN  As Long = 0
Public Const TIME_ZONE_ID_STANDARD As Long = 1
Public Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Public Const MaxLong    As Long = &H7FFFFFFF

Public Const OneMinute As Double = 1 / 1440

Public Const ParamNameServer As String = "Server"
Public Const ParamNamePort As String = "Port"
Public Const ParamNamePassword As String = "Password"
Public Const ParamNameConnectionRetryIntervalSecs As String = "Connection Retry Interval Secs"
Public Const ParamNameKeepConnection As String = "Keep Connection"
Public Const ParamNameProviderKey As String = "Provider Key"

'================================================================================
' Enums
'================================================================================

Public Enum ConnectionStates
    ConnNotConnected
    ConnConnecting
    ConnConnected
End Enum

Public Enum LocalErrorCodes
    LongJump = vbObjectError + 512
End Enum

Public Enum FieldTypes
    symbol = 1
    BidPrice = 2
    AskPrice = 3
    LastPrice = 4
    Change
    volume = 6
    openPrice
    prevClose = 8
    bidSize = 9
    askSize = 10
    lastSize = 11
    high = 12
    low = 13
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
    openInterest
    UPC
    Yield
    closePrice
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

Private Type QTAPITableEntry
    server          As String
    port            As Long
    ConnectionRetryIntervalSecs As Long
    providerKey     As String
    keepConnection  As Boolean  ' once this flag is set, the QTAPI instance
                                ' will only be disconnected by a call to
                                ' gReleaseQTAPIInstance with <forceDisconnect>
                                ' set to true or by a call to
                                ' gReleaseAllQTAPIInstances
    QTAPI           As QTAPI
    usageCount      As Long
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
' Global Variables
'================================================================================

'================================================================================
' Private Variables
'================================================================================

Private mQTAPITable() As QTAPITableEntry
Private mQTAPITableNextIndex As Long

Private mLogger As Logger

'================================================================================
' Properties
'================================================================================

Public Property Get gLogger() As Logger
If mLogger Is Nothing Then Set mLogger = GetLogger("log.serviceprovider.qtsp")
Set gLogger = mLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Function gCapabilities() As Long
gCapabilities = _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.PositionExact
End Function

Public Function gConvertESTToLocalTime(ByVal timestamp As Date) As Date
Dim currTZ As TIME_ZONE_INFORMATION
Dim estTZ As TIME_ZONE_INFORMATION
Dim localTime As SYSTEMTIME
Dim systime As SYSTEMTIME

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

localTime.wYear = Year(timestamp)
localTime.wMonth = Month(timestamp)
localTime.wDay = Day(timestamp)
localTime.wHour = Hour(timestamp)
localTime.wMinute = Minute(timestamp)
localTime.wSecond = Second(timestamp)

GetTimeZoneInformation currTZ

TzSpecificLocalTimeToSystemTime estTZ, localTime, systime

SystemTimeToTzSpecificLocalTime currTZ, systime, localTime

gConvertESTToLocalTime = DateSerial(localTime.wYear, _
                                localTime.wMonth, _
                                localTime.wDay) + _
                        TimeSerial(localTime.wHour, _
                                localTime.wMinute, _
                                localTime.wSecond)
End Function

Public Function gConvertLocalTimeToEST(ByVal timestamp As Date) As Date
Dim currTZ As TIME_ZONE_INFORMATION
Dim estTZ As TIME_ZONE_INFORMATION
Dim localTime As SYSTEMTIME
Dim systime As SYSTEMTIME

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

localTime.wYear = Year(timestamp)
localTime.wMonth = Month(timestamp)
localTime.wDay = Day(timestamp)
localTime.wHour = Hour(timestamp)
localTime.wMinute = Minute(timestamp)
localTime.wSecond = Second(timestamp)

GetTimeZoneInformation currTZ

TzSpecificLocalTimeToSystemTime currTZ, localTime, systime

SystemTimeToTzSpecificLocalTime estTZ, systime, localTime

gConvertLocalTimeToEST = DateSerial(localTime.wYear, _
                                localTime.wMonth, _
                                localTime.wDay) + _
                        TimeSerial(localTime.wHour, _
                                localTime.wMinute, _
                                localTime.wSecond)
End Function

Public Function gGetQTAPIInstance( _
                ByVal server As String, _
                ByVal port As Long, _
                ByVal password As String, _
                ByVal providerKey As String, _
                ByVal ConnectionRetryIntervalSecs As Long, _
                ByVal keepConnection As Boolean) As QTAPI
Dim i As Long

If mQTAPITableNextIndex = 0 Then
    ReDim mQTAPITable(5) As QTAPITableEntry
End If

For i = 0 To mQTAPITableNextIndex - 1
    If mQTAPITable(i).server = server And _
        mQTAPITable(i).port = port And _
        mQTAPITable(i).providerKey = providerKey And _
        mQTAPITable(i).ConnectionRetryIntervalSecs = ConnectionRetryIntervalSecs _
    Then
        Set gGetQTAPIInstance = mQTAPITable(i).QTAPI
        mQTAPITable(i).usageCount = mQTAPITable(i).usageCount + 1
        If keepConnection Then mQTAPITable(i).keepConnection = True
        Exit Function
    End If
Next

If mQTAPITableNextIndex > UBound(mQTAPITable) Then
    ReDim Preserve mQTAPITable(UBound(mQTAPITable) + 5) As QTAPITableEntry
End If

mQTAPITable(mQTAPITableNextIndex).server = server
mQTAPITable(mQTAPITableNextIndex).port = port
mQTAPITable(mQTAPITableNextIndex).providerKey = providerKey
mQTAPITable(mQTAPITableNextIndex).ConnectionRetryIntervalSecs = ConnectionRetryIntervalSecs
mQTAPITable(mQTAPITableNextIndex).usageCount = 1
mQTAPITable(mQTAPITableNextIndex).keepConnection = keepConnection
Set mQTAPITable(mQTAPITableNextIndex).QTAPI = New QTAPI
Set gGetQTAPIInstance = mQTAPITable(mQTAPITableNextIndex).QTAPI

mQTAPITableNextIndex = mQTAPITableNextIndex + 1

gGetQTAPIInstance.server = server
gGetQTAPIInstance.port = port
gGetQTAPIInstance.password = password
gGetQTAPIInstance.providerKey = providerKey
gGetQTAPIInstance.ConnectionRetryIntervalSecs = ConnectionRetryIntervalSecs
gGetQTAPIInstance.connect

End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub gReleaseAllQTAPIInstances()

Dim i As Long

For i = 0 To mQTAPITableNextIndex - 1
    mQTAPITable(i).usageCount = 0
    If Not mQTAPITable(i).QTAPI Is Nothing Then
        mQTAPITable(i).QTAPI.disconnect
        Set mQTAPITable(i).QTAPI = Nothing
    End If
    mQTAPITable(i).ConnectionRetryIntervalSecs = 0
    mQTAPITable(i).port = 0
    mQTAPITable(i).server = ""
    mQTAPITable(i).providerKey = ""
Next
                
End Sub

Public Sub gReleaseQTAPIInstance( _
                ByVal instance As QTAPI, _
                Optional ByVal forceDisconnect As Boolean)

Dim i As Long

For i = 0 To mQTAPITableNextIndex - 1
    If mQTAPITable(i).QTAPI Is instance Then
        mQTAPITable(i).usageCount = mQTAPITable(i).usageCount - 1
        If mQTAPITable(i).usageCount = 0 And _
            ((Not mQTAPITable(i).keepConnection) Or _
                forceDisconnect) _
        Then
            mQTAPITable(i).QTAPI.disconnect
            Set mQTAPITable(i).QTAPI = Nothing
            mQTAPITable(i).ConnectionRetryIntervalSecs = 0
            mQTAPITable(i).port = 0
            mQTAPITable(i).server = ""
            mQTAPITable(i).providerKey = ""
        End If
        Exit For
    End If
Next
                
End Sub
                
Public Function gSupports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
Select Case FormatIdentifier
Case TickfileFormatQuoteTracker, ""
    gSupports = (gCapabilities And capabilities)
End Select

End Function

