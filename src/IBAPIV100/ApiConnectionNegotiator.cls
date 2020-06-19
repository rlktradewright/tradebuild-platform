VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiConnectionNegotiator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ApiConnectionNegotiator"

' Client version history
'
'  6 = Added parentId to orderStatus
'  7 = The new execDetails event returned for an order filled status and reqExecDetails
'     Also market depth is available.
'  8 = Added lastFillPrice to orderStatus() event and permId to execution details
'  9 = Added 'averageCost', 'unrealizedPNL', and 'unrealizedPNL' to updatePortfolio event
' 10 = Added 'serverId' to the 'open order' & 'order status' events.
'      We send back all the API open orders upon connection.
'      Added new methods reqAllOpenOrders, reqAutoOpenOrders()
'      Added FA support - reqExecution has filter.
'                       - reqAccountUpdates takes acct code.
' 11 = Added permId to openOrder event.
' 12 = requsting open order attributes ignoreRth, hidden, and discretionary
' 13 = added goodAfterTime
' 14 = always send size on bid/ask/last tick
' 15 = send allocation description string on openOrder
' 16 = can receive account name in account and portfolio updates, and fa params in openOrder
' 17 = can receive liquidation field in exec reports, and notAutoAvailable field in mkt data
' 18 = can receive good till date field in open order messages, and request intraday backfill
' 19 = can receive rthOnly flag in ORDER_STATUS
' 20 = expects TWS time string on connection after server version >= 20.
' 21 = can receive bond contract details.
' 22 = can receive price magnifier in version 2 contract details message
' 23 = support for scanner
' 24 = can receive volatility order parameters in open order messages
' 25 = can receive HMDS query start and end times
' 26 = can receive option vols in option market data messages
' 27 = can receive delta neutral order type and delta neutral aux price in place order version 20: API 8.85
' 28 = can receive option model computation ticks: API 8.9
' 29 = can receive trail stop limit price in open order and can place them: API 8.91
' 30 = can receive extended bond contract def, new ticks, and trade count in bars
' 31 = can receive EFP extensions to scanner and market data, and combo legs on open orders
'    ; can receive RT bars
' 32 = can receive TickType.LAST_TIMESTAMP
'    ; can receive "whyHeld" in order status messages
' 33 = can receive ScaleNumComponents and ScaleComponentSize is open order messages
' 34 = can receive whatIf orders / order state
' 35 = can receive contId field for Contract objects
' 36 = can receive outsideRth field for Order objects
' 37 = can receive clearingAccount and clearingIntent for Order objects
' 38 = can receive multiplier and primaryExchange in portfolio updates
'    ; can receive cumQty and avgPrice in execution
'    ; can receive fundamental data
'    ; can receive underComp for Contract objects
'    ; can receive reqId and end marker in contractDetails/bondContractDetails
'    ; can receive ScaleInitComponentSize and ScaleSubsComponentSize for Order objects
' 39 = can receive underConId in contractDetails
' 40 = can receive algoStrategy/algoParams in openOrder
' 41 = can receive end marker for openOrder
'    ; can receive end marker for account download
'    ; can receive end marker for executions download
' 42 = can receive deltaNeutralValidation
' 43 = can receive longName(companyName)
'    ; can receive listingExchange
'    ; can receive RTVolume tick
' 44 = can receive end market for ticker snapshot
' 45 = can receive notHeld field in openOrder
' 46 = can receive contractMonth, industry, category, subcategory fields in contractDetails
'    ; can receive timeZoneId, tradingHours, liquidHours fields in contractDetails
' 47 = can receive gamma, vega, theta, undPrice fields in TICK_OPTION_COMPUTATION
' 48 = can receive exemptCode in openOrder
' 49 = can receive hedgeType and hedgeParam in openOrder
' 50 = can receive optOutSmartRouting field in openOrder
' 51 = can receive smartComboRoutingParams in openOrder
' 52 = can receive deltaNeutralConId, deltaNeutralSettlingFirm, deltaNeutralClearingAccount and deltaNeutralClearingIntent in openOrder
' 53 = can receive orderRef in execution
' 54 = can receive scale order fields (PriceAdjustValue, PriceAdjustInterval, ProfitOffset, AutoReset,
'      InitPosition, InitFillQty and RandomPercent) in openOrder
' 55 = can receive orderComboLegs (price) in openOrder
' 56 = can receive trailingPercent in openOrder
' 57 = can receive commissionReport message
' 58 = can receive CUSIP/ISIN/etc. in contractDescription/bondContractDescription
' 59 = can receive evRule, evMultiplier in contractDescription/bondContractDescription/executionDetails
'      can receive multiplier in executionDetails
' 60 = can receive deltaNeutralOpenClose, deltaNeutralShortSale, deltaNeutralShortSaleSlot and deltaNeutralDesignatedLocation in openOrder
' 61 = can receive multiplier in openOrder
'      can receive tradingClass in openOrder, updatePortfolio, execDetails and POSITION
' 62 = can receive avgCost in POSITION message

'@================================================================================
' Member variables
'@================================================================================

Private mReader                                     As BufferedReader
Private mWriter                                     As BufferedWriter

Private mServerVersion                              As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ServerVersion() As Long
Assert mServerVersion <> 0, "Server version not yet received"
ServerVersion = mServerVersion
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pReader As BufferedReader, _
                ByVal pWriter As BufferedWriter)
Set mReader = pReader
Set mWriter = pWriter
End Sub

Friend Sub NegotiateApiConnection()
Const ProcName As String = "NegotiateApiConnection"
On Error GoTo Err

gLog "Connecting to Tws: negotiating API connection", ModuleName, ProcName
mWriter.StartMessage "API"
mWriter.AddRawString "v" & ApiServerVersions.MinV100Plus & ".." & ApiServerVersions.Max, "SupportedServerVersions"
mWriter.Send forceLogMessage:=True

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Function ProcessInitialData( _
                ByVal pClientId As Long, _
                ByVal pOptionalCapabilities As String, _
                ByVal pTimeStamp As Date) As Long
Const ProcName As String = "ProcessInitialData"
On Error GoTo Err

gLog "Processing initial data for API connection", ModuleName, ProcName, , LogLevelSevere

gLog "Buffer contents: " & vbCrLf & gFormatBuffer(mReader.Buffer, mReader.BufferNextFreeIndex), ModuleName, ProcName, , LogLevelSevere

Dim lString1 As String
Dim lString2 As String

Dim lFieldNumber As Long
lFieldNumber = 1

mReader.BeginMessage

' With a new connection, the first data received is plain server version number,
' followed by the time at the server. However in some circumstances (eg after
' reconnection following a read timeout on the socket) TWS may send invalid information
' before the server version and server time. So we just read and discard until we
' identify the server version followed by a time that's close to the current time (say
' within 60 seconds - experiment has shown that 5 seconds is too short).

' Note that if there is incomplete data, then a call to GetString() will result in
' processing being aborted and then retried when more data arrives.

mReader.DiscardMessageLength

lString1 = mReader.GetString("Field 1")

Dim lServerTime As Date
Dim lTimezone As String

Const MaxDifferenceFromServerTime As Long = 60 * 60

Do
    lFieldNumber = lFieldNumber + 1
    lString2 = mReader.GetString("Field " & lFieldNumber)
    
    If IsInteger(lString1, ApiServerVersions.MinV100Plus, ApiServerVersions.Max) Then
        gLog "lString1 is potentially the server version: " & lString1, ModuleName, ProcName, , LogLevelSevere
        
        lServerTime = gGetDate(lString2, lTimezone)
        If lServerTime = 0# Then
            gLog "lString2 is not a date and time: " & lString2, ModuleName, ProcName, , LogLevelSevere
        ElseIf Abs(pTimeStamp - lServerTime) <= MaxDifferenceFromServerTime * OneSecond Then
            mServerVersion = CInt(lString1)
            Exit Do
        Else
            gLog "lString2 is not within " & MaxDifferenceFromServerTime & " seconds of receipt time (" & FormatTimestamp(pTimeStamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "): " & lString2, ModuleName, ProcName, , LogLevelSevere
        End If
    Else
        gLog "lString1 is not the server version which must be an integer <= " & ApiServerVersions.Max & ": " & lString1, ModuleName, ProcName, , LogLevelSevere
    End If
    lString1 = lString2
    
    Assert lFieldNumber <= 50, "Unable to recognise API connection completion sequence"
Loop

mReader.EndMessage

gLog "Server version", ModuleName, ProcName, CStr(mServerVersion), LogLevelSevere
gLog "TWS Time at connection", ModuleName, ProcName, FormatTimestamp(lServerTime, TimestampDateAndTimeISO8601 Or TimestampNoMillisecs) & " " & lTimezone, LogLevelSevere

StartAPI pClientId, pOptionalCapabilities
sendLogLevel

gLog "API connection negotiated - awaiting nextValidId", ModuleName, ProcName

ProcessInitialData = mServerVersion

Exit Function

Err:
If Err.Number = DataIncomplete Then
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Dim errSource As String: errSource = Err.source
    gLog "Expected initial data from TWS not found: awaiting more data", ModuleName, ProcName, , LogLevelSevere
    gHandleUnexpectedError Nothing, ProcName, ModuleName, pErrorNumber:=errNum, pErrorDesc:=errDesc, pErrorSource:=errSource
Else
    gHandleUnexpectedError Nothing, ProcName, ModuleName
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub sendLogLevel()
Const ProcName As String = "sendLogLevel"
On Error GoTo Err

Const VERSION = 1

mWriter.AddMessageId SET_SERVER_LOGLEVEL
mWriter.AddString VERSION, "Version"
mWriter.AddString TwsLogLevels.TwsLogLevelInformation, "Loglevel"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub StartAPI(ByVal pClientId As Long, pOptionalCapabilities As String)
Const VERSION As Integer = 2

mWriter.AddMessageId TwsSocketOutMsgTypes.StartAPI
mWriter.AddString VERSION, "Version"
mWriter.AddLong pClientId, "Client id"

mWriter.AddString pOptionalCapabilities, "Optional Capabilities"

mWriter.Send True
End Sub



