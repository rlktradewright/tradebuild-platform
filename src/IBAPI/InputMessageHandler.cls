VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputMessageHandler"
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

Event ApiConnectionEstablished(ByVal pServerVersion)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TwsHistDataStartInfo
    RequestId               As Long
    StartDate               As String
    EndDate                 As String
    BarCount                As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "InputMessageHandler"

Private Const ReqdServerVersion                     As Long = MIN_SERVER_VER_SCALE_ORDERS3

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mReader                          As BufferedReader
Attribute mReader.VB_VarHelpID = -1
Private mWriter                                     As BufferedWriter

Private mConnectionStatusConsumer                   As IConnectionStatusConsumer
Private mHistDataConsumer                           As IHistDataConsumer
Private mContractDetailsConsumer                    As IContractDetailsConsumer
Private mOrderInfoConsumer                          As IOrderInfoConsumer
Private mMarketDataConsumer                         As IMarketDataConsumer
Private mMarketDepthConsumer                        As IMarketDepthConsumer
Private mAccountDataConsumer                        As IAccountDataConsumer
Private mScannerDataConsumer                        As IScannerDataConsumer

Private mApiConnectionEstablished                   As Boolean

Private mStatsRecorder                              As PerformanceStatsRecorder

Private mServerVersion                              As Long

Private mNextOrderID                                As Long

Private mErrorAndNotificationConsumer               As IErrorAndNotificationConsumer
Private mProgramErrorHandler                        As IProgramErrorListener

Private mLogTwsMessages                             As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mNextOrderID = BaseOrderId
mLogTwsMessages = True
getElapsedTimer
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'================================================================================
' mReader Event Handlers
'================================================================================

Private Sub mReader_DataAdded()
Const ProcName As String = "mReader_DataAdded"
On Error GoTo Err

Static sRetryCount As Long

mReader.BeginCheckpointScope
    
If Not mApiConnectionEstablished Then negotiateConnection

Static sMessageID As TwsSocketInMsgTypes
Static sVersion As Long

checkmessages sMessageID, sVersion

sRetryCount = 0

mReader.EndCheckpointScope

Exit Sub

Err:
Dim errNum As Long
Dim errDesc As String
Dim errSource As String

errNum = Err.Number
errDesc = Err.Description
errSource = Err.source

If errNum = DataIncomplete Then
    sRetryCount = sRetryCount + 1
    Debug.Print "Data incomplete count: " & sRetryCount & "; CheckPointNumber: &H" & Hex$(mReader.CheckpointNumber)
    mReader.Rollback
    Exit Sub
End If

Dim s As String
s = "Error while processing input message: " & vbCrLf & _
    "Message ID=" & sMessageID & _
    " version=" & sVersion & _
    " MessageStartindex=" & mReader.MessageStartIndex & vbCrLf & _
    "InputBufferNextFreeIndex=" & mReader.BufferNextFreeIndex & _
    " InputParseIndex=" & mReader.InputParseIndex & vbCrLf & _
    "Buffer contents: " & vbCrLf & _
    gFormatBuffer(mReader.Buffer, mReader.BufferNextFreeIndex)

Debug.Print "Error " & errNum & ": " & errDesc & vbCrLf & s

gLog s, ModuleName, ProcName, pLogLevel:=LogLevelSevere
logSocketInputMessage True

gNotifyUnhandledError mProgramErrorHandler, pErrorNumber:=errNum, pErrorDesc:=errDesc, pErrorSource:=errSource, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let AccountDataConsumer(ByVal Value As IAccountDataConsumer)
Set mAccountDataConsumer = Value
End Property

Friend Property Get AccountDataConsumer() As IAccountDataConsumer
Set AccountDataConsumer = mAccountDataConsumer
End Property

Friend Property Get BufferedReader() As BufferedReader
Set BufferedReader = mReader
End Property

Friend Property Let ConnectionStatusConsumer(ByVal Value As IConnectionStatusConsumer)
Set mConnectionStatusConsumer = Value
End Property

Friend Property Get ConnectionStatusConsumer() As IConnectionStatusConsumer
Set ConnectionStatusConsumer = mConnectionStatusConsumer
End Property

Friend Property Let ContractDetailsConsumer(ByVal Value As IContractDetailsConsumer)
Set mContractDetailsConsumer = Value
End Property

Friend Property Get ContractDetailsConsumer() As IContractDetailsConsumer
Set ContractDetailsConsumer = mContractDetailsConsumer
End Property

Friend Property Let HistDataConsumer(ByVal Value As IHistDataConsumer)
Set mHistDataConsumer = Value
End Property

Friend Property Get HistDataConsumer() As IHistDataConsumer
Set HistDataConsumer = mHistDataConsumer
End Property

Friend Property Let MarketDataConsumer(ByVal Value As IMarketDataConsumer)
Set mMarketDataConsumer = Value
End Property

Friend Property Get MarketDataConsumer() As IMarketDataConsumer
Set MarketDataConsumer = mMarketDataConsumer
End Property

Friend Property Let MarketDepthConsumer(ByVal Value As IMarketDepthConsumer)
Set mMarketDepthConsumer = Value
End Property

Friend Property Get MarketDepthConsumer() As IMarketDepthConsumer
Set MarketDepthConsumer = mMarketDepthConsumer
End Property

Friend Property Let NextOrderId(ByVal Value As Long)
mNextOrderID = Value
End Property

Friend Property Get NextOrderId() As Long
NextOrderId = mNextOrderID
mNextOrderID = mNextOrderID + 1
End Property

Friend Property Let OrderInfoConsumer(ByVal Value As IOrderInfoConsumer)
Set mOrderInfoConsumer = Value
End Property

Friend Property Get OrderInfoConsumer() As IOrderInfoConsumer
Set OrderInfoConsumer = mOrderInfoConsumer
End Property

Friend Property Let ScannerDataConsumer(ByVal Value As IScannerDataConsumer)
Set mScannerDataConsumer = Value
End Property

Friend Property Get ScannerDataConsumer() As IScannerDataConsumer
Set ScannerDataConsumer = mScannerDataConsumer
End Property

'Friend Property Get SerialiseContractDetailsRequests() As Boolean
'SerialiseContractDetailsRequests = (mServerVersion < MIN_SERVER_VER_CONTRACT_DATA_CHAIN)
'End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pReader As BufferedReader, _
                ByVal pWriter As BufferedWriter, _
                ByVal pStatsRecorder As PerformanceStatsRecorder, _
                ByVal pErrorAndNotificationConsumer As IErrorAndNotificationConsumer, _
                ByVal pProgramErrorHandler As IProgramErrorListener)
Set mReader = pReader
Set mWriter = pWriter
Set mStatsRecorder = pStatsRecorder
mApiConnectionEstablished = False
Set mErrorAndNotificationConsumer = pErrorAndNotificationConsumer
Set mProgramErrorHandler = pProgramErrorHandler
End Sub

Friend Sub ProcessBuffer(ByVal pLog As Boolean)
Const ProcName As String = "ProcessBuffer"
On Error GoTo Err

mLogTwsMessages = pLog

Static sMessageID As TwsSocketInMsgTypes
Static sVersion As Long

checkmessages sMessageID, sVersion

mLogTwsMessages = True

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Reset()
mApiConnectionEstablished = False
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function adjustPrice( _
                ByVal Price As Double, _
                ByVal pCurrencyCode As String) As Double
Const ProcName As String = "adjustPrice"
On Error GoTo Err

If pCurrencyCode = "GBP" And _
    mServerVersion < 23 _
Then
    adjustPrice = Price * 100
Else
    adjustPrice = Price
End If

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub checkmessages(ByRef pMessageId As Long, ByRef pVersion As Long)
Const ProcName As String = "checkmessages"
On Error GoTo Err

Do While Not mReader.AtEnd
    mReader.BeginCheckpointScope
        mReader.BeginMessage mLogTwsMessages
        getPerformanceTimer.StartTiming
        
        getMessageIDandVersion mReader, pMessageId, pVersion
        getMessage pMessageId, pVersion
        
        updateStats pMessageId, getPerformanceTimer.ElapsedTimeMicroseconds
        mReader.EndMessage
    mReader.EndCheckpointScope
Loop

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountDownloadEnd(ByVal pVersion As Long)
Const ProcName As String = "getAccountDownloadEnd"
On Error GoTo Err

Dim accountName As String
accountName = mReader.GetString("Account name")

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.EndAcccountValue accountName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountSummary(ByVal pVersion As Long)
Const ProcName As String = "getAccountSummary"
On Error GoTo Err

Dim lRequestId As Long: lRequestId = mReader.GetLong("Request Id")
Dim lAccount As String: lAccount = mReader.GetString("Account")
Dim lTag As String: lTag = mReader.GetString("Tag")
Dim lValue As String: lValue = mReader.GetString("Value")
Dim lCurrency As String: lCurrency = mReader.GetString("Currency")

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyAccountSummary lRequestId, lAccount, lTag, lValue, lCurrency

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountSummaryEnd(ByVal pVersion As Long)
Const ProcName As String = "getAccountSummaryEnd"
On Error GoTo Err

Dim lRequestId As Long: lRequestId = mReader.GetLong("RequestId")

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.EndAccountSummary lRequestId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountTime(ByVal pVersion As Long)
Const ProcName As String = "getAccountTime"
On Error GoTo Err

Dim accountTime As String: accountTime = mReader.GetString("Account Time")

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyAccountTime accountTime

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountValue(ByVal pVersion As Long)
Const ProcName As String = "getAccountValue"
On Error GoTo Err

Dim key As String: key = mReader.GetString("Key")
Dim val As String: val = mReader.GetString("Value")
Dim cur As String: cur = mReader.GetString("Currency")
Dim accountName As String: If pVersion >= 2 Then accountName = mReader.GetString("Account Name")

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyAccountValue key, val, cur, accountName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getBondContractData(ByVal pVersion As Long)
' not currently supported
Const ProcName As String = "getBondContractData"
On Error GoTo Err

Dim lRequestId As Long
If (pVersion >= 3) Then lRequestId = mReader.GetLong("Req Id")
mReader.GetString "Symbol"
mReader.GetString "Sectype"
mReader.GetString "Cusip"
mReader.GetDouble "Coupon"
mReader.GetString "Maturity"
mReader.GetString "IssueDate"
mReader.GetString "Ratings"
mReader.GetString "BondType"
mReader.GetString "CouponType"
mReader.GetBoolean "Convertible"
mReader.GetBoolean "Callable"
mReader.GetBoolean "Putable"
mReader.GetString "DescAppend"
mReader.GetString "Exchange"
mReader.GetString "Currency"
mReader.GetString "MarketName"
mReader.GetString "TradingClass"
mReader.GetLong "conid"
mReader.GetDouble "MinTick"
mReader.GetString "OrderTypes"
mReader.GetString "ValidExchanges"
If (pVersion >= 2) Then
    mReader.GetString "Next Option Date"
    mReader.GetString "Next Option Type"
    mReader.GetBoolean "Next Option Partial"
    mReader.GetString "Notes"
End If
If pVersion >= 4 Then mReader.GetString "LongName"
If pVersion >= 6 Then
    mReader.GetString "EvRule"
    mReader.GetDouble "EvMultiplier"
End If
If pVersion >= 5 Then
    Dim lSecIdListCount As Long: lSecIdListCount = mReader.GetLong("SecIdListCount")
    If lSecIdListCount > 0 Then
        ReDim lSecIdList(lSecIdListCount - 1) As TwsTagValue
        Dim i As Long
        For i = 0 To lSecIdListCount - 1
            lSecIdList(i).Tag = mReader.GetString("Tag")
            lSecIdList(i).Value = mReader.GetString("Value")
        Next
    End If
End If

logSocketInputMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.NotifyError lRequestId, ErrorCodes.ErrUnsupportedOperationException, "Bonds not yet supported"

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getCallersRequestIdFromTwsContractRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsContractRequestId = pId - BaseContractRequestId
End Function

Private Function getCallersRequestIdFromTwsExecutionsRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsExecutionsRequestId = pId - BaseExecutionsRequestId
End Function

Private Function getCallersRequestIdFromTwsHistRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsHistRequestId = pId - BaseHistoricalDataRequestId
End Function

Private Function getCallersRequestIdFromTwsMarketDataRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsMarketDataRequestId = pId - BaseMarketDataRequestId
End Function

Private Function getCallersRequestIdFromTwsMarketDepthRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsMarketDepthRequestId = pId - BaseMarketDepthRequestId
End Function

Private Sub getCommissionReport(ByVal pVersion As Long)
Const ProcName As String = "getCommissionReport"
On Error GoTo Err

Dim lCommissionReport As New TwsCommissionReport
lCommissionReport.ExecId = mReader.GetString("ExecId")
lCommissionReport.Commission = mReader.GetDouble("Commission")
lCommissionReport.CurrencyCode = mReader.GetString("CurrencyCode")
lCommissionReport.RealizedPNL = mReader.GetDouble("RealizedPNL")
lCommissionReport.Yield = mReader.GetDouble("Yield")
lCommissionReport.YieldRedemptionDate = mReader.GetLong("YieldRedemptionDate")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getContractData(ByVal pVersion As Long)
Const ProcName As String = "getContractData"
On Error GoTo Err

Dim lReqId As Long
Dim lContract As New TwsContract
Dim lContractDetails As New TwsContractDetails

lContractDetails.Summary = lContract

If (pVersion >= 3) Then lReqId = getCallersRequestIdFromTwsContractRequestId(mReader.GetLong("Req Id"))

lContract.Symbol = mReader.GetString("Symbol")
lContract.Sectype = gTwsSecTypeFromString(mReader.GetString("Sec type"))
lContract.Expiry = mReader.GetString("Expiry")
lContract.Strike = mReader.GetDouble("Strike")
lContract.OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
lContract.Exchange = mReader.GetString("Exchange")
lContract.CurrencyCode = mReader.GetString("Currency")
lContract.LocalSymbol = mReader.GetString("Local Symbol")
lContractDetails.MarketName = mReader.GetString("Market Name")
lContract.TradingClass = mReader.GetString("Trading class")
lContract.ConId = mReader.GetLong("Contract id")

lContractDetails.MinTick = mReader.GetDouble("Minimum tick")
' note that IB gives very small tick sizes for some UK stocks, eg 0.000001 for VOD and TSCO
' (possibly all FTSE 100 stocks). This causes us tick encoding problems, so we up it to
' something that reflects reality!
If lContract.CurrencyCode = "GBP" And lContract.Sectype = TwsSecTypeStock And lContractDetails.MinTick < 0.0001 Then lContractDetails.MinTick = 0.0001

lContract.Multiplier = mReader.GetLong("Multiplier")
If lContract.Multiplier = 0 Then lContract.Multiplier = 1


lContractDetails.OrderTypes = mReader.GetString("Order types")
lContractDetails.ValidExchanges = mReader.GetString("Valid exchanges")
    
If pVersion >= 2 Then
    lContractDetails.PriceMagnifier = mReader.GetLong("Price magnifier")
Else
    If lContract.CurrencyCode = "GBP" Then
        lContractDetails.PriceMagnifier = 100
    Else
        lContractDetails.PriceMagnifier = 1
    End If
End If

lContractDetails.MinTick = lContractDetails.MinTick * lContractDetails.PriceMagnifier
    
If pVersion >= 4 Then lContractDetails.UnderConId = mReader.GetLong("UnderConId")

If pVersion >= 5 Then
   lContractDetails.LongName = mReader.GetString("Long name")
   lContract.PrimaryExch = mReader.GetString("Primary Exchange")
End If

If pVersion >= 6 Then
    lContractDetails.ContractMonth = mReader.GetString("ContractMonth")
    lContractDetails.Industry = mReader.GetString("Industry")
    lContractDetails.Category = mReader.GetString("Category")
    lContractDetails.Subcategory = mReader.GetString("Subcategory")
    lContractDetails.TimeZoneId = mReader.GetString("TimeZoneId")
    lContractDetails.TradingHours = mReader.GetString("TradingHours")
    lContractDetails.LiquidHours = mReader.GetString("LiquidHours")
 End If

If pVersion >= 8 Then
    lContractDetails.EvRule = mReader.GetString("EvRule")
    lContractDetails.EvMultiplier = mReader.GetDouble("EvMultiplier")
End If

If pVersion >= 7 Then
    Dim lSecIdListCount As Long: lSecIdListCount = mReader.GetLong("SecIdListCount")
    If lSecIdListCount > 0 Then
        ReDim lSecIdList(lSecIdListCount - 1) As TwsTagValue
        Dim i As Long
        For i = 0 To lSecIdListCount - 1
            lSecIdList(i).Tag = mReader.GetString("Tag")
            lSecIdList(i).Value = mReader.GetString("Value")
        Next
        lContractDetails.SecIdList = lSecIdList
    End If
End If

logSocketInputMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.NotifyContract lReqId, lContractDetails

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getContractEnd(ByVal pVersion As Long)
Const ProcName As String = "getContractEnd"
On Error GoTo Err

Dim reqId As Long
reqId = getCallersRequestIdFromTwsContractRequestId(mReader.GetLong("Req Id"))

logSocketInputMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.EndContractDetails reqId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getCurrentTime(ByVal pVersion As Long)
Const ProcName As String = "getCurrentTime"
On Error GoTo Err

Dim lSystemTime As Double: lSystemTime = mReader.GetDouble("SystemTime")
Dim lTime As Date: lTime = gGetDateFromUnixSystemTime(lSystemTime)

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getDeltaNeutralValidation(ByVal pVersion As Long)
Const ProcName As String = "getDeltaNeutralValidation"
On Error GoTo Err

Dim reqId As Long: reqId = mReader.GetLong("ReqID")

Dim ConId As Long: ConId = mReader.GetLong("ConID")
Dim Delta As Double: Delta = mReader.GetDouble("Delta")
Dim Price As Double: Price = mReader.GetDouble("Price")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getElapsedTimer() As ElapsedTimer
Static sElapsedTimer As ElapsedTimer
Const ProcName As String = "getElapsedTimer"
On Error GoTo Err

If sElapsedTimer Is Nothing Then
    Set sElapsedTimer = New ElapsedTimer
    sElapsedTimer.StartTiming
End If
Set getElapsedTimer = sElapsedTimer

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub getErrorMsg(ByVal pVersion As Long)
Const ProcName As String = "getErrorMsg"
On Error GoTo Err

Static sLostConnectionToIb As Boolean
Dim id As Long
Dim lErrorCode As Long
Dim lErrorMsg As String
Dim execFilter As TwsExecutionFilter
Dim index As Long


id = mReader.GetLong("Id")
Debug.Print "Error id: " & id

lErrorCode = mReader.GetLong("Error code")
Debug.Print "Error code: " & lErrorCode

lErrorMsg = mReader.GetString("Error msg")
Debug.Print "Error msg: " & lErrorMsg

logSocketInputMessage True

Select Case getIdType(id)
Case IdTypeContractData
    processContractError getCallersRequestIdFromTwsContractRequestId(id), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeRealtimeData
    If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyError getCallersRequestIdFromTwsMarketDataRequestId(id), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeMarketDepth
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.NotifyError getCallersRequestIdFromTwsMarketDepthRequestId(id), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeHistoricalData
    If lErrorCode = 165 Then
        gLog "Connected to IB Historical Market Data Service - " & lErrorMsg, ModuleName, ProcName
    Else
        If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.NotifyError getCallersRequestIdFromTwsHistRequestId(id), lErrorCode, lErrorMsg
    End If
    Exit Sub
Case IdTypeOrder
    If sLostConnectionToIb And lErrorCode = 200 Then Exit Sub
    If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyError id, lErrorCode, lErrorMsg
    Exit Sub
End Select

Select Case lErrorCode
Case 316
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.ResetMarketDepth pReEstablish:=True
Case 317
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.ResetMarketDepth pReEstablish:=False
Case 321    ' invalid Request
    gLog "Error (" & lErrorCode & "; id=" & id & ") from Tws: " & lErrorMsg, ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiError lErrorCode, lErrorMsg
Case 326    ' ClientID already in use

Case 1100, 2110
    ' connectivity between Tws and IB has been lost
    
    sLostConnectionToIb = True
    If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyIBServerConnectionClosed
    gLog "Connection to IB has been lost", ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
    
Case 1101
    ' connectivity between Tws and IB has been restored, but data has been
    ' lost, so need to reestablish market data and market depth requests

    sLostConnectionToIb = False
    
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.ResetMarketDepth True
    
    If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyIBServerConnectionRecovered True
    
    gLog "Connection to IB recovered: market data re-established", ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
    
Case 1102
    ' connectivity between Tws and IB has been restored without loss of data
    ' Now need to reassociate Order ids with Tws

    sLostConnectionToIb = False
    
    If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyIBServerConnectionRecovered False
    
    gLog "Connection to IB recovered: no loss of data", ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
    
Case 2103, 2104, 2105, 2106, 2107, 2108
    gLog lErrorMsg, ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
Case Else
    gLog "Error (" & lErrorCode & "; id=" & id & ") from Tws: " & lErrorMsg, ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiError lErrorCode, lErrorMsg
End Select

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getExecutionData(ByVal pVersion As Long)
Const ProcName As String = "getExecutionData"
On Error GoTo Err

Dim lReqId As Long: lReqId = -1
If (pVersion >= 7) Then lReqId = getCallersRequestIdFromTwsExecutionsRequestId(mReader.GetLong("lReqId"))

Dim lExecution As New TwsExecution
lExecution.OrderId = mReader.GetLong("OrderId")

Dim lContract As New TwsContract
With lContract
    If (pVersion >= 5) Then .ConId = mReader.GetLong("Con Id")
    .Symbol = mReader.GetString("Symbol")
    .Sectype = gTwsSecTypeFromString(mReader.GetString("Sec type"))
    .Expiry = mReader.GetString("Expiry")
    .Strike = mReader.GetDouble("Strike")
    .OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
    If pVersion >= 9 Then .Multiplier = mReader.GetLong("Multiplier")
    If .Multiplier = 0 Then .Multiplier = 1
    .Exchange = mReader.GetString("Exchange")
    .CurrencyCode = mReader.GetString("Currency")
    .LocalSymbol = mReader.GetString("Local Symbol")
    If pVersion >= 10 Then .TradingClass = mReader.GetString("TradingClass")
End With

With lExecution
    .ExecId = mReader.GetString("Exec id")
    .Time = mReader.GetString("Exec Time")
    .AcctNumber = mReader.GetString("Account id")
    .Exchange = mReader.GetString("Exchange")
    .Side = mReader.GetString("Side")
    .Shares = mReader.GetLong("Shares")
    .Price = mReader.GetDouble("Price")
    .Price = adjustPrice(.Price, lContract.CurrencyCode)
        
    If pVersion >= 2 Then .PermId = mReader.GetLong("Perm Id")
    
    If pVersion >= 3 Then .ClientID = mReader.GetLong("Client id")
       
    If pVersion >= 4 Then .Liquidation = mReader.GetBoolean("Liquidation")
    
    If (pVersion >= 6) Then
        .CumQty = mReader.GetLong("Cum Qty")
        .AvgPrice = mReader.GetDouble("Avg Price")
    End If
    
    If pVersion >= 8 Then .OrderRef = mReader.GetString("OrderRef")
    
    If pVersion >= 9 Then
        .EvRule = mReader.GetString("EvRule")
        .EvMultiplier = mReader.GetDouble("EvMultiplier")
    End If
End With

logSocketInputMessage True

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyExecution lReqId, lContract, lExecution

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getExecutionDataEnd(ByVal pVersion As Long)
Const ProcName As String = "getExecutionDataEnd"
On Error GoTo Err

Dim reqId  As Long

reqId = getCallersRequestIdFromTwsExecutionsRequestId(mReader.GetLong("reqId"))
logSocketInputMessage

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.EndExecutions reqId
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getFAData(ByVal pVersion As Long)
Const ProcName As String = "getFAData"
On Error GoTo Err

Dim DataType As TwsFADataTypes: DataType = mReader.GetLong("Data type")
Dim XMLData As String: XMLData = mReader.GetString("XML data")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getFundamentalData(ByVal pVersion As Long)
Const ProcName As String = "getFundamentalData"
On Error GoTo Err

Dim lReqId As Long: lReqId = mReader.GetLong("ReqID")
Dim lData As String: lData = mReader.GetString("Data")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getHistoricalData(ByVal pVersion As Long)
Const ProcName As String = "getHistoricalData"
On Error GoTo Err

Static sStartInfo As TwsHistDataStartInfo
Static sBar As TwsBar
Static i As Long

Dim lBar As TwsBar
Dim lReqId As Long

mReader.BeginCheckpointScope
    If mReader.GetLongCheckpointed(lReqId, "Request id") Then
        i = 0
        sBar = lBar
        sStartInfo.RequestId = getCallersRequestIdFromTwsHistRequestId(lReqId)
    End If
    
    If pVersion >= 2 Then mReader.GetStringCheckpointed sStartInfo.StartDate, "Start date"
    
    If pVersion >= 2 Then mReader.GetStringCheckpointed sStartInfo.EndDate, "End date"
    
    If mReader.GetLongCheckpointed(sStartInfo.BarCount, "Item count") Then
        Debug.Print "Bars retrieved: " & sStartInfo.BarCount
        
        If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.StartHistData sStartInfo.RequestId, sStartInfo.StartDate, sStartInfo.EndDate, sStartInfo.BarCount
    End If
    
    For i = i To sStartInfo.BarCount - 1
        mReader.BeginCheckpointScope
            mReader.GetStringCheckpointed sBar.timeStamp, "Bar date"
            mReader.GetDoubleCheckpointed sBar.OpenValue, "Open"
            mReader.GetDoubleCheckpointed sBar.HighValue, "High"
            mReader.GetDoubleCheckpointed sBar.LowValue, "Low"
            mReader.GetDoubleCheckpointed sBar.CloseValue, "Close"
            mReader.GetLongCheckpointed sBar.Volume, "Volume"
            mReader.GetDoubleCheckpointed sBar.WAP, "WAP"
            mReader.GetBooleanCheckpointed sBar.HasGaps, "Has gaps"
            If pVersion >= 3 Then mReader.GetLongCheckpointed sBar.TickVolume, "Tick volume"
            
            If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.NotifyBar sStartInfo.RequestId, sBar
            
        mReader.EndCheckpointScope
    Next
mReader.EndCheckpointScope

logSocketInputMessage

If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.EndHistData sStartInfo.RequestId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getIdType( _
                ByVal id As Long) As IdTypes
Const ProcName As String = "getIdType"
On Error GoTo Err

If id >= BaseOrderId Then
    getIdType = IdTypeOrder
ElseIf id >= BaseContractRequestId Then
    getIdType = IdTypeContractData
ElseIf id >= BaseExecutionsRequestId Then
    getIdType = IdTypeExecution
ElseIf id >= BaseHistoricalDataRequestId Then
    getIdType = IdTypeHistoricalData
ElseIf id >= BaseMarketDepthRequestId Then
    getIdType = IdTypeMarketDepth
ElseIf id >= 0 Then
    getIdType = IdTypeRealtimeData
Else
    getIdType = IdTypeNone
End If

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub getManagedAccounts(ByVal pVersion As Long)
Const ProcName As String = "getManagedAccounts"
On Error GoTo Err

Dim accountsList As String: accountsList = mReader.GetString("Accounts list")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMarketDataType(ByVal pVersion As Long)
Const ProcName As String = "getMarketDataType"
On Error GoTo Err

Dim lReqId As Long: lReqId = mReader.GetLong("ReqID")
Dim lMarketDataType As Long: lMarketDataType = mReader.GetLong("MarketDataType")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMarketDepth(ByVal pVersion As Long)
Const ProcName As String = "getMarketDepth"
On Error GoTo Err

Dim id As Long: id = mReader.GetLong("Id")
Dim lPosition As Long: lPosition = mReader.GetLong("POSITION")
Dim lOperation As TwsDOMOperations:  lOperation = mReader.GetLong("Operation")
Dim lSide As TwsDOMSides: lSide = mReader.GetLong("Side")
Dim lPrice As Double: lPrice = mReader.GetDouble("Price")
Dim lSize As Long: lSize = mReader.GetLong("Size")

logSocketInputMessage

If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.NotifyMarketDepth getCallersRequestIdFromTwsMarketDepthRequestId(id), lPosition, "", lOperation, lSide, lPrice, lSize

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMarketDepthL2(ByVal pVersion As Long)
Const ProcName As String = "getMarketDepthL2"
On Error GoTo Err

Dim id As Long: id = mReader.GetLong("Id")
Dim lPosition As Long: lPosition = mReader.GetLong("POSITION")
Dim lMarketMaker As String: lMarketMaker = mReader.GetString("Marketmaker")
Dim lOperation As TwsDOMOperations:  lOperation = mReader.GetLong("Operation")
Dim lSide As TwsDOMSides: lSide = mReader.GetLong("Side")
Dim lPrice As Double: lPrice = mReader.GetDouble("Price")
Dim lSize As Long: lSize = mReader.GetLong("Size")

logSocketInputMessage

If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.NotifyMarketDepth getCallersRequestIdFromTwsMarketDepthRequestId(id), lPosition, lMarketMaker, lOperation, lSide, lPrice, lSize

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMessage(ByVal pMessageId As TwsSocketInMsgTypes, ByVal pVersion As Long)
Const ProcName As String = "getMessage"
On Error GoTo Err

Select Case pMessageId
Case TwsSocketInMsgTypes.TICK_PRICE
    getTickPrice pVersion
Case TwsSocketInMsgTypes.TICK_SIZE
    getTickSize pVersion
Case TwsSocketInMsgTypes.ORDER_STATUS
    getOrderStatus pVersion
Case TwsSocketInMsgTypes.ERR_MSG
    getErrorMsg pVersion
Case TwsSocketInMsgTypes.OPEN_ORDER
    getOpenOrder pVersion
Case TwsSocketInMsgTypes.ACCT_VALUE
    getAccountValue pVersion
Case TwsSocketInMsgTypes.PORTFOLIO_VALUE
    getPortfolioValue pVersion
Case TwsSocketInMsgTypes.ACCT_UPDATE_TIME
    getAccountTime pVersion
Case TwsSocketInMsgTypes.NEXT_VALID_ID
    getNextValidId pVersion
Case TwsSocketInMsgTypes.CONTRACT_DATA
    getContractData pVersion
Case TwsSocketInMsgTypes.EXECUTION_DATA
    getExecutionData pVersion
Case TwsSocketInMsgTypes.MARKET_DEPTH
    getMarketDepth pVersion
Case TwsSocketInMsgTypes.MARKET_DEPTH_L2
    getMarketDepthL2 pVersion
Case TwsSocketInMsgTypes.NEWS_BULLETINS
    getNewsBulletins pVersion
Case TwsSocketInMsgTypes.MANAGED_ACCTS
    getManagedAccounts pVersion
Case TwsSocketInMsgTypes.RECEIVE_FA
    getFAData pVersion
Case TwsSocketInMsgTypes.HISTORICAL_DATA
    getHistoricalData pVersion
Case TwsSocketInMsgTypes.BOND_CONTRACT_DATA
    getBondContractData pVersion
Case TwsSocketInMsgTypes.SCANNER_PARAMETERS
    getScannerParameters pVersion
Case TwsSocketInMsgTypes.SCANNER_DATA
    getScannerData pVersion
Case TwsSocketInMsgTypes.TICK_OPTION_COMPUTATION
    getTickOptionComputation pVersion
Case TwsSocketInMsgTypes.TICK_GENERIC
    getTickGeneric pVersion
Case TwsSocketInMsgTypes.TICK_STRING
    getTickString pVersion
Case TwsSocketInMsgTypes.CONTRACT_DATA_END
    getContractEnd pVersion
Case TwsSocketInMsgTypes.OPEN_ORDER_END
    getOpenOrderEnd pVersion
Case TwsSocketInMsgTypes.ACCT_DOWNLOAD_END
    getAccountDownloadEnd pVersion
Case TwsSocketInMsgTypes.EXECUTION_DATA_END
    getExecutionDataEnd pVersion
Case TwsSocketInMsgTypes.DELTA_NEUTRAL_VALIDATION
    getDeltaNeutralValidation pVersion
Case TwsSocketInMsgTypes.TICK_SNAPSHOT_END
    getTickSnapshotEnd pVersion
Case TwsSocketInMsgTypes.POSITION
    getPosition pVersion
Case TwsSocketInMsgTypes.POSITION_END
    getPositionEnd pVersion
Case TwsSocketInMsgTypes.ACCOUNT_SUMMARY
    getAccountSummary pVersion
Case TwsSocketInMsgTypes.ACCOUNT_SUMMARY_END
    getAccountSummaryEnd pVersion
Case TwsSocketInMsgTypes.TICK_EFP
    getTickEFP pVersion
Case TwsSocketInMsgTypes.CURRENT_TIME
    getCurrentTime pVersion
Case TwsSocketInMsgTypes.REAL_TIME_BARS
    getRealtimeBars pVersion
Case TwsSocketInMsgTypes.FUNDAMENTAL_DATA
    getFundamentalData pVersion
Case TwsSocketInMsgTypes.MARKET_DATA_TYPE
    getMarketDataType pVersion
Case TwsSocketInMsgTypes.COMMISSION_REPORT
    getCommissionReport pVersion
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid message code from Tws"
End Select

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMessageIDandVersion(ByVal pReader As BufferedReader, ByRef pMessageId As Long, ByRef pVersion As Long)
Const ProcName As String = "getMessageIDandVersion"
On Error GoTo Err

If pReader.GetLongCheckpointed(pMessageId, "Msg id") Then
    pReader.AppendToMessage "("
    pReader.AppendToMessage gInputMessageIdToString(pMessageId)
    pReader.AppendToMessage ") "
End If
    
pReader.GetLongCheckpointed pVersion, "Version"

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getNewsBulletins(ByVal pVersion As Long)
Const ProcName As String = "getNewsBulletins"
On Error GoTo Err

Dim lMsgId As Long: lMsgId = mReader.GetLong("Id")
Dim msgType As Long: msgType = mReader.GetLong("Type")
Dim newsMessage As String: newsMessage = mReader.GetString("Message")
Dim originatingExch As String: originatingExch = mReader.GetString("Orig Exchange")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getNextValidId(ByVal pVersion As Long)
Const ProcName As String = "getNextValidId"
On Error GoTo Err

mNextOrderID = mReader.GetLong("Next id")

logSocketInputMessage

If mNextOrderID < BaseOrderId Then mNextOrderID = BaseOrderId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOpenOrder(ByVal pVersion As Long)
Const ProcName As String = "getOpenOrder"
On Error GoTo Err

Dim i As Long

Dim lContract As New TwsContract
Dim lOrder As New TwsOrder
Dim lOrderState As New TwsOrderState


lOrder.OrderId = mReader.GetLong("Id")

' read contract fields
With lContract
    If pVersion >= 17 Then lContract.ConId = mReader.GetLong("conId")
    .Symbol = mReader.GetString("Symbol")
    .Sectype = gTwsSecTypeFromString(mReader.GetString("Sec type"))
    .Expiry = mReader.GetString("Expiry")
    .Strike = mReader.GetDouble("Strike")
    .OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
    If pVersion >= 32 Then .Multiplier = mReader.GetLong("Multiplier")
    If .Multiplier = 0 Then .Multiplier = 1
    .Exchange = mReader.GetString("Exchange")
    .CurrencyCode = mReader.GetString("Currency")
    If pVersion >= 2 Then .LocalSymbol = mReader.GetString("Local Symbol")
    If pVersion >= 32 Then .TradingClass = mReader.GetString("Trading Class")
End With

' read Order fields
With lOrder
    .Action = gTwsOrderActionFromString(mReader.GetString("Action"))
    .TotalQuantity = mReader.GetLong("Quantity")
    .OrderType = gTwsOrderTypeFromString(mReader.GetString("Order type"))
    If .OrderType = TwsOrderTypeNone Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid OrderYype"
    If pVersion < 29 Then
        .LmtPrice = mReader.GetDouble("Limit price")
    Else
        .LmtPrice = mReader.GetDoubleMax("Limit price")
    End If
    If pVersion < 30 Then
        .AuxPrice = mReader.GetDouble("Aux price")
    Else
        .AuxPrice = mReader.GetDoubleMax("Aux price")
    End If
    .Tif = gTwsOrderTIFFromString(mReader.GetString("Time in force"))
    .OcaGroup = mReader.GetString("OCA group")
    .Account = mReader.GetString("Account")
    .OpenClose = mReader.GetString("Open/close")
    .Origin = mReader.GetLong("Origin")
    .OrderRef = mReader.GetString("Order ref")
    If pVersion >= 3 Then .ClientID = mReader.GetLong("Client id")
    
    If pVersion >= 4 Then
        .PermId = mReader.GetLong("Perm id")
        .OutsideRth = mReader.GetBoolean("Outside RTH")
        .Hidden = mReader.GetBoolean("Hidden")
        .DiscretionaryAmt = mReader.GetDouble("Discr amt")
    End If

    If pVersion >= 5 Then .GoodAfterTime = mReader.GetString("Good after Time")
    
    If pVersion >= 6 Then mReader.GetString "Shares allocation"
    
    If pVersion >= 7 Then
        .FaGroup = mReader.GetString("FA Group")
        .FaMethod = mReader.GetString("FA method")
        .FaPercentage = mReader.GetString("FA Percentage")
        .FaProfile = mReader.GetString("FA Profile")
    End If
    
    If pVersion >= 8 Then .GoodTillDate = mReader.GetString("Good till date")
    
    If pVersion >= 9 Then
        .Rule80A = mReader.GetString("Rule80A")
        .PercentOffset = mReader.GetDouble("Percent Offset")
        .SettlingFirm = mReader.GetString("Settling Firm")
        .ShortSaleSlot = mReader.GetLong("shortSaleSlot")
        .DesignatedLocation = mReader.GetString("designatedLocation")
        If pVersion >= 23 Then .ExemptCode = mReader.GetLong("ExemptCode")
        .AuctionStrategy = mReader.GetLong("Auction Strategy")
        .StartingPrice = mReader.GetDouble("starting Price")
        .StockRefPrice = mReader.GetDouble("stockRefPrice")
        .Delta = mReader.GetDouble("delta")
        .StockRangeLower = mReader.GetDouble("stockRangeLower")
        .StockRangeUpper = mReader.GetDouble("stockRangeUpper")
        .DisplaySize = mReader.GetLong("displaySize")
        'setproperty .providerProperties, "rthOnly"
        .BlockOrder = mReader.GetBoolean("blockOrder")
        .SweepToFill = mReader.GetBoolean("sweepToFill")
        .AllOrNone = mReader.GetBoolean("allOrNone")
        .MinQty = mReader.GetLong("minQty")
        .OcaType = mReader.GetLong("ocaType")
        .ETradeOnly = mReader.GetBoolean("eTradeOnly")
        .FirmQuoteOnly = mReader.GetBoolean("firmQuoteOnly")
        .NbboPriceCap = mReader.GetDouble("nbboPriceCap")
    End If

    If pVersion >= 10 Then
        .ParentId = mReader.GetLong("ParentId")
        .TriggerMethod = mReader.GetLong("TriggerMethod")
    End If

    If pVersion >= 11 Then
        .Volatility = mReader.GetDouble("Volatility")
        .VolatilityType = mReader.GetLong("VolatilityType")
        If pVersion = 11 Then
            .DeltaNeutralOrderType = IIf(mReader.GetLong("DeltaNeutralOrderType") = 0, TwsOrderTypes.TwsOrderTypeNone, TwsOrderTypes.TwsOrderTypeMarket)
        Else ' pVersion 12 and up
            Dim lDeltaNeutralOrderType As String
            lDeltaNeutralOrderType = mReader.GetString("DeltaNeutralOrderType")
            .DeltaNeutralOrderType = gTwsOrderTypeFromString(lDeltaNeutralOrderType)
            .DeltaNeutralAuxPrice = mReader.GetDoubleMax("DeltaNeutralAuxPrice")
            
            ' NB: the commented-out line below is incorrect because TWS actually sends the
            ' following fields even when it sends 'NONE' for DeltaNeutralOrderType
            'If pVersion >= 27 And .DeltaNeutralOrderType <> TwsOrderTypes.TwsOrderTypeNone Then
            If pVersion >= 27 And lDeltaNeutralOrderType <> "" Then
                .DeltaNeutralConId = mReader.GetLong("DeltaNeutralConId")
                .DeltaNeutralSettlingFirm = mReader.GetString("DeltaNeutralSettlingFirm")
                .DeltaNeutralClearingAccount = mReader.GetString("DeltaNeutralClearingAccount")
                .DeltaNeutralClearingIntent = mReader.GetString("DeltaNeutralClearingIntent")
            End If

            ' NB: the commented-out line below is incorrect because TWS actually sends the
            ' following fields even when it sends 'NONE' for DeltaNeutralOrderType
            'If pVersion >= 31 And .DeltaNeutralOrderType <> TwsOrderTypes.TwsOrderTypeNone Then
            If pVersion >= 31 And lDeltaNeutralOrderType <> "" Then
                .DeltaNeutralOpenClose = mReader.GetString("DeltaNeutralOpenClose")
                .DeltaNeutralShortSale = mReader.GetBoolean("DeltaNeutralShortSale")
                .DeltaNeutralShortSaleSlot = mReader.GetLong("DeltaNeutralShortSaleSlot")
                .DeltaNeutralDesignatedLocation = mReader.GetString("DeltaNeutralDesignatedLocation")
            End If
        End If
        .ContinuousUpdate = mReader.GetLong("ContinuousUpdate")
        If mServerVersion = 26 Then
            .StockRangeLower = mReader.GetDouble("StockRangeLower")
            .StockRangeUpper = mReader.GetDouble("StockRangeUpper")
        End If
        .ReferencePriceType = mReader.GetLong("ReferencePriceType")
    End If
    
    If pVersion >= 13 Then .TrailStopPrice = mReader.GetDoubleMax("TrailStopPrice")

    If pVersion >= 30 Then .TrailingPercent = mReader.GetDoubleMax("TrailingPercent")
    
    If pVersion >= 14 Then
        .BasisPoints = mReader.GetDoubleMax("BasisPoints")
        .BasisPointsType = mReader.GetLongMax("BasisPointsType")
        lContract.ComboLegsDescription = mReader.GetString("ComboLegsDescrip")
    End If

    If pVersion >= 29 Then
        Dim lComboLegsCount As Long: lComboLegsCount = mReader.GetLong("ComboLegsCount")
        If lComboLegsCount > 0 Then
            For i = 1 To lComboLegsCount
                Dim lComboLeg As New TwsComboLeg
                lComboLeg.ConId = mReader.GetLong("ConId" & i)
                lComboLeg.Ratio = mReader.GetLong("Ratio" & i)
                lComboLeg.Action = mReader.GetString("Action" & i)
                lComboLeg.Exchange = mReader.GetString("Exchange" & i)
                lComboLeg.OpenClose = mReader.GetLong("OpenClose" & i)
                lComboLeg.ShortSaleSlot = mReader.GetLong("ShortSaleSlot" & i)
                lComboLeg.DesignatedLocation = mReader.GetString("DesignatedLocation" & i)
                lComboLeg.ExemptCode = mReader.GetLong("ExemptCode" & i)

                lContract.ComboLegs.Add (lComboLeg)
            Next
        End If

        Dim lOrderComboLegsCount As Long: lOrderComboLegsCount = mReader.GetLong("OrderComboLegsCount")
        If lOrderComboLegsCount > 0 Then
            ReDim lOrderComboLegs(lOrderComboLegsCount - 1) As TwsOrderComboLeg
            For i = 0 To lOrderComboLegsCount - 1
                Set lOrderComboLegs(i) = New TwsOrderComboLeg
                lOrderComboLegs(i).Price = mReader.GetDouble("Price" & (i + 1))
            Next
            .OrderComboLegs = lOrderComboLegs
        End If
    End If

    If pVersion >= 26 Then
        Dim lSmartComboRoutingParamsCount As Long: lSmartComboRoutingParamsCount = mReader.GetLong("SmartComboRoutingParamsCount")
        If lSmartComboRoutingParamsCount > 0 Then
            ReDim lSmartComboRoutingParams(lSmartComboRoutingParamsCount - 1) As TwsTagValue
            For i = 0 To lSmartComboRoutingParamsCount - 1
                lSmartComboRoutingParams(i).Tag = mReader.GetString("Tag" & (i + 1))
                lSmartComboRoutingParams(i).Value = mReader.GetString("Value" & (i + 1))
            Next
            .SmartComboRoutingParams = lSmartComboRoutingParams
        End If
    End If

    If (pVersion >= 15) Then
        If (pVersion >= 20) Then
            .ScaleInitLevelSize = mReader.GetLongMax("Scale Init Level Size")
            .ScaleSubsLevelSize = mReader.GetLongMax("Scale Subs Level Size")
        Else
            .ScaleInitLevelSize = mReader.GetLongMax("Scale Init Level Size")
        End If
        .ScalePriceIncrement = mReader.GetDoubleMax("Scale Price Increment")
    End If
    
    If pVersion >= 28 And .ScalePriceIncrement > 0# And .ScalePriceIncrement <> MaxDouble Then
        .ScalePriceAdjustValue = mReader.GetDoubleMax("ScalePriceAdjustValue")
        .ScalePriceAdjustInterval = mReader.GetLongMax("ScalePriceAdjustInterval")
        .ScaleProfitOffset = mReader.GetDoubleMax("ScaleProfitOffset")
        .ScaleAutoReset = mReader.GetBoolean("ScaleAutoReset")
        .ScaleInitPosition = mReader.GetLongMax("ScaleInitPosition")
        .ScaleInitFillQty = mReader.GetLongMax("ScaleInitFillQty")
        .ScaleRandomPercent = mReader.GetBoolean("ScaleRandomPercent")
    End If
    
    If pVersion >= 24 Then
        .HedgeType = gTwsHedgeTypeFromString(mReader.GetString("HedgeType"))
        If .HedgeType <> TwsHedgeTypes.TwsHedgeTypeNone Then .HedgeParam = mReader.GetString("HedgeParam")
    End If
    
    If pVersion >= 25 Then .OptOutSmartRouting = mReader.GetBoolean("OptOutSmartRouting")
    
    If (pVersion >= 19) Then
        .ClearingAccount = mReader.GetString("Clearing Account")
        .ClearingIntent = mReader.GetString("Clearing Intent")
    End If
    
    If pVersion >= 22 Then .NotHeld = mReader.GetBoolean("Not held")
    
    If (pVersion >= 20) Then
        Dim lUnderComp As TwsUnderComp
        If CBool(mReader.GetLong("UnderComp")) Then
            lUnderComp.ConId = mReader.GetLong("UnderComp ConId")
            lUnderComp.Delta = mReader.GetDouble("UnderComp Delta")
            lUnderComp.Price = mReader.GetDouble("UnderComp Price")
            lContract.UnderComp = lUnderComp
        End If
    End If
    
    If pVersion >= 21 Then
        .AlgoStrategy = mReader.GetString("Algo strategy")
        If .AlgoStrategy <> "" Then
            Dim algoParamsCount As Long: algoParamsCount = mReader.GetLong("Algo params count")
            If algoParamsCount > 0 Then
                ReDim lAlgoParams(algoParamsCount - 1) As TwsTagValue
                For i = 0 To algoParamsCount - 1
                    lAlgoParams(i).Tag = mReader.GetString("Tag" & (i + 1))
                    lAlgoParams(i).Value = mReader.GetString("Value" & (i + 1))
                Next
                .AlgoParams = lAlgoParams
            End If
        End If
    End If
    
    If (pVersion >= 16) Then

        .WhatIf = mReader.GetBoolean("What If")

        With lOrderState
            .Status = mReader.GetString("Status")
            .InitMargin = mReader.GetString("Init Margin")
            .MaintMargin = mReader.GetString("Maint Margin")
            .EquityWithLoan = mReader.GetString("Equity With Loan")
            .Commission = mReader.GetDoubleMax("Commission")
            .MinCommission = mReader.GetDoubleMax("Min Commission")
            .MaxCommission = mReader.GetDoubleMax("Max Commission")
            .CommissionCurrency = mReader.GetString("Commission Currency")
            .WarningText = mReader.GetString("Warning Text")
        End With
    End If
End With

logSocketInputMessage True

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyOpenOrder lOrder.OrderId, lContract, lOrder, lOrderState

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOpenOrderEnd(ByVal pVersion As Long)
Const ProcName As String = "getOpenOrderEnd"
On Error GoTo Err

logSocketInputMessage

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.EndOpenOrders
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOrderStatus(ByVal pVersion As Long)
Const ProcName As String = "getOrderStatus"
On Error GoTo Err

' NB: lAvgFillPrice and lLastFillPrice ought to be subjected to adjustPrice,
' but we can't call that here because we don't know the currency code. But not really
' an issue since we are not likely to encounter a Tws with server pVersion < 23 !

Dim lorderId As Long: lorderId = mReader.GetLong("id")
Dim lStatus As String: lStatus = mReader.GetString("Status")
Dim lFilled As Long: lFilled = mReader.GetLong("Filled")
Dim lRemaining As Long: lRemaining = mReader.GetLong("Remaining")
Dim lAvgFillPrice As Double:  lAvgFillPrice = mReader.GetDouble("Avg fill price")
Dim lLastFillPrice As Double
Dim lClientID As Long
Dim lPermId As Long
Dim lParentId As Long
Dim lWhyHeld As String

If pVersion >= 2 Then lPermId = mReader.GetLong("Perm id")

If pVersion >= 3 Then lParentId = mReader.GetLong("Parent id")

If pVersion >= 4 Then lLastFillPrice = mReader.GetDouble("Last fill price")

If pVersion >= 5 Then lClientID = mReader.GetLong("Client id")

If pVersion >= 6 Then lWhyHeld = mReader.GetString("Why held")

logSocketInputMessage True

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyOrderStatus lorderId, lStatus, lFilled, lRemaining, lAvgFillPrice, lPermId, lParentId, lLastFillPrice, lClientID, lWhyHeld

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getPerformanceTimer() As ElapsedTimer
Static sPerformanceElapsedTimer As ElapsedTimer
Const ProcName As String = "getPerformanceTimer"
On Error GoTo Err

If sPerformanceElapsedTimer Is Nothing Then Set sPerformanceElapsedTimer = New ElapsedTimer
Set getPerformanceTimer = sPerformanceElapsedTimer

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub getPortfolioValue(ByVal pVersion As Long)
Const ProcName As String = "getPortfolioValue"
On Error GoTo Err

Dim lContract As New TwsContract

If pVersion >= 6 Then lContract.ConId = mReader.GetLong("Contract id")

lContract.Symbol = mReader.GetString("Symbol")
lContract.Sectype = gTwsSecTypeFromString(mReader.GetString("Sec type"))
lContract.Expiry = mReader.GetString("Expiry")
lContract.Strike = mReader.GetDouble("Strike")
lContract.OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))

If pVersion >= 7 Then
    lContract.Multiplier = mReader.GetLong("Multiplier")
    If lContract.Multiplier = 0 Then lContract.Multiplier = 1
    lContract.PrimaryExch = mReader.GetString("Primary exchg")
End If

lContract.CurrencyCode = mReader.GetString("Currency")

If pVersion >= 2 Then lContract.LocalSymbol = mReader.GetString("Local Symbol")
If pVersion >= 8 Then lContract.TradingClass = mReader.GetString("Trading Class")

Dim lPosition As Long: lPosition = mReader.GetLong("POSITION")
Dim lMarketPrice As Double: lMarketPrice = mReader.GetDouble("Market price")
Dim lMarketValue As Double: lMarketValue = mReader.GetDouble("Market Value")
Dim lAverageCost As Double
Dim lUnrealizedPNL As Double
Dim lRealizedPNL As Double
If pVersion >= 3 Then
    lAverageCost = mReader.GetDouble("Avg cost")
    lUnrealizedPNL = mReader.GetDouble("Unrealized PNL")
    lRealizedPNL = mReader.GetDouble("Realized PNL")
End If

Dim lAccountName As String: If pVersion >= 4 Then lAccountName = mReader.GetString("Account Name")

If pVersion = 6 And mServerVersion = 39 Then lContract.PrimaryExch = mReader.GetString("Primary Exchange")

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyPortfolioUpdate lContract, lPosition, lMarketPrice, lMarketValue, lAverageCost, lUnrealizedPNL, lRealizedPNL, lAccountName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getPosition(ByVal pVersion As Long)
Const ProcName As String = "getPosition"
On Error GoTo Err

Dim lAccount As String: lAccount = mReader.GetString("Account")
Dim lContract As New TwsContract
lContract.ConId = mReader.GetLong("Contract id")
lContract.Symbol = mReader.GetString("Symbol")
lContract.Sectype = gTwsSecTypeFromString(mReader.GetString("SecType"))
lContract.Expiry = mReader.GetString("Expiry")
lContract.Strike = mReader.GetDouble("Strike")
lContract.OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
lContract.Multiplier = mReader.GetLong("Multiplier")
If lContract.Multiplier = 0 Then lContract.Multiplier = 1
lContract.Exchange = mReader.GetString("Exchange")
lContract.CurrencyCode = mReader.GetString("Currency")
lContract.LocalSymbol = mReader.GetString("LocalSymbol")
lContract.TradingClass = mReader.GetString("Trading Class")

Dim lPosition As Long: lPosition = mReader.GetLong("Position")
Dim lAverageCost As Double
If pVersion >= 3 Then lAverageCost = mReader.GetDouble("Average Cost")

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyPosition lAccount, lContract, lPosition, lAverageCost

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getPositionEnd(ByVal pVersion As Long)
Const ProcName As String = "getPositionEnd"
On Error GoTo Err

logSocketInputMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.EndPosition

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getRealtimeBars(ByVal pVersion As Long)
Const ProcName As String = "getRealtimeBars"
On Error GoTo Err

Dim lReqId As Long: lReqId = mReader.GetLong("ReqId")
Dim lTime As Date: lTime = gGetDateFromUnixSystemTime(mReader.GetDouble("System Time"))
Dim lOpen As Double: lOpen = mReader.GetDouble("Open")
Dim lHigh As Double: lHigh = mReader.GetDouble("High")
Dim lLow As Double: lLow = mReader.GetDouble("Low")
Dim lClose As Double: lClose = mReader.GetDouble("Close")
Dim lVolume As Long: lVolume = mReader.GetLong("Volume")
Dim lWap As Double: lWap = mReader.GetDouble("Wap")
Dim lCount As Long: lCount = mReader.GetLong("Count")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getScannerData(ByVal pVersion As Long)
Const ProcName As String = "getScannerData"
On Error GoTo Err

Static sRequestId As Long
Static sNumberOfElements As Long

Static sContract As TwsContract
Static sContractDetails As TwsContractDetails

Static sRank As Long
Static sDistance As String
Static sBenchmark As String
Static sProjection As String
Static sLegs As String

mReader.BeginCheckpointScope
    mReader.GetLongCheckpointed sRequestId, "Request Id"
    mReader.GetLongCheckpointed sNumberOfElements, "Number of elements"
    
    Dim i As Long
    For i = 0 To sNumberOfElements - 1
        mReader.BeginCheckpointScope
            Dim lLongValue As Long
            Dim lStringValue As String
            Dim lDoubleValue As Double
            
            If mReader.GetLongCheckpointed(sRank, "Rank") Then
                Set sContract = New TwsContract
                Set sContractDetails = New TwsContractDetails
                sContractDetails.Summary = sContract
            End If
            
            If (pVersion >= 3) Then If mReader.GetLongCheckpointed(lLongValue, "ConId") Then sContract.ConId = lLongValue
            
            If mReader.GetStringCheckpointed(lStringValue, "Symbol") Then sContract.Symbol = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "Sec type") Then sContract.Sectype = gTwsSecTypeFromString(lStringValue)
            If mReader.GetStringCheckpointed(lStringValue, "Expiry") Then sContract.Expiry = lStringValue
            If mReader.GetDoubleCheckpointed(lDoubleValue, "Stirke") Then sContract.Strike = lDoubleValue
            If mReader.GetStringCheckpointed(lStringValue, "Right") Then sContract.OptRight = gTwsOptionRightFromString(lStringValue)
            If mReader.GetStringCheckpointed(lStringValue, "Exchange") Then sContract.Exchange = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "Currency") Then sContract.CurrencyCode = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "Local Symbol") Then sContract.LocalSymbol = lStringValue
            
            If mReader.GetStringCheckpointed(lStringValue, "Market name") Then sContractDetails.MarketName = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "TradingClass") Then sContract.TradingClass = lStringValue
            
            mReader.GetStringCheckpointed sDistance, "Distance"
            mReader.GetStringCheckpointed sBenchmark, "Benchmark"
            mReader.GetStringCheckpointed sProjection, "Projection"
            If (pVersion >= 2) Then mReader.GetStringCheckpointed sLegs, "Legs Str"
        
            If Not mScannerDataConsumer Is Nothing Then mScannerDataConsumer.NotifyScannerData sRequestId, sRank, sContractDetails, sDistance, sBenchmark, sProjection, sLegs
        mReader.EndCheckpointScope
    Next
mReader.EndCheckpointScope

logSocketInputMessage

If Not mScannerDataConsumer Is Nothing Then mScannerDataConsumer.EndScannerData sRequestId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getScannerParameters(ByVal pVersion As Long)
Const ProcName As String = "getScannerParameters"
On Error GoTo Err

Dim xml  As String: xml = mReader.GetString("XML")

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickEFP(ByVal pVersion As Long)
Const ProcName As String = "getTickEFP"
On Error GoTo Err

Dim lTickerId As Long: lTickerId = mReader.GetLong("Ticker Id")
Dim lTickType As Long: lTickType = mReader.GetLong("Tick Type")
Dim lBasisPoints  As Double: lBasisPoints = mReader.GetDouble("Basis Points")
Dim lFormattedBasisPoints As String: lFormattedBasisPoints = mReader.GetString("Formatted BasisP oints")
Dim lImpliedFuturesPrice As Double: lImpliedFuturesPrice = mReader.GetDouble("lImplied Futures Price")
Dim lHoldDays As Long: lHoldDays = mReader.GetLong("Hold Days")
Dim lFutureExpiry As String: lFutureExpiry = mReader.GetString("Future Expiry")
Dim lDividendImpact As Double: lDividendImpact = mReader.GetDouble("Dividend Impact")
Dim lDividendsToExpiry As Double: lDividendsToExpiry = mReader.GetDouble("Dividends To Expiry")

logSocketInputMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickEFP lTickerId, lTickType, lBasisPoints, lFormattedBasisPoints, lImpliedFuturesPrice, lHoldDays, lFutureExpiry, lDividendImpact, lDividendsToExpiry

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickGeneric(ByVal pVersion As Long)
Const ProcName As String = "getTickGeneric"
On Error GoTo Err

Dim lTickerId As Long: lTickerId = mReader.GetLong("tickerId")
Dim lTickType As Long: lTickType = mReader.GetLong("tickType")
Dim lValue As Double: lValue = mReader.GetDouble("Value")

logSocketInputMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickGeneric getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lValue

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickOptionComputation(ByVal pVersion As Long)
Const ProcName As String = "getTickOptionComputation"
On Error GoTo Err

Dim lTickerId As Long: lTickerId = mReader.GetLong("tickerId")
Dim lTickType As Long: lTickType = mReader.GetLong("tickType")
Dim lImpliedVol As Double: lImpliedVol = mReader.GetDouble("impliedVol")

If lImpliedVol < 0 Then
    ' -1 is the "not yet computed" indicator
    lImpliedVol = MaxDouble
End If

Dim lDelta As Double: lDelta = mReader.GetDouble("delta")
If Abs(lDelta) > 1 Then
    ' -2 is the "not yet computed" indicator
    lDelta = MaxDouble
End If

Dim lOptPrice As Double: lOptPrice = MaxDouble
Dim lPvDividend As Double: lPvDividend = MaxDouble
Dim lGamma As Double: lGamma = MaxDouble
Dim lVega As Double: lVega = MaxDouble
Dim lTheta As Double: lTheta = MaxDouble
Dim lUndPrice As Double: lUndPrice = MaxDouble

If pVersion >= 6 Or lTickType = TwsTickTypes.TwsTickModelOption Then   ' introduced in pVersion == 5
    lOptPrice = mReader.GetDouble("optPrice")
    If lOptPrice < 0 Then   ' -1 is the "not yet computed" indicator
        lOptPrice = MaxDouble
    End If
    lPvDividend = mReader.GetDouble("pvDividend")
    If lPvDividend < 0 Then    ' -1 is the "not yet computed" indicator
        lPvDividend = MaxDouble
    End If
End If

If (pVersion >= 6) Then
    lGamma = mReader.GetDouble("gamma")
    If (Abs(lGamma) > 1) Then ' -2 is the "not yet computed" indicator
        lGamma = MaxDouble
    End If
    lVega = mReader.GetDouble("Vega")
    If (Abs(lVega) > 1) Then ' -2 is the "not yet computed" indicator
        lVega = MaxDouble
    End If
    lTheta = mReader.GetDouble("theta")
    If (Abs(lTheta) > 1) Then ' -2 is the "not yet computed" indicator
        lTheta = MaxDouble
    End If
    lUndPrice = mReader.GetDouble("undPrice")
    If (lUndPrice < 0) Then ' -1 is the "not yet computed" indicator
        lUndPrice = MaxDouble
    End If
End If

logSocketInputMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickOptionComputation getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lImpliedVol, lDelta, lOptPrice, lPvDividend, lGamma, lVega, lTheta, lUndPrice

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickPrice(ByVal pVersion As Long)
Const ProcName As String = "getTickPrice"
On Error GoTo Err

Dim lTickType As Long
Dim lTickerId As Long
Dim lPrice As Double
Dim lSize As Long
Dim lCanAutoExecute As Boolean

lTickerId = mReader.GetLong("Ticker id")
lTickType = mReader.GetLong("Tick type")
lPrice = mReader.GetDouble("Price")

If pVersion >= 2 Then lSize = mReader.GetLong("Size")
If pVersion >= 3 Then lCanAutoExecute = mReader.GetBoolean("Can autoexecute")

logSocketInputMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickPrice getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lPrice, lSize, lCanAutoExecute

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickSize(ByVal pVersion As Long)
Const ProcName As String = "getTickSize"
Dim lTickType As Long
Dim lTickerId As Long
Dim lSize As Long

On Error GoTo Err

lTickerId = mReader.GetLong("Ticker id")
lTickType = mReader.GetLong("Tick type")
lSize = mReader.GetLong("Size")

logSocketInputMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickSize getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lSize

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName

End Sub

Private Sub getTickSnapshotEnd(ByVal pVersion As Long)
Const ProcName As String = "getTickSnapshotEnd"
On Error GoTo Err

Dim lReqId As Long: lReqId = mReader.GetLong("reqID")
logSocketInputMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.EndTickSnapshot getCallersRequestIdFromTwsMarketDataRequestId(lReqId)

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickString(ByVal pVersion As Long)
Const ProcName As String = "getTickString"
On Error GoTo Err

Dim lTickerId As Long: lTickerId = mReader.GetLong("tickerId")
Dim lTickType As Long: lTickType = mReader.GetLong("tickType")
Dim lValue As String: lValue = mReader.GetString("Value")

logSocketInputMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickString getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lValue

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub logSocketInputMessage(Optional ByVal pIgnoreLogLevel As Boolean)
Const ProcName As String = "logSocketInputMessage"
On Error GoTo Err

If Not mLogTwsMessages Then Exit Sub

If pIgnoreLogLevel Then
    gSocketLogger.Log mReader.Message, ProcName, ModuleName, LogLevelNormal
ElseIf gSocketLogger.IsLoggable(LogLevelHighDetail) Then
    gSocketLogger.Log mReader.Message, ProcName, ModuleName, LogLevelHighDetail
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub negotiateConnection()
Const ProcName As String = "negotiateConnection"
On Error GoTo Err

Dim lTrashLong As Long
Dim lTrashString As String
Dim lServerTime As String

mReader.BeginMessage True
' first data received is plain Server pVersion number
mReader.GetLongCheckpointed mServerVersion, "Server Version"

' in some circumstances, some versions of Tws send malformed
' ERR_MSG messages before sending the pVersion number, so we
' need to skip these if they occur. (They are malformed in that
' they don't include the id or error code.)

Do While mServerVersion = TwsSocketInMsgTypes.ERR_MSG
    mReader.BeginCheckpointScope
        mReader.GetLongCheckpointed lTrashLong, "Invalid pVersion"
            
        mReader.GetStringCheckpointed lTrashString, "Invalid errormsg"
        
        logSocketInputMessage
        mReader.EndMessage
        
        mReader.BeginMessage True
        mReader.GetLongCheckpointed mServerVersion, "Server pVersion"
    mReader.EndCheckpointScope
Loop

logSocketInputMessage
mReader.EndMessage

If mServerVersion >= 20 Then
    mReader.BeginMessage True
    If mReader.GetStringCheckpointed(lServerTime, "TWS Time") Then gLog "TWS Time at connection", ModuleName, ProcName, lServerTime
    logSocketInputMessage
    mReader.EndMessage
End If

Assert mServerVersion >= ReqdServerVersion, "TWS is out of date and needs to be upgraded"

mApiConnectionEstablished = True
RaiseEvent ApiConnectionEstablished(mServerVersion)
    
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub processContractError(ByVal pRequestId As Long, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "processContractError"
On Error GoTo Err

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.NotifyError pRequestId, pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub updateStats(ByVal pMessageId As Long, ByVal pMessageElapsedTime As Single)
Static sEventCount As Long

Const ProcName As String = "updateStats"
On Error GoTo Err

If mLogTwsMessages Then
    mStatsRecorder.UpdateMessageTypeStats pMessageId, pMessageElapsedTime
    sEventCount = sEventCount + 1
    
    
    If getElapsedTimer.ElapsedTimeMicroseconds >= 10000000 Then
        gLog "Event rate per second", ModuleName, ProcName, pMsgQualifier:=Format(sEventCount / 10, "0.0"), pLogLevel:=LogLevelDetail
        Debug.Print "Event rate per second = " & Format(sEventCount / 10, "0.0")
        sEventCount = 0
        getElapsedTimer.StartTiming
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

