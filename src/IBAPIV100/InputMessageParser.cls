VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputMessageParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Event PaperTradingDisclaimerNotAcknowledged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TwsHistDataStartInfo
    requestId               As Long
    StartDate               As String
    EndDate                 As String
    BarCount                As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "InputMessageParser"

'@================================================================================
' Member variables
'@================================================================================

Private mServerVersion                              As Long

Private mReader                                     As BufferedReader

Private mConnectionStatusConsumer                   As IConnectionStatusConsumer
Private mHistDataConsumer                           As IHistDataConsumer
Private mContractDetailsConsumer                    As IContractDetailsConsumer
Private mOrderInfoConsumer                          As IOrderInfoConsumer
Private mMarketDataConsumer                         As IMarketDataConsumer
Private mMarketDepthConsumer                        As IMarketDepthConsumer
Private mAccountDataConsumer                        As IAccountDataConsumer
Private mScannerDataConsumer                        As IScannerDataConsumer
Private mErrorAndNotificationConsumer               As IErrorAndNotificationConsumer

Private mIsTwsConnectedToIBServers                  As Boolean

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

Friend Property Let AccountDataConsumer(ByVal Value As IAccountDataConsumer)
Set mAccountDataConsumer = Value
End Property

Friend Property Let ConnectionStatusConsumer(ByVal Value As IConnectionStatusConsumer)
Set mConnectionStatusConsumer = Value
End Property

Friend Property Let ContractDetailsConsumer(ByVal Value As IContractDetailsConsumer)
Set mContractDetailsConsumer = Value
End Property

Friend Property Let HistDataConsumer(ByVal Value As IHistDataConsumer)
Set mHistDataConsumer = Value
End Property

Friend Property Let ErrorAndNotificationConsumer(ByVal Value As IErrorAndNotificationConsumer)
Set mErrorAndNotificationConsumer = Value
End Property

Friend Property Let MarketDataConsumer(ByVal Value As IMarketDataConsumer)
Set mMarketDataConsumer = Value
End Property

Friend Property Let MarketDepthConsumer(ByVal Value As IMarketDepthConsumer)
Set mMarketDepthConsumer = Value
End Property

Friend Property Get IsTwsConnectedToIBServers() As Boolean
IsTwsConnectedToIBServers = mIsTwsConnectedToIBServers
End Property

Friend Property Let OrderInfoConsumer(ByVal Value As IOrderInfoConsumer)
Set mOrderInfoConsumer = Value
End Property

Friend Property Let ScannerDataConsumer(ByVal Value As IScannerDataConsumer)
Set mScannerDataConsumer = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub GetMessage(ByVal pMessageId As TwsSocketInMsgTypes, ByVal pVersion As Long)
Const ProcName As String = "GetMessage"
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
Case TwsSocketInMsgTypes.TickRequestParams
    getTickRequestParams
Case TwsSocketInMsgTypes.MarketRule
    getMarketRule
Case TwsSocketInMsgTypes.OptionParameter
    getOptionParameter
Case TwsSocketInMsgTypes.OptionParameterEnd
    getOptionParameterEnd
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid message code from Tws"
End Select

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Begin( _
                ByVal pServerVersion As Long, _
                ByVal pReader As BufferedReader)
Const ProcName As String = "Begin"
On Error GoTo Err

mServerVersion = pServerVersion
Set mReader = pReader

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
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

Private Sub getAccountDownloadEnd(ByVal pVersion As Long)
Const ProcName As String = "getAccountDownloadEnd"
On Error GoTo Err

mReader.BeginLogMessage

Dim accountName As String
accountName = mReader.GetString("Account name")

mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.EndAccountValue accountName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountSummary(ByVal pVersion As Long)
Const ProcName As String = "getAccountSummary"
On Error GoTo Err

mReader.BeginLogMessage

Dim lRequestId As Long: lRequestId = mReader.GetLong("Request Id")
Dim lAccount As String: lAccount = mReader.GetString("Account")
Dim lTag As String: lTag = mReader.GetString("Tag")
Dim lValue As String: lValue = mReader.GetString("Value")
Dim lCurrency As String: lCurrency = mReader.GetString("Currency")

mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyAccountSummary lRequestId, lAccount, lTag, lValue, lCurrency

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountSummaryEnd(ByVal pVersion As Long)
Const ProcName As String = "getAccountSummaryEnd"
On Error GoTo Err

mReader.BeginLogMessage

Dim lRequestId As Long: lRequestId = mReader.GetLong("RequestId")

mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.EndAccountSummary lRequestId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountTime(ByVal pVersion As Long)
Const ProcName As String = "getAccountTime"
On Error GoTo Err

mReader.BeginLogMessage

Dim accountTime As String: accountTime = mReader.GetString("Account Time")

mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyAccountTime accountTime

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getAccountValue(ByVal pVersion As Long)
Const ProcName As String = "getAccountValue"
On Error GoTo Err

mReader.BeginLogMessage

Dim key As String: key = mReader.GetString("Key")
Dim val As String: val = mReader.GetString("Value")
Dim cur As String: cur = mReader.GetString("Currency")
Dim accountName As String: If pVersion >= 2 Then accountName = mReader.GetString("Account Name")

mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyAccountValue key, val, cur, accountName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getBondContractData(ByVal pVersion As Long)
' not currently supported
Const ProcName As String = "getBondContractData"
On Error GoTo Err

mReader.BeginLogMessage

Dim lRequestId As Long
If (pVersion >= 3) Then lRequestId = mReader.GetLong("Req Id")
lRequestId = GIdManager.GetCallerId(lRequestId, IdTypeContractData)
mReader.GetString "Symbol"
mReader.GetString "SecType"
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
    Dim lSecIdList As Parameters
    Set lSecIdList = mReader.GetParameters( _
                                mReader.GetLong("SecIdListCount"), _
                                "SecIdList")
End If

mReader.EndLogMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.NotifyError lRequestId, ErrorCodes.ErrUnsupportedOperationException, "Bonds not yet supported"

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getCommissionReport(ByVal pVersion As Long)
Const ProcName As String = "getCommissionReport"
On Error GoTo Err

mReader.BeginLogMessage

Dim lCommissionReport As New TwsCommissionReport
lCommissionReport.ExecId = mReader.GetString("ExecId")
lCommissionReport.Commission = mReader.GetDouble("Commission")
lCommissionReport.CurrencyCode = mReader.GetString("CurrencyCode")
lCommissionReport.RealizedPNL = mReader.GetDouble("RealizedPNL")
lCommissionReport.Yield = mReader.GetDouble("Yield")
lCommissionReport.YieldRedemptionDate = mReader.GetLong("YieldRedemptionDate")

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getContractData(ByVal pVersion As Long)
Const ProcName As String = "getContractData"
On Error GoTo Err

mReader.BeginLogMessage

Dim lReqId As Long
Dim lContractSpec As New TwsContractSpecifier
Dim lContract As New TwsContract

lContract.Specifier = lContractSpec

If (pVersion >= 3) Then lReqId = GIdManager.GetCallerId(mReader.GetLong("Req Id"), IdTypeContractData)

lContractSpec.Symbol = mReader.GetString("Symbol")
lContractSpec.SecType = gTwsSecTypeFromString(mReader.GetString("Sec type"))

Dim lExpiry As String: lExpiry = mReader.GetString("Expiry")
If lExpiry <> "" Then
    Dim ar() As String
    If InStr(1, lExpiry, "-") <> 0 Then
        ar = Split(lExpiry, "-")
    Else
        ar = Split(lExpiry, " ")
    End If
    If UBound(ar) >= 0 Then lContractSpec.Expiry = ar(0)
    If UBound(ar) >= 1 Then lContract.LastTradeTime = CDate(ar(1))
    If UBound(ar) >= 2 Then lContract.TimeZoneId = ar(2)
End If

lContractSpec.Strike = mReader.GetDouble("Strike")
lContractSpec.OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
lContractSpec.Exchange = mReader.GetString("Exchange")
lContractSpec.CurrencyCode = mReader.GetString("Currency")
lContractSpec.LocalSymbol = mReader.GetString("Local Symbol")
lContract.MarketName = mReader.GetString("Market Name")
lContractSpec.TradingClass = mReader.GetString("Trading class")
lContractSpec.ConId = mReader.GetLong("Contract id")

lContract.MinTick = mReader.GetDouble("Minimum tick")
' note that IB gives very small tick sizes for some UK stocks, eg 0.000001 for VOD and TSCO
' (possibly all FTSE 100 stocks). This causes us tick encoding problems, so we up it to
' something that reflects reality!
If lContractSpec.CurrencyCode = "GBP" And lContractSpec.SecType = TwsSecTypeStock And lContract.MinTick < 0.0001 Then lContract.MinTick = 0.0001

If mServerVersion >= ApiServerVersions.MD_SIZE_MULTIPLIER And _
    mServerVersion < ApiServerVersions.SIZE_RULES _
Then
    ' MDSizeMultiplier is no longer used
    mReader.GetLong ("MDSizeMultiplier")
End If

lContractSpec.Multiplier = mReader.GetDouble("Multiplier")

lContract.OrderTypes = mReader.GetString("Order types")
lContract.ValidExchanges = mReader.GetString("Valid exchanges")
    
lContract.PriceMagnifier = mReader.GetLong("Price magnifier")

lContract.MinTick = lContract.MinTick * lContract.PriceMagnifier
    
lContract.UnderConId = mReader.GetLong("UnderConId")

lContract.LongName = mReader.GetString("Long name")
lContractSpec.PrimaryExch = mReader.GetString("Primary Exchange")

lContract.ContractMonth = mReader.GetString("ContractMonth")
lContract.Industry = mReader.GetString("Industry")
lContract.Category = mReader.GetString("Category")
lContract.Subcategory = mReader.GetString("Subcategory")
lContract.TimeZoneId = mReader.GetString("TimeZoneId")
lContract.TradingHours = mReader.GetString("TradingHours")
lContract.LiquidHours = mReader.GetString("LiquidHours")

lContract.EvRule = mReader.GetString("EvRule")
lContract.EvMultiplier = mReader.GetDouble("EvMultiplier")

lContract.SecIdList = mReader.GetParameters( _
                                mReader.GetLong("SecIdListCount"), _
                                "SecIdList")

If mServerVersion >= ApiServerVersions.AGG_GROUP Then _
    lContract.AggGroup = mReader.GetLong("Agg Group")

If mServerVersion >= ApiServerVersions.UNDERLYING_INFO Then
    lContract.UnderSymbol = mReader.GetString("Under Symbol")
    lContract.UnderSecType = gTwsSecTypeFromString(mReader.GetString("UnderSecType"))
End If

If mServerVersion >= ApiServerVersions.MARKET_RULES Then lContract.MarketRuleIds = mReader.GetString("Market Rule Ids")

If mServerVersion >= ApiServerVersions.REAL_EXPIRATION_DATE Then lContract.RealExpirationDate = mReader.GetString("Real Expiration Date")

        If mServerVersion >= ApiServerVersions.STOCK_TYPE Then lContract.StockType = mReader.GetString("StockType")

        If mServerVersion >= ApiServerVersions.FRACTIONAL_SIZE_SUPPORT And mServerVersion < ApiServerVersions.SIZE_RULES Then
            ' SizeMinTick is no longer used
            mReader.GetDecimal "SizeMinTick"
        End If

        If mServerVersion >= ApiServerVersions.SIZE_RULES Then
            lContract.MinimumSize = mReader.GetDecimal("MinimumSize")
            lContract.SizeIncrement = mReader.GetDecimal("SizeIncrement")
            lContract.SuggestedSizeIncrement = mReader.GetDecimal("SuggestedSizeIncrement")
        End If

mReader.EndLogMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.NotifyContract lReqId, lContract

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getContractEnd(ByVal pVersion As Long)
Const ProcName As String = "getContractEnd"
On Error GoTo Err

mReader.BeginLogMessage

Dim reqId As Long
reqId = GIdManager.GetCallerId(mReader.GetLong("Req Id"), IdTypeContractData)

mReader.EndLogMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.EndContracts reqId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getCurrentTime(ByVal pVersion As Long)
Const ProcName As String = "getCurrentTime"
On Error GoTo Err

mReader.BeginLogMessage

Dim lSystemTime As Double: lSystemTime = mReader.GetDouble("SystemTime")
Dim lTime As Date: lTime = gGetDateFromUnixSystemTime(lSystemTime)

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getDeltaNeutralValidation(ByVal pVersion As Long)
Const ProcName As String = "getDeltaNeutralValidation"
On Error GoTo Err

mReader.BeginLogMessage

Dim reqId As Long: reqId = mReader.GetLong("ReqID")

Dim ConId As Long: ConId = mReader.GetLong("ConID")
Dim Delta As Double: Delta = mReader.GetDouble("Delta")
Dim Price As Double: Price = mReader.GetDouble("Price")

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getErrorMsg(ByVal pVersion As Long)
Const ProcName As String = "getErrorMsg"
On Error GoTo Err

mReader.BeginLogMessage True

Static sLostConnectionToIb As Boolean

Dim id As Long: id = mReader.GetLong("Id")
Debug.Print "Error id: " & id

Dim lErrorCode As Long: lErrorCode = mReader.GetLong("Error code")
Debug.Print "Error code: " & lErrorCode

Dim lErrorMsg As String: lErrorMsg = mReader.GetString("Error msg")
Debug.Print "Error msg: " & lErrorMsg

Dim advancedOrderRejectJson As String
If mServerVersion >= ApiServerVersions.ADVANCED_ORDER_REJECT Then
    advancedOrderRejectJson = mReader.GetString("Advanced order reject")
End If

mReader.EndLogMessage

' first handle some messages that relate to specific ids, but need to be handled
' generically and treated as program errors, since they indicate that something is
' wrong in this code
Select Case lErrorCode
Case 320, 321, 322, 323 ' server error when processing request
    gLog "Error (" & lErrorCode & "; id=" & id & ") from Tws: " & lErrorMsg, ModuleName, ProcName
    gNotifyUnhandledError Nothing, ProcName, ModuleName, , lErrorCode, "id=" & id & "; message=" & lErrorMsg
    Exit Sub
End Select

Select Case GIdManager.GetIdType(id)
Case IdTypeContractData
    processContractError GIdManager.GetCallerId(id, IdTypeContractData), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeMarketData
    If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyError GIdManager.GetCallerId(id, IdTypeMarketData), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeMarketDepth
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.NotifyError GIdManager.GetCallerId(id, IdTypeMarketDepth), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeHistoricalData
    If lErrorCode = 165 Then
        gLog "Connected to IB Historical Market Data Service - " & lErrorMsg, ModuleName, ProcName
    Else
        If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.NotifyError GIdManager.GetCallerId(id, IdTypeHistoricalData), lErrorCode, lErrorMsg
    End If
    Exit Sub
Case IdTypeOrder
    If sLostConnectionToIb And lErrorCode = 200 Then Exit Sub
    If advancedOrderRejectJson <> "" Then gLog "Order id: " & id & " ; advanced order rejection: " & vbCrLf & advancedOrderRejectJson, ModuleName, ProcName
    If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyError id, lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeScanner
    Const HistMarketDataServiceErrorMessagePrefix As String = "Historical Market Data Service error message:"
    If lErrorCode = 162 Then
        Dim l As Long: l = InStr(1, lErrorMsg, HistMarketDataServiceErrorMessagePrefix)
        If l <> 0 Then
            lErrorMsg = Mid$(lErrorMsg, l + Len(HistMarketDataServiceErrorMessagePrefix))
        End If
    End If
    If Not mScannerDataConsumer Is Nothing Then mScannerDataConsumer.NotifyError GIdManager.GetCallerId(id, IdTypeScanner), lErrorCode, lErrorMsg
    Exit Sub
End Select


Select Case lErrorCode
Case 316
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.ResetMarketDepth pReEstablish:=True
Case 317
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.ResetMarketDepth pReEstablish:=False
Case 326    ' ClientID already in use
    gLog "TWS/Gateway reports clientID already in use", ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
Case 1100, 2110
    ' connectivity between Tws and IB has been lost
    
    sLostConnectionToIb = True
    mIsTwsConnectedToIBServers = False
    
    gLog "Connection to IB has been lost", ModuleName, ProcName
    
    If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyIBServerConnectionClosed
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
    
Case 1101
    ' connectivity between Tws and IB has been restored, but data has been
    ' lost, so need to reestablish market data and market depth requests

    sLostConnectionToIb = False
    mIsTwsConnectedToIBServers = True
    
    gLog "Connection to IB recovered: market data re-established", ModuleName, ProcName
    
    If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.ResetMarketDepth True
    If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyIBServerConnectionRecovered True
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
    
Case 1102
    ' connectivity between Tws and IB has been restored without loss of data
    ' Now need to reassociate Order ids with Tws

    sLostConnectionToIb = False
    mIsTwsConnectedToIBServers = True
    
    gLog "Connection to IB recovered: no loss of data", ModuleName, ProcName
    
    If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyIBServerConnectionRecovered False
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
    
Case 2103, 2104, 2105, 2106, 2107, 2108
    gLog lErrorMsg, ModuleName, ProcName
    If Not mErrorAndNotificationConsumer Is Nothing Then mErrorAndNotificationConsumer.NotifyApiEvent lErrorCode, lErrorMsg
Case 10141  ' Paper trading disclaimer must first be accepted for API connection
    RaiseEvent PaperTradingDisclaimerNotAcknowledged
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

mReader.BeginLogMessage True

Dim lReqId As Long: lReqId = -1
If (pVersion >= 7) Then
    lReqId = mReader.GetLong("ReqId")
    If lReqId <> -1 Then lReqId = GIdManager.GetCallerId(lReqId, IdTypeExecution)
End If

Dim lExecution As New TwsExecution
lExecution.OrderId = mReader.GetLong("OrderId")

Dim lContract As New TwsContractSpecifier
With lContract
    If (pVersion >= 5) Then .ConId = mReader.GetLong("Con Id")
    .Symbol = mReader.GetString("Symbol")
    .SecType = gTwsSecTypeFromString(mReader.GetString("Sec type"))
    .Expiry = mReader.GetString("Expiry")
    .Strike = mReader.GetDouble("Strike")
    .OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
    If pVersion >= 9 Then .Multiplier = mReader.GetDouble("Multiplier")
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
    .Shares = mReader.GetDecimal("Shares")
    .Price = mReader.GetDouble("Price")
    .Price = adjustPrice(.Price, lContract.CurrencyCode)
        
    If pVersion >= 2 Then .PermId = mReader.GetLong("Perm Id")
    
    If pVersion >= 3 Then .ClientID = mReader.GetLong("Client id")
       
    If pVersion >= 4 Then .Liquidation = mReader.GetBoolean("Liquidation")
    
    If (pVersion >= 6) Then
        .CumQty = mReader.GetDecimal("Cum Qty")
        .AvgPrice = mReader.GetDouble("Avg Price")
    End If
    
    If pVersion >= 8 Then .OrderRef = mReader.GetString("OrderRef")
    
    If pVersion >= 9 Then
        .EvRule = mReader.GetString("EvRule")
        .EvMultiplier = mReader.GetDouble("EvMultiplier")
    End If

    If mServerVersion >= ApiServerVersions.MODELS_SUPPORT Then _
        .ModelCode = mReader.GetString("Model Code")
    If mServerVersion >= ApiServerVersions.LAST_LIQUIDITY Then _
        .LastLiquidity = mReader.GetLong("Last Liquidity")
End With

mReader.EndLogMessage

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyExecution lReqId, lContract, lExecution

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getExecutionDataEnd(ByVal pVersion As Long)
Const ProcName As String = "getExecutionDataEnd"
On Error GoTo Err

mReader.BeginLogMessage True

Dim reqId  As Long

reqId = GIdManager.GetCallerId(mReader.GetLong("reqId"), IdTypeExecution)
mReader.EndLogMessage

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.EndExecutions reqId
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getFAData(ByVal pVersion As Long)
Const ProcName As String = "getFAData"
On Error GoTo Err

mReader.BeginLogMessage

Dim DataType As TwsFADataTypes: DataType = mReader.GetLong("Data type")
Dim XMLData As String: XMLData = mReader.GetString("XML data")

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getFundamentalData(ByVal pVersion As Long)
Const ProcName As String = "getFundamentalData"
On Error GoTo Err

mReader.BeginLogMessage

Dim lReqId As Long: lReqId = mReader.GetLong("ReqID")
Dim lData As String: lData = mReader.GetString("Data")

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getHistoricalData(ByVal pVersion As Long)
Const ProcName As String = "getHistoricalData"
On Error GoTo Err

mReader.BeginLogMessage

Static sStartInfo As TwsHistDataStartInfo
Static sBar As TwsBar
Static i As Long

Dim lBar As TwsBar
Dim lReqId As Long
Dim lNotUsed As Boolean

mReader.BeginCheckpointScope
    If mReader.GetLongCheckpointed(lReqId, "Request id") Then
        i = 0
        sBar = lBar
        sStartInfo.requestId = GIdManager.GetCallerId(lReqId, IdTypeHistoricalData)
    End If
    
    If pVersion >= 2 Then
        mReader.GetStringCheckpointed sStartInfo.StartDate, "Start date"
        mReader.GetStringCheckpointed sStartInfo.EndDate, "End date"
    End If
    
    If mReader.GetLongCheckpointed(sStartInfo.BarCount, "Item count") Then
        Debug.Print "Bars retrieved: " & sStartInfo.BarCount
        
        If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.StartHistData sStartInfo.requestId, sStartInfo.StartDate, sStartInfo.EndDate, sStartInfo.BarCount
    End If
    
    For i = i To sStartInfo.BarCount - 1
        mReader.BeginCheckpointScope
            mReader.GetStringCheckpointed sBar.Timestamp, "Bar date"
            mReader.GetDoubleCheckpointed sBar.OpenValue, "Open"
            mReader.GetDoubleCheckpointed sBar.HighValue, "High"
            mReader.GetDoubleCheckpointed sBar.LowValue, "Low"
            mReader.GetDoubleCheckpointed sBar.CloseValue, "Close"
            mReader.GetDecimalCheckpointed sBar.Volume, "Volume"
            mReader.GetDecimalCheckpointed sBar.WAP, "WAP"
            If mServerVersion < ApiServerVersions.SYNT_REALTIME_BARS Then _
                mReader.GetBooleanCheckpointed lNotUsed, "Has gaps"
            If pVersion >= 3 Then mReader.GetLongCheckpointed sBar.TickVolume, "Tick volume"
            
            If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.NotifyBar sStartInfo.requestId, sBar
            
        mReader.EndCheckpointScope
    Next
mReader.EndCheckpointScope

mReader.EndLogMessage

If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.EndHistData sStartInfo.requestId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getManagedAccounts(ByVal pVersion As Long)
Const ProcName As String = "getManagedAccounts"
On Error GoTo Err

mReader.BeginLogMessage

Dim accountsList As String: accountsList = mReader.GetString("Accounts list")

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMarketDataType(ByVal pVersion As Long)
Const ProcName As String = "getMarketDataType"
On Error GoTo Err

mReader.BeginLogMessage

Dim lReqId As Long: lReqId = mReader.GetLong("Ticker id")
Dim lMarketDataType As Long: lMarketDataType = mReader.GetLong("MarketDataType")

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMarketDepth(ByVal pVersion As Long)
Const ProcName As String = "getMarketDepth"
On Error GoTo Err

mReader.BeginLogMessage

Dim id As Long: id = mReader.GetLong("Id")
Dim lPosition As Long: lPosition = mReader.GetLong("Position")
Dim lOperation As TwsDOMOperations:  lOperation = mReader.GetLong("Operation")
Dim lSide As TwsDOMSides: lSide = mReader.GetLong("Side")
Dim lPrice As Double: lPrice = mReader.GetDouble("Price")
Dim lSize As BoxedDecimal: Set lSize = mReader.GetDecimal("Size")

mReader.EndLogMessage

If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.NotifyMarketDepth GIdManager.GetCallerId(id, IdTypeMarketDepth), lPosition, "", lOperation, lSide, lPrice, lSize

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMarketDepthL2(ByVal pVersion As Long)
Const ProcName As String = "getMarketDepthL2"
On Error GoTo Err

mReader.BeginLogMessage

Dim id As Long: id = mReader.GetLong("Id")
Dim lPosition As Long: lPosition = mReader.GetLong("Position")
Dim lMarketMaker As String: lMarketMaker = mReader.GetString("Marketmaker")
Dim lOperation As TwsDOMOperations:  lOperation = mReader.GetLong("Operation")
Dim lSide As TwsDOMSides: lSide = mReader.GetLong("Side")
Dim lPrice As Double: lPrice = mReader.GetDouble("Price")
Dim lSize As BoxedDecimal: Set lSize = mReader.GetDecimal("Size")

mReader.EndLogMessage

If Not mMarketDepthConsumer Is Nothing Then mMarketDepthConsumer.NotifyMarketDepth GIdManager.GetCallerId(id, IdTypeMarketDepth), lPosition, lMarketMaker, lOperation, lSide, lPrice, lSize

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getMarketRule()
Const ProcName As String = "getMarketRule"
On Error GoTo Err

mReader.BeginLogMessage

Dim lMarketRuleId As Long: lMarketRuleId = mReader.GetLong("Market Rule Id")
Dim lIncrementsCount As Long: lIncrementsCount = mReader.GetLong("Increments Count")

ReDim lPriceincrements(lIncrementsCount - 1) As TwsPriceIncrement

Dim i As Long
For i = 0 To lIncrementsCount - 1
    lPriceincrements(i).LowEdge = mReader.GetDouble("Low Edge")
    lPriceincrements(i).Increment = mReader.GetDouble("Increment")
Next

mReader.EndLogMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.NotifyMarketRule lMarketRuleId, lPriceincrements

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getNewsBulletins(ByVal pVersion As Long)
Const ProcName As String = "getNewsBulletins"
On Error GoTo Err

mReader.BeginLogMessage

Dim lMsgId As Long: lMsgId = mReader.GetLong("Id")
Dim msgType As Long: msgType = mReader.GetLong("Type")
Dim newsMessage As String: newsMessage = mReader.GetString("Message")
Dim originatingExch As String: originatingExch = mReader.GetString("Orig Exchange")

mReader.EndLogMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getNextValidId(ByVal pVersion As Long)
Const ProcName As String = "getNextValidId"
On Error GoTo Err

' we force logging of this because its occurrence is important in the
' connection sequence, and it's nice to be able to see explicitly
' in the log when it occurs
mReader.BeginLogMessage True

Dim lNextOrderID As Long
lNextOrderID = mReader.GetLong("Next id")

mReader.EndLogMessage

If lNextOrderID <= GIdManager.GetNextOrderId(pPeekOnly:=True) Then
    ' no need to do anything - either this is a 'new' client id, or
    ' we're in a reconnection during an existing session so we've already
    ' got the right order id
Else
    ' add an arbitrary large number to cater for the case where the last
    ' TWS session was ended untidily, without TWS saving its settings
    lNextOrderID = lNextOrderID + 10000
    GIdManager.SetNextOrderId lNextOrderID
End If




' Note that we do this here rather than in processInitialData().
' This because there are various conditions that can arise after
' the initial API connection sequence is successful but before
' TWS will allow any API requests: eg TWS can object that the
' clientID is already in use if a connection is re-established
' quickly - it seems to take a while for TWS to tidy up after
' a disconnection. Once we get this callback, we know that TWS
' is happy.
gLog "API connection completed", ModuleName, ProcName
mIsTwsConnectedToIBServers = True
RaiseEvent ApiConnectionEstablished(mServerVersion)

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOpenOrder(ByVal pVersion As Long)
Const ProcName As String = "getOpenOrder"
On Error GoTo Err

mReader.BeginLogMessage True

Dim i As Long

Dim lContractSpec As New TwsContractSpecifier
Dim lOrder As New TwsOrder
Dim lOrderState As New TwsOrderState


lOrder.OrderId = mReader.GetLong("Id")

' read contract fields
With lContractSpec
    .ConId = mReader.GetLong("conId")
    .Symbol = mReader.GetString("Symbol")
    .SecType = gTwsSecTypeFromString(mReader.GetString("Sec type"))
    .Expiry = mReader.GetString("Expiry")
    .Strike = mReader.GetDouble("Strike")
    .OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
    .Multiplier = mReader.GetDouble("Multiplier")
    .Exchange = mReader.GetString("Exchange")
    .CurrencyCode = mReader.GetString("Currency")
    .LocalSymbol = mReader.GetString("Local Symbol")
    .TradingClass = mReader.GetString("Trading Class")
End With

' read Order fields
With lOrder
    .Action = gTwsOrderActionFromString(mReader.GetString("Action"))
    .TotalQuantity = mReader.GetDecimal("Quantity")
    .OrderType = gTwsOrderTypeFromString(mReader.GetString("Order type"))
    If .OrderType = TwsOrderTypeNone Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid OrderType"
    .LmtPrice = mReader.GetDoubleMax("Limit price")
    .AuxPrice = mReader.GetDoubleMax("Aux price")
    Dim tifString As String: tifString = mReader.GetString("Tif")
    If UCase$(tifString) = "OPG" Then
        If .OrderType = TwsOrderTypes.TwsOrderTypeMarket Then
            .OrderType = TwsOrderTypes.TwsOrderTypeMarketOnOpen
        ElseIf .OrderType = TwsOrderTypes.TwsOrderTypeLimit Then
            .OrderType = TwsOrderTypes.TwsOrderTypeLimitOnOpen
        End If
    Else
        .Tif = gTwsOrderTIFFromString(tifString)
    End If
    .OcaGroup = mReader.GetString("OCA group")
    .Account = mReader.GetString("Account")
    .OpenClose = mReader.GetString("Open/close")
    .Origin = mReader.GetLong("Origin")
    .OrderRef = mReader.GetString("Order ref")
    .ClientID = mReader.GetLong("Client id")
    
    .PermId = mReader.GetLong("Perm id")
    .OutsideRth = mReader.GetBoolean("Outside RTH")
    .Hidden = mReader.GetBoolean("Hidden")
    .DiscretionaryAmt = mReader.GetDouble("Discr amt")

    .GoodAfterTime = mReader.GetString("Good after Time")
    
    mReader.GetString "Shares allocation"
    
    .FaGroup = mReader.GetString("FA Group")
    .FaMethod = mReader.GetString("FA method")
    .FaPercentage = mReader.GetString("FA Percentage")
    If mServerVersion < ApiServerVersions.FA_PROFILE_DESUPPORT Then mReader.GetString ("FA Profile")
    
    If mServerVersion >= ApiServerVersions.MODELS_SUPPORT Then _
        .ModelCode = mReader.GetString("Model Code")

    .GoodTillDate = mReader.GetString("Good till date")
    
    .Rule80A = mReader.GetString("Rule80A")
    .PercentOffset = mReader.GetDoubleMax("Percent Offset")
    .SettlingFirm = mReader.GetString("Settling Firm")
    
    .ShortSaleSlot = mReader.GetLong("shortSaleSlot")
    .DesignatedLocation = mReader.GetString("designatedLocation")
    .ExemptCode = mReader.GetLong("ExemptCode")
    
    .AuctionStrategy = mReader.GetLong("Auction Strategy")
    
    .StartingPrice = mReader.GetDouble("starting Price")
    .StockRefPrice = mReader.GetDouble("stockRefPrice")
    .Delta = mReader.GetDouble("delta")
    
    .StockRangeLower = mReader.GetDouble("stockRangeLower")
    .StockRangeUpper = mReader.GetDouble("stockRangeUpper")
    
    .DisplaySize = mReader.GetLong("displaySize")
    
    .BlockOrder = mReader.GetBoolean("blockOrder")
    
    .SweepToFill = mReader.GetBoolean("sweepToFill")
    .AllOrNone = mReader.GetBoolean("allOrNone")
    .MinQty = mReader.GetLong("minQty")
    .OcaType = mReader.GetLong("ocaType")

    mReader.GetBoolean "eTradeOnly"      ' desupported
    mReader.GetBoolean "firmQuoteOnly"   ' desupported
    mReader.GetDouble "nbboPriceCap"     ' desupported
    
    .ParentId = mReader.GetLong("ParentId")
    
    .TriggerMethod = mReader.GetLong("TriggerMethod")

    .Volatility = mReader.GetDouble("Volatility")
    .VolatilityType = mReader.GetLong("VolatilityType")
    
    Dim lDeltaNeutralOrderType As String
    lDeltaNeutralOrderType = mReader.GetString("DeltaNeutralOrderType")
    .DeltaNeutralOrderType = gTwsOrderTypeFromString(lDeltaNeutralOrderType)
    .DeltaNeutralAuxPrice = mReader.GetDoubleMax("DeltaNeutralAuxPrice")
    
    If .DeltaNeutralOrderType <> TwsOrderTypeNone Then
        .DeltaNeutralConId = mReader.GetLong("DeltaNeutralConId")
        .DeltaNeutralSettlingFirm = mReader.GetString("DeltaNeutralSettlingFirm")
        .DeltaNeutralClearingAccount = mReader.GetString("DeltaNeutralClearingAccount")
        .DeltaNeutralClearingIntent = mReader.GetString("DeltaNeutralClearingIntent")
        .DeltaNeutralOpenClose = mReader.GetString("DeltaNeutralOpenClose")
        .DeltaNeutralShortSale = mReader.GetBoolean("DeltaNeutralShortSale")
        .DeltaNeutralShortSaleSlot = mReader.GetLong("DeltaNeutralShortSaleSlot")
        .DeltaNeutralDesignatedLocation = mReader.GetString("DeltaNeutralDesignatedLocation")
    End If
    
    .ContinuousUpdate = mReader.GetLong("ContinuousUpdate")
    .ReferencePriceType = mReader.GetLong("ReferencePriceType")
    
    .TrailStopPrice = mReader.GetDoubleMax("TrailStopPrice")
    .TrailingPercent = mReader.GetDoubleMax("TrailingPercent")
    
    .BasisPoints = mReader.GetDoubleMax("BasisPoints")
    .BasisPointsType = mReader.GetLongMax("BasisPointsType")
    
    lContractSpec.ComboLegsDescription = mReader.GetString("ComboLegsDescrip")

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

            lContractSpec.ComboLegs.Add (lComboLeg)
        Next
    End If

    Dim lOrderComboLegsCount As Long: lOrderComboLegsCount = mReader.GetLong("OrderComboLegsCount")
    If lOrderComboLegsCount > 0 Then
        ReDim lOrderComboLegs(lOrderComboLegsCount - 1) As TwsOrderComboLeg
        For i = 0 To lOrderComboLegsCount - 1
            Set lOrderComboLegs(i) = New TwsOrderComboLeg
            lOrderComboLegs(i).Price = mReader.GetDouble("Price" & (i + 1))
        Next
        .ComboLegs = lOrderComboLegs
    End If

    .SmartComboRoutingParams = mReader.GetParameters( _
                                            mReader.GetLong("SmartComboRoutingParamsCount"), _
                                            "SmartComboRoutingParams")

    .ScaleInitLevelSize = mReader.GetLongMax("Scale Init Level Size")
    .ScaleSubsLevelSize = mReader.GetLongMax("Scale Subs Level Size")
    .ScalePriceIncrement = mReader.GetDoubleMax("Scale Price Increment")
    
    If .ScalePriceIncrement > 0# And .ScalePriceIncrement <> MaxDouble Then
        .ScalePriceAdjustValue = mReader.GetDoubleMax("ScalePriceAdjustValue")
        .ScalePriceAdjustInterval = mReader.GetLongMax("ScalePriceAdjustInterval")
        .ScaleProfitOffset = mReader.GetDoubleMax("ScaleProfitOffset")
        .ScaleAutoReset = mReader.GetBoolean("ScaleAutoReset")
        .ScaleInitPosition = mReader.GetLongMax("ScaleInitPosition")
        .ScaleInitFillQty = mReader.GetLongMax("ScaleInitFillQty")
        .ScaleRandomPercent = mReader.GetBoolean("ScaleRandomPercent")
    End If
    
    .HedgeType = gTwsHedgeTypeFromString(mReader.GetString("HedgeType"))
    If .HedgeType <> TwsHedgeTypes.TwsHedgeTypeNone Then .HedgeParam = mReader.GetString("HedgeParam")
    
    .OptOutSmartRouting = mReader.GetBoolean("OptOutSmartRouting")
    
    .ClearingAccount = mReader.GetString("Clearing Account")
    .ClearingIntent = mReader.GetString("Clearing Intent")
    
    .NotHeld = mReader.GetBoolean("Not held")
    
    Dim lDeltaNeutralContract As TwsDeltaNeutralContract
    If CBool(mReader.GetLong("DeltaNeutral")) Then
        lDeltaNeutralContract.ConId = mReader.GetLong("Delta neutral ConId")
        lDeltaNeutralContract.Delta = mReader.GetDouble("Delta neutral Delta")
        lDeltaNeutralContract.Price = mReader.GetDouble("Delta neutral Price")
        lContractSpec.DeltaNeutralContract = lDeltaNeutralContract
    End If
    
    .AlgoStrategy = mReader.GetString("Algo strategy")
    If .AlgoStrategy <> "" Then
        Set .AlgoParams = mReader.GetParameters( _
                                        mReader.GetLong("Algo params count"), _
                                        "Algo params")
    End If
    
    .Solicited = mReader.GetBoolean("Solicited")

    .WhatIf = mReader.GetBoolean("What If")

    With lOrderState
        .Status = mReader.GetString("Status")
        If mServerVersion >= ApiServerVersions.WHAT_IF_EXT_FIELDS Then
            .InitMarginBefore = mReader.GetString("Init Margin before")
            .MaintMarginBefore = mReader.GetString("Maint Margin before")
            .EquityWithLoanBefore = mReader.GetString("Equity With Loan before")
            .InitMarginChange = mReader.GetString("Init Margin change")
            .MaintMarginChange = mReader.GetString("Maint Margin change")
            .EquityWithLoanChange = mReader.GetString("Equity With Loan change")
        End If
        
        .InitMarginAfter = mReader.GetString("Init Margin after")
        .MaintMarginAfter = mReader.GetString("Maint Margin after")
        .EquityWithLoanAfter = mReader.GetString("Equity With Loan after")
        
        .Commission = mReader.GetDoubleMax("Commission")
        .MinCommission = mReader.GetDoubleMax("Min Commission")
        .MaxCommission = mReader.GetDoubleMax("Max Commission")
        .CommissionCurrency = mReader.GetString("Commission Currency")
        .WarningText = mReader.GetString("Warning Text")
    End With

    .RandomizeSize = mReader.GetBoolean("Randomize Size")
    .RandomizePrice = mReader.GetBoolean("Randomize Price")

    If mServerVersion >= ApiServerVersions.PEGGED_TO_BENCHMARK Then
        If .OrderType = TwsOrderTypePeggedToBenchmark Then
            .ReferenceContractId = mReader.GetLong("ReferenceContractId")
            .IsPeggedChangeAmountDecrease = mReader.GetBoolean("IsPeggedChangeAmountDecrease")
            .PeggedChangeAmount = mReader.GetDoubleMax("PeggedChangeAmount")
            .ReferenceChangeAmount = mReader.GetDoubleMax("ReferenceChangeAmount")
            .ReferenceExchange = mReader.GetString("ReferenceExchange")
        End If

        Dim nConditions As Long: nConditions = mReader.GetLong("Conditions count")
        
        Assert nConditions = 0, "Conditions not supported"
'        If nConditions > 0 Then
'            For i = 1 To nConditions
'                Dim orderConditionType = CType(mReader.GetInt("OrderConditionType"), OrderConditionType)
'                Dim condition = OrderCondition.Create(orderConditionType)
'
'                condition.Deserialize(_Reader)
'                .Conditions.Add (condition)
'
'                .ConditionsIgnoreRth = mReader.GetBoolFromInt("Conditions Ignore Rth")
'                .ConditionsCancelOrder = mReader.GetBoolFromInt("Conditions Cancel Order")
'            Next
'        End If

        .AdjustedOrderType = gTwsOrderTypeFromString(mReader.GetString("AdjustedOrderType"))
        .TriggerPrice = mReader.GetDoubleMax("TriggerPrice")
        .TrailStopPrice = mReader.GetDoubleMax("TrailStopPrice")
        .LmtPriceOffset = mReader.GetDoubleMax("LmtPriceOffset")
        .AdjustedStopPrice = mReader.GetDoubleMax("AdjustedStopPrice")
        .AdjustedStopLimitPrice = mReader.GetDoubleMax("AdjustedStopLimitPrice")
        .AdjustedTrailingAmount = mReader.GetDoubleMax("AdjustedTrailingAmount")
        .AdjustableTrailingUnit = mReader.GetLong("AdjustableTrailingUnit")
    End If

    If mServerVersion >= ApiServerVersions.SOFT_DOLLAR_TIER Then
        .Tier.Name = mReader.GetString("Name")
        .Tier.Value = mReader.GetString("Value")
        .Tier.DisplayName = mReader.GetString("Display Name")
    End If

    If mServerVersion >= ApiServerVersions.CASH_QTY Then
        .CashQty = mReader.GetDoubleMax("Cash Qty")
    End If

    If mServerVersion >= ApiServerVersions.AUTO_PRICE_FOR_HEDGE Then
        .DontUseAutoPriceForHedge = mReader.GetBoolean("Dont Use Auto Price For Hedge")
    End If

    If mServerVersion >= ApiServerVersions.ORDER_CONTAINER Then
        .IsOmsContainer = mReader.GetBoolean("IsOmsContainer")
    End If

    If mServerVersion >= ApiServerVersions.D_PEG_ORDERS Then
        .DiscretionaryUpToLimitPrice = mReader.GetBoolean("DiscretionaryUpToLimitPrice")
    End If
    If mServerVersion >= ApiServerVersions.PRICE_MGMT_ALGO Then
        .UsePriceMgmtAlgo = mReader.GetBoolean("UsePriceMgmtAlgo")
    End If
    If mServerVersion >= ApiServerVersions.Duration Then
        .Duration = mReader.GetLongMax("Duration")
    End If
    If mServerVersion >= ApiServerVersions.POST_TO_ATS Then
        .PostToAts = mReader.GetLongMax("PostToAts")
    End If
    If mServerVersion >= ApiServerVersions.AUTO_CANCEL_PARENT Then
        .AutoCancelParent = mReader.GetBoolean("AutoCancelParent")
    End If

    If mServerVersion >= ApiServerVersions.PEGBEST_PEGMID_OFFSETS Then
        .MinimumTradeQuantity = mReader.GetLongMax("MinimumTradeQuantity")
        .MinimumCompeteSize = mReader.GetLongMax("MinimumCompeteSize")
        .CompeteAgainstBestOffset = mReader.GetDoubleMax("CompeteAgainstBestOffset")
        .MidOffsetAtWhole = mReader.GetDoubleMax("MidOffsetAtWhole")
        .MidOffsetAtHalf = mReader.GetDoubleMax("MidOffsetAtHalf")
    End If

End With

mReader.EndLogMessage

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyOpenOrder lOrder.OrderId, lContractSpec, lOrder, lOrderState

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOpenOrderEnd(ByVal pVersion As Long)
Const ProcName As String = "getOpenOrderEnd"
On Error GoTo Err

mReader.BeginLogMessage True
mReader.EndLogMessage

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.EndOpenOrders
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOptionParameter()
Const ProcName As String = "getOptionParameter"
On Error GoTo Err

mReader.BeginLogMessage

Dim lRequestId As Long: lRequestId = GIdManager.GetCallerId(mReader.GetLong("Request id"), IdTypeContractData)
Dim lExchange  As String: lExchange = mReader.GetString("Exchange")
Dim lUnderlyingConId As Long: lUnderlyingConId = mReader.GetLong("Und con id")
Dim lTradingClass As String: lTradingClass = mReader.GetString("Trading Class")
Dim lMult As String: lMult = mReader.GetString("Multiplier")
Dim lMultiplier As Double: lMultiplier = IIf(lMult = "", 1, CDbl(lMult))

Dim lExpsCount: lExpsCount = mReader.GetLong("Expirations Count")
ReDim lExpirations(lExpsCount - 1) As String

Dim i As Long
For i = 0 To lExpsCount - 1
    lExpirations(i) = mReader.GetString("Expiration")
Next

Dim lStrikesCount: lStrikesCount = mReader.GetLong("Strikes Count")
ReDim lStrikes(lStrikesCount - 1) As Double

For i = 0 To lStrikesCount - 1
    lStrikes(i) = mReader.GetDouble("Strike")
Next

mReader.EndLogMessage

If Not mContractDetailsConsumer Is Nothing Then
    mContractDetailsConsumer.NotifyOptionParameter lRequestId, _
                                                    lExchange, _
                                                    lUnderlyingConId, _
                                                    lTradingClass, _
                                                    lMultiplier, _
                                                    lExpirations, _
                                                    lStrikes
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOptionParameterEnd()
Const ProcName As String = "getOptionParameterEnd"
On Error GoTo Err

mReader.BeginLogMessage

Dim lRequestId  As Long
lRequestId = GIdManager.GetCallerId(mReader.GetLong("Request id"), IdTypeContractData)

mReader.EndLogMessage

If Not mContractDetailsConsumer Is Nothing Then mContractDetailsConsumer.EndOptionParameters lRequestId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getOrderStatus(ByVal pVersion As Long)
Const ProcName As String = "getOrderStatus"
On Error GoTo Err

mReader.BeginLogMessage True

' NB: lAvgFillPrice and lLastFillPrice ought to be subjected to adjustPrice,
' but we can't call that here because we don't know the currency code. But not really
' an issue since we are not likely to encounter a Tws with server pVersion < 23 !

Dim lorderId As Long: lorderId = mReader.GetLong("id")
Dim lStatus As String: lStatus = mReader.GetString("Status")
Dim lFilled As BoxedDecimal: Set lFilled = mReader.GetDecimal("Filled")
Dim lRemaining As BoxedDecimal: Set lRemaining = mReader.GetDecimal("Remaining")
Dim lAvgFillPrice As Double:  lAvgFillPrice = mReader.GetDouble("Avg fill price")
Dim lPermId As Long: lPermId = mReader.GetLong("Perm id")
Dim lParentId As Long: lParentId = mReader.GetLong("Parent id")
Dim lLastFillPrice As Double: lLastFillPrice = mReader.GetDouble("Last fill price")
Dim lClientID As Long: lClientID = mReader.GetLong("Client id")
Dim lWhyHeld As String: lWhyHeld = mReader.GetString("Why held")
Dim lmarketCapPrice As Double: lmarketCapPrice = MaxDouble

If mServerVersion >= ApiServerVersions.MARKET_CAP_PRICE Then
    lmarketCapPrice = mReader.GetDouble("Market cap price")
End If

mReader.EndLogMessage

If Not mOrderInfoConsumer Is Nothing Then mOrderInfoConsumer.NotifyOrderStatus lorderId, lStatus, lFilled, lRemaining, lAvgFillPrice, lPermId, lParentId, lLastFillPrice, lClientID, lWhyHeld, lmarketCapPrice

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getPortfolioValue(ByVal pVersion As Long)
Const ProcName As String = "getPortfolioValue"
On Error GoTo Err

mReader.BeginLogMessage

Dim lContract As New TwsContractSpecifier

lContract.ConId = mReader.GetLong("Contract id")

lContract.Symbol = mReader.GetString("Symbol")
lContract.SecType = gTwsSecTypeFromString(mReader.GetString("Sec type"))
lContract.Expiry = mReader.GetString("Expiry")
lContract.Strike = mReader.GetDouble("Strike")
lContract.OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))

lContract.Multiplier = mReader.GetDouble("Multiplier")
lContract.PrimaryExch = mReader.GetString("Primary exchg")

lContract.CurrencyCode = mReader.GetString("Currency")

lContract.LocalSymbol = mReader.GetString("Local Symbol")
lContract.TradingClass = mReader.GetString("Trading Class")

Dim lPosition As BoxedDecimal
Set lPosition = mReader.GetDecimal("Position")

Dim lMarketPrice As Double: lMarketPrice = mReader.GetDouble("Market price")
Dim lMarketValue As Double: lMarketValue = mReader.GetDouble("Market Value")
Dim lAverageCost As Double: lAverageCost = mReader.GetDouble("Avg cost")
Dim lUnrealizedPNL As Double: lUnrealizedPNL = mReader.GetDouble("Unrealized PNL")
Dim lRealizedPNL As Double: lRealizedPNL = mReader.GetDouble("Realized PNL")
Dim lAccountName As String: lAccountName = mReader.GetString("Account Name")

mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyPortfolioUpdate lContract, lPosition, lMarketPrice, lMarketValue, lAverageCost, lUnrealizedPNL, lRealizedPNL, lAccountName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getPosition(ByVal pVersion As Long)
Const ProcName As String = "getPosition"
On Error GoTo Err

mReader.BeginLogMessage True

Dim lAccount As String: lAccount = mReader.GetString("Account")
Dim lContract As New TwsContractSpecifier
lContract.ConId = mReader.GetLong("Contract id")
lContract.Symbol = mReader.GetString("Symbol")
lContract.SecType = gTwsSecTypeFromString(mReader.GetString("SecType"))
lContract.Expiry = mReader.GetString("Expiry")
lContract.Strike = mReader.GetDouble("Strike")
lContract.OptRight = gTwsOptionRightFromString(mReader.GetString("Right"))
lContract.Multiplier = mReader.GetDouble("Multiplier")
lContract.Exchange = mReader.GetString("Exchange")
lContract.CurrencyCode = mReader.GetString("Currency")
lContract.LocalSymbol = mReader.GetString("LocalSymbol")
lContract.TradingClass = mReader.GetString("Trading Class")

Dim lPosition As BoxedDecimal
Set lPosition = mReader.GetDecimal("Position")

Dim lAverageCost As Double
If pVersion >= 3 Then lAverageCost = mReader.GetDouble("Average Cost")

mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.NotifyPosition lAccount, lContract, lPosition, lAverageCost

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getPositionEnd(ByVal pVersion As Long)
Const ProcName As String = "getPositionEnd"
On Error GoTo Err

mReader.BeginLogMessage True
mReader.EndLogMessage

If Not mAccountDataConsumer Is Nothing Then mAccountDataConsumer.EndPosition

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getRealtimeBars(ByVal pVersion As Long)
Const ProcName As String = "getRealtimeBars"
On Error GoTo Err

mReader.BeginLogMessage

Dim lReqId As Long: lReqId = mReader.GetLong("ReqId")

Dim lBar As TwsBar

With lBar
    .Timestamp = gGetDateFromUnixSystemTime(mReader.GetDouble("Timestamp"))
    .OpenValue = mReader.GetDouble("Open")
    .HighValue = mReader.GetDouble("High")
    .LowValue = mReader.GetDouble("Low")
    .CloseValue = mReader.GetDouble("Close")
    Set .Volume = mReader.GetDecimal("Volume")
    Set .WAP = mReader.GetDecimal("Wap")
    .TickVolume = mReader.GetLong("TickVolume")
End With

mReader.EndLogMessage

If Not mHistDataConsumer Is Nothing Then mHistDataConsumer.NotifyRealtimeBar GIdManager.GetCallerId(lReqId, IdTypeHistoricalData), lBar

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getScannerData(ByVal pVersion As Long)
Const ProcName As String = "getScannerData"
On Error GoTo Err

mReader.BeginLogMessage

Static sRequestId As Long
Static sNumberOfElements As Long

Static sContractSpec As TwsContractSpecifier
Static sContract As TwsContract

Static sRank As Long
Static sDistance As String
Static sBenchmark As String
Static sProjection As String
Static sLegs As String

mReader.BeginCheckpointScope
    Dim lRequestId As Long
    mReader.GetLongCheckpointed lRequestId, "Request Id"
    sRequestId = GIdManager.GetCallerId(lRequestId, IdTypeScanner)
    mReader.GetLongCheckpointed sNumberOfElements, "Number of elements"
    
    Dim i As Long
    For i = 0 To sNumberOfElements - 1
        mReader.BeginCheckpointScope
            Dim lLongValue As Long
            Dim lStringValue As String
            Dim lDoubleValue As Double
            
            If mReader.GetLongCheckpointed(sRank, "Rank") Then
                Set sContractSpec = New TwsContractSpecifier
                Set sContract = New TwsContract
                sContract.Specifier = sContractSpec
            End If
            
            If (pVersion >= 3) Then If mReader.GetLongCheckpointed(lLongValue, "ConId") Then sContractSpec.ConId = lLongValue
            
            If mReader.GetStringCheckpointed(lStringValue, "Symbol") Then sContractSpec.Symbol = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "Sec type") Then sContractSpec.SecType = gTwsSecTypeFromString(lStringValue)
            If mReader.GetStringCheckpointed(lStringValue, "Expiry") Then sContractSpec.Expiry = lStringValue
            If mReader.GetDoubleCheckpointed(lDoubleValue, "Strike") Then sContractSpec.Strike = lDoubleValue
            If mReader.GetStringCheckpointed(lStringValue, "Right") Then sContractSpec.OptRight = gTwsOptionRightFromString(lStringValue)
            If mReader.GetStringCheckpointed(lStringValue, "Exchange") Then sContractSpec.Exchange = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "Currency") Then sContractSpec.CurrencyCode = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "Local Symbol") Then sContractSpec.LocalSymbol = lStringValue
            
            If mReader.GetStringCheckpointed(lStringValue, "Market name") Then sContract.MarketName = lStringValue
            If mReader.GetStringCheckpointed(lStringValue, "TradingClass") Then sContractSpec.TradingClass = lStringValue
            
            mReader.GetStringCheckpointed sDistance, "Distance"
            mReader.GetStringCheckpointed sBenchmark, "Benchmark"
            mReader.GetStringCheckpointed sProjection, "Projection"
            If (pVersion >= 2) Then mReader.GetStringCheckpointed sLegs, "Legs Str"
        
            If Not mScannerDataConsumer Is Nothing Then mScannerDataConsumer.NotifyScannerData sRequestId, sRank, sContract, sDistance, sBenchmark, sProjection, sLegs
        mReader.EndCheckpointScope
    Next
mReader.EndCheckpointScope

mReader.EndLogMessage

If Not mScannerDataConsumer Is Nothing Then mScannerDataConsumer.EndScannerData sRequestId

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getScannerParameters(ByVal pVersion As Long)
Const ProcName As String = "getScannerParameters"
On Error GoTo Err

mReader.BeginLogMessage

Dim xml  As String: xml = mReader.GetString("XML")

mReader.EndLogMessage

If Not mScannerDataConsumer Is Nothing Then mScannerDataConsumer.NotifyScannerParameters xml

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickEFP(ByVal pVersion As Long)
Const ProcName As String = "getTickEFP"
On Error GoTo Err

mReader.BeginLogMessage

Dim lTickerId As Long: lTickerId = mReader.GetLong("Ticker Id")
Dim lTickType As Long: lTickType = mReader.GetLong("Tick Type")
Dim lBasisPoints  As Double: lBasisPoints = mReader.GetDouble("Basis Points")
Dim lFormattedBasisPoints As String: lFormattedBasisPoints = mReader.GetString("Formatted BasisP oints")
Dim lImpliedFuturesPrice As Double: lImpliedFuturesPrice = mReader.GetDouble("lImplied Futures Price")
Dim lHoldDays As Long: lHoldDays = mReader.GetLong("Hold Days")
Dim lFutureExpiry As String: lFutureExpiry = mReader.GetString("Future Expiry")
Dim lDividendImpact As Double: lDividendImpact = mReader.GetDouble("Dividend Impact")
Dim lDividendsToExpiry As Double: lDividendsToExpiry = mReader.GetDouble("Dividends To Expiry")

mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickEFP lTickerId, lTickType, lBasisPoints, lFormattedBasisPoints, lImpliedFuturesPrice, lHoldDays, lFutureExpiry, lDividendImpact, lDividendsToExpiry

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickGeneric(ByVal pVersion As Long)
Const ProcName As String = "getTickGeneric"
On Error GoTo Err

mReader.BeginLogMessage

Dim lTickerId As Long: lTickerId = mReader.GetLong("Ticker id")
Dim lTickType As Long: lTickType = mReader.GetLong("Tick type")
Dim lValue As Double: lValue = mReader.GetDouble("Value")

mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickGeneric GIdManager.GetCallerId(lTickerId, IdTypeMarketData), lTickType, lValue

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickOptionComputation(ByVal pVersion As Long)
Const ProcName As String = "getTickOptionComputation"
On Error GoTo Err

mReader.BeginLogMessage

Dim lTickerId As Long: lTickerId = mReader.GetLong("Ticker id")
Dim lTickType As Long: lTickType = mReader.GetLong("Tick type")

Dim lTickAttrib As Long
If mServerVersion >= ApiServerVersions.PRICE_BASED_VOLATILITY Then lTickAttrib = mReader.GetLong("tickAttrib")

Dim lImpliedVol As Double: lImpliedVol = mReader.GetDouble("impliedVol")

If lImpliedVol = -1 Then
    ' -1 is the "not yet computed" indicator
    lImpliedVol = MaxDouble
End If

Dim lDelta As Double: lDelta = mReader.GetDouble("delta")
If lDelta = -2 Then
    ' -2 is the "not yet computed" indicator
    lDelta = MaxDouble
End If

Dim lOptPrice As Double: lOptPrice = MaxDouble
Dim lPvDividend As Double: lPvDividend = MaxDouble
Dim lGamma As Double: lGamma = MaxDouble
Dim lVega As Double: lVega = MaxDouble
Dim lTheta As Double: lTheta = MaxDouble
Dim lUndPrice As Double: lUndPrice = MaxDouble

If pVersion >= 6 Or lTickType = TwsTickTypes.TwsTickModelOption Or lTickType = TwsTickTypes.TwsTickDelayedModelOption Then     ' introduced in pVersion == 5
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
    If lGamma = -2 Then ' -2 is the "not yet computed" indicator
        lGamma = MaxDouble
    End If
    lVega = mReader.GetDouble("Vega")
    If lVega = -2 Then ' -2 is the "not yet computed" indicator
        lVega = MaxDouble
    End If
    lTheta = mReader.GetDouble("theta")
    If lTheta = -2 Then ' -2 is the "not yet computed" indicator
        lTheta = MaxDouble
    End If
    lUndPrice = mReader.GetDouble("undPrice")
    If lUndPrice = -1 Then ' -1 is the "not yet computed" indicator
        lUndPrice = MaxDouble
    End If
End If

mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickOptionComputation GIdManager.GetCallerId(lTickerId, IdTypeMarketData), lTickType, lImpliedVol, lDelta, lOptPrice, lPvDividend, lGamma, lVega, lTheta, lUndPrice

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickPrice(ByVal pVersion As Long)
Const ProcName As String = "getTickPrice"
On Error GoTo Err

mReader.BeginLogMessage

Dim lTickerId As Long: lTickerId = mReader.GetLong("Ticker id")
Dim lTickType As Long: lTickType = mReader.GetLong("Tick type")
Dim lPrice As Double: lPrice = mReader.GetDouble("Price")

Dim lSize As BoxedDecimal
If pVersion >= 2 Then Set lSize = mReader.GetDecimal("Size")

Dim lAttributes As TwsTickAttributes

If pVersion >= 3 Then
    Const CanAutoExecute = &H1&
    Const PastLimit = &H2&
    Const PreOpen = &H4&

    Dim lAttrMask As Long: lAttrMask = mReader.GetLong("Attr Mask")
    lAttributes.CanAutoExecute = (lAttrMask = 1)

    If mServerVersion >= ApiServerVersions.PAST_LIMIT Then
        lAttributes.CanAutoExecute = (lAttrMask And CanAutoExecute) > 0
        lAttributes.PastLimit = (lAttrMask And PastLimit) > 0
        If mServerVersion >= ApiServerVersions.PRE_OPEN_BID_ASK Then
            lAttributes.PreOpen = (lAttrMask And PreOpen) > 0
        End If
    End If
End If

mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickPrice GIdManager.GetCallerId(lTickerId, IdTypeMarketData), lTickType, lPrice, lSize, lAttributes

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickRequestParams()
Const ProcName As String = "getTickRequestParams"
On Error GoTo Err

mReader.BeginLogMessage

Dim lTickerId As Long: lTickerId = mReader.GetLong("Ticker Id")
Dim lMinTick As Double: lMinTick = mReader.GetDouble("Min tick")
Dim lBboExchange As String: lBboExchange = mReader.GetString("BBO Exchange")
Dim lSnapshotPermissions As Long: lSnapshotPermissions = mReader.GetLong("Snapshot permissions")

mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickRequestParams GIdManager.GetCallerId(lTickerId, IdTypeMarketData), lMinTick, lBboExchange, lSnapshotPermissions

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub


Private Sub getTickSize(ByVal pVersion As Long)
Const ProcName As String = "getTickSize"
On Error GoTo Err

mReader.BeginLogMessage

Dim lTickType As Long
Dim lTickerId As Long
Dim lSize As BoxedDecimal

lTickerId = mReader.GetLong("Ticker id")
lTickType = mReader.GetLong("Tick type")
Set lSize = mReader.GetDecimal("Size")

mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickSize GIdManager.GetCallerId(lTickerId, IdTypeMarketData), lTickType, lSize

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName

End Sub

Private Sub getTickSnapshotEnd(ByVal pVersion As Long)
Const ProcName As String = "getTickSnapshotEnd"
On Error GoTo Err

mReader.BeginLogMessage

Dim lReqId As Long: lReqId = mReader.GetLong("Ticker id")
mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.EndTickSnapshot GIdManager.GetCallerId(lReqId, IdTypeMarketData)

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub getTickString(ByVal pVersion As Long)
Const ProcName As String = "getTickString"
On Error GoTo Err

mReader.BeginLogMessage

Dim lTickerId As Long: lTickerId = mReader.GetLong("Ticker id")
Dim lTickType As Long: lTickType = mReader.GetLong("Tick type")
Dim lValue As String: lValue = mReader.GetString("Value")

mReader.EndLogMessage

If Not mMarketDataConsumer Is Nothing Then mMarketDataConsumer.NotifyTickString GIdManager.GetCallerId(lTickerId, IdTypeMarketData), lTickType, lValue

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





