Attribute VB_Name = "GIBEnhancedApi"
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
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GIBEnhancedAPI"

Private Const ExchangeCBOT                      As String = "CBOT"
Private Const ExchangeCME                       As String = "CME"
Private Const ExchangeECBOT                     As String = "ECBOT"
Private Const ExchangeGlobex                    As String = "GLOBEX"

Private Const ExchangeSmart                     As String = "SMART"
Private Const ExchangeSmartAUS                  As String = "SMARTAUS"
Private Const ExchangeSmartCAN                  As String = "SMARTCAN"
Private Const ExchangeSmartEUR                  As String = "SMARTEUR"
Private Const ExchangeSmartNASDAQ               As String = "SMARTNASDAQ"
Private Const ExchangeSmartNYSE                 As String = "SMARTNYSE"
Private Const ExchangeSmartUK                   As String = "SMARTUK"
Private Const ExchangeSmartUS                   As String = "SMARTUS"
Private Const ExchangeSmartQualified            As String = "SMART-"

Private Const PrimaryExchangeARCA               As String = "ARCA"
Private Const PrimaryExchangeASX                As String = "ASX"
Private Const PrimaryExchangeEBS                As String = "EBS"
Private Const PrimaryExchangeFWB                As String = "FWB"
Private Const PrimaryExchangeIBIS               As String = "IBIS"
Private Const PrimaryExchangeLSE                As String = "LSE"
Private Const PrimaryExchangeSWB                As String = "SWB"
Private Const PrimaryExchangeNASDAQ             As String = "NASDAQ"
Private Const PrimaryExchangeNYSE               As String = "NYSE"
Private Const PrimaryExchangeVENTURE            As String = "VENTURE"

Public Const ProviderPropertyContractID         As String = "Contract id"
Public Const ProviderPropertyOCAGroup           As String = "OCA group"
Public Const ProviderPropertyPriceMagnifier     As String = "PriceMagnifier"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

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

'@================================================================================
' Methods
'@================================================================================

Public Function ContractSpecToTwsContractSpec( _
                ByVal pContractSpecifier As IContractSpecifier) As TwsContractSpecifier
Const ProcName As String = "ContractSpecToTwsContractSpec"
On Error GoTo Err

Dim lComboLeg As comboLeg
Dim lTwsComboLeg As TwsComboLeg

Dim lTwsContractSpec As New TwsContractSpecifier

With lTwsContractSpec
    If Not pContractSpecifier.ProviderProperties Is Nothing Then
        .ConId = pContractSpecifier.ProviderProperties.GetParameterValue(ProviderPropertyContractID, "0")
    End If
    .CurrencyCode = pContractSpecifier.CurrencyCode
    Dim lExchange As String: lExchange = UCase$(pContractSpecifier.Exchange)
    If lExchange = ExchangeSmartAUS Then
        .PrimaryExch = PrimaryExchangeASX
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeSmartCAN Then
        .PrimaryExch = PrimaryExchangeVENTURE
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeSmartUK Then
        .PrimaryExch = PrimaryExchangeLSE
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeSmartNASDAQ Then
        .PrimaryExch = PrimaryExchangeNASDAQ
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeSmartNYSE Then
        .PrimaryExch = PrimaryExchangeNYSE
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeSmartUS Then
        If pContractSpecifier.Sectype <> SecTypeOption Then
            .PrimaryExch = PrimaryExchangeARCA
        End If
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeSmartEUR Then
        .PrimaryExch = PrimaryExchangeIBIS
        .Exchange = ExchangeSmart
    ElseIf InStr(1, lExchange, ExchangeSmartQualified) = 1 Then
        .PrimaryExch = Right$(lExchange, Len(lExchange) - Len(ExchangeSmartQualified))
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeGlobex Then
        .Exchange = ExchangeCME
    ElseIf lExchange = ExchangeECBOT Then
        .Exchange = ExchangeCBOT
    Else
        .Exchange = lExchange
    End If
    .Expiry = IIf(Len(pContractSpecifier.Expiry) >= 6, pContractSpecifier.Expiry, "")
    .LocalSymbol = pContractSpecifier.LocalSymbol
    .Multiplier = pContractSpecifier.Multiplier
    If pContractSpecifier.Sectype = SecTypeStock Then
        .Multiplier = 0
    Else
        Dim lPriceMagnifier As Long
        If Not pContractSpecifier.ProviderProperties Is Nothing Then
            lPriceMagnifier = pContractSpecifier.ProviderProperties.GetParameterValue(ProviderPropertyPriceMagnifier, "0")
            If lPriceMagnifier <> 0 Then .Multiplier = pContractSpecifier.Multiplier / lPriceMagnifier
        ElseIf UCase$(.CurrencyCode) = "GBP" Then
            .Multiplier = .Multiplier / 100
        End If
    End If
    .OptRight = OptionRightToTwsOptRight(pContractSpecifier.Right)
    .Sectype = SecTypeToTwsSecType(pContractSpecifier.Sectype)
    .Strike = pContractSpecifier.Strike
    .Symbol = pContractSpecifier.Symbol
    .TradingClass = pContractSpecifier.TradingClass
    If Not pContractSpecifier.ComboLegs Is Nothing Then
        For Each lComboLeg In pContractSpecifier.ComboLegs
            Set lTwsComboLeg = New TwsComboLeg
            With lTwsComboLeg
                .Action = IIf(lComboLeg.IsBuyLeg, TwsOrderActions.TwsOrderActionBuy, TwsOrderActions.TwsOrderActionSell)
                Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Combo contracts not supported"
                ' Need to fix this: the problem is that we need to do a contract details
                ' request to discover the contract id for the combo leg
            End With
        Next
    End If
End With

Set ContractSpecToTwsContractSpec = lTwsContractSpec

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractFutureToTwsContractFuture( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractFuture As IFuture, _
                ByVal pContractCache As ContractCache) As IFuture
Const ProcName As String = "ContractFutureToTwsContractFuture"
On Error GoTo Err

Dim lFutureBuilder As New TwsContractDtlsFutBldr
lFutureBuilder.Initialise pContractRequester, pContractFuture, pContractCache
Set ContractFutureToTwsContractFuture = lFutureBuilder.Future

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function ContractToTwsContract(ByVal pContract As IContract) As TwsContract
Const ProcName As String = "ContractToTwsContract"
On Error GoTo Err

Assert pContract.Specifier.Sectype <> SecTypeCombo, "Combo contracts not supported", ErrorCodes.ErrUnsupportedOperationException

Dim lContract As New TwsContract
Dim lContractSpec As TwsContractSpecifier
Set lContractSpec = ContractSpecToTwsContractSpec(pContract.Specifier)

With lContract
    .Specifier = lContractSpec
    .MinTick = pContract.TickSize
    .TimeZoneId = StandardTimezoneNameToTwsTimeZoneName(pContract.TimezoneName)
End With

Set ContractToTwsContract = lContract

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContracts( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pListener As IContractFetchListener, _
                ByVal pCookie As Variant, _
                ByVal pReturnTwsContracts As Boolean) As IFuture
Const ProcName As String = "FetchContracts"
On Error GoTo Err

If GIBEnhApi.Logger.IsLoggable(LogLevelDetail) Then GIBEnhApi.Log "Fetching " & _
                                                IIf(pReturnTwsContracts, "TWS", "") & _
                                                " contract details for", _
                                                ModuleName, ProcName, pContractSpecifier.ToString, LogLevelMediumDetail

Dim lFetcher As New ContractsRequestManager
Set FetchContracts = lFetcher.Fetch( _
                                pContractRequester, _
                                pContractCache, _
                                pContractSpecifier, _
                                pListener, _
                                pCookie, _
                                pReturnTwsContracts)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContractsSorted( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                ByVal pSortDescending As Boolean, _
                ByVal pCookie As Variant, _
                ByVal pReturnTwsContracts As Boolean) As IFuture
Const ProcName As String = "FetchContractsSorted"
On Error GoTo Err

If GIBEnhApi.Logger.IsLoggable(LogLevelDetail) Then GIBEnhApi.Log "Fetching sorted contract details for", ModuleName, ProcName, pContractSpecifier.ToString, LogLevelDetail

Dim lFetcher As New ContractsRequestManager
Set FetchContractsSorted = lFetcher.FetchSorted( _
                                pContractRequester, _
                                pContractCache, _
                                pContractSpecifier, _
                                pSortkeys, _
                                pSortDescending, _
                                pCookie, _
                                pReturnTwsContracts)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function GenerateTwsContractKey( _
                ByVal pContract As TwsContract) As String
Const ProcName As String = "GenerateTwsContractKey"
On Error GoTo Err

Dim lSpec As IContractSpecifier
Set lSpec = GIBEnhancedApi.TwsContractSpecToContractSpecifier(pContract.Specifier, pContract.PriceMagnifier)

GenerateTwsContractKey = lSpec.key

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetClient( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                Optional ByVal pSessionID As String, _
                Optional ByVal pConnectionRetryIntervalSecs As Long = 60, _
                Optional ByVal pLogApiMessages As ApiMessageLoggingOptions = ApiMessageLoggingOptionDefault, _
                Optional ByVal pLogRawApiMessages As ApiMessageLoggingOptions = ApiMessageLoggingOptionDefault, _
                Optional ByVal pLogApiMessageStats As Boolean = False, _
                Optional ByVal pDeferConnection As Boolean, _
                Optional ByVal pConnectionStateListener As ITwsConnectionStateListener, _
                Optional ByVal pProgramErrorHandler As IProgramErrorListener, _
                Optional ByVal pApiErrorListener As IErrorListener, _
                Optional ByVal pApiNotificationListener As INotificationListener) As Client
Const ProcName As String = "GetClient"
On Error GoTo Err

If pSessionID = "" Then pSessionID = GenerateGUIDString

Set GetClient = GClient.GetClient(pSessionID, _
                            pServer, _
                            pPort, _
                            pClientId, _
                            pConnectionRetryIntervalSecs, _
                            pLogApiMessages, _
                            pLogRawApiMessages, _
                            pLogApiMessageStats, _
                            pDeferConnection, _
                            pConnectionStateListener, _
                            pProgramErrorHandler, _
                            pApiErrorListener, _
                            pApiNotificationListener)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetSessionTimes(ByVal pSessionTimesString As String) As SessionTimes
Const ProcName As String = "GetSessionTimes"
On Error GoTo Err

If Len(pSessionTimesString) = 0 Then Exit Function

Const SessionTimes As String = "^(?:(\d{8}):(\d{2})(\d{2})-(\d{8}):(\d{2})(\d{2}))|(?:\d{8}:(CLOSED))$"
Dim lRegExp As RegExp: Set lRegExp = GIBEnhApi.RegExpProcessor
lRegExp.Pattern = SessionTimes

Dim lSessionTimesAr() As String: lSessionTimesAr = Split(pSessionTimesString, ";")

Dim lSessionTimes As SessionTimes
Dim lNumberOfSessionTimesProcessed As Long
Dim lSessionDate As Date
Dim i As Long
For i = 0 To UBound(lSessionTimesAr)
    Dim l As Variant: l = lSessionTimesAr(i)
    If getSessionTimesForDay(lRegExp, l, lSessionTimes, lSessionDate) Then
        If GetSessionTimes.StartTime = 0 Or lSessionTimes.StartTime < GetSessionTimes.StartTime Then
            GetSessionTimes.StartTime = lSessionTimes.StartTime
        End If
        If GetSessionTimes.EndTime = 0 Or lSessionTimes.EndTime > GetSessionTimes.EndTime Then
            GetSessionTimes.EndTime = lSessionTimes.EndTime
        End If
        lNumberOfSessionTimesProcessed = lNumberOfSessionTimesProcessed + 1
    End If
    If lNumberOfSessionTimesProcessed = 6 Then Exit For
Next

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsSmartExchange(ByVal pExchange As String) As Boolean
pExchange = UCase$(pExchange)
IsSmartExchange = (pExchange = ExchangeSmart Or _
                    pExchange = ExchangeSmartAUS Or _
                    pExchange = ExchangeSmartCAN Or _
                    pExchange = ExchangeSmartEUR Or _
                    pExchange = ExchangeSmartNASDAQ Or _
                    pExchange = ExchangeSmartNYSE Or _
                    pExchange = ExchangeSmartUK Or _
                    pExchange = ExchangeSmartUS Or _
                    InStr(1, pExchange, ExchangeSmartQualified) <> 0)
End Function

Private Function OptionRightToTwsOptRight(ByVal Value As OptionRights) As TwsOptionRights
Select Case Value
Case OptionRights.OptNone
    OptionRightToTwsOptRight = TwsOptRightNone
Case OptionRights.OptCall
    OptionRightToTwsOptRight = TwsOptRightCall
Case OptionRights.OptPut
    OptionRightToTwsOptRight = TwsOptRightPut
Case Else
    AssertArgument False, "Value is not a valid Option Right"
End Select
End Function

Private Function OrderActionFromString(ByVal Value As String) As OrderActions
Select Case UCase$(Value)
Case ""
    OrderActionFromString = OrderActionNone
Case "BUY"
    OrderActionFromString = OrderActionBuy
Case "SELL"
    OrderActionFromString = OrderActionSell
Case Else
    AssertArgument False, "Value is not a valid Order Action"
End Select
End Function

Private Function OrderActionToTwsOrderAction(ByVal Value As OrderActions) As TwsOrderActions
Select Case Value
Case OrderActions.OrderActionNone
    OrderActionToTwsOrderAction = TwsOrderActionNone
Case OrderActions.OrderActionBuy
    OrderActionToTwsOrderAction = TwsOrderActionBuy
Case OrderActions.OrderActionSell
    OrderActionToTwsOrderAction = TwsOrderActionSell
Case Else
    AssertArgument False, "Value is not a valid OrderAction"
End Select
End Function

Public Function OrderStatusFromString(ByVal Value As String) As OrderStatuses
Select Case UCase$(Value)
Case "CREATED"
    OrderStatusFromString = OrderStatusCreated
Case "REJECTED", "INACTIVE"
    OrderStatusFromString = OrderStatusRejected
Case "PENDINGSUBMIT", "APIPENDING"
    OrderStatusFromString = OrderStatusPendingSubmit
Case "PRESUBMITTED"
    OrderStatusFromString = OrderStatusPreSubmitted
Case "SUBMITTED"
    OrderStatusFromString = OrderStatusSubmitted
Case "PENDINGCANCEL"
    OrderStatusFromString = OrderStatusCancelling
Case "CANCELLED", "APICANCELLED"
    OrderStatusFromString = OrderStatusCancelled
Case "FILLED"
    OrderStatusFromString = OrderStatusFilled
Case Else
    AssertArgument False, "Invalid order status: " & Value
End Select
End Function

Private Function OrderTIFFromString(ByVal Value As String) As OrderTIFs
Select Case UCase$(Value)
Case ""
    OrderTIFFromString = OrderTIFNone
Case "DAY"
    OrderTIFFromString = OrderTIFDay
Case "GTC"
    OrderTIFFromString = OrderTIFGoodTillCancelled
Case "IOC"
    OrderTIFFromString = OrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Value is not a valid Order TIF"
End Select
End Function

Private Function OrderTIFToString(ByVal Value As OrderTIFs) As String
Select Case Value
Case OrderTIFs.OrderTIFNone
    OrderTIFToString = ""
Case OrderTIFs.OrderTIFDay
    OrderTIFToString = "DAY"
Case OrderTIFs.OrderTIFGoodTillCancelled
    OrderTIFToString = "GTC"
Case OrderTIFs.OrderTIFImmediateOrCancel
    OrderTIFToString = "IOC"
Case OrderTIFs.OrderTIFNone
    OrderTIFToString = ""
Case Else
    AssertArgument False, "Value is not a valid Order TIF"
End Select
End Function

Private Function OrderTIFToTwsOrderTIF(ByVal Value As OrderTIFs) As TwsOrderTIFs
Select Case Value
Case OrderTIFNone
    OrderTIFToTwsOrderTIF = TwsOrderTIFNone
Case OrderTIFDay
    OrderTIFToTwsOrderTIF = TwsOrderTIFDay
Case OrderTIFGoodTillCancelled
    OrderTIFToTwsOrderTIF = TwsOrderTIFGoodTillCancelled
Case OrderTIFImmediateOrCancel
    OrderTIFToTwsOrderTIF = TwsOrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Value is not a valid OrderTIF"
End Select
End Function

Public Function OrderToTwsOrder( _
                ByVal pOrder As IOrder, _
                ByVal pDataSource As IMarketDataSource) As TwsOrder
Const ProcName As String = "OrderToTwsOrder"
On Error GoTo Err

Set OrderToTwsOrder = New TwsOrder
With pOrder
    OrderToTwsOrder.Action = OrderActionToTwsOrderAction(.Action)
    OrderToTwsOrder.AllOrNone = .AllOrNone
    OrderToTwsOrder.BlockOrder = .BlockOrder
    OrderToTwsOrder.OrderId = .BrokerId
    OrderToTwsOrder.DiscretionaryAmt = .DiscretionaryAmount
    OrderToTwsOrder.DisplaySize = .DisplaySize
    If .GoodAfterTime <> 0 Then OrderToTwsOrder.GoodAfterTime = Format(.GoodAfterTime, "yyyymmdd hh:nn:ss") & IIf(.GoodAfterTimeTZ <> "", " " & StandardTimezoneNameToTwsTimeZoneName(.GoodAfterTimeTZ), "")
    If .GoodTillDate <> 0 Then OrderToTwsOrder.GoodTillDate = Format(.GoodTillDate, "yyyymmdd hh:nn:ss") & IIf(.GoodTillDateTZ <> "", " " & StandardTimezoneNameToTwsTimeZoneName(.GoodTillDateTZ), "")
    OrderToTwsOrder.Hidden = .Hidden
    OrderToTwsOrder.OutsideRth = .IgnoreRegularTradingHours
    OrderToTwsOrder.LmtPrice = .LimitPrice
    OrderToTwsOrder.MinQty = IIf(.MinimumQuantity = 0, GIBEnhApi.MaxLong, .MinimumQuantity)
    If Not .ProviderProperties Is Nothing Then
        OrderToTwsOrder.OcaGroup = .ProviderProperties.GetParameterValue(ProviderPropertyOCAGroup)
    End If
    OrderToTwsOrder.OrderType = OrderTypeToTwsOrderType(.OrderType)
    OrderToTwsOrder.Origin = .Origin
    OrderToTwsOrder.OrderRef = .OriginatorRef
    OrderToTwsOrder.OverridePercentageConstraints = .OverrideConstraints
    Set OrderToTwsOrder.TotalQuantity = .Quantity
    OrderToTwsOrder.SettlingFirm = .SettlingFirm
    OrderToTwsOrder.TriggerMethod = StopTriggerMethodToTwsTriggerMethod(.StopTriggerMethod)
    OrderToTwsOrder.SweepToFill = .SweepToFill
    OrderToTwsOrder.Tif = OrderTIFToTwsOrderTIF(.TimeInForce)
    If .OrderType = OrderTypeTrail Then
        OrderToTwsOrder.TrailStopPrice = .TriggerPrice
        OrderToTwsOrder.AuxPrice = Abs(getMarketPrice(pDataSource, pOrder.Action) - .TriggerPrice)
    ElseIf .OrderType = OrderTypeTrailLimit Then
        OrderToTwsOrder.TrailStopPrice = .TriggerPrice
        OrderToTwsOrder.AuxPrice = Abs(getMarketPrice(pDataSource, pOrder.Action) - .TriggerPrice)
    Else
        OrderToTwsOrder.AuxPrice = .TriggerPrice
    End If
End With

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function OrderTypeToTwsOrderType(ByVal Value As OrderTypes) As TwsOrderTypes
Const ProcName As String = "OrderTypeToTwsOrderType"
On Error GoTo Err

Select Case Value
Case OrderTypes.OrderTypeNone
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeNone
Case OrderTypes.OrderTypeMarket
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarket
Case OrderTypes.OrderTypeMarketOnClose
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case OrderTypes.OrderTypeLimit
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimit
Case OrderTypes.OrderTypeLimitOnClose
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case OrderTypes.OrderTypePeggedToMarket
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypePeggedToMarket
Case OrderTypes.OrderTypeStop
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeStop
Case OrderTypes.OrderTypeStopLimit
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeStopLimit
Case OrderTypes.OrderTypeTrail
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeTrail
Case OrderTypes.OrderTypeRelative
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeRelative
Case OrderTypes.OrderTypeMarketToLimit
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketToLimit
Case OrderTypes.OrderTypeLimitIfTouched
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitIfTouched
Case OrderTypes.OrderTypeMarketIfTouched
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketIfTouched
Case OrderTypes.OrderTypeTrailLimit
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeTrailLimit
Case OrderTypes.OrderTypeMarketWithProtection
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketWithProtection
Case OrderTypes.OrderTypeMarketOnOpen
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketOnOpen
Case OrderTypes.OrderTypeLimitOnOpen
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitOnOpen
Case OrderTypes.OrderTypePeggedToPrimary
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypePeggedToPrimary
Case OrderTypes.OrderTypeMidprice
    OrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMidprice
Case Else
    AssertArgument False, "Value is not a valid OrderType"
End Select

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RequestExecutions( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pReqId As Long, _
                ByVal pClientId As Long, _
                ByVal pFrom As Date)
Const ProcName As String = "RequestExecutions"
On Error GoTo Err

Dim lExecFilter As TwsExecutionFilter
Set lExecFilter = New TwsExecutionFilter
lExecFilter.ClientID = pClientId
lExecFilter.Time = pFrom
pTwsAPI.RequestExecutions pReqId, lExecFilter

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RequestOpenOrders(ByVal pTwsAPI As TwsAPI)
Const ProcName As String = "RequestOpenOrders"
On Error GoTo Err

pTwsAPI.RequestOpenOrders

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub


Private Function SecTypeToTwsSecType(ByVal Value As SecurityTypes) As TwsSecTypes
Select Case UCase$(Value)
Case SecurityTypes.SecTypeNone
    SecTypeToTwsSecType = TwsSecTypeNone
Case SecurityTypes.SecTypeStock
    SecTypeToTwsSecType = TwsSecTypeStock
Case SecurityTypes.SecTypeFuture
    SecTypeToTwsSecType = TwsSecTypeFuture
Case SecurityTypes.SecTypeOption
    SecTypeToTwsSecType = TwsSecTypeOption
Case SecurityTypes.SecTypeFuturesOption
    SecTypeToTwsSecType = TwsSecTypeFuturesOption
Case SecurityTypes.SecTypeCash
    SecTypeToTwsSecType = TwsSecTypeCash
Case SecurityTypes.SecTypeCombo
    SecTypeToTwsSecType = TwsSecTypeCombo
Case SecurityTypes.SecTypeIndex
    SecTypeToTwsSecType = TwsSecTypeIndex
Case SecurityTypes.SecTypeWarrant
    SecTypeToTwsSecType = TwsSecTypeWarrant
Case Else
    AssertArgument False, "Value is not a valid Security Type"
End Select
End Function

Public Function StandardTimezoneNameToTwsTimeZoneName(ByVal pTimezoneName As String) As String
Const ProcName As String = "StandardTimezoneNameToTwsTimeZoneName"
On Error GoTo Err

Select Case pTimezoneName
Case ""
    StandardTimezoneNameToTwsTimeZoneName = StandardTimezoneNameToTwsTimeZoneName(GetTimeZone("").StandardName)
Case "AUS Eastern Standard Time"
    StandardTimezoneNameToTwsTimeZoneName = "Australia/Sydney"
Case "Central Standard Time"
    StandardTimezoneNameToTwsTimeZoneName = "US/Central"
Case "China Standard Time"
    StandardTimezoneNameToTwsTimeZoneName = "Asia/Hong_Kong"
Case "GMT Standard Time"
    StandardTimezoneNameToTwsTimeZoneName = "Europe/London"
Case "Eastern Standard Time"
    StandardTimezoneNameToTwsTimeZoneName = "US/Eastern"
Case "Pacific Standard Time"
    StandardTimezoneNameToTwsTimeZoneName = "US/Pacific"
Case "Central Europe Standard Time", "Romance Standard Time"
    StandardTimezoneNameToTwsTimeZoneName = "MET"
Case "UTC"
    StandardTimezoneNameToTwsTimeZoneName = "UTC"
Case Else
    AssertArgument False, "Unrecognised timezone: " & pTimezoneName
End Select

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function StopTriggerMethodToTwsTriggerMethod(ByVal Value As OrderStopTriggerMethods) As TwsStopTriggerMethods
Select Case Value
Case OrderStopTriggerMethods.OrderStopTriggerDefault
    StopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDefault
Case OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk
    StopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDoubleBidAsk
Case OrderStopTriggerMethods.OrderStopTriggerLast
    StopTriggerMethodToTwsTriggerMethod = TwsStopTriggerLast
Case OrderStopTriggerMethods.OrderStopTriggerDoubleLast
    StopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDoubleLast
Case OrderStopTriggerMethods.OrderStopTriggerBidAsk
    StopTriggerMethodToTwsTriggerMethod = TwsStopTriggerBidAsk
Case OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk
    StopTriggerMethodToTwsTriggerMethod = TwsStopTriggerLastOrBidAsk
Case OrderStopTriggerMethods.OrderStopTriggerMidPoint
    StopTriggerMethodToTwsTriggerMethod = TwsStopTriggerMidPoint
Case Else
    AssertArgument False, "Value is not a valid StopTriggerMethod"
End Select
End Function
                
Public Function TwsContractSpecToContractSpecifier( _
                ByVal pTwsContractSpec As TwsContractSpecifier, _
                ByVal pPriceMagnifier As String) As IContractSpecifier
Const ProcName As String = "TwsContractSpecToContractSpecifier"
On Error GoTo Err

Dim lContractSpec As ContractSpecifier
With pTwsContractSpec
    Dim lExchange As String: lExchange = .Exchange
    If lExchange = ExchangeSmart And .PrimaryExch <> "" Then
        lExchange = ExchangeSmartQualified & .PrimaryExch
    End If
    
    Dim lMultiplier As Double
    If .Sectype = TwsSecTypeStock Then
        lMultiplier = 1 / pPriceMagnifier
    Else
        lMultiplier = .Multiplier / pPriceMagnifier
    End If
    Dim lTradingClass As String
    If .Sectype <> TwsSecTypeStock Then lTradingClass = .TradingClass
    Set lContractSpec = CreateContractSpecifier(.LocalSymbol, _
                                                .Symbol, _
                                                lTradingClass, _
                                                lExchange, _
                                                TwsSecTypeToSecType(.Sectype), _
                                                .CurrencyCode, _
                                                .Expiry, _
                                                lMultiplier, _
                                                .Strike, _
                                                TwsOptionRightToOptionRight(.OptRight))
    Dim p As New Parameters
    p.SetParameterValue ProviderPropertyContractID, .ConId
    p.SetParameterValue ProviderPropertyPriceMagnifier, pPriceMagnifier
    lContractSpec.ProviderProperties = p
End With

Set TwsContractSpecToContractSpecifier = lContractSpec

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function TwsContractToContract(ByVal pTwsContract As TwsContract) As IContract
Const ProcName As String = "TwsContractToContract"
On Error GoTo Err

Dim lBuilder As ContractBuilder

With pTwsContract
    Set lBuilder = CreateContractBuilder(GIBEnhancedApi.TwsContractSpecToContractSpecifier(.Specifier, .PriceMagnifier))
    
    Dim lExpiry As Date
    With .Specifier
        If .Expiry <> "" Then
            lExpiry = CDate(Left$(.Expiry, 4) & "/" & _
                                Mid$(.Expiry, 5, 2) & "/" & _
                                Right$(.Expiry, 2))
        End If
    End With
    If lExpiry <> 0 Then lBuilder.ExpiryDate = lExpiry + .LastTradeTime
    
    lBuilder.Description = .LongName
    lBuilder.TickSize = .MinTick
    lBuilder.TimezoneName = TwsTimezoneNameToStandardTimeZoneName(.TimeZoneId)
    
    Dim lSessionTimes As SessionTimes
    lSessionTimes = GetSessionTimes(.LiquidHours)
    lBuilder.sessionStartTime = lSessionTimes.StartTime
    lBuilder.sessionEndTime = lSessionTimes.EndTime
    
    lSessionTimes = GetSessionTimes(.TradingHours)
    lBuilder.FullSessionStartTime = lSessionTimes.StartTime
    lBuilder.FullSessionEndTime = lSessionTimes.EndTime
    
    Dim lProviderProps As New Parameters
    lProviderProps.SetParameterValue "Category", .Category
    lProviderProps.SetParameterValue "ContractMonth", .ContractMonth
    lProviderProps.SetParameterValue "EvMultiplier", .EvMultiplier
    lProviderProps.SetParameterValue "EvRule", .EvRule
    lProviderProps.SetParameterValue "Industry", .Industry
    lProviderProps.SetParameterValue "LiquidHours", .LiquidHours
    lProviderProps.SetParameterValue "MarketName", .MarketName
    lProviderProps.SetParameterValue "MarketRuleID", getMarketRuleID(pTwsContract)
    lProviderProps.SetParameterValue "OrderTypes", .OrderTypes
    lProviderProps.SetParameterValue "PriceMagnifier", .PriceMagnifier
    
    Dim s As String
    If Not .SecIdList Is Nothing Then
        Dim lParam As Parameter
        For Each lParam In .SecIdList
            s = s & IIf(s <> "", ";", "")
            s = s & lParam.Name & ":" & lParam.Value
        Next
    End If
    lProviderProps.SetParameterValue "SecIdList", s
    
    lProviderProps.SetParameterValue "Subcategory", .Subcategory
    lProviderProps.SetParameterValue "TradingHours", .TradingHours
    lProviderProps.SetParameterValue "UnderConId", .UnderConId
    lProviderProps.SetParameterValue "ValidExchanges", .ValidExchanges
    lBuilder.ProviderProperties = lProviderProps
End With

Set TwsContractToContract = lBuilder.Contract

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function TwsOptionRightToOptionRight(ByVal Value As TwsOptionRights) As OptionRights
Select Case Value
Case TwsOptRightNone
    TwsOptionRightToOptionRight = OptionRights.OptNone
Case TwsOptRightCall
    TwsOptionRightToOptionRight = OptionRights.OptCall
Case TwsOptRightPut
    TwsOptionRightToOptionRight = OptionRights.OptPut
Case Else
    AssertArgument False, "Value is not a valid Option Right"
End Select
End Function

Public Function TwsOrderActionToOrderAction(ByVal Value As TwsOrderActions) As OrderActions
Select Case Value
Case TwsOrderActionNone
    TwsOrderActionToOrderAction = OrderActionNone
Case TwsOrderActionBuy
    TwsOrderActionToOrderAction = OrderActionBuy
Case TwsOrderActionSell
    TwsOrderActionToOrderAction = OrderActionSell
Case Else
    AssertArgument False, "Value is not a valid Order Action"
End Select
End Function

Public Function TwsOrderTypeToOrderType(ByVal Value As TwsOrderTypes) As OrderTypes
Const ProcName As String = "TwsOrderTypeToOrderType"
On Error GoTo Err

Select Case Value
Case TwsOrderTypes.TwsOrderTypeNone
    TwsOrderTypeToOrderType = OrderTypeNone
Case TwsOrderTypes.TwsOrderTypeMarket
    TwsOrderTypeToOrderType = OrderTypeMarket
Case TwsOrderTypes.TwsOrderTypeMarketOnClose
    TwsOrderTypeToOrderType = OrderTypeMarketOnClose
Case TwsOrderTypes.TwsOrderTypeLimit
    TwsOrderTypeToOrderType = OrderTypeLimit
Case TwsOrderTypes.TwsOrderTypeLimitOnClose
    TwsOrderTypeToOrderType = OrderTypeLimitOnClose
Case TwsOrderTypes.TwsOrderTypePeggedToMarket
    TwsOrderTypeToOrderType = OrderTypePeggedToMarket
Case TwsOrderTypes.TwsOrderTypeStop
    TwsOrderTypeToOrderType = OrderTypeStop
Case TwsOrderTypes.TwsOrderTypeStopLimit
    TwsOrderTypeToOrderType = OrderTypeStopLimit
Case TwsOrderTypes.TwsOrderTypeTrail
    TwsOrderTypeToOrderType = OrderTypeTrail
Case TwsOrderTypes.TwsOrderTypeRelative
    TwsOrderTypeToOrderType = OrderTypeRelative
Case TwsOrderTypes.TwsOrderTypeMarketToLimit
    TwsOrderTypeToOrderType = OrderTypeMarketToLimit
Case TwsOrderTypes.TwsOrderTypeLimitIfTouched
    TwsOrderTypeToOrderType = OrderTypeLimitIfTouched
Case TwsOrderTypes.TwsOrderTypeMarketIfTouched
    TwsOrderTypeToOrderType = OrderTypeMarketIfTouched
Case TwsOrderTypes.TwsOrderTypeTrailLimit
    TwsOrderTypeToOrderType = OrderTypeTrailLimit
Case TwsOrderTypes.TwsOrderTypeMarketWithProtection
    TwsOrderTypeToOrderType = OrderTypeMarketWithProtection
Case TwsOrderTypes.TwsOrderTypeMarketOnOpen
    TwsOrderTypeToOrderType = OrderTypeMarketOnOpen
Case TwsOrderTypes.TwsOrderTypeLimitOnOpen
    TwsOrderTypeToOrderType = OrderTypeLimitOnOpen
Case TwsOrderTypes.TwsOrderTypePeggedToPrimary
    TwsOrderTypeToOrderType = OrderTypePeggedToPrimary
Case TwsOrderTypes.TwsOrderTypeMidprice
    TwsOrderTypeToOrderType = OrderTypeMidprice
Case Else
    TwsOrderTypeToOrderType = OrderTypeNone
End Select

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function TwsSecTypeToSecType(ByVal Value As TwsSecTypes) As SecurityTypes
Select Case Value
Case TwsSecTypeNone
    TwsSecTypeToSecType = SecurityTypes.SecTypeNone
Case TwsSecTypeStock
    TwsSecTypeToSecType = SecurityTypes.SecTypeStock
Case TwsSecTypeFuture
    TwsSecTypeToSecType = SecurityTypes.SecTypeFuture
Case TwsSecTypeOption
    TwsSecTypeToSecType = SecurityTypes.SecTypeOption
Case TwsSecTypeFuturesOption
    TwsSecTypeToSecType = SecurityTypes.SecTypeFuturesOption
Case TwsSecTypeCash
    TwsSecTypeToSecType = SecurityTypes.SecTypeCash
Case TwsSecTypeCombo
    TwsSecTypeToSecType = SecurityTypes.SecTypeCombo
Case TwsSecTypeIndex
    TwsSecTypeToSecType = SecurityTypes.SecTypeIndex
Case TwsSecTypeWarrant
    TwsSecTypeToSecType = SecurityTypes.SecTypeWarrant
Case Else
    AssertArgument False, "Value is not a valid Security Type"
End Select
End Function

Public Function TwsTriggerMethodToStopTriggerMethod(ByVal Value As TwsStopTriggerMethods) As OrderStopTriggerMethods
Select Case Value
Case TwsStopTriggerDefault
    TwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerDefault
Case TwsStopTriggerDoubleBidAsk
    TwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk
Case TwsStopTriggerLast
    TwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerLast
Case TwsStopTriggerDoubleLast
    TwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerDoubleLast
Case TwsStopTriggerBidAsk
    TwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerBidAsk
Case TwsStopTriggerLastOrBidAsk
    TwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk
Case TwsStopTriggerMidPoint
    TwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerMidPoint
Case Else
    AssertArgument False, "Value is not a valid TwsStopTriggerMethod"
End Select
End Function
                
Public Function TwsTimezoneNameToStandardTimeZoneName(ByVal pTimeZoneId As String) As String
Const ProcName As String = "TwsTimezoneNameToStandardTimeZoneName"
On Error GoTo Err

' Some of this list of timezone names comes from the Requesting Contract
' Details section of IBKR's API documentation at:
'
' https://interactivebrokers.github.io/tws-api/contract_details.html
'
' Other names have been collected over the years as IBKR has changed them (but
' are probably no longer likely to be encountered).
'
'Africa/Johannesburg
'
'Asia/Calcutta
'
'Australia/NSW
'
'Europe/Budapest
'Europe/Helsinki
'Europe/Moscow
'Europe/Riga
'Europe/Tallinn
'Europe/Warsaw
'Europe/Vilnius
'
'GB-Eire
'
'GMT
'
'Hongkong
'
'Israel
'
'Japan
'
'MET
'
'US/Central
'US/Eastern
'US/Pacific


Select Case UCase$(pTimeZoneId)
Case ""
    TwsTimezoneNameToStandardTimeZoneName = ""
Case "AUSTRALIA/NSW", _
        "AET", _
        "AEST (AUSTRALIAN EASTERN STANDARD TIME (NEW SOUTH WALES))"
    TwsTimezoneNameToStandardTimeZoneName = "AUS EASTERN STANDARD TIME"
Case "EUROPE/BUDAPEST", _
        "EUROPE/WARSAW", _
        "MET", _
        "MET (MIDDLE EUROPE TIME)"
    TwsTimezoneNameToStandardTimeZoneName = "CENTRAL EUROPE STANDARD TIME"
Case "CST", _
        "CTT", _
        "CST (CENTRAL STANDARD TIME)", _
        "US/CENTRAL"
    TwsTimezoneNameToStandardTimeZoneName = "CENTRAL STANDARD TIME"
Case "ASIA/HONG_KONG", _
        "HONGKONG"
    TwsTimezoneNameToStandardTimeZoneName = "CHINA STANDARD TIME"
Case "EUROPE/HELSINKI", _
        "EUROPE/RIGA", _
        "EUROPE/TALLINN", _
        "EUROPE/VILNIUS"
    TwsTimezoneNameToStandardTimeZoneName = "E. EUROPE STANDARD TIME"
Case "EST", _
        "EST5EDT", _
        "EST (EASTERN STANDARD TIME)", _
        "US/EASTERN"
    TwsTimezoneNameToStandardTimeZoneName = "EASTERN STANDARD TIME"
Case "EUROPE/LONDON", _
        "GB", _
        "GB-EIRE", _
        "GMT", _
        "GMT (GREENWICH MEAN TIME)", _
        "BRITISH SUMMER TIME", _
        "BST (BRITISH SUMMER TIME)"
    TwsTimezoneNameToStandardTimeZoneName = "GMT STANDARD TIME"
Case "ASIA/CALCUTTA"
    TwsTimezoneNameToStandardTimeZoneName = "INDIA STANDARD TIME"
Case "ISRAEL"
    TwsTimezoneNameToStandardTimeZoneName = "MIDDLE EAST STANDARD TIME"
Case "PST", _
        "US/PACIFIC", _
        "PACIFIC/PITCAIRN"
    TwsTimezoneNameToStandardTimeZoneName = "PACIFIC STANDARD TIME"
Case "EUROPE/MOSCOW"
    TwsTimezoneNameToStandardTimeZoneName = "RUSSIAN STANDARD TIME"
Case "AFRICA/JOHANNESBURG"
    TwsTimezoneNameToStandardTimeZoneName = "SOUTH AFRICA STANDARD TIME"
Case "JAPAN"
    TwsTimezoneNameToStandardTimeZoneName = "TOKYO STANDARD TIME"
Case "UTC"
    TwsTimezoneNameToStandardTimeZoneName = "UTC"
Case Else
    GIBEnhApi.Log "Unrecognised timezone: " & pTimeZoneId, ModuleName, ProcName, , LogLevelSevere
    TwsTimezoneNameToStandardTimeZoneName = ""
End Select

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getMarketPrice( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pAction As OrderActions) As Double
Const ProcName As String = "getMarketPrice"
On Error GoTo Err

Dim lMarketPrice As Double
If pDataSource.HasCurrentTick(TickTypeTrade) Then
    lMarketPrice = pDataSource.CurrentTick(TickTypeTrade).Price
ElseIf pDataSource.HasCurrentTick( _
                        IIf(pAction = OrderActionBuy, _
                            TickTypeAsk, _
                            TickTypeBid)) Then
    lMarketPrice = IIf(pAction = OrderActionBuy, _
                        pDataSource.CurrentTick(TickTypeAsk).Price, _
                        pDataSource.CurrentTick(TickTypeBid).Price)
ElseIf pDataSource.HasCurrentTick( _
                        IIf(pAction = OrderActionBuy, _
                            TickTypeBid, _
                            TickTypeAsk)) Then

    lMarketPrice = IIf(pAction = OrderActionBuy, _
                        pDataSource.CurrentTick(TickTypeBid).Price, _
                        pDataSource.CurrentTick(TickTypeAsk).Price)
End If

getMarketPrice = lMarketPrice

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getMarketRuleID( _
                ByVal pTwsContract As TwsContract) As Long
Const ProcName As String = "getMarketRuleID"
On Error GoTo Err

If pTwsContract.ValidExchanges = "" Then Exit Function

Dim lExchanges() As String
lExchanges = Split(pTwsContract.ValidExchanges, ",")

Dim i As Long
For i = 0 To UBound(lExchanges)
    If pTwsContract.Specifier.Exchange = lExchanges(i) Then Exit For
Next

Dim lMarketRuleIds() As String
lMarketRuleIds = Split(pTwsContract.MarketRuleIds, ",")

If pTwsContract.PriceMagnifier = 1 Then
    getMarketRuleID = lMarketRuleIds(i)
Else
    getMarketRuleID = -lMarketRuleIds(i)
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSessionTimesForDay( _
                ByVal pRegExp As RegExp, _
                ByVal pSessionTimesString As String, _
                ByRef pSessionTimes As SessionTimes, _
                ByRef pSessionDate As Date) As Boolean
Const ProcName As String = "getSessionTimesForDay"
On Error GoTo Err

Dim lMatches As MatchCollection
Set lMatches = pRegExp.Execute(pSessionTimesString)

If lMatches.Count <> 1 Then Exit Function

Dim lMatch As Match: Set lMatch = lMatches(0)
If lMatch.SubMatches(6) = "CLOSED" Then Exit Function

Dim lSessionDateStr As String

lSessionDateStr = lMatch.SubMatches(0)
Dim lSessionStartDate As Date
lSessionStartDate = CDate(Left$(lSessionDateStr, 4) & "/" & Mid$(lSessionDateStr, 5, 2) & "/" & Right$(lSessionDateStr, 2))

lSessionDateStr = lMatch.SubMatches(3)
Dim lSessionEndDate As Date
lSessionEndDate = CDate(Left$(lSessionDateStr, 4) & "/" & Mid$(lSessionDateStr, 5, 2) & "/" & Right$(lSessionDateStr, 2))

If lSessionStartDate = pSessionDate And lSessionEndDate = pSessionDate Then Exit Function
pSessionDate = lSessionEndDate

pSessionTimes.StartTime = CDate(lMatch.SubMatches(1) & ":" & lMatch.SubMatches(2))
pSessionTimes.EndTime = CDate(lMatch.SubMatches(4) & ":" & lMatch.SubMatches(5))
    
getSessionTimesForDay = True

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function






