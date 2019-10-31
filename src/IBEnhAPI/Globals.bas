Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                        As String = "IBEnhancedAPI27"
Private Const ModuleName                        As String = "Globals"

Public Const InvalidEnumValue                   As String = "*ERR*"
Public Const NullIndex                          As Long = -1

Public Const MaxLong                            As Long = &H7FFFFFFF
Public Const OneMicrosecond                     As Double = 1# / 86400000000#
Public Const OneMinute                          As Double = 1# / 1440#
Public Const OneSecond                          As Double = 1# / 86400#

Public Const NumDaysInWeek                      As Long = 5
Public Const NumDaysInMonth                     As Long = 22
Public Const NumDaysInYear                      As Long = 260
Public Const NumMonthsInYear                    As Long = 12

Public Const ExchangeSmart                      As String = "SMART"
Public Const ExchangeSmartCAN                   As String = "SMARTCAN"
Public Const ExchangeSmartEUR                   As String = "SMARTEUR"
Public Const ExchangeSmartNASDAQ                As String = "SMARTNASDAQ"
Public Const ExchangeSmartNYSE                  As String = "SMARTNYSE"
Public Const ExchangeSmartUK                    As String = "SMARTUK"
Public Const ExchangeSmartUS                    As String = "SMARTUS"
Public Const ExchangeSmartQualified             As String = "SMART/"

Public Const PrimaryExchangeARCA                As String = "ARCA"
Public Const PrimaryExchangeEBS                 As String = "EBS"
Public Const PrimaryExchangeFWB                 As String = "FWB"
Public Const PrimaryExchangeIBIS                As String = "IBIS"
Public Const PrimaryExchangeLSE                 As String = "LSE"
Public Const PrimaryExchangeSWB                 As String = "SWB"
Public Const PrimaryExchangeNASDAQ              As String = "NASDAQ"
Public Const PrimaryExchangeNYSE                As String = "NYSE"
Public Const PrimaryExchangeVENTURE             As String = "VENTURE"

Public Const ProviderPropertyOCAGroup           As String = "OCA group"

'================================================================================
' Enums
'================================================================================

Public Enum MarketDataReestablishmentModes
    ' only request market data for tickers that have not yet
    ' received it
    NewTickersOnly = 1
    
    ' cancel the current market data before re-requesting
    Cancel
    
    ' request market data for all tickers
    All
    
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Global variables
'================================================================================

'================================================================================
' Private variables
'================================================================================

Private mLogger As FormattingLogger

'================================================================================
' Properties
'================================================================================

Public Property Get gLogger() As FormattingLogger
If mLogger Is Nothing Then Set mLogger = CreateFormattingLogger("tradebuild.log.ibenhancedapi", ProjectName)
Set gLogger = mLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Function gContractSpecToTwsContract(ByVal pContractSpecifier As IContractSpecifier) As TwsContract
Const ProcName As String = "gContractSpecToTwsContract"
On Error GoTo Err

Dim lComboLeg As ComboLeg
Dim lTwsComboLeg As TwsComboLeg

Set gContractSpecToTwsContract = New TwsContract

With gContractSpecToTwsContract
    .CurrencyCode = pContractSpecifier.CurrencyCode
    Dim lExchange As String: lExchange = UCase$(pContractSpecifier.Exchange)
    If lExchange = ExchangeSmartCAN Then
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
        .PrimaryExch = PrimaryExchangeARCA
        .Exchange = ExchangeSmart
    ElseIf lExchange = ExchangeSmartEUR Then
        .PrimaryExch = PrimaryExchangeIBIS
        .Exchange = ExchangeSmart
    ElseIf InStr(1, lExchange, ExchangeSmartQualified) = 1 Then
        .PrimaryExch = Right$(lExchange, Len(lExchange) - Len(ExchangeSmartQualified))
        .Exchange = ExchangeSmart
    Else
        .Exchange = lExchange
    End If
    .Expiry = IIf(Len(pContractSpecifier.Expiry) >= 6, pContractSpecifier.Expiry, "")
    .IncludeExpired = True
    .LocalSymbol = pContractSpecifier.LocalSymbol
    .Multiplier = pContractSpecifier.Multiplier
    If .CurrencyCode = "GBP" And .Multiplier <> 1 Then .Multiplier = .Multiplier * 100
    .OptRight = gOptionRightToTwsOptRight(pContractSpecifier.Right)
    .SecType = gSecTypeToTwsSecType(pContractSpecifier.SecType)
    .Strike = pContractSpecifier.Strike
    .Symbol = pContractSpecifier.Symbol
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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gContractFutureToTwsContractDetails( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractFuture As IFuture, _
                ByVal pContractCache As TwsContractCache) As IFuture
Const ProcName As String = "gContractFutureToTwsContractDetails"
On Error GoTo Err

Dim lFutureBuilder As New TwsContractDtlsFutBldr
lFutureBuilder.Initialise pContractRequester, pContractFuture, pContractCache
Set gContractFutureToTwsContractDetails = lFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gContractToTwsContractDetails(ByVal pContract As IContract) As TwsContractDetails
Const ProcName As String = "gContractToTwsContractDetails"
On Error GoTo Err

Dim lContract As TwsContract
Dim lContractDetails As TwsContractDetails

Assert pContract.Specifier.SecType <> SecTypeCombo, "Combo contracts not supported", ErrorCodes.ErrUnsupportedOperationException

Set lContractDetails = New TwsContractDetails
Set lContract = gContractSpecToTwsContract(pContract.Specifier)

With lContractDetails
    .Summary = lContract
    .MinTick = pContract.TickSize
    .TimeZoneId = gStandardTimezoneNameToTwsTimeZoneName(pContract.TimezoneName)
End With

Set gContractToTwsContractDetails = lContractDetails

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFetchContracts( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pListener As IContractFetchListener, _
                ByVal pCookie As Variant, _
                ByVal pReturnTwsContracts As Boolean) As IFuture
Const ProcName As String = "gFetchContracts"
On Error GoTo Err

Dim lFetchTask As New ContractDetailsRequestTask

lFetchTask.Initialise pContractRequester, pContractCache, pContractSpecifier, pListener, pCookie, pReturnTwsContracts
If gLogger.IsLoggable(LogLevelDetail) Then gLog "Fetching contract details for", ModuleName, ProcName, pContractSpecifier.ToString, LogLevelDetail

StartTask lFetchTask, PriorityLow

Set gFetchContracts = lFetchTask.ContractsFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetObjectKey(ByVal pObject As Object) As String
gGetObjectKey = Hex$(ObjPtr(pObject))
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
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub gLog(ByRef pMsg As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err


gLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gOptionRightToTwsOptRight(ByVal Value As OptionRights) As TwsOptionRights
Select Case Value
Case OptionRights.OptNone
    gOptionRightToTwsOptRight = TwsOptRightNone
Case OptionRights.OptCall
    gOptionRightToTwsOptRight = TwsOptRightCall
Case OptionRights.OptPut
    gOptionRightToTwsOptRight = TwsOptRightPut
Case Else
    AssertArgument False, "Value is not a valid Option Right"
End Select
End Function

Public Function gOrderActionFromString(ByVal Value As String) As OrderActions
Select Case UCase$(Value)
Case ""
    gOrderActionFromString = OrderActionNone
Case "BUY"
    gOrderActionFromString = OrderActionBuy
Case "SELL"
    gOrderActionFromString = OrderActionSell
Case Else
    AssertArgument False, "Value is not a valid Order Action"
End Select
End Function

Public Function gOrderActionToTwsOrderAction(ByVal Value As OrderActions) As TwsOrderActions
Select Case Value
Case OrderActions.OrderActionNone
    gOrderActionToTwsOrderAction = TwsOrderActionNone
Case OrderActions.OrderActionBuy
    gOrderActionToTwsOrderAction = TwsOrderActionBuy
Case OrderActions.OrderActionSell
    gOrderActionToTwsOrderAction = TwsOrderActionSell
Case Else
    AssertArgument False, "Value is not a valid OrderAction"
End Select
End Function

Public Function gOrderStatusFromString(ByVal Value As String) As OrderStatuses
Select Case UCase$(Value)
Case "CREATED"
    gOrderStatusFromString = OrderStatusCreated
Case "REJECTED", "INACTIVE"
    gOrderStatusFromString = OrderStatusRejected
Case "PENDINGSUBMIT", "APIPENDING"
    gOrderStatusFromString = OrderStatusPendingSubmit
Case "PRESUBMITTED"
    gOrderStatusFromString = OrderStatusPreSubmitted
Case "SUBMITTED"
    gOrderStatusFromString = OrderStatusSubmitted
Case "PENDINGCANCEL"
    gOrderStatusFromString = OrderStatusCancelling
Case "CANCELLED", "APICANCELLED"
    gOrderStatusFromString = OrderStatusCancelled
Case "FILLED"
    gOrderStatusFromString = OrderStatusFilled
Case Else
    AssertArgument False, "Invalid order status: " & Value
End Select
End Function

Public Function gOrderTIFFromString(ByVal Value As String) As OrderTIFs
Select Case UCase$(Value)
Case ""
    gOrderTIFFromString = OrderTIFNone
Case "DAY"
    gOrderTIFFromString = OrderTIFDay
Case "GTC"
    gOrderTIFFromString = OrderTIFGoodTillCancelled
Case "IOC"
    gOrderTIFFromString = OrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Value is not a valid Order TIF"
End Select
End Function

Public Function gOrderTIFToString(ByVal Value As OrderTIFs) As String
Select Case Value
Case OrderTIFs.OrderTIFNone
    gOrderTIFToString = ""
Case OrderTIFs.OrderTIFDay
    gOrderTIFToString = "DAY"
Case OrderTIFs.OrderTIFGoodTillCancelled
    gOrderTIFToString = "GTC"
Case OrderTIFs.OrderTIFImmediateOrCancel
    gOrderTIFToString = "IOC"
Case OrderTIFs.OrderTIFNone
    gOrderTIFToString = ""
Case Else
    AssertArgument False, "Value is not a valid Order TIF"
End Select
End Function

Public Function gOrderTIFToTwsOrderTIF(ByVal Value As OrderTIFs) As TwsOrderTIFs
Select Case Value
Case OrderTIFNone
    gOrderTIFToTwsOrderTIF = TwsOrderTIFNone
Case OrderTIFDay
    gOrderTIFToTwsOrderTIF = TwsOrderTIFDay
Case OrderTIFGoodTillCancelled
    gOrderTIFToTwsOrderTIF = TwsOrderTIFGoodTillCancelled
Case OrderTIFImmediateOrCancel
    gOrderTIFToTwsOrderTIF = TwsOrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Value is not a valid OrderTIF"
End Select
End Function

Public Function gOrderToTwsOrder( _
                ByVal pOrder As IOrder, _
                ByVal pDataSource As IMarketDataSource) As TwsOrder
Const ProcName As String = "gOrderToTwsOrder"
On Error GoTo Err

Set gOrderToTwsOrder = New TwsOrder
With pOrder
    gOrderToTwsOrder.Action = gOrderActionToTwsOrderAction(.Action)
    gOrderToTwsOrder.AllOrNone = .AllOrNone
    gOrderToTwsOrder.BlockOrder = .BlockOrder
    gOrderToTwsOrder.OrderId = .BrokerId
    gOrderToTwsOrder.DiscretionaryAmt = .DiscretionaryAmount
    gOrderToTwsOrder.DisplaySize = .DisplaySize
    gOrderToTwsOrder.ETradeOnly = .ETradeOnly
    gOrderToTwsOrder.FirmQuoteOnly = .FirmQuoteOnly
    If .GoodAfterTime <> 0 Then gOrderToTwsOrder.GoodAfterTime = Format(.GoodAfterTime, "yyyymmdd hh:nn:ss") & IIf(.GoodAfterTimeTZ <> "", " " & gStandardTimezoneNameToTwsTimeZoneName(.GoodAfterTimeTZ), "")
    If .GoodTillDate <> 0 Then gOrderToTwsOrder.GoodTillDate = Format(.GoodTillDate, "yyyymmdd hh:nn:ss") & IIf(.GoodTillDateTZ <> "", " " & gStandardTimezoneNameToTwsTimeZoneName(.GoodTillDateTZ), "")
    gOrderToTwsOrder.Hidden = .Hidden
    gOrderToTwsOrder.OutsideRth = .IgnoreRegularTradingHours
    gOrderToTwsOrder.LmtPrice = .LimitPrice
    gOrderToTwsOrder.MinQty = IIf(.MinimumQuantity = 0, MaxLong, .MinimumQuantity)
    gOrderToTwsOrder.NbboPriceCap = IIf(.NbboPriceCap = 0, MaxDouble, .NbboPriceCap)
    gOrderToTwsOrder.OcaGroup = .ProviderProperties.GetParameterValue(ProviderPropertyOCAGroup)
    gOrderToTwsOrder.OrderType = gOrderTypeToTwsOrderType(.OrderType)
    gOrderToTwsOrder.Origin = .Origin
    gOrderToTwsOrder.OrderRef = .OriginatorRef
    gOrderToTwsOrder.OverridePercentageConstraints = .OverrideConstraints
    gOrderToTwsOrder.TotalQuantity = .Quantity
    gOrderToTwsOrder.SettlingFirm = .SettlingFirm
    gOrderToTwsOrder.TriggerMethod = gStopTriggerMethodToTwsTriggerMethod(.StopTriggerMethod)
    gOrderToTwsOrder.SweepToFill = .SweepToFill
    gOrderToTwsOrder.Tif = gOrderTIFToTwsOrderTIF(.TimeInForce)
    If .OrderType = OrderTypeTrail Then
        gOrderToTwsOrder.TrailStopPrice = .TriggerPrice
        gOrderToTwsOrder.AuxPrice = Abs(getMarketPrice(pDataSource, pOrder.Action) - .TriggerPrice)
    ElseIf .OrderType = OrderTypeTrailLimit Then
        gOrderToTwsOrder.TrailStopPrice = .TriggerPrice
        gOrderToTwsOrder.AuxPrice = Abs(getMarketPrice(pDataSource, pOrder.Action) - .TriggerPrice)
    Else
        gOrderToTwsOrder.AuxPrice = .TriggerPrice
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gOrderTypeToTwsOrderType(ByVal Value As OrderTypes) As TwsOrderTypes
Const ProcName As String = "gOrderTypeToTwsOrderType"
On Error GoTo Err

Select Case Value
Case OrderTypes.OrderTypeNone
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeNone
Case OrderTypes.OrderTypeMarket
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarket
Case OrderTypes.OrderTypeMarketOnClose
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case OrderTypes.OrderTypeLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimit
Case OrderTypes.OrderTypeLimitOnClose
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case OrderTypes.OrderTypePeggedToMarket
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypePeggedToMarket
Case OrderTypes.OrderTypeStop
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeStop
Case OrderTypes.OrderTypeStopLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeStopLimit
Case OrderTypes.OrderTypeTrail
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeTrail
Case OrderTypes.OrderTypeRelative
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeRelative
Case OrderTypes.OrderTypeMarketToLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketToLimit
Case OrderTypes.OrderTypeLimitIfTouched
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitIfTouched
Case OrderTypes.OrderTypeMarketIfTouched
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketIfTouched
Case OrderTypes.OrderTypeTrailLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeTrailLimit
Case OrderTypes.OrderTypeMarketWithProtection
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketWithProtection
Case OrderTypes.OrderTypeMarketOnOpen
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketOnOpen
Case OrderTypes.OrderTypeLimitOnOpen
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitOnOpen
Case OrderTypes.OrderTypePeggedToPrimary
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypePeggedToPrimary
Case Else
    AssertArgument False, "Value is not a valid OrderType"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gRequestExecutions( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pReqId As Long, _
                ByVal pClientId As Long, _
                ByVal pFrom As Date)
Const ProcName As String = "gRequestExecutions"
On Error GoTo Err

Dim lExecFilter As TwsExecutionFilter
Set lExecFilter = New TwsExecutionFilter
lExecFilter.ClientId = pClientId
lExecFilter.Time = pFrom
pTwsAPI.RequestExecutions pReqId, lExecFilter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gRequestOpenOrders(ByVal pTwsAPI As TwsAPI)
Const ProcName As String = "gRequestOpenOrders"
On Error GoTo Err

pTwsAPI.RequestOpenOrders

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function gSecTypeToTwsSecType(ByVal Value As SecurityTypes) As TwsSecTypes
Select Case UCase$(Value)
Case SecurityTypes.SecTypeNone
    gSecTypeToTwsSecType = TwsSecTypeNone
Case SecurityTypes.SecTypeStock
    gSecTypeToTwsSecType = TwsSecTypeStock
Case SecurityTypes.SecTypeFuture
    gSecTypeToTwsSecType = TwsSecTypeFuture
Case SecurityTypes.SecTypeOption
    gSecTypeToTwsSecType = TwsSecTypeOption
Case SecurityTypes.SecTypeFuturesOption
    gSecTypeToTwsSecType = TwsSecTypeFuturesOption
Case SecurityTypes.SecTypeCash
    gSecTypeToTwsSecType = TwsSecTypeCash
Case SecurityTypes.SecTypeCombo
    gSecTypeToTwsSecType = TwsSecTypeCombo
Case SecurityTypes.SecTypeIndex
    gSecTypeToTwsSecType = TwsSecTypeIndex
Case Else
    AssertArgument False, "Value is not a valid Security Type"
End Select
End Function

Public Function gStandardTimezoneNameToTwsTimeZoneName(ByVal pTimezoneName As String) As String
Const ProcName As String = "gStandardTimezoneNameToTwsTimeZoneName"
On Error GoTo Err

Select Case pTimezoneName
Case ""
    gStandardTimezoneNameToTwsTimeZoneName = ""
Case "AUS Eastern Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "AET"
Case "Central Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "CST"
Case "China Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "Asia/Hong_Kong"
Case "GMT Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "GMT"
Case "Eastern Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "EST"
Case "Pacific Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "Pacific/Pitcairn"
Case "Romance Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "MET"
Case Else
    AssertArgument False, "Unrecognised timezone: " & pTimezoneName
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gStopTriggerMethodToTwsTriggerMethod(ByVal Value As OrderStopTriggerMethods) As TwsStopTriggerMethods
Select Case Value
Case OrderStopTriggerMethods.OrderStopTriggerDefault
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDefault
Case OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDoubleBidAsk
Case OrderStopTriggerMethods.OrderStopTriggerLast
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerLast
Case OrderStopTriggerMethods.OrderStopTriggerDoubleLast
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDoubleLast
Case OrderStopTriggerMethods.OrderStopTriggerBidAsk
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerBidAsk
Case OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerLastOrBidAsk
Case OrderStopTriggerMethods.OrderStopTriggerMidPoint
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerMidPoint
Case Else
    AssertArgument False, "Value is not a valid StopTriggerMethod"
End Select
End Function
                
Public Function gTwsContractDetailsToContract(ByVal pTwsContractDetails As TwsContractDetails) As IContract
Const ProcName As String = "gTwsContractDetailsToContract"
On Error GoTo Err

Dim lBuilder As ContractBuilder

With pTwsContractDetails
    Set lBuilder = CreateContractBuilder(gTwsContractToContractSpecifier(.Summary, .PriceMagnifier))
    With .Summary
        If .Expiry <> "" Then
            lBuilder.ExpiryDate = CDate(Left$(.Expiry, 4) & "/" & _
                                        Mid$(.Expiry, 5, 2) & "/" & _
                                        Right$(.Expiry, 2))
        End If
    End With
    lBuilder.Description = .LongName
    lBuilder.TickSize = .MinTick
    lBuilder.TimezoneName = gTwsTimezoneNameToStandardTimeZoneName(.TimeZoneId)
End With

Set gTwsContractDetailsToContract = lBuilder.Contract

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTwsContractToContractSpecifier( _
                ByVal pTwsContract As TwsContract, _
                ByVal pPriceMagnifier) As IContractSpecifier
Const ProcName As String = "gTwsContractToContractSpecifier"
On Error GoTo Err

Dim lContractSpec As IContractSpecifier
With pTwsContract
    Dim lExchange As String: lExchange = .Exchange
    If lExchange = ExchangeSmart And .PrimaryExch <> "" Then
        lExchange = ExchangeSmartQualified & .PrimaryExch
    End If
    Set lContractSpec = CreateContractSpecifier(.LocalSymbol, _
                                                .Symbol, _
                                                lExchange, _
                                                gTwsSecTypeToSecType(.SecType), _
                                                .CurrencyCode, _
                                                .Expiry, _
                                                .Multiplier / pPriceMagnifier, _
                                                .Strike, _
                                                gTwsOptionRightToOptionRight(.OptRight))
End With

Set gTwsContractToContractSpecifier = lContractSpec

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTwsOptionRightToOptionRight(ByVal Value As TwsOptionRights) As OptionRights
Select Case Value
Case TwsOptRightNone
    gTwsOptionRightToOptionRight = OptionRights.OptNone
Case TwsOptRightCall
    gTwsOptionRightToOptionRight = OptionRights.OptCall
Case TwsOptRightPut
    gTwsOptionRightToOptionRight = OptionRights.OptPut
Case Else
    AssertArgument False, "Value is not a valid Option Right"
End Select
End Function

Public Function gTwsOrderActionToOrderAction(ByVal Value As TwsOrderActions) As OrderActions
Select Case Value
Case TwsOrderActionNone
    gTwsOrderActionToOrderAction = OrderActionNone
Case TwsOrderActionBuy
    gTwsOrderActionToOrderAction = OrderActionBuy
Case TwsOrderActionSell
    gTwsOrderActionToOrderAction = OrderActionSell
Case Else
    AssertArgument False, "Value is not a valid Order Action"
End Select
End Function

Public Function gTwsOrderTypeToOrderType(ByVal Value As TwsOrderTypes) As OrderTypes
Const ProcName As String = "gTwsOrderTypeToOrderType"
On Error GoTo Err

Select Case Value
Case TwsOrderTypes.TwsOrderTypeNone
    gTwsOrderTypeToOrderType = OrderTypeNone
Case TwsOrderTypes.TwsOrderTypeMarket
    gTwsOrderTypeToOrderType = OrderTypeMarket
Case TwsOrderTypes.TwsOrderTypeMarketOnClose
    gTwsOrderTypeToOrderType = OrderTypeMarketOnClose
Case TwsOrderTypes.TwsOrderTypeLimit
    gTwsOrderTypeToOrderType = OrderTypeLimit
Case TwsOrderTypes.TwsOrderTypeLimitOnClose
    gTwsOrderTypeToOrderType = OrderTypeLimitOnClose
Case TwsOrderTypes.TwsOrderTypePeggedToMarket
    gTwsOrderTypeToOrderType = OrderTypePeggedToMarket
Case TwsOrderTypes.TwsOrderTypeStop
    gTwsOrderTypeToOrderType = OrderTypeStop
Case TwsOrderTypes.TwsOrderTypeStopLimit
    gTwsOrderTypeToOrderType = OrderTypeStopLimit
Case TwsOrderTypes.TwsOrderTypeTrail
    gTwsOrderTypeToOrderType = OrderTypeTrail
Case TwsOrderTypes.TwsOrderTypeRelative
    gTwsOrderTypeToOrderType = OrderTypeRelative
Case TwsOrderTypes.TwsOrderTypeMarketToLimit
    gTwsOrderTypeToOrderType = OrderTypeMarketToLimit
Case TwsOrderTypes.TwsOrderTypeLimitIfTouched
    gTwsOrderTypeToOrderType = OrderTypeLimitIfTouched
Case TwsOrderTypes.TwsOrderTypeMarketIfTouched
    gTwsOrderTypeToOrderType = OrderTypeMarketIfTouched
Case TwsOrderTypes.TwsOrderTypeTrailLimit
    gTwsOrderTypeToOrderType = OrderTypeTrailLimit
Case TwsOrderTypes.TwsOrderTypeMarketWithProtection
    gTwsOrderTypeToOrderType = OrderTypeMarketWithProtection
Case TwsOrderTypes.TwsOrderTypeMarketOnOpen
    gTwsOrderTypeToOrderType = OrderTypeMarketOnOpen
Case TwsOrderTypes.TwsOrderTypeLimitOnOpen
    gTwsOrderTypeToOrderType = OrderTypeLimitOnOpen
Case TwsOrderTypes.TwsOrderTypePeggedToPrimary
    gTwsOrderTypeToOrderType = OrderTypePeggedToPrimary
Case Else
    gTwsOrderTypeToOrderType = OrderTypeNone
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTwsSecTypeToSecType(ByVal Value As TwsSecTypes) As SecurityTypes
Select Case Value
Case TwsSecTypeNone
    gTwsSecTypeToSecType = SecurityTypes.SecTypeNone
Case TwsSecTypeStock
    gTwsSecTypeToSecType = SecurityTypes.SecTypeStock
Case TwsSecTypeFuture
    gTwsSecTypeToSecType = SecurityTypes.SecTypeFuture
Case TwsSecTypeOption
    gTwsSecTypeToSecType = SecurityTypes.SecTypeOption
Case TwsSecTypeFuturesOption
    gTwsSecTypeToSecType = SecurityTypes.SecTypeFuturesOption
Case TwsSecTypeCash
    gTwsSecTypeToSecType = SecurityTypes.SecTypeCash
Case TwsSecTypeCombo
    gTwsSecTypeToSecType = SecurityTypes.SecTypeCombo
Case TwsSecTypeIndex
    gTwsSecTypeToSecType = SecurityTypes.SecTypeIndex
Case Else
    AssertArgument False, "Value is not a valid Security Type"
End Select
End Function

Public Function gTwsTriggerMethodToStopTriggerMethod(ByVal Value As TwsStopTriggerMethods) As OrderStopTriggerMethods
Select Case Value
Case TwsStopTriggerDefault
    gTwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerDefault
Case TwsStopTriggerDoubleBidAsk
    gTwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk
Case TwsStopTriggerLast
    gTwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerLast
Case TwsStopTriggerDoubleLast
    gTwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerDoubleLast
Case TwsStopTriggerBidAsk
    gTwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerBidAsk
Case TwsStopTriggerLastOrBidAsk
    gTwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk
Case TwsStopTriggerMidPoint
    gTwsTriggerMethodToStopTriggerMethod = OrderStopTriggerMethods.OrderStopTriggerMidPoint
Case Else
    AssertArgument False, "Value is not a valid TwsStopTriggerMethod"
End Select
End Function
                
Public Function gTwsTimezoneNameToStandardTimeZoneName(ByVal pTimeZoneId As String) As String
Const ProcName As String = "gTwsTimezoneNameToStandardTimeZoneName"
On Error GoTo Err

Select Case pTimeZoneId
Case ""
    gTwsTimezoneNameToStandardTimeZoneName = ""
Case "AET"
    gTwsTimezoneNameToStandardTimeZoneName = "AUS Eastern Standard Time"
Case "Asia/Hong_Kong"
    gTwsTimezoneNameToStandardTimeZoneName = "China Standard Time"
Case "CST", "CTT"
    gTwsTimezoneNameToStandardTimeZoneName = "Central Standard Time"
Case "GMT", "GB"
    gTwsTimezoneNameToStandardTimeZoneName = "GMT Standard Time"
Case "EST", "EST5EDT"
    gTwsTimezoneNameToStandardTimeZoneName = "Eastern Standard Time"
Case "PST", "Pacific/Pitcairn"
    gTwsTimezoneNameToStandardTimeZoneName = "Pacific Standard Time"
Case "MET"
    gTwsTimezoneNameToStandardTimeZoneName = "Romance Standard Time"
Case Else
    gLog "Unrecognised timezone: " & pTimeZoneId, ModuleName, ProcName, , LogLevelSevere
    gTwsTimezoneNameToStandardTimeZoneName = ""
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================

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
Else
    lMarketPrice = IIf(pAction = OrderActionSell, _
                        pDataSource.CurrentTick(TickTypeAsk).Price, _
                        pDataSource.CurrentTick(TickTypeBid).Price)
End If

getMarketPrice = lMarketPrice

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
