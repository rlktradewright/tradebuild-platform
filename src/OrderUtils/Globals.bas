Attribute VB_Name = "Globals"
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

Public Const ProjectName                            As String = "OrderUtils27"
Private Const ModuleName                            As String = "Globals"

Public Const MaxCurrency                            As Currency = 922337203685477.5807@
Public Const MaxDoubleValue                         As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const MinDate                                As Double = -657434#

Private Const StrOrderTypeNone                      As String = ""
Private Const StrOrderTypeMarket                    As String = "Market"
Private Const StrOrderTypeMarketOnClose             As String = "Market on Close"
Private Const StrOrderTypeLimit                     As String = "Limit"
Private Const StrOrderTypeLimitOnClose              As String = "Limit on Close"
Private Const StrOrderTypePegMarket                 As String = "Peg to Market"
Private Const StrOrderTypeStop                      As String = "Stop"
Private Const StrOrderTypeStopLimit                 As String = "Stop Limit"
Private Const StrOrderTypeTrail                     As String = "Trailing Stop"
Private Const StrOrderTypeRelative                  As String = "Relative"
Private Const StrOrderTypeVWAP                      As String = "VWAP"
Private Const StrOrderTypeMarketToLimit             As String = "Market to Limit"
Private Const StrOrderTypeQuote                     As String = "Quote"
Private Const StrOrderTypeAutoStop                  As String = "Auto Stop"
Private Const StrOrderTypeAutoLimit                 As String = "Auto Limit"
Private Const StrOrderTypeAdjust                    As String = "Adjust"
Private Const StrOrderTypeAlert                     As String = "Alert"
Private Const StrOrderTypeLimitIfTouched            As String = "Limit if Touched"
Private Const StrOrderTypeMarketIfTouched           As String = "Market if Touched"
Private Const StrOrderTypeTrailLimit                As String = "Trail Limit"
Private Const StrOrderTypeMarketWithProtection      As String = "Market with Protection"
Private Const StrOrderTypeMarketOnOpen              As String = "Market on Open"
Private Const StrOrderTypeLimitOnOpen               As String = "Limit on Open"
Private Const StrOrderTypePeggedToPrimary           As String = "Pegged to Primary"
Private Const StrOrderTypeMidprice                  As String = "Mid-Price"

Public Const BalancingOrderContextName              As String = "$balancing"
Public Const RecoveryOrderContextName               As String = "$recovery"

Public Const OrderInfoDelete                        As String = "DELETE"
Public Const OrderInfoData                          As String = "DATA"
Public Const OrderInfoComment                       As String = "COMMENT"

Public Const ProviderPropertyOCAGroup               As String = "OCA group"



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

Public Function gGetContractName( _
                ByVal pContractSpec As IContractSpecifier) As String
gGetContractName = pContractSpec.LocalSymbol & "@" & pContractSpec.Exchange
End Function

Public Property Get gEntryOrderTypes() As OrderTypes()
Static s() As OrderTypes
Static sInitialised As Boolean

If Not sInitialised Then
    sInitialised = True
    ReDim s(13) As OrderTypes
    s(0) = OrderTypeLimit
    s(1) = OrderTypeLimitIfTouched
    s(2) = OrderTypeLimitOnClose
    s(3) = OrderTypeLimitOnOpen
    s(4) = OrderTypeMarket
    s(5) = OrderTypeMarketIfTouched
    s(6) = OrderTypeMarketOnClose
    s(7) = OrderTypeMarketOnOpen
    s(8) = OrderTypeMarketToLimit
    s(9) = OrderTypeStop
    s(10) = OrderTypeStopLimit
    s(11) = OrderTypeTrail
    s(12) = OrderTypeTrailLimit
    s(13) = OrderTypeMidprice
End If
gEntryOrderTypes = s
End Property

Public Property Get gStopLossOrderTypes() As OrderTypes()
Static s() As OrderTypes
Static sInitialised As Boolean

If Not sInitialised Then
    sInitialised = True
    ReDim s(3) As OrderTypes
    s(0) = OrderTypeStop
    s(1) = OrderTypeStopLimit
    s(2) = OrderTypeTrail
    s(3) = OrderTypeTrailLimit
End If
gStopLossOrderTypes = s
End Property

Public Property Get gTargetOrderTypes() As OrderTypes()
Static s() As OrderTypes
Static sInitialised As Boolean

If Not sInitialised Then
    sInitialised = True
    ReDim s(6) As OrderTypes
    s(0) = OrderTypeLimit
    s(1) = OrderTypeLimitIfTouched
    s(2) = OrderTypeLimitOnClose
    s(3) = OrderTypeLimitOnOpen
    s(4) = OrderTypeMarketIfTouched
    s(5) = OrderTypeMarketOnClose
    s(6) = OrderTypeMarketOnOpen
End If
gTargetOrderTypes = s
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gBracketOrderRoleToString(ByVal pOrderRole As BracketOrderRoles) As String
Const ProcName As String = "gBracketOrderRoleToString"
On Error GoTo Err

Select Case pOrderRole
Case BracketOrderRoleNone
    gBracketOrderRoleToString = "None"
Case BracketOrderRoleEntry
    gBracketOrderRoleToString = "Entry"
Case BracketOrderRoleStopLoss
    gBracketOrderRoleToString = "Stop-loss"
Case BracketOrderRoleTarget
    gBracketOrderRoleToString = "Target"
Case BracketOrderRoleCloseout
    gBracketOrderRoleToString = "Closeout"
Case Else
    AssertArgument False, "Invalid order role"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketOrderStateToString(ByVal pState As BracketOrderStates) As String
Const ProcName As String = "gBracketOrderRoleToString"
On Error GoTo Err

Select Case pState
Case BracketOrderStateCreated
    gBracketOrderStateToString = "Created"
Case BracketOrderStateSubmitted
    gBracketOrderStateToString = "Submitted"
Case BracketOrderStateCancelling
    gBracketOrderStateToString = "Cancelling"
Case BracketOrderStateClosingOut
    gBracketOrderStateToString = "Closing out"
Case BracketOrderStateClosed
    gBracketOrderStateToString = "Closed"
Case BracketOrderStateAwaitingOtherOrderCancel
    gBracketOrderStateToString = "Awaiting order cancellation"
Case Else
    gBracketOrderStateToString = "*Unknown*"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketOrderToString( _
                ByVal pBracketOrder As IBracketOrder) As String
Const ProcName As String = "gBracketOrderToString"
On Error GoTo Err

Dim s As String
s = gOrderActionToString(pBracketOrder.EntryOrder.Action) & " " & _
    pBracketOrder.EntryOrder.Quantity & " " & _
    gGetOrderTypeAndPricesString(pBracketOrder.EntryOrder, pBracketOrder.Contract)

s = s & "; "
If Not pBracketOrder.StopLossOrder Is Nothing Then
    s = s & _
        gGetOrderTypeAndPricesString(pBracketOrder.StopLossOrder, pBracketOrder.Contract)
End If

s = s & "; "
If Not pBracketOrder.TargetOrder Is Nothing Then
    s = s & _
        gGetOrderTypeAndPricesString(pBracketOrder.TargetOrder, pBracketOrder.Contract)
End If

gBracketOrderToString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateBracketProfitCalculator( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pDataSource As IMarketDataSource) As BracketProfitCalculator
Const ProcName As String = "gCreateBracketProfitCalculator"
On Error GoTo Err

Set gCreateBracketProfitCalculator = New BracketProfitCalculator
gCreateBracketProfitCalculator.Initialise pBracketOrder, pDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalculateOffsettedPrice( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pSecType As SecurityTypes, _
                ByVal pOrderAction As OrderActions, _
                ByVal pTickSize As Double, _
                ByVal pBidPrice As Double, _
                ByVal pAskPrice As Double, _
                ByVal pTradePrice As Double, _
                ByVal pOptionModelPrice As Double) As Double
Const ProcName As String = "gCalculateOffsettedPrice"
On Error GoTo Err

If pPriceSpec.PriceType = PriceValueTypeEntry Or _
        pPriceSpec.PriceType = PriceValueTypeNone Then
    gCalculateOffsettedPrice = MaxDoubleValue
    Exit Function
End If

If pPriceSpec.PriceType = PriceValueTypeValue Then
    pPriceSpec.CheckPriceValid pTickSize, pSecType
End If

Dim lRoundingMode As TickRoundingModes
If pOrderAction = OrderActionBuy Then
    lRoundingMode = TickRoundingModeUp
Else
    lRoundingMode = TickRoundingModeDown
End If

Dim lPrice As Double

Select Case pPriceSpec.PriceType
Case PriceValueTypeNone
    Assert False
Case PriceValueTypeValue
    lPrice = pPriceSpec.Price
Case PriceValueTypeAsk
    AssertArgument pAskPrice <> MaxDouble, "Ask price not available"
    lPrice = pAskPrice
Case PriceValueTypeBid
    AssertArgument pBidPrice <> MaxDouble, "Bid price not available"
    lPrice = pBidPrice
Case PriceValueTypeLast
    AssertArgument pTradePrice <> MaxDouble, "Trade price not available"
    lPrice = pTradePrice
Case PriceValueTypeModel
    AssertArgument pOptionModelPrice <> MaxDouble, "Model price not available"
    lPrice = pOptionModelPrice
Case PriceValueTypeEntry
    Assert False
Case PriceValueTypeMid
    AssertArgument pAskPrice <> MaxDouble, "Ask price not available"
    AssertArgument pBidPrice <> MaxDouble, "Bid price not available"
    lPrice = (pAskPrice + pBidPrice) / 2#
Case PriceValueTypeBidOrAsk
    If pOrderAction = OrderActionBuy Then
        AssertArgument pBidPrice <> MaxDouble, "Bid price not available"
        lPrice = pBidPrice
    Else
        AssertArgument pAskPrice <> MaxDouble, "Ask price not available"
        lPrice = pAskPrice
    End If
End Select

Dim lResult As Double

Select Case pPriceSpec.OffsetType
Dim lOffset As Double
Case PriceOffsetTypeNone
    lResult = lPrice
Case PriceOffsetTypeIncrement
    lOffset = pPriceSpec.Offset
Case PriceOffsetTypeNumberOfTicks
    lOffset = pPriceSpec.Offset * pTickSize
Case PriceOffsetTypeBidAskPercent
    AssertArgument pAskPrice <> MaxDouble, "Ask price not available"
    AssertArgument pBidPrice <> MaxDouble, "Bid price not available"
    lOffset = pPriceSpec.Offset * (pAskPrice - pBidPrice) / 100#
Case PriceOffsetTypePercent
    lOffset = pPriceSpec.Offset * lPrice / 100#
End Select

If Not pPriceSpec.UseCloseoutSemantics And _
    pPriceSpec.PriceType <> PriceValueTypeBidOrAsk _
Then
    lResult = lPrice + lOffset
ElseIf pOrderAction = OrderActionBuy Then
    lResult = lPrice + lOffset
Else
    lResult = lPrice - lOffset
End If

gCalculateOffsettedPrice = gRoundToTickBoundary( _
                                        lResult, _
                                        pTickSize, _
                                        lRoundingMode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateOptionRolloverSpecification( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pInitialStrikeSelectionMode As OptionStrikeSelectionModes, _
                ByVal pInitialStrikeParameter As Double, _
                ByVal pInitialStrikeOperator As OptionStrikeSelectionOperators, _
                ByVal pRolloverStrikeSelectionMode As RolloverStrikeModes, _
                ByVal pRolloverStrikeParameter As Double, _
                ByVal pRolloverStrikeOperator As OptionStrikeSelectionOperators, _
                ByVal pRolloverQuantityMode As RolloverQuantityModes, _
                ByVal pRolloverQuantityParameter As BoxedDecimal, _
                ByVal pRolloverQuantityLotSize As Long, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pCloseOrderType As OrderTypes, _
                ByVal pCloseTimeoutSecs As Long, _
                ByVal pCloseLimitPriceSpec As PriceSpecifier, _
                ByVal pCloseTriggerPriceSpec As PriceSpecifier, _
                ByVal pEntryOrderType As OrderTypes, _
                ByVal pEntryTimeoutSecs As Long, _
                ByVal pEntryLimitPriceSpec As PriceSpecifier, _
                ByVal pEntryTriggerPriceSpec As PriceSpecifier) As RolloverSpecification
Const ProcName As String = "gCreateOptionRolloverSpecification"
On Error GoTo Err

Dim lRolloverSpecification As New RolloverSpecification
Set gCreateOptionRolloverSpecification = lRolloverSpecification. _
            setDays(pDays). _
            setTime(pTime). _
            setInitialStrikeSelectionMode(pInitialStrikeSelectionMode). _
            setInitialStrikeParameter(pInitialStrikeParameter). _
            setInitialStrikeOperator(pInitialStrikeOperator). _
            setRolloverStrikeSelectionMode(pRolloverStrikeSelectionMode). _
            setRolloverStrikeParameter(pRolloverStrikeParameter). _
            setRolloverStrikeOperator(pRolloverStrikeOperator). _
            setRolloverQuantityMode(pRolloverQuantityMode). _
            setRolloverQuantityParameter(pRolloverQuantityParameter). _
            setRolloverQuantityLotSize(pRolloverQuantityLotSize). _
            setUnderlyingExchangeName(pUnderlyingExchangeName). _
            setCloseOrderType(pCloseOrderType). _
            setCloseTimeoutSecs(pCloseTimeoutSecs). _
            setCloseLimitPriceSpec(pCloseLimitPriceSpec). _
            setCloseTriggerPriceSpec(pCloseTriggerPriceSpec). _
            setEntryOrderType(pEntryOrderType). _
            setEntryTimeoutSecs(pEntryTimeoutSecs). _
            setEntryLimitPriceSpec(pEntryLimitPriceSpec). _
            setEntryTriggerPriceSpec(pEntryTriggerPriceSpec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateRolloverSpecification( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pCloseOrderType As OrderTypes, _
                ByVal pCloseTimeoutSecs As Long, _
                ByVal pCloseLimitPriceSpec As PriceSpecifier, _
                ByVal pCloseTriggerPriceSpec As PriceSpecifier, _
                ByVal pEntryOrderType As OrderTypes, _
                ByVal pEntryTimeoutSecs As Long, _
                ByVal pEntryLimitPriceSpec As PriceSpecifier, _
                ByVal pEntryTriggerPriceSpec As PriceSpecifier) As RolloverSpecification
Const ProcName As String = "gCreateRolloverSpecification"
On Error GoTo Err

Dim lRolloverSpecification As New RolloverSpecification
Set gCreateRolloverSpecification = lRolloverSpecification. _
            setDays(pDays). _
            setTime(pTime). _
            setCloseOrderType(pCloseOrderType). _
            setCloseTimeoutSecs(pCloseTimeoutSecs). _
            setCloseLimitPriceSpec(pCloseLimitPriceSpec). _
            setCloseTriggerPriceSpec(pCloseTriggerPriceSpec). _
            setEntryOrderType(pEntryOrderType). _
            setEntryTimeoutSecs(pEntryTimeoutSecs). _
            setEntryLimitPriceSpec(pEntryLimitPriceSpec). _
            setEntryTriggerPriceSpec(pEntryTriggerPriceSpec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetOptionContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pAction As OrderActions, _
                ByVal pContractStore As IContractStore, _
                ByVal pSelectionMode As OptionStrikeSelectionModes, _
                ByVal pParameter As Long, _
                ByVal pOperator As OptionStrikeSelectionOperators, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pListener As IStateChangeListener, _
                ByVal pReferenceDate As Date) As IFuture
Const ProcName As String = "gGetOptionContract"
On Error GoTo Err

Dim lContractResolver As New OptionContractResolver
Set gGetOptionContract = lContractResolver.ResolveContract( _
                                                pContractSpec, _
                                                pAction, _
                                                pContractStore, _
                                                pSelectionMode, _
                                                pParameter, _
                                                pOperator, _
                                                pUnderlyingExchangeName, _
                                                pMarketDataManager, _
                                                pListener, _
                                                pReferenceDate)


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetOrderTypeAndPricesString( _
                ByVal pOrder As IOrder, _
                ByVal pContract As IContract) As String
Const ProcName As String = "gGetOrderTypeAndPricesString"
On Error GoTo Err

Dim s As String
s = gOrderTypeToShortString(pOrder.OrderType)

Select Case pOrder.OrderType
Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeMarketToLimit, _
        OrderTypeLimitOnOpen
    s = s & " " & gPriceOrSpecifierToString( _
                                pOrder.LimitPrice, _
                                pOrder.LimitPriceSpec, _
                                pContract)
Case OrderTypeStop, _
        OrderTypeMarketIfTouched, _
        OrderTypeTrail
    s = s & " " & gPriceOrSpecifierToString( _
                                pOrder.TriggerPrice, _
                                pOrder.TriggerPriceSpec, _
                                pContract)
Case OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeTrailLimit
    s = s & " " & gPriceOrSpecifierToString( _
                                pOrder.LimitPrice, _
                                pOrder.LimitPriceSpec, _
                                pContract) & _
        " " & gPriceOrSpecifierToString( _
                                pOrder.TriggerPrice, _
                                pOrder.TriggerPriceSpec, _
                                pContract)
Case OrderTypeMarketWithProtection

Case OrderTypeMarketOnOpen

Case OrderTypePeggedToPrimary

End Select

gGetOrderTypeAndPricesString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function gGetSourceDesignator( _
                ByRef pModuleName As String, _
                ByRef pProcedureName As String) As String
gGetSourceDesignator = "[" & ProjectName & "." & _
            pModuleName & ":" & _
            pProcedureName & "]"
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

Public Function gIsEntryOrderType(ByVal pOrderType As OrderTypes) As Boolean
Select Case pOrderType
Case OrderTypeLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitOnOpen, _
        OrderTypeMarket, _
        OrderTypeMarketIfTouched, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen, _
        OrderTypeMarketToLimit, _
        OrderTypeStop, _
        OrderTypeStopLimit, _
        OrderTypeTrail, _
        OrderTypeTrailLimit, _
        OrderTypeMidprice
    gIsEntryOrderType = True
End Select
End Function

Public Function gIsNullPriceSpecifier(ByVal pPriceSpec As PriceSpecifier) As Boolean
If pPriceSpec Is Nothing Then
    gIsNullPriceSpecifier = True
Else
    gIsNullPriceSpecifier = pPriceSpec.Price = MaxDoubleValue And _
                            pPriceSpec.PriceType = PriceValueTypeNone And _
                            pPriceSpec.Offset = 0 And _
                            pPriceSpec.OffsetType = PriceOffsetTypeNone
End If
End Function

Public Function gIsStopLossOrderType(ByVal pOrderType As OrderTypes) As Boolean
Select Case pOrderType
Case OrderTypeStop, _
        OrderTypeStopLimit, _
        OrderTypeTrail, _
        OrderTypeTrailLimit
    gIsStopLossOrderType = True
End Select
End Function

Public Function gIsTargetOrderType(ByVal pOrderType As OrderTypes) As Boolean
Select Case pOrderType
Case OrderTypeLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitOnOpen, _
        OrderTypeMarketIfTouched, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen
gIsTargetOrderType = True
End Select
End Function

Public Function gIsValidTIF(ByVal Value As OrderTIFs) As Boolean
Const ProcName As String = "gIsValidTIF"
On Error GoTo Err

Select Case Value
Case OrderTIFDay
    gIsValidTIF = True
Case OrderTIFGoodTillCancelled
    gIsValidTIF = True
Case OrderTIFImmediateOrCancel
    gIsValidTIF = True
Case Else
    gIsValidTIF = False
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gLog(ByVal pMsg As String, _
                ByVal pProcName As String, _
                ByVal pModName As String, _
                Optional ByVal pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err

Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tradebuild.log.orderutils", ProjectName)

sLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogBracketOrderProfileObject( _
                ByVal pData As BracketOrderProfile, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogBracketOrderProfileObject"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "bracketorderprofilestruct", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "bracketorderprofilestruct", pData, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogBracketOrderProfileString( _
                ByVal pData As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogBracketOrderProfileString"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "bracketorderprofilestring", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "bracketorderprofilestring", pData, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogBracketOrderRollover( _
                ByVal pData As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogBracketOrderRollover"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "bracketorderrollover", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated, True
Else
    logInfotypeData "bracketorderrollover", pData, pSimulated, pSource, pLogLevel, sLogger, True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogContractResolution(ByVal pMsg As String, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err

Static sLogger As Logger
If sLogger Is Nothing Then Set sLogger = GetLogger("tradebuild.log.orderutils.contractresolution")

sLogger.Log pLogLevel, pMsg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogDrawDown( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogDrawDown"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "drawdown", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "drawdown", pData, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogMaxLoss( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogMaxLoss"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "maxloss", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "maxloss", pData, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogMaxProfit( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogMaxProfit"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "maxprofit", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "maxprofit", pData, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogMoneyManagement( _
                ByVal pMessage As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogMoneyManagement"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "moneymanagement", pMessage, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "moneymanagement", pMessage, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogOrder( _
                ByVal pMessage As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogOrder"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "order", pMessage, pSimulated, pSource, pLogLevel, sLoggerSimulated, True
Else
    logInfotypeData "order", pMessage, pSimulated, pSource, pLogLevel, sLogger, True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogOrderDetail( _
                ByVal pMessage As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogOrderDetail"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "orderdetail", pMessage, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "orderdetail", pMessage, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogBracketOrderMessage( _
                ByVal pMessage As String, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pContract As IContract, _
                ByVal pKey As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pSource As Object)
Const ProcName As String = "gLogBracketOrderMessage"
On Error GoTo Err

Dim lTickPart As String

If pDataSource Is Nothing Then
ElseIf pDataSource.State <> MarketDataSourceStateRunning Then
Else
    lTickPart = vbCrLf & "    " & GetCurrentTickSummary(pDataSource) & "; "
End If

gLogOrder IIf(pIsSimulated, "(simulated) ", "") & _
            pMessage & vbCrLf & _
            "    Contract: " & gGetContractName(pContract.Specifier) & _
            IIf(pKey <> "", vbCrLf & "    Bracket id: " & pKey, "") & _
            lTickPart, _
        pIsSimulated, _
        pSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogOrderMessage( _
                ByVal pMessage As String, _
                ByVal pOrder As IOrder, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pContract As IContract, _
                ByVal pKey As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pSource As Object)
Const ProcName As String = "gLogOrderMessage"
On Error GoTo Err

gLogBracketOrderMessage pMessage & vbCrLf & _
                        "    BrokerId: " & pOrder.BrokerId & _
                        "; system id: " & pOrder.Id, _
                        pDataSource, _
                        pContract, _
                        pKey, _
                        pIsSimulated, _
                        pSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogPosition( _
                ByVal pPosition As BoxedDecimal, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogPosition"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "position", pPosition, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "position", pPosition, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogProfit( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogProfit"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "profit", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "profit", pData, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogTradeProfile( _
                ByVal pData As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogTradeProfile"
On Error GoTo Err

Static sLogger As Logger
Static sLoggerSimulated As Logger

If pSimulated Then
    logInfotypeData "tradeprofile", pData, pSimulated, pSource, pLogLevel, sLoggerSimulated
Else
    logInfotypeData "tradeprofile", pData, pSimulated, pSource, pLogLevel, sLogger
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Function gOptionStrikeSelectionModeFromString( _
                ByVal Value As String) As OptionStrikeSelectionModes
Select Case UCase$(Value)
Case ""
    gOptionStrikeSelectionModeFromString = OptionStrikeSelectionModeNone
Case "I"
    gOptionStrikeSelectionModeFromString = OptionStrikeSelectionModeIncrement
Case "$"
    gOptionStrikeSelectionModeFromString = OptionStrikeSelectionModeExpenditure
Case "D"
    gOptionStrikeSelectionModeFromString = OptionStrikeSelectionModeDelta
Case Else
    AssertArgument False, "Value is not a valid option strike selection mode"
End Select
End Function

Public Function gOptionStrikeSelectionModeToString( _
                ByVal Value As OptionStrikeSelectionModes) As String
Select Case Value
Case OptionStrikeSelectionModeNone
    gOptionStrikeSelectionModeToString = ""
Case OptionStrikeSelectionModeIncrement
    gOptionStrikeSelectionModeToString = "I"
Case OptionStrikeSelectionModeExpenditure
    gOptionStrikeSelectionModeToString = "$"
Case OptionStrikeSelectionModeDelta
    gOptionStrikeSelectionModeToString = "D"
Case Else
    AssertArgument False, "Value is not a valid option strike selection mode"
End Select
End Function

Public Function gOptionStrikeSelectionOperatorFromString(ByVal Value As String) As OptionStrikeSelectionOperators
Select Case UCase$(Value)
Case ""
    gOptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorNone
Case "<"
    gOptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorLT
Case "<="
    gOptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorLE
Case ">"
    gOptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorGT
Case ">="
    gOptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorGE
Case Else
    AssertArgument False, "Value is not a valid option strike selection operator"
End Select
End Function

Public Function gOptionStrikeSelectionOperatorToString(ByVal Value As OptionStrikeSelectionOperators) As String
Select Case Value
Case OptionStrikeSelectionOperatorNone
    gOptionStrikeSelectionOperatorToString = ""
Case OptionStrikeSelectionOperatorLT
    gOptionStrikeSelectionOperatorToString = "<"
Case OptionStrikeSelectionOperatorLE
    gOptionStrikeSelectionOperatorToString = "<="
Case OptionStrikeSelectionOperatorGT
    gOptionStrikeSelectionOperatorToString = ">"
Case OptionStrikeSelectionOperatorGE
    gOptionStrikeSelectionOperatorToString = ">="
Case Else
    AssertArgument False, "Value is not a valid option strike selection operator"
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

Public Function gOrderActionToString(ByVal Value As OrderActions) As String
Select Case Value
Case OrderActionBuy
    gOrderActionToString = "BUY"
Case OrderActionSell
    gOrderActionToString = "SELL"
Case OrderActionNone
    gOrderActionToString = ""
Case Else
    AssertArgument False, "Value is not a valid Order Action"
End Select
End Function

Public Function gOrderAttributeToString(ByVal Value As OrderAttributes) As String
Const ProcName As String = "gOrderAttributeToString"
On Error GoTo Err

Select Case Value
    Case OrderAttOpenClose
        gOrderAttributeToString = "OpenClose"
    Case OrderAttOrigin
        gOrderAttributeToString = "Origin"
    Case OrderAttOriginatorRef
        gOrderAttributeToString = "OriginatorRef"
    Case OrderAttBlockOrder
        gOrderAttributeToString = "BlockOrder"
    Case OrderAttSweepToFill
        gOrderAttributeToString = "SweepToFill"
    Case OrderAttDisplaySize
        gOrderAttributeToString = "DisplaySize"
    Case OrderAttIgnoreRTH
        gOrderAttributeToString = "IgnoreRTH"
    Case OrderAttHidden
        gOrderAttributeToString = "Hidden"
    Case OrderAttDiscretionaryAmount
        gOrderAttributeToString = "DiscretionaryAmount"
    Case OrderAttGoodAfterTime
        gOrderAttributeToString = "GoodAfterTime"
    Case OrderAttGoodTillDate
        gOrderAttributeToString = "GoodTillDate"
    'Case OrderAttRTHOnly
    '    gOrderAttributeToString = "RTHOnly"
    Case OrderAttRule80A
        gOrderAttributeToString = "Rule80A"
    Case OrderAttSettlingFirm
        gOrderAttributeToString = "SettlingFirm"
    Case OrderAttAllOrNone
        gOrderAttributeToString = "AllOrNone"
    Case OrderAttMinimumQuantity
        gOrderAttributeToString = "MinimumQuantity"
    Case OrderAttPercentOffset
        gOrderAttributeToString = "PercentOffset"
    Case OrderAttOverrideConstraints
        gOrderAttributeToString = "OverrideConstraints"
    Case OrderAttAction
        gOrderAttributeToString = "Action"
    Case OrderAttLimitPrice
        gOrderAttributeToString = "LimitPrice"
    Case OrderAttOrderType
        gOrderAttributeToString = "OrderType"
    Case OrderAttQuantity
        gOrderAttributeToString = "Quantity"
    Case OrderAttTimeInForce
        gOrderAttributeToString = "TimeInForce"
    Case OrderAttTriggerPrice
        gOrderAttributeToString = "TriggerPrice"
    Case OrderAttGoodAfterTimeTZ
        gOrderAttributeToString = "GoodAfterTimeTZ"
    Case OrderAttGoodTillDateTZ
        gOrderAttributeToString = "GoodTillDateTZ"
    Case OrderAttStopTriggerMethod
        gOrderAttributeToString = "StopTriggerMethod"
    Case Else
        gOrderAttributeToString = "***Unknown order attribute***"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gOrderStatusToString(ByVal pOrderStatus As OrderStatuses) As String
Select Case pOrderStatus
Case OrderStatusCreated
    gOrderStatusToString = "Created"
Case OrderStatusRejected
    gOrderStatusToString = "Rejected"
Case OrderStatusPendingSubmit
    gOrderStatusToString = "Pending submit"
Case OrderStatusPreSubmitted
    gOrderStatusToString = "Pre submitted"
Case OrderStatusSubmitted
    gOrderStatusToString = "Submitted"
Case OrderStatusFilled
    gOrderStatusToString = "Filled"
Case OrderStatusCancelling
    gOrderStatusToString = "Cancelling"
Case OrderStatusCancelled
    gOrderStatusToString = "Cancelled"
Case Else
    AssertArgument False, "Value is not a valid Order Status"
End Select
End Function

Public Function gOrderStopTriggerMethodToString(ByVal Value As OrderStopTriggerMethods) As String
Select Case Value
Case OrderStopTriggerDefault
    gOrderStopTriggerMethodToString = "Default"
Case OrderStopTriggerDoubleBidAsk
    gOrderStopTriggerMethodToString = "Double Bid/Ask"
Case OrderStopTriggerLast
    gOrderStopTriggerMethodToString = "Last"
Case OrderStopTriggerDoubleLast
    gOrderStopTriggerMethodToString = "Double Last"
Case OrderStopTriggerBidAsk
    gOrderStopTriggerMethodToString = "Bid/Ask"
Case OrderStopTriggerLastOrBidAsk
    gOrderStopTriggerMethodToString = "Last or Bid/Ask"
Case OrderStopTriggerMidPoint
    gOrderStopTriggerMethodToString = "Midpoint"
Case Else
    AssertArgument False, "Value is not a valid Order Stop Trigger Method"
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

Public Function gOrderTypeFromString(ByVal Value As String) As OrderTypes
Const ProcName As String = "gOrderTypeFromString"
On Error GoTo Err

Static sTypes As SortedDictionary
If sTypes Is Nothing Then
    Set sTypes = CreateSortedDictionary(KeyTypeString)
    
    sTypes.Add OrderTypeNone, UCase$(StrOrderTypeNone)
    sTypes.Add OrderTypeMarket, UCase$(StrOrderTypeMarket)
    sTypes.Add OrderTypeMarketOnClose, UCase$(StrOrderTypeMarketOnClose)
    sTypes.Add OrderTypeLimit, UCase$(StrOrderTypeLimit)
    sTypes.Add OrderTypeLimitOnClose, UCase$(StrOrderTypeLimitOnClose)
    sTypes.Add OrderTypePeggedToMarket, UCase$(StrOrderTypePegMarket)
    sTypes.Add OrderTypeStop, UCase$(StrOrderTypeStop)
    sTypes.Add OrderTypeStopLimit, UCase$(StrOrderTypeStopLimit)
    sTypes.Add OrderTypeTrail, UCase$(StrOrderTypeTrail)
    sTypes.Add OrderTypeRelative, UCase$(StrOrderTypeRelative)
    sTypes.Add OrderTypeMarketToLimit, UCase$(StrOrderTypeMarketToLimit)
    sTypes.Add OrderTypeLimitIfTouched, UCase$(StrOrderTypeLimitIfTouched)
    sTypes.Add OrderTypeMarketIfTouched, UCase$(StrOrderTypeMarketIfTouched)
    sTypes.Add OrderTypeTrailLimit, UCase$(StrOrderTypeTrailLimit)
    sTypes.Add OrderTypeMarketWithProtection, UCase$(StrOrderTypeMarketWithProtection)
    sTypes.Add OrderTypeMarketOnOpen, UCase$(StrOrderTypeMarketOnOpen)
    sTypes.Add OrderTypeLimitOnOpen, UCase$(StrOrderTypeLimitOnOpen)
    sTypes.Add OrderTypePeggedToPrimary, UCase$(StrOrderTypePeggedToPrimary)
    sTypes.Add OrderTypeMidprice, UCase$(StrOrderTypeMidprice)

    sTypes.Add OrderTypes.OrderTypeMarket, "MKT"
    sTypes.Add OrderTypes.OrderTypeMarketOnClose, "MKTCLS"
    sTypes.Add OrderTypes.OrderTypeLimit, "LMT"
    sTypes.Add OrderTypes.OrderTypeLimitOnClose, "LMTCLS"
    sTypes.Add OrderTypes.OrderTypePeggedToMarket, "PEGMKT"
    sTypes.Add OrderTypes.OrderTypeStop, "STP"
    sTypes.Add OrderTypes.OrderTypeStopLimit, "STPLMT"
    sTypes.Add OrderTypes.OrderTypeTrail, "TRAIL"
    sTypes.Add OrderTypes.OrderTypeRelative, "REL"
    sTypes.Add OrderTypes.OrderTypeMarketToLimit, "MTL"
    sTypes.Add OrderTypes.OrderTypeLimitIfTouched, "LIT"
    sTypes.Add OrderTypes.OrderTypeMarketIfTouched, "MIT"
    sTypes.Add OrderTypes.OrderTypeTrailLimit, "TRAILLMT"
    sTypes.Add OrderTypes.OrderTypeMarketWithProtection, "MKTPROT"
    sTypes.Add OrderTypes.OrderTypeMarketOnOpen, "MOO"
    sTypes.Add OrderTypes.OrderTypeLimitOnOpen, "LOO"
    sTypes.Add OrderTypes.OrderTypePeggedToPrimary, "PEGPRI"
    sTypes.Add OrderTypes.OrderTypeMidprice, "MIDPRICE"
End If

Dim lOrderType As OrderTypes: lOrderType = OrderTypeNone
sTypes.TryItem UCase$(Value), lOrderType
gOrderTypeFromString = lOrderType

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gOrderTypeToString(ByVal Value As OrderTypes) As String
Const ProcName As String = "GOrderTypeToString"
On Error GoTo Err

Select Case Value
Case OrderTypeNone
    gOrderTypeToString = StrOrderTypeNone
Case OrderTypeMarket
    gOrderTypeToString = StrOrderTypeMarket
Case OrderTypeMarketOnClose
    gOrderTypeToString = StrOrderTypeMarketOnClose
Case OrderTypeLimit
    gOrderTypeToString = StrOrderTypeLimit
Case OrderTypeLimitOnClose
    gOrderTypeToString = StrOrderTypeLimitOnClose
Case OrderTypePeggedToMarket
    gOrderTypeToString = StrOrderTypePegMarket
Case OrderTypeStop
    gOrderTypeToString = StrOrderTypeStop
Case OrderTypeStopLimit
    gOrderTypeToString = StrOrderTypeStopLimit
Case OrderTypeTrail
    gOrderTypeToString = StrOrderTypeTrail
Case OrderTypeRelative
    gOrderTypeToString = StrOrderTypeRelative
Case OrderTypeMarketToLimit
    gOrderTypeToString = StrOrderTypeMarketToLimit
Case OrderTypeLimitIfTouched
    gOrderTypeToString = StrOrderTypeLimitIfTouched
Case OrderTypeMarketIfTouched
    gOrderTypeToString = StrOrderTypeMarketIfTouched
Case OrderTypeTrailLimit
    gOrderTypeToString = StrOrderTypeTrailLimit
Case OrderTypeMarketWithProtection
    gOrderTypeToString = StrOrderTypeMarketWithProtection
Case OrderTypeMarketOnOpen
    gOrderTypeToString = StrOrderTypeMarketOnOpen
Case OrderTypeLimitOnOpen
    gOrderTypeToString = StrOrderTypeLimitOnOpen
Case OrderTypePeggedToPrimary
    gOrderTypeToString = StrOrderTypePeggedToPrimary
Case OrderTypeMidprice
    gOrderTypeToString = StrOrderTypeMidprice
Case Else
    AssertArgument False, "Invalid order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gOrderTypeToShortString(ByVal Value As OrderTypes) As String
Const ProcName As String = "gOrderTypeToShortString"
On Error GoTo Err

Select Case Value
Case OrderTypes.OrderTypeNone
    gOrderTypeToShortString = ""
Case OrderTypes.OrderTypeMarket
    gOrderTypeToShortString = "MKT"
Case OrderTypes.OrderTypeMarketOnClose
    gOrderTypeToShortString = "MKTCLS"
Case OrderTypes.OrderTypeLimit
    gOrderTypeToShortString = "LMT"
Case OrderTypes.OrderTypeLimitOnClose
    gOrderTypeToShortString = "LMTCLS"
Case OrderTypes.OrderTypePeggedToMarket
    gOrderTypeToShortString = "PEGMKT"
Case OrderTypes.OrderTypeStop
    gOrderTypeToShortString = "STP"
Case OrderTypes.OrderTypeStopLimit
    gOrderTypeToShortString = "STPLMT"
Case OrderTypes.OrderTypeTrail
    gOrderTypeToShortString = "TRAIL"
Case OrderTypes.OrderTypeRelative
    gOrderTypeToShortString = "REL"
Case OrderTypes.OrderTypeMarketToLimit
    gOrderTypeToShortString = "MTL"
Case OrderTypes.OrderTypeLimitIfTouched
    gOrderTypeToShortString = "LIT"
Case OrderTypes.OrderTypeMarketIfTouched
    gOrderTypeToShortString = "MIT"
Case OrderTypes.OrderTypeTrailLimit
    gOrderTypeToShortString = "TRAILLMT"
Case OrderTypes.OrderTypeMarketWithProtection
    gOrderTypeToShortString = "MKTPROT"
Case OrderTypes.OrderTypeMarketOnOpen
    gOrderTypeToShortString = "MOO"
Case OrderTypes.OrderTypeLimitOnOpen
    gOrderTypeToShortString = "LOO"
Case OrderTypes.OrderTypePeggedToPrimary
    gOrderTypeToShortString = "PEGPRI"
Case OrderTypes.OrderTypeMidprice
    gOrderTypeToShortString = "MIDPRICE"
Case Else
    AssertArgument False, "Value is not a valid Order Type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Property Get gRegExp() As RegExp
Static lRegexp As RegExp
If lRegexp Is Nothing Then Set lRegexp = New RegExp
Set gRegExp = lRegexp
End Property

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

''
' Synchronises an order to the specified order so that both are
' identical.
'
' @param  pTargetOrder the <code>order</code> that is to be synchronized
' @param  pSourceOrder the <code>order</code> to which the target order must be made identical
'@/
Public Sub gSyncToOrder(ByVal pTargetOrder As Order, ByVal pSourceOrder As IOrder)
Const ProcName As String = "gSyncToOrder"
On Error GoTo Err

With pTargetOrder
    .Action = pSourceOrder.Action
    .LimitPrice = pSourceOrder.LimitPrice
    .LimitPriceSpec = pSourceOrder.LimitPriceSpec
    .TriggerPrice = pSourceOrder.TriggerPrice
    .TriggerPriceSpec = pSourceOrder.TriggerPriceSpec
    .IgnoreRegularTradingHours = pSourceOrder.IgnoreRegularTradingHours
    
    .AllOrNone = pSourceOrder.AllOrNone
    .AveragePrice = pSourceOrder.AveragePrice
    .BlockOrder = pSourceOrder.BlockOrder
    .BrokerId = pSourceOrder.BrokerId
    .DiscretionaryAmount = pSourceOrder.DiscretionaryAmount
    .DisplaySize = pSourceOrder.DisplaySize
    .ErrorCode = pSourceOrder.ErrorCode
    .ErrorMessage = pSourceOrder.ErrorMessage
    .FillTime = pSourceOrder.FillTime
    .GoodAfterTime = pSourceOrder.GoodAfterTime
    .GoodAfterTimeTZ = pSourceOrder.GoodAfterTimeTZ
    .GoodTillDate = pSourceOrder.GoodTillDate
    .GoodTillDateTZ = pSourceOrder.GoodTillDateTZ
    .Hidden = pSourceOrder.Hidden
    .IsSimulated = pSourceOrder.IsSimulated
    .LastFillPrice = pSourceOrder.LastFillPrice
    .MinimumQuantity = pSourceOrder.MinimumQuantity
    .OrderType = pSourceOrder.OrderType
    .Origin = pSourceOrder.Origin
    .OriginatorRef = pSourceOrder.OriginatorRef
    .OverrideConstraints = pSourceOrder.OverrideConstraints
    .PercentOffset = pSourceOrder.PercentOffset
    .Quantity = pSourceOrder.Quantity
    .QuantityFilled = pSourceOrder.QuantityFilled
    .QuantityRemaining = pSourceOrder.QuantityRemaining
    .SettlingFirm = pSourceOrder.SettlingFirm
    .StopTriggerMethod = pSourceOrder.StopTriggerMethod
    .SweepToFill = pSourceOrder.SweepToFill
    .TimeInForce = pSourceOrder.TimeInForce

    ' do this last to prevent status influencing whether attributes are modifiable
    .Status = pSourceOrder.Status
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gVariantEquals(ByVal p1 As Variant, ByVal p2 As Variant) As Boolean
If IsMissing(p2) Or IsEmpty(p2) Then
    gVariantEquals = False
ElseIf IsNumeric(p1) And IsNumeric(p2) Then
    gVariantEquals = (p1 = p2)
ElseIf IsArray(p1) Then
    gVariantEquals = False
ElseIf IsObject(p1) And IsObject(p2) Then
    gVariantEquals = (p1 Is p2)
Else
    gVariantEquals = (p1 = p2)
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub logInfotypeData( _
                ByVal pInfoType As String, _
                ByRef pData As Variant, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                ByVal pLogLevel As LogLevels, _
                ByRef pLogger As Logger, _
                Optional ByVal pLogToParent As Boolean = False)
Const ProcName As String = "logInfotypeData"
On Error GoTo Err

If pLogger Is Nothing Then
    Set pLogger = GetLogger("position." & pInfoType & IIf(pSimulated, "Simulated", ""))
    pLogger.LogToParent = pLogToParent
End If
pLogger.Log pLogLevel, pData, pSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




