Attribute VB_Name = "GOrderUtils"
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

Private Const ModuleName                            As String = "GOrderUtils"

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

Public Property Get EntryOrderTypes() As OrderTypes()
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
EntryOrderTypes = s
End Property

Public Property Get StopLossOrderTypes() As OrderTypes()
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
StopLossOrderTypes = s
End Property

''
' Synchronises an order to the specified order so that both are
' identical.
'
' @param  pTargetOrder the <code>order</code> that is to be synchronized
' @param  pSourceOrder the <code>order</code> to which the target order must be made identical
'@/
Public Sub SyncToOrder(ByVal pTargetOrder As Order, ByVal pSourceOrder As IOrder)
Const ProcName As String = "SyncToOrder"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Property Get TargetOrderTypes() As OrderTypes()
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
TargetOrderTypes = s
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function BracketOrderRoleToString(ByVal pOrderRole As BracketOrderRoles) As String
Const ProcName As String = "BracketOrderRoleToString"
On Error GoTo Err

Select Case pOrderRole
Case BracketOrderRoleNone
    BracketOrderRoleToString = "None"
Case BracketOrderRoleEntry
    BracketOrderRoleToString = "Entry"
Case BracketOrderRoleStopLoss
    BracketOrderRoleToString = "Stop-loss"
Case BracketOrderRoleTarget
    BracketOrderRoleToString = "Target"
Case BracketOrderRoleCloseout
    BracketOrderRoleToString = "Closeout"
Case Else
    AssertArgument False, "Invalid order role"
End Select

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketOrderStateToString(ByVal pState As BracketOrderStates) As String
Const ProcName As String = "BracketOrderStateToString"
On Error GoTo Err

Select Case pState
Case BracketOrderStateCreated
    BracketOrderStateToString = "Created"
Case BracketOrderStateSubmitted
    BracketOrderStateToString = "Submitted"
Case BracketOrderStateCancelling
    BracketOrderStateToString = "Cancelling"
Case BracketOrderStateClosingOut
    BracketOrderStateToString = "Closing out"
Case BracketOrderStateClosed
    BracketOrderStateToString = "Closed"
Case BracketOrderStateAwaitingOtherOrderCancel
    BracketOrderStateToString = "Awaiting order cancellation"
Case Else
    BracketOrderStateToString = "*Unknown*"
End Select

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketOrderToString( _
                ByVal pBracketOrder As IBracketOrder) As String
Const ProcName As String = "BracketOrderToString"
On Error GoTo Err

Dim s As String
s = OrderActionToString(pBracketOrder.EntryOrder.Action) & " " & _
    pBracketOrder.EntryOrder.Quantity & " " & _
    getOrderTypeAndPricesString(pBracketOrder.EntryOrder, pBracketOrder.Contract)

s = s & "; "
If Not pBracketOrder.StopLossOrder Is Nothing Then
    s = s & _
        getOrderTypeAndPricesString(pBracketOrder.StopLossOrder, pBracketOrder.Contract)
End If

s = s & "; "
If Not pBracketOrder.TargetOrder Is Nothing Then
    s = s & _
        getOrderTypeAndPricesString(pBracketOrder.TargetOrder, pBracketOrder.Contract)
End If

BracketOrderToString = s

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBracketProfitCalculator( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pDataSource As IMarketDataSource) As BracketProfitCalculator
Const ProcName As String = "CreateBracketProfitCalculator"
On Error GoTo Err

Set CreateBracketProfitCalculator = New BracketProfitCalculator
CreateBracketProfitCalculator.Initialise pBracketOrder, pDataSource

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateOrderPersistenceDataStore(ByVal pRecoveryFilePath As String) As IOrderPersistenceDataStore
Const ProcName As String = "CreateOrderPersistenceDataStore"
On Error GoTo Err

Dim lDataStore As OrderPersistenceDataStore
Set lDataStore = New OrderPersistenceDataStore
lDataStore.Initialise pRecoveryFilePath
Set CreateOrderPersistenceDataStore = lDataStore

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetContractName( _
                ByVal pContractSpec As IContractSpecifier) As String
GetContractName = pContractSpec.LocalSymbol & "@" & pContractSpec.Exchange
End Function

Public Function IsEntryOrderType(ByVal pOrderType As OrderTypes) As Boolean
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
    IsEntryOrderType = True
End Select
End Function

Public Function IsNullPriceSpecifier(ByVal pPriceSpec As PriceSpecifier) As Boolean
If pPriceSpec Is Nothing Then
    IsNullPriceSpecifier = True
Else
    IsNullPriceSpecifier = pPriceSpec.Price = GOrderUtils.MaxDoubleValue And _
                            pPriceSpec.PriceType = PriceValueTypeNone And _
                            pPriceSpec.Offset = 0 And _
                            pPriceSpec.OffsetType = PriceOffsetTypeNone
End If
End Function

Public Function IsStopLossOrderType(ByVal pOrderType As OrderTypes) As Boolean
Select Case pOrderType
Case OrderTypeStop, _
        OrderTypeStopLimit, _
        OrderTypeTrail, _
        OrderTypeTrailLimit
    IsStopLossOrderType = True
End Select
End Function

Public Function IsTargetOrderType(ByVal pOrderType As OrderTypes) As Boolean
Select Case pOrderType
Case OrderTypeLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitOnOpen, _
        OrderTypeMarketIfTouched, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen
IsTargetOrderType = True
End Select
End Function

Public Sub LogBracketOrderMessage( _
                ByVal pMessage As String, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pContract As IContract, _
                ByVal pKey As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pSource As Object)
Const ProcName As String = "LogBracketOrderMessage"
On Error GoTo Err

Dim lTickPart As String

If pDataSource Is Nothing Then
ElseIf pDataSource.State <> MarketDataSourceStateRunning Then
Else
    lTickPart = vbCrLf & "    " & GetCurrentTickSummary(pDataSource) & "; "
End If

LogOrder IIf(pIsSimulated, "(simulated) ", "") & _
            pMessage & vbCrLf & _
            "    Contract: " & GetContractName(pContract.Specifier) & _
            IIf(pKey <> "", vbCrLf & "    Bracket id: " & pKey, "") & _
            lTickPart, _
        pIsSimulated, _
        pSource

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogOrderMessage( _
                ByVal pMessage As String, _
                ByVal pOrder As IOrder, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pContract As IContract, _
                ByVal pKey As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pSource As Object)
Const ProcName As String = "LogOrderMessage"
On Error GoTo Err

LogBracketOrderMessage pMessage & vbCrLf & _
                        "    BrokerId: " & pOrder.BrokerId & _
                        "; system id: " & pOrder.Id, _
                        pDataSource, _
                        pContract, _
                        pKey, _
                        pIsSimulated, _
                        pSource

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogPosition( _
                ByVal pPosition As BoxedDecimal, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogPosition"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogProfit( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogProfit"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogTradeProfile( _
                ByVal pData As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogTradeProfile"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Log(ByVal pMsg As String, _
                ByVal pProcName As String, _
                ByVal pModName As String, _
                Optional ByVal pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "Log"
On Error GoTo Err

Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tradebuild.log.orderutils", ProjectName)

sLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogBracketOrderProfileObject( _
                ByVal pData As BracketOrderProfile, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogBracketOrderProfileObject"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogBracketOrderProfileString( _
                ByVal pData As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogBracketOrderProfileString"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogBracketOrderRollover( _
                ByVal pData As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogBracketOrderRollover"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogContractResolution(ByVal pMsg As String, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "Log"
On Error GoTo Err

Static sLogger As Logger
If sLogger Is Nothing Then Set sLogger = GetLogger("tradebuild.log.orderutils.contractresolution")

sLogger.Log pLogLevel, pMsg

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogDrawDown( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogDrawDown"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogMaxLoss( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogMaxLoss"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogMaxProfit( _
                ByVal pData As Currency, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogMaxProfit"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogMoneyManagement( _
                ByVal pMessage As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogMoneyManagement"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogOrder( _
                ByVal pMessage As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogOrder"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogOrderDetail( _
                ByVal pMessage As String, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogOrderDetail"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewPriceSpecifier( _
                Optional ByVal pPrice As Double = MaxDoubleValue, _
                Optional ByVal pPriceString As String = "", _
                Optional ByVal pPriceType As PriceValueTypes = PriceValueTypeNone, _
                Optional ByVal pOffset As Double = 0#, _
                Optional ByVal pOffsetType As PriceOffsetTypes = PriceOffsetTypeNone) As PriceSpecifier
Set NewPriceSpecifier = gNewPriceSpecifier(pPrice, pPriceString, pPriceType, pOffset, pOffsetType)
End Function

Public Function OptionStrikeSelectionModeFromString( _
                ByVal Value As String) As OptionStrikeSelectionModes
Select Case UCase$(Value)
Case ""
    OptionStrikeSelectionModeFromString = OptionStrikeSelectionModeNone
Case "I"
    OptionStrikeSelectionModeFromString = OptionStrikeSelectionModeIncrement
Case "$"
    OptionStrikeSelectionModeFromString = OptionStrikeSelectionModeExpenditure
Case "D"
    OptionStrikeSelectionModeFromString = OptionStrikeSelectionModeDelta
Case Else
    AssertArgument False, "Value is not a valid option strike selection mode"
End Select
End Function

Public Function OptionStrikeSelectionModeToString( _
                ByVal Value As OptionStrikeSelectionModes) As String
Select Case Value
Case OptionStrikeSelectionModeNone
    OptionStrikeSelectionModeToString = ""
Case OptionStrikeSelectionModeIncrement
    OptionStrikeSelectionModeToString = "I"
Case OptionStrikeSelectionModeExpenditure
    OptionStrikeSelectionModeToString = "$"
Case OptionStrikeSelectionModeDelta
    OptionStrikeSelectionModeToString = "D"
Case Else
    AssertArgument False, "Value is not a valid option strike selection mode"
End Select
End Function

Public Function OptionStrikeSelectionOperatorFromString(ByVal Value As String) As OptionStrikeSelectionOperators
Select Case UCase$(Value)
Case ""
    OptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorNone
Case "<"
    OptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorLT
Case "<="
    OptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorLE
Case ">"
    OptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorGT
Case ">="
    OptionStrikeSelectionOperatorFromString = OptionStrikeSelectionOperatorGE
Case Else
    AssertArgument False, "Value is not a valid option strike selection operator"
End Select
End Function

Public Function OptionStrikeSelectionOperatorToString(ByVal Value As OptionStrikeSelectionOperators) As String
Select Case Value
Case OptionStrikeSelectionOperatorNone
    OptionStrikeSelectionOperatorToString = ""
Case OptionStrikeSelectionOperatorLT
    OptionStrikeSelectionOperatorToString = "<"
Case OptionStrikeSelectionOperatorLE
    OptionStrikeSelectionOperatorToString = "<="
Case OptionStrikeSelectionOperatorGT
    OptionStrikeSelectionOperatorToString = ">"
Case OptionStrikeSelectionOperatorGE
    OptionStrikeSelectionOperatorToString = ">="
Case Else
    AssertArgument False, "Value is not a valid option strike selection operator"
End Select
End Function

Public Function OrderActionFromString(ByVal Value As String) As OrderActions
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

Public Function OrderActionToString(ByVal Value As OrderActions) As String
Select Case Value
Case OrderActionBuy
    OrderActionToString = "BUY"
Case OrderActionSell
    OrderActionToString = "SELL"
Case OrderActionNone
    OrderActionToString = ""
Case Else
    AssertArgument False, "Value is not a valid Order Action"
End Select
End Function

Public Function OrderTIFFromString(ByVal Value As String) As OrderTIFs
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

Public Function OrderTIFToString(ByVal Value As OrderTIFs) As String
Select Case Value
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

Public Function CreateOptionRolloverSpecification( _
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
Const ProcName As String = "CreateOptionRolloverSpecification"
On Error GoTo Err

Dim lRolloverSpecification As New RolloverSpecification
Set CreateOptionRolloverSpecification = lRolloverSpecification. _
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateRolloverSpecification( _
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
Const ProcName As String = "CreateRolloverSpecification"
On Error GoTo Err

Set CreateRolloverSpecification = CreateRolloverSpec( _
                                                pDays, _
                                                pTime, _
                                                pCloseOrderType, _
                                                pCloseTimeoutSecs, _
                                                pCloseLimitPriceSpec, _
                                                pCloseTriggerPriceSpec, _
                                                pEntryOrderType, _
                                                pEntryTimeoutSecs, _
                                                pEntryLimitPriceSpec, _
                                                pEntryTriggerPriceSpec)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GenerateBracketOrderId() As String
GenerateBracketOrderId = GIdProvider.gNextId
End Function

Public Function GetOptionContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pAction As OrderActions, _
                ByVal pContractStore As IContractStore, _
                ByVal pSelectionMode As OptionStrikeSelectionModes, _
                ByVal pParameter As Long, _
                ByVal pOperator As OptionStrikeSelectionOperators, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                Optional ByVal pListener As IStateChangeListener, _
                Optional ByVal pReferenceDate As Date = MinDate) As IFuture
Const ProcName As String = "GetOptionContract"
On Error GoTo Err

Dim lContractResolver As New OptionContractResolver
Set GetOptionContract = lContractResolver.ResolveContract( _
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeFromString(ByVal Value As String) As OrderTypes
Const ProcName As String = "OrderTypeFromString"
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
OrderTypeFromString = lOrderType

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeToString(ByVal Value As OrderTypes) As String
Const ProcName As String = "OrderTypeToString"
On Error GoTo Err

Select Case Value
Case OrderTypeNone
    OrderTypeToString = StrOrderTypeNone
Case OrderTypeMarket
    OrderTypeToString = StrOrderTypeMarket
Case OrderTypeMarketOnClose
    OrderTypeToString = StrOrderTypeMarketOnClose
Case OrderTypeLimit
    OrderTypeToString = StrOrderTypeLimit
Case OrderTypeLimitOnClose
    OrderTypeToString = StrOrderTypeLimitOnClose
Case OrderTypePeggedToMarket
    OrderTypeToString = StrOrderTypePegMarket
Case OrderTypeStop
    OrderTypeToString = StrOrderTypeStop
Case OrderTypeStopLimit
    OrderTypeToString = StrOrderTypeStopLimit
Case OrderTypeTrail
    OrderTypeToString = StrOrderTypeTrail
Case OrderTypeRelative
    OrderTypeToString = StrOrderTypeRelative
Case OrderTypeMarketToLimit
    OrderTypeToString = StrOrderTypeMarketToLimit
Case OrderTypeLimitIfTouched
    OrderTypeToString = StrOrderTypeLimitIfTouched
Case OrderTypeMarketIfTouched
    OrderTypeToString = StrOrderTypeMarketIfTouched
Case OrderTypeTrailLimit
    OrderTypeToString = StrOrderTypeTrailLimit
Case OrderTypeMarketWithProtection
    OrderTypeToString = StrOrderTypeMarketWithProtection
Case OrderTypeMarketOnOpen
    OrderTypeToString = StrOrderTypeMarketOnOpen
Case OrderTypeLimitOnOpen
    OrderTypeToString = StrOrderTypeLimitOnOpen
Case OrderTypePeggedToPrimary
    OrderTypeToString = StrOrderTypePeggedToPrimary
Case OrderTypeMidprice
    OrderTypeToString = StrOrderTypeMidprice
Case Else
    AssertArgument False, "Invalid order type"
End Select

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeToShortString(ByVal Value As OrderTypes) As String
Const ProcName As String = "OrderTypeToShortString"
On Error GoTo Err

Select Case Value
Case OrderTypes.OrderTypeNone
    OrderTypeToShortString = ""
Case OrderTypes.OrderTypeMarket
    OrderTypeToShortString = "MKT"
Case OrderTypes.OrderTypeMarketOnClose
    OrderTypeToShortString = "MKTCLS"
Case OrderTypes.OrderTypeLimit
    OrderTypeToShortString = "LMT"
Case OrderTypes.OrderTypeLimitOnClose
    OrderTypeToShortString = "LMTCLS"
Case OrderTypes.OrderTypePeggedToMarket
    OrderTypeToShortString = "PEGMKT"
Case OrderTypes.OrderTypeStop
    OrderTypeToShortString = "STP"
Case OrderTypes.OrderTypeStopLimit
    OrderTypeToShortString = "STPLMT"
Case OrderTypes.OrderTypeTrail
    OrderTypeToShortString = "TRAIL"
Case OrderTypes.OrderTypeRelative
    OrderTypeToShortString = "REL"
Case OrderTypes.OrderTypeMarketToLimit
    OrderTypeToShortString = "MTL"
Case OrderTypes.OrderTypeLimitIfTouched
    OrderTypeToShortString = "LIT"
Case OrderTypes.OrderTypeMarketIfTouched
    OrderTypeToShortString = "MIT"
Case OrderTypes.OrderTypeTrailLimit
    OrderTypeToShortString = "TRAILLMT"
Case OrderTypes.OrderTypeMarketWithProtection
    OrderTypeToShortString = "MKTPROT"
Case OrderTypes.OrderTypeMarketOnOpen
    OrderTypeToShortString = "MOO"
Case OrderTypes.OrderTypeLimitOnOpen
    OrderTypeToShortString = "LOO"
Case OrderTypes.OrderTypePeggedToPrimary
    OrderTypeToShortString = "PEGPRI"
Case OrderTypes.OrderTypeMidprice
    OrderTypeToShortString = "MIDPRICE"
Case Else
    AssertArgument False, "Value is not a valid Order Type"
End Select

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderStatusToString(ByVal pOrderStatus As OrderStatuses) As String
Select Case pOrderStatus
Case OrderStatusCreated
    OrderStatusToString = "Created"
Case OrderStatusRejected
    OrderStatusToString = "Rejected"
Case OrderStatusPendingSubmit
    OrderStatusToString = "Pending submit"
Case OrderStatusPreSubmitted
    OrderStatusToString = "Pre submitted"
Case OrderStatusSubmitted
    OrderStatusToString = "Submitted"
Case OrderStatusFilled
    OrderStatusToString = "Filled"
Case OrderStatusCancelling
    OrderStatusToString = "Cancelling"
Case OrderStatusCancelled
    OrderStatusToString = "Cancelled"
Case Else
    AssertArgument False, "Value is not a valid Order Status"
End Select
End Function

Public Function OrderStopTriggerMethodToString(ByVal Value As OrderStopTriggerMethods) As String
Select Case Value
Case OrderStopTriggerDefault
    OrderStopTriggerMethodToString = "Default"
Case OrderStopTriggerDoubleBidAsk
    OrderStopTriggerMethodToString = "Double Bid/Ask"
Case OrderStopTriggerLast
    OrderStopTriggerMethodToString = "Last"
Case OrderStopTriggerDoubleLast
    OrderStopTriggerMethodToString = "Double Last"
Case OrderStopTriggerBidAsk
    OrderStopTriggerMethodToString = "Bid/Ask"
Case OrderStopTriggerLastOrBidAsk
    OrderStopTriggerMethodToString = "Last or Bid/Ask"
Case OrderStopTriggerMidPoint
    OrderStopTriggerMethodToString = "Midpoint"
Case Else
    AssertArgument False, "Value is not a valid Order Stop Trigger Method"
End Select
End Function

Public Function ParsePriceAndOffset( _
                ByRef pPriceSpec As PriceSpecifier, _
                ByVal pValue As String, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByRef pMessage As String, _
                Optional ByVal pUseCloseoutSemantics As Boolean = False) As Boolean
Const ProcName As String = "ParsePriceAndOffset"
On Error GoTo Err

ParsePriceAndOffset = gParsePriceAndOffset(pPriceSpec, pValue, pSecType, pTickSize, pUseCloseoutSemantics, pMessage)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function


Public Function PriceOffsetToString( _
                ByVal pOffset As Double, _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "PriceOffsetToString"
On Error GoTo Err

PriceOffsetToString = gPriceOffsetToString(pOffset, pOffsetType)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function PriceOffsetTypeToString( _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "PriceOffsetTypeToString"
On Error GoTo Err

PriceOffsetTypeToString = gPriceOffsetTypeToString(pOffsetType)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function PriceSpecifierToString( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pContract As IContract)
Const ProcName As String = "PriceSpecifierToString"
On Error GoTo Err

PriceSpecifierToString = gPriceSpecifierToString(pPriceSpec, _
                            pContract)
                            

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function TypedPriceToString( _
                ByVal pPrice As Double, _
                ByVal pPriceType As PriceValueTypes, _
                ByVal pContract As IContract) As String
Const ProcName As String = "TypedPriceToString"
On Error GoTo Err

TypedPriceToString = gTypedPriceToString(pPrice, pPriceType, pContract)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CreateRolloverSpec( _
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
Const ProcName As String = "CreateRolloverSpec"
On Error GoTo Err

Dim lRolloverSpecification As New RolloverSpecification
Set CreateRolloverSpec = lRolloverSpecification. _
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getOrderTypeAndPricesString( _
                ByVal pOrder As IOrder, _
                ByVal pContract As IContract) As String
Const ProcName As String = "getOrderTypeAndPricesString"
On Error GoTo Err

Dim s As String
s = OrderTypeToShortString(pOrder.OrderType)

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

getOrderTypeAndPricesString = s

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName

End Function

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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub







