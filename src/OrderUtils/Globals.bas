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

Private Const StrOrderTypeNone               As String = ""
Private Const StrOrderTypeMarket             As String = "Market"
Private Const StrOrderTypeMarketOnClose      As String = "Market on Close"
Private Const StrOrderTypeLimit              As String = "Limit"
Private Const StrOrderTypeLimitOnClose       As String = "Limit on Close"
Private Const StrOrderTypePegMarket          As String = "Peg to Market"
Private Const StrOrderTypeStop               As String = "Stop"
Private Const StrOrderTypeStopLimit          As String = "Stop Limit"
Private Const StrOrderTypeTrail              As String = "Trailing Stop"
Private Const StrOrderTypeRelative           As String = "Relative"
Private Const StrOrderTypeVWAP               As String = "VWAP"
Private Const StrOrderTypeMarketToLimit      As String = "Market to Limit"
Private Const StrOrderTypeQuote              As String = "Quote"
Private Const StrOrderTypeAutoStop           As String = "Auto Stop"
Private Const StrOrderTypeAutoLimit          As String = "Auto Limit"
Private Const StrOrderTypeAdjust             As String = "Adjust"
Private Const StrOrderTypeAlert              As String = "Alert"
Private Const StrOrderTypeLimitIfTouched     As String = "Limit if Touched"
Private Const StrOrderTypeMarketIfTouched    As String = "Market if Touched"
Private Const StrOrderTypeTrailLimit         As String = "Trail Limit"
Private Const StrOrderTypeMarketWithProtection As String = "Market with Protection"
Private Const StrOrderTypeMarketOnOpen       As String = "Market on Open"
Private Const StrOrderTypeLimitOnOpen        As String = "Limit on Open"
Private Const StrOrderTypePeggedToPrimary    As String = "Pegged to Primary"

Public Const BalancingOrderContextName              As String = "$balancing"
Public Const RecoveryOrderContextName               As String = "$recovery"

Public Const OrderInfoDelete                        As String = "DELETE"
Public Const OrderInfoData                          As String = "DATA"



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

Public Function gBracketEntryTypeFromString(ByVal Value As String) As BracketEntryTypes
Const ProcName As String = "gBracketEntryTypeFromString"
On Error GoTo Err

Select Case UCase$(Value)
Case "", "NONE"
    gBracketEntryTypeFromString = BracketEntryTypeNone
Case "MKT", "MARKET"
    gBracketEntryTypeFromString = BracketEntryTypeMarket
Case "MOO", "MARKET ON OPEN"
    gBracketEntryTypeFromString = BracketEntryTypeMarketOnOpen
Case "MOC", "MARKET ON CLOSE"
    gBracketEntryTypeFromString = BracketEntryTypeMarketOnClose
Case "MIT", "MARKET IF TOUCHED"
    gBracketEntryTypeFromString = BracketEntryTypeMarketIfTouched
Case "MTL", "MARKET TO LIMIT"
    gBracketEntryTypeFromString = BracketEntryTypeMarketToLimit
Case "BID", "BID PRICE"
    gBracketEntryTypeFromString = BracketEntryTypeBid
Case "ASK", "ASK PRICE"
    gBracketEntryTypeFromString = BracketEntryTypeAsk
Case "LAST", "LAST TRADE PRICE"
    gBracketEntryTypeFromString = BracketEntryTypeLast
Case "LMT", "LIMIT"
    gBracketEntryTypeFromString = BracketEntryTypeLimit
Case "LOO", "LIMIT ON OPEN"
    gBracketEntryTypeFromString = BracketEntryTypeLimitOnOpen
Case "LOC", "LIMIT ON CLOSE"
    gBracketEntryTypeFromString = BracketEntryTypeLimitOnClose
Case "LIT", "LIMIT IF TOUCHED"
    gBracketEntryTypeFromString = BracketEntryTypeLimitIfTouched
Case "STP", "STOP"
    gBracketEntryTypeFromString = BracketEntryTypeStop
Case "STPLMT", "STOP LIMIT"
    gBracketEntryTypeFromString = BracketEntryTypeStopLimit
Case Else
    AssertArgument False, "Invalid entry order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketEntryTypeToOrderType( _
                ByVal pBracketEntryType As BracketEntryTypes) As OrderTypes
Const ProcName As String = "gBracketEntryTypeToOrderType"
On Error GoTo Err

Select Case pBracketEntryType
Case BracketEntryTypeNone
    gBracketEntryTypeToOrderType = OrderTypeNone
Case BracketEntryTypeMarket
    gBracketEntryTypeToOrderType = OrderTypeMarket
Case BracketEntryTypeMarketOnOpen
    gBracketEntryTypeToOrderType = OrderTypeMarketOnOpen
Case BracketEntryTypeMarketOnClose
    gBracketEntryTypeToOrderType = OrderTypeMarketOnClose
Case BracketEntryTypeMarketIfTouched
    gBracketEntryTypeToOrderType = OrderTypeMarketIfTouched
Case BracketEntryTypeMarketToLimit
    gBracketEntryTypeToOrderType = OrderTypeMarketToLimit
Case BracketEntryTypeBid
    gBracketEntryTypeToOrderType = OrderTypeLimit
Case BracketEntryTypeAsk
    gBracketEntryTypeToOrderType = OrderTypeLimit
Case BracketEntryTypeLast
    gBracketEntryTypeToOrderType = OrderTypeLimit
Case BracketEntryTypeLimit
    gBracketEntryTypeToOrderType = OrderTypeLimit
Case BracketEntryTypeLimitOnOpen
    gBracketEntryTypeToOrderType = OrderTypeLimitOnOpen
Case BracketEntryTypeLimitOnClose
    gBracketEntryTypeToOrderType = OrderTypeLimitOnClose
Case BracketEntryTypeLimitIfTouched
    gBracketEntryTypeToOrderType = OrderTypeLimitIfTouched
Case BracketEntryTypeStop
    gBracketEntryTypeToOrderType = OrderTypeStop
Case BracketEntryTypeStopLimit
    gBracketEntryTypeToOrderType = OrderTypeStopLimit
Case Else
    AssertArgument False, "Invalid entry order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketEntryTypeToString(ByVal Value As BracketEntryTypes) As String
Const ProcName As String = "gBracketEntryTypeToString"
On Error GoTo Err

Select Case Value
Case BracketEntryTypeNone
    gBracketEntryTypeToString = ""
Case BracketEntryTypeMarket
    gBracketEntryTypeToString = "Market"
Case BracketEntryTypeMarketOnOpen
    gBracketEntryTypeToString = "Market on open"
Case BracketEntryTypeMarketOnClose
    gBracketEntryTypeToString = "Market on close"
Case BracketEntryTypeMarketIfTouched
    gBracketEntryTypeToString = "Market if touched"
Case BracketEntryTypeMarketToLimit
    gBracketEntryTypeToString = "Market to limit"
Case BracketEntryTypeBid
    gBracketEntryTypeToString = "Bid Price"
Case BracketEntryTypeAsk
    gBracketEntryTypeToString = "Ask Price"
Case BracketEntryTypeLast
    gBracketEntryTypeToString = "Last Trade Price"
Case BracketEntryTypeLimit
    gBracketEntryTypeToString = "Limit"
Case BracketEntryTypeLimitOnOpen
    gBracketEntryTypeToString = "Limit on open"
Case BracketEntryTypeLimitOnClose
    gBracketEntryTypeToString = "Limit on close"
Case BracketEntryTypeLimitIfTouched
    gBracketEntryTypeToString = "Limit if touched"
Case BracketEntryTypeStop
    gBracketEntryTypeToString = "Stop"
Case BracketEntryTypeStopLimit
    gBracketEntryTypeToString = "Stop limit"
Case Else
    AssertArgument False, "Invalid entry order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketEntryTypeToShortString(ByVal Value As BracketEntryTypes) As String
Const ProcName As String = "gBracketEntryTypeToShortString"
On Error GoTo Err

Select Case Value
Case BracketEntryTypeNone
    gBracketEntryTypeToShortString = ""
Case BracketEntryTypeMarket
    gBracketEntryTypeToShortString = "MKT"
Case BracketEntryTypeMarketOnOpen
    gBracketEntryTypeToShortString = "MOO"
Case BracketEntryTypeMarketOnClose
    gBracketEntryTypeToShortString = "MOC"
Case BracketEntryTypeMarketIfTouched
    gBracketEntryTypeToShortString = "MIT"
Case BracketEntryTypeMarketToLimit
    gBracketEntryTypeToShortString = "MTL"
Case BracketEntryTypeBid
    gBracketEntryTypeToShortString = "BID"
Case BracketEntryTypeAsk
    gBracketEntryTypeToShortString = "ASK"
Case BracketEntryTypeLast
    gBracketEntryTypeToShortString = "LAST"
Case BracketEntryTypeLimit
    gBracketEntryTypeToShortString = "LMT"
Case BracketEntryTypeLimitOnOpen
    gBracketEntryTypeToShortString = "LOO"
Case BracketEntryTypeLimitOnClose
    gBracketEntryTypeToShortString = "LOC"
Case BracketEntryTypeLimitIfTouched
    gBracketEntryTypeToShortString = "LIT"
Case BracketEntryTypeStop
    gBracketEntryTypeToShortString = "STP"
Case BracketEntryTypeStopLimit
    gBracketEntryTypeToShortString = "STPLMT"
Case Else
    AssertArgument False, "Invalid entry order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Public Function gBracketOrderToStringWithOffsets( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pIncludeOffsets As Boolean) As String
Const ProcName As String = "gBracketOrderToStringWithOffsets"
On Error GoTo Err

Dim s As String
s = gOrderActionToString(pBracketOrder.EntryOrder.Action) & " " & _
    pBracketOrder.EntryOrder.Quantity & " " & _
    gGetOrderTypeAndPricesString(pBracketOrder.EntryOrder, pBracketOrder.Contract, pIncludeOffsets)

s = s & "; "
If Not pBracketOrder.StopLossOrder Is Nothing Then
    s = s & _
        gGetOrderTypeAndPricesString(pBracketOrder.StopLossOrder, pBracketOrder.Contract, pIncludeOffsets)
End If

s = s & "; "
If Not pBracketOrder.TargetOrder Is Nothing Then
    s = s & _
        gGetOrderTypeAndPricesString(pBracketOrder.TargetOrder, pBracketOrder.Contract, pIncludeOffsets)
End If

gBracketOrderToStringWithOffsets = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketStopLossTypeFromString(ByVal Value As String) As BracketStopLossTypes
Const ProcName As String = "gBracketStopLossTypeToString"
On Error GoTo Err

Select Case UCase$(Value)
Case "", "NONE"
    gBracketStopLossTypeFromString = BracketStopLossTypeNone
Case "STP", "STOP"
    gBracketStopLossTypeFromString = BracketStopLossTypeStop
Case "STPLMT", "STOP LIMIT"
    gBracketStopLossTypeFromString = BracketStopLossTypeStopLimit
Case "BID", "BID PRICE"
    gBracketStopLossTypeFromString = BracketStopLossTypeBid
Case "ASK", "ASK PRICE"
    gBracketStopLossTypeFromString = BracketStopLossTypeAsk
Case "TRADE", "LAST TRADE PRICE"
    gBracketStopLossTypeFromString = BracketStopLossTypeLast
Case "TRAIL"
    gBracketStopLossTypeFromString = BracketStopLossTypeTrail
Case "TRAILLMT"
    gBracketStopLossTypeFromString = BracketStopLossTypeTrailLimit
Case "AUTO"
    gBracketStopLossTypeFromString = BracketStopLossTypeAuto
Case Else
    AssertArgument False, "Invalid stoploss order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketStopLossTypeToOrderType( _
                ByVal pBracketStopLossType As BracketStopLossTypes) As OrderTypes
Const ProcName As String = "gBracketStopLossTypeToOrderType"
On Error GoTo Err

Select Case pBracketStopLossType
Case BracketStopLossTypeNone
    gBracketStopLossTypeToOrderType = OrderTypeNone
Case BracketStopLossTypeStop
    gBracketStopLossTypeToOrderType = OrderTypeStop
Case BracketStopLossTypeStopLimit
    gBracketStopLossTypeToOrderType = OrderTypeStopLimit
Case BracketStopLossTypeTrail
    gBracketStopLossTypeToOrderType = OrderTypeTrail
Case BracketStopLossTypeTrailLimit
    gBracketStopLossTypeToOrderType = OrderTypeTrailLimit
Case BracketStopLossTypeBid
    gBracketStopLossTypeToOrderType = OrderTypeStop
Case BracketStopLossTypeAsk
    gBracketStopLossTypeToOrderType = OrderTypeStop
Case BracketStopLossTypeLast
    gBracketStopLossTypeToOrderType = OrderTypeStop
Case BracketStopLossTypeAuto
    gBracketStopLossTypeToOrderType = OrderTypeAutoStop
Case Else
    AssertArgument False, "Invalid stoploss order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketStopLossTypeToShortString(ByVal Value As BracketStopLossTypes) As String
Const ProcName As String = "gBracketStopLossTypeToShortString"
On Error GoTo Err

Select Case Value
Case BracketStopLossTypeNone
    gBracketStopLossTypeToShortString = "NONE"
Case BracketStopLossTypeStop
    gBracketStopLossTypeToShortString = "STP"
Case BracketStopLossTypeStopLimit
    gBracketStopLossTypeToShortString = "STPLMT"
Case BracketStopLossTypeBid
    gBracketStopLossTypeToShortString = "BID"
Case BracketStopLossTypeAsk
    gBracketStopLossTypeToShortString = "ASK"
Case BracketStopLossTypeLast
    gBracketStopLossTypeToShortString = "TRADE"
Case BracketStopLossTypeAuto
    gBracketStopLossTypeToShortString = "AUTO"
Case Else
    AssertArgument False, "Invalid stoploss order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketStopLossTypeToString(ByVal Value As BracketStopLossTypes)
Const ProcName As String = "gBracketStopLossTypeToString"
On Error GoTo Err

Select Case Value
Case BracketStopLossTypeNone
    gBracketStopLossTypeToString = "None"
Case BracketStopLossTypeStop
    gBracketStopLossTypeToString = "Stop"
Case BracketStopLossTypeStopLimit
    gBracketStopLossTypeToString = "Stop limit"
Case BracketStopLossTypeBid
    gBracketStopLossTypeToString = "Bid Price"
Case BracketStopLossTypeAsk
    gBracketStopLossTypeToString = "Ask Price"
Case BracketStopLossTypeLast
    gBracketStopLossTypeToString = "Last Trade Price"
Case BracketStopLossTypeAuto
    gBracketStopLossTypeToString = "Auto"
Case Else
    AssertArgument False, "Invalid stoploss order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketTargetTypeFromString(ByVal Value As String) As BracketTargetTypes
Const ProcName As String = "gBracketTargetTypeFromString"
On Error GoTo Err

Select Case UCase$(Value)
Case "NONE"
    gBracketTargetTypeFromString = BracketTargetTypeNone
Case "LMT", "LIMIT"
    gBracketTargetTypeFromString = BracketTargetTypeLimit
Case "LIT", "LIMIT IF TOUCHED"
    gBracketTargetTypeFromString = BracketTargetTypeLimitIfTouched
Case "MIT", "MARKET IF TOUCHED"
    gBracketTargetTypeFromString = BracketTargetTypeMarketIfTouched
Case "BID", "BID PRICE"
    gBracketTargetTypeFromString = BracketTargetTypeBid
Case "ASK", "ASK PRICE"
    gBracketTargetTypeFromString = BracketTargetTypeAsk
Case "LAST", "LAST TRADE PRICE"
    gBracketTargetTypeFromString = BracketTargetTypeLast
Case "AUTO"
    gBracketTargetTypeFromString = BracketTargetTypeAuto
Case Else
    AssertArgument False, "Invalid target order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a member of the BracketTargetTypes enumeration to the equivalent OrderTypes Value.
'
' @return           the OrderTypes Value corresponding to the parameter
' @param pBracketTargetType the BracketTargetTypes Value to be converted
' @ see
'
'@/
Public Function gBracketTargetTypeToOrderType( _
                ByVal pBracketTargetType As BracketTargetTypes) As OrderTypes
Const ProcName As String = "gBracketTargetTypeToOrderType"
On Error GoTo Err

Select Case pBracketTargetType
Case BracketTargetTypeNone
    gBracketTargetTypeToOrderType = OrderTypeNone
Case BracketTargetTypeLimit
    gBracketTargetTypeToOrderType = OrderTypeLimit
Case BracketTargetTypeLimitIfTouched
    gBracketTargetTypeToOrderType = OrderTypeLimitIfTouched
Case BracketTargetTypeMarketIfTouched
    gBracketTargetTypeToOrderType = OrderTypeMarketIfTouched
Case BracketTargetTypeBid
    gBracketTargetTypeToOrderType = OrderTypeLimit
Case BracketTargetTypeAsk
    gBracketTargetTypeToOrderType = OrderTypeLimit
Case BracketTargetTypeLast
    gBracketTargetTypeToOrderType = OrderTypeLimit
Case BracketTargetTypeAuto
    gBracketTargetTypeToOrderType = OrderTypeAutoLimit
Case Else
    AssertArgument False, "Invalid target order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketTargetTypeToShortString(ByVal Value As BracketTargetTypes) As String
Const ProcName As String = "gBracketTargetTypeToShortString"
On Error GoTo Err

Select Case Value
Case BracketTargetTypeNone
    gBracketTargetTypeToShortString = "NONE"
Case BracketTargetTypeLimit
    gBracketTargetTypeToShortString = "LMT"
Case BracketTargetTypeMarketIfTouched
    gBracketTargetTypeToShortString = "MIT"
Case BracketTargetTypeBid
    gBracketTargetTypeToShortString = "BID"
Case BracketTargetTypeAsk
    gBracketTargetTypeToShortString = "ASK"
Case BracketTargetTypeLast
    gBracketTargetTypeToShortString = "LAST"
Case BracketTargetTypeAuto
    gBracketTargetTypeToShortString = "AUTO"
Case Else
    AssertArgument False, "Invalid target order type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketTargetTypeToString(ByVal Value As BracketTargetTypes)
Const ProcName As String = "gBracketTargetTypeToString"
On Error GoTo Err

Select Case Value
Case BracketTargetTypeNone
    gBracketTargetTypeToString = "None"
Case BracketTargetTypeLimit
    gBracketTargetTypeToString = "Limit"
Case BracketTargetTypeMarketIfTouched
    gBracketTargetTypeToString = "Market if touched"
Case BracketTargetTypeBid
    gBracketTargetTypeToString = "Bid Price"
Case BracketTargetTypeAsk
    gBracketTargetTypeToString = "Ask Price"
Case BracketTargetTypeLast
    gBracketTargetTypeToString = "Last Trade Price"
Case BracketTargetTypeAuto
    gBracketTargetTypeToString = "Auto"
Case Else
    AssertArgument False, "Invalid target order type"
End Select

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

Public Function gGetOrderTypeAndPricesString( _
                ByVal pOrder As IOrder, _
                ByVal pContract As IContract, _
                Optional ByVal pIncludeOffsets As Boolean = False) As String
Const ProcName As String = "gGetOrderTypeAndPricesString"
On Error GoTo Err

Dim s As String
s = gOrderTypeToShortString(pOrder.OrderType)

Select Case pOrder.OrderType
Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeMarketToLimit, _
        OrderTypeLimitOnOpen
    s = s & " " & gPriceToString(pOrder.LimitPrice, _
                                pContract, _
                                pOrder.LimitPriceOffset, _
                                pOrder.LimitPriceOffsetType, _
                                pIncludeOffsets)
Case OrderTypeStop, _
        OrderTypeMarketIfTouched
    s = s & " " & gPriceToString(pOrder.TriggerPrice, _
                                pContract, _
                                pOrder.TriggerPriceOffset, _
                                pOrder.TriggerPriceOffset, _
                                pIncludeOffsets)
Case OrderTypeStopLimit, _
        OrderTypeLimitIfTouched
    s = s & " " & gPriceToString(pOrder.LimitPrice, _
                                pContract, _
                                pOrder.LimitPriceOffset, _
                                pOrder.LimitPriceOffsetType, _
                                pIncludeOffsets) & _
        " " & gPriceToString(pOrder.TriggerPrice, _
                                pContract, _
                                pOrder.TriggerPriceOffset, _
                                pOrder.TriggerPriceOffset, _
                                pIncludeOffsets)
Case OrderTypeTrail

Case OrderTypeRelative

Case OrderTypeVWAP

Case OrderTypeQuote

Case OrderTypeAutoStop
    s = s & " " & gPriceToString(MaxDouble, _
                                pContract, _
                                pOrder.TriggerPriceOffset, _
                                pOrder.TriggerPriceOffsetType, _
                                True)
Case OrderTypeAutoLimit
    s = s & " " & gPriceToString(MaxDouble, _
                                pContract, _
                                pOrder.LimitPriceOffset, _
                                pOrder.LimitPriceOffsetType, _
                                True)
Case OrderTypeAdjust

Case OrderTypeAlert

Case OrderTypeTrailLimit

Case OrderTypeMarketWithProtection

Case OrderTypeMarketOnOpen

Case OrderTypePeggedToPrimary

Case BracketEntryTypeAsk, _
        BracketEntryTypeBid, _
        BracketEntryTypeLast, _
        BracketStopLossTypeAsk, _
        BracketStopLossTypeBid, _
        BracketStopLossTypeLast, _
        BracketTargetTypeAsk, _
        BracketTargetTypeBid, _
        BracketTargetTypeLast
    If pIncludeOffsets Then
        s = s & _
            gPriceToString(MaxDouble, _
                            pContract, _
                            pOrder.LimitPriceOffset, _
                            pOrder.LimitPriceOffsetType, _
                            True) & _
            gPriceToString(MaxDouble, _
                            pContract, _
                            pOrder.TriggerPriceOffset, _
                            pOrder.TriggerPriceOffsetType, _
                            True)
    End If
Case BracketStopLossTypeAuto
    If pIncludeOffsets Then
        s = s & _
            gPriceToString(MaxDouble, _
                            pContract, _
                            pOrder.TriggerPriceOffset, _
                            pOrder.TriggerPriceOffsetType, _
                            True)
    End If
Case BracketTargetTypeAuto
    If pIncludeOffsets Then
        s = s & _
            gPriceToString(MaxDouble, _
                            pContract, _
                            pOrder.LimitPriceOffset, _
                            pOrder.LimitPriceOffsetType, _
                            True)
    End If
End Select

gGetOrderTypeAndPricesString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function gGetSignedQuantity(ByVal pExec As IExecutionReport) As Long
gGetSignedQuantity = IIf(pExec.Action = OrderActionBuy, pExec.Quantity, -pExec.Quantity)
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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "bracketorderprofilestruct", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "bracketorderprofilestring", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "drawdown", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "maxloss", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "maxprofit", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "moneymanagement", pMessage, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "order", pMessage, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "orderdetail", pMessage, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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
Dim lTimePart As String

If pDataSource Is Nothing Then
ElseIf pDataSource.State <> MarketDataSourceStateRunning Then
Else
    If pDataSource.IsTickReplay Then lTimePart = FormatTimestamp(pDataSource.Timestamp, TimestampDateAndTimeISO8601) & "  "
    lTickPart = GetCurrentTickSummary(pDataSource) & "; "
End If

gLogOrder lTimePart & _
            IIf(pIsSimulated, "(simulated) ", "") & _
            pMessage & vbCrLf & _
            "Contract: " & pContract.Specifier.LocalSymbol & "@" & pContract.Specifier.Exchange & vbCrLf & _
            IIf(pKey <> "", "Bracket id: " & pKey & vbCrLf, "") & _
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
                        "BrokerId: " & pOrder.BrokerId & _
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
                ByVal pPosition As Long, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogPosition"
On Error GoTo Err

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "position", pPosition, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "profit", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "tradeprofile", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

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
    Case OrderAttETradeOnly
        gOrderAttributeToString = "ETradeOnly"
    Case OrderAttFirmQuoteOnly
        gOrderAttributeToString = "FirmQuoteOnly"
    Case OrderAttNBBOPriceCap
        gOrderAttributeToString = "NBBOPriceCap"
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
Static sTypes As Collection
Const ProcName As String = "gOrderTypeFromString"
On Error GoTo Err

If sTypes Is Nothing Then
    Set sTypes = New Collection
    
    sTypes.Add OrderTypeNone, StrOrderTypeNone
    sTypes.Add OrderTypeMarket, StrOrderTypeMarket
    sTypes.Add OrderTypeMarketOnClose, StrOrderTypeMarketOnClose
    sTypes.Add OrderTypeLimit, StrOrderTypeLimit
    sTypes.Add OrderTypeLimitOnClose, StrOrderTypeLimitOnClose
    sTypes.Add OrderTypePeggedToMarket, StrOrderTypePegMarket
    sTypes.Add OrderTypeStop, StrOrderTypeStop
    sTypes.Add OrderTypeStopLimit, StrOrderTypeStopLimit
    sTypes.Add OrderTypeTrail, StrOrderTypeTrail
    sTypes.Add OrderTypeRelative, StrOrderTypeRelative
    sTypes.Add OrderTypeVWAP, StrOrderTypeVWAP
    sTypes.Add OrderTypeMarketToLimit, StrOrderTypeMarketToLimit
    sTypes.Add OrderTypeQuote, StrOrderTypeQuote
    sTypes.Add OrderTypeAdjust, StrOrderTypeAdjust
    sTypes.Add OrderTypeAlert, StrOrderTypeAlert
    sTypes.Add OrderTypeLimitIfTouched, StrOrderTypeLimitIfTouched
    sTypes.Add OrderTypeMarketIfTouched, StrOrderTypeMarketIfTouched
    sTypes.Add OrderTypeTrailLimit, StrOrderTypeTrailLimit
    sTypes.Add OrderTypeMarketWithProtection, StrOrderTypeMarketWithProtection
    sTypes.Add OrderTypeMarketOnOpen, StrOrderTypeMarketOnOpen
    sTypes.Add OrderTypeLimitOnOpen, StrOrderTypeLimitOnOpen
    sTypes.Add OrderTypePeggedToPrimary, StrOrderTypePeggedToPrimary
    sTypes.Add OrderTypes.OrderTypeAutoLimit, StrOrderTypeAutoLimit
    sTypes.Add OrderTypes.OrderTypeAutoStop, StrOrderTypeAutoStop

    sTypes.Add OrderTypes.OrderTypeMarket, "MKT"
    sTypes.Add OrderTypes.OrderTypeMarketOnClose, "MKTCLS"
    sTypes.Add OrderTypes.OrderTypeLimit, "LMT"
    sTypes.Add OrderTypes.OrderTypeLimitOnClose, "LMTCLS"
    sTypes.Add OrderTypes.OrderTypePeggedToMarket, "PEGMKT"
    sTypes.Add OrderTypes.OrderTypeStop, "STP"
    sTypes.Add OrderTypes.OrderTypeStopLimit, "STPLMT"
    sTypes.Add OrderTypes.OrderTypeTrail, "TRAIL"
    sTypes.Add OrderTypes.OrderTypeRelative, "REL"
    sTypes.Add OrderTypes.OrderTypeVWAP, "VWAP"
    sTypes.Add OrderTypes.OrderTypeMarketToLimit, "MTL"
    sTypes.Add OrderTypes.OrderTypeQuote, "QUOTE"
    sTypes.Add OrderTypes.OrderTypeAdjust, "ADJUST"
    sTypes.Add OrderTypes.OrderTypeAlert, "ALERT"
    sTypes.Add OrderTypes.OrderTypeLimitIfTouched, "LIT"
    sTypes.Add OrderTypes.OrderTypeMarketIfTouched, "MIT"
    sTypes.Add OrderTypes.OrderTypeTrailLimit, "TRAILLMT"
    sTypes.Add OrderTypes.OrderTypeMarketWithProtection, "MKTPROT"
    sTypes.Add OrderTypes.OrderTypeMarketOnOpen, "MOO"
    sTypes.Add OrderTypes.OrderTypeLimitOnOpen, "LOO"
    sTypes.Add OrderTypes.OrderTypePeggedToPrimary, "PEGPRI"
End If

gOrderTypeFromString = sTypes(Value)

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Type"
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
Case OrderTypeVWAP
    gOrderTypeToString = StrOrderTypeVWAP
Case OrderTypeMarketToLimit
    gOrderTypeToString = StrOrderTypeMarketToLimit
Case OrderTypeQuote
    gOrderTypeToString = StrOrderTypeQuote
Case OrderTypeAdjust
    gOrderTypeToString = StrOrderTypeAdjust
Case OrderTypeAlert
    gOrderTypeToString = StrOrderTypeAlert
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
Case OrderTypeAutoLimit
    gOrderTypeToString = StrOrderTypeAutoLimit
Case OrderTypeAutoStop
    gOrderTypeToString = StrOrderTypeAutoStop
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
Case OrderTypes.OrderTypeVWAP
    gOrderTypeToShortString = "VWAP"
Case OrderTypes.OrderTypeMarketToLimit
    gOrderTypeToShortString = "MTL"
Case OrderTypes.OrderTypeQuote
    gOrderTypeToShortString = "QUOTE"
Case OrderTypes.OrderTypeAdjust
    gOrderTypeToShortString = "ADJUST"
Case OrderTypes.OrderTypeAlert
    gOrderTypeToShortString = "ALERT"
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
Case OrderTypes.OrderTypeAutoLimit
    gOrderTypeToShortString = "AUTOLMT"
Case OrderTypes.OrderTypeAutoStop
    gOrderTypeToShortString = "AUTOSTP"
Case BracketEntryTypeAsk, BracketStopLossTypeAsk, BracketTargetTypeAsk
    gOrderTypeToShortString = "ASK"
Case BracketEntryTypeBid, BracketStopLossTypeBid, BracketTargetTypeBid
    gOrderTypeToShortString = "BID"
Case BracketEntryTypeLast, BracketStopLossTypeBid, BracketTargetTypeBid
    gOrderTypeToShortString = "LAST"
Case BracketStopLossTypeAuto, BracketTargetTypeAuto
    gOrderTypeToShortString = "AUTO"
Case Else
    AssertArgument False, "Value is not a valid Order Type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPriceOffsetToString( _
                ByVal pOffset As Double, _
                ByVal pOffsetType As PriceOffsetTypes)
Select Case pOffsetType
Case PriceOffsetTypeNone
    gPriceOffsetToString = ""
Case PriceOffsetTypeIncrement
    gPriceOffsetToString = "[" & pOffset & "]"
Case PriceOffsetTypeNumberOfTicks
    gPriceOffsetToString = "[" & CInt(pOffset) & "T]"
Case PriceOffsetTypeBidAskPercent
    gPriceOffsetToString = "[" & CInt(pOffset) & "%]"
Case Else
    AssertArgument False, "Value is not a valid Price Offset Type"
End Select
End Function

Public Function gPriceOffsetTypeToString( _
                ByVal pOffsetType As PriceOffsetTypes)
Select Case pOffsetType
Case PriceOffsetTypeNone
    gPriceOffsetTypeToString = "N/A"
Case PriceOffsetTypeIncrement
    gPriceOffsetTypeToString = ""
Case PriceOffsetTypeNumberOfTicks
    gPriceOffsetTypeToString = "T"
Case PriceOffsetTypeBidAskPercent
    gPriceOffsetTypeToString = "%"
Case Else
    AssertArgument False, "Value is not a valid Price Offset Type"
End Select
End Function

Public Function gPriceToString( _
                ByVal pPrice As Double, _
                ByVal pContract As IContract, _
                Optional ByVal pOffset As Double = 0#, _
                Optional ByVal pOffsetType As PriceOffsetTypes = PriceOffsetTypeNone, _
                Optional ByVal pIncludeOffset As Boolean = False) As String
Const ProcName As String = "gPriceToString"
On Error GoTo Err

If pPrice <> MaxDouble Then
    gPriceToString = FormatPrice(pPrice, pContract.Specifier.SecType, pContract.TickSize)
End If
If pIncludeOffset Then
    gPriceToString = gPriceToString & _
                    gPriceOffsetToString(pOffset, pOffsetType)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
Public Sub gSyncToOrder(ByVal pTargetOrder As IOrder, ByVal pSourceOrder As IOrder)
Const ProcName As String = "gSyncToOrder"
On Error GoTo Err

With pTargetOrder
    .Initialise pSourceOrder.GroupName, pSourceOrder.ContractSpecifier, pSourceOrder.OrderContext
    
    ' do this first because modifiability of other attributes may depend on the OrderType
    .OrderType = pSourceOrder.OrderType
    
    .Action = pSourceOrder.Action
    .AllOrNone = pSourceOrder.AllOrNone
    .AveragePrice = pSourceOrder.AveragePrice
    .BlockOrder = pSourceOrder.BlockOrder
    .BrokerId = pSourceOrder.BrokerId
    .DiscretionaryAmount = pSourceOrder.DiscretionaryAmount
    .DisplaySize = pSourceOrder.DisplaySize
    .ErrorCode = pSourceOrder.ErrorCode
    .ErrorMessage = pSourceOrder.ErrorMessage
    .ETradeOnly = pSourceOrder.ETradeOnly
    .FillTime = pSourceOrder.FillTime
    .FirmQuoteOnly = pSourceOrder.FirmQuoteOnly
    .GoodAfterTime = pSourceOrder.GoodAfterTime
    .GoodAfterTimeTZ = pSourceOrder.GoodAfterTimeTZ
    .GoodTillDate = pSourceOrder.GoodTillDate
    .GoodTillDateTZ = pSourceOrder.GoodTillDateTZ
    .Hidden = pSourceOrder.Hidden
    .Id = pSourceOrder.Id
    .IgnoreRegularTradingHours = pSourceOrder.IgnoreRegularTradingHours
    .IsSimulated = pSourceOrder.IsSimulated
    .LastFillPrice = pSourceOrder.LastFillPrice
    .LimitPrice = pSourceOrder.LimitPrice
    .LimitPriceOffset = pSourceOrder.LimitPriceOffset
    .LimitPriceOffsetType = pSourceOrder.LimitPriceOffsetType
    .MinimumQuantity = pSourceOrder.MinimumQuantity
    .NbboPriceCap = pSourceOrder.NbboPriceCap
'    .Offset = pSourceOrder.Offset
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
    .TriggerPrice = pSourceOrder.TriggerPrice
    .TriggerPriceOffset = pSourceOrder.TriggerPriceOffset
    .TriggerPriceOffsetType = pSourceOrder.TriggerPriceOffsetType

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
                ByRef pLogger As Logger)
Const ProcName As String = "logInfotypeData"
On Error GoTo Err

If pLogger Is Nothing Then
    Set pLogger = GetLogger("position." & pInfoType & IIf(pSimulated, "Simulated", ""))
    pLogger.LogToParent = False
End If
pLogger.Log pLogLevel, pData, pSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyCollectionMember( _
                ByVal pItem As Variant, _
                ByVal pSource As Object, _
                ByVal pListener As ICollectionChangeListener)
Dim ev As CollectionChangeEventData
Const ProcName As String = "notifyCollectionMember"
On Error GoTo Err

Set ev.Source = pSource
ev.changeType = CollItemAdded

gSetVariant ev.AffectedItem, pItem
pListener.Change ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




