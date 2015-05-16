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

Public Const RecoveryOrderContextName       As String = "$recovery"

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
    gBracketOrderRoleToString = "Stop loss"
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
Case BracketStopLossTypeBid
    gBracketStopLossTypeToOrderType = OrderTypeStop
Case BracketStopLossTypeAsk
    gBracketStopLossTypeToOrderType = OrderTypeStop
Case BracketStopLossTypeLast
    gBracketStopLossTypeToOrderType = OrderTypeStop
Case BracketStopLossTypeAuto
    gBracketStopLossTypeToOrderType = OrderTypeAutoStop
Case Else
    AssertArgument False, "Invalid entry type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketStopLossTypeToShortString(ByVal Value As BracketStopLossTypes)
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
    gBracketStopLossTypeToString = "Trade"
Case BracketStopLossTypeAuto
    gBracketStopLossTypeToString = "Auto"
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
    AssertArgument False, "Invalid entry type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBracketTargetTypeToShortString(ByVal Value As BracketTargetTypes)
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
    s = s & " " & gPriceToString(pOrder.LimitPrice, pContract)
Case OrderTypeStop, _
        OrderTypeMarketIfTouched
    s = s & " " & gPriceToString(pOrder.TriggerPrice, pContract)
Case OrderTypeStopLimit, _
        OrderTypeLimitIfTouched
    s = s & " " & gPriceToString(pOrder.LimitPrice, pContract) & _
        " " & gPriceToString(pOrder.TriggerPrice, pContract)
Case OrderTypeTrail

Case OrderTypeRelative

Case OrderTypeVWAP

Case OrderTypeQuote

Case OrderTypeAutoStop

Case OrderTypeAutoLimit

Case OrderTypeAdjust

Case OrderTypeAlert

Case OrderTypeTrailLimit

Case OrderTypeMarketWithProtection

Case OrderTypeMarketOnOpen

Case OrderTypePeggedToPrimary

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

Public Sub gLogBracketOrderProfileStruct( _
                ByVal pData As Variant, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogBracketOrderProfileStruct"
On Error GoTo Err

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "bracketorderprofilestruct", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogBracketOrderProfileString( _
                ByVal pData As Variant, _
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
                ByVal pData As Variant, _
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
                ByVal pData As Variant, _
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
                ByVal pData As Variant, _
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
                ByVal pData As Variant, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogMoneyManagement"
On Error GoTo Err

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "moneymanagement", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogOrder( _
                ByVal pData As Variant, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogOrder"
On Error GoTo Err

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "order", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogOrderDetail( _
                ByVal pData As Variant, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogOrderDetail"
On Error GoTo Err

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "orderdetail", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogPosition( _
                ByVal pData As Variant, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLogPosition"
On Error GoTo Err

Static lLogger As Logger
Static lLoggerSimulated As Logger

logInfotypeData "position", pData, pSimulated, pSource, pLogLevel, IIf(pSimulated, lLoggerSimulated, lLogger)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogProfit( _
                ByVal pData As Variant, _
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
                ByVal pData As Variant, _
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
Case Else
    AssertArgument False, "Value is not a valid Order Type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPriceToString( _
                ByVal pPrice As Double, _
                ByVal pContract As IContract) As String
Const ProcName As String = "gPriceToString"
On Error GoTo Err

gPriceToString = FormatPrice(pPrice, pContract.Specifier.SecType, pContract.TickSize)

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
    .Initialise .GroupName, pSourceOrder.ContextsName, pSourceOrder.ContractSpecifier, pSourceOrder.OrderContext
    
    ' do this first because modifiability of other attributes may depend on the OrderType
    If .IsAttributeModifiable(OrderAttOrderType) Then .OrderType = pSourceOrder.OrderType
    
    .Action = pSourceOrder.Action
    If .IsAttributeModifiable(OrderAttAllOrNone) Then .AllOrNone = pSourceOrder.AllOrNone
    .AveragePrice = pSourceOrder.AveragePrice
    If .IsAttributeModifiable(OrderAttBlockOrder) Then .BlockOrder = pSourceOrder.BlockOrder
    .BrokerId = pSourceOrder.BrokerId
    If .IsAttributeModifiable(OrderAttDiscretionaryAmount) Then .DiscretionaryAmount = pSourceOrder.DiscretionaryAmount
    If .IsAttributeModifiable(OrderAttDisplaySize) Then .DisplaySize = pSourceOrder.DisplaySize
    .ErrorCode = pSourceOrder.ErrorCode
    .ErrorMessage = pSourceOrder.ErrorMessage
    If .IsAttributeModifiable(OrderAttETradeOnly) Then .ETradeOnly = pSourceOrder.ETradeOnly
    .FillTime = pSourceOrder.FillTime
    If .IsAttributeModifiable(OrderAttFirmQuoteOnly) Then .FirmQuoteOnly = pSourceOrder.FirmQuoteOnly
    If .IsAttributeModifiable(OrderAttGoodAfterTime) Then .GoodAfterTime = pSourceOrder.GoodAfterTime
    If .IsAttributeModifiable(OrderAttGoodAfterTimeTZ) Then .GoodAfterTimeTZ = pSourceOrder.GoodAfterTimeTZ
    If .IsAttributeModifiable(OrderAttGoodTillDate) Then .GoodTillDate = pSourceOrder.GoodTillDate
    If .IsAttributeModifiable(OrderAttGoodTillDateTZ) Then .GoodTillDateTZ = pSourceOrder.GoodTillDateTZ
    If .IsAttributeModifiable(OrderAttHidden) Then .Hidden = pSourceOrder.Hidden
    .Id = pSourceOrder.Id
    If .IsAttributeModifiable(OrderAttIgnoreRTH) Then .IgnoreRegularTradingHours = pSourceOrder.IgnoreRegularTradingHours
    .IsSimulated = pSourceOrder.IsSimulated
    .LastFillPrice = pSourceOrder.LastFillPrice
    If .IsAttributeModifiable(OrderAttLimitPrice) Then .LimitPrice = pSourceOrder.LimitPrice
    If .IsAttributeModifiable(OrderAttMinimumQuantity) Then .MinimumQuantity = pSourceOrder.MinimumQuantity
    If .IsAttributeModifiable(OrderAttNBBOPriceCap) Then .NbboPriceCap = pSourceOrder.NbboPriceCap
    .Offset = pSourceOrder.Offset
    If .IsAttributeModifiable(OrderAttOrigin) Then .Origin = pSourceOrder.Origin
    If .IsAttributeModifiable(OrderAttOriginatorRef) Then .OriginatorRef = pSourceOrder.OriginatorRef
    If .IsAttributeModifiable(OrderAttOverrideConstraints) Then .OverrideConstraints = pSourceOrder.OverrideConstraints
    If .IsAttributeModifiable(OrderAttPercentOffset) Then .PercentOffset = pSourceOrder.PercentOffset
    If .IsAttributeModifiable(OrderAttQuantity) Then .Quantity = pSourceOrder.Quantity
    .QuantityFilled = pSourceOrder.QuantityFilled
    .QuantityRemaining = pSourceOrder.QuantityRemaining
    If .IsAttributeModifiable(OrderAttSettlingFirm) Then .SettlingFirm = pSourceOrder.SettlingFirm
    If .IsAttributeModifiable(OrderAttStopTriggerMethod) Then .StopTriggerMethod = pSourceOrder.StopTriggerMethod
    If .IsAttributeModifiable(OrderAttSweepToFill) Then .SweepToFill = pSourceOrder.SweepToFill
    If .IsAttributeModifiable(OrderAttTimeInForce) Then .TimeInForce = pSourceOrder.TimeInForce
    If .IsAttributeModifiable(OrderAttTriggerPrice) Then .TriggerPrice = pSourceOrder.TriggerPrice

    ' do this last to prevent status influencing whether attributes are modifiable
    .Status = pSourceOrder.Status
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub logInfotypeData( _
                ByVal pInfotype As String, _
                ByRef pData As Variant, _
                ByVal pSimulated As Boolean, _
                ByVal pSource As Object, _
                ByVal pLogLevel As LogLevels, _
                ByRef pLogger As Logger)
Const ProcName As String = "logInfotypeData"
On Error GoTo Err

If pLogger Is Nothing Then
    Set pLogger = GetLogger("position." & pInfotype & IIf(pSimulated, "Simulated", ""))
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
                ByVal pListener As CollectionChangeListener)
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




