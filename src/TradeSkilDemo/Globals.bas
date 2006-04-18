Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const LB_SETHORZEXTENT = &H194

Public Const StrOrderTypeNone As String = ""
Public Const StrOrderTypeMarket As String = "Market"
Public Const StrOrderTypeMarketClose As String = "Market on Close"
Public Const StrOrderTypeLimit As String = "Limit"
Public Const StrOrderTypeLimitClose As String = "Limit on Close"
Public Const StrOrderTypePegMarket As String = "Peg to Market"
Public Const StrOrderTypeStop As String = "Stop"
Public Const StrOrderTypeStopLimit As String = "Stop Limit"
Public Const StrOrderTypeTrail As String = "Trailing Stop"
Public Const StrOrderTypeRelative As String = "Relative"
Public Const StrOrderTypeVWAP As String = "VWAP"
Public Const StrOrderTypeMarketToLimit As String = "Market to Limit"
Public Const StrOrderTypeQuote As String = "Quote"
Public Const StrOrderTypeAutoStop As String = "Auto Stop"
Public Const StrOrderTypeAutoLimit As String = "Auto Limit"
Public Const StrOrderTypeAdjust As String = "Adjust"
Public Const StrOrderTypeAlert As String = "Alert"
Public Const StrOrderTypeLimitIfTouched As String = "Limit if Touched"
Public Const StrOrderTypeMarketIfTouched As String = "Market if Touched"
Public Const StrOrderTypeTrailLimit As String = "Trail Limit"
Public Const StrOrderTypeMarketWithProtection As String = "Market with Protection"
Public Const StrOrderTypeMarketOnOpen As String = "Market on Open"
Public Const StrOrderTypeLimitOnOpen As String = "Limit on Open"
Public Const StrOrderTypePeggedToPrimary As String = "Pegged to Primary"

Public Const StrOrderActionBuy As String = "BUY"
Public Const StrOrderActionSell As String = "SELL"

Public Const StrSecTypeStock As String = "Stock"
Public Const StrSecTypeFuture As String = "Future"
Public Const StrSecTypeOption As String = "Option"
Public Const StrSecTypeOptionFuture As String = "Option on futures"
Public Const StrSecTypeCash As String = "Cash"
Public Const StrSecTypeBag As String = "Bag"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

Public gMainForm As fTradeSkilDemo

Public gTradeBuildAPI As TradeBuildAPI

'================================================================================
' External function declarations
'================================================================================

Public Declare Function SendMessageByNum Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Sub InitCommonControls Lib "comctl32" ()

'================================================================================
' Variables
'================================================================================

'================================================================================
' Procedures
'================================================================================

Public Function entryTypeToString(ByVal value As EntryTypes) As String
Select Case value
Case EntryTypeMarket
    entryTypeToString = "Market"
Case EntryTypeMarketOnOpen
    entryTypeToString = "Market on open"
Case EntryTypeMarketOnClose
    entryTypeToString = "Market on close"
Case EntryTypeMarketIfTouched
    entryTypeToString = "Market if touched"
Case EntryTypeMarketToLimit
    entryTypeToString = "Market to limit"
Case EntryTypeBid
    entryTypeToString = "Bid price"
Case EntryTypeAsk
    entryTypeToString = "Ask price"
Case EntryTypeLast
    entryTypeToString = "Last trade price"
Case EntryTypeLimit
    entryTypeToString = "Limit"
Case EntryTypeLimitOnOpen
    entryTypeToString = "Limit on open"
Case EntryTypeLimitOnClose
    entryTypeToString = "Limit on close"
Case EntryTypeLimitIfTouched
    entryTypeToString = "Limit if touched"
Case EntryTypeStop
    entryTypeToString = "Stop"
Case EntryTypeStopLimit
    entryTypeToString = "Stop limit"
End Select
End Function

Public Function optionRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case "CALL"
    optionRightFromString = OptCall
Case "PUT"
    optionRightFromString = OptPut
Case Else
    optionRightFromString = OptNone
End Select
End Function

Public Function optionRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptCall
    optionRightToString = "CALL"
Case OptPut
    optionRightToString = "PUT"
Case OptNone
    optionRightToString = ""
End Select
End Function

Public Function orderActionFromString(ByVal value As String) As OrderActions
Select Case UCase$(value)
Case StrOrderActionBuy
    orderActionFromString = ActionBuy
Case StrOrderActionSell
    orderActionFromString = ActionSell
End Select
End Function

Public Function orderActionToString(ByVal value As OrderActions) As String
Select Case value
Case ActionBuy
    orderActionToString = StrOrderActionBuy
Case ActionSell
    orderActionToString = StrOrderActionSell
End Select
End Function

Public Function orderStatusToString(ByVal value As OrderStatuses) As String
Select Case value
Case OrderStatusCreated
    orderStatusToString = "Created"
Case OrderStatusPendingSubmit
    orderStatusToString = "Pendingsubmit"
Case orderstatuspresubmitted
    orderStatusToString = "Presubmitted"
Case orderstatussubmitted
    orderStatusToString = "Submitted"
Case orderstatuscancelling
    orderStatusToString = "Cancelling"
Case orderstatuscancelled
    orderStatusToString = "Cancelled"
Case orderstatusfilled
    orderStatusToString = "Filled"
End Select
End Function

Public Function orderTIFFromString(ByVal value As String) As OrderTifs
Select Case UCase$(value)
Case "DAY"
    orderTIFFromString = TIFDay
Case "GTC"
    orderTIFFromString = TIFGoodTillCancelled
Case "IOC"
    orderTIFFromString = TIFImmediateOrCancel
End Select
End Function

Public Function orderTIFToString(ByVal value As OrderTifs) As String
Select Case value
Case TIFDay
    orderTIFToString = "DAY"
Case TIFGoodTillCancelled
    orderTIFToString = "GTC"
Case TIFImmediateOrCancel
    orderTIFToString = "IOC"
End Select
End Function

Public Function orderTriggerMethodFromString(ByVal value As String) As TriggerMethods
Select Case UCase$(value)
Case "Default"
    orderTriggerMethodFromString = TriggerMethods.TriggerDefault
Case "Double bid/ask"
    orderTriggerMethodFromString = TriggerMethods.TriggerDoubleBidAsk
Case "Double last"
    orderTriggerMethodFromString = TriggerMethods.TriggerDoubleLast
Case "Last"
    orderTriggerMethodFromString = TriggerMethods.TriggerLast
End Select
End Function

Public Function orderTriggerMethodToString(ByVal value As TriggerMethods) As String
Select Case value
Case TriggerMethods.TriggerDefault
    orderTriggerMethodToString = "Default"
Case TriggerMethods.TriggerDoubleBidAsk
    orderTriggerMethodToString = "Double bid/ask"
Case TriggerMethods.TriggerDoubleLast
    orderTriggerMethodToString = "Double last"
Case TriggerMethods.TriggerLast
    orderTriggerMethodToString = "Last"
End Select
End Function

Public Function orderTypeFromString(ByVal value As String) As OrderTypes
Select Case UCase$(value)
Case StrOrderTypeNone
    orderTypeFromString = OrderTypeNone
Case StrOrderTypeMarket
    orderTypeFromString = OrderTypeMarket
Case StrOrderTypeMarketClose
    orderTypeFromString = OrderTypeMarketOnClose
Case StrOrderTypeLimit
    orderTypeFromString = OrderTypeLimit
Case StrOrderTypeLimitClose
    orderTypeFromString = OrderTypeLimitOnClose
Case StrOrderTypePegMarket
    orderTypeFromString = OrderTypePeggedToMarket
Case StrOrderTypeStop
    orderTypeFromString = OrderTypeStop
Case StrOrderTypeStopLimit
    orderTypeFromString = OrderTypeStopLimit
Case StrOrderTypeTrail
    orderTypeFromString = OrderTypeTrail
Case StrOrderTypeRelative
    orderTypeFromString = OrderTypeRelative
Case StrOrderTypeVWAP
    orderTypeFromString = OrderTypeVWAP
Case StrOrderTypeMarketToLimit
    orderTypeFromString = OrderTypeMarketToLimit
Case StrOrderTypeQuote
    orderTypeFromString = OrderTypeQuote
Case StrOrderTypeAdjust
    orderTypeFromString = OrderTypeAdjust
Case StrOrderTypeAlert
    orderTypeFromString = OrderTypeAlert
Case StrOrderTypeLimitIfTouched
    orderTypeFromString = OrderTypeLimitIfTouched
Case StrOrderTypeMarketIfTouched
    orderTypeFromString = OrderTypeMarketIfTouched
Case StrOrderTypeTrailLimit
    orderTypeFromString = OrderTypeTrailLimit
Case StrOrderTypeMarketWithProtection
    orderTypeFromString = OrderTypeMarketWithProtection
Case StrOrderTypeMarketOnOpen
    orderTypeFromString = OrderTypeMarketOnOpen
Case StrOrderTypeLimitOnOpen
    orderTypeFromString = OrderTypeLimitOnOpen
Case StrOrderTypePeggedToPrimary
    orderTypeFromString = OrderTypePeggedToPrimary
End Select

End Function
Public Function orderTypeToString(ByVal value As OrderTypes) As String
Select Case value
Case OrderTypeNone
    orderTypeToString = StrOrderTypeNone
Case OrderTypeMarket
    orderTypeToString = StrOrderTypeMarket
Case OrderTypeMarketOnClose
    orderTypeToString = StrOrderTypeMarketClose
Case OrderTypeLimit
    orderTypeToString = StrOrderTypeLimit
Case OrderTypeLimitOnClose
    orderTypeToString = StrOrderTypeLimitClose
Case OrderTypePeggedToMarket
    orderTypeToString = StrOrderTypePegMarket
Case OrderTypeStop
    orderTypeToString = StrOrderTypeStop
Case OrderTypeStopLimit
    orderTypeToString = StrOrderTypeStopLimit
Case OrderTypeTrail
    orderTypeToString = StrOrderTypeTrail
Case OrderTypeRelative
    orderTypeToString = StrOrderTypeRelative
Case OrderTypeVWAP
    orderTypeToString = StrOrderTypeVWAP
Case OrderTypeMarketToLimit
    orderTypeToString = StrOrderTypeMarketToLimit
Case OrderTypeQuote
    orderTypeToString = StrOrderTypeQuote
Case OrderTypeAdjust
    orderTypeToString = StrOrderTypeAdjust
Case OrderTypeAlert
    orderTypeToString = StrOrderTypeAlert
Case OrderTypeLimitIfTouched
    orderTypeToString = StrOrderTypeLimitIfTouched
Case OrderTypeMarketIfTouched
    orderTypeToString = StrOrderTypeMarketIfTouched
Case OrderTypeTrailLimit
    orderTypeToString = StrOrderTypeTrailLimit
Case OrderTypeMarketWithProtection
    orderTypeToString = StrOrderTypeMarketWithProtection
Case OrderTypeMarketOnOpen
    orderTypeToString = StrOrderTypeMarketOnOpen
Case OrderTypeLimitOnOpen
    orderTypeToString = StrOrderTypeLimitOnOpen
Case OrderTypePeggedToPrimary
    orderTypeToString = StrOrderTypePeggedToPrimary
End Select

End Function

Public Function secTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STOCK"
    secTypeFromString = SecTypeStock
Case "FUTURE"
    secTypeFromString = SecTypeFuture
Case "OPTION"
    secTypeFromString = SecTypeOption
Case "FUTURES OPTION"
    secTypeFromString = SecTypeFuturesOption
Case "CASH"
    secTypeFromString = SecTypeCash
Case "BAG"
    secTypeFromString = SecTypeBag
Case "INDEX"
    secTypeFromString = SecTypeIndex
End Select
End Function

Public Function secTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    secTypeToString = "Stock"
Case SecTypeFuture
    secTypeToString = "Future"
Case SecTypeOption
    secTypeToString = "Option"
Case SecTypeFuturesOption
    secTypeToString = "Futures Option"
Case SecTypeCash
    secTypeToString = "Cash"
Case SecTypeBag
    secTypeToString = "Bag"
Case SecTypeIndex
    secTypeToString = "Index"
End Select
End Function

Public Function stopTypeToString(ByVal value As StopTypes)
Select Case value
Case StopTypeNone
    stopTypeToString = "None"
Case StopTypeStop
    stopTypeToString = "Stop"
Case StopTypeStopLimit
    stopTypeToString = "Stop limit"
Case StopTypeBid
    stopTypeToString = "Bid price"
Case StopTypeAsk
    stopTypeToString = "Ask price"
Case StopTypeLast
    stopTypeToString = "Last trade price"
Case StopTypeAuto
    stopTypeToString = "Auto"
End Select
End Function

Public Function targetTypeToString(ByVal value As TargetTypes)
Select Case value
Case TargetTypeNone
    targetTypeToString = "None"
Case TargetTypeLimit
    targetTypeToString = "Limit"
Case TargetTypeMarketIfTouched
    targetTypeToString = "Market if touched"
Case TargetTypeBid
    targetTypeToString = "Bid price"
Case TargetTypeAsk
    targetTypeToString = "Ask price"
Case TargetTypeLast
    targetTypeToString = "Last trade price"
Case TargetTypeAuto
    targetTypeToString = "Auto"
End Select
End Function
