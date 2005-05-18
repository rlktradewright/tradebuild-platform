Attribute VB_Name = "Globals"
Option Explicit

Public Declare Function SendMessageByNum Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub InitCommonControls Lib "comctl32" ()

Public Const LB_SETHORZEXTENT = &H194

Public Const StrOrderTypeNone As String = ""
Public Const StrOrderTypeMarket As String = "Market"
Public Const StrOrderTypeMarketClose As String = "Market on Close"
Public Const StrOrderTypeLimit As String = "Limit"
Public Const StrOrderTypeLimitClose As String = "Limit on Close"
Public Const StrOrderTypePegMarket As String = "Peg to Market"
Public Const StrOrderTypeStop As String = "Stop"
Public Const StrOrderTypeStopLimit As String = "Stop Limit"
Public Const StrOrderTypeTrail As String = "Trail"
Public Const StrOrderTypeRelative As String = "Relative"
Public Const StrOrderTypeVWAP As String = "VWAP"
Public Const StrOrderTypeMarketToLimit As String = "Market to Limit"
Public Const StrOrderTypeQuote As String = "Quote"

Public Const StrOrderActionBuy As String = "BUY"
Public Const StrOrderActionSell As String = "SELL"

Public Const StrSecTypeStock As String = "Stock"
Public Const StrSecTypeFuture As String = "Future"
Public Const StrSecTypeOption As String = "Option"
Public Const StrSecTypeOptionFuture As String = "Option on futures"
Public Const StrSecTypeCash As String = "Cash"
Public Const StrSecTypeBag As String = "Bag"


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
Select Case value
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
Select Case value
Case StrOrderTypeNone
    orderTypeFromString = OrderTypeNone
Case StrOrderTypeMarket
    orderTypeFromString = OrderTypeMarket
Case StrOrderTypeMarketClose
    orderTypeFromString = OrderTypeMarketClose
Case StrOrderTypeLimit
    orderTypeFromString = OrderTypeLimit
Case StrOrderTypeLimitClose
    orderTypeFromString = OrderTypeLimitClose
Case StrOrderTypePegMarket
    orderTypeFromString = OrderTypePegMarket
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
End Select

End Function
Public Function orderTypeToString(ByVal value As OrderTypes) As String
Select Case value
Case OrderTypeNone
    orderTypeToString = StrOrderTypeNone
Case OrderTypeMarket
    orderTypeToString = StrOrderTypeMarket
Case OrderTypeMarketClose
    orderTypeToString = StrOrderTypeMarketClose
Case OrderTypeLimit
    orderTypeToString = StrOrderTypeLimit
Case OrderTypeLimitClose
    orderTypeToString = StrOrderTypeLimitClose
Case OrderTypePegMarket
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
End Select

End Function

Public Function secTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "Stock"
    secTypeFromString = SecTypeStock
Case "Future"
    secTypeFromString = SecTypeFuture
Case "Option"
    secTypeFromString = SecTypeOption
Case "Futures Option"
    secTypeFromString = SecTypeFuturesOption
Case "Cash"
    secTypeFromString = SecTypeCash
Case "Bag"
    secTypeFromString = SecTypeBag
Case "Index"
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


