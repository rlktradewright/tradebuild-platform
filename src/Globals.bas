Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                As String = "TradeBuild26"

Private Const ModuleName                As String = "Globals"

Public Const S_OK                           As Long = 0
Public Const NoValidID                      As Long = -1

Public Const DefaultStudyValue              As String = "$default"

Public Const MaxCurrency                    As Currency = 922337203685477.5807@
Public Const MinDouble                      As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MaxDouble                      As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const OneSecond                      As Double = 1.15740740740741E-05
Public Const OneMicroSecond                 As Double = 1.15740740740741E-11

Public Const MultiTaskingTimeQuantumMillisecs As Long = 20

Public Const BidInputName                   As String = "Bid"
Public Const AskInputName                   As String = "Ask"
Public Const OpenInterestInputName          As String = "Open interest"
Public Const TradeInputName                 As String = "Trade"
Public Const TickVolumeInputName            As String = "Tick Volume"
Public Const VolumeInputName                As String = "Total Volume"

Public Const StrOrderTypeNone               As String = ""
Public Const StrOrderTypeMarket             As String = "Market"
Public Const StrOrderTypeMarketClose        As String = "Market on Close"
Public Const StrOrderTypeLimit              As String = "Limit"
Public Const StrOrderTypeLimitClose         As String = "Limit on Close"
Public Const StrOrderTypePegMarket          As String = "Peg to Market"
Public Const StrOrderTypeStop               As String = "Stop"
Public Const StrOrderTypeStopLimit          As String = "Stop Limit"
Public Const StrOrderTypeTrail              As String = "Trailing Stop"
Public Const StrOrderTypeRelative           As String = "Relative"
Public Const StrOrderTypeVWAP               As String = "VWAP"
Public Const StrOrderTypeMarketToLimit      As String = "Market to Limit"
Public Const StrOrderTypeQuote              As String = "Quote"
Public Const StrOrderTypeAutoStop           As String = "Auto Stop"
Public Const StrOrderTypeAutoLimit          As String = "Auto Limit"
Public Const StrOrderTypeAdjust             As String = "Adjust"
Public Const StrOrderTypeAlert              As String = "Alert"
Public Const StrOrderTypeLimitIfTouched     As String = "Limit if Touched"
Public Const StrOrderTypeMarketIfTouched    As String = "Market if Touched"
Public Const StrOrderTypeTrailLimit         As String = "Trail Limit"
Public Const StrOrderTypeMarketWithProtection As String = "Market with Protection"
Public Const StrOrderTypeMarketOnOpen       As String = "Market on Open"
Public Const StrOrderTypeLimitOnOpen        As String = "Limit on Open"
Public Const StrOrderTypePeggedToPrimary    As String = "Pegged to Primary"

Public Const StrOrderActionBuy              As String = "Buy"
Public Const StrOrderActionSell             As String = "Sell"

'@================================================================================
' Enums
'@================================================================================

Public Enum TradeBuildListenValueTypes

    VTAll = -1  ' used by listenenrs to specify that they want to receive all
                ' types of listen data
    
    VTLog = 1
    VTTrace
    VTDebug

    VTProfitProfile
    VTSimulatedProfitProfile
    VTMoneyManagement
    VTOrderPlexProfileStruct
    VTSimulatedOrderPlexProfileStruct
    VTOrderPlexProfileString
    VTSimulatedOrderPlexProfileString
    VTOrder
    VTSimulatedOrder
    VTPosition
    VTSimulatedPosition
    VTTradeProfile
    VTSimulatedTradeProfile
    VTProfit
    VTSimulatedProfit
    VTDrawdown
    VTSimulatedDrawdown
    VTMaxProfit
    VTSimulatedMaxProfit
    VTOrderDetail
    VTOrderDetailSimulated
        
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

Public gTB As TradeBuildAPI

Private mLogLogger As Logger
Private mErrorLogger As Logger

Private mSpLogLogger As Logger

Private mTraceLogger As Logger

Private mDebugLogger As Logger

Private mProfitProfileLogger As Logger

Private mProfitProfileLoggerSimulated As Logger

Private mMoneyManagementLogger As Logger

Private mOrderPlexProfileStructLogger As Logger

Private mOrderPlexProfileStructLoggerSimulated As Logger

Private mOrderPlexProfileStringLogger As Logger

Private mOrderPlexProfileStringLoggerSimulated As Logger

Private mOrderLogger As Logger

Private mOrderLoggerSimulated As Logger

Private mPositionLogger As Logger

Private mPositionLoggerSimulated As Logger

Private mTradeProfileLogger As Logger

Private mTradeProfileLoggerSimulated As Logger

Private mProfitLogger As Logger

Private mProfitLoggerSimulated As Logger

Private mDrawdownLogger  As Logger

Private mDrawdownLoggerSimulated As Logger

Private mMaxProfitLogger As Logger

Private mMaxProfitLoggerSimulated As Logger

Private mOrderDetailLogger As Logger

Private mOrderDetailLoggerSimulated As Logger

Private mLogTokens(9) As String

Private mTracer As Tracer

'@================================================================================
' Procedures
'@================================================================================

Public Function GApiNotifyCodeToString(value As ApiNotifyCodes) As String
Select Case value
Case ApiNotifyServiceProviderError
    GApiNotifyCodeToString = "ApiNotifyServiceProviderError"
Case ApiNotifyTickfileEmpty
    GApiNotifyCodeToString = "ApiNotifyTickfileEmpty"
Case ApiNotifyTickfileInvalid
    GApiNotifyCodeToString = "ApiNotifyTickfileInvalid"
Case ApiNotifyTickfileVersionNotSupported
    GApiNotifyCodeToString = "ApiNotifyTickfileVersionNotSupported"
Case ApiNotifyTickfileContractDetailsInvalid
    GApiNotifyCodeToString = "ApiNotifyTickfileContractDetailsInvalid"
Case ApiNotifyTickfileNoContractDetails
    GApiNotifyCodeToString = "ApiNotifyTickfileNoContractDetails"
Case ApiNotifyTickfileDataSourceNotResponding
    GApiNotifyCodeToString = "ApiNotifyTickfileDataSourceNotResponding"
Case ApiNotifyTickfileDoesntExist
    GApiNotifyCodeToString = "ApiNotifyTickfileDoesntExist"
Case ApiNOtifyTickfileFormatNotSupported
    GApiNotifyCodeToString = "ApiNOtifyTickfileFormatNotSupported"
Case ApiNotifyTickfileContractSpecifierInvalid
    GApiNotifyCodeToString = "ApiNotifyTickfileContractSpecifierInvalid"
Case ApiNotifyCantWriteToTickfileDataStore
    GApiNotifyCodeToString = "ApiNotifyCantWriteToTickfileDataStore"
Case ApiNotifyRetryingConnectionToTickfileDataSource
    GApiNotifyCodeToString = "ApiNotifyRetryingConnectionToTickfileDataSource"
Case ApiNotifyConnectedToTickfileDataSource
    GApiNotifyCodeToString = "ApiNotifyConnectedToTickfileDataSource"
Case ApiNotifyReconnectingToTickfileDataSource
    GApiNotifyCodeToString = "ApiNotifyReconnectingToTickfileDataSource"
Case ApiNotifyLostConnectionToTickfileDataSource
    GApiNotifyCodeToString = "ApiNotifyLostConnectionToTickfileDataSource"
Case ApiNotifyNoHistoricDataSource
    GApiNotifyCodeToString = "ApiNotifyNoHistoricDataSource"
Case ApiNotifyCantConnectHistoricDataSource
    GApiNotifyCodeToString = "ApiNotifyCantConnectHistoricDataSource"
Case ApiNotifyConnectedToHistoricDataSource
    GApiNotifyCodeToString = "ApiNotifyConnectedToHistoricDataSource"
Case ApiNotifyDisconnectedFromHistoricDataSource
    GApiNotifyCodeToString = "ApiNotifyDisconnectedFromHistoricDataSource"
Case ApiNotifyRetryingConnectionToHistoricDataSource
    GApiNotifyCodeToString = "ApiNotifyRetryingConnectionToHistoricDataSource"
Case ApiNotifyLostConnectionToHistoricDataSource
    GApiNotifyCodeToString = "ApiNotifyLostConnectionToHistoricDataSource"
Case ApiNotifyReconnectingToHistoricDataSource
    GApiNotifyCodeToString = "ApiNotifyReconnectingToHistoricDataSource"
Case ApiNotifyHistoricDataRequestFailed
    GApiNotifyCodeToString = "ApiNotifyHistoricDataRequestFailed"
Case ApiNotifyInvalidRequest
    GApiNotifyCodeToString = "ApiNotifyInvalidRequest"
Case ApiNotifyCantConnectRealtimeDataSource
    GApiNotifyCodeToString = "ApiNotifyCantConnectRealtimeDataSource"
Case ApiNotifyConnectedToRealtimeDataSource
    GApiNotifyCodeToString = "ApiNotifyConnectedToRealtimeDataSource"
Case ApiNotifyLostConnectionToRealtimeDataSource
    GApiNotifyCodeToString = "ApiNotifyLostConnectionToRealtimeDataSource"
Case ApiNotifyNoRealtimeDataSource
    GApiNotifyCodeToString = "ApiNotifyNoRealtimeDataSource"
Case ApiNotifyReconnectingToRealtimeDataSource
    GApiNotifyCodeToString = "ApiNotifyReconnectingToRealtimeDataSource"
Case ApiNotifyDisconnectedFromRealtimeDataSource
    GApiNotifyCodeToString = "ApiNotifyDisconnectedFromRealtimeDataSource"
Case ApiNotifyRealtimeDataRequestFailed
    GApiNotifyCodeToString = "ApiNotifyRealtimeDataRequestFailed"
Case ApiNotifyRealtimeDataSourceNotResponding
    GApiNotifyCodeToString = "ApiNotifyRealtimeDataSourceNotResponding"
Case ApiNotifyCantConnectToBroker
    GApiNotifyCodeToString = "ApiNotifyCantConnectToBroker"
Case ApiNotifyConnectedToBroker
    GApiNotifyCodeToString = "ApiNotifyConnectedToBroker"
Case ApiNotifyRetryConnectToBroker
    GApiNotifyCodeToString = "ApiNotifyRetryConnectToBroker"
Case ApiNotifyLostConnectionToBroker
    GApiNotifyCodeToString = "ApiNotifyLostConnectionToBroker"
Case ApiNotifyReConnectingToBroker
    GApiNotifyCodeToString = "ApiNotifyReConnectingToBroker"
Case ApiNotifyDisconnectedFromBroker
    GApiNotifyCodeToString = "ApiNotifyDisconnectedFromBroker"
Case ApiNotifyNonSpecificNotification
    GApiNotifyCodeToString = "ApiNotifyNonSpecificNotification"
Case ApiNotifyCantWriteToHistoricDataStore
    GApiNotifyCodeToString = "ApiNotifyCantWriteToHistoricDataStore"
Case ApiNotifyTryLater
    GApiNotifyCodeToString = "ApiNotifyTryLater"
Case ApiNotifyCantConnectContractDataSource
    GApiNotifyCodeToString = "ApiNotifyCantConnectContractDataSource"
Case ApiNotifyConnectedToContractDataSource
    GApiNotifyCodeToString = "ApiNotifyConnectedToContractDataSource"
Case ApiNotifyDisconnectedFromContractDataSource
    GApiNotifyCodeToString = "ApiNotifyDisconnectedFromContractDataSource"
Case ApiNotifyLostConnectionToContractDataSource
    GApiNotifyCodeToString = "ApiNotifyLostConnectionToContractDataSource"
Case ApiNotifyReConnectingContractDataSource
    GApiNotifyCodeToString = "ApiNotifyReConnectingContractDataSource"
Case ApiNotifyRetryConnectContractDataSource
    GApiNotifyCodeToString = "ApiNotifyRetryConnectContractDataSource"
End Select
End Function

''
' Converts a member of the EntryOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pEntryOrderType the EntryOrderTypes value to be converted
' @ see
'
'@/
Public Function GEntryOrderTypeToOrderType( _
                ByVal pEntryOrderType As EntryOrderTypes) As OrderTypes
Select Case pEntryOrderType
Case EntryOrderTypeMarket
    GEntryOrderTypeToOrderType = OrderTypeMarket
Case EntryOrderTypeMarketOnOpen
    GEntryOrderTypeToOrderType = OrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
    GEntryOrderTypeToOrderType = OrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
    GEntryOrderTypeToOrderType = OrderTypeMarketIfTouched
Case EntryOrderTypeMarketToLimit
    GEntryOrderTypeToOrderType = OrderTypeMarketToLimit
Case EntryOrderTypeBid
    GEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeAsk
    GEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLast
    GEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLimit
    GEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLimitOnOpen
    GEntryOrderTypeToOrderType = OrderTypeLimitOnOpen
Case EntryOrderTypeLimitOnClose
    GEntryOrderTypeToOrderType = OrderTypeLimitOnClose
Case EntryOrderTypeLimitIfTouched
    GEntryOrderTypeToOrderType = OrderTypeLimitIfTouched
Case EntryOrderTypeStop
    GEntryOrderTypeToOrderType = OrderTypeStop
Case EntryOrderTypeStopLimit
    GEntryOrderTypeToOrderType = OrderTypeStopLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Module1::gEntryOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function GEntryOrderTypeToString(ByVal value As EntryOrderTypes) As String
Select Case value
Case EntryOrderTypeMarket
    GEntryOrderTypeToString = "Market"
Case EntryOrderTypeMarketOnOpen
    GEntryOrderTypeToString = "Market on open"
Case EntryOrderTypeMarketOnClose
    GEntryOrderTypeToString = "Market on close"
Case EntryOrderTypeMarketIfTouched
    GEntryOrderTypeToString = "Market if touched"
Case EntryOrderTypeMarketToLimit
    GEntryOrderTypeToString = "Market to limit"
Case EntryOrderTypeBid
    GEntryOrderTypeToString = "Bid price"
Case EntryOrderTypeAsk
    GEntryOrderTypeToString = "Ask price"
Case EntryOrderTypeLast
    GEntryOrderTypeToString = "Last Trade price"
Case EntryOrderTypeLimit
    GEntryOrderTypeToString = "Limit"
Case EntryOrderTypeLimitOnOpen
    GEntryOrderTypeToString = "Limit on open"
Case EntryOrderTypeLimitOnClose
    GEntryOrderTypeToString = "Limit on close"
Case EntryOrderTypeLimitIfTouched
    GEntryOrderTypeToString = "Limit if touched"
Case EntryOrderTypeStop
    GEntryOrderTypeToString = "Stop"
Case EntryOrderTypeStopLimit
    GEntryOrderTypeToString = "Stop limit"
End Select
End Function

Public Function GEntryOrderTypeToShortString(ByVal value As EntryOrderTypes) As String
Select Case value
Case EntryOrderTypeMarket
    GEntryOrderTypeToShortString = "MKT"
Case EntryOrderTypeMarketOnOpen
    GEntryOrderTypeToShortString = "MOO"
Case EntryOrderTypeMarketOnClose
    GEntryOrderTypeToShortString = "MOC"
Case EntryOrderTypeMarketIfTouched
    GEntryOrderTypeToShortString = "MIT"
Case EntryOrderTypeMarketToLimit
    GEntryOrderTypeToShortString = "MTL"
Case EntryOrderTypeBid
    GEntryOrderTypeToShortString = "BID"
Case EntryOrderTypeAsk
    GEntryOrderTypeToShortString = "ASK"
Case EntryOrderTypeLast
    GEntryOrderTypeToShortString = "LAST"
Case EntryOrderTypeLimit
    GEntryOrderTypeToShortString = "LMT"
Case EntryOrderTypeLimitOnOpen
    GEntryOrderTypeToShortString = "LOO"
Case EntryOrderTypeLimitOnClose
    GEntryOrderTypeToShortString = "LOC"
Case EntryOrderTypeLimitIfTouched
    GEntryOrderTypeToShortString = "LIT"
Case EntryOrderTypeStop
    GEntryOrderTypeToShortString = "STP"
Case EntryOrderTypeStopLimit
    GEntryOrderTypeToShortString = "STPLMT"
End Select
End Function

Public Sub GHandleFatalError( _
                ByRef pProcName As String, _
                ByRef pModuleName As String, _
                Optional ByVal pFailpoint As Long)
On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=pProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=pModuleName, pFailpoint:=pFailpoint

' NB: will never get to here so no need for Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, Err.source
End Sub

Public Function GIsValidTIF(ByVal value As OrderTifs) As Boolean
Select Case value
Case TIFDay
    GIsValidTIF = True
Case TIFGoodTillCancelled
    GIsValidTIF = True
Case TIFImmediateOrCancel
    GIsValidTIF = True
Case Else
    GIsValidTIF = False
End Select
End Function

Public Function GLegOpenCloseFromString(ByVal value As String) As LegOpenClose
Select Case UCase$(value)
Case ""
    GLegOpenCloseFromString = LegUnknownPos
Case "SAME"
    GLegOpenCloseFromString = LegSamePos
Case "OPEN"
    GLegOpenCloseFromString = LegOpenPos
Case "CLOSE"
    GLegOpenCloseFromString = LegClosePos
End Select
End Function

Public Function GLegOpenCloseToString(ByVal value As LegOpenClose) As String
Select Case value
Case LegSamePos
    GLegOpenCloseToString = "SAME"
Case LegOpenPos
    GLegOpenCloseToString = "OPEN"
Case LegClosePos
    GLegOpenCloseToString = "CLOSE"
End Select
End Function

Public Function GOrderActionFromString(ByVal value As String) As OrderActions
Select Case UCase$(value)
Case StrOrderActionBuy
    GOrderActionFromString = ActionBuy
Case StrOrderActionSell
    GOrderActionFromString = ActionSell
End Select
End Function

Public Function GOrderActionToString(ByVal value As OrderActions) As String
Select Case value
Case ActionBuy
    GOrderActionToString = StrOrderActionBuy
Case ActionSell
    GOrderActionToString = StrOrderActionSell
End Select
End Function

Public Function GOrderAttributeToString(ByVal value As OrderAttributes) As String
Select Case value
    Case OrderAttOpenClose
        GOrderAttributeToString = "OpenClose"
    Case OrderAttOrigin
        GOrderAttributeToString = "Origin"
    Case OrderAttOriginatorRef
        GOrderAttributeToString = "OriginatorRef"
    Case OrderAttBlockOrder
        GOrderAttributeToString = "BlockOrder"
    Case OrderAttSweepToFill
        GOrderAttributeToString = "SweepToFill"
    Case OrderAttDisplaySize
        GOrderAttributeToString = "DisplaySize"
    Case OrderAttIgnoreRTH
        GOrderAttributeToString = "IgnoreRTH"
    Case OrderAttHidden
        GOrderAttributeToString = "Hidden"
    Case OrderAttDiscretionaryAmount
        GOrderAttributeToString = "DiscretionaryAmount"
    Case OrderAttGoodAfterTime
        GOrderAttributeToString = "GoodAfterTime"
    Case OrderAttGoodTillDate
        GOrderAttributeToString = "GoodTillDate"
    Case OrderAttRTHOnly
        GOrderAttributeToString = "RTHOnly"
    Case OrderAttRule80A
        GOrderAttributeToString = "Rule80A"
    Case OrderAttSettlingFirm
        GOrderAttributeToString = "SettlingFirm"
    Case OrderAttAllOrNone
        GOrderAttributeToString = "AllOrNone"
    Case OrderAttMinimumQuantity
        GOrderAttributeToString = "MinimumQuantity"
    Case OrderAttPercentOffset
        GOrderAttributeToString = "PercentOffset"
    Case OrderAttETradeOnly
        GOrderAttributeToString = "ETradeOnly"
    Case OrderAttFirmQuoteOnly
        GOrderAttributeToString = "FirmQuoteOnly"
    Case OrderAttNBBOPriceCap
        GOrderAttributeToString = "NBBOPriceCap"
    Case OrderAttOverrideConstraints
        GOrderAttributeToString = "OverrideConstraints"
    Case OrderAttAction
        GOrderAttributeToString = "Action"
    Case OrderAttLimitPrice
        GOrderAttributeToString = "LimitPrice"
    Case OrderAttOrderType
        GOrderAttributeToString = "OrderType"
    Case OrderAttQuantity
        GOrderAttributeToString = "Quantity"
    Case OrderAttTimeInForce
        GOrderAttributeToString = "TimeInForce"
    Case OrderAttTriggerPrice
        GOrderAttributeToString = "TriggerPrice"
    Case OrderAttGoodAfterTimeTZ
        GOrderAttributeToString = "GoodAfterTimeTZ"
    Case OrderAttGoodTillDateTZ
        GOrderAttributeToString = "GoodTillDateTZ"
    Case OrderAttStopTriggerMethod
        GOrderAttributeToString = "StopTriggerMethod"
    Case Else
        GOrderAttributeToString = "***Unknown order attribute***"
End Select
End Function

Public Function GOrderStatusToString(ByVal value As OrderStatuses) As String
Select Case UCase$(value)
Case OrderStatusCreated
    GOrderStatusToString = "Created"
Case OrderStatusPendingSubmit
    GOrderStatusToString = "Pending Submit"
Case OrderStatusPreSubmitted
    GOrderStatusToString = "Presubmitted"
Case OrderStatusSubmitted
    GOrderStatusToString = "Submitted"
Case OrderStatusCancelling
    GOrderStatusToString = "Cancelling"
Case OrderStatusCancelled
    GOrderStatusToString = "Cancelled"
Case OrderStatusFilled
    GOrderStatusToString = "Filled"
End Select
End Function

Public Function GOrderStopTriggerMethodToString(ByVal value As StopTriggerMethods) As String
Select Case value
Case StopTriggerMethods.StopTriggerBidAsk
    GOrderStopTriggerMethodToString = "Bid/Ask"
Case StopTriggerMethods.StopTriggerDefault
    GOrderStopTriggerMethodToString = "Default"
Case StopTriggerMethods.StopTriggerDoubleBidAsk
    GOrderStopTriggerMethodToString = "Double Bid/Ask"
Case StopTriggerMethods.StopTriggerDoubleLast
    GOrderStopTriggerMethodToString = "Double last"
Case StopTriggerMethods.StopTriggerLast
    GOrderStopTriggerMethodToString = "Last"
Case StopTriggerMethods.StopTriggerLastOrBidAsk
    GOrderStopTriggerMethodToString = "Last or Bid/Ask"
Case StopTriggerMethods.StopTriggerMidPoint
    GOrderStopTriggerMethodToString = "Mid-point"
End Select
End Function

Public Function GOrderTIFToString(ByVal value As OrderTifs) As String
Select Case value
Case TIFDay
    GOrderTIFToString = "DAY"
Case TIFGoodTillCancelled
    GOrderTIFToString = "GTC"
Case TIFImmediateOrCancel
    GOrderTIFToString = "IOC"
End Select
End Function

Public Function GOrderTypeToString(ByVal value As OrderTypes) As String
Select Case value
Case OrderTypeNone
    GOrderTypeToString = StrOrderTypeNone
Case OrderTypeMarket
    GOrderTypeToString = StrOrderTypeMarket
Case OrderTypeMarketOnClose
    GOrderTypeToString = StrOrderTypeMarketClose
Case OrderTypeLimit
    GOrderTypeToString = StrOrderTypeLimit
Case OrderTypeLimitOnClose
    GOrderTypeToString = StrOrderTypeLimitClose
Case OrderTypePeggedToMarket
    GOrderTypeToString = StrOrderTypePegMarket
Case OrderTypeStop
    GOrderTypeToString = StrOrderTypeStop
Case OrderTypeStopLimit
    GOrderTypeToString = StrOrderTypeStopLimit
Case OrderTypeTrail
    GOrderTypeToString = StrOrderTypeTrail
Case OrderTypeRelative
    GOrderTypeToString = StrOrderTypeRelative
Case OrderTypeVWAP
    GOrderTypeToString = StrOrderTypeVWAP
Case OrderTypeMarketToLimit
    GOrderTypeToString = StrOrderTypeMarketToLimit
Case OrderTypeQuote
    GOrderTypeToString = StrOrderTypeQuote
Case OrderTypeAdjust
    GOrderTypeToString = StrOrderTypeAdjust
Case OrderTypeAlert
    GOrderTypeToString = StrOrderTypeAlert
Case OrderTypeLimitIfTouched
    GOrderTypeToString = StrOrderTypeLimitIfTouched
Case OrderTypeMarketIfTouched
    GOrderTypeToString = StrOrderTypeMarketIfTouched
Case OrderTypeTrailLimit
    GOrderTypeToString = StrOrderTypeTrailLimit
Case OrderTypeMarketWithProtection
    GOrderTypeToString = StrOrderTypeMarketWithProtection
Case OrderTypeMarketOnOpen
    GOrderTypeToString = StrOrderTypeMarketOnOpen
Case OrderTypeLimitOnOpen
    GOrderTypeToString = StrOrderTypeLimitOnOpen
Case OrderTypePeggedToPrimary
    GOrderTypeToString = StrOrderTypePeggedToPrimary
End Select

End Function

''
' Converts a member of the StopOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pStopOrderType the StopOrderTypes value to be converted
'
'@/
Public Function GStopOrderTypeToOrderType( _
                ByVal pStopOrderType As StopOrderTypes) As OrderTypes
Select Case pStopOrderType
Case StopOrderTypeNone
    GStopOrderTypeToOrderType = OrderTypeNone
Case StopOrderTypeStop
    GStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeStopLimit
    GStopOrderTypeToOrderType = OrderTypeStopLimit
Case StopOrderTypeBid
    GStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeAsk
    GStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeLast
    GStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeAuto
    GStopOrderTypeToOrderType = OrderTypeAutoStop
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Module1::gStopOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function GStopOrderTypeToShortString(ByVal value As StopOrderTypes)
Select Case value
Case StopOrderTypeNone
    GStopOrderTypeToShortString = "NONE"
Case StopOrderTypeStop
    GStopOrderTypeToShortString = "STP"
Case StopOrderTypeStopLimit
    GStopOrderTypeToShortString = "STPLMT"
Case StopOrderTypeBid
    GStopOrderTypeToShortString = "BID"
Case StopOrderTypeAsk
    GStopOrderTypeToShortString = "ASK"
Case StopOrderTypeLast
    GStopOrderTypeToShortString = "TRADE"
Case StopOrderTypeAuto
    GStopOrderTypeToShortString = "AUTO"
End Select
End Function

Public Function GStopOrderTypeToString(ByVal value As StopOrderTypes)
Select Case value
Case StopOrderTypeNone
    GStopOrderTypeToString = "None"
Case StopOrderTypeStop
    GStopOrderTypeToString = "Stop"
Case StopOrderTypeStopLimit
    GStopOrderTypeToString = "Stop limit"
Case StopOrderTypeBid
    GStopOrderTypeToString = "Bid price"
Case StopOrderTypeAsk
    GStopOrderTypeToString = "Ask price"
Case StopOrderTypeLast
    GStopOrderTypeToString = "LAST"
Case StopOrderTypeAuto
    GStopOrderTypeToString = "Auto"
End Select
End Function

''
' Converts a member of the TargetOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pTargetOrderType the TargetOrderTypes value to be converted
' @ see
'
'@/
Public Function GTargetOrderTypeToOrderType( _
                ByVal pTargetOrderType As TargetOrderTypes) As OrderTypes
Select Case pTargetOrderType
Case TargetOrderTypeNone
    GTargetOrderTypeToOrderType = OrderTypeNone
Case TargetOrderTypeLimit
    GTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeLimitIfTouched
    GTargetOrderTypeToOrderType = OrderTypeLimitIfTouched
Case TargetOrderTypeMarketIfTouched
    GTargetOrderTypeToOrderType = OrderTypeMarketIfTouched
Case TargetOrderTypeBid
    GTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeAsk
    GTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeLast
    GTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeAuto
    GTargetOrderTypeToOrderType = OrderTypeAutoLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Module1::gTargetOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function GTargetOrderTypeToShortString(ByVal value As TargetOrderTypes)
Select Case value
Case TargetOrderTypeNone
    GTargetOrderTypeToShortString = "NONE"
Case TargetOrderTypeLimit
    GTargetOrderTypeToShortString = "LMT"
Case TargetOrderTypeMarketIfTouched
    GTargetOrderTypeToShortString = "MIT"
Case TargetOrderTypeBid
    GTargetOrderTypeToShortString = "BID"
Case TargetOrderTypeAsk
    GTargetOrderTypeToShortString = "ASK"
Case TargetOrderTypeLast
    GTargetOrderTypeToShortString = "LAST"
Case TargetOrderTypeAuto
    GTargetOrderTypeToShortString = "AUTO"
End Select
End Function

Public Function GTargetOrderTypeToString(ByVal value As TargetOrderTypes)
Select Case value
Case TargetOrderTypeNone
    GTargetOrderTypeToString = "None"
Case TargetOrderTypeLimit
    GTargetOrderTypeToString = "Limit"
Case TargetOrderTypeMarketIfTouched
    GTargetOrderTypeToString = "Market if touched"
Case TargetOrderTypeBid
    GTargetOrderTypeToString = "Bid price"
Case TargetOrderTypeAsk
    GTargetOrderTypeToString = "Ask price"
Case TargetOrderTypeLast
    GTargetOrderTypeToString = "Last Trade price"
Case TargetOrderTypeAuto
    GTargetOrderTypeToString = "Auto"
End Select
End Function

Public Function GTickfileSpecifierToString(tfSpec As ITickfileSpecifier) As String
If tfSpec.Filename <> "" Then
    GTickfileSpecifierToString = tfSpec.Filename
Else
    GTickfileSpecifierToString = "Contract: " & _
                                Replace(tfSpec.Contract.specifier.ToString, vbCrLf, "; ") & _
                            ": From: " & FormatDateTime(tfSpec.FromDate, vbGeneralDate) & _
                            " To: " & FormatDateTime(tfSpec.ToDate, vbGeneralDate)
End If
End Function

Public Sub GLog(ByRef pMsg As String, _
                ByRef pProjName As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
If Not GLogLogger.IsLoggable(pLogLevel) Then Exit Sub
mLogTokens(0) = "["
mLogTokens(1) = pProjName
mLogTokens(2) = "."
mLogTokens(3) = pModName
mLogTokens(4) = ":"
mLogTokens(5) = pProcName
mLogTokens(6) = "] "
mLogTokens(7) = pMsg
If Len(pMsgQualifier) <> 0 Then
    mLogTokens(8) = ": "
    mLogTokens(9) = pMsgQualifier
Else
    mLogTokens(8) = vbNullString
    mLogTokens(9) = vbNullString
End If

GLogLogger.Log pLogLevel, Join(mLogTokens, "")
End Sub

Public Sub GLogData( _
                ByVal value As Variant, _
                ByVal valueType As Long, _
                ByVal source As Object)
Const ProcName As String = "GLogData"
Dim failpoint As Long
On Error GoTo Err

Select Case valueType
Case VTLog
    GLogLogger.Log LogLevelNormal, value
Case VTTrace
    GTraceLogger.Log LogLevelNormal, value, source
Case VTDebug
    GDebugLogger.Log LogLevelNormal, value, source
Case VTProfitProfile
    GProfitProfileLogger.Log LogLevelNormal, value, source
Case VTSimulatedProfitProfile
    GProfitProfileLoggerSimulated.Log LogLevelNormal, value, source
Case VTMoneyManagement
    GMoneyManagementLogger.Log LogLevelNormal, value, source
Case VTOrderPlexProfileStruct
    GOrderPlexProfileStructLogger.Log LogLevelNormal, value, source
Case VTSimulatedOrderPlexProfileStruct
    GOrderPlexProfileStructLoggerSimulated.Log LogLevelNormal, value, source
Case VTOrderPlexProfileString
    GOrderPlexProfileStringLogger.Log LogLevelNormal, value, source
Case VTSimulatedOrderPlexProfileString
    GOrderPlexProfileStringLoggerSimulated.Log LogLevelNormal, value, source
Case VTOrder
    GOrderLogger.Log LogLevelNormal, value, source
Case VTSimulatedOrder
    GOrderLoggerSimulated.Log LogLevelNormal, value, source
Case VTPosition
    GPositionLogger.Log LogLevelNormal, value, source
Case VTSimulatedPosition
    GPositionLoggerSimulated.Log LogLevelNormal, value, source
Case VTTradeProfile
    GTradeProfileLogger.Log LogLevelNormal, value, source
Case VTSimulatedTradeProfile
    GTradeProfileLoggerSimulated.Log LogLevelNormal, value, source
Case VTProfit
    GProfitLogger.Log LogLevelNormal, value, source
Case VTSimulatedProfit
    GProfitLoggerSimulated.Log LogLevelNormal, value, source
Case VTDrawdown
    GDrawdownLogger.Log LogLevelNormal, value, source
Case VTSimulatedDrawdown
    GDrawdownLoggerSimulated.Log LogLevelNormal, value, source
Case VTMaxProfit
    GMaxProfitLogger.Log LogLevelNormal, value, source
Case VTSimulatedMaxProfit
    GMaxProfitLoggerSimulated.Log LogLevelNormal, value, source
Case VTOrderDetail
    GOrderDetailLogger.Log LogLevelNormal, value, source
Case VTOrderDetailSimulated
    GOrderDetailLoggerSimulated.Log LogLevelNormal, value, source
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub
                            
Public Property Get GLogLogger() As Logger
If mLogLogger Is Nothing Then
    Set mLogLogger = GetLogger("tradebuild.log")
End If
Set GLogLogger = mLogLogger
End Property

Public Property Get GErrorLogger() As Logger
If mErrorLogger Is Nothing Then
    Set mErrorLogger = GetLogger("error")
End If
Set GErrorLogger = mErrorLogger
End Property

Public Property Get GSpLogLogger() As Logger
If mSpLogLogger Is Nothing Then
    Set mSpLogLogger = GetLogger("tradebuild.log.serviceprovider")
End If
Set GSpLogLogger = mSpLogLogger
End Property

Public Property Get GTraceLogger() As Logger
If mTraceLogger Is Nothing Then
    Set mTraceLogger = GetLogger("tradebuild.trace")
End If
Set GTraceLogger = mTraceLogger
End Property

Public Property Get GDebugLogger() As Logger
If mDebugLogger Is Nothing Then
    Set mDebugLogger = GetLogger("tradebuild.debug")
End If
Set GDebugLogger = mDebugLogger
End Property

Public Property Get GProfitProfileLogger() As Logger
If mProfitProfileLogger Is Nothing Then
    Set mProfitProfileLogger = GetLogger("tradebuild.ProfitProfile")
    mProfitProfileLogger.LogToParent = False
End If
Set GProfitProfileLogger = mProfitProfileLogger
End Property

Public Property Get GProfitProfileLoggerSimulated() As Logger
If mProfitProfileLoggerSimulated Is Nothing Then
    Set mProfitProfileLoggerSimulated = GetLogger("tradebuild.ProfitProfileSimulated")
    mProfitProfileLoggerSimulated.LogToParent = False
End If
Set GProfitProfileLoggerSimulated = mProfitProfileLoggerSimulated
End Property

Public Property Get GMoneyManagementLogger() As Logger
If mMoneyManagementLogger Is Nothing Then
    Set mMoneyManagementLogger = GetLogger("tradebuild.MoneyManagement")
End If
Set GMoneyManagementLogger = mMoneyManagementLogger
End Property

Public Property Get GOrderPlexProfileStructLogger() As Logger
If mOrderPlexProfileStructLogger Is Nothing Then
    Set mOrderPlexProfileStructLogger = GetLogger("tradebuild.OrderPlexProfileStruct")
    mOrderPlexProfileStructLogger.LogToParent = False
End If
Set GOrderPlexProfileStructLogger = mOrderPlexProfileStructLogger
End Property

Public Property Get GOrderPlexProfileStructLoggerSimulated() As Logger
If mOrderPlexProfileStructLoggerSimulated Is Nothing Then
    Set mOrderPlexProfileStructLoggerSimulated = GetLogger("tradebuild.SimulatedOrderPlexProfileStructSimulated")
    mOrderPlexProfileStructLoggerSimulated.LogToParent = False
End If
Set GOrderPlexProfileStructLoggerSimulated = mOrderPlexProfileStructLoggerSimulated
End Property

Public Property Get GOrderPlexProfileStringLogger() As Logger
If mOrderPlexProfileStringLogger Is Nothing Then
    Set mOrderPlexProfileStringLogger = GetLogger("tradebuild.OrderPlexProfileString")
    mOrderPlexProfileStringLogger.LogToParent = False
End If
Set GOrderPlexProfileStringLogger = mOrderPlexProfileStringLogger
End Property

Public Property Get GOrderPlexProfileStringLoggerSimulated() As Logger
If mOrderPlexProfileStringLoggerSimulated Is Nothing Then
    Set mOrderPlexProfileStringLoggerSimulated = GetLogger("tradebuild.OrderPlexProfileStringSimulated")
    mOrderPlexProfileStringLoggerSimulated.LogToParent = False
End If
Set GOrderPlexProfileStringLoggerSimulated = mOrderPlexProfileStringLoggerSimulated
End Property

Public Property Get GOrderLogger() As Logger
If mOrderLogger Is Nothing Then
    Set mOrderLogger = GetLogger("tradebuild.order")
End If
Set GOrderLogger = mOrderLogger
End Property

Public Property Get GOrderLoggerSimulated() As Logger
If mOrderLoggerSimulated Is Nothing Then
    Set mOrderLoggerSimulated = GetLogger("tradebuild.orderSimulated")
End If
Set GOrderLoggerSimulated = mOrderLoggerSimulated
End Property

Public Property Get GPositionLogger() As Logger
If mPositionLogger Is Nothing Then
    Set mPositionLogger = GetLogger("tradebuild.position")
    mPositionLogger.LogToParent = False
End If
Set GPositionLogger = mPositionLogger
End Property

Public Property Get GPositionLoggerSimulated() As Logger
If mPositionLoggerSimulated Is Nothing Then
    Set mPositionLoggerSimulated = GetLogger("tradebuild.positionSimulated")
    mPositionLoggerSimulated.LogToParent = False
End If
Set GPositionLoggerSimulated = mPositionLoggerSimulated
End Property

Public Property Get GTradeProfileLogger() As Logger
If mTradeProfileLogger Is Nothing Then
    Set mTradeProfileLogger = GetLogger("tradebuild.TradeProfile")
    mTradeProfileLogger.LogToParent = False
End If
Set GTradeProfileLogger = mTradeProfileLogger
End Property

Public Property Get GTradeProfileLoggerSimulated() As Logger
If mTradeProfileLoggerSimulated Is Nothing Then
    Set mTradeProfileLoggerSimulated = GetLogger("tradebuild.TradeProfileSimulated")
    mTradeProfileLoggerSimulated.LogToParent = False
End If
Set GTradeProfileLoggerSimulated = mTradeProfileLoggerSimulated
End Property

Public Property Get GProfitLogger() As Logger
If mProfitLogger Is Nothing Then
    Set mProfitLogger = GetLogger("tradebuild.Profit")
    mProfitLogger.LogToParent = False
End If
Set GProfitLogger = mProfitLogger
End Property

Public Property Get GProfitLoggerSimulated() As Logger
If mProfitLoggerSimulated Is Nothing Then
    Set mProfitLoggerSimulated = GetLogger("tradebuild.profitSimulated")
    mProfitLoggerSimulated.LogToParent = False
End If
Set GProfitLoggerSimulated = mProfitLoggerSimulated
End Property

Public Property Get GDrawdownLogger() As Logger
If mDrawdownLogger Is Nothing Then
    Set mDrawdownLogger = GetLogger("tradebuild.Drawdown")
    mDrawdownLogger.LogToParent = False
End If
Set GDrawdownLogger = mDrawdownLogger
End Property

Public Property Get GDrawdownLoggerSimulated() As Logger
If mDrawdownLoggerSimulated Is Nothing Then
    Set mDrawdownLoggerSimulated = GetLogger("tradebuild.drawdownSimulated")
    mDrawdownLoggerSimulated.LogToParent = False
End If
Set GDrawdownLoggerSimulated = mDrawdownLoggerSimulated
End Property

Public Property Get GMaxProfitLogger() As Logger
If mMaxProfitLogger Is Nothing Then
    Set mMaxProfitLogger = GetLogger("tradebuild.MaxProfit")
    mMaxProfitLogger.LogToParent = False
End If
Set GMaxProfitLogger = mMaxProfitLogger
End Property

Public Property Get GMaxProfitLoggerSimulated() As Logger
If mMaxProfitLoggerSimulated Is Nothing Then
    Set mMaxProfitLoggerSimulated = GetLogger("tradebuild.MaxProfitSimulated")
    mMaxProfitLoggerSimulated.LogToParent = False
End If
Set GMaxProfitLoggerSimulated = mMaxProfitLoggerSimulated
End Property

Public Property Get GOrderDetailLogger() As Logger
If mOrderDetailLogger Is Nothing Then
    Set mOrderDetailLogger = GetLogger("tradebuild.orderdetail")
    mOrderDetailLogger.LogToParent = False
End If
Set GOrderDetailLogger = mOrderDetailLogger
End Property

Public Property Get GOrderDetailLoggerSimulated() As Logger
If mOrderDetailLoggerSimulated Is Nothing Then
    Set mOrderDetailLoggerSimulated = GetLogger("tradebuild.orderdetailSimulated")
    mOrderDetailLoggerSimulated.LogToParent = False
End If
Set GOrderDetailLoggerSimulated = mOrderDetailLoggerSimulated
End Property

Public Property Get GTracer() As Tracer
If mTracer Is Nothing Then Set mTracer = GetTracer("tradebuild")
Set GTracer = mTracer
End Property

