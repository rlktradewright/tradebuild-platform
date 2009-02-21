Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                    As String = "TradeBuild26"

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
Public Const TickVolumeInputName            As String = "Tick volume"
Public Const VolumeInputName                As String = "Total volume"

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

'@================================================================================
' Procedures
'@================================================================================

Public Function gApiNotifyCodeToString(value As ApiNotifyCodes) As String
Select Case value
Case ApiNotifyServiceProviderError
    gApiNotifyCodeToString = "ApiNotifyServiceProviderError"
Case ApiNotifyTickfileEmpty
    gApiNotifyCodeToString = "ApiNotifyTickfileEmpty"
Case ApiNotifyTickfileInvalid
    gApiNotifyCodeToString = "ApiNotifyTickfileInvalid"
Case ApiNotifyTickfileVersionNotSupported
    gApiNotifyCodeToString = "ApiNotifyTickfileVersionNotSupported"
Case ApiNotifyContractDetailsInvalid
    gApiNotifyCodeToString = "ApiNotifyContractDetailsInvalid"
Case ApiNotifyNoContractDetails
    gApiNotifyCodeToString = "ApiNotifyNoContractDetails"
Case ApiNotifyTickfileDataSourceNotResponding
    gApiNotifyCodeToString = "ApiNotifyTickfileDataSourceNotResponding"
Case ApiNotifyCantWriteToTickfileDataStore
    gApiNotifyCodeToString = "ApiNotifyCantWriteToTickfileDataStore"
Case ApiNotifyRetryingConnectionToTickfileDataSource
    gApiNotifyCodeToString = "ApiNotifyRetryingConnectionToTickfileDataSource"
Case ApiNotifyConnectedToTickfileDataSource
    gApiNotifyCodeToString = "ApiNotifyConnectedToTickfileDataSource"
Case ApiNotifyNoHistoricDataSource
    gApiNotifyCodeToString = "ApiNotifyNoHistoricDataSource"
Case ApiNotifyCantConnectHistoricDataSource
    gApiNotifyCodeToString = "ApiNotifyCantConnectHistoricDataSource"
Case ApiNotifyConnectedToHistoricDataSource
    gApiNotifyCodeToString = "ApiNotifyConnectedToHistoricDataSource"
Case ApiNotifyRetryingConnectionToHistoricDataSource
    gApiNotifyCodeToString = "ApiNotifyRetryingConnectionToHistoricDataSource"
Case ApiNotifyContractSpecifierAmbiguous
    gApiNotifyCodeToString = "ApiNotifyContractSpecifierAmbiguous"
Case ApiNotifyContractSpecifierInvalid
    gApiNotifyCodeToString = "ApiNotifyContractSpecifierInvalid"
Case ApiNotifyMarketDepthNotAvailable
    gApiNotifyCodeToString = "ApiNotifyMarketDepthNotAvailable"
Case ApiNotifyInvalidRequest
    gApiNotifyCodeToString = "ApiNotifyInvalidRequest"
Case ApiNotifyFatalError
    gApiNotifyCodeToString = "ApiNotifyFatalError"
Case ApiNotifyCantConnectRealtimeDataSource
    gApiNotifyCodeToString = "ApiNotifyCantConnectRealtimeDataSource"
Case ApiNotifyConnectedToRealtimeDataSource
    gApiNotifyCodeToString = "ApiNotifyConnectedToRealtimeDataSource"
Case ApiNotifyLostConnectionToRealtimeDataSource
    gApiNotifyCodeToString = "ApiNotifyLostConnectionToRealtimeDataSource"
Case ApiNotifyNoRealtimeDataSource
    gApiNotifyCodeToString = "ApiNotifyNoRealtimeDataSource"
Case ApiNotifyReconnectingToRealtimeDataSource
    gApiNotifyCodeToString = "ApiNotifyReconnectingToRealtimeDataSource"
Case ApiNotifyDisconnectedFromRealtimeDataSource
    gApiNotifyCodeToString = "ApiNotifyDisconnectedFromRealtimeDataSource"
Case ApiNotifyRealtimeDataRequestFailed
    gApiNotifyCodeToString = "ApiNotifyRealtimeDataRequestFailed"
Case ApiNotifyCantConnectToBroker
    gApiNotifyCodeToString = "ApiNotifyCantConnectToBroker"
Case ApiNotifyConnectedToBroker
    gApiNotifyCodeToString = "ApiNotifyConnectedToBroker"
Case ApiNotifyRetryConnectToBroker
    gApiNotifyCodeToString = "ApiNotifyRetryConnectToBroker"
Case ApiNotifyLostConnectionToBroker
    gApiNotifyCodeToString = "ApiNotifyLostConnectionToBroker"
Case ApiNotifyReConnectingToBroker
    gApiNotifyCodeToString = "ApiNotifyReConnectingToBroker"
Case ApiNotifyInvalidOrder
    gApiNotifyCodeToString = "ApiNotifyInvalidOrder"
Case ApiNotifyInsufficientFunds
    gApiNotifyCodeToString = "ApiNotifyInsufficientFunds"
Case ApiNotifyDisconnectedFromBroker
    gApiNotifyCodeToString = "ApiNotifyDisconnectedFromBroker"
Case ApiNotifyNonSpecificNotification
    gApiNotifyCodeToString = "ApiNotifyNonSpecificNotification"
Case ApiNotifyContractExpired
    gApiNotifyCodeToString = "ApiNotifyContractExpired"
Case ApiNOtifyTickfileFormatNotSupported
    gApiNotifyCodeToString = "ApiNOtifyTickfileFormatNotSupported"
Case ApiNotifyContractDoesNotExist
    gApiNotifyCodeToString = "ApiNotifyContractDoesNotExist"
Case ApiNotifyCantWriteToHistoricDataStore
    gApiNotifyCodeToString = "ApiNotifyCantWriteToHistoricDataStore"
Case ApiNotifyTryLater
    gApiNotifyCodeToString = "ApiNotifyTryLater"
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
Public Function gEntryOrderTypeToOrderType( _
                ByVal pEntryOrderType As EntryOrderTypes) As OrderTypes
Select Case pEntryOrderType
Case EntryOrderTypeMarket
    gEntryOrderTypeToOrderType = OrderTypeMarket
Case EntryOrderTypeMarketOnOpen
    gEntryOrderTypeToOrderType = OrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
    gEntryOrderTypeToOrderType = OrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
    gEntryOrderTypeToOrderType = OrderTypeMarketIfTouched
Case EntryOrderTypeMarketToLimit
    gEntryOrderTypeToOrderType = OrderTypeMarketToLimit
Case EntryOrderTypeBid
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeAsk
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLast
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLimit
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLimitOnOpen
    gEntryOrderTypeToOrderType = OrderTypeLimitOnOpen
Case EntryOrderTypeLimitOnClose
    gEntryOrderTypeToOrderType = OrderTypeLimitOnClose
Case EntryOrderTypeLimitIfTouched
    gEntryOrderTypeToOrderType = OrderTypeLimitIfTouched
Case EntryOrderTypeStop
    gEntryOrderTypeToOrderType = OrderTypeStop
Case EntryOrderTypeStopLimit
    gEntryOrderTypeToOrderType = OrderTypeStopLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Module1::gEntryOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function gEntryOrderTypeToString(ByVal value As EntryOrderTypes) As String
Select Case value
Case EntryOrderTypeMarket
    gEntryOrderTypeToString = "Market"
Case EntryOrderTypeMarketOnOpen
    gEntryOrderTypeToString = "Market on open"
Case EntryOrderTypeMarketOnClose
    gEntryOrderTypeToString = "Market on close"
Case EntryOrderTypeMarketIfTouched
    gEntryOrderTypeToString = "Market if touched"
Case EntryOrderTypeMarketToLimit
    gEntryOrderTypeToString = "Market to limit"
Case EntryOrderTypeBid
    gEntryOrderTypeToString = "Bid price"
Case EntryOrderTypeAsk
    gEntryOrderTypeToString = "Ask price"
Case EntryOrderTypeLast
    gEntryOrderTypeToString = "Last trade price"
Case EntryOrderTypeLimit
    gEntryOrderTypeToString = "Limit"
Case EntryOrderTypeLimitOnOpen
    gEntryOrderTypeToString = "Limit on open"
Case EntryOrderTypeLimitOnClose
    gEntryOrderTypeToString = "Limit on close"
Case EntryOrderTypeLimitIfTouched
    gEntryOrderTypeToString = "Limit if touched"
Case EntryOrderTypeStop
    gEntryOrderTypeToString = "Stop"
Case EntryOrderTypeStopLimit
    gEntryOrderTypeToString = "Stop limit"
End Select
End Function

Public Function gEntryOrderTypeToShortString(ByVal value As EntryOrderTypes) As String
Select Case value
Case EntryOrderTypeMarket
    gEntryOrderTypeToShortString = "MKT"
Case EntryOrderTypeMarketOnOpen
    gEntryOrderTypeToShortString = "MOO"
Case EntryOrderTypeMarketOnClose
    gEntryOrderTypeToShortString = "MOC"
Case EntryOrderTypeMarketIfTouched
    gEntryOrderTypeToShortString = "MIT"
Case EntryOrderTypeMarketToLimit
    gEntryOrderTypeToShortString = "MTL"
Case EntryOrderTypeBid
    gEntryOrderTypeToShortString = "BID"
Case EntryOrderTypeAsk
    gEntryOrderTypeToShortString = "ASK"
Case EntryOrderTypeLast
    gEntryOrderTypeToShortString = "LAST"
Case EntryOrderTypeLimit
    gEntryOrderTypeToShortString = "LMT"
Case EntryOrderTypeLimitOnOpen
    gEntryOrderTypeToShortString = "LOO"
Case EntryOrderTypeLimitOnClose
    gEntryOrderTypeToShortString = "LOC"
Case EntryOrderTypeLimitIfTouched
    gEntryOrderTypeToShortString = "LIT"
Case EntryOrderTypeStop
    gEntryOrderTypeToShortString = "STP"
Case EntryOrderTypeStopLimit
    gEntryOrderTypeToShortString = "STPLMT"
End Select
End Function

Public Function gLegOpenCloseFromString(ByVal value As String) As LegOpenClose
Select Case UCase$(value)
Case ""
    gLegOpenCloseFromString = LegUnknownPos
Case "SAME"
    gLegOpenCloseFromString = LegSamePos
Case "OPEN"
    gLegOpenCloseFromString = LegOpenPos
Case "CLOSE"
    gLegOpenCloseFromString = LegClosePos
End Select
End Function

Public Function gLegOpenCloseToString(ByVal value As LegOpenClose) As String
Select Case value
Case LegSamePos
    gLegOpenCloseToString = "SAME"
Case LegOpenPos
    gLegOpenCloseToString = "OPEN"
Case LegClosePos
    gLegOpenCloseToString = "CLOSE"
End Select
End Function

Public Function gOrderActionFromString(ByVal value As String) As OrderActions
Select Case UCase$(value)
Case StrOrderActionBuy
    gOrderActionFromString = ActionBuy
Case StrOrderActionSell
    gOrderActionFromString = ActionSell
End Select
End Function

Public Function gOrderActionToString(ByVal value As OrderActions) As String
Select Case value
Case ActionBuy
    gOrderActionToString = StrOrderActionBuy
Case ActionSell
    gOrderActionToString = StrOrderActionSell
End Select
End Function

Public Function gOrderAttributeToString(ByVal value As OrderAttributes) As String
Select Case value
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
    Case OrderAttRTHOnly
        gOrderAttributeToString = "RTHOnly"
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
End Function

Public Function gOrderStatusToString(ByVal value As OrderStatuses) As String
Select Case UCase$(value)
Case OrderStatusCreated
    gOrderStatusToString = "Created"
Case OrderStatusPendingSubmit
    gOrderStatusToString = "Pending Submit"
Case OrderStatusPreSubmitted
    gOrderStatusToString = "Presubmitted"
Case OrderStatusSubmitted
    gOrderStatusToString = "Submitted"
Case OrderStatusCancelling
    gOrderStatusToString = "Cancelling"
Case OrderStatusCancelled
    gOrderStatusToString = "Cancelled"
Case OrderStatusFilled
    gOrderStatusToString = "Filled"
End Select
End Function

Public Function gOrderStopTriggerMethodToString(ByVal value As StopTriggerMethods) As String
Select Case value
Case StopTriggerMethods.StopTriggerBidAsk
    gOrderStopTriggerMethodToString = "Bid/ask"
Case StopTriggerMethods.StopTriggerDefault
    gOrderStopTriggerMethodToString = "Default"
Case StopTriggerMethods.StopTriggerDoubleBidAsk
    gOrderStopTriggerMethodToString = "Double bid/ask"
Case StopTriggerMethods.StopTriggerDoubleLast
    gOrderStopTriggerMethodToString = "Double last"
Case StopTriggerMethods.StopTriggerLast
    gOrderStopTriggerMethodToString = "Last"
Case StopTriggerMethods.StopTriggerLastOrBidAsk
    gOrderStopTriggerMethodToString = "Last or bid/ask"
Case StopTriggerMethods.StopTriggerMidPoint
    gOrderStopTriggerMethodToString = "Mid-point"
End Select
End Function

Public Function gOrderTIFToString(ByVal value As OrderTifs) As String
Select Case value
Case TIFDay
    gOrderTIFToString = "DAY"
Case TIFGoodTillCancelled
    gOrderTIFToString = "GTC"
Case TIFImmediateOrCancel
    gOrderTIFToString = "IOC"
End Select
End Function

Public Function gOrderTypeToString(ByVal value As OrderTypes) As String
Select Case value
Case OrderTypeNone
    gOrderTypeToString = StrOrderTypeNone
Case OrderTypeMarket
    gOrderTypeToString = StrOrderTypeMarket
Case OrderTypeMarketOnClose
    gOrderTypeToString = StrOrderTypeMarketClose
Case OrderTypeLimit
    gOrderTypeToString = StrOrderTypeLimit
Case OrderTypeLimitOnClose
    gOrderTypeToString = StrOrderTypeLimitClose
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
End Select

End Function

''
' Converts a member of the StopOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pStopOrderType the StopOrderTypes value to be converted
'
'@/
Public Function gStopOrderTypeToOrderType( _
                ByVal pStopOrderType As StopOrderTypes) As OrderTypes
Select Case pStopOrderType
Case StopOrderTypeNone
    gStopOrderTypeToOrderType = OrderTypeNone
Case StopOrderTypeStop
    gStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeStopLimit
    gStopOrderTypeToOrderType = OrderTypeStopLimit
Case StopOrderTypeBid
    gStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeAsk
    gStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeLast
    gStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeAuto
    gStopOrderTypeToOrderType = OrderTypeAutoStop
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Module1::gStopOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function gStopOrderTypeToString(ByVal value As StopOrderTypes)
Select Case value
Case StopOrderTypeNone
    gStopOrderTypeToString = "None"
Case StopOrderTypeStop
    gStopOrderTypeToString = "Stop"
Case StopOrderTypeStopLimit
    gStopOrderTypeToString = "Stop limit"
Case StopOrderTypeBid
    gStopOrderTypeToString = "Bid price"
Case StopOrderTypeAsk
    gStopOrderTypeToString = "Ask price"
Case StopOrderTypeLast
    gStopOrderTypeToString = "Last trade price"
Case StopOrderTypeAuto
    gStopOrderTypeToString = "Auto"
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
Public Function gTargetOrderTypeToOrderType( _
                ByVal pTargetOrderType As TargetOrderTypes) As OrderTypes
Select Case pTargetOrderType
Case TargetOrderTypeNone
    gTargetOrderTypeToOrderType = OrderTypeNone
Case TargetOrderTypeLimit
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeLimitIfTouched
    gTargetOrderTypeToOrderType = OrderTypeLimitIfTouched
Case TargetOrderTypeMarketIfTouched
    gTargetOrderTypeToOrderType = OrderTypeMarketIfTouched
Case TargetOrderTypeBid
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeAsk
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeLast
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeAuto
    gTargetOrderTypeToOrderType = OrderTypeAutoLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Module1::gTargetOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function gTargetOrderTypeToString(ByVal value As TargetOrderTypes)
Select Case value
Case TargetOrderTypeNone
    gTargetOrderTypeToString = "None"
Case TargetOrderTypeLimit
    gTargetOrderTypeToString = "Limit"
Case TargetOrderTypeMarketIfTouched
    gTargetOrderTypeToString = "Market if touched"
Case TargetOrderTypeBid
    gTargetOrderTypeToString = "Bid price"
Case TargetOrderTypeAsk
    gTargetOrderTypeToString = "Ask price"
Case TargetOrderTypeLast
    gTargetOrderTypeToString = "Last trade price"
Case TargetOrderTypeAuto
    gTargetOrderTypeToString = "Auto"
End Select
End Function

Public Function gTickfileSpecifierToString(tfSpec As ITickfileSpecifier) As String
If tfSpec.filename <> "" Then
    gTickfileSpecifierToString = tfSpec.filename
Else
    gTickfileSpecifierToString = "Contract: " & _
                                Replace(tfSpec.Contract.specifier.toString, vbCrLf, "; ") & _
                            ": From: " & FormatDateTime(tfSpec.FromDate, vbGeneralDate) & _
                            " To: " & FormatDateTime(tfSpec.ToDate, vbGeneralDate)
End If
End Function

Public Property Get gLogLogger() As Logger
If mLogLogger Is Nothing Then
    Set mLogLogger = GetLogger("tradebuild.log")
End If
Set gLogLogger = mLogLogger
End Property

Public Property Get gSpLogLogger() As Logger
If mSpLogLogger Is Nothing Then
    Set mSpLogLogger = GetLogger("tradebuild.log.serviceprovider")
End If
Set gSpLogLogger = mSpLogLogger
End Property

Public Property Get gTraceLogger() As Logger
If mTraceLogger Is Nothing Then
    Set mTraceLogger = GetLogger("tradebuild.trace")
End If
Set gTraceLogger = mTraceLogger
End Property

Public Property Get gDebugLogger() As Logger
If mDebugLogger Is Nothing Then
    Set mDebugLogger = GetLogger("tradebuild.debug")
End If
Set gDebugLogger = mDebugLogger
End Property

Public Property Get gProfitProfileLogger() As Logger
If mProfitProfileLogger Is Nothing Then
    Set mProfitProfileLogger = GetLogger("tradebuild.ProfitProfile")
    mProfitProfileLogger.logToParent = False
End If
Set gProfitProfileLogger = mProfitProfileLogger
End Property

Public Property Get gProfitProfileLoggerSimulated() As Logger
If mProfitProfileLoggerSimulated Is Nothing Then
    Set mProfitProfileLoggerSimulated = GetLogger("tradebuild.ProfitProfileSimulated")
    mProfitProfileLoggerSimulated.logToParent = False
End If
Set gProfitProfileLoggerSimulated = mProfitProfileLoggerSimulated
End Property

Public Property Get gMoneyManagementLogger() As Logger
If mMoneyManagementLogger Is Nothing Then
    Set mMoneyManagementLogger = GetLogger("tradebuild.MoneyManagement")
End If
Set gMoneyManagementLogger = mMoneyManagementLogger
End Property

Public Property Get gOrderPlexProfileStructLogger() As Logger
If mOrderPlexProfileStructLogger Is Nothing Then
    Set mOrderPlexProfileStructLogger = GetLogger("tradebuild.OrderPlexProfileStruct")
    mOrderPlexProfileStructLogger.logToParent = False
End If
Set gOrderPlexProfileStructLogger = mOrderPlexProfileStructLogger
End Property

Public Property Get gOrderPlexProfileStructLoggerSimulated() As Logger
If mOrderPlexProfileStructLoggerSimulated Is Nothing Then
    Set mOrderPlexProfileStructLoggerSimulated = GetLogger("tradebuild.SimulatedOrderPlexProfileStructSimulated")
    mOrderPlexProfileStructLoggerSimulated.logToParent = False
End If
Set gOrderPlexProfileStructLoggerSimulated = mOrderPlexProfileStructLoggerSimulated
End Property

Public Property Get gOrderPlexProfileStringLogger() As Logger
If mOrderPlexProfileStringLogger Is Nothing Then
    Set mOrderPlexProfileStringLogger = GetLogger("tradebuild.OrderPlexProfileString")
    mOrderPlexProfileStringLogger.logToParent = False
End If
Set gOrderPlexProfileStringLogger = mOrderPlexProfileStringLogger
End Property

Public Property Get gOrderPlexProfileStringLoggerSimulated() As Logger
If mOrderPlexProfileStringLoggerSimulated Is Nothing Then
    Set mOrderPlexProfileStringLoggerSimulated = GetLogger("tradebuild.OrderPlexProfileStringSimulated")
    mOrderPlexProfileStringLoggerSimulated.logToParent = False
End If
Set gOrderPlexProfileStringLoggerSimulated = mOrderPlexProfileStringLoggerSimulated
End Property

Public Property Get gOrderLogger() As Logger
If mOrderLogger Is Nothing Then
    Set mOrderLogger = GetLogger("tradebuild.order")
End If
Set gOrderLogger = mOrderLogger
End Property

Public Property Get gOrderLoggerSimulated() As Logger
If mOrderLoggerSimulated Is Nothing Then
    Set mOrderLoggerSimulated = GetLogger("tradebuild.orderSimulated")
End If
Set gOrderLoggerSimulated = mOrderLoggerSimulated
End Property

Public Property Get gPositionLogger() As Logger
If mPositionLogger Is Nothing Then
    Set mPositionLogger = GetLogger("tradebuild.position")
    mPositionLogger.logToParent = False
End If
Set gPositionLogger = mPositionLogger
End Property

Public Property Get gPositionLoggerSimulated() As Logger
If mPositionLoggerSimulated Is Nothing Then
    Set mPositionLoggerSimulated = GetLogger("tradebuild.positionSimulated")
    mPositionLoggerSimulated.logToParent = False
End If
Set gPositionLoggerSimulated = mPositionLoggerSimulated
End Property

Public Property Get gTradeProfileLogger() As Logger
If mTradeProfileLogger Is Nothing Then
    Set mTradeProfileLogger = GetLogger("tradebuild.TradeProfile")
    mTradeProfileLogger.logToParent = False
End If
Set gTradeProfileLogger = mTradeProfileLogger
End Property

Public Property Get gTradeProfileLoggerSimulated() As Logger
If mTradeProfileLoggerSimulated Is Nothing Then
    Set mTradeProfileLoggerSimulated = GetLogger("tradebuild.TradeProfileSimulated")
    mTradeProfileLoggerSimulated.logToParent = False
End If
Set gTradeProfileLoggerSimulated = mTradeProfileLoggerSimulated
End Property

Public Property Get gProfitLogger() As Logger
If mProfitLogger Is Nothing Then
    Set mProfitLogger = GetLogger("tradebuild.profit")
    mProfitLogger.logToParent = False
End If
Set gProfitLogger = mProfitLogger
End Property

Public Property Get gProfitLoggerSimulated() As Logger
If mProfitLoggerSimulated Is Nothing Then
    Set mProfitLoggerSimulated = GetLogger("tradebuild.profitSimulated")
    mProfitLoggerSimulated.logToParent = False
End If
Set gProfitLoggerSimulated = mProfitLoggerSimulated
End Property

Public Property Get gDrawdownLogger() As Logger
If mDrawdownLogger Is Nothing Then
    Set mDrawdownLogger = GetLogger("tradebuild.drawdown")
    mDrawdownLogger.logToParent = False
End If
Set gDrawdownLogger = mDrawdownLogger
End Property

Public Property Get gDrawdownLoggerSimulated() As Logger
If mDrawdownLoggerSimulated Is Nothing Then
    Set mDrawdownLoggerSimulated = GetLogger("tradebuild.drawdownSimulated")
    mDrawdownLoggerSimulated.logToParent = False
End If
Set gDrawdownLoggerSimulated = mDrawdownLoggerSimulated
End Property

Public Property Get gMaxProfitLogger() As Logger
If mMaxProfitLogger Is Nothing Then
    Set mMaxProfitLogger = GetLogger("tradebuild.MaxProfit")
    mMaxProfitLogger.logToParent = False
End If
Set gMaxProfitLogger = mMaxProfitLogger
End Property

Public Property Get gMaxProfitLoggerSimulated() As Logger
If mMaxProfitLoggerSimulated Is Nothing Then
    Set mMaxProfitLoggerSimulated = GetLogger("tradebuild.MaxProfitSimulated")
    mMaxProfitLoggerSimulated.logToParent = False
End If
Set gMaxProfitLoggerSimulated = mMaxProfitLoggerSimulated
End Property

Public Property Get gOrderDetailLogger() As Logger
If mOrderDetailLogger Is Nothing Then
    Set mOrderDetailLogger = GetLogger("tradebuild.orderdetail")
    mOrderDetailLogger.logToParent = False
End If
Set gOrderDetailLogger = mOrderDetailLogger
End Property

Public Property Get gOrderDetailLoggerSimulated() As Logger
If mOrderDetailLoggerSimulated Is Nothing Then
    Set mOrderDetailLoggerSimulated = GetLogger("tradebuild.orderdetailSimulated")
    mOrderDetailLoggerSimulated.logToParent = False
End If
Set gOrderDetailLoggerSimulated = mOrderDetailLoggerSimulated
End Property

