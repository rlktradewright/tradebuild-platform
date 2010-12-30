VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuild"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'' A GlobalMultiUse class that provides a variety of useful functions.

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum ApiNotifyCodes
    
    ' TradeBuild-specific
    ApiNotifyServiceProviderError = 900
                            ' Error notified by service provider and not
                            ' handled by TradeBuild
    ApiNotifyTickfileEmpty     ' Tickfile is empty
    ApiNotifyTickfileInvalid   ' Not a valid tickfile format
    ApiNotifyTickfileVersionNotSupported
                            ' The version of tickfile is not supported
    ApiNotifyTickfileContractDetailsInvalid  ' The tickfile contains contract details from
                            ' an old version of Contract class
    ApiNotifyTickfileNoContractDetails ' The tickfile contains no contract details
                            ' (eg an eSignal tickfile) and no contract
                            ' has been specified in the contract property
    ApiNotifyTickfileDataSourceNotResponding
    ApiNotifyTickfileDoesntExist
    ApiNOtifyTickfileFormatNotSupported
    ApiNotifyTickfileContractSpecifierInvalid
    
    ApiNotifyCantWriteToTickfileDataStore
    ApiNotifyRetryingConnectionToTickfileDataSource
    ApiNotifyConnectedToTickfileDataSource
    ApiNotifyReconnectingToTickfileDataSource
    ApiNotifyLostConnectionToTickfileDataSource
    
    ApiNotifyNoHistoricDataSource
    ApiNotifyCantConnectHistoricDataSource
    ApiNotifyConnectedToHistoricDataSource
    ApiNotifyDisconnectedFromHistoricDataSource
    ApiNotifyRetryingConnectionToHistoricDataSource
    ApiNotifyLostConnectionToHistoricDataSource
    ApiNotifyReconnectingToHistoricDataSource
    ApiNotifyHistoricDataRequestFailed
    
'    ApiNotifyMarketDepthNotAvailable
    
    ApiNotifyInvalidRequest
    
    ApiNotifyCantConnectRealtimeDataSource
    ApiNotifyConnectedToRealtimeDataSource
    ApiNotifyLostConnectionToRealtimeDataSource
    ApiNotifyNoRealtimeDataSource
    ApiNotifyReconnectingToRealtimeDataSource
    ApiNotifyDisconnectedFromRealtimeDataSource
    ApiNotifyRealtimeDataRequestFailed
    ApiNotifyRealtimeDataSourceNotResponding
    
    ApiNotifyCantConnectToBroker
    ApiNotifyConnectedToBroker
    ApiNotifyRetryConnectToBroker
    ApiNotifyLostConnectionToBroker
    ApiNotifyReConnectingToBroker
    ApiNotifyDisconnectedFromBroker
    
    ApiNotifyNonSpecificNotification
    
    ApiNotifyCantWriteToHistoricDataStore
    
    ApiNotifyTryLater
    
    ApiNotifyCantConnectContractDataSource
    ApiNotifyConnectedToContractDataSource
    ApiNotifyDisconnectedFromContractDataSource
    ApiNotifyLostConnectionToContractDataSource
    ApiNotifyReConnectingContractDataSource
    ApiNotifyRetryConnectContractDataSource
    
    ApiNotifyOrderRejected
    ApiNotifyOrderDeferred
    
    
End Enum

Public Enum BulletinTypes
    BulletinNews = 1
    BulletinExchangeUnavailable
    BulletinExchangeAvailable
End Enum

Public Enum PositionTypes
    PositionTypeLive = 1
    PositionTypeSimulated = 2
    PositionTypeAll = 3
End Enum

Public Enum ClosePositionFlags
    ClosePositionCancelOrders = 1
    ClosePositionWaitForCancel = 2
End Enum

Public Enum ConnectionStates
    ConnNotConnected
    ConnConnecting
    ConnConnected
    ConnReplaying
End Enum

Public Enum DOMEvents
    DOMNoEvents
    DOMRawEvents
    DOMProcessedEvents
    DOMBothEvents
End Enum

Public Enum DOMUpdateTypes
    DOMUpdateAsk
    DOMUpdateBid
    DOMUpdateLast
End Enum

Public Enum EntryOrderTypes
    EntryOrderTypeMarket = 1000
    EntryOrderTypeMarketOnOpen
    EntryOrderTypeMarketOnClose
    EntryOrderTypeMarketIfTouched
    EntryOrderTypeMarketToLimit
    EntryOrderTypeBid
    EntryOrderTypeAsk
    EntryOrderTypeLast
    EntryOrderTypeLimit
    EntryOrderTypeLimitOnOpen
    EntryOrderTypeLimitOnClose
    EntryOrderTypeLimitIfTouched
    EntryOrderTypeStop
    EntryOrderTypeStopLimit
End Enum

Public Enum FADataTypes
    FAGroups = 1
    FAProfile
    FAAccountAliases
End Enum

Public Enum LegOpenClose
    LegSamePos
    LegOpenPos
    LegClosePos
    LegUnknownPos
End Enum

Public Enum MarketDepthTypes
    MarketDepthAny
    MarketDepthProcessed
    MarketDepthRaw
    MarketDepthRawL2
End Enum

Public Enum OrderActions
    ActionNone
    ActionBuy = 1
    ActionSell
End Enum

Public Enum OrderAttributeIds
    OrderAttOpenClose = &H1&
    OrderAttOrigin = &H2&
    OrderAttOriginatorRef = &H4&
    OrderAttBlockOrder = &H8&
    OrderAttSweepToFill = &H10&
    OrderAttDisplaySize = &H20&
    OrderAttIgnoreRTH = &H40&
    OrderAttHidden = &H80&
    OrderAttDiscretionaryAmount = &H100&
    OrderAttGoodAfterTime = &H200&
    OrderAttGoodTillDate = &H400&
    'OrderAttRTHOnly = &H800&
    OrderAttRule80A = &H1000&
    OrderAttSettlingFirm = &H2000&
    OrderAttAllOrNone = &H4000&
    OrderAttMinimumQuantity = &H8000&
    OrderAttPercentOffset = &H10000
    OrderAttETradeOnly = &H20000
    OrderAttFirmQuoteOnly = &H40000
    OrderAttNBBOPriceCap = &H80000
    OrderAttOverrideConstraints = &H100000
    OrderAttAction = &H200000
    OrderAttLimitPrice = &H400000
    OrderAttOrderType = &H800000
    OrderAttQuantity = &H1000000
    OrderAttTimeInForce = &H2000000
    OrderAttTriggerPrice = &H4000000
    OrderAttGoodAfterTimeTZ = &H8000000
    OrderAttGoodTillDateTZ = &H10000000
    OrderAttStopTriggerMethod = &H20000000
End Enum

Public Enum OrderOrigins
    OriginCustomer
    OriginFirm
    OriginUnknown
End Enum

Public Enum OrderPlexChangeTypes
    OrderPlexCreated = 1
    OrderPlexCompleted
    OrderPlexSelfCancelled
    OrderPlexEntryOrderChanged
    OrderPlexStopOrderChanged
    OrderPlexTargetOrderChanged
    OrderPlexCloseoutOrderCreated
    OrderPlexCloseoutOrderChanged
    OrderPlexProfitThresholdExceeded
    OrderPlexLossThresholdExceeded
    OrderPlexDrawdownThresholdExceeded
    OrderPlexSizeChanged
    OrderPlexStateChanged
    OrderPlexChangesPending
    OrderPlexChangesApplied
    OrderPlexChangesCancelled
End Enum

Public Enum OrderStatuses
    OrderStatusCreated
    OrderStatusRejected         ' by the broker
    OrderStatusPendingSubmit
    OrderStatusPreSubmitted
    OrderStatusSubmitted
    OrderStatusFilled
    OrderStatusCancelling
    OrderStatusCancelled
End Enum

Public Enum OrderTifs
    TIFNone = 0
    TIFDay = 1
    TIFGoodTillCancelled = 2
    TIFImmediateOrCancel = 4
End Enum

Public Enum OrderTypes
    OrderTypeNone = 0
    OrderTypeMarket = TradeBuildSP.OrderTypes.OrderTypeMarket
    OrderTypeMarketOnClose = TradeBuildSP.OrderTypes.OrderTypeMarketOnClose
    OrderTypeLimit = TradeBuildSP.OrderTypes.OrderTypeLimit
    OrderTypeLimitOnClose = TradeBuildSP.OrderTypes.OrderTypeLimitOnClose
    OrderTypePeggedToMarket = TradeBuildSP.OrderTypes.OrderTypePeggedToMarket
    OrderTypeStop = TradeBuildSP.OrderTypes.OrderTypeStop
    OrderTypeStopLimit = TradeBuildSP.OrderTypes.OrderTypeStopLimit
    OrderTypeTrail = TradeBuildSP.OrderTypes.OrderTypeTrail
    OrderTypeRelative = TradeBuildSP.OrderTypes.OrderTypeRelative
    OrderTypeVWAP = TradeBuildSP.OrderTypes.OrderTypeVWAP
    OrderTypeMarketToLimit = TradeBuildSP.OrderTypes.OrderTypeMarketToLimit
    OrderTypeQuote = TradeBuildSP.OrderTypes.OrderTypeQuote
    OrderTypeAutoStop = TradeBuildSP.OrderTypes.OrderTypeAutoStop
    OrderTypeAutoLimit = TradeBuildSP.OrderTypes.OrderTypeAutoLimit
    OrderTypeAdjust = TradeBuildSP.OrderTypes.OrderTypeAdjust
    OrderTypeAlert = TradeBuildSP.OrderTypes.OrderTypeAlert
    OrderTypeLimitIfTouched = TradeBuildSP.OrderTypes.OrderTypeLimitIfTouched
    OrderTypeMarketIfTouched = TradeBuildSP.OrderTypes.OrderTypeMarketIfTouched
    OrderTypeTrailLimit = TradeBuildSP.OrderTypes.OrderTypeTrailLimit
    OrderTypeMarketWithProtection = TradeBuildSP.OrderTypes.OrderTypeMarketWithProtection
    OrderTypeMarketOnOpen = TradeBuildSP.OrderTypes.OrderTypeMarketOnOpen
    OrderTypeLimitOnOpen = TradeBuildSP.OrderTypes.OrderTypeLimitOnOpen
    OrderTypePeggedToPrimary = TradeBuildSP.OrderTypes.OrderTypePeggedToPrimary
End Enum

Public Enum PositionManagerChangeTypes
    PositionSizeChanged
    ProviderReadinessChanged
    PositionClosed
End Enum

Public Enum ProfitTypes
    ProfitTypeProfit
    ProfitTypeMaxProfit
    ProfitTypeMaxLoss
    ProfitTypeDrawdown
    ProfitTypeSessionProfit
    ProfitTypeSessionMaxProfit
    ProfitTypeSessionMaxLoss
    ProfitTypeSessionDrawdown
    ProfitTypeTradeProfit
    ProfitTypeTradeMaxProfit
    ProfitTypeTradeMaxLoss
    ProfitTypeTradeDrawdown
End Enum

Public Enum ReplayContinueModes
    ReplayContinue
    ReplayCancelThisFile
    ReplayCancelAllFiles
End Enum

Public Enum ServiceProviderRoles
    SPRealtimeData = 1
    SPPrimaryContractData = 2
    SPSecondaryContractData = 4
    SPHistoricalDataInput = 8
    SPHistoricalDataOutput = &H10
    SPBrokerSimulated = &H20
    SPTickfileInput = &H40
    SPTickfileOutput = &H80
    SPBrokerLive = &H100
End Enum

Public Enum StopOrderTypes
    StopOrderTypeNone = 2000
    StopOrderTypeStop
    StopOrderTypeStopLimit
    StopOrderTypeBid
    StopOrderTypeAsk
    StopOrderTypeLast
    StopOrderTypeAuto
End Enum

Public Enum StopTriggerMethods
    StopTriggerNone = 0
    StopTriggerDefault = 1
    StopTriggerDoubleBidAsk = 2
    StopTriggerLast = 4
    StopTriggerDoubleLast = 8
    StopTriggerBidAsk = 16
    StopTriggerLastOrBidAsk = 32
    StopTriggerMidPoint = 64
End Enum

Public Enum TargetOrderTypes
    TargetOrderTypeNone = 3000
    TargetOrderTypeLimit
    TargetOrderTypeLimitIfTouched
    TargetOrderTypeMarketIfTouched
    TargetOrderTypeBid
    TargetOrderTypeAsk
    TargetOrderTypeLast
    TargetOrderTypeAuto
End Enum

''
' Members of this enum govern the operation of a <code>Ticker</code> object.
'
' @param TickerOptOrdersAreLive
'   Orders placed via this Ticker are to be passed to the live order submission
'   service provider (if one has been configured).
' @param TickerOptUseExchangeTimeZone
'   Timestamps for ticks received via this Ticker are to be expressed in the
'   timezone of the relevant exchange, rather than the computer&39;s local
'   timezone. Note that this only has effect when the contract service provider
'   supplies exchange timezone information (the IB contract service provider does not).
' @param TickerOptWriteTickData
'   Indicates that tick Data for this Ticker (excluding market depth Data) is to
'   be written to a tickfile. Note that the nature of this tickfile is governed by
'   the currently configured Tickfile Service Provider.
' @param TickerOptWriteTradeBarData
'   Indicates that bar Data for trades notified via this Ticker is to
'   be written to a historical Data store. Note that the nature of this Data store, and
'   the timeframes of the bars recorded, is governed by the currently configured
'   Historical Data Service Provider.
' @param TickerOptWriteBidAndAskBarData
'   Indicates that bar Data for bids and asks notified via this Ticker is to
'   be written to a historical Data store. Note that the nature of this Data store, and
'   the timeframes of the bars recorded, is governed by the currently configured
'   Historical Data Service Provider.
' @param TickerOptUpdateLatestBar
'   Indicates that where the latest bar written to the historical Data store is still
'   open (ie the end of the bar has not yet been reached), then the Data from this Ticker
'   is to be used to update that bar. If not set, the previously recorded Data for that bar
'   is overwritten. This value is ignored if <code>TickerOptWriteTradeBarData</code> and
'   <code>TickerOptWriteBidAndAskBarData</code> are not set.
' @param TickerOptIncludeMarketDepthInTickfile
'   Indicates that market depth Data is to be included in the tickfile. This value is
'   ignored if <code>TickerOptWriteTickData</code> is not set.
' @param TickerOptNoImpliedTrades
'   If set, the Ticker does not attempt to deduce the existence of unreported trades
'   which may be indicated by the Ticker Data received from the Data source. For example,
'   if Volume is incremented by an amount equal to the last reported Trade, but no
'   Trade is reported, then an implied Trade exists - setting this option prevents
'   that implied Trade being notified.
' @param TickerOptNoVolumeAdjustments
'   If set, the Ticker reports all Volume ticks exactly as received from the Data source.
'   Otherwise, the Ticker may adjust Volume figures to ensure that Volume notified by the
'   Ticker is always increasing (this adjustment is to eliminate the backward movement in
'   total Volume that occurs sometimes with some Data sources).
'@/
Public Enum TickerOptions
    TickerOptOrdersAreLive = 1
    TickerOptUseExchangeTimeZone = 2
    TickerOptWriteTickData = 4
    TickerOptWriteTradeBarData = 8
    TickerOptWriteBidAndAskBarData = &H10
    TickerOptUpdateLatestBar = &H20
    TickerOptIncludeMarketDepthInTickfile = &H40
    TickerOptNoImpliedTrades = &H80
    TickerOptNoVolumeAdjustments = &H100
End Enum

Public Enum TickerStates
    TickerStateCreated
    TickerStateStarting
    TickerStateRunning
    TickerStatePaused
    TickerStateClosing
    TickerStateStopped
    TickerStateReady
End Enum

Public Enum TickfileFormatType
    FileBased = TradeBuildSP.TickfileFormatTypes.FileBased
    StreamBased = TradeBuildSP.TickfileFormatTypes.StreamBased
End Enum

Public Enum TimeframeStates
    TimeframeStateCreated
    TimeframeStateLoading
    TimeframeStateLoaded
    TimeframeStateFinished
End Enum

Public Enum ValueChanges
    ValueChangeNone
    ValueChangeUp
    ValueChangeDown
End Enum

Public Enum WriterNotifications
    WriterReady
    WriterNotReady
    WriterFileCreated
    WriterWriteOperationCompleted
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type ClosePositionsResult
    numberOfPositionsClosed     As Long
    numberOfPositionsNotClosed  As Long
End Type

Public Type InfoEventData
    source          As Object
    valueType       As Long
    Data            As Variant
End Type

Public Type PriceChangeEventData
    source          As Object
    Change          As Double
    ChangeString    As String
    ChangePercent   As Double
End Type

Public Type ProfitEventData
    source          As Object
    profitType    As Long
    ProfitAmount    As Currency
    currencyCode    As String
End Type

Public Type MarketDepthEventData
    source          As Object
    type            As DOMUpdateTypes
    price           As Double
    size            As Long
End Type

Public Type QuoteEventData
    source          As Object
    price           As Double
    size            As Long
    PriceChange     As ValueChanges
    sizeChange      As ValueChanges
    recentPriceChange     As ValueChanges   ' if PriceChange = ValueChangeNone
    recentSizeChange      As ValueChanges   ' if sizeChange = ValueChangeNone
End Type

Public Type RawMarketDepthEventData
    source          As Object
    position        As Long
    marketMaker     As String
    operation       As DOMOperations
    side            As DOMSides
    price           As Double
    size            As Long
End Type

Public Type TickfileFormatSpecifier
    FormatType      As TickfileFormatType
    capabilities    As Long
    FormalID        As String
    name            As String
    FileExtension   As String
End Type

Public Type TickWriterComment
    Timestamp       As Date
    Comment         As String
End Type

Public Type WriterEventData
    source          As Object
    action          As WriterNotifications
    Filename        As String
End Type

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Randomize
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get TradeBuildAPI() As TradeBuildAPI
If gTB Is Nothing Then Set gTB = New TradeBuildAPI
Set TradeBuildAPI = gTB
End Property

'@================================================================================
' Methods
'@================================================================================

''
'
' Converts a member of the EntryOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
'
' @param pEntryOrderType the EntryOrderTypes value to be converted
'
'@/
Public Function EntryOrderTypeToOrderType( _
                ByVal pEntryOrderType As EntryOrderTypes) As OrderTypes
EntryOrderTypeToOrderType = GEntryOrderTypeToOrderType(pEntryOrderType)
End Function

Public Function EntryOrderTypeToString( _
                ByVal value As EntryOrderTypes) As String
EntryOrderTypeToString = GEntryOrderTypeToString(value)
End Function

Public Function EntryOrderTypeToShortString( _
                ByVal value As EntryOrderTypes) As String
EntryOrderTypeToShortString = GEntryOrderTypeToShortString(value)
End Function

Public Function GetFormattedPriceFromQuoteEvent( _
                ByRef ev As QuoteEventData) As String
Dim lTicker As Ticker
Set lTicker = ev.source
GetFormattedPriceFromQuoteEvent = lTicker.FormatPrice(ev.price)
End Function

Public Function LegOpenCloseFromString( _
                ByVal value As String) As LegOpenClose
LegOpenCloseFromString = GLegOpenCloseFromString(value)
End Function

Public Function LegOpenCloseToString( _
                ByVal value As LegOpenClose) As String
LegOpenCloseToString = GLegOpenCloseToString(value)
End Function

Public Function OrderActionFromString( _
                ByVal value As String) As OrderActions
OrderActionFromString = GOrderActionFromString(value)
End Function

Public Function OrderActionToString( _
                ByVal value As OrderActions) As String
OrderActionToString = GOrderActionToString(value)
End Function

Public Function OrderAttributeToString(ByVal value As OrderAttributes) As String
OrderAttributeToString = GOrderAttributeToString(value)
End Function
Public Function OrderStatusToString( _
                ByVal value As OrderStatuses) As String
OrderStatusToString = GOrderStatusToString(value)
End Function

Public Function OrderStopTriggerMethodToString( _
                ByVal value As StopTriggerMethods) As String
OrderStopTriggerMethodToString = GOrderStopTriggerMethodToString(value)
End Function

Public Function OrderTIFToString( _
                ByVal value As OrderTifs) As String
OrderTIFToString = GOrderTIFToString(value)
End Function

Public Function OrderTypeToShortString( _
                ByVal value As OrderTypes) As String
OrderTypeToShortString = gOrderTypeToShortString(value)
End Function

Public Function OrderTypeToString( _
                ByVal value As OrderTypes) As String
OrderTypeToString = GOrderTypeToString(value)
End Function

''
' Converts a member of the StopOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pStopOrderType the StopOrderTypes value to be converted
' @see
'
'@/
Public Function StopOrderTypeToOrderType( _
                ByVal pStopOrderType As StopOrderTypes) As OrderTypes
StopOrderTypeToOrderType = GStopOrderTypeToOrderType(pStopOrderType)
End Function

Public Function StopOrderTypeToString( _
                ByVal value As StopOrderTypes)
StopOrderTypeToString = GStopOrderTypeToString(value)
End Function

''
' Converts a member of the <code>TargetOrderTypes</code> enumeration to the equivalent <code>OrderTypes</code> value.
'
' @return           the <code>OrderTypes</code> value corresponding to the parameter
' @param pTargetOrderType the <code>TargetOrderTypes</code> value to be converted
' @see
'
'@/
Public Function TargetOrderTypeToOrderType( _
                ByVal pTargetOrderType As TargetOrderTypes) As OrderTypes
TargetOrderTypeToOrderType = GTargetOrderTypeToOrderType(pTargetOrderType)
End Function

Public Function TargetOrderTypeToString( _
                ByVal value As TargetOrderTypes) As String
TargetOrderTypeToString = GTargetOrderTypeToString(value)
End Function

Public Function TickfileSpecifierToString( _
                ByVal TickfileSpec As TickfileSpecifier) As String
TickfileSpecifierToString = GTickfileSpecifierToString(TickfileSpec)
End Function

'@================================================================================
' Helper Functions
'@================================================================================



