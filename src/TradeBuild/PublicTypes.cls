VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type InfoEvent
    source          As Object
    valueType       As Long
    data            As Variant
End Type

Public Type NotificationEvent
    source          As Object
    eventCode       As Long
    eventMessage    As Variant
End Type

Public Type PriceChangeEvent
    source          As Object
    Change          As Double
    ChangeString    As String
    ChangePercent   As Double
End Type

Public Type ProfitEvent
    source          As Object
    profitType    As Long
    profitAmount    As Currency
    currencyCode    As String
End Type

Public Type MarketDepthEvent
    source          As Object
    side            As DOMSides
    price           As Double
    size            As Long
End Type

Public Type QuoteEvent
    source          As Object
    price           As Double
    priceString     As String
    size            As Long
    priceChange     As ValueChanges
    sizeChange      As ValueChanges
End Type

Public Type RawMarketDepthEvent
    source          As Object
    position        As Long
    marketMaker     As String
    operation       As DOMOperations
    side            As DOMSides
    price           As Double
    size            As Long
End Type

Public Type Tick
    timestamp       As Date
    TickType        As TradeBuild25.TickTypes
    TickPrice       As Double
    TickSize        As Long
    MDPosition      As Long
    MDMarketMaker   As String
    MDOperation     As Long
    MDSide          As Long
End Type

Public Type TickfileFormatSpecifier
    FormatType      As TradeBuild25.TickfileFormatType
    capabilities    As Long
    FormalID        As String
    name            As String
    FileExtension   As String
End Type

Public Type TickfileSpecifier
    tickfileFormatID    As String
    filename            As String
    Contract            As Contract
    From                As Date
    To                  As Date
    EntireSession       As Boolean
End Type

Public Type WriterEvent
    source          As Object
    action          As WriterNotifications
    filename        As String
End Type


'@================================================================================
' Member variables and constants
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
    ApiNotifyContractDetailsInvalid  ' The tickfile contains contract details from
                            ' an old version of Contract class
    ApiNotifyNoContractDetails ' The tickfile contains no contract details
                            ' (eg an eSignal tickfile) and no contract
                            ' has been specified in the contract property
    ApiNotifyTickfileDataSourceNotResponding
    ApiNotifyCantWriteToTickfileDataStore
    ApiNotifyRetryingConnectionToTickfileDataSource
    ApiNotifyConnectedToTickfileDataSource
    
    ApiNotifyNoHistoricDataSource
    ApiNotifyCantConnectHistoricDataSource
    ApiNotifyConnectedToHistoricDataSource
    ApiNotifyRetryingConnectionToHistoricDataSource
    
    ApiNotifyContractSpecifierAmbiguous
    ApiNotifyContractSpecifierInvalid
    
    ApiNotifyMarketDepthNotAvailable
    
    ApiNotifyInvalidRequest
    ApiNotifyFatalError           ' a fatal error has occurred within TradeBuild
    
    ApiNotifyCantConnectRealtimeDataSource
    ApiNotifyConnectedToRealtimeDataSource
    ApiNotifyLostConnectionToRealtimeDataSource
    ApiNotifyNoRealtimeDataSource
    ApiNotifyReconnectingToRealtimeDataSource
    ApiNotifyDisconnectedFromRealtimeDataSource
    ApiNotifyRealtimeDataRequestFailed
    
    ApiNotifyCantConnectToBroker
    ApiNotifyConnectedToBroker
    ApiNotifyRetryConnectToBroker
    ApiNotifyLostConnectionToBroker
    ApiNotifyReConnectingToBroker
    ApiNotifyInvalidOrder
    ApiNotifyInsufficientFunds
    ApiNotifyDisconnectedFromBroker
    
    ApiNotifyNonSpecificNotification
                            
End Enum

Public Enum BulletinTypes
    BulletinNews = 1
    BulletinExchangeUnavailable
    BulletinExchangeAvailable
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

Public Enum DOMOperations
    DOMInsert
    DOMUpdate
    DOMDelete
    DOMResetDisplay
End Enum

Public Enum DOMSides
    DOMAsk
    DOMBid
    DOMLast
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

Public Enum LogLevels
    LogLevelNone
    LogLevelLow
    LogLevelMedium
    LogLevelHigh
    LogLevelAll
End Enum

Public Enum MarketDepthTypes
    MarketDepthAny
    MarketDepthProcessed
    MarketDepthRaw
    MarketDepthRawL2
End Enum

Public Enum OptionRights
    OptNone = 0
    OptCall
    OptPut
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
    OrderAttRTHOnly = &H800&
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
    OrderStatusRejected
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
    OrderTypeMarket = 1&
    OrderTypeMarketOnClose = 2&
    OrderTypeLimit = 4&
    OrderTypeLimitOnClose = 8&
    OrderTypePeggedToMarket = &H10&
    OrderTypeStop = &H20&
    OrderTypeStopLimit = &H40&
    OrderTypeTrail = &H80&
    OrderTypeRelative = &H100&
    OrderTypeVWAP = &H200&
    OrderTypeMarketToLimit = &H400&
    OrderTypeQuote = &H800&
    OrderTypeAutoStop = &H1000&
    OrderTypeAutoLimit = &H2000&
    OrderTypeAdjust = &H4000&
    OrderTypeAlert = &H8000&
    OrderTypeLimitIfTouched = &H10000
    OrderTypeMarketIfTouched = &H20000
    OrderTypeTrailLimit = &H40000
    OrderTypeMarketWithProtection = &H80000
    OrderTypeMarketOnOpen = &H100000
    OrderTypeLimitOnOpen = &H200000
    OrderTypePeggedToPrimary = &H400000
End Enum

Public Enum PositionManagerChangeTypes
    PositionSizeChanged
End Enum

Public Enum ProfitTypes
    ProfitTypeProfit
    ProfitTypeMaxProfit
    ProfitTypeDrawdown
    ProfitTypeSessionProfit
    ProfitTypeSessionMaxProfit
    ProfitTypeSessionDrawdown
    ProfitTypetradeProfit
    ProfitTypeTradeMaxProfit
    ProfitTypetradeDrawdown
End Enum

Public Enum ReplayContinueModes
    ReplayContinue
    ReplayCancelThisFile
    ReplayCancelAllFiles
End Enum

Public Enum SecurityTypes
    SecTypeNone
    SecTypeStock
    SecTypeFuture
    SecTypeOption
    SecTypeFuturesOption
    SecTypeCash
    SecTypeBag
    SecTypeIndex
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

Public Enum TickTypes
    TickTypeBid = &H1&
    TickTypeAsk = &H2&
    TickTypeLast = &H4&
    TickTypeHigh = &H8&
    TickTypeLow = &H10&
    TickTypePrevClose = &H20&
    TickTypeVolume = &H40&
    TickTypeOpenInterest = &H80&
    TickTypeMarketDepth = &H100&
    TickTypeMarketDepthL2 = &H200&
    TickTypeMarketDepthReset = &H400&
End Enum

Public Enum TimeframeStates
    TimeframeStateCreated
    TimeframeStateLoading
    TimeframeStateLoaded
End Enum

Public Enum TradeBuildListenValueTypes

    VTAll = -1  ' used by listenenrs to specify that they want to receive all
                ' types of listen data
    
    VTLog = 1
    VTTrace
    VTDebug

    VTProfitProfile
    VTDummyProfitProfile
    VTMoneyManagement
    VTOrderPlexProfileStruct
    VTDummyOrderPlexProfileStruct
    VTOrderPlexProfileString
    VTDummyOrderPlexProfileString
    VTOrder
    VTDummyOrder
    VTPosition
    VTDummyPosition
    VTTradeProfile
    VTDummyTradeProfile
    VTProfit
    VTDummyProfit
    VTDrawdown
    VTDummyDrawdown
    VTMaxProfit
    VTDummyMaxProfit
    VTOrderDetail
    VTOrderDetailDummy
        
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
' Class Event Handlers
'@================================================================================

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


