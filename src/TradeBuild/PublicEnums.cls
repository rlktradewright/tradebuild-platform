VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

Public Type CurrentBarEvent
    Source          As Object
    currentBar      As TradeBuild.Bar
End Type

Public Type GenericEvent
    Source          As Object
End Type

Public Type GUID
    data(4) As Long
End Type

Public Type PriceChangeEvent
    Source          As Object
    Change          As Double
    ChangeString    As String
    ChangePercent   As Double
End Type

Public Type ProcessedMarketDepthEvent
    Source          As Object
    side            As DOMSides
    price           As Double
    size            As Long
End Type

Public Type QuoteEvent
    Source          As Object
    price           As Double
    priceString     As String
    size            As Long
End Type

Public Type RawMarketDepthEvent
    Source          As Object
    position        As Long
    marketMaker     As String
    operation       As DOMOperations
    side            As DOMSides
    price           As Double
    size            As Long
End Type

Public Type SortEntry
    key As String
    ObjectRef As Object
End Type

Public Type Tick
    timestamp       As Date
    TickType        As TradeBuild.TickTypes
    TickPrice       As Double
    TickSize        As Long
    MDPosition      As Long
    MDMarketMaker   As String
    MDOperation     As Long
    MDSide          As Long
End Type

Public Type TickfileFormatSpecifier
    FormatType      As TradeBuild.TickfileFormatType
    capabilities    As Long
    FormalID        As String
    Name            As String
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

Public Type TWSSocketMsgTickPrice
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    tickerid        As Long
    TickType        As TWSSocketTickTypes
    price           As Double
    size            As Long
End Type

Public Type TWSSocketMsgTickSize
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    tickerid        As Long
    TickType        As TWSSocketTickTypes
    size            As Long
End Type

Public Type TWSSocketMsgOrderStatus
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    orderId         As Long
    status          As String
    filled          As Long
    remaining       As Long
    avgFillPrice    As Double
    permId          As Long
    parentId        As Long
    lastFillPrice   As Double
    clientID        As Long
End Type

Public Type TWSSocketMsgErrMsg
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    id              As Long
    errorCode       As Long
    errorMsg        As String
End Type

Public Type TWSSocketMsgOpenOrder
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    orderId         As Long
    symbol          As String
    sectype         As String
    expiry          As String
    strike          As Double
    right           As String
    exchange        As String
    currencyCode    As String
    localSymbol     As String
    action          As String
    quantity        As Long
    orderType       As String
    limitPrice      As Double
    auxPrice        As Double
    timeInForce     As String
    ocaGroup        As String
    account         As String
    openClose       As String
    origin          As Long
    orderRef        As String
    clientID        As Long
    permId          As Long
    ignoreRTH       As Boolean
    hidden          As Boolean
    discretionaryAmt    As Double
    goodAfterTime       As String
    sharesAllocation    As String
    FAGroup             As String
    FAMethod            As String
    FAPercentage        As String
    FAProfile           As String
End Type

Public Type TWSSocketMsgAcctValue
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    key             As String
    value           As String
    currencyCode    As String
    accountName     As String
End Type

Public Type TWSSocketMsgPortfolioValue
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    symbol          As String
    sectype         As String
    expiry          As String
    strike          As Double
    right           As String
    currencyCode    As String
    localSymbol     As String
    position        As Long
    marketPrice     As Double
    marketValue     As Double
    averageCost     As Double
    unrealizedPNL   As Double
    realizedPNL     As Double
    accountName     As String
End Type

Public Type TWSSocketMsgAcctUpdateTime
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    accountTime     As String
End Type

Public Type TWSSocketMsgNextValidId
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    id              As Long
End Type

Public Type TWSSocketMsgContractData
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    symbol          As String
    sectype         As String
    expiry          As String
    strike          As Double
    right           As String
    exchange        As String
    currencyCode    As String
    localSymbol     As String
    marketName      As String
    tradingClass    As String
    contractID      As Long
    minimumTick     As Double
    multiplier      As Long
    OrderTypes      As String
    validExchanges  As String
End Type

Public Type TWSSocketMsgExecutionData
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    orderId         As Long
    symbol          As String
    sectype         As String
    expiry          As String
    strike          As Double
    right           As String
    exchange        As String
    currencyCode    As String
    localSymbol     As String
    execId          As String
    time            As String
    accountNumber   As String
    execExchange    As String
    side            As String
    quantity        As Long
    price           As Double
    permId          As Long
    clientID        As Long
End Type

Public Type TWSSocketMsgMarketDepth
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    tickerid        As Long
    position        As Long
    operation       As Long
    side            As Long
    price           As Double
    size            As Long
End Type

Public Type TWSSocketMsgMarketDepthL2
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    tickerid        As Long
    position        As Long
    marketMaker     As String
    operation       As Long
    side            As Long
    price           As Double
    size            As Long
End Type

Public Type TWSSocketMsgNewsBulletins
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    newsID          As Long
    newsType         As Long
    newsMessage     As String
    originatingExch As String
End Type

Public Type TWSSocketMsgManagedAccts
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    accountsList     As String
End Type

Public Type TWSSocketMsgReceiveFa
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    DataType        As Long
    XMLData         As String
End Type

Public Type WriterEvent
    Source          As Object
    action          As WriterNotifications
    filename        As String
End Type


'================================================================================
' Member variables and constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum ApiErrorCodes
    ' TWS-related codes
    MaxTickersReached = 1   ' Max number of tickers has been reached (101)
    AlreadyConnected        ' Already connected (501)
    ConnectFail             ' Couldn't connect to TWS.  Confirm that \"Enable ActiveX and Socket Clients\" is enabled on the \"Settings\" menu.(502)
    UpdateTWS               ' The TWS is out of date and must be upgraded.(503)
    NotConnected            ' Not connected(504)
    UnknownID               ' Fatal Error: Unknown message id(505)
    ZeroByteRead            ' Unexplained zero bytes read(506)
    NullStringRead          ' Null string read when expecting integer(507)
    NoBytesRead             ' Error: no bytes read or no null terminator found(508)
    SocketException         ' Exception caught while reading socket (509)
    FailCreateSocket        ' Failed to create socket(520)
    FailConnectTWS          ' Couldn't connect to TWS(521)
    
    ' TradeBuild-specific
    ServiceProviderErrorNotification
                            ' Error notified by service provider and not
                            ' handled by TradeBuild
    InvalidMessageCode = 900 ' Invalid message code received via socket
    TickfileEmpty           ' Tickfile is empty
    TickfileInvalid         ' Not a valid tickfile format
    TickfileVersionNotSupported
                            ' The version of tickfile is not supported
    TimedReplayNotAllowed   ' Can't do timed replay for this version
    DisconnectInvalid       ' Can't disconnect during tickfile replay
    ContractDetailsInvalid  ' The tickfile contains contract details from
                            ' an old version of Contract class
    UnrecognisedContract    ' Contract details in tickfile (old version)
                            ' not in database
    NoContractDetails       ' The tickfile contains no contract details
                            ' (eg an eSignal tickfile) and no contract
                            ' has been specified in the contract property
    TickfileNotCreated      ' couldn't create an output tickfile
    CantAccessCrescendoDatabase
                            ' the Crescendo database access software is not
                            ' installed or is not a compatible version
    TickerAlreadyRunningForThisContract
                            ' a ticker has already been started for the
                            ' contract specified in this ticker
    TickfileFormatNotDetermined
    TickfileDataSourceNotResponding
    RetryingConnectionToTickfileDataSource
    CantConnectHistoricDataSource
    RetryingConnectionToHistoricDataSource
    ContractSpecifierInvalid
    CantConnectRealtimeDataSource
                            
End Enum

Public Enum BarTypes
    BarTypeTrade = TradeBuildSP.BarTypes.trade
    BarTypeBid = TradeBuildSP.BarTypes.bid
    BarTypeAsk = TradeBuildSP.BarTypes.ask
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

Public Enum ErrorCodes
    ' generic run-time error codes defined by VB
    InvalidProcedureCall = 5
    Overflow = 6
    SubscriptOutOfRange = 9
    DivisionByZero = 11
    TypeMismatch = 13
    FileNotFound = 53
    FileAlreadyOpen = 55
    FileAlreadyExists = 58
    DiskFull = 61
    PermissionDenied = 70
    PathNotFound = 76
    InvalidObjectReference = 91
    
    InvalidPropertyValue = 380
    InvalidPropertyArrayIndex = 381
    
    ' non-generic error codes
    InvalidTickerID = vbObjectError + 512
    NotReceivingMarketDepth
    TickfileReplayProhibitsLiveOrders
    UnexpectedContract
    UnknownOrderID
    ContractDetailsReqNotAllowed
    InvalidOrderType
    InvalidOrderTypeInThisContext
    ContractCannotBeParsed
    NoInputTickFile
    CantCreateCrescendoTickfile
    ErrorOpeningTickfile
    CantAccessCrescendoDB
    NotImplemented
    AttemptToUseDeadTickerObject
    CantAddColumn
    CantGenerateColumns
    ColumnAlreadyAdded
    ColumnNameNotUnique
    TickerAlreadyInUse
    AlreadyConnected
    NoContractOrTickfile
    DatasourceHasNotBeenGenerated
    CantCreateGUID
    NotCorrectServiceProviderType
    NotUniqueServiceProviderName
    ServiceProviderNameInvalid
    InvalidServiceProviderHandle
    UnknownOrderTypeFromTWS

    ' generic error codes
    ArithmeticException = vbObjectError + 1024  ' an exceptional arithmetic condition has occurred
    ArrayIndexOutOfBoundsException  ' an array has been accessed with an illegal index
    ClassCastException              ' attempt to cast an object to class of which it is not an instance
    IllegalArgumentException        ' method has been passed an illegal or inappropriate argument
    IllegalStateException           ' a method has been invoked at an illegal or inappropriate time
    IndexOutOfBoundsException       ' an index of some sort (such as to an array, to a string, or to a vector) is out of range
    NullPointerException            ' attempt to use Nothing in a case where an object is required
    NumberFormatException           ' attempt to convert a string to one of the numeric types, but the string does not have the appropriate format
    RuntimeException                ' an unspecified runtime error has occurred
    SecurityException               ' a security violation has occurred
    UnsupportedOperationException   ' the requested operation is not supported



End Enum

Public Enum ExecSides
    SideBuy
    SideSell
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

Public Enum OrderOrigins
    OriginCustomer
    OriginFirm
    OriginUnknown
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
    TIFDay = 1
    TIFGoodTillCancelled
    TIFImmediateOrCancel
    TIFGoodTillDate
End Enum

Public Enum OrderTypes
    OrderTypeNone
    OrderTypeMarket
    OrderTypeMarketOnClose
    OrderTypeLimit
    OrderTypeLimitOnClose
    OrderTypePeggedToMarket
    OrderTypeStop
    OrderTypeStopLimit
    OrderTypeTrail
    OrderTypeRelative
    OrderTypeVWAP
    OrderTypeMarketToLimit
    OrderTypeQuote
    OrderTypeAutoStop
    OrderTypeAutoLimit
    OrderTypeAdjust
    OrderTypeAlert
    OrderTypeLimitIfTouched
    OrderTypeMarketIfTouched
    OrderTypeTrailLimit
    OrderTypeMarketWithProtection
    OrderTypeMarketOnOpen
    OrderTypeLimitOnOpen
    OrderTypePeggedToPrimary
End Enum

'Public Enum QuoteTypes
'    QuoteTypeAsk = 1
'    QuoteTypeBid
'    QuoteTypePrevClose
'    QuoteTypeHigh
'    QuoteTypeLow
'    QuoteTypeTrade
'    QuoteTypeVolume
'    QuoteTypeChange
'    QuoteTypeChangePercent
'End Enum

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

Public Enum StandardListenValueTypes
    DebugInfo = TradeBuildSP.StandardListenValueTypes.Debug
    LogInfo = TradeBuildSP.StandardListenValueTypes.Log
    TraceInfo = TradeBuildSP.StandardListenValueTypes.Trace
End Enum

Public Enum TickerColumnIds
    columnName = &H8&
    ColumnDescription = &H10&
    ColumnSymbol = &H20&
    ColumnSecType = &H40&
    ColumnExpiry = &H80&
    ColumnExchange = &H100&
    ColumnCurrency = &H200&
    ColumnRight = &H400&
    ColumnStrike = &H800&
    
    
    ColumnBid = &H1000&
    ColumnBidSize = &H2000&
    ColumnAsk = &H4000&
    ColumnAskSize = &H8000&
    ColumnTrade = &H10000
    ColumnTradeSize = &H20000
    ColumnVolume = &H40000
    ColumnHigh = &H80000
    ColumnLow = &H100000
    ColumnClose = &H200000
    ColumnChange = &H400000
    ColumnChangePercent = &H800000
    ColumnOpenInterest = &H1000000
End Enum

Public Enum TickerStateCodes
    Created
    Starting
    Running
    Paused
    Closing
    Dead
End Enum

Public Enum TickfileFormatType
    FileBased = TradeBuildSP.TickfileFormatTypes.FileBased
    StreamBased = TradeBuildSP.TickfileFormatTypes.StreamBased
End Enum

Public Enum TickTypes
    bid = &H1&
    ask = &H2&
    Last = &H4&
    high = &H8&
    Low = &H10&
    PrevClose = &H20&
    volume = &H40&
    openInterest = &H80&
    marketDepth = &H100&
    MarketDepthL2 = &H200&
    MarketDepthReset = &H400&
End Enum

Public Enum TimePeriodUnits
    Minute = 2
    Hour
    Day
    Week
    Month
    LunarMonth
    Year
End Enum

Public Enum TimestampFormats
    TimestampTimeOnly
    TimestampDateOnly
    TimestampDateAndTime
    TimestampCustom
End Enum

Public Enum TradeBuildListenValueTypes

    ValueTypeTickerQuote = 100
    ValueTypeTickerMarketDepth
    ValueTypeBarEvent

    
    
        
    ' Listen valueType identifiers for the IB TWS Service provider (these
    ' should be moved to the IB TWS Service Provider's type library when
    ' it is constructed)
        
    ValueTypeTWSSocketMsgTickPrice = 1000
    ValueTypeTWSSocketMsgTickSize
    ValueTypeTWSSocketMsgOrderStatus
    ValueTypeTWSSocketMsgErrMsg
    ValueTypeTWSSocketMsgOpenOrder
    ValueTypeTWSSocketMsgAcctValue
    ValueTypeTWSSocketMsgPortfolioValue
    ValueTypeTWSSocketMsgAcctUpdateTime
    ValueTypeTWSSocketMsgNextValidId
    ValueTypeTWSSocketMsgContractData
    ValueTypeTWSSocketMsgExecutionData
    ValueTypeTWSSocketMsgMarketDepth
    ValueTypeTWSSocketMsgMarketDepthL2
    ValueTypeTWSSocketMsgNewsBulletins
    ValueTypeTWSSocketMsgManagedAccts
    ValueTypeTWSSocketMsgReceiveFa
End Enum

Public Enum TriggerMethods
    TriggerDefault
    TriggerDoubleBidAsk
    TriggerLast
    TriggerDoubleLast
End Enum

Public Enum TWSSocketInMsgTypes
    TICK_PRICE = 1
    TICK_SIZE = 2
    ORDER_STATUS = 3
    ERR_MSG = 4
    OPEN_ORDER = 5
    ACCT_VALUE = 6
    PORTFOLIO_VALUE = 7
    ACCT_UPDATE_TIME = 8
    NEXT_VALID_ID = 9
    CONTRACT_DATA = 10
    EXECUTION_DATA = 11
    MARKET_DEPTH = 12
    MARKET_DEPTH_L2 = 13
    NEWS_BULLETINS = 14
    MANAGED_ACCTS = 15
    RECEIVE_FA = 16
End Enum

Public Enum TWSSocketOutMsgTypes
    REQ_MKT_DATA = 1
    CANCEL_MKT_DATA = 2
    PLACE_ORDER = 3
    CANCEL_ORDER = 4
    REQ_OPEN_ORDERS = 5
    REQ_ACCT_DATA = 6
    REQ_EXECUTIONS = 7
    REQ_IDS = 8
    REQ_CONTRACT_DATA = 9
    REQ_MKT_DEPTH = 10
    CANCEL_MKT_DEPTH = 11
    REQ_NEWS_BULLETINS = 12
    CANCEL_NEWS_BULLETINS = 13
    SET_SERVER_LOGLEVEL = 14
    REQ_AUTO_OPEN_ORDERS = 15
    REQ_ALL_OPEN_ORDERS = 16
    REQ_MANAGED_ACCTS = 17
    REQ_FA = 18
    REPLACE_FA = 19
End Enum

Public Enum TWSSocketTickTypes
    TICK_BID_SIZE   ' 0
    TICK_BID        ' 1
    TICK_ASK        ' 2
    TICK_ASK_SIZE   ' 3
    TICK_LAST       ' 4
    TICK_LAST_SIZE  ' 5
    TICK_HIGH       ' 6
    TICK_LOW        ' 7
    TICK_VOLUME     ' 8
    TICK_CLOSE      ' 9
End Enum

Public Enum WriterNotifications
    WriterReady
    WriterNotReady
    WriterFileCreated
    WriteWriteOperationCompleted
End Enum

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


