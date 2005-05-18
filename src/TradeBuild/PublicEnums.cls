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

Public Type GUID
    data(4) As Long
End Type

Public Type SortEntry
    key As String
    ObjectRef As Object
End Type

Public Type TickerListenData
    TickerID As Long
    price As Double
    Size As Long
End Type

Public Type TickerListenMDData
    TickerID As Long
    side As DOMSides
    price As Double
    Size As Long
    marketDataReRequested As Boolean
End Type

Public Type TickerListenRawMDData
    TickerID As Long
    position As Long
    marketMaker As String
    operation As DOMOperations
    side As DOMSides
    price As Double
    Size As Long
End Type

Public Type TickfileFormatSpecifier
    FormatType      As TradeBuild.TickfileFormatType
    capabilities    As Long
    FormalID        As String
    Name            As String
    FileExtension   As String
End Type

Public Type TickfileSpecifier
    TickfileFormatID    As String
    Filename            As String
    contractSpecifier   As contractSpecifier
    From                As Date
    To                  As Date
    EntireSession       As Boolean
End Type

Public Type TWSSocketMsgTickPrice
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    TickerID        As Long
    TickType        As TWSSocketTickTypes
    price           As Double
    Size            As Long
End Type

Public Type TWSSocketMsgTickSize
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    TickerID        As Long
    TickType        As TWSSocketTickTypes
    Size            As Long
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
    secType         As String
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
    secType         As String
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
    secType         As String
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
    secType         As String
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
    TickerID        As Long
    position        As Long
    operation       As Long
    side            As Long
    price           As Double
    Size            As Long
End Type

Public Type TWSSocketMsgMarketDepthL2
    msgType         As TWSSocketInMsgTypes
    Version         As Long
    TickerID        As Long
    position        As Long
    marketMaker     As String
    operation       As Long
    side            As Long
    price           As Double
    Size            As Long
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
    TickfleVersionNotSupported
                            ' The version of tickfile is not supported
    TickfileVersionNotSupported = TickfleVersionNotSupported    ' correct spelling!
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
End Enum

Public Enum DOMSides
    DOMAsk
    DOMBid
    DOMLast
End Enum

Public Enum ErrorCodes
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

Public Enum OptionRights
    OptNone = 0
    OptCall
    OptPut
End Enum
    
Public Enum OrderActions
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
End Enum

Public Enum OrderTypes
    MinOrderType = 0
    OrderTypeNone
    OrderTypeMarket
    OrderTypeMarketClose
    OrderTypeLimit
    OrderTypeLimitClose
    OrderTypePegMarket
    OrderTypeStop
    OrderTypeStopLimit
    OrderTypeTrail
    OrderTypeRelative
    OrderTypeVWAP
    OrderTypeMarketToLimit
    OrderTypeQuote
    MaxOrderType = OrderTypeQuote
End Enum

Public Enum ReplayContinueModes
    ReplayContinue
    ReplayCancelThisFile
    ReplayCancelAllFiles
End Enum

Public Enum SecurityTypes
    SecTypeStock = 1
    SecTypeFuture
    SecTypeOption
    SecTypeFuturesOption
    SecTypeCash
    SecTypeBag
    SecTypeIndex
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
End Enum

Public Enum TickerStateCodes
    Created
    Starting
    Running
    Paused
    Dead
End Enum

Public Enum TickfileFormatType
    FileBased = TradeBuildSP.TickfileFormatType.FileBased
    StreamBased = TradeBuildSP.TickfileFormatType.StreamBased
End Enum

Public Enum TickTypes
    bid = &H1&
    ask = &H2&
    Last = &H4&
    High = &H8&
    Low = &H10&
    PrevClose = &H20&
    Volume = &H40&
    OpenInterest = &H80&
    marketDepth = &H100&
    MarketDepthL2 = &H200&
    MarketDepthReset = &H400&
End Enum

Public Enum TimestampFormats
    TimestampTimeOnly
    TimestampDateOnly
    TimestampDateAndTime
    TimestampCustom
End Enum

Public Enum TradeBuildListenValueTypes

    ValueTypeTradeBuildAsk = 100
    ValueTypeTradeBuildAskSize
    ValueTypeTradeBuildBid
    ValueTypeTradeBuildBidSize
    ValueTypeTradeBuildClose
    ValueTypeTradeBuildHigh
    ValueTypeTradeBuildLow
    ValueTypeTradeBuildTrade
    ValueTypeTradeBuildVolume
    ValueTypeTradeBuildMarketdepth
    ValueTypeTradeBuildRawMarketdepth
    ValueTypeTradeBuildRawMarketdepthL2
    ValueTypeTradeBuildChange
    ValueTypeTradeBuildChangePercent

    
    
        
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


