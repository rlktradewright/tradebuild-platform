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

'================================================================================
' Enums
'================================================================================

Public Enum ConnectionStates
    ConnNotConnected
    ConnConnecting
    ConnConnected
End Enum

Public Enum FADataTypes
    FAGroups = 1
    FAProfile
    FAAccountAliases
End Enum

Public Enum ListenValueTypes

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
    ValueTypeRawSocketData
End Enum

Public Enum TickTypes
    None
    Bid
    Ask
    Last
    volume
    prevClose
    High
    Low
    openInterest
End Enum

Public Enum TWSBarSizes
    BarSize1Second = 1
    BarSize5Second = 2
    BarSize15Second = 3
    BarSize30Second = 4
    BarSize1Minute = 5
    BarSize2Minutes = 6
    BarSize5Minutes = 7
    BarSize15Minutes = 8
    BarSize30Minutes = 9
    BarSize1Hour = 10
    BarSize1Day = 11
End Enum

Public Enum TWSLogLevels
    TWSLogLevelSystem
    TWSLogLevelError
    TWSLogLevelWarning
    TWSLogLevelInformation
    TWSLogLevelDetail
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
    HISTORICAL_DATA = 17
    BOND_CONTRACT_DATA = 18
    SCANNER_PARAMETERS = 19
    SCANNER_DATA = 20
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
    REQ_HISTORICAL_DATA = 20
    EXERCISE_OPTIONS = 21
    REQ_SCANNER_SUBSCRIPTION = 22
    CANCEL_SCANNER_SUBSCRIPTION = 23
    REQ_SCANNER_PARAMETERS = 24
    CANCEL_HISTORICAL_DATA = 25
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
' Types
'================================================================================

Public Type HistoricBar
    timestamp       As Date
    openValue       As Double
    highValue       As Double
    lowValue        As Double
    closeValue      As Double
    volume          As Long
    WAP             As Double
    hasGaps         As Boolean
End Type

Public Type OrderIdMapEntry
    order                   As IOrder
    orderSubmitter          As orderSubmitter
End Type

Public Type TWSContractDetailsRequestQueueEntry
    tradeBuildContracts         As TradeBuildSP.IContracts
    localContracts              As Collection
    requester                   As ContractDetailsRequester
End Type

Public Type TWSHistoricalDataRequestQueueEntry
    id                      As Long
    contractSpec            As TradeBuildSP.IContractSpecifier
    endDateTime             As String
    barSizeSetting          As TWSBarSizes
    duration                As String
    whatToShow              As String
End Type

Public Type TWSSocketMsgTickPrice
    msgType         As TWSSocketInMsgTypes
    version         As Long
    tickerid        As Long
    tickType        As TWSSocketTickTypes
    price           As Double
    size            As Long
End Type

Public Type TWSSocketMsgTickSize
    msgType         As TWSSocketInMsgTypes
    version         As Long
    tickerid        As Long
    tickType        As TWSSocketTickTypes
    size            As Long
End Type

Public Type TWSSocketMsgOrderStatus
    msgType         As TWSSocketInMsgTypes
    version         As Long
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
    version         As Long
    id              As Long
    errorCode       As Long
    errorMsg        As String
End Type

Public Type TWSSocketMsgOpenOrder
    msgType         As TWSSocketInMsgTypes
    version         As Long
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
    version         As Long
    key             As String
    value           As String
    currencyCode    As String
    accountName     As String
End Type

Public Type TWSSocketMsgPortfolioValue
    msgType         As TWSSocketInMsgTypes
    version         As Long
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
    version         As Long
    accountTime     As String
End Type

Public Type TWSSocketMsgNextValidId
    msgType         As TWSSocketInMsgTypes
    version         As Long
    id              As Long
End Type

Public Type TWSSocketMsgContractData
    msgType         As TWSSocketInMsgTypes
    version         As Long
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
    version         As Long
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
    version         As Long
    tickerid        As Long
    position        As Long
    operation       As Long
    side            As Long
    price           As Double
    size            As Long
End Type

Public Type TWSSocketMsgMarketDepthL2
    msgType         As TWSSocketInMsgTypes
    version         As Long
    tickerid        As Long
    position        As Long
    marketmaker     As String
    operation       As Long
    side            As Long
    price           As Double
    size            As Long
End Type

Public Type TWSSocketMsgNewsBulletins
    msgType         As TWSSocketInMsgTypes
    version         As Long
    newsID          As Long
    newsType         As Long
    newsMessage     As String
    originatingExch As String
End Type

Public Type TWSSocketMsgManagedAccts
    msgType         As TWSSocketInMsgTypes
    version         As Long
    accountsList     As String
End Type

Public Type TWSSocketMsgReceiveFa
    msgType         As TWSSocketInMsgTypes
    version         As Long
    DataType        As Long
    XMLData         As String
End Type


