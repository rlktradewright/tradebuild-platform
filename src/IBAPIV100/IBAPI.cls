VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Constants
'@================================================================================

Private Const ModuleName                            As String = "IBAPI"

Private Const TICK_BID_SIZE                      As Long = 0
Private Const TICK_BID                           As Long = 1
Private Const TICK_ASK                           As Long = 2
Private Const TICK_ASK_SIZE                      As Long = 3
Private Const TICK_LAST                          As Long = 4
Private Const TICK_LAST_SIZE                     As Long = 5
Private Const TICK_High                          As Long = 6
Private Const TICK_LOW                           As Long = 7
Private Const TICK_VOLUME                        As Long = 8
Private Const TICK_CLOSE                         As Long = 9
Private Const TICK_BID_OPTION                    As Long = 10
Private Const TICK_ASK_OPTION                    As Long = 11
Private Const TICK_LAST_OPTION                   As Long = 12
Private Const TICK_MODEL_OPTION                  As Long = 13
Private Const TICK_OPEN                          As Long = 14
Private Const TICK_LOW_13_WEEK                   As Long = 15
Private Const TICK_HIGH_13_WEEK                  As Long = 16
Private Const TICK_LOW_26_WEEK                   As Long = 17
Private Const TICK_HIGH_26_WEEK                  As Long = 18
Private Const TICK_LOW_52_WEEK                   As Long = 19
Private Const TICK_HIGH_52_WEEK                  As Long = 20
Private Const TICK_AVG_VOLUME                    As Long = 21
Private Const TICK_OPEN_INTEREST                 As Long = 22
Private Const TICK_OPTION_HISTORICAL_VOL         As Long = 23
Private Const TICK_OPTION_IMPLIED_VOL            As Long = 24
Private Const TICK_OPTION_BID_EXCH               As Long = 25
Private Const TICK_OPTION_ASK_EXCH               As Long = 26
Private Const TICK_OPTION_CALL_OPEN_INTEREST     As Long = 27
Private Const TICK_OPTION_PUT_OPEN_INTEREST      As Long = 28
Private Const TICK_OPTION_CALL_VOLUME            As Long = 29
Private Const TICK_OPTION_PUT_VOLUME             As Long = 30
Private Const TICK_INDEX_FUTURE_PREMIUM          As Long = 31
Private Const TICK_BID_EXCH                      As Long = 32
Private Const TICK_ASK_EXCH                      As Long = 33
Private Const TICK_AUCTION_VOLUME                As Long = 34
Private Const TICK_AUCTION_PRICE                 As Long = 35
Private Const TICK_AUCTION_IMBALANCE             As Long = 36
Private Const TICK_MARK_PRICE                    As Long = 37
Private Const TICK_BID_EFP_COMPUTATION           As Long = 38
Private Const TICK_ASK_EFP_COMPUTATION           As Long = 39
Private Const TICK_LAST_EFP_COMPUTATION          As Long = 40
Private Const TICK_OPEN_EFP_COMPUTATION          As Long = 41
Private Const TICK_HIGH_EFP_COMPUTATION          As Long = 42
Private Const TICK_LOW_EFP_COMPUTATION           As Long = 43
Private Const TICK_CLOSE_EFP_COMPUTATION         As Long = 44
Private Const TICK_LAST_TIMESTAMP                As Long = 45
Private Const TICK_SHORTABLE                     As Long = 46

'@================================================================================
' Enums
'@================================================================================

Public Enum ApiMessageLoggingOptions
    ApiMessageLoggingOptionNone
    ApiMessageLoggingOptionDefault
    ApiMessageLoggingOptionAlways
End Enum

Public Enum TwsAuctionStrategies
    TwsAuctionStrategyMatch = 1
    TwsAuctionStrategyImprovement = 2
    TwsAuctionStrategyTransparent = 3
End Enum

Public Enum TwsConnectionStates
    TwsConnNotConnected
    TwsConnConnecting
    TwsConnConnected
    TwsConnFailed
End Enum

Public Enum TwsDOMOperations
    TwsDOMInsert
    TwsDOMUpdate
    TwsDOMDelete
End Enum

Public Enum TwsDOMSides
    TwsDOMAsk
    TwsDOMBid
End Enum

Public Enum TwsFADataTypes
    TwsFAGroups = 1
    TwsFaProfile
    TwsFAAccountAliases
End Enum

Public Enum TwsHedgeTypes
    TwsHedgeTypeNone
    TwsHedgeTypeDelta
    TwsHedgeTypeBeta
    TwsHedgeTypeFX
    TwsHedgeTypePair
End Enum

Public Enum TwsLogLevels
    TwsLogLevelSystem = 1
    TwsLogLevelError
    TwsLogLevelWarning
    TwsLogLevelInformation
    TwsLogLevelDetail
End Enum

Public Enum TwsLegOpenCloseCodes
    TwsLegOpenCloseSame = 0        ' open/close leg value is same as combo
    TwsLegOpenCloseOpen = 1
    TwsLegOpenCloseClose = 2
    TwsLegOpenCloseUnknown = 3
End Enum

Public Enum TwsLiquidityTypes
    TwsLiquidityTypeNone
    TwsLiquidityTypeAddedLiquidity
    TwsLiquidityTypeRemovedLiquidity
    TwsLiquidityTypeLiquidityRoutedOut
End Enum

Public Enum TwsOcaTypes
    TwsOcaTypeCancelWithBlock = 1
    TwsOcaTypeReduceWithBlock = 2
    TwsOcaTypeReduceNonBlock = 3
End Enum

Public Enum TwsOptionRights
    TwsOptRightNone = 0
    TwsOptRightCall
    TwsOptRightPut
End Enum

Public Enum TwsOrderActions
    TwsOrderActionNone
    TwsOrderActionBuy = 1
    TwsOrderActionSell
    TwsOrderActionSellShort
End Enum
    
Public Enum TwsOrderTIFs
    TwsOrderTIFNone
    TwsOrderTIFDay
    TwsOrderTIFGoodTillCancelled
    TwsOrderTIFImmediateOrCancel
End Enum

Public Enum TwsOrderTypes
    TwsOrderTypeNone = 0
    TwsOrderTypeMarket
    TwsOrderTypeMarketOnClose
    TwsOrderTypeLimit
    TwsOrderTypeLimitOnClose
    TwsOrderTypePeggedToMarket
    TwsOrderTypeStop
    TwsOrderTypeStopLimit
    TwsOrderTypeTrail
    TwsOrderTypeRelative
    TwsOrderTypeVWAP
    TwsOrderTypeMarketToLimit
    TwsOrderTypeQuote
    TwsOrderTypeAdjust
    TwsOrderTypeAlert
    TwsOrderTypeLimitIfTouched
    TwsOrderTypeMarketIfTouched
    TwsOrderTypeTrailLimit
    TwsOrderTypeMarketWithProtection
    TwsOrderTypeMarketOnOpen
    TwsOrderTypeLimitOnOpen
    TwsOrderTypePeggedToPrimary
    TwsOrderTypeVol
    TwsOrderTypePeggedToBenchmark
End Enum

Public Enum TwsSecTypes
    TwsSecTypeNone
    TwsSecTypeStock
    TwsSecTypeFuture
    TwsSecTypeOption
    TwsSecTypeFuturesOption
    TwsSecTypeCash
    TwsSecTypeCombo
    TwsSecTypeIndex
End Enum

Public Enum TwsShortSaleSlotCodes
    TwsShortSaleSlotNotApplicable
    TwsShortSaleSlotClearingBroker = 1
    TwsShortSaleSlotThirdParty = 2
End Enum

Public Enum TwsStopTriggerMethods
    TwsStopTriggerDefault = 0
    TwsStopTriggerDoubleBidAsk = 1
    TwsStopTriggerLast = 2
    TwsStopTriggerDoubleLast = 3
    TwsStopTriggerBidAsk = 4
    TwsStopTriggerLastOrBidAsk = 7
    TwsStopTriggerMidPoint = 8
End Enum

Public Enum TwsTickTypes
    TwsTickBidSize
    TwsTickBid
    TwsTickAsk
    TwsTickAskSize
    TwsTickLast
    TwsTickLastSize
    TwsTickHigh
    TwsTickLow
    TwsTickVolume
    TwsTickClose
    TwsTickBidOption
    TwsTickAskOption
    TwsTickLastOption
    TwsTickModelOption
    TwsTickOpen
    TwsTickLow13Week
    TwsTickHigh13Week
    TwsTickLow26Week
    TwsTickHigh26Week
    TwsTickLow52Week
    TwsTickHigh52Week
    TwsTickAvgVolume
    TwsTickOpenInterest
    TwsTickOptionHistoricalVol
    TwsTickOptionImpliedVol
    TwsTickOptionBidExch
    TwsTickOptionAskExch
    TwsTickOptionCallOpenInterest
    TwsTickOptionPutOpenInterest
    TwsTickOptionCallVolume
    TwsTickOptionPutVolume
    TwsTickIndexFuturePremium
    TwsTickBidExch
    TwsTickAskExch
    TwsTickAuctionVolume
    TwsTickAuctionPrice
    TwsTickAuctionImbalance
    TwsTickMarkPrice
    TwsTickBidEfpComputation
    TwsTickAskEfpComputation
    TwsTickLastEfpComputation
    TwsTickHighEfpComputation
    TwsTickLowEfpComputation
    TwsTickOpenEfpComputation
    TwsTickCloseEfpComputation
    TwsTickLastTimestamp
    TwsTickShortable
    TwsTickFundamentals
    TwsTickRTVolume
    TwsTickHalted
    TwsTickBidYield
    TwsTickAskYield
    TwsTickLastYield
    TwsTickCustOptComp
    TwsTickTrades
    TwsTickTradesPerMin
    TwsTickVolumePerMin
    TwsTickLastRTHTrade
    TwsTickRTHistoricalVol
    TwsTickIBDividends
    TwsTickBondFactorMultiplier
    TwsTickRegulatoryImbalance
    TwsTickNewsTick
    TwsTickShortTermVolume3Min
    TwsTickShortTermVolume5Min
    TwsTickShortTermVolume10Min
    TwsTickDelayedBid
    TwsTickDelayedAsk
    TwsTickDelayedLast
    TwsTickDelayedBidSize
    TwsTickDelayedAskSize
    TwsTickDelayedLastSize
    TwsTickDelayedHigh
    TwsTickDelayedLow
    TwsTickDelayedVolume
    TwsTickDelayedClose
    TwsTickDelayedOpen
    TwsTickRTTrdVolume
    TwsTickCreditmanMarkPrice
    TwsTickCreditmanSlowMarkPrice
    TwsTickDelayedBidOption
    TwsTickDelayedAskOption
    TwsTickDelayedLastOption
    TwsTickDelayedModelOption
    TwsTickLastExchange
    TwsTickLastRegulatoryTime
    TwsTickFuturesOpenInterest
    TwsTickAverageOptionVolume
    TwsTickDelayedLastTimestamp
    
    TwsTickMax = TwsTickDelayedLastTimestamp
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TwsHistoricalDataRequest
    ContractSpec            As TwsContractSpecifier
    EndDateTime             As String
    BarSizeSetting          As String
    Duration                As String
    WhatToShow              As String
    MainSessionOnly         As Boolean
End Type

Public Type TwsBar
    Timestamp               As String
    OpenValue               As Double
    HighValue               As Double
    LowValue                As Double
    CloseValue              As Double
    Volume                  As Long
    TickVolume              As Long
    WAP                     As Double
    HasGaps                 As Boolean
End Type

Public Type TwsPriceIncrement
    LowEdge                 As Double
    Increment               As Double
End Type

Public Type TwsRealtimeBar
    Timestamp               As String
    OpenValue               As Double
    HighValue               As Double
    LowValue                As Double
    CloseValue              As Double
    Volume                  As Long
    TickVolume              As Long
    WAP                     As Double
End Type

Public Type TwsTickAttributes
    CanAutoExecute          As Boolean
    PastLimit               As Boolean
    PreOpen                 As Boolean
    Unreported              As Boolean
    BidPastLow              As Boolean
    AskPastHigh             As Boolean
End Type

Public Type TwsDeltaNeutralContract
    ConId                   As Long
    Delta                   As Double
    Price                   As Double
End Type

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

Public Function GetAPI( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                Optional ByVal pConnectionRetryIntervalSecs As Long = 10, _
                Optional ByVal pLogApiMessages As ApiMessageLoggingOptions = ApiMessageLoggingOptionDefault, _
                Optional ByVal pLogRawApiMessages As ApiMessageLoggingOptions = ApiMessageLoggingOptionDefault, _
                Optional ByVal pLogApiMessageStats As Boolean = False) As TwsAPI
Const ProcName As String = "GetAPI"
On Error GoTo Err

Set GetAPI = gGetApi(pServer, pPort, pClientId, pConnectionRetryIntervalSecs, pLogApiMessages, pLogRawApiMessages, pLogApiMessageStats)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsDateStringToDate( _
                ByRef pDateString As String, _
                Optional ByRef pTimezoneName As String) As Date
Const ProcName As String = "TwsDateStringToDate"
On Error GoTo Err

TwsDateStringToDate = gGetDate(pDateString, pTimezoneName)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsHedgeTypeFromString(ByVal pValue As String) As TwsHedgeTypes
Const ProcName As String = "TwsHedgeTypeFromString"
On Error GoTo Err

TwsHedgeTypeFromString = gTwsHedgeTypeFromString(pValue)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsHedgeTypeToString(ByVal pValue As TwsHedgeTypes) As String
Const ProcName As String = "TwsHedgeTypeToString"
On Error GoTo Err

TwsHedgeTypeToString = gTwsHedgeTypeToString(pValue)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOptionRightFromString( _
                ByVal Value As String) As TwsOptionRights
Const ProcName As String = "TwsOptionRightFromString"

On Error GoTo Err

TwsOptionRightFromString = gTwsOptionRightFromString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOptionRightToString( _
                ByVal Value As TwsOptionRights) As String
Const ProcName As String = "TwsOptionRightToString"

On Error GoTo Err

TwsOptionRightToString = gTwsOptionRightToString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderActionFromString( _
                ByVal Value As String) As TwsOrderActions
Const ProcName As String = "TwsOrderActionFromString"
On Error GoTo Err

TwsOrderActionFromString = gTwsOrderActionFromString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderActionToString( _
                ByVal Value As TwsOrderActions) As String
Const ProcName As String = "TwsOrderActionToString"
On Error GoTo Err

TwsOrderActionToString = gTwsOrderActionToString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderTypeFromString(ByVal Value As String) As TwsOrderTypes
Const ProcName As String = "TwsOrderTypeFromString"
On Error GoTo Err

TwsOrderTypeFromString = gTwsOrderTypeFromString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderTypeToString(ByVal Value As TwsOrderTypes) As String
Const ProcName As String = "TwsOrderTypeToString"
On Error GoTo Err

TwsOrderTypeToString = gTwsOrderTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsSecTypeFromString( _
                ByVal Value As String) As TwsSecTypes
Const ProcName As String = "TwsSecTypeFromString"

On Error GoTo Err

TwsSecTypeFromString = gTwsSecTypeFromString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsSecTypeToShortString( _
                ByVal Value As TwsSecTypes) As String
Const ProcName As String = "TwsSecTypeToShortString"

On Error GoTo Err

TwsSecTypeToShortString = gTwsSecTypeToShortString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsSecTypeToString( _
                ByVal Value As TwsSecTypes) As String
Const ProcName As String = "TwsSecTypeToString"

On Error GoTo Err

TwsSecTypeToString = gTwsSecTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsShortSaleSlotFromString( _
                ByVal Value As String) As TwsShortSaleSlotCodes
Const ProcName As String = "TwsSortSaleSlotFromString"
On Error GoTo Err

TwsShortSaleSlotFromString = gTwsShortSaleSlotFromString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

''
' Brief description
'
' @remarks
'
' @return
'
' @param name
'
' @see
'
'@/
'
Public Function TwsShortSaleSlotToString( _
                ByVal Value As TwsShortSaleSlotCodes) As String
Const ProcName As String = "TwsSortSaleSlotToString"
On Error GoTo Err

TwsShortSaleSlotToString = gTwsShortSaleSlotToString(Value)

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





