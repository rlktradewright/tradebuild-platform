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

'@================================================================================
' Enums
'@================================================================================

Public Enum TwsApiMessageLoggingOptions
    TwsApiMessageLoggingOptionNone
    TWSApiMessageLoggingOptionDefault
    TWSApiMessageLoggingOptionAlways
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

Public Enum TwsLoggingLevels
    TwsLoggingLevelSystem = 1
    TwsLoggingLevelError
    TwsLoggingLevelWarning
    TwsLoggingLevelInformation
    TwsLoggingLevelDetail
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

Public Enum TwsMarketDataTypes
    TwsMarketDataTypeNone
    TwsMarketDataTypeRealtime
    TwsMarketDataTypeFrozen
    TwsMarketDataTypeDelayed
    TwsMarketDataTypeDelayedFrozen
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
    TwsOrderTIFGoodTillDate
    TwsOrderTIFFillOrKill
    TwsOrderTIFDayTillCancelled
    TwsOrderTIFAuction
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
    TwsOrderTypeAuction
    TwsOrderTypePeggedToStock
    TwsOrderTypeBoxTop
    TwsOrderTypePassiveRelative
    TwsOrderTypePeggedToBest
    TwsOrderTypePeggedToMidpoint
    TwsOrderTypeStopWithProtection
    TwsOrderTypeRelativeLimitCombo
    TwsOrderTypeRelativeMarketCombo
    TwsOrderTypeMidprice
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
    TwsSecTypeWarrant
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
    TwsTickShortableShares
    TwsTickDelayedHalted
    TwsTickReuters2MutualFunds
    TwsTickEtfNavClose
    TwsTickEtfNavPriorClose
    TwsTickEtfNavBid
    TwsTickEtfNavAsk
    TwsTickEtfNavLast
    TwsTickEtfFrozenNavLast
    TwsTickEtfNavHigh
    TwsTickEtfNavLow
    TwsTickSocialMarketAnalytics
    TwsTickEstimatedIpoMidpoint
    TwsTickFinalIpoLast
    TwsTickDelayedYieldBid
    TwsTickDelayedYieldAsk
    
    TwsTickMax = TwsTickDelayedYieldAsk
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
    Volume                  As BoxedDecimal
    TickVolume              As Long
    WAP                     As BoxedDecimal
End Type

Public Type TwsPriceIncrement
    LowEdge                 As Double
    Increment               As Double
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
                Optional ByVal pLogApiMessages As TwsApiMessageLoggingOptions = TWSApiMessageLoggingOptionDefault, _
                Optional ByVal pLogRawApiMessages As TwsApiMessageLoggingOptions = TWSApiMessageLoggingOptionDefault, _
                Optional ByVal pLogApiMessageStats As Boolean = False) As TwsAPI
Const ProcName As String = "GetAPI"
On Error GoTo Err

Set GetAPI = GIBApi.GetAPI(pServer, pPort, pClientId, pConnectionRetryIntervalSecs, pLogApiMessages, pLogRawApiMessages, pLogApiMessageStats)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsDateStringToDate( _
                ByRef pDateString As String, _
                Optional ByRef pTimezoneName As String) As Date
Const ProcName As String = "TwsDateStringToDate"
On Error GoTo Err

TwsDateStringToDate = GIBApi.TwsDateStringToDate(pDateString, pTimezoneName)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsHedgeTypeFromString(ByVal pValue As String) As TwsHedgeTypes
Const ProcName As String = "TwsHedgeTypeFromString"
On Error GoTo Err

TwsHedgeTypeFromString = GIBApi.TwsHedgeTypeFromString(pValue)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsHedgeTypeToString(ByVal pValue As TwsHedgeTypes) As String
Const ProcName As String = "TwsHedgeTypeToString"
On Error GoTo Err

TwsHedgeTypeToString = GIBApi.TwsHedgeTypeToString(pValue)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOptionRightFromString( _
                ByVal Value As String) As TwsOptionRights
Const ProcName As String = "TwsOptionRightFromString"

On Error GoTo Err

TwsOptionRightFromString = GIBApi.TwsOptionRightFromString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOptionRightToString( _
                ByVal Value As TwsOptionRights) As String
Const ProcName As String = "TwsOptionRightToString"

On Error GoTo Err

TwsOptionRightToString = GIBApi.TwsOptionRightToString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderActionFromString( _
                ByVal Value As String) As TwsOrderActions
Const ProcName As String = "TwsOrderActionFromString"
On Error GoTo Err

TwsOrderActionFromString = GIBApi.TwsOrderActionFromString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderActionToString( _
                ByVal Value As TwsOrderActions) As String
Const ProcName As String = "TwsOrderActionToString"
On Error GoTo Err

TwsOrderActionToString = GIBApi.TwsOrderActionToString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderTypeFromString(ByVal Value As String) As TwsOrderTypes
Const ProcName As String = "TwsOrderTypeFromString"
On Error GoTo Err

TwsOrderTypeFromString = GIBApi.TwsOrderTypeFromString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsOrderTypeToString(ByVal Value As TwsOrderTypes) As String
Const ProcName As String = "TwsOrderTypeToString"
On Error GoTo Err

TwsOrderTypeToString = GIBApi.TwsOrderTypeToString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsSecTypeFromString( _
                ByVal Value As String) As TwsSecTypes
Const ProcName As String = "TwsSecTypeFromString"

On Error GoTo Err

TwsSecTypeFromString = GIBApi.TwsSecTypeFromString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsSecTypeToShortString( _
                ByVal Value As TwsSecTypes) As String
Const ProcName As String = "TwsSecTypeToShortString"

On Error GoTo Err

TwsSecTypeToShortString = GIBApi.TwsSecTypeToShortString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsSecTypeToString( _
                ByVal Value As TwsSecTypes) As String
Const ProcName As String = "TwsSecTypeToString"

On Error GoTo Err

TwsSecTypeToString = GIBApi.TwsSecTypeToString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsShortSaleSlotFromString( _
                ByVal Value As String) As TwsShortSaleSlotCodes
Const ProcName As String = "TwsSortSaleSlotFromString"
On Error GoTo Err

TwsShortSaleSlotFromString = GIBApi.TwsShortSaleSlotFromString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
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

TwsShortSaleSlotToString = GIBApi.TwsShortSaleSlotToString(Value)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





