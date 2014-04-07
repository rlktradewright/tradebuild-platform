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

Private Const ModuleName                            As String = "IBTwsSP"

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
    TwsTickAsk = TICK_ASK
    TwsTickAskSize = TICK_ASK_SIZE
    TwsTickBid = TICK_BID
    TwsTickBidSize = TICK_BID_SIZE
    TwsTickLast = TICK_LAST
    TwsTickLastSize = TICK_LAST_SIZE
    TwsTickHigh = TICK_High
    TwsTickLow = TICK_LOW
    TwsTickVolume = TICK_VOLUME
    TwsTickOpen = TICK_OPEN
    TwsTickClose = TICK_CLOSE
    TwsTickBidOption = TICK_BID_OPTION
    TwsTickAskOption = TICK_ASK_OPTION
    TwsTickLastOption = TICK_LAST_OPTION
    TwsTickModelOption = TICK_MODEL_OPTION
    TwsTickLow13Week = TICK_LOW_13_WEEK
    TwsTickHigh13Week = TICK_HIGH_13_WEEK
    TwsTickLow26Week = TICK_LOW_26_WEEK
    TwsTickHigh26Week = TICK_HIGH_26_WEEK
    TwsTickLow52Week = TICK_LOW_52_WEEK
    TwsTickHigh52Week = TICK_HIGH_52_WEEK
    TwsTickAvgVolume = TICK_AVG_VOLUME
    TwsTickOpenInterest = TICK_OPEN_INTEREST
    TwsTickOptionHistoricalVol = TICK_OPTION_HISTORICAL_VOL
    TwsTickOptionImpliedVol = TICK_OPTION_IMPLIED_VOL
    TwsTickOptionBidExch = TICK_OPTION_BID_EXCH
    TwsTickOptionAskExch = TICK_OPTION_ASK_EXCH
    TwsTickOptionCallOpenInterest = TICK_OPTION_CALL_OPEN_INTEREST
    TwsTickOptionPutOpenInterest = TICK_OPTION_PUT_OPEN_INTEREST
    TwsTickOptionCallVolume = TICK_OPTION_CALL_VOLUME
    TwsTickOptionPutVolume = TICK_OPTION_PUT_VOLUME
    TwsTickIndexFuturePremium = TICK_INDEX_FUTURE_PREMIUM
    TwsTickBidExch = TICK_BID_EXCH
    TwsTickAskExch = TICK_ASK_EXCH
    TwsTickAuctionVolume = TICK_AUCTION_VOLUME
    TwsTickAuctionPrice = TICK_AUCTION_PRICE
    TwsTickAuctionImbalance = TICK_AUCTION_IMBALANCE
    TwsTickMarkPrice = TICK_MARK_PRICE
    TwsTickBidEfpComputation = TICK_BID_EFP_COMPUTATION
    TwsTickAskEfpComputation = TICK_ASK_EFP_COMPUTATION
    TwsTickLastEfpComputation = TICK_LAST_EFP_COMPUTATION
    TwsTickHighEfpComputation = TICK_HIGH_EFP_COMPUTATION
    TwsTickLowEfpComputation = TICK_LOW_EFP_COMPUTATION
    TwsTickOpenEfpComputation = TICK_OPEN_EFP_COMPUTATION
    TwsTickCloseEfpComputation = TICK_CLOSE_EFP_COMPUTATION
    TwsTickLastTimestamp = TICK_LAST_TIMESTAMP
    TwsTickShortable = TICK_SHORTABLE
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TwsHistoricalDataRequest
    'Id                      As Long
    contract                As TwsContract
    EndDateTime             As String
    BarSizeSetting          As String
    Duration                As String
    WhatToShow              As String
    RequestTime             As Date
End Type

Public Type TwsBar
    timeStamp               As String
    OpenValue               As Double
    HighValue               As Double
    LowValue                As Double
    CloseValue              As Double
    Volume                  As Long
    TickVolume              As Long
    WAP                     As Double
    HasGaps                 As Boolean
End Type

Public Type TwsHistDataStartInfo
    RequestId               As Long
    StartDate               As String
    EndDate                 As String
    BarCount                As Long
End Type

Public Type TwsTagValue
    Tag                     As String
    Value                   As String
End Type

Public Type TwsUnderComp
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
                Optional ByVal pConnectionRetryIntervalSecs As Long = 10) As TwsAPI
Const ProcName As String = "GetAPI"
On Error GoTo Err

Set GetAPI = gGetApi(pServer, pPort, pClientId, pConnectionRetryIntervalSecs)

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





