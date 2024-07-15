Attribute VB_Name = "GIBApi"
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

Private Const ModuleName                            As String = "GIBApi"

Private Const InvalidEnumValue                   As String = "*ERR*"

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

Public Enum ApiServerVersions
    RANDOMIZE_SIZE_AND_PRICE = 76
    MinV100Plus = 100
    FRACTIONAL_POSITIONS = 101
    PEGGED_TO_BENCHMARK = 102
    MODELS_SUPPORT = 103
    SEC_DEF_OPT_PARAMS_REQ = 104
    EXT_OPERATOR = 105
    SOFT_DOLLAR_TIER = 106
    REQ_FAMILY_CODES = 107
    REQ_MATCHING_SYMBOLS = 108
    PAST_LIMIT = 109
    MD_SIZE_MULTIPLIER = 110
    CASH_QTY = 111
    REQ_MKT_DEPTH_EXCHANGES = 112
    TICK_NEWS = 113
    SMART_COMPONENTS = 114
    REQ_NEWS_PROVIDERS = 115
    REQ_NEWS_ARTICLE = 116
    REQ_HISTORICAL_NEWS = 117
    REQ_HEAD_TIMESTAMP = 118
    REQ_HISTOGRAM_DATA = 119
    SERVICE_DATA_TYPE = 120
    AGG_GROUP = 121
    UNDERLYING_INFO = 122
    CANCEL_HEADTIMESTAMP = 123
    SYNT_REALTIME_BARS = 124
    CFD_REROUTE = 125
    MARKET_RULES = 126
    PNL = 127
    NEWS_QUERY_ORIGINS = 128
    UNREALIZED_PNL = 129
    HISTORICAL_TICKS = 130
    MARKET_CAP_PRICE = 131
    PRE_OPEN_BID_ASK = 132
    REAL_EXPIRATION_DATE = 134
    REALIZED_PNL = 135
    LAST_LIQUIDITY = 136
    TICK_BY_TICK = 137
    DECISION_MAKER = 138
    MIFID_EXECUTION = 139
    TICK_BY_TICK_IGNORE_SIZE = 140
    AUTO_PRICE_FOR_HEDGE = 141
    WHAT_IF_EXT_FIELDS = 142
    SCANNER_GENERIC_OPTS = 143
    API_BIND_ORDER = 144
    ORDER_CONTAINER = 145
    SMART_DEPTH = 146
    D_PEG_ORDERS = 148
    MKT_DEPTH_PRIM_EXCHANGE = 149
    COMPLETED_ORDERS = 150
    PRICE_MGMT_ALGO = 151
    STOCK_TYPE = 152
    ENCODE_MSG_ASCII7 = 153
    SEND_ALL_FAMILY_CODES = 154
    NO_DEFAULT_OPEN_CLOSE = 155
    PRICE_BASED_VOLATILITY = 156
    REPLACE_FA_END = 157
    Duration = 158
    MARKET_DATA_IN_SHARES = 159
    POST_TO_ATS = 160
    WSHE_CALENDAR = 161
    AUTO_CANCEL_PARENT = 162
    FRACTIONAL_SIZE_SUPPORT = 163
    SIZE_RULES = 164
    HISTORICAL_SCHEDULE = 165
    ADVANCED_ORDER_REJECT = 166
    USER_INFO = 167
    CRYPTO_AGGREGATED_TRADES = 168
    MANUAL_ORDER_TIME = 169
    PEGBEST_PEGMID_OFFSETS = 170
    WSH_EVENT_DATA_FILTERS = 171
    IPO_PRICES = 172
    WSH_EVENT_DATA_FILTERS_DATE = 173
    INSTRUMENT_TIMEZONE = 174
    MARKET_DATA_IN_SHARES_1 = 175
    BOND_ISSUERID = 176
    FA_PROFILE_DESUPPORT = 177

    ' Max for API 973.07

    Max = FA_PROFILE_DESUPPORT
End Enum

Public Enum InternalErrorCodes
    DataIncomplete = vbObjectError + 4327   ' let's hope nothing else uses this number!
End Enum

Public Enum TwsSocketInMsgTypes
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
    TICK_OPTION_COMPUTATION = 21
    TICK_GENERIC = 45
    TICK_STRING = 46
    TICK_EFP = 47
    CURRENT_TIME = 49
    REAL_TIME_BARS = 50
    FUNDAMENTAL_DATA = 51
    CONTRACT_DATA_END = 52
    OPEN_ORDER_END = 53
    ACCT_DOWNLOAD_END = 54
    EXECUTION_DATA_END = 55
    DELTA_NEUTRAL_VALIDATION = 56
    TICK_SNAPSHOT_END = 57
    MARKET_DATA_TYPE = 58
    COMMISSION_REPORT = 59
    Position = 61
    POSITION_END = 62
    ACCOUNT_SUMMARY = 63
    ACCOUNT_SUMMARY_END = 64
    VERIFYMESSAGEAPI = 65
    VERIFYCOMPLETED = 66
    DISPLAYGROUPLIST = 67
    DISPLAYGROUPUPDATED = 68
    VERIFYANDAUTHMESSAGEAPI = 69
    VERIFYANDAUTHCOMPLETED = 70
    POSITIONMULTI = 71
    POSITIONMULTIEND = 72
    ACCOUNTUPDATEMULTI = 73
    ACCOUNTUPDATEMULTIEND = 74

    ' Messages from here on don't have a version number
    MaxIdWithVersion = ACCOUNTUPDATEMULTIEND

    OptionParameter = 75
    OptionParameterEnd = 76
    SOFTDOLLARTIERS = 77
    FAMILYCODES = 78
    SYMBOLSAMPLES = 79
    MARKETDEPTHEXCHANGES = 80
    TickRequestParams = 81
    SMARTCOMPONENTS = 82
    NEWSARTICLE = 83
    TICKNEWS = 84
    NEWSPROVIDERS = 85
    HISTORICALNEWS = 86
    HISTORICALNEWSEND = 87
    HEADTIMESTAMP = 88
    HISTOGRAMDATA = 89
    HISTORICALDATAUPDATE = 90
    REROUTEMARKETDATA = 91
    REROUTEMARKETDEPTH = 92
    MarketRule = 93
    PNL = 94
    PNLSINGLE = 95
    HISTORICALTICKMIDPOINT = 96
    HISTORICALTICKBIDASK = 97
    HISTORICALTICKLAST = 98
    TICKBYTICK = 99
    OrderBound = 100
    CompletedOrder = 101
    CompletedOrdersEnd = 102
    ReplaceFAEnd = 103
    WshMetaData = 104
    WshEventData = 105
    HistoricalSchedule = 106
    UserInformation = 107

    MAX_SOCKET_INMSG
End Enum

Public Enum TwsSocketOutMsgTypes
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
    REQ_CURRENT_TIME = 49
    REQ_REAL_TIME_BARS = 50
    CANCEL_REAL_TIME_BARS = 51
    REQ_FUNDAMENTAL_DATA = 52
    CANCEL_FUNDAMENTAL_DATA = 53
    REQ_CALC_IMPLIED_VOLAT = 54
    REQ_CALC_OPTION_PRICE = 55
    CANCEL_CALC_IMPLIED_VOLAT = 56
    CANCEL_CALC_OPTION_PRICE = 57
    REQ_GLOBAL_CANCEL = 58
    REQ_MARKET_DATA_TYPE = 59
    REQ_POSITIONS = 61
    REQ_ACCOUNT_SUMMARY = 62
    CANCEL_ACCOUNT_SUMMARY = 63
    CANCEL_POSITIONS = 64
    VerifyRequest = 65
    VerifyMessage = 66
    QueryDisplayGroups = 67
    SubscribeToGroupEvents = 68
    UpdateDisplayGroup = 69
    UnsubscribeFromGroupEvents = 70
    StartAPI = 71
    VerifyAndAuthRequest = 72
    VerifyAndAuthMessage = 73
    RequestPositionsMulti = 74
    CancelPositionsMulti = 75
    RequestAccountDataMulti = 76
    CancelAccountUpdatesMulti = 77
    RequestOptionParameters = 78
    RequestSoftDollarTiers = 79
    RequestFamilyCodes = 80
    RequestMatchingSymbols = 81
    RequestMarketDepthExchanges = 82
    RequestSmartComponents = 83
    RequestNewsArticle = 84
    RequestNewsProviders = 85
    RequestHistoricalNews = 86
    RequestHeadTimestamp = 87
    RequestHistogramData = 88
    CancelHistogramData = 89
    CancelHeadTimestamp = 90
    RequestMarketRule = 91
    RequestPnL = 92
    CancelPnL = 93
    RequestPnLSingle = 94
    CancelPnLSingle = 95
    RequestHistoricalTickData = 96
    RequestTickByTickData = 97
    CancelTickByTickData = 98
    RequestCompletedOrders = 99
    RequestWshMetaData = 100
    CancelWshMetaData = 101
    RequestWshEventData = 102
    CancelWshEventData = 103
    RequestUserInformation = 104
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private Property Get InputMessageIdMap() As SortedDictionary
Static sMap As SortedDictionary
If sMap Is Nothing Then
    Set sMap = CreateSortedDictionary(KeyTypeInteger)
    setupInputMessageIdMap sMap
End If
Set InputMessageIdMap = sMap
End Property

Private Property Get OutputMessageIdMap() As SortedDictionary
Static sMap As SortedDictionary
If sMap Is Nothing Then
    Set sMap = CreateSortedDictionary(KeyTypeInteger)
    setupOutputMessageIdMap sMap
End If
Set OutputMessageIdMap = sMap
End Property

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

Public Function ByteBufferToString( _
                ByRef pHeader As String, _
                ByRef pBuffer() As Byte) As String
Const ProcName As String = "ByteBufferToString"
On Error GoTo Err

ByteBufferToString = pHeader & Replace(StrConv(pBuffer, vbUnicode), Chr$(0), "_")

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function ContractHasExpired(ByVal pContractSpec As TwsContractSpecifier) As Boolean
If pContractSpec.SecType = TwsSecTypeCash Or _
    pContractSpec.SecType = TwsSecTypeIndex Or _
    pContractSpec.SecType = TwsSecTypeStock _
Then
    ContractHasExpired = False
    Exit Function
End If
    
ContractHasExpired = contractExpiryToDate(pContractSpec) < Now
End Function

Public Function FormatBuffer( _
                ByRef pBuffer() As Byte, _
                ByVal pBufferNextFreeIndex As Long) As String
Dim s As StringBuilder
Dim i As Long
Dim j As Long

Const ProcName As String = "FormatBuffer"
On Error GoTo Err

Set s = CreateStringBuilder
Do While i < pBufferNextFreeIndex
    s.Append Format(i, "0000  ")
    For j = i To i + 49
        If j = pBufferNextFreeIndex Then Exit For
        s.Append IIf(pBuffer(j) <> 0, Chr$(pBuffer(j)), "_")
    Next
    i = i + 50
    If j < pBufferNextFreeIndex Then s.AppendLine ""
Loop
FormatBuffer = s.ToString

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function GetAPI( _
                ByVal pServer As String, _
                ByVal pPort As String, _
                ByVal pClientId As Long, _
                Optional ByVal pConnectionRetryIntervalSecs As Long = 10, _
                Optional ByVal pLogApiMessages As TwsApiMessageLoggingOptions = TWSApiMessageLoggingOptionDefault, _
                Optional ByVal pLogRawApiMessages As TwsApiMessageLoggingOptions = TWSApiMessageLoggingOptionDefault, _
                Optional ByVal pLogApiMessageStats As Boolean = False) As TwsAPI
Const ProcName As String = "GetApi"
On Error GoTo Err

Set GetAPI = New TwsAPI
GetAPI.Initialise pServer, pPort, pClientId, pLogApiMessages, pLogRawApiMessages, pLogApiMessageStats
GetAPI.ConnectionRetryIntervalSecs = pConnectionRetryIntervalSecs

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

' format is yyyymmdd [hh:mm:ss [timezone]]. There can be more than one space
' between the date, time and timezone parts
Public Function GetTwsDate( _
                ByVal pDateString As String, _
                Optional ByRef pTimezoneName As String) As Date
Const ProcName As String = "GetTwsDate"
On Error GoTo Err

If pDateString = "" Then Exit Function

Dim ar() As String
ar = Split(pDateString, " ")

Dim lDatePart As String
lDatePart = ar(0)

Dim lTimePart As String
Dim lTimezoneName As String

Dim i As Long
For i = 1 To UBound(ar)
    If ar(i) <> "" Then
        If lTimePart = "" Then
            lTimePart = ar(i)
        Else
            lTimezoneName = lTimezoneName & " " & ar(i)
        End If
    End If
Next

GetTwsDate = CDate(Left$(lDatePart, 4) & "/" & _
                Mid$(lDatePart, 5, 2) & "/" & _
                Mid$(lDatePart, 7, 2))

If Len(lTimePart) <> 0 Then GetTwsDate = GetTwsDate + CDate(lTimePart)

If Not IsMissing(pTimezoneName) Then pTimezoneName = Trim$(lTimezoneName)

Exit Function

Err:
If Err.Number <> VBErrorCodes.VbErrTypeMismatch Then
    GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End If
GetTwsDate = CDate(0#)
End Function

Public Function GetDateFromUnixSystemTime(ByVal pSystemTime As Double) As Date
Const ProcName As String = "GetDateFromUnixSystemTime"
On Error GoTo Err

GetDateFromUnixSystemTime = CDate(CDbl((2209161600@ + pSystemTime) / (86400@)))

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function InputMessageIdToString( _
                ByVal msgId As TwsSocketInMsgTypes) As String
Const ProcName As String = "InputMessageIdToString"
On Error GoTo Err

Select Case msgId
Case TICK_SIZE
    InputMessageIdToString = "TickSize"
Case TICK_STRING
    InputMessageIdToString = "TickString"
Case MARKET_DEPTH
    InputMessageIdToString = "MarketDepth"
Case TICK_PRICE
    InputMessageIdToString = "TickPrice"
Case Else
    Dim lName As String
    If InputMessageIdMap.TryItem(msgId, lName) Then
        InputMessageIdToString = lName
    Else
        InputMessageIdToString = "?????"
    End If
End Select

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function LongToNetworkBytes(ByVal pNumber As Long) As Byte()
Dim b(3) As Byte
b(0) = ((pNumber And &HFF000000) / &H1000000) And &HFF&
b(1) = (pNumber And &HFF0000) / &H10000
b(2) = (pNumber And &HFF00&) / &H100&
b(3) = (pNumber And &HFF&)
LongToNetworkBytes = b
End Function

Public Function NetworkBytesToLong(ByRef pBytes() As Byte) As Long
Dim l As Long
l = pBytes(0) * &H1000000 + _
    pBytes(1) * &H10000 + _
    pBytes(2) * &H100& + _
    pBytes(3)
NetworkBytesToLong = l
End Function

Public Function OutputMessageIdToString( _
                ByVal msgId As TwsSocketOutMsgTypes) As String
Const ProcName As String = "OutputMessageIdToString"
On Error GoTo Err

Dim lName As String
If OutputMessageIdMap.TryItem(msgId, lName) Then
    OutputMessageIdToString = lName
Else
    OutputMessageIdToString = "?????"
End If

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub SetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Private Function TruncateTimeToNextMinute(ByVal pTimeStamp As Date) As Date
Const ProcName As String = "TruncateTimeToNextMinute"
On Error GoTo Err

TruncateTimeToNextMinute = Int((pTimeStamp + GIB.OneMinute - GIB.OneMicrosecond) / GIB.OneMinute) * GIB.OneMinute

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsDateStringToDate( _
                ByRef pDateString As String, _
                Optional ByRef pTimezoneName As String) As Date
Const ProcName As String = "TwsDateStringToDate"
On Error GoTo Err

TwsDateStringToDate = GetTwsDate(pDateString, pTimezoneName)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsHedgeTypeFromString(ByVal pValue As String) As TwsHedgeTypes
Select Case UCase$(pValue)
Case ""
    TwsHedgeTypeFromString = TwsHedgeTypeNone
Case "D", "DELTA"
    TwsHedgeTypeFromString = TwsHedgeTypeDelta
Case "B", "BETA"
    TwsHedgeTypeFromString = TwsHedgeTypeBeta
Case "F"
    TwsHedgeTypeFromString = TwsHedgeTypeFX
Case "P", "PAIR"
    TwsHedgeTypeFromString = TwsHedgeTypePair
Case Else
    AssertArgument False, "Value is not a valid Hedge Type"
End Select
End Function

Public Function TwsHedgeTypeToString(ByVal pValue As TwsHedgeTypes) As String
Select Case pValue
Case TwsHedgeTypeNone
    TwsHedgeTypeToString = ""
Case TwsHedgeTypeDelta
    TwsHedgeTypeToString = "D"
Case TwsHedgeTypeBeta
    TwsHedgeTypeToString = "B"
Case TwsHedgeTypeFX
    TwsHedgeTypeToString = "F"
Case TwsHedgeTypePair
    TwsHedgeTypeToString = "P"
Case Else
    AssertArgument False, "Value is not a valid Hedge Type"
End Select
End Function

Public Function TwsOptionRightToString(ByVal Value As TwsOptionRights) As String
Select Case Value
Case TwsOptRightNone
    TwsOptionRightToString = ""
Case TwsOptRightCall
    TwsOptionRightToString = "Call"
Case TwsOptRightPut
    TwsOptionRightToString = "Put"
Case Else
    TwsOptionRightToString = InvalidEnumValue
End Select
End Function

Public Function TwsOptionRightFromString(ByVal Value As String) As TwsOptionRights
Select Case UCase$(Value)
Case "", "?", "0"
    TwsOptionRightFromString = TwsOptRightNone
Case "CALL", "C"
    TwsOptionRightFromString = TwsOptRightCall
Case "PUT", "P"
    TwsOptionRightFromString = TwsOptRightPut
Case Else
    AssertArgument False, "Value is not a valid Option Right"
End Select
End Function

Public Function TwsOrderActionFromString(ByVal Value As String) As TwsOrderActions
Select Case UCase$(Value)
Case ""
    TwsOrderActionFromString = TwsOrderActionNone
Case "BUY"
    TwsOrderActionFromString = TwsOrderActionBuy
Case "SELL"
    TwsOrderActionFromString = TwsOrderActionSell
Case "SSHORT"
    TwsOrderActionFromString = TwsOrderActionSellShort
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function

Public Function TwsOrderActionToString(ByVal Value As TwsOrderActions) As String
Select Case Value
Case TwsOrderActionBuy
    TwsOrderActionToString = "BUY"
Case TwsOrderActionSell
    TwsOrderActionToString = "SELL"
Case TwsOrderActionSellShort
    TwsOrderActionToString = "SSHORT"
Case TwsOrderActionNone
    TwsOrderActionToString = ""
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function
Public Function TwsOrderTIFFromString(ByVal Value As String) As TwsOrderTIFs
Select Case UCase$(Value)
Case ""
    TwsOrderTIFFromString = TwsOrderTIFNone
Case "DAY"
    TwsOrderTIFFromString = TwsOrderTIFDay
Case "GTC"
    TwsOrderTIFFromString = TwsOrderTIFGoodTillCancelled
Case "IOC"
    TwsOrderTIFFromString = TwsOrderTIFImmediateOrCancel
Case "FOK"
    TwsOrderTIFFromString = TwsOrderTIFFillOrKill
Case "DTC"
    TwsOrderTIFFromString = TwsOrderTIFDayTillCancelled
Case "AUC"
    TwsOrderTIFFromString = TwsOrderTIFAuction
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Tws Order TIF"
End Select
End Function

Public Function TwsOrderTIFToString(ByVal Value As TwsOrderTIFs) As String
Select Case Value
Case TwsOrderTIFs.TwsOrderTIFDay
    TwsOrderTIFToString = "DAY"
Case TwsOrderTIFs.TwsOrderTIFGoodTillCancelled
    TwsOrderTIFToString = "GTC"
Case TwsOrderTIFs.TwsOrderTIFImmediateOrCancel
    TwsOrderTIFToString = "IOC"
Case TwsOrderTIFGoodTillDate
    TwsOrderTIFToString = "GTD"
Case TwsOrderTIFFillOrKill
    TwsOrderTIFToString = "FOK"
Case TwsOrderTIFDayTillCancelled
    TwsOrderTIFToString = "DTC"
Case TwsOrderTIFAuction
    TwsOrderTIFToString = "AUC"
Case TwsOrderTIFs.TwsOrderTIFNone
    TwsOrderTIFToString = ""
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Tws Order TIF"
End Select
End Function

Public Function TwsOrderTypeFromString(ByVal Value As String) As TwsOrderTypes
Select Case UCase$(Value)
Case ""
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeNone
Case "MKT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarket
Case "MOC"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case "LMT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimit
Case "LOC"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case "PEG MKT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToMarket
Case "PEGMKT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToMarket
Case "STP"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStop
Case "STP LMT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStopLimit
Case "STPLMT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStopLimit
Case "TRAIL"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeTrail
Case "REL"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeRelative
Case "VWAP"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeVWAP
Case "MTL"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketToLimit
Case "RFQ"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeQuote
Case "ADJUST"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeAdjust
Case "ALERT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeAlert
Case "LIT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitIfTouched
Case "MIT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketIfTouched
Case "TRAIL LIMIT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeTrailLimit
Case "TRAILLIMIT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeTrailLimit
Case "MKT PROT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketWithProtection
Case "MKTPROT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketWithProtection
Case "MKT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnOpen
Case "LMT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnOpen
Case "REL"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToPrimary
Case "VOL"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeVol
Case "PEG BENCH"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToBenchmark
Case "PEGBENCH"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToBenchmark
Case "AUC"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeAuction
Case "PEG STK"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToStock
Case "PEGSTK"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToStock
Case "BOX TOP"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeBoxTop
Case "BOXTOP"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeBoxTop
Case "PASSV REL"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePassiveRelative
Case "PASSVREL"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePassiveRelative
Case "PEG MID"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToMidpoint
Case "PEGMID"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToMidpoint
Case "STP PRT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStopWithProtection
Case "STPPRT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStopWithProtection
Case "REL + LMT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeRelativeLimitCombo
Case "REL+LMT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeRelativeLimitCombo
Case "REL + MKT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeRelativeMarketCombo
Case "REL+MKT"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeRelativeMarketCombo
Case "PEG BEST"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToBest
Case "PEGBEST"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToBest
Case "MIDPRICE"
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMidprice
Case Else
    TwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeNone
End Select
End Function

Public Function TwsOrderTypeToString(ByVal Value As TwsOrderTypes) As String
Const ProcName As String = "TwsOrderTypeToString"
On Error GoTo Err

Select Case Value
Case TwsOrderTypes.TwsOrderTypeNone
    TwsOrderTypeToString = ""
Case TwsOrderTypes.TwsOrderTypeMarket
    TwsOrderTypeToString = "MKT"
Case TwsOrderTypes.TwsOrderTypeMarketOnClose
    TwsOrderTypeToString = "MOC"
Case TwsOrderTypes.TwsOrderTypeLimit
    TwsOrderTypeToString = "LMT"
Case TwsOrderTypes.TwsOrderTypeLimitOnClose
    TwsOrderTypeToString = "LOC"
Case TwsOrderTypes.TwsOrderTypePeggedToMarket
    TwsOrderTypeToString = "PEG MKT"
Case TwsOrderTypes.TwsOrderTypeStop
    TwsOrderTypeToString = "STP"
Case TwsOrderTypes.TwsOrderTypeStopLimit
    TwsOrderTypeToString = "STP LMT"
Case TwsOrderTypes.TwsOrderTypeTrail
    TwsOrderTypeToString = "TRAIL"
Case TwsOrderTypes.TwsOrderTypeRelative
    TwsOrderTypeToString = "REL"
Case TwsOrderTypes.TwsOrderTypeVWAP
    TwsOrderTypeToString = "VWAP"
Case TwsOrderTypes.TwsOrderTypeMarketToLimit
    TwsOrderTypeToString = "MTL"
Case TwsOrderTypes.TwsOrderTypeQuote
    TwsOrderTypeToString = "RFQ"
Case TwsOrderTypes.TwsOrderTypeAdjust
    TwsOrderTypeToString = "ADJUST"
Case TwsOrderTypes.TwsOrderTypeAlert
    TwsOrderTypeToString = "ALERT"
Case TwsOrderTypes.TwsOrderTypeLimitIfTouched
    TwsOrderTypeToString = "LIT"
Case TwsOrderTypes.TwsOrderTypeMarketIfTouched
    TwsOrderTypeToString = "MIT"
Case TwsOrderTypes.TwsOrderTypeTrailLimit
    TwsOrderTypeToString = "TRAIL LIMIT"
Case TwsOrderTypes.TwsOrderTypeMarketWithProtection
    TwsOrderTypeToString = "MKT PROT"
Case TwsOrderTypes.TwsOrderTypeMarketOnOpen
    TwsOrderTypeToString = "MKT"
Case TwsOrderTypes.TwsOrderTypeLimitOnOpen
    TwsOrderTypeToString = "LMT"
Case TwsOrderTypes.TwsOrderTypePeggedToPrimary
    TwsOrderTypeToString = "REL"
Case TwsOrderTypes.TwsOrderTypeVol
    TwsOrderTypeToString = "VOL"
Case TwsOrderTypes.TwsOrderTypePeggedToBenchmark
    TwsOrderTypeToString = "PEG BENCH"
Case TwsOrderTypes.TwsOrderTypeAuction
    TwsOrderTypeToString = "AUC"
Case TwsOrderTypes.TwsOrderTypePeggedToStock
    TwsOrderTypeToString = "PEG STK"
Case TwsOrderTypes.TwsOrderTypeBoxTop
    TwsOrderTypeToString = "BOX TOP"
Case TwsOrderTypes.TwsOrderTypePassiveRelative
    TwsOrderTypeToString = "PASSV REL"
Case TwsOrderTypes.TwsOrderTypePeggedToMidpoint
    TwsOrderTypeToString = "PEG MID"
Case TwsOrderTypes.TwsOrderTypeStopWithProtection
    TwsOrderTypeToString = "STP PRT"
Case TwsOrderTypes.TwsOrderTypeRelativeLimitCombo
    TwsOrderTypeToString = "REL + LMT"
Case TwsOrderTypes.TwsOrderTypeRelativeMarketCombo
    TwsOrderTypeToString = "REL + MKT"
Case TwsOrderTypes.TwsOrderTypePeggedToBest
    TwsOrderTypeToString = "PEG BEST"
Case TwsOrderTypes.TwsOrderTypeMidprice
    TwsOrderTypeToString = "MIDPRICE"
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException
End Select

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function TwsSecTypeFromString(ByVal Value As String) As TwsSecTypes
Select Case UCase$(Value)
Case "", "UNK"
    TwsSecTypeFromString = TwsSecTypeNone
Case "STOCK", "STK"
    TwsSecTypeFromString = TwsSecTypeStock
Case "FUTURE", "FUT"
    TwsSecTypeFromString = TwsSecTypeFuture
Case "OPTION", "OPT"
    TwsSecTypeFromString = TwsSecTypeOption
Case "FUTURES OPTION", "FOP", "EC"  ' "EC" is event contract - see:
                                    ' https://www.cmegroup.com/content/dam/cmegroup/notices/ser/2022/03/SER-8968.pdf
                                    ' https://www.cmegroup.com/activetrader/event-contracts/contract-specifications-for-event-contracts.html
    TwsSecTypeFromString = TwsSecTypeFuturesOption
Case "CASH"
    TwsSecTypeFromString = TwsSecTypeCash
Case "BAG", "COMBO", "CMB"
    TwsSecTypeFromString = TwsSecTypeCombo
Case "INDEX", "IND"
    TwsSecTypeFromString = TwsSecTypeIndex
Case "WARRANT", "WAR"
    TwsSecTypeFromString = TwsSecTypeWarrant
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Security Type"
End Select
End Function

Public Function TwsSecTypeToShortString(ByVal Value As TwsSecTypes) As String
Select Case Value
Case TwsSecTypeStock
    TwsSecTypeToShortString = "STK"
Case TwsSecTypeFuture
    TwsSecTypeToShortString = "FUT"
Case TwsSecTypeOption
    TwsSecTypeToShortString = "OPT"
Case TwsSecTypeFuturesOption
    TwsSecTypeToShortString = "FOP"
Case TwsSecTypeCash
    TwsSecTypeToShortString = "CASH"
Case TwsSecTypeCombo
    TwsSecTypeToShortString = "BAG"
Case TwsSecTypeIndex
    TwsSecTypeToShortString = "IND"
Case TwsSecTypeNone
    TwsSecTypeToShortString = ""
Case TwsSecTypeWarrant
    TwsSecTypeToShortString = "WAR"
Case Else
    TwsSecTypeToShortString = InvalidEnumValue
End Select
End Function

Public Function TwsSecTypeToString(ByVal Value As TwsSecTypes) As String
Select Case Value
Case TwsSecTypeStock
    TwsSecTypeToString = "Stock"
Case TwsSecTypeFuture
    TwsSecTypeToString = "Future"
Case TwsSecTypeOption
    TwsSecTypeToString = "Option"
Case TwsSecTypeFuturesOption
    TwsSecTypeToString = "Futures Option"
Case TwsSecTypeCash
    TwsSecTypeToString = "Cash"
Case TwsSecTypeCombo
    TwsSecTypeToString = "Combo"
Case TwsSecTypeIndex
    TwsSecTypeToString = "Index"
Case TwsSecTypeNone
    TwsSecTypeToString = ""
Case TwsSecTypeWarrant
    TwsSecTypeToString = "Warrant"
Case Else
    TwsSecTypeToString = InvalidEnumValue
End Select
End Function

Public Function TwsShortSaleSlotFromString( _
                ByVal Value As String) As TwsShortSaleSlotCodes
Select Case UCase$(Value)
Case "", "N/A", "NOT APPLICABLE"
    TwsShortSaleSlotFromString = TwsShortSaleSlotNotApplicable
Case "B", "CB", "BROKER", "CLEARING BROKER"
    TwsShortSaleSlotFromString = TwsShortSaleSlotClearingBroker
Case "T", "TP", "THIRD PARTY"
    TwsShortSaleSlotFromString = TwsShortSaleSlotThirdParty
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Short Sale Slot"
End Select
End Function

Public Function TwsShortSaleSlotToString( _
                ByVal Value As TwsShortSaleSlotCodes) As String
Select Case Value
Case TwsShortSaleSlotNotApplicable
    TwsShortSaleSlotToString = "N/A"
Case TwsShortSaleSlotClearingBroker
    TwsShortSaleSlotToString = "BROKER"
Case TwsShortSaleSlotThirdParty
    TwsShortSaleSlotToString = "THIRD PARTY"
Case Else
    TwsShortSaleSlotToString = InvalidEnumValue
End Select
End Function

Private Function TwsStopTriggerMethodFromString(ByVal Value As String) As TwsStopTriggerMethods
Select Case Value
Case "Default"
    TwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDefault
Case "Double Bid/Ask"
    TwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk
Case "Last"
    TwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerLast
Case "Double Last"
    TwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDoubleLast
Case "Bid/Ask"
    TwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerBidAsk
Case "Last or Bid/Ask"
    TwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk
Case "Midpoint"
    TwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerMidPoint
End Select
End Function

Private Function TwsStopTriggerMethodToString(ByVal Value As TwsStopTriggerMethods) As String
Select Case Value
Case TwsStopTriggerMethods.TwsStopTriggerBidAsk
    TwsStopTriggerMethodToString = "Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerDefault
    TwsStopTriggerMethodToString = "Default"
Case TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk
    TwsStopTriggerMethodToString = "Double Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerDoubleLast
    TwsStopTriggerMethodToString = "Double Last"
Case TwsStopTriggerMethods.TwsStopTriggerLast
    TwsStopTriggerMethodToString = "Last"
Case TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk
    TwsStopTriggerMethodToString = "Last or Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerMidPoint
    TwsStopTriggerMethodToString = "Midpoint"
End Select
End Function

Public Sub WSAIoctlCompletionRoutine( _
                ByVal dwError As Long, _
                ByVal cbTransferred As Long, _
                ByVal lpOverlapped As Long, _
                ByVal dwFlags As Long)
                
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function contractExpiryToDate( _
                ByVal pContractSpec As TwsContractSpecifier) As Date
Const ProcName As String = "contractExpiryToDate"
On Error GoTo Err

If Len(pContractSpec.Expiry) = 8 Then
    contractExpiryToDate = CDate(Left$(pContractSpec.Expiry, 4) & "/" & Mid$(pContractSpec.Expiry, 5, 2) & "/" & Right$(pContractSpec.Expiry, 2))
ElseIf Len(pContractSpec.Expiry) = 6 Then
    contractExpiryToDate = CDate(Left$(pContractSpec.Expiry, 4) & "/" & Mid$(pContractSpec.Expiry, 5, 2) & "/" & "01")
End If

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub addInputMessageIdMapEntry( _
                ByVal pMap As SortedDictionary, _
                ByVal pMessageId As TwsSocketInMsgTypes, _
                ByVal pMessageName As String)
pMap.Add pMessageName, pMessageId
End Sub

Private Sub addOutputMessageIdMapEntry( _
                ByVal pMap As SortedDictionary, _
                ByVal pMessageId As TwsSocketOutMsgTypes, _
                ByVal pMessageName As String)
pMap.Add pMessageName, pMessageId
End Sub

Private Sub setupInputMessageIdMap(ByVal pMap As SortedDictionary)
addInputMessageIdMapEntry pMap, TICK_PRICE, "TICKPRICE"
addInputMessageIdMapEntry pMap, TICK_SIZE, "TICKSIZE"
addInputMessageIdMapEntry pMap, ORDER_STATUS, "ORDERSTATUS"
addInputMessageIdMapEntry pMap, ERR_MSG, "ERRORMESSAGE"
addInputMessageIdMapEntry pMap, OPEN_ORDER, "OPENORDER"
addInputMessageIdMapEntry pMap, ACCT_VALUE, "ACCOUNTVALUE"
addInputMessageIdMapEntry pMap, PORTFOLIO_VALUE, "PORTFOLIOVALUE"
addInputMessageIdMapEntry pMap, ACCT_UPDATE_TIME, "ACCOUNTUPDATETIME"
addInputMessageIdMapEntry pMap, NEXT_VALID_ID, "NEXTVALIDID"
addInputMessageIdMapEntry pMap, CONTRACT_DATA, "CONTRACTDATA"
addInputMessageIdMapEntry pMap, EXECUTION_DATA, "EXECUTIONDATA"
addInputMessageIdMapEntry pMap, MARKET_DEPTH, "MARKETDEPTH"
addInputMessageIdMapEntry pMap, MARKET_DEPTH_L2, "MARKETDEPTHL2"
addInputMessageIdMapEntry pMap, NEWS_BULLETINS, "NEWBULLETIN"
addInputMessageIdMapEntry pMap, MANAGED_ACCTS, "MANAGEDACCOUNTS"
addInputMessageIdMapEntry pMap, RECEIVE_FA, "RECEIVEFA"
addInputMessageIdMapEntry pMap, HISTORICAL_DATA, "HISTORICALDATA"
addInputMessageIdMapEntry pMap, BOND_CONTRACT_DATA, "BONDCONTRACTDATA"
addInputMessageIdMapEntry pMap, SCANNER_PARAMETERS, "SCANNERPARAMETERS"
addInputMessageIdMapEntry pMap, SCANNER_DATA, "SCANNERDATA"
addInputMessageIdMapEntry pMap, TICK_OPTION_COMPUTATION, "OPTIONCOMPUTATION"
addInputMessageIdMapEntry pMap, TICK_GENERIC, "GENERIC"
addInputMessageIdMapEntry pMap, TICK_STRING, "STRING"
addInputMessageIdMapEntry pMap, TICK_EFP, "EFP"
addInputMessageIdMapEntry pMap, CURRENT_TIME, "CURRENTTIME"
addInputMessageIdMapEntry pMap, REAL_TIME_BARS, "REALTIMEBAR"
addInputMessageIdMapEntry pMap, FUNDAMENTAL_DATA, "FUNDAMENTALDATA"
addInputMessageIdMapEntry pMap, CONTRACT_DATA_END, "CONTRACTDATAEND"
addInputMessageIdMapEntry pMap, OPEN_ORDER_END, "OPENORDEREND"
addInputMessageIdMapEntry pMap, ACCT_DOWNLOAD_END, "ACCOUNTDOWNLOADEND"
addInputMessageIdMapEntry pMap, EXECUTION_DATA_END, "EXECUTIONDATAEND"
addInputMessageIdMapEntry pMap, DELTA_NEUTRAL_VALIDATION, "DELTANEUTRALVALIDN"
addInputMessageIdMapEntry pMap, TICK_SNAPSHOT_END, "TICKSNAPSHOTEND"
addInputMessageIdMapEntry pMap, MARKET_DATA_TYPE, "MARKETDATATYPE"
addInputMessageIdMapEntry pMap, COMMISSION_REPORT, "COMMISSIONREPORT"
addInputMessageIdMapEntry pMap, Position, "POSITION"
addInputMessageIdMapEntry pMap, POSITION_END, "POSITIONEND"
addInputMessageIdMapEntry pMap, ACCOUNT_SUMMARY, "ACCOUNTSUMMARY"
addInputMessageIdMapEntry pMap, ACCOUNT_SUMMARY_END, "ACCOUNTSUMMARYEND"
addInputMessageIdMapEntry pMap, VERIFYMESSAGEAPI, "VERIFYMESSAGEAPI"
addInputMessageIdMapEntry pMap, VERIFYCOMPLETED, "VERIFYCOMPLETED"
addInputMessageIdMapEntry pMap, DISPLAYGROUPLIST, "DISPLAYGROUPLIST"
addInputMessageIdMapEntry pMap, DISPLAYGROUPUPDATED, "DISPLAYGROUPUPD"
addInputMessageIdMapEntry pMap, VERIFYANDAUTHMESSAGEAPI, "VERIFY/AUTHMESSAGEAPI"
addInputMessageIdMapEntry pMap, VERIFYANDAUTHCOMPLETED, "VERIFY/AUTHCOMPLETED"
addInputMessageIdMapEntry pMap, POSITIONMULTI, "POSITIONMULTI"
addInputMessageIdMapEntry pMap, POSITIONMULTIEND, "POSITIONMULTIEND"
addInputMessageIdMapEntry pMap, ACCOUNTUPDATEMULTI, "ACCOUNTUPDATEMULTI"
addInputMessageIdMapEntry pMap, ACCOUNTUPDATEMULTIEND, "ACCOUNTUPDATEMULTIEND"
addInputMessageIdMapEntry pMap, OptionParameter, "OPTIONPARAMETER"
addInputMessageIdMapEntry pMap, OptionParameterEnd, "OPTIONPARAMETEREND"
addInputMessageIdMapEntry pMap, SOFTDOLLARTIERS, "SOFTDOLLARTIERS"
addInputMessageIdMapEntry pMap, FAMILYCODES, "FAMILYCODES"
addInputMessageIdMapEntry pMap, SYMBOLSAMPLES, "SYMBOLSAMPLES"
addInputMessageIdMapEntry pMap, MARKETDEPTHEXCHANGES, "MARKETDEPTHEXCHANGES"
addInputMessageIdMapEntry pMap, TickRequestParams, "TICKREQUESTPARAMS"
addInputMessageIdMapEntry pMap, SMARTCOMPONENTS, "SMARTCOMPONENTS"
addInputMessageIdMapEntry pMap, NEWSARTICLE, "NEWSARTICLE"
addInputMessageIdMapEntry pMap, TICKNEWS, "TICKNEWS"
addInputMessageIdMapEntry pMap, NEWSPROVIDERS, "NEWSPROVIDERS"
addInputMessageIdMapEntry pMap, HISTORICALNEWS, "HISTORICALNEWS"
addInputMessageIdMapEntry pMap, HISTORICALNEWSEND, "HISTORICALNEWSEND"
addInputMessageIdMapEntry pMap, HEADTIMESTAMP, "HEADTIMESTAMP"
addInputMessageIdMapEntry pMap, HISTOGRAMDATA, "HISTOGRAMDATA"
addInputMessageIdMapEntry pMap, HISTORICALDATAUPDATE, "HISTORICALDATAUPDATE"
addInputMessageIdMapEntry pMap, REROUTEMARKETDATA, "REROUTEMARKETDATA"
addInputMessageIdMapEntry pMap, REROUTEMARKETDEPTH, "REROUTEMARKETDEPTH"
addInputMessageIdMapEntry pMap, MarketRule, "MARKETRULE"
addInputMessageIdMapEntry pMap, TwsSocketInMsgTypes.PNL, "PNL"
addInputMessageIdMapEntry pMap, PNLSINGLE, "PNLSINGLE"
addInputMessageIdMapEntry pMap, HISTORICALTICKMIDPOINT, "HISTORICALTICKMIDPOINT"
addInputMessageIdMapEntry pMap, HISTORICALTICKBIDASK, "HISTORICALTICKBIDASK"
addInputMessageIdMapEntry pMap, HISTORICALTICKLAST, "HISTORICALTICKLAST"
addInputMessageIdMapEntry pMap, TICKBYTICK, "TICKBYTICK"
End Sub

Private Sub setupOutputMessageIdMap(ByVal pMap As SortedDictionary)
addOutputMessageIdMapEntry pMap, REQ_MKT_DATA, "REQ_MKT_DATA"
addOutputMessageIdMapEntry pMap, CANCEL_MKT_DATA, "CANCEL_MKT_DATA"
addOutputMessageIdMapEntry pMap, PLACE_ORDER, "PLACE_ORDER"
addOutputMessageIdMapEntry pMap, CANCEL_ORDER, "CANCEL_ORDER"
addOutputMessageIdMapEntry pMap, REQ_OPEN_ORDERS, "REQ_OPEN_ORDERS"
addOutputMessageIdMapEntry pMap, REQ_ACCT_DATA, "REQ_ACCT_DATA"
addOutputMessageIdMapEntry pMap, REQ_EXECUTIONS, "REQ_EXECUTIONS"
addOutputMessageIdMapEntry pMap, REQ_IDS, "REQ_IDS"
addOutputMessageIdMapEntry pMap, REQ_CONTRACT_DATA, "REQ_CONTRACT_DATA"
addOutputMessageIdMapEntry pMap, REQ_MKT_DEPTH, "REQ_MKT_DEPTH"
addOutputMessageIdMapEntry pMap, CANCEL_MKT_DEPTH, "CANCEL_MKT_DEPTH"
addOutputMessageIdMapEntry pMap, REQ_NEWS_BULLETINS, "REQ_NEWS_BULLETINS"
addOutputMessageIdMapEntry pMap, CANCEL_NEWS_BULLETINS, "CANCEL_NEWS_BULLETINS"
addOutputMessageIdMapEntry pMap, SET_SERVER_LOGLEVEL, "SET_SERVER_LOGLEVEL"
addOutputMessageIdMapEntry pMap, REQ_AUTO_OPEN_ORDERS, "REQ_AUTO_OPEN_ORDERS"
addOutputMessageIdMapEntry pMap, REQ_ALL_OPEN_ORDERS, "REQ_ALL_OPEN_ORDERS"
addOutputMessageIdMapEntry pMap, REQ_MANAGED_ACCTS, "REQ_MANAGED_ACCTS"
addOutputMessageIdMapEntry pMap, REQ_FA, "REQ_FA"
addOutputMessageIdMapEntry pMap, REPLACE_FA, "REPLACE_FA"
addOutputMessageIdMapEntry pMap, REQ_HISTORICAL_DATA, "REQ_HISTORICAL_DATA"
addOutputMessageIdMapEntry pMap, EXERCISE_OPTIONS, "EXERCISE_OPTIONS"
addOutputMessageIdMapEntry pMap, REQ_SCANNER_SUBSCRIPTION, "REQ_SCANNER_SUBSCRIPTION"
addOutputMessageIdMapEntry pMap, CANCEL_SCANNER_SUBSCRIPTION, "CANCEL_SCANNER_SUBSCRIPTION"
addOutputMessageIdMapEntry pMap, REQ_SCANNER_PARAMETERS, "REQ_SCANNER_PARAMETERS"
addOutputMessageIdMapEntry pMap, CANCEL_HISTORICAL_DATA, "CANCEL_HISTORICAL_DATA"
addOutputMessageIdMapEntry pMap, REQ_CURRENT_TIME, "REQ_CURRENT_TIME"
addOutputMessageIdMapEntry pMap, REQ_REAL_TIME_BARS, "REQ_REAL_TIME_BARS"
addOutputMessageIdMapEntry pMap, CANCEL_REAL_TIME_BARS, "CANCEL_REAL_TIME_BARS"
addOutputMessageIdMapEntry pMap, REQ_FUNDAMENTAL_DATA, "REQ_FUNDAMENTAL_DATA"
addOutputMessageIdMapEntry pMap, CANCEL_FUNDAMENTAL_DATA, "CANCEL_FUNDAMENTAL_DATA"
addOutputMessageIdMapEntry pMap, REQ_CALC_IMPLIED_VOLAT, "REQ_CALC_IMPLIED_VOLAT"
addOutputMessageIdMapEntry pMap, REQ_CALC_OPTION_PRICE, "REQ_CALC_OPTION_PRICE"
addOutputMessageIdMapEntry pMap, CANCEL_CALC_IMPLIED_VOLAT, "CANCEL_CALC_IMPLIED_VOLAT"
addOutputMessageIdMapEntry pMap, CANCEL_CALC_OPTION_PRICE, "CANCEL_CALC_OPTION_PRICE"
addOutputMessageIdMapEntry pMap, REQ_GLOBAL_CANCEL, "REQ_GLOBAL_CANCEL"
addOutputMessageIdMapEntry pMap, REQ_MARKET_DATA_TYPE, "REQ_MARKET_DATA_TYPE"
addOutputMessageIdMapEntry pMap, REQ_POSITIONS, "REQ_POSITIONS"
addOutputMessageIdMapEntry pMap, REQ_ACCOUNT_SUMMARY, "REQ_ACCOUNT_SUMMARY"
addOutputMessageIdMapEntry pMap, CANCEL_ACCOUNT_SUMMARY, "CANCEL_ACCOUNT_SUMMARY"
addOutputMessageIdMapEntry pMap, CANCEL_POSITIONS, "CANCEL_POSITIONS"
addOutputMessageIdMapEntry pMap, VerifyRequest, "VERIFY_REQUEST"
addOutputMessageIdMapEntry pMap, VerifyMessage, "VERIFY_MESSAGE"
addOutputMessageIdMapEntry pMap, QueryDisplayGroups, "QUERY_DISPLAY_GROUPS"
addOutputMessageIdMapEntry pMap, SubscribeToGroupEvents, "SUBSCRIBE_TO_GROUP_EVENTS"
addOutputMessageIdMapEntry pMap, UpdateDisplayGroup, "UPDATE_DISPLAY_GROUP"
addOutputMessageIdMapEntry pMap, UnsubscribeFromGroupEvents, "UNSUBSCRIBE_FROM_GROUP_EVENTS"
addOutputMessageIdMapEntry pMap, StartAPI, "START_API"
addOutputMessageIdMapEntry pMap, VerifyAndAuthRequest, "VERIFY_AND_AUTH_REQ"
addOutputMessageIdMapEntry pMap, VerifyAndAuthMessage, "VERIFY_AND_AUTH_MESSAGE"
addOutputMessageIdMapEntry pMap, RequestPositionsMulti, "REQ_POSITIONS_MULTI"
addOutputMessageIdMapEntry pMap, CancelPositionsMulti, "CANCEL_POSITIONS_MULTI"
addOutputMessageIdMapEntry pMap, RequestAccountDataMulti, "REQ_ACCOUNT_UPDATES_MULTI"
addOutputMessageIdMapEntry pMap, CancelAccountUpdatesMulti, "CANCEL_ACCOUNT_UPDATES_MULTI"
addOutputMessageIdMapEntry pMap, RequestOptionParameters, "REQ_SEC_DEF_OPT_PARAMS"
addOutputMessageIdMapEntry pMap, RequestSoftDollarTiers, "REQ_SOFT_DOLLAR_TIERS"
addOutputMessageIdMapEntry pMap, REQ_FAMILY_CODES, "REQ_FAMILY_CODES"
addOutputMessageIdMapEntry pMap, REQ_MATCHING_SYMBOLS, "REQ_MATCHING_SYMBOLS"
addOutputMessageIdMapEntry pMap, REQ_MKT_DEPTH_EXCHANGES, "REQ_MKT_DEPTH_EXCHANGES"
addOutputMessageIdMapEntry pMap, RequestSmartComponents, "RequestSmartComponents"
addOutputMessageIdMapEntry pMap, RequestNewsArticle, "RequestNewsArticle"
addOutputMessageIdMapEntry pMap, RequestNewsProviders, "RequestNewsProviders"
addOutputMessageIdMapEntry pMap, RequestHistoricalNews, "RequestHistoricalNews"
addOutputMessageIdMapEntry pMap, RequestHeadTimestamp, "RequestHeadTimestamp"
addOutputMessageIdMapEntry pMap, RequestHistogramData, "RequestHistogramData"
addOutputMessageIdMapEntry pMap, CancelHistogramData, "CancelHistogramData"
addOutputMessageIdMapEntry pMap, CancelHeadTimestamp, "CancelHeadTimestamp"
addOutputMessageIdMapEntry pMap, RequestMarketRule, "RequestMarketRule"
addOutputMessageIdMapEntry pMap, RequestPnL, "RequestPnL"
addOutputMessageIdMapEntry pMap, CancelPnL, "CancelPnL"
addOutputMessageIdMapEntry pMap, RequestPnLSingle, "RequestPnLSingle"
addOutputMessageIdMapEntry pMap, CancelPnLSingle, "CancelPnLSingle"
addOutputMessageIdMapEntry pMap, RequestHistoricalTickData, "RequestHistoricalTickData"
addOutputMessageIdMapEntry pMap, RequestTickByTickData, "RequestTickByTickData"
addOutputMessageIdMapEntry pMap, CancelTickByTickData, "CancelTickByTickData"
End Sub


