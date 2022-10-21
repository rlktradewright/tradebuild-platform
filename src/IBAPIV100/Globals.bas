Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                        As String = "IBAPIV100"
Private Const ModuleName                        As String = "Globals"

Public Const InvalidEnumValue                   As String = "*ERR*"
Public Const NullIndex                          As Long = -1

Public Const MaxLong                            As Long = &H7FFFFFFF
Public Const OneMicrosecond                     As Double = 1# / 86400000000#
Public Const OneMinute                          As Double = 1# / 1440#
Public Const OneSecond                          As Double = 1# / 86400#

Public Const NumDaysInWeek                      As Long = 5
Public Const NumDaysInMonth                     As Long = 22
Public Const NumDaysInYear                      As Long = 260
Public Const NumMonthsInYear                    As Long = 12

Public Const TwsLogLevelDetailString            As String = "Detail"
Public Const TwsLogLevelErrorString             As String = "Error"
Public Const TwsLogLevelInformationString       As String = "Information"
Public Const TwsLogLevelSystemString            As String = "System"
Public Const TwsLogLevelWarningString           As String = "Warning"

Public Const BaseMarketDataRequestId            As Long = 0
Public Const BaseMarketDepthRequestId           As Long = &H400000
Public Const BaseScannerRequestId               As Long = &H410000
Public Const BaseHistoricalDataRequestId        As Long = &H420000
Public Const BaseExecutionsRequestId            As Long = &H430000
Public Const BaseContractRequestId              As Long = &H4400000
Public Const BaseOrderId                        As Long = &H10000000

Public Const MaxCallersMarketDataRequestId      As Long = BaseMarketDepthRequestId - 1
Public Const MaxCallersMarketDepthRequestId     As Long = BaseScannerRequestId - BaseMarketDepthRequestId - 1
Public Const MaxCallersScannerRequestId         As Long = BaseHistoricalDataRequestId - BaseScannerRequestId - 1
Public Const MaxCallersHistoricalDataRequestId  As Long = BaseExecutionsRequestId - BaseHistoricalDataRequestId - 1
Public Const MaxCallersExecutionsRequestId      As Long = BaseContractRequestId - BaseExecutionsRequestId - 1
Public Const MaxCallersContractRequestId        As Long = BaseOrderId - BaseContractRequestId - 1
Public Const MaxCallersOrderId                  As Long = &H7FFFFFFF - BaseOrderId - 1

'================================================================================
' Enums
'================================================================================

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

    ' Max for API 973.07

    Max = SCANNER_GENERIC_OPTS
End Enum

Public Enum InternalErrorCodes
    DataIncomplete = vbObjectError + 4327   ' let's hope nothing else uses this number!
End Enum

Public Enum IdTypes
    IdTypeNone
    IdTypeRealtimeData
    IdTypeMarketDepth
    IdTypeHistoricalData
    IdTypeOrder
    IdTypeContractData
    IdTypeExecution
    IdTypeScanner
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
    POSITION = 61
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

    ' Max for API 973.07

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

    ' Max for API 973.07

End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' External function declarations
'================================================================================

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                Destination As Any, _
                source As Any, _
                ByVal length As Long)
                            
Public Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                Destination As Any, _
                source As Any, _
                ByVal length As Long)
                            
'================================================================================
' Private variables
'================================================================================

Private mLogger As FormattingLogger

Private mInputMessageIdMap                              As SortedDictionary

Private mOutputMessageIdMap                             As SortedDictionary

'================================================================================
' Properties
'================================================================================

Public Function gGetCallersRequestIdFromTwsContractRequestId(ByVal pId As Long) As Long
gGetCallersRequestIdFromTwsContractRequestId = pId - BaseContractRequestId
End Function

Public Function gGetCallersRequestIdFromTwsExecutionsRequestId(ByVal pId As Long) As Long
gGetCallersRequestIdFromTwsExecutionsRequestId = pId - BaseExecutionsRequestId
End Function

Public Function gGetCallersRequestIdFromTwsHistRequestId(ByVal pId As Long) As Long
gGetCallersRequestIdFromTwsHistRequestId = pId - BaseHistoricalDataRequestId
End Function

Public Function gGetCallersRequestIdFromTwsMarketDataRequestId(ByVal pId As Long) As Long
gGetCallersRequestIdFromTwsMarketDataRequestId = pId - BaseMarketDataRequestId
End Function

Public Function gGetCallersRequestIdFromTwsMarketDepthRequestId(ByVal pId As Long) As Long
gGetCallersRequestIdFromTwsMarketDepthRequestId = pId - BaseMarketDepthRequestId
End Function

Public Function gGetCallersRequestIdFromTwsScannerRequestId(ByVal pId As Long) As Long
gGetCallersRequestIdFromTwsScannerRequestId = pId - BaseScannerRequestId
End Function

Public Function gGetIdType( _
                ByVal id As Long) As IdTypes
Const ProcName As String = "ggGetIdType"
On Error GoTo Err

If id >= BaseOrderId Then
    gGetIdType = IdTypeOrder
ElseIf id >= BaseContractRequestId Then
    gGetIdType = IdTypeContractData
ElseIf id >= BaseExecutionsRequestId Then
    gGetIdType = IdTypeExecution
ElseIf id >= BaseHistoricalDataRequestId Then
    gGetIdType = IdTypeHistoricalData
ElseIf id >= BaseScannerRequestId Then
    gGetIdType = IdTypeScanner
ElseIf id >= BaseMarketDepthRequestId Then
    gGetIdType = IdTypeMarketDepth
ElseIf id >= 0 Then
    gGetIdType = IdTypeRealtimeData
Else
    gGetIdType = IdTypeNone
End If

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gGetTwsContractRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersContractRequestId, "Max request id is " & MaxCallersContractRequestId
gGetTwsContractRequestIdFromCallersRequestId = pId + BaseContractRequestId
End Function

Public Function gGetTwsExecutionsRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersExecutionsRequestId, "Max request id is " & MaxCallersExecutionsRequestId
gGetTwsExecutionsRequestIdFromCallersRequestId = pId + BaseExecutionsRequestId
End Function

Public Function gGetTwsHistRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersHistoricalDataRequestId, "Max request id is " & MaxCallersHistoricalDataRequestId
gGetTwsHistRequestIdFromCallersRequestId = pId + BaseHistoricalDataRequestId
End Function

Public Function gGetTwsMarketDataRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersMarketDataRequestId, "Max request id is " & MaxCallersMarketDataRequestId
gGetTwsMarketDataRequestIdFromCallersRequestId = pId + BaseMarketDataRequestId
End Function

Public Function gGetTwsMarketDepthRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersMarketDepthRequestId, "Max request id is " & MaxCallersMarketDepthRequestId
gGetTwsMarketDepthRequestIdFromCallersRequestId = pId + BaseMarketDepthRequestId
End Function

Public Function gGetTwsScannerRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersScannerRequestId, "Max request id is " & MaxCallersScannerRequestId
gGetTwsScannerRequestIdFromCallersRequestId = pId + BaseScannerRequestId
End Function

Public Property Get gLogger() As FormattingLogger
If mLogger Is Nothing Then Set mLogger = CreateFormattingLogger("tradebuild.log.ibapi", ProjectName)
Set gLogger = mLogger
End Property

Public Property Get gSocketLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tradebuild.log.ibapi.socket", ProjectName)
Set gSocketLogger = sLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Function gByteBufferToString( _
                ByRef pHeader As String, _
                ByRef pBuffer() As Byte) As String
Const ProcName As String = "gByteBufferToString"
On Error GoTo Err

gByteBufferToString = pHeader & Replace(StrConv(pBuffer, vbUnicode), Chr$(0), "_")

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gContractHasExpired(ByVal pContractSpec As TwsContractSpecifier) As Boolean
If pContractSpec.SecType = TwsSecTypeCash Or _
    pContractSpec.SecType = TwsSecTypeIndex Or _
    pContractSpec.SecType = TwsSecTypeStock _
Then
    gContractHasExpired = False
    Exit Function
End If
    
gContractHasExpired = contractExpiryToDate(pContractSpec) < Now
End Function

Public Function gFormatBuffer( _
                ByRef pBuffer() As Byte, _
                ByVal pBufferNextFreeIndex As Long) As String
Dim s As StringBuilder
Dim i As Long
Dim j As Long

Const ProcName As String = "gFormatBuffer"
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
gFormatBuffer = s.ToString

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gGetApi( _
                ByVal pServer As String, _
                ByVal pPort As String, _
                ByVal pClientId As Long, _
                ByVal pConnectionRetryIntervalSecs As Long, _
                ByVal pLogApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogApiMessageStats As Boolean) As TwsAPI
Const ProcName As String = "gGetApi"
On Error GoTo Err

Set gGetApi = New TwsAPI
gGetApi.Initialise pServer, pPort, pClientId, pLogApiMessages, pLogRawApiMessages, pLogApiMessageStats
gGetApi.ConnectionRetryIntervalSecs = pConnectionRetryIntervalSecs

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

' format is yyyymmdd [hh:mm:ss [timezone]]. There can be more than one space
' between the date, time and timezone parts
Public Function gGetDate( _
                ByVal pDateString As String, _
                Optional ByRef pTimezoneName As String) As Date
Const ProcName As String = "gGetDate"
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

gGetDate = CDate(Left$(lDatePart, 4) & "/" & _
                Mid$(lDatePart, 5, 2) & "/" & _
                Mid$(lDatePart, 7, 2))

If Len(lTimePart) <> 0 Then gGetDate = gGetDate + CDate(lTimePart)

If Not IsMissing(pTimezoneName) Then pTimezoneName = lTimezoneName

Exit Function

Err:
If Err.Number <> VBErrorCodes.VbErrTypeMismatch Then
    gHandleUnexpectedError Nothing, ProcName, ModuleName
End If
gGetDate = CDate(0#)
End Function

Public Function gGetDateFromUnixSystemTime(ByVal pSystemTime As Double) As Date
Const ProcName As String = "gGetDateFromUnixSystemTime"
On Error GoTo Err

gGetDateFromUnixSystemTime = CDate(CDbl((2209161600@ + pSystemTime) / (86400@)))

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByVal pErrorHandler As IProgramErrorListener, _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

If Not pErrorHandler Is Nothing Then
    On Error GoTo HandleError
    ' ensure the calling proc's details are included in the error source
    HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, True, False, errNum, errDesc, errSource
    ' will never get here!
End If

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource

' will never get here!
Exit Sub

HandleError:
Dim ev As ErrorEventData
ev.ErrorCode = Err.Number
errNum = Err.Number

ev.ErrorMessage = Err.Description
errDesc = Err.Description

ev.ErrorSource = Err.source
errSource = Err.source

pErrorHandler.NotifyUnexpectedProgramError ev

' if we get to here, the error handler hasn't re-raised the error, so we will
Err.Raise errNum, errSource, errDesc
End Sub

Public Sub gNotifyUnhandledError( _
                ByVal pErrorHandler As IProgramErrorListener, _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

If Not pErrorHandler Is Nothing Then
    On Error GoTo HandleError
    ' ensure the calling proc's details are included in the error source
    HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, True, False, errNum, errDesc, errSource
End If

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource

' will never get here!
Exit Sub

HandleError:
Dim ev As ErrorEventData
ev.ErrorCode = Err.Number

ev.ErrorMessage = Err.Description

ev.ErrorSource = Err.source

pErrorHandler.NotifyUnhandledProgramError ev

' if we get to here, the error handler hasn't called the unhandled error mechanism, so we will
UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gInputMessageIdToString( _
                ByVal msgId As TwsSocketInMsgTypes) As String
Const ProcName As String = "gInputMessageIdToString"
On Error GoTo Err

Select Case msgId
Case TICK_SIZE
    gInputMessageIdToString = "TickSize"
Case TICK_STRING
    gInputMessageIdToString = "TickString"
Case MARKET_DEPTH
    gInputMessageIdToString = "MarketDepth"
Case TICK_PRICE
    gInputMessageIdToString = "TickPrice"
Case Else
    Dim lName As String
    If mInputMessageIdMap.TryItem(msgId, lName) Then
        gInputMessageIdToString = lName
    Else
        gInputMessageIdToString = "?????"
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Sub gLog(ByRef pMsg As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err


gLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Public Function gLongToNetworkBytes(ByVal pNumber As Long) As Byte()
Dim b(3) As Byte
b(0) = ((pNumber And &HFF000000) / &H1000000) And &HFF&
b(1) = (pNumber And &HFF0000) / &H10000
b(2) = (pNumber And &HFF00&) / &H100&
b(3) = (pNumber And &HFF&)
gLongToNetworkBytes = b
End Function

Public Function gOutputMessageIdToString( _
                ByVal msgId As TwsSocketOutMsgTypes) As String
Const ProcName As String = "gOutputMessageIdToString"
On Error GoTo Err

Dim lName As String
If mOutputMessageIdMap.TryItem(msgId, lName) Then
    gOutputMessageIdToString = lName
Else
    gOutputMessageIdToString = "?????"
End If

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gRoundTimeToSecond( _
                ByVal pTimeStamp As Date) As Date
Const ProcName As String = "gRoundTimeToSecond"
gRoundTimeToSecond = Int((pTimeStamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function gTruncateTimeToNextMinute(ByVal pTimeStamp As Date) As Date
Const ProcName As String = "gTruncateTimeToNextMinute"
On Error GoTo Err

gTruncateTimeToNextMinute = Int((pTimeStamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gTruncateTimeToMinute(ByVal pTimeStamp As Date) As Date
Const ProcName As String = "gTruncateTimeToMinute"
On Error GoTo Err

gTruncateTimeToMinute = Int((pTimeStamp + OneMicrosecond) / OneMinute) * OneMinute

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gTwsHedgeTypeFromString(ByVal pValue As String) As TwsHedgeTypes
Select Case UCase$(pValue)
Case ""
    gTwsHedgeTypeFromString = TwsHedgeTypeNone
Case "D", "DELTA"
    gTwsHedgeTypeFromString = TwsHedgeTypeDelta
Case "B", "BETA"
    gTwsHedgeTypeFromString = TwsHedgeTypeBeta
Case "F"
    gTwsHedgeTypeFromString = TwsHedgeTypeFX
Case "P", "PAIR"
    gTwsHedgeTypeFromString = TwsHedgeTypePair
Case Else
    AssertArgument False, "Value is not a valid Hedge Type"
End Select
End Function

Public Function gTwsHedgeTypeToString(ByVal pValue As TwsHedgeTypes) As String
Select Case pValue
Case TwsHedgeTypeNone
    gTwsHedgeTypeToString = ""
Case TwsHedgeTypeDelta
    gTwsHedgeTypeToString = "D"
Case TwsHedgeTypeBeta
    gTwsHedgeTypeToString = "B"
Case TwsHedgeTypeFX
    gTwsHedgeTypeToString = "F"
Case TwsHedgeTypePair
    gTwsHedgeTypeToString = "P"
Case Else
    AssertArgument False, "Value is not a valid Hedge Type"
End Select
End Function

Public Function gTwsOptionRightFromString(ByVal Value As String) As TwsOptionRights
Select Case UCase$(Value)
Case "", "?", "0"
    gTwsOptionRightFromString = TwsOptRightNone
Case "CALL", "C"
    gTwsOptionRightFromString = TwsOptRightCall
Case "PUT", "P"
    gTwsOptionRightFromString = TwsOptRightPut
Case Else
    AssertArgument False, "Value is not a valid Option Right"
End Select
End Function

Public Function gTwsOptionRightToString(ByVal Value As TwsOptionRights) As String
Select Case Value
Case TwsOptRightNone
    gTwsOptionRightToString = ""
Case TwsOptRightCall
    gTwsOptionRightToString = "Call"
Case TwsOptRightPut
    gTwsOptionRightToString = "Put"
Case Else
    gTwsOptionRightToString = InvalidEnumValue
End Select
End Function

Public Function gTwsOrderActionFromString(ByVal Value As String) As TwsOrderActions
Select Case UCase$(Value)
Case ""
    gTwsOrderActionFromString = TwsOrderActionNone
Case "BUY"
    gTwsOrderActionFromString = TwsOrderActionBuy
Case "SELL"
    gTwsOrderActionFromString = TwsOrderActionSell
Case "SSHORT"
    gTwsOrderActionFromString = TwsOrderActionSellShort
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function

Public Function gTwsOrderActionToString(ByVal Value As TwsOrderActions) As String
Select Case Value
Case TwsOrderActionBuy
    gTwsOrderActionToString = "BUY"
Case TwsOrderActionSell
    gTwsOrderActionToString = "SELL"
Case TwsOrderActionSellShort
    gTwsOrderActionToString = "SSHORT"
Case TwsOrderActionNone
    gTwsOrderActionToString = ""
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function

Public Function gTwsOrderTIFFromString(ByVal Value As String) As TwsOrderTIFs
Select Case UCase$(Value)
Case ""
    gTwsOrderTIFFromString = TwsOrderTIFNone
Case "DAY"
    gTwsOrderTIFFromString = TwsOrderTIFDay
Case "GTC"
    gTwsOrderTIFFromString = TwsOrderTIFGoodTillCancelled
Case "IOC"
    gTwsOrderTIFFromString = TwsOrderTIFImmediateOrCancel
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Tws Order TIF"
End Select
End Function

Public Function gTwsOrderTIFToString(ByVal Value As TwsOrderTIFs) As String
Select Case Value
Case TwsOrderTIFs.TwsOrderTIFDay
    gTwsOrderTIFToString = "DAY"
Case TwsOrderTIFs.TwsOrderTIFGoodTillCancelled
    gTwsOrderTIFToString = "GTC"
Case TwsOrderTIFs.TwsOrderTIFImmediateOrCancel
    gTwsOrderTIFToString = "IOC"
Case TwsOrderTIFs.TwsOrderTIFNone
    gTwsOrderTIFToString = ""
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Tws Order TIF"
End Select
End Function

Public Function gTwsOrderTypeFromString(ByVal Value As String) As TwsOrderTypes
Select Case UCase$(Value)
Case "MKT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarket
Case "MKTCLS"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case "LMT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimit
Case "LMTCLS"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case "PEGMKT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToMarket
Case "STP"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStop
Case "STPLMT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStopLimit
Case "STP LMT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStopLimit
Case "TRAIL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeTrail
Case "REL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeRelative
Case "VWAP"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeVWAP
Case "MTL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketToLimit
Case "RFQ"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeQuote
Case "ADJUST"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeAdjust
Case "ALERT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeAlert
Case "LIT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitIfTouched
Case "MIT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketIfTouched
Case "TRAIL LIMIT", "TRAILLMT"
    ' Note that we get both spellings in the API, eg "TRAIL LIMIT" is OrderStatus
    ' and "TRAILLMT" in permitted order types
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeTrailLimit
Case "MKTPROT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketWithProtection
Case "MOO"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnOpen
Case "MOC"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case "LOO"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnOpen
Case "LOC"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case "PEGPRI"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToPrimary
Case "VOL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeVol
Case "PEG BENCH"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToBenchmark
Case Else
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeNone
End Select
End Function

Public Function gTwsOrderTypeToString(ByVal Value As TwsOrderTypes) As String
Const ProcName As String = "gTwsOrderTypeToString"
On Error GoTo Err

Select Case Value
Case TwsOrderTypes.TwsOrderTypeNone
    gTwsOrderTypeToString = ""
Case TwsOrderTypes.TwsOrderTypeMarket
    gTwsOrderTypeToString = "MKT"
Case TwsOrderTypes.TwsOrderTypeMarketOnClose
    gTwsOrderTypeToString = "MOC"
Case TwsOrderTypes.TwsOrderTypeLimit
    gTwsOrderTypeToString = "LMT"
Case TwsOrderTypes.TwsOrderTypeLimitOnClose
    gTwsOrderTypeToString = "LOC"
Case TwsOrderTypes.TwsOrderTypePeggedToMarket
    gTwsOrderTypeToString = "PEG MKT"
Case TwsOrderTypes.TwsOrderTypeStop
    gTwsOrderTypeToString = "STP"
Case TwsOrderTypes.TwsOrderTypeStopLimit
    gTwsOrderTypeToString = "STPLMT"
Case TwsOrderTypes.TwsOrderTypeTrail
    gTwsOrderTypeToString = "TRAIL"
Case TwsOrderTypes.TwsOrderTypeRelative
    gTwsOrderTypeToString = "REL"
Case TwsOrderTypes.TwsOrderTypeVWAP
    gTwsOrderTypeToString = "VWAP"
Case TwsOrderTypes.TwsOrderTypeMarketToLimit
    gTwsOrderTypeToString = "MTL"
Case TwsOrderTypes.TwsOrderTypeQuote
    gTwsOrderTypeToString = "QUOTE"
Case TwsOrderTypes.TwsOrderTypeAdjust
    gTwsOrderTypeToString = "ADJUST"
Case TwsOrderTypes.TwsOrderTypeAlert
    gTwsOrderTypeToString = "ALERT"
Case TwsOrderTypes.TwsOrderTypeLimitIfTouched
    gTwsOrderTypeToString = "LIT"
Case TwsOrderTypes.TwsOrderTypeMarketIfTouched
    gTwsOrderTypeToString = "MIT"
Case TwsOrderTypes.TwsOrderTypeTrailLimit
    gTwsOrderTypeToString = "TRAIL LIMIT"
Case TwsOrderTypes.TwsOrderTypeMarketWithProtection
    gTwsOrderTypeToString = "MKTPROT"
Case TwsOrderTypes.TwsOrderTypeMarketOnOpen
    gTwsOrderTypeToString = "MOO"
Case TwsOrderTypes.TwsOrderTypeLimitOnOpen
    gTwsOrderTypeToString = "LOO"
Case TwsOrderTypes.TwsOrderTypePeggedToPrimary
    gTwsOrderTypeToString = "PEGPRI"
Case TwsOrderTypes.TwsOrderTypeVol
    gTwsOrderTypeToString = "VOL"
Case TwsOrderTypes.TwsOrderTypePeggedToBenchmark
    gTwsOrderTypeToString = "PEG BENCH"
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException
End Select

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gTwsSecTypeFromString(ByVal Value As String) As TwsSecTypes
Select Case UCase$(Value)
Case ""
    gTwsSecTypeFromString = TwsSecTypeNone
Case "STOCK", "STK"
    gTwsSecTypeFromString = TwsSecTypeStock
Case "FUTURE", "FUT"
    gTwsSecTypeFromString = TwsSecTypeFuture
Case "OPTION", "OPT"
    gTwsSecTypeFromString = TwsSecTypeOption
Case "FUTURES OPTION", "FOP"
    gTwsSecTypeFromString = TwsSecTypeFuturesOption
Case "CASH"
    gTwsSecTypeFromString = TwsSecTypeCash
Case "BAG", "COMBO", "CMB"
    gTwsSecTypeFromString = TwsSecTypeCombo
Case "INDEX", "IND"
    gTwsSecTypeFromString = TwsSecTypeIndex
Case "WARRANT", "WAR"
    gTwsSecTypeFromString = TwsSecTypeWarrant
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Security Type"
End Select
End Function

Public Function gTwsSecTypeToString(ByVal Value As TwsSecTypes) As String
Select Case Value
Case TwsSecTypeStock
    gTwsSecTypeToString = "Stock"
Case TwsSecTypeFuture
    gTwsSecTypeToString = "Future"
Case TwsSecTypeOption
    gTwsSecTypeToString = "Option"
Case TwsSecTypeFuturesOption
    gTwsSecTypeToString = "Futures Option"
Case TwsSecTypeCash
    gTwsSecTypeToString = "Cash"
Case TwsSecTypeCombo
    gTwsSecTypeToString = "Combo"
Case TwsSecTypeIndex
    gTwsSecTypeToString = "Index"
Case TwsSecTypeNone
    gTwsSecTypeToString = ""
Case TwsSecTypeWarrant
    gTwsSecTypeToString = "Warrant"
Case Else
    gTwsSecTypeToString = InvalidEnumValue
End Select
End Function

Public Function gTwsSecTypeToShortString(ByVal Value As TwsSecTypes) As String
Select Case Value
Case TwsSecTypeStock
    gTwsSecTypeToShortString = "STK"
Case TwsSecTypeFuture
    gTwsSecTypeToShortString = "FUT"
Case TwsSecTypeOption
    gTwsSecTypeToShortString = "OPT"
Case TwsSecTypeFuturesOption
    gTwsSecTypeToShortString = "FOP"
Case TwsSecTypeCash
    gTwsSecTypeToShortString = "CASH"
Case TwsSecTypeCombo
    gTwsSecTypeToShortString = "BAG"
Case TwsSecTypeIndex
    gTwsSecTypeToShortString = "IND"
Case TwsSecTypeNone
    gTwsSecTypeToShortString = ""
Case TwsSecTypeWarrant
    gTwsSecTypeToShortString = "WAR"
Case Else
    gTwsSecTypeToShortString = InvalidEnumValue
End Select
End Function

Public Function gTwsShortSaleSlotFromString( _
                ByVal Value As String) As TwsShortSaleSlotCodes
Select Case UCase$(Value)
Case "", "N/A", "NOT APPLICABLE"
    gTwsShortSaleSlotFromString = TwsShortSaleSlotNotApplicable
Case "B", "CB", "BROKER", "CLEARING BROKER"
    gTwsShortSaleSlotFromString = TwsShortSaleSlotClearingBroker
Case "T", "TP", "THIRD PARTY"
    gTwsShortSaleSlotFromString = TwsShortSaleSlotThirdParty
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Short Sale Slot"
End Select
End Function

Public Function gTwsShortSaleSlotToString( _
                ByVal Value As TwsShortSaleSlotCodes) As String
Select Case Value
Case TwsShortSaleSlotNotApplicable
    gTwsShortSaleSlotToString = "N/A"
Case TwsShortSaleSlotClearingBroker
    gTwsShortSaleSlotToString = "BROKER"
Case TwsShortSaleSlotThirdParty
    gTwsShortSaleSlotToString = "THIRD PARTY"
Case Else
    gTwsShortSaleSlotToString = InvalidEnumValue
End Select
End Function

Public Function gTwsStopTriggerMethodFromString(ByVal Value As String) As TwsStopTriggerMethods
Select Case Value
Case "Default"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDefault
Case "Double Bid/Ask"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk
Case "Last"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerLast
Case "Double Last"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDoubleLast
Case "Bid/Ask"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerBidAsk
Case "Last or Bid/Ask"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk
Case "Midpoint"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerMidPoint
End Select
End Function

Public Function gTwsStopTriggerMethodToString(ByVal Value As TwsStopTriggerMethods) As String
Select Case Value
Case TwsStopTriggerMethods.TwsStopTriggerBidAsk
    gTwsStopTriggerMethodToString = "Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerDefault
    gTwsStopTriggerMethodToString = "Default"
Case TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk
    gTwsStopTriggerMethodToString = "Double Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerDoubleLast
    gTwsStopTriggerMethodToString = "Double Last"
Case TwsStopTriggerMethods.TwsStopTriggerLast
    gTwsStopTriggerMethodToString = "Last"
Case TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk
    gTwsStopTriggerMethodToString = "Last or Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerMidPoint
    gTwsStopTriggerMethodToString = "Midpoint"
End Select
End Function

Public Sub gWSAIoctlCompletionRoutine( _
                ByVal dwError As Long, _
                ByVal cbTransferred As Long, _
                ByVal lpOverlapped As Long, _
                ByVal dwFlags As Long)
                
End Sub

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

Set mInputMessageIdMap = CreateSortedDictionary(KeyTypeInteger)
setupInputMessageIdMap

Set mOutputMessageIdMap = CreateSortedDictionary(KeyTypeInteger)
setupOutputMessageIdMap

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub


'================================================================================
' Helper Functions
'================================================================================

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
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub addInputMessageIdMapEntry( _
                ByVal pMessageId As TwsSocketInMsgTypes, _
                ByVal pMessageName As String)
mInputMessageIdMap.Add pMessageName, pMessageId
End Sub

Private Sub addOutputMessageIdMapEntry( _
                ByVal pMessageId As TwsSocketOutMsgTypes, _
                ByVal pMessageName As String)
mOutputMessageIdMap.Add pMessageName, pMessageId
End Sub

Private Sub setupInputMessageIdMap()
addInputMessageIdMapEntry TICK_PRICE, "TICKPRICE"
addInputMessageIdMapEntry TICK_SIZE, "TICKSIZE"
addInputMessageIdMapEntry ORDER_STATUS, "ORDERSTATUS"
addInputMessageIdMapEntry ERR_MSG, "ERRORMESSAGE"
addInputMessageIdMapEntry OPEN_ORDER, "OPENORDER"
addInputMessageIdMapEntry ACCT_VALUE, "ACCOUNTVALUE"
addInputMessageIdMapEntry PORTFOLIO_VALUE, "PORTFOLIOVALUE"
addInputMessageIdMapEntry ACCT_UPDATE_TIME, "ACCOUNTUPDATETIME"
addInputMessageIdMapEntry NEXT_VALID_ID, "NEXTVALIDID"
addInputMessageIdMapEntry CONTRACT_DATA, "CONTRACTDATA"
addInputMessageIdMapEntry EXECUTION_DATA, "EXECUTIONDATA"
addInputMessageIdMapEntry MARKET_DEPTH, "MARKETDEPTH"
addInputMessageIdMapEntry MARKET_DEPTH_L2, "MARKETDEPTHL2"
addInputMessageIdMapEntry NEWS_BULLETINS, "NEWBULLETIN"
addInputMessageIdMapEntry MANAGED_ACCTS, "MANAGEDACCOUNTS"
addInputMessageIdMapEntry RECEIVE_FA, "RECEIVEFA"
addInputMessageIdMapEntry HISTORICAL_DATA, "HISTORICALDATA"
addInputMessageIdMapEntry BOND_CONTRACT_DATA, "BONDCONTRACTDATA"
addInputMessageIdMapEntry SCANNER_PARAMETERS, "SCANNERPARAMETERS"
addInputMessageIdMapEntry SCANNER_DATA, "SCANNERDATA"
addInputMessageIdMapEntry TICK_OPTION_COMPUTATION, "OPTIONCOMPUTATION"
addInputMessageIdMapEntry TICK_GENERIC, "GENERIC"
addInputMessageIdMapEntry TICK_STRING, "STRING"
addInputMessageIdMapEntry TICK_EFP, "EFP"
addInputMessageIdMapEntry CURRENT_TIME, "CURRENTTIME"
addInputMessageIdMapEntry REAL_TIME_BARS, "REALTIMEBAR"
addInputMessageIdMapEntry FUNDAMENTAL_DATA, "FUNDAMENTALDATA"
addInputMessageIdMapEntry CONTRACT_DATA_END, "CONTRACTDATAEND"
addInputMessageIdMapEntry OPEN_ORDER_END, "OPENORDEREND"
addInputMessageIdMapEntry ACCT_DOWNLOAD_END, "ACCOUNTDOWNLOADEND"
addInputMessageIdMapEntry EXECUTION_DATA_END, "EXECUTIONDATAEND"
addInputMessageIdMapEntry DELTA_NEUTRAL_VALIDATION, "DELTANEUTRALVALIDN"
addInputMessageIdMapEntry TICK_SNAPSHOT_END, "TICKSNAPSHOTEND"
addInputMessageIdMapEntry MARKET_DATA_TYPE, "MARKETDATATYPE"
addInputMessageIdMapEntry COMMISSION_REPORT, "COMMISSIONREPORT"
addInputMessageIdMapEntry POSITION, "POSITION"
addInputMessageIdMapEntry POSITION_END, "POSITIONEND"
addInputMessageIdMapEntry ACCOUNT_SUMMARY, "ACCOUNTSUMMARY"
addInputMessageIdMapEntry ACCOUNT_SUMMARY_END, "ACCOUNTSUMMARYEND"
addInputMessageIdMapEntry VERIFYMESSAGEAPI, "VERIFYMESSAGEAPI"
addInputMessageIdMapEntry VERIFYCOMPLETED, "VERIFYCOMPLETED"
addInputMessageIdMapEntry DISPLAYGROUPLIST, "DISPLAYGROUPLIST"
addInputMessageIdMapEntry DISPLAYGROUPUPDATED, "DISPLAYGROUPUPD"
addInputMessageIdMapEntry VERIFYANDAUTHMESSAGEAPI, "VERIFY/AUTHMESSAGEAPI"
addInputMessageIdMapEntry VERIFYANDAUTHCOMPLETED, "VERIFY/AUTHCOMPLETED"
addInputMessageIdMapEntry POSITIONMULTI, "POSITIONMULTI"
addInputMessageIdMapEntry POSITIONMULTIEND, "POSITIONMULTIEND"
addInputMessageIdMapEntry ACCOUNTUPDATEMULTI, "ACCOUNTUPDATEMULTI"
addInputMessageIdMapEntry ACCOUNTUPDATEMULTIEND, "ACCOUNTUPDATEMULTIEND"
addInputMessageIdMapEntry OptionParameter, "OPTIONPARAMETER"
addInputMessageIdMapEntry OptionParameterEnd, "OPTIONPARAMETEREND"
addInputMessageIdMapEntry SOFTDOLLARTIERS, "SOFTDOLLARTIERS"
addInputMessageIdMapEntry FAMILYCODES, "FAMILYCODES"
addInputMessageIdMapEntry SYMBOLSAMPLES, "SYMBOLSAMPLES"
addInputMessageIdMapEntry MARKETDEPTHEXCHANGES, "MARKETDEPTHEXCHANGES"
addInputMessageIdMapEntry TickRequestParams, "TICKREQUESTPARAMS"
addInputMessageIdMapEntry SMARTCOMPONENTS, "SMARTCOMPONENTS"
addInputMessageIdMapEntry NEWSARTICLE, "NEWSARTICLE"
addInputMessageIdMapEntry TICKNEWS, "TICKNEWS"
addInputMessageIdMapEntry NEWSPROVIDERS, "NEWSPROVIDERS"
addInputMessageIdMapEntry HISTORICALNEWS, "HISTORICALNEWS"
addInputMessageIdMapEntry HISTORICALNEWSEND, "HISTORICALNEWSEND"
addInputMessageIdMapEntry HEADTIMESTAMP, "HEADTIMESTAMP"
addInputMessageIdMapEntry HISTOGRAMDATA, "HISTOGRAMDATA"
addInputMessageIdMapEntry HISTORICALDATAUPDATE, "HISTORICALDATAUPDATE"
addInputMessageIdMapEntry REROUTEMARKETDATA, "REROUTEMARKETDATA"
addInputMessageIdMapEntry REROUTEMARKETDEPTH, "REROUTEMARKETDEPTH"
addInputMessageIdMapEntry MarketRule, "MARKETRULE"
addInputMessageIdMapEntry TwsSocketInMsgTypes.PNL, "PNL"
addInputMessageIdMapEntry PNLSINGLE, "PNLSINGLE"
addInputMessageIdMapEntry HISTORICALTICKMIDPOINT, "HISTORICALTICKMIDPOINT"
addInputMessageIdMapEntry HISTORICALTICKBIDASK, "HISTORICALTICKBIDASK"
addInputMessageIdMapEntry HISTORICALTICKLAST, "HISTORICALTICKLAST"
addInputMessageIdMapEntry TICKBYTICK, "TICKBYTICK"
End Sub

Private Sub setupOutputMessageIdMap()
addOutputMessageIdMapEntry REQ_MKT_DATA, "REQ_MKT_DATA"
addOutputMessageIdMapEntry CANCEL_MKT_DATA, "CANCEL_MKT_DATA"
addOutputMessageIdMapEntry PLACE_ORDER, "PLACE_ORDER"
addOutputMessageIdMapEntry CANCEL_ORDER, "CANCEL_ORDER"
addOutputMessageIdMapEntry REQ_OPEN_ORDERS, "REQ_OPEN_ORDERS"
addOutputMessageIdMapEntry REQ_ACCT_DATA, "REQ_ACCT_DATA"
addOutputMessageIdMapEntry REQ_EXECUTIONS, "REQ_EXECUTIONS"
addOutputMessageIdMapEntry REQ_IDS, "REQ_IDS"
addOutputMessageIdMapEntry REQ_CONTRACT_DATA, "REQ_CONTRACT_DATA"
addOutputMessageIdMapEntry REQ_MKT_DEPTH, "REQ_MKT_DEPTH"
addOutputMessageIdMapEntry CANCEL_MKT_DEPTH, "CANCEL_MKT_DEPTH"
addOutputMessageIdMapEntry REQ_NEWS_BULLETINS, "REQ_NEWS_BULLETINS"
addOutputMessageIdMapEntry CANCEL_NEWS_BULLETINS, "CANCEL_NEWS_BULLETINS"
addOutputMessageIdMapEntry SET_SERVER_LOGLEVEL, "SET_SERVER_LOGLEVEL"
addOutputMessageIdMapEntry REQ_AUTO_OPEN_ORDERS, "REQ_AUTO_OPEN_ORDERS"
addOutputMessageIdMapEntry REQ_ALL_OPEN_ORDERS, "REQ_ALL_OPEN_ORDERS"
addOutputMessageIdMapEntry REQ_MANAGED_ACCTS, "REQ_MANAGED_ACCTS"
addOutputMessageIdMapEntry REQ_FA, "REQ_FA"
addOutputMessageIdMapEntry REPLACE_FA, "REPLACE_FA"
addOutputMessageIdMapEntry REQ_HISTORICAL_DATA, "REQ_HISTORICAL_DATA"
addOutputMessageIdMapEntry EXERCISE_OPTIONS, "EXERCISE_OPTIONS"
addOutputMessageIdMapEntry REQ_SCANNER_SUBSCRIPTION, "REQ_SCANNER_SUBSCRIPTION"
addOutputMessageIdMapEntry CANCEL_SCANNER_SUBSCRIPTION, "CANCEL_SCANNER_SUBSCRIPTION"
addOutputMessageIdMapEntry REQ_SCANNER_PARAMETERS, "REQ_SCANNER_PARAMETERS"
addOutputMessageIdMapEntry CANCEL_HISTORICAL_DATA, "CANCEL_HISTORICAL_DATA"
addOutputMessageIdMapEntry REQ_CURRENT_TIME, "REQ_CURRENT_TIME"
addOutputMessageIdMapEntry REQ_REAL_TIME_BARS, "REQ_REAL_TIME_BARS"
addOutputMessageIdMapEntry CANCEL_REAL_TIME_BARS, "CANCEL_REAL_TIME_BARS"
addOutputMessageIdMapEntry REQ_FUNDAMENTAL_DATA, "REQ_FUNDAMENTAL_DATA"
addOutputMessageIdMapEntry CANCEL_FUNDAMENTAL_DATA, "CANCEL_FUNDAMENTAL_DATA"
addOutputMessageIdMapEntry REQ_CALC_IMPLIED_VOLAT, "REQ_CALC_IMPLIED_VOLAT"
addOutputMessageIdMapEntry REQ_CALC_OPTION_PRICE, "REQ_CALC_OPTION_PRICE"
addOutputMessageIdMapEntry CANCEL_CALC_IMPLIED_VOLAT, "CANCEL_CALC_IMPLIED_VOLAT"
addOutputMessageIdMapEntry CANCEL_CALC_OPTION_PRICE, "CANCEL_CALC_OPTION_PRICE"
addOutputMessageIdMapEntry REQ_GLOBAL_CANCEL, "REQ_GLOBAL_CANCEL"
addOutputMessageIdMapEntry REQ_MARKET_DATA_TYPE, "REQ_MARKET_DATA_TYPE"
addOutputMessageIdMapEntry REQ_POSITIONS, "REQ_POSITIONS"
addOutputMessageIdMapEntry REQ_ACCOUNT_SUMMARY, "REQ_ACCOUNT_SUMMARY"
addOutputMessageIdMapEntry CANCEL_ACCOUNT_SUMMARY, "CANCEL_ACCOUNT_SUMMARY"
addOutputMessageIdMapEntry CANCEL_POSITIONS, "CANCEL_POSITIONS"
addOutputMessageIdMapEntry VerifyRequest, "VERIFY_REQUEST"
addOutputMessageIdMapEntry VerifyMessage, "VERIFY_MESSAGE"
addOutputMessageIdMapEntry QueryDisplayGroups, "QUERY_DISPLAY_GROUPS"
addOutputMessageIdMapEntry SubscribeToGroupEvents, "SUBSCRIBE_TO_GROUP_EVENTS"
addOutputMessageIdMapEntry UpdateDisplayGroup, "UPDATE_DISPLAY_GROUP"
addOutputMessageIdMapEntry UnsubscribeFromGroupEvents, "UNSUBSCRIBE_FROM_GROUP_EVENTS"
addOutputMessageIdMapEntry StartAPI, "START_API"
addOutputMessageIdMapEntry VerifyAndAuthRequest, "VERIFY_AND_AUTH_REQ"
addOutputMessageIdMapEntry VerifyAndAuthMessage, "VERIFY_AND_AUTH_MESSAGE"
addOutputMessageIdMapEntry RequestPositionsMulti, "REQ_POSITIONS_MULTI"
addOutputMessageIdMapEntry CancelPositionsMulti, "CANCEL_POSITIONS_MULTI"
addOutputMessageIdMapEntry RequestAccountDataMulti, "REQ_ACCOUNT_UPDATES_MULTI"
addOutputMessageIdMapEntry CancelAccountUpdatesMulti, "CANCEL_ACCOUNT_UPDATES_MULTI"
addOutputMessageIdMapEntry RequestOptionParameters, "REQ_SEC_DEF_OPT_PARAMS"
addOutputMessageIdMapEntry RequestSoftDollarTiers, "REQ_SOFT_DOLLAR_TIERS"
addOutputMessageIdMapEntry REQ_FAMILY_CODES, "REQ_FAMILY_CODES"
addOutputMessageIdMapEntry REQ_MATCHING_SYMBOLS, "REQ_MATCHING_SYMBOLS"
addOutputMessageIdMapEntry REQ_MKT_DEPTH_EXCHANGES, "REQ_MKT_DEPTH_EXCHANGES"
addOutputMessageIdMapEntry RequestSmartComponents, "RequestSmartComponents"
addOutputMessageIdMapEntry RequestNewsArticle, "RequestNewsArticle"
addOutputMessageIdMapEntry RequestNewsProviders, "RequestNewsProviders"
addOutputMessageIdMapEntry RequestHistoricalNews, "RequestHistoricalNews"
addOutputMessageIdMapEntry RequestHeadTimestamp, "RequestHeadTimestamp"
addOutputMessageIdMapEntry RequestHistogramData, "RequestHistogramData"
addOutputMessageIdMapEntry CancelHistogramData, "CancelHistogramData"
addOutputMessageIdMapEntry CancelHeadTimestamp, "CancelHeadTimestamp"
addOutputMessageIdMapEntry RequestMarketRule, "RequestMarketRule"
addOutputMessageIdMapEntry RequestPnL, "RequestPnL"
addOutputMessageIdMapEntry CancelPnL, "CancelPnL"
addOutputMessageIdMapEntry RequestPnLSingle, "RequestPnLSingle"
addOutputMessageIdMapEntry CancelPnLSingle, "CancelPnLSingle"
addOutputMessageIdMapEntry RequestHistoricalTickData, "RequestHistoricalTickData"
addOutputMessageIdMapEntry RequestTickByTickData, "RequestTickByTickData"
addOutputMessageIdMapEntry CancelTickByTickData, "CancelTickByTickData"
End Sub
