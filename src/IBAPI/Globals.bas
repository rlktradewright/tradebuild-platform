Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                        As String = "IBAPI970"
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
Public Const BaseHistoricalDataRequestId        As Long = &H800000
Public Const BaseExecutionsRequestId            As Long = &H810000
Public Const BaseContractRequestId              As Long = &H1000000
Public Const BaseOrderId                        As Long = &H10000000

Public Const MaxCallersMarketDataRequestId      As Long = BaseMarketDepthRequestId - 1
Public Const MaxCallersMarketDepthRequestId     As Long = BaseHistoricalDataRequestId - BaseMarketDepthRequestId - 1
Public Const MaxCallersHistoricalDataRequestId  As Long = BaseExecutionsRequestId - BaseHistoricalDataRequestId - 1
Public Const MaxCallersExecutionsRequestId      As Long = BaseContractRequestId - BaseExecutionsRequestId - 1
Public Const MaxCallersContractRequestId        As Long = BaseOrderId - BaseContractRequestId - 1
Public Const MaxCallersOrderId                  As Long = &H7FFFFFFF - BaseOrderId - 1

'Public Const MIN_SERVER_VER_REAL_TIME_BARS              As Long = 34
'Public Const MIN_SERVER_VER_SCALE_ORDERS                As Long = 35
'Public Const MIN_SERVER_VER_SNAPSHOT_MKT_DATA           As Long = 35
'Public Const MIN_SERVER_VER_SSHORT_COMBO_LEGS           As Long = 35
'Public Const MIN_SERVER_VER_WHAT_IF_ORDERS              As Long = 36
'Public Const MIN_SERVER_VER_CONTRACT_CONID              As Long = 37
'Public Const MIN_SERVER_VER_PTA_ORDERS                  As Long = 39
'Public Const MIN_SERVER_VER_FUNDAMENTAL_DATA            As Long = 40
'Public Const MIN_SERVER_VER_UNDER_COMP                  As Long = 40
'Public Const MIN_SERVER_VER_CONTRACT_DATA_CHAIN         As Long = 40
'Public Const MIN_SERVER_VER_SCALE_ORDERS2               As Long = 40
'Public Const MIN_SERVER_VER_ALGO_ORDERS                 As Long = 41
'Public Const MIN_SERVER_VER_EXECUTION_DATA_CHAIN        As Long = 42
'Public Const MIN_SERVER_VER_NOT_HELD                    As Long = 44
'Public Const MIN_SERVER_VER_SEC_ID_TYPE                 As Long = 45
'Public Const MIN_SERVER_VER_PLACE_ORDER_CONID           As Long = 46
'Public Const MIN_SERVER_VER_REQ_MKT_DATA_CONID          As Long = 47
'Public Const MIN_SERVER_VER_REQ_CALC_IMPLIED_VOLAT      As Long = 49
'Public Const MIN_SERVER_VER_REQ_CALC_OPTION_PRICE       As Long = 50
'Public Const MIN_SERVER_VER_CANCEL_CALC_IMPLIED_VOLAT   As Long = 50
'Public Const MIN_SERVER_VER_CANCEL_CALC_OPTION_PRICE    As Long = 50
'Public Const MIN_SERVER_VER_SSHORTX_OLD                 As Long = 51
'Public Const MIN_SERVER_VER_SSHORTX                     As Long = 52
'Public Const MIN_SERVER_VER_REQ_GLOBAL_CANCEL           As Long = 53
'Public Const MIN_SERVER_VER_HEDGE_ORDERS                As Long = 54
'Public Const MIN_SERVER_VER_REQ_MARKET_DATA_TYPE        As Long = 55
'Public Const MIN_SERVER_VER_OPT_OUT_SMART_ROUTING       As Long = 56
'Public Const MIN_SERVER_VER_SMART_COMBO_ROUTING_PARAMS  As Long = 57
'Public Const MIN_SERVER_VER_DELTA_NEUTRAL_CONID         As Long = 58
Public Const MIN_SERVER_VER_SCALE_ORDERS3               As Long = 60
Public Const MIN_SERVER_VER_ORDER_COMBO_LEGS_PRICE      As Long = 61
Public Const MIN_SERVER_VER_TRAILING_PERCENT            As Long = 62
Public Const MIN_SERVER_VER_DELTA_NEUTRAL_OPEN_CLOSE    As Long = 66
Public Const MIN_SERVER_VER_ACCT_SUMMARY                As Long = 67
Public Const MIN_SERVER_VER_TRADING_CLASS               As Long = 68
Public Const MIN_SERVER_VER_SCALE_TABLE                 As Long = 69

'================================================================================
' Enums
'================================================================================

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

'================================================================================
' Properties
'================================================================================

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

Public Function gContractHasExpired(ByVal pContract As TwsContract) As Boolean
If pContract.Sectype = TwsSecTypeCash Or _
    pContract.Sectype = TwsSecTypeIndex Or _
    pContract.Sectype = TwsSecTypeStock _
Then
    gContractHasExpired = False
    Exit Function
End If
    
gContractHasExpired = contractExpiryToDate(pContract) < Now
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
                ByVal pConnectionRetryIntervalSecs As Long) As TwsAPI
Const ProcName As String = "gGetApi"
On Error GoTo Err

Set gGetApi = New TwsAPI
gGetApi.Initialise pServer, pPort, pClientId
gGetApi.ConnectionRetryIntervalSecs = pConnectionRetryIntervalSecs

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

' format is yyyymmdd [hh:mm:ss [timezone]]
Public Function gGetDate( _
                ByVal pDateString As String, _
                Optional ByRef pTimezoneName As String) As Date
Const ProcName As String = "gGetDate"
On Error GoTo Err

If pDateString = "" Then Exit Function

If Len(pDateString) = 8 Then
    gGetDate = CDate(Left$(pDateString, 4) & "/" & _
                    Mid$(pDateString, 5, 2) & "/" & _
                    Mid$(pDateString, 7, 2))
ElseIf Len(pDateString) >= 17 Then
    gGetDate = CDate(Left$(pDateString, 4) & "/" & _
                        Mid$(pDateString, 5, 2) & "/" & _
                        Mid$(pDateString, 7, 2) & " " & _
                        Mid$(pDateString, 11, 8))
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid date string format"
End If

If IsMissing(pTimezoneName) Then Exit Function

If Len(pDateString) > 17 Then
    pTimezoneName = Trim$(Right$(pDateString, Len(pDateString) - 17))
Else
    pTimezoneName = ""
End If

Exit Function

Err:
If Err.Number <> VBErrorCodes.VbErrTypeMismatch And _
    Err.Number <> ErrorCodes.ErrIllegalArgumentException Then
    gHandleUnexpectedError Nothing, ProcName, ModuleName
End If
End Function

Public Function gGetDateFromUnixSystemTime(ByVal pSystemTime As Double) As Date
Const ProcName As String = "gGetDateFromUnixSystemTime"
On Error GoTo Err

gGetDateFromUnixSystemTime = CDate((2209161600000@ + pSystemTime) / (86400000@))

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
Case TICK_PRICE
    gInputMessageIdToString = "TICK_PRICE"
Case TICK_SIZE
    gInputMessageIdToString = "TICK_SIZE"
Case ORDER_STATUS
    gInputMessageIdToString = "ORDER_STATUS"
Case ERR_MSG
    gInputMessageIdToString = "ERR_MSG"
Case OPEN_ORDER
    gInputMessageIdToString = "OPEN_ORDER"
Case ACCT_VALUE
    gInputMessageIdToString = "ACCT_Value"
Case PORTFOLIO_VALUE
    gInputMessageIdToString = "PORTFOLIO_Value"
Case ACCT_UPDATE_TIME
    gInputMessageIdToString = "ACCT_UPDATE_TIME"
Case NEXT_VALID_ID
    gInputMessageIdToString = "NEXT_VALID_ID"
Case CONTRACT_DATA
    gInputMessageIdToString = "CONTRACT_DATA"
Case EXECUTION_DATA
    gInputMessageIdToString = "EXECUTION_DATA"
Case MARKET_DEPTH
    gInputMessageIdToString = "MARKET_DEPTH"
Case MARKET_DEPTH_L2
    gInputMessageIdToString = "MARKET_DEPTH_L2"
Case NEWS_BULLETINS
    gInputMessageIdToString = "NEWS_BULLETINS"
Case MANAGED_ACCTS
    gInputMessageIdToString = "MANAGED_ACCTS"
Case RECEIVE_FA
    gInputMessageIdToString = "RECEIVE_FA"
Case HISTORICAL_DATA
    gInputMessageIdToString = "HISTORICAL_DATA"
Case BOND_CONTRACT_DATA
    gInputMessageIdToString = "BOND_CONTRACT_DATA"
Case SCANNER_PARAMETERS
    gInputMessageIdToString = "SCANNER_PARAMETERS"
Case SCANNER_DATA
    gInputMessageIdToString = "SCANNER_DATA"
Case TICK_OPTION_COMPUTATION
    gInputMessageIdToString = "TICK_OPTION_COMPUTATION"
Case TICK_GENERIC
    gInputMessageIdToString = "TICK_GENERIC"
Case TICK_STRING
    gInputMessageIdToString = "TICK_STRING"
Case TICK_EFP
    gInputMessageIdToString = "TICK_EFP"
Case CURRENT_TIME
    gInputMessageIdToString = "CURRENT_TIME"
Case REAL_TIME_BARS
    gInputMessageIdToString = "REAL_TIME_BARS"
Case FUNDAMENTAL_DATA
    gInputMessageIdToString = "FUNDAMENTAL_DATA"
Case CONTRACT_DATA_END
    gInputMessageIdToString = "CONTRACT_DATA_END"
Case OPEN_ORDER_END
    gInputMessageIdToString = "OPEN_ORDER_END"
Case ACCT_DOWNLOAD_END
    gInputMessageIdToString = "ACCT_DOWNLOAD_END"
Case EXECUTION_DATA_END
    gInputMessageIdToString = "EXECUTION_DATA_END"
Case DELTA_NEUTRAL_VALIDATION
    gInputMessageIdToString = "DELTA_NEUTRAL_VALIDATION"
Case TICK_SNAPSHOT_END
    gInputMessageIdToString = "TICK_SNAPSHOT_END"
Case MARKET_DATA_TYPE
    gInputMessageIdToString = "MARKET_DATA_TYPE"
Case COMMISSION_REPORT
    gInputMessageIdToString = "COMMISSION_REPORT"
Case POSITION
    gInputMessageIdToString = "POSITION"
Case POSITION_END
    gInputMessageIdToString = "POSITION_END"
Case ACCOUNT_SUMMARY
    gInputMessageIdToString = "ACCOUNT_SUMMARY"
Case ACCOUNT_SUMMARY_END
    gInputMessageIdToString = "ACCOUNT_SUMMARY_END"
Case Else
    gInputMessageIdToString = "?????"
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

Public Function gOutputMessageIdToString( _
                ByVal msgId As TwsSocketOutMsgTypes) As String
Const ProcName As String = "gOutputMessageIdToString"
On Error GoTo Err

Select Case msgId
Case REQ_MKT_DATA
    gOutputMessageIdToString = "REQ_MKT_DATA"
Case CANCEL_MKT_DATA
    gOutputMessageIdToString = "CANCEL_MKT_DATA"
Case PLACE_ORDER
    gOutputMessageIdToString = "PLACE_ORDER"
Case CANCEL_ORDER
    gOutputMessageIdToString = "CANCEL_ORDER"
Case REQ_OPEN_ORDERS
    gOutputMessageIdToString = "REQ_OPEN_ORDERS"
Case REQ_ACCT_DATA
    gOutputMessageIdToString = "REQ_ACCT_DATA"
Case REQ_EXECUTIONS
    gOutputMessageIdToString = "REQ_EXECUTIONS"
Case REQ_IDS
    gOutputMessageIdToString = "REQ_IDS"
Case REQ_CONTRACT_DATA
    gOutputMessageIdToString = "REQ_CONTRACT_DATA"
Case REQ_MKT_DEPTH
    gOutputMessageIdToString = "REQ_MKT_DEPTH"
Case CANCEL_MKT_DEPTH
    gOutputMessageIdToString = "CANCEL_MKT_DEPTH"
Case REQ_NEWS_BULLETINS
    gOutputMessageIdToString = "REQ_NEWS_BULLETINS"
Case CANCEL_NEWS_BULLETINS
    gOutputMessageIdToString = "CANCEL_NEWS_BULLETINS"
Case SET_SERVER_LOGLEVEL
    gOutputMessageIdToString = "SET_SERVER_LOGLEVEL"
Case REQ_AUTO_OPEN_ORDERS
    gOutputMessageIdToString = "REQ_AUTO_OPEN_ORDERS"
Case REQ_ALL_OPEN_ORDERS
    gOutputMessageIdToString = "REQ_ALL_OPEN_ORDERS"
Case REQ_MANAGED_ACCTS
    gOutputMessageIdToString = "REQ_MANAGED_ACCTS"
Case REQ_FA
    gOutputMessageIdToString = "REQ_FA"
Case REPLACE_FA
    gOutputMessageIdToString = "REPLACE_FA"
Case REQ_HISTORICAL_DATA
    gOutputMessageIdToString = "REQ_HISTORICAL_DATA"
Case EXERCISE_OPTIONS
    gOutputMessageIdToString = "EXERCISE_OPTIONS"
Case REQ_SCANNER_SUBSCRIPTION
    gOutputMessageIdToString = "REQ_SCANNER_SUBSCRIPTION"
Case CANCEL_SCANNER_SUBSCRIPTION
    gOutputMessageIdToString = "CANCEL_SCANNER_SUBSCRIPTION"
Case REQ_SCANNER_PARAMETERS
    gOutputMessageIdToString = "REQ_SCANNER_PARAMETERS"
Case CANCEL_HISTORICAL_DATA
    gOutputMessageIdToString = "CANCEL_HISTORICAL_DATA"
Case REQ_CURRENT_TIME
    gOutputMessageIdToString = "REQ_CURRENT_TIME"
Case REQ_REAL_TIME_BARS
    gOutputMessageIdToString = "REQ_REAL_TIME_BARS"
Case CANCEL_REAL_TIME_BARS
    gOutputMessageIdToString = "CANCEL_REAL_TIME_BARS"
Case Else
    gOutputMessageIdToString = "?????"
End Select

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gRoundTimeToSecond( _
                ByVal timeStamp As Date) As Date
Const ProcName As String = "gRoundTimeToSecond"
gRoundTimeToSecond = Int((timeStamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function gSocketInMsgTypeToString( _
                ByVal Value As TwsSocketInMsgTypes) As String
Const ProcName As String = "gSocketInMsgTypeToString"
On Error GoTo Err

Select Case Value
Case TICK_PRICE
    gSocketInMsgTypeToString = "Tick price          "
Case TICK_SIZE
    gSocketInMsgTypeToString = "Tick Size           "
Case ORDER_STATUS
    gSocketInMsgTypeToString = "Order status        "
Case ERR_MSG
    gSocketInMsgTypeToString = "Error message       "
Case OPEN_ORDER
    gSocketInMsgTypeToString = "Open Order          "
Case ACCT_VALUE
    gSocketInMsgTypeToString = "Account Value       "
Case PORTFOLIO_VALUE
    gSocketInMsgTypeToString = "Portfolio Value     "
Case ACCT_UPDATE_TIME
    gSocketInMsgTypeToString = "Account update Time "
Case NEXT_VALID_ID
    gSocketInMsgTypeToString = "Next valid id       "
Case CONTRACT_DATA
    gSocketInMsgTypeToString = "Contract data       "
Case EXECUTION_DATA
    gSocketInMsgTypeToString = "Execution data      "
Case MARKET_DEPTH
    gSocketInMsgTypeToString = "Market depth        "
Case MARKET_DEPTH_L2
    gSocketInMsgTypeToString = "Market depth L2     "
Case NEWS_BULLETINS
    gSocketInMsgTypeToString = "New bulletin        "
Case MANAGED_ACCTS
    gSocketInMsgTypeToString = "Managed accounts    "
Case RECEIVE_FA
    gSocketInMsgTypeToString = "Receive FA          "
Case HISTORICAL_DATA
    gSocketInMsgTypeToString = "Historical data     "
Case BOND_CONTRACT_DATA
    gSocketInMsgTypeToString = "Bond contract data  "
Case SCANNER_PARAMETERS
    gSocketInMsgTypeToString = "scanner parameters  "
Case SCANNER_DATA
    gSocketInMsgTypeToString = "Scanner data        "
Case TICK_OPTION_COMPUTATION
    gSocketInMsgTypeToString = "Option computation  "
Case TICK_GENERIC
    gSocketInMsgTypeToString = "Generic             "
Case TICK_STRING
    gSocketInMsgTypeToString = "String              "
Case TICK_EFP
    gSocketInMsgTypeToString = "EFP                 "
Case CURRENT_TIME
    gSocketInMsgTypeToString = "Current Time        "
Case REAL_TIME_BARS
    gSocketInMsgTypeToString = "Realtime bar        "
Case FUNDAMENTAL_DATA
    gSocketInMsgTypeToString = "Fundamental data    "
Case CONTRACT_DATA_END
    gSocketInMsgTypeToString = "Contract data end   "
Case OPEN_ORDER_END
    gSocketInMsgTypeToString = "Open Order end      "
Case ACCT_DOWNLOAD_END
    gSocketInMsgTypeToString = "Account download end"
Case EXECUTION_DATA_END
    gSocketInMsgTypeToString = "Execution data end  "
Case DELTA_NEUTRAL_VALIDATION
    gSocketInMsgTypeToString = "Delta neutral validn"
Case TICK_SNAPSHOT_END
    gSocketInMsgTypeToString = "Tick snapshot end   "
Case MARKET_DATA_TYPE
    gSocketInMsgTypeToString = "Market data type"
Case COMMISSION_REPORT
    gSocketInMsgTypeToString = "Commission report"
Case POSITION
    gSocketInMsgTypeToString = "Position"
Case POSITION_END
    gSocketInMsgTypeToString = "Position end"
Case ACCOUNT_SUMMARY
    gSocketInMsgTypeToString = "Account summary"
Case ACCOUNT_SUMMARY_END
    gSocketInMsgTypeToString = "Account summary end"
Case Else
    gSocketInMsgTypeToString = "Msg type " & Format(Value, "00                  ")
End Select

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName

End Function

Public Function gTruncateTimeToNextMinute(ByVal timeStamp As Date) As Date
Const ProcName As String = "gTruncateTimeToNextMinute"
On Error GoTo Err

gTruncateTimeToNextMinute = Int((timeStamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gTruncateTimeToMinute(ByVal timeStamp As Date) As Date
Const ProcName As String = "gTruncateTimeToMinute"
On Error GoTo Err

gTruncateTimeToMinute = Int((timeStamp + OneMicrosecond) / OneMinute) * OneMinute

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
Case "", "?"
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
Case "TRAILLMT"
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
    gTwsOrderTypeToString = "MKTCLS"
Case TwsOrderTypes.TwsOrderTypeLimit
    gTwsOrderTypeToString = "LMT"
Case TwsOrderTypes.TwsOrderTypeLimitOnClose
    gTwsOrderTypeToString = "LMTCLS"
Case TwsOrderTypes.TwsOrderTypePeggedToMarket
    gTwsOrderTypeToString = "PEGMKT"
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
    gTwsOrderTypeToString = "TRAILLMT"
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
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException
End Select

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function gTwsSecTypeFromString(ByVal Value As String) As TwsSecTypes
Select Case UCase$(Value)
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


'================================================================================
' Helper Functions
'================================================================================

Private Function contractExpiryToDate( _
                ByVal pContract As TwsContract) As Date
Const ProcName As String = "contractExpiryToDate"
On Error GoTo Err

If Len(pContract.Expiry) = 8 Then
    contractExpiryToDate = CDate(Left$(pContract.Expiry, 4) & "/" & Mid$(pContract.Expiry, 5, 2) & "/" & Right$(pContract.Expiry, 2))
ElseIf Len(pContract.Expiry) = 6 Then
    contractExpiryToDate = CDate(Left$(pContract.Expiry, 4) & "/" & Mid$(pContract.Expiry, 5, 2) & "/" & "01")
End If

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function



