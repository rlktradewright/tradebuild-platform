Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                    As String = "IBTWSSP26"
Private Const ModuleName                As String = "Globals"


Public Const MaxLong As Long = &H7FFFFFFF
Public Const OneMicrosecond As Double = 1# / 86400000000#
Public Const OneMinute As Double = 1# / 1440#
Public Const OneSecond As Double = 1# / 86400#

Public Const ContractInfoSPName As String = "IB TWS Contract Info Service Provider"
Public Const HistoricDataSPName As String = "IB TWS Historic Data Service Provider"
Public Const RealtimeDataSPName As String = "IB TWS Realtime Data Service Provider"
Public Const OrderSubmissionSPName As String = "IB TWS Order Submission Service Provider"

Public Const ProviderKey As String = "TWS"

Public Const ParamNameClientId As String = "Client Id"
Public Const ParamNameConnectionRetryIntervalSecs As String = "Connection Retry Interval Secs"
Public Const ParamNameKeepConnection As String = "Keep Connection"
Public Const ParamNamePort As String = "Port"
Public Const ParamNameProviderKey As String = "Provider Key"
Public Const ParamNameRole As String = "Role"
Public Const ParamNameServer As String = "Server"
Public Const ParamNameTwsLogLevel As String = "TWS Log Level"

Public Const TWSLogLevelDetailString        As String = "Detail"
Public Const TWSLogLevelErrorString         As String = "Error"
Public Const TWSLogLevelInformationString   As String = "Information"
Public Const TWSLogLevelSystemString        As String = "System"
Public Const TWSLogLevelWarningString       As String = "Warning"

'================================================================================
' Enums
'================================================================================

Public Enum ConnectionStates
    ConnNotConnected
    ConnConnecting
    ConnReSynching
    ConnConnected
End Enum

Public Enum FADataTypes
    FAGroups = 1
    FAProfile
    FAAccountAliases
End Enum

Public Enum TWSLogLevels
    TWSLogLevelSystem = 1
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
    TICK_OPTION_COMPUTATION = 21
    TICK_GENERIC = 45
    TICK_STRING = 46
    TICK_EFP = 47
    CURRENT_TIME = 49
    REAL_TIME_BARS = 50
    FUNDAMENTAL_DATA = 51
    CONTRACT_DATA_END = 52
    MAX_SOCKET_INMSG
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
    REQ_CURRENT_TIME = 49
    REQ_REAL_TIME_BARS = 50
    CANCEL_REAL_TIME_BARS = 51
End Enum

Public Enum TWSSocketTickTypes
    TICK_BID_SIZE                   ' 0
    TICK_BID                        ' 1
    TICK_ASK                        ' 2
    TICK_ASK_SIZE                   ' 3
    TICK_LAST                       ' 4
    TICK_LAST_SIZE                  ' 5
    TICK_HIGH                       ' 6
    TICK_LOW                        ' 7
    TICK_VOLUME                     ' 8
    TICK_CLOSE                      ' 9
    TICK_BID_OPTION                 ' 10
    TICK_ASK_OPTION                 ' 11
    TICK_LAST_OPTION                ' 12
    TICK_MODEL_OPTION               ' 13
    TICK_OPEN                       ' 14
    TICK_LOW_13_WEEK                ' 15
    TICK_HIGH_13_WEEK               ' 16
    TICK_LOW_26_WEEK                ' 17
    TICK_HIGH_26_WEEK               ' 18
    TICK_LOW_52_WEEK                ' 19
    TICK_HIGH_52_WEEK               ' 20
    TICK_AVG_VOLUME                 ' 21
    TICK_OPEN_INTEREST              ' 22
    TICK_OPTION_HISTORICAL_VOL      ' 23
    TICK_OPTION_IMPLIED_VOL         ' 24
    TICK_OPTION_BID_EXCH            ' 25
    TICK_OPTION_ASK_EXCH            ' 26
    TICK_OPTION_CALL_OPEN_INTEREST  ' 27
    TICK_OPTION_PUT_OPEN_INTEREST   ' 28
    TICK_OPTION_CALL_VOLUME         ' 29
    TICK_OPTION_PUT_VOLUME          ' 30
    TICK_INDEX_FUTURE_PREMIUM       ' 31
    TICK_BID_EXCH                   ' 32
    TICK_ASK_EXCH                   ' 33
    TICK_AUCTION_VOLUME             ' 34
    TICK_AUCTION_PRICE              ' 35
    TICK_AUCTION_IMBALANCE          ' 36
    TICK_MARK_PRICE                 ' 37
    TICK_BID_EFP_COMPUTATION        ' 38
    TICK_ASK_EFP_COMPUTATION        ' 39
    TICK_LAST_EFP_COMPUTATION       ' 40
    TICK_OPEN_EFP_COMPUTATION       ' 41
    TICK_HIGH_EFP_COMPUTATION       ' 42
    TICK_LOW_EFP_COMPUTATION        ' 43
    TICK_CLOSE_EFP_COMPUTATION      ' 44
    TICK_LAST_TIMESTAMP             ' 45
    TICK_SHORTABLE                  ' 46
End Enum

'================================================================================
' Types
'================================================================================

Private Type TWSAPITableEntry
    Server          As String
    Port            As Long
    clientID        As Long
    ProviderKey     As String
    ConnectionRetryIntervalSecs As Long
'    KeepConnection  As Boolean  ' once this flag is set, the TWSAPI instance
'                                ' will only be disconnected by a call to
'                                ' gReleaseTWSAPIInstance with <forceDisconnect>
'                                ' set to true (and the usageCount is zero),
'                                ' or by a call to gReleaseAllTWSAPIInstances
    TWSAPI          As TWSAPI
    usageCount      As Long
End Type

'================================================================================
' Global variables
'================================================================================

'================================================================================
' Private variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mTWSAPITable() As TWSAPITableEntry
Private mTWSAPITableNextIndex As Long

Private mRandomClientIds As Collection

Private mLogger As Logger

Private mLogTokens(9) As String

'================================================================================
' Properties
'================================================================================

Public Property Let gCommonServiceConsumer( _
                ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property

Public Sub gLog(ByRef pMsg As String, _
                ByRef pProjName As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
If Not gLogger.IsLoggable(pLogLevel) Then Exit Sub
mLogTokens(0) = "["
mLogTokens(1) = pProjName
mLogTokens(2) = "."
mLogTokens(3) = pModName
mLogTokens(4) = ":"
mLogTokens(5) = pProcName
mLogTokens(6) = "] "
mLogTokens(7) = pMsg
If Len(pMsgQualifier) <> 0 Then
    mLogTokens(8) = ": "
    mLogTokens(9) = pMsgQualifier
Else
    mLogTokens(8) = vbNullString
    mLogTokens(9) = vbNullString
End If

gLogger.Log pLogLevel, Join(mLogTokens, "")
End Sub

Public Property Get gLogger() As Logger
If mLogger Is Nothing Then Set mLogger = GetLogger("tradebuild.log.serviceprovider.ibtwssp")
Set gLogger = mLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Function gGetTWSAPIInstance( _
                ByVal Server As String, _
                ByVal Port As Long, _
                ByVal clientID As Long, _
                ByVal ProviderKey As String, _
                ByVal ConnectionRetryIntervalSecs As Long, _
                ByVal TWSLogLevel As TWSLogLevels) As TWSAPI
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If mTWSAPITableNextIndex = 0 Then
    ReDim mTWSAPITable(1) As TWSAPITableEntry
End If

If clientID < 0 Then clientID = getRandomClientId(clientID & ProviderKey)

Server = UCase$(Server)

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).Server = Server And _
        mTWSAPITable(i).Port = Port And _
        mTWSAPITable(i).clientID = clientID And _
        mTWSAPITable(i).ProviderKey = ProviderKey _
    Then
        Set gGetTWSAPIInstance = mTWSAPITable(i).TWSAPI
        mTWSAPITable(i).usageCount = mTWSAPITable(i).usageCount + 1
        If ConnectionRetryIntervalSecs > 0 And _
            ConnectionRetryIntervalSecs < mTWSAPITable(i).ConnectionRetryIntervalSecs _
        Then
            mTWSAPITable(i).ConnectionRetryIntervalSecs = ConnectionRetryIntervalSecs
        End If
        Exit Function
    End If
Next

If mTWSAPITableNextIndex > UBound(mTWSAPITable) Then
    ReDim Preserve mTWSAPITable(2 * (UBound(mTWSAPITable) + 1) - 1) As TWSAPITableEntry
End If

mTWSAPITable(mTWSAPITableNextIndex).Server = Server
mTWSAPITable(mTWSAPITableNextIndex).Port = Port
mTWSAPITable(mTWSAPITableNextIndex).clientID = clientID
mTWSAPITable(mTWSAPITableNextIndex).ProviderKey = ProviderKey
mTWSAPITable(mTWSAPITableNextIndex).ConnectionRetryIntervalSecs = ConnectionRetryIntervalSecs
mTWSAPITable(mTWSAPITableNextIndex).usageCount = 1
Set mTWSAPITable(mTWSAPITableNextIndex).TWSAPI = New TWSAPI
Set gGetTWSAPIInstance = mTWSAPITable(mTWSAPITableNextIndex).TWSAPI

mTWSAPITableNextIndex = mTWSAPITableNextIndex + 1

gGetTWSAPIInstance.commonServiceConsumer = mCommonServiceConsumer
gGetTWSAPIInstance.Server = Server
gGetTWSAPIInstance.Port = Port
gGetTWSAPIInstance.clientID = clientID
gGetTWSAPIInstance.ProviderKey = ProviderKey
gGetTWSAPIInstance.ConnectionRetryIntervalSecs = ConnectionRetryIntervalSecs
gGetTWSAPIInstance.TWSLogLevel = TWSLogLevel
gGetTWSAPIInstance.Connect

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gGetTWSAPIInstance", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Public Function gHistDataCapabilities() As Long
gHistDataCapabilities = 0
End Function

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
gHistDataSupports = (gHistDataCapabilities And capabilities)
End Function

Public Function gInputMessageIdToString( _
                ByVal msgId As TWSSocketInMsgTypes) As String
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
    gInputMessageIdToString = "ACCT_VALUE"
Case PORTFOLIO_VALUE
    gInputMessageIdToString = "PORTFOLIO_VALUE"
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
Case Else
    gInputMessageIdToString = "?????"
End Select
                
End Function

Public Function gOutputMessageIdToString( _
                ByVal msgId As TWSSocketOutMsgTypes) As String
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
End Function

Public Function gParseClientId( _
                value As String) As Long
Dim failpoint As Long
On Error GoTo Err

If value = "" Then
    gParseClientId = -1
ElseIf Not IsInteger(value) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Client Id' parameter: value must be an integer"
Else
    gParseClientId = CLng(value)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gParseClientId", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gParseConnectionRetryInterval( _
                value As String) As Long
Dim failpoint As Long
On Error GoTo Err

If value = "" Then
    gParseConnectionRetryInterval = 0
ElseIf Not IsInteger(value, 0) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Connection Retry Interval Secs' parameter: value must be an integer >= 0"
Else
    gParseConnectionRetryInterval = CLng(value)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gParseConnectionRetryInterval", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gParseKeepConnection( _
                value As String) As Boolean
On Error GoTo Err
If value = "" Then
    gParseKeepConnection = False
Else
    gParseKeepConnection = CBool(value)
End If
Exit Function

Err:
Err.Raise ErrorCodes.ErrIllegalArgumentException, _
        , _
        "Invalid 'Keep Connection' parameter: value must be 'true' or 'false'"
End Function

Public Function gParsePort( _
                value As String) As Long
Dim failpoint As Long
On Error GoTo Err

If value = "" Then
    gParsePort = 7496
ElseIf Not IsInteger(value, 1024, 65535) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Port' parameter: value must be an integer >= 1024 and <=65535"
Else
    gParsePort = CLng(value)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gParsePort", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gParseRole( _
                value As String) As String

Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(value)
Case "", "P", "PR", "PRIM", "PRIMARY"
    gParseRole = "PRIMARY"
Case "S", "SEC", "SECOND", "SECONDARY"
    gParseRole = "SECONDARY"
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Role' parameter: value must be one of 'P', 'PR', 'PRIM', 'PRIMARY', 'S', 'SEC', 'SECOND', or 'SECONDARY'"
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gParseRole", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gParseTwsLogLevel( _
                value As String) As TWSLogLevels
Dim failpoint As Long
On Error GoTo Err

If value = "" Then
    gParseTwsLogLevel = TWSLogLevelError
Else
    gParseTwsLogLevel = gTwsLogLevelFromString(value)
End If
Exit Function

Err:
If Err.number = ErrorCodes.ErrIllegalArgumentException Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Tws Log Level' parameter: value must be one of " & _
            TWSLogLevelSystemString & ", " & _
            TWSLogLevelErrorString & ", " & _
            TWSLogLevelWarningString & ", " & _
            TWSLogLevelInformationString & " or " & _
            TWSLogLevelDetailString
End If

HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gParseTwsLogLevel", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gRealtimeDataCapabilities() As Long
gRealtimeDataCapabilities = TradeBuildSP.RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPosition
End Function

Public Function gRealtimeDataSupports(ByVal capabilities As Long) As Boolean
gRealtimeDataSupports = (gRealtimeDataCapabilities And capabilities)
End Function

Public Sub gReleaseAllTWSAPIInstances()

Dim i As Long

Dim failpoint As Long
On Error GoTo Err

For i = 0 To mTWSAPITableNextIndex - 1
    mTWSAPITable(i).usageCount = 0
    If Not mTWSAPITable(i).TWSAPI Is Nothing Then
        mTWSAPITable(i).TWSAPI.Disconnect "release all", True
        Set mTWSAPITable(i).TWSAPI = Nothing
    End If
    mTWSAPITable(i).clientID = 0
    mTWSAPITable(i).ConnectionRetryIntervalSecs = 0
    mTWSAPITable(i).Port = 0
    mTWSAPITable(i).Server = ""
    mTWSAPITable(i).ProviderKey = ""
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gReleaseAllTWSAPIInstances", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Sub

Public Sub gReleaseTWSAPIInstance( _
                ByVal instance As TWSAPI, _
                Optional ByVal forceDisconnect As Boolean)

Dim i As Long

Dim failpoint As Long
On Error GoTo Err

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).TWSAPI Is instance Then
        mTWSAPITable(i).usageCount = mTWSAPITable(i).usageCount - 1
        If mTWSAPITable(i).usageCount = 0 Or _
            forceDisconnect _
        Then
            If mTWSAPITable(i).TWSAPI.connectionState <> ConnNotConnected Then
                mTWSAPITable(i).TWSAPI.Disconnect "release", forceDisconnect
            End If
            Set mTWSAPITable(i).TWSAPI = Nothing
            mTWSAPITable(i).clientID = 0
            mTWSAPITable(i).ConnectionRetryIntervalSecs = 0
            mTWSAPITable(i).Port = 0
            mTWSAPITable(i).Server = ""
            mTWSAPITable(i).ProviderKey = ""
        End If
        Exit For
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gReleaseTWSAPIInstance", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Sub

Public Function gRoundTimeToSecond( _
                ByVal timestamp As Date) As Date
gRoundTimeToSecond = Int((timestamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function

Public Function gSocketInMsgTypeToString( _
                ByVal value As TWSSocketInMsgTypes) As String
Select Case value
Case TICK_PRICE
    gSocketInMsgTypeToString = "Tick price          "
Case TICK_SIZE
    gSocketInMsgTypeToString = "Tick size           "
Case ORDER_STATUS
    gSocketInMsgTypeToString = "Order status        "
Case ERR_MSG
    gSocketInMsgTypeToString = "Error message       "
Case OPEN_ORDER
    gSocketInMsgTypeToString = "Open order          "
Case ACCT_VALUE
    gSocketInMsgTypeToString = "Account value       "
Case PORTFOLIO_VALUE
    gSocketInMsgTypeToString = "Portfolio value     "
Case ACCT_UPDATE_TIME
    gSocketInMsgTypeToString = "Account update time "
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
    gSocketInMsgTypeToString = "Current time        "
Case REAL_TIME_BARS
    gSocketInMsgTypeToString = "Realtime bar        "
Case FUNDAMENTAL_DATA
    gSocketInMsgTypeToString = "Fundamental data    "
Case CONTRACT_DATA_END
    gSocketInMsgTypeToString = "Contract data end   "
Case Else
    gSocketInMsgTypeToString = "Msg type " & Format(value, "00         ")
End Select

End Function
                
Public Function gTruncateTimeToNextMinute(ByVal timestamp As Date) As Date
gTruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTruncateTimeToMinute(ByVal timestamp As Date) As Date
gTruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTwsLogLevelFromString( _
                ByVal value As String) As TWSLogLevels
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(value)
Case UCase$(TWSLogLevelDetailString)
    gTwsLogLevelFromString = TWSLogLevelDetail
Case UCase$(TWSLogLevelErrorString)
    gTwsLogLevelFromString = TWSLogLevelError
Case UCase$(TWSLogLevelInformationString)
    gTwsLogLevelFromString = TWSLogLevelInformation
Case UCase$(TWSLogLevelSystemString)
    gTwsLogLevelFromString = TWSLogLevelSystem
Case UCase$(TWSLogLevelWarningString)
    gTwsLogLevelFromString = TWSLogLevelWarning
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="gTwsLogLevelFromString", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function clientIdAlreadyInUse( _
                ByVal value As Long) As Boolean
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).clientID = value Then
        clientIdAlreadyInUse = True
        Exit Function
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="clientIdAlreadyInUse", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Function

Private Function getRandomClientId( _
                ByVal designator As String) As Long
                
Dim failpoint As Long
On Error GoTo Err

If mRandomClientIds Is Nothing Then
    Set mRandomClientIds = New Collection
    Randomize
End If

' first see if a clientId has already been generated for this designator

On Error Resume Next
getRandomClientId = mRandomClientIds(CStr(designator))
On Error GoTo Err

If getRandomClientId <> 0 Then
    Exit Function   ' clientId already exists for this designator
End If

getRandomClientId = Rnd * (&H7FFFFFFF - &H70000000) + &H70000000

Do While clientIdAlreadyInUse(getRandomClientId)
    getRandomClientId = Rnd * (&H7FFFFFFF - &H70000000) + &H70000000
Loop

mRandomClientIds.add getRandomClientId, CStr(designator)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="getRandomClientId", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function


