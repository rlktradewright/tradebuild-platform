Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const OneMicrosecond As Double = 1# / 86400000000#
Public Const OneMinute As Double = 1# / 1440#
Public Const OneSecond As Double = 1# / 86400#

Public Const ContractInfoSPName As String = "IB TWS Contract Info Service Provider"
Public Const HistoricDataSPName As String = "IB TWS Historic Data Service Provider"
Public Const RealtimeDataSPName As String = "IB TWS Realtime Data Service Provider"
Public Const OrderSubmissionSPName As String = "IB TWS Order Submission Service Provider"

Public Const providerKey As String = "TWS"

Public Const ParamNameClientId As String = "Client Id"
Public Const ParamNameConnectionRetryIntervalSecs As String = "Connection Retry Interval Secs"
Public Const ParamNameKeepConnection As String = "Keep Connection"
Public Const ParamNameLogLevel As String = "Log Level"
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

'================================================================================
' Types
'================================================================================

Private Type TWSAPITableEntry
    server          As String
    port            As Long
    clientID        As Long
    providerKey     As String
    connectionRetryIntervalSecs As Long
    keepConnection  As Boolean  ' once this flag is set, the TWSAPI instance
                                ' will only be disconnected by a call to
                                ' gReleaseTWSAPIInstance with <forceDisconnect>
                                ' set to true (and the usageCount is zero),
                                ' or by a call to gReleaseAllTWSAPIInstances
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

'================================================================================
' Procedures
'================================================================================

Public Property Let gCommonServiceConsumer( _
                ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property

Public Function gGetTWSAPIInstance( _
                ByVal server As String, _
                ByVal port As Long, _
                ByVal clientID As Long, _
                ByVal providerKey As String, _
                ByVal connectionRetryIntervalSecs As Long, _
                ByVal keepConnection As Boolean, _
                ByVal logLevel As LogLevels, _
                ByVal TWSLogLevel As TWSLogLevels) As TWSAPI
Dim i As Long

If mTWSAPITableNextIndex = 0 Then
    ReDim mTWSAPITable(1) As TWSAPITableEntry
End If

If clientID < 0 Then clientID = getRandomClientId(clientID & providerKey)

server = UCase$(server)

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).server = server And _
        mTWSAPITable(i).port = port And _
        mTWSAPITable(i).clientID = clientID And _
        mTWSAPITable(i).providerKey = providerKey _
    Then
        Set gGetTWSAPIInstance = mTWSAPITable(i).TWSAPI
        mTWSAPITable(i).usageCount = mTWSAPITable(i).usageCount + 1
        If Not mTWSAPITable(i).keepConnection Then mTWSAPITable(i).keepConnection = keepConnection
        If connectionRetryIntervalSecs > 0 And _
            connectionRetryIntervalSecs < mTWSAPITable(i).connectionRetryIntervalSecs _
        Then
            mTWSAPITable(i).connectionRetryIntervalSecs = connectionRetryIntervalSecs
        End If
        If gGetTWSAPIInstance.logLevel < LogLevelAll Then gGetTWSAPIInstance.logLevel = logLevel
        Exit Function
    End If
Next

If mTWSAPITableNextIndex > UBound(mTWSAPITable) Then
    ReDim Preserve mTWSAPITable(2 * (UBound(mTWSAPITable) + 1) - 1) As TWSAPITableEntry
End If

mTWSAPITable(mTWSAPITableNextIndex).server = server
mTWSAPITable(mTWSAPITableNextIndex).port = port
mTWSAPITable(mTWSAPITableNextIndex).clientID = clientID
mTWSAPITable(mTWSAPITableNextIndex).providerKey = providerKey
mTWSAPITable(mTWSAPITableNextIndex).connectionRetryIntervalSecs = connectionRetryIntervalSecs
mTWSAPITable(mTWSAPITableNextIndex).usageCount = 1
mTWSAPITable(mTWSAPITableNextIndex).keepConnection = keepConnection
Set mTWSAPITable(mTWSAPITableNextIndex).TWSAPI = New TWSAPI
Set gGetTWSAPIInstance = mTWSAPITable(mTWSAPITableNextIndex).TWSAPI

mTWSAPITableNextIndex = mTWSAPITableNextIndex + 1

gGetTWSAPIInstance.commonServiceConsumer = mCommonServiceConsumer
gGetTWSAPIInstance.server = server
gGetTWSAPIInstance.port = port
gGetTWSAPIInstance.clientID = clientID
gGetTWSAPIInstance.providerKey = providerKey
gGetTWSAPIInstance.connectionRetryIntervalSecs = connectionRetryIntervalSecs
gGetTWSAPIInstance.logLevel = logLevel
gGetTWSAPIInstance.TWSLogLevel = TWSLogLevel
gGetTWSAPIInstance.Connect

End Function

Public Function gHistDataCapabilities() As Long
gHistDataCapabilities = 0
End Function

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
gHistDataSupports = (gHistDataCapabilities And capabilities)
End Function

Public Function gParseClientId( _
                value As String) As Long
If value = "" Then gParseClientId = -1
If Not IsInteger(value) Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Client Id' parameter: value must be an integer"
End If
gParseClientId = CLng(value)
End Function

Public Function gParseConnectionRetryInterval( _
                value As String) As Long
If value = "" Then gParseConnectionRetryInterval = 0
If Not IsInteger(value, 0) Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Connection Retry Interval Secs' parameter: value must be an integer >= 0"
End If
gParseConnectionRetryInterval = CLng(value)
End Function

Public Function gParseKeepConnection( _
                value As String) As Boolean
On Error GoTo err
If value = "" Then gParseKeepConnection = False
gParseKeepConnection = CBool(value)
Exit Function

err:
err.Raise ErrorCodes.ErrIllegalArgumentException, _
        , _
        "Invalid 'Keep Connection' parameter: value must be 'true' or 'false'"
End Function

Public Function gParseLogLevel( _
                value As String) As Long
If value = "" Then gParseLogLevel = LogLevels.LogLevelLow
If Not IsInteger(value, 0, 4) Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Log Level' parameter: value must be a positive integer less than 5 "
End If
gParseLogLevel = CLng(value)
End Function

Public Function gParsePort( _
                value As String) As Long
If value = "" Then gParsePort = 7496
If Not IsInteger(value, 1024, 65535) Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Port' parameter: value must be an integer >= 1024 and <=65535"
End If
gParsePort = CLng(value)
End Function

Public Function gParseRole( _
                value As String) As String

Select Case UCase$(value)
Case "", "P", "PR", "PRIM", "PRIMARY"
    gParseRole = "PRIMARY"
Case "S", "SEC", "SECOND", "SECONDARY"
    gParseRole = "SECONDARY"
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Invalid 'Role' parameter: value must be one of 'P', 'PR', 'PRIM', 'PRIMARY', 'S', 'SEC', 'SECOND', or 'SECONDARY'"
End Select
End Function

Public Function gParseTwsLogLevel( _
                value As String) As TWSLogLevels
On Error GoTo err
If value = "" Then gParseTwsLogLevel = TWSLogLevelError
gParseTwsLogLevel = gTwsLogLevelFromString(value)
Exit Function

err:
err.Raise ErrorCodes.ErrIllegalArgumentException, _
        , _
        "Invalid 'Tws Log Level' parameter: value must be one of " & _
        TWSLogLevelSystemString & ", " & _
        TWSLogLevelErrorString & ", " & _
        TWSLogLevelWarningString & ", " & _
        TWSLogLevelInformationString & " or " & _
        TWSLogLevelDetailString
End Function

Public Function gRealtimeDataCapabilities() As Long
gRealtimeDataCapabilities = TradeBuildSP.RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPosition
End Function

Public Function gRealtimeDataSupports(ByVal capabilities As Long) As Boolean
gRealtimeDataSupports = (gRealtimeDataCapabilities And capabilities)
End Function

Public Sub gReleaseAllTWSAPIInstances()

Dim i As Long

For i = 0 To mTWSAPITableNextIndex - 1
    mTWSAPITable(i).usageCount = 0
    If Not mTWSAPITable(i).TWSAPI Is Nothing Then
        mTWSAPITable(i).TWSAPI.disconnect "release all", False
        Set mTWSAPITable(i).TWSAPI = Nothing
    End If
    mTWSAPITable(i).clientID = 0
    mTWSAPITable(i).connectionRetryIntervalSecs = 0
    mTWSAPITable(i).port = 0
    mTWSAPITable(i).server = ""
    mTWSAPITable(i).keepConnection = False
    mTWSAPITable(i).providerKey = ""
Next
                
End Sub

Public Sub gReleaseTWSAPIInstance( _
                ByVal instance As TWSAPI, _
                Optional ByVal forceDisconnect As Boolean)

Dim i As Long

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).TWSAPI Is instance Then
        mTWSAPITable(i).usageCount = mTWSAPITable(i).usageCount - 1
        If (mTWSAPITable(i).usageCount = 0 And _
                (Not mTWSAPITable(i).keepConnection)) Or _
            forceDisconnect _
        Then
            If mTWSAPITable(i).TWSAPI.connectionState <> ConnNotConnected Then
                mTWSAPITable(i).TWSAPI.disconnect "release", forceDisconnect
            End If
            Set mTWSAPITable(i).TWSAPI = Nothing
            mTWSAPITable(i).clientID = 0
            mTWSAPITable(i).connectionRetryIntervalSecs = 0
            mTWSAPITable(i).port = 0
            mTWSAPITable(i).server = ""
            mTWSAPITable(i).keepConnection = False
            mTWSAPITable(i).providerKey = ""
        End If
        Exit For
    End If
Next
                
End Sub

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
    err.Raise ErrorCodes.ErrIllegalArgumentException
End Select
End Function

'================================================================================
' Helper Functions
'================================================================================

Public Function clientIdAlreadyInUse( _
                ByVal value As Long) As Boolean
Dim i As Long
For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).clientID = value Then
        clientIdAlreadyInUse = True
        Exit Function
    End If
Next
                
End Function

Public Function getRandomClientId( _
                ByVal designator As String) As Long
                
If mRandomClientIds Is Nothing Then
    Set mRandomClientIds = New Collection
    Randomize
End If

' first see if a clientId has already been generated for this designator

On Error Resume Next
getRandomClientId = mRandomClientIds(CStr(designator))
On Error GoTo 0

If getRandomClientId <> 0 Then
    Exit Function   ' clientId already exists for this designator
End If

getRandomClientId = Rnd * (&H7FFFFFFF - &H7000000) + &H7000000

Do While clientIdAlreadyInUse(getRandomClientId)
    getRandomClientId = Rnd * (&H7FFFFFFF - &H7000000) + &H7000000
Loop

mRandomClientIds.add getRandomClientId, CStr(designator)

End Function

Public Function gRoundTimeToSecond( _
                ByVal timestamp As Date) As Date
gRoundTimeToSecond = Int((timestamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function



