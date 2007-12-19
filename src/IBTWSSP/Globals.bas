Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const OneMicrosecond As Double = 1# / 86400000000#
Public Const OneMinute As Double = 1# / 1440#

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
Public Const ParamNameServer As String = "Server"
Public Const ParamNameTwsLogLevel As String = "TWS Log Level"

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
                ByVal keepConnection As Boolean) As TWSAPI
Dim i As Long

If mTWSAPITableNextIndex = 0 Then
    ReDim mTWSAPITable(5) As TWSAPITableEntry
End If

If clientID < 0 Then clientID = getRandomClientId(clientID)

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).server = server And _
        mTWSAPITable(i).port = port And _
        mTWSAPITable(i).clientID = clientID And _
        mTWSAPITable(i).providerKey = providerKey And _
        mTWSAPITable(i).connectionRetryIntervalSecs = connectionRetryIntervalSecs _
    Then
        Set gGetTWSAPIInstance = mTWSAPITable(i).TWSAPI
        mTWSAPITable(i).usageCount = mTWSAPITable(i).usageCount + 1
        mTWSAPITable(i).keepConnection = keepConnection
        Exit Function
    End If
Next

If mTWSAPITableNextIndex > UBound(mTWSAPITable) Then
    ReDim Preserve mTWSAPITable(UBound(mTWSAPITable) + 5) As TWSAPITableEntry
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
gGetTWSAPIInstance.Connect

End Function

Public Function gHistDataCapabilities() As Long
gHistDataCapabilities = 0
End Function

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
gHistDataSupports = (gHistDataCapabilities And capabilities)
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
        mTWSAPITable(i).TWSAPI.disconnect "release all"
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
                mTWSAPITable(i).TWSAPI.disconnect "release"
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
                
Public Function gTruncateTimeToNextMinute(ByVal timestamp As Date) As Date
gTruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTruncateTimeToMinute(ByVal timestamp As Date) As Date
gTruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
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
                ByVal designator As Long) As Long
                
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



