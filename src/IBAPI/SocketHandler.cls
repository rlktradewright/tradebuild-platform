VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SocketHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' Raised when a successful connection to Tws has been achieved.
Event Connected()

' Raised when connection to Tws has failed and the connection
' retry interval is zero.
Event ConnectFailed(ByVal pDescription As String, _
                ByVal pRetryInterval As Long)

' Raised when an attempt to connect to Tws is initiated.
Event Connecting()

' Raised when the connection to Tws is disconnected.
Event ConnectionClosed(ByVal pDescription As String)

' Raised when we Disconnect from Tws
Event Disconnected(ByVal pReason As String)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "SocketHandler"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mSocket                          As Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mWinsockForm                                As WinsockForm

Private mServer                                     As String
Private mPort                                       As Long

Private mRetryingConnection                         As Boolean

Private WithEvents mConnectionTLI                   As TimerListItem
Attribute mConnectionTLI.VB_VarHelpID = -1

Private mBufferedReader                             As BufferedReader

Private mIsConnecting                               As Boolean
Private mIsConnected                                As Boolean

Private mConnectionRetryIntervalSecs                As Long
Private mTempConnectionRetryIntervalSecs            As Long

Private mProgramErrorHandler                        As IProgramErrorListener

Private mRawApiMessageLogLevel                      As LogLevels

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'================================================================================
' mConnectionTLI Event Handlers
'================================================================================

Private Sub mConnectionTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mConnectionTLI_StateChange"
On Error GoTo Err
If ev.State = TimerListItemStates.TimerListItemStateExpired Then
    Set mConnectionTLI = Nothing
    Connect
ElseIf ev.State = TimerListItemStates.TimerListItemStateCancelled Then
    Set mConnectionTLI = Nothing
End If
Exit Sub
Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mSocket Event Handlers
'================================================================================

Private Sub mSocket_Close()
Const ProcName As String = "mSocket_Close"
On Error GoTo Err
ReleaseSocket
handleTwsDisconnection "closed by peer", False
Exit Sub
Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocket_Connect()
Const ProcName As String = "mSocket_Connect"
On Error GoTo Err

Dim ka As TcpKeepalive
ka.Onoff = 1
ka.KeepaliveInterval = 1000
ka.KeepaliveTime = 15000

Dim ov As WSAOVERLAPPED
Dim oBuf(10) As Byte
Dim bytesReturned As Long

Dim result As Long
result = WSAIoctl(mSocket.SocketHandle, _
        SIO_KEEPALIVE_VALS, _
        VarPtr(ka), _
        Len(ka), _
        VarPtr(oBuf(0)), _
        0, _
        bytesReturned, _
        ov, _
        AddressOf WSAIoctlCompletionRoutine)
If result <> 0 Then
    result = Err.LastDllError
    Assert result <> WSA_IO_PENDING, "WSAIoctl returned " & result
End If
        
mIsConnecting = False
mIsConnected = True

mBufferedReader.ClearBuffer

RaiseEvent Connected

Exit Sub

Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocket_DataArrival( _
                ByVal bytesTotal As Long)
Const ProcName As String = "mSocket_DataArrival"
On Error GoTo Err

If mSocket.State <> StateConstants.sckConnected Then Exit Sub

Dim lSocketInputData() As Byte
mSocket.GetData lSocketInputData, vbArray + vbByte

Dim lBytesReceived As Long
lBytesReceived = UBound(lSocketInputData) + 1

If GIB.SocketLogger.IsLoggable(mRawApiMessageLogLevel) Then
    GIB.SocketLogger.Log ByteBufferToString("In buf: ", lSocketInputData), ProcName, ModuleName, mRawApiMessageLogLevel
End If

mBufferedReader.AddData lSocketInputData

Exit Sub
Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocket_Error(ByVal pNumber As Integer, _
                        ByRef pDescription As String, _
                        ByVal Scode As Long, _
                        ByVal pSource As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        ByRef CancelDisplay As Boolean)
Const ProcName As String = "mSocket_Error"
Dim errorNum As MSWinsockLib.ErrorConstants

On Error GoTo Err
errorNum = pNumber
Select Case errorNum
'Case sckInvalidPropertyValue

'Case sckGetNotSupported

'Case sckSetNotSupported

'Case sckOutOfMemory

'Case sckBadState

'Case sckInvalidArg

'Case sckSuccess

'Case sckUnsupported

'Case sckInvalidOp

'Case sckOutOfRange

'Case sckWrongProtocol

'Case sckOpCanceled

'Case sckInvalidArgument

'Case sckWouldBlock

'Case sckInProgress

'Case sckAlreadyComplete

'Case sckNotSocket

'Case sckMsgTooBig

'Case sckPortNotSupported

'Case sckAddressInUse

Case sckAddressNotAvailable, _
    sckNetworkSubsystemFailed, _
    sckNetworkUnreachable, _
    sckNetReset, _
    sckConnectAborted, _
    sckConnectionRefused, _
    sckHostNotFound, _
    sckHostNotFoundTryAgain, _
    sckTimedout, _
    sckConnectionReset
    
    ReleaseSocket
    If Not mIsConnected Then
        mIsConnecting = False
        GIB.Log "Failed to connect to Tws" & _
            IIf(mConnectionRetryIntervalSecs <> 0, " - retrying in " & mConnectionRetryIntervalSecs & " seconds ", " ") & _
            pDescription & ": " & getConnectionString, _
            ModuleName, ProcName
        
        RaiseEvent ConnectFailed(pDescription, mConnectionRetryIntervalSecs)
        If mConnectionRetryIntervalSecs <> 0 Then retryConnection
    Else
        mIsConnected = False
        GIB.Log "Socket error " & pNumber & ": " & _
                    pDescription & ": " & _
                    getConnectionString, _
                    ModuleName, ProcName
        handleTwsDisconnection pDescription, False
    End If

'Case sckNoBufferSpace

'Case sckAlreadyConnected

'Case sckNotConnected

'Case sckSocketShutdown

'Case sckNotInitialized

'Case sckNonRecoverableError

'Case sckNoData
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Socket error " & pNumber & ": " & _
                                                        pDescription & ": " & _
                                                        getConnectionString
End Select

Exit Sub
Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ConnectionRetryIntervalSecs(ByVal Value As Long)
mConnectionRetryIntervalSecs = Value
If mConnectionRetryIntervalSecs = 0 And mRetryingConnection Then
    mConnectionTLI.Cancel
    Set mConnectionTLI = Nothing
End If
End Property

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Friend Property Let TempConnectionRetryIntervalSecs(ByVal Value As Long)
mTempConnectionRetryIntervalSecs = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Connect()
Const ProcName As String = "Connect"
On Error GoTo Err

If mIsConnected Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Already connected"

createsocket

Dim s As String
s = "Connecting to Tws: " & getConnectionString
GIB.Log s, ModuleName, ProcName

mSocket.Connect mServer, mPort
mIsConnecting = True

RaiseEvent Connecting

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Disconnect( _
                ByVal pReason As String, _
                ByVal pClosedByApplication As Boolean)
Const ProcName As String = "Disconnect"
On Error GoTo Err

If Not (mIsConnecting Or mIsConnected) Then Exit Sub
mIsConnecting = False
mIsConnected = False

If pClosedByApplication And Not mConnectionTLI Is Nothing Then
    mConnectionTLI.Cancel
    Set mConnectionTLI = Nothing
End If

ReleaseSocket

GIB.Log "Disconnected from: " & getConnectionString & ": " & pReason, ModuleName, ProcName
handleTwsDisconnection pReason, pClosedByApplication
RaiseEvent Disconnected(pReason)
Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pBufferedReader As BufferedReader, _
                ByVal pProgramErrorHandler As IProgramErrorListener, _
                ByVal pLogRawApiMessages As TWSApiMessageLoggingOptions)
If pServer = "" Then pServer = "127.0.0.1"
mServer = pServer
mPort = pPort
Set mBufferedReader = pBufferedReader
Set mProgramErrorHandler = pProgramErrorHandler
If pLogRawApiMessages = TWSApiMessageLoggingOptionAlways Then
    mRawApiMessageLogLevel = LogLevelNormal
ElseIf pLogRawApiMessages = TWSApiMessageLoggingOptionDefault Then
    mRawApiMessageLogLevel = LogLevelMediumDetail
Else
    mRawApiMessageLogLevel = LogLevelNone
End If
End Sub

Friend Sub SendBytes(ByRef pData() As Byte)
Const ProcName As String = "SendBytes"
On Error GoTo Err

mSocket.SendData pData

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub SendLong(ByVal pData As Long)
Const ProcName As String = "SendLong"
On Error GoTo Err

mSocket.SendData pData

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createsocket()
Const ProcName As String = "createsocket"
On Error GoTo Err

If mWinsockForm Is Nothing Then Set mWinsockForm = New WinsockForm
Set mSocket = mWinsockForm.Winsock1

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getConnectionString() As String
getConnectionString = "server=" & mServer & _
                        " port=" & mPort
End Function

Private Sub handleTwsDisconnection(ByVal pMessage As String, ByVal pClosedByApplication As Boolean)
Const ProcName As String = "handleTwsDisconnection"
On Error GoTo Err

GIB.Log "Connection to Tws closed: " & pMessage & ": " & getConnectionString, ModuleName, ProcName

Set mSocket = Nothing
mIsConnected = False

RaiseEvent ConnectionClosed(pMessage)
If Not pClosedByApplication Then retryConnection    ' 60

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub ReleaseSocket()
Const ProcName As String = "ReleaseSocket"
On Error GoTo Err

If Not mSocket Is Nothing Then
    GIB.Log "Releasing socket: " & getConnectionString, ModuleName, ProcName
    mSocket.Close
    Set mSocket = Nothing
    Unload mWinsockForm
    Set mWinsockForm = Nothing
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub retryConnection(Optional ByVal pRetryInterval As Long = 0)
Const ProcName As String = "retryConnection"
On Error GoTo Err

Dim lRetryInterval As Long
If mConnectionRetryIntervalSecs <> 0 Then lRetryInterval = mConnectionRetryIntervalSecs
If pRetryInterval <> 0 Then lRetryInterval = pRetryInterval
If mTempConnectionRetryIntervalSecs <> 0 And mTempConnectionRetryIntervalSecs < lRetryInterval Then lRetryInterval = mTempConnectionRetryIntervalSecs

If lRetryInterval <> 0 Then
    GIB.Log "Reconnecting in " & lRetryInterval & " seconds", ModuleName, ProcName
    Set mConnectionTLI = GetGlobalTimerList.Add(Nothing, lRetryInterval, ExpiryTimeUnitSeconds)
    mRetryingConnection = True
    mTempConnectionRetryIntervalSecs = 0
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub





