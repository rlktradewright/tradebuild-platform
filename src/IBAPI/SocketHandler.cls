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

Private Type TcpKeepalive
    Onoff As Long
    KeepaliveTime As Long
    KeepaliveInterval As Long
End Type

Private Type WSAOVERLAPPED
    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type


Const SOL_SOCKET = 65535      ' Options for socket level.
Const SO_KEEPALIVE = &H8&     ' Keep connections alive.

Const IOC_IN                    As Long = &H80000000
Const IOC_VENDOR                As Long = &H18000000
Const SIO_KEEPALIVE_VALS        As Long = IOC_IN Or IOC_VENDOR Or 4

Const WSA_IO_PENDING            As Long = 997

Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private Declare Function WSAIoctl Lib "ws2_32.dll" ( _
                ByVal s As Long, _
                ByVal dwIoControlCode As Long, _
                ByVal lpvInBuffer As Long, _
                ByVal cbInBuffer As Long, _
                ByVal lpvOutBuffer As Long, _
                ByVal cbOutBuffer, _
                ByRef lpcbBytesReturned As Long, _
                ByRef lpOverlapped As WSAOVERLAPPED, _
                ByVal lpCompletionRoutine) As Long
  
Private Declare Function setsockopt Lib "wsock32.dll" ( _
                ByVal s As Long, _
                ByVal level As Long, _
                ByVal optname As Long, _
                ByVal optval As Long, _
                ByVal optlen As Long) As Long

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

Private WithEvents mConnectionTimer                 As IntervalTimer
Attribute mConnectionTimer.VB_VarHelpID = -1

Private mBufferedReader                             As BufferedReader

Private mIsConnecting                               As Boolean
Private mIsConnected                                As Boolean

Private mConnectionRetryIntervalSecs                As Long

Private mProgramErrorHandler                        As IProgramErrorListener

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'================================================================================
' mConnectionTimer Event Handlers
'================================================================================

Private Sub mConnectionTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mConnectionTimer_TimerExpired"
On Error GoTo Err
Set mConnectionTimer = Nothing
Connect
Exit Sub
Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mSocket Event Handlers
'================================================================================

Private Sub mSocket_Close()
Const ProcName As String = "mSocket_Close"
On Error GoTo Err
mSocket.Close
handleTwsDisconnection "closed by peer", False
Exit Sub
Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocket_Connect()
Const ProcName As String = "mSocket_Connect"
On Error GoTo Err

If mRetryingConnection Then
    ' this means we have been waiting for Tws to start. If we try to start tickers
    ' too soon after Tws accepts socket connections, Tws gives a 'max number of
    ' tickers reached' error. So we'll cancel this connection and connect again in
    ' 10 seconds
    gLog "Tws is now available - connecting again in 10 seconds to allow Tws Time to initialise", ModuleName, ProcName
    mSocket.Close
    Set mConnectionTimer = CreateIntervalTimer(10, ExpiryTimeUnitSeconds)
    mConnectionTimer.StartTimer
    mRetryingConnection = False
    Exit Sub
End If

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
        12, _
        VarPtr(oBuf(0)), _
        0, _
        bytesReturned, _
        ov, _
        AddressOf gWSAIoctlCompletionRoutine)
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
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
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

If gSocketLogger.IsLoggable(LogLevelMediumDetail) Then
    gSocketLogger.Log showBuffer("In buf: ", lSocketInputData), ProcName, ModuleName, LogLevels.LogLevelMediumDetail
End If

mBufferedReader.AddData lSocketInputData

Exit Sub
Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
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
        gLog "Failed to connect to Tws" & _
            IIf(mConnectionRetryIntervalSecs <> 0, " - retrying in " & mConnectionRetryIntervalSecs & " seconds", "") & _
            pDescription & ": " & getConnectionString, _
            ModuleName, ProcName
        
        RaiseEvent ConnectFailed(pDescription, mConnectionRetryIntervalSecs)
        If mConnectionRetryIntervalSecs <> 0 Then retryConnection
    Else
        gLog "Socket error " & pNumber & ": " & _
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
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ConnectionRetryIntervalSecs(ByVal Value As Long)
mConnectionRetryIntervalSecs = Value
If mConnectionRetryIntervalSecs = 0 And mRetryingConnection Then Set mConnectionTimer = Nothing
End Property

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
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
gLog s, ModuleName, ProcName

mSocket.Connect mServer, mPort
mIsConnecting = True

RaiseEvent Connecting

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Disconnect( _
                ByVal pReason As String)
Const ProcName As String = "Disconnect"
On Error GoTo Err

If Not (mIsConnecting Or mIsConnected) Then Exit Sub
mIsConnecting = False
mIsConnected = False

If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer

ReleaseSocket

gLog "Disconnected from: " & getConnectionString & ": " & pReason, ModuleName, ProcName
handleTwsDisconnection "closed by application", True
RaiseEvent Disconnected(pReason)
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pBufferedReader As BufferedReader, _
                ByVal pProgramErrorHandler As IProgramErrorListener)
If pServer = "" Then pServer = "127.0.0.1"
mServer = pServer
mPort = pPort
Set mBufferedReader = pBufferedReader
Set mProgramErrorHandler = pProgramErrorHandler
End Sub

Friend Sub Send(ByRef pData() As Byte)
Const ProcName As String = "Send"
On Error GoTo Err

If gSocketLogger.IsLoggable(LogLevelMediumDetail) Then
    gSocketLogger.Log showBuffer("Out buf: ", pData), ProcName, ModuleName, LogLevels.LogLevelMediumDetail
End If
mSocket.SendData pData

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
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
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getConnectionString() As String
getConnectionString = "server=" & mServer & _
                        " port=" & mPort
End Function

Private Sub handleTwsDisconnection(ByVal pMessage As String, ByVal pClosedByApplication As Boolean)
Const ProcName As String = "handleTwsDisconnection"
On Error GoTo Err

gLog "Connection to Tws closed: " & pMessage & ": " & getConnectionString, ModuleName, ProcName

Set mSocket = Nothing
mIsConnected = False

RaiseEvent ConnectionClosed(pMessage)
If Not pClosedByApplication Then retryConnection

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub ReleaseSocket()
Const ProcName As String = "ReleaseSocket"
On Error GoTo Err

If Not mSocket Is Nothing Then
    mSocket.Close
    Set mSocket = Nothing
    Unload mWinsockForm
    Set mWinsockForm = Nothing
End If

mIsConnected = False

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub retryConnection()
Const ProcName As String = "retryConnection"
On Error GoTo Err

If mConnectionRetryIntervalSecs <> 0 Then
    gLog "Reconnecting in " & mConnectionRetryIntervalSecs & " seconds", ModuleName, ProcName
    Set mConnectionTimer = CreateIntervalTimer(mConnectionRetryIntervalSecs, ExpiryTimeUnitSeconds)
    mConnectionTimer.StartTimer
    mRetryingConnection = True
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function showBuffer( _
                ByRef pHeader As String, _
                ByRef pBuffer() As Byte) As String
Const ProcName As String = "showBuffer"
On Error GoTo Err

showBuffer = pHeader & Replace(StrConv(pBuffer, vbUnicode), Chr$(0), "_")

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function




