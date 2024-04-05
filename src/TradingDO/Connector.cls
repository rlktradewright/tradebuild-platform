VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connector"
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

Event Connected( _
                ByVal pConnection As Connection)

Event ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "Connector"

Private Const ConnectCompletionTimeoutMillisecs                                As Long = 5000

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                 As TradingDB

Private WithEvents mConnection              As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private mConnected                          As Boolean

' need the following because sometimes when running async under heavy Load,
' ADO appears not to notify the connection completion event. So need to retry
' the connection if we don't get the event within a certain time.
Private WithEvents mConnectionTLI           As TimerListItem
Attribute mConnectionTLI.VB_VarHelpID = -1

Private mCursorlocation                     As CursorLocationEnum

Private mCancelled                          As Boolean

Private mFutureBuilder                      As FutureBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
On Error Resume Next
If Not mConnection Is Nothing Then
    If IsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
End If
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@===============================================================================
' mConnection Event Handlers
'@===============================================================================

Private Sub mConnection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnection_ConnectComplete"
On Error GoTo Err

If mCancelled Then Exit Sub

If Not mConnectionTLI Is Nothing Then mConnectionTLI.Cancel
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    GTDO.Logger.Log "Connector: connection error occurred" & vbCrLf & _
                                GenerateErrorMessage(pError) & _
                                GenerateConnectionErrorMessages(pConnection), _
                ProcName, ModuleName, LogLevelSevere
    Dim s As String
    s = ModuleName & ": connection error " & vbCrLf & _
        GenerateErrorMessage(pError) & _
        GenerateConnectionErrorMessages(pConnection)
    RaiseEvent ConnectFailed(pError.Number, s)
    If Not mFutureBuilder Is Nothing Then
        mFutureBuilder.Fail pError.Number, s, ProjectName & "." & ModuleName & ":" & ProcName
    End If
Else
    GTDO.Logger.Log "Connector: connection succeeded", ProcName, ModuleName, LogLevelHighDetail
    mConnected = True
    RaiseEvent Connected(mConnection)
    If Not mFutureBuilder Is Nothing Then
        mFutureBuilder.Value = mConnection
        mFutureBuilder.Complete
    End If
End If

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mConnection_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnection_InfoMessage"
On Error GoTo Err

GTDO.Logger.Log "Connector: InfoMessage event received" & vbCrLf & _
                            GenerateErrorMessage(pError) & _
                            GenerateConnectionErrorMessages(pConnection), _
            ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' mConnectionTLI Event Handlers
'@===============================================================================

Private Sub mConnectionTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mConnectionTLI_StateChange"
On Error GoTo Err

If mConnected Or IsStateSet(mConnection.State, adStateOpen) Then Exit Sub
If ev.State <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub

GTDO.Logger.Log "Connect timed out: cancelling connection", ProcName, ModuleName, LogLevelNormal

' note that mConnection.Cancel causes mConnection_ConnectComplete to be called
' immediately (ie synchronously)
mCancelled = True
mConnection.Cancel
mCancelled = False

Set mConnection = Nothing

GTDO.Logger.Log "reconnecting", ProcName, ModuleName, LogLevelNormal
connectToDB True, True

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Connection() As ADODB.Connection
Const ProcName As String = "connection"

On Error GoTo Err

Assert mConnected, "Connection not yet established or connection failed"

Set Connection = mConnection

Exit Property

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Connect() As Connection
Const ProcName As String = "Connect"
On Error GoTo Err

connectToDB False, False
Set Connect = mConnection

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ConnectAsync(Optional ByVal pUseTimeout As Boolean = True)
Const ProcName As String = "ConnectAsync"
On Error GoTo Err

connectToDB True, pUseTimeout

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ConnectFuture(Optional ByVal pUseTimeout As Boolean = True) As IFuture
Const ProcName As String = "ConnectFuture"
On Error GoTo Err

connectToDB True, pUseTimeout
Set mFutureBuilder = New FutureBuilder
Set ConnectFuture = mFutureBuilder.Future

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal cursorlocation As CursorLocationEnum)
Set mDB = db
mCursorlocation = cursorlocation
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub connectToDB(ByVal pAsync As Boolean, ByVal pUseTimeout As Boolean)
Const ProcName As String = "connectToDB"
Dim failpoint As String
On Error GoTo Err

If Not mConnection Is Nothing Then
    failpoint = "100"
    If IsStateSet(mConnection.State, adStateOpen) Then
        failpoint = "200"
        mConnection.Close
    Else
        failpoint = "300"
        mConnection.Cancel
    End If
    Set mConnection = Nothing
End If

failpoint = "400"
Set mConnection = New ADODB.Connection

failpoint = "500"
mConnection.cursorlocation = mCursorlocation

failpoint = "600"
mConnection.Open mDB.ConnectionString, , , IIf(pAsync, adAsyncConnect, -1)

failpoint = "700"
If pAsync And pUseTimeout Then Set mConnectionTLI = GetGlobalTimerList.Add(Empty, ConnectCompletionTimeoutMillisecs, ExpiryTimeUnitMilliseconds)

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName, failpoint
End Sub

