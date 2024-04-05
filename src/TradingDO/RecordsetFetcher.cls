VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecordsetFetcher"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "RecordsetFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                         As TradingDB

Private mCommand                                    As ADODB.Command
Private mCommandType                                As CommandTypeEnum
Private mCursorType                                 As CursorTypeEnum

Private mConnector                                  As Connector

Private WithEvents mConnection                      As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRS                              As ADODB.Recordset
Attribute mRS.VB_VarHelpID = -1

Private mAsynchronous                               As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mFutureBuilder                              As FutureBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@===============================================================================
' mConnection Event Handlers
'@===============================================================================

Private Sub mConnection_ExecuteComplete( _
                ByVal RecordsAffected As Long, _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pCommand As ADODB.Command, _
                ByVal pRecordset As ADODB.Recordset, _
                ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnection_ExecuteComplete"
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    Dim s As String
    s = "Fetch error " & vbCrLf & _
        GenerateErrorMessage(pError) & _
        GenerateConnectionErrorMessages(pConnection)
        
    If mAsynchronous Then
        mFutureBuilder.Fail pError.Number, s, ProjectName & "." & ModuleName & ":" & ProcName
    Else
        Err.Raise pError.Number, , s
    End If
    Exit Sub
End If

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@===============================================================================
' mFutureWaiter Event Handlers
'@===============================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf ev.Future.IsAvailable Then
    Set mConnection = ev.Future.Value
    Set mRS = Execute(mConnection, mCommand, mCommandType, mCursorType, mAsynchronous)
End If

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' mRs Event Handlers
'@===============================================================================

Private Sub mRs_FetchComplete( _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Const ProcName As String = "mRs_FetchComplete"
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    Dim s As String
    s = "Fetch error " & vbCrLf & _
        GenerateErrorMessage(pError) & _
        GenerateConnectionErrorMessages(mConnection)
    
    Debug.Print "RecordsetFetcher: close connection(error)"
    If IsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    If mAsynchronous Then
        mFutureBuilder.Fail pError.Number, _
                            s, _
                            ProjectName & "." & ModuleName & ":" & ProcName
    Else
        Err.Raise pError.Number, , s
    End If
    Exit Sub
End If

If mAsynchronous Then
    mFutureBuilder.Value = mRS
    mFutureBuilder.Complete
End If

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function FetchRecordset( _
                ByVal pCommand As ADODB.Command, _
                ByVal pCommandType As CommandTypeEnum, _
                ByVal pCursorType As CursorTypeEnum) As Recordset
Const ProcName As String = "FetchRecordset"
On Error GoTo Err

Set mCommand = pCommand
mCommandType = pCommandType
mCursorType = pCursorType

If mConnector Is Nothing Then
    Set mConnector = mDB.GetConnector
    Set mConnection = mConnector.Connect
End If
Set FetchRecordset = Execute(mConnection, mCommand, mCommandType, mCursorType, False)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function FetchRecordsetAsync( _
                ByVal pCommand As ADODB.Command, _
                ByVal pCommandType As CommandTypeEnum, _
                ByVal pCursorType As CursorTypeEnum) As IFuture
Const ProcName As String = "FetchRecordsetAsync"
On Error GoTo Err

mAsynchronous = True
Set mFutureBuilder = New FutureBuilder

Set mCommand = pCommand
mCommandType = pCommandType
mCursorType = pCursorType

If mConnector Is Nothing Then
    Set mConnector = mDB.GetConnector
    mFutureWaiter.Add mConnector.ConnectFuture(True)
Else
    Set mRS = Execute(mConnection, mCommand, mCommandType, mCursorType, mAsynchronous)
End If

Set FetchRecordsetAsync = mFutureBuilder.Future

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise(ByVal pDB As TradingDB)
Set mDB = pDB
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function Execute( _
                ByVal pConnection As ADODB.Connection, _
                ByVal pCmd As ADODB.Command, _
                ByVal pCommandType As CommandTypeEnum, _
                ByVal pCursorType As CursorTypeEnum, _
                ByVal pAsync As Boolean) As Recordset
Const ProcName As String = "Execute"
On Error GoTo Err

'GTDO.Logger.Log "RecordsetFetcher initiating fetch for " & mInstrument.ShortName, ProcName, ModuleName

pCmd.ActiveConnection = pConnection

Dim lRS As New ADODB.Recordset
lRS.LockType = adLockReadOnly
lRS.cursorlocation = adUseClient

If pAsync Then
    lRS.Properties("Initial Fetch Size") = 0
    lRS.Open pCmd, _
                , _
                pCursorType, _
                , _
                pCommandType + _
                ExecuteOptionEnum.adAsyncFetchNonBlocking + _
                ExecuteOptionEnum.adAsyncExecute
Else
    lRS.Open pCmd, _
                , _
                pCursorType, _
                , _
                pCommandType
End If

Set Execute = lRS

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function





