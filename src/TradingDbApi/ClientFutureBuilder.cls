VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClientFutureBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ModuleName                            As String = "DBClientFutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mFutureBuilder                              As New FutureBuilder

Private mSelfRef                                    As Object

Private mDBClient                                   As DBClient

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mUseSynchronousReads                        As Boolean
Private mUseSynchronousWrites                       As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' mFutureWaiter Interface Members
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
Else
    Dim lDBClient As New DBClient
    lDBClient.Initialise ev.Future.Value, mUseSynchronousReads, mUseSynchronousWrites
    mFutureBuilder.Value = lDBClient
    mFutureBuilder.Complete
End If

Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Future() As IFuture
Set Future = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pDatabaseType As DatabaseTypes, _
                ByVal pServer As String, _
                ByVal pDatabaseName As String, _
                ByVal pUsername As String, _
                ByVal pPassword As String, _
                ByVal pUseSynchronousReads As Boolean, _
                ByVal pUseSynchronousWrites As Boolean, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

mFutureBuilder.Cookie = pCookie

mUseSynchronousReads = pUseSynchronousReads
mUseSynchronousWrites = pUseSynchronousWrites
mFutureWaiter.Add CreateTradingDBFuture(CreateConnectionParams( _
                                        pDatabaseType, _
                                        pServer, _
                                        pDatabaseName, _
                                        pUsername, _
                                        pPassword), _
                                    pCookie)


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




