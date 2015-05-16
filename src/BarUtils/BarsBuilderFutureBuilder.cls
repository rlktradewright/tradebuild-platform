VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarsBuilderFutureBuilder"
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

Private Const ModuleName                            As String = "BarsBuilderFutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mFutureBuilder                              As New FutureBuilder
Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mBarTimePeriod                              As TimePeriod
Private mTickSize                                   As Double
Private mNumberOfBarsToCache                        As Long
Private mSave                                       As Boolean

Private mSession                                    As Session
Private mBars                                       As Bars

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSelfRef = Me
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

Dim lBarsBuilder As New BarsBuilder
If mBars Is Nothing Then
    lBarsBuilder.Initialise mBarTimePeriod, mSession, mTickSize, mNumberOfBarsToCache, mSave
Else
    lBarsBuilder.InitialiseWithInitialBars mBars, mSession, mTickSize
End If
mFutureBuilder.Value = lBarsBuilder
mFutureBuilder.Complete

Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsPending Then Exit Sub

If ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    Set mSelfRef = Nothing
    mFutureWaiter.Clear
ElseIf ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
    Set mSelfRef = Nothing
    mFutureWaiter.Clear
ElseIf ev.Future.IsAvailable Then
    If TypeOf ev.Future.Value Is Session Then
        Set mSession = ev.Future.Value
    ElseIf TypeOf ev.Future.Value Is Bars Then
        Set mBars = ev.Future.Value
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Future() As IFuture
Const ProcName As String = "Future"
On Error GoTo Err

Set Future = mFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSessionFuture As IFuture, _
                ByVal pTickSize As Double, _
                ByVal pNumberOfBarsToCache As Long, _
                ByVal pSave As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mBarTimePeriod = pBarTimePeriod
mTickSize = pTickSize
mNumberOfBarsToCache = pNumberOfBarsToCache
mSave = pSave

mFutureWaiter.Add pSessionFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseWithInitialBars( _
                ByVal pBarsFuture As IFuture, _
                ByVal pSessionFuture As IFuture, _
                ByVal pTickSize As Double)
Const ProcName As String = "InitialiseWithInitialBars"
On Error GoTo Err

mTickSize = pTickSize

mFutureWaiter.Add pBarsFuture
mFutureWaiter.Add pSessionFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




