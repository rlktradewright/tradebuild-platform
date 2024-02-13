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

Private mTickSize                                   As Double

Private mSession                                    As Session
Private mBars                                       As Bars

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    mFutureWaiter.Clear
ElseIf ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
    mFutureWaiter.Clear
ElseIf TypeOf ev.Future.Value Is Bars Then
    gLogger.Log "Got bars", ProcName, ModuleName, LogLevelHighDetail
    Set mBars = ev.Future.Value

    Dim lBarsBuilder As New BarsBuilder
    lBarsBuilder.InitialiseWithInitialBars mBars, mSession, mTickSize
    
    gLogger.Log "BarsBuilder future builder completed", ProcName, ModuleName, LogLevelHighDetail
    mFutureBuilder.Value = lBarsBuilder
    mFutureBuilder.Complete
    
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
BarsGlobals.gNotifyUnhandledError ProcName, ModuleName
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
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                ByVal pTickSize As Double, _
                ByVal pNumberOfBarsToCache As Long, _
                ByVal pNormaliseDailyTimestamps As Boolean, _
                ByVal pSave As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Dim lBarsBuilder As New BarsBuilder
lBarsBuilder.Initialise pBarTimePeriod, pSession, pTickSize, pNumberOfBarsToCache, pNormaliseDailyTimestamps, pSave

gLogger.Log "BarsBuilder future builder completed", ProcName, ModuleName, LogLevelHighDetail
mFutureBuilder.Value = lBarsBuilder
mFutureBuilder.Complete

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseWithInitialBars( _
                ByVal pBarsFuture As IFuture, _
                ByVal pSession As Session, _
                ByVal pTickSize As Double)
Const ProcName As String = "InitialiseWithInitialBars"
On Error GoTo Err

mTickSize = pTickSize

mFutureWaiter.Add pBarsFuture
Set mSession = pSession
Set mSelfRef = Me

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




