VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timeframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event BarsLoaded()

Event StateChange(ByRef ev As StateChangeEventData)


'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                        As String = "Timeframe"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mState                              As TimeframeStates

Private mKey                                As String

Private mStudyBase                          As IStudyBase

Private mInitialBarsFuture                  As IFuture
Private mBarsFuture                         As IFuture
Private mBarStudy                           As IBarStudy

Private mStateChangeListeners               As Listeners

Private mIsHistorical                       As Boolean

Private mContractFuture                     As IFuture

Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStateChangeListeners = New Listeners
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

fireStateChange TimeframeStateLoaded
RaiseEvent BarsLoaded

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsPending Then Exit Sub
If ev.Future.IsCancelled Then Exit Sub

If ev.Future.IsFaulted Then Err.Raise ev.Future.ErrorNumber, ev.Future.ErrorSource, ev.Future.ErrorMessage

Dim lInitialBars As Bars
Set lInitialBars = mInitialBarsFuture.Value
logAMessage "initial bars loaded: " & lInitialBars.Count, ProcName
logAMessage "starting BarStudy", ProcName
mStudyBase.StudyManager.StartStudy mBarStudy, 0

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get IsHistorical() As Boolean
IsHistorical = mIsHistorical
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Get State() As TimeframeStates
State = mState
End Property

Public Property Get BarsFuture() As IFuture
Set BarsFuture = mBarsFuture
End Property

Public Property Get BarStudy() As IBarStudy
Set BarStudy = mBarStudy
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddStateChangeListener( _
                ByVal pListener As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mInitialBarsFuture Is Nothing Then
    If mInitialBarsFuture.IsPending Then mInitialBarsFuture.Cancel
    If mBarsFuture.IsPending Then mBarsFuture.Cancel
End If

fireStateChange TimeframeStateFinished

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStudyBase As IStudyBase, _
                ByVal pContractFuture As IFuture, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pKey As String, _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pBarType As BarTypes, _
                ByVal pFromTime As Date, _
                ByVal pToTime As Date, _
                ByVal pStateChangeListener As StateChangeListener, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pExcludeCurrentbar As Boolean, _
                ByVal pClockFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

If Not pStateChangeListener Is Nothing Then AddStateChangeListener pStateChangeListener

mKey = pKey
Set mStudyBase = pStudyBase
Set mContractFuture = pContractFuture

Dim lFetchBarsNeeded As Boolean
If pMaxNumberOfBars = 0 And pToTime = 0 Then
ElseIf pMaxNumberOfBars = 0 And pToTime <> 0 Then
    lFetchBarsNeeded = True
    mIsHistorical = True
ElseIf pToTime <> 0 Then
    lFetchBarsNeeded = True
    mIsHistorical = True
Else
    lFetchBarsNeeded = True
End If

Set mBarStudy = CreateBarStudy(pBarTimePeriod, mStudyBase)
Set mBarsFuture = mBarStudy.BarsFuture

If lFetchBarsNeeded And Not mContractFuture Is Nothing And Not pHistDataStore Is Nothing Then
    Set mInitialBarsFuture = pHistDataStore.FetchBars(CreateBarDataSpecifierFuture(pContractFuture, _
                                                                pBarTimePeriod, _
                                                                pFromTime, _
                                                                pToTime, _
                                                                pMaxNumberOfBars, _
                                                                pBarType, _
                                                                pClockFuture, _
                                                                pExcludeCurrentbar, _
                                                                pIncludeBarsOutsideSession), _
                                                    mBarStudy.InitialBarFetchListener)
    mFutureWaiter.Add mInitialBarsFuture
    
    
    fireStateChange TimeframeStateLoading
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener( _
                ByVal pListener As StateChangeListener)
mStateChangeListeners.Remove pListener
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange(ByVal newState As TimeframeStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

If newState = mState Then Exit Sub
mState = newState

If gLogger.IsLoggable(LogLevelDetail) Then
    Dim s As String
    Select Case newState
'    Case TimeframeStateCreated
'        s = s & "created"
    Case TimeframeStateLoading
        s = s & "loading"
    Case TimeframeStateLoaded
        If mInitialBarsFuture Is Nothing Then
            s = s & "loaded without initial bars"
        Else
            s = s & "loaded with initial bars"
        End If
    Case TimeframeStateFinished
        s = s & "finished"
    End Select
    logAMessage s, ProcName
End If

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = newState

mStateChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mStateChangeListeners.Count
    Dim lListener As StateChangeListener
    Set lListener = mStateChangeListeners.GetListener(i)
    lListener.Change ev
Next
RaiseEvent StateChange(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getName() As String
Const ProcName As String = "getName"
On Error GoTo Err

If Not mContractFuture.IsAvailable Then
    getName = "????/" & mKey
Else
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    getName = lContract.Specifier.LocalSymbol & "/" & mKey
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logAMessage(ByVal pMessage As String, ByVal pProcName As String, Optional ByVal pLogLevel As LogLevels = LogLevels.LogLevelDetail)
gLogger.Log "Timeframe (" & getName & ") " & pMessage, pProcName, ModuleName, pLogLevel
End Sub




