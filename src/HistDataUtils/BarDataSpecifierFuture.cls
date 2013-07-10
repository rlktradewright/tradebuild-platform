VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarDataSpecifierFuture"
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

Implements IFuture

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

Private Const ModuleName                            As String = "BarDataSpecifierFuture"

'@================================================================================
' Member variables
'@================================================================================

Private mContractFuture                             As IFuture
Private mClockFuture                                As IFuture

Private mBarSpecifier                               As New BarDataSpecifier

Private mExcludeCurrentBar                          As Boolean

Private mIsCancelled                                As Boolean
Private mIsAvailable                                As Boolean
Private mIsFaulted                                  As Boolean
Private mErrorNumber                                As Long
Private mErrorMessage                               As String
Private mErrorSource                                As String

Private mListeners                                  As New EnumerableCollection

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IFuture Interface Members
'@================================================================================

Private Sub IFuture_AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IFuture_AddStateChangeListener"
On Error GoTo Err

mListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IFuture_Cancel()
Const ProcName As String = "IFuture_Cancel"
On Error GoTo Err

If Not IFuture_IsPending Then Exit Sub

mContractFuture.RemoveStateChangeListener Me
If Not mClockFuture Is Nothing Then mClockFuture.RemoveStateChangeListener Me
mIsCancelled = True
fireStateChange FutureStateCancelled
mListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IFuture_ErrorMessage() As String
IFuture_ErrorMessage = mErrorMessage
End Property

Private Property Get IFuture_ErrorNumber() As Long
IFuture_ErrorNumber = mErrorNumber
End Property

Private Property Get IFuture_ErrorSource() As String
IFuture_ErrorSource = mErrorSource
End Property

Private Property Get IFuture_IsAvailable() As Boolean
IFuture_IsAvailable = mIsAvailable
End Property

Private Property Get IFuture_IsCancelled() As Boolean
IFuture_IsCancelled = mIsCancelled
End Property

Private Property Get IFuture_IsFaulted() As Boolean
IFuture_IsFaulted = mIsFaulted
End Property

Private Property Get IFuture_IsPending() As Boolean
IFuture_IsPending = Not (mIsAvailable Or mIsCancelled Or mIsFaulted)
End Property

Private Sub IFuture_RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IFuture_RemoveStateChangeListener"
On Error GoTo Err

mListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IFuture_Value() As Object
Const ProcName As String = "IFuture_Value"
On Error GoTo Err

Assert Not mIsCancelled, "BarDataSpecifier creation was cancelled"
Assert Not mIsFaulted, "BarDataSpecifier creation failed: error=" & mErrorNumber & "; description=" & mErrorMessage & "; source=" & mErrorSource
Assert mIsAvailable, "BarDataSpecifier creation has not yet completed"

Set IFuture_Value = mBarSpecifier

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

    setupSpecifier
    fireStateChange FutureStateCompleted
    mListeners.Clear

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not checkFuture(ev.Future) Then
    mFutureWaiter.Clear
    mListeners.Clear
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pToTime As Date, _
                ByVal pFromTime As Date, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pBarType As BarTypes, _
                ByVal pClockFuture As IFuture, _
                ByVal pExcludeCurrentBar As Boolean, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pCustomSessionStartTime As Date, _
                ByVal pCustomSessionEndTime As Date)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractFuture = pContractFuture
Set mClockFuture = pClockFuture

mBarSpecifier.BarTimePeriod = pBarTimePeriod
mBarSpecifier.BarType = pBarType
mBarSpecifier.FromDate = pFromTime
mBarSpecifier.ToDate = pToTime
mBarSpecifier.SetCustomSessionTimes pCustomSessionStartTime, pCustomSessionEndTime
mBarSpecifier.IncludeBarsOutsideSession = pIncludeBarsOutsideSession
mBarSpecifier.MaxNumberOfBars = pMaxNumberOfBars
mExcludeCurrentBar = pExcludeCurrentBar

mFutureWaiter.Add mContractFuture
If Not mClockFuture Is Nothing Then mFutureWaiter.Add mClockFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function checkFuture(ByVal pFuture As IFuture) As Boolean
If pFuture.IsCancelled Then
    mIsCancelled = True
    checkFuture = False
    fireStateChange FutureStateCancelled
ElseIf pFuture.IsFaulted Then
    mIsFaulted = True
    mErrorNumber = ErrorCodes.ErrIllegalStateException
    mErrorMessage = "BarDataSpecifierFuture failed to complete"
    mErrorSource = pFuture.ErrorSource
    checkFuture = False
    fireStateChange FutureStateErrored
Else
    checkFuture = True
End If
End Function

Private Sub fireStateChange(ByVal pState As FutureStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData

Set ev.Source = Me
ev.State = pState

Dim lListener As StateChangeListener
For Each lListener In mListeners
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupSpecifier()
Const ProcName As String = "setupSpecifier"
On Error GoTo Err

mBarSpecifier.Contract = mContractFuture.Value

If mBarSpecifier.ToDate = 0 Then
    Dim lClock As Clock
    If Not mClockFuture Is Nothing Then
        Set lClock = mClockFuture.Value
    Else
        Set lClock = GetClock(mBarSpecifier.Contract.TimezoneName)
    End If
    
    Dim lTimestamp As Date
    lTimestamp = lClock.TimeStamp
    
    Dim lSessionStartTime As Date
    lSessionStartTime = IIf(mBarSpecifier.CustomSessionStartTime <> 0, mBarSpecifier.CustomSessionStartTime, mBarSpecifier.Contract.SessionStartTime)
    
    Dim lSessionEndTime As Date
    lSessionEndTime = IIf(mBarSpecifier.CustomSessionEndTime <> 0, mBarSpecifier.CustomSessionEndTime, mBarSpecifier.Contract.SessionEndTime)
    
    Dim lSessionTimes As SessionTimes
    lSessionTimes = GetSessionTimes(lTimestamp, lSessionStartTime, lSessionEndTime)
    If mExcludeCurrentBar Then
        mBarSpecifier.ToDate = BarStartTime(lTimestamp, _
                                    mBarSpecifier.BarTimePeriod, _
                                    lSessionTimes.StartTime)
    Else
        ' ensure we get the latest bar
        mBarSpecifier.ToDate = BarEndTime(lTimestamp, _
                                    mBarSpecifier.BarTimePeriod, _
                                    lSessionTimes.StartTime, _
                                    lSessionTimes.EndTime)
    End If
End If

mIsAvailable = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




