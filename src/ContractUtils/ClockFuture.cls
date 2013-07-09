VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClockFuture"
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
Implements StateChangeListener

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

Private Const ModuleName                            As String = "ClockFuture"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private mContractFuture                             As IFuture
Private mIsSimulated                                As Boolean
Private mClockRate                                  As Single

Private mClock                                      As Clock

Private mIsCancelled                                As Boolean
Private mIsAvailable                                As Boolean
Private mIsFaulted                                  As Boolean
Private mErrorNumber                                As Long
Private mErrorMessage                               As String
Private mErrorSource                                As String

Private mListeners                                  As New EnumerableCollection

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

Assert Not mIsCancelled, "Clock creation was cancelled"
Assert Not mIsFaulted, "Clock creation failed: error=" & mErrorNumber & "; description=" & mErrorMessage & "; source=" & mErrorSource
Assert mIsAvailable, "Clock creation has not yet completed"

Set IFuture_Value = mClock

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Dim lState As FutureStates
lState = ev.State

If lState <> FutureStatePending Then handleContract

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsSimulated As Boolean, _
                ByVal pClockRate As Single)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractFuture = pContractFuture
mIsSimulated = pIsSimulated

If mContractFuture.IsPending Then
    mContractFuture.AddStateChangeListener Me
Else
    handleContract
    mListeners.Clear
End If


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

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

Private Sub handleContract()
Const ProcName As String = "handleContract"
On Error GoTo Err

If mContractFuture.IsAvailable Then
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    If mIsSimulated Then
        Set mClock = CreateSimulatedClock(mClockRate, lContract.TimezoneName)
    Else
        Set mClock = GetClock(lContract.TimezoneName)
    End If
    fireStateChange FutureStateCompleted
ElseIf mContractFuture.IsCancelled Then
    mIsCancelled = True
    fireStateChange FutureStateCancelled
ElseIf mContractFuture.IsFaulted Then
    mIsFaulted = True
    mErrorNumber = ErrorCodes.ErrIllegalStateException
    mErrorMessage = "ContractFuture failed to complete"
    mErrorSource = mContractFuture.ErrorSource
    fireStateChange FutureStateErrored
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


