VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SessionFuture"
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

Private Const ModuleName                            As String = "SessionFuture"

'@================================================================================
' Member variables
'@================================================================================

Private mSession                                    As New Session

Private mListeners                                  As New EnumerableCollection

Private mIsCancelled                                As Boolean
Private mIsAvailable                                As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
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

Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IFuture_ErrorMessage() As String

End Property

Private Property Get IFuture_ErrorNumber() As Long

End Property

Private Property Get IFuture_ErrorSource() As String

End Property

Private Property Get IFuture_IsAvailable() As Boolean
IFuture_IsAvailable = mIsAvailable
End Property

Private Property Get IFuture_IsCancelled() As Boolean
IFuture_IsCancelled = mIsCancelled
End Property

Private Property Get IFuture_IsFaulted() As Boolean
IFuture_IsFaulted = False
End Property

Private Property Get IFuture_IsPending() As Boolean
IFuture_IsPending = Not (mIsAvailable Or mIsCancelled)
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

Assert Not mIsCancelled, "Session creation was cancelled"
Assert mIsAvailable, "Session creation not yet completed"

Set IFuture_Value = mSession

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get IsAvailable() As Boolean
IsAvailable = mIsAvailable
End Property

Friend Property Get IsCancelled() As Boolean
IsCancelled = mIsCancelled
End Property

Friend Property Get Session() As Session
Set Session = mSession
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

If mIsAvailable Or mIsCancelled Then Exit Sub

mIsCancelled = True
fireStateChange FutureStateCancelled
mListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Complete()
Const ProcName As String = "Complete"
On Error GoTo Err

mIsAvailable = True
fireStateChange FutureStateCompleted
mListeners.Clear

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




