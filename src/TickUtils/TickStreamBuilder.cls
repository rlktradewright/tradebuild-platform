VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickStreamBuilder"
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

Event StateChange(ev As StateChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickStreamBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mTickStream                                 As TickStream

Private mStateChangeListeners                       As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mTickStream = New TickStream
mTickStream.State = TickStreamStateCreated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get TickStream() As ITickStream
Set TickStream = mTickStream
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mTickStream.State = TickStreamStateFinished

fireStateChange TickStreamStateFinished
mStateChangeListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStreamId As Long, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

mTickStream.Initialise Me, pStreamId, pContractFuture, pClockFuture
mTickStream.State = TickStreamStateReady
fireStateChange TickStreamStateReady

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyFinished()
Const ProcName As String = "NotifyFinished"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyStart()
Const ProcName As String = "NotifyStart"
On Error GoTo Err

fireStateChange TickStreamStateRunning
mTickStream.State = TickStreamStateRunning

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyTick(pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

If mTickStream.State = TickStreamStateReady Then Exit Sub

Assert mTickStream.State = TickStreamStateRunning, "Tick stream is not running"
mTickStream.NotifyTick pTick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange(ByVal pState As TickStreamStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = pState

Dim lListener As StateChangeListener
For Each lListener In mStateChangeListeners.CurrentListeners
    lListener.Change ev
Next

RaiseEvent StateChange(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


