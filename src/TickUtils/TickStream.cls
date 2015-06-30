VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickStream"
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

Implements IGenericTickSource
Implements ITickStream

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

Private Const ModuleName                            As String = "TickStream"

'@================================================================================
' Member variables
'@================================================================================

Private mBuilder                                    As TickStreamBuilder

Private mContractFuture                             As IFuture
Private mClockFuture                                As IFuture

Private mGenericTickListeners                       As New Listeners
Private mStateChangeListeners                       As New Listeners

Private mStreamId                                   As Long

Private mState                                      As TickStreamStates

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickSource Interface Members
'@================================================================================

Private Sub IGenericTickSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IGenericTickSource_AddGenericTickListener"
On Error GoTo Err

AddGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGenericTickSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IGenericTickSource_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITickStream Interface Members
'@================================================================================

Private Sub ITickStream_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "ITickStream_AddGenericTickListener"
On Error GoTo Err

AddGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "ITickStream_AddStateChangeListener"
On Error GoTo Err

AddStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickStream_ClockFuture() As IFuture
Set ITickStream_ClockFuture = ClockFuture
End Property

Private Property Get ITickStream_ContractFuture() As IFuture
Set ITickStream_ContractFuture = ContractFuture
End Property

Private Sub ITickStream_Finish()
Const ProcName As String = "ITickStream_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "ITickStream_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "ITickStream_RemoveStateChangeListener"
On Error GoTo Err

RemoveStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_Start()
Const ProcName As String = "ITickStream_Start"
On Error GoTo Err

Start

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickStream_State() As TickStreamStates
ITickStream_State = State
End Property

Private Property Get ITickStream_StreamId() As Long
ITickStream_StreamId = StreamId
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ClockFuture() As IFuture
Set ClockFuture = mClockFuture
End Property

Public Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Friend Property Let State(ByVal Value As TickStreamStates)
Const ProcName As String = "State"
On Error GoTo Err

If Value = mState Then Exit Property
mState = Value

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = mState

mStateChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mStateChangeListeners.Count
    Dim lListener As IStateChangeListener
    Set lListener = mStateChangeListeners.GetListener(i)
    lListener.Change ev
Next

RaiseEvent StateChange(ev)

If mState = TickStreamStateFinished Then
    mStateChangeListeners.Clear
    mGenericTickListeners.Clear
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get State() As TickStreamStates
State = mState
End Property

Public Property Get StreamId() As Long
StreamId = mStreamId
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
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

If mState = TickStreamStateFinished Then Exit Sub
mBuilder.Finish
Set mBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBuilder As TickStreamBuilder, _
                ByVal pStreamId As Long, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

mState = TickStreamStateCreated
Set mBuilder = pBuilder
mStreamId = pStreamId
Set mContractFuture = pContractFuture
Set mClockFuture = pClockFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NoMoreTicks()
Const ProcName As String = "NoMoreTicks"
On Error GoTo Err

If mState = TickStreamStateFinished Then Exit Sub
mState = TickStreamStateFinished

Dim ev As GenericTickEventData
ev.StreamId = StreamId
ev.Tick.Timestamp = MaxDate
ev.Tick.TickType = -1
Set ev.Source = Me

mGenericTickListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mGenericTickListeners.Count
    Dim lListener As IGenericTickListener
    Set lListener = mGenericTickListeners.GetListener(i)
    lListener.NoMoreTicks ev
Next

State = TickStreamStateFinished

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

If mState = TickStreamStateFinished Then Exit Sub

Dim ev As GenericTickEventData
ev.StreamId = StreamId
ev.Tick = pTick
Set ev.Source = Me

mGenericTickListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mGenericTickListeners.Count
    Dim lListener As IGenericTickListener
    Set lListener = mGenericTickListeners.GetListener(i)
    lListener.NotifyTick ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Start()
Const ProcName As String = "Start"
On Error GoTo Err

Assert mState = TickStreamStateReady, "Tick stream is already started or is finished"
mBuilder.NotifyStart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




