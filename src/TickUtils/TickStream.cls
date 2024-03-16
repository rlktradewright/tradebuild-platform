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
Private mOptionComputationListeners                 As New Listeners
Private mStateChangeListeners                       As New Listeners

Private mStreamId                                   As Long

Private mState                                      As TickStreamStates

Private mInitGenericTickListeners                   As Boolean
Private mCurrentGenericTickListeners()              As Object
Private mSomeGenericTickListeners                   As Boolean

Private mInitOptionComputationListeners             As Boolean
Private mCurrentOptionComputationListeners()        As Object
Private mSomeOptionComputationListeners             As Boolean

Private mMinimumPriceVariation                      As Double

Private mIsDelayed                                  As Boolean

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
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGenericTickSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IGenericTickSource_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
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
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_AddOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "ITickStream_AddOptionComputationListener"
On Error GoTo Err

AddOptionComputationListener pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "ITickStream_AddStateChangeListener"
On Error GoTo Err

AddStateChangeListener pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
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
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickStream_IsDelayed() As Boolean
ITickStream_IsDelayed = mIsDelayed
End Property

Private Property Get ITickStream_MinimumPriceVariation() As Double
ITickStream_MinimumPriceVariation = MinimumPriceVariation
End Property

Private Sub ITickStream_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "ITickStream_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_RemoveOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "ITickStream_RemoveOptionComputationListener"
On Error GoTo Err

RemoveOptionComputationListener pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "ITickStream_RemoveStateChangeListener"
On Error GoTo Err

RemoveStateChangeListener pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_Start()
Const ProcName As String = "ITickStream_Start"
On Error GoTo Err

Start

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
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

Public Property Get MinimumPriceVariation() As Double
Const ProcName As String = "MinimumPriceVariation"
On Error GoTo Err

MinimumPriceVariation = GetMinimumPriceVariation
Assert MinimumPriceVariation > 0, "MinimumPriceVariation has not been set"

Exit Property

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let State(ByVal Value As TickStreamStates)
Const ProcName As String = "State"
On Error GoTo Err

If Value = mState Then Exit Property
mState = Value

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = mState

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mStateChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mStateChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

RaiseEvent StateChange(ev)

If mState = TickStreamStateFinished Then
    mStateChangeListeners.Clear
    mGenericTickListeners.Clear
    mOptionComputationListeners.Clear
End If

Exit Property

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
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
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "AddOptionComputationListener"
On Error GoTo Err

mOptionComputationListeners.Add pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If mState = TickStreamStateFinished Then Exit Sub
mBuilder.Finish
Set mBuilder = Nothing

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetMinimumPriceVariation() As Double
GetMinimumPriceVariation = mMinimumPriceVariation
End Function

Friend Sub Initialise( _
                ByVal pBuilder As TickStreamBuilder, _
                ByVal pStreamId As Long, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pIsDelayed As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mState = TickStreamStateCreated
Set mBuilder = pBuilder
mStreamId = pStreamId
Set mContractFuture = pContractFuture
Set mClockFuture = pClockFuture
mIsDelayed = pIsDelayed

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NoMoreTicks()
Const ProcName As String = "NoMoreTicks"
On Error GoTo Err

If mState = TickStreamStateFinished Then Exit Sub

updateCurrentGenericTickListeners
If mSomeGenericTickListeners Then
    Dim ev As GenericTickEventData
    ev.StreamId = StreamId
    ev.Tick.Timestamp = MaxDate
    ev.Tick.TickType = -1
    Set ev.Source = Me
    
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentGenericTickListeners)
        Set lListener = mCurrentGenericTickListeners(i)
        lListener.NoMoreTicks ev
    Next
End If


updateCurrentOptionComputationListeners
If mSomeOptionComputationListeners Then
    Dim ev1 As OptionComputationEventData
    ev1.StreamId = StreamId
    ev1.Computation.ComputationType = OptionComputationTypeNone
    Set ev1.Source = Me
    Dim lListener1 As IOptionComputationListener
    For i = 0 To UBound(mCurrentOptionComputationListeners)
        Set lListener1 = mCurrentOptionComputationListeners(i)
        lListener1.NoMoreComputations ev1
    Next
End If

State = TickStreamStateFinished

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyOptionComputation(ByRef pComputation As OptionComputation)
Const ProcName As String = "NotifyOptionComputation"
On Error GoTo Err

If mState = TickStreamStateFinished Then
    Exit Sub
End If

updateCurrentOptionComputationListeners

If Not mSomeOptionComputationListeners Then
    Exit Sub
End If

Dim ev As OptionComputationEventData
ev.StreamId = StreamId
ev.Computation = pComputation
Set ev.Source = Me

Dim lListener As IOptionComputationListener
Dim i As Long
For i = 0 To UBound(mCurrentOptionComputationListeners)
    Set lListener = mCurrentOptionComputationListeners(i)
    lListener.NotifyComputation ev
Next

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

If mState = TickStreamStateFinished Then Exit Sub
updateCurrentGenericTickListeners

If Not mSomeGenericTickListeners Then Exit Sub

Dim ev As GenericTickEventData
ev.StreamId = StreamId
ev.Tick = pTick
Set ev.Source = Me

Dim lListener As IGenericTickListener
Dim i As Long
For i = 0 To UBound(mCurrentGenericTickListeners)
    Set lListener = mCurrentGenericTickListeners(i)
    lListener.NotifyTick ev
Next

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Remove pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "RemoveOptionComputationListener"
On Error GoTo Err

mOptionComputationListeners.Remove pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetMinimumPriceVariation(ByVal Value As Double)
mMinimumPriceVariation = Value
End Sub

Public Sub Start()
Const ProcName As String = "Start"
On Error GoTo Err

Assert mState = TickStreamStateReady, "Tick stream is already started or is finished"

State = TickStreamStateRunning
mBuilder.NotifyStart

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub updateCurrentGenericTickListeners()
If Not mInitGenericTickListeners Or Not mGenericTickListeners.Valid Then
    mInitGenericTickListeners = True
    mSomeGenericTickListeners = mGenericTickListeners.GetCurrentListeners(mCurrentGenericTickListeners)
End If
End Sub

Private Sub updateCurrentOptionComputationListeners()
If Not mInitOptionComputationListeners Or Not mOptionComputationListeners.Valid Then
    mInitOptionComputationListeners = True
    mSomeOptionComputationListeners = mOptionComputationListeners.GetCurrentListeners(mCurrentOptionComputationListeners)
End If
End Sub
