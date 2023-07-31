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

If mTickStream.State = TickStreamStateFinished Then Exit Sub

mTickStream.NoMoreTicks
mTickStream.State = TickStreamStateFinished

fireStateChange TickStreamStateFinished
mStateChangeListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsMinimumPriceVariationSet() As Boolean
IsMinimumPriceVariationSet = (mTickStream.GetMinimumPriceVariation <> 0)
End Function

Friend Sub Initialise( _
                ByVal pStreamId As Long, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pIsDelayed As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mTickStream.Initialise Me, pStreamId, pContractFuture, pClockFuture, pIsDelayed
mTickStream.State = TickStreamStateReady
fireStateChange TickStreamStateReady

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyOptionComputation(pComputation As OptionComputation)
Const ProcName As String = "NotifyOptionComputation"
On Error GoTo Err

If mTickStream.State = TickStreamStateReady Then Exit Sub

Assert mTickStream.State = TickStreamStateRunning, "Tick stream is not running"

If pComputation.ComputationType = OptionComputationTypeModel Then
    Dim lTick As GenericTick
    lTick.Timestamp = pComputation.Timestamp
    
    lTick.TickType = TickTypeOptionModelPrice
    lTick.Price = pComputation.OptionPrice
    mTickStream.NotifyTick lTick

    lTick.TickType = TickTypeOptionModelImpliedVolatility
    lTick.Price = pComputation.ImpliedVolatility
    mTickStream.NotifyTick lTick

    lTick.TickType = TickTypeOptionModelDelta
    lTick.Price = pComputation.Delta
    mTickStream.NotifyTick lTick

    lTick.TickType = TickTypeOptionModelGamma
    lTick.Price = pComputation.Gamma
    mTickStream.NotifyTick lTick

    lTick.TickType = TickTypeOptionModelVega
    lTick.Price = pComputation.Vega
    mTickStream.NotifyTick lTick

    lTick.TickType = TickTypeOptionModelTheta
    lTick.Price = pComputation.Theta
    mTickStream.NotifyTick lTick

    lTick.TickType = TickTypeOptionModelUnderlyingPrice
    lTick.Price = pComputation.UnderlyingPrice
    mTickStream.NotifyTick lTick

End If

mTickStream.NotifyOptionComputation pComputation

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyStart()
Const ProcName As String = "NotifyStart"
On Error GoTo Err

fireStateChange TickStreamStateRunning

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

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetMinimumPriceVariation(ByVal Value As Double)
Const ProcName As String = "SetMinimumPriceVariation"
On Error GoTo Err

Assert mTickStream.GetMinimumPriceVariation = 0#, "MinimumPriceVariation has already been set"
Assert Value > 0, "MinimumPriceVariation must be > 0"

mTickStream.SetMinimumPriceVariation Value

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


