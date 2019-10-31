VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlaceOrdersTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type OrderPlacersEntry
    Next                As Long
    OrderPlacer         As OrderPlacer
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "PlaceOrdersTask"

Private Const NullIndex                             As Long = -1

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext
Private mOrderPlacers()                             As OrderPlacersEntry
Private mNextUnusedIndex                            As Long
Private mFirstFreeIndex                             As Long

Private mListHeadIndex                              As Long
Private mListTailIndex                              As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mOrderPlacers(7) As OrderPlacersEntry
mListHeadIndex = NullIndex
mFirstFreeIndex = NullIndex
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mListHeadIndex = NullIndex Then
    mTaskContext.Suspend -1
    Exit Sub
End If

Dim lOrderPlacer As OrderPlacer
Set lOrderPlacer = mOrderPlacers(mListHeadIndex).OrderPlacer
releaseEntry mListHeadIndex

If Not lOrderPlacer.PlaceOrder Then
    ' the order context is not ready enough, so the
    ' orderPlacer will now listen for changes to its readiness
Else
    mTaskContext.Suspend 25
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddContractProcessor( _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pStageOrders As Boolean)
Const ProcName As String = "AddContractProcessor"
On Error GoTo Err

addContractProc pContractProcessor, pStageOrders

mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddContractProcessors( _
                ByVal pContractProcessors As EnumerableCollection, _
                ByVal pStageOrders As Boolean)
Const ProcName As String = "AddContractProcessors"
On Error GoTo Err

Dim lContractProcessor As ContractProcessor
For Each lContractProcessor In pContractProcessors
    addContractProc lContractProcessor, pStageOrders
Next

mTaskContext.Continue

gWriteLineToStdOut StageOrdersCommand & " " & IIf(pStageOrders, "YES", "NO")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddOrderPlacer(ByVal pOrderPlacer As OrderPlacer)
Const ProcName As String = "AddOrderPlacer"
On Error GoTo Err

Set mOrderPlacers(allocateEntry).OrderPlacer = pOrderPlacer

mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addContractProc( _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pStageOrders As Boolean)
Const ProcName As String = "addContractProc"
On Error GoTo Err

Dim lBOS As BracketOrderSpecification
For Each lBOS In pContractProcessor.BracketOrders
    Dim lOrderPlacer As OrderPlacer
    Set lOrderPlacer = New OrderPlacer
    lOrderPlacer.Initialise lBOS, pContractProcessor.Contract, pContractProcessor.OrderContext, pStageOrders
    Set mOrderPlacers(allocateEntry).OrderPlacer = lOrderPlacer
Next
pContractProcessor.ClearOrders

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function allocateEntry() As Long
Dim lIndex As Long
lIndex = allocateEntryIndex

With mOrderPlacers(lIndex)
    .Next = NullIndex
    Set .OrderPlacer = Nothing
End With

If mListHeadIndex = NullIndex Then
    mListHeadIndex = lIndex
    mListTailIndex = lIndex
Else
    mOrderPlacers(mListTailIndex).Next = lIndex
    mListTailIndex = lIndex
End If

allocateEntry = lIndex
End Function

Private Function allocateEntryIndex() As Long
Const ProcName As String = "allocateEntryIndex"

If mFirstFreeIndex <> NullIndex Then
    allocateEntryIndex = mFirstFreeIndex
    mFirstFreeIndex = mOrderPlacers(mFirstFreeIndex).Next
    Exit Function
End If

If mNextUnusedIndex > UBound(mOrderPlacers) Then
    ReDim Preserve mOrderPlacers(1 To 2 * UBound(mOrderPlacers)) As OrderPlacersEntry
Else
    allocateEntryIndex = mNextUnusedIndex
    mNextUnusedIndex = mNextUnusedIndex + 1
    Exit Function
End If

allocateEntryIndex = mNextUnusedIndex
mNextUnusedIndex = mNextUnusedIndex + 1
End Function

Private Sub releaseEntry( _
                ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

If pIndex = mListHeadIndex Then mListHeadIndex = mOrderPlacers(pIndex).Next

With mOrderPlacers(pIndex)
    Set .OrderPlacer = Nothing
    .Next = mFirstFreeIndex
End With

mFirstFreeIndex = pIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




