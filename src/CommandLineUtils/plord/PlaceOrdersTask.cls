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

Private mGroups                                     As Groups

Private mMoneyManager                               As MoneyManager

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
ITask_TaskName = mTaskContext.Name
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

If pStageOrders Then
    gWriteLineToConsole gCommands.StageOrdersCommand.Name & " YES"
End If

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

Friend Sub Initialise( _
                ByVal pGroups As Groups, _
                ByVal pMoneyManager As MoneyManager)
Set mGroups = pGroups
Set mMoneyManager = pMoneyManager
End Sub

Friend Sub RemoveOrderPlacer(ByVal pOrderPlacer As OrderPlacer)
Const ProcName As String = "RemoveOrderPlacer"
On Error GoTo Err

Dim lIndex As Long: lIndex = findEntryIndex(pOrderPlacer)
If lIndex <> NullIndex Then releaseEntry lIndex

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

pContractProcessor.CanSubmitPendingOrders = True

If Not pContractProcessor.IsReady Then Exit Sub

Dim lBOS As BracketOrderSpecification
For Each lBOS In pContractProcessor.BracketOrders
    If mGroups.Item(pContractProcessor.GroupName).OrderPlacers.Contains(lBOS.ID) Then
        mGroups.Item(pContractProcessor.GroupName).OrderPlacers.Remove (lBOS.ID)
    End If
    Dim lOrderPlacer As OrderPlacer: Set lOrderPlacer = New OrderPlacer
    mGroups.Item(pContractProcessor.GroupName).OrderPlacers.Add lOrderPlacer, lBOS.ID
    lOrderPlacer.Initialise _
                    lBOS, _
                    pContractProcessor.Contract, _
                    pContractProcessor.OrderContext, _
                    pStageOrders, _
                    mMoneyManager, _
                    mGroups.Item(pContractProcessor.GroupName)

    Set mOrderPlacers(allocateEntry).OrderPlacer = lOrderPlacer
    
    Dim lEntry As New LiveOrderEntry
    If Not gLiveOrders.TryItem(lBOS.ID, lEntry) Then
        lEntry.Key = lBOS.ID
        gLiveOrders.Add lEntry
    End If
    lEntry.BracketOrderSpec = lBOS
    lEntry.GroupName = pContractProcessor.GroupName
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
    allocateEntryIndex = mNextUnusedIndex
    mNextUnusedIndex = mNextUnusedIndex + 1
    Exit Function
End If
    
ReDim Preserve mOrderPlacers(2 * (UBound(mOrderPlacers) + 1) - 1) As OrderPlacersEntry

allocateEntryIndex = mNextUnusedIndex
mNextUnusedIndex = mNextUnusedIndex + 1
End Function

Private Function findEntryIndex(ByVal pOrderPlacer As OrderPlacer) As Long
Const ProcName As String = "findEntryIndex"
On Error GoTo Err

Dim lIndex As Long: lIndex = mListHeadIndex

Do While lIndex <> NullIndex
    If mOrderPlacers(lIndex).OrderPlacer Is pOrderPlacer Then
        findEntryIndex = lIndex
        Exit Function
    End If
    lIndex = mOrderPlacers(lIndex).Next
Loop

findEntryIndex = NullIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub releaseEntry( _
                ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

If pIndex = mListHeadIndex Then
    mListHeadIndex = mOrderPlacers(pIndex).Next
Else
    Dim lIndex As Long: lIndex = mListHeadIndex
    
    Do While lIndex <> NullIndex
        If mOrderPlacers(lIndex).Next = pIndex Then
            mOrderPlacers(lIndex).Next = mOrderPlacers(pIndex).Next
            Exit Do
        End If
        lIndex = mOrderPlacers(lIndex).Next
    Loop
End If

With mOrderPlacers(pIndex)
    Set .OrderPlacer = Nothing
    .Next = mFirstFreeIndex
End With

mFirstFreeIndex = pIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




