VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CloseoutOrderContextTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IStateChangeListener
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "CloseoutOrderContextTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mBracketOrdersToCloseout                    As New EnumerableCollection

Private mOrderContextName                             As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
Set lBracketOrder = ev.Source

Dim lState As BracketOrderStates
lState = ev.State

If lState = BracketOrderStateClosed Then
    mBracketOrdersToCloseout.Remove lBracketOrder
    If mBracketOrdersToCloseout.Count = 0 Then
        GOrderUtils.Log "Closeout completed for order context: " & mOrderContextName, ProcName, ModuleName
        mTaskContext.Finish Empty, False
    End If
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static sEn As Enumerator
If sEn Is Nothing Then
    If mBracketOrdersToCloseout.Count = 0 Then
        GOrderUtils.Log "Closeout completed for order context: " & mOrderContextName, ProcName, ModuleName
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    Set sEn = mBracketOrdersToCloseout.Enumerator
End If

If Not sEn.MoveNext Then
    mTaskContext.Suspend -1
    Exit Sub
End If

Dim lBracketOrder As IBracketOrder
Set lBracketOrder = sEn.Current
lBracketOrder.CloseOut
lBracketOrder.AddStateChangeListener Me

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Const ProcName As String = "ITask_TaskContext"
On Error GoTo Err

Set mTaskContext = Value
Dim lOrderContext As OrderContext
Set lOrderContext = mTaskContext.Data
mOrderContextName = lOrderContext.Name

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In lOrderContext.ActiveBracketOrders
    mBracketOrdersToCloseout.Add lBracketOrder
Next

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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

'@================================================================================
' Helper Functions
'@================================================================================




