VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ExecutionsSummary 
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   ScaleHeight     =   3810
   ScaleWidth      =   5535
   Begin MSComctlLib.ListView ExecutionsList 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Filled orders"
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "ExecutionsSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements CollectionChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "ExecutionsSummary"

' Percentage widths of the Open Orders columns
Private Const ExecutionsExecIdWidth = 25
Private Const ExecutionsOrderIDWidth = 10
Private Const ExecutionsActionWidth = 8
Private Const ExecutionsQuantityWidth = 8
Private Const ExecutionsSymbolWidth = 8
Private Const ExecutionsPriceWidth = 10
Private Const ExecutionsTimeWidth = 23

'@================================================================================
' Enums
'@================================================================================

Private Enum ExecutionsColumns
    ExecId = 1
    OrderId
    Action
    Quantity
    Symbol
    Price
    Time
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mMonitoredExecutions            As Collection

'@================================================================================
' UserControl Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"

On Error GoTo Err

Set mMonitoredExecutions = New Collection

ExecutionsList.Left = 0
ExecutionsList.Top = 0

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.ExecId, , "Exec Id"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.OrderId, , "ID"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Action, , "Action"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Quantity, , "Quant"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Symbol, , "Symb"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Price, , "Price"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Time, , "Time"

ExecutionsList.SortKey = ExecutionsColumns.Time - 1
ExecutionsList.SortOrder = lvwDescending

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName

End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"

On Error GoTo Err

ExecutionsList.Height = UserControl.Height
ExecutionsList.Width = UserControl.Width

ExecutionsList.ColumnHeaders(ExecutionsColumns.ExecId).Width = _
    ExecutionsExecIdWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.OrderId).Width = _
    ExecutionsOrderIDWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.Action).Width = _
    ExecutionsActionWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.Quantity).Width = _
    ExecutionsQuantityWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.Symbol).Width = _
    ExecutionsSymbolWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.Price).Width = _
    ExecutionsPriceWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.Time).Width = _
    ExecutionsTimeWidth * ExecutionsList.Width / 100

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName

End Sub

Private Sub UserControl_Terminate()
Debug.Print "ExecutionsSummary control terminated"
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change( _
                ev As CollectionChangeEventData)
Const ProcName As String = "CollectionChangeListener_Change"
On Error GoTo Err

If ev.changeType <> CollItemAdded Then Exit Sub

addExecution ev.AffectedItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ExecutionsList_ColumnClick(ByVal columnHeader As columnHeader)
Const ProcName As String = "ExecutionsList_ColumnClick"

On Error GoTo Err

If ExecutionsList.SortKey = columnHeader.index - 1 Then
    ExecutionsList.SortOrder = 1 - ExecutionsList.SortOrder
Else
    ExecutionsList.SortKey = columnHeader.index - 1
    ExecutionsList.SortOrder = lvwAscending
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

Public Sub Clear()
Const ProcName As String = "Clear"

On Error GoTo Err

ExecutionsList.ListItems.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim i As Long
For i = mMonitoredExecutions.Count To 1 Step -1
    Dim lExecs As Executions
    Set lExecs = mMonitoredExecutions(i)
    lExecs.RemoveCollectionChangeListener Me
    mMonitoredExecutions.Remove i
Next

Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub MonitorExecutions( _
                ByVal pExecutions As Executions)
Const ProcName As String = "MonitorExecutions"
On Error GoTo Err

pExecutions.AddCollectionChangeListener Me
mMonitoredExecutions.Add pExecutions

Dim lExec As Execution
For Each lExec In pExecutions
    addExecution lExec
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addExecution(ByVal pExec As Execution)
Const ProcName As String = "addExecution"
On Error GoTo Err

Dim lListItem As ListItem
On Error Resume Next
Set lListItem = ExecutionsList.ListItems(pExec.Id)
On Error GoTo Err

If lListItem Is Nothing Then
    Set lListItem = ExecutionsList.ListItems.Add(, pExec.Id, pExec.Id)
End If

lListItem.SubItems(ExecutionsColumns.Action - 1) = IIf(pExec.Action = OrderActionBuy, "BUY", "SELL")
lListItem.SubItems(ExecutionsColumns.OrderId - 1) = pExec.BrokerId
lListItem.SubItems(ExecutionsColumns.Price - 1) = pExec.Price
lListItem.SubItems(ExecutionsColumns.Quantity - 1) = pExec.Quantity
lListItem.SubItems(ExecutionsColumns.Symbol - 1) = pExec.SecurityName
lListItem.SubItems(ExecutionsColumns.Time - 1) = pExec.FillTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

