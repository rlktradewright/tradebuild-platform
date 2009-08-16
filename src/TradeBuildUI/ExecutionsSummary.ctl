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
    execId = 1
    orderId
    Action
    quantity
    symbol
    price
    Time
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mMonitoredWorkspaces            As Collection
Private mSimulated                      As Boolean

'@================================================================================
' UserControl Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Set mMonitoredWorkspaces = New Collection

ExecutionsList.Left = 0
ExecutionsList.Top = 0

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.execId, , "Exec id"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.orderId, , "ID"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Action, , "Action"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.quantity, , "Quant"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.symbol, , "Symb"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.price, , "Price"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Time, , "Time"

ExecutionsList.SortKey = ExecutionsColumns.Time - 1
ExecutionsList.SortOrder = lvwDescending

End Sub

Private Sub UserControl_Resize()
ExecutionsList.Height = UserControl.Height
ExecutionsList.Width = UserControl.Width

ExecutionsList.ColumnHeaders(ExecutionsColumns.execId).Width = _
    ExecutionsExecIdWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.orderId).Width = _
    ExecutionsOrderIDWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.Action).Width = _
    ExecutionsActionWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.quantity).Width = _
    ExecutionsQuantityWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.symbol).Width = _
    ExecutionsSymbolWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.price).Width = _
    ExecutionsPriceWidth * ExecutionsList.Width / 100

ExecutionsList.ColumnHeaders(ExecutionsColumns.Time).Width = _
    ExecutionsTimeWidth * ExecutionsList.Width / 100

End Sub

Private Sub UserControl_Terminate()
Debug.Print "ExecutionsSummary control terminated"
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change( _
                ev As CollectionChangeEvent)
Dim exec As Execution
Dim listItem As listItem

If ev.changeType <> CollItemAdded Then Exit Sub

Set exec = ev.affectedItem

On Error Resume Next
Set listItem = ExecutionsList.ListItems(exec.execId)
On Error GoTo 0

If listItem Is Nothing Then
    Set listItem = ExecutionsList.ListItems.Add(, exec.execId, exec.execId)
End If

listItem.SubItems(ExecutionsColumns.Action - 1) = IIf(exec.Action = ActionBuy, "BUY", "SELL")
listItem.SubItems(ExecutionsColumns.orderId - 1) = exec.orderBrokerId
listItem.SubItems(ExecutionsColumns.price - 1) = exec.price
listItem.SubItems(ExecutionsColumns.quantity - 1) = exec.quantity
listItem.SubItems(ExecutionsColumns.symbol - 1) = exec.contractSpecifier.localSymbol
listItem.SubItems(ExecutionsColumns.Time - 1) = exec.Time

End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ExecutionsList_ColumnClick(ByVal columnHeader As columnHeader)
If ExecutionsList.SortKey = columnHeader.index - 1 Then
    ExecutionsList.SortOrder = 1 - ExecutionsList.SortOrder
Else
    ExecutionsList.SortKey = columnHeader.index - 1
    ExecutionsList.SortOrder = lvwAscending
End If
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Simulated(ByVal value As Boolean)
If mMonitoredWorkspaces.count > 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "simulated", _
            "Property must be set before any workspaces are monitored"
End If

mSimulated = value
PropertyChanged "simulated"
End Property

Public Property Get Simulated() As Boolean
Simulated = mSimulated
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Clear()
ExecutionsList.ListItems.Clear
End Sub

Public Sub Finish()
Dim i As Long
Dim lWorkspace As Workspace

Dim failpoint As Long
On Error GoTo Err

For i = mMonitoredWorkspaces.count To 1 Step -1
    Set lWorkspace = mMonitoredWorkspaces(i)
    lWorkspace.Executions.removeCollectionChangeListener Me
    mMonitoredWorkspaces.remove i
Next

Clear

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Finish" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Public Sub monitorWorkspace( _
                ByVal pWorkspace As Workspace)
If mSimulated Then
    pWorkspace.ExecutionsSimulated.addCollectionChangeListener Me
Else
    pWorkspace.Executions.addCollectionChangeListener Me
End If
mMonitoredWorkspaces.Add pWorkspace
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================



