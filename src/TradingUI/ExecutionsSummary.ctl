VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
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
      Appearance      =   0
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

Implements ICollectionChangeListener
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                        As String = "ExecutionsSummary"

' Percentage widths of the Open Orders columns
Private Const ExecutionsExecIdWidth = 25
Private Const ExecutionsOrderIDWidth = 10
Private Const ExecutionsActionWidth = 8
Private Const ExecutionsQuantityWidth = 8
Private Const ExecutionsSymbolWidth = 8
Private Const ExecutionsPriceWidth = 10
Private Const ExecutionsTimeWidth = 23

Private Const PropNameBackcolor                 As String = "Backcolor"
Private Const PropNameForecolor                 As String = "Forecolor"

'@================================================================================
' Enums
'@================================================================================

Private Enum ExecutionsColumns
    Time = 1
    Action
    Quantity
    Symbol
    Price
    ExecId
    OrderId
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mExecutionsCollection                               As New EnumerableCollection
Private mPositionManagersCollection                         As New EnumerableCollection

Private mTheme                                              As ITheme

'@================================================================================
' UserControl Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

ExecutionsList.Left = 0
ExecutionsList.Top = 0

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Time, , "Time"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Action, , "Action"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Quantity, , "Quant"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Symbol, , "Symb"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Price, , "Price"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.ExecId, , "Exec Id"
ExecutionsList.ColumnHeaders.Add ExecutionsColumns.OrderId, , "ID"

ExecutionsList.SortKey = ExecutionsColumns.Time - 1
ExecutionsList.SortOrder = lvwDescending

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_InitProperties()
BackColor = vbWindowBackground
ForeColor = vbWindowText
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
BackColor = PropBag.ReadProperty(PropNameBackcolor, vbWindowBackground)
ForeColor = PropBag.ReadProperty(PropNameForecolor, vbWindowText)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameBackcolor, BackColor, vbWindowBackground
PropBag.WriteProperty PropNameForecolor, ForeColor, vbWindowText
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
' ICollectionChangeListener Interface Members
'@================================================================================

Private Sub ICollectionChangeListener_Change( _
                ev As CollectionChangeEventData)
Const ProcName As String = "ICollectionChangeListener_Change"
On Error GoTo Err

If ev.changeType <> CollItemAdded Then Exit Sub

If TypeOf ev.AffectedItem Is IExecutionReport Then
    addExecution ev.AffectedItem
ElseIf TypeOf ev.AffectedItem Is PositionManager Then
    Dim lPm As PositionManager
    Set lPm = ev.AffectedItem
    MonitorExecutions lPm.Executions
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

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

Public Property Let BackColor(ByVal Value As OLE_COLOR)
ExecutionsList.BackColor = Value
PropertyChanged PropNameBackcolor
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
BackColor = ExecutionsList.BackColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
ExecutionsList.ForeColor = Value
PropertyChanged PropNameForecolor
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
ForeColor = ExecutionsList.ForeColor
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

BackColor = mTheme.TextBackColor
ForeColor = mTheme.TextForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

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

Dim lExecs As Executions
For Each lExecs In mExecutionsCollection
    lExecs.RemoveCollectionChangeListener Me
Next
mExecutionsCollection.Clear

Dim lPms As PositionManagers
For Each lPms In mPositionManagersCollection
    lPms.RemoveCollectionChangeListener Me
Next
mPositionManagersCollection.Clear

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
mExecutionsCollection.Add pExecutions

Dim lExec As IExecutionReport
For Each lExec In pExecutions
    addExecution lExec
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Public Sub MonitorPositions( _
                ByVal pPositionManagers As PositionManagers)
Const ProcName As String = "MonitorPosition"
On Error GoTo Err

pPositionManagers.AddCollectionChangeListener Me
mPositionManagersCollection.Add pPositionManagers

Dim lPm As PositionManager
For Each lPm In pPositionManagers
    MonitorExecutions lPm.Executions
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addExecution(ByVal pExec As IExecutionReport)
Const ProcName As String = "addExecution"
On Error GoTo Err

Dim lListItem As ListItem
On Error Resume Next
Set lListItem = ExecutionsList.ListItems(pExec.Id)
On Error GoTo Err

If lListItem Is Nothing Then
    Set lListItem = ExecutionsList.ListItems.Add(, pExec.Id, pExec.FillTime)
End If

lListItem.SubItems(ExecutionsColumns.Action - 1) = IIf(pExec.Action = OrderActionBuy, "BUY", "SELL")
lListItem.SubItems(ExecutionsColumns.OrderId - 1) = pExec.BrokerId
lListItem.SubItems(ExecutionsColumns.Price - 1) = pExec.Price
lListItem.SubItems(ExecutionsColumns.Quantity - 1) = pExec.Quantity
lListItem.SubItems(ExecutionsColumns.Symbol - 1) = pExec.SecurityName
lListItem.SubItems(ExecutionsColumns.ExecId - 1) = pExec.Id

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

