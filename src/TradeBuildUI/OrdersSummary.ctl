VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl OrdersSummary 
   Alignable       =   -1  'True
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12810
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   12810
   Begin VB.TextBox EditText 
      Height          =   285
      Left            =   11160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ImageList OrderPlexImageList 
      Left            =   11880
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrdersSummary.ctx":0000
            Key             =   "Expand"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrdersSummary.ctx":0452
            Key             =   "Contract"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid OrderPlexGrid 
      Height          =   3900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6879
      _Version        =   393216
      Rows            =   0
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColorFixed  =   12632256
      MergeCells      =   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "OrdersSummary"
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
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements ChangeListener
Implements CollectionChangeListener
Implements ProfitListener
Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Click()
Event SelectionChanged()
                
'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "OrdersSummary"

Private Const RowDataOrderPlexBase As Long = &H100
Private Const RowDataPositionManagerBase As Long = &H1000000

'@================================================================================
' Enums
'@================================================================================

Private Enum OPGridColumns
    symbol
    ExpandIndicator
    OtherColumns    ' keep this entry last
End Enum

Private Enum OPGridOrderPlexColumns
    CreationTime = OPGridColumns.OtherColumns
    Size
    profit
    MaxProfit
    Drawdown
    currencyCode
End Enum

Private Enum OPGridPositionColumns
    exchange = OPGridColumns.OtherColumns
    Size
    profit
    MaxProfit
    Drawdown
    currencyCode
End Enum

Private Enum OPGridOrderColumns
    typeInPlex = OPGridColumns.OtherColumns
    Action
    Quantity
    OrderType
    Price
    AuxPrice
    Status
    Size
    QuantityRemaining
    AveragePrice
    LastFillTime
    LastFillPrice
    Id
    BrokerId
End Enum

Private Enum OPGridColumnWidths
    ExpandIndicatorWidth = 3
    SymbolWidth = 15
End Enum

Private Enum OPGridOrderPlexColumnWidths
    CreationTimeWidth = 15
    SizeWidth = 6
    ProfitWidth = 9
    MaxProfitWidth = 9
    DrawdownWidth = 9
    CurrencyCodeWidth = 4
End Enum

Private Enum OPGridPositionColumnWidths
    ExchangeWidth = 9
    SizeWidth = 6
    ProfitWidth = 9
    MaxProfitWidth = 9
    DrawdownWidth = 9
    CurrencyCodeWidth = 5
End Enum

Private Enum OPGridOrderColumnWidths
    TypeInPlexWidth = 9
    SizeWidth = 6
    AveragePriceWidth = 9
    StatusWidth = 13
    ActionWidth = 4
    QuantityWidth = 6
    QuantityRemainingWidth = 5
    OrderTypeWidth = 5
    PriceWidth = 9
    AuxPriceWidth = 9
    LastFillTimeWidth = 15
    LastFillPriceWidth = 9
    IdWidth = 40
    BrokerIdWidth = 11
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type OrderPlexGridMappingEntry
    op                  As OrderPlex
    
    ' indicates whether this entry in the grid is expanded
    isExpanded          As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    gridIndex           As Long
                                
    ' offset from gridIndex of line in OrdersGrid relating to
    ' the corresponding order: -1 means  it's not in the grid
    entryGridOffset      As Long
    stopGridOffset       As Long
    targetGridOffset     As Long
    closeoutGridOffset   As Long
    
End Type

Private Type PositionManagerGridMappingEntry
    
    ' indicates whether this entry in the grid is expanded
    isExpanded          As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    gridIndex           As Long
                                
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mSelectedOrderPlexGridRow                       As Long
Private mSelectedOrderPlex                              As OrderPlex
Private mSelectedOrderIndex                             As Long

Private mOrderPlexGridMappingTable()                    As OrderPlexGridMappingEntry
Private mMaxOrderPlexGridMappingTableIndex              As Long

Private mPositionManagerGridMappingTable()              As PositionManagerGridMappingEntry
Private mMaxPositionManagerGridMappingTableIndex        As Long

' the index of the first entry in the order plex frid that relates to
' order plexes (rather than header rows, currency totals etc)
Private mFirstOrderPlexGridRowIndex                     As Long

Private mLetterWidth                                    As Single
Private mDigitWidth                                     As Single

Private mMonitoredWorkspaces                            As Collection

Private mSimulated                                      As Boolean

Private mInitialised                                    As Boolean

Private mIsEditing                                        As Boolean
Private mEditedOrderPlex                                As OrderPlex
Private mEditedOrderIndex                               As Long
Private mEditedCol                                      As Long

'@================================================================================
' User Control Event Handlers
'@================================================================================

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
Const ProcName As String = "UserControl_AccessKeyPress"
On Error GoTo Err

If Not mIsEditing Then Exit Sub

handleEditingTerminationKey KeyAscii

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub UserControl_Initialize()
Dim widthString As String

Const ProcName As String = "UserControl_Initialize"
Dim failpoint As String
On Error GoTo Err

Set mMonitoredWorkspaces = New Collection

widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = UserControl.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = UserControl.TextWidth(widthString) / Len(widthString)

setupOrderPlexGrid

ReDim mOrderPlexGridMappingTable(3) As OrderPlexGridMappingEntry
mMaxOrderPlexGridMappingTableIndex = -1

ReDim mPositionManagerGridMappingTable(3) As PositionManagerGridMappingEntry
mMaxPositionManagerGridMappingTableIndex = -1

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_InitProperties()
mSimulated = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Simulated = PropBag.ReadProperty("simulated", False)
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

OrderPlexGrid.Width = UserControl.Width
OrderPlexGrid.Height = UserControl.Height

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub UserControl_Terminate()
Debug.Print "OrdersSummary control terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "simulated", mSimulated, False
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEvent)
Const ProcName As String = "ChangeListener_Change"
Dim failpoint As String
On Error GoTo Err

If TypeOf ev.Source Is OrderPlex Then
    Dim opChangeType As OrderPlexChangeTypes
    Dim op As OrderPlex
    Dim opIndex As Long
    
    Set op = ev.Source
    
    opIndex = findOrderPlexTableIndex(op)
    
    With mOrderPlexGridMappingTable(opIndex)
    
        opChangeType = ev.changeType
        
        Select Case opChangeType
        Case OrderPlexChangeTypes.OrderPlexCreated
        
        Case OrderPlexChangeTypes.OrderPlexCompleted
            If op Is mEditedOrderPlex Then endEdit
            If op.Size = 0 Then op.RemoveChangeListener Me
        Case OrderPlexChangeTypes.OrderPlexSelfCancelled
            If op Is mEditedOrderPlex Then endEdit
            If op.Size = 0 Then op.RemoveChangeListener Me
        Case OrderPlexChangeTypes.OrderPlexEntryOrderChanged
            If op Is mEditedOrderPlex Then endEdit
            displayOrderValues .gridIndex + .entryGridOffset, op.entryOrder
        Case OrderPlexChangeTypes.OrderPlexStopOrderChanged
            If op Is mEditedOrderPlex Then endEdit
            displayOrderValues .gridIndex + .stopGridOffset, op.stopOrder
        Case OrderPlexChangeTypes.OrderPlexTargetOrderChanged
            If op Is mEditedOrderPlex Then endEdit
            displayOrderValues .gridIndex + .targetGridOffset, op.targetOrder
        Case OrderPlexChangeTypes.OrderPlexCloseoutOrderCreated
            If op Is mEditedOrderPlex Then endEdit
            If .targetGridOffset >= 0 Then
                .closeoutGridOffset = .targetGridOffset + 1
            ElseIf .stopGridOffset >= 0 Then
                .closeoutGridOffset = .stopGridOffset + 1
            ElseIf .entryGridOffset >= 0 Then
                .closeoutGridOffset = .entryGridOffset + 1
            Else
                .closeoutGridOffset = 1
            End If
            
            addOrderEntryToOrderPlexGrid .gridIndex + .closeoutGridOffset, _
                                    .op.Contract.Specifier.localSymbol, _
                                    op.CloseoutOrder, _
                                    opIndex, _
                                    "Closeout"
        Case OrderPlexChangeTypes.OrderPlexCloseoutOrderChanged
            If op Is mEditedOrderPlex Then endEdit
            displayOrderValues .gridIndex + .closeoutGridOffset, _
                                                op.CloseoutOrder
        Case OrderPlexChangeTypes.OrderPlexProfitThresholdExceeded
    
        Case OrderPlexChangeTypes.OrderPlexLossThresholdExceeded
    
        Case OrderPlexChangeTypes.OrderPlexDrawdownThresholdExceeded
    
        Case OrderPlexChangeTypes.OrderPlexSizeChanged
            If op Is mEditedOrderPlex Then endEdit
            OrderPlexGrid.TextMatrix(.gridIndex, OPGridOrderPlexColumns.Size) = op.Size
        Case OrderPlexChangeTypes.OrderPlexStateChanged
            If op Is mEditedOrderPlex Then endEdit
            If op.State = OrderPlexStateCodes.OrderPlexStateSubmitted Then
                OrderPlexGrid.TextMatrix(.gridIndex, OPGridOrderPlexColumns.CreationTime) = formattedTime(op.CreationTime)
            End If
            If op.State <> OrderPlexStateCodes.OrderPlexStateCreated And _
                op.State <> OrderPlexStateCodes.OrderPlexStateSubmitted _
            Then
                ' the order plex is now in a state where it can't be modified.
                ' If it's the currently selected order plex, make it not so.
                If op Is mSelectedOrderPlex Then
                    invertEntryColors mSelectedOrderPlexGridRow
                    mSelectedOrderPlexGridRow = -1
                    Set mSelectedOrderPlex = Nothing
                    RaiseEvent SelectionChanged
                End If
            End If
        End Select
    End With
ElseIf TypeOf ev.Source Is PositionManager Then
    Dim pmChangeType As PositionManagerChangeTypes
    Dim pm As PositionManager
    Dim pmIndex As Long
    
    Set pm = ev.Source
    pmChangeType = ev.changeType
    
    
    Select Case pmChangeType
    Case PositionManagerChangeTypes.PositionSizeChanged
        pmIndex = findPositionManagerTableIndex(pm)
        OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(pmIndex).gridIndex, _
                                OPGridPositionColumns.Size) = pm.PositionSize
    End Select
End If

adjustEditBox

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEvent)

Const ProcName As String = "CollectionChangeListener_Change"
Dim failpoint As String
On Error GoTo Err

If TypeOf ev.Source Is OrderPlexes Then
    Dim op As OrderPlex
    Set op = ev.affectedItem
    
    Select Case ev.changeType
    Case CollItemAdded
        op.AddChangeListener Me
        op.AddProfitListener Me
    Case CollItemRemoved
        op.RemoveChangeListener Me
        op.RemoveProfitListener Me
    End Select
ElseIf TypeOf ev.Source Is Tickers Then
    Dim lTicker As Ticker
    Set lTicker = ev.affectedItem
    
    If lTicker.State = TickerStateReady Or lTicker.State = TickerStateRunning Then
        Select Case ev.changeType
        Case CollItemAdded
            listenForProfit lTicker
        Case CollItemRemoved
            ' nothing to do here as the Ticker has already
            ' tidied everything up
        End Select
    Else
        lTicker.addStateChangeListener Me
    End If
End If

adjustEditBox

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' ProfitListener Interface Members
'@================================================================================

Private Sub ProfitListener_profitAmount(ev As ProfitEvent)
Dim rowIndex As Long

Const ProcName As String = "ProfitListener_profitAmount"
Dim failpoint As String
On Error GoTo Err

If TypeOf ev.Source Is OrderPlex Then
    Dim opProfitType As ProfitTypes
    Dim op As OrderPlex
    Dim opIndex As Long
    
    Set op = ev.Source
    
    opIndex = findOrderPlexTableIndex(op)
    rowIndex = mOrderPlexGridMappingTable(opIndex).gridIndex
    opProfitType = ev.profitType
    
    Select Case opProfitType
    Case ProfitTypes.ProfitTypeProfit
        displayProfitValue ev.ProfitAmount, rowIndex, OPGridOrderPlexColumns.profit
    Case ProfitTypes.ProfitTypeMaxProfit
        displayProfitValue ev.ProfitAmount, rowIndex, OPGridOrderPlexColumns.MaxProfit
    Case ProfitTypes.ProfitTypeDrawdown
        displayProfitValue -ev.ProfitAmount, rowIndex, OPGridOrderPlexColumns.Drawdown
    End Select

ElseIf TypeOf ev.Source Is PositionManager Then
    Dim pmProfitType As ProfitTypes
    Dim pm As PositionManager
    Dim pmIndex As Long
    
    Set pm = ev.Source
    
    pmIndex = findPositionManagerTableIndex(pm)
    rowIndex = mPositionManagerGridMappingTable(pmIndex).gridIndex
    pmProfitType = ev.profitType
    
    Select Case pmProfitType
    Case ProfitTypes.ProfitTypeSessionProfit
        displayProfitValue ev.ProfitAmount, rowIndex, OPGridPositionColumns.profit
    Case ProfitTypes.ProfitTypeSessionMaxProfit
        displayProfitValue ev.ProfitAmount, rowIndex, OPGridPositionColumns.MaxProfit
    Case ProfitTypes.ProfitTypeSessionDrawdown
        displayProfitValue -ev.ProfitAmount, rowIndex, OPGridPositionColumns.Drawdown
    Case ProfitTypes.ProfitTypeTradeProfit
    Case ProfitTypes.ProfitTypeTradeMaxProfit
    Case ProfitTypes.ProfitTypeTradeDrawdown
    End Select
End If

adjustEditBox

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As TWUtilities30.StateChangeEvent)
If TypeOf ev.Source Is Ticker Then
    If ev.State = TickerStates.TickerStateReady Then
        Dim lTicker As Ticker
        Set lTicker = ev.Source
        listenForProfit lTicker
        lTicker.removeStateChangeListener Me
    End If
End If
End Sub

'@================================================================================
' Form Control Event Handlers
'@================================================================================

Private Sub EditText_KeyDown(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "EditText_KeyDown"
Dim failpoint As String
On Error GoTo Err

handleEditingTerminationKey KeyCode

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub EditText_KeyPress(KeyAscii As Integer)
' Delete returns to get rid of beep.
If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub

Private Sub OrderPlexGrid_Click()
Dim row As Long
Dim rowdata As Long
Dim op As OrderPlex
Dim index As Long
Dim selectedOrder As Order

Const ProcName As String = "OrderPlexGrid_Click"
Dim failpoint As String
On Error GoTo Err

row = OrderPlexGrid.row

If OrderPlexGrid.MouseCol = OPGridColumns.symbol Then
    RaiseEvent Click
    Exit Sub
End If

If OrderPlexGrid.MouseCol = OPGridColumns.ExpandIndicator Then
    expandOrContract
    adjustEditBox
Else

    invertEntryColors mSelectedOrderPlexGridRow
    
    mSelectedOrderPlexGridRow = -1
    
    OrderPlexGrid.row = row
    rowdata = OrderPlexGrid.rowdata(row)
    If rowdata < RowDataPositionManagerBase And _
        rowdata >= RowDataOrderPlexBase _
    Then
        index = rowdata - RowDataOrderPlexBase
        Set op = mOrderPlexGridMappingTable(index).op
        If op.State = OrderPlexStateCodes.OrderPlexStateCreated Or _
            op.State = OrderPlexStateCodes.OrderPlexStateSubmitted _
        Then
            
            mSelectedOrderPlexGridRow = row
            Set mSelectedOrderPlex = op
            invertEntryColors mSelectedOrderPlexGridRow
            
            mSelectedOrderIndex = mSelectedOrderPlexGridRow - mOrderPlexGridMappingTable(index).gridIndex
            If mSelectedOrderIndex <> 0 Then
                Set selectedOrder = op.Order(mSelectedOrderIndex)
                If selectedOrder.IsModifiable Then
                    If (OrderPlexGrid.MouseCol = OPGridOrderColumns.Price And _
                            selectedOrder.IsAttributeModifiable(OrderAttributeIds.OrderAttLimitPrice)) Or _
                        (OrderPlexGrid.MouseCol = OPGridOrderColumns.AuxPrice And _
                            selectedOrder.IsAttributeModifiable(OrderAttributeIds.OrderAttTriggerPrice)) Or _
                        (OrderPlexGrid.MouseCol = OPGridOrderColumns.Quantity And _
                        selectedOrder.IsAttributeModifiable(OrderAttributeIds.OrderAttQuantity)) _
                    Then
                        mIsEditing = True
                        Set mEditedOrderPlex = op
                        mEditedOrderIndex = mSelectedOrderIndex
                        mEditedCol = OrderPlexGrid.MouseCol
                        OrderPlexGrid.col = mEditedCol
                        
                        EditText.Text = OrderPlexGrid.Text
                        EditText.SelStart = 0
                        EditText.SelLength = Len(EditText.Text)
                        EditText.Visible = True
                        EditText.SetFocus
                                                
                        adjustEditBox
                    
                    End If
                End If
            End If
        End If
    End If
End If
RaiseEvent Click
RaiseEvent SelectionChanged

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderPlexGrid_Scroll()
Const ProcName As String = "OrderPlexGrid_Scroll"
Dim failpoint As String
On Error GoTo Err

adjustEditBox

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
Dim failpoint As String
On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get IsEditing() As Boolean
IsEditing = mIsEditing
End Property

Public Property Get IsSelectedItemModifiable() As Boolean
Dim selectedOrder As Order

Const ProcName As String = "IsSelectedItemModifiable"
Dim failpoint As String
On Error GoTo Err

If mSelectedOrderIndex = 0 Then Exit Property

Set selectedOrder = mSelectedOrderPlex.Order(mSelectedOrderIndex)
If Not selectedOrder Is Nothing Then
    IsSelectedItemModifiable = selectedOrder.IsModifiable
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get SelectedItem() As OrderPlex
Set SelectedItem = mSelectedOrderPlex
End Property

Public Property Get SelectedOrderIndex() As Long
SelectedOrderIndex = mSelectedOrderIndex
End Property

Public Property Let Simulated(ByVal value As Boolean)
Const ProcName As String = "Simulated"
Dim failpoint As String
On Error GoTo Err

If mMonitoredWorkspaces.Count > 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Property must be set before any workspaces are monitored"
End If

mSimulated = value
PropertyChanged "simulated"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get Simulated() As Boolean
Simulated = mSimulated
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Dim i As Long
Dim lWorkspace As Workspace
Dim lTicker As Ticker

Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

For i = 0 To mMaxOrderPlexGridMappingTableIndex
    If Not mOrderPlexGridMappingTable(i).op Is Nothing Then
        mOrderPlexGridMappingTable(i).op.RemoveChangeListener Me
        mOrderPlexGridMappingTable(i).op.RemoveProfitListener Me
        Set mOrderPlexGridMappingTable(i).op = Nothing
    End If
Next

For i = mMonitoredWorkspaces.Count To 1 Step -1
    Set lWorkspace = mMonitoredWorkspaces(i)
    lWorkspace.Tickers.RemoveCollectionChangeListener Me
    For Each lTicker In lWorkspace.Tickers
        If lTicker.State <> TickerStateClosing And _
            lTicker.State <> TickerStateStopped _
        Then
            If mSimulated Then
                If Not lTicker.PositionManagerSimulated Is Nothing Then
                    lTicker.PositionManagerSimulated.RemoveChangeListener Me
                    lTicker.PositionManagerSimulated.RemoveProfitListener Me
                End If
            Else
                If Not lTicker.PositionManager Is Nothing Then
                    lTicker.PositionManager.RemoveChangeListener Me
                    lTicker.PositionManager.RemoveProfitListener Me
                End If
            End If
        End If
    Next
    lWorkspace.OrderPlexes.RemoveCollectionChangeListener Me
    mMonitoredWorkspaces.Remove i
Next

OrderPlexGrid.Clear
mInitialised = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub monitorWorkspace( _
                ByVal pWorkspace As Workspace)
Const ProcName As String = "monitorWorkspace"
Dim failpoint As String
On Error GoTo Err

If Not mInitialised Then setupOrderPlexGrid

pWorkspace.Tickers.AddCollectionChangeListener Me
'If mSimulated Then
'    pWorkspace.OrderPlexesSimulated.AddCollectionChangeListener Me
'Else
'    pWorkspace.OrderPlexes.AddCollectionChangeListener Me
'End If
mMonitoredWorkspaces.Add pWorkspace

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================

Private Function addEntryToOrderPlexGrid( _
                ByVal symbol As String, _
                Optional ByVal before As Boolean, _
                Optional ByVal index As Long = -1) As Long
Dim i As Long

Const ProcName As String = "addEntryToOrderPlexGrid"
Dim failpoint As String
On Error GoTo Err

If index < 0 Then
    For i = mFirstOrderPlexGridRowIndex To OrderPlexGrid.Rows - 1
        If (before And _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) >= symbol) Or _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = "" _
        Then
            index = i
            Exit For
        ElseIf (Not before And _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) > symbol) Or _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = "" _
        Then
            index = i
            Exit For
        End If
    Next
    
    If index < 0 Then
        OrderPlexGrid.addItem ""
        index = OrderPlexGrid.Rows - 1
    ElseIf OrderPlexGrid.TextMatrix(index, OPGridColumns.symbol) = "" Then
        OrderPlexGrid.TextMatrix(index, OPGridColumns.symbol) = symbol
    Else
        OrderPlexGrid.addItem "", index
    End If
Else
    OrderPlexGrid.addItem "", index
End If

OrderPlexGrid.TextMatrix(index, OPGridColumns.symbol) = symbol
If index < OrderPlexGrid.Rows - 1 Then
    ' this new entry has displaced one or more existing entries so
    ' the OrderPlexGridMappingTable and PositionManageGridMappingTable indexes
    ' need to be adjusted
    For i = 0 To mMaxOrderPlexGridMappingTableIndex
        If mOrderPlexGridMappingTable(i).gridIndex >= index Then
            mOrderPlexGridMappingTable(i).gridIndex = mOrderPlexGridMappingTable(i).gridIndex + 1
        End If
    Next
    For i = 0 To mMaxPositionManagerGridMappingTableIndex
        If mPositionManagerGridMappingTable(i).gridIndex >= index Then
            mPositionManagerGridMappingTable(i).gridIndex = mPositionManagerGridMappingTable(i).gridIndex + 1
        End If
    Next
End If

addEntryToOrderPlexGrid = index

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function addOrderPlexEntryToOrderPlexGrid( _
                ByVal symbol As String, _
                ByVal orderPlexTableIndex As Long) As Long
Dim index As Long

Const ProcName As String = "addOrderPlexEntryToOrderPlexGrid"
Dim failpoint As String
On Error GoTo Err

index = addEntryToOrderPlexGrid(symbol, False)

OrderPlexGrid.rowdata(index) = orderPlexTableIndex + RowDataOrderPlexBase

OrderPlexGrid.row = index
OrderPlexGrid.col = OPGridColumns.ExpandIndicator
OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture

OrderPlexGrid.col = OPGridOrderPlexColumns.profit
OrderPlexGrid.CellBackColor = &HC0C0C0
OrderPlexGrid.CellForeColor = vbWhite

OrderPlexGrid.col = OPGridOrderPlexColumns.MaxProfit
OrderPlexGrid.CellBackColor = &HC0C0C0
OrderPlexGrid.CellForeColor = vbWhite

OrderPlexGrid.col = OPGridOrderPlexColumns.Drawdown
OrderPlexGrid.CellBackColor = &HC0C0C0
OrderPlexGrid.CellForeColor = vbWhite

addOrderPlexEntryToOrderPlexGrid = index

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function
                
Private Sub addOrderEntryToOrderPlexGrid( _
                ByVal index As Long, _
                ByVal symbol As String, _
                ByVal pOrder As Order, _
                ByVal orderPlexTableIndex As Long, _
                ByVal typeInPlex As String)


Const ProcName As String = "addOrderEntryToOrderPlexGrid"
Dim failpoint As String
On Error GoTo Err

index = addEntryToOrderPlexGrid(symbol, False, index)

OrderPlexGrid.rowdata(index) = orderPlexTableIndex + RowDataOrderPlexBase

OrderPlexGrid.TextMatrix(index, OPGridOrderColumns.typeInPlex) = typeInPlex

displayOrderValues index, pOrder

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Sub adjustEditBox()
Dim opIndex As Long
Const ProcName As String = "adjustEditBox"
Dim failpoint As String
On Error GoTo Err

If mIsEditing Then
    opIndex = findOrderPlexTableIndex(mEditedOrderPlex)
    OrderPlexGrid.row = mOrderPlexGridMappingTable(opIndex).gridIndex + mEditedOrderIndex
    OrderPlexGrid.col = mEditedCol
    
    EditText.Move OrderPlexGrid.Left + OrderPlexGrid.CellLeft + 8, _
                OrderPlexGrid.Top + OrderPlexGrid.CellTop + 8, _
                OrderPlexGrid.CellWidth - 16, _
                OrderPlexGrid.CellHeight - 16
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Function contractOrderPlexEntry( _
                ByVal index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Dim lIndex As Long

Const ProcName As String = "contractOrderPlexEntry"
Dim failpoint As String
On Error GoTo Err

With mOrderPlexGridMappingTable(index)
    
    If mIsEditing And .op Is mEditedOrderPlex Then endEdit
    
    If .entryGridOffset >= 0 Then
        lIndex = .gridIndex + .entryGridOffset
        OrderPlexGrid.rowHeight(lIndex) = 0
    End If
    If .stopGridOffset >= 0 Then
        lIndex = .gridIndex + .stopGridOffset
        OrderPlexGrid.rowHeight(lIndex) = 0
    End If
    If .targetGridOffset >= 0 Then
        lIndex = .gridIndex + .targetGridOffset
        OrderPlexGrid.rowHeight(lIndex) = 0
    End If
    If .closeoutGridOffset >= 0 Then
        lIndex = .gridIndex + .closeoutGridOffset
        OrderPlexGrid.rowHeight(lIndex) = 0
    End If
    
    If Not preserveCurrentExpandedState Then
        .isExpanded = False
        OrderPlexGrid.row = .gridIndex
        OrderPlexGrid.col = OPGridColumns.ExpandIndicator
        OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
        Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Expand").Picture
    End If
End With

contractOrderPlexEntry = lIndex

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Sub contractPositionManagerEntry(ByVal index As Long)
Dim i As Long
Dim symbol As String
Dim lOpEntryIndex As Long

Const ProcName As String = "contractPositionManagerEntry"
Dim failpoint As String
On Error GoTo Err

mPositionManagerGridMappingTable(index).isExpanded = False
OrderPlexGrid.row = mPositionManagerGridMappingTable(index).gridIndex
OrderPlexGrid.col = OPGridColumns.ExpandIndicator
OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Expand").Picture

symbol = OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(index).gridIndex, OPGridColumns.symbol)
i = mPositionManagerGridMappingTable(index).gridIndex + 1
Do While OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = symbol
    OrderPlexGrid.rowHeight(i) = 0
    lOpEntryIndex = OrderPlexGrid.rowdata(i) - RowDataOrderPlexBase
    i = contractOrderPlexEntry(lOpEntryIndex, True) + 1
Loop

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub displayOrderValues( _
                ByVal gridIndex As Long, _
                ByVal pOrder As Order)
Dim lTicker As Ticker

Const ProcName As String = "displayOrderValues"
Dim failpoint As String
On Error GoTo Err

Set lTicker = pOrder.Ticker

OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Action) = OrderActionToString(pOrder.Action)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.AuxPrice) = lTicker.FormatPrice(pOrder.TriggerPrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.AveragePrice) = lTicker.FormatPrice(pOrder.AveragePrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Id) = pOrder.Id
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.LastFillPrice) = lTicker.FormatPrice(pOrder.LastFillPrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.LastFillTime) = formattedTime(pOrder.FillTime)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.OrderType) = OrderTypeToShortString(pOrder.OrderType)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Price) = lTicker.FormatPrice(pOrder.LimitPrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Quantity) = pOrder.Quantity
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.QuantityRemaining) = pOrder.QuantityRemaining
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Size) = IIf(pOrder.QuantityFilled <> 0, pOrder.QuantityFilled, 0)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Status) = OrderStatusToString(pOrder.Status)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.BrokerId) = pOrder.BrokerId

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub displayProfitValue( _
                ByVal profit As Currency, _
                ByVal rowIndex As Long, _
                ByVal colIndex As Long)
Const ProcName As String = "displayProfitValue"
Dim failpoint As String
On Error GoTo Err

OrderPlexGrid.row = rowIndex
OrderPlexGrid.col = colIndex
OrderPlexGrid.Text = Format(profit, "0.00")
If profit >= 0 Then
    OrderPlexGrid.CellForeColor = CPositiveProfitColor
Else
    OrderPlexGrid.CellForeColor = CNegativeProfitColor
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub endEdit()
Const ProcName As String = "endEdit"
Dim failpoint As String
On Error GoTo Err

EditText.Text = ""
EditText.Visible = False
mIsEditing = False
Set mEditedOrderPlex = Nothing
mEditedOrderIndex = -1
mEditedCol = -1
OrderPlexGrid.SetFocus

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub
                
Private Sub expandOrContract()
Dim rowdata As Long
Dim index As Long
Dim expanded As Boolean

Const ProcName As String = "expandOrContract"
Dim failpoint As String
On Error GoTo Err

rowdata = OrderPlexGrid.rowdata(OrderPlexGrid.MouseRow)
If rowdata >= RowDataPositionManagerBase Then
    index = rowdata - RowDataPositionManagerBase
    expanded = mPositionManagerGridMappingTable(index).isExpanded
    If expanded Then
        contractPositionManagerEntry index
    Else
        expandPositionManagerEntry index
    End If
ElseIf rowdata >= RowDataOrderPlexBase Then
    index = rowdata - RowDataOrderPlexBase
    expanded = mOrderPlexGridMappingTable(index).isExpanded
    If OrderPlexGrid.row <> mOrderPlexGridMappingTable(index).gridIndex Then
        ' clicked on an order entry
        Exit Sub
    End If
    If expanded Then
        contractOrderPlexEntry index
    Else
        expandOrderPlexEntry index
    End If
Else
    Exit Sub
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Function expandOrderPlexEntry( _
                ByVal index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Dim lIndex As Long


Const ProcName As String = "expandOrderPlexEntry"
Dim failpoint As String
On Error GoTo Err

With mOrderPlexGridMappingTable(index)
    
    If .entryGridOffset >= 0 Then
        lIndex = .gridIndex + .entryGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.rowHeight(lIndex) = -1
    End If
    If .stopGridOffset >= 0 Then
        lIndex = .gridIndex + .stopGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.rowHeight(lIndex) = -1
    End If
    If .targetGridOffset >= 0 Then
        lIndex = .gridIndex + .targetGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.rowHeight(lIndex) = -1
    End If
    If .closeoutGridOffset >= 0 Then
        lIndex = .gridIndex + .closeoutGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.rowHeight(lIndex) = -1
    End If
    
    If Not preserveCurrentExpandedState Then
        .isExpanded = True
        OrderPlexGrid.row = .gridIndex
        OrderPlexGrid.col = OPGridColumns.ExpandIndicator
        OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
        Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture
    End If
End With

expandOrderPlexEntry = lIndex

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Sub expandPositionManagerEntry(ByVal index As Long)
Dim i As Long
Dim symbol As String
Dim lOpEntryIndex As Long

Const ProcName As String = "expandPositionManagerEntry"
Dim failpoint As String
On Error GoTo Err

mPositionManagerGridMappingTable(index).isExpanded = True
OrderPlexGrid.row = mPositionManagerGridMappingTable(index).gridIndex
OrderPlexGrid.col = OPGridColumns.ExpandIndicator
OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture

symbol = OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(index).gridIndex, OPGridColumns.symbol)
i = mPositionManagerGridMappingTable(index).gridIndex + 1
Do While OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = symbol
    OrderPlexGrid.rowHeight(i) = -1
    lOpEntryIndex = OrderPlexGrid.rowdata(i) - RowDataOrderPlexBase
    i = expandOrderPlexEntry(lOpEntryIndex, True) + 1
Loop

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Function findOrderPlexTableIndex(ByVal op As OrderPlex) As Long
Dim opIndex As Long
Dim lOrder As Order
Dim symbol As String

' first make sure the relevant PositionManager entry is set up
Const ProcName As String = "findOrderPlexTableIndex"
Dim failpoint As String
On Error GoTo Err

findPositionManagerTableIndex op.PositionManager

symbol = op.Contract.Specifier.localSymbol
opIndex = op.IndexApplication
If opIndex > UBound(mOrderPlexGridMappingTable) Then
    ReDim Preserve mOrderPlexGridMappingTable(2 * (UBound(mOrderPlexGridMappingTable) + 1) - 1) As OrderPlexGridMappingEntry
End If
If opIndex > mMaxOrderPlexGridMappingTableIndex Then mMaxOrderPlexGridMappingTableIndex = opIndex

With mOrderPlexGridMappingTable(opIndex)
    If .op Is Nothing Then
        
        .isExpanded = True
        .entryGridOffset = -1
        .stopGridOffset = -1
        .targetGridOffset = -1
        .closeoutGridOffset = -1
        
        Set .op = op
        .gridIndex = addOrderPlexEntryToOrderPlexGrid(op.Contract.Specifier.localSymbol, opIndex)
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridOrderPlexColumns.CreationTime) = formattedTime(op.CreationTime)
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridOrderPlexColumns.currencyCode) = op.Contract.Specifier.currencyCode
        
        Set lOrder = op.entryOrder
        If Not lOrder Is Nothing Then
            .entryGridOffset = 1
            addOrderEntryToOrderPlexGrid .gridIndex + .entryGridOffset, _
                                    symbol, _
                                    lOrder, _
                                    opIndex, _
                                    "Entry"
        End If
        
        Set lOrder = op.stopOrder
        If Not lOrder Is Nothing Then
            If .entryGridOffset >= 0 Then
                .stopGridOffset = .entryGridOffset + 1
            Else
                .stopGridOffset = 1
            End If
            addOrderEntryToOrderPlexGrid .gridIndex + .stopGridOffset, _
                                    symbol, _
                                    lOrder, _
                                    opIndex, _
                                    "Stop"
        End If
        
        Set lOrder = op.targetOrder
        If Not lOrder Is Nothing Then
            If .stopGridOffset >= 0 Then
                .targetGridOffset = .stopGridOffset + 1
            ElseIf .entryGridOffset >= 0 Then
                .targetGridOffset = .entryGridOffset + 1
            Else
                .targetGridOffset = 1
            End If
            addOrderEntryToOrderPlexGrid .gridIndex + .targetGridOffset, _
                                    symbol, _
                                    lOrder, _
                                    opIndex, _
                                    "Target"
        End If
    End If
End With
findOrderPlexTableIndex = opIndex

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function findPositionManagerTableIndex(ByVal pm As PositionManager) As Long
Dim pmIndex As Long

Const ProcName As String = "findPositionManagerTableIndex"
Dim failpoint As String
On Error GoTo Err

pmIndex = pm.IndexApplication
Do While pmIndex > UBound(mPositionManagerGridMappingTable)
    ReDim Preserve mPositionManagerGridMappingTable(2 * (UBound(mPositionManagerGridMappingTable) + 1) - 1) As PositionManagerGridMappingEntry
Loop
If pmIndex > mMaxPositionManagerGridMappingTableIndex Then mMaxPositionManagerGridMappingTableIndex = pmIndex

findPositionManagerTableIndex = pmIndex

With mPositionManagerGridMappingTable(pmIndex)
    If .gridIndex = 0 Then
        .gridIndex = addEntryToOrderPlexGrid(pm.Ticker.Contract.Specifier.localSymbol, True)
        OrderPlexGrid.rowdata(.gridIndex) = pmIndex + RowDataPositionManagerBase
        OrderPlexGrid.row = .gridIndex
        OrderPlexGrid.col = 1
        OrderPlexGrid.colSel = OrderPlexGrid.Cols - 1
        OrderPlexGrid.FillStyle = FillStyleSettings.flexFillRepeat
        OrderPlexGrid.CellBackColor = &HC0C0C0
        OrderPlexGrid.CellForeColor = vbWhite
        OrderPlexGrid.CellFontBold = True
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridPositionColumns.exchange) = pm.Ticker.Contract.Specifier.exchange
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridPositionColumns.currencyCode) = pm.Ticker.Contract.Specifier.currencyCode
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridPositionColumns.Size) = pm.PositionSize
        OrderPlexGrid.col = OPGridColumns.ExpandIndicator
        OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
        Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture
        .isExpanded = True
    End If
End With

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function formattedTime(ByVal pTime As Date) As String
formattedTime = IIf(pTime = 0, _
                    "", _
                    IIf(Int(pTime) = Int(Now), _
                        FormatTimestamp(pTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs), _
                        FormatTimestamp(pTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)))
End Function

Private Sub handleEditingTerminationKey(ByVal KeyCode As Long)
Const ProcName As String = "handleEditingTerminationKey"
On Error GoTo Err

Select Case KeyCode
Case KeyCodeConstants.vbKeyEscape   ' ESC: hide, return focus to MSHFlexGrid.
    endEdit
Case KeyCodeConstants.vbKeyReturn   ' ENTER return focus to MSHFlexGrid.
    updateOrderPlex
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub invertEntryColors(ByVal rowNumber As Long)
Dim foreColor As Long
Dim backColor As Long
Dim i As Long

Const ProcName As String = "invertEntryColors"
Dim failpoint As String
On Error GoTo Err

If rowNumber < 0 Then Exit Sub

OrderPlexGrid.row = rowNumber

For i = OPGridColumns.OtherColumns To OrderPlexGrid.Cols - 1
    OrderPlexGrid.col = i
    foreColor = IIf(OrderPlexGrid.CellForeColor = 0, OrderPlexGrid.foreColor, OrderPlexGrid.CellForeColor)
    If foreColor = SystemColorConstants.vbWindowText Then
        OrderPlexGrid.CellForeColor = SystemColorConstants.vbHighlightText
    ElseIf foreColor = SystemColorConstants.vbHighlightText Then
        OrderPlexGrid.CellForeColor = SystemColorConstants.vbWindowText
    ElseIf foreColor > 0 Then
        OrderPlexGrid.CellForeColor = IIf((foreColor Xor &HFFFFFF) = 0, 1, foreColor Xor &HFFFFFF)
    End If
    
    backColor = IIf(OrderPlexGrid.CellBackColor = 0, OrderPlexGrid.backColor, OrderPlexGrid.CellBackColor)
    If backColor = SystemColorConstants.vbWindowBackground Then
        OrderPlexGrid.CellBackColor = SystemColorConstants.vbHighlight
    ElseIf backColor = SystemColorConstants.vbHighlight Then
        OrderPlexGrid.CellBackColor = SystemColorConstants.vbWindowBackground
    ElseIf backColor > 0 Then
        OrderPlexGrid.CellBackColor = IIf((backColor Xor &HFFFFFF) = 0, 1, backColor Xor &HFFFFFF)
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Sub listenForProfit( _
                ByVal pTicker As Ticker)
If mSimulated Then
    pTicker.PositionManagerSimulated.AddChangeListener Me
    pTicker.PositionManagerSimulated.AddProfitListener Me
    pTicker.PositionManagerSimulated.OrderPlexes.AddCollectionChangeListener Me
Else
    pTicker.PositionManager.AddChangeListener Me
    pTicker.PositionManager.AddProfitListener Me
    pTicker.PositionManager.OrderPlexes.AddCollectionChangeListener Me
End If
End Sub

Private Sub setupOrderPlexGrid()
Const ProcName As String = "setupOrderPlexGrid"
Dim failpoint As String
On Error GoTo Err

With OrderPlexGrid
    mSelectedOrderPlexGridRow = -1
    .AllowUserResizing = flexResizeBoth
    
    .Cols = 0
    .Rows = 20
    .FixedRows = 3
    ' .FixedCols = 1
    
    setupOrderPlexGridColumn 0, OPGridColumns.ExpandIndicator, OPGridColumnWidths.ExpandIndicatorWidth, "", True, AlignmentSettings.flexAlignCenterCenter
    setupOrderPlexGridColumn 0, OPGridColumns.symbol, OPGridColumnWidths.SymbolWidth, "Symbol", True, AlignmentSettings.flexAlignLeftCenter
    
    setupOrderPlexGridColumn 0, OPGridPositionColumns.currencyCode, OPGridPositionColumnWidths.CurrencyCodeWidth, "Curr", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.Drawdown, OPGridPositionColumnWidths.DrawdownWidth, "Drawdown", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.exchange, OPGridPositionColumnWidths.ExchangeWidth, "Exchange", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.MaxProfit, OPGridPositionColumnWidths.MaxProfitWidth, "Max", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.profit, OPGridPositionColumnWidths.ProfitWidth, "Profit", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.Size, OPGridPositionColumnWidths.SizeWidth, "Size", False, AlignmentSettings.flexAlignRightCenter
    
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.CreationTime, OPGridOrderPlexColumnWidths.CreationTimeWidth, "Creation Time", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.currencyCode, OPGridOrderPlexColumnWidths.CurrencyCodeWidth, "Curr", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.Drawdown, OPGridOrderPlexColumnWidths.DrawdownWidth, "Drawdown", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.MaxProfit, OPGridOrderPlexColumnWidths.MaxProfitWidth, "Max", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.profit, OPGridOrderPlexColumnWidths.ProfitWidth, "Profit", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.Size, OPGridOrderPlexColumnWidths.SizeWidth, "Size", False, AlignmentSettings.flexAlignRightCenter
    
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Action, OPGridOrderColumnWidths.ActionWidth, "Action", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.AuxPrice, OPGridOrderColumnWidths.AuxPriceWidth, "Trigger", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.AveragePrice, OPGridOrderColumnWidths.AveragePriceWidth, "Avg fill", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Id, OPGridOrderColumnWidths.IdWidth, "Id", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.LastFillPrice, OPGridOrderColumnWidths.LastFillPriceWidth, "Last fill", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.LastFillTime, OPGridOrderColumnWidths.LastFillTimeWidth, "Last fill time", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.OrderType, OPGridOrderColumnWidths.OrderTypeWidth, "Type", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Price, OPGridOrderColumnWidths.PriceWidth, "Price", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Quantity, OPGridOrderColumnWidths.QuantityWidth, "Qty", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.QuantityRemaining, OPGridOrderColumnWidths.QuantityRemainingWidth, "Rem", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Size, OPGridOrderColumnWidths.SizeWidth, "Size", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Status, OPGridOrderColumnWidths.StatusWidth, "Status", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.typeInPlex, OPGridOrderColumnWidths.TypeInPlexWidth, "Mode", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.BrokerId, OPGridOrderColumnWidths.BrokerIdWidth, "Broker Id", True, AlignmentSettings.flexAlignLeftCenter
    
    .MergeCells = flexMergeFree
    .MergeCol(OPGridColumns.symbol) = True
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightAlways
    .FocusRect = flexFocusNone
    
    mFirstOrderPlexGridRowIndex = 3
End With

EditText.Text = ""

mInitialised = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub setupOrderPlexGridColumn( _
                ByVal rowNumber As Long, _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As AlignmentSettings)
    
Dim lColumnWidth As Long
Dim i As Long

Const ProcName As String = "setupOrderPlexGridColumn"
Dim failpoint As String
On Error GoTo Err

With OrderPlexGrid
    .row = rowNumber
    If (columnNumber + 1) > .Cols Then
        For i = .Cols To columnNumber
            .Cols = i + 1
            .colWidth(i) = 0
        Next
    End If
    
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
    If .colWidth(columnNumber) < lColumnWidth Then
        .colWidth(columnNumber) = lColumnWidth
    End If
    
    .ColAlignment(columnNumber) = align
    .TextMatrix(rowNumber, columnNumber) = columnHeader
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub updateOrderPlex()
Dim orderNumber As Long
Dim Price As Double

Const ProcName As String = "updateOrderPlex"
Dim failpoint As String
On Error GoTo Err

If Not EditText.Visible Then Exit Sub

orderNumber = mSelectedOrderPlexGridRow - mOrderPlexGridMappingTable(OrderPlexGrid.rowdata(OrderPlexGrid.row) - RowDataOrderPlexBase).gridIndex
If OrderPlexGrid.col = OPGridOrderColumns.Price Then
    If mSelectedOrderPlex.Contract.ParsePrice(EditText.Text, Price) Then
        mSelectedOrderPlex.NewOrderPrice(orderNumber) = Price
    End If
ElseIf OrderPlexGrid.col = OPGridOrderColumns.AuxPrice Then
    If mSelectedOrderPlex.Contract.ParsePrice(EditText.Text, Price) Then
        mSelectedOrderPlex.NewOrderTriggerPrice(orderNumber) = Price
    End If
ElseIf OrderPlexGrid.col = OPGridOrderColumns.Quantity Then
    If IsNumeric(EditText.Text) Then
        mSelectedOrderPlex.NewQuantity = EditText.Text
    End If
End If
    
If mSelectedOrderPlex.dirty Then mSelectedOrderPlex.Update

endEdit

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

