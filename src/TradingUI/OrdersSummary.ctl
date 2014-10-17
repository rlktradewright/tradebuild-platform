VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
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
   Begin MSComctlLib.ImageList BracketOrderImageList 
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
   Begin MSFlexGridLib.MSFlexGrid BracketOrderGrid 
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
      BackColorFixed  =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   14737632
      GridLinesFixed  =   1
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
Implements IProfitListener

'@================================================================================
' Events
'@================================================================================

Event Click()
Event SelectionChanged()
                
'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "OrdersSummary"

Private Const RowDataBracketOrderBase       As Long = &H100
Private Const RowDataPositionManagerBase    As Long = &H1000000

'@================================================================================
' Enums
'@================================================================================

Private Enum BracketOrderGridColumns
    Symbol
    ExpandIndicator
    OtherColumns    ' keep this entry last
    
    'BracketOrder Columns
    BracketCreationTime = OtherColumns
    BracketSize
    BracketProfit
    BracketMaxProfit
    BracketDrawdown
    BracketCurrencyCode
    
    'Position Columns
    PositionExchange = OtherColumns
    PositionSize
    PositionProfit
    PositionMaxProfit
    PositionDrawdown
    PositionCurrencyCode

    'Order Columns
    OrderMode = OtherColumns
    OrderAction
    OrderQuantity
    OrderType
    OrderPrice
    OrderAuxPrice
    OrderStatus
    OrderSize
    OrderQuantityRemaining
    OrderAveragePrice
    OrderLastFillTime
    OrderLastFillPrice
    OrderId
    OrderBrokerId
End Enum

Private Enum BracketOrderGridColumnWidths
    ExpandIndicatorWidth = 3
    SymbolWidth = 15

    BracketCreationTimeWidth = 15
    BracketSizeWidth = 6
    BracketProfitWidth = 9
    BracketMaxProfitWidth = 9
    BracketDrawdownWidth = 9
    BracketCurrencyCodeWidth = 4

    PositionExchangeWidth = 9
    PositionSizeWidth = 6
    PositionProfitWidth = 9
    PositionMaxProfitWidth = 9
    PositionDrawdownWidth = 9
    PositionCurrencyCodeWidth = 5

    OrderModeWidth = 9
    OrderSizeWidth = 6
    OrderAveragePriceWidth = 9
    OrderStatusWidth = 13
    OrderActionWidth = 4
    OrderQuantityWidth = 6
    OrderQuantityRemainingWidth = 5
    OrderTypeWidth = 5
    OrderPriceWidth = 9
    OrderAuxPriceWidth = 9
    OrderLastFillTimeWidth = 15
    OrderLastFillPriceWidth = 9
    OrderIdWidth = 40
    OrderBrokerIdWidth = 11
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type BracketOrderGridMappingEntry
    BracketOrder        As IBracketOrder
    ProfitCalculator    As BracketProfitCalculator
    
    ' indicates whether this entry in the grid is expanded
    IsExpanded          As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    GridIndex           As Long
                                
    ' offset from gridIndex of line in OrdersGrid relating to
    ' the corresponding order: -1 means  it's not in the grid
    EntryGridOffset     As Long
    StopLossGridOffset  As Long
    TargetGridOffset    As Long
    CloseoutGridOffset  As Long
    TickSize            As Double
    secType             As SecurityTypes
    
End Type

Private Type PositionManagerGridMappingEntry
    
    ' indicates whether this entry in the grid is expanded
    IsExpanded          As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    GridIndex           As Long
                                
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManager                                  As IMarketDataManager

Private mSelectedBracketOrder                               As IBracketOrder

Private mBracketOrderGridMappingTable()                     As BracketOrderGridMappingEntry
Private mMaxBracketOrderGridMappingTableIndex               As Long

Private mPositionManagerGridMappingTable()                  As PositionManagerGridMappingEntry
Private mMaxPositionManagerGridMappingTableIndex            As Long

' the index of the first entry in the bracket order grid that relates to
' bracket orders (rather than header rows, currency totals etc)
Private mFirstBracketOrderGridRowIndex                      As Long

Private mLetterWidth                                        As Single
Private mDigitWidth                                         As Single

Private mPositionManagersCollection                         As New EnumerableCollection
Private mMonitoredPositions                                 As New EnumerableCollection
    
Private mIsEditing                                          As Boolean
Private mEditedCol                                          As Long

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
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

Dim widthString As String

widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = UserControl.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = UserControl.TextWidth(widthString) / Len(widthString)

setupBracketOrderGrid

ReDim mBracketOrderGridMappingTable(3) As BracketOrderGridMappingEntry
mMaxBracketOrderGridMappingTableIndex = -1

ReDim mPositionManagerGridMappingTable(3) As PositionManagerGridMappingEntry
mMaxPositionManagerGridMappingTableIndex = -1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

BracketOrderGrid.Width = UserControl.Width
BracketOrderGrid.Height = UserControl.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Terminate()
Debug.Print "OrdersSummary control terminated"
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

If TypeOf ev.Source Is IBracketOrder Then
    Dim lBracketOrder As IBracketOrder
    Set lBracketOrder = ev.Source
    
    Dim lBracketOrderIndex As Long
    lBracketOrderIndex = findBracketOrderGridMappingIndex(lBracketOrder)
    
    With mBracketOrderGridMappingTable(lBracketOrderIndex)
    
        Dim lBracketOrderChangeType As BracketOrderChangeTypes
        lBracketOrderChangeType = ev.changeType
        
        Select Case lBracketOrderChangeType
        Case BracketOrderChangeTypes.BracketOrderCreated
            
        Case BracketOrderChangeTypes.BracketOrderCompleted
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            If lBracketOrder.Size = 0 Then lBracketOrder.RemoveChangeListener Me
        Case BracketOrderChangeTypes.BracketOrderSelfCancelled
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            If lBracketOrder.Size = 0 Then lBracketOrder.RemoveChangeListener Me
        Case BracketOrderChangeTypes.BracketOrderEntryOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .EntryGridOffset, lBracketOrder.EntryOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderStopOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .StopLossGridOffset, lBracketOrder.StopLossOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderTargetOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .TargetGridOffset, lBracketOrder.TargetOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderCloseoutOrderCreated
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            If .TargetGridOffset >= 0 Then
                .CloseoutGridOffset = .TargetGridOffset + 1
            ElseIf .StopLossGridOffset >= 0 Then
                .CloseoutGridOffset = .StopLossGridOffset + 1
            ElseIf .EntryGridOffset >= 0 Then
                .CloseoutGridOffset = .EntryGridOffset + 1
            Else
                .CloseoutGridOffset = 1
            End If
            
            addOrderEntryToBracketOrderGrid .GridIndex + .CloseoutGridOffset, _
                                    .BracketOrder.Contract.Specifier.LocalSymbol, _
                                    lBracketOrder.CloseoutOrder, _
                                    lBracketOrderIndex, _
                                    "Closeout", _
                                    .secType, _
                                    .TickSize
        Case BracketOrderChangeTypes.BracketOrderCloseoutOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .CloseoutGridOffset, lBracketOrder.CloseoutOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderSizeChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            GridColumn(.GridIndex, BracketSize) = lBracketOrder.Size
        Case BracketOrderChangeTypes.BracketOrderStateChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            If lBracketOrder.State = BracketOrderStates.BracketOrderStateSubmitted Then
                GridColumn(.GridIndex, BracketCreationTime) = formattedTime(lBracketOrder.CreationTime)
            End If
            If lBracketOrder.State <> BracketOrderStates.BracketOrderStateCreated And _
                lBracketOrder.State <> BracketOrderStates.BracketOrderStateSubmitted _
            Then
                ' the bracket order is now in a state where it can't be modified.
                ' If it's the currently selected bracket order, make it not so.
                If lBracketOrder Is mSelectedBracketOrder Then
                    invertEntryColors getBracketOrderGridIndex(mSelectedBracketOrder)
                    Set mSelectedBracketOrder = Nothing
                    RaiseEvent SelectionChanged
                End If
            End If
        End Select
    End With
ElseIf TypeOf ev.Source Is PositionManager Then
    Dim pm As PositionManager
    Set pm = ev.Source
    
    Dim pmChangeType As PositionManagerChangeTypes
    pmChangeType = ev.changeType
    Select Case pmChangeType
    Case PositionManagerChangeTypes.PositionSizeChanged
        Dim pmIndex As Long
        pmIndex = findPositionManagerGridMappingIndex(pm)
        showPositionManagerEntry pm
        GridColumn(mPositionManagerGridMappingTable(pmIndex).GridIndex, _
                                PositionSize) = pm.PositionSize
    End Select
End If

adjustEditBox

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEventData)
0 Const ProcName As String = "CollectionChangeListener_Change"
On Error GoTo Err

If TypeOf ev.Source Is BracketOrders Then
    If IsEmpty(ev.AffectedItem) Then Exit Sub
    
    Dim lBracketOrder As IBracketOrder
    Set lBracketOrder = ev.AffectedItem
    
    Select Case ev.changeType
    Case CollItemAdded
        Dim lPm As PositionManager
        For Each lPm In mMonitoredPositions
            If lPm.BracketOrders Is ev.Source Then
                showPositionManagerEntry lPm
                Exit For
            End If
        Next
    
        addBracketOrder lBracketOrder, lPm
    Case CollItemRemoved
        Dim lBracketOrderIndex As Long
        lBracketOrderIndex = findBracketOrderGridMappingIndex(lBracketOrder)
        lBracketOrder.RemoveChangeListener Me
        mBracketOrderGridMappingTable(lBracketOrderIndex).ProfitCalculator.RemoveProfitListener Me
        Set mBracketOrderGridMappingTable(lBracketOrderIndex).ProfitCalculator = Nothing
    End Select
ElseIf TypeOf ev.Source Is PositionManagers Then
    If ev.changeType <> CollItemAdded Then Exit Sub
    
    Dim lPositionManager As PositionManager
    Set lPositionManager = ev.AffectedItem
    lPositionManager.AddChangeListener Me
    lPositionManager.AddProfitListener Me
    lPositionManager.BracketOrders.AddCollectionChangeListener Me
    mMonitoredPositions.Add lPositionManager
End If

adjustEditBox

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IProfitListener Interface Members
'@================================================================================

Private Sub IProfitListener_NotifyProfit(ev As ProfitEventData)
Const ProcName As String = "IProfitListener_NotifyProfit"
On Error GoTo Err

Dim rowIndex As Long

If TypeOf ev.Source Is BracketProfitCalculator Then
    Dim lProfitCalculator As BracketProfitCalculator
    Set lProfitCalculator = ev.Source
    
    Dim lBracketOrder As IBracketOrder
    Set lBracketOrder = lProfitCalculator.BracketOrder
    
    Dim lBracketOrderIndex As Long
    lBracketOrderIndex = findBracketOrderGridMappingIndex(lBracketOrder)
    rowIndex = mBracketOrderGridMappingTable(lBracketOrderIndex).GridIndex
    
    Dim lBOProfitType As ProfitTypes
    lBOProfitType = ev.ProfitTypes
    
    If lBOProfitType And ProfitTypes.ProfitTypeProfit Then _
        displayProfitValue lProfitCalculator.Profit, rowIndex, BracketProfit
    If lBOProfitType And ProfitTypes.ProfitTypeMaxProfit Then _
        displayProfitValue lProfitCalculator.MaxProfit, rowIndex, BracketMaxProfit
    If lBOProfitType And ProfitTypes.ProfitTypeDrawdown Then _
        displayProfitValue -lProfitCalculator.Drawdown, rowIndex, BracketDrawdown

ElseIf TypeOf ev.Source Is PositionManager Then
    Dim lPositionManager As PositionManager
    Set lPositionManager = ev.Source
    
    showPositionManagerEntry lPositionManager
    
    Dim lPositionManagerIndex As Long
    lPositionManagerIndex = findPositionManagerGridMappingIndex(lPositionManager)
    rowIndex = mPositionManagerGridMappingTable(lPositionManagerIndex).GridIndex
    
    Dim lPMProfitType As ProfitTypes
    lPMProfitType = ev.ProfitTypes
    
    If lPMProfitType Or ProfitTypes.ProfitTypeSessionProfit Then _
        displayProfitValue lPositionManager.Profit, rowIndex, PositionProfit
    If lPMProfitType Or ProfitTypes.ProfitTypeSessionMaxProfit Then _
        displayProfitValue lPositionManager.MaxProfit, rowIndex, PositionMaxProfit
    If lPMProfitType Or ProfitTypes.ProfitTypeSessionDrawdown Then _
        displayProfitValue -lPositionManager.Drawdown, rowIndex, PositionDrawdown
End If

adjustEditBox

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Form Control Event Handlers
'@================================================================================

Private Sub EditText_KeyDown(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "EditText_KeyDown"

On Error GoTo Err

handleEditingTerminationKey KeyCode

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub EditText_KeyPress(KeyAscii As Integer)
' Delete returns to get rid of beep.
If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub

Private Sub BracketOrderGrid_Click()
Const ProcName As String = "BracketOrderGrid_Click"
On Error GoTo Err

Dim lSelectionChanged As Boolean

If BracketOrderGrid.MouseCol = Symbol Then
    RaiseEvent Click
    Exit Sub
End If

If BracketOrderGrid.MouseCol = ExpandIndicator Then
    expandOrContract
    adjustEditBox
    RaiseEvent Click
    Exit Sub
End If

BracketOrderGrid.col = BracketOrderGrid.MouseCol

Dim lBracketOrder As IBracketOrder
Set lBracketOrder = getSelectedBracketOrder

If Not mSelectedBracketOrder Is Nothing Then
    invertEntryColors getBracketOrderGridIndex(mSelectedBracketOrder)
    Set mSelectedBracketOrder = Nothing
    lSelectionChanged = True
End If

If lBracketOrder Is Nothing Then
    RaiseEvent Click
    If lSelectionChanged Then RaiseEvent SelectionChanged
    Exit Sub
End If

If lBracketOrder.State <> BracketOrderStates.BracketOrderStateCreated And _
    lBracketOrder.State <> BracketOrderStates.BracketOrderStateSubmitted _
Then
    RaiseEvent Click
    If lSelectionChanged Then RaiseEvent SelectionChanged
    Exit Sub
End If

Set mSelectedBracketOrder = lBracketOrder
lSelectionChanged = True
invertEntryColors getBracketOrderGridIndex(mSelectedBracketOrder)

Dim lSelectedOrder As IOrder
Set lSelectedOrder = getSelectedOrder
    
If lSelectedOrder Is Nothing Then
    RaiseEvent Click
    RaiseEvent SelectionChanged
    Exit Sub
End If

If Not lSelectedOrder.IsModifiable Then
    RaiseEvent Click
    RaiseEvent SelectionChanged
    Exit Sub
End If

If (BracketOrderGrid.col = OrderPrice And _
        lSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttLimitPrice)) Or _
    (BracketOrderGrid.col = OrderAuxPrice And _
        lSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttTriggerPrice)) Or _
    (BracketOrderGrid.col = OrderQuantity And _
        lSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttQuantity)) _
Then
    mIsEditing = True
    mEditedCol = BracketOrderGrid.col
    BracketOrderGrid.col = mEditedCol
    
    EditText.Text = BracketOrderGrid.Text
    EditText.SelStart = 0
    EditText.SelLength = Len(EditText.Text)
    EditText.Visible = True
    EditText.SetFocus
                            
    adjustEditBox

End If
    
RaiseEvent Click
RaiseEvent SelectionChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BracketOrderGrid_Scroll()
Const ProcName As String = "BracketOrderGrid_Scroll"
On Error GoTo Err

adjustEditBox

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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
On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsEditing() As Boolean
IsEditing = mIsEditing
End Property

Public Property Get IsSelectedItemModifiable() As Boolean
Const ProcName As String = "IsSelectedItemModifiable"
On Error GoTo Err

If mSelectedBracketOrder Is Nothing Then Exit Property

Dim lSelectedOrder As IOrder
Set lSelectedOrder = getSelectedOrder
If Not lSelectedOrder Is Nothing Then
    IsSelectedItemModifiable = lSelectedOrder.IsModifiable
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelectedItem() As IBracketOrder
Set SelectedItem = mSelectedBracketOrder
End Property

Public Property Get SelectedOrderIndex() As Long
SelectedOrderIndex = getBracketOrderGridIndex(mSelectedBracketOrder)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mMarketDataManager = Nothing

Dim i As Long
For i = 0 To mMaxBracketOrderGridMappingTableIndex
    If Not mBracketOrderGridMappingTable(i).BracketOrder Is Nothing Then
        mBracketOrderGridMappingTable(i).BracketOrder.RemoveChangeListener Me
        mBracketOrderGridMappingTable(i).ProfitCalculator.RemoveProfitListener Me
        Set mBracketOrderGridMappingTable(i).BracketOrder = Nothing
        Set mBracketOrderGridMappingTable(i).ProfitCalculator = Nothing
        mBracketOrderGridMappingTable(i).CloseoutGridOffset = 0
        mBracketOrderGridMappingTable(i).EntryGridOffset = 0
        mBracketOrderGridMappingTable(i).GridIndex = 0
        mBracketOrderGridMappingTable(i).IsExpanded = False
        mBracketOrderGridMappingTable(i).secType = SecTypeNone
        mBracketOrderGridMappingTable(i).StopLossGridOffset = 0
        mBracketOrderGridMappingTable(i).TargetGridOffset = 0
        mBracketOrderGridMappingTable(i).TickSize = 0#
    End If
Next

mMaxBracketOrderGridMappingTableIndex = 0
mMaxPositionManagerGridMappingTableIndex = 0
mFirstBracketOrderGridRowIndex = 0

Dim lPositionManager As PositionManager
For Each lPositionManager In mMonitoredPositions
    If Not lPositionManager.IsFinished Then
        lPositionManager.RemoveProfitListener Me
        lPositionManager.RemoveChangeListener Me
        lPositionManager.BracketOrders.RemoveCollectionChangeListener Me
    End If
Next
mMonitoredPositions.Clear

Dim lPositionManagers As PositionManagers
For Each lPositionManagers In mPositionManagersCollection
    lPositionManagers.RemoveCollectionChangeListener Me
Next
mPositionManagersCollection.Clear

BracketOrderGrid.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise(ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pMarketDataManager Is Nothing, "pMarketDataManager must be supplied"

Finish

Set mMarketDataManager = pMarketDataManager

setupBracketOrderGrid

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub MonitorPositions( _
                ByVal pPositionManagers As PositionManagers)
Const ProcName As String = "MonitorPosition"
On Error GoTo Err

Assert Not mMarketDataManager Is Nothing, "Initialise method has not been called"

pPositionManagers.AddCollectionChangeListener Me
mPositionManagersCollection.Add pPositionManagers

Dim lPositionManager As PositionManager
For Each lPositionManager In pPositionManagers
    mMonitoredPositions.Add lPositionManager

    If lPositionManager.BracketOrders.Count <> 0 Or lPositionManager.PositionSize <> 0 Or lPositionManager.PendingPositionSize <> 0 Then
        showPositionManagerEntry lPositionManager
    End If
    
    lPositionManager.AddChangeListener Me
    lPositionManager.AddProfitListener Me
    lPositionManager.BracketOrders.AddCollectionChangeListener Me
    
    Dim lAnyActiveBracketOrders As Boolean
    Dim lBracketOrder As IBracketOrder
    For Each lBracketOrder In lPositionManager.BracketOrders
        addBracketOrder lBracketOrder, lPositionManager
        If lBracketOrder.State = BracketOrderStateClosed Then
            contractBracketOrderEntry findBracketOrderGridMappingIndex(lBracketOrder)
        Else
            lAnyActiveBracketOrders = True
        End If
    Next
    
    If Not lAnyActiveBracketOrders Then contractPositionManagerEntry findPositionManagerGridMappingIndex(lPositionManager)
    
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addBracketOrder(ByVal pBracketOrder As IBracketOrder, ByVal pPositionManager As PositionManager)
Const ProcName As String = "addBracketOrder"
On Error GoTo Err

pBracketOrder.AddChangeListener Me
displayBracketOrder pBracketOrder, pPositionManager

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function addBracketOrderEntryToBracketOrderGrid( _
                ByVal pSymbol As String, _
                ByVal pBracketOrderGridMappingIndex As Long, _
                ByVal pPositionManagerGridIndex As Long) As Long
Const ProcName As String = "addBracketOrderEntryToBracketOrderGrid"
On Error GoTo Err

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

Dim index As Long
index = addEntryToBracketOrderGrid(pPositionManagerGridIndex + 1, pSymbol, True)

BracketOrderGrid.RowData(index) = pBracketOrderGridMappingIndex + RowDataBracketOrderBase

BracketOrderGrid.Row = index
BracketOrderGrid.col = ExpandIndicator
BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture

BracketOrderGrid.col = BracketProfit
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

BracketOrderGrid.col = BracketMaxProfit
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

BracketOrderGrid.col = BracketDrawdown
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

addBracketOrderEntryToBracketOrderGrid = index

BracketOrderGrid.Row = lPrevRow
BracketOrderGrid.col = lPrevCol

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                
Private Function addEntryToBracketOrderGrid( _
                ByVal pStartIndex As Long, _
                ByVal pSymbol As String, _
                Optional ByVal pBefore As Boolean, _
                Optional ByVal pIndex As Long = -1) As Long
Const ProcName As String = "addEntryToBracketOrderGrid"
On Error GoTo Err

Dim i As Long

If pStartIndex = 0 Then pStartIndex = mFirstBracketOrderGridRowIndex

If pIndex < 0 Then
    For i = pStartIndex To BracketOrderGrid.Rows - 1
        If (pBefore And _
            GridColumn(i, Symbol) >= pSymbol) Or _
            GridColumn(i, Symbol) = "" _
        Then
            pIndex = i
            Exit For
        ElseIf (Not pBefore And _
            GridColumn(i, Symbol) > pSymbol) Or _
            GridColumn(i, Symbol) = "" _
        Then
            pIndex = i
            Exit For
        End If
    Next
    
    If pIndex < 0 Then
        BracketOrderGrid.addItem ""
        pIndex = BracketOrderGrid.Rows - 1
    ElseIf GridColumn(pIndex, Symbol) = "" Then
        GridColumn(pIndex, Symbol) = pSymbol
    Else
        BracketOrderGrid.addItem "", pIndex
    End If
Else
    BracketOrderGrid.addItem "", pIndex
End If

GridColumn(pIndex, Symbol) = pSymbol
If pIndex < BracketOrderGrid.Rows - 1 Then
    ' this new entry has displaced one or more existing entries so
    ' the BracketOrderGridMappingTable and PositionManageGridMappingTable indexes
    ' need to be adjusted
    For i = 0 To mMaxBracketOrderGridMappingTableIndex
        If mBracketOrderGridMappingTable(i).GridIndex >= pIndex Then
            mBracketOrderGridMappingTable(i).GridIndex = mBracketOrderGridMappingTable(i).GridIndex + 1
        End If
    Next
    For i = 0 To mMaxPositionManagerGridMappingTableIndex
        If mPositionManagerGridMappingTable(i).GridIndex >= pIndex Then
            mPositionManagerGridMappingTable(i).GridIndex = mPositionManagerGridMappingTable(i).GridIndex + 1
        End If
    Next
    If pIndex <= BracketOrderGrid.Row Then BracketOrderGrid.Row = BracketOrderGrid.Row + 1
End If

addEntryToBracketOrderGrid = pIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub addOrderEntryToBracketOrderGrid( _
                ByVal pIndex As Long, _
                ByVal pSymbol As String, _
                ByVal pOrder As IOrder, _
                ByVal pBracketOrderTableIndex As Long, _
                ByVal pOrderMode As String, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double)
Const ProcName As String = "addOrderEntryToBracketOrderGrid"
On Error GoTo Err

pIndex = addEntryToBracketOrderGrid(0, pSymbol, False, pIndex)

BracketOrderGrid.RowData(pIndex) = pBracketOrderTableIndex + RowDataBracketOrderBase

GridColumn(pIndex, OrderMode) = pOrderMode

displayOrderValues pIndex, pOrder, pSecType, pTickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub adjustEditBox()
Const ProcName As String = "adjustEditBox"
On Error GoTo Err

If Not mIsEditing Then Exit Sub

'Dim opIndex As Long
'opIndex = findBracketOrderGridMappingIndex(mSelectedBracketOrder)
'BracketOrderGrid.Row = mBracketOrderGridMappingTable(opIndex).GridIndex + mEditedOrderIndex
'BracketOrderGrid.col = mEditedCol

EditText.Move BracketOrderGrid.Left + BracketOrderGrid.CellLeft + 8, _
            BracketOrderGrid.Top + BracketOrderGrid.Celltop + 8, _
            BracketOrderGrid.CellWidth - 16, _
            BracketOrderGrid.CellHeight - 16

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function contractBracketOrderEntry( _
                ByVal index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Const ProcName As String = "contractBracketOrderEntry"
On Error GoTo Err

Dim lIndex As Long

With mBracketOrderGridMappingTable(index)
    
    If mIsEditing And .BracketOrder Is mSelectedBracketOrder Then endEdit
    
    If .EntryGridOffset >= 0 Then
        lIndex = .GridIndex + .EntryGridOffset
        BracketOrderGrid.rowHeight(lIndex) = 0
    End If
    If .StopLossGridOffset >= 0 Then
        lIndex = .GridIndex + .StopLossGridOffset
        BracketOrderGrid.rowHeight(lIndex) = 0
    End If
    If .TargetGridOffset >= 0 Then
        lIndex = .GridIndex + .TargetGridOffset
        BracketOrderGrid.rowHeight(lIndex) = 0
    End If
    If .CloseoutGridOffset >= 0 Then
        lIndex = .GridIndex + .CloseoutGridOffset
        BracketOrderGrid.rowHeight(lIndex) = 0
    End If
    
    If Not preserveCurrentExpandedState Then
        Dim lPrevRow As Long
        lPrevRow = BracketOrderGrid.Row
        Dim lPrevCol As Long
        lPrevCol = BracketOrderGrid.col
        
        .IsExpanded = False
        BracketOrderGrid.Row = .GridIndex
        BracketOrderGrid.col = ExpandIndicator
        BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Expand").Picture
        
        BracketOrderGrid.Row = lPrevRow
        BracketOrderGrid.col = lPrevCol
    End If
End With

contractBracketOrderEntry = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub contractPositionManagerEntry(ByVal index As Long)
Const ProcName As String = "contractPositionManagerEntry"
On Error GoTo Err

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

mPositionManagerGridMappingTable(index).IsExpanded = False
BracketOrderGrid.Row = mPositionManagerGridMappingTable(index).GridIndex
BracketOrderGrid.col = ExpandIndicator
BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Expand").Picture

Dim lSymbol As String
lSymbol = GridColumn(mPositionManagerGridMappingTable(index).GridIndex, Symbol)

Dim i As Long
i = mPositionManagerGridMappingTable(index).GridIndex + 1
Do While GridColumn(i, Symbol) = lSymbol And gridRowIsBracketOrder(i)
    BracketOrderGrid.rowHeight(i) = 0
    
    Dim lBracketOrderIndex As Long
    lBracketOrderIndex = getBracketOrderGridMappingIndexFromRowIndex(i)
    i = contractBracketOrderEntry(lBracketOrderIndex, True) + 1
Loop

BracketOrderGrid.Row = lPrevRow
BracketOrderGrid.col = lPrevCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displayBracketOrder(ByVal pBracketOrder As IBracketOrder, ByVal pPositionManager As PositionManager)
Const ProcName As String = "displayBracketOrder"
On Error GoTo Err

Dim lSymbol As String
lSymbol = pBracketOrder.Contract.Specifier.LocalSymbol

Dim lIndex As Long
lIndex = findBracketOrderGridMappingIndex(pBracketOrder)

With mBracketOrderGridMappingTable(lIndex)
    If .BracketOrder Is Nothing Then
        
        .IsExpanded = True
        .EntryGridOffset = -1
        .StopLossGridOffset = -1
        .TargetGridOffset = -1
        .CloseoutGridOffset = -1
        
        Set .BracketOrder = pBracketOrder
        .TickSize = pBracketOrder.Contract.TickSize
        .secType = pBracketOrder.Contract.Specifier.secType
        .GridIndex = addBracketOrderEntryToBracketOrderGrid( _
                                pBracketOrder.Contract.Specifier.LocalSymbol, _
                                lIndex, _
                                mPositionManagerGridMappingTable(findPositionManagerGridMappingIndex(pPositionManager)).GridIndex)
        GridColumn(.GridIndex, BracketCreationTime) = formattedTime(pBracketOrder.CreationTime)
        GridColumn(.GridIndex, BracketCurrencyCode) = pBracketOrder.Contract.Specifier.CurrencyCode
        
        Dim lDataSource As IMarketDataSource
        Set lDataSource = pPositionManager.DataSource
        Set .ProfitCalculator = CreateBracketProfitCalculator(pBracketOrder, lDataSource)
        .ProfitCalculator.AddProfitListener Me
        
        Dim lOrder As IOrder
        Set lOrder = pBracketOrder.EntryOrder
        If Not lOrder Is Nothing Then
            .EntryGridOffset = 1
            addOrderEntryToBracketOrderGrid .GridIndex + .EntryGridOffset, _
                                    lSymbol, _
                                    lOrder, _
                                    lIndex, _
                                    "Entry", _
                                    .secType, _
                                    .TickSize
        End If
        
        Set lOrder = pBracketOrder.StopLossOrder
        If Not lOrder Is Nothing Then
            If .EntryGridOffset >= 0 Then
                .StopLossGridOffset = .EntryGridOffset + 1
            Else
                .StopLossGridOffset = 1
            End If
            addOrderEntryToBracketOrderGrid .GridIndex + .StopLossGridOffset, _
                                    lSymbol, _
                                    lOrder, _
                                    lIndex, _
                                    "Stop Loss", _
                                    .secType, _
                                    .TickSize
        End If
        
        Set lOrder = pBracketOrder.TargetOrder
        If Not lOrder Is Nothing Then
            If .StopLossGridOffset >= 0 Then
                .TargetGridOffset = .StopLossGridOffset + 1
            ElseIf .EntryGridOffset >= 0 Then
                .TargetGridOffset = .EntryGridOffset + 1
            Else
                .TargetGridOffset = 1
            End If
            addOrderEntryToBracketOrderGrid .GridIndex + .TargetGridOffset, _
                                    lSymbol, _
                                    lOrder, _
                                    lIndex, _
                                    "Target", _
                                    .secType, _
                                    .TickSize
        End If
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displayOrderValues( _
                ByVal pGridIndex As Long, _
                ByVal pOrder As IOrder, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double)
Const ProcName As String = "displayOrderValues"
On Error GoTo Err

GridColumn(pGridIndex, OrderAction) = OrderActionToString(pOrder.Action)
GridColumn(pGridIndex, OrderAuxPrice) = FormatPrice(pOrder.TriggerPrice, pSecType, pTickSize)
GridColumn(pGridIndex, OrderAveragePrice) = FormatPrice(pOrder.AveragePrice, pSecType, pTickSize)
GridColumn(pGridIndex, OrderId) = pOrder.Id
GridColumn(pGridIndex, OrderLastFillPrice) = FormatPrice(pOrder.LastFillPrice, pSecType, pTickSize)
GridColumn(pGridIndex, OrderLastFillTime) = formattedTime(pOrder.FillTime)
GridColumn(pGridIndex, OrderType) = OrderTypeToShortString(pOrder.OrderType)
GridColumn(pGridIndex, OrderPrice) = FormatPrice(pOrder.LimitPrice, pSecType, pTickSize)
GridColumn(pGridIndex, OrderQuantity) = pOrder.Quantity
GridColumn(pGridIndex, OrderQuantityRemaining) = pOrder.QuantityRemaining
GridColumn(pGridIndex, OrderSize) = IIf(pOrder.QuantityFilled <> 0, pOrder.QuantityFilled, 0)
GridColumn(pGridIndex, OrderStatus) = OrderStatusToString(pOrder.Status)
GridColumn(pGridIndex, OrderBrokerId) = pOrder.BrokerId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displayProfitValue( _
                ByVal pProfit As Currency, _
                ByVal pRowIndex As Long, _
                ByVal pColIndex As Long)
Const ProcName As String = "displayProfitValue"
On Error GoTo Err

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

BracketOrderGrid.Row = pRowIndex
BracketOrderGrid.col = pColIndex
BracketOrderGrid.Text = Format(pProfit, "0.00")
If pProfit >= 0 Then
    BracketOrderGrid.CellForeColor = CPositiveProfitColor
Else
    BracketOrderGrid.CellForeColor = CNegativeProfitColor
End If

BracketOrderGrid.Row = lPrevRow
BracketOrderGrid.col = lPrevCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub endEdit()
Const ProcName As String = "endEdit"
On Error GoTo Err

EditText.Text = ""
EditText.Visible = False
mIsEditing = False
mEditedCol = -1
BracketOrderGrid.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Private Sub expandOrContract()
Const ProcName As String = "expandOrContract"
On Error GoTo Err

Dim index As Long
Dim expanded As Boolean

If gridRowIsPositionManager(BracketOrderGrid.Row) Then
    index = getPositionManagerGridMappingIndexFromRowIndex(BracketOrderGrid.Row)
    expanded = mPositionManagerGridMappingTable(index).IsExpanded
    If expanded Then
        contractPositionManagerEntry index
    Else
        expandPositionManagerEntry index
    End If
ElseIf gridRowIsBracketOrder(BracketOrderGrid.Row) Then
    index = getBracketOrderGridMappingIndexFromRowIndex(BracketOrderGrid.Row)
    expanded = mBracketOrderGridMappingTable(index).IsExpanded
    If BracketOrderGrid.Row <> mBracketOrderGridMappingTable(index).GridIndex Then
        ' clicked on an order entry
        Exit Sub
    End If
    If expanded Then
        contractBracketOrderEntry index
    Else
        expandBracketOrderEntry index
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function expandBracketOrderEntry( _
                ByVal index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Const ProcName As String = "expandBracketOrderEntry"
On Error GoTo Err

Dim lIndex As Long

With mBracketOrderGridMappingTable(index)
    
    If .EntryGridOffset >= 0 Then
        lIndex = .GridIndex + .EntryGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then BracketOrderGrid.rowHeight(lIndex) = -1
    End If
    If .StopLossGridOffset >= 0 Then
        lIndex = .GridIndex + .StopLossGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then BracketOrderGrid.rowHeight(lIndex) = -1
    End If
    If .TargetGridOffset >= 0 Then
        lIndex = .GridIndex + .TargetGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then BracketOrderGrid.rowHeight(lIndex) = -1
    End If
    If .CloseoutGridOffset >= 0 Then
        lIndex = .GridIndex + .CloseoutGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then BracketOrderGrid.rowHeight(lIndex) = -1
    End If
    
    If Not preserveCurrentExpandedState Then
        Dim lPrevRow As Long
        lPrevRow = BracketOrderGrid.Row
        Dim lPrevCol As Long
        lPrevCol = BracketOrderGrid.col
        
        .IsExpanded = True
        BracketOrderGrid.Row = .GridIndex
        BracketOrderGrid.col = ExpandIndicator
        BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture
    
        BracketOrderGrid.Row = lPrevRow
        BracketOrderGrid.col = lPrevRow
    End If
End With

expandBracketOrderEntry = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub expandPositionManagerEntry(ByVal index As Long)
Const ProcName As String = "expandPositionManagerEntry"
On Error GoTo Err

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

mPositionManagerGridMappingTable(index).IsExpanded = True
BracketOrderGrid.Row = mPositionManagerGridMappingTable(index).GridIndex
BracketOrderGrid.col = ExpandIndicator
BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture

Dim lSymbol As String
lSymbol = GridColumn(mPositionManagerGridMappingTable(index).GridIndex, Symbol)

Dim i As Long
i = mPositionManagerGridMappingTable(index).GridIndex + 1
Do While GridColumn(i, Symbol) = lSymbol And gridRowIsBracketOrder(i)
    BracketOrderGrid.rowHeight(i) = -1
    
    Dim lBracketOrderIndex As Long
    lBracketOrderIndex = getBracketOrderGridMappingIndexFromRowIndex(i)
    i = expandBracketOrderEntry(lBracketOrderIndex, True) + 1
Loop

BracketOrderGrid.Row = lPrevRow
BracketOrderGrid.col = lPrevCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findBracketOrderGridMappingIndex(ByVal pBracketOrder As IBracketOrder) As Long
 Const ProcName As String = "findBracketOrderGridMappingIndex"
On Error GoTo Err

Dim lBracketOrderIndex As Long
lBracketOrderIndex = pBracketOrder.ApplicationIndex
Do While lBracketOrderIndex > UBound(mBracketOrderGridMappingTable)
    ReDim Preserve mBracketOrderGridMappingTable(2 * (UBound(mBracketOrderGridMappingTable) + 1) - 1) As BracketOrderGridMappingEntry
Loop
If lBracketOrderIndex > mMaxBracketOrderGridMappingTableIndex Then mMaxBracketOrderGridMappingTableIndex = lBracketOrderIndex

findBracketOrderGridMappingIndex = lBracketOrderIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function findPositionManagerGridMappingIndex(ByVal pm As PositionManager) As Long
Const ProcName As String = "findPositionManagerGridMappingIndex"
On Error GoTo Err

Dim pmIndex As Long
pmIndex = pm.ApplicationIndex

Do While pmIndex > UBound(mPositionManagerGridMappingTable)
    ReDim Preserve mPositionManagerGridMappingTable(2 * (UBound(mPositionManagerGridMappingTable) + 1) - 1) As PositionManagerGridMappingEntry
Loop
If pmIndex > mMaxPositionManagerGridMappingTableIndex Then mMaxPositionManagerGridMappingTableIndex = pmIndex

findPositionManagerGridMappingIndex = pmIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function formattedTime(ByVal pTime As Date) As String
formattedTime = IIf(pTime = 0, _
                    "", _
                    IIf(Int(pTime) = Int(Now), _
                        FormatTimestamp(pTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs), _
                        FormatTimestamp(pTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)))
End Function

Private Function getBracketOrderGridIndex(ByVal pBracketOrder) As Long
getBracketOrderGridIndex = mBracketOrderGridMappingTable(findBracketOrderGridMappingIndex(pBracketOrder)).GridIndex
End Function

Private Function getBracketOrderGridMappingIndexFromRowIndex(ByVal pRowIndex As Long) As Long
getBracketOrderGridMappingIndexFromRowIndex = BracketOrderGrid.RowData(pRowIndex) - RowDataBracketOrderBase
End Function

Private Function getPositionManagerGridMappingIndexFromRowIndex(ByVal pRowIndex As Long) As Long
getPositionManagerGridMappingIndexFromRowIndex = BracketOrderGrid.RowData(pRowIndex) - RowDataPositionManagerBase
End Function

Private Function getSelectedOrder() As IOrder
Const ProcName As String = "getSelectedOrder"
On Error GoTo Err

Dim lIndex As Long
lIndex = getSelectedBracketOrderGridMappingIndex
If lIndex = NullIndex Then Exit Function

Dim lOrderOffset As Long
lOrderOffset = BracketOrderGrid.Row - mBracketOrderGridMappingTable(lIndex).GridIndex
            
Select Case lOrderOffset
Case mBracketOrderGridMappingTable(lIndex).EntryGridOffset
    Set getSelectedOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder.EntryOrder
Case mBracketOrderGridMappingTable(lIndex).StopLossGridOffset
    Set getSelectedOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder.StopLossOrder
Case mBracketOrderGridMappingTable(lIndex).TargetGridOffset
    Set getSelectedOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder.TargetOrder
Case mBracketOrderGridMappingTable(lIndex).CloseoutGridOffset
    Set getSelectedOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder.CloseoutOrder
Case Else
    Set getSelectedOrder = Nothing
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSelectedBracketOrder() As IBracketOrder
Dim lIndex As Long
lIndex = getSelectedBracketOrderGridMappingIndex
If lIndex = NullIndex Then Exit Function

Set getSelectedBracketOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder
End Function

Private Function getSelectedBracketOrderGridMappingIndex() As Long
getSelectedBracketOrderGridMappingIndex = NullIndex

If Not gridRowIsBracketOrder(BracketOrderGrid.Row) Then Exit Function

getSelectedBracketOrderGridMappingIndex = getBracketOrderGridMappingIndexFromRowIndex(BracketOrderGrid.Row)
End Function

Private Property Let GridColumn(ByVal pRowIndex As Long, ByVal pColumnIndex As Long, ByVal value As String)
BracketOrderGrid.TextMatrix(pRowIndex, pColumnIndex) = value
End Property

Private Property Get GridColumn(ByVal pRowIndex As Long, ByVal pColumnIndex As Long) As String
GridColumn = BracketOrderGrid.TextMatrix(pRowIndex, pColumnIndex)
End Property

Private Function gridRowIsBracketOrder(ByVal pRowIndex As Long) As Boolean
gridRowIsBracketOrder = (BracketOrderGrid.RowData(pRowIndex) >= RowDataBracketOrderBase And _
                            BracketOrderGrid.RowData(pRowIndex) < RowDataPositionManagerBase)
End Function

Private Function gridRowIsOrder(ByVal pRowIndex As Long) As Boolean
gridRowIsOrder = (BracketOrderGrid.RowData(pRowIndex) < RowDataBracketOrderBase)
End Function

Private Function gridRowIsPositionManager(ByVal pRowIndex As Long) As Boolean
gridRowIsPositionManager = (BracketOrderGrid.RowData(pRowIndex) >= RowDataPositionManagerBase)
End Function

Private Sub handleEditingTerminationKey(ByVal KeyCode As Long)
Const ProcName As String = "handleEditingTerminationKey"
On Error GoTo Err

Select Case KeyCode
Case KeyCodeConstants.vbKeyEscape   ' ESC: hide, return focus to MSHFlexGrid.
    endEdit
Case KeyCodeConstants.vbKeyReturn   ' ENTER return focus to MSHFlexGrid.
    updateBracketOrder
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub invertEntryColors(ByVal rowNumber As Long)
Const ProcName As String = "invertEntryColors"
On Error GoTo Err

If rowNumber < 0 Then Exit Sub

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

BracketOrderGrid.Row = rowNumber

Dim i As Long
For i = OtherColumns To BracketOrderGrid.Cols - 1
    BracketOrderGrid.col = i
    Dim lForeColor As Long
    lForeColor = IIf(BracketOrderGrid.CellForeColor = 0, BracketOrderGrid.ForeColor, BracketOrderGrid.CellForeColor)
    If lForeColor = SystemColorConstants.vbWindowText Then
        BracketOrderGrid.CellForeColor = SystemColorConstants.vbHighlightText
    ElseIf lForeColor = SystemColorConstants.vbHighlightText Then
        BracketOrderGrid.CellForeColor = SystemColorConstants.vbWindowText
    ElseIf lForeColor > 0 Then
        BracketOrderGrid.CellForeColor = IIf((lForeColor Xor &HFFFFFF) = 0, 1, lForeColor Xor &HFFFFFF)
    End If
    
    Dim lBackColor As Long
    lBackColor = IIf(BracketOrderGrid.CellBackColor = 0, BracketOrderGrid.BackColor, BracketOrderGrid.CellBackColor)
    If lBackColor = SystemColorConstants.vbWindowBackground Then
        BracketOrderGrid.CellBackColor = SystemColorConstants.vbHighlight
    ElseIf lBackColor = SystemColorConstants.vbHighlight Then
        BracketOrderGrid.CellBackColor = SystemColorConstants.vbWindowBackground
    ElseIf lBackColor > 0 Then
        BracketOrderGrid.CellBackColor = IIf((lBackColor Xor &HFFFFFF) = 0, 1, lBackColor Xor &HFFFFFF)
    End If
Next

BracketOrderGrid.Row = lPrevRow
BracketOrderGrid.col = lPrevCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupBracketOrderGrid()
Const ProcName As String = "setupBracketOrderGrid"
On Error GoTo Err

With BracketOrderGrid
    Set mSelectedBracketOrder = Nothing
    
    .Redraw = False
    .AllowUserResizing = flexResizeBoth
    
    .Cols = 0
    .Rows = 20
    .FixedRows = 3
    ' .FixedCols = 1
    
    setupBracketOrderGridColumn 0, ExpandIndicator, BracketOrderGridColumnWidths.ExpandIndicatorWidth, "", True, MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    setupBracketOrderGridColumn 0, Symbol, BracketOrderGridColumnWidths.SymbolWidth, "Symbol", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    
    setupBracketOrderGridColumn 0, PositionCurrencyCode, PositionCurrencyCodeWidth, "Curr", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 0, PositionDrawdown, PositionDrawdownWidth, "Drawdown", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 0, PositionExchange, PositionExchangeWidth, "Exchange", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 0, PositionMaxProfit, PositionMaxProfitWidth, "Max", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 0, PositionProfit, PositionProfitWidth, "Profit", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 0, PositionSize, PositionSizeWidth, "Size", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    
    setupBracketOrderGridColumn 1, BracketCreationTime, BracketCreationTimeWidth, "Creation Time", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketCurrencyCode, BracketCurrencyCodeWidth, "Curr", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 1, BracketDrawdown, BracketDrawdownWidth, "Drawdown", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketMaxProfit, BracketMaxProfitWidth, "Max", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketProfit, BracketProfitWidth, "Profit", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketSize, BracketSizeWidth, "Size", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    
    setupBracketOrderGridColumn 2, OrderAction, OrderActionWidth, "Action", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, OrderAuxPrice, OrderAuxPriceWidth, "Trigger", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderAveragePrice, OrderAveragePriceWidth, "Avg fill", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderId, OrderIdWidth, "Id", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, OrderLastFillPrice, OrderLastFillPriceWidth, "Last fill", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderLastFillTime, OrderLastFillTimeWidth, "Last fill time", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderType, OrderTypeWidth, "Type", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, OrderPrice, OrderPriceWidth, "Price", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderQuantity, OrderQuantityWidth, "Qty", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderQuantityRemaining, OrderQuantityRemainingWidth, "Rem", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderSize, OrderSizeWidth, "Size", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, OrderStatus, OrderStatusWidth, "Status", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, OrderMode, OrderModeWidth, "Mode", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, OrderBrokerId, OrderBrokerIdWidth, "Broker Id", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    
    .MergeCells = flexMergeFree
    .MergeCol(Symbol) = True
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightAlways
    .FocusRect = flexFocusNone
    
    .Redraw = True
    
    mFirstBracketOrderGridRowIndex = 3
End With

EditText.Text = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupBracketOrderGridColumn( _
                ByVal rowNumber As Long, _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As MSFlexGridLib.AlignmentSettings)
Const ProcName As String = "setupBracketOrderGridColumn"
On Error GoTo Err

Dim lColumnWidth As Long
Dim i As Long

With BracketOrderGrid
    .Row = rowNumber
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showPositionManagerEntry(ByVal pPositionManager As PositionManager)
Const ProcName As String = "showPositionManagerEntry"
On Error GoTo Err

Dim lIndex As Long
lIndex = findPositionManagerGridMappingIndex(pPositionManager)

If mPositionManagerGridMappingTable(lIndex).GridIndex <> 0 Then Exit Sub

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

With mPositionManagerGridMappingTable(lIndex)
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = gGetContractFromContractFuture(pPositionManager.ContractFuture).Specifier
    .GridIndex = addEntryToBracketOrderGrid(0, lContractSpec.LocalSymbol, True)
    BracketOrderGrid.RowData(.GridIndex) = lIndex + RowDataPositionManagerBase
    BracketOrderGrid.Row = .GridIndex
    BracketOrderGrid.col = 1
    BracketOrderGrid.ColSel = BracketOrderGrid.Cols - 1
    BracketOrderGrid.FillStyle = MSFlexGridLib.FillStyleSettings.flexFillRepeat
    BracketOrderGrid.CellBackColor = &HC0C0C0
    BracketOrderGrid.CellForeColor = vbWhite
    BracketOrderGrid.CellFontBold = True
    GridColumn(.GridIndex, PositionExchange) = lContractSpec.Exchange
    BracketOrderGrid.TextMatrix(.GridIndex, PositionCurrencyCode) = lContractSpec.CurrencyCode
    BracketOrderGrid.TextMatrix(.GridIndex, PositionSize) = pPositionManager.PositionSize
    BracketOrderGrid.col = ExpandIndicator
    BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture
    .IsExpanded = True
End With

BracketOrderGrid.Row = lPrevRow
BracketOrderGrid.col = lPrevCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateBracketOrder()
Const ProcName As String = "updateBracketOrder"
On Error GoTo Err

Dim lPrice As Double

If Not EditText.Visible Then Exit Sub

Dim lOrder As IOrder
Set lOrder = getSelectedOrder
If BracketOrderGrid.col = OrderPrice Then
    If ParsePrice(EditText.Text, mSelectedBracketOrder.Contract.Specifier.secType, mSelectedBracketOrder.Contract.TickSize, lPrice) Then
        lOrder.LimitPrice = lPrice
    End If
ElseIf BracketOrderGrid.col = OrderAuxPrice Then
    If ParsePrice(EditText.Text, mSelectedBracketOrder.Contract.Specifier.secType, mSelectedBracketOrder.Contract.TickSize, lPrice) Then
        lOrder.TriggerPrice = lPrice
    End If
ElseIf BracketOrderGrid.col = OrderQuantity Then
    If IsInteger(EditText.Text, 0) Then
        lOrder.Quantity = CLng(EditText.Text)
    End If
End If
    
If mSelectedBracketOrder.IsDirty Then mSelectedBracketOrder.Update

endEdit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

