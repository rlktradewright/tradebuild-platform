VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl OrdersSummary 
   Alignable       =   -1  'True
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12810
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   12810
   Begin VB.TextBox MessageText 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Text            =   "Message Text"
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox EditText 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   11160
      TabIndex        =   0
      Text            =   "EditText"
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

Implements IChangeListener
Implements ICollectionChangeListener
Implements IBracketOrderErrorListener
Implements IBracketOrderMsgListener
Implements IProfitListener
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Click()
Event SelectionChanged()
                
'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "OrdersSummary"

Private Const RowDataOrderRoleMask          As Long = &HF&

Private Const RowDataBracketOrderBase       As Long = &H10&
Private Const RowDataBracketOrderMask       As Long = &HFFFF0

Private Const RowDataPositionManagerBase    As Long = &H100000
Private Const RowDataPositionManagerMask    As Long = &H7FF00000

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
    
    LastColumn = OrderBrokerId
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
    
    ' indicates whether this entry's messages are visible
    ' (for example it may be expanded but not visible because
    ' the PositionManager entry may not be expanded)
    IsMessagesVisible   As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    GridIndex           As Long
                                
    ' offset from GridIndex of line in OrdersGrid relating to
    ' the corresponding order: -1 means  it's not in the grid
    EntryGridOffset     As Long
    EntryMessageIndex   As Long
    
    StopLossGridOffset  As Long
    StopLossMessageIndex   As Long
    
    TargetGridOffset    As Long
    TargetMessageIndex  As Long
    
    CloseoutGridOffset  As Long
    CloseoutMessageIndex   As Long
    
    TickSize            As Double
    SecType             As SecurityTypes
    
End Type

Private Type PositionManagerGridMappingEntry
    
    ' indicates whether this entry in the grid is expanded
    IsExpanded          As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    GridIndex           As Long
                                
End Type

Private Type MessageMappingEntry
    BracketOrderIndex   As Long
    OrderRole           As BracketOrderRoles
    MessageTextIndex    As Long
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManager                                  As IMarketDataManager

Private mSelectedBracketOrder                               As IBracketOrder
Private mSelectedOrder                                      As IOrder

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
Private mEditedOrderRole                                    As BracketOrderRoles

Private mTheme                                              As ITheme

Private mMessageMappingTable()                              As MessageMappingEntry
Private mMaxMessageMappingTableIndex                        As Long

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

ReDim mMessageMappingTable(3) As MessageMappingEntry
mMaxMessageMappingTableIndex = 0

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

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
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
            If lBracketOrder.Size = 0 Then
                ' Don't stop listening - in case broker sends late messages
'                lBracketOrder.RemoveChangeListener Me
'                lBracketOrder.RemoveBracketOrderErrorListener Me
'                lBracketOrder.RemoveBracketOrderMessageListener Me
            End If
        Case BracketOrderChangeTypes.BracketOrderSelfCancelled
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            If lBracketOrder.Size = 0 Then
                lBracketOrder.RemoveChangeListener Me
                lBracketOrder.RemoveBracketOrderErrorListener Me
                lBracketOrder.RemoveBracketOrderMessageListener Me
            End If
        Case BracketOrderChangeTypes.BracketOrderEntryOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .EntryGridOffset, lBracketOrder.EntryOrder, .SecType, .TickSize
            If lBracketOrder.EntryOrder.ErrorMessage = "" And _
                lBracketOrder.EntryOrder.Message = "" Then clearMessage lBracketOrder.EntryOrder, lBracketOrderIndex
        Case BracketOrderChangeTypes.BracketOrderStopLossOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .StopLossGridOffset, lBracketOrder.StopLossOrder, .SecType, .TickSize
            If lBracketOrder.StopLossOrder.ErrorMessage = "" And _
                lBracketOrder.StopLossOrder.Message = "" Then clearMessage lBracketOrder.StopLossOrder, lBracketOrderIndex
        Case BracketOrderChangeTypes.BracketOrderTargetOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .TargetGridOffset, lBracketOrder.TargetOrder, .SecType, .TickSize
            If lBracketOrder.TargetOrder.ErrorMessage = "" And _
                lBracketOrder.TargetOrder.Message = "" Then clearMessage lBracketOrder.TargetOrder, lBracketOrderIndex
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
                                    BracketOrderRoleCloseout, _
                                    .SecType, _
                                    .TickSize
        Case BracketOrderChangeTypes.BracketOrderCloseoutOrderChanged
            If lBracketOrder Is mSelectedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .CloseoutGridOffset, lBracketOrder.CloseoutOrder, .SecType, .TickSize
            If lBracketOrder.CloseoutOrder.ErrorMessage = "" And _
                lBracketOrder.CloseoutOrder.Message = "" Then clearMessage lBracketOrder.CloseoutOrder, lBracketOrderIndex
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
' ICollectionChangeListener Interface Members
'@================================================================================

Private Sub ICollectionChangeListener_Change(ev As CollectionChangeEventData)
0 Const ProcName As String = "ICollectionChangeListener_Change"
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
        lBracketOrder.RemoveBracketOrderErrorListener Me
        lBracketOrder.RemoveBracketOrderMessageListener Me
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
' IBracketOrderErrorListener Interface Members
'@================================================================================

Private Sub IBracketOrderErrorListener_NotifyBracketOrderError(ev As BracketOrderErrorEventData)
Const ProcName As String = "IBracketOrderErrorListener_NotifyBracketOrderError"
On Error GoTo Err

Dim lBracketOrderIndex As Long
lBracketOrderIndex = findBracketOrderGridMappingIndex(ev.Source)

clearMessage ev.AffectedOrder, lBracketOrderIndex
setupMessage ev.AffectedOrder, ev.AffectedOrder.ErrorMessage, lBracketOrderIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IBracketOrderMessageListener Interface Members
'@================================================================================

Private Sub IBracketOrderMsgListener_NotifyBracketOrderMessage(ev As BracketOrderMessageEventData)
Const ProcName As String = "IBracketOrderMsgListener_NotifyBracketOrderMessage"
On Error GoTo Err

Dim lBracketOrderIndex As Long
lBracketOrderIndex = findBracketOrderGridMappingIndex(ev.Source)

clearMessage ev.AffectedOrder, lBracketOrderIndex
setupMessage ev.AffectedOrder, ev.AffectedOrder.Message, lBracketOrderIndex

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
    
    If CBool(lBOProfitType And ProfitTypes.ProfitTypeProfit) Then _
        displayProfitValue lProfitCalculator.Profit, rowIndex, BracketProfit
    If CBool(lBOProfitType And ProfitTypes.ProfitTypeMaxProfit) Then _
        displayProfitValue lProfitCalculator.MaxProfit, rowIndex, BracketMaxProfit
    If CBool(lBOProfitType And ProfitTypes.ProfitTypeDrawdown) Then _
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
    
    If CBool(lPMProfitType And ProfitTypes.ProfitTypeSessionProfit) Then _
        displayProfitValue lPositionManager.Profit, rowIndex, PositionProfit
    If CBool(lPMProfitType And ProfitTypes.ProfitTypeSessionMaxProfit) Then _
        displayProfitValue lPositionManager.MaxProfit, rowIndex, PositionMaxProfit
    If CBool(lPMProfitType And ProfitTypes.ProfitTypeSessionDrawdown) Then _
        displayProfitValue -lPositionManager.Drawdown, rowIndex, PositionDrawdown
End If

adjustEditBox

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

If mIsEditing Then Exit Sub

If BracketOrderGrid.MouseCol = Symbol Then
    RaiseEvent Click
    Exit Sub
End If

If BracketOrderGrid.MouseCol = ExpandIndicator Then
    expandOrContract
    RaiseEvent Click
    Exit Sub
End If

BracketOrderGrid.col = BracketOrderGrid.MouseCol

Dim lBracketOrder As IBracketOrder
Set lBracketOrder = getSelectedBracketOrder

Dim lSelectionChanged As Boolean
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

Set mSelectedOrder = getSelectedOrder
    
If mSelectedOrder Is Nothing Then
    RaiseEvent Click
    RaiseEvent SelectionChanged
    Exit Sub
End If

If Not mSelectedOrder.IsModifiable Then
    RaiseEvent Click
    RaiseEvent SelectionChanged
    Exit Sub
End If

If (BracketOrderGrid.col = OrderPrice And _
        mSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttLimitPrice)) Or _
    (BracketOrderGrid.col = OrderAuxPrice And _
        mSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttTriggerPrice)) Or _
    (BracketOrderGrid.col = OrderQuantity And _
        mSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttQuantity)) _
Then
    mIsEditing = True
    mEditedCol = BracketOrderGrid.col
    mEditedOrderRole = SelectedOrderRole
    
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

Dim i As Long
For i = 1 To mMaxMessageMappingTableIndex
    displayMessage i
Next

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
                ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = Value
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

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Get SelectedItem() As IBracketOrder
Set SelectedItem = mSelectedBracketOrder
End Property

Public Property Get SelectedOrderRole() As BracketOrderRoles
SelectedOrderRole = BracketOrderGrid.RowData(BracketOrderGrid.Row) And RowDataOrderRoleMask
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

BracketOrderGrid.BackColorBkg = mTheme.TextBackColor
BracketOrderGrid.BackColor = mTheme.TextBackColor
BracketOrderGrid.BackColorFixed = mTheme.GridBackColorFixed
BracketOrderGrid.ForeColor = mTheme.GridForeColor
BracketOrderGrid.ForeColorFixed = mTheme.GridForeColorFixed
BracketOrderGrid.GridColor = mTheme.TextBackColor
BracketOrderGrid.GridColorFixed = mTheme.GridLineColorFixed
If Not mTheme.GridFont Is Nothing Then Set BracketOrderGrid.Font = mTheme.GridFont

If Not mTheme.TextFont Is Nothing Then Set EditText.Font = mTheme.TextFont

Dim i As Long
For i = 1 To mMaxMessageMappingTableIndex
    If mMessageMappingTable(i).MessageTextIndex <> 0 Then
        MessageText(mMessageMappingTable(i).MessageTextIndex).BackColor = mTheme.TextBackColor
        MessageText(mMessageMappingTable(i).MessageTextIndex).ForeColor = mTheme.AlertForeColor
        If Not mTheme.AlertFont Is Nothing Then Set MessageText(mMessageMappingTable(i).MessageTextIndex).Font = mTheme.AlertFont
    End If
Next

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

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mMarketDataManager = Nothing

Dim i As Long
For i = 1 To mMaxBracketOrderGridMappingTableIndex
    If Not mBracketOrderGridMappingTable(i).BracketOrder Is Nothing Then
        mBracketOrderGridMappingTable(i).BracketOrder.RemoveChangeListener Me
        mBracketOrderGridMappingTable(i).BracketOrder.RemoveBracketOrderErrorListener Me
        mBracketOrderGridMappingTable(i).BracketOrder.RemoveBracketOrderMessageListener Me
        mBracketOrderGridMappingTable(i).ProfitCalculator.RemoveProfitListener Me
        Set mBracketOrderGridMappingTable(i).BracketOrder = Nothing
        Set mBracketOrderGridMappingTable(i).ProfitCalculator = Nothing
        mBracketOrderGridMappingTable(i).CloseoutGridOffset = 0
        mBracketOrderGridMappingTable(i).EntryGridOffset = 0
        mBracketOrderGridMappingTable(i).GridIndex = 0
        mBracketOrderGridMappingTable(i).IsExpanded = False
        mBracketOrderGridMappingTable(i).SecType = SecTypeNone
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
Const ProcName As String = "MonitorPositions"
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
            collapseBracketOrderEntry findBracketOrderGridMappingIndex(lBracketOrder)
        Else
            lAnyActiveBracketOrders = True
        End If
    Next
    
    If Not lAnyActiveBracketOrders Then collapsePositionManagerEntry findPositionManagerGridMappingIndex(lPositionManager)
    
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
pBracketOrder.AddBracketOrderErrorListener Me
pBracketOrder.AddBracketOrderMessageListener Me
displayBracketOrder pBracketOrder, pPositionManager

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function addBracketOrderEntryToBracketOrderGrid( _
                ByVal pSymbol As String, _
                ByVal pPositionManagerGridMappingIndex As Long, _
                ByVal pRowData As Long) As Long
Const ProcName As String = "addBracketOrderEntryToBracketOrderGrid"
On Error GoTo Err

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

Dim Index As Long
Index = addEntryToBracketOrderGrid(pPositionManagerGridMappingIndex + 1, pSymbol, True)

BracketOrderGrid.RowData(Index) = pRowData

BracketOrderGrid.Row = Index

setExpandOrCollapseIcon Index, False

BracketOrderGrid.col = BracketProfit
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

BracketOrderGrid.col = BracketMaxProfit
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

BracketOrderGrid.col = BracketDrawdown
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

addBracketOrderEntryToBracketOrderGrid = Index

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
        BracketOrderGrid.AddItem ""
        pIndex = BracketOrderGrid.Rows - 1
    ElseIf GridColumn(pIndex, Symbol) = "" Then
        GridColumn(pIndex, Symbol) = pSymbol
    Else
        BracketOrderGrid.AddItem "", pIndex
    End If
Else
    BracketOrderGrid.AddItem "", pIndex
End If

GridColumn(pIndex, Symbol) = pSymbol
If pIndex < BracketOrderGrid.Rows - 1 Then
    ' this new entry has displaced one or more existing entries so
    ' the BracketOrderGridMappingTable and PositionManageGridMappingTable indexes
    ' need to be adjusted
    For i = 1 To mMaxBracketOrderGridMappingTableIndex
        If mBracketOrderGridMappingTable(i).GridIndex >= pIndex Then
            mBracketOrderGridMappingTable(i).GridIndex = mBracketOrderGridMappingTable(i).GridIndex + 1
        End If
    Next
    For i = 1 To mMaxPositionManagerGridMappingTableIndex
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
                ByVal pRowData As Long, _
                ByVal pOrderRole As BracketOrderRoles, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double)
Const ProcName As String = "addOrderEntryToBracketOrderGrid"
On Error GoTo Err

pIndex = addEntryToBracketOrderGrid(0, pSymbol, False, pIndex)

BracketOrderGrid.RowData(pIndex) = pRowData

GridColumn(pIndex, OrderMode) = BracketOrderRoleToString(pOrderRole)

displayOrderValues pIndex, pOrder, pSecType, pTickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub adjustEditBox()
Const ProcName As String = "adjustEditBox"
On Error GoTo Err

If Not mIsEditing Then Exit Sub

If (mEditedCol = OrderPrice And _
        Not mSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttLimitPrice)) Or _
    (mEditedCol = OrderAuxPrice And _
        Not mSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttTriggerPrice)) Or _
    (mEditedCol = OrderQuantity And _
        Not mSelectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttQuantity)) _
Then
    endEdit
    Exit Sub
End If

Dim lIndex As Long
lIndex = findBracketOrderGridMappingIndex(mSelectedBracketOrder)

Dim lOffset As Long
Select Case mEditedOrderRole
Case BracketOrderRoleEntry
    lOffset = mBracketOrderGridMappingTable(lIndex).EntryGridOffset
Case BracketOrderRoleStopLoss
    lOffset = mBracketOrderGridMappingTable(lIndex).StopLossGridOffset
Case BracketOrderRoleTarget
    lOffset = mBracketOrderGridMappingTable(lIndex).TargetGridOffset
Case BracketOrderRoleCloseout
    lOffset = mBracketOrderGridMappingTable(lIndex).CloseoutGridOffset
End Select

BracketOrderGrid.Row = mBracketOrderGridMappingTable(lIndex).GridIndex + lOffset
BracketOrderGrid.col = mEditedCol

EditText.Move BracketOrderGrid.Left + BracketOrderGrid.CellLeft + 8, _
            BracketOrderGrid.Top + BracketOrderGrid.Celltop + 8, _
            BracketOrderGrid.CellWidth - 16, _
            BracketOrderGrid.CellHeight - 16

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function allocateMessageIndex( _
                ByVal pBracketOrderIndex As Long, _
                ByVal pOrderRole As BracketOrderRoles, _
                ByVal pMessage As String, _
                ByVal pSymbol As String, _
                ByVal pIsExpanded As Boolean) As Long
Const ProcName As String = "allocateMessageIndex"
On Error GoTo Err

mMaxMessageMappingTableIndex = mMaxMessageMappingTableIndex + 1
If mMaxMessageMappingTableIndex > UBound(mMessageMappingTable) Then ReDim Preserve mMessageMappingTable(2 * (UBound(mMessageMappingTable) + 1) - 1) As MessageMappingEntry

Load MessageText(MessageText.UBound + 1)
With mMessageMappingTable(mMaxMessageMappingTableIndex)
    .MessageTextIndex = MessageText.UBound
    .BracketOrderIndex = pBracketOrderIndex
    .OrderRole = pOrderRole
End With

MessageText(MessageText.UBound).Text = pMessage
If Not mTheme Is Nothing Then
    MessageText(MessageText.UBound).BackColor = mTheme.TextBackColor
    MessageText(MessageText.UBound).ForeColor = mTheme.AlertForeColor
    If Not mTheme.AlertFont Is Nothing Then Set MessageText(MessageText.UBound).Font = mTheme.AlertFont
End If

Dim lMessageRow As Long
With mBracketOrderGridMappingTable(pBracketOrderIndex)
    Select Case pOrderRole
    Case BracketOrderRoleEntry
        lMessageRow = .GridIndex + .EntryGridOffset + 1
    Case BracketOrderRoleStopLoss
        lMessageRow = .GridIndex + .StopLossGridOffset + 1
    Case BracketOrderRoleTarget
        lMessageRow = .GridIndex + .TargetGridOffset + 1
    Case BracketOrderRoleCloseout
        lMessageRow = .GridIndex + .CloseoutGridOffset + 1
    End Select
End With

addEntryToBracketOrderGrid 0, pSymbol, False, lMessageRow
If Not pIsExpanded Then BracketOrderGrid.RowHeight(lMessageRow) = 0

allocateMessageIndex = mMaxMessageMappingTableIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub clearMessage(ByVal pOrder As IOrder, ByVal pBracketOrderIndex As Long)
Const ProcName As String = "clearMessage"
On Error GoTo Err

With mBracketOrderGridMappingTable(pBracketOrderIndex)
    Dim lRole As BracketOrderRoles
    lRole = getOrderRole(pBracketOrderIndex, pOrder)
    Select Case lRole
    Case BracketOrderRoleEntry
        If .EntryMessageIndex <> 0 Then
            If .StopLossGridOffset <> NullIndex Then .StopLossGridOffset = .StopLossGridOffset - 1
            If .TargetGridOffset <> NullIndex Then .TargetGridOffset = .TargetGridOffset - 1
            If .CloseoutGridOffset <> NullIndex Then .CloseoutGridOffset = .CloseoutGridOffset - 1
            deleteMessage .EntryMessageIndex
            .EntryMessageIndex = 0
        End If
    Case BracketOrderRoleStopLoss
        If .StopLossMessageIndex <> 0 Then
            If .TargetGridOffset <> NullIndex Then .TargetGridOffset = .TargetGridOffset - 1
            If .CloseoutGridOffset <> NullIndex Then .CloseoutGridOffset = .CloseoutGridOffset - 1
            deleteMessage .StopLossMessageIndex
            .StopLossMessageIndex = 0
        End If
    Case BracketOrderRoleTarget
        If .TargetMessageIndex <> 0 Then
            If .CloseoutGridOffset <> NullIndex Then .CloseoutGridOffset = .CloseoutGridOffset - 1
            deleteMessage .TargetMessageIndex
            .TargetMessageIndex = 0
        End If
    Case BracketOrderRoleCloseout
        If .CloseoutMessageIndex <> 0 Then
            deleteMessage .CloseoutMessageIndex
            .CloseoutMessageIndex = 0
        End If
    End Select
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function collapseBracketOrderEntry( _
                ByVal Index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Const ProcName As String = "collapseBracketOrderEntry"
On Error GoTo Err

Dim lIndex As Long
Dim lLastIndex As Long

With mBracketOrderGridMappingTable(Index)

    .IsMessagesVisible = False
    
    If mIsEditing And .BracketOrder Is mSelectedBracketOrder Then endEdit
    
    If .EntryGridOffset > 0 Then
        lIndex = .GridIndex + .EntryGridOffset
        BracketOrderGrid.RowHeight(lIndex) = 0
        If .EntryMessageIndex <> 0 Then
            lIndex = lIndex + 1
            hideMessage .EntryMessageIndex, lIndex
        End If
        lLastIndex = lIndex
    End If
    If .StopLossGridOffset > 0 Then
        lIndex = .GridIndex + .StopLossGridOffset
        BracketOrderGrid.RowHeight(lIndex) = 0
        If .StopLossMessageIndex <> 0 Then
            lIndex = lIndex + 1
            hideMessage .StopLossMessageIndex, lIndex
        End If
        If lIndex > lLastIndex Then lLastIndex = lIndex
    End If
    If .TargetGridOffset > 0 Then
        lIndex = .GridIndex + .TargetGridOffset
        BracketOrderGrid.RowHeight(lIndex) = 0
        If .TargetMessageIndex <> 0 Then
            lIndex = lIndex + 1
            hideMessage .TargetMessageIndex, lIndex
        End If
        If lIndex > lLastIndex Then lLastIndex = lIndex
    End If
    If .CloseoutGridOffset > 0 Then
        lIndex = .GridIndex + .CloseoutGridOffset
        BracketOrderGrid.RowHeight(lIndex) = 0
        If .CloseoutMessageIndex <> 0 Then
            lIndex = lIndex + 1
            hideMessage .CloseoutMessageIndex, lIndex
        End If
        If lIndex > lLastIndex Then lLastIndex = lIndex
    End If
    
    If Not preserveCurrentExpandedState Then
        .IsExpanded = False
        setExpandOrCollapseIcon .GridIndex, True
    End If
End With

collapseBracketOrderEntry = lLastIndex + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub collapsePositionManagerEntry(ByVal Index As Long)
Const ProcName As String = "collapsePositionManagerEntry"
On Error GoTo Err

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

Dim lPmRow As Long
lPmRow = mPositionManagerGridMappingTable(Index).GridIndex

mPositionManagerGridMappingTable(Index).IsExpanded = False
BracketOrderGrid.Row = lPmRow

setExpandOrCollapseIcon lPmRow, True

Dim lSymbol As String
lSymbol = GridColumn(lPmRow, Symbol)

Dim i As Long
i = lPmRow + 1
Do While GridColumn(i, Symbol) = lSymbol And gridRowIsBracketOrder(i)
    BracketOrderGrid.RowHeight(i) = 0
    i = collapseBracketOrderEntry(getBracketOrderGridMappingIndexFromRowIndex(i), True)
Loop

BracketOrderGrid.Row = lPrevRow
BracketOrderGrid.col = lPrevCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteMessage(ByVal pMessageIndex As Long)
Const ProcName As String = "deleteMessage"
On Error GoTo Err

If Not mBracketOrderGridMappingTable(mMessageMappingTable(pMessageIndex).BracketOrderIndex).IsMessagesVisible Then Exit Sub

Dim lMessageRow As Long
With mBracketOrderGridMappingTable(mMessageMappingTable(pMessageIndex).BracketOrderIndex)
    Select Case mMessageMappingTable(pMessageIndex).OrderRole
    Case BracketOrderRoles.BracketOrderRoleEntry
        lMessageRow = .GridIndex + .EntryGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleStopLoss
        lMessageRow = .GridIndex + .StopLossGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleTarget
        lMessageRow = .GridIndex + .TargetGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleCloseout
        lMessageRow = .GridIndex + .CloseoutGridOffset + 1
    End Select
End With

BracketOrderGrid.RemoveItem lMessageRow
MessageText(mMessageMappingTable(pMessageIndex).MessageTextIndex).Visible = False
mMessageMappingTable(pMessageIndex).BracketOrderIndex = 0
mMessageMappingTable(pMessageIndex).MessageTextIndex = 0
mMessageMappingTable(pMessageIndex).OrderRole = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displayBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pPositionManager As PositionManager)
Const ProcName As String = "displayBracketOrder"
On Error GoTo Err

Dim lSymbol As String
lSymbol = pBracketOrder.Contract.Specifier.LocalSymbol

Dim lIndex As Long
lIndex = findBracketOrderGridMappingIndex(pBracketOrder)

If Not mBracketOrderGridMappingTable(lIndex).BracketOrder Is Nothing Then Exit Sub

With mBracketOrderGridMappingTable(lIndex)
    .IsExpanded = True
    .IsMessagesVisible = True
    .EntryGridOffset = NullIndex
    .StopLossGridOffset = NullIndex
    .TargetGridOffset = NullIndex
    .CloseoutGridOffset = NullIndex
    
    Set .BracketOrder = pBracketOrder
    .TickSize = pBracketOrder.Contract.TickSize
    .SecType = pBracketOrder.Contract.Specifier.SecType
    
    Dim lPositionManagerGridMappingIndex As Long
    lPositionManagerGridMappingIndex = mPositionManagerGridMappingTable(findPositionManagerGridMappingIndex(pPositionManager)).GridIndex
    .GridIndex = addBracketOrderEntryToBracketOrderGrid( _
                            pBracketOrder.Contract.Specifier.LocalSymbol, _
                            lPositionManagerGridMappingIndex, _
                            generateRowData(lPositionManagerGridMappingIndex, lIndex, BracketOrderRoleNone))
    GridColumn(.GridIndex, BracketCreationTime) = formattedTime(pBracketOrder.CreationTime)
    GridColumn(.GridIndex, BracketCurrencyCode) = pBracketOrder.Contract.Specifier.CurrencyCode
    
    Dim lDataSource As IMarketDataSource
    Set lDataSource = pPositionManager.DataSource
    Set .ProfitCalculator = CreateBracketProfitCalculator(pBracketOrder, lDataSource)
    .ProfitCalculator.AddProfitListener Me
    
    If Not pBracketOrder.EntryOrder Is Nothing Then
        .EntryGridOffset = 1
        addOrderEntryToBracketOrderGrid .GridIndex + .EntryGridOffset, _
                                lSymbol, _
                                pBracketOrder.EntryOrder, _
                                generateRowData(lPositionManagerGridMappingIndex, lIndex, BracketOrderRoles.BracketOrderRoleEntry), _
                                BracketOrderRoles.BracketOrderRoleEntry, _
                                .SecType, _
                                .TickSize
    End If
    
    If Not pBracketOrder.StopLossOrder Is Nothing Then
        If .EntryGridOffset >= 0 Then
            .StopLossGridOffset = .EntryGridOffset + 1
        Else
            .StopLossGridOffset = 1
        End If
        addOrderEntryToBracketOrderGrid .GridIndex + .StopLossGridOffset, _
                                lSymbol, _
                                pBracketOrder.StopLossOrder, _
                                generateRowData(lPositionManagerGridMappingIndex, lIndex, BracketOrderRoles.BracketOrderRoleStopLoss), _
                                BracketOrderRoles.BracketOrderRoleStopLoss, _
                                .SecType, _
                                .TickSize
    End If
    
    If Not pBracketOrder.TargetOrder Is Nothing Then
        If .StopLossGridOffset >= 0 Then
            .TargetGridOffset = .StopLossGridOffset + 1
        ElseIf .EntryGridOffset >= 0 Then
            .TargetGridOffset = .EntryGridOffset + 1
        Else
            .TargetGridOffset = 1
        End If
        addOrderEntryToBracketOrderGrid .GridIndex + .TargetGridOffset, _
                                lSymbol, _
                                pBracketOrder.TargetOrder, _
                                generateRowData(lPositionManagerGridMappingIndex, lIndex, BracketOrderRoles.BracketOrderRoleTarget), _
                                BracketOrderRoles.BracketOrderRoleTarget, _
                                .SecType, _
                                .TickSize
    End If

    If Not pBracketOrder.CloseoutOrder Is Nothing Then
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
                                lSymbol, _
                                pBracketOrder.CloseoutOrder, _
                                generateRowData(lPositionManagerGridMappingIndex, lIndex, BracketOrderRoles.BracketOrderRoleCloseout), _
                                BracketOrderRoles.BracketOrderRoleCloseout, _
                                .SecType, _
                                .TickSize
    End If

    If Not pBracketOrder.EntryOrder Is Nothing Then _
                    setupMessage pBracketOrder.EntryOrder, _
                                getMessage(pBracketOrder.EntryOrder), _
                                lIndex
    If Not pBracketOrder.StopLossOrder Is Nothing Then _
                    setupMessage pBracketOrder.StopLossOrder, _
                                getMessage(pBracketOrder.StopLossOrder), _
                                lIndex
    If Not pBracketOrder.TargetOrder Is Nothing Then _
                    setupMessage pBracketOrder.TargetOrder, _
                                getMessage(pBracketOrder.TargetOrder), _
                                lIndex

    If Not pBracketOrder.CloseoutOrder Is Nothing Then _
                    setupMessage pBracketOrder.CloseoutOrder, _
                                getMessage(pBracketOrder.CloseoutOrder), _
                                lIndex

End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displayMessage(ByVal pMessageIndex As Long)
Const ProcName As String = "displayMessage"
On Error GoTo Err

If mMessageMappingTable(pMessageIndex).MessageTextIndex = 0 Then Exit Sub
If Not mBracketOrderGridMappingTable(mMessageMappingTable(pMessageIndex).BracketOrderIndex).IsMessagesVisible Then Exit Sub

Dim lMessageRow As Long
With mBracketOrderGridMappingTable(mMessageMappingTable(pMessageIndex).BracketOrderIndex)
    Select Case mMessageMappingTable(pMessageIndex).OrderRole
    Case BracketOrderRoles.BracketOrderRoleEntry
        lMessageRow = .GridIndex + .EntryGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleStopLoss
        lMessageRow = .GridIndex + .StopLossGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleTarget
        lMessageRow = .GridIndex + .TargetGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleCloseout
        lMessageRow = .GridIndex + .CloseoutGridOffset + 1
    End Select
End With
    
If BracketOrderGrid.RowIsVisible(lMessageRow) Then
    Dim lCurrRow As Long
    lCurrRow = BracketOrderGrid.Row
    Dim lCurrCol As Long
    lCurrCol = BracketOrderGrid.col
    
    BracketOrderGrid.col = BracketOrderGridColumns.OrderAction
    BracketOrderGrid.Row = lMessageRow
    MessageText(mMessageMappingTable(pMessageIndex).MessageTextIndex).Move BracketOrderGrid.ColPos(BracketOrderGridColumns.OrderAction), _
                                                                    BracketOrderGrid.Celltop, _
                                                                    BracketOrderGrid.Width - BracketOrderGrid.ColPos(BracketOrderGridColumns.OrderAction), _
                                                                    BracketOrderGrid.RowHeight(lMessageRow)
    MessageText(mMessageMappingTable(pMessageIndex).MessageTextIndex).Visible = True
    MessageText(mMessageMappingTable(pMessageIndex).MessageTextIndex).ZOrder 0
    
    BracketOrderGrid.col = lCurrCol
    BracketOrderGrid.Row = lCurrRow
Else
    MessageText(mMessageMappingTable(pMessageIndex).MessageTextIndex).Visible = False
End If

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
mEditedOrderRole = BracketOrderRoleNone
BracketOrderGrid.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Private Sub expandOrContract()
Const ProcName As String = "expandOrContract"
On Error GoTo Err

Dim Index As Long
Dim expanded As Boolean

If gridRowIsPositionManager(BracketOrderGrid.Row) Then
    Index = getPositionManagerGridMappingIndexFromRowIndex(BracketOrderGrid.Row)
    expanded = mPositionManagerGridMappingTable(Index).IsExpanded
    If expanded Then
        collapsePositionManagerEntry Index
    Else
        expandPositionManagerEntry Index
    End If
ElseIf gridRowIsBracketOrder(BracketOrderGrid.Row) Then
    Index = getBracketOrderGridMappingIndexFromRowIndex(BracketOrderGrid.Row)
    expanded = mBracketOrderGridMappingTable(Index).IsExpanded
    If BracketOrderGrid.Row <> mBracketOrderGridMappingTable(Index).GridIndex Then
        ' clicked on an order entry
        Exit Sub
    End If
    If expanded Then
        collapseBracketOrderEntry Index
    Else
        expandBracketOrderEntry Index
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function expandBracketOrderEntry( _
                ByVal Index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Const ProcName As String = "expandBracketOrderEntry"
On Error GoTo Err

Dim lLastIndex As Long
Dim lIndex As Long

With mBracketOrderGridMappingTable(Index)
    If preserveCurrentExpandedState Then
        .IsMessagesVisible = .IsExpanded
    Else
        .IsMessagesVisible = True
    End If
    
    If .EntryGridOffset > 0 Then
        lIndex = .GridIndex + .EntryGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then
            BracketOrderGrid.RowHeight(lIndex) = -1
            If .EntryMessageIndex <> 0 Then
                lIndex = lIndex + 1
                displayMessage .EntryMessageIndex
            End If
        End If
        lLastIndex = lIndex
    End If
    If .StopLossGridOffset > 0 Then
        lIndex = .GridIndex + .StopLossGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then
            BracketOrderGrid.RowHeight(lIndex) = -1
            If .StopLossMessageIndex <> 0 Then
                lIndex = lIndex + 1
                displayMessage .StopLossMessageIndex
            End If
        End If
        If lIndex > lLastIndex Then lLastIndex = lIndex
    End If
    If .TargetGridOffset > 0 Then
        lIndex = .GridIndex + .TargetGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then
            BracketOrderGrid.RowHeight(lIndex) = -1
            If .TargetMessageIndex <> 0 Then
                lIndex = lIndex + 1
                displayMessage .TargetMessageIndex
            End If
        End If
        If lIndex > lLastIndex Then lLastIndex = lIndex
    End If
    If .CloseoutGridOffset > 0 Then
        lIndex = .GridIndex + .CloseoutGridOffset
        If Not preserveCurrentExpandedState Or .IsExpanded Then
            BracketOrderGrid.RowHeight(lIndex) = -1
            If .CloseoutMessageIndex <> 0 Then
                lIndex = lIndex + 1
                displayMessage .EntryMessageIndex
            End If
        End If
        If lIndex > lLastIndex Then lLastIndex = lIndex
    End If
    
    If Not preserveCurrentExpandedState Then
        .IsExpanded = True
        setExpandOrCollapseIcon .GridIndex, False
    End If
End With

expandBracketOrderEntry = lLastIndex + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub expandPositionManagerEntry(ByVal Index As Long)
Const ProcName As String = "expandPositionManagerEntry"
On Error GoTo Err

mPositionManagerGridMappingTable(Index).IsExpanded = True

setExpandOrCollapseIcon mPositionManagerGridMappingTable(Index).GridIndex, False

Dim lSymbol As String
lSymbol = GridColumn(mPositionManagerGridMappingTable(Index).GridIndex, Symbol)

Dim i As Long
i = mPositionManagerGridMappingTable(Index).GridIndex + 1
Do While GridColumn(i, Symbol) = lSymbol And gridRowIsBracketOrder(i)
    BracketOrderGrid.RowHeight(i) = -1
    i = expandBracketOrderEntry(getBracketOrderGridMappingIndexFromRowIndex(i), True)
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findBracketOrderGridMappingIndex(ByVal pBracketOrder As IBracketOrder) As Long
Const ProcName As String = "findBracketOrderGridMappingIndex"
On Error GoTo Err

Dim lBracketOrderIndex As Long
lBracketOrderIndex = pBracketOrder.ApplicationIndex + 1
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
pmIndex = pm.ApplicationIndex + 1

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

Private Function generateRowData( _
                ByVal pPositionManagerGridMappingIndex As Long, _
                ByVal pBracketOrderGridMappingIndex As Long, _
                ByVal pOrderRole As BracketOrderRoles) As Long
generateRowData = pPositionManagerGridMappingIndex * RowDataPositionManagerBase + _
                pBracketOrderGridMappingIndex * RowDataBracketOrderBase + _
                pOrderRole
End Function

Private Function getBracketOrderGridIndex(ByVal pBracketOrder) As Long
getBracketOrderGridIndex = mBracketOrderGridMappingTable(findBracketOrderGridMappingIndex(pBracketOrder)).GridIndex
End Function

Private Function getBracketOrderGridMappingIndexFromRowIndex(ByVal pRowIndex As Long) As Long
getBracketOrderGridMappingIndexFromRowIndex = (BracketOrderGrid.RowData(pRowIndex) And RowDataBracketOrderMask) / RowDataBracketOrderBase
End Function

Private Function getPositionManagerGridMappingIndexFromRowIndex(ByVal pRowIndex As Long) As Long
getPositionManagerGridMappingIndexFromRowIndex = (BracketOrderGrid.RowData(pRowIndex) And RowDataPositionManagerMask) / RowDataPositionManagerBase
End Function

Private Function getMessage(ByVal pOrder As IOrder) As String
If pOrder.ErrorMessage <> "" Then
    getMessage = pOrder.ErrorMessage
ElseIf pOrder.Message <> "" Then
    getMessage = pOrder.Message
End If
End Function

Private Function getOrderRole(ByVal pBracketOrderIndex As Long, ByVal pOrder As IOrder) As BracketOrderRoles
Dim lBracketOrder As IBracketOrder
Set lBracketOrder = mBracketOrderGridMappingTable(pBracketOrderIndex).BracketOrder
If pOrder Is lBracketOrder.EntryOrder Then
    getOrderRole = BracketOrderRoleEntry
ElseIf pOrder Is lBracketOrder.StopLossOrder Then
    getOrderRole = BracketOrderRoleStopLoss
ElseIf pOrder Is lBracketOrder.TargetOrder Then
    getOrderRole = BracketOrderRoleTarget
ElseIf pOrder Is lBracketOrder.CloseoutOrder Then
    getOrderRole = BracketOrderRoleCloseout
End If
End Function

Private Function getSelectedOrder() As IOrder
Const ProcName As String = "getSelectedOrder"
On Error GoTo Err

Dim lIndex As Long
lIndex = getSelectedBracketOrderGridMappingIndex
If lIndex = NullIndex Then Exit Function

Select Case SelectedOrderRole
Case BracketOrderRoles.BracketOrderRoleEntry
    Set getSelectedOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder.EntryOrder
Case BracketOrderRoles.BracketOrderRoleStopLoss
    Set getSelectedOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder.StopLossOrder
Case BracketOrderRoles.BracketOrderRoleTarget
    Set getSelectedOrder = mBracketOrderGridMappingTable(lIndex).BracketOrder.TargetOrder
Case BracketOrderRoles.BracketOrderRoleCloseout
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
getSelectedBracketOrderGridMappingIndex = getBracketOrderGridMappingIndexFromRowIndex(BracketOrderGrid.Row)
End Function

Private Property Let GridColumn(ByVal pRowIndex As Long, ByVal pColumnIndex As Long, ByVal Value As String)
BracketOrderGrid.TextMatrix(pRowIndex, pColumnIndex) = Value
End Property

Private Property Get GridColumn(ByVal pRowIndex As Long, ByVal pColumnIndex As Long) As String
GridColumn = BracketOrderGrid.TextMatrix(pRowIndex, pColumnIndex)
End Property

Private Function gridRowIsBracketOrder(ByVal pRowIndex As Long) As Boolean
gridRowIsBracketOrder = ((BracketOrderGrid.RowData(pRowIndex) And RowDataBracketOrderMask) <> 0) And _
                        ((BracketOrderGrid.RowData(pRowIndex) And RowDataOrderRoleMask) = 0)
End Function

Private Function gridRowIsOrder(ByVal pRowIndex As Long) As Boolean
gridRowIsOrder = (BracketOrderGrid.RowData(pRowIndex) And RowDataOrderRoleMask) <> 0
End Function

Private Function gridRowIsPositionManager(ByVal pRowIndex As Long) As Boolean
gridRowIsPositionManager = ((BracketOrderGrid.RowData(pRowIndex) And RowDataPositionManagerMask) <> 0) And _
                        ((BracketOrderGrid.RowData(pRowIndex) And RowDataBracketOrderMask) = 0) And _
                        ((BracketOrderGrid.RowData(pRowIndex) And RowDataOrderRoleMask) = 0)
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

Private Sub hideMessage(ByVal pMessageIndex As Long, ByVal pMessageRow As Long)
Const ProcName As String = "hideMessage"
On Error GoTo Err

MessageText(mMessageMappingTable(pMessageIndex).MessageTextIndex).Visible = False
BracketOrderGrid.RowHeight(pMessageRow) = 0

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

Private Sub setExpandOrCollapseIcon(ByVal pGridRow As Long, ByVal pShowExpand As Boolean)
Const ProcName As String = "setExpandOrCollapseIcon"
On Error GoTo Err

Dim lPrevRow As Long
lPrevRow = BracketOrderGrid.Row
Dim lPrevCol As Long
lPrevCol = BracketOrderGrid.col

BracketOrderGrid.Row = pGridRow
BracketOrderGrid.col = ExpandIndicator
BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages(IIf(pShowExpand, "Expand", "Contract")).Picture

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
    .Rows = 100
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
                ByVal Align As MSFlexGridLib.AlignmentSettings)
Const ProcName As String = "setupBracketOrderGridColumn"
On Error GoTo Err

Dim lColumnWidth As Long
Dim i As Long

With BracketOrderGrid
    .Row = rowNumber
    If (columnNumber + 1) > .Cols Then
        For i = .Cols To columnNumber
            .Cols = i + 1
            .ColWidth(i) = 0
        Next
    End If
    
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
    If .ColWidth(columnNumber) < lColumnWidth Then
        .ColWidth(columnNumber) = lColumnWidth
    End If
    
    .ColAlignment(columnNumber) = Align
    .TextMatrix(rowNumber, columnNumber) = columnHeader
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupMessage(ByVal pOrder As IOrder, ByVal pMessage As String, ByVal pBracketOrderIndex As Long)
Const ProcName As String = "setupMessage"
On Error GoTo Err

If pMessage = "" Then Exit Sub

With mBracketOrderGridMappingTable(pBracketOrderIndex)
    Dim lSymbol As String
    lSymbol = .BracketOrder.Contract.Specifier.LocalSymbol

    Dim lRole As BracketOrderRoles
    lRole = getOrderRole(pBracketOrderIndex, pOrder)
    
    Dim lMessageIndex As Long
    lMessageIndex = allocateMessageIndex(pBracketOrderIndex, lRole, pMessage, lSymbol, .IsExpanded)
    
    Select Case lRole
    Case BracketOrderRoles.BracketOrderRoleEntry
        .EntryMessageIndex = lMessageIndex
        If .StopLossGridOffset <> NullIndex Then .StopLossGridOffset = .StopLossGridOffset + 1
        If .TargetGridOffset <> NullIndex Then .TargetGridOffset = .TargetGridOffset + 1
        If .CloseoutGridOffset <> NullIndex Then .CloseoutGridOffset = .CloseoutGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleStopLoss
        .StopLossMessageIndex = lMessageIndex
        If .TargetGridOffset <> NullIndex Then .TargetGridOffset = .TargetGridOffset + 1
        If .CloseoutGridOffset <> NullIndex Then .CloseoutGridOffset = .CloseoutGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleTarget
        .TargetMessageIndex = lMessageIndex
        If .CloseoutGridOffset <> NullIndex Then .CloseoutGridOffset = .CloseoutGridOffset + 1
    Case BracketOrderRoles.BracketOrderRoleCloseout
        .CloseoutMessageIndex = lMessageIndex
    End Select
        
    displayMessage lMessageIndex
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
    BracketOrderGrid.RowData(.GridIndex) = generateRowData(lIndex, 0, 0)
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
    setExpandOrCollapseIcon .GridIndex, False
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

If Not EditText.Visible Then Exit Sub

Dim lOrder As IOrder
Set lOrder = getSelectedOrder

Dim lPrice As Double
Dim lPriceSpec As PriceSpecifier
If mEditedCol = OrderPrice Or mEditedCol = OrderAuxPrice Then
    If ParsePrice(EditText.Text, mSelectedBracketOrder.Contract.Specifier.SecType, mSelectedBracketOrder.Contract.TickSize, lPrice) Then
        Set lPriceSpec = NewPriceSpecifier(lPrice, PriceValueTypeValue)
    End If
End If

If mEditedCol = OrderPrice Then
    If Not lPriceSpec Is Nothing Then
        Select Case SelectedOrderRole
        Case BracketOrderRoleEntry
            mSelectedBracketOrder.SetNewEntryLimitPrice lPriceSpec
        Case BracketOrderRoleStopLoss
            mSelectedBracketOrder.SetNewStopLossLimitPrice lPriceSpec
        Case BracketOrderRoleTarget
            mSelectedBracketOrder.SetNewTargetLimitPrice lPriceSpec
        End Select
    End If
ElseIf mEditedCol = OrderAuxPrice Then
    If Not lPriceSpec Is Nothing Then
        Select Case SelectedOrderRole
        Case BracketOrderRoleEntry
            mSelectedBracketOrder.SetNewEntryTriggerPrice lPriceSpec
        Case BracketOrderRoleStopLoss
            mSelectedBracketOrder.SetNewStopLossTriggerPrice lPriceSpec
        Case BracketOrderRoleTarget
            mSelectedBracketOrder.SetNewTargetTriggerPrice lPriceSpec
        End Select
    End If
ElseIf mEditedCol = OrderQuantity Then
    If IsInteger(EditText.Text, 0) Then
        Select Case SelectedOrderRole
        Case BracketOrderRoleEntry
            mSelectedBracketOrder.SetNewEntryQuantity CLng(EditText.Text)
        Case BracketOrderRoleStopLoss
            mSelectedBracketOrder.SetNewStopLossQuantity CLng(EditText.Text)
        Case BracketOrderRoleTarget
            mSelectedBracketOrder.SetNewTargetQuantity CLng(EditText.Text)
        End Select
    End If
End If
    
If mSelectedBracketOrder.IsDirty Then mSelectedBracketOrder.Update

endEdit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

