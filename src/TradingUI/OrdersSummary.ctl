VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
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

Private Const ModuleName                As String = "OrdersSummary"

Private Const RowDataBracketOrderBase As Long = &H100
Private Const RowDataPositionManagerBase As Long = &H1000000

'@================================================================================
' Enums
'@================================================================================

Private Enum BracketOrderGridColumns
    Symbol
    ExpandIndicator
    OtherColumns    ' keep this entry last
End Enum

Private Enum BracketOrderGridBracketOrderColumns
    CreationTime = BracketOrderGridColumns.OtherColumns
    Size
    Profit
    MaxProfit
    Drawdown
    CurrencyCode
End Enum

Private Enum BracketOrderGridPositionColumns
    Exchange = BracketOrderGridColumns.OtherColumns
    Size
    Profit
    MaxProfit
    Drawdown
    CurrencyCode
End Enum

Private Enum BracketOrderGridOrderColumns
    OrderMode = BracketOrderGridColumns.OtherColumns
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

Private Enum BracketOrderGridColumnWidths
    ExpandIndicatorWidth = 3
    SymbolWidth = 15
End Enum

Private Enum BracketOrderGridBracketOrderColumnWidths
    CreationTimeWidth = 15
    SizeWidth = 6
    ProfitWidth = 9
    MaxProfitWidth = 9
    DrawdownWidth = 9
    CurrencyCodeWidth = 4
End Enum

Private Enum BracketOrderGridPositionColumnWidths
    ExchangeWidth = 9
    SizeWidth = 6
    ProfitWidth = 9
    MaxProfitWidth = 9
    DrawdownWidth = 9
    CurrencyCodeWidth = 5
End Enum

Private Enum BracketOrderGridOrderColumnWidths
    OrderModeWidth = 9
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
    StopGridOffset      As Long
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

Private mSelectedBracketOrderGridRow                        As Long
Private mSelectedBracketOrder                               As IBracketOrder
Private mSelectedOrderIndex                                 As Long

Private mBracketOrderGridMappingTable()                     As BracketOrderGridMappingEntry
Private mMaxBracketOrderGridMappingTableIndex               As Long

Private mPositionManagerGridMappingTable()                  As PositionManagerGridMappingEntry
Private mMaxPositionManagerGridMappingTableIndex            As Long

' the index of the first entry in the bracket order grid that relates to
' bracket orders (rather than header rows, currency totals etc)
Private mFirstBracketOrderGridRowIndex                      As Long

Private mLetterWidth                                        As Single
Private mDigitWidth                                         As Single

Private mMonitoredPositions                                 As EnumerableCollection
    
Private mIsEditing                                          As Boolean
Private mEditedBracketOrder                                 As IBracketOrder
Private mEditedOrderIndex                                   As Long
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

Set mMonitoredPositions = New EnumerableCollection

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
    Dim lBracketOrderChangeType As BracketOrderChangeTypes
    Dim lBracketOrder As IBracketOrder
    Dim lBracketOrderIndex As Long
    
    Set lBracketOrder = ev.Source
    
    lBracketOrderIndex = findBracketOrderTableIndex(lBracketOrder)
    
    With mBracketOrderGridMappingTable(lBracketOrderIndex)
    
        lBracketOrderChangeType = ev.changeType
        
        Select Case lBracketOrderChangeType
        Case BracketOrderChangeTypes.BracketOrderCreated
            
        Case BracketOrderChangeTypes.BracketOrderCompleted
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            If lBracketOrder.Size = 0 Then lBracketOrder.RemoveChangeListener Me
        Case BracketOrderChangeTypes.BracketOrderSelfCancelled
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            If lBracketOrder.Size = 0 Then lBracketOrder.RemoveChangeListener Me
        Case BracketOrderChangeTypes.BracketOrderEntryOrderChanged
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .EntryGridOffset, lBracketOrder.EntryOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderStopOrderChanged
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .StopGridOffset, lBracketOrder.StopLossOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderTargetOrderChanged
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .TargetGridOffset, lBracketOrder.TargetOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderCloseoutOrderCreated
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            If .TargetGridOffset >= 0 Then
                .CloseoutGridOffset = .TargetGridOffset + 1
            ElseIf .StopGridOffset >= 0 Then
                .CloseoutGridOffset = .StopGridOffset + 1
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
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            displayOrderValues .GridIndex + .CloseoutGridOffset, lBracketOrder.CloseoutOrder, .secType, .TickSize
        Case BracketOrderChangeTypes.BracketOrderSizeChanged
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            BracketOrderGrid.TextMatrix(.GridIndex, BracketOrderGridBracketOrderColumns.Size) = lBracketOrder.Size
        Case BracketOrderChangeTypes.BracketOrderStateChanged
            If lBracketOrder Is mEditedBracketOrder Then endEdit
            If lBracketOrder.State = BracketOrderStates.BracketOrderStateSubmitted Then
                BracketOrderGrid.TextMatrix(.GridIndex, BracketOrderGridBracketOrderColumns.CreationTime) = formattedTime(lBracketOrder.CreationTime)
            End If
            If lBracketOrder.State <> BracketOrderStates.BracketOrderStateCreated And _
                lBracketOrder.State <> BracketOrderStates.BracketOrderStateSubmitted _
            Then
                ' the bracket order is now in a state where it can't be modified.
                ' If it's the currently selected bracket order, make it not so.
                If lBracketOrder Is mSelectedBracketOrder Then
                    invertEntryColors mSelectedBracketOrderGridRow
                    mSelectedBracketOrderGridRow = -1
                    Set mSelectedBracketOrder = Nothing
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
        showPositionManagerEntry pm
        BracketOrderGrid.TextMatrix(mPositionManagerGridMappingTable(pmIndex).GridIndex, _
                                BracketOrderGridPositionColumns.Size) = pm.PositionSize
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
Const ProcName As String = "CollectionChangeListener_Change"
On Error GoTo Err

If TypeOf ev.Source Is BracketOrders Then
    Dim lBracketOrder As IBracketOrder
    
    If IsEmpty(ev.AffectedItem) Then Exit Sub
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
    
        addBracketOrder lBracketOrder
    Case CollItemRemoved
        Dim lBracketOrderIndex As Long
        lBracketOrderIndex = findBracketOrderTableIndex(lBracketOrder)
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
    lBracketOrderIndex = findBracketOrderTableIndex(lBracketOrder)
    rowIndex = mBracketOrderGridMappingTable(lBracketOrderIndex).GridIndex
    
    Dim lBOProfitType As ProfitTypes
    lBOProfitType = ev.ProfitType
    
    Select Case lBOProfitType
    Case ProfitTypes.ProfitTypeProfit
        displayProfitValue ev.ProfitAmount, rowIndex, BracketOrderGridBracketOrderColumns.Profit
    Case ProfitTypes.ProfitTypeMaxProfit
        displayProfitValue ev.ProfitAmount, rowIndex, BracketOrderGridBracketOrderColumns.MaxProfit
    Case ProfitTypes.ProfitTypeDrawdown
        displayProfitValue -ev.ProfitAmount, rowIndex, BracketOrderGridBracketOrderColumns.Drawdown
    End Select

ElseIf TypeOf ev.Source Is PositionManager Then
    Dim lPositionManager As PositionManager
    Set lPositionManager = ev.Source
    
    showPositionManagerEntry lPositionManager
    
    Dim lPositionManagerIndex As Long
    lPositionManagerIndex = findPositionManagerTableIndex(lPositionManager)
    rowIndex = mPositionManagerGridMappingTable(lPositionManagerIndex).GridIndex
    
    Dim lPMProfitType As ProfitTypes
    lPMProfitType = ev.ProfitType
    
    Select Case lPMProfitType
    Case ProfitTypes.ProfitTypeSessionProfit
        displayProfitValue ev.ProfitAmount, rowIndex, BracketOrderGridPositionColumns.Profit
    Case ProfitTypes.ProfitTypeSessionMaxProfit
        displayProfitValue ev.ProfitAmount, rowIndex, BracketOrderGridPositionColumns.MaxProfit
    Case ProfitTypes.ProfitTypeSessionDrawdown
        displayProfitValue -ev.ProfitAmount, rowIndex, BracketOrderGridPositionColumns.Drawdown
    Case ProfitTypes.ProfitTypeTradeProfit
    Case ProfitTypes.ProfitTypeTradeMaxProfit
    Case ProfitTypes.ProfitTypeTradeDrawdown
    End Select
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

Dim lRow As Long
Dim lRowdata As Long
Dim op As IBracketOrder
Dim index As Long
Dim selectedOrder As IOrder

lRow = BracketOrderGrid.Row

If BracketOrderGrid.MouseCol = BracketOrderGridColumns.Symbol Then
    RaiseEvent Click
    Exit Sub
End If

If BracketOrderGrid.MouseCol = BracketOrderGridColumns.ExpandIndicator Then
    expandOrContract
    adjustEditBox
Else

    invertEntryColors mSelectedBracketOrderGridRow
    
    mSelectedBracketOrderGridRow = -1
    
    BracketOrderGrid.Row = lRow
    lRowdata = BracketOrderGrid.RowData(lRow)
    If lRowdata < RowDataPositionManagerBase And _
        lRowdata >= RowDataBracketOrderBase _
    Then
        index = lRowdata - RowDataBracketOrderBase
        Set op = mBracketOrderGridMappingTable(index).BracketOrder
        If op.State = BracketOrderStates.BracketOrderStateCreated Or _
            op.State = BracketOrderStates.BracketOrderStateSubmitted _
        Then
            
            mSelectedBracketOrderGridRow = lRow
            Set mSelectedBracketOrder = op
            invertEntryColors mSelectedBracketOrderGridRow
            
            mSelectedOrderIndex = mSelectedBracketOrderGridRow - mBracketOrderGridMappingTable(index).GridIndex
            If mSelectedOrderIndex <> 0 Then
                Set selectedOrder = op.Order(mSelectedOrderIndex)
                If selectedOrder.IsModifiable Then
                    If (BracketOrderGrid.MouseCol = BracketOrderGridOrderColumns.Price And _
                            selectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttLimitPrice)) Or _
                        (BracketOrderGrid.MouseCol = BracketOrderGridOrderColumns.AuxPrice And _
                            selectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttTriggerPrice)) Or _
                        (BracketOrderGrid.MouseCol = BracketOrderGridOrderColumns.Quantity And _
                        selectedOrder.IsAttributeModifiable(OrderAttributes.OrderAttQuantity)) _
                    Then
                        mIsEditing = True
                        Set mEditedBracketOrder = op
                        mEditedOrderIndex = mSelectedOrderIndex
                        mEditedCol = BracketOrderGrid.MouseCol
                        BracketOrderGrid.col = mEditedCol
                        
                        EditText.Text = BracketOrderGrid.Text
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

Dim selectedOrder As IOrder

If mSelectedOrderIndex = 0 Then Exit Property

Set selectedOrder = mSelectedBracketOrder.Order(mSelectedOrderIndex)
If Not selectedOrder Is Nothing Then
    IsSelectedItemModifiable = selectedOrder.IsModifiable
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelectedItem() As IBracketOrder
Set SelectedItem = mSelectedBracketOrder
End Property

Public Property Get SelectedOrderIndex() As Long
SelectedOrderIndex = mSelectedOrderIndex
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim i As Long
For i = 0 To mMaxBracketOrderGridMappingTableIndex
    If Not mBracketOrderGridMappingTable(i).BracketOrder Is Nothing Then
        mBracketOrderGridMappingTable(i).BracketOrder.RemoveChangeListener Me
        mBracketOrderGridMappingTable(i).ProfitCalculator.RemoveProfitListener Me
        Set mBracketOrderGridMappingTable(i).BracketOrder = Nothing
    End If
Next

Dim en As Enumerator
Set en = mMonitoredPositions.Enumerator
Do While en.MoveNext
    Dim lPm As PositionManager
    Set lPm = en.Current
    If Not lPm.IsFinished Then
        lPm.RemoveChangeListener Me
        lPm.RemoveProfitListener Me
        lPm.BracketOrders.RemoveCollectionChangeListener Me
    End If
    en.Remove
Loop

BracketOrderGrid.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise(ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pMarketDataManager Is Nothing, "pMarketDataManager must be supplied"
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

Dim lPositionManager As PositionManager
For Each lPositionManager In pPositionManagers
    mMonitoredPositions.Add lPositionManager

    Dim lBracketOrder As IBracketOrder
    
    If lPositionManager.BracketOrders.Count <> 0 Or lPositionManager.PositionSize <> 0 Or lPositionManager.PendingPositionSize <> 0 Then
        showPositionManagerEntry lPositionManager
    End If
    
    lPositionManager.AddProfitListener Me
    
    Dim lAnyActiveBracketOrders As Boolean
    For Each lBracketOrder In lPositionManager.BracketOrders
        addBracketOrder lBracketOrder
        If lBracketOrder.State = BracketOrderStateClosed Then
            contractBracketOrderEntry findBracketOrderTableIndex(lBracketOrder)
        Else
            lAnyActiveBracketOrders = True
        End If
    Next
    
    If Not lAnyActiveBracketOrders Then contractPositionManagerEntry findPositionManagerTableIndex(lPositionManager)
    
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "addBracketOrder"
On Error GoTo Err

pBracketOrder.AddChangeListener Me
displayBracketOrder pBracketOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function addBracketOrderEntryToBracketOrderGrid( _
                ByVal pSymbol As String, _
                ByVal BracketOrderTableIndex As Long) As Long
Const ProcName As String = "addBracketOrderEntryToBracketOrderGrid"
On Error GoTo Err

Dim index As Long
index = addEntryToBracketOrderGrid(pSymbol, False)

BracketOrderGrid.RowData(index) = BracketOrderTableIndex + RowDataBracketOrderBase

BracketOrderGrid.Row = index
BracketOrderGrid.col = BracketOrderGridColumns.ExpandIndicator
BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture

BracketOrderGrid.col = BracketOrderGridBracketOrderColumns.Profit
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

BracketOrderGrid.col = BracketOrderGridBracketOrderColumns.MaxProfit
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

BracketOrderGrid.col = BracketOrderGridBracketOrderColumns.Drawdown
BracketOrderGrid.CellBackColor = &HC0C0C0
BracketOrderGrid.CellForeColor = vbWhite

addBracketOrderEntryToBracketOrderGrid = index

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                
Private Function addEntryToBracketOrderGrid( _
                ByVal pSymbol As String, _
                Optional ByVal pBefore As Boolean, _
                Optional ByVal pIndex As Long = -1) As Long
Const ProcName As String = "addEntryToBracketOrderGrid"
On Error GoTo Err

Dim i As Long

If pIndex < 0 Then
    For i = mFirstBracketOrderGridRowIndex To BracketOrderGrid.Rows - 1
        If (pBefore And _
            BracketOrderGrid.TextMatrix(i, BracketOrderGridColumns.Symbol) >= pSymbol) Or _
            BracketOrderGrid.TextMatrix(i, BracketOrderGridColumns.Symbol) = "" _
        Then
            pIndex = i
            Exit For
        ElseIf (Not pBefore And _
            BracketOrderGrid.TextMatrix(i, BracketOrderGridColumns.Symbol) > pSymbol) Or _
            BracketOrderGrid.TextMatrix(i, BracketOrderGridColumns.Symbol) = "" _
        Then
            pIndex = i
            Exit For
        End If
    Next
    
    If pIndex < 0 Then
        BracketOrderGrid.addItem ""
        pIndex = BracketOrderGrid.Rows - 1
    ElseIf BracketOrderGrid.TextMatrix(pIndex, BracketOrderGridColumns.Symbol) = "" Then
        BracketOrderGrid.TextMatrix(pIndex, BracketOrderGridColumns.Symbol) = pSymbol
    Else
        BracketOrderGrid.addItem "", pIndex
    End If
Else
    BracketOrderGrid.addItem "", pIndex
End If

BracketOrderGrid.TextMatrix(pIndex, BracketOrderGridColumns.Symbol) = pSymbol
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

pIndex = addEntryToBracketOrderGrid(pSymbol, False, pIndex)

BracketOrderGrid.RowData(pIndex) = pBracketOrderTableIndex + RowDataBracketOrderBase

BracketOrderGrid.TextMatrix(pIndex, BracketOrderGridOrderColumns.OrderMode) = pOrderMode

displayOrderValues pIndex, pOrder, pSecType, pTickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub adjustEditBox()
Const ProcName As String = "adjustEditBox"
On Error GoTo Err

Dim opIndex As Long

If mIsEditing Then
    opIndex = findBracketOrderTableIndex(mEditedBracketOrder)
    BracketOrderGrid.Row = mBracketOrderGridMappingTable(opIndex).GridIndex + mEditedOrderIndex
    BracketOrderGrid.col = mEditedCol
    
    EditText.Move BracketOrderGrid.Left + BracketOrderGrid.CellLeft + 8, _
                BracketOrderGrid.Top + BracketOrderGrid.Celltop + 8, _
                BracketOrderGrid.CellWidth - 16, _
                BracketOrderGrid.CellHeight - 16
End If

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
    
    If mIsEditing And .BracketOrder Is mEditedBracketOrder Then endEdit
    
    If .EntryGridOffset >= 0 Then
        lIndex = .GridIndex + .EntryGridOffset
        BracketOrderGrid.rowHeight(lIndex) = 0
    End If
    If .StopGridOffset >= 0 Then
        lIndex = .GridIndex + .StopGridOffset
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
        .IsExpanded = False
        BracketOrderGrid.Row = .GridIndex
        BracketOrderGrid.col = BracketOrderGridColumns.ExpandIndicator
        BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Expand").Picture
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

Dim i As Long
Dim lSymbol As String
Dim lOpEntryIndex As Long

mPositionManagerGridMappingTable(index).IsExpanded = False
BracketOrderGrid.Row = mPositionManagerGridMappingTable(index).GridIndex
BracketOrderGrid.col = BracketOrderGridColumns.ExpandIndicator
BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Expand").Picture

lSymbol = BracketOrderGrid.TextMatrix(mPositionManagerGridMappingTable(index).GridIndex, BracketOrderGridColumns.Symbol)
i = mPositionManagerGridMappingTable(index).GridIndex + 1
Do While BracketOrderGrid.TextMatrix(i, BracketOrderGridColumns.Symbol) = lSymbol
    BracketOrderGrid.rowHeight(i) = 0
    lOpEntryIndex = BracketOrderGrid.RowData(i) - RowDataBracketOrderBase
    i = contractBracketOrderEntry(lOpEntryIndex, True) + 1
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displayBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "displayBracketOrder"
On Error GoTo Err

Dim lSymbol As String
lSymbol = pBracketOrder.Contract.Specifier.LocalSymbol

Dim lIndex As Long
lIndex = findBracketOrderTableIndex(pBracketOrder)

With mBracketOrderGridMappingTable(lIndex)
    If .BracketOrder Is Nothing Then
        
        .IsExpanded = True
        .EntryGridOffset = -1
        .StopGridOffset = -1
        .TargetGridOffset = -1
        .CloseoutGridOffset = -1
        
        Set .BracketOrder = pBracketOrder
        .TickSize = pBracketOrder.Contract.TickSize
        .secType = pBracketOrder.Contract.Specifier.secType
        .GridIndex = addBracketOrderEntryToBracketOrderGrid(pBracketOrder.Contract.Specifier.LocalSymbol, lIndex)
        BracketOrderGrid.TextMatrix(.GridIndex, BracketOrderGridBracketOrderColumns.CreationTime) = formattedTime(pBracketOrder.CreationTime)
        BracketOrderGrid.TextMatrix(.GridIndex, BracketOrderGridBracketOrderColumns.CurrencyCode) = pBracketOrder.Contract.Specifier.CurrencyCode
        
        Dim lDataSource As IMarketDataSource
        Set lDataSource = mMarketDataManager.CreateMarketDataSource(CreateFuture(pBracketOrder.Contract), False)
        lDataSource.StartMarketData
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
                .StopGridOffset = .EntryGridOffset + 1
            Else
                .StopGridOffset = 1
            End If
            addOrderEntryToBracketOrderGrid .GridIndex + .StopGridOffset, _
                                    lSymbol, _
                                    lOrder, _
                                    lIndex, _
                                    "Stop Loss", _
                                    .secType, _
                                    .TickSize
        End If
        
        Set lOrder = pBracketOrder.TargetOrder
        If Not lOrder Is Nothing Then
            If .StopGridOffset >= 0 Then
                .TargetGridOffset = .StopGridOffset + 1
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

BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.Action) = OrderActionToString(pOrder.Action)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.AuxPrice) = FormatPrice(pOrder.TriggerPrice, pSecType, pTickSize)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.AveragePrice) = FormatPrice(pOrder.AveragePrice, pSecType, pTickSize)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.Id) = pOrder.Id
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.LastFillPrice) = FormatPrice(pOrder.LastFillPrice, pSecType, pTickSize)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.LastFillTime) = formattedTime(pOrder.FillTime)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.OrderType) = OrderTypeToShortString(pOrder.OrderType)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.Price) = FormatPrice(pOrder.LimitPrice, pSecType, pTickSize)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.Quantity) = pOrder.Quantity
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.QuantityRemaining) = pOrder.QuantityRemaining
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.Size) = IIf(pOrder.QuantityFilled <> 0, pOrder.QuantityFilled, 0)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.Status) = OrderStatusToString(pOrder.Status)
BracketOrderGrid.TextMatrix(pGridIndex, BracketOrderGridOrderColumns.BrokerId) = pOrder.BrokerId

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

BracketOrderGrid.Row = pRowIndex
BracketOrderGrid.col = pColIndex
BracketOrderGrid.Text = Format(pProfit, "0.00")
If pProfit >= 0 Then
    BracketOrderGrid.CellForeColor = CPositiveProfitColor
Else
    BracketOrderGrid.CellForeColor = CNegativeProfitColor
End If

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
Set mEditedBracketOrder = Nothing
mEditedOrderIndex = -1
mEditedCol = -1
BracketOrderGrid.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Private Sub expandOrContract()
Const ProcName As String = "expandOrContract"
On Error GoTo Err

Dim RowData As Long
Dim index As Long
Dim expanded As Boolean

RowData = BracketOrderGrid.RowData(BracketOrderGrid.MouseRow)
If RowData >= RowDataPositionManagerBase Then
    index = RowData - RowDataPositionManagerBase
    expanded = mPositionManagerGridMappingTable(index).IsExpanded
    If expanded Then
        contractPositionManagerEntry index
    Else
        expandPositionManagerEntry index
    End If
ElseIf RowData >= RowDataBracketOrderBase Then
    index = RowData - RowDataBracketOrderBase
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
Else
    Exit Sub
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
    If .StopGridOffset >= 0 Then
        lIndex = .GridIndex + .StopGridOffset
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
        .IsExpanded = True
        BracketOrderGrid.Row = .GridIndex
        BracketOrderGrid.col = BracketOrderGridColumns.ExpandIndicator
        BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
        Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture
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

Dim i As Long
Dim lSymbol As String
Dim lOpEntryIndex As Long

mPositionManagerGridMappingTable(index).IsExpanded = True
BracketOrderGrid.Row = mPositionManagerGridMappingTable(index).GridIndex
BracketOrderGrid.col = BracketOrderGridColumns.ExpandIndicator
BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture

lSymbol = BracketOrderGrid.TextMatrix(mPositionManagerGridMappingTable(index).GridIndex, BracketOrderGridColumns.Symbol)
i = mPositionManagerGridMappingTable(index).GridIndex + 1
Do While BracketOrderGrid.TextMatrix(i, BracketOrderGridColumns.Symbol) = lSymbol
    BracketOrderGrid.rowHeight(i) = -1
    lOpEntryIndex = BracketOrderGrid.RowData(i) - RowDataBracketOrderBase
    i = expandBracketOrderEntry(lOpEntryIndex, True) + 1
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findBracketOrderTableIndex(ByVal pBracketOrder As IBracketOrder) As Long
Const ProcName As String = "findBracketOrderTableIndex"
On Error GoTo Err

Dim lBracketOrderIndex As Long
lBracketOrderIndex = pBracketOrder.ApplicationIndex
Do While lBracketOrderIndex > UBound(mBracketOrderGridMappingTable)
    ReDim Preserve mBracketOrderGridMappingTable(2 * (UBound(mBracketOrderGridMappingTable) + 1) - 1) As BracketOrderGridMappingEntry
Loop
If lBracketOrderIndex > mMaxBracketOrderGridMappingTableIndex Then mMaxBracketOrderGridMappingTableIndex = lBracketOrderIndex

findBracketOrderTableIndex = lBracketOrderIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function findPositionManagerTableIndex(ByVal pm As PositionManager) As Long
Const ProcName As String = "findPositionManagerTableIndex"
On Error GoTo Err

Dim pmIndex As Long
pmIndex = pm.ApplicationIndex

Do While pmIndex > UBound(mPositionManagerGridMappingTable)
    ReDim Preserve mPositionManagerGridMappingTable(2 * (UBound(mPositionManagerGridMappingTable) + 1) - 1) As PositionManagerGridMappingEntry
Loop
If pmIndex > mMaxPositionManagerGridMappingTableIndex Then mMaxPositionManagerGridMappingTableIndex = pmIndex

findPositionManagerTableIndex = pmIndex

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

Dim lForeColor As Long
Dim lBackColor As Long
Dim i As Long

If rowNumber < 0 Then Exit Sub

BracketOrderGrid.Row = rowNumber

For i = BracketOrderGridColumns.OtherColumns To BracketOrderGrid.Cols - 1
    BracketOrderGrid.col = i
    lForeColor = IIf(BracketOrderGrid.CellForeColor = 0, BracketOrderGrid.ForeColor, BracketOrderGrid.CellForeColor)
    If lForeColor = SystemColorConstants.vbWindowText Then
        BracketOrderGrid.CellForeColor = SystemColorConstants.vbHighlightText
    ElseIf lForeColor = SystemColorConstants.vbHighlightText Then
        BracketOrderGrid.CellForeColor = SystemColorConstants.vbWindowText
    ElseIf lForeColor > 0 Then
        BracketOrderGrid.CellForeColor = IIf((lForeColor Xor &HFFFFFF) = 0, 1, lForeColor Xor &HFFFFFF)
    End If
    
    lBackColor = IIf(BracketOrderGrid.CellBackColor = 0, BracketOrderGrid.BackColor, BracketOrderGrid.CellBackColor)
    If lBackColor = SystemColorConstants.vbWindowBackground Then
        BracketOrderGrid.CellBackColor = SystemColorConstants.vbHighlight
    ElseIf lBackColor = SystemColorConstants.vbHighlight Then
        BracketOrderGrid.CellBackColor = SystemColorConstants.vbWindowBackground
    ElseIf lBackColor > 0 Then
        BracketOrderGrid.CellBackColor = IIf((lBackColor Xor &HFFFFFF) = 0, 1, lBackColor Xor &HFFFFFF)
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupBracketOrderGrid()
Const ProcName As String = "setupBracketOrderGrid"
On Error GoTo Err

With BracketOrderGrid
    mSelectedBracketOrderGridRow = -1
    
    .Redraw = False
    .AllowUserResizing = flexResizeBoth
    
    .Cols = 0
    .Rows = 20
    .FixedRows = 3
    ' .FixedCols = 1
    
    setupBracketOrderGridColumn 0, BracketOrderGridColumns.ExpandIndicator, BracketOrderGridColumnWidths.ExpandIndicatorWidth, "", True, MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    setupBracketOrderGridColumn 0, BracketOrderGridColumns.Symbol, BracketOrderGridColumnWidths.SymbolWidth, "Symbol", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    
    setupBracketOrderGridColumn 0, BracketOrderGridPositionColumns.CurrencyCode, BracketOrderGridPositionColumnWidths.CurrencyCodeWidth, "Curr", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 0, BracketOrderGridPositionColumns.Drawdown, BracketOrderGridPositionColumnWidths.DrawdownWidth, "Drawdown", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 0, BracketOrderGridPositionColumns.Exchange, BracketOrderGridPositionColumnWidths.ExchangeWidth, "Exchange", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 0, BracketOrderGridPositionColumns.MaxProfit, BracketOrderGridPositionColumnWidths.MaxProfitWidth, "Max", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 0, BracketOrderGridPositionColumns.Profit, BracketOrderGridPositionColumnWidths.ProfitWidth, "Profit", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 0, BracketOrderGridPositionColumns.Size, BracketOrderGridPositionColumnWidths.SizeWidth, "Size", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    
    setupBracketOrderGridColumn 1, BracketOrderGridBracketOrderColumns.CreationTime, BracketOrderGridBracketOrderColumnWidths.CreationTimeWidth, "Creation Time", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketOrderGridBracketOrderColumns.CurrencyCode, BracketOrderGridBracketOrderColumnWidths.CurrencyCodeWidth, "Curr", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 1, BracketOrderGridBracketOrderColumns.Drawdown, BracketOrderGridBracketOrderColumnWidths.DrawdownWidth, "Drawdown", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketOrderGridBracketOrderColumns.MaxProfit, BracketOrderGridBracketOrderColumnWidths.MaxProfitWidth, "Max", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketOrderGridBracketOrderColumns.Profit, BracketOrderGridBracketOrderColumnWidths.ProfitWidth, "Profit", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 1, BracketOrderGridBracketOrderColumns.Size, BracketOrderGridBracketOrderColumnWidths.SizeWidth, "Size", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.Action, BracketOrderGridOrderColumnWidths.ActionWidth, "Action", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.AuxPrice, BracketOrderGridOrderColumnWidths.AuxPriceWidth, "Trigger", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.AveragePrice, BracketOrderGridOrderColumnWidths.AveragePriceWidth, "Avg fill", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.Id, BracketOrderGridOrderColumnWidths.IdWidth, "Id", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.LastFillPrice, BracketOrderGridOrderColumnWidths.LastFillPriceWidth, "Last fill", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.LastFillTime, BracketOrderGridOrderColumnWidths.LastFillTimeWidth, "Last fill time", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.OrderType, BracketOrderGridOrderColumnWidths.OrderTypeWidth, "Type", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.Price, BracketOrderGridOrderColumnWidths.PriceWidth, "Price", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.Quantity, BracketOrderGridOrderColumnWidths.QuantityWidth, "Qty", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.QuantityRemaining, BracketOrderGridOrderColumnWidths.QuantityRemainingWidth, "Rem", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.Size, BracketOrderGridOrderColumnWidths.SizeWidth, "Size", False, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.Status, BracketOrderGridOrderColumnWidths.StatusWidth, "Status", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.OrderMode, BracketOrderGridOrderColumnWidths.OrderModeWidth, "Mode", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    setupBracketOrderGridColumn 2, BracketOrderGridOrderColumns.BrokerId, BracketOrderGridOrderColumnWidths.BrokerIdWidth, "Broker Id", True, MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
    
    .MergeCells = flexMergeFree
    .MergeCol(BracketOrderGridColumns.Symbol) = True
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
lIndex = findPositionManagerTableIndex(pPositionManager)

If mPositionManagerGridMappingTable(lIndex).GridIndex <> 0 Then Exit Sub

With mPositionManagerGridMappingTable(lIndex)
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = gGetContractFromContractFuture(pPositionManager.ContractFuture).Specifier
    .GridIndex = addEntryToBracketOrderGrid(lContractSpec.LocalSymbol, True)
    BracketOrderGrid.RowData(.GridIndex) = lIndex + RowDataPositionManagerBase
    BracketOrderGrid.Row = .GridIndex
    BracketOrderGrid.col = 1
    BracketOrderGrid.ColSel = BracketOrderGrid.Cols - 1
    BracketOrderGrid.FillStyle = MSFlexGridLib.FillStyleSettings.flexFillRepeat
    BracketOrderGrid.CellBackColor = &HC0C0C0
    BracketOrderGrid.CellForeColor = vbWhite
    BracketOrderGrid.CellFontBold = True
    BracketOrderGrid.TextMatrix(.GridIndex, BracketOrderGridPositionColumns.Exchange) = lContractSpec.Exchange
    BracketOrderGrid.TextMatrix(.GridIndex, BracketOrderGridPositionColumns.CurrencyCode) = lContractSpec.CurrencyCode
    BracketOrderGrid.TextMatrix(.GridIndex, BracketOrderGridPositionColumns.Size) = pPositionManager.PositionSize
    BracketOrderGrid.col = BracketOrderGridColumns.ExpandIndicator
    BracketOrderGrid.CellPictureAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    Set BracketOrderGrid.CellPicture = BracketOrderImageList.ListImages("Contract").Picture
    .IsExpanded = True
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateBracketOrder()
Const ProcName As String = "updateBracketOrder"
On Error GoTo Err

Dim orderNumber As Long
Dim Price As Double

If Not EditText.Visible Then Exit Sub

orderNumber = mSelectedBracketOrderGridRow - mBracketOrderGridMappingTable(BracketOrderGrid.RowData(BracketOrderGrid.Row) - RowDataBracketOrderBase).GridIndex
If BracketOrderGrid.col = BracketOrderGridOrderColumns.Price Then
    If ParsePrice(EditText.Text, mSelectedBracketOrder.Contract.Specifier.secType, mSelectedBracketOrder.Contract.TickSize, Price) Then
        mSelectedBracketOrder.SetNewOrderPrice orderNumber, Price
    End If
ElseIf BracketOrderGrid.col = BracketOrderGridOrderColumns.AuxPrice Then
    If ParsePrice(EditText.Text, mSelectedBracketOrder.Contract.Specifier.secType, mSelectedBracketOrder.Contract.TickSize, Price) Then
        mSelectedBracketOrder.SetNewOrderTriggerPrice orderNumber, Price
    End If
ElseIf BracketOrderGrid.col = BracketOrderGridOrderColumns.Quantity Then
    If IsNumeric(EditText.Text) Then
        mSelectedBracketOrder.SetNewQuantity EditText.Text
    End If
End If
    
If mSelectedBracketOrder.IsDirty Then mSelectedBracketOrder.Update

endEdit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

