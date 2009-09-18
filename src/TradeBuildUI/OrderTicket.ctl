VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OrderTicket 
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   ScaleHeight     =   6195
   ScaleWidth      =   8790
   Begin VB.CheckBox SimulateOrdersCheck 
      Caption         =   "Simulate orders"
      Height          =   195
      Left            =   3480
      TabIndex        =   73
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ticker"
      Height          =   1815
      Left            =   240
      TabIndex        =   52
      Top             =   3840
      Width           =   3015
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   105
         ScaleHeight     =   1455
         ScaleWidth      =   2655
         TabIndex        =   53
         Top             =   240
         Width           =   2655
         Begin VB.TextBox VolumeText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox HighText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox LowText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox LastSizeText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox AskSizeText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox BidSizeText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox BidText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox LastText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox AskText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Bid"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Ask"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Last"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "Volume"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "High"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Low"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1200
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton UndoButton 
      Caption         =   "&Undo"
      Height          =   495
      Left            =   7560
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order"
      Height          =   2895
      Left            =   240
      TabIndex        =   41
      Top             =   840
      Width           =   3015
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   105
         ScaleHeight     =   2535
         ScaleWidth      =   2895
         TabIndex        =   42
         Top             =   240
         Width           =   2895
         Begin VB.TextBox StopPriceText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   6
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox OffsetText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   5
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox OffsetValueText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox OrderIDText 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   0
            Width           =   2535
         End
         Begin VB.TextBox PriceText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   4
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox TypeCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderTicket.ctx":0000
            Left            =   960
            List            =   "OrderTicket.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox QuantityText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   2
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox ActionCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderTicket.ctx":0004
            Left            =   960
            List            =   "OrderTicket.ctx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Offset (ticks)"
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Id"
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "Stop price"
            Height          =   255
            Left            =   0
            TabIndex        =   49
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Price"
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Type"
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Quantity"
            Height          =   255
            Left            =   0
            TabIndex        =   46
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Action"
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   4815
      Left            =   3360
      TabIndex        =   31
      Top             =   840
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   3735
         TabIndex        =   32
         Top             =   240
         Width           =   3735
         Begin VB.CheckBox IgnoreRthCheck 
            Caption         =   "Ignore RTH"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   8
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox OrderRefText 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   9
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox OverrideCheck 
            Caption         =   "Override"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   23
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox MinQuantityText 
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   15
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox FirmQuoteOnlyCheck 
            Caption         =   "Firm quote only"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   21
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox ETradeOnlyCheck 
            Caption         =   "eTrade only"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   20
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CheckBox AllOrNoneCheck 
            Caption         =   "All or none"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   18
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox GoodTillDateTZText 
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   13
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox GoodAfterTimeTZText 
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   11
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox GoodTillDateText 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   12
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox GoodAfterTimeText 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   10
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox DiscrAmountText 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   16
            Top             =   1800
            Width           =   735
         End
         Begin VB.CheckBox HiddenCheck 
            Caption         =   "Hidden"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   22
            Top             =   3000
            Width           =   855
         End
         Begin VB.ComboBox TriggerMethodCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderTicket.ctx":0008
            Left            =   1200
            List            =   "OrderTicket.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox DisplaySizeText 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   14
            Top             =   1440
            Width           =   735
         End
         Begin VB.CheckBox SweepToFillCheck 
            Caption         =   "SweepToFill"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   24
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CheckBox BlockOrderCheck 
            Caption         =   "Block order"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   19
            Top             =   2520
            Width           =   1095
         End
         Begin VB.ComboBox TIFCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderTicket.ctx":000C
            Left            =   1200
            List            =   "OrderTicket.ctx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Min qty"
            Height          =   375
            Left            =   2040
            TabIndex        =   40
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Good till date"
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Good after time"
            Height          =   255
            Left            =   0
            TabIndex        =   38
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "Discr amount"
            Height          =   255
            Left            =   0
            TabIndex        =   37
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Trigger method"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Display size"
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Order ref"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "TIF"
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.ComboBox OrderSchemeCombo 
      Height          =   315
      ItemData        =   "OrderTicket.ctx":0010
      Left            =   1320
      List            =   "OrderTicket.ctx":0012
      TabIndex        =   0
      Text            =   "Simple order"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton PlaceOrdersButton 
      Caption         =   "&Place orders"
      Height          =   495
      Left            =   7560
      TabIndex        =   25
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton ResetButton 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   7560
      TabIndex        =   28
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CompleteOrdersButton 
      Caption         =   "Complete &order"
      Height          =   495
      Left            =   7560
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton ModifyButton 
      Caption         =   "&Modify"
      Height          =   495
      Left            =   7560
      TabIndex        =   29
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip BracketTabStrip 
      Height          =   5280
      Left            =   120
      TabIndex        =   69
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9313
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Entry"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop loss"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Target"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label SymbolLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   72
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Order scheme"
      Height          =   255
      Left            =   240
      TabIndex        =   71
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label OrderSimulationLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   70
      Top             =   5760
      Width           =   7335
   End
End
Attribute VB_Name = "OrderTicket"
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
Implements QuoteListener

'@================================================================================
' Events
'@================================================================================

Event CaptionChanged(ByVal caption As String)

'@================================================================================
' Constants
'@================================================================================

Private Const NotReadyMessage                   As String = "Not ready for placing orders"

Private Const OrdersLiveMessage                 As String = "Orders are LIVE !!"
Private Const OrdersSimulatedMessage            As String = "Orders are simulated"

'@================================================================================
' Enums
'@================================================================================

Private Enum BracketIndexes
    BracketEntryOrder
    BracketStopOrder
    BracketTargetOrder
End Enum

Private Enum BracketTabs
    TabEntryOrder = 1
    TabStopOrder
    TabTargetOrder
End Enum

Private Enum OrderSchemes
    SimpleOrder
    Bracketorder
    OCAOrder
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTicker                              As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private WithEvents mOrderContext                        As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1

Private mContract                                       As Contract

Private mOrderAction                                    As OrderActions

Private WithEvents mOrderPlex                           As OrderPlex
Attribute mOrderPlex.VB_VarHelpID = -1

Private mCurrentBrackerOrderIndex                       As BracketIndexes

Private mInvalidControls(2)                             As Control

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub UserControl_Initialize()

setupOrderSchemeCombo

loadOrderFields BracketIndexes.BracketStopOrder
loadOrderFields BracketIndexes.BracketTargetOrder

setupActionCombo BracketIndexes.BracketEntryOrder
setupActionCombo BracketIndexes.BracketStopOrder
setupActionCombo BracketIndexes.BracketTargetOrder

End Sub

Private Sub UserControl_Terminate()
Finish
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEvent)
Dim op As OrderPlex

Set op = ev.Source

Select Case ev.changeType
Case OrderPlexChangeTypes.OrderPlexChangesApplied
    ModifyButton.Enabled = False
    UndoButton.Enabled = False
Case OrderPlexChangeTypes.OrderPlexChangesCancelled
    ModifyButton.Enabled = False
    UndoButton.Enabled = False
Case OrderPlexChangeTypes.OrderPlexChangesPending
    ModifyButton.Enabled = True
    UndoButton.Enabled = True
Case OrderPlexChangeTypes.OrderPlexCompleted
    reset
    clearOrderPlex
Case OrderPlexChangeTypes.OrderPlexSelfCancelled
    reset
    clearOrderPlex
Case OrderPlexChangeTypes.OrderPlexEntryOrderChanged
    setOrderFieldValues op.entryOrder, BracketIndexes.BracketEntryOrder
Case OrderPlexChangeTypes.OrderPlexStopOrderChanged
    setOrderFieldValues op.stopOrder, BracketIndexes.BracketStopOrder
Case OrderPlexChangeTypes.OrderPlexTargetOrderChanged
    setOrderFieldValues op.targetOrder, BracketIndexes.BracketTargetOrder
Case OrderPlexChangeTypes.OrderPlexCloseoutOrderCreated
Case OrderPlexChangeTypes.OrderPlexCloseoutOrderChanged
Case OrderPlexChangeTypes.OrderPlexProfitThresholdExceeded
Case OrderPlexChangeTypes.OrderPlexLossThresholdExceeded
Case OrderPlexChangeTypes.OrderPlexDrawdownThresholdExceeded
Case OrderPlexChangeTypes.OrderPlexSizeChanged
Case OrderPlexChangeTypes.OrderPlexStateChanged
End Select
End Sub

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)
AskText = GetFormattedPriceFromQuoteEvent(ev)
AskSizeText = ev.size
setPriceFields
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
BidText = GetFormattedPriceFromQuoteEvent(ev)
BidSizeText = ev.size
setPriceFields
End Sub

Private Sub QuoteListener_high(ev As QuoteEvent)
HighText = GetFormattedPriceFromQuoteEvent(ev)
End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)
LowText = GetFormattedPriceFromQuoteEvent(ev)
End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_previousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_sessionOpen(ev As tradebuild26.QuoteEvent)

End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
LastText = GetFormattedPriceFromQuoteEvent(ev)
LastSizeText = ev.size
setPriceFields
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)
VolumeText = ev.size
End Sub

'@================================================================================
' Form Control Event Handlers
'@================================================================================

Private Sub ActionCombo_Click(ByRef index As Integer)
setAction index
End Sub

Private Sub BracketTabStrip_Click()
mCurrentBrackerOrderIndex = BracketTabStrip.SelectedItem.index - 1
showOrderFields mCurrentBrackerOrderIndex
End Sub

Private Sub CancelButton_Click()
If Not mOrderPlex Is Nothing Then
    mOrderPlex.Cancel True
End If
clearOrderPlex
reset
End Sub

Private Sub CompleteOrdersButton_Click()
'Dim i As Long
'Dim order As order
'
'mOCAOrders.Add mEntryOrder
'For i = 1 To mOCAOrders.Count
'    Set order = mOCAOrders(i)
'    placeOrder order, IIf(i = mOCAOrders.Count, True, False), True
'Next
'
'Set mOCAOrders = Nothing
'
'OrderIDText = ""
'OcaGroupText = ""
'OcaGroupText.Visible = False
'OCAGroupLabel.Visible = False
'OrderSchemeCombo.Enabled = True
'OrderSchemeCombo.ListIndex = SimpleOrder
End Sub

Private Sub ModifyButton_Click()
If Not isValidOrder(BracketEntryOrder) Then Exit Sub
setOrderAttributes mOrderPlex.entryOrder, BracketIndexes.BracketEntryOrder
If Not mOrderPlex.stopOrder Is Nothing Then
    If Not isValidOrder(BracketStopOrder) Then Exit Sub
    setOrderAttributes mOrderPlex.stopOrder, BracketIndexes.BracketStopOrder
End If
If Not mOrderPlex.targetOrder Is Nothing Then
    If Not isValidOrder(BracketTargetOrder) Then Exit Sub
    setOrderAttributes mOrderPlex.targetOrder, BracketIndexes.BracketTargetOrder
End If
mOrderPlex.Update
End Sub

Private Sub OffsetText_Change(index As Integer)
If IsNumeric(OffsetText(index)) Then
    OffsetValueText(index) = OffsetText(index) * mContract.tickSize
Else
    OffsetValueText(index) = ""
End If
setPriceField index
End Sub

Private Sub OrderSchemeCombo_Click()
setOrderScheme comboItemData(OrderSchemeCombo)
End Sub

Private Sub PlaceOrdersButton_Click()
Dim op As OrderPlex

Select Case comboItemData(OrderSchemeCombo)
Case OrderSchemes.SimpleOrder
    If Not isValidOrder(BracketEntryOrder) Then Exit Sub
    
    If comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder)) = OrderActions.ActionBuy Then
        Set op = mOrderContext.CreateBuyOrderPlex( _
                                    QuantityText(BracketIndexes.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketEntryOrder) = "", 0, OffsetText(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketEntryOrder)), _
                                    StopOrderTypes.StopOrderTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    TargetOrderTypes.TargetOrderTypeNone, _
                                    0, _
                                    0, _
                                    0)
    Else
        Set op = mOrderContext.CreateSellOrderPlex( _
                                    QuantityText(BracketIndexes.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketEntryOrder) = "", 0, OffsetText(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketEntryOrder)), _
                                    StopOrderTypes.StopOrderTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    TargetOrderTypes.TargetOrderTypeNone, _
                                    0, _
                                    0, _
                                    0)
        
    End If
    
    setOrderAttributes op.entryOrder, BracketIndexes.BracketEntryOrder
    mOrderContext.executeOrderPlex op
Case OrderSchemes.Bracketorder
    If Not isValidOrder(BracketEntryOrder) Then Exit Sub
    If Not isValidOrder(BracketStopOrder) Then Exit Sub
    If Not isValidOrder(BracketTargetOrder) Then Exit Sub
    
    If comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder)) = OrderActions.ActionBuy Then
        Set op = mOrderContext.CreateBuyOrderPlex( _
                                    QuantityText(BracketIndexes.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketEntryOrder) = "", 0, OffsetText(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketEntryOrder)), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketStopOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketStopOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketStopOrder) = "", 0, OffsetText(BracketIndexes.BracketStopOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketStopOrder)), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketTargetOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketTargetOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketTargetOrder) = "", 0, OffsetText(BracketIndexes.BracketTargetOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketTargetOrder)))
    Else
        Set op = mOrderContext.CreateSellOrderPlex( _
                                    QuantityText(BracketIndexes.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketEntryOrder) = "", 0, OffsetText(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketEntryOrder)), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketStopOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketStopOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketStopOrder) = "", 0, OffsetText(BracketIndexes.BracketStopOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketStopOrder)), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketTargetOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketTargetOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketTargetOrder) = "", 0, OffsetText(BracketIndexes.BracketTargetOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketTargetOrder)))
    End If
    
    setOrderAttributes op.entryOrder, BracketIndexes.BracketEntryOrder
    If Not op.stopOrder Is Nothing Then
        setOrderAttributes op.stopOrder, BracketIndexes.BracketStopOrder
    End If
    If Not op.targetOrder Is Nothing Then
        setOrderAttributes op.targetOrder, BracketIndexes.BracketTargetOrder
    End If
    mOrderContext.executeOrderPlex op
Case OrderSchemes.OCAOrder
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Select

End Sub

Private Sub PriceText_Validate( _
                index As Integer, _
                Cancel As Boolean)
Dim price As Double

' allow blank price to prevent user irritation if they place the caret
' in the price field when the order type is limit, and then decide they
' want to change the order type - if space is not allowed then they
' would have to enter a valid price before being able to get to the order
' type combo
If PriceText(index) = "" Then Exit Sub

If (comboItemData(ActionCombo(index)) = OrderActions.ActionNone And _
        PriceText(index) <> "" _
    ) Or _
    Not mContract.ParsePrice(PriceText(index), price) _
Then
    Cancel = True
    Exit Sub
End If

If Not mOrderPlex Is Nothing Then
    Select Case index
    Case BracketIndexes.BracketEntryOrder
        mOrderPlex.newEntryPrice = price
    Case BracketIndexes.BracketStopOrder
        mOrderPlex.newStopPrice = price
    Case BracketIndexes.BracketTargetOrder
        mOrderPlex.newTargetPrice = price
    End Select
End If
End Sub

Private Sub QuantityText_Validate( _
                index As Integer, _
                Cancel As Boolean)
Dim quantity As Long
Dim max As Long

If comboItemData(ActionCombo(index)) <> OrderActions.ActionNone And _
    Not IsNumeric(QuantityText(index)) _
Then
    Cancel = True
    Exit Sub
End If

Select Case mContract.specifier.secType
Case SecTypeStock
    max = 100000
Case SecTypeFuture
    max = 100
Case SecTypeOption
    max = 100
Case SecTypeFuturesOption
    max = 100
Case SecTypeCash
    max = 10000000
Case SecTypeCombo
    max = 100
Case SecTypeIndex
    max = 0
End Select

If Not IsInteger(QuantityText(index), 1, max) Then
    Cancel = True
    Exit Sub
End If

quantity = CLng(QuantityText(index))

If mOrderPlex Is Nothing Then
    If quantity = 0 Then
        Cancel = True
        Exit Sub
    End If
    
    If comboItemData(OrderSchemeCombo) = OrderSchemes.Bracketorder Then
        Select Case index
        Case BracketIndexes.BracketEntryOrder
            QuantityText(BracketIndexes.BracketStopOrder) = quantity
            QuantityText(BracketIndexes.BracketTargetOrder) = quantity
        Case BracketIndexes.BracketStopOrder
            QuantityText(BracketIndexes.BracketEntryOrder) = quantity
            QuantityText(BracketIndexes.BracketTargetOrder) = quantity
        Case BracketIndexes.BracketTargetOrder
            QuantityText(BracketIndexes.BracketEntryOrder) = quantity
            QuantityText(BracketIndexes.BracketStopOrder) = quantity
        End Select
    End If
    
Else
    mOrderPlex.newQuantity = quantity
End If
End Sub

Private Sub ResetButton_Click()
clearOrderPlex
reset
End Sub

Private Sub SimulateOrdersCheck_Click()
If SimulateOrdersCheck.value = vbUnchecked Then
    Set mOrderContext = mTicker.DefaultOrderContext
Else
    Set mOrderContext = mTicker.DefaultOrderContextSimulated
End If
setupTicker
End Sub

Private Sub StopPriceText_Validate( _
                index As Integer, _
                Cancel As Boolean)
Dim price As Double

If (comboItemData(ActionCombo(index)) = OrderActions.ActionNone And _
        StopPriceText(index) <> "" _
    ) Or _
    Not mTicker.ParsePrice(StopPriceText(index), price) _
Then
    Cancel = True
    Exit Sub
End If

If Not mOrderPlex Is Nothing Then
    Select Case index
    Case BracketIndexes.BracketEntryOrder
        mOrderPlex.newEntryTriggerPrice = price
    Case BracketIndexes.BracketStopOrder
        mOrderPlex.newStopTriggerPrice = price
    Case BracketIndexes.BracketTargetOrder
        mOrderPlex.newTargetTriggerPrice = price
    End Select
End If
End Sub

Private Sub TypeCombo_Click(index As Integer)
configureOrderFields index
setPriceField index
End Sub

Private Sub UndoButton_Click()
mOrderPlex.cancelChanges
End Sub

'@================================================================================
' mOrderContext Event Handlers
'@================================================================================

Private Sub mOrderContext_NotReady()
disableAll NotReadyMessage
End Sub

Private Sub mOrderContext_Ready()
OrderSchemeCombo.Enabled = True
setupTicker
End Sub

'@================================================================================
' mOrderPlex Event Handlers
'@================================================================================

Private Sub mOrderPlex_EntryOrderFilled()
disableOrderFields BracketIndexes.BracketEntryOrder
End Sub

Private Sub mOrderPlex_StopOrderFilled()
disableOrderFields BracketIndexes.BracketStopOrder
End Sub

Private Sub mOrderPlex_TargetOrderFilled()
disableOrderFields BracketIndexes.BracketTargetOrder
End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_StateChange(ev As StateChangeEvent)

Select Case ev.State
Case TickerStateCreated

Case TickerStateStarting

Case TickerStateRunning

Case TickerStatePaused

Case TickerStateClosing

Case TickerStateStopped
    disableAll "Ticker has been stopped"
    Set mOrderContext = Nothing
    Set mTicker = Nothing
End Select
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
UserControl.Enabled = value
PropertyChanged "Enabled"
End Property

Public Property Let Ticker(ByVal value As Ticker)

If value Is mTicker Then Exit Property

If Not mTicker Is Nothing Then mTicker.RemoveQuoteListener Me

Set mTicker = value
If mTicker.OrdersAreLive Then
    Set mOrderContext = mTicker.DefaultOrderContext
Else
    Set mOrderContext = mTicker.DefaultOrderContextSimulated
End If
If mOrderContext.isReady Then
    setupTicker
Else
    disableAll NotReadyMessage
End If

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
On Error GoTo Err
If Not mTicker Is Nothing Then
    mTicker.RemoveQuoteListener Me
    Set mTicker = Nothing
End If
clearOrderPlex
Exit Sub
Err:
'ignore any errors
End Sub

Public Sub showOrderPlex( _
                ByVal value As OrderPlex, _
                ByVal selectedOrderNumber As Long)

Dim entryOrder As Order
Dim stopOrder As Order
Dim targetOrder As Order

clearOrderPlex

Set mOrderPlex = value
Ticker = mOrderPlex.Ticker

SimulateOrdersCheck.Enabled = False     ' can't allow the simulation mode to be changed

Set entryOrder = mOrderPlex.entryOrder
Set stopOrder = mOrderPlex.stopOrder
Set targetOrder = mOrderPlex.targetOrder

If stopOrder Is Nothing And targetOrder Is Nothing Then
    RaiseEvent CaptionChanged("Change a single order")
Else
    RaiseEvent CaptionChanged("Change a bracket order")
End If

OrderSchemeCombo.Enabled = False
BracketTabStrip.Visible = True
If selectedOrderNumber <> 0 Then
    If Not entryOrder Is Nothing Then
        selectedOrderNumber = selectedOrderNumber - 1
        If selectedOrderNumber = 0 Then BracketTabStrip.Tabs(BracketTabs.TabEntryOrder).Selected = True
    End If
    If Not stopOrder Is Nothing Then
        selectedOrderNumber = selectedOrderNumber - 1
        If selectedOrderNumber = 0 Then BracketTabStrip.Tabs(BracketTabs.TabStopOrder).Selected = True
    End If
    If Not targetOrder Is Nothing Then
        selectedOrderNumber = selectedOrderNumber - 1
        If selectedOrderNumber = 0 Then BracketTabStrip.Tabs(BracketTabs.TabTargetOrder).Selected = True
    End If
Else
    If isOrderModifiable(entryOrder) Then
        BracketTabStrip.Tabs(BracketTabs.TabEntryOrder).Selected = True
    ElseIf isOrderModifiable(stopOrder) Then
        BracketTabStrip.Tabs(BracketTabs.TabStopOrder).Selected = True
    ElseIf isOrderModifiable(targetOrder) Then
        BracketTabStrip.Tabs(BracketTabs.TabTargetOrder).Selected = True
    End If
End If

setOrderFieldValues entryOrder, BracketIndexes.BracketEntryOrder
setOrderFieldValues stopOrder, BracketIndexes.BracketStopOrder
setOrderFieldValues targetOrder, BracketIndexes.BracketTargetOrder

ModifyButton.Move PlaceOrdersButton.Left, PlaceOrdersButton.Top
ModifyButton.Visible = True
ModifyButton.Enabled = False

PlaceOrdersButton.Visible = False

CancelButton.Visible = True

UndoButton.Move CompleteOrdersButton.Left, CompleteOrdersButton.Top
UndoButton.Enabled = False
UndoButton.Visible = True

CompleteOrdersButton.Visible = False

ResetButton.Visible = True

mOrderPlex.AddChangeListener Me
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addItemToCombo( _
                ByVal combo As ComboBox, _
                ByVal itemText As String, _
                ByVal itemData As Long)
combo.addItem itemText
combo.itemData(combo.ListCount - 1) = itemData
End Sub

Private Sub clearOrderFields(ByVal index As Long)
enableOrderFields index
OrderIDText(index) = ""
ActionCombo(index).ListIndex = 0
Select Case mContract.specifier.secType
Case SecTypeStock
    QuantityText(index) = 100
Case SecTypeFuture
    QuantityText(index) = 1
Case SecTypeOption
    QuantityText(index) = 1
Case SecTypeFuturesOption
    QuantityText(index) = 1
Case SecTypeCash
    QuantityText(index) = 25000
Case SecTypeCombo
    QuantityText(index) = 1
Case SecTypeIndex
    QuantityText(index) = 0
End Select

' don't set TypeCombo(Index) as it will affect other fields and there
' is no sensible value to set it to
PriceText(index) = ""
StopPriceText(index) = ""
OffsetText(index) = ""
TIFCombo(index).ListIndex = 0
TriggerMethodCombo(index).ListIndex = 0
IgnoreRthCheck(index) = vbUnchecked
OrderRefText(index) = ""
AllOrNoneCheck(index) = vbUnchecked
BlockOrderCheck(index) = vbUnchecked
ETradeOnlyCheck(index) = vbUnchecked
FirmQuoteOnlyCheck(index) = vbUnchecked
HiddenCheck(index) = vbUnchecked
OverrideCheck(index) = vbUnchecked
SweepToFillCheck(index) = vbUnchecked
DisplaySizeText(index) = ""
MinQuantityText(index) = ""
DiscrAmountText(index) = ""
GoodAfterTimeText(index) = ""
GoodAfterTimeTZText(index) = ""
GoodTillDateText(index) = ""
GoodTillDateTZText(index) = ""
End Sub

Private Sub clearOrderPlex()
If Not mOrderPlex Is Nothing Then
    mOrderPlex.RemoveChangeListener Me
    Set mOrderPlex = Nothing
End If
End Sub

Private Function comboItemData(ByVal combo As ComboBox) As Long
comboItemData = combo.itemData(combo.ListIndex)
End Function

Private Sub configureOrderFields( _
                ByVal orderIndex As Long)
Select Case orderIndex
Case BracketIndexes.BracketEntryOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case EntryOrderTypeMarket
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeMarketOnOpen
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeMarketOnClose
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeMarketIfTouched
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeMarketToLimit
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case EntryOrderTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case EntryOrderTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case EntryOrderTypeLimit
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeLimitOnOpen
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeLimitOnClose
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeLimitIfTouched
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeStop
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryOrderTypeStopLimit
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    End Select
Case BracketIndexes.BracketStopOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case StopOrderTypeNone
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case StopOrderTypeStop
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case StopOrderTypeStopLimit
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case StopOrderTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case StopOrderTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case StopOrderTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case StopOrderTypeAuto
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    End Select
Case BracketIndexes.BracketTargetOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case TargetOrderTypeNone
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetOrderTypeLimit
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetOrderTypeLimitIfTouched
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetOrderTypeMarketIfTouched
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetOrderTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case TargetOrderTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case TargetOrderTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case TargetOrderTypeAuto
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    End Select
End Select
End Sub

Private Sub disableAll( _
                ByVal message As String)
OrderSchemeCombo.Enabled = False

PlaceOrdersButton.Enabled = False
CompleteOrdersButton.Enabled = False
ResetButton.Enabled = False
ModifyButton.Enabled = False
UndoButton.Enabled = False

disableOrderFields BracketIndexes.BracketEntryOrder
disableOrderFields BracketIndexes.BracketStopOrder
disableOrderFields BracketIndexes.BracketTargetOrder

SymbolLabel.caption = ""
AskText = ""
AskSizeText = ""
BidText = ""
BidSizeText = ""
LastText = ""
LastSizeText = ""
VolumeText = ""
HighText = ""
LowText = ""

OrderSimulationLabel = message
End Sub

Private Sub disableControl(ByVal field As Control)
field.Enabled = False
If TypeOf field Is CheckBox Or _
    TypeOf field Is OptionButton Then Exit Sub
    
field.backColor = SystemColorConstants.vbButtonFace
End Sub

Private Sub disableOrderFields(ByVal index As Long)
disableControl ActionCombo(index)
disableControl QuantityText(index)
disableControl TypeCombo(index)
disableControl PriceText(index)
disableControl OffsetText(index)
disableControl StopPriceText(index)
disableControl IgnoreRthCheck(index)
disableControl TIFCombo(index)
disableControl OrderRefText(index)
disableControl AllOrNoneCheck(index)
disableControl BlockOrderCheck(index)
disableControl ETradeOnlyCheck(index)
disableControl FirmQuoteOnlyCheck(index)
disableControl HiddenCheck(index)
disableControl OverrideCheck(index)
disableControl SweepToFillCheck(index)
disableControl DisplaySizeText(index)
disableControl MinQuantityText(index)
disableControl TriggerMethodCombo(index)
disableControl DiscrAmountText(index)
disableControl GoodAfterTimeText(index)
disableControl GoodAfterTimeTZText(index)
disableControl GoodTillDateText(index)
disableControl GoodTillDateTZText(index)
End Sub

Private Sub enableControl(ByVal field As Control)
field.Enabled = True
If TypeOf field Is CheckBox Or _
    TypeOf field Is OptionButton Then Exit Sub
    
field.backColor = SystemColorConstants.vbWindowBackground
End Sub

Private Sub enableOrderFields(ByVal index As Long)
enableControl ActionCombo(index)
enableControl QuantityText(index)
enableControl TypeCombo(index)
enableControl PriceText(index)
enableControl OffsetText(index)
enableControl StopPriceText(index)
enableControl IgnoreRthCheck(index)
enableControl TIFCombo(index)
enableControl OrderRefText(index)
enableControl AllOrNoneCheck(index)
enableControl BlockOrderCheck(index)
enableControl ETradeOnlyCheck(index)
enableControl FirmQuoteOnlyCheck(index)
enableControl HiddenCheck(index)
enableControl OverrideCheck(index)
enableControl SweepToFillCheck(index)
enableControl DisplaySizeText(index)
enableControl MinQuantityText(index)
enableControl TriggerMethodCombo(index)
enableControl DiscrAmountText(index)
enableControl GoodAfterTimeText(index)
enableControl GoodAfterTimeTZText(index)
enableControl GoodTillDateText(index)
enableControl GoodTillDateTZText(index)
End Sub

Private Function getPrice( _
                ByVal priceString As String) As Double
Dim price As Double
mContract.ParsePrice priceString, price
getPrice = price
End Function

Private Function isOrderModifiable(ByVal pOrder As Order) As Boolean
If pOrder Is Nothing Then Exit Function
isOrderModifiable = pOrder.isModifiable
End Function

Private Function isPrice( _
                ByVal priceString As String) As Boolean
Dim price As Double
isPrice = mContract.ParsePrice(priceString, price)
End Function

Private Function isValidOrder( _
                ByVal index As Long) As Boolean

If Not mInvalidControls(index) Is Nothing Then mInvalidControls(index).backColor = vbButtonFace

If comboItemData(ActionCombo(index)) = OrderActions.ActionNone Then
    isValidOrder = True
    Exit Function
End If

Select Case index
Case BracketEntryOrder
    If Not IsInteger(QuantityText(index), 0) Then setInvalidControl QuantityText(index), index: Exit Function
    If QuantityText(index) = 0 And mOrderPlex Is Nothing Then setInvalidControl QuantityText(index), index: Exit Function
    
    Select Case comboItemData(TypeCombo(index))
    Case EntryOrderTypeMarket, EntryOrderTypeMarketOnOpen, EntryOrderTypeMarketOnClose
        ' other field values don't matter
    Case EntryOrderTypeMarketIfTouched, EntryOrderTypeStop
        If Not isPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
    Case EntryOrderTypeMarketToLimit, EntryOrderTypeLimit, EntryOrderTypeLimitOnOpen, EntryOrderTypeLimitOnClose
        If Not isPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case EntryOrderTypeBid, EntryOrderTypeAsk, EntryOrderTypeLast
        If OffsetText(index) <> "" Then
            If Not IsInteger(OffsetText(index), -100, 100) Then setInvalidControl OffsetText(index), index: Exit Function
        End If
    Case EntryOrderTypeLimitIfTouched, EntryOrderTypeStopLimit
        If Not isPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
        If Not isPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    End Select
Case BracketStopOrder
    If comboItemData(TypeCombo(index)) = StopOrderTypeNone Then
        isValidOrder = True
        Exit Function
    End If
    
    If Not IsInteger(QuantityText(index), 1) Then setInvalidControl QuantityText(index), index: Exit Function
    
    Select Case comboItemData(TypeCombo(index))
    Case StopOrderTypeStop
        If Not isPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
    Case StopOrderTypeStopLimit
        If Not isPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
        If Not isPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case StopOrderTypeBid, StopOrderTypeAsk, StopOrderTypeLast, StopOrderTypeAuto
        If OffsetText(index) <> "" Then
            If Not IsInteger(OffsetText(index), -100, 100) Then setInvalidControl OffsetText(index), index: Exit Function
        End If
    End Select
Case BracketTargetOrder
    If comboItemData(TypeCombo(index)) = TargetOrderTypeNone Then
        isValidOrder = True
        Exit Function
    End If
    
    If Not IsInteger(QuantityText(index), 1) Then setInvalidControl QuantityText(index), index: Exit Function
    
    Select Case comboItemData(TypeCombo(index))
    Case TargetOrderTypeLimit
        If Not isPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case TargetOrderTypeLimitIfTouched
        If Not isPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
        If Not isPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case TargetOrderTypeMarketIfTouched
        If Not isPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
    Case TargetOrderTypeBid, TargetOrderTypeAsk, TargetOrderTypeLast, TargetOrderTypeAuto
        If OffsetText(index) <> "" Then
            If Not IsInteger(OffsetText(index), -100, 100) Then setInvalidControl OffsetText(index), index: Exit Function
        End If
    End Select
End Select

If DisplaySizeText(index) <> "" Then
    If Not IsInteger(DisplaySizeText(index), 1) Then setInvalidControl DisplaySizeText(index), index: Exit Function
End If

If MinQuantityText(index) <> "" Then
    If Not IsInteger(MinQuantityText(index), 1) Then setInvalidControl MinQuantityText(index), index: Exit Function
End If

If DiscrAmountText(index) <> "" Then
    If Not isPrice(DiscrAmountText(index)) Then setInvalidControl DiscrAmountText(index), index: Exit Function
End If

isValidOrder = True
End Function

Private Sub loadOrderFields(ByVal index As Long)
load OrderIDText(index)
load ActionCombo(index)
load QuantityText(index)
load TypeCombo(index)
load PriceText(index)
load StopPriceText(index)
load IgnoreRthCheck(index)
load OffsetText(index)
load OffsetValueText(index)
load TIFCombo(index)
load OrderRefText(index)
load AllOrNoneCheck(index)
load BlockOrderCheck(index)
load ETradeOnlyCheck(index)
load FirmQuoteOnlyCheck(index)
load HiddenCheck(index)
load OverrideCheck(index)
load SweepToFillCheck(index)
load DisplaySizeText(index)
load MinQuantityText(index)
load TriggerMethodCombo(index)
load DiscrAmountText(index)
load GoodAfterTimeText(index)
load GoodAfterTimeTZText(index)
load GoodTillDateText(index)
load GoodTillDateTZText(index)
End Sub

Private Sub reset()
clearOrderFields BracketIndexes.BracketEntryOrder
clearOrderFields BracketIndexes.BracketStopOrder
clearOrderFields BracketIndexes.BracketTargetOrder

If mTicker.OrdersAreLive Then
    SimulateOrdersCheck.Enabled = True
Else
    SimulateOrdersCheck.value = vbChecked
    SimulateOrdersCheck.Enabled = False
End If

OrderSchemeCombo.Enabled = True
selectComboEntry OrderSchemeCombo, OrderSchemes.Bracketorder
setOrderScheme OrderSchemes.Bracketorder

selectComboEntry ActionCombo(BracketIndexes.BracketEntryOrder), _
                OrderActions.ActionBuy
setAction BracketIndexes.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndexes.BracketEntryOrder), _
                EntryOrderTypes.EntryOrderTypeLimit
setOrderFieldsEnabling BracketIndexes.BracketEntryOrder, Nothing
configureOrderFields BracketIndexes.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndexes.BracketStopOrder), _
                StopOrderTypes.StopOrderTypeStop
setOrderFieldsEnabling BracketIndexes.BracketStopOrder, Nothing
configureOrderFields BracketIndexes.BracketStopOrder

selectComboEntry TypeCombo(BracketIndexes.BracketTargetOrder), _
                TargetOrderTypes.TargetOrderTypeNone
setOrderFieldsEnabling BracketIndexes.BracketTargetOrder, Nothing
configureOrderFields BracketIndexes.BracketTargetOrder

BracketTabStrip.Tabs(BracketTabs.TabEntryOrder).Selected = True

End Sub

Private Sub selectComboEntry( _
                ByVal combo As ComboBox, _
                ByVal itemData As Long)
Dim i As Long

For i = 0 To combo.ListCount - 1
    If combo.itemData(i) = itemData Then
        combo.ListIndex = i
        Exit For
    End If
Next
End Sub

Private Sub setAction( _
                ByVal index As Long)
mOrderAction = comboItemData(ActionCombo(index))
If comboItemData(OrderSchemeCombo) = OrderSchemes.Bracketorder And _
    index = BracketIndexes.BracketEntryOrder _
Then
    If comboItemData(ActionCombo(index)) = OrderActions.ActionSell Then
        selectComboEntry ActionCombo(BracketIndexes.BracketStopOrder), OrderActions.ActionBuy
        selectComboEntry ActionCombo(BracketIndexes.BracketTargetOrder), OrderActions.ActionBuy
    Else
        selectComboEntry ActionCombo(BracketIndexes.BracketStopOrder), OrderActions.ActionSell
        selectComboEntry ActionCombo(BracketIndexes.BracketTargetOrder), OrderActions.ActionSell
    End If
    disableControl ActionCombo(BracketIndexes.BracketStopOrder)
    disableControl ActionCombo(BracketIndexes.BracketTargetOrder)
End If
End Sub

Private Sub setInvalidControl( _
                ByVal pControl As Control, _
                ByVal index As Long)
Set mInvalidControls(index) = pControl
If BracketTabStrip.Visible Then BracketTabStrip.Tabs(index + 1).Selected = True
pControl.backColor = ErroredFieldColor
End Sub

'/**
' Sets the attributes for an order from the specified fields on the control
'
' @param pOrder     the <code>order</code> whose attributes are to be set
' @param orderIndex the index of the order page whose fields are the source of
'                   the attribute values
'
'*/
Private Sub setOrderAttributes( _
                ByVal pOrder As Order, _
                ByVal orderIndex As Long)

With pOrder
    If pOrder.isAttributeModifiable(OrderAttAllOrNone) Then .allOrNone = (AllOrNoneCheck(orderIndex) = vbChecked)
    If pOrder.isAttributeModifiable(OrderAttBlockOrder) Then .blockOrder = (BlockOrderCheck(orderIndex) = vbChecked)
    If pOrder.isAttributeModifiable(OrderAttDiscretionaryAmount) Then .discretionaryAmount = IIf(DiscrAmountText(orderIndex) = "", 0, DiscrAmountText(orderIndex))
    If pOrder.isAttributeModifiable(OrderAttDisplaySize) Then .displaySize = IIf(DisplaySizeText(orderIndex) = "", 0, DisplaySizeText(orderIndex))
    If pOrder.isAttributeModifiable(OrderAttETradeOnly) Then .eTradeOnly = (ETradeOnlyCheck(orderIndex) = vbChecked)
    If pOrder.isAttributeModifiable(OrderAttFirmQuoteOnly) Then .firmQuoteOnly = (FirmQuoteOnlyCheck(orderIndex) = vbChecked)
    If pOrder.isAttributeModifiable(OrderAttGoodAfterTime) Then .goodAfterTime = IIf(GoodAfterTimeText(orderIndex) = "", 0, GoodAfterTimeText(orderIndex))
    If pOrder.isAttributeModifiable(OrderAttGoodAfterTimeTZ) Then .goodAfterTimeTZ = GoodAfterTimeTZText(orderIndex)
    If pOrder.isAttributeModifiable(OrderAttGoodTillDate) Then .goodTillDate = IIf(GoodTillDateText(orderIndex) = "", 0, GoodTillDateText(orderIndex))
    If pOrder.isAttributeModifiable(OrderAttGoodTillDateTZ) Then .goodTillDateTZ = GoodTillDateTZText(orderIndex)
    If pOrder.isAttributeModifiable(OrderAttHidden) Then .Hidden = (HiddenCheck(orderIndex) = vbChecked)
    If pOrder.isAttributeModifiable(OrderAttIgnoreRTH) Then .ignoreRegularTradingHours = (IgnoreRthCheck(orderIndex) = vbChecked)
    'If pOrder.isAttributeModifiable(OrderAttLimitPrice) Then .limitPrice = IIf(PriceText(orderIndex) = "", 0, PriceText(orderIndex))
    If pOrder.isAttributeModifiable(OrderAttMinimumQuantity) Then .minimumQuantity = IIf(MinQuantityText(orderIndex) = "", 0, MinQuantityText(orderIndex))
    'If pOrder.isAttributeModifiable(OrderAttOrderType) Then .orderType = comboItemData(TypeCombo(orderIndex))
    If pOrder.isAttributeModifiable(OrderAttOriginatorRef) Then .originatorRef = OrderRefText(orderIndex)
    If pOrder.isAttributeModifiable(OrderAttOverrideConstraints) Then .overrideConstraints = (OverrideCheck(orderIndex) = vbChecked)
    If pOrder.isAttributeModifiable(OrderAttQuantity) Then .quantity = QuantityText(orderIndex)
    If pOrder.isAttributeModifiable(OrderAttStopTriggerMethod) Then .StopTriggerMethod = comboItemData(TriggerMethodCombo(orderIndex))
    If pOrder.isAttributeModifiable(OrderAttSweepToFill) Then .SweepToFill = (SweepToFillCheck(orderIndex) = vbChecked)
    If pOrder.isAttributeModifiable(OrderAttTimeInForce) Then .timeInForce = comboItemData(TIFCombo(orderIndex))
    'If pOrder.isAttributeModifiable(OrderAttTriggerPrice) Then .triggerPrice = IIf(StopPriceText(orderIndex) = "", 0, StopPriceText(orderIndex))
End With
End Sub

Private Sub setOrderFieldValues( _
                ByVal pOrder As Order, _
                ByVal orderIndex As Long)
If pOrder Is Nothing Then
    disableOrderFields orderIndex
    Exit Sub
End If

clearOrderFields orderIndex

With pOrder
    setOrderId orderIndex, .id
    
    ActionCombo(orderIndex).Text = OrderActionToString(.Action)
    QuantityText(orderIndex) = .quantity
    TypeCombo(orderIndex).Text = OrderTypeToString(.orderType)
    PriceText(orderIndex) = IIf(.limitPrice <> 0, .limitPrice, "")
    StopPriceText(orderIndex) = IIf(.triggerPrice <> 0, .triggerPrice, "")
    IgnoreRthCheck(orderIndex) = IIf(.ignoreRegularTradingHours, vbChecked, vbUnchecked)
    TIFCombo(orderIndex) = OrderTIFToString(.timeInForce)
    OrderRefText(orderIndex) = .originatorRef
    AllOrNoneCheck(orderIndex) = IIf(.allOrNone, vbChecked, vbUnchecked)
    BlockOrderCheck(orderIndex) = IIf(.blockOrder, vbChecked, vbUnchecked)
    ETradeOnlyCheck(orderIndex) = IIf(.eTradeOnly, vbChecked, vbUnchecked)
    FirmQuoteOnlyCheck(orderIndex) = IIf(.firmQuoteOnly, vbChecked, vbUnchecked)
    HiddenCheck(orderIndex) = IIf(.Hidden, vbChecked, vbUnchecked)
    OverrideCheck(orderIndex) = IIf(.overrideConstraints, vbChecked, vbUnchecked)
    SweepToFillCheck(orderIndex) = IIf(.SweepToFill, vbChecked, vbUnchecked)
    DisplaySizeText(orderIndex) = IIf(.displaySize <> 0, .displaySize, "")
    MinQuantityText(orderIndex) = IIf(.minimumQuantity <> 0, .displaySize, "")
    If .StopTriggerMethod <> 0 Then TriggerMethodCombo(orderIndex) = OrderStopTriggerMethodToString(.StopTriggerMethod)
    DiscrAmountText(orderIndex) = IIf(.discretionaryAmount <> 0, .discretionaryAmount, "")
    GoodAfterTimeText(orderIndex) = IIf(.goodAfterTime <> 0, FormatDateTime(.goodAfterTime, vbGeneralDate), "")
    GoodAfterTimeTZText(orderIndex) = .goodAfterTimeTZ
    GoodTillDateText(orderIndex) = IIf(.goodTillDate <> 0, FormatDateTime(.goodTillDate, vbGeneralDate), "")
    GoodTillDateTZText(orderIndex) = .goodTillDateTZ
    
    ' do this last because it sets the various fields attributes
    TypeCombo(orderIndex).Text = OrderTypeToString(.orderType)
End With

If Not isOrderModifiable(pOrder) Then
    disableOrderFields orderIndex
Else
    setOrderFieldsEnabling orderIndex, pOrder
End If
End Sub

Private Sub setOrderFieldEnabling( _
                ByVal pControl As Control, _
                ByVal orderAtt As OrderAttributeIds, _
                ByVal pOrder As Order)
If Not pOrder Is Nothing Then
    If pOrder.isAttributeModifiable(orderAtt) Then
        enableControl pControl
    Else
        disableControl pControl
    End If
ElseIf mOrderContext.isAttributeSupported(orderAtt) Then
    enableControl pControl
Else
    disableControl pControl
End If
End Sub

Private Sub setOrderFieldsEnabling( _
                ByVal index As Long, _
                ByVal pOrder As Order)
setOrderFieldEnabling ActionCombo(index), OrderAttAction, pOrder
setOrderFieldEnabling QuantityText(index), OrderAttQuantity, pOrder
setOrderFieldEnabling TypeCombo(index), OrderAttOrderType, pOrder
setOrderFieldEnabling PriceText(index), OrderAttLimitPrice, pOrder
setOrderFieldEnabling StopPriceText(index), OrderAttTriggerPrice, pOrder
setOrderFieldEnabling IgnoreRthCheck(index), OrderAttIgnoreRTH, pOrder
setOrderFieldEnabling TIFCombo(index), OrderAttTimeInForce, pOrder
setOrderFieldEnabling OrderRefText(index), OrderAttOriginatorRef, pOrder
setOrderFieldEnabling AllOrNoneCheck(index), OrderAttAllOrNone, pOrder
setOrderFieldEnabling BlockOrderCheck(index), OrderAttBlockOrder, pOrder
setOrderFieldEnabling ETradeOnlyCheck(index), OrderAttETradeOnly, pOrder
setOrderFieldEnabling FirmQuoteOnlyCheck(index), OrderAttFirmQuoteOnly, pOrder
setOrderFieldEnabling HiddenCheck(index), OrderAttHidden, pOrder
setOrderFieldEnabling OverrideCheck(index), OrderAttOverrideConstraints, pOrder
setOrderFieldEnabling SweepToFillCheck(index), OrderAttSweepToFill, pOrder
setOrderFieldEnabling DisplaySizeText(index), OrderAttDisplaySize, pOrder
setOrderFieldEnabling MinQuantityText(index), OrderAttMinimumQuantity, pOrder
setOrderFieldEnabling TriggerMethodCombo(index), OrderAttStopTriggerMethod, pOrder
setOrderFieldEnabling DiscrAmountText(index), OrderAttDiscretionaryAmount, pOrder
setOrderFieldEnabling GoodAfterTimeText(index), OrderAttGoodAfterTime, pOrder
setOrderFieldEnabling GoodAfterTimeTZText(index), OrderAttGoodAfterTimeTZ, pOrder
setOrderFieldEnabling GoodTillDateText(index), OrderAttGoodTillDate, pOrder
setOrderFieldEnabling GoodTillDateTZText(index), OrderAttGoodTillDateTZ, pOrder
End Sub

Private Sub setOrderId( _
                ByVal index As Long, _
                ByVal id As String)
enableControl OrderIDText(index)
OrderIDText(index) = id
disableControl OrderIDText(index)
End Sub

Private Sub setOrderScheme( _
                ByVal pOrderScheme As OrderSchemes)
Select Case pOrderScheme
Case OrderSchemes.SimpleOrder
    RaiseEvent CaptionChanged("Create a simple order")
    BracketTabStrip.Visible = False
    PlaceOrdersButton.Enabled = True
    PlaceOrdersButton.Visible = True
    CompleteOrdersButton.Visible = False
    ModifyButton.Visible = False
    UndoButton.Visible = False
    ResetButton.Enabled = True
    ResetButton.Enabled = True
    showOrderFields BracketIndexes.BracketEntryOrder
    
Case OrderSchemes.Bracketorder
    RaiseEvent CaptionChanged("Create a bracket order")
    BracketTabStrip.Visible = True
    PlaceOrdersButton.Enabled = True
    PlaceOrdersButton.Visible = True
    CompleteOrdersButton.Visible = False
    ModifyButton.Visible = False
    UndoButton.Visible = False
    ResetButton.Enabled = True
    ResetButton.Enabled = True
    BracketTabStrip.Tabs(BracketTabs.TabEntryOrder).Selected = True
Case OrderSchemes.OCAOrder
    Dim OCAId As Long
    RaiseEvent CaptionChanged("Create a 'one cancels all' group")
    BracketTabStrip.Visible = False
'    If mOCAOrders Is Nothing Then Set mOCAOrders = New Collection
    PlaceOrdersButton.Visible = True
    CompleteOrdersButton.Visible = True
    ModifyButton.Visible = False
    UndoButton.Visible = False
End Select
End Sub

Private Sub setPriceField( _
                index As Integer)
Dim basePrice As Double
Dim offset As Double

Select Case index
Case BracketIndexes.BracketEntryOrder
    Select Case comboItemData(TypeCombo(index))
    Case EntryOrderTypeBid
        basePrice = mTicker.BidPrice
    Case EntryOrderTypeAsk
        basePrice = mTicker.AskPrice
    Case EntryOrderTypeLast
        basePrice = mTicker.TradePrice
    Case Else
        Exit Sub
    End Select
Case BracketIndexes.BracketStopOrder
    Select Case comboItemData(TypeCombo(index))
    Case StopOrderTypeBid
        basePrice = mTicker.BidPrice
    Case StopOrderTypeAsk
        basePrice = mTicker.AskPrice
    Case StopOrderTypeLast
        basePrice = mTicker.TradePrice
    Case StopOrderTypeAuto
        basePrice = 0
    Case Else
        Exit Sub
    End Select
Case BracketIndexes.BracketTargetOrder
    Select Case comboItemData(TypeCombo(index))
    Case TargetOrderTypeBid
        basePrice = mTicker.BidPrice
    Case TargetOrderTypeAsk
        basePrice = mTicker.AskPrice
    Case TargetOrderTypeLast
        basePrice = mTicker.TradePrice
    Case TargetOrderTypeAuto
        basePrice = 0
    Case Else
        Exit Sub
    End Select
End Select

If IsNumeric(OffsetText(index)) Then
    offset = OffsetText(index) * mContract.tickSize
End If

PriceText(index) = mTicker.FormatPrice(basePrice + offset)
End Sub

Private Sub setPriceFields()
setPriceField BracketIndexes.BracketEntryOrder
setPriceField BracketIndexes.BracketStopOrder
setPriceField BracketIndexes.BracketTargetOrder
End Sub

Private Sub setupActionCombo(ByVal index As Long)
ActionCombo(index).Clear
If index <> BracketIndexes.BracketEntryOrder Then
    addItemToCombo ActionCombo(index), _
                OrderActionToString(OrderActions.ActionNone), _
                OrderActions.ActionNone
End If
addItemToCombo ActionCombo(index), _
            OrderActionToString(OrderActions.ActionBuy), _
            OrderActions.ActionBuy
addItemToCombo ActionCombo(index), _
            OrderActionToString(OrderActions.ActionSell), _
            OrderActions.ActionSell
End Sub

Private Sub setupOrderSchemeCombo()
OrderSchemeCombo.Clear
addItemToCombo OrderSchemeCombo, _
            "Bracket order", _
            OrderSchemes.Bracketorder
addItemToCombo OrderSchemeCombo, _
            "Simple order", _
            OrderSchemes.SimpleOrder
'addItemToCombo OrderSchemeCombo, _
'            "OCA order", _
'            OrderSchemes.OCAOrder
'addItemToCombo OrderSchemeCombo, _
'            "Combination order", _
'            OrderSchemes.CombinationOrder
OrderSchemeCombo.ListIndex = 0
End Sub

Private Sub setupTicker()
Set mContract = mTicker.Contract

SymbolLabel.caption = mContract.specifier.localSymbol & _
                        " on " & _
                        mContract.specifier.exchange
                        
setupTifCombo BracketIndexes.BracketEntryOrder
setupTifCombo BracketIndexes.BracketStopOrder
setupTifCombo BracketIndexes.BracketTargetOrder

setupTriggerMethodCombo BracketIndexes.BracketEntryOrder
setupTriggerMethodCombo BracketIndexes.BracketStopOrder
setupTriggerMethodCombo BracketIndexes.BracketTargetOrder

setupTypeCombo BracketIndexes.BracketEntryOrder
setupTypeCombo BracketIndexes.BracketStopOrder
setupTypeCombo BracketIndexes.BracketTargetOrder

reset

mTicker.RemoveQuoteListener Me
mTicker.AddQuoteListener Me
showTickerValues

If mOrderContext.IsSimulated Then
    OrderSimulationLabel.caption = OrdersSimulatedMessage
Else
    OrderSimulationLabel.caption = OrdersLiveMessage
End If

End Sub

Private Sub setupTifCombo(ByVal index As Long)
Dim permittedTifs As Long

permittedTifs = mOrderContext.permittedOrderTifs

TIFCombo(index).Clear

If permittedTifs And OrderTifs.TIFDay Then
    addItemToCombo TIFCombo(index), _
                OrderTIFToString(OrderTifs.TIFDay), _
                OrderTifs.TIFDay
End If
If permittedTifs And OrderTifs.TIFGoodTillCancelled Then
    addItemToCombo TIFCombo(index), _
                OrderTIFToString(OrderTifs.TIFGoodTillCancelled), _
                OrderTifs.TIFGoodTillCancelled
End If
If permittedTifs And OrderTifs.TIFImmediateOrCancel Then
    addItemToCombo TIFCombo(index), _
                OrderTIFToString(OrderTifs.TIFImmediateOrCancel), _
                OrderTifs.TIFImmediateOrCancel
End If

TIFCombo(0).ListIndex = 0
End Sub

Private Sub setupTriggerMethodCombo(ByVal index As Long)
Dim permittedTriggers As Long

permittedTriggers = mOrderContext.permittedStopTriggerMethods

TriggerMethodCombo(index).Clear

If permittedTriggers And StopTriggerMethods.StopTriggerDefault Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(StopTriggerMethods.StopTriggerDefault), _
                StopTriggerMethods.StopTriggerDefault
End If
If permittedTriggers And StopTriggerMethods.StopTriggerLast Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(StopTriggerMethods.StopTriggerLast), _
                StopTriggerMethods.StopTriggerLast
End If
If permittedTriggers And StopTriggerMethods.StopTriggerBidAsk Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(StopTriggerMethods.StopTriggerBidAsk), _
                StopTriggerMethods.StopTriggerBidAsk
End If
If permittedTriggers And StopTriggerMethods.StopTriggerDoubleBidAsk Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(StopTriggerMethods.StopTriggerDoubleBidAsk), _
                StopTriggerMethods.StopTriggerDoubleBidAsk
End If
If permittedTriggers And StopTriggerMethods.StopTriggerDoubleLast Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(StopTriggerMethods.StopTriggerDoubleLast), _
                StopTriggerMethods.StopTriggerDoubleLast
End If
If permittedTriggers And StopTriggerMethods.StopTriggerLastOrBidAsk Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(StopTriggerMethods.StopTriggerLastOrBidAsk), _
                StopTriggerMethods.StopTriggerLastOrBidAsk
End If
If permittedTriggers And StopTriggerMethods.StopTriggerMidPoint Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(StopTriggerMethods.StopTriggerMidPoint), _
                StopTriggerMethods.StopTriggerMidPoint
End If

TriggerMethodCombo(index).ListIndex = 0
End Sub

Private Sub setupTypeCombo(ByVal index As Long)
Dim permittedOrderTypes As Long

permittedOrderTypes = mOrderContext.permittedOrderTypes

TypeCombo(index).Clear

If index = BracketIndexes.BracketEntryOrder Then
    If permittedOrderTypes And OrderTypes.OrderTypeLimit Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeLimit), _
                    EntryOrderTypes.EntryOrderTypeLimit
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeMarket Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeMarket), _
                    EntryOrderTypes.EntryOrderTypeMarket
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeStop Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeStop), _
                    EntryOrderTypes.EntryOrderTypeStop
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeStopLimit Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeStopLimit), _
                    EntryOrderTypes.EntryOrderTypeStopLimit
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeLimit Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeBid), _
                    EntryOrderTypes.EntryOrderTypeBid
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeAsk), _
                    EntryOrderTypes.EntryOrderTypeAsk
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeLast), _
                    EntryOrderTypes.EntryOrderTypeLast
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeLimitOnOpen Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeLimitOnOpen), _
                    EntryOrderTypes.EntryOrderTypeLimitOnOpen
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeMarketOnOpen Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeMarketOnOpen), _
                    EntryOrderTypes.EntryOrderTypeMarketOnOpen
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeLimitOnClose Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeLimitOnClose), _
                    EntryOrderTypes.EntryOrderTypeLimitOnClose
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeMarketOnClose Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeMarketOnClose), _
                    EntryOrderTypes.EntryOrderTypeMarketOnClose
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeLimitIfTouched Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeLimitIfTouched), _
                    EntryOrderTypes.EntryOrderTypeLimitIfTouched
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeMarketIfTouched Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeMarketIfTouched), _
                    EntryOrderTypes.EntryOrderTypeMarketIfTouched
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeMarketToLimit Then
        addItemToCombo TypeCombo(index), _
                    EntryOrderTypeToString(EntryOrderTypes.EntryOrderTypeMarketToLimit), _
                    EntryOrderTypes.EntryOrderTypeMarketToLimit
    End If
ElseIf index = BracketIndexes.BracketStopOrder Then
    addItemToCombo TypeCombo(index), _
                StopOrderTypeToString(StopOrderTypes.StopOrderTypeNone), _
                StopOrderTypes.StopOrderTypeNone
    If permittedOrderTypes And OrderTypes.OrderTypeStop Then
        addItemToCombo TypeCombo(index), _
                    StopOrderTypeToString(StopOrderTypes.StopOrderTypeStop), _
                    StopOrderTypes.StopOrderTypeStop
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeStopLimit Then
        addItemToCombo TypeCombo(index), _
                    StopOrderTypeToString(StopOrderTypes.StopOrderTypeStopLimit), _
                    StopOrderTypes.StopOrderTypeStopLimit
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeLimit Then
        addItemToCombo TypeCombo(index), _
                    StopOrderTypeToString(StopOrderTypes.StopOrderTypeBid), _
                    StopOrderTypes.StopOrderTypeBid
        addItemToCombo TypeCombo(index), _
                    StopOrderTypeToString(StopOrderTypes.StopOrderTypeAsk), _
                    StopOrderTypes.StopOrderTypeAsk
        addItemToCombo TypeCombo(index), _
                    StopOrderTypeToString(StopOrderTypes.StopOrderTypeLast), _
                    StopOrderTypes.StopOrderTypeLast
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeStop Then
        addItemToCombo TypeCombo(index), _
                    StopOrderTypeToString(StopOrderTypes.StopOrderTypeAuto), _
                    StopOrderTypes.StopOrderTypeAuto
    End If
ElseIf index = BracketIndexes.BracketTargetOrder Then
    addItemToCombo TypeCombo(index), _
                TargetOrderTypeToString(TargetOrderTypes.TargetOrderTypeNone), _
                TargetOrderTypes.TargetOrderTypeNone
    If permittedOrderTypes And OrderTypes.OrderTypeLimit Then
        addItemToCombo TypeCombo(index), _
                    TargetOrderTypeToString(TargetOrderTypes.TargetOrderTypeLimit), _
                    TargetOrderTypes.TargetOrderTypeLimit
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeMarketIfTouched Then
        addItemToCombo TypeCombo(index), _
                    TargetOrderTypeToString(TargetOrderTypes.TargetOrderTypeMarketIfTouched), _
                    TargetOrderTypes.TargetOrderTypeMarketIfTouched
    End If
    If permittedOrderTypes And OrderTypes.OrderTypeLimit Then
        addItemToCombo TypeCombo(index), _
                    TargetOrderTypeToString(TargetOrderTypes.TargetOrderTypeBid), _
                    TargetOrderTypes.TargetOrderTypeBid
        addItemToCombo TypeCombo(index), _
                    TargetOrderTypeToString(TargetOrderTypes.TargetOrderTypeAsk), _
                    TargetOrderTypes.TargetOrderTypeAsk
        addItemToCombo TypeCombo(index), _
                    TargetOrderTypeToString(TargetOrderTypes.TargetOrderTypeLast), _
                    TargetOrderTypes.TargetOrderTypeLast
        addItemToCombo TypeCombo(index), _
                    TargetOrderTypeToString(TargetOrderTypes.TargetOrderTypeAuto), _
                    TargetOrderTypes.TargetOrderTypeAuto
    End If
End If

TypeCombo(index).ListIndex = 0
End Sub

Private Sub showOrderFields(ByVal index As Long)
Dim i As Long
For i = 0 To ActionCombo.Count - 1
    If i = index Then
        OrderIDText(i).Visible = True
        ActionCombo(i).Visible = True
        QuantityText(i).Visible = True
        TypeCombo(i).Visible = True
        PriceText(i).Visible = True
        OffsetText(i).Visible = True
        OffsetValueText(i).Visible = True
        StopPriceText(i).Visible = True
        IgnoreRthCheck(i).Visible = True
        TIFCombo(i).Visible = True
        OrderRefText(i).Visible = True
        AllOrNoneCheck(i).Visible = True
        BlockOrderCheck(i).Visible = True
        ETradeOnlyCheck(i).Visible = True
        FirmQuoteOnlyCheck(i).Visible = True
        HiddenCheck(i).Visible = True
        OverrideCheck(i).Visible = True
        SweepToFillCheck(i).Visible = True
        DisplaySizeText(i).Visible = True
        MinQuantityText(i).Visible = True
        TriggerMethodCombo(i).Visible = True
        DiscrAmountText(i).Visible = True
        GoodAfterTimeText(i).Visible = True
        GoodAfterTimeTZText(i).Visible = True
        GoodTillDateText(i).Visible = True
        GoodTillDateTZText(i).Visible = True
    Else
        OrderIDText(i).Visible = False
        ActionCombo(i).Visible = False
        QuantityText(i).Visible = False
        TypeCombo(i).Visible = False
        PriceText(i).Visible = False
        OffsetText(i).Visible = False
        OffsetValueText(i).Visible = False
        StopPriceText(i).Visible = False
        IgnoreRthCheck(i).Visible = False
        TIFCombo(i).Visible = False
        OrderRefText(i).Visible = False
        AllOrNoneCheck(i).Visible = False
        BlockOrderCheck(i).Visible = False
        ETradeOnlyCheck(i).Visible = False
        FirmQuoteOnlyCheck(i).Visible = False
        HiddenCheck(i).Visible = False
        OverrideCheck(i).Visible = False
        SweepToFillCheck(i).Visible = False
        DisplaySizeText(i).Visible = False
        MinQuantityText(i).Visible = False
        TriggerMethodCombo(i).Visible = False
        DiscrAmountText(i).Visible = False
        GoodAfterTimeText(i).Visible = False
        GoodAfterTimeTZText(i).Visible = False
        GoodTillDateText(i).Visible = False
        GoodTillDateTZText(i).Visible = False
    End If
Next
End Sub

Private Sub showTickerValues()
AskText.Text = mTicker.FormatPrice(mTicker.AskPrice)
AskSizeText.Text = mTicker.AskSize
BidText.Text = mTicker.FormatPrice(mTicker.BidPrice)
BidSizeText.Text = mTicker.BidSize
LastText.Text = mTicker.FormatPrice(mTicker.TradePrice)
LastSizeText.Text = mTicker.TradeSize
VolumeText.Text = mTicker.Volume
HighText.Text = mTicker.FormatPrice(mTicker.HighPrice)
LowText.Text = mTicker.FormatPrice(mTicker.LowPrice)
setPriceFields
End Sub




