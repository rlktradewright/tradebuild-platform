VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OrderForm 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Ticker"
      Height          =   1815
      Left            =   120
      TabIndex        =   52
      Top             =   3840
      Width           =   2895
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
            Alignment       =   2  'Center
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox HighText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox LowText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox LastSizeText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox AskSizeText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox BidSizeText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox BidText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox LastText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox AskText 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Bid"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Ask"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Last"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "Volume"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "High"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Low"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   1200
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton UndoButton 
      Caption         =   "Undo"
      Height          =   495
      Left            =   7440
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order"
      Height          =   2895
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   2895
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   105
         ScaleHeight     =   2535
         ScaleWidth      =   2655
         TabIndex        =   35
         Top             =   240
         Width           =   2655
         Begin VB.TextBox OffsetText 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   51
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox OffsetValueText 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox OrderIDText 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox StopPriceText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   5
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox PriceText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   4
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox TypeCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderForm.frx":0000
            Left            =   960
            List            =   "OrderForm.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox QuantityText 
            Alignment       =   1  'Right Justify
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   2
            Text            =   "1"
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox ActionCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderForm.frx":0004
            Left            =   960
            List            =   "OrderForm.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Offset (ticks)"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Order id"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Stop price"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Price"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Type"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Quantity"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Action"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   4815
      Left            =   3120
      TabIndex        =   22
      Top             =   840
      Width           =   4095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   3855
         TabIndex        =   23
         Top             =   240
         Width           =   3855
         Begin VB.TextBox OcaGroupText 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox GoodAfterTimeText 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   15
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox DiscrAmountText 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   14
            Top             =   3240
            Width           =   975
         End
         Begin VB.CheckBox HiddenCheck 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   13
            Top             =   2880
            Width           =   255
         End
         Begin VB.CheckBox IgnoreRTHCheck 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   12
            Top             =   2520
            Width           =   255
         End
         Begin VB.ComboBox TriggerMethodCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderForm.frx":0008
            Left            =   1320
            List            =   "OrderForm.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox DisplaySizeText 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   10
            Top             =   1800
            Width           =   975
         End
         Begin VB.CheckBox SweepToFillCheck 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   9
            Top             =   1440
            Width           =   255
         End
         Begin VB.CheckBox BlockOrderCheck 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   8
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox OrderRefText 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox TIFCombo 
            Height          =   315
            Index           =   0
            ItemData        =   "OrderForm.frx":000C
            Left            =   1320
            List            =   "OrderForm.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label Label7 
            Caption         =   "OCA group"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Good after time"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "Discr amount"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Hidden"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Ignore RTH"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Trigger method"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Display size"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Sweep to fill"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Block order"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Order ref"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "TIF"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.ComboBox OrderSchemeCombo 
      Height          =   315
      ItemData        =   "OrderForm.frx":0010
      Left            =   1200
      List            =   "OrderForm.frx":0012
      TabIndex        =   0
      Text            =   "Simple order"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton PlaceOrdersButton 
      Caption         =   "&Place orders"
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton ResetButton 
      Cancel          =   -1  'True
      Caption         =   "&Reset"
      Height          =   495
      Left            =   7440
      TabIndex        =   21
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CompleteOrdersButton 
      Caption         =   "Complete &order"
      Height          =   495
      Left            =   7440
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton ModifyButton 
      Caption         =   "&Modify"
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7440
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip BracketTabStrip 
      Height          =   5280
      Left            =   0
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9313
      MultiRow        =   -1  'True
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
      Left            =   4920
      TabIndex        =   48
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label13 
      Caption         =   "Order scheme"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label OrderSimulationLabel 
      Alignment       =   2  'Center
      Caption         =   "Orders are simulated"
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
      Left            =   0
      TabIndex        =   42
      Top             =   5760
      Width           =   7335
   End
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuild.ChangeListener
Implements TradeBuild.QuoteListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Private Enum BracketIndices
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
    BracketOrder
    OCAOrder
    CombinationOrder
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTicker As Ticker

Private mTickSize As Double

Private mOrderAction As TradeBuild.OrderActions

Private WithEvents mOrderPlex As TradeBuild.OrderPlex
Attribute mOrderPlex.VB_VarHelpID = -1

Private mCurrentBrackerOrderIndex As BracketIndices

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()

Me.Left = 0
Me.Top = Screen.Height - Me.Height

setupOrderSchemeCombo

loadOrderFields BracketIndices.BracketStopOrder
loadOrderFields BracketIndices.BracketTargetOrder

setupActionCombo BracketIndices.BracketEntryOrder
setupActionCombo BracketIndices.BracketStopOrder
setupActionCombo BracketIndices.BracketTargetOrder

setupTypeCombo BracketIndices.BracketEntryOrder
setupTypeCombo BracketIndices.BracketStopOrder
setupTypeCombo BracketIndices.BracketTargetOrder

setupTifCombo BracketIndices.BracketEntryOrder
setupTifCombo BracketIndices.BracketStopOrder
setupTifCombo BracketIndices.BracketTargetOrder

TriggerMethodCombo(0).AddItem orderTriggerMethodToString(TriggerMethods.TriggerDefault)
TriggerMethodCombo(0).AddItem orderTriggerMethodToString(TriggerMethods.TriggerDoubleBidAsk)
TriggerMethodCombo(0).AddItem orderTriggerMethodToString(TriggerMethods.TriggerDoubleLast)
TriggerMethodCombo(0).AddItem orderTriggerMethodToString(TriggerMethods.TriggerLast)

reset

End Sub

Private Sub Form_Unload(cancel As Integer)
mTicker.removeQuoteListener Me
reset
End Sub

'================================================================================
' ChangeListener Interface Members
'================================================================================

Private Sub ChangeListener_Change(ev As TradeBuild.ChangeEvent)
Dim op As TradeBuild.OrderPlex

Set op = ev.source

Select Case ev.ChangeType
Case OrderPlexChangeTypes.ChangesApplied
    ModifyButton.Enabled = False
    UndoButton.Enabled = False
Case OrderPlexChangeTypes.ChangesCancelled
    ModifyButton.Enabled = False
    UndoButton.Enabled = False
Case OrderPlexChangeTypes.ChangesPending
    ModifyButton.Enabled = True
    UndoButton.Enabled = True
Case OrderPlexChangeTypes.Completed
    reset
Case OrderPlexChangeTypes.SelfCancelled
    reset
Case OrderPlexChangeTypes.EntryOrderChanged
    setOrderFields op.entryOrder, BracketIndices.BracketEntryOrder
Case OrderPlexChangeTypes.StopOrderChanged
    setOrderFields op.stopOrder, BracketIndices.BracketStopOrder
Case OrderPlexChangeTypes.TargetOrderChanged
    setOrderFields op.targetOrder, BracketIndices.BracketTargetOrder
Case OrderPlexChangeTypes.CloseoutOrderCreated
Case OrderPlexChangeTypes.CloseoutOrderChanged
Case OrderPlexChangeTypes.ProfitThresholdExceeded
Case OrderPlexChangeTypes.LossThresholdExceeded
Case OrderPlexChangeTypes.DrawdownThresholdExceeded
Case OrderPlexChangeTypes.SizeChanged
Case OrderPlexChangeTypes.StateChanged
End Select
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As TradeBuild.QuoteEvent)
AskText = ev.priceString
AskSizeText = ev.size
setPriceField BracketIndices.BracketEntryOrder
setPriceField BracketIndices.BracketStopOrder
setPriceField BracketIndices.BracketTargetOrder
End Sub

Private Sub QuoteListener_bid(ev As TradeBuild.QuoteEvent)
BidText = ev.priceString
BidSizeText = ev.size
setPriceField BracketIndices.BracketEntryOrder
setPriceField BracketIndices.BracketStopOrder
setPriceField BracketIndices.BracketTargetOrder
End Sub

Private Sub QuoteListener_high(ev As TradeBuild.QuoteEvent)
HighText = ev.priceString
End Sub

Private Sub QuoteListener_Low(ev As TradeBuild.QuoteEvent)
LowText = ev.priceString
End Sub

Private Sub QuoteListener_openInterest(ev As TradeBuild.QuoteEvent)

End Sub

Private Sub QuoteListener_previousClose(ev As TradeBuild.QuoteEvent)

End Sub

Private Sub QuoteListener_trade(ev As TradeBuild.QuoteEvent)
LastText = ev.priceString
LastSizeText = ev.size
setPriceField BracketIndices.BracketEntryOrder
setPriceField BracketIndices.BracketStopOrder
setPriceField BracketIndices.BracketTargetOrder
End Sub

Private Sub QuoteListener_volume(ev As TradeBuild.QuoteEvent)
VolumeText = ev.size
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub ActionCombo_Click(ByRef index As Integer)
setAction index
End Sub

Private Sub BracketTabStrip_Click()
mCurrentBrackerOrderIndex = BracketTabStrip.SelectedItem.index - 1
showOrderFields mCurrentBrackerOrderIndex
End Sub

Private Sub CancelButton_Click()
If Not mOrderPlex Is Nothing Then
    mOrderPlex.cancel True
End If
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
mOrderPlex.Update
End Sub

Private Sub OffsetText_Change(index As Integer)
If IsNumeric(OffsetText(index)) Then
    OffsetValueText(index) = OffsetText(index) * mTickSize
Else
    OffsetValueText(index) = ""
End If
setPriceField index
End Sub

Private Sub OrderSchemeCombo_Click()
setOrderScheme comboItemData(OrderSchemeCombo)
End Sub

Private Sub PlaceOrdersButton_Click()
Dim op As TradeBuild.OrderPlex

Select Case comboItemData(OrderSchemeCombo)
Case OrderSchemes.SimpleOrder
    If comboItemData(ActionCombo(BracketIndices.BracketEntryOrder)) = OrderActions.ActionBuy Then
        Set op = mTicker.defaultOrderContext.Buy( _
                                    QuantityText(BracketIndices.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndices.BracketEntryOrder)), _
                                    IIf(PriceText(BracketIndices.BracketEntryOrder) = "", 0, PriceText(BracketIndices.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketEntryOrder) = "", 0, OffsetText(BracketIndices.BracketEntryOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketEntryOrder) = "", 0, StopPriceText(BracketIndices.BracketEntryOrder)), _
                                    StopTypes.StopTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    TargetTypes.TargetTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    0, _
                                    Nothing)
    Else
        Set op = mTicker.defaultOrderContext.Sell( _
                                    QuantityText(BracketIndices.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndices.BracketEntryOrder)), _
                                    IIf(PriceText(BracketIndices.BracketEntryOrder) = "", 0, PriceText(BracketIndices.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketEntryOrder) = "", 0, OffsetText(BracketIndices.BracketEntryOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketEntryOrder) = "", 0, StopPriceText(BracketIndices.BracketEntryOrder)), _
                                    StopTypes.StopTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    TargetTypes.TargetTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    0, _
                                    Nothing)
        
    End If
Case OrderSchemes.BracketOrder
    If comboItemData(ActionCombo(BracketIndices.BracketEntryOrder)) = OrderActions.ActionBuy Then
        Set op = mTicker.defaultOrderContext.Buy( _
                                    QuantityText(BracketIndices.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndices.BracketEntryOrder)), _
                                    IIf(PriceText(BracketIndices.BracketEntryOrder) = "", 0, PriceText(BracketIndices.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketEntryOrder) = "", 0, OffsetText(BracketIndices.BracketEntryOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketEntryOrder) = "", 0, StopPriceText(BracketIndices.BracketEntryOrder)), _
                                    comboItemData(TypeCombo(BracketIndices.BracketStopOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketStopOrder) = "", 0, StopPriceText(BracketIndices.BracketStopOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketStopOrder) = "", 0, OffsetText(BracketIndices.BracketStopOrder)), _
                                    IIf(PriceText(BracketIndices.BracketStopOrder) = "", 0, PriceText(BracketIndices.BracketStopOrder)), _
                                    comboItemData(TypeCombo(BracketIndices.BracketTargetOrder)), _
                                    IIf(PriceText(BracketIndices.BracketTargetOrder) = "", 0, PriceText(BracketIndices.BracketTargetOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketTargetOrder) = "", 0, OffsetText(BracketIndices.BracketTargetOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketTargetOrder) = "", 0, StopPriceText(BracketIndices.BracketTargetOrder)), _
                                    0, _
                                    Nothing)
    Else
        Set op = mTicker.defaultOrderContext.Sell( _
                                    QuantityText(BracketIndices.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndices.BracketEntryOrder)), _
                                    IIf(PriceText(BracketIndices.BracketEntryOrder) = "", 0, PriceText(BracketIndices.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketEntryOrder) = "", 0, OffsetText(BracketIndices.BracketEntryOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketEntryOrder) = "", 0, StopPriceText(BracketIndices.BracketEntryOrder)), _
                                    comboItemData(TypeCombo(BracketIndices.BracketStopOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketStopOrder) = "", 0, StopPriceText(BracketIndices.BracketStopOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketStopOrder) = "", 0, OffsetText(BracketIndices.BracketStopOrder)), _
                                    IIf(PriceText(BracketIndices.BracketStopOrder) = "", 0, PriceText(BracketIndices.BracketStopOrder)), _
                                    comboItemData(TypeCombo(BracketIndices.BracketTargetOrder)), _
                                    IIf(PriceText(BracketIndices.BracketTargetOrder) = "", 0, PriceText(BracketIndices.BracketTargetOrder)), _
                                    IIf(OffsetText(BracketIndices.BracketTargetOrder) = "", 0, OffsetText(BracketIndices.BracketTargetOrder)), _
                                    IIf(StopPriceText(BracketIndices.BracketTargetOrder) = "", 0, StopPriceText(BracketIndices.BracketTargetOrder)), _
                                    0, _
                                    Nothing)
    End If
Case OrderSchemes.OCAOrder
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Case OrderSchemes.CombinationOrder
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Select

op.addChangeListener gMainForm
op.addProfitListener gMainForm
End Sub

Private Sub PriceText_Validate( _
                index As Integer, _
                cancel As Boolean)
Dim price As Double

If (comboItemData(ActionCombo(index)) = OrderActions.ActionNone And _
        PriceText(index) <> "" _
    ) Or _
    Not mTicker.parsePrice(PriceText(index), price) _
Then
    cancel = True
    Exit Sub
End If

If Not mOrderPlex Is Nothing Then
    Select Case index
    Case BracketIndices.BracketEntryOrder
        mOrderPlex.newEntryPrice = price
    Case BracketIndices.BracketStopOrder
        mOrderPlex.newStopPrice = price
    Case BracketIndices.BracketTargetOrder
        mOrderPlex.newTargetPrice = price
    End Select
End If
End Sub

Private Sub QuantityText_Validate( _
                index As Integer, _
                cancel As Boolean)
Dim quantity As Long

If comboItemData(ActionCombo(index)) <> OrderActions.ActionNone And _
    Not IsNumeric(QuantityText(index)) _
Then
    cancel = True
    Exit Sub
End If

quantity = CLng(QuantityText(index))

If CDbl(QuantityText(index)) - quantity <> 0 Then
    cancel = True
    Exit Sub
End If

If quantity < 0 Then
    cancel = True
    Exit Sub
End If

If mOrderPlex Is Nothing Then
    If quantity = 0 Then
        cancel = True
        Exit Sub
    End If
    
    If index = BracketIndices.BracketEntryOrder And _
        comboItemData(OrderSchemeCombo) = OrderSchemes.BracketOrder _
    Then
        QuantityText(BracketIndices.BracketStopOrder) = quantity
        QuantityText(BracketIndices.BracketTargetOrder) = quantity
    End If
Else
    mOrderPlex.newQuantity = quantity
End If
End Sub

Private Sub ResetButton_Click()
reset
End Sub

Private Sub StopPriceText_Validate( _
                index As Integer, _
                cancel As Boolean)
Dim price As Double

If (comboItemData(ActionCombo(index)) = OrderActions.ActionNone And _
        StopPriceText(index) <> "" _
    ) Or _
    Not mTicker.parsePrice(StopPriceText(index), price) _
Then
    cancel = True
    Exit Sub
End If

If Not mOrderPlex Is Nothing Then
    Select Case index
    Case BracketIndices.BracketEntryOrder
        mOrderPlex.newEntryTriggerPrice = price
    Case BracketIndices.BracketStopOrder
        mOrderPlex.newStopTriggerPrice = price
    Case BracketIndices.BracketTargetOrder
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

'================================================================================
' mOrderPlex Event Handlers
'================================================================================

Private Sub mOrderPlex_EntryOrderFilled()
disableOrderFields BracketIndices.BracketEntryOrder
End Sub

Private Sub mOrderPlex_StopOrderFilled()
disableOrderFields BracketIndices.BracketStopOrder
End Sub

Private Sub mOrderPlex_TargetOrderFilled()
disableOrderFields BracketIndices.BracketTargetOrder
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let ordersAreSimulated(ByVal value As Boolean)
If value Then
    OrderSimulationLabel.Caption = "Orders are simulated"
Else
    OrderSimulationLabel.Caption = "Orders are LIVE !!"
End If
End Property

Public Property Let Ticker(ByVal value As Ticker)
Dim orderType As Variant

Set mTicker = value
mTickSize = mTicker.Contract.TickSize

SymbolLabel.Caption = mTicker.Contract.specifier.localSymbol & _
                        " on " & _
                        mTicker.Contract.specifier.exchange
                        
mTicker.addQuoteListener Me
showTickerValues
End Property

'================================================================================
' Methods
'================================================================================

Public Sub showOrderPlex( _
                ByVal value As OrderPlex, _
                ByVal selectedOrderNumber As Long)

Dim entryOrder As TradeBuild.Order
Dim stopOrder As TradeBuild.Order
Dim targetOrder As TradeBuild.Order

Set mOrderPlex = value
Ticker = mOrderPlex.Ticker
Set entryOrder = mOrderPlex.entryOrder
Set stopOrder = mOrderPlex.stopOrder
Set targetOrder = mOrderPlex.targetOrder

If stopOrder Is Nothing And targetOrder Is Nothing Then
    Me.Caption = "Change a single order"
Else
    Me.Caption = "Change a bracket order"
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

setOrderFields entryOrder, BracketIndices.BracketEntryOrder
setOrderFields stopOrder, BracketIndices.BracketStopOrder
setOrderFields targetOrder, BracketIndices.BracketTargetOrder

ModifyButton.Move PlaceOrdersButton.Left, PlaceOrdersButton.Top
ModifyButton.Visible = True
ModifyButton.Enabled = False

PlaceOrdersButton.Visible = False

CancelButton.Move CompleteOrdersButton.Left, CompleteOrdersButton.Top
CancelButton.Visible = True

UndoButton.Move ResetButton.Left, ResetButton.Top
UndoButton.Enabled = False
UndoButton.Visible = True

ResetButton.Visible = False

mOrderPlex.addChangeListener Me
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addItemToCombo( _
                ByVal combo As ComboBox, _
                ByVal itemText As String, _
                ByVal itemData As Long)
combo.AddItem itemText
combo.itemData(combo.ListCount - 1) = itemData
End Sub

Private Sub clearOrderFields(ByVal index As Long)
enableOrderFields index
OrderIDText(index) = ""
ActionCombo(index).ListIndex = 0
QuantityText(index) = 1
' don't set TypeCombo(Index) as it will affect other fields and there
' is no sensible value to set it to
PriceText(index) = ""
StopPriceText(index) = ""
OffsetText(index) = ""
TIFCombo(index).ListIndex = 0
End Sub

Private Function comboItemData(ByVal combo As ComboBox) As Long
comboItemData = combo.itemData(combo.ListIndex)
End Function

Private Sub configureOrderFields( _
                ByVal orderIndex As Long)
Select Case orderIndex
Case BracketIndices.BracketEntryOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case EntryTypeMarket
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeMarketOnOpen
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeMarketOnClose
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeMarketIfTouched
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeMarketToLimit
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case EntryTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case EntryTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case EntryTypeLimit
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeLimitOnOpen
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeLimitOnClose
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeLimitIfTouched
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeStop
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case EntryTypeStopLimit
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    End Select
Case BracketIndices.BracketStopOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case StopTypeNone
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case StopTypeStop
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case StopTypeStopLimit
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case StopTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case StopTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case StopTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case StopTypeAuto
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    End Select
Case BracketIndices.BracketTargetOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case TargetTypeNone
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetTypeLimit
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetTypeLimitIfTouched
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetTypeMarketIfTouched
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case TargetTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case TargetTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case TargetTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case TargetTypeAuto
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    End Select
End Select
End Sub


Private Sub disableControl(ByVal field As Control)
field.Enabled = False
field.backColor = SystemColorConstants.vbButtonFace
End Sub

Private Sub disableOrderFields(ByVal index As Long)
disableControl ActionCombo(index)
disableControl QuantityText(index)
disableControl TypeCombo(index)
disableControl PriceText(index)
disableControl StopPriceText(index)
disableControl OffsetText(index)
disableControl TIFCombo(index)
End Sub

Private Sub enableControl(ByVal field As Control)
field.Enabled = True
field.backColor = SystemColorConstants.vbWindowBackground
End Sub

Private Sub enableOrderFields(ByVal index As Long)
enableControl ActionCombo(index)
enableControl QuantityText(index)
enableControl TypeCombo(index)
enableControl PriceText(index)
enableControl StopPriceText(index)
enableControl OffsetText(index)
enableControl TIFCombo(index)
End Sub

Private Function isOrderModifiable(ByVal pOrder As TradeBuild.Order) As Boolean
If pOrder Is Nothing Then Exit Function
isOrderModifiable = pOrder.isModifiable
End Function

Private Sub loadOrderFields(ByVal index As Long)
Load OrderIDText(index)
Load ActionCombo(index)
Load QuantityText(index)
Load TypeCombo(index)
Load PriceText(index)
Load StopPriceText(index)
Load OffsetText(index)
Load OffsetValueText(index)
Load TIFCombo(index)
Load OcaGroupText(index)
End Sub

Private Sub reset()
clearOrderFields BracketIndices.BracketEntryOrder
clearOrderFields BracketIndices.BracketStopOrder
clearOrderFields BracketIndices.BracketTargetOrder

OrderSchemeCombo.Enabled = True
selectComboEntry OrderSchemeCombo, OrderSchemes.BracketOrder
setOrderScheme OrderSchemes.BracketOrder

selectComboEntry ActionCombo(BracketIndices.BracketEntryOrder), _
                OrderActions.ActionBuy
setAction BracketIndices.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndices.BracketEntryOrder), _
                EntryTypes.EntryTypeLimit
configureOrderFields BracketIndices.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndices.BracketStopOrder), _
                StopTypes.StopTypeNone
configureOrderFields BracketIndices.BracketStopOrder

selectComboEntry TypeCombo(BracketIndices.BracketTargetOrder), _
                TargetTypes.TargetTypeNone
configureOrderFields BracketIndices.BracketTargetOrder

BracketTabStrip.Tabs(BracketTabs.TabEntryOrder).Selected = True

If Not mOrderPlex Is Nothing Then
    mOrderPlex.removeChangeListener Me
    Set mOrderPlex = Nothing
End If
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
If comboItemData(OrderSchemeCombo) = OrderSchemes.BracketOrder And _
    index = BracketIndices.BracketEntryOrder _
Then
    If comboItemData(ActionCombo(index)) = OrderActions.ActionSell Then
        selectComboEntry ActionCombo(BracketIndices.BracketStopOrder), OrderActions.ActionBuy
        selectComboEntry ActionCombo(BracketIndices.BracketTargetOrder), OrderActions.ActionBuy
    Else
        selectComboEntry ActionCombo(BracketIndices.BracketStopOrder), OrderActions.ActionSell
        selectComboEntry ActionCombo(BracketIndices.BracketTargetOrder), OrderActions.ActionSell
    End If
    disableControl ActionCombo(BracketIndices.BracketStopOrder)
    disableControl ActionCombo(BracketIndices.BracketTargetOrder)
End If
End Sub

Private Sub setOrderFields( _
                ByVal pOrder As TradeBuild.Order, _
                ByVal orderIndex As Long)
If pOrder Is Nothing Then
    disableOrderFields orderIndex
    Exit Sub
End If

clearOrderFields orderIndex

With pOrder
    setOrderId orderIndex, .id
    If .ocaGroup <> "" Then OcaGroupText(orderIndex) = .ocaGroup
    
    ActionCombo(orderIndex).Text = orderActionToString(.Action)
    QuantityText(orderIndex) = .quantity
    TypeCombo(orderIndex).Text = orderTypeToString(.orderType)
    PriceText(orderIndex) = IIf(.limitPrice <> 0, .limitPrice, "")
    StopPriceText(orderIndex) = IIf(.triggerPrice <> 0, .triggerPrice, "")
    TIFCombo(orderIndex) = orderTIFToString(.timeInForce)
'    OrderRefText(orderIndex) = .orderRef
'    BlockOrderCheck(orderIndex) = IIf(.blockOrder, vbChecked, vbUnchecked)
'    SweepToFillCheck(orderIndex) = IIf(.sweepToFill, vbChecked, vbUnchecked)
'    DisplaySizeText(orderIndex) = IIf(.displaySize <> 0, .displaySize, "")
'    TriggerMethodCombo(orderIndex) = orderTriggerMethodToString(.triggerMethod)
'    IgnoreRTHCheck(orderIndex) = IIf(.ignoreRTH, vbChecked, vbUnchecked)
'    HiddenCheck(orderIndex) = IIf(.Hidden, vbChecked, vbUnchecked)
'    DiscrAmountText(orderIndex) = IIf(.discretionaryAmt <> 0, .discretionaryAmt, "")
'    GoodAfterTimeText(orderIndex) = .goodAfterTime
    
    ' do this last because it sets the various fields attributes
    TypeCombo(orderIndex).Text = orderTypeToString(.orderType)
End With

If Not isOrderModifiable(pOrder) Then
    disableOrderFields orderIndex
End If
End Sub

Private Sub setOrderId( _
                ByVal index As Long, _
                ByVal id As Long)
enableControl OrderIDText(index)
OrderIDText(index) = id
disableControl OrderIDText(index)
End Sub

Private Sub setOrderScheme( _
                ByVal pOrderScheme As OrderSchemes)
Select Case pOrderScheme
Case OrderSchemes.SimpleOrder
    Me.Caption = "Create a single order"
    BracketTabStrip.Visible = False
    PlaceOrdersButton.Visible = True
    CompleteOrdersButton.Visible = False
    ModifyButton.Visible = False
    UndoButton.value = False
    showOrderFields BracketIndices.BracketEntryOrder
    
Case OrderSchemes.BracketOrder
    Me.Caption = "Create a bracket order"
    BracketTabStrip.Visible = True
    PlaceOrdersButton.Visible = True
    CompleteOrdersButton.Visible = False
    ModifyButton.Visible = False
    UndoButton.value = False
    BracketTabStrip.Tabs(BracketTabs.TabEntryOrder).Selected = True
Case OrderSchemes.OCAOrder
    Dim OCAId As Long
    Me.Caption = "Create a 'one cancels all' group"
    BracketTabStrip.Visible = False
'    If mOCAOrders Is Nothing Then Set mOCAOrders = New Collection
    PlaceOrdersButton.Visible = True
    CompleteOrdersButton.Visible = True
    ModifyButton.Visible = False
    UndoButton.value = False
Case OrderSchemes.CombinationOrder
    ' not implemented
    ' ??? whyis this any different from creating a single order
    ' on a combination ticker???
    OrderSchemeCombo.ListIndex = SimpleOrder
End Select
End Sub

Private Sub setPriceField( _
                index As Integer)
Dim basePrice As Double
Dim offset As Double

Select Case index
Case BracketIndices.BracketEntryOrder
    Select Case comboItemData(TypeCombo(index))
    Case EntryTypeBid
        basePrice = mTicker.BidPrice
    Case EntryTypeAsk
        basePrice = mTicker.AskPrice
    Case EntryTypeLast
        basePrice = mTicker.TradePrice
    Case Else
        Exit Sub
    End Select
Case BracketIndices.BracketStopOrder
    Select Case comboItemData(TypeCombo(index))
    Case StopTypeBid
        basePrice = mTicker.BidPrice
    Case StopTypeAsk
        basePrice = mTicker.AskPrice
    Case StopTypeLast
        basePrice = mTicker.TradePrice
    Case StopTypeAuto
        basePrice = 0
    Case Else
        Exit Sub
    End Select
Case BracketIndices.BracketTargetOrder
    Select Case comboItemData(TypeCombo(index))
    Case TargetTypeBid
        basePrice = mTicker.BidPrice
    Case TargetTypeAsk
        basePrice = mTicker.AskPrice
    Case TargetTypeLast
        basePrice = mTicker.TradePrice
    Case TargetTypeAuto
        basePrice = 0
    Case Else
        Exit Sub
    End Select
End Select

If IsNumeric(OffsetText(index)) Then
    offset = OffsetText(index) * mTickSize
End If

PriceText(index) = mTicker.formatPrice(basePrice + offset)
End Sub

Private Sub setupActionCombo(ByVal index As Long)
If index <> BracketIndices.BracketEntryOrder Then
    addItemToCombo ActionCombo(index), _
                orderActionToString(OrderActions.ActionNone), _
                OrderActions.ActionNone
End If
addItemToCombo ActionCombo(index), _
            orderActionToString(OrderActions.ActionBuy), _
            OrderActions.ActionBuy
addItemToCombo ActionCombo(index), _
            orderActionToString(OrderActions.ActionSell), _
            OrderActions.ActionSell
End Sub

Private Sub setupOrderSchemeCombo()
addItemToCombo OrderSchemeCombo, _
            "Bracket order", _
            OrderSchemes.BracketOrder
addItemToCombo OrderSchemeCombo, _
            "Simple order", _
            OrderSchemes.SimpleOrder
addItemToCombo OrderSchemeCombo, _
            "OCA order", _
            OrderSchemes.OCAOrder
addItemToCombo OrderSchemeCombo, _
            "Combination order", _
            OrderSchemes.CombinationOrder
OrderSchemeCombo.ListIndex = 0
End Sub

Private Sub setupTifCombo(ByVal index As Long)
addItemToCombo TIFCombo(index), _
            orderTIFToString(OrderTifs.TIFDay), _
            OrderTifs.TIFDay
addItemToCombo TIFCombo(index), _
            orderTIFToString(OrderTifs.TIFGoodTillCancelled), _
            OrderTifs.TIFGoodTillCancelled
addItemToCombo TIFCombo(index), _
            orderTIFToString(OrderTifs.TIFImmediateOrCancel), _
            OrderTifs.TIFImmediateOrCancel
TIFCombo(0).ListIndex = 0
End Sub

Private Sub setupTypeCombo(ByVal index As Long)

If index = BracketIndices.BracketEntryOrder Then
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeLimit), _
                EntryTypes.EntryTypeLimit
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeMarket), _
                EntryTypes.EntryTypeMarket
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeStop), _
                EntryTypes.EntryTypeStop
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeStopLimit), _
                EntryTypes.EntryTypeStopLimit
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeBid), _
                EntryTypes.EntryTypeBid
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeAsk), _
                EntryTypes.EntryTypeAsk
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeLast), _
                EntryTypes.EntryTypeLast
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeLimitOnOpen), _
                EntryTypes.EntryTypeLimitOnOpen
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeMarketOnOpen), _
                EntryTypes.EntryTypeMarketOnOpen
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeLimitOnClose), _
                EntryTypes.EntryTypeLimitOnClose
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeMarketOnClose), _
                EntryTypes.EntryTypeMarketOnClose
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeLimitIfTouched), _
                EntryTypes.EntryTypeLimitIfTouched
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeMarketIfTouched), _
                EntryTypes.EntryTypeMarketIfTouched
    addItemToCombo TypeCombo(index), _
                entryTypeToString(EntryTypes.EntryTypeMarketToLimit), _
                EntryTypes.EntryTypeMarketToLimit
ElseIf index = BracketIndices.BracketStopOrder Then
    addItemToCombo TypeCombo(index), _
                stopTypeToString(StopTypes.StopTypeNone), _
                StopTypes.StopTypeNone
    addItemToCombo TypeCombo(index), _
                stopTypeToString(StopTypes.StopTypeStop), _
                StopTypes.StopTypeStop
    addItemToCombo TypeCombo(index), _
                stopTypeToString(StopTypes.StopTypeStopLimit), _
                StopTypes.StopTypeStopLimit
    addItemToCombo TypeCombo(index), _
                stopTypeToString(StopTypes.StopTypeBid), _
                StopTypes.StopTypeBid
    addItemToCombo TypeCombo(index), _
                stopTypeToString(StopTypes.StopTypeAsk), _
                StopTypes.StopTypeAsk
    addItemToCombo TypeCombo(index), _
                stopTypeToString(StopTypes.StopTypeLast), _
                StopTypes.StopTypeLast
    addItemToCombo TypeCombo(index), _
                stopTypeToString(StopTypes.StopTypeAuto), _
                StopTypes.StopTypeAuto
ElseIf index = BracketIndices.BracketTargetOrder Then
    addItemToCombo TypeCombo(index), _
                targetTypeToString(TargetTypes.TargetTypeNone), _
                TargetTypes.TargetTypeNone
    addItemToCombo TypeCombo(index), _
                targetTypeToString(TargetTypes.TargetTypeLimit), _
                TargetTypes.TargetTypeLimit
    addItemToCombo TypeCombo(index), _
                targetTypeToString(TargetTypes.TargetTypeMarketIfTouched), _
                TargetTypes.TargetTypeMarketIfTouched
    addItemToCombo TypeCombo(index), _
                targetTypeToString(TargetTypes.TargetTypeBid), _
                TargetTypes.TargetTypeBid
    addItemToCombo TypeCombo(index), _
                targetTypeToString(TargetTypes.TargetTypeAsk), _
                TargetTypes.TargetTypeAsk
    addItemToCombo TypeCombo(index), _
                targetTypeToString(TargetTypes.TargetTypeLast), _
                TargetTypes.TargetTypeLast
    addItemToCombo TypeCombo(index), _
                targetTypeToString(TargetTypes.TargetTypeAuto), _
                TargetTypes.TargetTypeAuto

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
        StopPriceText(i).Visible = True
        OffsetText(i).Visible = True
        OffsetValueText(i).Visible = True
    Else
        OrderIDText(i).Visible = False
        ActionCombo(i).Visible = False
        QuantityText(i).Visible = False
        TypeCombo(i).Visible = False
        PriceText(i).Visible = False
        StopPriceText(i).Visible = False
        OffsetText(i).Visible = False
        OffsetValueText(i).Visible = False
    End If
Next
End Sub

Private Sub showTickerValues()
AskText.Text = mTicker.AskPriceString
AskSizeText.Text = mTicker.AskSize
BidText.Text = mTicker.BidPriceString
BidSizeText.Text = mTicker.bidSize
LastText.Text = mTicker.TradePriceString
LastSizeText.Text = mTicker.TradeSize
VolumeText.Text = mTicker.Volume
HighText.Text = mTicker.highPriceString
LowText.Text = mTicker.lowPriceString
End Sub
