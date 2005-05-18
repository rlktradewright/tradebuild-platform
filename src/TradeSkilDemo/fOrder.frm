VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fOrder 
   Caption         =   "Create an order"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelOrderButton 
      Caption         =   "&Cancel order"
      Height          =   495
      Left            =   7440
      TabIndex        =   20
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton ModifyOrderButton 
      Caption         =   "&Modify order"
      Height          =   495
      Left            =   7440
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CompleteOrderButton 
      Caption         =   "Complete &order"
      Height          =   495
      Left            =   7440
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox OCAGroupText 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox OrderIDText 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton ResetButton 
      Cancel          =   -1  'True
      Caption         =   "&Reset"
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton PlaceOrderButton 
      Caption         =   "&Place order"
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox OrderSchemeCombo 
      Height          =   315
      ItemData        =   "fOrder.frx":0000
      Left            =   1680
      List            =   "fOrder.frx":0002
      TabIndex        =   0
      Text            =   "Simple order"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   4815
      Left            =   3240
      TabIndex        =   23
      Top             =   1200
      Width           =   4095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   3855
         TabIndex        =   31
         Top             =   240
         Width           =   3855
         Begin VB.ComboBox TIFCombo 
            Height          =   315
            ItemData        =   "fOrder.frx":0004
            Left            =   1320
            List            =   "fOrder.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   2295
         End
         Begin VB.TextBox OrderRefText 
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox BlockOrderCheck 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1320
            TabIndex        =   8
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox SweepToFillCheck 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1320
            TabIndex        =   9
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox DisplaySizeText 
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Top             =   1800
            Width           =   975
         End
         Begin VB.ComboBox TriggerMethodCombo 
            Height          =   315
            ItemData        =   "fOrder.frx":0008
            Left            =   1320
            List            =   "fOrder.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2160
            Width           =   2295
         End
         Begin VB.CheckBox IgnoreRTHCheck 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            Top             =   2520
            Width           =   255
         End
         Begin VB.CheckBox HiddenCheck 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1320
            TabIndex        =   13
            Top             =   2880
            Width           =   255
         End
         Begin VB.TextBox DiscrAmountText 
            Height          =   285
            Left            =   1320
            TabIndex        =   14
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox GoodAfterTimeText 
            Height          =   285
            Left            =   1320
            TabIndex        =   15
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "TIF"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Order ref"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Block order"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Sweep to fill"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Display size"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Trigger method"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Ignore RTH"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Hidden"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Discr amount"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Good after time"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   3600
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Security"
      Height          =   2535
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   2775
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   2535
         TabIndex        =   48
         Top             =   240
         Width           =   2535
         Begin VB.TextBox SymbolText 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox ExpiryText 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox ExchangeText 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox RightText 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox SecTypeText 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox StrikeText 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Symbol"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Type"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Expiry"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Right"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Exchange"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Strike"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1800
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order"
      Height          =   2175
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   2535
         TabIndex        =   42
         Top             =   240
         Width           =   2535
         Begin VB.ComboBox ActionCombo 
            Height          =   315
            ItemData        =   "fOrder.frx":000C
            Left            =   960
            List            =   "fOrder.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox QuantityText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Text            =   "1"
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox TypeCombo 
            Height          =   315
            ItemData        =   "fOrder.frx":0010
            Left            =   960
            List            =   "fOrder.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox PriceText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox AuxPriceText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Action"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Quantity"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Type"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Price"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Aux price"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.TabStrip BracketTabStrip 
      Height          =   495
      Left            =   240
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   8295
   End
   Begin VB.Label OCAGroupLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "OCA group"
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Order id"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Order scheme"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "fOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event cancelOrder(ByVal orderID)
Event createOrder(ByRef order As order)  ' causes the main form to create a new order
Event nextOCAID(ByRef id As Long) ' used to get the next OCA group id from the main form
Event placeOrder( _
                ByVal pOrder As order, _
                ByVal pContractSpecifier As ContractSpecifier, _
                ByVal passToTWS As Boolean)

Private mContract As Contract

Private mEntryOrder As order
Private mStopOrder As order
Private mTargetOrder As order
Private mModifiedOrder As order
Private mCurrentOrder As order

Private mOCAOrders As Collection

Private mInitialising As Boolean

Private Enum OrderSchemes
    SimpleOrder
    BracketOrder
    OCAOrder
    CombinationOrder
End Enum

Private Enum BracketTabs
    EntryOrder = 1
    StopOrder
    TargetOrder
End Enum

Private Enum BracketOrderComponents
    EntryOrder
    StopOrder
    TargetOrder
End Enum

Private mBracketOrderComponent As BracketOrderComponents

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()

Me.Left = 0
Me.Top = Screen.Height - Me.Height

mInitialising = True

OrderSchemeCombo.AddItem "Simple order", OrderSchemes.SimpleOrder
OrderSchemeCombo.ItemData(OrderSchemes.SimpleOrder) = OrderSchemes.SimpleOrder

OrderSchemeCombo.AddItem "Bracket order", OrderSchemes.BracketOrder
OrderSchemeCombo.ItemData(OrderSchemes.BracketOrder) = OrderSchemes.BracketOrder

OrderSchemeCombo.AddItem "OCA order", OrderSchemes.OCAOrder
OrderSchemeCombo.ItemData(OrderSchemes.OCAOrder) = OrderSchemes.OCAOrder

OrderSchemeCombo.AddItem "Combination order", OrderSchemes.CombinationOrder
OrderSchemeCombo.ItemData(OrderSchemes.CombinationOrder) = OrderSchemes.CombinationOrder

OrderSchemeCombo.ListIndex = OrderSchemes.SimpleOrder

ActionCombo.AddItem orderActionToString(OrderActions.ActionBuy)
ActionCombo.AddItem orderActionToString(OrderActions.ActionSell)
ActionCombo.ListIndex = 0

TIFCombo.AddItem orderTIFToString(OrderTifs.TIFDay)
TIFCombo.AddItem orderTIFToString(OrderTifs.TIFGoodTillCancelled)
TIFCombo.AddItem orderTIFToString(OrderTifs.TIFImmediateOrCancel)
TIFCombo.ListIndex = 0

TriggerMethodCombo.AddItem orderTriggerMethodToString(TriggerMethods.TriggerDefault)
TriggerMethodCombo.AddItem orderTriggerMethodToString(TriggerMethods.TriggerDoubleBidAsk)
TriggerMethodCombo.AddItem orderTriggerMethodToString(TriggerMethods.TriggerDoubleLast)
TriggerMethodCombo.AddItem orderTriggerMethodToString(TriggerMethods.TriggerLast)
TriggerMethodCombo.ListIndex = 0

mInitialising = False
createEntryOrder
End Sub

Private Sub Form_Unload(cancel As Integer)
reset
End Sub

Private Sub ActionCombo_Click()

If mInitialising Then Exit Sub

mCurrentOrder.Action = orderActionFromString(ActionCombo)
If OrderSchemeCombo.ListIndex = OrderSchemes.BracketOrder And _
    BracketTabStrip.SelectedItem.Index = BracketTabs.EntryOrder Then
    If Not mStopOrder Is Nothing Then
        mStopOrder.Action = orderActionFromString(IIf(ActionCombo = StrOrderActionSell, _
                                                StrOrderActionBuy, _
                                                StrOrderActionSell))
    End If
    If Not mTargetOrder Is Nothing Then
        mTargetOrder.Action = orderActionFromString(IIf(ActionCombo = StrOrderActionSell, _
                                                StrOrderActionBuy, _
                                                StrOrderActionSell))
    End If
End If
End Sub

Private Sub AuxPriceText_Change()
mCurrentOrder.auxPrice = IIf(AuxPriceText = "", 0, AuxPriceText)
End Sub

Private Sub BracketTabStrip_Click()

If mInitialising Then Exit Sub

ActionCombo.Enabled = True
QuantityText.Enabled = True

Select Case BracketTabStrip.SelectedItem.Index
Case BracketTabs.EntryOrder
    Set mCurrentOrder = mEntryOrder
    TypeCombo = orderTypeToString(IIf(mCurrentOrder.orderType = orderTypes.OrderTypeNone, _
                                    orderTypes.OrderTypeLimit, _
                                    mCurrentOrder.orderType))
Case BracketTabs.StopOrder
    If mStopOrder Is Nothing Then
        RaiseEvent createOrder(mStopOrder)
        mStopOrder.Action = IIf(mEntryOrder.Action = OrderActions.ActionSell, _
                                OrderActions.ActionBuy, _
                                OrderActions.ActionSell)
        mStopOrder.quantity = QuantityText
    End If
    Set mCurrentOrder = mStopOrder
    TypeCombo = orderTypeToString(IIf(mCurrentOrder.orderType = orderTypes.OrderTypeNone, _
                                    orderTypes.OrderTypeStop, _
                                    mCurrentOrder.orderType))
Case BracketTabs.TargetOrder
    If mTargetOrder Is Nothing Then
        RaiseEvent createOrder(mTargetOrder)
        mTargetOrder.Action = IIf(mEntryOrder.Action = OrderActions.ActionSell, _
                                    OrderActions.ActionBuy, _
                                    OrderActions.ActionSell)
        mTargetOrder.quantity = QuantityText
    End If
    Set mCurrentOrder = mTargetOrder
    TypeCombo = orderTypeToString(IIf(mCurrentOrder.orderType = orderTypes.OrderTypeNone, _
                                    orderTypes.OrderTypeLimit, _
                                    mCurrentOrder.orderType))
End Select

OrderIDText = mCurrentOrder.id
ActionCombo = orderActionToString(mCurrentOrder.Action)
If BracketTabStrip.SelectedItem.Index <> BracketTabs.EntryOrder Then ActionCombo.Enabled = False

If BracketTabStrip.SelectedItem.Index <> BracketTabs.EntryOrder Then QuantityText.Enabled = False

PriceText = IIf(mCurrentOrder.limitPrice = 0, "", mCurrentOrder.limitPrice)
AuxPriceText = IIf(mCurrentOrder.auxPrice = 0, "", mCurrentOrder.auxPrice)

End Sub

Private Sub CancelOrderButton_Click()
RaiseEvent cancelOrder(mModifiedOrder.id)
End Sub

Private Sub CompleteOrderButton_Click()
Dim i As Long
Dim order As order

mOCAOrders.Add mEntryOrder
For i = 1 To mOCAOrders.Count
    Set order = mOCAOrders(i)
    placeOrder order, IIf(i = mOCAOrders.Count, True, False), True
Next

Set mOCAOrders = Nothing

OrderIDText = ""
OCAGroupText = ""
OCAGroupText.Visible = False
OCAGroupLabel.Visible = False
OrderSchemeCombo.Enabled = True
OrderSchemeCombo.ListIndex = SimpleOrder
End Sub

Private Sub ModifyOrderButton_Click()
placeOrder mModifiedOrder, True, True
End Sub

Private Sub OrderSchemeCombo_Click()

If mInitialising Then Exit Sub

Select Case OrderSchemeCombo.ListIndex
Case SimpleOrder
    OCAGroupText.Visible = False
    OCAGroupLabel.Visible = False
    BracketTabStrip.Visible = False
    PlaceOrderButton.Visible = True
    CompleteOrderButton.Visible = False
    ModifyOrderButton.Visible = False
    createEntryOrder
Case BracketOrder
    OrderSchemeCombo.Enabled = False
    OCAGroupText.Visible = False
    OCAGroupLabel.Visible = False
    BracketTabStrip.Visible = True
    PlaceOrderButton.Visible = True
    CompleteOrderButton.Visible = False
    ModifyOrderButton.Visible = False
    createEntryOrder
    BracketTabStrip.Tabs(BracketTabs.EntryOrder).Selected = True
Case OCAOrder
    Dim OCAId As Long
    OrderSchemeCombo.Enabled = False
    OCAGroupText.Visible = True
    OCAGroupLabel.Visible = True
    BracketTabStrip.Visible = False
    If mOCAOrders Is Nothing Then Set mOCAOrders = New Collection
    PlaceOrderButton.Visible = True
    CompleteOrderButton.Visible = True
    ModifyOrderButton.Visible = False
    RaiseEvent nextOCAID(OCAId)
    OCAGroupText = OCAId
    createEntryOrder
Case CombinationOrder
    
    MsgBox "This facility is not yet implemented", vbInformation, "Sorry"
End Select

End Sub

Private Sub PlaceOrderButton_Click()
Select Case OrderSchemeCombo.ListIndex
Case OrderSchemes.SimpleOrder
    placeOrder mEntryOrder, True, True
    createEntryOrder
Case OrderSchemes.BracketOrder
    placeOrder mEntryOrder, False, True
    If Not mStopOrder Is Nothing Then
        If mStopOrder.orderType <> orderTypes.OrderTypeNone Then
            mStopOrder.parentId = mEntryOrder.id
            If mTargetOrder Is Nothing Then
                placeOrder mStopOrder, True, True
            ElseIf mTargetOrder.orderType <> orderTypes.OrderTypeNone Then
                placeOrder mStopOrder, False, True
            Else
                placeOrder mStopOrder, True, True
            End If
        End If
        Set mStopOrder = Nothing
    End If
    If Not mTargetOrder Is Nothing Then
        If mTargetOrder.orderType <> orderTypes.OrderTypeNone Then
            mTargetOrder.parentId = mEntryOrder.id
            placeOrder mTargetOrder, True, True
        End If
        Set mTargetOrder = Nothing
    End If
    
    BracketTabStrip.Visible = False
    ActionCombo.Enabled = True
    QuantityText.Enabled = True
    OrderSchemeCombo.Enabled = True
    OrderSchemeCombo.ListIndex = SimpleOrder

Case OrderSchemes.OCAOrder
    mOCAOrders.Add mEntryOrder
    placeOrder mEntryOrder, False, False
    createEntryOrder
Case OrderSchemes.CombinationOrder
End Select
End Sub

Private Sub PriceText_Change()
mCurrentOrder.limitPrice = IIf(PriceText = "", 0, PriceText)
End Sub

Private Sub QuantityText_Change()
mCurrentOrder.quantity = IIf(QuantityText = "", 0, QuantityText)
If Not mStopOrder Is Nothing Then mStopOrder.quantity = mEntryOrder.quantity
If Not mTargetOrder Is Nothing Then mTargetOrder.quantity = mEntryOrder.quantity
End Sub

Private Sub ResetButton_Click()
reset
End Sub

Private Sub TypeCombo_Click()

If mInitialising Then Exit Sub

Select Case TypeCombo
Case StrOrderTypeMarket
    PriceText.Enabled = False
    AuxPriceText.Enabled = False
Case StrOrderTypeMarketClose
    PriceText.Enabled = False
    AuxPriceText.Enabled = False
Case StrOrderTypeLimit
    PriceText.Enabled = True
    AuxPriceText.Enabled = False
Case StrOrderTypeLimitClose
    PriceText.Enabled = True
    AuxPriceText.Enabled = False
Case StrOrderTypePegMarket
    PriceText.Enabled = False
    AuxPriceText.Enabled = False
Case StrOrderTypeStop
    PriceText.Enabled = False
    AuxPriceText.Enabled = True
Case StrOrderTypeStopLimit
    PriceText.Enabled = True
    AuxPriceText.Enabled = True
Case StrOrderTypeTrail
    PriceText.Enabled = False
    AuxPriceText.Enabled = True
Case StrOrderTypeRelative
    PriceText.Enabled = True
    AuxPriceText.Enabled = True
Case StrOrderTypeVWAP
    PriceText.Enabled = False
    AuxPriceText.Enabled = False
Case StrOrderTypeMarketToLimit
    PriceText.Enabled = False
    AuxPriceText.Enabled = False
Case StrOrderTypeQuote
    PriceText.Enabled = False
    AuxPriceText.Enabled = False
End Select

If Not mCurrentOrder Is Nothing Then
    mCurrentOrder.orderType = orderTypeFromString(TypeCombo)
End If
End Sub


'=================================================================================
'
' Properties
'
'=================================================================================

Public Property Let Contract(ByVal value As Contract)
Dim orderType As Variant

Set mContract = value
SymbolText = mContract.specifier.symbol
SecTypeText = secTypeToString(mContract.specifier.secType)
ExpiryText = Left$(mContract.specifier.expiry, 6)
ExchangeText = mContract.specifier.exchange
StrikeText = IIf(mContract.specifier.strike = 0, "", mContract.specifier.strike)
RightText = optionRightToString(mContract.specifier.Right)
TypeCombo.Clear
For Each orderType In mContract.orderTypes
    TypeCombo.AddItem orderTypeToString(orderType)
Next
TypeCombo.ListIndex = 0
End Property

Public Property Let order(ByVal value As order)

ActionCombo.Enabled = True
QuantityText.Enabled = True
OrderSchemeCombo.Enabled = True
OrderSchemeCombo.ListIndex = SimpleOrder
BracketTabStrip.Visible = False
PlaceOrderButton.Enabled = True
CompleteOrderButton.Visible = False
OrderIDText = ""
OCAGroupText = ""
Set mStopOrder = Nothing
Set mTargetOrder = Nothing
Set mOCAOrders = Nothing

Set mEntryOrder = Nothing
Set mModifiedOrder = value
With mModifiedOrder
    OrderIDText = .id
    If .ocaGroup <> "" Then
        OCAGroupText = .ocaGroup
        OCAGroupText.Visible = True
        OCAGroupLabel.Visible = True
    End If
    ActionCombo.Text = orderActionToString(.Action)
    QuantityText = .quantity
    TypeCombo.Text = orderTypeToString(.orderType)
    PriceText = IIf(.limitPrice <> 0, .limitPrice, "")
    AuxPriceText = IIf(.auxPrice <> 0, .auxPrice, "")
    TIFCombo = orderTIFToString(.timeInForce)
    OrderRefText = .orderRef
    BlockOrderCheck = IIf(.blockOrder, vbChecked, vbUnchecked)
    SweepToFillCheck = IIf(.sweepToFill, vbChecked, vbUnchecked)
    DisplaySizeText = IIf(.displaySize <> 0, .displaySize, "")
    TriggerMethodCombo.ListIndex = .triggerMethod
    IgnoreRTHCheck = IIf(.ignoreRTH, vbChecked, vbUnchecked)
    HiddenCheck = IIf(.Hidden, vbChecked, vbUnchecked)
    DiscrAmountText = IIf(.discretionaryAmt <> 0, .discretionaryAmt, "")
    GoodAfterTimeText = .goodAfterTime
End With

Set mCurrentOrder = mModifiedOrder

ModifyOrderButton.Move PlaceOrderButton.Left, PlaceOrderButton.Top
ModifyOrderButton.Visible = True
PlaceOrderButton.Visible = False

CancelOrderButton.Move ResetButton.Left, ResetButton.Top
CancelOrderButton.Visible = True
ResetButton.Visible = False

Me.Caption = "Modify an order"
End Property

Public Property Let ordersAreSimulated(ByVal value As Boolean)
If value Then
    OrderSimulationLabel.Caption = "Orders are simulated"
Else
    OrderSimulationLabel.Caption = "Orders are LIVE !!"
End If
End Property

'=================================================================================
'
' Methods
'
'=================================================================================

Public Sub orderCompleted(ByVal value As order)
If value Is mModifiedOrder Then
    ' this occurs when a order is being modified and is then
    ' cancelled from the main form or is filled
    reset
End If
    

End Sub

'=================================================================================
'
' Helper functions
'
'=================================================================================

Private Sub createEntryOrder()

RaiseEvent createOrder(mEntryOrder)

OrderIDText = mEntryOrder.id

mEntryOrder.Action = orderActionFromString(ActionCombo)
mEntryOrder.quantity = QuantityText
If TypeCombo <> "" Then
    mEntryOrder.orderType = orderTypeFromString(TypeCombo)
End If
mEntryOrder.limitPrice = IIf(PriceText = "", 0, PriceText)
mEntryOrder.auxPrice = IIf(AuxPriceText = "", 0, AuxPriceText)

Set mCurrentOrder = mEntryOrder

End Sub

Private Function orderTypeFromString(ByVal value As String) As orderTypes
Select Case value
Case StrOrderTypeMarket
    orderTypeFromString = orderTypes.OrderTypeMarket
Case StrOrderTypeMarketClose
    orderTypeFromString = orderTypes.OrderTypeMarketClose
Case StrOrderTypeLimit
    orderTypeFromString = orderTypes.OrderTypeLimit
Case StrOrderTypeLimitClose
    orderTypeFromString = orderTypes.OrderTypeLimitClose
Case StrOrderTypePegMarket
    orderTypeFromString = orderTypes.OrderTypePegMarket
Case StrOrderTypeStop
    orderTypeFromString = orderTypes.OrderTypeStop
Case StrOrderTypeStopLimit
    orderTypeFromString = orderTypes.OrderTypeStopLimit
Case StrOrderTypeTrail
    orderTypeFromString = orderTypes.OrderTypeTrail
Case StrOrderTypeRelative
    orderTypeFromString = orderTypes.OrderTypeRelative
Case StrOrderTypeVWAP
    orderTypeFromString = orderTypes.OrderTypeVWAP
Case StrOrderTypeMarketToLimit
    orderTypeFromString = orderTypes.OrderTypeMarketToLimit
Case StrOrderTypeQuote
    orderTypeFromString = orderTypes.OrderTypeQuote
End Select
End Function

Private Sub placeOrder(ByVal pOrder As order, _
                        ByVal transmit As Boolean, _
                        ByVal passToTWS As Boolean)
Dim lContractSpecifier As ContractSpecifier

With pOrder
    .openClose = "O"
    .blockOrder = BlockOrderCheck
    .discretionaryAmt = IIf(DiscrAmountText <> "", DiscrAmountText, 0)
    .displaySize = IIf(DisplaySizeText <> "", DisplaySizeText, 0)
    .goodAfterTime = GoodAfterTimeText
    .Hidden = HiddenCheck
    .ignoreRTH = IgnoreRTHCheck
    .ocaGroup = OCAGroupText
    .orderRef = OrderRefText
    .sweepToFill = SweepToFillCheck
    .timeInForce = orderTIFFromString(TIFCombo)
    .triggerMethod = orderTriggerMethodFromString(TriggerMethodCombo)
    
    Set lContractSpecifier = mContract.specifier
    
    .transmit = transmit
End With

RaiseEvent placeOrder(pOrder, lContractSpecifier, passToTWS)

End Sub

Private Sub reset()
ActionCombo.Enabled = True
QuantityText.Enabled = True
OrderSchemeCombo.Enabled = True
If OrderSchemeCombo.ListIndex = SimpleOrder Then
    createEntryOrder
Else
    OrderSchemeCombo.ListIndex = SimpleOrder ' NB this creates a new entry order
End If
BracketTabStrip.Visible = False
CompleteOrderButton.Visible = False
ModifyOrderButton.Visible = False
PlaceOrderButton.Visible = True
CancelOrderButton.Visible = False
ResetButton.Visible = True
OCAGroupText = ""
Set mStopOrder = Nothing
Set mTargetOrder = Nothing
Set mOCAOrders = Nothing
Me.Caption = "Create an order"
End Sub
