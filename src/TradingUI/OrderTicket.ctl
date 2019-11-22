VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#33.0#0"; "TWControls40.ocx"
Begin VB.UserControl OrderTicket 
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   ScaleHeight     =   6195
   ScaleWidth      =   8790
   Begin VB.OptionButton BracketOrderOption 
      Caption         =   "&Bracket order"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1560
      TabIndex        =   32
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton SimpleOrderOption 
      Caption         =   "&Simple order"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox SimulateOrdersCheck 
      Caption         =   "S&imulate orders"
      Height          =   195
      Left            =   3480
      TabIndex        =   33
      Top             =   120
      Width           =   1455
   End
   Begin TWControls40.TWButton UndoButton 
      Height          =   495
      Left            =   7560
      TabIndex        =   30
      Top             =   5250
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Undo"
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
   Begin TWControls40.TWButton PlaceOrdersButton 
      Height          =   495
      Left            =   7560
      TabIndex        =   25
      Top             =   1020
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Place orders"
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
   Begin TWControls40.TWButton ResetButton 
      Height          =   495
      Left            =   7560
      TabIndex        =   28
      Top             =   2820
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Reset"
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
   Begin TWControls40.TWButton CompleteOrdersButton 
      Height          =   495
      Left            =   7560
      TabIndex        =   26
      Top             =   1620
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Complete &order"
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
   Begin TWControls40.TWButton ModifyButton 
      Height          =   495
      Left            =   7560
      TabIndex        =   29
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Modify"
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
   Begin TWControls40.TWButton CancelButton 
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   2220
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Cancel orders"
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
   Begin VB.PictureBox BackPicture 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   8775
      TabIndex        =   72
      Top             =   795
      Width           =   8775
      Begin VB.Frame Frame2 
         Caption         =   "Ticker"
         Height          =   1815
         Left            =   120
         TabIndex        =   52
         Top             =   3180
         Width           =   3135
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   105
            ScaleHeight     =   1455
            ScaleWidth      =   2655
            TabIndex        =   53
            Top             =   240
            Width           =   2655
            Begin VB.Label VolumeLabel 
               Height          =   255
               Left            =   960
               TabIndex        =   62
               Top             =   720
               Width           =   855
            End
            Begin VB.Label HighLabel 
               Height          =   255
               Left            =   960
               TabIndex        =   61
               Top             =   960
               Width           =   855
            End
            Begin VB.Label LowLabel 
               Height          =   255
               Left            =   960
               TabIndex        =   60
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label LastSizeLabel 
               Height          =   255
               Left            =   1920
               TabIndex        =   59
               Top             =   240
               Width           =   735
            End
            Begin VB.Label AskSizeLabel 
               Height          =   255
               Left            =   1920
               TabIndex        =   58
               Top             =   0
               Width           =   735
            End
            Begin VB.Label BidSizeLabel 
               Height          =   255
               Left            =   1920
               TabIndex        =   57
               Top             =   480
               Width           =   735
            End
            Begin VB.Label BidLabel 
               Height          =   255
               Left            =   960
               TabIndex        =   56
               Top             =   480
               Width           =   855
            End
            Begin VB.Label LastLabel 
               Height          =   255
               Left            =   960
               TabIndex        =   55
               Top             =   240
               Width           =   855
            End
            Begin VB.Label AskLabel 
               Height          =   255
               Left            =   960
               TabIndex        =   54
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
      Begin VB.Frame Frame1 
         Caption         =   "Order"
         Height          =   3015
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   3135
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   2685
            Left            =   105
            ScaleHeight     =   2685
            ScaleWidth      =   2895
            TabIndex        =   45
            Top             =   240
            Width           =   2895
            Begin VB.TextBox TriggerOffsetText 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   0
               Left            =   1890
               TabIndex        =   6
               Top             =   2070
               Width           =   600
            End
            Begin TWControls40.TWImageCombo TypeCombo 
               Height          =   330
               Index           =   0
               Left            =   960
               TabIndex        =   2
               Top             =   1080
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "OrderTicket.ctx":0000
               Text            =   ""
            End
            Begin TWControls40.TWImageCombo ActionCombo 
               Height          =   330
               Index           =   0
               Left            =   960
               TabIndex        =   0
               Top             =   360
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "OrderTicket.ctx":001C
               Text            =   ""
            End
            Begin VB.TextBox TriggerPriceText 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   5
               Top             =   2070
               Width           =   855
            End
            Begin VB.TextBox LimitOffsetText 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   0
               Left            =   1890
               TabIndex        =   4
               Top             =   1440
               Width           =   600
            End
            Begin VB.TextBox LimitPriceText 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   3
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox QuantityText 
               Alignment       =   1  'Right Justify
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   1
               Top             =   720
               Width           =   855
            End
            Begin VB.Label CurrentTriggerPriceLabel 
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   75
               Top             =   2385
               Width           =   855
            End
            Begin VB.Label CurrentLimitPriceLabel 
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   74
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label OrderIdLabel 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   71
               Top             =   0
               Width           =   2535
            End
            Begin VB.Label Label6 
               Caption         =   "Id"
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label5 
               Caption         =   "Trigger price"
               Height          =   255
               Left            =   0
               TabIndex        =   50
               Top             =   2070
               Width           =   900
            End
            Begin VB.Label Label4 
               Caption         =   "Limit Price"
               Height          =   255
               Left            =   0
               TabIndex        =   49
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Type"
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Quantity"
               Height          =   255
               Left            =   0
               TabIndex        =   47
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Action"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   46
               Top             =   360
               Width           =   855
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Options"
         Height          =   4875
         Left            =   3360
         TabIndex        =   34
         Top             =   120
         Width           =   3975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   120
            ScaleHeight     =   4455
            ScaleWidth      =   3735
            TabIndex        =   35
            Top             =   240
            Width           =   3735
            Begin TWControls40.TWImageCombo TriggerMethodCombo 
               Height          =   330
               Index           =   0
               Left            =   1200
               TabIndex        =   17
               Top             =   2160
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "OrderTicket.ctx":0038
               Text            =   ""
            End
            Begin TWControls40.TWImageCombo TIFCombo 
               Height          =   330
               Index           =   0
               Left            =   1200
               TabIndex        =   7
               Top             =   0
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "OrderTicket.ctx":0054
               Text            =   ""
            End
            Begin VB.CheckBox IgnoreRthCheck 
               Caption         =   "Ignore RTH"
               Height          =   255
               Index           =   0
               Left            =   2520
               TabIndex        =   8
               Top             =   0
               Width           =   1215
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
               Width           =   1410
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
               Width           =   1335
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Min qty"
               Height          =   375
               Left            =   2040
               TabIndex        =   43
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "Good till date"
               Height          =   255
               Left            =   0
               TabIndex        =   42
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label21 
               Caption         =   "Good after time"
               Height          =   255
               Left            =   0
               TabIndex        =   41
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "Discr amount"
               Height          =   255
               Left            =   0
               TabIndex        =   40
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label17 
               Caption         =   "Trigger method"
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "Display size"
               Height          =   255
               Left            =   0
               TabIndex        =   38
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Order ref"
               Height          =   255
               Left            =   0
               TabIndex        =   37
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label10 
               Caption         =   "TIF"
               Height          =   255
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Width           =   855
            End
         End
      End
      Begin VB.Label OrderSimulationLabel 
         Alignment       =   2  'Center
         Caption         =   "Qazzly wazzox"
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
         TabIndex        =   73
         Top             =   5010
         Width           =   7215
      End
   End
   Begin MSComctlLib.TabStrip BracketTabStrip 
      Height          =   5760
      Left            =   0
      TabIndex        =   69
      Top             =   480
      Visible         =   0   'False
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   10160
      MultiRow        =   -1  'True
      Style           =   2
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
      TabIndex        =   70
      Top             =   120
      Width           =   3615
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

Implements IChangeListener
Implements IGenericTickListener
Implements IStateChangeListener
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event CaptionChanged(ByVal Caption As String)
Event NeedSimulatedOrderContext()
Event NeedLiveOrderContext()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                        As String = "OrderTicket"

Private Const NotReadyMessage                   As String = "Not ready for placing orders"
Private Const NoContractMessage                 As String = NotReadyMessage & ": no contract"

Private Const OrdersLiveMessage                 As String = "Orders are LIVE !!"
Private Const OrdersSimulatedMessage            As String = "Orders are simulated"

'@================================================================================
' Enums
'@================================================================================

Private Enum BracketIndexes
    BracketEntryOrder
    BracketStopLossOrder
    BracketTargetOrder
End Enum

Private Enum BracketTabs
    TabEntryOrder = 1
    TabStopLossOrder
    TabTargetOrder
End Enum

Private Enum OrderSchemes
    SimpleOrder
    BracketOrder
    OCAOrder
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                                     As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1
Private WithEvents mActiveOrderContext                  As OrderContext
Attribute mActiveOrderContext.VB_VarHelpID = -1

Private mLiveOrderContext                               As OrderContext
Attribute mLiveOrderContext.VB_VarHelpID = -1
Private mSimulatedOrderContext                          As OrderContext
Attribute mSimulatedOrderContext.VB_VarHelpID = -1

Private mContract                                       As IContract

Private mBracketOrder                                   As IBracketOrder
Attribute mBracketOrder.VB_VarHelpID = -1

Private mEntryLimitPriceSpec                            As PriceSpecifier
Private mEntryTriggerPriceSpec                          As PriceSpecifier

Private mStopLossLimitPriceSpec                         As PriceSpecifier
Private mStopLossTriggerPriceSpec                       As PriceSpecifier

Private mTargetLimitPriceSpec                           As PriceSpecifier
Private mTargetTriggerPriceSpec                         As PriceSpecifier

Private mCurrentBracketOrderIndex                       As BracketIndexes

Private mInvalidTexts(2)                                As TextBox

Private mMode                                           As OrderTicketModes

Private mTheme                                          As ITheme

Private mPriceToleranceTicks                            As Long

Private mAskPrice                                       As Double
Private mBidPrice                                       As Double
Private mTradePrice                                     As Double

Private mMaxDiscretionaryAmountTicks                    As Long

Private mFutureBuilder                                  As FutureBuilder
Private WithEvents mFutureWaiter                        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

mPriceToleranceTicks = 100
mMaxDiscretionaryAmountTicks = 10

BracketOrderOption.Value = True
setOrderScheme BracketOrder

loadOrderFields BracketIndexes.BracketStopLossOrder
loadOrderFields BracketIndexes.BracketTargetOrder

setupActionCombo BracketIndexes.BracketEntryOrder
setupActionCombo BracketIndexes.BracketStopLossOrder
setupActionCombo BracketIndexes.BracketTargetOrder

disableAll NoContractMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
Clear
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

Dim op As IBracketOrder
Set op = ev.Source

Select Case ev.changeType
Case BracketOrderChangeTypes.BracketOrderChangesApplied
    ModifyButton.Enabled = False
    UndoButton.Enabled = False
Case BracketOrderChangeTypes.BracketOrderChangesCancelled
    ModifyButton.Enabled = False
    UndoButton.Enabled = False
Case BracketOrderChangeTypes.BracketOrderChangesPending
    ModifyButton.Enabled = True
    UndoButton.Enabled = True
Case BracketOrderChangeTypes.BracketOrderCompleted
    'clearBracketOrder
    Set mBracketOrder = Nothing
    setupControls
Case BracketOrderChangeTypes.BracketOrderSelfCancelled
    'clearBracketOrder
    Set mBracketOrder = Nothing
    setupControls
Case BracketOrderChangeTypes.BracketOrderEntryOrderChanged
    If op.EntryOrder.Status = OrderStatusFilled Then disableOrderFields BracketIndexes.BracketEntryOrder
    setOrderFieldValues op.EntryOrder, BracketIndexes.BracketEntryOrder
Case BracketOrderChangeTypes.BracketOrderStopLossOrderChanged
    If op.StopLossOrder.Status = OrderStatusFilled Then disableOrderFields BracketIndexes.BracketStopLossOrder
    setOrderFieldValues op.StopLossOrder, BracketIndexes.BracketStopLossOrder
Case BracketOrderChangeTypes.BracketOrderTargetOrderChanged
    If op.TargetOrder.Status = OrderStatusFilled Then disableOrderFields BracketIndexes.BracketTargetOrder
    setOrderFieldValues op.TargetOrder, BracketIndexes.BracketTargetOrder
Case BracketOrderChangeTypes.BracketOrderCloseoutOrderCreated
Case BracketOrderChangeTypes.BracketOrderCloseoutOrderChanged
Case BracketOrderChangeTypes.BracketOrderSizeChanged
Case BracketOrderChangeTypes.BracketOrderStateChanged
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Dim lPriceText As String
lPriceText = priceToString(ev.Tick.Price)

Select Case ev.Tick.TickType
Case TickTypeBid
    mBidPrice = ev.Tick.Price
    BidLabel = lPriceText
    BidSizeLabel = ev.Tick.Size
    
    If mFutureBuilder Is Nothing Then
        setPriceFields
    ElseIf ticksAvailable Then
        mFutureBuilder.Complete
        Set mFutureBuilder = Nothing
    End If
Case TickTypeAsk
    mAskPrice = ev.Tick.Price
    AskLabel = lPriceText
    AskSizeLabel = ev.Tick.Size
    
    If mFutureBuilder Is Nothing Then
        setPriceFields
    ElseIf ticksAvailable Then
        mFutureBuilder.Complete
        Set mFutureBuilder = Nothing
    End If
Case TickTypeHighPrice
    HighLabel = lPriceText
Case TickTypeLowPrice
    LowLabel = lPriceText
Case TickTypeTrade
    mTradePrice = ev.Tick.Price
    LastLabel = lPriceText
    LastSizeLabel = ev.Tick.Size

    If mFutureBuilder Is Nothing Then
        setPriceFields
    ElseIf ticksAvailable Then
        mFutureBuilder.Complete
        Set mFutureBuilder = Nothing
    End If
Case TickTypeVolume
    VolumeLabel = ev.Tick.Size
End Select

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
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim lState As MarketDataSourceStates
lState = ev.State

Select Case lState
Case MarketDataSourceStateCreated

Case MarketDataSourceStateReady

Case MarketDataSourceStateRunning

Case MarketDataSourceStatePaused

Case MarketDataSourceStateStopped
    Clear
Case MarketDataSourceStateFinished
    Clear
Case MarketDataSourceStateError

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ActionCombo_Click(ByRef Index As Integer)
Const ProcName As String = "ActionCombo_Click"
On Error GoTo Err

setAction Index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BracketOrderOption_Click()
Const ProcName As String = "BracketOrderOption_Click"
On Error GoTo Err

setOrderScheme BracketOrder

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BracketTabStrip_Click()
Const ProcName As String = "BracketTabStrip_Click"
On Error GoTo Err

mCurrentBracketOrderIndex = BracketTabStrip.SelectedItem.Index - 1
showOrderFields mCurrentBracketOrderIndex

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"
On Error GoTo Err

If Not mBracketOrder Is Nothing Then mBracketOrder.Cancel True
clearBracketOrder
setupControls

CancelButton.Visible = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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
'OrderIdLabel = ""
'OcaGroupText = ""
'OcaGroupText.Visible = False
'OCAGroupLabel.Visible = False
'OrderSchemeCombo.Enabled = True
'OrderSchemeCombo.ListIndex = SimpleOrder
End Sub

Private Sub LimitOffsetText_KeyDown( _
                Index As Integer, _
                KeyCode As Integer, _
                Shift As Integer)
Const ProcName As String = "LimitOffsetText_KeyDown"
On Error GoTo Err

If KeyCode = KeyCodeConstants.vbKeyReturn Then
    validateLimitPriceFields Index
    setCurrentLimitPriceField Index
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub LimitOffsetText_Validate( _
                Index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "LimitOffsetText_Change"
On Error GoTo Err

If Not validateLimitPriceFields(Index) Then Cancel = True
setCurrentLimitPriceField Index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub LimitPriceText_KeyDown( _
                Index As Integer, _
                KeyCode As Integer, _
                Shift As Integer)
Const ProcName As String = "LimitPriceText_KeyDown"
On Error GoTo Err

If KeyCode = KeyCodeConstants.vbKeyReturn Then
    validateLimitPriceFields Index
    setCurrentLimitPriceField Index
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub LimitPriceText_Validate( _
                Index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "LimitPriceText_Validate"
On Error GoTo Err

If Not validateLimitPriceFields(Index) Then Cancel = True
setCurrentLimitPriceField Index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ModifyButton_Click()
Const ProcName As String = "ModifyButton_Click"
On Error GoTo Err

If Not isValidOrder(BracketEntryOrder) Then Exit Sub
setOrderAttributes mBracketOrder.EntryOrder, BracketIndexes.BracketEntryOrder
If Not mBracketOrder.StopLossOrder Is Nothing Then
    If Not isValidOrder(BracketStopLossOrder) Then Exit Sub
    setOrderAttributes mBracketOrder.StopLossOrder, BracketIndexes.BracketStopLossOrder
End If
If Not mBracketOrder.TargetOrder Is Nothing Then
    If Not isValidOrder(BracketTargetOrder) Then Exit Sub
    setOrderAttributes mBracketOrder.TargetOrder, BracketIndexes.BracketTargetOrder
End If
mBracketOrder.Update

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub PlaceOrdersButton_Click()
Const ProcName As String = "PlaceOrdersButton_Click"
On Error GoTo Err

If Not isValidOrder(BracketEntryOrder) Then Exit Sub

Dim lOrderType As OrderTypes
lOrderType = comboItemData(TypeCombo(BracketIndexes.BracketEntryOrder))

Dim lEntryOrder As IOrder
Set lEntryOrder = mActiveOrderContext.CreateEntryOrder( _
                        lOrderType, _
                        mEntryLimitPriceSpec, _
                        mEntryTriggerPriceSpec _
                )

Dim op As IBracketOrder
If SimpleOrderOption.Value Then
    Set op = mActiveOrderContext.CreateBracketOrder( _
                    comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder)), _
                    QuantityText(BracketIndexes.BracketEntryOrder), _
                    lEntryOrder _
                )
    
    setOrderAttributes op.EntryOrder, BracketIndexes.BracketEntryOrder
    mActiveOrderContext.ExecuteBracketOrder op
ElseIf BracketOrderOption.Value Then
    If Not isValidOrder(BracketStopLossOrder) Then Exit Sub
    If Not isValidOrder(BracketTargetOrder) Then Exit Sub
    
    lOrderType = comboItemData(TypeCombo(BracketIndexes.BracketStopLossOrder))
    Dim lStopLossOrder As IOrder
    If lOrderType <> OrderTypeNone Then _
        Set lStopLossOrder = mActiveOrderContext.CreateStopLossOrder( _
                                lOrderType, _
                                mStopLossLimitPriceSpec, _
                                mStopLossTriggerPriceSpec _
                        )
    
    lOrderType = comboItemData(TypeCombo(BracketIndexes.BracketTargetOrder))
    Dim lTargetOrder As IOrder
    If lOrderType <> OrderTypeNone Then _
        Set lTargetOrder = mActiveOrderContext.CreateTargetOrder( _
                            lOrderType, _
                            mTargetLimitPriceSpec, _
                            mTargetTriggerPriceSpec _
                    )
    
    Set op = mActiveOrderContext.CreateBracketOrder( _
                    comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder)), _
                    QuantityText(BracketIndexes.BracketEntryOrder), _
                    lEntryOrder, _
                    lStopLossOrder, _
                    lTargetOrder _
                )
    
    setOrderAttributes op.EntryOrder, BracketIndexes.BracketEntryOrder
    If Not op.StopLossOrder Is Nothing Then
        setOrderAttributes op.StopLossOrder, BracketIndexes.BracketStopLossOrder
    End If
    If Not op.TargetOrder Is Nothing Then
        setOrderAttributes op.TargetOrder, BracketIndexes.BracketTargetOrder
    End If
    mActiveOrderContext.ExecuteBracketOrder op
    
    Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabEntryOrder)
Else
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub QuantityText_Validate( _
                Index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "QuantityText_Validate"
On Error GoTo Err

If comboItemData(ActionCombo(Index)) <> OrderActions.OrderActionNone And _
    Not IsNumeric(QuantityText(Index)) _
Then
    highlightText QuantityText(Index)
    Cancel = True
    Exit Sub
End If

Dim max As Long
Dim min As Long

Select Case mContract.Specifier.secType
Case SecTypeStock
    min = 10
    max = 100000
Case SecTypeFuture
    min = 1
    max = 100
Case SecTypeOption
    min = 1
    max = 100
Case SecTypeFuturesOption
    min = 1
    max = 100
Case SecTypeCash
    min = 1000
    max = 10000000
Case SecTypeCombo
    min = 1
    max = 100
Case SecTypeIndex
    min = 0
    max = 0
End Select

If Not IsInteger(QuantityText(Index), min, max) Then
    highlightText QuantityText(Index)
    Cancel = True
    Exit Sub
End If

Dim Quantity As Long
Quantity = CLng(QuantityText(Index))

If mBracketOrder Is Nothing Then
    If Quantity = 0 Then
        highlightText QuantityText(Index)
        Cancel = True
        Exit Sub
    End If
    
    If BracketOrderOption.Value Then
        Select Case Index
        Case BracketIndexes.BracketEntryOrder
            QuantityText(BracketIndexes.BracketStopLossOrder) = Quantity
            QuantityText(BracketIndexes.BracketTargetOrder) = Quantity
        Case BracketIndexes.BracketStopLossOrder
            QuantityText(BracketIndexes.BracketEntryOrder) = Quantity
            QuantityText(BracketIndexes.BracketTargetOrder) = Quantity
        Case BracketIndexes.BracketTargetOrder
            QuantityText(BracketIndexes.BracketEntryOrder) = Quantity
            QuantityText(BracketIndexes.BracketStopLossOrder) = Quantity
        End Select
    End If
    
Else
    mBracketOrder.SetNewEntryQuantity Quantity
End If

unHighlightText QuantityText(Index)


Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ResetButton_Click()
Const ProcName As String = "ResetButton_Click"
On Error GoTo Err

clearBracketOrder
setupControls

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub SimpleOrderOption_Click()
Const ProcName As String = "SimpleOrderOption_Click"
On Error GoTo Err

setOrderScheme SimpleOrder

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub SimulateOrdersCheck_Click()
Const ProcName As String = "SimulateOrdersCheck_Click"
On Error GoTo Err

chooseOrderContext

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TriggerOffsetText_KeyDown( _
                Index As Integer, _
                KeyCode As Integer, _
                Shift As Integer)
Const ProcName As String = "TriggerOffsetText_KeyDown"
On Error GoTo Err

If KeyCode = KeyCodeConstants.vbKeyReturn Then
    validateTriggerPriceFields Index
    setCurrentTriggerPriceField Index
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TriggerOffsetText_Validate( _
                Index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "TriggerOffsetText_Validate"
On Error GoTo Err

If Not validateTriggerPriceFields(Index) Then Cancel = True
setCurrentTriggerPriceField Index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TriggerPriceText_KeyDown( _
                Index As Integer, _
                KeyCode As Integer, _
                Shift As Integer)
Const ProcName As String = "TriggerPriceText_KeyDown"
On Error GoTo Err

If KeyCode = KeyCodeConstants.vbKeyReturn Then
    validateTriggerPriceFields Index
    setCurrentTriggerPriceField Index
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TriggerPriceText_Validate( _
                Index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "TriggerPriceText_Validate"
On Error GoTo Err

If Not validateTriggerPriceFields(Index) Then Cancel = True
setCurrentTriggerPriceField Index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TypeCombo_Click(Index As Integer)
Const ProcName As String = "TypeCombo_Click"
On Error GoTo Err

configureOrderFields Index
setCurrentLimitPriceField Index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UndoButton_Click()
Const ProcName As String = "UndoButton_Click"
On Error GoTo Err

mBracketOrder.CancelChanges

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mActiveOrderContext Event Handlers
'@================================================================================

Private Sub mActiveOrderContext_Change(ev As ChangeEventData)
Const ProcName As String = "mActiveOrderContext_Change"
On Error GoTo Err

Dim lChangeType As OrderContextChangeTypes
lChangeType = ev.changeType

Select Case lChangeType
Case OrderContextReadyStateChanged
    If mActiveOrderContext.IsReady Then
        SimpleOrderOption.Enabled = True
        BracketOrderOption.Enabled = True
        setupControls
    Else
        disableAll NotReadyMessage
    End If
Case OrderContextActiveStateChanged

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then Err.Raise ev.Future.ErrorNumber, ev.Future.ErrorSource, ev.Future.ErrorMessage

If ev.Future.IsAvailable Then doSetupControls
Set mFutureBuilder = Nothing
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Property Let MaxDiscretionaryAmountTicks(ByVal Value As Long)
Const ProcName As String = "MaxDiscretionaryAmountTicks"
On Error GoTo Err

AssertArgument Value >= 0, "MaxDiscretionaryAmountTicks"
mMaxDiscretionaryAmountTicks = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MaxDiscretionaryAmountTicks() As Long
MaxDiscretionaryAmountTicks = mMaxDiscretionaryAmountTicks
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let PriceToleranceTicks(ByVal Value As Long)
Const ProcName As String = "PriceToleranceTicks"
On Error GoTo Err

AssertArgument Value >= 0, "PriceToleranceTicks"
mPriceToleranceTicks = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PriceToleranceTicks() As Long
PriceToleranceTicks = mPriceToleranceTicks
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor
UserControl.ForeColor = mTheme.ForeColor
gApplyTheme mTheme, UserControl.Controls

OrderSimulationLabel.ForeColor = mTheme.AlertForeColor

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

If Not mFutureBuilder Is Nothing Then
    mFutureBuilder.Cancel
    Set mFutureBuilder = Nothing
End If

clearOrderContexts
clearDataSource
clearBracketOrder
clearControls
disableAll NoContractMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetLiveOrderContext( _
                ByVal pOrderContext As OrderContext)
Const ProcName As String = "SetLiveOrderContext"
On Error GoTo Err

AssertArgument mMode = OrderTicketModeLiveOnly Or mMode = OrderTicketModeLiveAndSimulated, "LiveOrderContext invalid in this mode"
AssertArgument Not pOrderContext Is Nothing, "LiveOrderContext is Nothing"
AssertArgument Not pOrderContext.IsSimulated, "LiveOrderContext is simulated"
If Not mSimulatedOrderContext Is Nothing Then AssertArgument gGetContractFromContractFuture(mSimulatedOrderContext.ContractFuture).Specifier.Equals(gGetContractFromContractFuture(pOrderContext.ContractFuture).Specifier), "Live and Simulated order contexts must use the same contract"

Set mLiveOrderContext = pOrderContext

setActiveOrderContext mLiveOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetMode(ByVal pMode As OrderTicketModes)
Const ProcName As String = "SetMode"
On Error GoTo Err

setModeNoPromptForOrderContext pMode

chooseOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetSimulatedOrderContext( _
                ByVal pOrderContext As OrderContext)
Const ProcName As String = "SetSimulatedOrderContext"
On Error GoTo Err

AssertArgument mMode = OrderTicketModeSimulatedOnly Or mMode = OrderTicketModeLiveAndSimulated, "SimulatedOrderContext invalid in this mode"
AssertArgument Not pOrderContext Is Nothing, "SimulatedOrderContext is Nothing"
AssertArgument pOrderContext.IsSimulated, "SimulatedOrderContext is not simulated"
If Not mLiveOrderContext Is Nothing Then AssertArgument gGetContractFromContractFuture(pOrderContext.ContractFuture).Specifier.Equals(gGetContractFromContractFuture(mLiveOrderContext.ContractFuture).Specifier), "Live and Simulated order contexts must use the same contract"

Set mSimulatedOrderContext = pOrderContext

setActiveOrderContext mSimulatedOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pRole As BracketOrderRoles)
Const ProcName As String = "ShowBracketOrder"
On Error GoTo Err

clearBracketOrder

Set mBracketOrder = pBracketOrder
If mBracketOrder.IsSimulated Then
    setModeNoPromptForOrderContext OrderTicketModeSimulatedOnly
    SetSimulatedOrderContext mBracketOrder.OrderContext
Else
    setModeNoPromptForOrderContext OrderTicketModeLiveOnly
    SetLiveOrderContext mBracketOrder.OrderContext
End If

Dim lEntryOrder As IOrder
Set lEntryOrder = mBracketOrder.EntryOrder
Set mEntryLimitPriceSpec = lEntryOrder.LimitPriceSpec
Set mEntryTriggerPriceSpec = lEntryOrder.TriggerPriceSpec

Dim lStopLossOrder As IOrder
Set lStopLossOrder = mBracketOrder.StopLossOrder
Set mStopLossLimitPriceSpec = lStopLossOrder.LimitPriceSpec
Set mStopLossTriggerPriceSpec = lStopLossOrder.TriggerPriceSpec

Dim lTargetOrder As IOrder
Set lTargetOrder = mBracketOrder.TargetOrder
Set mTargetLimitPriceSpec = lTargetOrder.LimitPriceSpec
Set mTargetTriggerPriceSpec = lTargetOrder.TriggerPriceSpec

If lStopLossOrder Is Nothing And lTargetOrder Is Nothing Then
    AssertArgument pRole = BracketOrderRoleEntry, "pRole must be BracketOrderRoleEntry for a standalone order"
    BracketTabStrip.Visible = False
    RaiseEvent CaptionChanged("Change a single order")
Else
    BracketTabStrip.Visible = True
    RaiseEvent CaptionChanged("Change a bracket order")
End If

SimpleOrderOption.Enabled = False
BracketOrderOption.Enabled = False

Select Case pRole
Case BracketOrderRoleEntry
    Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabEntryOrder)
Case BracketOrderRoleStopLoss
    Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabStopLossOrder)
Case BracketOrderRoleTarget
    Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabTargetOrder)
Case Else
    AssertArgument False, "Invalid pRole"
End Select

setOrderFieldValues lEntryOrder, BracketIndexes.BracketEntryOrder
setOrderFieldValues lStopLossOrder, BracketIndexes.BracketStopLossOrder
setOrderFieldValues lTargetOrder, BracketIndexes.BracketTargetOrder

configureOrderFields BracketIndexes.BracketEntryOrder
configureOrderFields BracketIndexes.BracketStopLossOrder
configureOrderFields BracketIndexes.BracketTargetOrder

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

mBracketOrder.AddChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addItemToCombo( _
                ByVal pCombo As TWImageCombo, _
                ByVal pItemText As String, _
                ByVal pItemData As Long)
Const ProcName As String = "addItemToCombo"
On Error GoTo Err

pCombo.ComboItems.Add , , pItemText
pCombo.ComboItems(pCombo.ComboItems.Count).Tag = pItemData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub chooseOrderContext()
Const ProcName As String = "chooseOrderContext"
On Error GoTo Err

If SimulateOrdersCheck.Value = vbUnchecked Then
    If mLiveOrderContext Is Nothing Then
        RaiseEvent NeedLiveOrderContext
    Else
        setActiveOrderContext mLiveOrderContext
    End If
Else
    If mSimulatedOrderContext Is Nothing Then
        RaiseEvent NeedSimulatedOrderContext
    Else
        setActiveOrderContext mSimulatedOrderContext
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearBracketOrder()
Const ProcName As String = "clearBracketOrder"
On Error GoTo Err

If mBracketOrder Is Nothing Then Exit Sub

mBracketOrder.RemoveChangeListener Me
Set mBracketOrder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearControls()
Const ProcName As String = "clearControls"
On Error GoTo Err

SymbolLabel.Caption = ""
                        
clearPriceFields

clearOrderFields BracketIndexes.BracketEntryOrder
clearOrderFields BracketIndexes.BracketStopLossOrder
clearOrderFields BracketIndexes.BracketTargetOrder

clearDataSourceValues

OrderSimulationLabel.Caption = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearDataSource()
Const ProcName As String = "clearDataSource"
On Error GoTo Err

If Not mDataSource Is Nothing Then
    mDataSource.RemoveGenericTickListener Me
    mDataSource.RemoveStateChangeListener Me
    Set mDataSource = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearOrderContexts()
Set mActiveOrderContext = Nothing
Set mLiveOrderContext = Nothing
Set mSimulatedOrderContext = Nothing
End Sub

Private Sub clearOrderFields(ByVal pIndex As Long)
Const ProcName As String = "clearOrderFields"
On Error GoTo Err

enableOrderFields pIndex
OrderIdLabel(pIndex) = ""
setComboListIndex ActionCombo(pIndex), 1

QuantityText(pIndex) = 0

' don't set TypeCombo(pIndex) as it will affect other fields and there
' is no sensible value to set it to
LimitPriceText(pIndex) = ""
LimitOffsetText(pIndex) = ""
TriggerPriceText(pIndex) = ""
TriggerOffsetText(pIndex) = ""
If TIFCombo(pIndex).ComboItems.Count <> 0 Then setComboListIndex TIFCombo(pIndex), 1
If TriggerMethodCombo(pIndex).ComboItems.Count <> 0 Then setComboListIndex TriggerMethodCombo(pIndex), 1
IgnoreRthCheck(pIndex) = vbUnchecked
OrderRefText(pIndex) = ""
AllOrNoneCheck(pIndex) = vbUnchecked
BlockOrderCheck(pIndex) = vbUnchecked
ETradeOnlyCheck(pIndex) = vbUnchecked
FirmQuoteOnlyCheck(pIndex) = vbUnchecked
HiddenCheck(pIndex) = vbUnchecked
OverrideCheck(pIndex) = vbUnchecked
SweepToFillCheck(pIndex) = vbUnchecked
DisplaySizeText(pIndex) = ""
MinQuantityText(pIndex) = ""
DiscrAmountText(pIndex) = ""
GoodAfterTimeText(pIndex) = ""
GoodAfterTimeTZText(pIndex) = ""
GoodTillDateText(pIndex) = ""
GoodTillDateTZText(pIndex) = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearPriceFields()
Const ProcName As String = "clearPriceFields"
On Error GoTo Err

LimitPriceText(BracketIndexes.BracketEntryOrder) = ""
LimitOffsetText(BracketIndexes.BracketEntryOrder) = ""
TriggerPriceText(BracketIndexes.BracketEntryOrder) = ""
TriggerOffsetText(BracketIndexes.BracketEntryOrder) = ""
CurrentLimitPriceLabel(BracketIndexes.BracketEntryOrder) = ""
CurrentTriggerPriceLabel(BracketIndexes.BracketEntryOrder) = ""

LimitPriceText(BracketIndexes.BracketStopLossOrder) = ""
LimitOffsetText(BracketIndexes.BracketStopLossOrder) = ""
TriggerPriceText(BracketIndexes.BracketStopLossOrder) = ""
TriggerOffsetText(BracketIndexes.BracketStopLossOrder) = ""
CurrentLimitPriceLabel(BracketIndexes.BracketStopLossOrder) = ""
CurrentTriggerPriceLabel(BracketIndexes.BracketStopLossOrder) = ""

LimitPriceText(BracketIndexes.BracketTargetOrder) = ""
LimitOffsetText(BracketIndexes.BracketTargetOrder) = ""
TriggerPriceText(BracketIndexes.BracketTargetOrder) = ""
TriggerOffsetText(BracketIndexes.BracketTargetOrder) = ""
CurrentLimitPriceLabel(BracketIndexes.BracketTargetOrder) = ""
CurrentTriggerPriceLabel(BracketIndexes.BracketTargetOrder) = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearDataSourceValues()
Const ProcName As String = "clearDataSourceValues"
On Error GoTo Err

AskLabel.Caption = ""
AskSizeLabel.Caption = ""
BidLabel.Caption = ""
BidSizeLabel.Caption = ""
LastLabel.Caption = ""
LastSizeLabel.Caption = ""
VolumeLabel.Caption = ""
HighLabel.Caption = ""
LowLabel.Caption = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearInvalidText( _
                ByVal pText As TextBox, _
                ByVal pIndex As Long)
Const ProcName As String = "clearInvalidText"
On Error GoTo Err

If Not mTheme Is Nothing Then
    If Not mTheme.AlertFont Is Nothing Then Set pText.Font = mTheme.AlertFont
    pText.ForeColor = mTheme.AlertForeColor
Else
    pText.BackColor = ErroredFieldColor
End If
Set mInvalidTexts(pIndex) = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function comboItemData(ByVal pCombo As TWImageCombo) As Long
Const ProcName As String = "comboItemData"
On Error GoTo Err

If pCombo.SelectedItem Is Nothing Then Exit Function
comboItemData = pCombo.SelectedItem.Tag

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub configureOrderFields( _
                ByVal pIndex As Long)
Const ProcName As String = "configureOrderFields"
On Error GoTo Err

Select Case comboItemData(TypeCombo(pIndex))
Case OrderTypeMarket
    configurePriceFields pIndex, False, False
Case OrderTypeMarketOnOpen
    configurePriceFields pIndex, False, False
Case OrderTypeMarketOnClose
    configurePriceFields pIndex, False, False
Case OrderTypeMarketIfTouched
    configurePriceFields pIndex, False, True
Case OrderTypeMarketToLimit
    configurePriceFields pIndex, False, False
Case OrderTypeLimit
    configurePriceFields pIndex, True, False
Case OrderTypeLimitOnOpen
    configurePriceFields pIndex, True, False
Case OrderTypeLimitOnClose
    configurePriceFields pIndex, True, False
Case OrderTypeLimitIfTouched
    configurePriceFields pIndex, True, True
Case OrderTypeStop
    configurePriceFields pIndex, False, True
Case OrderTypeStopLimit
    configurePriceFields pIndex, True, True
Case OrderTypeTrail
    configurePriceFields pIndex, False, True
Case OrderTypeTrailLimit
    configurePriceFields pIndex, True, True
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub configurePriceFields( _
                ByVal pIndex As Long, _
                ByVal pAllowLimitPrice As Boolean, _
                ByVal pAllowTriggerPrice As Boolean)
If pAllowLimitPrice Then
    enableControl LimitPriceText(pIndex)
    enableControl LimitOffsetText(pIndex)
Else
    disableControl LimitPriceText(pIndex)
    disableControl LimitOffsetText(pIndex)
End If

If pAllowTriggerPrice Then
    enableControl TriggerPriceText(pIndex)
    enableControl TriggerOffsetText(pIndex)
Else
    disableControl TriggerPriceText(pIndex)
    disableControl TriggerOffsetText(pIndex)
End If
End Sub

Public Function createPriceSpec( _
                ByVal pPriceText As String, _
                ByVal pOffsetText As String, _
                ByRef pPriceSpec As PriceSpecifier) As Boolean
Const ProcName As String = "createPriceSpec"
On Error GoTo Err

Dim lPriceString As String
lPriceString = pPriceText
If pOffsetText <> "" Then lPriceString = lPriceString & "[" & pOffsetText & "]"
createPriceSpec = mActiveOrderContext.ParsePriceAndOffset(pPriceSpec, lPriceString)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub disableAll( _
                ByVal pMessage As String)
Const ProcName As String = "disableAll"
On Error GoTo Err

SimpleOrderOption.Enabled = False
BracketOrderOption.Enabled = False

PlaceOrdersButton.Enabled = False
CompleteOrdersButton.Enabled = False
ResetButton.Enabled = False
ModifyButton.Enabled = False
UndoButton.Enabled = False

disableOrderFields BracketIndexes.BracketEntryOrder
disableOrderFields BracketIndexes.BracketStopLossOrder
disableOrderFields BracketIndexes.BracketTargetOrder

SymbolLabel.Caption = ""
AskLabel = ""
AskSizeLabel = ""
BidLabel = ""
BidSizeLabel = ""
LastLabel = ""
LastSizeLabel = ""
VolumeLabel = ""
HighLabel = ""
LowLabel = ""

OrderSimulationLabel = pMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableControl(ByVal pField As Control)
Const ProcName As String = "disableControl"
On Error GoTo Err

pField.Enabled = False
If TypeOf pField Is CheckBox Or _
    TypeOf pField Is OptionButton Then Exit Sub
    
If mTheme Is Nothing Then
    pField.BackColor = SystemColorConstants.vbButtonFace
Else
    pField.BackColor = mTheme.DisabledBackColor
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableOrderFields(ByVal pIndex As Long)
Const ProcName As String = "disableOrderFields"
On Error GoTo Err

disableControl ActionCombo(pIndex)
disableControl QuantityText(pIndex)
disableControl TypeCombo(pIndex)
disableControl LimitPriceText(pIndex)
disableControl LimitOffsetText(pIndex)
disableControl TriggerPriceText(pIndex)
disableControl TriggerOffsetText(pIndex)
disableControl IgnoreRthCheck(pIndex)
disableControl TIFCombo(pIndex)
disableControl OrderRefText(pIndex)
disableControl AllOrNoneCheck(pIndex)
disableControl BlockOrderCheck(pIndex)
disableControl ETradeOnlyCheck(pIndex)
disableControl FirmQuoteOnlyCheck(pIndex)
disableControl HiddenCheck(pIndex)
disableControl OverrideCheck(pIndex)
disableControl SweepToFillCheck(pIndex)
disableControl DisplaySizeText(pIndex)
disableControl MinQuantityText(pIndex)
disableControl TriggerMethodCombo(pIndex)
disableControl DiscrAmountText(pIndex)
disableControl GoodAfterTimeText(pIndex)
disableControl GoodAfterTimeTZText(pIndex)
disableControl GoodTillDateText(pIndex)
disableControl GoodTillDateTZText(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doSetupControls()
Const ProcName As String = "doSetupControls"
On Error GoTo Err

SymbolLabel.Caption = mContract.Specifier.LocalSymbol & _
                        " on " & _
                        mContract.Specifier.Exchange
                        
setupTifCombo BracketIndexes.BracketEntryOrder
setupTifCombo BracketIndexes.BracketStopLossOrder
setupTifCombo BracketIndexes.BracketTargetOrder

setupTriggerMethodCombo BracketIndexes.BracketEntryOrder
setupTriggerMethodCombo BracketIndexes.BracketStopLossOrder
setupTriggerMethodCombo BracketIndexes.BracketTargetOrder

setupTypeCombo BracketIndexes.BracketEntryOrder
setupTypeCombo BracketIndexes.BracketStopLossOrder
setupTypeCombo BracketIndexes.BracketTargetOrder

reset

setQuantity BracketIndexes.BracketEntryOrder
setQuantity BracketIndexes.BracketStopLossOrder
setQuantity BracketIndexes.BracketTargetOrder

showDataSourceValues

If mActiveOrderContext.IsSimulated Then
    OrderSimulationLabel.Caption = OrdersSimulatedMessage
Else
    OrderSimulationLabel.Caption = OrdersLiveMessage
End If

ActionCombo(0).SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableControl(ByVal pField As Control)
Const ProcName As String = "enableControl"
On Error GoTo Err

pField.Enabled = True
If TypeOf pField Is CheckBox Or _
    TypeOf pField Is OptionButton Then Exit Sub
    
If mTheme Is Nothing Then
    pField.BackColor = SystemColorConstants.vbWindowBackground
Else
    pField.BackColor = mTheme.TextBackColor
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableOrderFields(ByVal pIndex As Long)
Const ProcName As String = "enableOrderFields"
On Error GoTo Err

If pIndex = BracketIndexes.BracketEntryOrder Then enableControl ActionCombo(pIndex)
enableControl QuantityText(pIndex)
enableControl TypeCombo(pIndex)
enableControl LimitPriceText(pIndex)
enableControl LimitOffsetText(pIndex)
enableControl TriggerPriceText(pIndex)
enableControl TriggerOffsetText(pIndex)
enableControl IgnoreRthCheck(pIndex)
enableControl TIFCombo(pIndex)
enableControl OrderRefText(pIndex)
enableControl AllOrNoneCheck(pIndex)
enableControl BlockOrderCheck(pIndex)
enableControl ETradeOnlyCheck(pIndex)
enableControl FirmQuoteOnlyCheck(pIndex)
enableControl HiddenCheck(pIndex)
enableControl OverrideCheck(pIndex)
enableControl SweepToFillCheck(pIndex)
enableControl DisplaySizeText(pIndex)
enableControl MinQuantityText(pIndex)
enableControl TriggerMethodCombo(pIndex)
enableControl DiscrAmountText(pIndex)
enableControl GoodAfterTimeText(pIndex)
enableControl GoodAfterTimeTZText(pIndex)
enableControl GoodTillDateText(pIndex)
enableControl GoodTillDateTZText(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getPrice( _
                ByVal pPriceString As String) As Double
Const ProcName As String = "getPrice"
On Error GoTo Err

Dim Price As Double
priceFromString pPriceString, Price
getPrice = Price

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub highlightText( _
                ByVal pText As TextBox)
Const ProcName As String = "highlightText"
On Error GoTo Err

If Not mTheme Is Nothing Then
    If Not mTheme.AlertFont Is Nothing Then Set pText.Font = mTheme.AlertFont
    pText.ForeColor = mTheme.AlertForeColor
Else
    pText.BackColor = ErroredFieldColor
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isOrderModifiable(ByVal pOrder As IOrder) As Boolean
Const ProcName As String = "isOrderModifiable"
On Error GoTo Err

If pOrder Is Nothing Then Exit Function
isOrderModifiable = pOrder.IsModifiable

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidPrice( _
                ByVal pPriceString As String, ByVal pBasePrice As Double) As Boolean
Const ProcName As String = "isValidPrice"
On Error GoTo Err

Dim lPrice As Double
If Not priceFromString(pPriceString, lPrice) Then Exit Function

If pBasePrice = 0# Then
    isValidPrice = True
Else
    isValidPrice = Int(Abs(lPrice - pBasePrice) / mContract.TickSize) <= mPriceToleranceTicks
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidOrder( _
                ByVal pIndex As Long) As Boolean
Const ProcName As String = "isValidOrder"
On Error GoTo Err

If Not mInvalidTexts(pIndex) Is Nothing Then unHighlightText mInvalidTexts(pIndex)

If comboItemData(ActionCombo(pIndex)) = OrderActions.OrderActionNone Then
    isValidOrder = True
    Exit Function
End If

Select Case pIndex
Case BracketEntryOrder
    If Not IsInteger(QuantityText(pIndex), 0) Then setInvalidText QuantityText(pIndex), pIndex: Exit Function
    If QuantityText(pIndex) = 0 And mBracketOrder Is Nothing Then setInvalidText QuantityText(pIndex), pIndex: Exit Function
    
    Select Case comboItemData(TypeCombo(pIndex))
    Case OrderTypeMarket, _
            OrderTypeMarketOnOpen, _
            OrderTypeMarketOnClose, _
            OrderTypeMarketToLimit
        ' other field values don't matter
    Case OrderTypeMarketIfTouched, _
            OrderTypeStop, _
            OrderTypeTrail
        If Not mEntryTriggerPriceSpec.IsValid Then
            setInvalidText TriggerPriceText(pIndex), pIndex
            setInvalidText TriggerOffsetText(pIndex), pIndex
            Exit Function
        End If
    Case OrderTypeLimit, _
            OrderTypeLimitOnOpen, _
            OrderTypeLimitOnClose
        If Not mEntryLimitPriceSpec.IsValid Then
            setInvalidText LimitPriceText(pIndex), pIndex
            setInvalidText LimitOffsetText(pIndex), pIndex
            Exit Function
        End If
    Case OrderTypeLimitIfTouched, _
            OrderTypeStopLimit, _
            OrderTypeTrailLimit
        If Not mEntryLimitPriceSpec.IsValid Then
            setInvalidText LimitPriceText(pIndex), pIndex
            setInvalidText LimitOffsetText(pIndex), pIndex
            Exit Function
        End If
        If Not mEntryTriggerPriceSpec.IsValid Then
            setInvalidText TriggerPriceText(pIndex), pIndex
            setInvalidText TriggerOffsetText(pIndex), pIndex
            Exit Function
        End If
    End Select
Case BracketStopLossOrder
    If comboItemData(TypeCombo(pIndex)) = OrderTypeNone Then
        isValidOrder = True
        Exit Function
    End If
    
    If Not IsInteger(QuantityText(pIndex), 1) Then setInvalidText QuantityText(pIndex), pIndex: Exit Function
    
    Select Case comboItemData(TypeCombo(pIndex))
    Case OrderTypeStop, _
            OrderTypeTrail
        If Not mStopLossTriggerPriceSpec.IsValid Then
            setInvalidText TriggerPriceText(pIndex), pIndex
            setInvalidText TriggerOffsetText(pIndex), pIndex
            Exit Function
        End If
    Case OrderTypeStopLimit, _
            OrderTypeTrailLimit
        If Not mStopLossLimitPriceSpec.IsValid Then
            setInvalidText LimitPriceText(pIndex), pIndex
            setInvalidText LimitOffsetText(pIndex), pIndex
            Exit Function
        End If
        If Not mStopLossTriggerPriceSpec.IsValid Then
            setInvalidText TriggerPriceText(pIndex), pIndex
            setInvalidText TriggerOffsetText(pIndex), pIndex
            Exit Function
        End If
    End Select
Case BracketTargetOrder
    If comboItemData(TypeCombo(pIndex)) = OrderTypeNone Then
        isValidOrder = True
        Exit Function
    End If
    
    If Not IsInteger(QuantityText(pIndex), 1) Then setInvalidText QuantityText(pIndex), pIndex: Exit Function
    
    Select Case comboItemData(TypeCombo(pIndex))
    Case OrderTypeLimit
        If Not mTargetLimitPriceSpec.IsValid Then
            setInvalidText LimitPriceText(pIndex), pIndex
            setInvalidText LimitOffsetText(pIndex), pIndex
            Exit Function
        End If
    Case OrderTypeLimitIfTouched
        If Not mTargetLimitPriceSpec.IsValid Then
            setInvalidText LimitPriceText(pIndex), pIndex
            setInvalidText LimitOffsetText(pIndex), pIndex
            Exit Function
        End If
        If Not mTargetTriggerPriceSpec.IsValid Then
            setInvalidText LimitPriceText(pIndex), pIndex
            setInvalidText LimitOffsetText(pIndex), pIndex
            Exit Function
        End If
    Case OrderTypeMarketIfTouched
        If Not mTargetTriggerPriceSpec.IsValid Then
            setInvalidText LimitPriceText(pIndex), pIndex
            setInvalidText LimitOffsetText(pIndex), pIndex
            Exit Function
        End If
    End Select
End Select

If DisplaySizeText(pIndex) <> "" Then
    If Not IsInteger(DisplaySizeText(pIndex), 1) Then setInvalidText DisplaySizeText(pIndex), pIndex: Exit Function
End If

If MinQuantityText(pIndex) <> "" Then
    If Not IsInteger(MinQuantityText(pIndex), 1) Then setInvalidText MinQuantityText(pIndex), pIndex: Exit Function
End If

If DiscrAmountText(pIndex) <> "" Then
    If Not isValidPrice(DiscrAmountText(pIndex), 0#) Then setInvalidText DiscrAmountText(pIndex), pIndex: Exit Function
    Dim lPrice As Double
    priceFromString DiscrAmountText(pIndex), lPrice
    If Int(Abs(lPrice) / mContract.TickSize) > mMaxDiscretionaryAmountTicks Then setInvalidText DiscrAmountText(pIndex), pIndex: Exit Function
End If

If GoodAfterTimeText(pIndex) <> "" Then
    If Not IsDate(GoodAfterTimeText(pIndex)) Then setInvalidText GoodAfterTimeText(pIndex), pIndex
End If

If GoodTillDateText(pIndex) <> "" Then
    If Not IsDate(GoodTillDateText(pIndex)) Then setInvalidText GoodTillDateText(pIndex), pIndex
End If

isValidOrder = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadOrderFields(ByVal pIndex As Long)
Const ProcName As String = "loadOrderFields"
On Error GoTo Err

Load OrderIdLabel(pIndex)
Load ActionCombo(pIndex)
Load QuantityText(pIndex)
Load TypeCombo(pIndex)
Load LimitPriceText(pIndex)
Load LimitOffsetText(pIndex)
Load CurrentLimitPriceLabel(pIndex)
Load TriggerPriceText(pIndex)
Load TriggerOffsetText(pIndex)
Load CurrentTriggerPriceLabel(pIndex)
Load IgnoreRthCheck(pIndex)
Load TIFCombo(pIndex)
Load OrderRefText(pIndex)
Load AllOrNoneCheck(pIndex)
Load BlockOrderCheck(pIndex)
Load ETradeOnlyCheck(pIndex)
Load FirmQuoteOnlyCheck(pIndex)
Load HiddenCheck(pIndex)
Load OverrideCheck(pIndex)
Load SweepToFillCheck(pIndex)
Load DisplaySizeText(pIndex)
Load MinQuantityText(pIndex)
Load TriggerMethodCombo(pIndex)
Load DiscrAmountText(pIndex)
Load GoodAfterTimeText(pIndex)
Load GoodAfterTimeTZText(pIndex)
Load GoodTillDateText(pIndex)
Load GoodTillDateTZText(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function priceFromString( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "priceFromString"
On Error GoTo Err

priceFromString = ParsePrice(pPriceString, mContract.Specifier.secType, mContract.TickSize, pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function priceToString(ByVal pPrice As Double) As String
Const ProcName As String = "priceToString"
On Error GoTo Err

priceToString = FormatPrice(pPrice, mContract.Specifier.secType, mContract.TickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub reset()
Const ProcName As String = "reset"
On Error GoTo Err

clearOrderFields BracketIndexes.BracketEntryOrder
clearOrderFields BracketIndexes.BracketStopLossOrder
clearOrderFields BracketIndexes.BracketTargetOrder

SimpleOrderOption.Enabled = True
BracketOrderOption.Enabled = True
BracketOrderOption.Value = True

setOrderScheme OrderSchemes.BracketOrder

selectComboEntry ActionCombo(BracketIndexes.BracketEntryOrder), _
                OrderActions.OrderActionBuy
setAction BracketIndexes.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndexes.BracketEntryOrder), _
                OrderTypes.OrderTypeLimit
setOrderFieldsEnabling BracketIndexes.BracketEntryOrder, Nothing
configureOrderFields BracketIndexes.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndexes.BracketStopLossOrder), _
                OrderTypes.OrderTypeStop
setOrderFieldsEnabling BracketIndexes.BracketStopLossOrder, Nothing
configureOrderFields BracketIndexes.BracketStopLossOrder

selectComboEntry TypeCombo(BracketIndexes.BracketTargetOrder), _
                OrderTypes.OrderTypeNone
setOrderFieldsEnabling BracketIndexes.BracketTargetOrder, Nothing
configureOrderFields BracketIndexes.BracketTargetOrder

Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabEntryOrder)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub selectComboEntry( _
                ByVal pCombo As TWImageCombo, _
                ByVal ItemData As Long)
Const ProcName As String = "selectComboEntry"
On Error GoTo Err

Dim i As Long
For i = 1 To pCombo.ComboItems.Count
    If pCombo.ComboItems(i).Tag = ItemData Then
        Set pCombo.SelectedItem = pCombo.ComboItems(i)
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setAction( _
                ByVal pIndex As Long)
Const ProcName As String = "setAction"
On Error GoTo Err

If BracketOrderOption.Value And pIndex = BracketIndexes.BracketEntryOrder Then
    If comboItemData(ActionCombo(pIndex)) = OrderActions.OrderActionSell Then
        selectComboEntry ActionCombo(BracketIndexes.BracketStopLossOrder), OrderActions.OrderActionBuy
        selectComboEntry ActionCombo(BracketIndexes.BracketTargetOrder), OrderActions.OrderActionBuy
    Else
        selectComboEntry ActionCombo(BracketIndexes.BracketStopLossOrder), OrderActions.OrderActionSell
        selectComboEntry ActionCombo(BracketIndexes.BracketTargetOrder), OrderActions.OrderActionSell
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setActiveOrderContext(ByVal pOrderContext As OrderContext)
Const ProcName As String = "setActiveOrderContext"
On Error GoTo Err

If Not mFutureBuilder Is Nothing Then
    mFutureBuilder.Cancel
    Set mFutureBuilder = Nothing
End If

If pOrderContext Is mActiveOrderContext Then Exit Sub

If Not mDataSource Is Nothing Then
    mDataSource.RemoveGenericTickListener Me
    mDataSource.RemoveStateChangeListener Me
End If

Set mActiveOrderContext = pOrderContext
Set mContract = gGetContractFromContractFuture(mActiveOrderContext.ContractFuture)

Set mDataSource = mActiveOrderContext.DataSource
If Not mDataSource Is Nothing Then
    If Not mDataSource.IsMarketDataRequested Then mDataSource.StartMarketData
    mDataSource.AddGenericTickListener Me
    mDataSource.AddStateChangeListener Me
    If ticksAvailable Then
        mAskPrice = mDataSource.CurrentTick(TickTypeAsk).Price
        mBidPrice = mDataSource.CurrentTick(TickTypeBid).Price
        mTradePrice = mDataSource.CurrentTick(TickTypeTrade).Price
    End If
End If

If mActiveOrderContext.IsReady Then
    setupControls
Else
    disableAll NotReadyMessage
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setComboListIndex(ByVal pCombo As TWImageCombo, ByVal pListIndex As Long)
Const ProcName As String = "setComboListIndex"
On Error GoTo Err

Set pCombo.SelectedItem = pCombo.ComboItems(pListIndex)
pCombo.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCurrentLimitPriceField( _
                pIndex As Integer)
Const ProcName As String = "setCurrentLimitPriceField"
On Error GoTo Err

Dim lOrderType As OrderTypes: lOrderType = comboItemData(TypeCombo(pIndex))
If lOrderType = OrderTypeNone Then
    CurrentLimitPriceLabel(pIndex) = ""
    Exit Sub
End If

Dim lOrderAction As OrderActions
lOrderAction = comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder))

Dim lPrice As Double

Select Case pIndex
Case BracketIndexes.BracketEntryOrder
    If mEntryLimitPriceSpec Is Nothing Then CurrentLimitPriceLabel(pIndex) = "": Exit Sub
    lPrice = mActiveOrderContext.CalculateOffsettedPrice( _
                                                    mEntryLimitPriceSpec, _
                                                    mContract.Specifier.secType, _
                                                    lOrderAction)
Case BracketIndexes.BracketStopLossOrder
    If mStopLossLimitPriceSpec Is Nothing Then CurrentLimitPriceLabel(pIndex) = "": Exit Sub
    lPrice = mActiveOrderContext.CalculateOffsettedPrice( _
                                                    mStopLossLimitPriceSpec, _
                                                    mContract.Specifier.secType, _
                                                    lOrderAction)
Case BracketIndexes.BracketTargetOrder
    If mTargetLimitPriceSpec Is Nothing Then CurrentLimitPriceLabel(pIndex) = "": Exit Sub
    lPrice = mActiveOrderContext.CalculateOffsettedPrice( _
                                                    mTargetLimitPriceSpec, _
                                                    mContract.Specifier.secType, _
                                                    lOrderAction)
End Select

CurrentLimitPriceLabel(pIndex) = IIf(lPrice = MaxDouble, _
                                    "N/A", _
                                    FormatPrice(lPrice, _
                                                mContract.Specifier.secType, _
                                                mContract.TickSize))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCurrentTriggerPriceField( _
                pIndex As Integer)
Const ProcName As String = "setCurrentTriggerPriceField"
On Error GoTo Err

Dim lOrderType As OrderTypes: lOrderType = comboItemData(TypeCombo(pIndex))
If lOrderType = OrderTypeNone Then
    CurrentTriggerPriceLabel(pIndex) = ""
    Exit Sub
End If

Dim lOrderAction As OrderActions
lOrderAction = comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder))

Dim lPrice As Double

Select Case pIndex
Case BracketIndexes.BracketEntryOrder
    If mEntryTriggerPriceSpec Is Nothing Then CurrentTriggerPriceLabel(pIndex) = "": Exit Sub
    lPrice = mActiveOrderContext.CalculateOffsettedPrice( _
                                                    mEntryTriggerPriceSpec, _
                                                    mContract.Specifier.secType, _
                                                    lOrderAction)
Case BracketIndexes.BracketStopLossOrder
    If mStopLossTriggerPriceSpec Is Nothing Then CurrentTriggerPriceLabel(pIndex) = "": Exit Sub
    lPrice = mActiveOrderContext.CalculateOffsettedPrice( _
                                                    mStopLossTriggerPriceSpec, _
                                                    mContract.Specifier.secType, _
                                                    lOrderAction)
Case BracketIndexes.BracketTargetOrder
    If mTargetTriggerPriceSpec Is Nothing Then CurrentTriggerPriceLabel(pIndex) = "": Exit Sub
    lPrice = mActiveOrderContext.CalculateOffsettedPrice( _
                                                    mTargetTriggerPriceSpec, _
                                                    mContract.Specifier.secType, _
                                                    lOrderAction)
End Select

CurrentTriggerPriceLabel(pIndex) = IIf(lPrice = MaxDouble, _
                                    "N/A", _
                                    FormatPrice(lPrice, _
                                                mContract.Specifier.secType, _
                                                mContract.TickSize))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInvalidText( _
                ByVal pText As TextBox, _
                ByVal pIndex As Long)
Const ProcName As String = "setInvalidText"
On Error GoTo Err

Set mInvalidTexts(pIndex) = pText
highlightText pText

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setModeNoPromptForOrderContext(ByVal pMode As OrderTicketModes)
Const ProcName As String = "setModeNoPromptForOrderContext"
On Error GoTo Err

mMode = pMode

Select Case mMode
Case OrderTicketModeLiveOnly
    SimulateOrdersCheck.Value = vbUnchecked
    SimulateOrdersCheck.Visible = False
Case OrderTicketModeSimulatedOnly
    SimulateOrdersCheck.Value = vbChecked
    SimulateOrdersCheck.Visible = False
Case OrderTicketModeLiveAndSimulated
    SimulateOrdersCheck.Value = vbUnchecked
    SimulateOrdersCheck.Visible = True
Case Else
    AssertArgument False, "Invalid mode"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'/**
' Sets the attributes for an order from the specified fields on the control
'
' @param pOrder     the <code>order</code> whose attributes are to be set
' @param pIndex the index of the order page whose fields are the source of
'                   the attribute values
'
'*/
Private Sub setOrderAttributes( _
                ByVal pOrder As IOrder, _
                ByVal pIndex As Long)
Const ProcName As String = "setOrderAttributes"
On Error GoTo Err

With pOrder
    If pOrder.IsAttributeModifiable(OrderAttAllOrNone) Then .AllOrNone = (AllOrNoneCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttBlockOrder) Then .BlockOrder = (BlockOrderCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttDiscretionaryAmount) Then .DiscretionaryAmount = IIf(DiscrAmountText(pIndex) = "", 0, DiscrAmountText(pIndex))
    If pOrder.IsAttributeModifiable(OrderAttDisplaySize) Then .displaySize = IIf(DisplaySizeText(pIndex) = "", 0, DisplaySizeText(pIndex))
    If pOrder.IsAttributeModifiable(OrderAttETradeOnly) Then .ETradeOnly = (ETradeOnlyCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttFirmQuoteOnly) Then .FirmQuoteOnly = (FirmQuoteOnlyCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttGoodAfterTime) Then If GoodAfterTimeText(pIndex) <> "" Then .GoodAfterTime = CDate(GoodAfterTimeText(pIndex))
    If pOrder.IsAttributeModifiable(OrderAttGoodAfterTimeTZ) Then .GoodAfterTimeTZ = GoodAfterTimeTZText(pIndex)
    If pOrder.IsAttributeModifiable(OrderAttGoodTillDate) Then If GoodTillDateText(pIndex) <> "" Then .GoodTillDate = CDate(GoodTillDateText(pIndex))
    If pOrder.IsAttributeModifiable(OrderAttGoodTillDateTZ) Then .GoodTillDateTZ = GoodTillDateTZText(pIndex)
    If pOrder.IsAttributeModifiable(OrderAttHidden) Then .Hidden = (HiddenCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttIgnoreRTH) Then .IgnoreRegularTradingHours = (IgnoreRthCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttMinimumQuantity) Then .MinimumQuantity = IIf(MinQuantityText(pIndex) = "", 0, MinQuantityText(pIndex))
    If pOrder.IsAttributeModifiable(OrderAttOriginatorRef) Then .OriginatorRef = OrderRefText(pIndex)
    If pOrder.IsAttributeModifiable(OrderAttOverrideConstraints) Then .OverrideConstraints = (OverrideCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttQuantity) Then .Quantity = QuantityText(pIndex)
    If pOrder.IsAttributeModifiable(OrderAttStopTriggerMethod) Then .StopTriggerMethod = comboItemData(TriggerMethodCombo(pIndex))
    If pOrder.IsAttributeModifiable(OrderAttSweepToFill) Then .SweepToFill = (SweepToFillCheck(pIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttTimeInForce) Then .TimeInForce = comboItemData(TIFCombo(pIndex))
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderFieldValues( _
                ByVal pOrder As IOrder, _
                ByVal pIndex As Long)
Const ProcName As String = "setOrderFieldValues"
On Error GoTo Err

If pOrder Is Nothing Then
    disableOrderFields pIndex
    Exit Sub
End If

clearOrderFields pIndex

With pOrder
    setOrderId pIndex, .Id
    
    selectComboEntry ActionCombo(pIndex), .Action
    QuantityText(pIndex) = .Quantity
    selectComboEntry TypeCombo(pIndex), .OrderType
    LimitPriceText(pIndex) = IIf(.LimitPrice <> MaxDouble, .LimitPrice, "")
    TriggerPriceText(pIndex) = IIf(.TriggerPrice <> MaxDouble, .TriggerPrice, "")
    IgnoreRthCheck(pIndex) = IIf(.IgnoreRegularTradingHours, vbChecked, vbUnchecked)
    selectComboEntry TIFCombo(pIndex), .TimeInForce
    OrderRefText(pIndex) = .OriginatorRef
    AllOrNoneCheck(pIndex) = IIf(.AllOrNone, vbChecked, vbUnchecked)
    BlockOrderCheck(pIndex) = IIf(.BlockOrder, vbChecked, vbUnchecked)
    ETradeOnlyCheck(pIndex) = IIf(.ETradeOnly, vbChecked, vbUnchecked)
    FirmQuoteOnlyCheck(pIndex) = IIf(.FirmQuoteOnly, vbChecked, vbUnchecked)
    HiddenCheck(pIndex) = IIf(.Hidden, vbChecked, vbUnchecked)
    OverrideCheck(pIndex) = IIf(.OverrideConstraints, vbChecked, vbUnchecked)
    SweepToFillCheck(pIndex) = IIf(.SweepToFill, vbChecked, vbUnchecked)
    DisplaySizeText(pIndex) = IIf(.displaySize <> 0, .displaySize, "")
    MinQuantityText(pIndex) = IIf(.MinimumQuantity <> 0, .displaySize, "")
    If .StopTriggerMethod <> 0 Then TriggerMethodCombo(pIndex) = OrderStopTriggerMethodToString(.StopTriggerMethod)
    DiscrAmountText(pIndex) = IIf(.DiscretionaryAmount <> 0, .DiscretionaryAmount, "")
    GoodAfterTimeText(pIndex) = IIf(.GoodAfterTime <> 0, FormatDateTime(.GoodAfterTime, vbGeneralDate), "")
    GoodAfterTimeTZText(pIndex) = .GoodAfterTimeTZ
    GoodTillDateText(pIndex) = IIf(.GoodTillDate <> 0, FormatDateTime(.GoodTillDate, vbGeneralDate), "")
    GoodTillDateTZText(pIndex) = .GoodTillDateTZ
    
    ' do this last because it sets the various fields attributes
    selectComboEntry TypeCombo(pIndex), .OrderType
End With

If Not isOrderModifiable(pOrder) Then
    disableOrderFields pIndex
Else
    setOrderFieldsEnabling pIndex, pOrder
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderFieldEnabling( _
                ByVal pControl As Control, _
                ByVal pOrderAtt As OrderAttributes, _
                ByVal pOrder As IOrder)
Const ProcName As String = "setOrderFieldEnabling"
On Error GoTo Err

If Not pOrder Is Nothing Then
    If pOrder.IsAttributeModifiable(pOrderAtt) Then
        enableControl pControl
    Else
        disableControl pControl
    End If
ElseIf mActiveOrderContext.IsOrderAttributeSupported(pOrderAtt) Then
    enableControl pControl
Else
    disableControl pControl
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderFieldsEnabling( _
                ByVal pIndex As Long, _
                ByVal pOrder As IOrder)
Const ProcName As String = "setOrderFieldsEnabling"
On Error GoTo Err

If pIndex = BracketIndexes.BracketEntryOrder Then setOrderFieldEnabling ActionCombo(pIndex), OrderAttAction, pOrder
setOrderFieldEnabling QuantityText(pIndex), OrderAttQuantity, pOrder
setOrderFieldEnabling TypeCombo(pIndex), OrderAttOrderType, pOrder
setOrderFieldEnabling LimitPriceText(pIndex), OrderAttLimitPrice, pOrder
setOrderFieldEnabling LimitOffsetText(pIndex), OrderAttLimitPrice, pOrder
setOrderFieldEnabling TriggerPriceText(pIndex), OrderAttTriggerPrice, pOrder
setOrderFieldEnabling TriggerOffsetText(pIndex), OrderAttTriggerPrice, pOrder
setOrderFieldEnabling IgnoreRthCheck(pIndex), OrderAttIgnoreRTH, pOrder
setOrderFieldEnabling TIFCombo(pIndex), OrderAttTimeInForce, pOrder
setOrderFieldEnabling OrderRefText(pIndex), OrderAttOriginatorRef, pOrder
setOrderFieldEnabling AllOrNoneCheck(pIndex), OrderAttAllOrNone, pOrder
setOrderFieldEnabling BlockOrderCheck(pIndex), OrderAttBlockOrder, pOrder
setOrderFieldEnabling ETradeOnlyCheck(pIndex), OrderAttETradeOnly, pOrder
setOrderFieldEnabling FirmQuoteOnlyCheck(pIndex), OrderAttFirmQuoteOnly, pOrder
setOrderFieldEnabling HiddenCheck(pIndex), OrderAttHidden, pOrder
setOrderFieldEnabling OverrideCheck(pIndex), OrderAttOverrideConstraints, pOrder
setOrderFieldEnabling SweepToFillCheck(pIndex), OrderAttSweepToFill, pOrder
setOrderFieldEnabling DisplaySizeText(pIndex), OrderAttDisplaySize, pOrder
setOrderFieldEnabling MinQuantityText(pIndex), OrderAttMinimumQuantity, pOrder
setOrderFieldEnabling TriggerMethodCombo(pIndex), OrderAttStopTriggerMethod, pOrder
setOrderFieldEnabling DiscrAmountText(pIndex), OrderAttDiscretionaryAmount, pOrder
setOrderFieldEnabling GoodAfterTimeText(pIndex), OrderAttGoodAfterTime, pOrder
setOrderFieldEnabling GoodAfterTimeTZText(pIndex), OrderAttGoodAfterTimeTZ, pOrder
setOrderFieldEnabling GoodTillDateText(pIndex), OrderAttGoodTillDate, pOrder
setOrderFieldEnabling GoodTillDateTZText(pIndex), OrderAttGoodTillDateTZ, pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderId( _
                ByVal pIndex As Long, _
                ByVal pId As String)
Const ProcName As String = "setOrderId"
On Error GoTo Err

OrderIdLabel(pIndex).Caption = pId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderScheme( _
                ByVal pOrderScheme As OrderSchemes)
Const ProcName As String = "setOrderScheme"
On Error GoTo Err

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
    
Case OrderSchemes.BracketOrder
    RaiseEvent CaptionChanged("Create a bracket order")
    BracketTabStrip.Visible = True
    PlaceOrdersButton.Enabled = True
    PlaceOrdersButton.Visible = True
    CompleteOrdersButton.Visible = False
    ModifyButton.Visible = False
    UndoButton.Visible = False
    ResetButton.Enabled = True
    ResetButton.Enabled = True
    Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabEntryOrder)
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPriceFields()
Const ProcName As String = "setPriceFields"
On Error GoTo Err

setCurrentLimitPriceField BracketIndexes.BracketEntryOrder
setCurrentTriggerPriceField BracketIndexes.BracketEntryOrder
setCurrentLimitPriceField BracketIndexes.BracketStopLossOrder
setCurrentTriggerPriceField BracketIndexes.BracketStopLossOrder
setCurrentLimitPriceField BracketIndexes.BracketTargetOrder
setCurrentTriggerPriceField BracketIndexes.BracketTargetOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setQuantity(ByVal pIndex As Long)
Select Case mContract.Specifier.secType
Case SecTypeStock
    QuantityText(pIndex) = 100
Case SecTypeFuture
    QuantityText(pIndex) = 1
Case SecTypeOption
    QuantityText(pIndex) = 1
Case SecTypeFuturesOption
    QuantityText(pIndex) = 1
Case SecTypeCash
    QuantityText(pIndex) = 25000
Case SecTypeCombo
    QuantityText(pIndex) = 1
Case SecTypeIndex
    QuantityText(pIndex) = 0
End Select
End Sub

Private Sub setupActionCombo(ByVal pIndex As Long)
Const ProcName As String = "setupActionCombo"
On Error GoTo Err

ActionCombo(pIndex).ComboItems.Clear
If pIndex <> BracketIndexes.BracketEntryOrder Then
    addItemToCombo ActionCombo(pIndex), _
                OrderActionToString(OrderActions.OrderActionNone), _
                OrderActions.OrderActionNone
    disableControl ActionCombo(pIndex)
End If
addItemToCombo ActionCombo(pIndex), _
            OrderActionToString(OrderActions.OrderActionBuy), _
            OrderActions.OrderActionBuy
addItemToCombo ActionCombo(pIndex), _
            OrderActionToString(OrderActions.OrderActionSell), _
            OrderActions.OrderActionSell

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupControls()
Const ProcName As String = "setupControls"
On Error GoTo Err

If ticksAvailable Then
    doSetupControls
ElseIf mFutureBuilder Is Nothing Then
    Set mFutureBuilder = New FutureBuilder
    Set mFutureWaiter = New FutureWaiter
    mFutureWaiter.Add mFutureBuilder.Future
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ticksAvailable() As Boolean
Const ProcName As String = "ticksAvailable"
On Error GoTo Err

ticksAvailable = mDataSource.HasCurrentTick(TickTypeAsk) And _
                mDataSource.HasCurrentTick(TickTypeBid) And _
                mDataSource.HasCurrentTick(TickTypeTrade)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupTifCombo(ByVal pIndex As Long)
Const ProcName As String = "setupTifCombo"
On Error GoTo Err

TIFCombo(pIndex).ComboItems.Clear

If mActiveOrderContext.IsOrderTifSupported(OrderTIFs.OrderTIFDay) Then
    addItemToCombo TIFCombo(pIndex), _
                OrderTIFToString(OrderTIFs.OrderTIFDay), _
                OrderTIFs.OrderTIFDay
End If
If mActiveOrderContext.IsOrderTifSupported(OrderTIFs.OrderTIFGoodTillCancelled) Then
    addItemToCombo TIFCombo(pIndex), _
                OrderTIFToString(OrderTIFs.OrderTIFGoodTillCancelled), _
                OrderTIFs.OrderTIFGoodTillCancelled
End If
If mActiveOrderContext.IsOrderTifSupported(OrderTIFs.OrderTIFImmediateOrCancel) Then
    addItemToCombo TIFCombo(pIndex), _
                OrderTIFToString(OrderTIFs.OrderTIFImmediateOrCancel), _
                OrderTIFs.OrderTIFImmediateOrCancel
End If

setComboListIndex TIFCombo(pIndex), 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTriggerMethodCombo(ByVal pIndex As Long)
Const ProcName As String = "setupTriggerMethodCombo"
On Error GoTo Err

TriggerMethodCombo(pIndex).ComboItems.Clear

If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerDefault) Then
    addItemToCombo TriggerMethodCombo(pIndex), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerDefault), _
                OrderStopTriggerMethods.OrderStopTriggerDefault
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerLast) Then
    addItemToCombo TriggerMethodCombo(pIndex), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerLast), _
                OrderStopTriggerMethods.OrderStopTriggerLast
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerBidAsk) Then
    addItemToCombo TriggerMethodCombo(pIndex), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerBidAsk), _
                OrderStopTriggerMethods.OrderStopTriggerBidAsk
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk) Then
    addItemToCombo TriggerMethodCombo(pIndex), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk), _
                OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerDoubleLast) Then
    addItemToCombo TriggerMethodCombo(pIndex), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerDoubleLast), _
                OrderStopTriggerMethods.OrderStopTriggerDoubleLast
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk) Then
    addItemToCombo TriggerMethodCombo(pIndex), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk), _
                OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerMidPoint) Then
    addItemToCombo TriggerMethodCombo(pIndex), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerMidPoint), _
                OrderStopTriggerMethods.OrderStopTriggerMidPoint
End If

setComboListIndex TriggerMethodCombo(pIndex), 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTypeCombo(ByVal pIndex As Long)
Const ProcName As String = "setupTypeCombo"
On Error GoTo Err

TypeCombo(pIndex).ComboItems.Clear

If pIndex = BracketIndexes.BracketEntryOrder Then
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimit) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeLimit), _
                    OrderTypes.OrderTypeLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarket) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeMarket), _
                    OrderTypes.OrderTypeMarket
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStop) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeStop), _
                    OrderTypes.OrderTypeStop
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStopLimit) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeStopLimit), _
                    OrderTypes.OrderTypeStopLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStop) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeTrail), _
                    OrderTypes.OrderTypeTrail
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStopLimit) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeTrailLimit), _
                    OrderTypes.OrderTypeTrailLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimitOnOpen) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeLimitOnOpen), _
                    OrderTypes.OrderTypeLimitOnOpen
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketOnOpen) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeMarketOnOpen), _
                    OrderTypes.OrderTypeMarketOnOpen
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimitOnClose) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeLimitOnClose), _
                    OrderTypes.OrderTypeLimitOnClose
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketOnClose) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeMarketOnClose), _
                    OrderTypes.OrderTypeMarketOnClose
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimitIfTouched) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeLimitIfTouched), _
                    OrderTypes.OrderTypeLimitIfTouched
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketIfTouched) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeMarketIfTouched), _
                    OrderTypes.OrderTypeMarketIfTouched
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketToLimit) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeMarketToLimit), _
                    OrderTypes.OrderTypeMarketToLimit
    End If
ElseIf pIndex = BracketIndexes.BracketStopLossOrder Then
    addItemToCombo TypeCombo(pIndex), _
                OrderTypeToString(OrderTypes.OrderTypeNone), _
                OrderTypes.OrderTypeNone
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStop) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeStop), _
                    OrderTypes.OrderTypeStop
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStopLimit) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeStopLimit), _
                    OrderTypes.OrderTypeStopLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStop) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeTrail), _
                    OrderTypes.OrderTypeTrail
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStopLimit) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeTrailLimit), _
                    OrderTypes.OrderTypeTrailLimit
    End If
ElseIf pIndex = BracketIndexes.BracketTargetOrder Then
    addItemToCombo TypeCombo(pIndex), _
                OrderTypeToString(OrderTypes.OrderTypeNone), _
                OrderTypes.OrderTypeNone
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimit) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeLimit), _
                    OrderTypes.OrderTypeLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketIfTouched) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeMarketIfTouched), _
                    OrderTypes.OrderTypeMarketIfTouched
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketIfTouched) Then
        addItemToCombo TypeCombo(pIndex), _
                    OrderTypeToString(OrderTypes.OrderTypeLimitIfTouched), _
                    OrderTypes.OrderTypeLimitIfTouched
    End If
End If

setComboListIndex TypeCombo(pIndex), 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showOrderFields(ByVal pIndex As Long)
Const ProcName As String = "showOrderFields"
On Error GoTo Err

Dim i As Long
Dim lVisible As Boolean
For i = 0 To ActionCombo.Count - 1
    lVisible = (i = pIndex)
    OrderIdLabel(i).Visible = lVisible
    ActionCombo(i).Visible = lVisible
    QuantityText(i).Visible = lVisible
    TypeCombo(i).Visible = lVisible
    LimitPriceText(i).Visible = lVisible
    LimitOffsetText(i).Visible = lVisible
    CurrentLimitPriceLabel(i).Visible = lVisible
    TriggerPriceText(i).Visible = lVisible
    TriggerOffsetText(i).Visible = lVisible
    CurrentTriggerPriceLabel(i).Visible = lVisible
    IgnoreRthCheck(i).Visible = lVisible
    TIFCombo(i).Visible = lVisible
    OrderRefText(i).Visible = lVisible
    AllOrNoneCheck(i).Visible = lVisible
    BlockOrderCheck(i).Visible = lVisible
    ETradeOnlyCheck(i).Visible = lVisible
    FirmQuoteOnlyCheck(i).Visible = lVisible
    HiddenCheck(i).Visible = lVisible
    OverrideCheck(i).Visible = lVisible
    SweepToFillCheck(i).Visible = lVisible
    DisplaySizeText(i).Visible = lVisible
    MinQuantityText(i).Visible = lVisible
    TriggerMethodCombo(i).Visible = lVisible
    DiscrAmountText(i).Visible = lVisible
    GoodAfterTimeText(i).Visible = lVisible
    GoodAfterTimeTZText(i).Visible = lVisible
    GoodTillDateText(i).Visible = lVisible
    GoodTillDateTZText(i).Visible = lVisible
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showDataSourceValues()
Const ProcName As String = "showDataSourceValues"
On Error GoTo Err

If mDataSource Is Nothing Then Exit Sub

If mDataSource.HasCurrentTick(TickTypeAsk) Then
    AskLabel.Caption = priceToString(mDataSource.CurrentTick(TickTypeAsk).Price)
    AskSizeLabel.Caption = mDataSource.CurrentTick(TickTypeAsk).Size
End If

If mDataSource.HasCurrentTick(TickTypeBid) Then
    BidLabel.Caption = priceToString(mDataSource.CurrentTick(TickTypeBid).Price)
    BidSizeLabel.Caption = mDataSource.CurrentTick(TickTypeBid).Size
End If

If mDataSource.HasCurrentTick(TickTypeTrade) Then
    LastLabel.Caption = priceToString(mDataSource.CurrentTick(TickTypeTrade).Price)
    LastSizeLabel.Caption = mDataSource.CurrentTick(TickTypeTrade).Size
End If

If mDataSource.HasCurrentTick(TickTypeVolume) Then
    VolumeLabel.Caption = mDataSource.CurrentTick(TickTypeVolume).Size
End If

If mDataSource.HasCurrentTick(TickTypeHighPrice) Then
    HighLabel.Caption = priceToString(mDataSource.CurrentTick(TickTypeHighPrice).Price)
End If

If mDataSource.HasCurrentTick(TickTypeLowPrice) Then
    LowLabel.Caption = priceToString(mDataSource.CurrentTick(TickTypeLowPrice).Price)
End If

setPriceFields

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unHighlightText( _
                ByVal pText As TextBox)
Const ProcName As String = "unHighlightText"
On Error GoTo Err

If mTheme Is Nothing Then
    pText.BackColor = vbButtonFace
Else
    gApplyThemeToControl mTheme, pText
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function validateLimitPriceFields( _
                ByVal pIndex As Integer) As Boolean
Const ProcName As String = "validateLimitPriceFields"
On Error GoTo Err

If LimitPriceText(pIndex) = "" And LimitOffsetText(pIndex) = "" Then
    ' allow blank price to prevent user irritation if they place the caret
    ' in the limit price or offset field when the order type is limit, and then
    ' decide they want to change the order type - if space is not allowed then they
    ' would have to enter a valid price before being able to get to the order
    ' type combo
    validateLimitPriceFields = True
    Exit Function
End If

If comboItemData(ActionCombo(pIndex)) = OrderActions.OrderActionNone And _
    (LimitPriceText(pIndex) <> "" Or LimitOffsetText(pIndex) <> "") _
Then
    Exit Function
End If
    
Dim lPriceSpec As PriceSpecifier
If Not createPriceSpec(LimitPriceText(pIndex), LimitOffsetText(pIndex), lPriceSpec) Then
    highlightText LimitPriceText(pIndex)
    highlightText LimitOffsetText(pIndex)
    Exit Function
End If
unHighlightText LimitPriceText(pIndex)
unHighlightText LimitOffsetText(pIndex)

Select Case pIndex
Case BracketIndexes.BracketEntryOrder
    Set mEntryLimitPriceSpec = lPriceSpec
    If Not mBracketOrder Is Nothing Then mBracketOrder.SetNewEntryLimitPrice lPriceSpec
Case BracketIndexes.BracketStopLossOrder
    Set mStopLossLimitPriceSpec = lPriceSpec
    If Not mBracketOrder Is Nothing Then mBracketOrder.SetNewStopLossLimitPrice lPriceSpec
Case BracketIndexes.BracketTargetOrder
    Set mTargetLimitPriceSpec = lPriceSpec
    If Not mBracketOrder Is Nothing Then mBracketOrder.SetNewTargetLimitPrice lPriceSpec
End Select

validateLimitPriceFields = True

Exit Function

Err:
gNotifyUnhandledError ProcName, ModuleName
End Function

Private Function validateTriggerPriceFields( _
                ByVal pIndex As Integer) As Boolean
Const ProcName As String = "validateTriggerPriceFields"
On Error GoTo Err

If TriggerPriceText(pIndex) = "" And TriggerOffsetText(pIndex) = "" Then
    ' allow blank price to prevent user irritation if they place the caret
    ' in the trigger price or offset field when the order type is limit, and then
    ' decide they want to change the order type - if space is not allowed then they
    ' would have to enter a valid price before being able to get to the order
    ' type combo
    validateTriggerPriceFields = True
    Exit Function
End If

If comboItemData(ActionCombo(pIndex)) = OrderActions.OrderActionNone And _
    (TriggerPriceText(pIndex) <> "" Or TriggerOffsetText(pIndex) <> "") _
Then
    Exit Function
End If
    
Dim lPriceSpec As PriceSpecifier
If Not createPriceSpec(TriggerPriceText(pIndex), TriggerOffsetText(pIndex), lPriceSpec) Then
    highlightText TriggerPriceText(pIndex)
    highlightText TriggerOffsetText(pIndex)
    Exit Function
End If
unHighlightText TriggerPriceText(pIndex)
unHighlightText TriggerOffsetText(pIndex)

Select Case pIndex
Case BracketIndexes.BracketEntryOrder
    Set mEntryTriggerPriceSpec = lPriceSpec
    If Not mBracketOrder Is Nothing Then mBracketOrder.SetNewEntryTriggerPrice lPriceSpec
Case BracketIndexes.BracketStopLossOrder
    Set mStopLossTriggerPriceSpec = lPriceSpec
    If Not mBracketOrder Is Nothing Then mBracketOrder.SetNewStopLossTriggerPrice lPriceSpec
Case BracketIndexes.BracketTargetOrder
    Set mTargetTriggerPriceSpec = lPriceSpec
    If Not mBracketOrder Is Nothing Then mBracketOrder.SetNewTargetTriggerPrice lPriceSpec
End Select

validateTriggerPriceFields = True

Exit Function

Err:
gNotifyUnhandledError ProcName, ModuleName
End Function




