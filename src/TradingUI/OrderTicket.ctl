VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
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
      TabIndex        =   31
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton SimpleOrderOption 
      Caption         =   "&Simple order"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox SimulateOrdersCheck 
      Caption         =   "S&imulate orders"
      Height          =   195
      Left            =   3480
      TabIndex        =   32
      Top             =   120
      Width           =   1455
   End
   Begin TWControls40.TWButton UndoButton 
      Height          =   495
      Left            =   7560
      TabIndex        =   29
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
      TabIndex        =   24
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
      TabIndex        =   27
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
      TabIndex        =   25
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
      TabIndex        =   28
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
      TabIndex        =   26
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
      TabIndex        =   73
      Top             =   795
      Width           =   8775
      Begin VB.Frame Frame2 
         Caption         =   "Ticker"
         Height          =   1815
         Left            =   120
         TabIndex        =   53
         Top             =   3120
         Width           =   3135
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   105
            ScaleHeight     =   1455
            ScaleWidth      =   2655
            TabIndex        =   54
            Top             =   240
            Width           =   2655
            Begin VB.TextBox VolumeText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox HighText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox LowText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox LastSizeText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox AskSizeText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   0
               Width           =   735
            End
            Begin VB.TextBox BidSizeText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox BidText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox LastText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox AskText 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label22 
               Caption         =   "Bid"
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Ask"
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label11 
               Caption         =   "Last"
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label25 
               Caption         =   "Volume"
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label24 
               Caption         =   "High"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label23 
               Caption         =   "Low"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   1200
               Width           =   855
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Order"
         Height          =   2895
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   3135
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   2535
            Left            =   105
            ScaleHeight     =   2535
            ScaleWidth      =   2895
            TabIndex        =   44
            Top             =   240
            Width           =   2895
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
            Begin VB.TextBox StopPriceText 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   5
               Top             =   2160
               Width           =   855
            End
            Begin VB.TextBox OffsetText 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   4
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
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1800
               Width           =   855
            End
            Begin VB.TextBox PriceText 
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
            Begin VB.Label OrderIdLabel 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   72
               Top             =   0
               Width           =   2535
            End
            Begin VB.Label Label8 
               Caption         =   "Offset (ticks)"
               Height          =   255
               Left            =   0
               TabIndex        =   52
               Top             =   1800
               Width           =   975
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
               Caption         =   "Stop price"
               Height          =   255
               Left            =   0
               TabIndex        =   50
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Price"
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
         Height          =   4815
         Left            =   3360
         TabIndex        =   33
         Top             =   120
         Width           =   3975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   120
            ScaleHeight     =   4455
            ScaleWidth      =   3735
            TabIndex        =   34
            Top             =   240
            Width           =   3735
            Begin TWControls40.TWImageCombo TriggerMethodCombo 
               Height          =   330
               Index           =   0
               Left            =   1200
               TabIndex        =   16
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
               TabIndex        =   6
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
               TabIndex        =   7
               Top             =   0
               Width           =   1215
            End
            Begin VB.TextBox OrderRefText 
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   8
               Top             =   360
               Width           =   2535
            End
            Begin VB.CheckBox OverrideCheck 
               Caption         =   "Override"
               Height          =   255
               Index           =   0
               Left            =   2400
               TabIndex        =   22
               Top             =   3000
               Width           =   1335
            End
            Begin VB.TextBox MinQuantityText 
               Height          =   285
               Index           =   0
               Left            =   2760
               TabIndex        =   14
               Top             =   1440
               Width           =   975
            End
            Begin VB.CheckBox FirmQuoteOnlyCheck 
               Caption         =   "Firm quote only"
               Height          =   255
               Index           =   0
               Left            =   2400
               TabIndex        =   20
               Top             =   2760
               Width           =   1410
            End
            Begin VB.CheckBox ETradeOnlyCheck 
               Caption         =   "eTrade only"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   19
               Top             =   2760
               Width           =   1215
            End
            Begin VB.CheckBox AllOrNoneCheck 
               Caption         =   "All or none"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   17
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox GoodTillDateTZText 
               Height          =   285
               Index           =   0
               Left            =   2760
               TabIndex        =   12
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox GoodAfterTimeTZText 
               Height          =   285
               Index           =   0
               Left            =   2760
               TabIndex        =   10
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox GoodTillDateText 
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   11
               Top             =   1080
               Width           =   1575
            End
            Begin VB.TextBox GoodAfterTimeText 
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   9
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox DiscrAmountText 
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   15
               Top             =   1800
               Width           =   735
            End
            Begin VB.CheckBox HiddenCheck 
               Caption         =   "Hidden"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   21
               Top             =   3000
               Width           =   855
            End
            Begin VB.TextBox DisplaySizeText 
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   13
               Top             =   1440
               Width           =   735
            End
            Begin VB.CheckBox SweepToFillCheck 
               Caption         =   "SweepToFill"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   23
               Top             =   3240
               Width           =   1215
            End
            Begin VB.CheckBox BlockOrderCheck 
               Caption         =   "Block order"
               Height          =   255
               Index           =   0
               Left            =   2400
               TabIndex        =   18
               Top             =   2520
               Width           =   1335
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Min qty"
               Height          =   375
               Left            =   2040
               TabIndex        =   42
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "Good till date"
               Height          =   255
               Left            =   0
               TabIndex        =   41
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label21 
               Caption         =   "Good after time"
               Height          =   255
               Left            =   0
               TabIndex        =   40
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "Discr amount"
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label17 
               Caption         =   "Trigger method"
               Height          =   255
               Left            =   0
               TabIndex        =   38
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "Display size"
               Height          =   255
               Left            =   0
               TabIndex        =   37
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Order ref"
               Height          =   255
               Left            =   0
               TabIndex        =   36
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label10 
               Caption         =   "TIF"
               Height          =   255
               Left            =   0
               TabIndex        =   35
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
         TabIndex        =   74
         Top             =   5010
         Width           =   7215
      End
   End
   Begin MSComctlLib.TabStrip BracketTabStrip 
      Height          =   5760
      Left            =   0
      TabIndex        =   70
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
      TabIndex        =   71
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
Implements IThemeable
Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

Event CaptionChanged(ByVal caption As String)
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

Private mCurrentBracketOrderIndex                       As BracketIndexes

Private mInvalidControls(2)                             As Control

Private mMode                                           As OrderTicketModes

Private mTheme                                          As ITheme

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

BracketOrderOption.value = True
setOrderScheme BracketOrder

loadOrderFields BracketIndexes.BracketStopOrder
loadOrderFields BracketIndexes.BracketTargetOrder

setupActionCombo BracketIndexes.BracketEntryOrder
setupActionCombo BracketIndexes.BracketStopOrder
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
' ChangeListener Interface Members
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
Case BracketOrderChangeTypes.BracketOrderStopOrderChanged
    If op.StopLossOrder.Status = OrderStatusFilled Then disableOrderFields BracketIndexes.BracketStopOrder
    setOrderFieldValues op.StopLossOrder, BracketIndexes.BracketStopOrder
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
    BidText = lPriceText
    BidSizeText = ev.Tick.Size
    setPriceFields
Case TickTypeAsk
    AskText = lPriceText
    AskSizeText = ev.Tick.Size
    setPriceFields
Case TickTypeHighPrice
    HighText = lPriceText
Case TickTypeLowPrice
    LowText = lPriceText
Case TickTypeTrade
    LastText = lPriceText
    LastSizeText = ev.Tick.Size
    setPriceFields
Case TickTypeVolume
    VolumeText = ev.Tick.Size
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

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

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

Private Sub ActionCombo_Click(ByRef index As Integer)
Const ProcName As String = "ActionCombo_Click"
On Error GoTo Err

setAction index

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

mCurrentBracketOrderIndex = BracketTabStrip.SelectedItem.index - 1
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

Private Sub ModifyButton_Click()
Const ProcName As String = "ModifyButton_Click"
On Error GoTo Err

If Not isValidOrder(BracketEntryOrder) Then Exit Sub
setOrderAttributes mBracketOrder.EntryOrder, BracketIndexes.BracketEntryOrder
If Not mBracketOrder.StopLossOrder Is Nothing Then
    If Not isValidOrder(BracketStopOrder) Then Exit Sub
    setOrderAttributes mBracketOrder.StopLossOrder, BracketIndexes.BracketStopOrder
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

Private Sub OffsetText_Change(index As Integer)
Const ProcName As String = "OffsetText_Change"
On Error GoTo Err

If IsNumeric(OffsetText(index)) Then
    OffsetValueText(index) = OffsetText(index) * mContract.TickSize
Else
    OffsetValueText(index) = ""
End If
setPriceField index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub PlaceOrdersButton_Click()
Const ProcName As String = "PlaceOrdersButton_Click"
On Error GoTo Err

Dim op As IBracketOrder

If SimpleOrderOption.value Then
    If Not isValidOrder(BracketEntryOrder) Then Exit Sub
    
    If comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder)) = OrderActions.OrderActionBuy Then
        Set op = mActiveOrderContext.CreateBuyBracketOrder( _
                                    QuantityText(BracketIndexes.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketEntryOrder) = "", 0, OffsetText(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketEntryOrder)), _
                                    BracketStopLossTypes.BracketStopLossTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    BracketTargetTypes.BracketTargetTypeNone, _
                                    0, _
                                    0, _
                                    0)
    Else
        Set op = mActiveOrderContext.CreateSellBracketOrder( _
                                    QuantityText(BracketIndexes.BracketEntryOrder), _
                                    comboItemData(TypeCombo(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(PriceText(BracketIndexes.BracketEntryOrder)), _
                                    IIf(OffsetText(BracketIndexes.BracketEntryOrder) = "", 0, OffsetText(BracketIndexes.BracketEntryOrder)), _
                                    getPrice(StopPriceText(BracketIndexes.BracketEntryOrder)), _
                                    BracketStopLossTypes.BracketStopLossTypeNone, _
                                    0, _
                                    0, _
                                    0, _
                                    BracketTargetTypes.BracketTargetTypeNone, _
                                    0, _
                                    0, _
                                    0)
        
    End If
    
    setOrderAttributes op.EntryOrder, BracketIndexes.BracketEntryOrder
    mActiveOrderContext.ExecuteBracketOrder op
ElseIf BracketOrderOption.value Then
    If Not isValidOrder(BracketEntryOrder) Then Exit Sub
    If Not isValidOrder(BracketStopOrder) Then Exit Sub
    If Not isValidOrder(BracketTargetOrder) Then Exit Sub
    
    If comboItemData(ActionCombo(BracketIndexes.BracketEntryOrder)) = OrderActions.OrderActionBuy Then
        Set op = mActiveOrderContext.CreateBuyBracketOrder( _
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
        Set op = mActiveOrderContext.CreateSellBracketOrder( _
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
    
    setOrderAttributes op.EntryOrder, BracketIndexes.BracketEntryOrder
    If Not op.StopLossOrder Is Nothing Then
        setOrderAttributes op.StopLossOrder, BracketIndexes.BracketStopOrder
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

Private Sub PriceText_Validate( _
                index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "PriceText_Validate"
On Error GoTo Err

' allow blank Price to prevent user irritation if they place the caret
' in the Price field when the order type is limit, and then decide they
' want to change the order type - if space is not allowed then they
' would have to enter a valid Price before being able to get to the order
' type combo
If PriceText(index) = "" Then Exit Sub

Dim lPrice As Double
If (comboItemData(ActionCombo(index)) = OrderActions.OrderActionNone And _
        PriceText(index) <> "" _
    ) Or _
    Not priceFromString(PriceText(index), lPrice) Or _
    lPrice <= 0 _
Then
    Cancel = True
    Exit Sub
End If

If Not mBracketOrder Is Nothing Then
    Select Case index
    Case BracketIndexes.BracketEntryOrder
        mBracketOrder.SetNewEntryPrice lPrice
    Case BracketIndexes.BracketStopOrder
        mBracketOrder.SetNewStopLossPrice lPrice
    Case BracketIndexes.BracketTargetOrder
        mBracketOrder.SetNewTargetPrice lPrice
    End Select
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub QuantityText_Validate( _
                index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "QuantityText_Validate"
On Error GoTo Err

If comboItemData(ActionCombo(index)) <> OrderActions.OrderActionNone And _
    Not IsNumeric(QuantityText(index)) _
Then
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

If Not IsInteger(QuantityText(index), min, max) Then
    Cancel = True
    Exit Sub
End If

Dim Quantity As Long
Quantity = CLng(QuantityText(index))

If mBracketOrder Is Nothing Then
    If Quantity = 0 Then
        Cancel = True
        Exit Sub
    End If
    
    If BracketOrderOption.value Then
        Select Case index
        Case BracketIndexes.BracketEntryOrder
            QuantityText(BracketIndexes.BracketStopOrder) = Quantity
            QuantityText(BracketIndexes.BracketTargetOrder) = Quantity
        Case BracketIndexes.BracketStopOrder
            QuantityText(BracketIndexes.BracketEntryOrder) = Quantity
            QuantityText(BracketIndexes.BracketTargetOrder) = Quantity
        Case BracketIndexes.BracketTargetOrder
            QuantityText(BracketIndexes.BracketEntryOrder) = Quantity
            QuantityText(BracketIndexes.BracketStopOrder) = Quantity
        End Select
    End If
    
Else
    mBracketOrder.SetNewQuantity Quantity
End If

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

Private Sub StopPriceText_Validate( _
                index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "StopPriceText_Validate"
On Error GoTo Err

Dim lPrice As Double
If (comboItemData(ActionCombo(index)) = OrderActions.OrderActionNone And _
        StopPriceText(index) <> "" _
    ) Or _
    Not priceFromString(StopPriceText(index), lPrice) Or _
    lPrice < 0 _
Then
    Cancel = True
    Exit Sub
End If

If Not mBracketOrder Is Nothing Then
    Select Case index
    Case BracketIndexes.BracketEntryOrder
        mBracketOrder.SetNewEntryTriggerPrice lPrice
    Case BracketIndexes.BracketStopOrder
        mBracketOrder.SetNewStopLossTriggerPrice lPrice
    Case BracketIndexes.BracketTargetOrder
        mBracketOrder.SetNewTargetTriggerPrice lPrice
    End Select
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TypeCombo_Click(index As Integer)
Const ProcName As String = "TypeCombo_Click"
On Error GoTo Err

configureOrderFields index
setPriceField index

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
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
UserControl.BackColor = mTheme.BackColor
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

Dim lStopOrder As IOrder
Set lStopOrder = mBracketOrder.StopLossOrder

Dim lTargetOrder As IOrder
Set lTargetOrder = mBracketOrder.TargetOrder

If lStopOrder Is Nothing And lTargetOrder Is Nothing Then
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
    Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabStopOrder)
Case BracketOrderRoleTarget
    Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabTargetOrder)
Case Else
    AssertArgument False, "Invalid pRole"
End Select

setOrderFieldValues lEntryOrder, BracketIndexes.BracketEntryOrder
setOrderFieldValues lStopOrder, BracketIndexes.BracketStopOrder
setOrderFieldValues lTargetOrder, BracketIndexes.BracketTargetOrder

configureOrderFields BracketIndexes.BracketEntryOrder
configureOrderFields BracketIndexes.BracketStopOrder
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
                ByVal combo As TWImageCombo, _
                ByVal itemText As String, _
                ByVal ItemData As Long)
Const ProcName As String = "addItemToCombo"
On Error GoTo Err

combo.ComboItems.Add , , itemText
combo.ComboItems(combo.ComboItems.Count).Tag = ItemData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub chooseOrderContext()
Const ProcName As String = "chooseOrderContext"
On Error GoTo Err

If SimulateOrdersCheck.value = vbUnchecked Then
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

SymbolLabel.caption = ""
                        
clearPriceFields

clearOrderFields BracketIndexes.BracketEntryOrder
clearOrderFields BracketIndexes.BracketStopOrder
clearOrderFields BracketIndexes.BracketTargetOrder

clearDataSourceValues

OrderSimulationLabel.caption = ""

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

Private Sub clearOrderFields(ByVal index As Long)
Const ProcName As String = "clearOrderFields"
On Error GoTo Err

enableOrderFields index
OrderIdLabel(index) = ""
setComboListIndex ActionCombo(index), 1

QuantityText(index) = 0

' don't set TypeCombo(Index) as it will affect other fields and there
' is no sensible value to set it to
PriceText(index) = ""
StopPriceText(index) = ""
OffsetText(index) = ""
If TIFCombo(index).ComboItems.Count <> 0 Then setComboListIndex TIFCombo(index), 1
If TriggerMethodCombo(index).ComboItems.Count <> 0 Then setComboListIndex TriggerMethodCombo(index), 1
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearPriceFields()
Const ProcName As String = "clearPriceFields"
On Error GoTo Err

PriceText(BracketIndexes.BracketEntryOrder) = ""
PriceText(BracketIndexes.BracketStopOrder) = ""
PriceText(BracketIndexes.BracketTargetOrder) = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearDataSourceValues()
Const ProcName As String = "clearDataSourceValues"
On Error GoTo Err

AskText.Text = ""
AskSizeText.Text = ""
BidText.Text = ""
BidSizeText.Text = ""
LastText.Text = ""
LastSizeText.Text = ""
VolumeText.Text = ""
HighText.Text = ""
LowText.Text = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function comboItemData(ByVal combo As TWImageCombo) As Long
Const ProcName As String = "comboItemData"
On Error GoTo Err

If combo.SelectedItem Is Nothing Then Exit Function
comboItemData = combo.SelectedItem.Tag

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub configureOrderFields( _
                ByVal orderIndex As Long)
Const ProcName As String = "configureOrderFields"
On Error GoTo Err

Select Case orderIndex
Case BracketIndexes.BracketEntryOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case BracketEntryTypeMarket
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeMarketOnOpen
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeMarketOnClose
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeMarketIfTouched
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeMarketToLimit
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketEntryTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketEntryTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketEntryTypeLimit
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeLimitOnOpen
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeLimitOnClose
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeLimitIfTouched
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeStop
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketEntryTypeStopLimit
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    End Select
Case BracketIndexes.BracketStopOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case BracketStopLossTypeNone
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketStopLossTypeStop
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketStopLossTypeStopLimit
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketStopLossTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketStopLossTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketStopLossTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketStopLossTypeAuto
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    End Select
Case BracketIndexes.BracketTargetOrder
    Select Case comboItemData(TypeCombo(orderIndex))
    Case BracketTargetTypeNone
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketTargetTypeLimit
        enableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketTargetTypeLimitIfTouched
        enableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketTargetTypeMarketIfTouched
        disableControl PriceText(orderIndex)
        enableControl StopPriceText(orderIndex)
        disableControl OffsetText(orderIndex)
    Case BracketTargetTypeBid
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketTargetTypeAsk
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketTargetTypeLast
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    Case BracketTargetTypeAuto
        disableControl PriceText(orderIndex)
        disableControl StopPriceText(orderIndex)
        enableControl OffsetText(orderIndex)
    End Select
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableAll( _
                ByVal message As String)
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableControl(ByVal field As Control)
Const ProcName As String = "disableControl"
On Error GoTo Err

field.Enabled = False
If TypeOf field Is CheckBox Or _
    TypeOf field Is OptionButton Then Exit Sub
    
If mTheme Is Nothing Then
    field.BackColor = SystemColorConstants.vbButtonFace
Else
    field.BackColor = mTheme.DisabledBackColor
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableOrderFields(ByVal index As Long)
Const ProcName As String = "disableOrderFields"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableControl(ByVal field As Control)
Const ProcName As String = "enableControl"
On Error GoTo Err

field.Enabled = True
If TypeOf field Is CheckBox Or _
    TypeOf field Is OptionButton Then Exit Sub
    
If mTheme Is Nothing Then
    field.BackColor = SystemColorConstants.vbWindowBackground
Else
    field.BackColor = mTheme.TextBackColor
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableOrderFields(ByVal index As Long)
Const ProcName As String = "enableOrderFields"
On Error GoTo Err

If index = BracketIndexes.BracketEntryOrder Then enableControl ActionCombo(index)
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getPrice( _
                ByVal priceString As String) As Double
Const ProcName As String = "getPrice"
On Error GoTo Err

Dim Price As Double
priceFromString priceString, Price
getPrice = Price

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
                ByVal priceString As String) As Boolean
Const ProcName As String = "isValidPrice"
On Error GoTo Err

Dim Price As Double
isValidPrice = priceFromString(priceString, Price)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidOrder( _
                ByVal index As Long) As Boolean
Const ProcName As String = "isValidOrder"
On Error GoTo Err

If Not mInvalidControls(index) Is Nothing Then
    If mTheme Is Nothing Then
        mInvalidControls(index).BackColor = vbButtonFace
    Else
        mInvalidControls(index).BackColor = mTheme.DisabledBackColor
    End If
End If

If comboItemData(ActionCombo(index)) = OrderActions.OrderActionNone Then
    isValidOrder = True
    Exit Function
End If

Select Case index
Case BracketEntryOrder
    If Not IsInteger(QuantityText(index), 0) Then setInvalidControl QuantityText(index), index: Exit Function
    If QuantityText(index) = 0 And mBracketOrder Is Nothing Then setInvalidControl QuantityText(index), index: Exit Function
    
    Select Case comboItemData(TypeCombo(index))
    Case BracketEntryTypeMarket, BracketEntryTypeMarketOnOpen, BracketEntryTypeMarketOnClose
        ' other field values don't matter
    Case BracketEntryTypeMarketIfTouched, BracketEntryTypeStop
        If Not isValidPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
    Case BracketEntryTypeMarketToLimit, BracketEntryTypeLimit, BracketEntryTypeLimitOnOpen, BracketEntryTypeLimitOnClose
        If Not isValidPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case BracketEntryTypeBid, BracketEntryTypeAsk, BracketEntryTypeLast
        If OffsetText(index) <> "" Then
            If Not IsInteger(OffsetText(index), -100, 100) Then setInvalidControl OffsetText(index), index: Exit Function
        End If
    Case BracketEntryTypeLimitIfTouched, BracketEntryTypeStopLimit
        If Not isValidPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
        If Not isValidPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    End Select
Case BracketStopOrder
    If comboItemData(TypeCombo(index)) = BracketStopLossTypeNone Then
        isValidOrder = True
        Exit Function
    End If
    
    If Not IsInteger(QuantityText(index), 1) Then setInvalidControl QuantityText(index), index: Exit Function
    
    Select Case comboItemData(TypeCombo(index))
    Case BracketStopLossTypeStop
        If Not isValidPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
    Case BracketStopLossTypeStopLimit
        If Not isValidPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
        If Not isValidPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case BracketStopLossTypeBid, BracketStopLossTypeAsk, BracketStopLossTypeLast, BracketStopLossTypeAuto
        If OffsetText(index) <> "" Then
            If Not IsInteger(OffsetText(index), -100, 100) Then setInvalidControl OffsetText(index), index: Exit Function
        End If
    End Select
Case BracketTargetOrder
    If comboItemData(TypeCombo(index)) = BracketStopLossTypeNone Then
        isValidOrder = True
        Exit Function
    End If
    
    If Not IsInteger(QuantityText(index), 1) Then setInvalidControl QuantityText(index), index: Exit Function
    
    Select Case comboItemData(TypeCombo(index))
    Case BracketTargetTypeLimit
        If Not isValidPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case BracketTargetTypeLimitIfTouched
        If Not isValidPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
        If Not isValidPrice(PriceText(index)) Then setInvalidControl PriceText(index), index: Exit Function
    Case BracketTargetTypeMarketIfTouched
        If Not isValidPrice(StopPriceText(index)) Then setInvalidControl StopPriceText(index), index: Exit Function
    Case BracketTargetTypeBid, BracketTargetTypeAsk, BracketTargetTypeLast, BracketTargetTypeAuto
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
    If Not isValidPrice(DiscrAmountText(index)) Then setInvalidControl DiscrAmountText(index), index: Exit Function
End If

isValidOrder = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadOrderFields(ByVal index As Long)
Const ProcName As String = "loadOrderFields"
On Error GoTo Err

Load OrderIdLabel(index)
Load ActionCombo(index)
Load QuantityText(index)
Load TypeCombo(index)
Load PriceText(index)
Load StopPriceText(index)
Load IgnoreRthCheck(index)
Load OffsetText(index)
Load OffsetValueText(index)
Load TIFCombo(index)
Load OrderRefText(index)
Load AllOrNoneCheck(index)
Load BlockOrderCheck(index)
Load ETradeOnlyCheck(index)
Load FirmQuoteOnlyCheck(index)
Load HiddenCheck(index)
Load OverrideCheck(index)
Load SweepToFillCheck(index)
Load DisplaySizeText(index)
Load MinQuantityText(index)
Load TriggerMethodCombo(index)
Load DiscrAmountText(index)
Load GoodAfterTimeText(index)
Load GoodAfterTimeTZText(index)
Load GoodTillDateText(index)
Load GoodTillDateTZText(index)

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
clearOrderFields BracketIndexes.BracketStopOrder
clearOrderFields BracketIndexes.BracketTargetOrder

SimpleOrderOption.Enabled = True
BracketOrderOption.Enabled = True
BracketOrderOption.value = True

setOrderScheme OrderSchemes.BracketOrder

selectComboEntry ActionCombo(BracketIndexes.BracketEntryOrder), _
                OrderActions.OrderActionBuy
setAction BracketIndexes.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndexes.BracketEntryOrder), _
                BracketEntryTypes.BracketEntryTypeLimit
setOrderFieldsEnabling BracketIndexes.BracketEntryOrder, Nothing
configureOrderFields BracketIndexes.BracketEntryOrder

selectComboEntry TypeCombo(BracketIndexes.BracketStopOrder), _
                BracketStopLossTypes.BracketStopLossTypeStop
setOrderFieldsEnabling BracketIndexes.BracketStopOrder, Nothing
configureOrderFields BracketIndexes.BracketStopOrder

selectComboEntry TypeCombo(BracketIndexes.BracketTargetOrder), _
                BracketTargetTypes.BracketTargetTypeNone
setOrderFieldsEnabling BracketIndexes.BracketTargetOrder, Nothing
configureOrderFields BracketIndexes.BracketTargetOrder

Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(BracketTabs.TabEntryOrder)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub selectComboEntry( _
                ByVal combo As TWImageCombo, _
                ByVal ItemData As Long)
Const ProcName As String = "selectComboEntry"
On Error GoTo Err

Dim i As Long
For i = 1 To combo.ComboItems.Count
    If combo.ComboItems(i).Tag = ItemData Then
        Set combo.SelectedItem = combo.ComboItems(i)
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setAction( _
                ByVal index As Long)
Const ProcName As String = "setAction"
On Error GoTo Err

If BracketOrderOption.value And index = BracketIndexes.BracketEntryOrder Then
    If comboItemData(ActionCombo(index)) = OrderActions.OrderActionSell Then
        selectComboEntry ActionCombo(BracketIndexes.BracketStopOrder), OrderActions.OrderActionBuy
        selectComboEntry ActionCombo(BracketIndexes.BracketTargetOrder), OrderActions.OrderActionBuy
    Else
        selectComboEntry ActionCombo(BracketIndexes.BracketStopOrder), OrderActions.OrderActionSell
        selectComboEntry ActionCombo(BracketIndexes.BracketTargetOrder), OrderActions.OrderActionSell
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setActiveOrderContext(ByVal value As OrderContext)
Const ProcName As String = "setActiveOrderContext"
On Error GoTo Err

If value Is mActiveOrderContext Then Exit Sub

If Not mDataSource Is Nothing Then
    mDataSource.RemoveGenericTickListener Me
    mDataSource.RemoveStateChangeListener Me
End If

Set mActiveOrderContext = value
Set mContract = gGetContractFromContractFuture(mActiveOrderContext.ContractFuture)

Set mDataSource = mActiveOrderContext.DataSource
If Not mDataSource Is Nothing Then
    mDataSource.AddGenericTickListener Me
    mDataSource.AddStateChangeListener Me
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInvalidControl( _
                ByVal pControl As Control, _
                ByVal index As Long)
Const ProcName As String = "setInvalidControl"
On Error GoTo Err

Set mInvalidControls(index) = pControl
If BracketTabStrip.Visible Then Set BracketTabStrip.SelectedItem = BracketTabStrip.Tabs(index + 1)
pControl.BackColor = ErroredFieldColor

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
    SimulateOrdersCheck.value = vbUnchecked
    SimulateOrdersCheck.Visible = False
Case OrderTicketModeSimulatedOnly
    SimulateOrdersCheck.value = vbChecked
    SimulateOrdersCheck.Visible = False
Case OrderTicketModeLiveAndSimulated
    SimulateOrdersCheck.value = vbUnchecked
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
' @param orderIndex the index of the order page whose fields are the source of
'                   the attribute values
'
'*/
Private Sub setOrderAttributes( _
                ByVal pOrder As IOrder, _
                ByVal orderIndex As Long)

Const ProcName As String = "setOrderAttributes"

On Error GoTo Err

With pOrder
    If pOrder.IsAttributeModifiable(OrderAttAllOrNone) Then .AllOrNone = (AllOrNoneCheck(orderIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttBlockOrder) Then .BlockOrder = (BlockOrderCheck(orderIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttDiscretionaryAmount) Then .DiscretionaryAmount = IIf(DiscrAmountText(orderIndex) = "", 0, DiscrAmountText(orderIndex))
    If pOrder.IsAttributeModifiable(OrderAttDisplaySize) Then .displaySize = IIf(DisplaySizeText(orderIndex) = "", 0, DisplaySizeText(orderIndex))
    If pOrder.IsAttributeModifiable(OrderAttETradeOnly) Then .ETradeOnly = (ETradeOnlyCheck(orderIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttFirmQuoteOnly) Then .FirmQuoteOnly = (FirmQuoteOnlyCheck(orderIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttGoodAfterTime) Then .GoodAfterTime = IIf(GoodAfterTimeText(orderIndex) = "", 0, GoodAfterTimeText(orderIndex))
    If pOrder.IsAttributeModifiable(OrderAttGoodAfterTimeTZ) Then .GoodAfterTimeTZ = GoodAfterTimeTZText(orderIndex)
    If pOrder.IsAttributeModifiable(OrderAttGoodTillDate) Then .GoodTillDate = IIf(GoodTillDateText(orderIndex) = "", 0, GoodTillDateText(orderIndex))
    If pOrder.IsAttributeModifiable(OrderAttGoodTillDateTZ) Then .GoodTillDateTZ = GoodTillDateTZText(orderIndex)
    If pOrder.IsAttributeModifiable(OrderAttHidden) Then .Hidden = (HiddenCheck(orderIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttIgnoreRTH) Then .IgnoreRegularTradingHours = (IgnoreRthCheck(orderIndex) = vbChecked)
    'If pOrder.isAttributeModifiable(OrderAttLimitPrice) Then .limitPrice = IIf(PriceText(orderIndex) = "", 0, PriceText(orderIndex))
    If pOrder.IsAttributeModifiable(OrderAttMinimumQuantity) Then .MinimumQuantity = IIf(MinQuantityText(orderIndex) = "", 0, MinQuantityText(orderIndex))
    'If pOrder.isAttributeModifiable(OrderAttOrderType) Then .orderType = comboItemData(TypeCombo(orderIndex))
    If pOrder.IsAttributeModifiable(OrderAttOriginatorRef) Then .OriginatorRef = OrderRefText(orderIndex)
    If pOrder.IsAttributeModifiable(OrderAttOverrideConstraints) Then .OverrideConstraints = (OverrideCheck(orderIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttQuantity) Then .Quantity = QuantityText(orderIndex)
    If pOrder.IsAttributeModifiable(OrderAttStopTriggerMethod) Then .StopTriggerMethod = comboItemData(TriggerMethodCombo(orderIndex))
    If pOrder.IsAttributeModifiable(OrderAttSweepToFill) Then .SweepToFill = (SweepToFillCheck(orderIndex) = vbChecked)
    If pOrder.IsAttributeModifiable(OrderAttTimeInForce) Then .TimeInForce = comboItemData(TIFCombo(orderIndex))
    'If pOrder.isAttributeModifiable(OrderAttTriggerPrice) Then .triggerPrice = IIf(StopPriceText(orderIndex) = "", 0, StopPriceText(orderIndex))
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderFieldValues( _
                ByVal pOrder As IOrder, _
                ByVal orderIndex As Long)
Const ProcName As String = "setOrderFieldValues"
On Error GoTo Err

If pOrder Is Nothing Then
    disableOrderFields orderIndex
    Exit Sub
End If

clearOrderFields orderIndex

With pOrder
    setOrderId orderIndex, .Id
    
    selectComboEntry ActionCombo(orderIndex), .Action
    QuantityText(orderIndex) = .Quantity
    selectComboEntry TypeCombo(orderIndex), .OrderType
    PriceText(orderIndex) = IIf(.LimitPrice <> 0, .LimitPrice, "")
    StopPriceText(orderIndex) = IIf(.TriggerPrice <> 0, .TriggerPrice, "")
    IgnoreRthCheck(orderIndex) = IIf(.IgnoreRegularTradingHours, vbChecked, vbUnchecked)
    selectComboEntry TIFCombo(orderIndex), .TimeInForce
    OrderRefText(orderIndex) = .OriginatorRef
    AllOrNoneCheck(orderIndex) = IIf(.AllOrNone, vbChecked, vbUnchecked)
    BlockOrderCheck(orderIndex) = IIf(.BlockOrder, vbChecked, vbUnchecked)
    ETradeOnlyCheck(orderIndex) = IIf(.ETradeOnly, vbChecked, vbUnchecked)
    FirmQuoteOnlyCheck(orderIndex) = IIf(.FirmQuoteOnly, vbChecked, vbUnchecked)
    HiddenCheck(orderIndex) = IIf(.Hidden, vbChecked, vbUnchecked)
    OverrideCheck(orderIndex) = IIf(.OverrideConstraints, vbChecked, vbUnchecked)
    SweepToFillCheck(orderIndex) = IIf(.SweepToFill, vbChecked, vbUnchecked)
    DisplaySizeText(orderIndex) = IIf(.displaySize <> 0, .displaySize, "")
    MinQuantityText(orderIndex) = IIf(.MinimumQuantity <> 0, .displaySize, "")
    If .StopTriggerMethod <> 0 Then TriggerMethodCombo(orderIndex) = OrderStopTriggerMethodToString(.StopTriggerMethod)
    DiscrAmountText(orderIndex) = IIf(.DiscretionaryAmount <> 0, .DiscretionaryAmount, "")
    GoodAfterTimeText(orderIndex) = IIf(.GoodAfterTime <> 0, FormatDateTime(.GoodAfterTime, vbGeneralDate), "")
    GoodAfterTimeTZText(orderIndex) = .GoodAfterTimeTZ
    GoodTillDateText(orderIndex) = IIf(.GoodTillDate <> 0, FormatDateTime(.GoodTillDate, vbGeneralDate), "")
    GoodTillDateTZText(orderIndex) = .GoodTillDateTZ
    
    ' do this last because it sets the various fields attributes
    selectComboEntry TypeCombo(orderIndex), .OrderType
End With

If Not isOrderModifiable(pOrder) Then
    disableOrderFields orderIndex
Else
    setOrderFieldsEnabling orderIndex, pOrder
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
                ByVal index As Long, _
                ByVal pOrder As IOrder)
Const ProcName As String = "setOrderFieldsEnabling"
On Error GoTo Err

If index = BracketIndexes.BracketEntryOrder Then setOrderFieldEnabling ActionCombo(index), OrderAttAction, pOrder
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderId( _
                ByVal index As Long, _
                ByVal Id As String)
Const ProcName As String = "setOrderId"
On Error GoTo Err

'enableControl OrderIdLabel(index)
OrderIdLabel(index).caption = Id
'disableControl OrderIdLabel(index)

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

Private Sub setPriceField( _
                index As Integer)
Const ProcName As String = "setPriceField"
On Error GoTo Err

Dim lBasePrice As Double

Select Case index
Case BracketIndexes.BracketEntryOrder
    Select Case comboItemData(TypeCombo(index))
    Case BracketEntryTypeBid
        lBasePrice = mDataSource.CurrentTick(TickTypeBid).Price
    Case BracketEntryTypeAsk
        lBasePrice = mDataSource.CurrentTick(TickTypeAsk).Price
    Case BracketEntryTypeLast
        lBasePrice = mDataSource.CurrentTick(TickTypeTrade).Price
    Case Else
        Exit Sub
    End Select
Case BracketIndexes.BracketStopOrder
    Select Case comboItemData(TypeCombo(index))
    Case BracketStopLossTypeBid
        lBasePrice = mDataSource.CurrentTick(TickTypeBid).Price
    Case BracketStopLossTypeAsk
        lBasePrice = mDataSource.CurrentTick(TickTypeAsk).Price
    Case BracketStopLossTypeLast
        lBasePrice = mDataSource.CurrentTick(TickTypeTrade).Price
    Case BracketStopLossTypeAuto
        lBasePrice = 0
    Case Else
        Exit Sub
    End Select
Case BracketIndexes.BracketTargetOrder
    Select Case comboItemData(TypeCombo(index))
    Case BracketTargetTypeBid
        lBasePrice = mDataSource.CurrentTick(TickTypeBid).Price
    Case BracketTargetTypeAsk
        lBasePrice = mDataSource.CurrentTick(TickTypeAsk).Price
    Case BracketTargetTypeLast
        lBasePrice = mDataSource.CurrentTick(TickTypeTrade).Price
    Case BracketTargetTypeAuto
        lBasePrice = 0
    Case Else
        Exit Sub
    End Select
End Select

Dim lOffset As Double
If IsNumeric(OffsetText(index)) Then lOffset = OffsetText(index) * mContract.TickSize

PriceText(index) = priceToString(lBasePrice + lOffset)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPriceFields()
Const ProcName As String = "setPriceFields"
On Error GoTo Err

setPriceField BracketIndexes.BracketEntryOrder
setPriceField BracketIndexes.BracketStopOrder
setPriceField BracketIndexes.BracketTargetOrder

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

Private Sub setupActionCombo(ByVal index As Long)
Const ProcName As String = "setupActionCombo"
On Error GoTo Err

ActionCombo(index).ComboItems.Clear
If index <> BracketIndexes.BracketEntryOrder Then
    addItemToCombo ActionCombo(index), _
                OrderActionToString(OrderActions.OrderActionNone), _
                OrderActions.OrderActionNone
    disableControl ActionCombo(index)
End If
addItemToCombo ActionCombo(index), _
            OrderActionToString(OrderActions.OrderActionBuy), _
            OrderActions.OrderActionBuy
addItemToCombo ActionCombo(index), _
            OrderActionToString(OrderActions.OrderActionSell), _
            OrderActions.OrderActionSell

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupControls()
Const ProcName As String = "setupControls"
On Error GoTo Err

SymbolLabel.caption = mContract.Specifier.LocalSymbol & _
                        " on " & _
                        mContract.Specifier.Exchange
                        
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

setQuantity BracketIndexes.BracketEntryOrder
setQuantity BracketIndexes.BracketStopOrder
setQuantity BracketIndexes.BracketTargetOrder

showDataSourceValues

If mActiveOrderContext.IsSimulated Then
    OrderSimulationLabel.caption = OrdersSimulatedMessage
Else
    OrderSimulationLabel.caption = OrdersLiveMessage
End If

ActionCombo(0).SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTifCombo(ByVal index As Long)
Const ProcName As String = "setupTifCombo"
On Error GoTo Err

TIFCombo(index).ComboItems.Clear

If mActiveOrderContext.IsOrderTifSupported(OrderTIFs.OrderTIFDay) Then
    addItemToCombo TIFCombo(index), _
                OrderTIFToString(OrderTIFs.OrderTIFDay), _
                OrderTIFs.OrderTIFDay
End If
If mActiveOrderContext.IsOrderTifSupported(OrderTIFs.OrderTIFGoodTillCancelled) Then
    addItemToCombo TIFCombo(index), _
                OrderTIFToString(OrderTIFs.OrderTIFGoodTillCancelled), _
                OrderTIFs.OrderTIFGoodTillCancelled
End If
If mActiveOrderContext.IsOrderTifSupported(OrderTIFs.OrderTIFImmediateOrCancel) Then
    addItemToCombo TIFCombo(index), _
                OrderTIFToString(OrderTIFs.OrderTIFImmediateOrCancel), _
                OrderTIFs.OrderTIFImmediateOrCancel
End If

setComboListIndex TIFCombo(index), 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTriggerMethodCombo(ByVal index As Long)
Const ProcName As String = "setupTriggerMethodCombo"
On Error GoTo Err

TriggerMethodCombo(index).ComboItems.Clear

If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerDefault) Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerDefault), _
                OrderStopTriggerMethods.OrderStopTriggerDefault
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerLast) Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerLast), _
                OrderStopTriggerMethods.OrderStopTriggerLast
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerBidAsk) Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerBidAsk), _
                OrderStopTriggerMethods.OrderStopTriggerBidAsk
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk) Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk), _
                OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerDoubleLast) Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerDoubleLast), _
                OrderStopTriggerMethods.OrderStopTriggerDoubleLast
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk) Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk), _
                OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk
End If
If mActiveOrderContext.IsStopTriggerMethodSupported(OrderStopTriggerMethods.OrderStopTriggerMidPoint) Then
    addItemToCombo TriggerMethodCombo(index), _
                OrderStopTriggerMethodToString(OrderStopTriggerMethods.OrderStopTriggerMidPoint), _
                OrderStopTriggerMethods.OrderStopTriggerMidPoint
End If

setComboListIndex TriggerMethodCombo(index), 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTypeCombo(ByVal index As Long)
Const ProcName As String = "setupTypeCombo"
On Error GoTo Err

TypeCombo(index).ComboItems.Clear

If index = BracketIndexes.BracketEntryOrder Then
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimit) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeLimit), _
                    BracketEntryTypes.BracketEntryTypeLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarket) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeMarket), _
                    BracketEntryTypes.BracketEntryTypeMarket
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStop) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeStop), _
                    BracketEntryTypes.BracketEntryTypeStop
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStopLimit) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeStopLimit), _
                    BracketEntryTypes.BracketEntryTypeStopLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimit) And _
        Not mDataSource Is Nothing _
    Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeBid), _
                    BracketEntryTypes.BracketEntryTypeBid
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeAsk), _
                    BracketEntryTypes.BracketEntryTypeAsk
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeLast), _
                    BracketEntryTypes.BracketEntryTypeLast
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimitOnOpen) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeLimitOnOpen), _
                    BracketEntryTypes.BracketEntryTypeLimitOnOpen
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketOnOpen) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeMarketOnOpen), _
                    BracketEntryTypes.BracketEntryTypeMarketOnOpen
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimitOnClose) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeLimitOnClose), _
                    BracketEntryTypes.BracketEntryTypeLimitOnClose
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketOnClose) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeMarketOnClose), _
                    BracketEntryTypes.BracketEntryTypeMarketOnClose
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimitIfTouched) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeLimitIfTouched), _
                    BracketEntryTypes.BracketEntryTypeLimitIfTouched
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketIfTouched) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeMarketIfTouched), _
                    BracketEntryTypes.BracketEntryTypeMarketIfTouched
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketToLimit) Then
        addItemToCombo TypeCombo(index), _
                    BracketEntryTypeToString(BracketEntryTypes.BracketEntryTypeMarketToLimit), _
                    BracketEntryTypes.BracketEntryTypeMarketToLimit
    End If
ElseIf index = BracketIndexes.BracketStopOrder Then
    addItemToCombo TypeCombo(index), _
                BracketStopLossTypeToString(BracketStopLossTypes.BracketStopLossTypeNone), _
                BracketStopLossTypes.BracketStopLossTypeNone
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStop) Then
        addItemToCombo TypeCombo(index), _
                    BracketStopLossTypeToString(BracketStopLossTypes.BracketStopLossTypeStop), _
                    BracketStopLossTypes.BracketStopLossTypeStop
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStopLimit) Then
        addItemToCombo TypeCombo(index), _
                    BracketStopLossTypeToString(BracketStopLossTypes.BracketStopLossTypeStopLimit), _
                    BracketStopLossTypes.BracketStopLossTypeStopLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimit) And _
        Not mDataSource Is Nothing _
    Then
        addItemToCombo TypeCombo(index), _
                    BracketStopLossTypeToString(BracketStopLossTypes.BracketStopLossTypeBid), _
                    BracketStopLossTypes.BracketStopLossTypeBid
        addItemToCombo TypeCombo(index), _
                    BracketStopLossTypeToString(BracketStopLossTypes.BracketStopLossTypeAsk), _
                    BracketStopLossTypes.BracketStopLossTypeAsk
        addItemToCombo TypeCombo(index), _
                    BracketStopLossTypeToString(BracketStopLossTypes.BracketStopLossTypeLast), _
                    BracketStopLossTypes.BracketStopLossTypeLast
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeStop) Then
        addItemToCombo TypeCombo(index), _
                    BracketStopLossTypeToString(BracketStopLossTypes.BracketStopLossTypeAuto), _
                    BracketStopLossTypes.BracketStopLossTypeAuto
    End If
ElseIf index = BracketIndexes.BracketTargetOrder Then
    addItemToCombo TypeCombo(index), _
                BracketTargetTypeToString(BracketTargetTypes.BracketTargetTypeNone), _
                BracketTargetTypes.BracketTargetTypeNone
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimit) Then
        addItemToCombo TypeCombo(index), _
                    BracketTargetTypeToString(BracketTargetTypes.BracketTargetTypeLimit), _
                    BracketTargetTypes.BracketTargetTypeLimit
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeMarketIfTouched) Then
        addItemToCombo TypeCombo(index), _
                    BracketTargetTypeToString(BracketTargetTypes.BracketTargetTypeMarketIfTouched), _
                    BracketTargetTypes.BracketTargetTypeMarketIfTouched
    End If
    If mActiveOrderContext.IsOrderTypeSupported(OrderTypes.OrderTypeLimit) And _
        Not mDataSource Is Nothing _
    Then
        addItemToCombo TypeCombo(index), _
                    BracketTargetTypeToString(BracketTargetTypes.BracketTargetTypeBid), _
                    BracketTargetTypes.BracketTargetTypeBid
        addItemToCombo TypeCombo(index), _
                    BracketTargetTypeToString(BracketTargetTypes.BracketTargetTypeAsk), _
                    BracketTargetTypes.BracketTargetTypeAsk
        addItemToCombo TypeCombo(index), _
                    BracketTargetTypeToString(BracketTargetTypes.BracketTargetTypeLast), _
                    BracketTargetTypes.BracketTargetTypeLast
        addItemToCombo TypeCombo(index), _
                    BracketTargetTypeToString(BracketTargetTypes.BracketTargetTypeAuto), _
                    BracketTargetTypes.BracketTargetTypeAuto
    End If
End If

setComboListIndex TypeCombo(index), 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showOrderFields(ByVal index As Long)
Const ProcName As String = "showOrderFields"
On Error GoTo Err

Dim i As Long
For i = 0 To ActionCombo.Count - 1
    If i = index Then
        OrderIdLabel(i).Visible = True
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
        OrderIdLabel(i).Visible = False
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showDataSourceValues()
Const ProcName As String = "showDataSourceValues"
On Error GoTo Err

If mDataSource Is Nothing Then Exit Sub

AskText.Text = priceToString(mDataSource.CurrentTick(TickTypeAsk).Price)
AskSizeText.Text = mDataSource.CurrentTick(TickTypeAsk).Size
BidText.Text = priceToString(mDataSource.CurrentTick(TickTypeBid).Price)
BidSizeText.Text = mDataSource.CurrentTick(TickTypeBid).Size
LastText.Text = priceToString(mDataSource.CurrentTick(TickTypeTrade).Price)
LastSizeText.Text = mDataSource.CurrentTick(TickTypeTrade).Size
VolumeText.Text = mDataSource.CurrentTick(TickTypeVolume).Size
HighText.Text = priceToString(mDataSource.CurrentTick(TickTypeHighPrice).Price)
LowText.Text = priceToString(mDataSource.CurrentTick(TickTypeLowPrice).Price)
setPriceFields

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




