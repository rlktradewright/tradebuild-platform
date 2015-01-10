VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#279.1#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#27.1#0"; "TWControls40.ocx"
Begin VB.UserControl InfoPanel 
   Appearance      =   0  'Flat
   BackColor       =   &H00CDF3FF&
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12285
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   12285
   Begin VB.PictureBox HidePicture 
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   11940
      MouseIcon       =   "InfoPanel.ctx":0000
      MousePointer    =   99  'Custom
      Picture         =   "InfoPanel.ctx":0152
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      ToolTipText     =   "Hide Information Panel"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox UnpinPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   11580
      MouseIcon       =   "InfoPanel.ctx":06DC
      MousePointer    =   99  'Custom
      Picture         =   "InfoPanel.ctx":082E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      ToolTipText     =   "Unpin Information Panel"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox PinPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   11640
      MouseIcon       =   "InfoPanel.ctx":0DB8
      MousePointer    =   99  'Custom
      Picture         =   "InfoPanel.ctx":0F0A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      ToolTipText     =   "Pin Information Panel"
      Top             =   30
      Width           =   240
   End
   Begin TabDlg.SSTab InfoSSTab 
      Height          =   4455
      Left            =   -2
      TabIndex        =   0
      Top             =   300
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1. Orders"
      TabPicture(0)   =   "InfoPanel.ctx":1494
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "OrdersPicture"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2. Executions"
      TabPicture(1)   =   "InfoPanel.ctx":14B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ExecutionsPicture"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3. Log"
      TabPicture(2)   =   "InfoPanel.ctx":14CC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LogText"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox LogText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   -75000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Status messages"
         Top             =   295
         Width           =   12195
      End
      Begin VB.PictureBox OrdersPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4160
         Left            =   0
         ScaleHeight     =   4155
         ScaleWidth      =   12255
         TabIndex        =   12
         Top             =   295
         Width           =   12255
         Begin TWControls40.TWButton ClosePositionsButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            DefaultBorderColor=   15793920
            ForeColor       =   13684944
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Close all positions!"
            ForeColor       =   13684944
         End
         Begin TWControls40.TWButton OrderTicketButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   14
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            DefaultBorderColor=   15793920
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Order Ticket"
         End
         Begin TWControls40.TWButton CancelOrderPlexButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   1
            Top             =   2370
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            DefaultBorderColor=   15793920
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Cancel"
         End
         Begin TWControls40.TWButton ModifyOrderPlexButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   2
            Top             =   1770
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            DefaultBorderColor=   15793920
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Modify"
         End
         Begin TradingUI27.OrdersSummary LiveOrdersSummary 
            Height          =   3735
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6588
         End
         Begin TradingUI27.OrdersSummary SimulatedOrdersSummary 
            Height          =   3615
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   6376
         End
         Begin TradingUI27.OrdersSummary TickfileOrdersSummary 
            Height          =   3615
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   6376
         End
         Begin MSComctlLib.TabStrip OrdersSummaryTabStrip 
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   3690
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            MultiRow        =   -1  'True
            Style           =   2
            Placement       =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Live"
                  Object.ToolTipText     =   "Show live orders"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Simulated"
                  Object.ToolTipText     =   "Show simulated orders"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Tickfile"
                  Object.ToolTipText     =   "Show tickfile orders"
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox ExecutionsPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   -75000
         ScaleHeight     =   4155
         ScaleWidth      =   12255
         TabIndex        =   16
         Top             =   300
         Width           =   12255
         Begin TradingUI27.ExecutionsSummary LiveExecutionsSummary 
            Height          =   3750
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   6615
         End
         Begin TradingUI27.ExecutionsSummary SimulatedExecutionsSummary 
            Height          =   3750
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   6615
         End
         Begin TradingUI27.ExecutionsSummary TickfileExecutionsSummary 
            Height          =   3750
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   6615
         End
         Begin MSComctlLib.TabStrip ExecutionsSummaryTabStrip 
            Height          =   375
            Left            =   0
            TabIndex        =   18
            Top             =   3750
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            MultiRow        =   -1  'True
            Style           =   2
            Placement       =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Live"
                  Object.ToolTipText     =   "Show live executions"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Simulated"
                  Object.ToolTipText     =   "Show simulated executions"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Tickfile"
                  Object.ToolTipText     =   "Show executions against tickfiles"
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "InfoPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IThemeable
Implements LogListener
Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Hide()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Pin()
Event Unpin()

'@================================================================================
' Enums
'@================================================================================

Private Enum InfoTabIndexNumbers
    InfoTabIndexOrders
    InfoTabIndexExecutions
    InfoTabIndexLog
End Enum

Private Enum OrdersTabIndexNumbers
    OrdersTabIndexLive = 1
    OrdersTabIndexSimulated
    OrderTabIndexTickfile
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "InfoPanel"

Private Const ExecutionsTabCaptionLive              As String = "Live"
Private Const ExecutionsTabCaptionSimulated         As String = "Simulated"
Private Const ExecutionsTabCaptionTickfile          As String = "Tickfile"

Private Const MinimumHeightTwips                    As Long = 2985
Private Const MinimumWidthTwips                     As Long = 4215

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildAPI                              As TradeBuildAPI
Private mAppInstanceConfig                          As ConfigurationSection

Private WithEvents mTickerGrid                      As TickerGrid
Attribute mTickerGrid.VB_VarHelpID = -1
Private WithEvents mTickers                         As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mOrderTicket                                As fOrderTicket

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent Mouseup(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

If UserControl.Height < MinimumHeightTwips Then UserControl.Height = MinimumHeightTwips
If UserControl.Width < MinimumWidthTwips Then UserControl.Width = MinimumWidthTwips

UnpinPicture.Left = UserControl.Width - 47 * Screen.TwipsPerPixelX
PinPicture.Left = UserControl.Width - 47 * Screen.TwipsPerPixelX
HidePicture.Left = UserControl.Width - 23 * Screen.TwipsPerPixelX

InfoSSTab.Height = UserControl.Height - InfoSSTab.Top + 2 * Screen.TwipsPerPixelY
InfoSSTab.Width = UserControl.Width + 4 * Screen.TwipsPerPixelX

OrdersPicture.Height = InfoSSTab.Height - 2 * Screen.TwipsPerPixelY - OrdersPicture.Top
OrdersPicture.Width = InfoSSTab.Width

ExecutionsPicture.Height = OrdersPicture.Height
ExecutionsPicture.Width = OrdersPicture.Width

OrderTicketButton.Left = InfoSSTab.Width - OrderTicketButton.Width - 120 - 2 * Screen.TwipsPerPixelX
ModifyOrderPlexButton.Left = OrderTicketButton.Left
CancelOrderPlexButton.Left = OrderTicketButton.Left
ClosePositionsButton.Left = OrderTicketButton.Left

OrdersSummaryTabStrip.Top = OrdersPicture.Height - OrdersSummaryTabStrip.Height
ExecutionsSummaryTabStrip.Top = OrdersSummaryTabStrip.Top

LiveOrdersSummary.Width = ModifyOrderPlexButton.Left - 120 - 120
LiveOrdersSummary.Height = OrdersSummaryTabStrip.Top - LiveOrdersSummary.Top

SimulatedOrdersSummary.Width = LiveOrdersSummary.Width
SimulatedOrdersSummary.Height = LiveOrdersSummary.Height

TickfileOrdersSummary.Width = LiveOrdersSummary.Width
TickfileOrdersSummary.Height = LiveOrdersSummary.Height

LogText.Width = InfoSSTab.Width - 4 * Screen.TwipsPerPixelX
LogText.Height = InfoSSTab.Height - 2 * Screen.TwipsPerPixelY - LogText.Top

LiveExecutionsSummary.Width = InfoSSTab.Width - 4 * Screen.TwipsPerPixelX
LiveExecutionsSummary.Height = ExecutionsPicture.Height - ExecutionsSummaryTabStrip.Top

SimulatedExecutionsSummary.Width = LiveExecutionsSummary.Width
SimulatedExecutionsSummary.Height = LiveExecutionsSummary.Height

TickfileExecutionsSummary.Width = LiveExecutionsSummary.Width
TickfileExecutionsSummary.Height = LiveExecutionsSummary.Height

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

'================================================================================
' LogListener Interface Members
'================================================================================

Private Sub LogListener_Finish()
'nothing to do
End Sub

Private Sub LogListener_Notify(ByVal Logrec As LogRecord)
Const ProcName As String = "LogListener_Notify"
On Error GoTo Err

If Len(LogText.Text) >= 32767 Then
    ' clear some space at the start of the textbox
    LogText.SelStart = 0
    LogText.SelLength = 16384
    LogText.SelText = ""
End If

LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = formatLogRecord(Logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Dim lTicker As Ticker
Set lTicker = ev.Source

Select Case ev.State
Case MarketDataSourceStates.MarketDataSourceStateCreated

Case MarketDataSourceStates.MarketDataSourceStateReady
Case MarketDataSourceStates.MarketDataSourceStateRunning
    If lTicker Is getSelectedDataSource Then
        If lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled Then OrderTicketButton.Enabled = True
    End If
    
Case MarketDataSourceStates.MarketDataSourceStatePaused

Case MarketDataSourceStates.MarketDataSourceStateStopped
    If getSelectedDataSource Is Nothing Then OrderTicketButton.Enabled = False
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub CancelOrderPlexButton_Click()
Const ProcName As String = "CancelOrderPlexButton_Click"
On Error GoTo Err

Dim op As IBracketOrder

If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    Set op = LiveOrdersSummary.SelectedItem
ElseIf OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexSimulated Then
    Set op = SimulatedOrdersSummary.SelectedItem
Else
    Set op = TickfileOrdersSummary.SelectedItem
End If
If Not op Is Nothing Then op.Cancel True

CancelOrderPlexButton.Enabled = False
ModifyOrderPlexButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ClosePositionsButton_Click()
Const ProcName As String = "ClosePositionsButton_Click"
On Error GoTo Err

If Not mTradeBuildAPI.ClosingPositions Then
    If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
        mTradeBuildAPI.CloseAllPositions PositionTypeLive, _
                                        ClosePositionCancelOrders Or ClosePositionWaitForCancel
    Else
        mTradeBuildAPI.CloseAllPositions PositionTypeSimulated, _
                                        ClosePositionCancelOrders Or ClosePositionWaitForCancel
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ExecutionsSummaryTabStrip_Click()
Const ProcName As String = "ExecutionsSummaryTabStrip_Click"
On Error GoTo Err

Static currIndex As Long
If ExecutionsSummaryTabStrip.SelectedItem.Index = currIndex Then Exit Sub

Select Case ExecutionsSummaryTabStrip.SelectedItem.caption
Case ExecutionsTabCaptionLive
    LiveExecutionsSummary.Visible = True
    SimulatedExecutionsSummary.Visible = False
    TickfileExecutionsSummary.Visible = False
Case ExecutionsTabCaptionSimulated
    LiveExecutionsSummary.Visible = False
    SimulatedExecutionsSummary.Visible = True
    TickfileExecutionsSummary.Visible = False
Case ExecutionsTabCaptionTickfile
    LiveExecutionsSummary.Visible = False
    SimulatedExecutionsSummary.Visible = False
    TickfileExecutionsSummary.Visible = True
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub HidePicture_Click()
RaiseEvent Hide
End Sub

Private Sub InfoSSTab_Click(PreviousTab As Integer)
Const ProcName As String = "InfoSSTAB_Click"
On Error GoTo Err

Select Case InfoSSTab.Tab
Case InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexLog
Case InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexOrders
    If ModifyOrderPlexButton.Enabled Then
        ModifyOrderPlexButton.Default = True
    Else
        If CancelOrderPlexButton.Enabled Then CancelOrderPlexButton.Default = True
    End If
Case InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexExecutions
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveOrdersSummary_SelectionChanged()
Const ProcName As String = "LiveOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection LiveOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ModifyOrderPlexButton_Click()
Const ProcName As String = "ModifyOrderPlexButton_Click"
On Error GoTo Err

Dim os As OrdersSummary

If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    Set os = LiveOrdersSummary
ElseIf OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexSimulated Then
    Set os = SimulatedOrdersSummary
Else
    Set os = TickfileOrdersSummary
End If

If os.SelectedItem Is Nothing Then
    ModifyOrderPlexButton.Enabled = False
ElseIf os.IsSelectedItemModifiable Then
    mOrderTicket.Show vbModeless, Me
    mOrderTicket.ShowBracketOrder os.SelectedItem, os.SelectedOrderIndex
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrdersSummaryTabStrip_Click()
Const ProcName As String = "OrdersSummaryTabStrip_Click"
On Error GoTo Err

Static currIndex As Long
If OrdersSummaryTabStrip.SelectedItem.Index = currIndex Then Exit Sub

Select Case OrdersSummaryTabStrip.SelectedItem.Index
Case OrdersTabIndexNumbers.OrdersTabIndexLive
    LiveOrdersSummary.Visible = True
    SimulatedOrdersSummary.Visible = False
    TickfileOrdersSummary.Visible = False
    setOrdersSelection LiveOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrdersTabIndexLive
Case OrdersTabIndexNumbers.OrdersTabIndexSimulated
    LiveOrdersSummary.Visible = False
    SimulatedOrdersSummary.Visible = True
    TickfileOrdersSummary.Visible = False
    setOrdersSelection SimulatedOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrdersTabIndexSimulated
Case OrdersTabIndexNumbers.OrderTabIndexTickfile
    LiveOrdersSummary.Visible = False
    SimulatedOrdersSummary.Visible = False
    TickfileOrdersSummary.Visible = True
    setOrdersSelection TickfileOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrderTabIndexTickfile
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderTicketButton_Click()
Const ProcName As String = "OrderTicketButton_Click"
On Error GoTo Err

showOrderTicket

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub PinPicture_Click()
RaiseEvent Pin
End Sub

Private Sub SimulatedOrdersSummary_SelectionChanged()
Const ProcName As String = "SimulatedOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection SimulatedOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickfileOrdersSummary_SelectionChanged()
Const ProcName As String = "SimulatedOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection TickfileOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UnpinPicture_Click()
RaiseEvent Unpin
End Sub

'================================================================================
' mTickerGrid Event Handlers
'================================================================================

Private Sub mTickerGrid_TickerSelectionChanged()
Const ProcName As String = "mTickerGrid_TickerSelectionChanged"
On Error GoTo Err

handleSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

Dim lTicker As Ticker

Select Case ev.ChangeType
Case CollItemAdded
    Set lTicker = ev.AffectedItem
    lTicker.AddStateChangeListener Me
Case CollItemRemoved
    Set lTicker = ev.AffectedItem
    lTicker.RemoveStateChangeListener Me
Case CollItemChanged

Case CollOrderChanged

Case CollCollectionCleared

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get MinimumHeight() As Long
MinimumHeight = MinimumHeightTwips
End Property

Public Property Get MinimumWidth() As Long
MinimumWidth = MinimumWidthTwips
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If Value Is Nothing Then Exit Property

Set mTheme = Value
gApplyTheme mTheme, UserControl.Controls

ClosePositionsButton.BackColor = &HFF&
ClosePositionsButton.ForeColor = &HD0D0D0
ClosePositionsButton.MouseoverColor = &H80&
ClosePositionsButton.PushedColor = &H78D0&

PinPicture.BackColor = UserControl.BackColor
UnpinPicture.BackColor = UserControl.BackColor
HidePicture.BackColor = UserControl.BackColor

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

LiveOrdersSummary.Finish
SimulatedOrdersSummary.Finish
LiveExecutionsSummary.Finish
SimulatedExecutionsSummary.Finish
stopLogging

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pPinned As Boolean, _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByVal pTickerGrid As TickerGrid, _
                ByVal pOrderTicket As fOrderTicket)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pPinned Then
    UnpinPicture.Visible = True
    PinPicture.Visible = False
Else
    UnpinPicture.Visible = False
    PinPicture.Visible = True
End If

setupLogging

Set mTradeBuildAPI = pTradeBuildAPI
Set mTickers = mTradeBuildAPI.Tickers
Set mTickerGrid = pTickerGrid
Set mAppInstanceConfig = pAppInstanceConfig
Set mOrderTicket = pOrderTicket

LogMessage "Loading configuration: Setting up order summaries"
setupOrderSummaries

LogMessage "Loading configuration: Setting up execution summaries"
setupExecutionSummaries

InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexOrders

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub MonitorTickfilePositions( _
                ByVal pTickfileDataManager As TickfileDataManager, _
                ByVal pPositionManagers As PositionManagers)
Const ProcName As String = "MonitorTickfilePositions"
On Error GoTo Err

TickfileOrdersSummary.Initialise pTickfileDataManager
TickfileOrdersSummary.MonitorPositions pPositionManagers
TickfileExecutionsSummary.MonitorPositions pPositionManagers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
On Error GoTo Err

Static formatter As LogFormatter
If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSelectedDataSource() As IMarketDataSource
Const ProcName As String = "getSelectedDataSource"
On Error GoTo Err

If mTickerGrid.SelectedTickers.Count = 1 Then Set getSelectedDataSource = mTickerGrid.SelectedTickers.Item(1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleSelectedTickers()
Const ProcName As String = "handleSelectedTickers"
On Error GoTo Err

If mTickerGrid.SelectedTickers.Count = 0 Then
    OrderTicketButton.Enabled = False
Else
    OrderTicketButton.Enabled = False
    
    Dim lTicker As Ticker
    Set lTicker = getSelectedDataSource
    If lTicker.State = MarketDataSourceStateRunning Then
        Dim lContract As IContract
        Set lContract = lTicker.ContractFuture.Value
        If lContract.Specifier.SecType <> SecTypeIndex Then
            If (lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled) Then OrderTicketButton.Enabled = True
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub stopLogging()
Const ProcName As String = "stopLogging"
On Error GoTo Err

GetLogger("log").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrdersSelection( _
                ByVal pOrdersSummary As OrdersSummary)
Const ProcName As String = "setOrdersSelection"
On Error GoTo Err

If pOrdersSummary.IsEditing Then
    pOrdersSummary.Default = True
    Exit Sub
End If

pOrdersSummary.Default = False

Dim selection As IBracketOrder
Set selection = pOrdersSummary.SelectedItem

If selection Is Nothing Then
    CancelOrderPlexButton.Enabled = False
    ModifyOrderPlexButton.Enabled = False
Else
    If pOrdersSummary.SelectedOrderIndex = 0 Then
        CancelOrderPlexButton.Enabled = True
    Else
        CancelOrderPlexButton.Enabled = False
    End If
    If pOrdersSummary.IsSelectedItemModifiable Then
        ModifyOrderPlexButton.Enabled = True
        ModifyOrderPlexButton.Default = True
    Else
        ModifyOrderPlexButton.Enabled = False
        ModifyOrderPlexButton.Default = False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupExecutionSummaries()
Const ProcName As String = "setupExecutionSummaries"
On Error GoTo Err

Do While ExecutionsSummaryTabStrip.Tabs.Count > 0
    ExecutionsSummaryTabStrip.Tabs.Remove 1
Loop

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    SimulatedExecutionsSummary.Visible = True
    ExecutionsSummaryTabStrip.Tabs.Add 1, , ExecutionsTabCaptionSimulated
Else
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated
    SimulatedExecutionsSummary.Visible = False
    LiveExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveExecutionsSummary.Visible = True
    ExecutionsSummaryTabStrip.Tabs.Add 1, , ExecutionsTabCaptionLive
    ExecutionsSummaryTabStrip.Tabs.Add 2, , ExecutionsTabCaptionSimulated
End If

If Not mTradeBuildAPI.TickfileStoreInput Is Nothing Then
    TickfileExecutionsSummary.Visible = False
    ExecutionsSummaryTabStrip.Tabs.Add ExecutionsSummaryTabStrip.Tabs.Count + 1, , ExecutionsTabCaptionTickfile
End If

If ExecutionsSummaryTabStrip.Tabs.Count = 1 Then
    ExecutionsSummaryTabStrip.Visible = False
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top + ExecutionsSummaryTabStrip.Height - SimulatedExecutionsSummary.Top
Else
    ExecutionsSummaryTabStrip.Visible = True
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    LiveExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    TickfileExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupLogging()
Const ProcName As String = "setupLogging"
On Error GoTo Err

GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupOrderSummaries()
Const ProcName As String = "setupOrderSummaries"
On Error GoTo Err

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top + OrdersSummaryTabStrip.Height - SimulatedOrdersSummary.Top
    SimulatedOrdersSummary.Visible = True
    
    LiveOrdersSummary.Visible = False
    
    OrdersSummaryTabStrip.Visible = False
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexSimulated).Selected = True
Else
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top - SimulatedOrdersSummary.Top
    
    LiveOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
    LiveOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveOrdersSummary.Height = SimulatedOrdersSummary.Height
    
    OrdersSummaryTabStrip.Visible = True
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexLive).Selected = True
End If

SimulatedOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
SimulatedOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showOrderTicket()
Const ProcName As String = "showOrderTicket"
On Error GoTo Err

If getSelectedDataSource Is Nothing Then
    gModelessMsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
Else
    mOrderTicket.Show vbModeless, Me
    mOrderTicket.Ticker = getSelectedDataSource
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




