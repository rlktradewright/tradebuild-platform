VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#292.0#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
         TabIndex        =   11
         Top             =   295
         Width           =   12255
         Begin VB.PictureBox OrdersSummaryOptionsPicture 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   3375
            TabIndex        =   17
            Top             =   3720
            Width           =   3375
            Begin VB.OptionButton TickfileOrdersOption 
               Caption         =   "Tickfile"
               Height          =   255
               Left            =   2160
               TabIndex        =   20
               Top             =   75
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.OptionButton SimulatedOrdersOption 
               Caption         =   "Simulated"
               Height          =   255
               Left            =   960
               TabIndex        =   19
               Top             =   75
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton LiveOrdersOption 
               Caption         =   "Live"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   75
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin TWControls40.TWButton ClosePositionsButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Caption         =   "Close all positions!"
            DefaultBorderColor=   15793920
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   13684944
         End
         Begin TWControls40.TWButton OrderTicketButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   13
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Caption         =   "Order Ticket"
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
         End
         Begin TWControls40.TWButton CancelOrderPlexButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   1
            Top             =   2370
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Caption         =   "&Cancel"
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
         End
         Begin TWControls40.TWButton ModifyOrderPlexButton 
            Height          =   495
            Left            =   11160
            TabIndex        =   2
            Top             =   1770
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Caption         =   "&Modify"
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
         End
         Begin TradingUI27.OrdersSummary LiveOrdersSummary 
            Height          =   3735
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6588
         End
         Begin TradingUI27.OrdersSummary SimulatedOrdersSummary 
            Height          =   3615
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   6376
         End
         Begin TradingUI27.OrdersSummary TickfileOrdersSummary 
            Height          =   3615
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   6376
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
         TabIndex        =   15
         Top             =   300
         Width           =   12255
         Begin VB.PictureBox ExecutionsSummaryOptionsPicture 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   3375
            TabIndex        =   21
            Top             =   3720
            Width           =   3375
            Begin VB.OptionButton LiveExecutionsOption 
               Caption         =   "Live"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   75
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.OptionButton SimulatedExecutionsOption 
               Caption         =   "Simulated"
               Height          =   255
               Left            =   960
               TabIndex        =   23
               Top             =   75
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton TickfileExecutionsOption 
               Caption         =   "Tickfile"
               Height          =   255
               Left            =   2160
               TabIndex        =   22
               Top             =   75
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin TradingUI27.ExecutionsSummary LiveExecutionsSummary 
            Height          =   3750
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   6615
         End
         Begin TradingUI27.ExecutionsSummary SimulatedExecutionsSummary 
            Height          =   3750
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   6615
         End
         Begin TradingUI27.ExecutionsSummary TickfileExecutionsSummary 
            Height          =   3750
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   6615
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
Implements ILogListener
Implements IStateChangeListener

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

OrdersSummaryOptionsPicture.Top = OrdersPicture.Height - OrdersSummaryOptionsPicture.Height
ExecutionsSummaryOptionsPicture.Top = OrdersSummaryOptionsPicture.Top

LiveOrdersSummary.Width = ModifyOrderPlexButton.Left - 120 - 120
LiveOrdersSummary.Height = OrdersSummaryOptionsPicture.Top - LiveOrdersSummary.Top

SimulatedOrdersSummary.Width = LiveOrdersSummary.Width
SimulatedOrdersSummary.Height = LiveOrdersSummary.Height

TickfileOrdersSummary.Width = LiveOrdersSummary.Width
TickfileOrdersSummary.Height = LiveOrdersSummary.Height

LogText.Width = InfoSSTab.Width - 4 * Screen.TwipsPerPixelX
LogText.Height = InfoSSTab.Height - 2 * Screen.TwipsPerPixelY - LogText.Top

LiveExecutionsSummary.Width = InfoSSTab.Width - 4 * Screen.TwipsPerPixelX
LiveExecutionsSummary.Height = ExecutionsPicture.Height - ExecutionsSummaryOptionsPicture.Top

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
' ILogListener Interface Members
'================================================================================

Private Sub ILogListener_Finish()
'nothing to do
End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
Const ProcName As String = "ILogListener_Notify"
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

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
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

If LiveOrdersOption.Value Then
    Set op = LiveOrdersSummary.SelectedItem
ElseIf SimulatedOrdersOption Then
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
    If LiveOrdersOption.Value Then
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

Private Sub LiveExecutionsOption_Click()
Const ProcName As String = "LiveExecutionsOption_Click"
On Error GoTo Err

LiveExecutionsSummary.Visible = True
SimulatedExecutionsSummary.Visible = False
TickfileExecutionsSummary.Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub LiveOrdersOption_Click()
Const ProcName As String = "LiveOrdersOption_Click"
On Error GoTo Err

LiveOrdersSummary.Visible = True
SimulatedOrdersSummary.Visible = False
TickfileOrdersSummary.Visible = False
setOrdersSelection LiveOrdersSummary

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

If LiveOrdersOption.Value Then
    Set os = LiveOrdersSummary
ElseIf SimulatedOrdersOption Then
    Set os = SimulatedOrdersSummary
Else
    Set os = TickfileOrdersSummary
End If

If os.SelectedItem Is Nothing Then
    ModifyOrderPlexButton.Enabled = False
ElseIf os.IsSelectedItemModifiable Then
    mOrderTicket.Show vbModeless, Me
    mOrderTicket.ShowBracketOrder os.SelectedItem, os.SelectedOrderRole
End If

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

Private Sub SimulatedExecutionsOption_Click()
Const ProcName As String = "SimulatedExecutionsOption_Click"
On Error GoTo Err

LiveExecutionsSummary.Visible = False
SimulatedExecutionsSummary.Visible = True
TickfileExecutionsSummary.Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub SimulatedOrdersOption_Click()
Const ProcName As String = "SimulatedOrdersOption_Click"
On Error GoTo Err

LiveOrdersSummary.Visible = False
SimulatedOrdersSummary.Visible = True
TickfileOrdersSummary.Visible = False
setOrdersSelection SimulatedOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub SimulatedOrdersSummary_SelectionChanged()
Const ProcName As String = "SimulatedOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection SimulatedOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickfileExecutionsOption_Click()
Const ProcName As String = "TickfileExecutionsOption_Click"
On Error GoTo Err

LiveExecutionsSummary.Visible = False
SimulatedExecutionsSummary.Visible = False
TickfileExecutionsSummary.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub TickfileOrdersOption_Click()
Const ProcName As String = "TickfileOrdersOption_Click"
On Error GoTo Err

LiveOrdersSummary.Visible = False
SimulatedOrdersSummary.Visible = False
TickfileOrdersSummary.Visible = True
setOrdersSelection TickfileOrdersSummary

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

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If mTheme Is Value Then Exit Property
Set mTheme = Value
If mTheme Is Nothing Then Exit Property

gApplyTheme mTheme, UserControl.Controls

ClosePositionsButton.BackColor = &HFF&
ClosePositionsButton.ForeColor = &HD0D0D0
ClosePositionsButton.MouseOverBackColor = &H80&
ClosePositionsButton.PushedBackColor = &H78D0&

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

Static formatter As ILogFormatter
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

OrderTicketButton.Enabled = False

If mTickerGrid.SelectedTickers.Count = 0 Then Exit Sub
    
OrderTicketButton.Enabled = False

Dim lTicker As Ticker
Set lTicker = getSelectedDataSource

If lTicker Is Nothing Then Exit Sub
If lTicker.State <> MarketDataSourceStateRunning Then Exit Sub
    
Dim lContract As IContract
Set lContract = lTicker.ContractFuture.Value
If lContract.Specifier.SecType = SecTypeIndex Then Exit Sub

If (lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled) Then OrderTicketButton.Enabled = True

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
    If pOrdersSummary.SelectedOrderRole = BracketOrderRoleNone Then
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

LiveExecutionsSummary.Visible = False
SimulatedExecutionsSummary.Visible = False
TickfileExecutionsSummary.Visible = False

Dim lLeft As Long
lLeft = 8 * Screen.TwipsPerPixelX

Dim lNumber As Long

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    SimulatedExecutionsSummary.Visible = True
    
    SimulatedExecutionsOption.Left = lLeft
    SimulatedExecutionsOption.Value = True
    SimulatedExecutionsOption.Visible = True
    lLeft = lLeft + SimulatedExecutionsOption.Width + 8 * Screen.TwipsPerPixelX
    
    lNumber = 1
Else
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated
    
    LiveExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveExecutionsSummary.Visible = True
    
    LiveExecutionsOption.Left = lLeft
    LiveExecutionsOption.Value = True
    LiveExecutionsOption.Visible = True
    lLeft = lLeft + LiveExecutionsOption.Width + 8 * Screen.TwipsPerPixelX
    
    SimulatedExecutionsOption.Left = lLeft
    SimulatedExecutionsOption = False
    SimulatedExecutionsOption.Visible = True
    lLeft = lLeft + SimulatedExecutionsOption.Width + 8 * Screen.TwipsPerPixelX
    
    lNumber = 2
End If

If Not mTradeBuildAPI.TickfileStoreInput Is Nothing Then
    TickfileExecutionsSummary.Visible = True
    
    TickfileExecutionsOption.Left = lLeft
    TickfileExecutionsOption.Value = False
    TickfileExecutionsOption.Visible = True
    
    lNumber = lNumber + 1
End If

If lNumber = 1 Then
    ExecutionsSummaryOptionsPicture.Visible = False
    SimulatedExecutionsSummary.Height = ExecutionsSummaryOptionsPicture.Top + ExecutionsSummaryOptionsPicture.Height - SimulatedExecutionsSummary.Top
Else
    ExecutionsSummaryOptionsPicture.Visible = True
    SimulatedExecutionsSummary.Height = ExecutionsSummaryOptionsPicture.Top - SimulatedExecutionsSummary.Top
    LiveExecutionsSummary.Height = SimulatedExecutionsSummary.Height
    TickfileExecutionsSummary.Height = SimulatedExecutionsSummary.Height
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

SimulatedOrdersSummary.Visible = False
LiveOrdersSummary.Visible = False
TickfileOrdersSummary.Visible = False

SimulatedOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
SimulatedOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated

Dim lLeft As Long
lLeft = 8 * Screen.TwipsPerPixelX

Dim lNumber As Long

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedOrdersSummary.Visible = True
    
    SimulatedOrdersOption.Left = lLeft
    SimulatedOrdersOption.Value = True
    SimulatedOrdersOption.Visible = True
    lLeft = lLeft + SimulatedOrdersOption.Width + 8 * Screen.TwipsPerPixelX
    
    lNumber = 1
Else
    LiveOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
    LiveOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveOrdersSummary.Visible = True
    
    LiveOrdersOption.Left = lLeft
    LiveOrdersOption.Value = True
    LiveOrdersOption.Visible = True
    lLeft = lLeft + LiveOrdersOption.Width + 8 * Screen.TwipsPerPixelX
    
    SimulatedOrdersOption.Left = lLeft
    SimulatedOrdersOption.Value = False
    SimulatedOrdersOption.Visible = True
    lLeft = lLeft + SimulatedOrdersOption.Width + 8 * Screen.TwipsPerPixelX
    
    lNumber = 2
End If

If Not mTradeBuildAPI.TickfileStoreInput Is Nothing Then
    TickfileOrdersSummary.Visible = True
    
    TickfileOrdersOption.Left = lLeft
    TickfileOrdersOption.Value = False
    TickfileOrdersOption.Visible = True
    
    lNumber = lNumber + 1
End If

If lNumber = 1 Then
    OrdersSummaryOptionsPicture.Visible = False
    SimulatedOrdersSummary.Height = OrdersSummaryOptionsPicture.Top + OrdersSummaryOptionsPicture.Height - SimulatedOrdersSummary.Top
Else
    OrdersSummaryOptionsPicture.Visible = True
    SimulatedOrdersSummary.Height = OrdersSummaryOptionsPicture.Top - SimulatedOrdersSummary.Top
    LiveOrdersSummary.Height = SimulatedOrdersSummary.Height
    TickfileOrdersSummary.Height = SimulatedOrdersSummary.Height
End If
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showOrderTicket()
Const ProcName As String = "showOrderTicket"
On Error GoTo Err

If getSelectedDataSource Is Nothing Then
    gModelessMsgBox "No ticker selected - please select a ticker", vbExclamation, mTheme, "Error"
Else
    mOrderTicket.Show vbModeless, Me
    mOrderTicket.Ticker = getSelectedDataSource
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




