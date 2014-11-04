VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#263.0#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#18.0#0"; "TWControls40.ocx"
Begin VB.UserControl FeaturesPanel 
   BackColor       =   &H00CDF3FF&
   ClientHeight    =   9675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   DefaultCancel   =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   4065
   Begin TabDlg.SSTab FeaturesSSTab 
      Height          =   9015
      Left            =   -30
      TabIndex        =   1
      Top             =   660
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FeaturesPanel.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TickersPicture"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FeaturesPanel.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LiveChartPicture"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FeaturesPanel.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "HistChartPicture"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "FeaturesPanel.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ReplayTickerPicture"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "FeaturesPanel.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ConfigPicture"
      Tab(4).ControlCount=   1
      Begin VB.PictureBox ConfigPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   51
         Top             =   0
         Width           =   4125
         Begin VB.TextBox CurrentConfigNameText 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   420
            Width           =   3375
         End
         Begin VB.CommandButton ConfigEditorButton 
            Caption         =   "Show config editor"
            Height          =   375
            Left            =   1920
            TabIndex        =   52
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Current configuration is:"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.PictureBox ReplayTickerPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   41
         Top             =   0
         Width           =   4125
         Begin VB.ComboBox ReplaySpeedCombo 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   2760
            Width           =   2775
         End
         Begin VB.CommandButton StopReplayButton 
            Caption         =   "St&op"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3360
            TabIndex        =   45
            ToolTipText     =   "Stop tickfile replay"
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton PauseReplayButton 
            Caption         =   "P&ause"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2640
            TabIndex        =   44
            ToolTipText     =   "Pause tickfile replay"
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton PlayTickFileButton 
            Caption         =   "&Play"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1920
            TabIndex        =   43
            ToolTipText     =   "Start or resume tickfile replay"
            Top             =   3240
            Width           =   615
         End
         Begin TradingUI27.TickfileOrganiser TickfileOrganiser1 
            Height          =   2520
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   4445
         End
         Begin MSComctlLib.ProgressBar ReplayProgressBar 
            Height          =   135
            Left            =   120
            TabIndex        =   47
            Top             =   4200
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   238
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label20 
            Caption         =   "Replay speed"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label ReplayContractLabel 
            Height          =   975
            Left            =   120
            TabIndex        =   49
            Top             =   4440
            Width           =   3855
         End
         Begin VB.Label ReplayProgressLabel 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   3960
            Width           =   3855
         End
      End
      Begin VB.PictureBox HistChartPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   24
         Top             =   0
         Width           =   4125
         Begin VB.TextBox NumHistBarsText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3000
            TabIndex        =   31
            Text            =   "500"
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox HistSessionOnlyCheck 
            Caption         =   "Session only"
            Height          =   375
            Left            =   2760
            TabIndex        =   30
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Caption         =   "Change chart styles"
            Height          =   1335
            Left            =   120
            TabIndex        =   25
            Top             =   7320
            Width           =   3855
            Begin VB.PictureBox ChangeHistChartStylesPicture 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   60
               ScaleHeight     =   975
               ScaleWidth      =   3735
               TabIndex        =   26
               Top             =   240
               Width           =   3735
               Begin VB.CommandButton ChangeHistChartStylesButton 
                  Caption         =   "Change ALL historical chart styles"
                  Height          =   495
                  Left            =   480
                  TabIndex        =   27
                  Top             =   480
                  Width           =   2775
               End
               Begin VB.Label Label9 
                  Caption         =   "Click this button to change the style of all existing historical charts to the style selected above."
                  Height          =   495
                  Left            =   120
                  TabIndex        =   28
                  Top             =   0
                  Width           =   3495
               End
            End
         End
         Begin TradingUI27.ContractSearch HistContractSearch 
            Height          =   4455
            Left            =   120
            TabIndex        =   29
            Top             =   2760
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7858
         End
         Begin TradingUI27.TimeframeSelector HistTimeframeSelector 
            Height          =   330
            Left            =   1920
            TabIndex        =   32
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker ToDatePicker 
            Height          =   375
            Left            =   1920
            TabIndex        =   33
            Top             =   1800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyy-MM-dd HH:mm"
            Format          =   20774915
            CurrentDate     =   39365
         End
         Begin MSComCtl2.DTPicker FromDatePicker 
            Height          =   375
            Left            =   1920
            TabIndex        =   34
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   128
            CalendarForeColor=   16711680
            CalendarTitleBackColor=   16777215
            CalendarTitleForeColor=   12632256
            CalendarTrailingForeColor=   65280
            CheckBox        =   -1  'True
            CustomFormat    =   "yyy-MM-dd HH:mm"
            Format          =   20774915
            CurrentDate     =   39365
         End
         Begin TWControls40.TWImageCombo HistChartStylesCombo 
            Height          =   330
            Left            =   1920
            TabIndex        =   35
            Top             =   2280
            Width           =   2055
            _ExtentX        =   3625
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
            MouseIcon       =   "FeaturesPanel.ctx":008C
            Text            =   ""
         End
         Begin VB.Label Label5 
            Caption         =   "To"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "From"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Timeframe"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Number of history bars"
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Style"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Width           =   1455
         End
      End
      Begin VB.PictureBox LiveChartPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   11
         Top             =   0
         Width           =   4125
         Begin VB.TextBox NumHistoryBarsText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3000
            TabIndex        =   19
            Text            =   "500"
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox SessionOnlyCheck 
            Caption         =   "Session only"
            Height          =   375
            Left            =   2760
            TabIndex        =   18
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CommandButton ChartButton 
            Caption         =   "Show &Chart"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3000
            TabIndex        =   17
            Top             =   2040
            Width           =   975
         End
         Begin VB.Frame Frame1 
            Caption         =   "Change chart styles"
            Height          =   1335
            Left            =   120
            TabIndex        =   13
            Top             =   3360
            Width           =   3855
            Begin VB.PictureBox ChangeLiveChartStylesPicture 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   60
               ScaleHeight     =   975
               ScaleWidth      =   3735
               TabIndex        =   14
               Top             =   240
               Width           =   3735
               Begin VB.CommandButton ChangeLiveChartStylesButton 
                  Caption         =   "Change ALL live chart styles"
                  Height          =   495
                  Left            =   480
                  TabIndex        =   15
                  Top             =   480
                  Width           =   2775
               End
               Begin VB.Label Label7 
                  Caption         =   "Click this button to change the style of all existing live charts to the style selected above."
                  Height          =   495
                  Left            =   120
                  TabIndex        =   16
                  Top             =   0
                  Width           =   3495
               End
            End
         End
         Begin TWControls40.TWImageCombo LiveChartStylesCombo 
            Height          =   330
            Left            =   1920
            TabIndex        =   12
            Top             =   1560
            Width           =   2055
            _ExtentX        =   3625
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
            MouseIcon       =   "FeaturesPanel.ctx":00A8
            Text            =   ""
         End
         Begin TradingUI27.TimeframeSelector LiveChartTimeframeSelector 
            Height          =   330
            Left            =   1920
            TabIndex        =   20
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
         End
         Begin VB.Label Label18 
            Caption         =   "Timeframe"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Number of history bars"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Style"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.PictureBox TickersPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   0
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   7
         Top             =   0
         Width           =   4125
         Begin TradingUI27.ContractSearch LiveContractSearch 
            Height          =   5415
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   9551
         End
         Begin VB.CommandButton ChartButton1 
            Caption         =   "Chart"
            Enabled         =   0   'False
            Height          =   375
            Left            =   720
            TabIndex        =   2
            Top             =   6360
            Width           =   975
         End
         Begin VB.CommandButton StopTickerButton 
            Appearance      =   0  'Flat
            Caption         =   "Sto&p"
            Enabled         =   0   'False
            Height          =   375
            Left            =   720
            TabIndex        =   5
            Top             =   5880
            Width           =   975
         End
         Begin VB.CommandButton OrderTicketButton 
            Caption         =   "&Order ticket"
            Enabled         =   0   'False
            Height          =   375
            Left            =   720
            TabIndex        =   4
            Top             =   6840
            Width           =   975
         End
         Begin VB.CommandButton MarketDepthButton 
            Caption         =   "&Mkt depth"
            Enabled         =   0   'False
            Height          =   375
            Left            =   720
            TabIndex        =   3
            Top             =   7320
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox HidePicture 
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3720
      MouseIcon       =   "FeaturesPanel.ctx":00C4
      MousePointer    =   99  'Custom
      Picture         =   "FeaturesPanel.ctx":0216
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      ToolTipText     =   "Hide Features Panel"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox UnpinPicture 
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3360
      MouseIcon       =   "FeaturesPanel.ctx":07A0
      MousePointer    =   99  'Custom
      Picture         =   "FeaturesPanel.ctx":08F2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      ToolTipText     =   "Unpin Features Panel"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox PinPicture 
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3360
      MouseIcon       =   "FeaturesPanel.ctx":0E7C
      MousePointer    =   99  'Custom
      Picture         =   "FeaturesPanel.ctx":0FCE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      ToolTipText     =   "Pin Features Panel"
      Top             =   30
      Width           =   240
   End
   Begin MSComctlLib.TabStrip FeaturesTabStrip 
      Height          =   640
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   1138
      TabWidthStyle   =   1
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Tickers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Live chart"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Historical chart"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Replay tickfiles"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Config"
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
Attribute VB_Name = "FeaturesPanel"
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

Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Hide()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute Mouseup.VB_UserMemId = -607
Event Pin()
Event Unpin()

'@================================================================================
' Enums
'@================================================================================

Private Enum FeaturesTabIndexNumbers
    FeaturesTabIndexTickers
    FeaturesTabIndexLiveCharts
    FeaturesTabIndexHistoricalCharts
    FeaturesTabIndexTickfileReplay
    FeaturesTabIndexConfig
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "FeaturesPane"

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildAPI                              As TradeBuildAPI
Private mAppInstanceConfig                          As ConfigurationSection

Private WithEvents mTickerGrid                      As TickerGrid
Attribute mTickerGrid.VB_VarHelpID = -1
Private mTickfileOrdersSummary                      As OrdersSummary
Private mTickfileExecutionsSummary                  As ExecutionsSummary
Private mChartForms                                 As ChartForms
Private mOrderTicket                                As fOrderTicket

Private WithEvents mReplayController                As ReplayController
Attribute mReplayController.VB_VarHelpID = -1
Private WithEvents mTickfileReplayTC                As TaskController
Attribute mTickfileReplayTC.VB_VarHelpID = -1

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
'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

OrderTicketButton.Enabled = Not (getSelectedDataSource Is Nothing)

Dim lDataSource As IMarketDataSource
Set lDataSource = ev.Source

Select Case ev.State
Case MarketDataSourceStates.MarketDataSourceStateCreated

Case MarketDataSourceStates.MarketDataSourceStateReady
Case MarketDataSourceStates.MarketDataSourceStateRunning
    If lDataSource Is getSelectedDataSource Then
        MarketDepthButton.Enabled = True
        ChartButton.Enabled = True
        ChartButton1.Enabled = True
    End If
    
Case MarketDataSourceStates.MarketDataSourceStatePaused

Case MarketDataSourceStates.MarketDataSourceStateStopped
    If getSelectedDataSource Is Nothing Then
        StopTickerButton.Enabled = False
        MarketDepthButton.Enabled = False
        ChartButton.Enabled = False
        ChartButton1.Enabled = False
    End If
    
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ChangeHistChartStylesButton_Click()
Const ProcName As String = "ChangeHistChartStylesButton_Click"
On Error GoTo Err

setAllChartStyles HistChartStylesCombo.Text, True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChangeLiveChartStylesButton_Click()
Const ProcName As String = "ChangeLiveChartStylesButton_Click"
On Error GoTo Err

setAllChartStyles LiveChartStylesCombo.Text, False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartButton_Click()
Const ProcName As String = "ChartButton_Click"
On Error GoTo Err

Dim lTicker As Ticker
For Each lTicker In mTickerGrid.SelectedTickers
    createChart lTicker
Next

clearSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartButton1_Click()
Const ProcName As String = "ChartButton1_Click"
On Error GoTo Err

ChartButton_Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FeaturesSSTab_Click(PreviousTab As Integer)
Const ProcName As String = "FeaturesSSTab_Click"
On Error GoTo Err

Select Case FeaturesSSTab.Tab
Case FeaturesTabIndexNumbers.FeaturesTabIndexConfig
    ConfigEditorButton.SetFocus
Case FeaturesTabIndexNumbers.FeaturesTabIndexHistoricalCharts
    HistContractSearch.SetFocus
Case FeaturesTabIndexNumbers.FeaturesTabIndexLiveCharts
    LiveChartTimeframeSelector.SetFocus
    If mTickerGrid.SelectedTickers.Count > 0 Then ChartButton.Default = True
Case FeaturesTabIndexNumbers.FeaturesTabIndexTickers
    LiveContractSearch.SetFocus
    If mTickerGrid.SelectedTickers.Count > 0 Then ChartButton1.Default = True
Case FeaturesTabIndexNumbers.FeaturesTabIndexTickfileReplay
    If Not mReplayController Is Nothing Then
        If PlayTickFileButton.Enabled Then
            PlayTickFileButton.Default = True
        ElseIf StopReplayButton.Enabled Then
            StopReplayButton.Default = True
        End If
    End If
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FeaturesTabStrip_Click()
Const ProcName As String = "FeaturesTabStrip_Click"
On Error GoTo Err

FeaturesSSTab.SetFocus
FeaturesSSTab.Tab = FeaturesTabStrip.SelectedItem.Index - 1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HidePicture_Click()
RaiseEvent Hide
End Sub

Private Sub HistChartStylesCombo_Click()
Const ProcName As String = "HistChartStylesCombo_Change"
On Error GoTo Err

mAppInstanceConfig.SetSetting ConfigSettingAppCurrentHistChartStyle, HistChartStylesCombo.Text

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub HistContractSearch_Action()
Const ProcName As String = "HistContractSearch_Action"
On Error GoTo Err

createHistoricCharts HistContractSearch.SelectedContracts

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HistContractSearch_NoContracts()
Const ProcName As String = "HistContractSearch_NoContracts"
On Error GoTo Err

gModelessMsgBox "No contracts found", vbExclamation, "Attention"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveChartStylesCombo_Click()
Const ProcName As String = "LiveChartStylesCombo_Change"
On Error GoTo Err

mAppInstanceConfig.SetSetting ConfigSettingAppCurrentChartStyle, LiveChartStylesCombo.Text

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub LiveChartTimeframeSelector_Click()
Const ProcName As String = "LiveChartTimeframeSelector_Click"
On Error GoTo Err

setChartButtonTooltip

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveContractSearch_Action()
Const ProcName As String = "LiveContractSearch_Action"
On Error GoTo Err

Dim lPreferredRow As Long
lPreferredRow = CLng(LiveContractSearch.Cookie)

Dim lContract As IContract
For Each lContract In LiveContractSearch.SelectedContracts
    mTickerGrid.StartTickerFromContract lContract, lPreferredRow
    If lPreferredRow <> 0 Then lPreferredRow = lPreferredRow + 1
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveContractSearch_NoContracts()
Const ProcName As String = "LiveContractSearch_NoContracts"
On Error GoTo Err

ModelessMsgBox "No contracts found", vbExclamation, "Attention", UserControl.ContainerHwnd

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub MarketDepthButton_Click()
Const ProcName As String = "MarketDepthButton_Click"
On Error GoTo Err

Dim lTicker As Ticker
For Each lTicker In mTickerGrid.SelectedTickers
    showMarketDepthForm lTicker
Next

clearSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NumHistBarsText_Validate(Cancel As Boolean)
Const ProcName As String = "NumHistBarsText_Validate"
On Error GoTo Err

If Not IsInteger(NumHistBarsText.Text, 0, 2000) Then Cancel = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NumHistoryBarsText_Validate(Cancel As Boolean)
Const ProcName As String = "NumHistoryBarsText_Validate"
On Error GoTo Err

If Not IsInteger(NumHistoryBarsText.Text, 0, 2000) Then Cancel = True

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

Private Sub PauseReplayButton_Click()
Const ProcName As String = "PauseReplayButton_Click"
On Error GoTo Err

PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
LogMessage "Tickfile replay paused"
mReplayController.PauseReplay

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub PinPicture_Click()
RaiseEvent Pin
End Sub

Private Sub PlayTickFileButton_Click()
Const ProcName As String = "PlayTickFileButton_Click"
On Error GoTo Err

PlayTickFileButton.Enabled = False
PauseReplayButton.Enabled = True
StopReplayButton.Enabled = True
ReplayProgressBar.Visible = True

If mReplayController Is Nothing Then
    TickfileOrganiser1.Enabled = False
    
    Dim lTickfileDataManager As TickfileDataManager

    Set lTickfileDataManager = CreateTickDataManager(TickfileOrganiser1.TickFileSpecifiers, _
                                                mTradeBuildAPI.TickfileStoreInput, _
                                                mTradeBuildAPI.StudyLibraryManager, _
                                                mTradeBuildAPI.ContractStorePrimary, _
                                                mTradeBuildAPI.ContractStoreSecondary, _
                                                MarketDataSourceOptUseExchangeTimeZone, _
                                                , _
                                                , _
                                                CInt(ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)), _
                                                250)
    mTickfileOrdersSummary.Initialise lTickfileDataManager
    
    Dim lOrderManager As New OrderManager
    mTickfileOrdersSummary.MonitorPositions lOrderManager.PositionManagersSimulated
    mTickfileExecutionsSummary.MonitorPositions lOrderManager.PositionManagersSimulated
    
    Set mReplayController = lTickfileDataManager.ReplayController
    
    Dim lTickers As Tickers
    Set lTickers = CreateTickers(lTickfileDataManager, mTradeBuildAPI.StudyLibraryManager, mTradeBuildAPI.HistoricalDataStoreInput, lOrderManager, , mTradeBuildAPI.OrderSubmitterFactorySimulated)
    
    Dim i As Long
    For i = 1 To TickfileOrganiser1.TickfileCount
        Dim lTicker As Ticker
        Set lTicker = lTickers.CreateTicker(mReplayController.TickStream(i - 1).ContractFuture, False)
        mTickerGrid.AddTickerFromDataSource lTicker
    Next
    
    LogMessage "Tickfile replay started"
    Set mTickfileReplayTC = mReplayController.StartReplay
ElseIf mReplayController.ReplayInProgress Then
    LogMessage "Tickfile replay resumed"
    mReplayController.ResumeReplay
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ReplaySpeedCombo_Click()
Const ProcName As String = "ReplaySpeedCombo_Click"
On Error GoTo Err

If Not mReplayController Is Nothing Then
    mReplayController.ReplaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub StopReplayButton_Click()
Const ProcName As String = "StopReplayButton_Click"
On Error GoTo Err

stopTickfileReplay

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub StopTickerButton_Click()
Const ProcName As String = "StopTickerButton_Click"
On Error GoTo Err

StopSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UnpinPicture_Click()
RaiseEvent Unpin
End Sub

'================================================================================
' mReplayController Event Handlers
'================================================================================

Private Sub mReplayController_ReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Long)
Const ProcName As String = "mReplayController_ReplayProgress"
On Error GoTo Err

ReplayProgressBar.Value = pPercentComplete
ReplayProgressLabel.caption = pTickfileTimestamp & _
                                "  Processed " & _
                                pEventsPlayed & _
                                " events"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTickfileReplayTC Event Handlers
'================================================================================

Private Sub mTickfileReplayTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTickfileReplayTC_Completed"
On Error GoTo Err

Set mReplayController = Nothing

MarketDepthButton.Enabled = False
PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
StopReplayButton.Enabled = False

ReplayProgressBar.Value = 0
ReplayProgressBar.Visible = False
ReplayContractLabel.caption = ""
ReplayProgressLabel.caption = ""

TickfileOrganiser1.Enabled = True

LogMessage "Tickfile replay completed"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

'@================================================================================
' Properties
'@================================================================================

Public Property Let BackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "BackColor"
On Error GoTo Err

TickersPicture.BackColor = Value
LiveChartPicture.BackColor = Value
ChangeLiveChartStylesPicture.BackColor = Value
ChangeHistChartStylesPicture.BackColor = Value
HistChartPicture.BackColor = Value
ReplayTickerPicture.BackColor = Value
ConfigPicture.BackColor = Value

FromDatePicker.CalendarBackColor = Value
ToDatePicker.CalendarBackColor = Value

LiveContractSearch.BackColor = Value
HistContractSearch.BackColor = Value

On Error Resume Next
Dim lControl As Control
For Each lControl In UserControl.Controls
    If TypeOf lControl Is CommandButton Or _
        TypeOf lControl Is PictureBox Or _
        TypeOf lControl Is TextBox Or _
        TypeOf lControl Is ComboBox Or _
        TypeOf lControl Is TWImageCombo _
    Then
    Else
        lControl.BackColor = Value
    End If
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
BackColor = TickersPicture.BackColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "ForeColor"
On Error GoTo Err

LiveContractSearch.ForeColor = Value
HistContractSearch.ForeColor = Value


FromDatePicker.CalendarForeColor = Value
ToDatePicker.CalendarForeColor = Value

On Error Resume Next
Dim lControl As Control
For Each lControl In UserControl.Controls
    If TypeOf lControl Is Label Or _
        TypeOf lControl Is CheckBox Or _
        TypeOf lControl Is Frame Or _
        TypeOf lControl Is TimeframeSelector _
    Then lControl.ForeColor = Value
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
ForeColor = TickersPicture.BackColor
End Property

Public Property Let TextboxBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "TextboxBackColor"
On Error GoTo Err

LiveContractSearch.TextboxBackColor = Value
HistContractSearch.TextboxBackColor = Value

FromDatePicker.CalendarTitleBackColor = Value
ToDatePicker.CalendarTitleBackColor = Value

Dim lControl As Control
For Each lControl In UserControl.Controls
    If TypeOf lControl Is TextBox Or _
        TypeOf lControl Is ComboBox Or _
        TypeOf lControl Is TWImageCombo _
    Then lControl.BackColor = Value
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextboxBackColor() As OLE_COLOR

End Property

Public Property Let TextboxForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "TextboxForeColor"
On Error GoTo Err

LiveContractSearch.TextboxForeColor = Value
HistContractSearch.TextboxForeColor = Value

FromDatePicker.CalendarTitleForeColor = Value
ToDatePicker.CalendarTitleForeColor = Value

FromDatePicker.CalendarTrailingForeColor = toneDown(Value)
ToDatePicker.CalendarTrailingForeColor = toneDown(Value)

Dim lControl As Control
For Each lControl In UserControl.Controls
    If TypeOf lControl Is TextBox Or _
        TypeOf lControl Is ComboBox Or _
        TypeOf lControl Is TWImageCombo _
    Then lControl.ForeColor = Value
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextboxForeColor() As OLE_COLOR

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

LogMessage "Stopping tickfile replay"
' prevent event handler being fired on completion, which would
' reload the form again
Set mTickfileReplayTC = Nothing
stopTickfileReplay


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pPinned As Boolean, _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByVal pTickerGrid As TickerGrid, _
                ByVal pTickfileOrdersSummary As OrdersSummary, _
                ByVal pTickfileExecutionsSummary As ExecutionsSummary, _
                ByVal pChartForms As ChartForms, _
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

Set mTradeBuildAPI = pTradeBuildAPI
Set mAppInstanceConfig = pAppInstanceConfig
Set mTickerGrid = pTickerGrid
Set mTickfileOrdersSummary = pTickfileOrdersSummary
Set mTickfileExecutionsSummary = pTickfileExecutionsSummary
Set mChartForms = pChartForms
Set mOrderTicket = pOrderTicket

LogMessage "Initialising Features Panel: Setting up contract search"
setupContractSearch

setupReplaySpeedCombo

LogMessage "Initialising Features Panel: Setting up tickfile organiser"
setupTickfileOrganiser

LogMessage "Initialising Features Panel: Setting up timeframeselectors"
setupTimeframeSelectors

LogMessage "Initialising Features Panel: setting current chart styles"
setCurrentChartStyles

FromDatePicker.Value = DateAdd("m", -1, Now)
FromDatePicker.Value = Empty    ' clear the checkbox
ToDatePicker.Value = Now

CurrentConfigNameText = mAppInstanceConfig.InstanceQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadContractsForUserChoice( _
                ByVal pContracts As IContracts, _
                ByVal pPreferredTickerGridIndex)
Const ProcName As String = "LoadContractsForUserChoice"
On Error GoTo Err

LiveContractSearch.LoadContracts pContracts, pPreferredTickerGridIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowTickersPane()
Const ProcName As String = "ShowTickersPane"
On Error GoTo Err

If FeaturesSSTab.Tab <> FeaturesTabIndexNumbers.FeaturesTabIndexTickers Then FeaturesTabStrip.Tabs(FeaturesTabIndexNumbers.FeaturesTabIndexTickers + 1).Selected = True
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearSelectedTickers()
Const ProcName As String = "clearSelectedTickers"
On Error GoTo Err

mTickerGrid.DeselectSelectedTickers
handleSelectedTickers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createChart(ByVal pTicker As Ticker)
Const ProcName As String = "createChart"
On Error GoTo Err

If Not pTicker.State = MarketDataSourceStateRunning Then Exit Sub

Dim tp As TimePeriod
Set tp = LiveChartTimeframeSelector.TimePeriod

Dim lConfig As ConfigurationSection

If Not pTicker.IsTickReplay Then
    Set lConfig = mAppInstanceConfig.AddConfigurationSection(ConfigSectionCharts)
End If

mChartForms.Add pTicker, _
                tp, _
                pTicker.Timeframes, _
                mTradeBuildAPI.BarFormatterLibManager, _
                mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator, _
                lConfig, _
                CreateChartSpecifier(CLng(NumHistoryBarsText.Text), Not (SessionOnlyCheck = vbChecked)), _
                ChartStylesManager.Item(LiveChartStylesCombo.SelectedItem.Text), _
                gMainForm

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createHistoricCharts( _
                ByVal pContracts As IContracts)
Const ProcName As String = "createHistoricCharts"
On Error GoTo Err

Dim lConfig As ConfigurationSection
Set lConfig = mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionHistoricCharts)

Dim lContract As IContract
For Each lContract In pContracts
    Dim fromDate As Date
    If IsNull(FromDatePicker.Value) Then
        fromDate = CDate(0)
    Else
        fromDate = DateSerial(FromDatePicker.Year, FromDatePicker.Month, FromDatePicker.Day) + _
                    TimeSerial(FromDatePicker.Hour, FromDatePicker.Minute, 0)
    End If
    
    Dim toDate As Date
    If IsNull(ToDatePicker.Value) Then
        toDate = Now
    Else
        toDate = DateSerial(ToDatePicker.Year, ToDatePicker.Month, ToDatePicker.Day) + _
                    TimeSerial(ToDatePicker.Hour, ToDatePicker.Minute, 0)
    End If
    
    mChartForms.AddHistoric HistTimeframeSelector.TimePeriod, _
                        CreateFuture(lContract), _
                        mTradeBuildAPI.StudyLibraryManager.CreateStudyManager, _
                        mTradeBuildAPI.HistoricalDataStoreInput, _
                        mTradeBuildAPI.BarFormatterLibManager, _
                        lConfig, _
                        CreateChartSpecifier(CLng(NumHistBarsText.Text), Not (HistSessionOnlyCheck = vbChecked), fromDate, toDate), _
                        ChartStylesManager.Item(LiveChartStylesCombo.SelectedItem.Text), _
                        gMainForm

Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
    StopTickerButton.Enabled = False
    ChartButton.Enabled = False
    ChartButton1.Enabled = False
    MarketDepthButton.Enabled = False
    OrderTicketButton.Enabled = False
Else
    StopTickerButton.Enabled = True
    
    ChartButton.Enabled = False
    ChartButton1.Enabled = False
    MarketDepthButton.Enabled = False
    OrderTicketButton.Enabled = False
    
    If FeaturesSSTab.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexLiveCharts Then
        ChartButton.Default = True
    ElseIf FeaturesSSTab.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexTickers Then
        ChartButton1.Default = True
    End If
    
    Dim lTicker As Ticker
    Set lTicker = getSelectedDataSource
    If lTicker Is Nothing Then
    ElseIf lTicker.State = MarketDataSourceStateRunning Then
        ChartButton.Enabled = True
        ChartButton1.Enabled = True
        Dim lContract As IContract
        Set lContract = lTicker.ContractFuture.Value
        If (lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled) And lContract.Specifier.SecType <> SecTypeIndex Then
            OrderTicketButton.Enabled = True
            MarketDepthButton.Enabled = True
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadStyleComboItems(ByVal pComboItems As ComboItems)
Const ProcName As String = "loadStyleComboItems"
On Error GoTo Err

pComboItems.Clear

Dim lStyle As ChartStyle
For Each lStyle In ChartStylesManager
    pComboItems.Add , lStyle.name, lStyle.name
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setAllChartStyles(ByVal pStyleName As String, ByVal pHistorical As Boolean)
Const ProcName As String = "setAllChartStyles"
On Error GoTo Err

mChartForms.SetStyle ChartStylesManager.Item(pStyleName), pHistorical

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setChartButtonTooltip()
Const ProcName As String = "setChartButtonTooltip"
On Error GoTo Err

Dim tp As TimePeriod
Set tp = LiveChartTimeframeSelector.TimePeriod

ChartButton.ToolTipText = "Show " & tp.ToString & " chart"
ChartButton1.ToolTipText = ChartButton.ToolTipText

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCurrentChartStyle( _
                ByVal pComboItems As ComboItems, _
                ByVal pConfigSettingName As String)
Const ProcName As String = "setCurrentChartStyle"
On Error GoTo Err

loadStyleComboItems pComboItems

Dim lStyleName As String
lStyleName = mAppInstanceConfig.GetSetting(pConfigSettingName, "")

If ChartStylesManager.Contains(lStyleName) Then
    pComboItems.Item(lStyleName).Selected = True
Else
    pComboItems.Item(ChartStyleNameAppDefault).Selected = True
    mAppInstanceConfig.SetSetting pConfigSettingName, ChartStyleNameAppDefault
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCurrentChartStyles()
Const ProcName As String = "setCurrentChartStyles"
On Error GoTo Err

setCurrentChartStyle LiveChartStylesCombo.ComboItems, ConfigSettingAppCurrentChartStyle
setCurrentChartStyle HistChartStylesCombo.ComboItems, ConfigSettingAppCurrentHistChartStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupContractSearch()
Const ProcName As String = "setupContractSearch"
On Error GoTo Err

LiveContractSearch.Initialise mTradeBuildAPI.ContractStorePrimary, mTradeBuildAPI.ContractStoreSecondary
LiveContractSearch.IncludeHistoricalContracts = False

HistContractSearch.Initialise mTradeBuildAPI.ContractStorePrimary, mTradeBuildAPI.ContractStoreSecondary
HistContractSearch.IncludeHistoricalContracts = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupReplaySpeedCombo()
Const ProcName As String = "setupReplaySpeedCombo"
On Error GoTo Err

ReplaySpeedCombo.AddItem "Continuous"
ReplaySpeedCombo.ItemData(0) = 0
ReplaySpeedCombo.AddItem "Actual speed"
ReplaySpeedCombo.ItemData(1) = 1
ReplaySpeedCombo.AddItem "2x Actual speed"
ReplaySpeedCombo.ItemData(2) = 2
ReplaySpeedCombo.AddItem "4x Actual speed"
ReplaySpeedCombo.ItemData(3) = 4
ReplaySpeedCombo.AddItem "8x Actual speed"
ReplaySpeedCombo.ItemData(4) = 8

ReplaySpeedCombo.Text = "Actual speed"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTickfileOrganiser()
Const ProcName As String = "setupTickfileOrganiser"
On Error GoTo Err

TickfileOrganiser1.Initialise mTradeBuildAPI.TickfileStoreInput, mTradeBuildAPI.ContractStorePrimary, mTradeBuildAPI.ContractStoreSecondary
TickfileOrganiser1.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTimeframeSelectors()
Const ProcName As String = "setupTimeframeSelectors"
On Error GoTo Err

' now set up the timeframe selectors, which depends on what timeframes the historical data service
' provider supports (it obtains this info from TradeBuild)
LiveChartTimeframeSelector.Initialise mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator     ' use the default settings built-in to the control
LiveChartTimeframeSelector.SelectTimeframe GetTimePeriod(5, TimePeriodMinute)
HistTimeframeSelector.Initialise mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator
HistTimeframeSelector.SelectTimeframe GetTimePeriod(5, TimePeriodMinute)

setChartButtonTooltip

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showMarketDepthForm(ByVal pTicker As Ticker)
Const ProcName As String = "showMarketDepthForm"
On Error GoTo Err

If Not pTicker.State = MarketDataSourceStateRunning Then Exit Sub

Dim mktDepthForm As New fMarketDepth
mktDepthForm.numberOfRows = 100
mktDepthForm.Ticker = pTicker

mktDepthForm.Show vbModeless

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

Private Sub StopSelectedTickers()
Const ProcName As String = "StopSelectedTickers"
On Error GoTo Err

Dim lTickers As SelectedTickers
Set lTickers = mTickerGrid.SelectedTickers

mTickerGrid.StopSelectedTickers

Dim lTicker As IMarketDataSource
For Each lTicker In lTickers
    lTicker.Finish
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub stopTickfileReplay()
Const ProcName As String = "stopTickfileReplay"
On Error GoTo Err

PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
StopReplayButton.Enabled = False
ChartButton.Enabled = False
ChartButton1.Enabled = False
If Not mReplayController Is Nothing Then
    mReplayController.StopReplay
    Set mReplayController = Nothing
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Function toneDown(ByVal pColor As Long) As Long
If (pColor And &H80000000) Then pColor = GetSysColor(pColor And &HFFFFFF)

toneDown = (((pColor And &HFF0000) / &H20000) And &HFF0000) + _
            (((pColor And &HFF00) / &H200) And &HFF00) + _
            ((pColor And &HFF) / &H2)
End Function



