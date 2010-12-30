VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#243.0#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fTradeSkilDemo 
   Caption         =   "TradeSkil Demo Edition Version 2.6"
   ClientHeight    =   9960
   ClientLeft      =   225
   ClientTop       =   345
   ClientWidth     =   16665
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   16665
   Begin VB.PictureBox ShowFeaturesPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   16320
      MouseIcon       =   "fTradeSkilDemo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":0152
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   57
      ToolTipText     =   "Show features"
      Top             =   9345
      Width           =   480
   End
   Begin VB.PictureBox HideFeaturesPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   16320
      MouseIcon       =   "fTradeSkilDemo.frx":0594
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":06E6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   56
      ToolTipText     =   "Hide features"
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox ShowControlsPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      MouseIcon       =   "fTradeSkilDemo.frx":0B28
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":0C7A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   55
      ToolTipText     =   "Show controls"
      Top             =   440
      Width           =   480
   End
   Begin VB.PictureBox HideControlsPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3900
      MouseIcon       =   "fTradeSkilDemo.frx":10BC
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":120E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   54
      ToolTipText     =   "Hide controls"
      Top             =   440
      Width           =   480
   End
   Begin TabDlg.SSTab FeaturesSSTAB 
      Height          =   4455
      Left            =   4320
      TabIndex        =   6
      Top             =   5040
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&1. Orders"
      TabPicture(0)   =   "fTradeSkilDemo.frx":1650
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SimulatedOrdersSummary"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LiveOrdersSummary"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "OrdersSummaryTabStrip"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ModifyOrderPlexButton"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CancelOrderPlexButton"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ClosePositionsButton"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "OrderTicket1Button"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&2. Executions"
      TabPicture(1)   =   "fTradeSkilDemo.frx":166C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LiveExecutionsSummary"
      Tab(1).Control(1)=   "ExecutionsSummaryTabStrip"
      Tab(1).Control(2)=   "SimulatedExecutionsSummary"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&3. Log"
      TabPicture(2)   =   "fTradeSkilDemo.frx":1688
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LogText"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton OrderTicket1Button 
         Caption         =   "Order  Ticket"
         Height          =   495
         Left            =   11160
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox LogText 
         Height          =   3975
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Status messages"
         Top             =   120
         Width           =   11955
      End
      Begin VB.CommandButton ClosePositionsButton 
         Caption         =   "Close all positions!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11160
         TabIndex        =   10
         Top             =   3510
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderPlexButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   11160
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderPlexButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   11160
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin MSComctlLib.TabStrip OrdersSummaryTabStrip 
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         MultiRow        =   -1  'True
         Style           =   1
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
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
         EndProperty
      End
      Begin TradeBuildUI26.OrdersSummary LiveOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradeBuildUI26.OrdersSummary SimulatedOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradeBuildUI26.ExecutionsSummary LiveExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   51
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6376
      End
      Begin MSComctlLib.TabStrip ExecutionsSummaryTabStrip 
         Height          =   375
         Left            =   -74880
         TabIndex        =   52
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         MultiRow        =   -1  'True
         Style           =   1
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
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
         EndProperty
      End
      Begin TradeBuildUI26.ExecutionsSummary SimulatedExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   53
         Top             =   120
         Width           =   11995
         _ExtentX        =   21167
         _ExtentY        =   6376
      End
   End
   Begin MSComctlLib.TabStrip ControlsTabStrip 
      Height          =   580
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1032
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
   End
   Begin TabDlg.SSTab ControlsSSTab 
      Height          =   9015
      Left            =   120
      TabIndex        =   31
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   15901
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   2
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "fTradeSkilDemo.frx":16A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LiveContractSearch"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "fTradeSkilDemo.frx":16C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChartButton"
      Tab(1).Control(1)=   "SessionOnlyCheck"
      Tab(1).Control(2)=   "NumHistoryBarsText"
      Tab(1).Control(3)=   "LiveChartTimeframeSelector"
      Tab(1).Control(4)=   "Label22"
      Tab(1).Control(5)=   "Label18"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "fTradeSkilDemo.frx":16DC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "HistContractSearch"
      Tab(2).Control(1)=   "HistSessionOnlyCheck"
      Tab(2).Control(2)=   "NumHistBarsText"
      Tab(2).Control(3)=   "HistTimeframeSelector"
      Tab(2).Control(4)=   "ToDatePicker"
      Tab(2).Control(5)=   "FromDatePicker"
      Tab(2).Control(6)=   "Label3"
      Tab(2).Control(7)=   "Label2"
      Tab(2).Control(8)=   "Label4"
      Tab(2).Control(9)=   "Label5"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "fTradeSkilDemo.frx":16F8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SkipReplayButton"
      Tab(3).Control(1)=   "PlayTickFileButton"
      Tab(3).Control(2)=   "PauseReplayButton"
      Tab(3).Control(3)=   "StopReplayButton"
      Tab(3).Control(4)=   "ReplaySpeedCombo"
      Tab(3).Control(5)=   "SelectTickfilesButton"
      Tab(3).Control(6)=   "ClearTickfileListButton"
      Tab(3).Control(7)=   "TickfileList"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "ReplayProgressBar"
      Tab(3).Control(9)=   "ReplayProgressLabel"
      Tab(3).Control(10)=   "ReplayContractLabel"
      Tab(3).Control(11)=   "Label20"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "fTradeSkilDemo.frx":1714
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label6"
      Tab(4).Control(1)=   "CurrentConfigNameText"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "ConfigEditorButton"
      Tab(4).ControlCount=   3
      Begin TradeBuildUI26.ContractSearch HistContractSearch 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   20
         Top             =   3000
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8916
         IncludeHistoricalContracts=   -1  'True
      End
      Begin TradeBuildUI26.ContractSearch LiveContractSearch 
         Height          =   5415
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9551
      End
      Begin VB.CommandButton ConfigEditorButton 
         Caption         =   "Show config editor"
         Height          =   375
         Left            =   -72840
         TabIndex        =   29
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox CurrentConfigNameText 
         Height          =   285
         Left            =   -74640
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3375
      End
      Begin VB.CommandButton SkipReplayButton 
         Caption         =   "S&kip"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -72360
         TabIndex        =   26
         ToolTipText     =   "Pause tickfile replay"
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73800
         TabIndex        =   24
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73080
         TabIndex        =   25
         ToolTipText     =   "Pause tickfile replay"
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -71640
         TabIndex        =   27
         ToolTipText     =   "Stop tickfile replay"
         Top             =   3240
         Width           =   615
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "fTradeSkilDemo.frx":1730
         Left            =   -73800
         List            =   "fTradeSkilDemo.frx":1732
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CommandButton SelectTickfilesButton 
         Caption         =   "..."
         Height          =   375
         Left            =   -72120
         TabIndex        =   21
         ToolTipText     =   "Select tickfile(s)"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton ClearTickfileListButton 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71520
         TabIndex        =   22
         ToolTipText     =   "Clear tickfile list"
         Top             =   360
         Width           =   495
      End
      Begin VB.ListBox TickfileList 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   840
         Width           =   3855
      End
      Begin VB.CheckBox HistSessionOnlyCheck 
         Caption         =   "Session only"
         Height          =   375
         Left            =   -72240
         TabIndex        =   17
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox NumHistBarsText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72000
         TabIndex        =   16
         Text            =   "500"
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Selected tickers"
         Height          =   1215
         Left            =   120
         TabIndex        =   34
         Top             =   6000
         Width           =   3615
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   3255
            TabIndex        =   35
            Top             =   240
            Width           =   3255
            Begin VB.CommandButton Chart1Button 
               Caption         =   "Chart"
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   1
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton StopTickerButton 
               Caption         =   "Sto&p"
               Enabled         =   0   'False
               Height          =   375
               Left            =   1080
               TabIndex        =   4
               Top             =   480
               Width           =   975
            End
            Begin VB.CommandButton OrderTicketButton 
               Caption         =   "&Order ticket"
               Enabled         =   0   'False
               Height          =   375
               Left            =   2160
               TabIndex        =   3
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton MarketDepthButton 
               Caption         =   "&Mkt depth"
               Enabled         =   0   'False
               Height          =   375
               Left            =   1080
               TabIndex        =   2
               Top             =   0
               Width           =   975
            End
         End
      End
      Begin VB.CommandButton ChartButton 
         Caption         =   "Show &Chart"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72000
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox SessionOnlyCheck 
         Caption         =   "Session only"
         Height          =   375
         Left            =   -72240
         TabIndex        =   13
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox NumHistoryBarsText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72000
         TabIndex        =   12
         Text            =   "500"
         Top             =   840
         Width           =   975
      End
      Begin TradeBuildUI26.TimeframeSelector LiveChartTimeframeSelector 
         Height          =   330
         Left            =   -73080
         TabIndex        =   11
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
      End
      Begin TradeBuildUI26.TimeframeSelector HistTimeframeSelector 
         Height          =   330
         Left            =   -73080
         TabIndex        =   15
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker ToDatePicker 
         Height          =   375
         Left            =   -73080
         TabIndex        =   19
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd HH:mm"
         Format          =   70647811
         CurrentDate     =   39365
      End
      Begin MSComCtl2.DTPicker FromDatePicker 
         Height          =   375
         Left            =   -73080
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd HH:mm"
         Format          =   70647811
         CurrentDate     =   39365
      End
      Begin MSComctlLib.ProgressBar ReplayProgressBar 
         Height          =   135
         Left            =   -74880
         TabIndex        =   42
         Top             =   4440
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label6 
         Caption         =   "Current configuration is:"
         Height          =   375
         Left            =   -74640
         TabIndex        =   45
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   -74880
         TabIndex        =   44
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   975
         Left            =   -74880
         TabIndex        =   43
         Top             =   4680
         Width           =   3855
      End
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74880
         TabIndex        =   41
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Number of history bars"
         Height          =   495
         Left            =   -74880
         TabIndex        =   39
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Timeframe"
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Number of history bars"
         Height          =   375
         Left            =   -74880
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Timeframe"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   9585
      Width           =   16665
      _ExtentX        =   29395
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   23733
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "timezone"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Key             =   "datetime"
         EndProperty
      EndProperty
   End
   Begin TradeBuildUI26.TickerGrid TickerGrid1 
      Height          =   4815
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8493
      AllowUserReordering=   3
      RowSizingMode   =   1
      Rows            =   100
      RowBackColorOdd =   16316664
      RowBackColorEven=   15658734
      HighLight       =   0
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0
      FillStyle       =   1
      Cols            =   24
   End
End
Attribute VB_Name = "fTradeSkilDemo"
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
' Interfaces
'================================================================================

Implements LogListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================
    
Private Const ModuleName                    As String = "fTradeSkilDemo"

'================================================================================
' Enums
'================================================================================

Private Enum ControlsTabIndexNumbers
    ControlsTabIndexTickers
    ControlsTabIndexLiveCharts
    ControlsTabIndexHistoricalCharts
    ControlsTabIndexTickfileReplay
    ControlsTabIndexConfig
End Enum

Private Enum ExecutionsTabIndexNumbers
    ExecutionsTabIndexLive = 1
    ExecutionsTabIndexSimulated
End Enum

Private Enum FeaturesTabIndexNumbers
    FeaturesTabIndexOrders
    FeaturesTabIndexExecutions
    FeaturesTabIndexLog
End Enum

Private Enum OrdersTabIndexNumbers
    OrdersTabIndexLive = 1
    OrdersTabIndexSimulated
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI               As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1

Private WithEvents mTickers                     As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private WithEvents mTickfileManager             As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1

Private WithEvents mCurrentClock                As Clock
Attribute mCurrentClock.VB_VarHelpID = -1

Private mConfigEditor                           As fConfigEditor
Attribute mConfigEditor.VB_VarHelpID = -1

Private mControlsHidden                         As Boolean
Private mFeaturesHidden                         As Boolean

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
' ensure we get the Windows XP look and feel if running on XP
InitCommonControls

Set mTradeBuildAPI = TradeBuildAPI

End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"

On Error GoTo Err

setupLogging

setCurrentClock getDefaultClock

setupReplaySpeedCombo

FromDatePicker.value = DateAdd("m", -1, Now)
FromDatePicker.value = Empty    ' clear the checkbox
ToDatePicker.value = Now

SendMessage TickfileList.hWnd, LB_SETHORZEXTENT, 2000, 0

LogMessage "Main form loaded successfully", LogLevelDetail

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub Form_QueryUnload( _
                cancel As Integer, _
                UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"

On Error GoTo Err

updateInstanceSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
Static prevHeight As Long
Static prevWidth As Long


On Error GoTo Err

If Me.WindowState = FormWindowStateConstants.vbMinimized Then Exit Sub

If Me.Width < ControlsSSTab.Width + 120 Then Me.Width = ControlsSSTab.Width + 120
If Me.Height < 9555 Then Me.Height = 9555

If Me.Width = prevWidth And Me.Height = prevHeight Then Exit Sub

prevWidth = Me.Width
prevHeight = Me.Height

Resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(cancel As Integer)
Const ProcName As String = "Form_Unload"
Dim lTicker As Ticker
Dim f As Form


On Error GoTo Err

LogMessage "Unloading program"

Set mCurrentClock = Nothing

finishUIControls

For Each f In Forms
    If Not f Is Me Then Unload f
Next

LogMessage "Stopping tickers"
If Not mTickers Is Nothing Then
    For Each lTicker In mTickers
        lTicker.StopTicker
    Next
End If

saveSettings

mTradeBuildAPI.ServiceProviders.RemoveAll

killLogging

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' LogListener Interface Members
'================================================================================

Private Sub LogListener_Finish()
'nothing to do
End Sub

Private Sub LogListener_Notify(ByVal Logrec As TWUtilities30.LogRecord)

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub CancelOrderPlexButton_Click()
Const ProcName As String = "CancelOrderPlexButton_Click"
Dim op As OrderPlex


On Error GoTo Err

If OrdersSummaryTabStrip.SelectedItem.index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    Set op = LiveOrdersSummary.SelectedItem
Else
    Set op = SimulatedOrdersSummary.SelectedItem
End If
If Not op Is Nothing Then op.cancel True

CancelOrderPlexButton.Enabled = False
ModifyOrderPlexButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ChartButton_Click()
Const ProcName As String = "ChartButton_Click"
Dim lTicker As Ticker


On Error GoTo Err

For Each lTicker In TickerGrid1.SelectedTickers
    createChart lTicker
Next

clearSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub Chart1Button_Click()
Const ProcName As String = "Chart1Button_Click"
Dim lTicker As Ticker


On Error GoTo Err

For Each lTicker In TickerGrid1.SelectedTickers
    createChart lTicker
Next

clearSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ClearTickfileListButton_Click()
Const ProcName As String = "ClearTickfileListButton_Click"

On Error GoTo Err

TickfileList.Clear
ClearTickfileListButton.Enabled = False
mTickfileManager.ClearTickfileSpecifiers
PlayTickFileButton.Enabled = False
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = False
StopReplayButton.Enabled = False
ChartButton.Enabled = False
Chart1Button.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ClosePositionsButton_Click()
Const ProcName As String = "ClosePositionsButton_Click"

On Error GoTo Err

If Not mTradeBuildAPI.ClosingPositions Then
    If OrdersSummaryTabStrip.SelectedItem.index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
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

Private Sub ConfigEditorButton_Click()
Const ProcName As String = "ConfigEditorButton_Click"

On Error GoTo Err

showConfigEditor

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ControlsSSTab_Click(PreviousTab As Integer)
Const ProcName As String = "ControlsSSTab_Click"

On Error GoTo Err

Select Case ControlsSSTab.Tab
Case ControlsTabIndexNumbers.ControlsTabIndexConfig
    ConfigEditorButton.SetFocus
Case ControlsTabIndexNumbers.ControlsTabIndexHistoricalCharts
    HistContractSearch.SetFocus
Case ControlsTabIndexNumbers.ControlsTabIndexLiveCharts
    LiveChartTimeframeSelector.SetFocus
    If TickerGrid1.SelectedTickers.Count > 0 Then ChartButton.Default = True
Case ControlsTabIndexNumbers.ControlsTabIndexTickers
    LiveContractSearch.SetFocus
    If TickerGrid1.SelectedTickers.Count > 0 Then Chart1Button.Default = True
Case ControlsTabIndexNumbers.ControlsTabIndexTickfileReplay
    If mTickfileManager Is Nothing Then
        SelectTickfilesButton.Default = True
    Else
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

Private Sub ControlsTabStrip_Click()
Const ProcName As String = "ControlsTabStrip_Click"

On Error GoTo Err

ControlsSSTab.SetFocus
ControlsSSTab.Tab = ControlsTabStrip.SelectedItem.index - 1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub FeaturesSSTAB_Click(PreviousTab As Integer)
Const ProcName As String = "FeaturesSSTAB_Click"

On Error GoTo Err

Select Case FeaturesSSTAB.Tab
Case FeaturesSSTAB.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexLog
Case FeaturesSSTAB.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexOrders
    If ModifyOrderPlexButton.Enabled Then
        ModifyOrderPlexButton.Default = True
    Else
        If CancelOrderPlexButton.Enabled Then CancelOrderPlexButton.Default = True
    End If
Case FeaturesSSTAB.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexExecutions
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub HideControlsPicture_Click()
Const ProcName As String = "HideControlsPicture_Click"

On Error GoTo Err

hideControls

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub HideFeaturesPicture_Click()
Const ProcName As String = "HideFeaturesPicture_Click"

On Error GoTo Err

hideFeatures

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

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

MsgBox "No contracts found", vbExclamation, "Attention"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub LiveContractSearch_Action()
Const ProcName As String = "LiveContractSearch_Action"

On Error GoTo Err

mTickers.StartTickersFromContracts TickerOptOrdersAreLive + TickerOptUseExchangeTimeZone, _
                                    LiveContractSearch.SelectedContracts

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveContractSearch_NoContracts()
Const ProcName As String = "LiveContractSearch_NoContracts"

On Error GoTo Err

MsgBox "No contracts found", vbExclamation, "Attention"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub LiveChartTimeframeSelector_Click()
Const ProcName As String = "LiveChartTimeframeSelector_Click"

On Error GoTo Err

setChartButtonTooltip

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

Private Sub MarketDepthButton_Click()
Const ProcName As String = "MarketDepthButton_Click"
Dim lTicker As Ticker


On Error GoTo Err

For Each lTicker In TickerGrid1.SelectedTickers
    showMarketDepthForm lTicker
Next

clearSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ModifyOrderPlexButton_Click()
Const ProcName As String = "ModifyOrderPlexButton_Click"

On Error GoTo Err

If OrdersSummaryTabStrip.SelectedItem.index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    If LiveOrdersSummary.SelectedItem Is Nothing Then
        ModifyOrderPlexButton.Enabled = False
    ElseIf LiveOrdersSummary.IsSelectedItemModifiable Then
        getOrderTicket.showOrderPlex LiveOrdersSummary.SelectedItem, LiveOrdersSummary.SelectedOrderIndex
    End If
Else
    If SimulatedOrdersSummary.SelectedItem Is Nothing Then
        ModifyOrderPlexButton.Enabled = False
    ElseIf SimulatedOrdersSummary.IsSelectedItemModifiable Then
        getOrderTicket.showOrderPlex SimulatedOrdersSummary.SelectedItem, SimulatedOrdersSummary.SelectedOrderIndex
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub NumHistBarsText_Validate(cancel As Boolean)
Const ProcName As String = "NumHistBarsText_Validate"

On Error GoTo Err

If Not IsInteger(NumHistBarsText.Text, 0, 2000) Then cancel = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub NumHistoryBarsText_Validate(cancel As Boolean)
Const ProcName As String = "NumHistoryBarsText_Validate"

On Error GoTo Err

If Not IsInteger(NumHistoryBarsText.Text, 0, 2000) Then cancel = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub OrdersSummaryTabStrip_Click()
Const ProcName As String = "OrdersSummaryTabStrip_Click"
Static currIndex As Long


On Error GoTo Err

If OrdersSummaryTabStrip.SelectedItem.index = currIndex Then Exit Sub

Select Case OrdersSummaryTabStrip.SelectedItem.index
Case OrdersTabIndexNumbers.OrdersTabIndexLive
    LiveOrdersSummary.Visible = True
    SimulatedOrdersSummary.Visible = False
    setOrdersSelection LiveOrdersSummary
Case OrdersTabIndexNumbers.OrdersTabIndexSimulated
    LiveOrdersSummary.Visible = False
    SimulatedOrdersSummary.Visible = True
    setOrdersSelection SimulatedOrdersSummary
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderTicket1Button_Click()
Const ProcName As String = "OrderTicket1Button_Click"

On Error GoTo Err

If getSelectedTicker Is Nothing Then
    MsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
Else
    getOrderTicket.Ticker = getSelectedTicker
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub OrderTicketButton_Click()
Const ProcName As String = "OrderTicketButton_Click"

On Error GoTo Err

If getSelectedTicker Is Nothing Then
    MsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
Else
    getOrderTicket.Ticker = getSelectedTicker
End If

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
mTickfileManager.PauseReplay

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub PlayTickFileButton_Click()
Const ProcName As String = "PlayTickFileButton_Click"

On Error GoTo Err

PlayTickFileButton.Enabled = False
SelectTickfilesButton.Enabled = False
ClearTickfileListButton.Enabled = False
PauseReplayButton.Enabled = True
SkipReplayButton.Enabled = True
StopReplayButton.Enabled = True
ReplayProgressBar.Visible = True

If mTickfileManager.Ticker Is Nothing Then
    mTickfileManager.ReplayProgressEventIntervalMillisecs = 250
    LogMessage "Tickfile replay started"
Else
    LogMessage "Tickfile replay resumed"
End If
mTickfileManager.ReplaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)

mTickfileManager.StartReplay

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ReplaySpeedCombo_Click()
Const ProcName As String = "ReplaySpeedCombo_Click"

On Error GoTo Err

If Not mTickfileManager Is Nothing Then
    mTickfileManager.ReplaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub SelectTickfilesButton_Click()
Const ProcName As String = "SelectTickfilesButton_Click"


Dim tickfiles As TickFileSpecifiers
Dim tfs As TickfileSpecifier
Dim userCancelled As Boolean


On Error GoTo Err

Set tickfiles = SelectTickfiles(userCancelled)
If userCancelled Then Exit Sub

Set mTickfileManager = mTickers.CreateTickFileManager(TickerOptions.TickerOptUseExchangeTimeZone)

mTickfileManager.TickFileSpecifiers = tickfiles

TickfileList.Clear
For Each tfs In tickfiles
    TickfileList.AddItem tfs.filename
Next
checkOkToStartReplay
ClearTickfileListButton.Enabled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ShowControlsPicture_Click()
Const ProcName As String = "ShowControlsPicture_Click"

On Error GoTo Err

showControls

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub showFeaturesPicture_Click()
Const ProcName As String = "showFeaturesPicture_Click"

On Error GoTo Err

showFeatures

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub SimulatedOrdersSummary_Click()
Const ProcName As String = "SimulatedOrdersSummary_Click"

On Error GoTo Err

setOrdersSelection SimulatedOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub SkipReplayButton_Click()
Const ProcName As String = "SkipReplayButton_Click"

On Error GoTo Err

LogMessage "Tickfile skipped"
mTickfileManager.SkipTickfile

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub StopReplayButton_Click()
Const ProcName As String = "StopReplayButton_Click"

On Error GoTo Err

PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = True
StopReplayButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
ChartButton.Enabled = False
Chart1Button.Enabled = False
mTickfileManager.Ticker.StopTicker

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub StopTickerButton_Click()
Const ProcName As String = "StopTickerButton_Click"

On Error GoTo Err

TickerGrid1.stopSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub TickerGrid1_SelectionChanged()
Const ProcName As String = "TickerGrid1_SelectionChanged"

On Error GoTo Err

handleSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub TickerGrid1_TickerStarted(ByVal row As Long)
Const ProcName As String = "TickerGrid1_TickerStarted"

On Error GoTo Err

TickerGrid1.deselectSelectedTickers
TickerGrid1.selectTicker row
handleSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' mCurrentClock Event Handlers
'================================================================================

Private Sub mCurrentClock_Tick()
Const ProcName As String = "mCurrentClock_Tick"

On Error GoTo Err

displayTime

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_StateChange( _
                ev As StateChangeEventData)
Const ProcName As String = "mTickers_StateChange"
Dim lTicker As Ticker


On Error GoTo Err

OrderTicketButton.Enabled = Not (getSelectedTicker Is Nothing)
OrderTicket1Button.Enabled = OrderTicketButton.Enabled

Set lTicker = ev.Source

Select Case ev.state
Case TickerStateCreated
Case TickerStateStarting

Case TickerStateReady
    If lTicker Is getSelectedTicker Then setCurrentClock lTicker.Clock
Case TickerStateRunning
    If lTicker Is getSelectedTicker Then
        MarketDepthButton.Enabled = True
        ChartButton.Enabled = True
        Chart1Button.Enabled = True
    End If
    
Case TickerStatePaused

Case TickerStateClosing

Case TickerStateStopped
    If getSelectedTicker Is Nothing Then
        StopTickerButton.Enabled = False
        MarketDepthButton.Enabled = False
        ChartButton.Enabled = False
        Chart1Button.Enabled = False
    Else
        setCurrentClock getSelectedTicker.Clock
    End If
    
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' mTickfileManager Event Handlers
'================================================================================

Private Sub mTickfileManager_QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As Contract, _
                continueMode As ReplayContinueModes)
Const ProcName As String = "mTickfileManager_QueryReplayNextTickfile"

On Error GoTo Err

If tickfileIndex <> 0 Then setCurrentClock getDefaultClock

ReplayProgressBar.Min = 0
ReplayProgressBar.Max = 100
ReplayProgressBar.value = 0
TickfileList.ListIndex = tickfileIndex - 1
ReplayContractLabel.caption = Replace(pContract.Specifier.ToString, vbCrLf, "; ")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub mTickfileManager_ReplayCompleted()
Const ProcName As String = "mTickfileManager_ReplayCompleted"

On Error GoTo Err

MarketDepthButton.Enabled = False
PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = False
StopReplayButton.Enabled = False

SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
ReplayProgressBar.value = 0
ReplayProgressBar.Visible = False
ReplayContractLabel.caption = ""
ReplayProgressLabel.caption = ""

LogMessage "Tickfile replay completed"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub mTickfileManager_ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)
Const ProcName As String = "mTickfileManager_ReplayProgress"

On Error GoTo Err

ReplayProgressBar.value = percentComplete
ReplayProgressLabel.caption = tickfileTimestamp & _
                                "  Processed " & _
                                eventsPlayed & _
                                " events"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub mTickfileManager_TickerAllocated(ByVal pTicker As Ticker)
Const ProcName As String = "mTickfileManager_TickerAllocated"

On Error GoTo Err

pTicker.DOMEventsRequired = DOMProcessedEvents

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

' fires when an unrecoverable error has occurred in TradeBuild.
Private Sub mTradeBuildAPI_Error( _
                ByRef ev As ErrorEventData)
Const ProcName As String = "mTradeBuildAPI_Error"
On Error Resume Next    ' ignore any further errors

' TradeBuild will have already logged the error so no need for us to do it
gHandleFatalError

End Sub

Private Sub mTradeBuildAPI_Notification( _
                ByRef ev As NotificationEventData)
Const ProcName As String = "mTradeBuildAPI_Notification"
Dim spError As ServiceProviderError


On Error GoTo Err

Select Case ev.eventCode
Case ApiNotifyCodes.ApiNotifyInvalidRequest
    LogMessage "Request failed: " & _
                ev.eventMessage & vbCrLf
    gModelessMsgBox "Request failed: " & _
                ev.eventMessage & vbCrLf, _
                MsgBoxExclamation, _
                "Attention"
Case ApiNotifyCodes.ApiNotifyOrderRejected
    LogMessage "Order rejected: " & _
                ev.eventMessage & vbCrLf
    gModelessMsgBox "Order rejected: " & _
                ev.eventMessage & vbCrLf, _
                MsgBoxExclamation, _
                "Attention"
Case ApiNotifyCodes.ApiNotifyOrderDeferred
    LogMessage "Order deferred: " & _
                ev.eventMessage & vbCrLf
    gModelessMsgBox "Order deferred: " & _
                ev.eventMessage & vbCrLf, _
                MsgBoxExclamation, _
                "Attention"
Case ApiNotifyCodes.ApiNotifyServiceProviderError
    Set spError = mTradeBuildAPI.GetServiceProviderError
    LogMessage "Error from " & _
                        spError.ServiceProviderName & _
                        ": code " & spError.errorCode & _
                        ": " & spError.Message

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Sub initialise( _
                ByVal editConfig As Boolean)
Const ProcName As String = "initialise"

On Error GoTo Err

Set mConfigEditor = New fConfigEditor

If editConfig Then
    ' show the configuration editor and don't attempt any other configuration
    showConfigEditor
    Exit Sub
End If

loadAppInstanceConfig

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function LoadConfig( _
                ByVal configToLoad As TWUtilities30.ConfigurationSection) As Boolean
Const ProcName As String = "LoadConfig"



On Error GoTo Err

updateInstanceSettings
saveSettings

finishUIControls

closeChartsAndMarketDepthForms

CurrentConfigNameText.Text = ""
Me.caption = gAppTitle

mTradeBuildAPI.EndSession

gSetPermittedServiceProviderRoles

Set gAppInstanceConfig = configToLoad

If ConfigureTradeBuild(gConfigStore, gAppInstanceConfig.InstanceQualifier) Then
    Unload mConfigEditor
    loadAppInstanceConfig
    Set mConfigEditor = New fConfigEditor
    LoadConfig = True
Else
    MsgBox "The configuration cannot be loaded", _
            vbExclamation, _
            "Attention!"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub MakeVisible()
Const ProcName As String = "MakeVisible"

On Error GoTo Err

Me.Show
If Not mControlsHidden Then ControlsTabStrip.Tabs(1).Selected = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyInstanceSettings()
Const ProcName As String = "applyInstanceSettings"

On Error GoTo Err

LogMessage "Loading configuration: positioning main form"
Select Case gAppInstanceConfig.GetSetting(ConfigSettingMainFormWindowstate, WindowStateNormal)
Case WindowStateMaximized
    Me.WindowState = FormWindowStateConstants.vbMaximized
Case WindowStateMinimized
    Me.WindowState = FormWindowStateConstants.vbMinimized
Case WindowStateNormal
    Me.left = CLng(gAppInstanceConfig.GetSetting(ConfigSettingMainFormLeft, 0)) * Screen.TwipsPerPixelX
    Me.Top = CLng(gAppInstanceConfig.GetSetting(ConfigSettingMainFormTop, 0)) * Screen.TwipsPerPixelY
    Me.Width = CLng(gAppInstanceConfig.GetSetting(ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
    Me.Height = CLng(gAppInstanceConfig.GetSetting(ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
End Select

mControlsHidden = CBool(gAppInstanceConfig.GetSetting(ConfigSettingMainFormControlsHidden, CStr(False)))
If mControlsHidden Then
    hideControls
Else
    showControls
End If

mFeaturesHidden = CBool(gAppInstanceConfig.GetSetting(ConfigSettingMainFormFeaturesHidden, CStr(False)))
If mFeaturesHidden Then
    hideFeatures
Else
    showFeatures
End If

LogMessage "Loading configuration: starting tickers"
TickerGrid1.LoadFromConfig gAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionTickerGrid)

LogMessage "Loading configuration: loading default study configurations"
LoadDefaultStudyConfigurationsFromConfig gAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionDefaultStudyConfigs)

LogMessage "Loading configuration: starting charts"
Dim chartConfig As ConfigurationSection
For Each chartConfig In gAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionCharts)
    createChartFromConfig chartConfig
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub checkOkToStartReplay()
Const ProcName As String = "checkOkToStartReplay"

On Error GoTo Err

If TickfileList.ListCount <> 0 Then
    PlayTickFileButton.Enabled = True
Else
    PlayTickFileButton.Enabled = False
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub clearSelectedTickers()
Const ProcName As String = "clearSelectedTickers"

On Error GoTo Err

TickerGrid1.deselectSelectedTickers
handleSelectedTickers

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub closeChartsAndMarketDepthForms()
Const ProcName As String = "closeChartsAndMarketDepthForms"
Dim f As Form

On Error GoTo Err

For Each f In Forms
    If TypeOf f Is fChart Or TypeOf f Is fMarketDepth Then Unload f
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub createChart(ByVal pTicker As Ticker)
Const ProcName As String = "createChart"
Dim chartForm As fChart
Dim tp As timePeriod


On Error GoTo Err

If Not pTicker.state = TickerStateRunning Then Exit Sub

Set tp = LiveChartTimeframeSelector.TimeframeDesignator
Set chartForm = New fChart
chartForm.showChart pTicker, _
                    createChartSpec(tp, CLng(NumHistoryBarsText.Text), SessionOnlyCheck = vbChecked)
chartForm.Show vbModeless

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errDesc As String: errDesc = Err.Description
Dim errSource As String: errSource = Err.Source

Unload chartForm

On Error GoTo ErrErr
Err.Raise errNumber, errSource, errDesc
ErrErr:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub createChartFromConfig( _
                ByVal chartConfig As ConfigurationSection)
Const ProcName As String = "createChartFromConfig"
Dim chartForm As fChart

On Error GoTo Err

Set chartForm = New fChart
If chartForm.LoadFromConfig(chartConfig) Then
    chartForm.Show vbModeless
Else
    Unload chartForm
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errDesc As String: errDesc = Err.Description
Dim errSource As String: errSource = Err.Source

Unload chartForm

On Error GoTo ErrErr
Err.Raise errNumber, errSource, errDesc
ErrErr:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function createChartSpec( _
                ByVal timePeriod As timePeriod, _
                ByVal initialNumberOfBars As Long, _
                ByVal sessionOnly As Boolean) As ChartSpecifier
Const ProcName As String = "createChartSpec"
Static defaultRegionStyle As ChartRegionStyle
Static volumeRegionStyle As ChartRegionStyle
Static xAxisRegionStyle As ChartRegionStyle
Static defaultYAxisRegionStyle As ChartRegionStyle
Static defaultBarsStyle As BarStyle
Static defaultVolumeStyle As DataPointStyle

On Error GoTo Err

ReDim GradientFillColors(1) As Long

If defaultRegionStyle Is Nothing Then
    Set defaultRegionStyle = New ChartRegionStyle
    defaultRegionStyle.Autoscaling = True
    GradientFillColors(0) = RGB(192, 192, 192)
    GradientFillColors(1) = RGB(248, 248, 248)
    defaultRegionStyle.BackGradientFillColors = GradientFillColors
    'defaultRegionStyle.GridLineStyle.Color = &HC0C0C0
    defaultRegionStyle.GridlineSpacingY = 1.8
    defaultRegionStyle.HasGrid = True
    defaultRegionStyle.CursorSnapsToTickBoundaries = True
End If

If volumeRegionStyle Is Nothing Then
    Set volumeRegionStyle = defaultRegionStyle.Clone
    volumeRegionStyle.GridlineSpacingY = 0.8
    volumeRegionStyle.MinimumHeight = 10
    volumeRegionStyle.IntegerYScale = True
End If

If xAxisRegionStyle Is Nothing Then
    Set xAxisRegionStyle = defaultRegionStyle.Clone
    xAxisRegionStyle.HasGrid = False
    xAxisRegionStyle.HasGridText = True
    GradientFillColors(0) = RGB(230, 236, 207)
    GradientFillColors(1) = RGB(222, 236, 215)
    xAxisRegionStyle.BackGradientFillColors = GradientFillColors
End If

If defaultYAxisRegionStyle Is Nothing Then
    Set defaultYAxisRegionStyle = defaultRegionStyle.Clone
    GradientFillColors(0) = RGB(234, 246, 254)
    GradientFillColors(1) = RGB(226, 246, 255)
    defaultYAxisRegionStyle.BackGradientFillColors = GradientFillColors
    defaultYAxisRegionStyle.HasGrid = False
End If

If defaultBarsStyle Is Nothing Then
    Set defaultBarsStyle = New BarStyle
    defaultBarsStyle.Thickness = 2
    defaultBarsStyle.Width = 0.6
    defaultBarsStyle.DisplayMode = BarDisplayModeCandlestick
    defaultBarsStyle.DownColor = &H43FC2
    defaultBarsStyle.IncludeInAutoscale = True
    defaultBarsStyle.OutlineThickness = 1
    defaultBarsStyle.SolidUpBody = False
    defaultBarsStyle.TailThickness = 1
    defaultBarsStyle.UpColor = &H1D9311
End If

If defaultVolumeStyle Is Nothing Then
    Set defaultVolumeStyle = New DataPointStyle
    defaultVolumeStyle.DisplayMode = DataPointDisplayModeHistogram
    defaultVolumeStyle.DownColor = &H43FC2
    defaultVolumeStyle.HistogramBarWidth = 0.6
    defaultVolumeStyle.IncludeInAutoscale = True
    defaultVolumeStyle.LineStyle = LineSolid
    defaultVolumeStyle.LineThickness = 1
    defaultVolumeStyle.PointStyle = PointRound
    defaultVolumeStyle.UpColor = &H1D9311
End If

Set createChartSpec = CreateChartSpecifier(timePeriod, _
                       initialNumberOfBars, _
                       Not sessionOnly, _
                       20, _
                       defaultRegionStyle, _
                       volumeRegionStyle, _
                       xAxisRegionStyle, _
                       defaultYAxisRegionStyle, _
                       defaultBarsStyle, _
                       defaultVolumeStyle)

createChartSpec.TwipsPerBar = 100

createChartSpec.ChartBackColor = &H7F7FFF

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub createHistoricCharts( _
                ByVal pContracts As Contracts)
Const ProcName As String = "createHistoricCharts"
Dim lTicker As Ticker
Dim fromDate As Date
Dim toDate As Date
Dim chartForm As fChart
Dim lContract As Contract


On Error GoTo Err

For Each lContract In pContracts
    Set lTicker = mTickers.Add(TickerOptions.TickerOptUseExchangeTimeZone)
    lTicker.LoadTicker lContract.Specifier
    
    If IsNull(FromDatePicker.value) Then
        fromDate = CDate(0)
    Else
        fromDate = DateSerial(FromDatePicker.Year, FromDatePicker.Month, FromDatePicker.Day) + _
                    TimeSerial(FromDatePicker.Hour, FromDatePicker.Minute, 0)
    End If
    
    If IsNull(ToDatePicker.value) Then
        toDate = Now
    Else
        toDate = DateSerial(ToDatePicker.Year, ToDatePicker.Month, ToDatePicker.Day) + _
                    TimeSerial(ToDatePicker.Hour, ToDatePicker.Minute, 0)
    End If
    
    Set chartForm = New fChart
    chartForm.showHistoricalChart lTicker, _
                        createChartSpec(HistTimeframeSelector.TimeframeDesignator, CLng(NumHistBarsText.Text), HistSessionOnlyCheck = vbChecked), _
                        fromDate, _
                        toDate
    chartForm.Show vbModeless
    chartForm.Visible = True

Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub displayTime()
Const ProcName As String = "displayTime"
Dim theTime As Date

On Error GoTo Err

theTime = mCurrentClock.TimeStamp
StatusBar1.Panels("datetime") = FormatDateTime(theTime, vbShortDate) & _
                Format(theTime, " hh:mm:ss")

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub finishUIControls()
Const ProcName As String = "finishUIControls"

On Error GoTo Err

LiveOrdersSummary.Finish
SimulatedOrdersSummary.Finish
LiveExecutionsSummary.Finish
SimulatedExecutionsSummary.Finish
TickerGrid1.Finish

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
Static formatter As LogFormatter

On Error GoTo Err

If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getDefaultClock() As Clock
Const ProcName As String = "getDefaultClock"
Static lClock As Clock

On Error GoTo Err

If lClock Is Nothing Then Set lClock = GetClock("") ' create a clock running local time
Set getDefaultClock = lClock

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getOrderTicket() As fOrderTicket
Const ProcName As String = "getOrderTicket"
Static lOrderTicket As fOrderTicket

On Error GoTo Err

If lOrderTicket Is Nothing Then
    Set lOrderTicket = New fOrderTicket
End If
lOrderTicket.Show vbModeless
Set getOrderTicket = lOrderTicket

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getSelectedTicker() As Ticker
Const ProcName As String = "getSelectedTicker"

On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 1 Then Set getSelectedTicker = TickerGrid1.SelectedTickers.Item(1)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub handleSelectedTickers()
Const ProcName As String = "handleSelectedTickers"
Dim lTicker As Ticker


On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 0 Then
    StopTickerButton.Enabled = False
    ChartButton.Enabled = False
    Chart1Button.Enabled = False
    MarketDepthButton.Enabled = False
    OrderTicketButton.Enabled = False
    OrderTicket1Button.Enabled = False
    setCurrentClock getDefaultClock
Else
    StopTickerButton.Enabled = True
    ChartButton.Enabled = True
    Chart1Button.Enabled = True
    MarketDepthButton.Enabled = True
    
    If ControlsSSTab.Tab = ControlsTabIndexNumbers.ControlsTabIndexLiveCharts Then
        ChartButton.Default = True
    ElseIf ControlsSSTab.Tab = ControlsTabIndexNumbers.ControlsTabIndexTickers Then
        Chart1Button.Default = True
    End If
    
    Set lTicker = getSelectedTicker
    If Not lTicker Is Nothing Then
        If lTicker.state = TickerStateRunning Then
            setCurrentClock lTicker.Clock
            OrderTicketButton.Enabled = True
            OrderTicket1Button.Enabled = True
        Else
            setCurrentClock getDefaultClock
            
            OrderTicketButton.Enabled = False
            OrderTicket1Button.Enabled = False
            ChartButton.Enabled = False
            Chart1Button.Enabled = False
            MarketDepthButton.Enabled = False
        End If
    Else
        setCurrentClock getDefaultClock
        OrderTicketButton.Enabled = False
        OrderTicket1Button.Enabled = False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub hideControls()
Const ProcName As String = "hideControls"

On Error GoTo Err

ControlsSSTab.Visible = False
ControlsTabStrip.Visible = False
ShowControlsPicture.Visible = True
HideControlsPicture.Visible = False
mControlsHidden = True
Resize
Me.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub hideFeatures()
Const ProcName As String = "hideFeatures"

On Error GoTo Err

FeaturesSSTAB.Visible = False
ShowFeaturesPicture.Visible = True
HideFeaturesPicture.Visible = False
mFeaturesHidden = True
Resize
Me.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub killLogging()
Const ProcName As String = "killLogging"

On Error GoTo Err

GetLogger("log").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub loadAppInstanceConfig()
Const ProcName As String = "loadAppInstanceConfig"



On Error GoTo Err

LogMessage "Loading configuration: " & gAppInstanceConfig.InstanceQualifier

Set mTickers = mTradeBuildAPI.Tickers

LogMessage "Loading configuration: Setting up ticker grid"
setupTickerGrid

LogMessage "Loading configuration: Setting up order summaries"
setupOrderSummaries

LogMessage "Loading configuration: Setting up execution summaries"
setupExecutionSummaries

LogMessage "Loading configuration: Setting up timeframeselectors"
setupTimeframeSelectors

LogMessage "Recovering orders from last session"
mTradeBuildAPI.RecoverOrders gAppInstanceConfig.InstanceQualifier

applyInstanceSettings

FeaturesSSTAB.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexOrders

LogMessage "Loaded configuration: " & gAppInstanceConfig.InstanceQualifier
CurrentConfigNameText = gAppInstanceConfig.InstanceQualifier
Me.caption = gAppTitle & _
            " - " & gAppInstanceConfig.InstanceQualifier

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Resize()
Const ProcName As String = "Resize"
Dim left As Long


On Error GoTo Err

If mControlsHidden Then
    left = ShowControlsPicture.Width + 60
Else
    left = 120 + ControlsTabStrip.Width + 120
End If

StatusBar1.Top = Me.ScaleHeight - StatusBar1.Height

ControlsSSTab.Height = StatusBar1.Top - ControlsSSTab.Top

FeaturesSSTAB.Move left, _
                    StatusBar1.Top - FeaturesSSTAB.Height, _
                    Me.ScaleWidth - left - 120

HideFeaturesPicture.Move FeaturesSSTAB.left + FeaturesSSTAB.Width - 255, _
                        FeaturesSSTAB.Top
ShowFeaturesPicture.Move HideFeaturesPicture.left, _
                        StatusBar1.Top - 240

TickerGrid1.Move left, _
                TickerGrid1.Top, _
                Me.ScaleWidth - left - 120, _
                IIf(mFeaturesHidden, ShowFeaturesPicture.Top - 60, FeaturesSSTAB.Top - 120) - TickerGrid1.Top

If OrderTicket1Button.left >= 0 Then
    OrderTicket1Button.left = FeaturesSSTAB.Width - OrderTicket1Button.Width - 120
    ModifyOrderPlexButton.left = FeaturesSSTAB.Width - ModifyOrderPlexButton.Width - 120
    CancelOrderPlexButton.left = FeaturesSSTAB.Width - CancelOrderPlexButton.Width - 120
    ClosePositionsButton.left = FeaturesSSTAB.Width - CancelOrderPlexButton.Width - 120
    
    LiveOrdersSummary.Width = ModifyOrderPlexButton.left - 120 - 120
    SimulatedOrdersSummary.Width = ModifyOrderPlexButton.left - 120 - 120
Else
    OrderTicket1Button.left = FeaturesSSTAB.Width - OrderTicket1Button.Width - 120 - SSTabInactiveControlAdjustment
    ModifyOrderPlexButton.left = FeaturesSSTAB.Width - ModifyOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    CancelOrderPlexButton.left = FeaturesSSTAB.Width - CancelOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    ClosePositionsButton.left = FeaturesSSTAB.Width - CancelOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    
    LiveOrdersSummary.Width = ModifyOrderPlexButton.left + SSTabInactiveControlAdjustment - 120 - 120
    SimulatedOrdersSummary.Width = ModifyOrderPlexButton.left + SSTabInactiveControlAdjustment - 120 - 120
End If

LogText.Width = FeaturesSSTAB.Width - 120 - 120
LiveExecutionsSummary.Width = FeaturesSSTAB.Width - 120 - 120
SimulatedExecutionsSummary.Width = FeaturesSSTAB.Width - 120 - 120

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub saveSettings()
Const ProcName As String = "saveSettings"

On Error GoTo Err

If gConfigStore.dirty Then
    LogMessage "Saving configuration"
    gConfigStore.Save
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setChartButtonTooltip()
Const ProcName As String = "setChartButtonTooltip"
Dim tp As timePeriod


On Error GoTo Err

Set tp = LiveChartTimeframeSelector.TimeframeDesignator

ChartButton.ToolTipText = "Show " & _
                        tp.ToString & _
                        " chart"
Chart1Button.ToolTipText = ChartButton.ToolTipText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setCurrentClock( _
                ByVal pClock As Clock)
Const ProcName As String = "setCurrentClock"

On Error GoTo Err

Set mCurrentClock = pClock
StatusBar1.Panels("timezone") = mCurrentClock.TimeZone.StandardName
displayTime

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setOrdersSelection( _
                ByVal pOrdersSummary As OrdersSummary)
Const ProcName As String = "setOrdersSelection"
Dim selection As OrderPlex


On Error GoTo Err

If pOrdersSummary.IsEditing Then
    pOrdersSummary.Default = True
    Exit Sub
End If

pOrdersSummary.Default = False

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupExecutionSummaries()
Const ProcName As String = "setupExecutionSummaries"

On Error GoTo Err

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedExecutionsSummary.Simulated = True
    SimulatedExecutionsSummary.MonitorWorkspace mTradeBuildAPI.DefaultWorkSpace
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top + ExecutionsSummaryTabStrip.Height - SimulatedExecutionsSummary.Top
    SimulatedExecutionsSummary.Visible = True
    
    LiveExecutionsSummary.Visible = False
    
    ExecutionsSummaryTabStrip.Visible = False
Else
    SimulatedExecutionsSummary.Simulated = True
    SimulatedExecutionsSummary.MonitorWorkspace mTradeBuildAPI.DefaultWorkSpace
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    
    LiveExecutionsSummary.Simulated = False
    LiveExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    LiveExecutionsSummary.MonitorWorkspace mTradeBuildAPI.DefaultWorkSpace
    
    ExecutionsSummaryTabStrip.Visible = True
    ExecutionsSummaryTabStrip.Tabs.Item(ExecutionsTabIndexLive).Selected = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupLogging()
Const ProcName As String = "setupLogging"

On Error GoTo Err

GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setupOrderSummaries()
Const ProcName As String = "setupOrderSummaries"

On Error GoTo Err

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedOrdersSummary.Finish
    SimulatedOrdersSummary.Simulated = True
    SimulatedOrdersSummary.MonitorWorkspace mTradeBuildAPI.DefaultWorkSpace
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top + OrdersSummaryTabStrip.Height - SimulatedOrdersSummary.Top
    SimulatedOrdersSummary.Visible = True
    
    LiveOrdersSummary.Visible = False
    
    OrdersSummaryTabStrip.Visible = False
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexSimulated).Selected = True
Else
    SimulatedOrdersSummary.Finish
    SimulatedOrdersSummary.Simulated = True
    SimulatedOrdersSummary.MonitorWorkspace mTradeBuildAPI.DefaultWorkSpace
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top - SimulatedOrdersSummary.Top
    
    LiveOrdersSummary.Finish
    LiveOrdersSummary.Simulated = False
    LiveOrdersSummary.Height = OrdersSummaryTabStrip.Top - SimulatedOrdersSummary.Top
    LiveOrdersSummary.MonitorWorkspace mTradeBuildAPI.DefaultWorkSpace
    
    OrdersSummaryTabStrip.Visible = True
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexLive).Selected = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupReplaySpeedCombo()
Const ProcName As String = "setupReplaySpeedCombo"

On Error GoTo Err

ReplaySpeedCombo.AddItem "10 Second intervals"
ReplaySpeedCombo.ItemData(0) = -10000
ReplaySpeedCombo.AddItem "5 Second intervals"
ReplaySpeedCombo.ItemData(1) = -5000
ReplaySpeedCombo.AddItem "1 Second intervals"
ReplaySpeedCombo.ItemData(2) = -1000
ReplaySpeedCombo.AddItem "0.5 second intervals"
ReplaySpeedCombo.ItemData(3) = -500
ReplaySpeedCombo.AddItem "Continuous"
ReplaySpeedCombo.ItemData(4) = 0
ReplaySpeedCombo.AddItem "Actual speed"
ReplaySpeedCombo.ItemData(5) = 1
ReplaySpeedCombo.AddItem "2x Actual speed"
ReplaySpeedCombo.ItemData(6) = 2
ReplaySpeedCombo.AddItem "4x Actual speed"
ReplaySpeedCombo.ItemData(7) = 4
ReplaySpeedCombo.AddItem "8x Actual speed"
ReplaySpeedCombo.ItemData(8) = 8

ReplaySpeedCombo.Text = "Actual speed"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupTickerGrid()
Const ProcName As String = "setupTickerGrid"

On Error GoTo Err

TickerGrid1.Finish
TickerGrid1.MonitorWorkspace mTradeBuildAPI.DefaultWorkSpace

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupTimeframeSelectors()
Const ProcName As String = "setupTimeframeSelectors"
' now set up the timeframe selectors, which depends on what timeframes the historical data service
' provider supports (it obtains this info from TradeBuild)

On Error GoTo Err

LiveChartTimeframeSelector.initialise   ' use the default settings built-in to the control
LiveChartTimeframeSelector.selectTimeframe GetTimePeriod(5, TimePeriodMinute)
HistTimeframeSelector.initialise
HistTimeframeSelector.selectTimeframe GetTimePeriod(5, TimePeriodMinute)

setChartButtonTooltip

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub showConfigEditor()
Const ProcName As String = "showConfigEditor"

On Error GoTo Err

mConfigEditor.Show vbModeless

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub showControls()
Const ProcName As String = "showControls"

On Error GoTo Err

mControlsHidden = False
Resize
Me.Refresh
ControlsTabStrip.Visible = True
ControlsSSTab.Visible = True
ShowControlsPicture.Visible = False
HideControlsPicture.Visible = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub showFeatures()
Const ProcName As String = "showFeatures"

On Error GoTo Err

mFeaturesHidden = False
Resize
Me.Refresh
FeaturesSSTAB.Visible = True
ShowFeaturesPicture.Visible = False
HideFeaturesPicture.Visible = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub showMarketDepthForm(ByVal pTicker As Ticker)
Const ProcName As String = "showMarketDepthForm"
Dim mktDepthForm As fMarketDepth


On Error GoTo Err

If Not pTicker.state = TickerStateRunning Then Exit Sub

Set mktDepthForm = New fMarketDepth
mktDepthForm.numberOfRows = 100
mktDepthForm.Ticker = pTicker
mktDepthForm.Show vbModeless

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub updateInstanceSettings()
Const ProcName As String = "updateInstanceSettings"

On Error GoTo Err

If Not gAppInstanceConfig Is Nothing Then
    gAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionMainForm
    Select Case Me.WindowState
    Case FormWindowStateConstants.vbMaximized
        gAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMaximized
    Case FormWindowStateConstants.vbMinimized
        gAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMinimized
    Case FormWindowStateConstants.vbNormal
        gAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateNormal
        gAppInstanceConfig.SetSetting ConfigSettingMainFormLeft, Me.left / Screen.TwipsPerPixelX
        gAppInstanceConfig.SetSetting ConfigSettingMainFormTop, Me.Top / Screen.TwipsPerPixelY
        gAppInstanceConfig.SetSetting ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX
        gAppInstanceConfig.SetSetting ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY
    End Select
    
    gAppInstanceConfig.SetSetting ConfigSettingMainFormControlsHidden, CStr(mControlsHidden)
    gAppInstanceConfig.SetSetting ConfigSettingMainFormFeaturesHidden, CStr(mFeaturesHidden)
    
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


