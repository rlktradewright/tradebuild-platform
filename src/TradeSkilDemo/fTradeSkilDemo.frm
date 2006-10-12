VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1942F60B-EC12-45BB-8FBB-89AA1BB15BA4}#16.0#0"; "TradeBuildUI.ocx"
Begin VB.Form fTradeSkilDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TradeSkil Demo Edition"
   ClientHeight    =   6855
   ClientLeft      =   210
   ClientTop       =   330
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   14385
   Begin VB.TextBox DateTimeText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   11880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox CloseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox LowText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox HighText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox VolumeText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox AskSizeText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox LastSizeText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox AskText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox LastText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox BidText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox BidSizeText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox NameText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   360
      Width           =   2280
   End
   Begin VB.CommandButton ChartButton 
      Caption         =   "C&hart"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   44
      ToolTipText     =   "Display a chart"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton MarketDepthButton 
      Caption         =   "&Market depth"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   43
      ToolTipText     =   "Display the market depth"
      Top             =   0
      Width           =   975
   End
   Begin TabDlg.SSTab MainSSTAB 
      Height          =   4335
      Left            =   120
      TabIndex        =   55
      Top             =   960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&1. Configuration"
      TabPicture(0)   =   "fTradeSkilDemo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ConfigureButton"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&2. Tickers"
      TabPicture(1)   =   "fTradeSkilDemo.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3. Orders"
      TabPicture(2)   =   "fTradeSkilDemo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "OrdersSummary1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "OrderButton"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CancelOrderButton"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ModifyOrderButton"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&4. Executions"
      TabPicture(3)   =   "fTradeSkilDemo.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ExecutionsSummary1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5. Replay tickfiles"
      TabPicture(4)   =   "fTradeSkilDemo.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ReplaySpeedCombo"
      Tab(4).Control(1)=   "TickfileList"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "StopReplayButton"
      Tab(4).Control(3)=   "PauseReplayButton"
      Tab(4).Control(4)=   "ClearTickfileListButton"
      Tab(4).Control(5)=   "SelectTickfilesButton"
      Tab(4).Control(6)=   "PlayTickFileButton"
      Tab(4).Control(7)=   "SkipReplayButton"
      Tab(4).Control(8)=   "ReplayProgressBar"
      Tab(4).Control(9)=   "ReplayContractLabel"
      Tab(4).Control(10)=   "ReplayProgressLabel"
      Tab(4).Control(11)=   "Label20"
      Tab(4).Control(12)=   "Label19"
      Tab(4).ControlCount=   13
      Begin TradeBuildUI.ExecutionsSummary ExecutionsSummary1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   106
         Top             =   360
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   6800
      End
      Begin TradeBuildUI.OrdersSummary OrdersSummary1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   105
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6800
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "fTradeSkilDemo.frx":008C
         Left            =   -73800
         List            =   "fTradeSkilDemo.frx":008E
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Service Providers"
         Height          =   3495
         Left            =   -74640
         TabIndex        =   91
         Top             =   480
         Width           =   10335
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   120
            ScaleHeight     =   3135
            ScaleWidth      =   10095
            TabIndex        =   92
            Top             =   240
            Width           =   10095
            Begin VB.CheckBox UseQTTickfileSPCheck 
               Caption         =   "Enable tick data replay from QuoteTracker"
               Height          =   195
               Left            =   2280
               TabIndex        =   24
               Top             =   2640
               Width           =   6015
            End
            Begin VB.CheckBox UseTickDBSPCheck 
               Caption         =   "Enable tick data storage/replay from database"
               Height          =   195
               Left            =   2280
               TabIndex        =   23
               Top             =   2400
               Width           =   6015
            End
            Begin VB.CheckBox UseTickfileSPCheck 
               Caption         =   "Enable tick data storage/replay from files"
               Height          =   195
               Left            =   2280
               TabIndex        =   22
               Top             =   2160
               Value           =   1  'Checked
               Width           =   6015
            End
            Begin VB.TextBox SecContractDataPortText 
               Height          =   285
               Left            =   5400
               TabIndex        =   11
               Text            =   "7496"
               Top             =   960
               Width           =   1335
            End
            Begin VB.ComboBox SecContractDataCombo 
               Height          =   315
               ItemData        =   "fTradeSkilDemo.frx":0090
               Left            =   2280
               List            =   "fTradeSkilDemo.frx":009D
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   960
               Width           =   1725
            End
            Begin VB.TextBox SecContractDataServerText 
               Height          =   285
               Left            =   4080
               TabIndex        =   10
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox SecContractDataClientIdText 
               Height          =   285
               Left            =   6720
               TabIndex        =   12
               Text            =   "123"
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox ContractDataPortText 
               Height          =   285
               Left            =   5400
               TabIndex        =   7
               Text            =   "7496"
               Top             =   600
               Width           =   1335
            End
            Begin VB.ComboBox ContractDataCombo 
               Height          =   315
               ItemData        =   "fTradeSkilDemo.frx":00BA
               Left            =   2280
               List            =   "fTradeSkilDemo.frx":00C4
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   600
               Width           =   1725
            End
            Begin VB.TextBox ContractDataServerText 
               Height          =   285
               Left            =   4080
               TabIndex        =   6
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox ContractDataClientIdText 
               Height          =   285
               Left            =   6720
               TabIndex        =   8
               Text            =   "123"
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox BrokerPortText 
               Height          =   285
               Left            =   5400
               TabIndex        =   20
               Text            =   "7496"
               Top             =   1680
               Width           =   1335
            End
            Begin VB.ComboBox BrokerCombo 
               Height          =   315
               ItemData        =   "fTradeSkilDemo.frx":00D9
               Left            =   2280
               List            =   "fTradeSkilDemo.frx":00E3
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1680
               Width           =   1725
            End
            Begin VB.TextBox BrokerServerText 
               Height          =   285
               Left            =   4080
               TabIndex        =   19
               Top             =   1680
               Width           =   1335
            End
            Begin VB.TextBox BrokerClientIdText 
               Height          =   285
               Left            =   6720
               TabIndex        =   21
               Text            =   "123"
               Top             =   1680
               Width           =   1335
            End
            Begin VB.TextBox HistDataClientIdText 
               Height          =   285
               Left            =   6720
               TabIndex        =   16
               Text            =   "123"
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox HistDataServerText 
               Height          =   285
               Left            =   4080
               TabIndex        =   14
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox HistDataPasswordText 
               Height          =   285
               Left            =   8040
               TabIndex        =   17
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox HistDataCombo 
               Height          =   315
               ItemData        =   "fTradeSkilDemo.frx":00FE
               Left            =   2280
               List            =   "fTradeSkilDemo.frx":010B
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1320
               Width           =   1725
            End
            Begin VB.TextBox HistDataPortText 
               Height          =   285
               Left            =   5400
               TabIndex        =   15
               Text            =   "7496"
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox DataPasswordText 
               Height          =   285
               Left            =   8040
               TabIndex        =   4
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox DataClientIdText 
               Height          =   285
               Left            =   6720
               TabIndex        =   3
               Text            =   "123"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox DataPortText 
               Height          =   285
               Left            =   5400
               TabIndex        =   2
               Text            =   "7496"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox DataServerText 
               Height          =   285
               Left            =   4080
               TabIndex        =   1
               Top             =   240
               Width           =   1335
            End
            Begin VB.ComboBox RealtimeDataCombo 
               Height          =   315
               ItemData        =   "fTradeSkilDemo.frx":012E
               Left            =   2280
               List            =   "fTradeSkilDemo.frx":0138
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "Secondary Contract Data"
               Height          =   375
               Left            =   0
               TabIndex        =   90
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Primary Contract Data"
               Height          =   255
               Left            =   360
               TabIndex        =   101
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "Broker"
               Height          =   255
               Left            =   360
               TabIndex        =   100
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Historical Data"
               Height          =   255
               Left            =   360
               TabIndex        =   99
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Realtime Data"
               Height          =   255
               Left            =   360
               TabIndex        =   98
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label31 
               Caption         =   "Password"
               Height          =   255
               Left            =   8160
               TabIndex        =   97
               Top             =   0
               Width           =   690
            End
            Begin VB.Label Label33 
               Caption         =   "Client id"
               Height          =   255
               Left            =   6960
               TabIndex        =   96
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label34 
               Caption         =   "Port"
               Height          =   255
               Left            =   5760
               TabIndex        =   95
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Server"
               Height          =   255
               Index           =   2
               Left            =   4440
               TabIndex        =   94
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Provider"
               Height          =   255
               Index           =   0
               Left            =   3000
               TabIndex        =   93
               Top             =   0
               Width           =   1095
            End
         End
      End
      Begin VB.ListBox TickfileList 
         Height          =   1230
         Left            =   -74400
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   600
         Width           =   6855
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68160
         TabIndex        =   85
         ToolTipText     =   "Stop tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69600
         TabIndex        =   84
         ToolTipText     =   "Pause tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton ClearTickfileListButton 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67440
         TabIndex        =   83
         ToolTipText     =   "Clear tickfile list"
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton SelectTickfilesButton 
         Caption         =   "..."
         Height          =   375
         Left            =   -67440
         TabIndex        =   82
         ToolTipText     =   "Select tickfile(s)"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -70320
         TabIndex        =   81
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton SkipReplayButton 
         Caption         =   "S&kip"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68880
         TabIndex        =   80
         ToolTipText     =   "Pause tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton ConfigureButton 
         Caption         =   "&Configure"
         Height          =   375
         Left            =   -63840
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   60
         ScaleHeight     =   3855
         ScaleWidth      =   13935
         TabIndex        =   57
         Top             =   360
         Width           =   13935
         Begin TradeBuildUI.TickerGrid TickerGrid1 
            Height          =   3855
            Left            =   3960
            TabIndex        =   104
            Top             =   0
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   6800
         End
         Begin VB.ComboBox TimeframeCombo 
            Height          =   315
            ItemData        =   "fTradeSkilDemo.frx":014F
            Left            =   2880
            List            =   "fTradeSkilDemo.frx":0151
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton GridMarketDepthButton 
            Caption         =   "Market depth"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2880
            TabIndex        =   37
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton GridChartButton 
            Caption         =   "Chart"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2880
            TabIndex        =   36
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton StopTickerButton 
            Caption         =   "Sto&p ticker"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   38
            Top             =   2760
            Width           =   975
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ticker management"
            Height          =   3855
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   2775
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   3495
               Left            =   120
               ScaleHeight     =   3495
               ScaleWidth      =   2535
               TabIndex        =   59
               Top             =   240
               Width           =   2535
               Begin VB.TextBox ExchangeText 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   30
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.TextBox LocalSymbolText 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   26
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.TextBox CurrencyText 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   31
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.TextBox StrikePriceText 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   32
                  Top             =   2160
                  Width           =   1335
               End
               Begin VB.TextBox ExpiryText 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.TextBox SymbolText 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   27
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox TypeCombo 
                  Height          =   315
                  ItemData        =   "fTradeSkilDemo.frx":0153
                  Left            =   1200
                  List            =   "fTradeSkilDemo.frx":0155
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   705
                  Width           =   1335
               End
               Begin VB.ComboBox RightCombo 
                  Height          =   315
                  Left            =   1200
                  Style           =   2  'Dropdown List
                  TabIndex        =   33
                  Top             =   2520
                  Width           =   855
               End
               Begin VB.CommandButton StartTickerButton 
                  Caption         =   "&Start ticker"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   34
                  Top             =   3000
                  Width           =   975
               End
               Begin VB.Label Label29 
                  Caption         =   "Short name"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   67
                  Top             =   0
                  Width           =   855
               End
               Begin VB.Label Label26 
                  Caption         =   "Currency"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   66
                  Top             =   1800
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Caption         =   "Exchange"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   65
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.Label Label5 
                  Caption         =   "Expiry"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   64
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Caption         =   "Type"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   63
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label Label3 
                  Caption         =   "Symbol"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   62
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label17 
                  Caption         =   "Strike price"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   61
                  Top             =   2160
                  Width           =   855
               End
               Begin VB.Label Label21 
                  Caption         =   "Right"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   60
                  Top             =   2520
                  Width           =   855
               End
            End
         End
      End
      Begin VB.CommandButton OrderButton 
         Caption         =   "&Order ticket"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   42
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ReplayProgressBar 
         Height          =   135
         Left            =   -74400
         TabIndex        =   87
         Top             =   2760
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   975
         Left            =   -74400
         TabIndex        =   102
         Top             =   3000
         Width           =   6855
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   -74400
         TabIndex        =   103
         Top             =   2520
         Width           =   6855
      End
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74400
         TabIndex        =   89
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Select tickfile(s)"
         Height          =   255
         Left            =   -74280
         TabIndex        =   88
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox StatusText 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Status messages"
      Top             =   5400
      Width           =   14175
   End
   Begin VB.Label Label27 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   360
      TabIndex        =   56
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Close"
      Height          =   255
      Left            =   10560
      TabIndex        =   53
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Low"
      Height          =   255
      Left            =   9600
      TabIndex        =   52
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "High"
      Height          =   255
      Left            =   8760
      TabIndex        =   51
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Volume"
      Height          =   255
      Left            =   7800
      TabIndex        =   50
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Last/Size"
      Height          =   255
      Left            =   4920
      TabIndex        =   49
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Ask size"
      Height          =   255
      Left            =   6840
      TabIndex        =   48
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Ask"
      Height          =   255
      Left            =   5760
      TabIndex        =   47
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Bid"
      Height          =   255
      Left            =   3960
      TabIndex        =   46
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Bid size"
      Height          =   255
      Left            =   3000
      TabIndex        =   45
      Top             =   120
      Width           =   855
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements InfoListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const TabIndexConfiguration As Long = 0
Private Const TabIndexTickers As Long = 1
Private Const TabIndexOrders As Long = 2
Private Const TabIndexExecutions As Long = 3
Private Const TabIndexReplayTickfiles As Long = 0


Private Const Timeframe1min As String = "1 min"
Private Const Timeframe2min As String = "2 min"
Private Const Timeframe5min As String = "5 min"
Private Const Timeframe15min As String = "15 min"
Private Const Timeframe30min As String = "30 min"
Private Const Timeframe1hour As String = "1 hour"
Private Const Timeframe1day As String = "Daily"

'================================================================================
' Enums
'================================================================================

Private Enum MainSSTABTabNumbers
    Configuration
    Tickers
    Orders
    ReplayTickfiles
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1
Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mListenerKey As String

Private mSimulateOrders As Boolean

Private WithEvents mTickers As Tickers
Attribute mTickers.VB_VarHelpID = -1
Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private WithEvents mTickfileManager As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1

Private mOrderForm As OrderForm
Attribute mOrderForm.VB_VarHelpID = -1

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
' ensure we get the Windows XP look and feel if running on XP
InitCommonControls
End Sub

Private Sub Form_Load()

' position at top left corner of the screen
Me.Top = 0
Me.Left = 0

Set mTradeBuildAPI = New TradeBuildAPI

' listen for log info from TradeBuild components
mListenerKey = mTradeBuildAPI.addInfoListener(Me, TradeBuildListenValueTypes.VTLog)

Set mTickers = mTradeBuildAPI.Tickers

OrdersSummary1.monitorWorkspace mTradeBuildAPI.defaultWorkSpace
ExecutionsSummary1.monitorWorkspace mTradeBuildAPI.defaultWorkSpace
TickerGrid1.monitorWorkspace mTradeBuildAPI.defaultWorkSpace

Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = True
mTimer.TimerIntervalMillisecs = 500
mTimer.StartTimer

TypeCombo.AddItem ""
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.AddItem OptionRightToString(OptionRights.OptCall)
RightCombo.AddItem OptionRightToString(OptionRights.OptPut)

TimeframeCombo.AddItem Timeframe1min
TimeframeCombo.AddItem Timeframe2min
TimeframeCombo.AddItem Timeframe5min
TimeframeCombo.AddItem Timeframe15min
TimeframeCombo.AddItem Timeframe30min
TimeframeCombo.AddItem Timeframe1hour
TimeframeCombo.AddItem Timeframe1day

TimeframeCombo.Text = Timeframe5min

RealtimeDataCombo.Text = "TWS"
ContractDataCombo.Text = "TWS"
SecContractDataCombo.Text = "(none)"
HistDataCombo.Text = "TWS"
BrokerCombo.Text = "Simulated"


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

setupServiceProviders

End Sub

Private Sub Form_Unload(cancel As Integer)
Dim i As Integer
Dim lTicker As Ticker

If Not mTimer Is Nothing Then mTimer.StopTimer

For i = Forms.Count - 1 To 0 Step -1
   Unload Forms(i)
Next

If Not mTradeBuildAPI Is Nothing Then
    For Each lTicker In mTickers
        lTicker.StopTicker
    Next
    If mListenerKey <> "" Then mTradeBuildAPI.RemoveListener mListenerKey
    Set mTradeBuildAPI = Nothing
End If

OrdersSummary1.finish
ExecutionsSummary1.finish
TickerGrid1.finish

End Sub

'================================================================================
' IListener Interface Members
'================================================================================

Private Sub InfoListener_notify( _
                ByRef ev As InfoEvent)
If ev.valueType = TradeBuildListenValueTypes.VTLog Then
    logMessage CStr(ev.Data)
End If
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub BrokerCombo_Click()
If BrokerCombo.Text = "IB via TWS" Then
    BrokerServerText.Enabled = True
    BrokerPortText.Enabled = True
    BrokerPortText = 7496
    BrokerClientIdText.Enabled = True
    If BrokerClientIdText = "" Then BrokerClientIdText = 123
Else
    BrokerServerText.Enabled = False
    BrokerPortText.Enabled = False
    BrokerClientIdText.Enabled = False
End If
End Sub

Private Sub BrokerClientIdText_Validate(cancel As Boolean)
If Not gIsInteger(BrokerClientIdText.Text) Then cancel = True
End Sub

Private Sub BrokerPortText_Validate(cancel As Boolean)
If Not gIsInteger(BrokerPortText.Text, 1024, 65535) Then cancel = True
If BrokerPortText.Text = "" Then cancel = True
End Sub

Private Sub CancelOrderButton_Click()
Dim op As OrderPlex

Set op = OrdersSummary1.SelectedItem
If Not op Is Nothing Then op.cancel

CancelOrderButton.Enabled = False
ModifyOrderButton.Enabled = False
End Sub

Private Sub ChartButton_Click()
createChart mTicker
End Sub

Private Sub ClearTickfileListButton_Click()
TickfileList.Clear
ClearTickfileListButton.Enabled = False
mTickfileManager.ClearTickfileSpecifiers
PlayTickFileButton.Enabled = False
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = False
StopReplayButton.Enabled = False
ChartButton.Enabled = False
End Sub

Private Sub ConfigureButton_Click()

ExecutionsSummary1.Clear

removeServiceProviders

setupServiceProviders

logMessage "Configuration updated"

End Sub

Private Sub ContractDataCombo_Click()
If ContractDataCombo.Text = "TradeBuild" Then
    ContractDataServerText.Enabled = False
    ContractDataPortText.Enabled = False
    ContractDataClientIdText.Enabled = False
ElseIf ContractDataCombo.Text = "TWS" Then
    ContractDataServerText.Enabled = True
    ContractDataPortText.Enabled = True
    ContractDataPortText = 7496
    ContractDataClientIdText.Enabled = True
    If ContractDataClientIdText = "" Then ContractDataClientIdText = 123
End If
End Sub

Private Sub ContractDataCombo_Validate(cancel As Boolean)
If ContractDataCombo.Text = SecContractDataCombo.Text Then cancel = True
End Sub

Private Sub CurrencyText_Change()
checkOkToStartTicker
End Sub

Private Sub DataClientIdText_Validate(cancel As Boolean)
If Not gIsInteger(DataClientIdText.Text) Then cancel = True
End Sub

Private Sub DataPortText_Validate(cancel As Boolean)
If Not gIsInteger(DataPortText.Text, 0, 65535) Then cancel = True
If DataPortText.Text = "" Then cancel = True
End Sub

Private Sub ExchangeText_Change()
checkOkToStartTicker
End Sub

Private Sub ExpiryText_Change()
checkOkToStartTicker
End Sub

Private Sub GridChartButton_Click()
Dim lTicker As Ticker

For Each lTicker In TickerGrid1.SelectedTickers
    createChart lTicker
Next
End Sub

Private Sub GridMarketDepthButton_Click()
Dim lTicker As Ticker

For Each lTicker In TickerGrid1.SelectedTickers
    showMarketDepthForm lTicker
Next
End Sub

Private Sub HistDataCombo_Click()
If HistDataCombo.Text = "TradeBuild" Then
    HistDataServerText.Enabled = False
    HistDataPortText.Enabled = False
    HistDataClientIdText.Enabled = False
    HistDataPasswordText.Enabled = False
ElseIf HistDataCombo.Text = "TWS" Then
    HistDataServerText.Enabled = True
    HistDataPortText.Enabled = True
    HistDataPortText = 7496
    HistDataClientIdText.Enabled = True
    HistDataPasswordText.Enabled = False
    If HistDataClientIdText = "" Then HistDataClientIdText = 123
ElseIf HistDataCombo.Text = "QuoteTracker" Then
    HistDataServerText.Enabled = True
    HistDataPortText.Enabled = True
    HistDataPortText = 16240
    HistDataClientIdText.Enabled = False
    HistDataPasswordText.Enabled = True
Else
    HistDataServerText.Enabled = False
    HistDataPortText.Enabled = False
    HistDataClientIdText.Enabled = False
    HistDataPasswordText.Enabled = False
End If
End Sub

Private Sub HistDataPortText_Validate(cancel As Boolean)
If Not gIsInteger(HistDataPortText.Text, 0, 65535) Then cancel = True
If HistDataPortText.Text = "" Then cancel = True
End Sub

Private Sub HistDataClientIdText_Validate(cancel As Boolean)
If Not gIsInteger(HistDataClientIdText.Text) Then cancel = True
End Sub

Private Sub LocalSymbolText_Change()
checkOkToStartTicker
End Sub

Private Sub MainSSTAB_Click(PreviousTab As Integer)
If MainSSTAB.Tab = TabIndexTickers Then
    LocalSymbolText.SetFocus
ElseIf MainSSTAB.Tab = TabIndexConfiguration Then
    RealtimeDataCombo.SetFocus
ElseIf MainSSTAB.Tab = TabIndexOrders Then
    If OrderButton.Enabled Then
        OrderButton.SetFocus
    ElseIf ModifyOrderButton.Enabled Then
        ModifyOrderButton.SetFocus
    ElseIf CancelOrderButton.Enabled Then
        CancelOrderButton.SetFocus
    End If
ElseIf MainSSTAB.Tab = TabIndexReplayTickfiles Then
    If mTickfileManager Is Nothing Then
        SelectTickfilesButton.SetFocus
    ElseIf PlayTickFileButton.Enabled Then
        PlayTickFileButton.SetFocus
    ElseIf StopReplayButton.Enabled Then
        StopReplayButton.SetFocus
    End If
End If
End Sub

Private Sub MarketDepthButton_Click()
showMarketDepthForm mTicker
End Sub

Private Sub ModifyOrderButton_Click()
Dim op As OrderPlex

Set op = OrdersSummary1.SelectedItem
If op Is Nothing Then
    ModifyOrderButton.Enabled = False
ElseIf OrdersSummary1.isSelectedItemModifiable Then
    If mOrderForm Is Nothing Then Set mOrderForm = New OrderForm
    mOrderForm.Show vbModeless
    mOrderForm.showOrderPlex op, OrdersSummary1.selectedOrderIndex
End If
End Sub

Private Sub OrderButton_Click()
If mTicker Is Nothing Then
    MsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
    Exit Sub
End If
If mOrderForm Is Nothing Then Set mOrderForm = New OrderForm
mOrderForm.Show vbModeless
mOrderForm.ordersAreSimulated = mSimulateOrders
mOrderForm.Ticker = mTicker
End Sub

Private Sub OrdersSummary1_SelectionChanged()
Dim selection As OrderPlex

Set selection = OrdersSummary1.SelectedItem

If selection Is Nothing Then
    CancelOrderButton.Enabled = False
    ModifyOrderButton.Enabled = False
Else
    CancelOrderButton.Enabled = True
    If OrdersSummary1.isSelectedItemModifiable Then
        ModifyOrderButton.Enabled = True
    End If
End If
End Sub

Private Sub PauseReplayButton_Click()
PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
logMessage "Tickfile replay paused"
mTickfileManager.PauseReplay
End Sub

Private Sub PlayTickFileButton_Click()

PlayTickFileButton.Enabled = False
SelectTickfilesButton.Enabled = False
ClearTickfileListButton.Enabled = False
PauseReplayButton.Enabled = True
SkipReplayButton.Enabled = True
StopReplayButton.Enabled = True
ReplayProgressBar.Visible = True
ConfigureButton.Enabled = False

If mTickfileManager.Ticker Is Nothing Then
    mTickfileManager.ReplayProgressEventIntervalMillisecs = 250
    logMessage "Tickfile replay started"
Else
    logMessage "Tickfile replay resumed"
End If
mTickfileManager.replaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)

mTickfileManager.StartReplay
End Sub

Private Sub RealtimeDataCombo_Click()
If RealtimeDataCombo.Text = "TWS" Then
    DataPortText = 7496
    DataClientIdText.Enabled = True
    DataPasswordText.Enabled = False
    If DataClientIdText = "" Then DataClientIdText = 123
ElseIf RealtimeDataCombo.Text = "QuoteTracker" Then
    DataPortText = 16240
    DataClientIdText.Enabled = False
    DataPasswordText.Enabled = True
End If
End Sub

Private Sub ReplaySpeedCombo_Click()
If Not mTickfileManager Is Nothing Then
    mTickfileManager.replaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)
End If
End Sub

Private Sub RightCombo_Click()
checkOkToStartTicker
End Sub

Private Sub SecContractDataCombo_Click()
If SecContractDataCombo.Text = "TradeBuild" Then
    SecContractDataServerText.Enabled = False
    SecContractDataPortText.Enabled = False
    SecContractDataClientIdText.Enabled = False
ElseIf SecContractDataCombo.Text = "TWS" Then
    SecContractDataServerText.Enabled = True
    SecContractDataPortText.Enabled = True
    SecContractDataPortText = 7496
    SecContractDataClientIdText.Enabled = True
    If SecContractDataClientIdText = "" Then SecContractDataClientIdText = 123
Else
    SecContractDataServerText.Enabled = False
    SecContractDataPortText.Enabled = False
    SecContractDataClientIdText.Enabled = False
End If
End Sub

Private Sub SecContractDataCombo_Validate(cancel As Boolean)
If ContractDataCombo.Text = SecContractDataCombo.Text Then cancel = True
End Sub

Private Sub SelectTickfilesButton_Click()
Set mTickfileManager = mTickers.createTickFileManager
mTickfileManager.ShowTickfileSelectionDialogue
End Sub

Private Sub SkipReplayButton_Click()
logMessage "Tickfile skipped"
mTickfileManager.SkipTickfile
End Sub

Private Sub StartTickerButton_Click()
Dim lTicker As Ticker
Dim lContractSpecifier As ContractSpecifier

If ExchangeText <> "" Then
    If Not mTradeBuildAPI.IsValidExchangeCode(ExchangeText) Then
        logMessage "Invalid exchange - " & ExchangeText
    Exit Sub
    End If
End If

Set lContractSpecifier = mTradeBuildAPI.newContractSpecifier( _
                                LocalSymbolText, _
                                SymbolText, _
                                ExchangeText, _
                                SecTypeFromString(TypeCombo), _
                                CurrencyText, _
                                ExpiryText, _
                                IIf(StrikePriceText = "", 0, StrikePriceText), _
                                OptionRightFromString(RightCombo))

StartTickerButton.Enabled = False

Set lTicker = createTicker
lTicker.DOMEventsRequired = DOMEvents.DOMNoEvents
lTicker.StartTicker lContractSpecifier

If LocalSymbolText <> "" Then
    LocalSymbolText.SetFocus
Else
    SymbolText.SetFocus
End If
End Sub

Private Sub StopReplayButton_Click()

PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = True
StopReplayButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
ChartButton.Enabled = False
mTicker.StopTicker
End Sub

Private Sub StopTickerButton_Click()
TickerGrid1.StopSelectedTickers
End Sub

Private Sub SymbolText_Change()
checkOkToStartTicker
End Sub

Private Sub TickerGrid1_Click()

If TickerGrid1.SelectedTickers.Count = 0 Then
    StopTickerButton.Enabled = False
    GridChartButton.Enabled = False
    GridMarketDepthButton.Enabled = False
Else
    StopTickerButton.Enabled = True
    GridChartButton.Enabled = True
    GridMarketDepthButton.Enabled = True
    MarketDepthButton.Enabled = True
    ChartButton.Enabled = True
    
    If TickerGrid1.SelectedTickers.Count = 1 Then
        
        Set mTicker = TickerGrid1.SelectedTickers.Item(1)
        
        If Not mTicker.PositionManager Is Nothing And _
            mTicker.defaultOrderContext.IsReady _
        Then
            OrderButton.Enabled = True
        End If
        
        NameText = mTicker.Contract.specifier.localSymbol
        BidSizeText = mTicker.bidSize
        BidText = mTicker.BidPriceString
        AskSizeText = mTicker.AskSize
        AskText = mTicker.AskPriceString
        LastSizeText = mTicker.TradeSize
        LastText = mTicker.TradePriceString
        VolumeText = mTicker.Volume
        HighText = mTicker.highPriceString
        LowText = mTicker.lowPriceString
        CloseText = mTicker.closePriceString
    End If
End If

End Sub

Private Sub TimeframeCombo_Click()
ChartButton.caption = TimeframeCombo.Text & "Chart"
End Sub

Private Sub TypeCombo_Click()

Select Case SecTypeFromString(TypeCombo)
Case SecurityTypes.SecTypeNone
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = True
    RightCombo.Enabled = True
Case SecurityTypes.SecTypeFuture
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case SecurityTypes.SecTypeStock
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case SecurityTypes.SecTypeOption
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = True
    RightCombo.Enabled = True
Case SecurityTypes.SecTypeFuturesOption
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = True
    RightCombo.Enabled = True
Case SecurityTypes.SecTypeCash
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case SecurityTypes.SecTypeIndex
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case SecurityTypes.SecTypeBag
    logMessage "BAG type is not implemented"
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
End Select

checkOkToStartTicker
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_ask(ev As TradeBuild.QuoteEvent)
AskText = ev.priceString
AskSizeText = ev.Size
setForeColor AskText, ev.priceChange
setForeColor AskSizeText, ev.sizeChange
End Sub

Private Sub mTicker_bid(ev As TradeBuild.QuoteEvent)
BidText = ev.priceString
BidSizeText = ev.Size
setForeColor BidText, ev.priceChange
setForeColor BidSizeText, ev.sizeChange
End Sub

Private Sub mTicker_high(ev As TradeBuild.QuoteEvent)
HighText = ev.priceString
End Sub

Private Sub mTicker_Low(ev As TradeBuild.QuoteEvent)
LowText = ev.priceString
End Sub

Private Sub mTicker_previousClose(ev As TradeBuild.QuoteEvent)
CloseText = ev.priceString
End Sub

Private Sub mTicker_trade(ev As TradeBuild.QuoteEvent)
LastText = ev.priceString
LastSizeText = ev.Size
setForeColor LastText, ev.priceChange
setForeColor LastSizeText, ev.sizeChange
End Sub

Private Sub mTicker_volume(ev As TradeBuild.QuoteEvent)
VolumeText = ev.Size
End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_Error( _
                ByRef ev As ErrorEvent)
Dim lTicker As Ticker

On Error GoTo err

Set lTicker = ev.source

Select Case ev.errorCode
Case ApiErrorCodes.ApiErrContractSpecifierAmbiguous
    logMessage "Ambiguous contract details(" & ev.errorMsg & "):" & _
                        lTicker.Contracts.ContractSpecifier.ToString
    StartTickerButton.Enabled = True
Case ApiErrorCodes.ApiErrContractSpecifierInvalid
    logMessage "Invalid contract details (" & ev.errorMsg & "):" & _
                        lTicker.Contracts.ContractSpecifier.ToString
    StartTickerButton.Enabled = True
Case ApiErrorCodes.ApiErrMarketDepthNotAvailable
    logMessage "No market depth for " & _
                        lTicker.Contract.specifier.localSymbol & _
                        ": " & ev.errorMsg
Case Else
    logMessage "Error " & ev.errorCode & ": " & ev.errorMsg
End Select

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickers_tickerError"
End Sub

Private Sub mTickers_StateChange( _
                ev As TradeBuild.StateChangeEvent)
Dim lTicker As Ticker

On Error GoTo err

Set lTicker = ev.source

Select Case ev.State
Case TickerStateCreated
    ConfigureButton.Enabled = False
Case TickerStateStarting

Case TickerStateRunning
    If lTicker Is mTicker Then
        MarketDepthButton.Enabled = True
        ChartButton.Enabled = True
        If Not lTicker.PositionManager Is Nothing And _
            lTicker.defaultOrderContext.IsReady _
        Then
            OrderButton.Enabled = True
        End If
    
        NameText = lTicker.Contract.specifier.localSymbol
        
    End If
    
    StartTickerButton.Enabled = True

Case TickerStatePaused

Case TickerStateClosing

Case TickerStateStopped
    StopTickerButton.Enabled = False
    MarketDepthButton.Enabled = False
    GridChartButton.Enabled = False
    GridMarketDepthButton.Enabled = False
    
    If lTicker Is mTicker Then
        clearTickerFields
        Set mTicker = Nothing
    End If
    
    If mTickers.Count = 0 Then ConfigureButton.Enabled = True
End Select

Exit Sub

err:
handleFatalError err.Number, err.Description, "mTickers_TickerStateEvent"
End Sub

'================================================================================
' mTickfileManager Event Handlers
'================================================================================

Private Sub mTickfileManager_Error( _
                ByRef ev As ErrorEvent)
On Error GoTo err
logMessage "Error " & ev.errorCode & ": " & ev.errorMsg

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickfileManager_Error"
End Sub

Private Sub mTickfileManager_QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As TradeBuild.Contract, _
                continueMode As TradeBuild.ReplayContinueModes)
On Error GoTo err

If tickfileIndex <> 0 Then
    clearTickerFields
    Set mTicker = Nothing
End If

ReplayProgressBar.Min = 0
ReplayProgressBar.Max = 100
ReplayProgressBar.value = 0
TickfileList.ListIndex = tickfileIndex
ReplayContractLabel.caption = Replace(pContract.specifier.ToString, vbCrLf, "; ")

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickfileManager_QueryReplayNextTickfile"
End Sub

Private Sub mTickfileManager_ReplayCompleted()
On Error GoTo err

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

logMessage "Tickfile replay completed"

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickfileManager_ReplayCompleted"
End Sub

Private Sub mTickfileManager_ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)
On Error GoTo err
ReplayProgressBar.value = percentComplete
ReplayProgressLabel.caption = tickfileTimestamp & _
                                "  Processed " & _
                                eventsPlayed & _
                                " events"

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickfileManager_ReplayProgress"
End Sub

Private Sub mTickfileManager_TickerAllocated(ByVal pTicker As TradeBuild.Ticker)
On Error GoTo err

Set mTicker = pTicker
mTicker.DOMEventsRequired = DOMProcessedEvents
mTicker.includeMarketDepthInTickfile = True

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickfileManager_TickerAllocated"
End Sub

Private Sub mTickfileManager_TickfilesSelected()
Dim tickfiles() As TradeBuild.TickfileSpecifier
Dim i As Long
On Error GoTo err
TickfileList.Clear
tickfiles = mTickfileManager.TickfileSpecifiers
For i = 0 To UBound(tickfiles)
    TickfileList.AddItem tickfiles(i).FileName
Next
checkOkToStartReplay
ClearTickfileListButton.Enabled = True

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickfileManager_TickfilesSelected"
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Dim theTime As Date
If Not mTicker Is Nothing Then
    theTime = mTicker.TimeStamp
Else
    theTime = GetTimestamp
End If
DateTimeText = FormatDateTime(theTime, vbShortDate) & vbCrLf & Format(theTime, "hh:mm:ss")
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_Error( _
                ByRef ev As ErrorEvent)
Dim spError As ServiceProviderError

On Error GoTo err

Select Case ev.errorCode
Case ApiErrorCodes.ApiErrServiceProviderError
    Set spError = mTradeBuildAPI.getServiceProviderError
    logMessage "Error from " & _
                        spError.serviceProviderName & _
                        ": code " & spError.errorCode & _
                        ": " & spError.message

Case ApiErrorCodes.ApiErrFatalError
    handleFatalError ev.errorCode, ev.errorMsg, "mTradeBuildAPI_errorMessage"
Case Else
    logMessage "Error " & ev.errorCode & ": " & ev.errorMsg
End Select


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_errorMessage"
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub checkOkToStartReplay()
If TickfileList.ListCount <> 0 Then
    PlayTickFileButton.Enabled = True
Else
    PlayTickFileButton.Enabled = False
End If
End Sub

Private Sub checkOkToStartTicker()
If LocalSymbolText <> "" Or _
    SymbolText <> "" _
Then
    StartTickerButton.Enabled = True
Else
    StartTickerButton.Enabled = False
End If
End Sub

Private Sub clearTickerFields()
NameText = ""
BidText = ""
BidSizeText = ""
AskText = ""
AskSizeText = ""
LastText = ""
LastSizeText = ""
VolumeText = ""
HighText = ""
LowText = ""
CloseText = ""
ChartButton.Enabled = False
End Sub

Private Sub createChart(ByVal pTicker As Ticker)
Dim chartform As fChart2
Dim mins As Long

Select Case TimeframeCombo.Text
Case Timeframe1min
    mins = 1
Case Timeframe2min
    mins = 2
Case Timeframe5min
    mins = 5
Case Timeframe15min
    mins = 15
Case Timeframe30min
    mins = 30
Case Timeframe1hour
    mins = 60
Case Timeframe1day
    mins = 1440
End Select

Set chartform = New fChart2
chartform.showChart pTicker, 500, 40, mins, TimePeriodUnits.TimePeriodMinute
chartform.Visible = True
End Sub

Private Function createTicker() As Ticker
Set createTicker = mTickers.Add(Not mSimulateOrders)
End Function

Private Sub handleFatalError(ByVal errNum As Long, _
                            ByVal Description As String, _
                            ByVal source As String)
Set mTicker = Nothing
removeServiceProviders

If mListenerKey <> "" Then mTradeBuildAPI.RemoveListener mListenerKey

Set mTradeBuildAPI = Nothing

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please note the error message below and email it to support@tradewright.com" & vbCrLf & _
        "Error number: " & errNum & vbCrLf & _
        "Description: " & Description & vbCrLf & _
        "Source: fTradeSkilDemo::" & source, _
        vbCritical, _
        "Fatal error"
Unload Me
End Sub

Private Sub removeServiceProviders()
mTradeBuildAPI.ServiceProviders.RemoveAll
End Sub


Private Sub setForeColor( _
                ByVal pControl As Control, _
                ByVal change As TradeBuild.ValueChanges)
If change = ValueChangeUp Then
    pControl.ForeColor = IncreasedValueColor
ElseIf change = ValueChangeDown Then
    pControl.ForeColor = DecreasedValueColor
End If
End Sub

Private Sub setupServiceProviders()
Dim realtimeServiceProvider As Object
Dim histDataServiceProvider As Object
Dim liveOrderServiceProvider As Object
Dim TWSContractServiceProvider As Object


If UseTickDBSPCheck = vbChecked Then
    ' enable historical tick data storage/retrieval to/from TradeBuild's database
    mTradeBuildAPI.ServiceProviders.Add CreateObject("TBInfoBase.TickfileServiceProvider"), LogLevelLow
End If

If UseTickfileSPCheck = vbChecked Then
    ' enable historical tick data storage and retrieval to/from various file formats
    mTradeBuildAPI.ServiceProviders.Add CreateObject("TickfileSP.TickfileServiceProvider"), LogLevelLow
End If

If UseQTTickfileSPCheck = vbChecked Then
    ' enable historical tick data retrieval from QuoteTracker
    mTradeBuildAPI.ServiceProviders.Add CreateObject("QTSP.QTTickfileServiceProvider"), LogLevelLow
End If

' enable the use of TradeBuild's built-in technical indicators
mTradeBuildAPI.ServiceProviders.Add CreateObject("BuiltInStudiesSP.StudyServiceProvider"), LogLevelLow, "Built in"

If RealtimeDataCombo.Text = "TWS" Then
    ' set up TWS realtime data service provider
    Set realtimeServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.RealtimeDataServiceProvider"), LogLevelLow)
    realtimeServiceProvider.Server = DataServerText
    realtimeServiceProvider.Port = DataPortText
    realtimeServiceProvider.clientID = DataClientIdText
    realtimeServiceProvider.providerKey = "IB"
    realtimeServiceProvider.keepConnection = True

ElseIf RealtimeDataCombo.Text = "QuoteTracker" Then
    ' set up QuoteTrackerT realtime data service provider
    Set realtimeServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("QTSP.QTRealtimeDataServiceProvider"), LogLevelLow)
    realtimeServiceProvider.QTServer = DataServerText
    realtimeServiceProvider.QTPort = DataPortText
    realtimeServiceProvider.password = DataPasswordText
    realtimeServiceProvider.providerKey = "QTIB"
    realtimeServiceProvider.keepConnection = True
End If

If ContractDataCombo.Text = "TradeBuild" Then
    mTradeBuildAPI.ServiceProviders.Add CreateObject("TBInfoBase.ContractInfoServiceProvider"), LogLevelLow
ElseIf ContractDataCombo.Text = "TWS" Then
    Set TWSContractServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.ContractInfoServiceProvider"), LogLevelLow)
    TWSContractServiceProvider.Server = ContractDataServerText
    TWSContractServiceProvider.Port = ContractDataPortText
    TWSContractServiceProvider.clientID = ContractDataClientIdText
    TWSContractServiceProvider.providerKey = "IB"
    TWSContractServiceProvider.keepConnection = True
End If

' if required, add a secondary contract info service provider - typically used to be
' able to start tickers for instruments that aren't defined in the primary contracts
' data source
If SecContractDataCombo.Text = "TradeBuild" Then
    mTradeBuildAPI.ServiceProviders.Add CreateObject("TBInfoBase.ContractInfoServiceProvider"), LogLevelLow
ElseIf SecContractDataCombo.Text = "TWS" Then
    Set TWSContractServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.ContractInfoServiceProvider"), LogLevelLow)
    TWSContractServiceProvider.Server = SecContractDataServerText
    TWSContractServiceProvider.Port = SecContractDataPortText
    TWSContractServiceProvider.clientID = SecContractDataClientIdText
    TWSContractServiceProvider.providerKey = "IB"
    TWSContractServiceProvider.keepConnection = True
End If

If HistDataCombo.Text = "TradeBuild" Then
    Set histDataServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("TBInfoBase.HistDataServiceProvider"), LogLevelLow)
ElseIf HistDataCombo.Text = "TWS" Then
    Set histDataServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.HistDataServiceProvider"), LogLevelLow)
    histDataServiceProvider.Server = HistDataServerText
    histDataServiceProvider.Port = HistDataPortText
    histDataServiceProvider.clientID = HistDataClientIdText
    histDataServiceProvider.providerKey = "IB"
    histDataServiceProvider.keepConnection = True
ElseIf HistDataCombo.Text = "QuoteTracker" Then
    Set histDataServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("QTSP.QTHistDataServiceProvider"), LogLevelLow)
    histDataServiceProvider.QTServer = HistDataServerText
    histDataServiceProvider.QTPort = HistDataPortText
    histDataServiceProvider.password = HistDataPasswordText
    histDataServiceProvider.providerKey = "QTIB"
    histDataServiceProvider.keepConnection = True
End If

If BrokerCombo.Text = "IB via TWS" Then
    ' set up TWS live order submission service provider
    Set liveOrderServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.OrderSubmissionServiceProvider"), LogLevelLow)
    liveOrderServiceProvider.Server = BrokerServerText
    liveOrderServiceProvider.Port = BrokerPortText
    liveOrderServiceProvider.clientID = BrokerClientIdText
    liveOrderServiceProvider.providerKey = "IB"
    liveOrderServiceProvider.keepConnection = True
    mSimulateOrders = False
Else
    mSimulateOrders = True
End If

End Sub

Private Sub showMarketDepthForm(ByVal pTicker As Ticker)
Dim mktDepthForm As fMarketDepth

Set mktDepthForm = New fMarketDepth
mktDepthForm.numberOfRows = 100
mktDepthForm.Ticker = pTicker
mktDepthForm.Show vbModeless
End Sub

Private Sub logMessage(message As String)
Dim timeString As String
timeString = FormatDateTime(Now, vbLongTime) & "  "
StatusText.Text = IIf(StatusText.Text <> "", _
                        StatusText.Text & vbCrLf & timeString & message, _
                        timeString & message)
StatusText.SelStart = Len(StatusText.Text)
StatusText.SelLength = 0
End Sub

