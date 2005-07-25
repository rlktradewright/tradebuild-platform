VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fTradeSkilDemo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9345
   ClientLeft      =   210
   ClientTop       =   330
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   14385
   Begin VB.CommandButton ChartButton 
      Caption         =   "C&hart"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   36
      ToolTipText     =   "Display a chart"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton MarketDepthButton 
      Caption         =   "&Market depth"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   35
      ToolTipText     =   "Display the market depth"
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox DataList 
      Height          =   2400
      ItemData        =   "fTradeSkilDemo.frx":0000
      Left            =   120
      List            =   "fTradeSkilDemo.frx":0007
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Raw socket data"
      Top             =   6840
      Width           =   14175
   End
   Begin TabDlg.SSTab MainSSTAB 
      Height          =   4335
      Left            =   120
      TabIndex        =   57
      Top             =   960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&1. Connection"
      TabPicture(0)   =   "fTradeSkilDemo.frx":0015
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&2. Tickers"
      TabPicture(1)   =   "fTradeSkilDemo.frx":0031
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3. Orders"
      TabPicture(2)   =   "fTradeSkilDemo.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "OrderButton"
      Tab(2).Control(1)=   "CancelOrderButton"
      Tab(2).Control(2)=   "ModifyOrderButton"
      Tab(2).Control(3)=   "OpenOrdersList"
      Tab(2).Control(4)=   "ExecutionsList"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "&4. Replay tickfiles"
      TabPicture(3)   =   "fTradeSkilDemo.frx":0069
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label20"
      Tab(3).Control(1)=   "ReplayContractLabel"
      Tab(3).Control(2)=   "Label19"
      Tab(3).Control(3)=   "ReplayProgressLabel"
      Tab(3).Control(4)=   "ReplayProgressBar"
      Tab(3).Control(5)=   "ReplayMarketDepthCheck"
      Tab(3).Control(6)=   "RewriteCheck"
      Tab(3).Control(7)=   "ReplaySpeedCombo"
      Tab(3).Control(8)=   "TickfileList"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "StopReplayButton"
      Tab(3).Control(10)=   "PauseReplayButton"
      Tab(3).Control(11)=   "ClearTickfileListButton"
      Tab(3).Control(12)=   "SelectTickfilesButton"
      Tab(3).Control(13)=   "PlayTickFileButton"
      Tab(3).Control(14)=   "SkipReplayButton"
      Tab(3).ControlCount=   15
      Begin VB.CommandButton SkipReplayButton 
         Caption         =   "S&kip"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68880
         TabIndex        =   32
         ToolTipText     =   "Pause tickfile replay"
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -70320
         TabIndex        =   30
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton SelectTickfilesButton 
         Caption         =   "..."
         Height          =   375
         Left            =   -67440
         TabIndex        =   25
         ToolTipText     =   "Select tickfile(s)"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton ClearTickfileListButton 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67440
         TabIndex        =   26
         ToolTipText     =   "Clear tickfile list"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69600
         TabIndex        =   31
         ToolTipText     =   "Pause tickfile replay"
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68160
         TabIndex        =   33
         ToolTipText     =   "Stop tickfile replay"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ListBox TickfileList 
         Height          =   1230
         Left            =   -74400
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   720
         Width           =   6855
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "fTradeSkilDemo.frx":0085
         Left            =   -73800
         List            =   "fTradeSkilDemo.frx":00B4
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Adjust tickfile replay speed"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox RewriteCheck 
         Caption         =   "Rewrite"
         Height          =   255
         Left            =   -72000
         TabIndex        =   28
         Top             =   2100
         Width           =   1095
      End
      Begin VB.CheckBox ReplayMarketDepthCheck 
         Caption         =   "Show market depth"
         Height          =   255
         Left            =   -72000
         TabIndex        =   29
         Top             =   2340
         Width           =   1695
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   -74880
         ScaleHeight     =   3855
         ScaleWidth      =   13935
         TabIndex        =   74
         Top             =   360
         Width           =   13935
         Begin VB.CommandButton FullChartButton 
            Caption         =   "Full chart"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2880
            TabIndex        =   93
            Top             =   1680
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid TickerGrid 
            Height          =   3735
            Left            =   3960
            TabIndex        =   34
            Top             =   120
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   6588
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
            WrapCellPointer =   -1  'True
            RowDividerStyle =   0
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.CheckBox SummaryCheck 
            Caption         =   "Check1"
            Height          =   195
            Left            =   3720
            TabIndex        =   20
            Top             =   1320
            Width           =   255
         End
         Begin VB.CommandButton GridMarketDepthButton 
            Caption         =   "Market depth"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2880
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton GridChartButton 
            Caption         =   "Chart"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2880
            TabIndex        =   18
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton StopTickerButton 
            Caption         =   "Sto&p ticker"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   21
            Top             =   2760
            Width           =   975
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ticker management"
            Height          =   3855
            Left            =   0
            TabIndex        =   75
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
               TabIndex        =   76
               Top             =   240
               Width           =   2535
               Begin VB.TextBox CurrencyText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   12
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.TextBox StrikePriceText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   13
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.TextBox ExchangeText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   11
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.TextBox ExpiryText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   10
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.TextBox SymbolText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   8
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.ComboBox TypeCombo 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "fTradeSkilDemo.frx":0158
                  Left            =   1200
                  List            =   "fTradeSkilDemo.frx":015A
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.CheckBox RecordCheck 
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   15
                  ToolTipText     =   "Write the ticker data to a tickfile for playback later"
                  Top             =   2520
                  Width           =   255
               End
               Begin VB.ComboBox RightCombo 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "fTradeSkilDemo.frx":015C
                  Left            =   1200
                  List            =   "fTradeSkilDemo.frx":015E
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Top             =   2160
                  Width           =   855
               End
               Begin VB.CheckBox MarketDepthCheck 
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   16
                  ToolTipText     =   "Write the ticker data to a tickfile for playback later"
                  Top             =   2880
                  Width           =   255
               End
               Begin VB.CommandButton StartTickerButton 
                  Caption         =   "&Start ticker"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   1560
                  TabIndex        =   17
                  Top             =   3120
                  Width           =   975
               End
               Begin VB.Label Label26 
                  Caption         =   "Currency"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   85
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Caption         =   "Exchange"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   84
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label Label5 
                  Caption         =   "Expiry"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   83
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Caption         =   "Type"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   82
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label3 
                  Caption         =   "Symbol"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   81
                  Top             =   0
                  Width           =   855
               End
               Begin VB.Label Label18 
                  Caption         =   "Record tickfile"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   80
                  Top             =   2520
                  Width           =   1455
               End
               Begin VB.Label Label17 
                  Caption         =   "Strike price"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   79
                  Top             =   1800
                  Width           =   855
               End
               Begin VB.Label Label21 
                  Caption         =   "Right"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   78
                  Top             =   2160
                  Width           =   855
               End
               Begin VB.Label Label22 
                  Caption         =   "Include market depth"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   77
                  Top             =   2880
                  Width           =   1455
               End
            End
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Summary"
            Height          =   255
            Left            =   2880
            TabIndex        =   91
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.CommandButton OrderButton 
         Caption         =   "&Order ticket"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Socket data"
         Height          =   975
         Left            =   5160
         TabIndex        =   60
         Top             =   480
         Width           =   4455
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   4215
            TabIndex        =   66
            Top             =   240
            Width           =   4215
            Begin VB.CheckBox SocketDataCheck 
               Height          =   255
               Left            =   1800
               TabIndex        =   6
               ToolTipText     =   "Write the ticker data to a tickfile for playback later"
               Top             =   0
               Width           =   255
            End
            Begin VB.CheckBox LogDataCheck 
               Height          =   255
               Left            =   1800
               TabIndex        =   7
               ToolTipText     =   "Write the ticker data to a tickfile for playback later"
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label23 
               Caption         =   "Display"
               Height          =   375
               Left            =   360
               TabIndex        =   68
               Top             =   0
               Width           =   1455
            End
            Begin VB.Label Label24 
               Caption         =   "Log to file"
               Height          =   255
               Left            =   360
               TabIndex        =   67
               Top             =   360
               Width           =   1455
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Connection details"
         Height          =   2175
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   4455
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   120
            ScaleHeight     =   1815
            ScaleWidth      =   4215
            TabIndex        =   62
            Top             =   240
            Width           =   4215
            Begin VB.CheckBox SimulateOrdersCheck 
               Height          =   255
               Left            =   1800
               TabIndex        =   3
               ToolTipText     =   "Write the ticker data to a tickfile for playback later"
               Top             =   1320
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.TextBox ServerText 
               Height          =   285
               Left            =   1800
               TabIndex        =   0
               Top             =   0
               Width           =   1335
            End
            Begin VB.TextBox ClientIDText 
               Height          =   285
               Left            =   1800
               TabIndex        =   2
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox PortText 
               Height          =   285
               Left            =   1800
               TabIndex        =   1
               Text            =   "7496"
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton ConnectButton 
               Caption         =   "&Connect"
               Enabled         =   0   'False
               Height          =   375
               Left            =   3240
               TabIndex        =   4
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton DisconnectButton 
               Caption         =   "&Disconnect"
               Enabled         =   0   'False
               Height          =   375
               Left            =   3240
               TabIndex        =   5
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label25 
               Caption         =   "Simulate orders"
               Height          =   375
               Left            =   360
               TabIndex        =   69
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Server"
               Height          =   255
               Left            =   360
               TabIndex        =   65
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label2 
               Caption         =   "Client id"
               Height          =   255
               Left            =   360
               TabIndex        =   64
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "Port"
               Height          =   255
               Left            =   360
               TabIndex        =   63
               Top             =   360
               Width           =   615
            End
         End
      End
      Begin MSComctlLib.ListView OpenOrdersList 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   72
         ToolTipText     =   "Open orders"
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ExecutionsList 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   73
         ToolTipText     =   "Filled orders"
         Top             =   2520
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2990
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ProgressBar ReplayProgressBar 
         Height          =   135
         Left            =   -74400
         TabIndex        =   87
         Top             =   2880
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   -74400
         TabIndex        =   92
         Top             =   2640
         Width           =   5655
      End
      Begin VB.Label Label19 
         Caption         =   "Select tickfile(s)"
         Height          =   255
         Left            =   -74280
         TabIndex        =   90
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   855
         Left            =   -74400
         TabIndex        =   89
         Top             =   3120
         Width           =   5655
      End
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74400
         TabIndex        =   88
         Top             =   2160
         Width           =   615
      End
   End
   Begin VB.TextBox StatusText 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Status messages"
      Top             =   5400
      Width           =   14175
   End
   Begin VB.Label Label27 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   360
      TabIndex        =   71
      Top             =   120
      Width           =   855
   End
   Begin VB.Label SymbolLabel 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   70
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label DateTimeLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   12360
      TabIndex        =   61
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Close"
      Height          =   255
      Left            =   10560
      TabIndex        =   55
      Top             =   120
      Width           =   855
   End
   Begin VB.Label CloseLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   54
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Low"
      Height          =   255
      Left            =   9600
      TabIndex        =   53
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LowLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   52
      Top             =   360
      Width           =   975
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
   Begin VB.Label HighLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   50
      Top             =   360
      Width           =   975
   End
   Begin VB.Label LastSizeLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   49
      Top             =   600
      Width           =   975
   End
   Begin VB.Label VolumeLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   48
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Volume"
      Height          =   255
      Left            =   7800
      TabIndex        =   47
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LastLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   46
      Top             =   360
      Width           =   975
   End
   Begin VB.Label AskSizeLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   45
      Top             =   360
      Width           =   975
   End
   Begin VB.Label AskLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   44
      Top             =   360
      Width           =   975
   End
   Begin VB.Label BidLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   43
      Top             =   360
      Width           =   975
   End
   Begin VB.Label BidSizeLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   42
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Last/Size"
      Height          =   255
      Left            =   4920
      TabIndex        =   41
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Ask size"
      Height          =   255
      Left            =   6840
      TabIndex        =   40
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Ask"
      Height          =   255
      Left            =   5760
      TabIndex        =   39
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Bid"
      Height          =   255
      Left            =   3960
      TabIndex        =   38
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Bid size"
      Height          =   255
      Left            =   3000
      TabIndex        =   37
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


Private WithEvents mTradeBuildAPI As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1
Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private WithEvents mTickers As Tickers
Attribute mTickers.VB_VarHelpID = -1
Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mTickerFormatString As String

Private mTimestamp As Date

Private WithEvents mOrderForm As fOrder
Attribute mOrderForm.VB_VarHelpID = -1

Private mContractCol As Collection
Private mCurrentContract As Contract

Private mOrdersCol As Collection

Private Enum TickerGridColumns
    Key
    order
    TickerName
    currencyCode
    bidSize
    bid
    ask
    AskSize
    trade
    TradeSize
    Volume
    Change
    ChangePercent
    HighPrice
    LowPrice
    ClosePrice
    Description
    symbol
    secType
    expiry
    exchange
    OptionRight
    strike
End Enum

' Percentage widths of the TickerGrid columns
Private Enum TickerGridColumnWidths
    NameWidth = 15
    CurrencyWidth = 5
    BidSizeWidth = 5
    bidWidth = 10
    askWidth = 10
    AskSizeWidth = 5
    tradeWidth = 10
    TradeSizeWidth = 5
    VolumeWidth = 10
    ChangeWidth = 8
    ChangePercentWidth = 8
    highWidth = 10
    lowWidth = 10
    closeWidth = 10
    descriptionWidth = 20
    SymbolWidth = 5
    SecTypeWidth = 10
    ExpiryWidth = 10
    ExchangeWidth = 10
    OptionRightWidth = 5
    StrikeWidth = 8
End Enum

Private Enum TickerGridSummaryColumns
    Key
    order
    TickerName
    bidSize
    bid
    ask
    AskSize
    trade
    TradeSize
    Volume
    Change
    ChangePercent
End Enum

' Percentage widths of the TickerGrid columns (summary mode)
Private Enum TickerGridSummaryColumnWidths
    NameWidth = 15
    BidSizeWidth = 5
    bidWidth = 10
    askWidth = 10
    AskSizeWidth = 5
    tradeWidth = 10
    TradeSizeWidth = 5
    VolumeWidth = 10
    ChangeWidth = 8
    ChangePercentWidth = 8
End Enum

Private Enum OpenOrdersColumns
    orderID = 1
    status
    Action
    quantity
    symbol
    orderType
    price
    auxPrice
    parentId
    ocaGroup
End Enum

' Percentage widths of the Open Orders columns
Private Const OpenOrdersOrderIDWidth = 9
Private Const OpenOrdersStatusWidth = 12
Private Const OpenOrdersActionWidth = 7
Private Const OpenOrdersQuantityWidth = 8
Private Const OpenOrdersSymbolWidth = 8
Private Const OpenOrdersOrdertypeWidth = 10
Private Const OpenOrdersPriceWidth = 10
Private Const OpenOrdersAuxPriceWidth = 10
Private Const OpenOrdersParentIDWidth = 9
Private Const OpenOrdersOCAGroupWidth = 9

Private Enum ExecutionsColumns
    execId = 1
    orderID
    Action
    quantity
    symbol
    price
    Time
End Enum

' Percentage widths of the Open Orders columns
Private Const ExecutionsExecIdWidth = 25
Private Const ExecutionsOrderIDWidth = 10
Private Const ExecutionsActionWidth = 8
Private Const ExecutionsQuantityWidth = 8
Private Const ExecutionsSymbolWidth = 8
Private Const ExecutionsPriceWidth = 10
Private Const ExecutionsTimeWidth = 23

Private Const StandardFormHeight = 7230
Private Const ExtendedFormHeight = 9750

Private Enum MainSSTABTabNumbers
    Connection
    Tickers
    Orders
    ReplayTickfiles
End Enum



Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()

Dim TickfileSP As TickfileSP.TickfileServiceProvider
Dim TBContractInfoSP As TBInfoBase.ContractInfoServiceProvider
Dim TBSQLDBTickfileSP As TBInfoBase.TickfileServiceProvider

Me.Top = 0
Me.Left = 0
Me.Height = StandardFormHeight

Set mTradeBuildAPI = New TradeBuildAPI

Set TBContractInfoSP = New TBInfoBase.ContractInfoServiceProvider
mTradeBuildAPI.ServiceProviders.Add TBContractInfoSP

mTradeBuildAPI.ServiceProviders.Add New TBInfoBase.HistDataServiceProvider

Set TBSQLDBTickfileSP = New TBInfoBase.TickfileServiceProvider
mTradeBuildAPI.ServiceProviders.Add TBSQLDBTickfileSP

Set TickfileSP = New TickfileSP.TickfileServiceProvider
mTradeBuildAPI.ServiceProviders.Add TickfileSP

Set mTickers = mTradeBuildAPI.Tickers
setupDefaultTickerGrid

Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = True
mTimer.TimerIntervalMillisecs = 500
mTimer.StartTimer

TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.AddItem optionRightToString(OptionRights.OptCall)
RightCombo.AddItem optionRightToString(OptionRights.OptPut)

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.orderID, , "ID"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.orderID).width = _
    OpenOrdersOrderIDWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.status, , "Status"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.status).width = _
    OpenOrdersStatusWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.Action, , "Action"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.Action).width = _
    OpenOrdersActionWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.quantity, , "Quant"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.quantity).width = _
    OpenOrdersQuantityWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.symbol, , "Symb"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.symbol).width = _
    OpenOrdersSymbolWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.orderType, , "Type"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.orderType).width = _
    OpenOrdersOrdertypeWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.price, , "Price"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.price).width = _
    OpenOrdersPriceWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.auxPrice, , "Aux"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.auxPrice).width = _
    OpenOrdersAuxPriceWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.parentId, , "Parent"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.parentId).width = _
    OpenOrdersParentIDWidth * OpenOrdersList.width / 100

OpenOrdersList.ColumnHeaders.Add OpenOrdersColumns.ocaGroup, , "OCA"
OpenOrdersList.ColumnHeaders(OpenOrdersColumns.ocaGroup).width = _
    OpenOrdersOCAGroupWidth * OpenOrdersList.width / 100


ExecutionsList.ColumnHeaders.Add ExecutionsColumns.execId, , "Exec id"
ExecutionsList.ColumnHeaders(ExecutionsColumns.execId).width = _
    ExecutionsExecIdWidth * ExecutionsList.width / 100

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.orderID, , "ID"
ExecutionsList.ColumnHeaders(ExecutionsColumns.orderID).width = _
    ExecutionsOrderIDWidth * ExecutionsList.width / 100

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Action, , "Action"
ExecutionsList.ColumnHeaders(ExecutionsColumns.Action).width = _
    ExecutionsActionWidth * ExecutionsList.width / 100

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.quantity, , "Quant"
ExecutionsList.ColumnHeaders(ExecutionsColumns.quantity).width = _
    ExecutionsQuantityWidth * ExecutionsList.width / 100

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.symbol, , "Symb"
ExecutionsList.ColumnHeaders(ExecutionsColumns.symbol).width = _
    ExecutionsSymbolWidth * ExecutionsList.width / 100

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.price, , "Price"
ExecutionsList.ColumnHeaders(ExecutionsColumns.price).width = _
    ExecutionsPriceWidth * ExecutionsList.width / 100

ExecutionsList.ColumnHeaders.Add ExecutionsColumns.Time, , "Time"
ExecutionsList.ColumnHeaders(ExecutionsColumns.Time).width = _
    ExecutionsTimeWidth * ExecutionsList.width / 100


ExecutionsList.SortKey = ExecutionsColumns.Time - 1
ExecutionsList.SortOrder = lvwDescending

ReplaySpeedCombo.Text = "Actual speed"
End Sub

Private Sub Form_Unload(cancel As Integer)
Dim i As Integer
Dim lTicker As Ticker

If Not mTradeBuildAPI Is Nothing Then
    For Each lTicker In mTickers
        lTicker.StopTicker
    Next
    If mTradeBuildAPI.connectionState = ConnConnected Or _
        mTradeBuildAPI.connectionState = ConnConnecting _
    Then
        mTradeBuildAPI.disconnect
    End If
    Set mTradeBuildAPI = Nothing
End If
For i = Forms.Count - 1 To 0 Step -1
   Unload Forms(i)
Next
End Sub

Private Sub CancelOrderButton_Click()
mTradeBuildAPI.cancelOrder CLng(Right$(OpenOrdersList.SelectedItem.Key, Len(OpenOrdersList.SelectedItem.Key) - 1))
CancelOrderButton.Enabled = False
ModifyOrderButton.Enabled = False
End Sub

Private Sub ChartButton_Click()
Dim tickerAppData As TickerApplicationData

Set tickerAppData = mTicker.ApplicationData
If Not tickerAppData.chartForm Is Nothing Then
    tickerAppData.chartForm.Show vbModeless
End If
End Sub

Private Sub ClearTickfileListButton_Click()
TickfileList.Clear
ClearTickfileListButton.Enabled = False
mTicker.clearTickfileNames
PlayTickFileButton.Enabled = False
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = False
StopReplayButton.Enabled = False
ChartButton.Enabled = False
End Sub

Private Sub ClientIDText_Change()
checkOKToConnect
End Sub

Private Sub ConnectButton_Click()
mTradeBuildAPI.simulateOrders = (SimulateOrdersCheck = vbChecked)
SimulateOrdersCheck.Enabled = False
mTradeBuildAPI.Connect IIf(ServerText = "", "127.0.0.1", ServerText), PortText, ClientIDText
writeStatusMessage "Attempting connection to " & _
                    IIf(ServerText = "", "local server", ServerText) & _
                    "; port=" & PortText & _
                    "; client id=" & ClientIDText
ConnectButton.Enabled = False
DisconnectButton.Enabled = True
PlayTickFileButton.Enabled = False
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = False
StopReplayButton.Enabled = False
SelectTickfilesButton.Enabled = False
ClearTickfileListButton.Enabled = False
TickfileList.Clear
End Sub

Private Sub CurrencyText_Change()
checkOkToStartTicker
End Sub

Private Sub DisconnectButton_Click()

clearTickerFields

ConnectButton.Enabled = True
DisconnectButton.Enabled = False
SimulateOrdersCheck.Enabled = True
OrderButton.Enabled = False

StartTickerButton.Enabled = False
StopTickerButton.Enabled = False
GridChartButton.Enabled = False
FullChartButton.Enabled = False
GridMarketDepthButton.Enabled = False
MarketDepthButton.Enabled = False
ChartButton.Enabled = False

SymbolText.Enabled = False
TypeCombo.Enabled = False
ExpiryText.Enabled = False
ExchangeText.Enabled = False
CurrencyText.Enabled = False
StrikePriceText.Enabled = False
RightCombo.Enabled = False
RecordCheck.Enabled = False
MarketDepthCheck.Enabled = False

Set mTicker = Nothing

OpenOrdersList.ListItems.Clear
ExecutionsList.ListItems.Clear

If Not mOrderForm Is Nothing Then Unload mOrderForm
Set mOrderForm = Nothing

mTradeBuildAPI.disconnect
ConnectButton.SetFocus
End Sub

Private Sub ExchangeText_Change()
checkOkToStartTicker
End Sub

Private Sub ExecutionsList_ColumnClick(ByVal ColumnHeader As ColumnHeader)
If ExecutionsList.SortKey = ColumnHeader.Index - 1 Then
    ExecutionsList.SortOrder = 1 - ExecutionsList.SortOrder
Else
    ExecutionsList.SortKey = ColumnHeader.Index - 1
    ExecutionsList.SortOrder = lvwAscending
End If
End Sub

Private Sub ExpiryText_Change()
checkOkToStartTicker
End Sub

Private Sub FullChartButton_Click()
Dim Ticker As Ticker
Dim bookmark As Variant
Dim chartForm As fChart1

For Each bookmark In TickerGrid.SelBookmarks
    TickerGrid.bookmark = bookmark           ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set Ticker = mTickers(TickerGrid.Text)
    Set chartForm = New fChart1
    chartForm.minimumTicksHeight = 40
    chartForm.InitialNumberOfBars = 200
    chartForm.barLength = 5
    chartForm.Ticker = Ticker
    chartForm.Visible = True
Next


End Sub

Private Sub GridChartButton_Click()
Dim Ticker As Ticker
Dim tickerAppData As TickerApplicationData
Dim bookmark As Variant

For Each bookmark In TickerGrid.SelBookmarks
    TickerGrid.bookmark = bookmark           ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set Ticker = mTickers(TickerGrid.Text)
    Set tickerAppData = Ticker.ApplicationData
    If Not tickerAppData.chartForm Is Nothing Then
        tickerAppData.chartForm.Show vbModeless
    End If
Next

End Sub

Private Sub GridMarketDepthButton_Click()
Dim Ticker As Ticker
Dim tickerAppData As TickerApplicationData
Dim bookmark As Variant

GridMarketDepthButton.Enabled = False

For Each bookmark In TickerGrid.SelBookmarks
    TickerGrid.bookmark = bookmark           ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set Ticker = mTickers(TickerGrid.Text)
    showMarketDepthForm Ticker
    If Ticker Is mTicker Then MarketDepthButton.Enabled = False
Next
End Sub

Private Sub MarketDepthButton_Click()
showMarketDepthForm mTicker
MarketDepthButton.Enabled = False
GridMarketDepthButton.Enabled = False
End Sub

Private Sub ModifyOrderButton_Click()
Dim theListitem As listItem
Dim Contract As Contract

Set theListitem = OpenOrdersList.SelectedItem
On Error Resume Next
Set Contract = mContractCol(theListitem.SubItems(OpenOrdersColumns.symbol - 1))
On Error GoTo 0
If Contract Is Nothing Then
    MsgBox "Can't modify this order - no contract details", vbExclamation, "Error"
    Exit Sub
End If

If mOrderForm Is Nothing Then Set mOrderForm = New fOrder
mOrderForm.Contract = Contract
mOrderForm.order = mOrdersCol(OpenOrdersList.SelectedItem.Key)
mOrderForm.Show vbModeless

End Sub

Private Sub OpenOrdersList_Click()
Dim status As OrderStatuses
Dim selectedOrder As order
If Not OpenOrdersList.SelectedItem Is Nothing Then
    Set selectedOrder = mOrdersCol(OpenOrdersList.SelectedItem.Key)
    status = selectedOrder.status
    Select Case status
    Case OrderStatuses.OrderStatusFilled, OrderStatuses.OrderStatusCancelled
        CancelOrderButton.Enabled = False
        ModifyOrderButton.Enabled = False
    Case OrderStatuses.OrderStatusRejected
        CancelOrderButton.Enabled = False
        ModifyOrderButton.Enabled = True
    Case Else
        CancelOrderButton.Enabled = True
        ModifyOrderButton.Enabled = True
    End Select
End If
End Sub

Private Sub OrderButton_Click()
If mCurrentContract Is Nothing Then
    MsgBox "No contract details available - please start a ticker", vbExclamation, "Error"
    Exit Sub
End If
If mOrderForm Is Nothing Then Set mOrderForm = New fOrder
mOrderForm.Contract = mCurrentContract
mOrderForm.ordersAreSimulated = mTradeBuildAPI.simulateOrders
mOrderForm.Show vbModeless
End Sub

Private Sub PauseReplayButton_Click()
PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
writeStatusMessage "Tickfile replay paused"
mTicker.PauseTicker
End Sub

Private Sub PlayTickFileButton_Click()
ServerText.Enabled = False
PortText.Enabled = False
ClientIDText.Enabled = False
SocketDataCheck.Enabled = False
LogDataCheck.Enabled = False
SimulateOrdersCheck.Enabled = False

PlayTickFileButton.Enabled = False
SelectTickfilesButton.Enabled = False
ClearTickfileListButton.Enabled = False
RewriteCheck.Enabled = False
ReplayMarketDepthCheck.Enabled = False
PauseReplayButton.Enabled = True
SkipReplayButton.Enabled = True
StopReplayButton.Enabled = True
ReplayProgressBar.Visible = True
ConnectButton.Enabled = False

If mTicker.State = TickerStateCodes.Paused Then
    writeStatusMessage "Tickfile replay resumed"
Else
    writeStatusMessage "Tickfile replay started"
End If
mTicker.ReplayProgressEventFrequency = 10
mTicker.replaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)
mTicker.StartTicker Nothing, DOMProcessedEvents, (RewriteCheck = vbChecked)
End Sub

Private Sub PortText_Change()
checkOKToConnect
End Sub

Private Sub RecordCheck_Click()
If RecordCheck = vbChecked Then
    MarketDepthCheck.Enabled = True
Else
    MarketDepthCheck.Enabled = False
End If
End Sub

Private Sub ReplaySpeedCombo_Click()
If Not mTicker Is Nothing Then
    mTicker.replaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)
End If
End Sub

Private Sub RightCombo_Click()
checkOkToStartTicker
End Sub

Private Sub SelectTickfilesButton_Click()
Set mTicker = mTickers.Add(Format(1000000000 * Rnd), "0")
mTicker.outputTickFilePath = App.Path
mTicker.selectTickFiles
End Sub

Private Sub SkipReplayButton_Click()
writeStatusMessage "Tickfile skipped"
clearTickerAppData mTicker
clearTickerFields
mTicker.SkipTicker
End Sub

Private Sub SocketDataCheck_Click()
If SocketDataCheck = vbChecked Then
    Me.Height = ExtendedFormHeight
    DataList.Visible = True
Else
    Me.Height = StandardFormHeight
    DataList.Visible = False
End If
End Sub

Private Sub StartTickerButton_Click()
Dim lTicker As Ticker
Dim lContractSpecifier As contractSpecifier

Set lContractSpecifier = New contractSpecifier
lContractSpecifier.symbol = SymbolText
lContractSpecifier.secType = secTypeFromString(TypeCombo)
lContractSpecifier.expiry = IIf(lContractSpecifier.secType = SecurityTypes.SecTypeFuture Or _
                                lContractSpecifier.secType = SecurityTypes.SecTypeFuturesOption Or _
                                lContractSpecifier.secType = SecurityTypes.SecTypeOption, _
                                ExpiryText, _
                                "")
lContractSpecifier.exchange = ExchangeText
lContractSpecifier.currencyCode = CurrencyText
If lContractSpecifier.secType = SecurityTypes.SecTypeFuturesOption Or _
    lContractSpecifier.secType = SecurityTypes.SecTypeOption _
Then
    lContractSpecifier.strike = IIf(StrikePriceText = "", 0, StrikePriceText)
    If RightCombo.Text <> "" Then
        lContractSpecifier.Right = optionRightFromString(RightCombo)
    End If
End If

StartTickerButton.Enabled = False

Set lTicker = mTickers.Add(Format(CLng(1000000000 * Rnd)), "0")
lTicker.outputTickFilePath = App.Path

lTicker.StartTicker lContractSpecifier, _
                    DOMEvents.DOMNoEvents, _
                    RecordCheck = vbChecked, _
                    RecordCheck = vbChecked And MarketDepthCheck = vbChecked

SymbolText.SetFocus
End Sub

Private Sub StopReplayButton_Click()

clearTickerAppData mTicker
clearTickerFields

PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = True
StopReplayButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
RewriteCheck.Enabled = True
ReplayMarketDepthCheck.Enabled = False
ChartButton.Enabled = False
mTicker.StopTicker
checkOKToConnect
End Sub

Private Sub StopTickerButton_Click()
Dim Ticker As Ticker
Dim bookmark As Variant

SymbolText.SetFocus

For Each bookmark In TickerGrid.SelBookmarks
    TickerGrid.bookmark = bookmark           ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set Ticker = mTickers(TickerGrid.Text)
    Ticker.StopTicker
Next

'MsgBox "Here"
End Sub

Private Sub SummaryCheck_Click()
If SummaryCheck = vbChecked Then
    setupSummaryTickerGrid
Else
    setupDefaultTickerGrid
End If
    
End Sub

Private Sub SymbolText_Change()
checkOkToStartTicker
End Sub

Private Sub TickerGrid_Error(ByVal DataError As Integer, Response As Integer)
writeStatusMessage "Ticker grid error " & DataError & ": " & TickerGrid.ErrorText
Response = 0    ' prevents the grid displaying an error message
End Sub

Private Sub TickerGrid_SelChange(cancel As Integer)
Dim Ticker As Ticker
Dim bookmark As Variant
Dim tickerAppData As TickerApplicationData

If TickerGrid.SelStartCol <> -1 Then
    StopTickerButton.Enabled = False
    GridChartButton.Enabled = False
    FullChartButton.Enabled = False
    GridMarketDepthButton.Enabled = False
Else
    ' the user has clicked on the record selectors
    StopTickerButton.Enabled = True
    GridChartButton.Enabled = True
    FullChartButton.Enabled = True
    
    If TickerGrid.SelBookmarks.Count = 1 Then
        
        TickerGrid.col = 0                       ' make the cell containing the key current
        
        Set mTicker = mTickers(TickerGrid.Text)
        Set tickerAppData = mTicker.ApplicationData
        
        If tickerAppData Is Nothing Then
            GridChartButton.Enabled = False
            FullChartButton.Enabled = False
            MarketDepthButton.Enabled = False
            GridMarketDepthButton.Enabled = False
        Else
            If tickerAppData.MarketDepthForm Is Nothing Then
                MarketDepthButton.Enabled = True
                GridMarketDepthButton.Enabled = True
            Else
                MarketDepthButton.Enabled = False
                GridMarketDepthButton.Enabled = False
            End If
            If Not tickerAppData.chartForm Is Nothing Then
                ChartButton.Enabled = True
            End If
        End If
        
        Set mCurrentContract = mTicker.Contract
        If mCurrentContract.NumberOfDecimals = 0 Then
            mTickerFormatString = "0"
        Else
            mTickerFormatString = "0." & String(mCurrentContract.NumberOfDecimals, "0")
        End If
        
        SymbolLabel.Caption = mCurrentContract.specifier.localSymbol
        BidSizeLabel.Caption = mTicker.bidSize
        BidLabel.Caption = Format(mTicker.BidPrice, mTickerFormatString)
        AskSizeLabel.Caption = mTicker.AskSize
        AskLabel.Caption = Format(mTicker.AskPrice, mTickerFormatString)
        LastSizeLabel.Caption = mTicker.TradeSize
        LastLabel.Caption = Format(mTicker.TradePrice, mTickerFormatString)
        VolumeLabel.Caption = mTicker.Volume
        HighLabel.Caption = Format(mTicker.HighPrice, mTickerFormatString)
        LowLabel.Caption = Format(mTicker.LowPrice, mTickerFormatString)
        CloseLabel.Caption = Format(mTicker.ClosePrice, mTickerFormatString)
    Else
        MarketDepthButton.Enabled = False
        GridMarketDepthButton.Enabled = False
    End If
End If

End Sub

Private Sub TypeCombo_Click()

Select Case secTypeFromString(TypeCombo)
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
    writeStatusMessage "BAG type is not implemented"
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
End Select

checkOkToStartTicker
End Sub




Private Sub mOrderForm_cancelOrder(ByVal orderID As Variant)
mTradeBuildAPI.cancelOrder orderID
CancelOrderButton.Enabled = False
ModifyOrderButton.Enabled = False
End Sub

Private Sub mOrderForm_createOrder(ByRef order As order)
Set order = mTradeBuildAPI.createOrder
End Sub

Private Sub mOrderForm_nextOCAID(id As Long)
Randomize
id = Fix(Rnd * 1999999999 + 1)
End Sub

Private Sub mOrderForm_placeOrder(ByVal pOrder As order, _
                                ByVal pContractSpecifier As contractSpecifier, _
                                ByVal passToTWS As Boolean)
openOrder pContractSpecifier, pOrder
If passToTWS Then mTradeBuildAPI.placeOrder pContractSpecifier, pOrder
End Sub

Private Sub mTicker_Application(ByVal timestamp As Date, data As Object)
Dim Ticker As Ticker
Dim tickerAppData As TickerApplicationData
Dim bookmark As Variant

' this fires when the market depth form for this ticker is unloaded. This may be
' either because the user closed the market depth form, or because the user
' stopped the ticker.

If mTicker.State = TickerStateCodes.Dead Then Exit Sub

MarketDepthButton.Enabled = True

' if the current selection in the ticker grid is this ticker, then enable
' the GridMarketDepthButton
For Each bookmark In TickerGrid.SelBookmarks
    TickerGrid.bookmark = bookmark           ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set Ticker = mTickers(TickerGrid.Text)
    If Ticker Is mTicker Then GridMarketDepthButton.Enabled = True
Next
End Sub

Private Sub mTicker_ask(ByVal timestamp As Date, _
                        ByVal price As Double, _
                        ByVal size As Long)
Dim msg As String
On Error GoTo err
mTimestamp = timestamp
AskLabel.Caption = Format(price, mTickerFormatString)
AskSizeLabel.Caption = size

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_ask"
End Sub

Private Sub mTicker_bid(ByVal timestamp As Date, _
                        ByVal price As Double, _
                        ByVal size As Long)
Dim msg As String
On Error GoTo err
mTimestamp = timestamp
BidLabel.Caption = Format(price, mTickerFormatString)
BidSizeLabel.Caption = size

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_bid"
End Sub

Private Sub mTicker_ContractInvalid(ByVal contractSpecifier As TradeBuild.contractSpecifier)
writeStatusMessage "Invalid contract details:" & vbCrLf & _
                    Replace(contractSpecifier.ToString, vbCrLf, "; ")
StartTickerButton.Enabled = True
End Sub

Private Sub mTicker_errorMessage(ByVal timestamp As Date, _
                                ByVal id As Long, _
                                ByVal errorCode As TradeBuild.ApiErrorCodes, _
                                ByVal errorMsg As String)
On Error GoTo err
mTimestamp = timestamp
writeStatusMessage "Error " & errorCode & ": " & id & ": " & errorMsg

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_errorMessage"
End Sub

Private Sub mTicker_high(ByVal timestamp As Date, _
                        ByVal price As Double)
On Error GoTo err
mTimestamp = timestamp
HighLabel.Caption = Format(price, mTickerFormatString)

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_high"
End Sub

Private Sub mTicker_low(ByVal timestamp As Date, _
                        ByVal price As Double)
On Error GoTo err
mTimestamp = timestamp
LowLabel.Caption = Format(price, mTickerFormatString)

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_low"
End Sub

Private Sub mTicker_OutputTickfileCreated(ByVal timestamp As Date, _
                            ByVal Filename As String)
writeStatusMessage "Created output tickfile: " & Filename
End Sub

Private Sub mTicker_previousClose(ByVal timestamp As Date, _
                                ByVal price As Double)
On Error GoTo err
mTimestamp = timestamp
CloseLabel.Caption = Format(price, mTickerFormatString)

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_previousClose"
End Sub

Private Sub mTicker_replayCompleted(ByVal timestamp As Date)
On Error GoTo err
mTimestamp = timestamp

clearTickerFields

If Not mOrderForm Is Nothing Then
    Unload mOrderForm
    Set mOrderForm = Nothing
End If
OrderButton.Enabled = False
ChartButton.Enabled = False
MarketDepthButton.Enabled = False
PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = False
StopReplayButton.Enabled = False

SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
RewriteCheck.Enabled = True
ReplayMarketDepthCheck.Enabled = False
ReplayProgressBar.value = 0
ReplayProgressBar.Visible = False
ReplayContractLabel.Caption = ""
ReplayProgressLabel.Caption = ""

ServerText.Enabled = True
PortText.Enabled = True
ClientIDText.Enabled = True
SocketDataCheck.Enabled = True
LogDataCheck.Enabled = True
SimulateOrdersCheck.Enabled = True
checkOKToConnect

writeStatusMessage "Tickfile replay completed"

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_replayCompleted"
End Sub

Private Sub mTicker_replayNextTickfile(ByVal tickfileIndex As Long, _
                                    ByVal tickfileName As String, _
                                    ByVal tickfileSizeBytes As Long, _
                                    ByVal pContract As Contract, _
                                    ByRef continueMode As ReplayContinueModes)
On Error GoTo err

clearTickerAppData mTicker
clearTickerFields

OpenOrdersList.ListItems.Clear
ExecutionsList.ListItems.Clear

Set mOrdersCol = New Collection
Set mContractCol = New Collection
Set mCurrentContract = pContract
mContractCol.Add pContract, pContract.specifier.localSymbol

ReplayProgressBar.Min = 0
ReplayProgressBar.Max = 100
ReplayProgressBar.value = 0
TickfileList.ListIndex = tickfileIndex
ReplayContractLabel.Caption = Replace(pContract.specifier.ToString, vbCrLf, "; ")


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_replayNextTickfile"
End Sub

Private Sub mTicker_replayProgress(ByVal tickfileTimestamp As Date, _
                                    ByVal eventsPlayed As Long, _
                                    ByVal percentComplete As Single)
On Error GoTo err
mTimestamp = tickfileTimestamp
ReplayProgressBar.value = percentComplete
ReplayProgressBar.Refresh
ReplayProgressLabel.Caption = tickfileTimestamp & _
                                "  Processed " & _
                                eventsPlayed & _
                                " events"

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_replayProgress"
End Sub

Private Sub mTicker_TickfilesSelected()
Dim tickfiles() As TradeBuild.TickfileSpecifier
Dim i As Long
TickfileList.Clear
tickfiles = mTicker.TickfileSpecifiers
For i = 0 To UBound(tickfiles)
    TickfileList.AddItem tickfiles(i).Filename
Next
checkOkToStartReplay
ClearTickfileListButton.Enabled = True
End Sub

Private Sub mTicker_trade(ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
Dim msg As String
On Error GoTo err
mTimestamp = timestamp
LastLabel.Caption = Format(price, mTickerFormatString)
LastSizeLabel.Caption = size

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_trade"
End Sub

Private Sub mTicker_volume(ByVal timestamp As Date, _
                            ByVal size As Long)
On Error GoTo err
mTimestamp = timestamp
VolumeLabel.Caption = size

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_volume"
End Sub


Private Sub mTickers_contractInvalid(ByVal pTicker As Ticker, _
                ByVal contractSpec As contractSpecifier)
writeStatusMessage "Invalid contract details:" & vbCrLf & _
                    Replace(contractSpec.ToString, vbCrLf, "; ")
StartTickerButton.Enabled = True
End Sub

Private Sub mTickers_DOMReset(ByVal Key As String, _
                            ByVal timestamp As Date, _
                            ByVal marketDepthReRequested As Boolean)

Dim tickerAppData As TickerApplicationData
Dim Ticker As Ticker

On Error GoTo err
mTimestamp = timestamp

Set Ticker = mTickers(Key)
Set tickerAppData = Ticker.ApplicationData

tickerAppData.MarketDepthForm.setDOMCell mTimestamp, DOMSides.DOMLast, CDbl(LastLabel), LastSizeLabel
tickerAppData.MarketDepthForm.setDOMCell mTimestamp, DOMSides.DOMAsk, CDbl(AskLabel), AskSizeLabel
tickerAppData.MarketDepthForm.setDOMCell mTimestamp, DOMSides.DOMBid, CDbl(BidLabel), BidSizeLabel

If marketDepthReRequested Then
    writeStatusMessage Ticker.Contract.specifier.localSymbol & ": market depth reset and data re-requested"
Else
    writeStatusMessage Ticker.Contract.specifier.localSymbol & ": market depth reset and continuing"
End If


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickers_DOMReset"
End Sub

Private Sub mTickers_DuplicateTickerRequest(ByVal pTicker As TradeBuild.Ticker, _
                                           ByVal contractSpec As contractSpecifier)
writeStatusMessage "A ticker is already running for contract: " & _
                    Replace(contractSpec.ToString, vbCrLf, "; ")
StartTickerButton.Enabled = True
End Sub

Private Sub mTickers_MarketDepthNotAvailable(ByVal Key As String, _
                            ByVal timestamp As Date, _
                            ByVal reason As String)
Dim tickerAppData As TickerApplicationData
Dim Ticker As Ticker

On Error GoTo err
mTimestamp = timestamp

Set Ticker = mTickers(Key)

writeStatusMessage "No market depth for " & _
                    Ticker.Contract.specifier.localSymbol & _
                    ": " & reason

Set tickerAppData = Ticker.ApplicationData
Unload tickerAppData.MarketDepthForm
Set tickerAppData.MarketDepthForm = Nothing

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickers_MarketDepthNotAvailable"
End Sub

Private Sub mTickers_TickerReady( _
                ByVal pTicker As Ticker)
Dim tickerAppData As TickerApplicationData

On Error GoTo err

If pTicker Is mTicker Then
    Set mCurrentContract = mTicker.Contract
    MarketDepthButton.Enabled = True
    ChartButton.Enabled = True

    SymbolLabel.Caption = mCurrentContract.specifier.localSymbol
    
    If mCurrentContract.NumberOfDecimals = 0 Then
        mTickerFormatString = "0"
    Else
        mTickerFormatString = "0." & String(mTicker.Contract.NumberOfDecimals, "0")
    End If
End If

On Error Resume Next
If mContractCol.Item(pTicker.Contract.specifier.localSymbol) Is Nothing Then
    mContractCol.Add pTicker.Contract, pTicker.Contract.specifier.localSymbol
End If

On Error GoTo err

If pTicker.ApplicationData Is Nothing Then
    Set tickerAppData = New TickerApplicationData
    pTicker.ApplicationData = tickerAppData
    Set tickerAppData.TickerProxy = pTicker.Proxy
End If

Set tickerAppData.chartForm = New fChart
tickerAppData.chartForm.Contract = pTicker.Contract
tickerAppData.chartForm.minimumTicksHeight = 40
tickerAppData.chartForm.barLength = 1
tickerAppData.chartForm.Visible = False

tickerAppData.ChartListenerKey = pTicker.AddListener(tickerAppData.chartForm, _
                                                TradeBuildListenValueTypes.ValueTypeTradeBuildTrade)

' NB: the market depth form isn't created until the
' user clicks the Market Depth button.
    
StartTickerButton.Enabled = True

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_Ready"
End Sub

Private Sub mTickers_TickerRemoved(ByVal pTicker As Ticker)

StopTickerButton.Enabled = False
MarketDepthButton.Enabled = False
GridChartButton.Enabled = False
FullChartButton.Enabled = False
GridMarketDepthButton.Enabled = False

If pTicker Is mTicker Then clearTickerFields

clearTickerAppData pTicker
End Sub

Private Sub mTimer_TimerExpired()
Dim theTime As Date
If Not mTicker Is Nothing Then
    theTime = mTicker.timestamp
Else
    theTime = GetTimestamp
End If
DateTimeLabel.Caption = Format(theTime, "dd/mm/yy") & vbCrLf & Format(theTime, "hh:mm:ss")
End Sub

Private Sub mTradeBuildAPI_connected(ByVal timestamp As Date)
Dim execFilter As ExecutionFilter

On Error GoTo err

Set mContractCol = New Collection
Set mOrdersCol = New Collection
OpenOrdersList.ListItems.Clear
ExecutionsList.ListItems.Clear

writeStatusMessage "Connected to " & _
                    IIf(ServerText = "", "local server", ServerText) & _
                    "; port=" & PortText & _
                    "; client id=" & ClientIDText
ConnectButton.Enabled = False
DisconnectButton.Enabled = True
OrderButton.Enabled = True

ServerText.Enabled = False
PortText.Enabled = False
ClientIDText.Enabled = False

MainSSTAB.Tab = MainSSTABTabNumbers.Tickers
SymbolText.Enabled = True
SymbolText.SetFocus
TypeCombo.Enabled = True
ExpiryText.Enabled = IIf(TypeCombo.Text = StrSecTypeFuture Or _
                        TypeCombo.Text = StrSecTypeOption Or _
                        TypeCombo.Text = StrSecTypeOptionFuture, _
                        True, _
                        False)
ExchangeText.Enabled = True
CurrencyText.Enabled = True
StrikePriceText.Enabled = IIf(TypeCombo.Text = StrSecTypeOption Or _
                        TypeCombo.Text = StrSecTypeOptionFuture, _
                        True, _
                        False)
RightCombo.Enabled = IIf(TypeCombo.Text = StrSecTypeOption Or _
                        TypeCombo.Text = StrSecTypeOptionFuture, _
                        True, _
                        False)
RecordCheck.Enabled = True
If RecordCheck = vbChecked Then MarketDepthCheck.Enabled = True

checkOkToStartTicker

Set execFilter = New ExecutionFilter
execFilter.clientId = ClientIDText
mTradeBuildAPI.RequestExecutions execFilter

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_connected"
End Sub

Private Sub mTradeBuildAPI_connectFailed(ByVal timestamp As Date, ByVal Description As String, ByVal retrying As Boolean)
ConnectButton.Enabled = True
DisconnectButton.Enabled = False
writeStatusMessage "Connection attempt failed"
End Sub

Private Sub mTradeBuildAPI_connecting(ByVal timestamp As Date)
On Error GoTo err
DataList.Clear
writeStatusMessage "Connecting"

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_connecting"
End Sub

Private Sub mTradeBuildAPI_connectionToIBClosed(ByVal timestamp As Date)
writeStatusMessage "Connection from TWS to IB has been lost"
End Sub

Private Sub mTradeBuildAPI_connectionToIBRecovered(ByVal timestamp As Date)
writeStatusMessage "Connection from TWS to IB has been restored successfully"
End Sub

Private Sub mTradeBuildAPI_connectionToTWSClosed(ByVal timestamp As Date)
On Error GoTo err

mTimestamp = timestamp
checkOKToConnect
DisconnectButton.Enabled = False
SimulateOrdersCheck.Enabled = True
OrderButton.Enabled = False
StartTickerButton.Enabled = False
StopTickerButton.Enabled = False
GridChartButton.Enabled = False
FullChartButton.Enabled = False
GridMarketDepthButton.Enabled = False
MarketDepthButton.Enabled = False
ChartButton.Enabled = False
Set mTicker = Nothing

ServerText.Enabled = True
PortText.Enabled = True
ClientIDText.Enabled = True

SymbolText.Enabled = False
TypeCombo.Enabled = False
ExpiryText.Enabled = False
ExchangeText.Enabled = False
CurrencyText.Enabled = False
StrikePriceText.Enabled = False
RightCombo.Enabled = False
RecordCheck.Enabled = False
MarketDepthCheck.Enabled = False

SelectTickfilesButton.Enabled = True

OpenOrdersList.ListItems.Clear
ExecutionsList.ListItems.Clear

If Not mOrderForm Is Nothing Then Unload mOrderForm
Set mOrderForm = Nothing

writeStatusMessage "Connection closed"

checkOkToStartReplay

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_connectionClosed"
End Sub

Private Sub mTradeBuildAPI_dataReceived(ByVal timestamp As Date)
Dim data As String
Static widthPx As Long
Dim width As Long
Dim fs As New FileSystemObject
Static log As TextStream
Dim logFileName As String

On Error GoTo err
mTimestamp = timestamp
If SocketDataCheck = vbChecked Or LogDataCheck = vbChecked Then
    data = mTradeBuildAPI.socketData
End If

If SocketDataCheck = vbChecked Then
    
    ' set the scrolling width of the list box if need be
    width = Me.TextWidth(data & "  ")
    If Me.ScaleMode = vbTwips Then
        ' If using Twips then change to pixels
        width = width / Screen.TwipsPerPixelX
    End If
    If width > widthPx Then
        widthPx = width
        SendMessageByNum DataList.hwnd, LB_SETHORZEXTENT, widthPx, 0
    End If
    
    
    DataList.AddItem data
    If DataList.ListCount > 10 Then DataList.TopIndex = DataList.ListCount - 10
End If

If LogDataCheck = vbChecked Then
    
    If log Is Nothing Then
        logFileName = App.Path & "\datalog" & Format(Now, "yyyymmddhhnnss") & ".txt"
        Set log = fs.CreateTextFile(logFileName, True)
        writeStatusMessage "Socket data logged to " & logFileName
    End If
    
    log.WriteLine FormatTimestamp(timestamp, TimestampFormats.TimestampDateAndTime) & "  " & data

End If
Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_dataReceived"
End Sub

Private Sub mTradeBuildAPI_errorMessage(ByVal timestamp As Date, _
                        ByVal id As Long, _
                        ByVal errorCode As ApiErrorCodes, _
                        ByVal errorMsg As String)
Dim listItem As listItem
Dim spError As ServiceProviderError

On Error GoTo err

mTimestamp = timestamp

Select Case errorCode
Case ApiErrorCodes.ServiceProviderErrorNotification
    Set spError = mTradeBuildAPI.getServiceProviderError
    writeStatusMessage "Error from " & _
                        spError.serviceProviderName & _
                        ": code " & spError.errorCode & _
                        ": id " & id & ": " _
                        & spError.message

Case Else
    writeStatusMessage "Error " & errorCode & ": " & id & ": " & errorMsg
End Select


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_errorMessage"
End Sub

Private Sub mTradeBuildAPI_executionDetails(ByVal timestamp As Date, _
                        ByVal id As Long, _
                        ByVal pContractSpecifier As contractSpecifier, _
                        ByVal exec As Execution)
Dim listItem As listItem
On Error GoTo err

mTimestamp = timestamp

On Error Resume Next
Set listItem = ExecutionsList.ListItems(CStr(exec.execId))
On Error GoTo err

If listItem Is Nothing Then
    Set listItem = ExecutionsList.ListItems.Add(, CStr(exec.execId), CStr(exec.execId))
End If

listItem.SubItems(ExecutionsColumns.Action - 1) = IIf(exec.side = ExecSides.SideBuy, "BUY", "SELL")
listItem.SubItems(ExecutionsColumns.orderID - 1) = exec.orderID
listItem.SubItems(ExecutionsColumns.price - 1) = exec.price
listItem.SubItems(ExecutionsColumns.quantity - 1) = exec.quantity
listItem.SubItems(ExecutionsColumns.symbol - 1) = pContractSpecifier.localSymbol
listItem.SubItems(ExecutionsColumns.Time - 1) = exec.Time


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_executionDetails"
End Sub

Private Sub mTradeBuildAPI_openOrder(ByVal timestamp As Date, _
                            ByVal pContractSpecifier As contractSpecifier, _
                            ByVal pOrder As order)
On Error GoTo err


mTimestamp = timestamp
openOrder pContractSpecifier, pOrder


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_openOrder"
End Sub

Private Sub mTradeBuildAPI_orderStatus(ByVal timestamp As Date, _
                                ByVal id As Long, _
                                ByVal status As OrderStatuses, _
                                ByVal filled As Long, _
                                ByVal remaining As Long, _
                                ByVal avgFillPrice As Double, _
                                ByVal permId As Long, _
                                ByVal parentId As Long, _
                                ByVal lastFillPrice As Double, _
                                ByVal clientId As Long)
Dim listItem As listItem
Dim lOrder As order
Dim orderKey As String

On Error GoTo err

mTimestamp = timestamp

orderKey = "A" & CStr(id)

On Error Resume Next
Set listItem = OpenOrdersList.ListItems(orderKey)
On Error GoTo err

If listItem Is Nothing Then
    Set listItem = OpenOrdersList.ListItems.Add(, orderKey, CStr(id))
End If

listItem.SubItems(OpenOrdersColumns.status - 1) = orderStatusToString(status)
listItem.SubItems(OpenOrdersColumns.quantity - 1) = remaining

Set lOrder = mOrdersCol(orderKey)

lOrder.status = status
lOrder.quantityFilled = filled
lOrder.quantity = remaining
lOrder.averagePrice = avgFillPrice
lOrder.permId = permId
lOrder.lastFillPrice = lastFillPrice

If (status = OrderStatuses.OrderStatusCancelled Or status = OrderStatuses.OrderStatusFilled) _
    And (Not mOrderForm Is Nothing) _
Then
    Set lOrder = mOrdersCol(orderKey)
    mOrderForm.orderCompleted lOrder
End If


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_orderStatus"
End Sub



Private Sub checkOKToConnect()
If PortText <> "" And ClientIDText <> "" And _
    mTradeBuildAPI.connectionState <> ConnReplaying _
Then
    ConnectButton.Enabled = True
Else
    ConnectButton.Enabled = False
End If
End Sub

Private Sub checkOkToStartReplay()
If TickfileList.ListCount <> 0 And _
    mTradeBuildAPI.connectionState <> ConnConnected And _
    mTradeBuildAPI.connectionState <> ConnConnecting _
Then
    PlayTickFileButton.Enabled = True
Else
    PlayTickFileButton.Enabled = False
End If
End Sub

Private Sub checkOkToStartTicker()
If SymbolText <> "" And _
    TypeCombo.Text <> "" And _
    IIf(TypeCombo.Text = StrSecTypeFuture Or _
        TypeCombo.Text = StrSecTypeOption Or _
        TypeCombo.Text = StrSecTypeOptionFuture, _
        ExpiryText <> "", _
        True) And _
    IIf(TypeCombo.Text = StrSecTypeOption Or _
        TypeCombo.Text = StrSecTypeOptionFuture, _
        StrikePriceText <> "", _
        True) And _
    IIf(TypeCombo.Text = StrSecTypeOption Or _
        TypeCombo.Text = StrSecTypeOptionFuture, _
        RightCombo <> "", _
        True) And _
    ExchangeText <> "" _
Then
    StartTickerButton.Enabled = True
Else
    StartTickerButton.Enabled = False
End If
End Sub

Private Sub clearTickerAppData(ByVal pTicker As Ticker)
Dim tickerAppData As TickerApplicationData

Set tickerAppData = pTicker.ApplicationData

If tickerAppData Is Nothing Then Exit Sub

If Not tickerAppData.MarketDepthForm Is Nothing Then
    Unload tickerAppData.MarketDepthForm
    Set tickerAppData.MarketDepthForm = Nothing
End If
If Not tickerAppData.chartForm Is Nothing Then
    Unload tickerAppData.chartForm
    Set tickerAppData.chartForm = Nothing
    pTicker.RemoveListener tickerAppData.ChartListenerKey
End If
pTicker.ApplicationData = Nothing
End Sub

Private Sub clearTickerFields()
SymbolLabel.Caption = ""
BidLabel.Caption = ""
BidSizeLabel.Caption = ""
AskLabel.Caption = ""
AskSizeLabel.Caption = ""
LastLabel.Caption = ""
LastSizeLabel.Caption = ""
VolumeLabel.Caption = ""
HighLabel.Caption = ""
LowLabel.Caption = ""
CloseLabel.Caption = ""
ChartButton.Enabled = False
End Sub

Private Sub handleFatalError(ByVal errNum As Long, _
                            ByVal Description As String, _
                            ByVal source As String)
If Not mTicker Is Nothing Then
    Set mTicker = Nothing
Else
    mTradeBuildAPI.disconnect
End If
Set mTradeBuildAPI = Nothing

MsgBox "A fatal error has occurred. The program will close" & vbCrLf & _
        "Error number: " & errNum & vbCrLf & _
        "Description: " & Description & vbCrLf & _
        "Source: fTradeSkilDemo::" & source, _
        vbCritical, _
        "Fatal error"
Unload Me
End Sub

Private Sub openOrder(ByVal pContractSpecifier As contractSpecifier, _
                ByVal pOrder As order)

Dim listItem As listItem
Dim orderKey As String

orderKey = "A" & CStr(pOrder.id)

On Error Resume Next
Set listItem = OpenOrdersList.ListItems(orderKey)
On Error GoTo 0

If listItem Is Nothing Then
    Set listItem = OpenOrdersList.ListItems.Add(, orderKey, CStr(pOrder.id))
End If

On Error Resume Next
If mOrdersCol(orderKey) Is Nothing Then
    mOrdersCol.Add pOrder, orderKey
End If
On Error GoTo 0

On Error Resume Next
If mContractCol(pContractSpecifier.localSymbol) Is Nothing Then
    mTradeBuildAPI.RequestContract pContractSpecifier
End If
On Error GoTo 0

If LCase$(listItem.SubItems(OpenOrdersColumns.status - 1)) = "filled" Then
    OpenOrdersList.ListItems.Remove (orderKey)
    If OpenOrdersList.SelectedItem Is Nothing Then
        ModifyOrderButton.Enabled = False
        CancelOrderButton.Enabled = False
    End If
    Exit Sub
End If

listItem.SubItems(OpenOrdersColumns.symbol - 1) = pContractSpecifier.localSymbol
listItem.SubItems(OpenOrdersColumns.Action - 1) = IIf(pOrder.Action = OrderActions.ActionBuy, "BUY", "SELL")
If pOrder.auxPrice <> 0 Then listItem.SubItems(OpenOrdersColumns.auxPrice - 1) = pOrder.auxPrice
listItem.SubItems(OpenOrdersColumns.ocaGroup - 1) = pOrder.ocaGroup
listItem.SubItems(OpenOrdersColumns.orderType - 1) = orderTypeToString(pOrder.orderType)
If pOrder.limitPrice <> 0 Then listItem.SubItems(OpenOrdersColumns.price - 1) = pOrder.limitPrice
listItem.SubItems(OpenOrdersColumns.quantity - 1) = pOrder.quantity
If pOrder.parentId <> 0 Then listItem.SubItems(OpenOrdersColumns.parentId - 1) = pOrder.parentId

listItem.EnsureVisible
End Sub

Private Sub setupDefaultTickerGrid()
With mTickers
    .ClearColumns
    
    .AddColumn TickerColumnIds.columnName, "Name"
    .AddColumn TickerColumnIds.ColumnCurrency, "Currency"
    .AddColumn TickerColumnIds.ColumnBidSize, "Bid size"
    .AddColumn TickerColumnIds.ColumnBid, "Bid"
    .AddColumn TickerColumnIds.ColumnAsk, "Ask"
    .AddColumn TickerColumnIds.ColumnAskSize, "Ask size"
    .AddColumn TickerColumnIds.ColumnTrade, "Last"
    .AddColumn TickerColumnIds.ColumnTradeSize, "Last size"
    .AddColumn TickerColumnIds.ColumnVolume, "Volume"
    .AddColumn TickerColumnIds.ColumnChange, "Change"
    .AddColumn TickerColumnIds.ColumnChangePercent, "% Change"
    .AddColumn TickerColumnIds.ColumnHigh, "High"
    .AddColumn TickerColumnIds.ColumnLow, "Low"
    .AddColumn TickerColumnIds.ColumnClose, "Close"
    .AddColumn TickerColumnIds.ColumnDescription, "Description"
    .AddColumn TickerColumnIds.ColumnSymbol, "Symbol"
    .AddColumn TickerColumnIds.ColumnSecType, "Sec type"
    .AddColumn TickerColumnIds.ColumnExpiry, "Expiry"
    .AddColumn TickerColumnIds.ColumnExchange, "Exchange"
    .AddColumn TickerColumnIds.ColumnRight, "Right"
    .AddColumn TickerColumnIds.ColumnStrike, "Strike"

    .Generate
End With
Set TickerGrid.DataSource = mTickers

Dim col As Column
Set col = TickerGrid.Columns(TickerGridColumns.Key)
col.Visible = False
Set col = TickerGrid.Columns(TickerGridColumns.order)
col.Visible = False
Set col = TickerGrid.Columns(TickerGridColumns.TickerName)
col.width = TickerGridColumnWidths.NameWidth * TickerGrid.width / 100
col.Alignment = dbgLeft
Set col = TickerGrid.Columns(TickerGridColumns.currencyCode)
col.width = TickerGridColumnWidths.CurrencyWidth * TickerGrid.width / 100
col.Alignment = dbgLeft
Set col = TickerGrid.Columns(TickerGridColumns.bidSize)
col.width = TickerGridColumnWidths.BidSizeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.bid)
col.width = TickerGridColumnWidths.bidWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.ask)
col.width = TickerGridColumnWidths.askWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.AskSize)
col.width = TickerGridColumnWidths.AskSizeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.trade)
col.width = TickerGridColumnWidths.tradeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.TradeSize)
col.width = TickerGridColumnWidths.TradeSizeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.Volume)
col.width = TickerGridColumnWidths.VolumeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.Change)
col.width = TickerGridColumnWidths.ChangeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.ChangePercent)
col.width = TickerGridColumnWidths.ChangePercentWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.HighPrice)
col.width = TickerGridColumnWidths.highWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.LowPrice)
col.width = TickerGridColumnWidths.lowWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.ClosePrice)
col.width = TickerGridColumnWidths.closeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.symbol)
col.width = TickerGridColumnWidths.SymbolWidth * TickerGrid.width / 100
col.Alignment = dbgCenter
Set col = TickerGrid.Columns(TickerGridColumns.secType)
col.width = TickerGridColumnWidths.SecTypeWidth * TickerGrid.width / 100
col.Alignment = dbgCenter
Set col = TickerGrid.Columns(TickerGridColumns.expiry)
col.width = TickerGridColumnWidths.ExpiryWidth * TickerGrid.width / 100
col.Alignment = dbgCenter
Set col = TickerGrid.Columns(TickerGridColumns.exchange)
col.width = TickerGridColumnWidths.ExchangeWidth * TickerGrid.width / 100
col.Alignment = dbgLeft
Set col = TickerGrid.Columns(TickerGridColumns.OptionRight)
col.width = TickerGridColumnWidths.OptionRightWidth * TickerGrid.width / 100
col.Alignment = dbgCenter
Set col = TickerGrid.Columns(TickerGridColumns.strike)
col.width = TickerGridColumnWidths.StrikeWidth * TickerGrid.width / 100
col.Alignment = dbgRight

End Sub

Private Sub setupSummaryTickerGrid()
With mTickers
    .ClearColumns
    
    .AddColumn TickerColumnIds.columnName, "Name"
    .AddColumn TickerColumnIds.ColumnBidSize, "Bid size"
    .AddColumn TickerColumnIds.ColumnBid, "Bid"
    .AddColumn TickerColumnIds.ColumnAsk, "Ask"
    .AddColumn TickerColumnIds.ColumnAskSize, "Ask size"
    .AddColumn TickerColumnIds.ColumnTrade, "Last"
    .AddColumn TickerColumnIds.ColumnTradeSize, "Last size"
    .AddColumn TickerColumnIds.ColumnVolume, "Volume"
    .AddColumn TickerColumnIds.ColumnChange, "Change"
    .AddColumn TickerColumnIds.ColumnChangePercent, "% Change"

    .Generate
End With
Set TickerGrid.DataSource = mTickers

Dim col As Column
Set col = TickerGrid.Columns(TickerGridSummaryColumns.Key)
col.Visible = False
Set col = TickerGrid.Columns(TickerGridSummaryColumns.order)
col.Visible = False
Set col = TickerGrid.Columns(TickerGridSummaryColumns.TickerName)
col.width = TickerGridSummaryColumnWidths.NameWidth * TickerGrid.width / 100
col.Alignment = dbgLeft
Set col = TickerGrid.Columns(TickerGridSummaryColumns.bidSize)
col.width = TickerGridSummaryColumnWidths.BidSizeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.bid)
col.width = TickerGridSummaryColumnWidths.bidWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.ask)
col.width = TickerGridSummaryColumnWidths.askWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.AskSize)
col.width = TickerGridSummaryColumnWidths.AskSizeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.trade)
col.width = TickerGridSummaryColumnWidths.tradeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.TradeSize)
col.width = TickerGridSummaryColumnWidths.TradeSizeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.Volume)
col.width = TickerGridSummaryColumnWidths.VolumeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.Change)
col.width = TickerGridSummaryColumnWidths.ChangeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridSummaryColumns.ChangePercent)
col.width = TickerGridSummaryColumnWidths.ChangePercentWidth * TickerGrid.width / 100
col.Alignment = dbgRight

End Sub

Private Sub showMarketDepthForm(ByVal pTicker As Ticker)
Dim msg As String
Dim tickerAppData As TickerApplicationData
Dim mktDepthForm As fMarketDepth

Set tickerAppData = pTicker.ApplicationData


If tickerAppData.MarketDepthForm Is Nothing Then
    Set mktDepthForm = New fMarketDepth
    Set tickerAppData.MarketDepthForm = mktDepthForm
End If

mktDepthForm.Contract = pTicker.Contract
If pTicker.TradePrice <> 0 Then
    mktDepthForm.initialPrice = pTicker.TradePrice
ElseIf pTicker.BidPrice <> 0 Then
    mktDepthForm.initialPrice = pTicker.BidPrice <> 0
ElseIf pTicker.AskPrice <> 0 Then
    mktDepthForm.initialPrice = pTicker.AskPrice
End If
mktDepthForm.numberOfRows = 100
mktDepthForm.numberOfVisibleRows = 20

If pTicker.TradePrice <> 0 Then
    mktDepthForm.setDOMCell mTimestamp, DOMSides.DOMLast, pTicker.TradePrice, pTicker.TradeSize
End If
If pTicker.AskPrice <> 0 Then
    mktDepthForm.setDOMCell mTimestamp, DOMSides.DOMAsk, pTicker.AskPrice, pTicker.AskSize
End If
If pTicker.BidPrice <> 0 Then
    mktDepthForm.setDOMCell mTimestamp, DOMSides.DOMBid, pTicker.BidPrice, pTicker.bidSize
End If

tickerAppData.MarketDepthListenerKey = pTicker.AddListener(mktDepthForm, _
                                            TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth)

mktDepthForm.Show vbModeless

pTicker.RequestMarketDepth DOMEvents.DOMProcessedEvents, _
                        IIf(RecordCheck = vbChecked, True, False)

End Sub

Private Sub writeStatusMessage(message As String)
Dim timeString As String
timeString = FormatDateTime(Now, vbLongTime) & "  "
StatusText.Text = IIf(StatusText.Text <> "", _
                        StatusText.Text & vbCrLf & timeString & message, _
                        timeString & message)
StatusText.SelStart = Len(StatusText.Text)
StatusText.SelLength = 0
End Sub

