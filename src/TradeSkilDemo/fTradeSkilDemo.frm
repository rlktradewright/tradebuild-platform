VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   Begin VB.TextBox DateTimeText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   12240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox CloseText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox LowText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox HighText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox VolumeText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox AskSizeText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox LastSizeText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox AskText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox LastText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox BidText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox BidSizeText 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox NameText 
      Height          =   255
      Left            =   360
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   360
      Width           =   2280
   End
   Begin VB.CommandButton ChartButton 
      Caption         =   "C&hart"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   38
      ToolTipText     =   "Display a chart"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton MarketDepthButton 
      Caption         =   "&Market depth"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   37
      ToolTipText     =   "Display the market depth"
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox DataList 
      Height          =   2400
      ItemData        =   "fTradeSkilDemo.frx":0000
      Left            =   120
      List            =   "fTradeSkilDemo.frx":0007
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Raw socket data"
      Top             =   6840
      Width           =   14175
   End
   Begin TabDlg.SSTab MainSSTAB 
      Height          =   4335
      Left            =   120
      TabIndex        =   49
      Top             =   960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&1. Connection"
      TabPicture(0)   =   "fTradeSkilDemo.frx":0015
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ConnectButton"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DisconnectButton"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&2. Tickers"
      TabPicture(1)   =   "fTradeSkilDemo.frx":0031
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3. Orders"
      TabPicture(2)   =   "fTradeSkilDemo.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "EditText"
      Tab(2).Control(1)=   "OrderPlexImageList"
      Tab(2).Control(2)=   "OrderPlexGrid"
      Tab(2).Control(3)=   "OrderButton"
      Tab(2).Control(4)=   "CancelOrderButton"
      Tab(2).Control(5)=   "ModifyOrderButton"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&4. Executions"
      TabPicture(3)   =   "fTradeSkilDemo.frx":0069
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ReplayContractLabel"
      Tab(3).Control(1)=   "ReplayProgressLabel"
      Tab(3).Control(2)=   "ExecutionsList"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "&4. Replay tickfiles"
      TabPicture(4)   =   "fTradeSkilDemo.frx":0085
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label19"
      Tab(4).Control(1)=   "Label20"
      Tab(4).Control(2)=   "ReplayProgressBar"
      Tab(4).Control(3)=   "SkipReplayButton"
      Tab(4).Control(4)=   "PlayTickFileButton"
      Tab(4).Control(5)=   "SelectTickfilesButton"
      Tab(4).Control(6)=   "ClearTickfileListButton"
      Tab(4).Control(7)=   "PauseReplayButton"
      Tab(4).Control(8)=   "StopReplayButton"
      Tab(4).Control(9)=   "TickfileList"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "ReplaySpeedCombo"
      Tab(4).Control(11)=   "RewriteCheck"
      Tab(4).Control(12)=   "ReplayMarketDepthCheck"
      Tab(4).ControlCount=   13
      Begin VB.TextBox EditText 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -62280
         TabIndex        =   119
         Text            =   "Text1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComctlLib.ImageList OrderPlexImageList 
         Left            =   -62040
         Top             =   3360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTradeSkilDemo.frx":00A1
               Key             =   "Expand"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fTradeSkilDemo.frx":04F3
               Key             =   "Contract"
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox ReplayMarketDepthCheck 
         Caption         =   "Show market depth"
         Height          =   255
         Left            =   -72000
         TabIndex        =   114
         Top             =   2220
         Width           =   1695
      End
      Begin VB.CheckBox RewriteCheck 
         Caption         =   "Rewrite"
         Height          =   255
         Left            =   -72000
         TabIndex        =   113
         Top             =   1980
         Width           =   1095
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "fTradeSkilDemo.frx":0945
         Left            =   -73800
         List            =   "fTradeSkilDemo.frx":0974
         Style           =   2  'Dropdown List
         TabIndex        =   112
         ToolTipText     =   "Adjust tickfile replay speed"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ListBox TickfileList 
         Height          =   1230
         Left            =   -74400
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   600
         Width           =   6855
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68160
         TabIndex        =   110
         ToolTipText     =   "Stop tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69600
         TabIndex        =   109
         ToolTipText     =   "Pause tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton ClearTickfileListButton 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67440
         TabIndex        =   108
         ToolTipText     =   "Clear tickfile list"
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton SelectTickfilesButton 
         Caption         =   "..."
         Height          =   375
         Left            =   -67440
         TabIndex        =   107
         ToolTipText     =   "Select tickfile(s)"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -70320
         TabIndex        =   106
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton SkipReplayButton 
         Caption         =   "S&kip"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68880
         TabIndex        =   105
         ToolTipText     =   "Pause tickfile replay"
         Top             =   1920
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid OrderPlexGrid 
         Height          =   3900
         Left            =   -74880
         TabIndex        =   104
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6879
         _Version        =   393216
         Rows            =   0
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColorFixed  =   12632256
         MergeCells      =   2
         BorderStyle     =   0
         Appearance      =   0
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
      Begin VB.Frame Frame5 
         Caption         =   "Historical Data Source Connection details"
         Height          =   2175
         Left            =   7320
         TabIndex        =   97
         Top             =   480
         Width           =   3495
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   90
            ScaleHeight     =   1815
            ScaleWidth      =   3375
            TabIndex        =   98
            Top             =   225
            Width           =   3375
            Begin VB.ComboBox HistDataSourceCombo 
               Height          =   315
               ItemData        =   "fTradeSkilDemo.frx":0A18
               Left            =   1800
               List            =   "fTradeSkilDemo.frx":0A25
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   0
               Width           =   1365
            End
            Begin VB.TextBox HistPasswordText 
               Height          =   285
               Left            =   1800
               TabIndex        =   13
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox HistServerText 
               Height          =   285
               Left            =   1800
               TabIndex        =   10
               Top             =   330
               Width           =   1335
            End
            Begin VB.TextBox HistClientIdText 
               Height          =   285
               Left            =   1800
               TabIndex        =   12
               Top             =   810
               Width           =   1335
            End
            Begin VB.TextBox HistPortText 
               Height          =   285
               Left            =   1800
               TabIndex        =   11
               Text            =   "7496"
               Top             =   570
               Width           =   1335
            End
            Begin VB.Label Label35 
               Caption         =   "Password"
               Height          =   255
               Left            =   315
               TabIndex        =   103
               Top             =   1125
               Width           =   690
            End
            Begin VB.Label Label1 
               Caption         =   "Server"
               Height          =   255
               Index           =   4
               Left            =   315
               TabIndex        =   102
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Data source"
               Height          =   255
               Index           =   1
               Left            =   315
               TabIndex        =   101
               Top             =   15
               Width           =   1095
            End
            Begin VB.Label Label32 
               Caption         =   "Client id"
               Height          =   255
               Left            =   315
               TabIndex        =   100
               Top             =   870
               Width           =   615
            End
            Begin VB.Label Label30 
               Caption         =   "Port"
               Height          =   255
               Left            =   315
               TabIndex        =   99
               Top             =   615
               Width           =   615
            End
         End
      End
      Begin VB.CommandButton DisconnectButton 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton ConnectButton 
         Caption         =   "&Connect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   2880
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Realtime Data Source Connection details"
         Height          =   2175
         Left            =   3720
         TabIndex        =   90
         Top             =   480
         Width           =   3495
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   90
            ScaleHeight     =   1815
            ScaleWidth      =   3375
            TabIndex        =   91
            Top             =   225
            Width           =   3375
            Begin VB.TextBox DataSourcePortText 
               Height          =   285
               Left            =   1800
               TabIndex        =   6
               Text            =   "7496"
               Top             =   570
               Width           =   1335
            End
            Begin VB.TextBox DataSourceClientIdText 
               Height          =   285
               Left            =   1800
               TabIndex        =   7
               Top             =   810
               Width           =   1335
            End
            Begin VB.TextBox DataSourceServerText 
               Height          =   285
               Left            =   1800
               TabIndex        =   5
               Top             =   330
               Width           =   1335
            End
            Begin VB.TextBox DataSourcePasswordText 
               Height          =   285
               Left            =   1800
               TabIndex        =   8
               Top             =   1080
               Width           =   1335
            End
            Begin VB.ComboBox RealtimeDataSourceCombo 
               Height          =   315
               ItemData        =   "fTradeSkilDemo.frx":0A48
               Left            =   1800
               List            =   "fTradeSkilDemo.frx":0A52
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   0
               Width           =   1365
            End
            Begin VB.Label Label34 
               Caption         =   "Port"
               Height          =   255
               Left            =   315
               TabIndex        =   96
               Top             =   615
               Width           =   615
            End
            Begin VB.Label Label33 
               Caption         =   "Client id"
               Height          =   255
               Left            =   315
               TabIndex        =   95
               Top             =   870
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Data source"
               Height          =   255
               Index           =   3
               Left            =   315
               TabIndex        =   94
               Top             =   15
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Server"
               Height          =   255
               Index           =   2
               Left            =   315
               TabIndex        =   93
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label31 
               Caption         =   "Password"
               Height          =   255
               Left            =   315
               TabIndex        =   92
               Top             =   1125
               Width           =   690
            End
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   -74940
         ScaleHeight     =   3855
         ScaleWidth      =   13935
         TabIndex        =   61
         Top             =   360
         Width           =   13935
         Begin MSDataGridLib.DataGrid TickerGrid 
            Height          =   3735
            Left            =   3960
            TabIndex        =   36
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
            TabIndex        =   31
            Top             =   1320
            Width           =   255
         End
         Begin VB.CommandButton GridMarketDepthButton 
            Caption         =   "Market depth"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2880
            TabIndex        =   30
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton GridChartButton 
            Caption         =   "Chart"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2880
            TabIndex        =   29
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton StopTickerButton 
            Caption         =   "Sto&p ticker"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   32
            Top             =   2760
            Width           =   975
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ticker management"
            Height          =   3855
            Left            =   0
            TabIndex        =   62
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
               TabIndex        =   63
               Top             =   240
               Width           =   2535
               Begin VB.TextBox LocalSymbolText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   18
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.TextBox CurrencyText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   23
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.TextBox StrikePriceText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   24
                  Top             =   2160
                  Width           =   1335
               End
               Begin VB.TextBox ExchangeText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   22
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.TextBox ExpiryText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   21
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.TextBox SymbolText 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   19
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox TypeCombo 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "fTradeSkilDemo.frx":0A69
                  Left            =   1200
                  List            =   "fTradeSkilDemo.frx":0A6B
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   705
                  Width           =   1335
               End
               Begin VB.CheckBox RecordCheck 
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   26
                  ToolTipText     =   "Write the ticker data to a tickfile for playback later"
                  Top             =   2880
                  Width           =   255
               End
               Begin VB.ComboBox RightCombo 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "fTradeSkilDemo.frx":0A6D
                  Left            =   1200
                  List            =   "fTradeSkilDemo.frx":0A6F
                  Style           =   2  'Dropdown List
                  TabIndex        =   25
                  Top             =   2520
                  Width           =   855
               End
               Begin VB.CheckBox MarketDepthCheck 
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   27
                  ToolTipText     =   "Write the ticker data to a tickfile for playback later"
                  Top             =   3120
                  Width           =   255
               End
               Begin VB.CommandButton StartTickerButton 
                  Caption         =   "&Start ticker"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   1560
                  TabIndex        =   28
                  Top             =   3120
                  Width           =   975
               End
               Begin VB.Label Label29 
                  Caption         =   "Short name"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   76
                  Top             =   0
                  Width           =   855
               End
               Begin VB.Label Label26 
                  Caption         =   "Currency"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   72
                  Top             =   1800
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Caption         =   "Exchange"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   71
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.Label Label5 
                  Caption         =   "Expiry"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Caption         =   "Type"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   69
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label Label3 
                  Caption         =   "Symbol"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   68
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label18 
                  Caption         =   "Record tickfile"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   67
                  Top             =   2880
                  Width           =   1455
               End
               Begin VB.Label Label17 
                  Caption         =   "Strike price"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   66
                  Top             =   2160
                  Width           =   855
               End
               Begin VB.Label Label21 
                  Caption         =   "Right"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   65
                  Top             =   2520
                  Width           =   855
               End
               Begin VB.Label Label22 
                  Caption         =   "Include market depth"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   64
                  Top             =   3120
                  Width           =   1455
               End
            End
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Summary"
            Height          =   255
            Left            =   2880
            TabIndex        =   74
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.CommandButton OrderButton 
         Caption         =   "&Order ticket"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   35
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Socket data"
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   2760
         Width           =   3495
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   3300
            TabIndex        =   56
            Top             =   240
            Width           =   3300
            Begin VB.CheckBox SocketDataCheck 
               Height          =   255
               Left            =   1800
               TabIndex        =   16
               ToolTipText     =   "Write the ticker data to a tickfile for playback later"
               Top             =   0
               Width           =   255
            End
            Begin VB.CheckBox LogDataCheck 
               Height          =   255
               Left            =   1800
               TabIndex        =   17
               ToolTipText     =   "Write the ticker data to a tickfile for playback later"
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label23 
               Caption         =   "Display"
               Height          =   375
               Left            =   360
               TabIndex        =   58
               Top             =   0
               Width           =   1455
            End
            Begin VB.Label Label24 
               Caption         =   "Log to file"
               Height          =   255
               Left            =   360
               TabIndex        =   57
               Top             =   360
               Width           =   1455
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "TWS Connection details"
         Height          =   2175
         Left            =   120
         TabIndex        =   51
         Top             =   480
         Width           =   3495
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   90
            ScaleHeight     =   1815
            ScaleWidth      =   3375
            TabIndex        =   53
            Top             =   225
            Width           =   3375
            Begin VB.TextBox ServerText 
               Height          =   285
               Left            =   1800
               TabIndex        =   0
               Top             =   330
               Width           =   1335
            End
            Begin VB.CheckBox SimulateOrdersCheck 
               Height          =   255
               Left            =   1800
               TabIndex        =   3
               ToolTipText     =   "Write the ticker data to a tickfile for playback later"
               Top             =   1425
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.TextBox ClientIDText 
               Height          =   285
               Left            =   1800
               TabIndex        =   2
               Top             =   810
               Width           =   1335
            End
            Begin VB.TextBox PortText 
               Height          =   285
               Left            =   1800
               TabIndex        =   1
               Text            =   "7496"
               Top             =   570
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Server"
               Height          =   255
               Index           =   0
               Left            =   315
               TabIndex        =   89
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label25 
               Caption         =   "Simulate orders"
               Height          =   375
               Left            =   315
               TabIndex        =   59
               Top             =   1425
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Client id"
               Height          =   255
               Left            =   315
               TabIndex        =   55
               Top             =   870
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "Port"
               Height          =   255
               Left            =   315
               TabIndex        =   54
               Top             =   615
               Width           =   615
            End
         End
      End
      Begin MSComctlLib.ProgressBar ReplayProgressBar 
         Height          =   135
         Left            =   -74400
         TabIndex        =   115
         Top             =   2760
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView ExecutionsList 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   118
         ToolTipText     =   "Filled orders"
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6800
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
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74400
         TabIndex        =   117
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Select tickfile(s)"
         Height          =   255
         Left            =   -74280
         TabIndex        =   116
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   -74400
         TabIndex        =   75
         Top             =   2640
         Width           =   5655
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   855
         Left            =   -74400
         TabIndex        =   73
         Top             =   3120
         Width           =   5655
      End
   End
   Begin VB.TextBox StatusText 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Status messages"
      Top             =   5400
      Width           =   14175
   End
   Begin VB.Label Label27 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   360
      TabIndex        =   60
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Close"
      Height          =   255
      Left            =   10560
      TabIndex        =   47
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Low"
      Height          =   255
      Left            =   9600
      TabIndex        =   46
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "High"
      Height          =   255
      Left            =   8760
      TabIndex        =   45
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Volume"
      Height          =   255
      Left            =   7800
      TabIndex        =   44
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Last/Size"
      Height          =   255
      Left            =   4920
      TabIndex        =   43
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Ask size"
      Height          =   255
      Left            =   6840
      TabIndex        =   42
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Ask"
      Height          =   255
      Left            =   5760
      TabIndex        =   41
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Bid"
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Bid size"
      Height          =   255
      Left            =   3000
      TabIndex        =   39
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

Implements TradeBuild.QuoteListener
Implements TradeBuild.ChangeListener
Implements TradeBuild.ProfitListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

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

Private Const RowDataOrderPlexBase As Long = &H100
Private Const RowDataPositionManagerBase As Long = &H1000000

'================================================================================
' Enums
'================================================================================

Private Enum ExecutionsColumns
    execId = 1
    orderId
    Action
    quantity
    symbol
    price
    Time
End Enum

Private Enum MainSSTABTabNumbers
    Connection
    Tickers
    Orders
    ReplayTickfiles
End Enum

Private Enum OPGridColumns
    symbol
    ExpandIndicator
    OtherColumns    ' keep this entry last
End Enum

Private Enum OPGridOrderPlexColumns
    creationTime = OPGridColumns.OtherColumns
    size
    profit
    MaxProfit
    drawdown
    currencyCode
End Enum

Private Enum OPGridPositionColumns
    exchange = OPGridColumns.OtherColumns
    size
    profit
    MaxProfit
    drawdown
    currencyCode
End Enum

Private Enum OPGridOrderColumns
    typeInPlex = OPGridColumns.OtherColumns
    size
    averagePrice
    Status
    Action
    quantity
    orderType
    price
    auxPrice
    LastFillTime
    lastFillPrice
    id
    VendorId
End Enum

Private Enum OPGridColumnWidths
    ExpandIndicatorWidth = 3
    SymbolWidth = 15
End Enum

Private Enum OPGridOrderPlexColumnWidths
    CreationTimeWidth = 17
    SizeWidth = 5
    ProfitWidth = 8
    MaxProfitWidth = 8
    DrawdownWidth = 8
    CurrencyCodeWidth = 3
End Enum

Private Enum OPGridPositionColumnWidths
    ExchangeWidth = 9
    SizeWidth = 5
    ProfitWidth = 8
    MaxProfitWidth = 8
    DrawdownWidth = 8
    CurrencyCodeWidth = 5
End Enum

Private Enum OPGridOrderColumnWidths
    TypeInPlexWidth = 9
    SizeWidth = 5
    AveragePriceWidth = 9
    StatusWidth = 15
    ActionWidth = 5
    QuantityWidth = 7
    OrderTypeWidth = 7
    PriceWidth = 9
    AuxPriceWidth = 9
    LastFillTimeWidth = 17
    LastFillPriceWidth = 9
    IdWidth = 10
    VendorIdWidth = 10
End Enum

Private Enum TickerGridColumns
    Key
    Order
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
    highPrice
    lowPrice
    closePrice
    Description
    symbol
    sectype
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
    Order
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

'================================================================================
' Types
'================================================================================

Private Type OrderPlexGridMappingEntry
    op                  As TradeBuild.OrderPlex
    
    ' indicates whether this entry in the grid is expanded
    isExpanded          As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    gridIndex           As Long
                                
    ' offset from gridIndex of line in OrdersGrid relating to
    ' the corresponding order: -1 means  it's not in the grid
    entryGridOffset      As Long
    stopGridOffset       As Long
    targetGridOffset     As Long
    closeoutGridOffset   As Long
    
End Type

Private Type PositionManagerGridMappingEntry
    pm                  As TradeBuild.PositionManager
    
    ' indicates whether this entry in the grid is expanded
    isExpanded          As Boolean
    
    ' index of first line in OrdersGrid relating to this entry
    gridIndex           As Long
                                
End Type

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1
Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mTWSContractServiceProvider As Object
Private mRealtimeServiceProvider As Object
Private mHistDataServiceProvider As Object

Private WithEvents mTickers As Tickers
Attribute mTickers.VB_VarHelpID = -1
Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private WithEvents mTickfileManager As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1
Private mTimestamp As Date

Private mOrderForm As OrderForm
Attribute mOrderForm.VB_VarHelpID = -1

Private mSelectedOrderPlexGridRow As Long
Private mSelectedOrderPlex As TradeBuild.OrderPlex
Private mSelectedOrder As TradeBuild.Order

Private mContractCol As Collection
Private mCurrentContract As Contract

Private mOrdersCol As Collection

Private mOrderPlexGridMappingTable() As OrderPlexGridMappingEntry
Private mMaxOrderPlexGridMappingTableIndex As Long

Private mPositionManagerGridMappingTable() As PositionManagerGridMappingEntry
Private mMaxPositionManagerGridMappingTableIndex As Long

' the index of the first entry in the order plex frid that relates to
' order plexes (rather than header rows, currency totals etc)
Private mFirstOrderPlexGridRowIndex As Long

Private mLetterWidth As Single
Private mDigitWidth As Single

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()

Dim widthString As String
widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = Me.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = Me.TextWidth(widthString) / Len(widthString)

Set gMainForm = Me

Me.Top = 0
Me.Left = 0
Me.Height = StandardFormHeight

Set mTradeBuildAPI = New TradeBuildAPI
Set gTradeBuildAPI = mTradeBuildAPI
mTradeBuildAPI.ConnectionRetryIntervalSecs = 10

mTradeBuildAPI.ServiceProviders.Add CreateObject("TBInfoBase.ContractInfoServiceProvider")

mTradeBuildAPI.ServiceProviders.Add CreateObject("TBInfoBase.TickfileServiceProvider")

mTradeBuildAPI.ServiceProviders.Add CreateObject("TickfileSP.TickfileServiceProvider")

mTradeBuildAPI.ServiceProviders.Add CreateObject("QTSP.QTTickfileServiceProvider")

Set mTickers = mTradeBuildAPI.Tickers

setupDefaultTickerGrid
setupOrderPlexGrid

ReDim mOrderPlexGridMappingTable(50) As OrderPlexGridMappingEntry
mMaxOrderPlexGridMappingTableIndex = -1

ReDim mPositionManagerGridMappingTable(20) As PositionManagerGridMappingEntry
mMaxPositionManagerGridMappingTableIndex = -1

Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = True
mTimer.TimerIntervalMillisecs = 500
mTimer.StartTimer

TypeCombo.AddItem ""
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.AddItem optionRightToString(OptionRights.OptCall)
RightCombo.AddItem optionRightToString(OptionRights.OptPut)

ExecutionsList.columnHeaders.Add ExecutionsColumns.execId, , "Exec id"
ExecutionsList.columnHeaders(ExecutionsColumns.execId).width = _
    ExecutionsExecIdWidth * ExecutionsList.width / 100

ExecutionsList.columnHeaders.Add ExecutionsColumns.orderId, , "ID"
ExecutionsList.columnHeaders(ExecutionsColumns.orderId).width = _
    ExecutionsOrderIDWidth * ExecutionsList.width / 100

ExecutionsList.columnHeaders.Add ExecutionsColumns.Action, , "Action"
ExecutionsList.columnHeaders(ExecutionsColumns.Action).width = _
    ExecutionsActionWidth * ExecutionsList.width / 100

ExecutionsList.columnHeaders.Add ExecutionsColumns.quantity, , "Quant"
ExecutionsList.columnHeaders(ExecutionsColumns.quantity).width = _
    ExecutionsQuantityWidth * ExecutionsList.width / 100

ExecutionsList.columnHeaders.Add ExecutionsColumns.symbol, , "Symb"
ExecutionsList.columnHeaders(ExecutionsColumns.symbol).width = _
    ExecutionsSymbolWidth * ExecutionsList.width / 100

ExecutionsList.columnHeaders.Add ExecutionsColumns.price, , "Price"
ExecutionsList.columnHeaders(ExecutionsColumns.price).width = _
    ExecutionsPriceWidth * ExecutionsList.width / 100

ExecutionsList.columnHeaders.Add ExecutionsColumns.Time, , "Time"
ExecutionsList.columnHeaders(ExecutionsColumns.Time).width = _
    ExecutionsTimeWidth * ExecutionsList.width / 100


ExecutionsList.SortKey = ExecutionsColumns.Time - 1
ExecutionsList.SortOrder = lvwDescending

RealtimeDataSourceCombo.Text = "TWS"
HistDataSourceCombo.Text = "TradeBuild"

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
        removeServiceProviders
        mTradeBuildAPI.disconnect
    End If
    Set mTradeBuildAPI = Nothing
End If
For i = Forms.Count - 1 To 0 Step -1
   Unload Forms(i)
Next
End Sub

'================================================================================
' ChangeListener Interface Members
'================================================================================

Private Sub ChangeListener_Change(ev As TradeBuild.ChangeEvent)
If TypeOf ev.source Is TradeBuild.OrderPlex Then
    Dim opChangeType As TradeBuild.OrderPlexChangeTypes
    Dim op As TradeBuild.OrderPlex
    Dim opIndex As Long
    
    Set op = ev.source
    
    opIndex = findOrderPlexTableIndex(op)
    
    With mOrderPlexGridMappingTable(opIndex)
    
        opChangeType = ev.ChangeType
        
        Select Case opChangeType
        Case OrderPlexChangeTypes.Created
        
        Case OrderPlexChangeTypes.Completed
    
        Case OrderPlexChangeTypes.SelfCancelled
    
        Case OrderPlexChangeTypes.EntryOrderChanged
            displayOrderValuesInOrderPlexGrid .gridIndex + .entryGridOffset, op.entryOrder
        Case OrderPlexChangeTypes.StopOrderChanged
            displayOrderValuesInOrderPlexGrid .gridIndex + .stopGridOffset, op.stopOrder
        Case OrderPlexChangeTypes.TargetOrderChanged
            displayOrderValuesInOrderPlexGrid .gridIndex + .targetGridOffset, op.targetOrder
        Case OrderPlexChangeTypes.CloseoutOrderCreated
            If .targetGridOffset >= 0 Then
                .closeoutGridOffset = .targetGridOffset + 1
            ElseIf .stopGridOffset >= 0 Then
                .closeoutGridOffset = .stopGridOffset + 1
            ElseIf .entryGridOffset >= 0 Then
                .closeoutGridOffset = .entryGridOffset + 1
            Else
                .closeoutGridOffset = 1
            End If
            
            addOrderEntryToOrderPlexGrid .gridIndex + .closeoutGridOffset, _
                                    .op.Contract.specifier.symbol, _
                                    op.closeoutOrder, _
                                    opIndex, _
                                    "Closeout"
        Case OrderPlexChangeTypes.CloseoutOrderChanged
            displayOrderValuesInOrderPlexGrid .gridIndex + .targetGridOffset, _
                                                op.closeoutOrder
        Case OrderPlexChangeTypes.ProfitThresholdExceeded
    
        Case OrderPlexChangeTypes.LossThresholdExceeded
    
        Case OrderPlexChangeTypes.DrawdownThresholdExceeded
    
        Case OrderPlexChangeTypes.SizeChanged
            OrderPlexGrid.TextMatrix(.gridIndex, OPGridOrderPlexColumns.size) = op.size
        Case OrderPlexChangeTypes.StateChanged
            If op.State <> OrderPlexStateCodes.Created And _
                op.State <> OrderPlexStateCodes.Submitted _
            Then
                ' the order plex is now in a state where it can't be modified. If it's
                ' the currently selected order plex, make it not so.
                If op Is mSelectedOrderPlex Then
                    invertEntryColors mSelectedOrderPlexGridRow
                    mSelectedOrderPlexGridRow = -1
                    Set mSelectedOrderPlex = Nothing
                    ModifyOrderButton.Enabled = False
                End If
            End If
        End Select
    End With
ElseIf TypeOf ev.source Is TradeBuild.PositionManager Then
    Dim pmChangeType As TradeBuild.PositionManagerChangeTypes
    Dim pm As TradeBuild.PositionManager
    Dim pmIndex As Long
    
    Set pm = ev.source
    
    pmIndex = findPositionManagerTableIndex(pm)
    
    With mPositionManagerGridMappingTable(pmIndex)
    
        pmChangeType = ev.ChangeType
        
        Select Case pmChangeType
        Case PositionManagerChangeTypes.PositionSizeChanged
            OrderPlexGrid.TextMatrix(.gridIndex, OPGridPositionColumns.size) = pm.positionSize
        End Select
    End With
End If
End Sub

'================================================================================
' ProfitListener Interface Members
'================================================================================

Private Sub ProfitListener_profitAmount(ev As TradeBuild.ProfitEvent)
If TypeOf ev.source Is TradeBuild.OrderPlex Then
    Dim opProfitType As TradeBuild.OrderPlexProfitTypes
    Dim op As TradeBuild.OrderPlex
    Dim opIndex As Long
    
    Set op = ev.source
    
    opIndex = findOrderPlexTableIndex(op)
    
    opProfitType = ev.profitType
    
    Select Case opProfitType
    Case TradeBuild.OrderPlexProfitTypes.profit
        OrderPlexGrid.TextMatrix(mOrderPlexGridMappingTable(opIndex).gridIndex, _
                                OPGridOrderPlexColumns.profit) = Format(ev.profitAmount, "0.00")
    Case TradeBuild.OrderPlexProfitTypes.MaxProfit
        OrderPlexGrid.TextMatrix(mOrderPlexGridMappingTable(opIndex).gridIndex, _
                                OPGridOrderPlexColumns.MaxProfit) = IIf(ev.profitAmount <> 0, Format(ev.profitAmount, "0.00"), "")
    Case TradeBuild.OrderPlexProfitTypes.drawdown
        OrderPlexGrid.TextMatrix(mOrderPlexGridMappingTable(opIndex).gridIndex, _
                                OPGridOrderPlexColumns.drawdown) = IIf(ev.profitAmount <> 0, Format(ev.profitAmount, "0.00"), "")
    End Select

ElseIf TypeOf ev.source Is TradeBuild.PositionManager Then
    Dim pmProfitType As TradeBuild.PositionProfitTypes
    Dim pm As TradeBuild.PositionManager
    Dim pmIndex As Long
    
    Set pm = ev.source
    
    pmIndex = findPositionManagerTableIndex(pm)
    
    pmProfitType = ev.profitType
    
    Select Case pmProfitType
    Case TradeBuild.PositionProfitTypes.SessionProfit
        OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(pmIndex).gridIndex, _
                                OPGridPositionColumns.profit) = Format(ev.profitAmount, "0.00")
    Case TradeBuild.PositionProfitTypes.SessionMaxProfit
        OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(pmIndex).gridIndex, _
                                OPGridPositionColumns.MaxProfit) = IIf(ev.profitAmount <> 0, Format(ev.profitAmount, "0.00"), "")
    Case TradeBuild.PositionProfitTypes.SessionDrawdown
        OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(pmIndex).gridIndex, _
                                OPGridPositionColumns.drawdown) = IIf(ev.profitAmount <> 0, Format(ev.profitAmount, "0.00"), "")
    Case TradeBuild.PositionProfitTypes.tradeProfit
    Case TradeBuild.PositionProfitTypes.TradeMaxProfit
    Case TradeBuild.PositionProfitTypes.tradeDrawdown
    End Select
End If
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As TradeBuild.QuoteEvent)
On Error GoTo err
mTimestamp = mTicker.timestamp
AskText = ev.priceString
AskSizeText = ev.size

Exit Sub
err:
handleFatalError err.Number, err.Description, "QuoteListener_ask"
End Sub

Private Sub QuoteListener_bid(ev As TradeBuild.QuoteEvent)
On Error GoTo err
mTimestamp = mTicker.timestamp
BidText = ev.priceString
BidSizeText = ev.size

Exit Sub
err:
handleFatalError err.Number, err.Description, "QuoteListener_bid"
End Sub

Private Sub QuoteListener_high(ev As TradeBuild.QuoteEvent)
On Error GoTo err
mTimestamp = mTicker.timestamp
HighText = ev.priceString

Exit Sub
err:
handleFatalError err.Number, err.Description, "QuoteListener_high"
End Sub

Private Sub QuoteListener_Low(ev As TradeBuild.QuoteEvent)
On Error GoTo err
mTimestamp = mTicker.timestamp
LowText = ev.priceString

Exit Sub
err:
handleFatalError err.Number, err.Description, "QuoteListener_low"
End Sub

Private Sub QuoteListener_openInterest(ev As TradeBuild.QuoteEvent)

End Sub

Private Sub QuoteListener_previousClose(ev As TradeBuild.QuoteEvent)
On Error GoTo err
mTimestamp = mTicker.timestamp
CloseText = ev.priceString

Exit Sub
err:
handleFatalError err.Number, err.Description, "QuoteListener_previousClose"
End Sub

Private Sub QuoteListener_trade(ev As TradeBuild.QuoteEvent)
On Error GoTo err
mTimestamp = mTicker.timestamp
LastText = ev.priceString
LastSizeText = ev.size

Exit Sub
err:
handleFatalError err.Number, err.Description, "QuoteListener_trade"
End Sub

Private Sub QuoteListener_volume(ev As TradeBuild.QuoteEvent)
On Error GoTo err
mTimestamp = mTicker.timestamp
VolumeText = ev.size

Exit Sub
err:
handleFatalError err.Number, err.Description, "QuoteListener_volume"
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub CancelOrderButton_Click()
Dim rowdata As Long
Dim index As Long

rowdata = OrderPlexGrid.rowdata(mSelectedOrderPlexGridRow)
index = rowdata - RowDataOrderPlexBase

mOrderPlexGridMappingTable(index).op.cancel True

invertEntryColors mSelectedOrderPlexGridRow

CancelOrderButton.Enabled = False
ModifyOrderButton.Enabled = False
End Sub

Private Sub ChartButton_Click()
createChart mTicker
GridChartButton.Enabled = False
ChartButton.Enabled = False
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

Private Sub ClientIDText_Change()
checkOKToConnect
End Sub

Private Sub ConnectButton_Click()
Dim sp As Object

Set mContractCol = New Collection
Set mOrdersCol = New Collection
setupOrderPlexGrid
ExecutionsList.ListItems.Clear

mTradeBuildAPI.simulateOrders = (SimulateOrdersCheck = vbChecked)
SimulateOrdersCheck.Enabled = False
mTradeBuildAPI.Connect IIf(ServerText = "", "127.0.0.1", ServerText), PortText, ClientIDText
writeStatusMessage "Attempting connection to " & _
                    IIf(ServerText = "", "local server", ServerText) & _
                    "; port=" & PortText & _
                    "; client id=" & ClientIDText

If RealtimeDataSourceCombo.Text = "TWS" Then
    ' set up TWS realtime data service provider
    Set mRealtimeServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.RealtimeDataServiceProvider"))
    mRealtimeServiceProvider.Server = DataSourceServerText
    mRealtimeServiceProvider.Port = DataSourcePortText
    mRealtimeServiceProvider.clientID = DataSourceClientIdText
    mRealtimeServiceProvider.providerKey = "IB"
    mRealtimeServiceProvider.keepConnection = True
ElseIf RealtimeDataSourceCombo.Text = "QuoteTracker" Then
    ' set up QT realtime data service provider
    Set mRealtimeServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("QTSP.QTRealtimeDataServiceProvider"))
    mRealtimeServiceProvider.QTServer = DataSourceServerText
    mRealtimeServiceProvider.QTPort = DataSourcePortText
    mRealtimeServiceProvider.password = DataSourcePasswordText
    mRealtimeServiceProvider.providerKey = "QTIB"
    mRealtimeServiceProvider.keepConnection = True
End If

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
GridMarketDepthButton.Enabled = False
MarketDepthButton.Enabled = False
ChartButton.Enabled = False

LocalSymbolText.Enabled = False
SymbolText.Enabled = False
TypeCombo.Enabled = False
ExpiryText.Enabled = False
ExchangeText.Enabled = False
CurrencyText.Enabled = False
StrikePriceText.Enabled = False
RightCombo.Enabled = False
RecordCheck.Enabled = False
MarketDepthCheck.Enabled = False

If Not mTicker Is Nothing Then
    mTicker.removeQuoteListener Me
    Set mTicker = Nothing
End If

setupOrderPlexGrid
ExecutionsList.ListItems.Clear

If Not mOrderForm Is Nothing Then Unload mOrderForm
Set mOrderForm = Nothing

removeServiceProviders
mTradeBuildAPI.disconnect
ConnectButton.SetFocus
End Sub

Private Sub ExchangeText_Change()
checkOkToStartTicker
End Sub

Private Sub ExecutionsList_ColumnClick(ByVal columnHeader As columnHeader)
If ExecutionsList.SortKey = columnHeader.index - 1 Then
    ExecutionsList.SortOrder = 1 - ExecutionsList.SortOrder
Else
    ExecutionsList.SortKey = columnHeader.index - 1
    ExecutionsList.SortOrder = lvwAscending
End If
End Sub

Private Sub ExpiryText_Change()
checkOkToStartTicker
End Sub

Private Sub GridChartButton_Click()
Dim lTicker As Ticker
Dim bookmark As Variant

For Each bookmark In TickerGrid.SelBookmarks
    TickerGrid.bookmark = bookmark           ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set lTicker = mTickers(TickerGrid.Text)
    createChart lTicker
Next

GridChartButton.Enabled = False
ChartButton.Enabled = False
End Sub

Private Sub GridMarketDepthButton_Click()
Dim Ticker As Ticker
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

Private Sub HistDataSourceCombo_Click()
If HistDataSourceCombo.Text = "TWS" Then
    HistPortText = 7496
    HistClientIdText.Enabled = True
    HistPasswordText.Enabled = False
ElseIf HistDataSourceCombo.Text = "QuoteTracker" Then
    HistPortText = 16240
    HistClientIdText.Enabled = False
    HistPasswordText.Enabled = True
End If
End Sub

Private Sub LocalSymbolText_Change()
checkOkToStartTicker
End Sub

Private Sub MarketDepthButton_Click()
showMarketDepthForm mTicker
MarketDepthButton.Enabled = False
GridMarketDepthButton.Enabled = False
End Sub

Private Sub ModifyOrderButton_Click()
Dim rowdata As Long
Dim index As Long

rowdata = OrderPlexGrid.rowdata(mSelectedOrderPlexGridRow)
index = rowdata - RowDataOrderPlexBase

If mOrderForm Is Nothing Then Set mOrderForm = New OrderForm

mOrderForm.showOrderPlex mOrderPlexGridMappingTable(index).op, _
                        mSelectedOrderPlexGridRow - mOrderPlexGridMappingTable(index).gridIndex
mOrderForm.Show vbModeless

End Sub

Private Sub OrderButton_Click()
If mTicker Is Nothing Then
    MsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
    Exit Sub
End If
If mOrderForm Is Nothing Then Set mOrderForm = New OrderForm
mOrderForm.ordersAreSimulated = mTradeBuildAPI.simulateOrders
mOrderForm.Show vbModeless
mOrderForm.Ticker = mTicker
End Sub

Private Sub OrderPlexGrid_Click()
Dim row As Long
Dim rowdata As Long
Dim op As TradeBuild.OrderPlex
Dim index As Long
Dim orderIndex As Long

row = OrderPlexGrid.row

If OrderPlexGrid.MouseCol = OPGridColumns.symbol Then Exit Sub

If OrderPlexGrid.MouseCol = OPGridColumns.ExpandIndicator Then
    expandOrContract
Else

    invertEntryColors mSelectedOrderPlexGridRow
    
    mSelectedOrderPlexGridRow = -1
    CancelOrderButton.Enabled = False
    ModifyOrderButton.Enabled = False
    
    OrderPlexGrid.row = row
    rowdata = OrderPlexGrid.rowdata(row)
    If rowdata < RowDataPositionManagerBase And _
        rowdata >= RowDataOrderPlexBase _
    Then
        index = rowdata - RowDataOrderPlexBase
        Set op = mOrderPlexGridMappingTable(index).op
        If op.State = OrderPlexStateCodes.Created Or _
            op.State = OrderPlexStateCodes.Submitted _
        Then
            
            mSelectedOrderPlexGridRow = row
            Set mSelectedOrderPlex = op
            invertEntryColors mSelectedOrderPlexGridRow
            
            CancelOrderButton.Enabled = True
            ModifyOrderButton.Enabled = True
            
            orderIndex = mSelectedOrderPlexGridRow - mOrderPlexGridMappingTable(index).gridIndex
            If orderIndex = 0 Then Exit Sub
            
            Set mSelectedOrder = op.Order(orderIndex)
            If mSelectedOrder.isModifiable Then
                If (OrderPlexGrid.MouseCol = OPGridOrderColumns.price And _
                        mSelectedOrder.isAttributeModifiable(OrderAttributeIds.limitPrice)) Or _
                    (OrderPlexGrid.MouseCol = OPGridOrderColumns.auxPrice And _
                        mSelectedOrder.isAttributeModifiable(OrderAttributeIds.triggerPrice)) Or _
                    (OrderPlexGrid.MouseCol = OPGridOrderColumns.quantity And _
                    mSelectedOrder.isAttributeModifiable(OrderAttributeIds.quantity)) _
                Then
                    OrderPlexGrid.col = OrderPlexGrid.MouseCol
                    EditText.Move OrderPlexGrid.Left + OrderPlexGrid.CellLeft + 8, _
                                OrderPlexGrid.Top + OrderPlexGrid.CellTop + 8, _
                                OrderPlexGrid.CellWidth - 16, _
                                OrderPlexGrid.CellHeight - 16
                    EditText.Text = OrderPlexGrid.Text
                    EditText.SelStart = 0
                    EditText.SelLength = Len(EditText.Text)
                    EditText.Visible = True
                    EditText.SetFocus
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub OrderPlexGrid_LeaveCell()
Dim orderNumber As Long
Dim price As Double

If Not EditText.Visible Then Exit Sub

orderNumber = mSelectedOrderPlexGridRow - mOrderPlexGridMappingTable(OrderPlexGrid.rowdata(OrderPlexGrid.row) - RowDataOrderPlexBase).gridIndex
If OrderPlexGrid.col = OPGridOrderColumns.price Then
    If mSelectedOrderPlex.Contract.parsePrice(EditText.Text, price) Then
        mSelectedOrderPlex.newOrderPrice(orderNumber) = price
    End If
ElseIf OrderPlexGrid.col = OPGridOrderColumns.auxPrice Then
    If mSelectedOrderPlex.Contract.parsePrice(EditText.Text, price) Then
        mSelectedOrderPlex.newOrderTriggerPrice(orderNumber) = price
    End If
ElseIf OrderPlexGrid.col = OPGridOrderColumns.quantity Then
    If IsNumeric(EditText.Text) Then
        mSelectedOrderPlex.newQuantity = EditText.Text
    End If
End If
    
If mSelectedOrderPlex.dirty Then mSelectedOrderPlex.Update

EditText.Visible = False
End Sub

Private Sub OrderPlexGrid_Scroll()
If EditText.Visible Then
    EditText.Move OrderPlexGrid.Left + OrderPlexGrid.CellLeft + 8, _
                OrderPlexGrid.Top + OrderPlexGrid.CellTop + 8, _
                OrderPlexGrid.CellWidth - 16, _
                OrderPlexGrid.CellHeight - 16
End If
End Sub

Private Sub PauseReplayButton_Click()
PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
writeStatusMessage "Tickfile replay paused"
mTickfileManager.PauseReplay
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
OrderButton.Enabled = True

setupTWSContractServiceProvider

mTradeBuildAPI.simulateOrders = True

If Not mTicker Is Nothing Then
    writeStatusMessage "Tickfile replay resumed"
Else
    writeStatusMessage "Tickfile replay started"
    mTickfileManager.ReplayProgressEventIntervalMillisecs = 250
End If
mTickfileManager.replaySpeed = ReplaySpeedCombo.itemData(ReplaySpeedCombo.ListIndex)

mTickfileManager.StartReplay
End Sub

Private Sub PortText_Change()
checkOKToConnect
End Sub

Private Sub RealtimeDataSourceCombo_Click()
If RealtimeDataSourceCombo.Text = "TWS" Then
    DataSourcePortText = 7496
    DataSourceClientIdText.Enabled = True
    DataSourcePasswordText.Enabled = False
ElseIf RealtimeDataSourceCombo.Text = "QuoteTracker" Then
    DataSourcePortText = 16240
    DataSourceClientIdText.Enabled = False
    DataSourcePasswordText.Enabled = True
End If
End Sub

Private Sub RecordCheck_Click()
If RecordCheck = vbChecked Then
    MarketDepthCheck.Enabled = True
Else
    MarketDepthCheck.Enabled = False
End If
End Sub

Private Sub ReplaySpeedCombo_Click()
If Not mTickfileManager Is Nothing Then
    mTickfileManager.replaySpeed = ReplaySpeedCombo.itemData(ReplaySpeedCombo.ListIndex)
End If
End Sub

Private Sub RightCombo_Click()
checkOkToStartTicker
End Sub

Private Sub SelectTickfilesButton_Click()
Set mTickfileManager = mTickers.createTickFileManager
mTickfileManager.ShowTickfileSelectionDialogue
End Sub

Private Sub SkipReplayButton_Click()
writeStatusMessage "Tickfile skipped"
clearTickerAppData mTicker
clearTickerFields
mTickfileManager.SkipTickfile
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
Dim lContractSpecifier As ContractSpecifier

Set lContractSpecifier = mTradeBuildAPI.newContractSpecifier( _
                                LocalSymbolText, _
                                SymbolText, _
                                ExchangeText, _
                                secTypeFromString(TypeCombo), _
                                CurrencyText, _
                                ExpiryText, _
                                IIf(StrikePriceText = "", 0, StrikePriceText), _
                                optionRightFromString(RightCombo))

StartTickerButton.Enabled = False

setupTWSContractServiceProvider

Set lTicker = createTicker
lTicker.DOMEventsRequired = DOMEvents.DOMNoEvents
lTicker.writeToTickfile = (RecordCheck = vbChecked)
lTicker.includeMarketDepthInTickfile = (RecordCheck = vbChecked And MarketDepthCheck = vbChecked)
lTicker.StartTicker lContractSpecifier

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

For Each bookmark In TickerGrid.SelBookmarks
    TickerGrid.bookmark = bookmark           ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set Ticker = mTickers(TickerGrid.Text)
    Ticker.StopTicker
    Ticker.PositionManager.removeProfitListener Me
    Ticker.PositionManager.removeChangeListener Me
Next

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
Dim tickerAppData As TickerApplicationData

If TickerGrid.SelStartCol <> -1 Then
    StopTickerButton.Enabled = False
    GridChartButton.Enabled = False
    GridMarketDepthButton.Enabled = False
Else
    ' the user has clicked on the record selectors
    StopTickerButton.Enabled = True
    GridChartButton.Enabled = True
    
    If TickerGrid.SelBookmarks.Count = 1 Then
        
        TickerGrid.col = 0                       ' make the cell containing the key current
        
        If Not mTicker Is Nothing Then mTicker.removeQuoteListener Me
        Set mTicker = mTickers(TickerGrid.Text)
        mTicker.addQuoteListener Me
        Set tickerAppData = mTicker.ApplicationData
        
        If tickerAppData.MarketDepthForm Is Nothing Then
            MarketDepthButton.Enabled = True
            GridMarketDepthButton.Enabled = True
        Else
            MarketDepthButton.Enabled = False
            GridMarketDepthButton.Enabled = False
        End If
        If tickerAppData.chartform Is Nothing Then
            ChartButton.Enabled = True
            GridChartButton.Enabled = True
        Else
            ChartButton.Enabled = False
            GridChartButton.Enabled = False
        End If
        
        Set mCurrentContract = mTicker.Contract
        
        NameText = mCurrentContract.specifier.localSymbol
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
    Else
        MarketDepthButton.Enabled = False
        GridMarketDepthButton.Enabled = False
        ChartButton.Enabled = False
        GridChartButton.Enabled = False
    End If
End If

End Sub

Private Sub TypeCombo_Click()

Select Case secTypeFromString(TypeCombo)
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
    writeStatusMessage "BAG type is not implemented"
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
End Select

checkOkToStartTicker
End Sub

'================================================================================
' mOrderForm Event Handlers
'================================================================================

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_Application(ByVal timestamp As Date, ByVal data As Variant)
Dim Ticker As Ticker
Dim eventCode As ApplicationEventCodes

' this fires when the market depth form or the chart from for this ticker are
' unloaded. This may be either because the user closed the form, or because the
' user stopped the ticker.

If mTicker.State = TickerStateCodes.Dead Then Exit Sub

eventCode = CLng(data)

Select Case eventCode
Case MarketDepthFormClosed
    MarketDepthButton.Enabled = True
Case ChartFormClosed
    ChartButton.Enabled = True
End Select

' if the current selection in the ticker grid is this ticker, then enable
' the GridMarketDepthButton or GridChartButton
If TickerGrid.SelBookmarks.Count = 1 Then
    TickerGrid.bookmark = TickerGrid.SelBookmarks(0)    ' select the row
    TickerGrid.col = 0                       ' make the cell containing the key current
    Set Ticker = mTickers(TickerGrid.Text)
    If Ticker Is mTicker Then
        Select Case eventCode
        Case MarketDepthFormClosed
            MarketDepthButton.Enabled = True
            GridMarketDepthButton.Enabled = True
        Case ChartFormClosed
            GridChartButton.Enabled = True
        End Select
    End If
End If

End Sub

Private Sub mTicker_ContractInvalid( _
                ByVal ContractSpecifier As TradeBuild.ContractSpecifier, _
                ByVal reason As String)
On Error GoTo err
writeStatusMessage "Invalid contract details (" & reason & "):" & vbCrLf & _
                    Replace(ContractSpecifier.ToString, vbCrLf, "; ")
StartTickerButton.Enabled = True

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_ContractInvalid"
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

Private Sub mTicker_OutputTickfileCreated(ByVal timestamp As Date, _
                            ByVal Filename As String)
writeStatusMessage "Created output tickfile: " & Filename
End Sub

Private Sub mTicker_volume(ByVal timestamp As Date, _
                            ByVal size As Long)
On Error GoTo err
mTimestamp = timestamp
VolumeText = size

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_volume"
End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_ContractAmbiguous( _
                ByVal pTicker As TradeBuild.Ticker, _
                ByVal contracts As TradeBuild.contracts)
writeStatusMessage "Ambiguous contract details:" & vbCrLf & _
                    Replace(contracts.ContractSpecifier.ToString, vbCrLf, "; ")
StartTickerButton.Enabled = True
End Sub

Private Sub mTickers_contractInvalid(ByVal pTicker As Ticker, _
                ByVal contractSpec As ContractSpecifier, _
                ByVal reason As String)
writeStatusMessage "Invalid contract details (" & reason & "):" & vbCrLf & _
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

tickerAppData.MarketDepthForm.setDOMCell mTimestamp, DOMSides.DOMLast, CDbl(LastText), LastSizeText
tickerAppData.MarketDepthForm.setDOMCell mTimestamp, DOMSides.DOMAsk, CDbl(AskText), AskSizeText
tickerAppData.MarketDepthForm.setDOMCell mTimestamp, DOMSides.DOMBid, CDbl(BidText), BidSizeText

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
                                           ByVal contractSpec As ContractSpecifier)
writeStatusMessage "A ticker is already running for contract: " & _
                    Replace(contractSpec.ToString, vbCrLf, "; ")
StartTickerButton.Enabled = True
End Sub

Private Sub mTickers_MarketDepthNotAvailable( _
                            ByVal pTicker As TradeBuild.Ticker, _
                            ByVal reason As String)
Dim tickerAppData As TickerApplicationData

On Error GoTo err

writeStatusMessage "No market depth for " & _
                    pTicker.Contract.specifier.localSymbol & _
                    ": " & reason

Set tickerAppData = pTicker.ApplicationData
Unload tickerAppData.MarketDepthForm
Set tickerAppData.MarketDepthForm = Nothing

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickers_MarketDepthNotAvailable"
End Sub

Private Sub mTickers_TickerReady( _
                ByVal pTicker As Ticker)

On Error GoTo err

If pTicker Is mTicker Then
    Set mCurrentContract = mTicker.Contract
    MarketDepthButton.Enabled = True
    ChartButton.Enabled = True

    NameText = mCurrentContract.specifier.localSymbol
    
End If

pTicker.PositionManager.addProfitListener Me
pTicker.PositionManager.addChangeListener Me

On Error Resume Next
If mContractCol.Item(pTicker.Contract.specifier.localSymbol) Is Nothing Then
    mContractCol.Add pTicker.Contract, pTicker.Contract.specifier.localSymbol
End If

On Error GoTo err

StartTickerButton.Enabled = True

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTicker_Ready"
End Sub

Private Sub mTickers_TickerRemoved(ByVal pTicker As Ticker)

' The following seems to be needed to prevent the TickerGrid_Error
' event being fired. Otherwise, disabling StopTickerButton causes the focus
' to go the the TickerGrid, which then causes an error.
If LocalSymbolText.Enabled Then
    LocalSymbolText.SetFocus
Else
    ReplaySpeedCombo.SetFocus
End If

StopTickerButton.Enabled = False
MarketDepthButton.Enabled = False
GridChartButton.Enabled = False
GridMarketDepthButton.Enabled = False

If pTicker Is mTicker Then
    clearTickerFields
    mTicker.removeQuoteListener Me
    mTicker.PositionManager.removeProfitListener Me
    mTicker.PositionManager.removeChangeListener Me
    Set mTicker = Nothing
End If

clearTickerAppData pTicker
pTicker.ApplicationData = Empty
End Sub

'================================================================================
' mTickfileManager Event Handlers
'================================================================================

Private Sub mTickfileManager_errorMessage( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As TradeBuild.ApiErrorCodes, _
                ByVal errorMsg As String)
On Error GoTo err
mTimestamp = timestamp
writeStatusMessage "Error " & errorCode & ": " & id & ": " & errorMsg

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTickfileManager_errorMessage"
End Sub

Private Sub mTickfileManager_QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As TradeBuild.Contract, _
                continueMode As TradeBuild.ReplayContinueModes)
On Error GoTo err

If tickfileIndex <> 0 Then
    clearTickerAppData mTicker
    clearTickerFields
    Set mTicker = Nothing
End If

setupOrderPlexGrid
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
handleFatalError err.Number, err.Description, "mTickfileManager_QueryReplayNextTickfile"
End Sub

Private Sub mTickfileManager_ReplayCompleted()
On Error GoTo err

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
handleFatalError err.Number, err.Description, "mTickfileManager_ReplayCompleted"
End Sub

Private Sub mTickfileManager_ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)
On Error GoTo err
mTimestamp = tickfileTimestamp
ReplayProgressBar.value = percentComplete
ReplayProgressLabel.Caption = tickfileTimestamp & _
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
mTicker.addQuoteListener Me
initialiseTicker mTicker
mTicker.DOMEventsRequired = DOMProcessedEvents
mTicker.writeToTickfile = (RewriteCheck = vbChecked)
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
    TickfileList.AddItem tickfiles(i).Filename
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
    theTime = mTicker.timestamp
Else
    theTime = GetTimestamp
End If
DateTimeText = Format(theTime, "dd/mm/yy") & vbCrLf & Format(theTime, "hh:mm:ss")
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_connected(ByVal timestamp As Date)
Dim execFilter As ExecutionFilter

On Error GoTo err

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
LocalSymbolText.Enabled = True
LocalSymbolText.SetFocus
SymbolText.Enabled = True
TypeCombo.Enabled = True
ExpiryText.Enabled = IIf(TypeCombo.Text = "" Or _
                        TypeCombo.Text = StrSecTypeFuture Or _
                        TypeCombo.Text = StrSecTypeOption Or _
                        TypeCombo.Text = StrSecTypeOptionFuture, _
                        True, _
                        False)
ExchangeText.Enabled = True
CurrencyText.Enabled = True
StrikePriceText.Enabled = IIf(TypeCombo.Text = "" Or _
                        TypeCombo.Text = StrSecTypeOption Or _
                        TypeCombo.Text = StrSecTypeOptionFuture, _
                        True, _
                        False)
RightCombo.Enabled = IIf(TypeCombo.Text = "" Or _
                        TypeCombo.Text = StrSecTypeOption Or _
                        TypeCombo.Text = StrSecTypeOptionFuture, _
                        True, _
                        False)
RecordCheck.Enabled = True
If RecordCheck = vbChecked Then MarketDepthCheck.Enabled = True

checkOkToStartTicker

Set execFilter = New ExecutionFilter
execFilter.clientID = ClientIDText
mTradeBuildAPI.RequestExecutions execFilter

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_connected"
End Sub

Private Sub mTradeBuildAPI_connectFailed( _
                ByVal timestamp As Date, _
                ByVal Description As String, _
                ByVal retrying As Boolean)
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

Private Sub mTradeBuildAPI_connectionToTWSClosed( _
                ByVal timestamp As Date, _
                ByVal reconnecting As Boolean)
On Error GoTo err

mTimestamp = timestamp

If reconnecting Then
    writeStatusMessage "Connection closed - attempting to reconnect"
    Exit Sub
End If

checkOKToConnect
DisconnectButton.Enabled = False
SimulateOrdersCheck.Enabled = True
OrderButton.Enabled = False
StartTickerButton.Enabled = False
StopTickerButton.Enabled = False
GridChartButton.Enabled = False
GridMarketDepthButton.Enabled = False
MarketDepthButton.Enabled = False
ChartButton.Enabled = False
Set mTicker = Nothing

ServerText.Enabled = True
PortText.Enabled = True
ClientIDText.Enabled = True

LocalSymbolText.Enabled = False
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

setupOrderPlexGrid
ExecutionsList.ListItems.Clear

If Not mOrderForm Is Nothing Then Unload mOrderForm
Set mOrderForm = Nothing

writeStatusMessage "Connection closed"

checkOkToStartReplay

Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_connectionToTWSClosed"
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
                        ByVal pContractSpecifier As ContractSpecifier, _
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
listItem.SubItems(ExecutionsColumns.orderId - 1) = exec.orderId
listItem.SubItems(ExecutionsColumns.price - 1) = exec.price
listItem.SubItems(ExecutionsColumns.quantity - 1) = exec.quantity
listItem.SubItems(ExecutionsColumns.symbol - 1) = pContractSpecifier.localSymbol
listItem.SubItems(ExecutionsColumns.Time - 1) = exec.Time


Exit Sub
err:
handleFatalError err.Number, err.Description, "mTradeBuildAPI_executionDetails"
End Sub

'Private Sub mTradeBuildAPI_openOrder(ByVal timestamp As Date, _
'                            ByVal pContractSpecifier As ContractSpecifier, _
'                            ByVal pOrder As Order)
'On Error GoTo err
'
'
'mTimestamp = timestamp
'openOrder pContractSpecifier, pOrder
'
'
'Exit Sub
'err:
'handleFatalError err.Number, err.Description, "mTradeBuildAPI_openOrder"
'End Sub

'Private Sub mTradeBuildAPI_orderStatus(ByVal timestamp As Date, _
'                                ByVal id As Long, _
'                                ByVal status As OrderStatuses, _
'                                ByVal filled As Long, _
'                                ByVal remaining As Long, _
'                                ByVal avgFillPrice As Double, _
'                                ByVal permId As Long, _
'                                ByVal parentId As Long, _
'                                ByVal lastFillPrice As Double, _
'                                ByVal clientId As Long)
'Dim listItem As listItem
'Dim lOrder As Order
'Dim orderKey As String
'
'On Error GoTo err
'
'mTimestamp = timestamp
'
'orderKey = "A" & CStr(id)
'
'On Error Resume Next
'Set listItem = OpenOrdersList.ListItems(orderKey)
'On Error GoTo err
'
'If listItem Is Nothing Then
'    Set listItem = OpenOrdersList.ListItems.Add(, orderKey, CStr(id))
'End If
'
'listItem.SubItems(OpenOrdersColumns.status - 1) = orderStatusToString(status)
'listItem.SubItems(OpenOrdersColumns.Quantity - 1) = remaining
'
'Set lOrder = mOrdersCol(orderKey)
'
'lOrder.status = status
'lOrder.quantityFilled = filled
'lOrder.Quantity = remaining
'lOrder.averagePrice = avgFillPrice
'lOrder.permId = permId
'lOrder.lastFillPrice = lastFillPrice
'
'Exit Sub
'err:
'handleFatalError err.Number, err.Description, "mTradeBuildAPI_orderStatus"
'End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Sub logMessage(ByVal message As String)
writeStatusMessage message
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function addEntryToOrderPlexGrid( _
                ByVal symbol As String, _
                Optional ByVal before As Boolean, _
                Optional ByVal index As Long = -1) As Long
Dim i As Long

If index < 0 Then
    For i = mFirstOrderPlexGridRowIndex To OrderPlexGrid.Rows - 1
        If (before And _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) >= symbol) Or _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = "" _
        Then
            index = i
            Exit For
        ElseIf (Not before And _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) > symbol) Or _
            OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = "" _
        Then
            index = i
            Exit For
        End If
    Next
    
    If index < 0 Then
        OrderPlexGrid.AddItem ""
        index = OrderPlexGrid.Rows - 1
    ElseIf OrderPlexGrid.TextMatrix(index, OPGridColumns.symbol) = "" Then
        OrderPlexGrid.TextMatrix(index, OPGridColumns.symbol) = symbol
    Else
        OrderPlexGrid.AddItem "", index
    End If
Else
    OrderPlexGrid.AddItem "", index
End If

OrderPlexGrid.TextMatrix(index, OPGridColumns.symbol) = symbol
If index < OrderPlexGrid.Rows - 1 Then
    ' this new entry has displaced one or more existing entries so
    ' the OrderPlexGridMappingTable and PositionManageGridMappingTable indexes
    ' need to be adjusted
    For i = 0 To mMaxOrderPlexGridMappingTableIndex
        If mOrderPlexGridMappingTable(i).gridIndex >= index Then
            mOrderPlexGridMappingTable(i).gridIndex = mOrderPlexGridMappingTable(i).gridIndex + 1
        End If
    Next
    For i = 0 To mMaxPositionManagerGridMappingTableIndex
        If mPositionManagerGridMappingTable(i).gridIndex >= index Then
            mPositionManagerGridMappingTable(i).gridIndex = mPositionManagerGridMappingTable(i).gridIndex + 1
        End If
    Next
End If

addEntryToOrderPlexGrid = index
End Function

Private Function addOrderPlexEntryToOrderPlexGrid( _
                ByVal symbol As String, _
                ByVal orderPlexTableIndex As Long) As Long
Dim index As Long

index = addEntryToOrderPlexGrid(symbol, False)

OrderPlexGrid.rowdata(index) = orderPlexTableIndex + RowDataOrderPlexBase

OrderPlexGrid.row = index
OrderPlexGrid.col = OPGridColumns.ExpandIndicator
OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture

OrderPlexGrid.col = OPGridOrderPlexColumns.profit
OrderPlexGrid.CellBackColor = &HC0C0C0
OrderPlexGrid.CellForeColor = vbWhite

OrderPlexGrid.col = OPGridOrderPlexColumns.MaxProfit
OrderPlexGrid.CellBackColor = &HC0C0C0
OrderPlexGrid.CellForeColor = vbWhite

OrderPlexGrid.col = OPGridOrderPlexColumns.drawdown
OrderPlexGrid.CellBackColor = &HC0C0C0
OrderPlexGrid.CellForeColor = vbWhite

addOrderPlexEntryToOrderPlexGrid = index
End Function
                
Private Sub addOrderEntryToOrderPlexGrid( _
                ByVal index As Long, _
                ByVal symbol As String, _
                ByVal pOrder As TradeBuild.Order, _
                ByVal orderPlexTableIndex As Long, _
                ByVal typeInPlex As String)


index = addEntryToOrderPlexGrid(symbol, False, index)

OrderPlexGrid.rowdata(index) = orderPlexTableIndex + RowDataOrderPlexBase

OrderPlexGrid.TextMatrix(index, OPGridOrderColumns.typeInPlex) = typeInPlex

displayOrderValuesInOrderPlexGrid index, pOrder

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
If LocalSymbolText <> "" Or _
    SymbolText <> "" _
Then
'If SymbolText <> "" And _
'    TypeCombo.Text <> "" And _
'    IIf(TypeCombo.Text = StrSecTypeFuture Or _
'        TypeCombo.Text = StrSecTypeOption Or _
'        TypeCombo.Text = StrSecTypeOptionFuture, _
'        ExpiryText <> "", _
'        True) And _
'    IIf(TypeCombo.Text = StrSecTypeOption Or _
'        TypeCombo.Text = StrSecTypeOptionFuture, _
'        StrikePriceText <> "", _
'        True) And _
'    IIf(TypeCombo.Text = StrSecTypeOption Or _
'        TypeCombo.Text = StrSecTypeOptionFuture, _
'        RightCombo <> "", _
'        True) And _
'    ExchangeText <> "" _
'Then
    StartTickerButton.Enabled = True
Else
    StartTickerButton.Enabled = False
End If
End Sub

Private Sub clearTickerAppData(ByVal pTicker As Ticker)
Dim tickerAppData As TickerApplicationData

If pTicker Is Nothing Then Exit Sub

Set tickerAppData = pTicker.ApplicationData

If tickerAppData Is Nothing Then Exit Sub

If Not tickerAppData.MarketDepthForm Is Nothing Then
    Unload tickerAppData.MarketDepthForm
    Set tickerAppData.MarketDepthForm = Nothing
End If
If Not tickerAppData.chartform Is Nothing Then
    Unload tickerAppData.chartform
    Set tickerAppData.chartform = Nothing
End If
'pTicker.ApplicationData = Nothing
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

Private Function contractOrderPlexEntry( _
                ByVal index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Dim lIndex As Long

With mOrderPlexGridMappingTable(index)
    If .entryGridOffset >= 0 Then
        lIndex = .gridIndex + .entryGridOffset
        OrderPlexGrid.RowHeight(lIndex) = 0
    End If
    If .stopGridOffset >= 0 Then
        lIndex = .gridIndex + .stopGridOffset
        OrderPlexGrid.RowHeight(lIndex) = 0
    End If
    If .targetGridOffset >= 0 Then
        lIndex = .gridIndex + .targetGridOffset
        OrderPlexGrid.RowHeight(lIndex) = 0
    End If
    If .closeoutGridOffset >= 0 Then
        lIndex = .gridIndex + .closeoutGridOffset
        OrderPlexGrid.RowHeight(lIndex) = 0
    End If
    
    If Not preserveCurrentExpandedState Then
        .isExpanded = False
        OrderPlexGrid.row = .gridIndex
        OrderPlexGrid.col = OPGridColumns.ExpandIndicator
        OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
        Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Expand").Picture
    End If
End With

contractOrderPlexEntry = lIndex
End Function

Private Sub contractPositionManagerEntry(ByVal index As Long)
Dim i As Long
Dim symbol As String
Dim lOpEntryIndex As Long

mPositionManagerGridMappingTable(index).isExpanded = False
OrderPlexGrid.row = mPositionManagerGridMappingTable(index).gridIndex
OrderPlexGrid.col = OPGridColumns.ExpandIndicator
OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Expand").Picture

symbol = OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(index).gridIndex, OPGridColumns.symbol)
i = mPositionManagerGridMappingTable(index).gridIndex + 1
Do While OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = symbol
    OrderPlexGrid.RowHeight(i) = 0
    lOpEntryIndex = OrderPlexGrid.rowdata(i) - RowDataOrderPlexBase
    i = contractOrderPlexEntry(lOpEntryIndex, True) + 1
Loop
End Sub

Private Sub createChart(ByVal pTicker As Ticker)
Dim chartform As fChart1

If Not pTicker.ApplicationData.chartform Is Nothing Then Exit Sub

If mHistDataServiceProvider Is Nothing Then
    ' set up TWS historical data service provider
    If HistDataSourceCombo.Text = "TradeBuild" Then
        Set mHistDataServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("TBInfoBase.HistDataServiceProvider"))
    ElseIf HistDataSourceCombo.Text = "TWS" Then
        Set mHistDataServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.HistDataServiceProvider"))
        mHistDataServiceProvider.Server = HistServerText
        mHistDataServiceProvider.Port = HistPortText
        mHistDataServiceProvider.clientID = HistClientIdText
        mHistDataServiceProvider.keepConnection = True
    ElseIf HistDataSourceCombo.Text = "QuoteTracker" Then
        Set mHistDataServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("QTSP.QTHistDataServiceProvider"))
        mHistDataServiceProvider.QTServer = HistServerText
        mHistDataServiceProvider.QTPort = HistPortText
        mHistDataServiceProvider.password = HistPasswordText
        mHistDataServiceProvider.providerKey = "QTIB"
        mHistDataServiceProvider.keepConnection = True
    End If
End If

Set chartform = New fChart1
chartform.minimumTicksHeight = 40
chartform.InitialNumberOfBars = 500
chartform.barLength = 1
chartform.Ticker = pTicker
chartform.Visible = True
Set pTicker.ApplicationData.chartform = chartform
End Sub

Private Function createTicker() As Ticker
Set createTicker = mTickers.Add(Format(CLng(1000000000 * Rnd)), "0")
initialiseTicker createTicker
End Function

Private Sub displayOrderValuesInOrderPlexGrid( _
                ByVal gridIndex As Long, _
                ByVal pOrder As Order)
Dim lTicker As TradeBuild.Ticker

Set lTicker = pOrder.Ticker

OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Action) = orderActionToString(pOrder.Action)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.auxPrice) = lTicker.formatPrice(pOrder.triggerPrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.averagePrice) = lTicker.formatPrice(pOrder.averagePrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.id) = pOrder.id
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.lastFillPrice) = lTicker.formatPrice(pOrder.lastFillPrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.LastFillTime) = IIf(pOrder.fillTime <> 0, pOrder.fillTime, "")
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.orderType) = orderTypeToString(pOrder.orderType)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.price) = lTicker.formatPrice(pOrder.limitPrice, True)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.quantity) = pOrder.quantity
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.size) = IIf(pOrder.quantityFilled <> 0, pOrder.quantityFilled, 0)
OrderPlexGrid.TextMatrix(gridIndex, OPGridOrderColumns.Status) = orderStatusToString(pOrder.Status)
End Sub

Private Sub expandOrContract()
Dim rowdata As Long
Dim index As Long
Dim expanded As Boolean

rowdata = OrderPlexGrid.rowdata(OrderPlexGrid.MouseRow)
If rowdata >= RowDataPositionManagerBase Then
    index = rowdata - RowDataPositionManagerBase
    expanded = mPositionManagerGridMappingTable(index).isExpanded
    If expanded Then
        contractPositionManagerEntry index
    Else
        expandPositionManagerEntry index
    End If
ElseIf rowdata >= RowDataOrderPlexBase Then
    index = rowdata - RowDataOrderPlexBase
    expanded = mOrderPlexGridMappingTable(index).isExpanded
    If OrderPlexGrid.row <> mOrderPlexGridMappingTable(index).gridIndex Then
        ' clicked on an order entry
        Exit Sub
    End If
    If expanded Then
        contractOrderPlexEntry index
    Else
        expandOrderPlexEntry index
    End If
Else
    Exit Sub
End If
End Sub

Private Function expandOrderPlexEntry( _
                ByVal index As Long, _
                Optional ByVal preserveCurrentExpandedState As Boolean) As Long
Dim lIndex As Long


With mOrderPlexGridMappingTable(index)
    
    If .entryGridOffset >= 0 Then
        lIndex = .gridIndex + .entryGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.RowHeight(lIndex) = -1
    End If
    If .stopGridOffset >= 0 Then
        lIndex = .gridIndex + .stopGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.RowHeight(lIndex) = -1
    End If
    If .targetGridOffset >= 0 Then
        lIndex = .gridIndex + .targetGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.RowHeight(lIndex) = -1
    End If
    If .closeoutGridOffset >= 0 Then
        lIndex = .gridIndex + .closeoutGridOffset
        If Not preserveCurrentExpandedState Or .isExpanded Then OrderPlexGrid.RowHeight(lIndex) = -1
    End If
    
    If Not preserveCurrentExpandedState Then
        .isExpanded = True
        OrderPlexGrid.row = .gridIndex
        OrderPlexGrid.col = OPGridColumns.ExpandIndicator
        OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
        Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture
    End If
End With

expandOrderPlexEntry = lIndex
End Function

Private Sub expandPositionManagerEntry(ByVal index As Long)
Dim i As Long
Dim symbol As String
Dim lOpEntryIndex As Long

mPositionManagerGridMappingTable(index).isExpanded = True
OrderPlexGrid.row = mPositionManagerGridMappingTable(index).gridIndex
OrderPlexGrid.col = OPGridColumns.ExpandIndicator
OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture

symbol = OrderPlexGrid.TextMatrix(mPositionManagerGridMappingTable(index).gridIndex, OPGridColumns.symbol)
i = mPositionManagerGridMappingTable(index).gridIndex + 1
Do While OrderPlexGrid.TextMatrix(i, OPGridColumns.symbol) = symbol
    OrderPlexGrid.RowHeight(i) = -1
    lOpEntryIndex = OrderPlexGrid.rowdata(i) - RowDataOrderPlexBase
    i = expandOrderPlexEntry(lOpEntryIndex, True) + 1
Loop
End Sub

Private Function findOrderPlexTableIndex(ByVal op As TradeBuild.OrderPlex) As Long
Dim opIndex As Long
Dim lOrder As TradeBuild.Order
Dim symbol As String

' first make sure the relevant PositionManager entry is set up
findPositionManagerTableIndex op.Ticker.PositionManager

symbol = op.Contract.specifier.localSymbol
opIndex = op.indexApplication
If opIndex > UBound(mOrderPlexGridMappingTable) Then
    ReDim Preserve mOrderPlexGridMappingTable(UBound(mOrderPlexGridMappingTable) + 50) As OrderPlexGridMappingEntry
End If
If opIndex > mMaxOrderPlexGridMappingTableIndex Then mMaxOrderPlexGridMappingTableIndex = opIndex

With mOrderPlexGridMappingTable(opIndex)
    If .op Is Nothing Then
        
        .isExpanded = True
        .entryGridOffset = -1
        .stopGridOffset = -1
        .targetGridOffset = -1
        .closeoutGridOffset = -1
        
        Set .op = op
        .gridIndex = addOrderPlexEntryToOrderPlexGrid(op.Contract.specifier.localSymbol, opIndex)
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridOrderPlexColumns.creationTime) = op.creationTime
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridOrderPlexColumns.currencyCode) = op.Contract.specifier.currencyCode
        
        Set lOrder = op.entryOrder
        If Not lOrder Is Nothing Then
            .entryGridOffset = 1
            addOrderEntryToOrderPlexGrid .gridIndex + .entryGridOffset, _
                                    symbol, _
                                    lOrder, _
                                    opIndex, _
                                    "Entry"
        End If
        
        Set lOrder = op.stopOrder
        If Not lOrder Is Nothing Then
            If .entryGridOffset >= 0 Then
                .stopGridOffset = .entryGridOffset + 1
            Else
                .stopGridOffset = 1
            End If
            addOrderEntryToOrderPlexGrid .gridIndex + .stopGridOffset, _
                                    symbol, _
                                    lOrder, _
                                    opIndex, _
                                    "Stop"
        End If
        
        Set lOrder = op.targetOrder
        If Not lOrder Is Nothing Then
            If .stopGridOffset >= 0 Then
                .targetGridOffset = .stopGridOffset + 1
            ElseIf .entryGridOffset >= 0 Then
                .targetGridOffset = .entryGridOffset + 1
            Else
                .targetGridOffset = 1
            End If
            addOrderEntryToOrderPlexGrid .gridIndex + .targetGridOffset, _
                                    symbol, _
                                    lOrder, _
                                    opIndex, _
                                    "Target"
        End If
    End If
End With
findOrderPlexTableIndex = opIndex
End Function

Private Function findPositionManagerTableIndex(ByVal pm As TradeBuild.PositionManager) As Long
Dim pmIndex As Long
Dim symbol As String

symbol = pm.Ticker.Contract.specifier.localSymbol
pmIndex = pm.indexApplication
If pmIndex > UBound(mPositionManagerGridMappingTable) Then
    ReDim Preserve mPositionManagerGridMappingTable(UBound(mPositionManagerGridMappingTable) + 20) As PositionManagerGridMappingEntry
End If
If pmIndex > mMaxPositionManagerGridMappingTableIndex Then mMaxPositionManagerGridMappingTableIndex = pmIndex

With mPositionManagerGridMappingTable(pmIndex)
    If .pm Is Nothing Then
        Set .pm = pm
        .gridIndex = addEntryToOrderPlexGrid(pm.Ticker.Contract.specifier.localSymbol, True)
        OrderPlexGrid.rowdata(.gridIndex) = pmIndex + RowDataPositionManagerBase
        OrderPlexGrid.row = .gridIndex
        OrderPlexGrid.col = 1
        OrderPlexGrid.ColSel = OrderPlexGrid.Cols - 1
        OrderPlexGrid.FillStyle = FillStyleSettings.flexFillRepeat
        OrderPlexGrid.CellBackColor = &HC0C0C0
        OrderPlexGrid.CellForeColor = vbWhite
        OrderPlexGrid.CellFontBold = True
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridPositionColumns.exchange) = pm.Ticker.Contract.specifier.exchange
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridPositionColumns.currencyCode) = pm.Ticker.Contract.specifier.currencyCode
        OrderPlexGrid.TextMatrix(.gridIndex, OPGridPositionColumns.size) = pm.positionSize
        OrderPlexGrid.col = OPGridColumns.ExpandIndicator
        OrderPlexGrid.CellPictureAlignment = AlignmentSettings.flexAlignCenterCenter
        Set OrderPlexGrid.CellPicture = OrderPlexImageList.ListImages("Contract").Picture
        .isExpanded = True
    End If
End With
findPositionManagerTableIndex = pmIndex
End Function

Private Sub handleFatalError(ByVal errNum As Long, _
                            ByVal Description As String, _
                            ByVal source As String)
If Not mTicker Is Nothing Then
    Set mTicker = Nothing
Else
    removeServiceProviders
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

Private Sub initialiseTicker(ByVal pTicker As Ticker)
pTicker.outputTickfilePath = App.Path
pTicker.ApplicationData = New TickerApplicationData
Set pTicker.ApplicationData.TickerProxy = pTicker.Proxy
End Sub

'Private Sub openOrder(ByVal pContractSpecifier As ContractSpecifier, _
'                ByVal pOrder As Order)
'
'Dim listItem As listItem
'Dim orderKey As String
'
'orderKey = "A" & CStr(pOrder.id)
'
'On Error Resume Next
'Set listItem = OpenOrdersList.ListItems(orderKey)
'On Error GoTo 0
'
'If listItem Is Nothing Then
'    Set listItem = OpenOrdersList.ListItems.Add(, orderKey, CStr(pOrder.id))
'End If
'
'On Error Resume Next
'If mOrdersCol(orderKey) Is Nothing Then
'    mOrdersCol.Add pOrder, orderKey
'End If
'On Error GoTo 0
'
'On Error Resume Next
'If mContractCol(pContractSpecifier.localSymbol) Is Nothing Then
'    mTradeBuildAPI.RequestContract pContractSpecifier
'End If
'On Error GoTo 0
'
'If LCase$(listItem.SubItems(OpenOrdersColumns.status - 1)) = "filled" Then
'    OpenOrdersList.ListItems.Remove (orderKey)
'    If OpenOrdersList.SelectedItem Is Nothing Then
'        ModifyOrderButton.Enabled = False
'        CancelOrderButton.Enabled = False
'    End If
'    Exit Sub
'End If
'
'listItem.SubItems(OpenOrdersColumns.symbol - 1) = pContractSpecifier.localSymbol
'listItem.SubItems(OpenOrdersColumns.Action - 1) = IIf(pOrder.Action = OrderActions.ActionBuy, "BUY", "SELL")
'If pOrder.triggerPrice <> 0 Then listItem.SubItems(OpenOrdersColumns.auxPrice - 1) = pOrder.triggerPrice
'listItem.SubItems(OpenOrdersColumns.ocaGroup - 1) = pOrder.ocaGroup
'listItem.SubItems(OpenOrdersColumns.orderType - 1) = orderTypeToString(pOrder.orderType)
'If pOrder.limitPrice <> 0 Then listItem.SubItems(OpenOrdersColumns.price - 1) = pOrder.limitPrice
'listItem.SubItems(OpenOrdersColumns.quantity - 1) = pOrder.quantity
'If pOrder.parentId <> 0 Then listItem.SubItems(OpenOrdersColumns.parentId - 1) = pOrder.parentId
'
'listItem.EnsureVisible
'End Sub

Private Sub invertEntryColors(ByVal rowNumber As Long)
Dim foreColor As Long
Dim backColor As Long
Dim i As Long

If rowNumber < 0 Then Exit Sub

OrderPlexGrid.row = rowNumber

For i = OPGridColumns.OtherColumns To OrderPlexGrid.Cols - 1
    OrderPlexGrid.col = i
    foreColor = IIf(OrderPlexGrid.CellForeColor = 0, OrderPlexGrid.foreColor, OrderPlexGrid.CellForeColor)
    If foreColor = SystemColorConstants.vbWindowText Then
        OrderPlexGrid.CellForeColor = SystemColorConstants.vbHighlightText
    ElseIf foreColor = SystemColorConstants.vbHighlightText Then
        OrderPlexGrid.CellForeColor = SystemColorConstants.vbWindowText
    ElseIf foreColor > 0 Then
        OrderPlexGrid.CellForeColor = IIf((foreColor Xor &HFFFFFF) = 0, 1, foreColor Xor &HFFFFFF)
    End If
    
    backColor = IIf(OrderPlexGrid.CellBackColor = 0, OrderPlexGrid.backColor, OrderPlexGrid.CellBackColor)
    If backColor = SystemColorConstants.vbWindowBackground Then
        OrderPlexGrid.CellBackColor = SystemColorConstants.vbHighlight
    ElseIf backColor = SystemColorConstants.vbHighlight Then
        OrderPlexGrid.CellBackColor = SystemColorConstants.vbWindowBackground
    ElseIf backColor > 0 Then
        OrderPlexGrid.CellBackColor = IIf((backColor Xor &HFFFFFF) = 0, 1, backColor Xor &HFFFFFF)
    End If
Next

End Sub

Private Sub removeServiceProviders()
If Not mTWSContractServiceProvider Is Nothing Then
    mTradeBuildAPI.ServiceProviders.Remove mTWSContractServiceProvider
    Set mTWSContractServiceProvider = Nothing
End If
If Not mHistDataServiceProvider Is Nothing Then
    mTradeBuildAPI.ServiceProviders.Remove mHistDataServiceProvider
    Set mHistDataServiceProvider = Nothing
End If
If Not mRealtimeServiceProvider Is Nothing Then
    mTradeBuildAPI.ServiceProviders.Remove mRealtimeServiceProvider
    Set mRealtimeServiceProvider = Nothing
End If
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
Set col = TickerGrid.Columns(TickerGridColumns.Order)
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
Set col = TickerGrid.Columns(TickerGridColumns.highPrice)
col.width = TickerGridColumnWidths.highWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.lowPrice)
col.width = TickerGridColumnWidths.lowWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.closePrice)
col.width = TickerGridColumnWidths.closeWidth * TickerGrid.width / 100
col.Alignment = dbgRight
Set col = TickerGrid.Columns(TickerGridColumns.symbol)
col.width = TickerGridColumnWidths.SymbolWidth * TickerGrid.width / 100
col.Alignment = dbgCenter
Set col = TickerGrid.Columns(TickerGridColumns.sectype)
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

Private Sub setupOrderPlexGrid()
With OrderPlexGrid
    mSelectedOrderPlexGridRow = -1
    .AllowUserResizing = flexResizeBoth
    
    .Cols = 0
    .Rows = 20
    .FixedRows = 3
    ' .FixedCols = 1
    
    setupOrderPlexGridColumn 0, OPGridColumns.ExpandIndicator, OPGridColumnWidths.ExpandIndicatorWidth, "", True, AlignmentSettings.flexAlignCenterCenter
    setupOrderPlexGridColumn 0, OPGridColumns.symbol, OPGridColumnWidths.SymbolWidth, "Symbol", True, AlignmentSettings.flexAlignLeftCenter
    
    setupOrderPlexGridColumn 0, OPGridPositionColumns.currencyCode, OPGridPositionColumnWidths.CurrencyCodeWidth, "Curr", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.drawdown, OPGridPositionColumnWidths.DrawdownWidth, "Drawdown", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.exchange, OPGridPositionColumnWidths.ExchangeWidth, "Exchange", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.MaxProfit, OPGridPositionColumnWidths.MaxProfitWidth, "Max", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.profit, OPGridPositionColumnWidths.ProfitWidth, "Profit", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 0, OPGridPositionColumns.size, OPGridPositionColumnWidths.SizeWidth, "Size", False, AlignmentSettings.flexAlignRightCenter
    
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.creationTime, OPGridOrderPlexColumnWidths.CreationTimeWidth, "Creation Time", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.currencyCode, OPGridOrderPlexColumnWidths.CurrencyCodeWidth, "Curr", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.drawdown, OPGridOrderPlexColumnWidths.DrawdownWidth, "Drawdown", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.MaxProfit, OPGridOrderPlexColumnWidths.MaxProfitWidth, "Max", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.profit, OPGridOrderPlexColumnWidths.ProfitWidth, "Profit", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 1, OPGridOrderPlexColumns.size, OPGridOrderPlexColumnWidths.SizeWidth, "Size", False, AlignmentSettings.flexAlignRightCenter
    
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Action, OPGridOrderColumnWidths.ActionWidth, "Action", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.auxPrice, OPGridOrderColumnWidths.AuxPriceWidth, "Trigger", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.averagePrice, OPGridOrderColumnWidths.AveragePriceWidth, "Avg", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.id, OPGridOrderColumnWidths.IdWidth, "Id", True, AlignmentSettings.flexAlignCenterCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.lastFillPrice, OPGridOrderColumnWidths.LastFillPriceWidth, "Fill", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.LastFillTime, OPGridOrderColumnWidths.LastFillTimeWidth, "Last fill time", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.orderType, OPGridOrderColumnWidths.OrderTypeWidth, "Order type", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.price, OPGridOrderColumnWidths.PriceWidth, "Price", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.quantity, OPGridOrderColumnWidths.QuantityWidth, "Rem Qty", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.size, OPGridOrderColumnWidths.SizeWidth, "Size", False, AlignmentSettings.flexAlignRightCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.Status, OPGridOrderColumnWidths.StatusWidth, "Status", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.typeInPlex, OPGridOrderColumnWidths.TypeInPlexWidth, "Mode", True, AlignmentSettings.flexAlignLeftCenter
    setupOrderPlexGridColumn 2, OPGridOrderColumns.VendorId, OPGridOrderColumnWidths.VendorIdWidth, "Vendor id", True, AlignmentSettings.flexAlignCenterCenter
    
    .MergeCells = flexMergeFree
    .MergeCol(OPGridColumns.symbol) = True
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightAlways
    .FocusRect = flexFocusNone
    
    mFirstOrderPlexGridRowIndex = 3
End With

EditText.Text = ""
End Sub

Private Sub setupOrderPlexGridColumn( _
                ByVal rowNumber As Long, _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As AlignmentSettings)
    
Dim lColumnWidth As Long

With OrderPlexGrid
    .row = rowNumber
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .ColWidth(columnNumber) = 0
    End If
    
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
    If .ColWidth(columnNumber) < lColumnWidth Then
        .ColWidth(columnNumber) = lColumnWidth
    End If
    
    .ColAlignment(columnNumber) = align
    .TextMatrix(rowNumber, columnNumber) = columnHeader
End With
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
Set col = TickerGrid.Columns(TickerGridSummaryColumns.Order)
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

Private Sub setupTWSContractServiceProvider()
If mTWSContractServiceProvider Is Nothing Then
    Set mTWSContractServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.ContractInfoServiceProvider"))
    mTWSContractServiceProvider.Server = ServerText
    mTWSContractServiceProvider.Port = PortText
    If IsNumeric(ClientIDText.Text) Then
        mTWSContractServiceProvider.clientID = CLng(ClientIDText) + 1
    Else
        mTWSContractServiceProvider.clientID = Int(&H7FFFFFFF * Rnd) + 1
    End If
End If
End Sub

Private Sub showMarketDepthForm(ByVal pTicker As Ticker)
Dim tickerAppData As TickerApplicationData
Dim mktDepthForm As fMarketDepth

Set tickerAppData = pTicker.ApplicationData

If Not tickerAppData.MarketDepthForm Is Nothing Then Exit Sub

Set mktDepthForm = New fMarketDepth
Set tickerAppData.MarketDepthForm = mktDepthForm

mktDepthForm.numberOfRows = 100
mktDepthForm.numberOfVisibleRows = 20
mktDepthForm.Ticker = pTicker

pTicker.RequestMarketDepth DOMEvents.DOMProcessedEvents, _
                        IIf(RecordCheck = vbChecked, True, False)

mktDepthForm.Show vbModeless
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

