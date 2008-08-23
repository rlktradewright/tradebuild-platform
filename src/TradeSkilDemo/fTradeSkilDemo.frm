VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#100.0#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fTradeSkilDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TradeSkil Demo Edition Version 2.6"
   ClientHeight    =   6855
   ClientLeft      =   210
   ClientTop       =   330
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   14385
   Begin VB.TextBox TimeZoneText 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox DateTimeText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   11880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   51
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
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   43
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
      TabIndex        =   42
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
      TabIndex        =   41
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   360
      Width           =   2280
   End
   Begin VB.CommandButton ChartButton 
      Caption         =   "C&hart"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   11
      ToolTipText     =   "Display a chart"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton MarketDepthButton 
      Caption         =   "&Market depth"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   10
      ToolTipText     =   "Display the market depth"
      Top             =   0
      Width           =   975
   End
   Begin TabDlg.SSTab MainSSTAB 
      Height          =   4455
      Left            =   120
      TabIndex        =   36
      Top             =   960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&1. Configuration"
      TabPicture(0)   =   "fTradeSkilDemo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CurrentConfigNameText"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ConfigManager1"
      Tab(0).Control(2)=   "ConfigureButton"
      Tab(0).Control(3)=   "Label1"
      Tab(0).ControlCount=   4
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
      Tab(2).Control(1)=   "OrderButton"
      Tab(2).Control(2)=   "CancelOrderPlexButton"
      Tab(2).Control(3)=   "ModifyOrderPlexButton"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&4. Executions"
      TabPicture(3)   =   "fTradeSkilDemo.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ExecutionsSummary1"
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
      TabCaption(5)   =   "&6. Historical charts"
      TabPicture(5)   =   "fTradeSkilDemo.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Chart specifier"
         Height          =   3975
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   6375
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   3645
            Left            =   120
            ScaleHeight     =   3645
            ScaleWidth      =   6015
            TabIndex        =   66
            Top             =   240
            Width           =   6015
            Begin VB.TextBox NumHistBarsText 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2400
               TabIndex        =   69
               Text            =   "500"
               Top             =   840
               Width           =   975
            End
            Begin VB.CommandButton HistChartButton 
               Caption         =   "Show &Chart"
               Enabled         =   0   'False
               Height          =   375
               Left            =   4920
               TabIndex        =   73
               Top             =   0
               Width           =   975
            End
            Begin VB.CheckBox HistSessionOnlyCheck 
               Caption         =   "Session only"
               Height          =   375
               Left            =   2400
               TabIndex        =   70
               Top             =   1200
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin TradeBuildUI26.ContractSpecBuilder HistContractSpecBuilder 
               Height          =   2895
               Left            =   0
               TabIndex        =   67
               Top             =   0
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   5106
            End
            Begin TradeBuildUI26.TimeframeSelector HistTimeframeSelector 
               Height          =   330
               Left            =   2400
               TabIndex        =   68
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   582
            End
            Begin MSComCtl2.DTPicker ToDatePicker 
               Height          =   375
               Left            =   2400
               TabIndex        =   72
               Top             =   2520
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyy-MM-dd HH:mm"
               Format          =   20774915
               CurrentDate     =   39365
            End
            Begin MSComCtl2.DTPicker FromDatePicker 
               Height          =   375
               Left            =   2400
               TabIndex        =   71
               Top             =   1800
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyy-MM-dd HH:mm"
               Format          =   20774915
               CurrentDate     =   39365
            End
            Begin VB.Label Label5 
               Caption         =   "To"
               Height          =   255
               Left            =   2400
               TabIndex        =   77
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "From"
               Height          =   255
               Left            =   2400
               TabIndex        =   76
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Timeframe"
               Height          =   255
               Left            =   2400
               TabIndex        =   75
               Top             =   0
               Width           =   735
            End
            Begin VB.Label Label3 
               Caption         =   "# history bars"
               Height          =   255
               Left            =   2400
               TabIndex        =   74
               Top             =   600
               Width           =   975
            End
         End
      End
      Begin VB.TextBox CurrentConfigNameText 
         Height          =   285
         Left            =   -64560
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1140
         Width           =   3615
      End
      Begin TradeSkilDemo26.ConfigManager ConfigManager1 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7223
      End
      Begin TradeBuildUI26.ExecutionsSummary ExecutionsSummary1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   7011
      End
      Begin TradeBuildUI26.OrdersSummary OrdersSummary1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   7011
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "fTradeSkilDemo.frx":00A8
         Left            =   -73800
         List            =   "fTradeSkilDemo.frx":00AA
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2460
         Width           =   2775
      End
      Begin VB.ListBox TickfileList 
         Height          =   1620
         Left            =   -74400
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   660
         Width           =   6855
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68160
         TabIndex        =   21
         ToolTipText     =   "Stop tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69600
         TabIndex        =   19
         ToolTipText     =   "Pause tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton ClearTickfileListButton 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67440
         TabIndex        =   16
         ToolTipText     =   "Clear tickfile list"
         Top             =   1140
         Width           =   495
      End
      Begin VB.CommandButton SelectTickfilesButton 
         Caption         =   "..."
         Height          =   375
         Left            =   -67440
         TabIndex        =   15
         ToolTipText     =   "Select tickfile(s)"
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -70320
         TabIndex        =   18
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton SkipReplayButton 
         Caption         =   "S&kip"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68880
         TabIndex        =   20
         ToolTipText     =   "Pause tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton ConfigureButton 
         Caption         =   "Load &Configuration"
         Height          =   375
         Left            =   -64560
         TabIndex        =   25
         ToolTipText     =   "Set this configuration"
         Top             =   420
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   4035
         Left            =   60
         ScaleHeight     =   4035
         ScaleWidth      =   13935
         TabIndex        =   38
         Top             =   360
         Width           =   13935
         Begin TradeBuildUI26.TickerGrid TickerGrid1 
            Height          =   3855
            Left            =   3960
            TabIndex        =   62
            Top             =   120
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   6800
         End
         Begin VB.Frame Frame3 
            Caption         =   "Charts"
            Height          =   3015
            Left            =   2280
            TabIndex        =   58
            Top             =   0
            Width           =   1575
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   2775
               Left            =   120
               ScaleHeight     =   2775
               ScaleWidth      =   1395
               TabIndex        =   59
               Top             =   180
               Width           =   1395
               Begin TradeBuildUI26.TimeframeSelector TimeframeSelector1 
                  Height          =   330
                  Left            =   0
                  TabIndex        =   5
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   582
               End
               Begin VB.CheckBox SessionOnlyCheck 
                  Caption         =   "Session only"
                  Height          =   375
                  Left            =   0
                  TabIndex        =   7
                  Top             =   1320
                  Value           =   1  'Checked
                  Width           =   1335
               End
               Begin VB.CommandButton GridChartButton 
                  Caption         =   "Show &Chart"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   360
                  TabIndex        =   8
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.TextBox NumHistoryBarsText 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   0
                  TabIndex        =   6
                  Text            =   "500"
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label Label22 
                  Caption         =   "# history bars"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   61
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label Label18 
                  Caption         =   "Timeframe"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   60
                  Top             =   120
                  Width           =   735
               End
            End
         End
         Begin VB.CommandButton GridMarketDepthButton 
            Caption         =   "&Mkt depth"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2760
            TabIndex        =   9
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton StopTickerButton 
            Caption         =   "Sto&p"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   22
            Top             =   3240
            Width           =   855
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ticker management"
            Height          =   3975
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   2175
            Begin TradeBuildUI26.ContractSpecBuilder ContractSpecBuilder1 
               Height          =   2895
               Left            =   120
               TabIndex        =   0
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   5106
            End
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   3495
               Left            =   120
               ScaleHeight     =   3495
               ScaleWidth      =   1935
               TabIndex        =   2
               Top             =   240
               Width           =   1935
               Begin VB.CommandButton StartTickerButton 
                  Caption         =   "&Start"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   0
                  TabIndex        =   3
                  Top             =   3000
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
         TabIndex        =   12
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderPlexButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   14
         Top             =   1620
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderPlexButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -62280
         TabIndex        =   13
         Top             =   1020
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ReplayProgressBar 
         Height          =   135
         Left            =   -74400
         TabIndex        =   53
         Top             =   3180
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Current configuration is:"
         Height          =   375
         Left            =   -64560
         TabIndex        =   63
         Top             =   1260
         Width           =   2295
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   975
         Left            =   -74400
         TabIndex        =   56
         Top             =   3420
         Width           =   6855
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   -74400
         TabIndex        =   57
         Top             =   2940
         Width           =   6855
      End
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74400
         TabIndex        =   55
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Select tickfile(s)"
         Height          =   255
         Left            =   -74280
         TabIndex        =   54
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.TextBox LogText 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Status messages"
      Top             =   5400
      Width           =   14175
   End
   Begin VB.Label Label27 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Close"
      Height          =   255
      Left            =   10560
      TabIndex        =   34
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Low"
      Height          =   255
      Left            =   9600
      TabIndex        =   33
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "High"
      Height          =   255
      Left            =   8760
      TabIndex        =   32
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Volume"
      Height          =   255
      Left            =   7800
      TabIndex        =   31
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Last/Size"
      Height          =   255
      Left            =   4920
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Ask size"
      Height          =   255
      Left            =   6840
      TabIndex        =   29
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Ask"
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Bid"
      Height          =   255
      Left            =   3960
      TabIndex        =   27
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Bid size"
      Height          =   255
      Left            =   3000
      TabIndex        =   26
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
' Interfaces
'================================================================================

Implements LogListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================
    
Private Const AttributeNameAppConfigName    As String = "Name"

Private Const ConfigNameTradeBuild          As String = "TradeBuild"

Private Const DefaultConfigName             As String = "Default config"

' This is a locally defined 'error' code that can be raised to indicate that
' the program should exit because of some condition. Before raising it,
' the user must be notified via a suitable message.
'Private Const UnloadNotifyException         As Long = vbObjectError + 512

'================================================================================
' Enums
'================================================================================

Private Enum TabIndexNumbers
    TabIndexConfiguration
    TabIndexTickers
    TabIndexOrders
    TabIndexExecutions
    TabIndexReplayTickfiles
    TabIndexHistoricalCharts
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
Private WithEvents mTicker                      As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private WithEvents mTickfileManager             As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1

Private WithEvents mCurrentClock                As Clock
Attribute mCurrentClock.VB_VarHelpID = -1

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
' ensure we get the Windows XP look and feel if running on XP
InitCommonControls
End Sub

Private Sub Form_Load()

On Error GoTo Err

'DefaultLogLevel = TWUtilities30.LogLevelNormal

' position at top left of screen
Me.Left = 0
Me.Top = 0

setupLogging

Set mTradeBuildAPI = TradeBuildAPI

Set mTickers = mTradeBuildAPI.Tickers

setCurrentClock getDefaultClock

OrdersSummary1.monitorWorkspace mTradeBuildAPI.defaultWorkSpace
ExecutionsSummary1.monitorWorkspace mTradeBuildAPI.defaultWorkSpace
TickerGrid1.monitorWorkspace mTradeBuildAPI.defaultWorkSpace

setupReplaySpeedCombo

FromDatePicker.value = DateAdd("m", -1, Now)
FromDatePicker.value = Empty    ' clear the checkbox
ToDatePicker.value = Now

Exit Sub

Err:
handleFatalError Err.Number, _
                Err.Description, _
                Err.source

End Sub

Private Sub Form_QueryUnload( _
                cancel As Integer, _
                UnloadMode As Integer)
If ConfigManager1.changesPending Then
    If MsgBox("Apply these changes?" & vbCrLf & _
            "If you click No, your changes to this configuration item will be lost", _
            vbYesNo Or vbQuestion, _
            "Attention!") = vbYes Then
        ConfigManager1.applyPendingChanges
    End If
End If
If ConfigManager1.dirty Then
    If MsgBox("Permanently save configuration changes?" & vbCrLf & _
            "If you click No, all configuration changes since the last save will be removed from the configuration file", _
            vbYesNo Or vbQuestion, _
            "Attention!") = vbYes Then
        ConfigManager1.saveConfigFile
    End If
End If
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(cancel As Integer)
Dim lTicker As Ticker
Dim f As Form

logMessage "Unloading program"

logMessage "Stopping tickers"
If Not mTickers Is Nothing Then
    For Each lTicker In mTickers
        lTicker.stopTicker
    Next
    Set mTradeBuildAPI = Nothing
End If

OrdersSummary1.finish
ExecutionsSummary1.finish
TickerGrid1.finish

For Each f In Forms
    Unload f
Next

TerminateTWUtilities
End Sub

'================================================================================
' LogListener Interface Members
'================================================================================

Private Sub LogListener_finish()
'nothing to do
End Sub

Private Sub LogListener_Notify(ByVal logrec As TWUtilities30.LogRecord)
LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = formatLogRecord(logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub CancelOrderPlexButton_Click()
Dim op As OrderPlex

Set op = OrdersSummary1.SelectedItem
If Not op Is Nothing Then op.cancel True

CancelOrderPlexButton.Enabled = False
ModifyOrderPlexButton.Enabled = False
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

Private Sub ConfigManager1_GotFocus()
ConfigManager1.Default = True
ConfigManager1.cancel = True
End Sub

Private Sub ConfigManager1_LostFocus()
ConfigManager1.Default = False
ConfigManager1.cancel = False
End Sub

Private Sub ConfigManager1_SelectedItemChanged()
checkOkToLoadConfiguration
End Sub

Private Sub ConfigureButton_Click()
loadAppConfig ConfigManager1.selectedAppConfig
End Sub

Private Sub ContractSpecBuilder1_GotFocus()
StartTickerButton.Default = True
End Sub

Private Sub ContractSpecBuilder1_LostFocus()
StartTickerButton.Default = False
End Sub

Private Sub ContractSpecBuilder1_NotReady()
StartTickerButton.Enabled = False
End Sub

Private Sub ContractSpecBuilder1_Ready()
StartTickerButton.Enabled = True
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

Private Sub HistChartButton_Click()
Dim lTicker As Ticker
Dim fromDate As Date
Dim toDate As Date
Dim chartForm As fChart2

Set lTicker = mTickers.Add(TickerOptions.TickerOptUseExchangeTimeZone)
lTicker.loadTicker HistContractSpecBuilder.contractSpecifier

HistContractSpecBuilder.SetFocus

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

Set chartForm = New fChart2
chartForm.showHistoricalChart lTicker, _
                    NumHistBarsText, _
                    fromDate, _
                    toDate, _
                    IIf(HistSessionOnlyCheck = vbChecked, False, True), _
                    20, _
                    HistTimeframeSelector.timeframeDesignator
chartForm.Show vbModeless
chartForm.Visible = True
End Sub

Private Sub HistContractSpecBuilder_NotReady()
HistChartButton.Enabled = False
End Sub

Private Sub HistContractSpecBuilder_ready()
HistChartButton.Enabled = True
End Sub

Private Sub HistTimeframeSelector_Change()
setHistChartButtonTooltip
End Sub

Private Sub MainSSTAB_Click(PreviousTab As Integer)
If MainSSTAB.Tab = TabIndexTickers Then
    ContractSpecBuilder1.SetFocus
ElseIf MainSSTAB.Tab = TabIndexConfiguration Then
    If ConfigManager1.Visible Then ConfigManager1.SetFocus
ElseIf MainSSTAB.Tab = TabIndexOrders Then
    If OrderButton.Enabled Then
        OrderButton.SetFocus
    ElseIf ModifyOrderPlexButton.Enabled Then
        ModifyOrderPlexButton.SetFocus
    ElseIf CancelOrderPlexButton.Enabled Then
        CancelOrderPlexButton.SetFocus
    End If
ElseIf MainSSTAB.Tab = TabIndexReplayTickfiles Then
    If mTickfileManager Is Nothing Then
        SelectTickfilesButton.SetFocus
    ElseIf PlayTickFileButton.Enabled Then
        PlayTickFileButton.SetFocus
    ElseIf StopReplayButton.Enabled Then
        StopReplayButton.SetFocus
    End If
ElseIf MainSSTAB.Tab = TabIndexHistoricalCharts Then
    HistContractSpecBuilder.SetFocus
End If
End Sub

Private Sub MarketDepthButton_Click()
showMarketDepthForm mTicker
End Sub

Private Sub ModifyOrderPlexButton_Click()
If OrdersSummary1.SelectedItem Is Nothing Then
    ModifyOrderPlexButton.Enabled = False
ElseIf OrdersSummary1.isSelectedItemModifiable Then
    getOrderForm.showOrderPlex OrdersSummary1.SelectedItem, OrdersSummary1.selectedOrderIndex
End If
End Sub

Private Sub NumHistBarsText_Validate(cancel As Boolean)
If Not IsInteger(NumHistBarsText.Text, 0, 2000) Then cancel = True
End Sub

Private Sub NumHistoryBarsText_Validate(cancel As Boolean)
If Not IsInteger(NumHistoryBarsText.Text, 0, 2000) Then cancel = True
End Sub

Private Sub OrderButton_Click()
If mTicker Is Nothing Then
    MsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
    Exit Sub
End If
getOrderForm.Ticker = mTicker
End Sub

Private Sub OrdersSummary1_SelectionChanged()
Dim selection As OrderPlex

Set selection = OrdersSummary1.SelectedItem

If selection Is Nothing Then
    CancelOrderPlexButton.Enabled = False
    ModifyOrderPlexButton.Enabled = False
Else
    If OrdersSummary1.selectedOrderIndex = 0 Then
        CancelOrderPlexButton.Enabled = True
    Else
        CancelOrderPlexButton.Enabled = False
    End If
    If OrdersSummary1.isSelectedItemModifiable Then
        ModifyOrderPlexButton.Enabled = True
    Else
        ModifyOrderPlexButton.Enabled = False
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

Private Sub ReplaySpeedCombo_Click()
If Not mTickfileManager Is Nothing Then
    mTickfileManager.replaySpeed = ReplaySpeedCombo.ItemData(ReplaySpeedCombo.ListIndex)
End If
End Sub

Private Sub SelectTickfilesButton_Click()
Dim tickfiles As TickFileSpecifiers
Dim tfs As TickfileSpecifier
Dim userCancelled As Boolean

Set tickfiles = SelectTickfiles(userCancelled)
If userCancelled Then Exit Sub

Set mTickfileManager = mTickers.createTickFileManager(TickerOptions.TickerOptUseExchangeTimeZone)

mTickfileManager.TickFileSpecifiers = tickfiles

TickfileList.Clear
For Each tfs In tickfiles
    TickfileList.AddItem tfs.filename
Next
checkOkToStartReplay
ClearTickfileListButton.Enabled = True

End Sub

Private Sub SkipReplayButton_Click()
logMessage "Tickfile skipped"
mTickfileManager.SkipTickfile
End Sub

Private Sub StartTickerButton_Click()
Dim lTicker As Ticker

Set lTicker = createTicker
lTicker.DOMEventsRequired = DOMEvents.DOMNoEvents
lTicker.startTicker ContractSpecBuilder1.contractSpecifier

ContractSpecBuilder1.SetFocus
End Sub

Private Sub StopReplayButton_Click()

PlayTickFileButton.Enabled = True
PauseReplayButton.Enabled = False
SkipReplayButton.Enabled = True
StopReplayButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
ChartButton.Enabled = False
mTicker.stopTicker
End Sub

Private Sub StopTickerButton_Click()
TickerGrid1.stopSelectedTickers
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
        
        clearTickerFields
        
        Set mTicker = TickerGrid1.SelectedTickers.Item(1)
        
        If mTicker.State = TickerStateRunning Then
            If mTicker.defaultOrderContext.IsReady Then
                OrderButton.Enabled = True
            End If
            
            setCurrentClock mTicker.Clock
            
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
        Else
            
            setCurrentClock getDefaultClock
            
            GridChartButton.Enabled = False
            GridMarketDepthButton.Enabled = False
            OrderButton.Enabled = False
        End If
    End If
End If

End Sub

Private Sub TimeframeSelector1_Click()
setChartButtonTooltip
End Sub

'================================================================================
' mCurrentClock Event Handlers
'================================================================================

Private Sub mCurrentClock_Tick()
displayTime
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_ask(ev As QuoteEvent)
AskText = ev.priceString
AskSizeText = ev.Size
setForeColor AskText, ev.priceChange
setForeColor AskSizeText, ev.sizeChange
End Sub

Private Sub mTicker_bid(ev As QuoteEvent)
BidText = ev.priceString
BidSizeText = ev.Size
setForeColor BidText, ev.priceChange
setForeColor BidSizeText, ev.sizeChange
End Sub

Private Sub mTicker_high(ev As QuoteEvent)
HighText = ev.priceString
End Sub

Private Sub mTicker_Low(ev As QuoteEvent)
LowText = ev.priceString
End Sub

Private Sub mTicker_previousClose(ev As QuoteEvent)
CloseText = ev.priceString
End Sub

Private Sub mTicker_trade(ev As QuoteEvent)
LastText = ev.priceString
LastSizeText = ev.Size
setForeColor LastText, ev.priceChange
setForeColor LastSizeText, ev.sizeChange
End Sub

Private Sub mTicker_volume(ev As QuoteEvent)
VolumeText = ev.Size
End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_Notification( _
                ByRef ev As NotificationEvent)
Dim lTicker As Ticker

On Error GoTo Err

Set lTicker = ev.source

Select Case ev.eventCode
Case ApiNotifyCodes.ApiNotifyContractDoesNotExist
    logMessage "Contract does not exist: " & _
                        lTicker.Contracts.contractSpecifier.toString
    modelessMsgBox "No such contract is available - you specified: " & vbCrLf & _
                        vbTab & lTicker.Contracts.contractSpecifier.toString, _
                    MsgBoxCritical, _
                    "Attention"
Case ApiNotifyCodes.ApiNotifyContractExpired
    logMessage "Contract has expired: " & _
                        lTicker.Contracts.contractSpecifier.toString
    modelessMsgBox "Contract has expired - you specified: " & vbCrLf & _
                        vbTab & lTicker.Contracts.contractSpecifier.toString, _
                    MsgBoxCritical, _
                    "Attention"
Case ApiNotifyCodes.ApiNotifyContractSpecifierAmbiguous
    logMessage "Ambiguous contract details (" & ev.eventMessage & "): " & _
                        lTicker.Contracts.contractSpecifier.toString
    modelessMsgBox "Ambiguous contract details - you specified: " & vbCrLf & _
                        vbTab & lTicker.Contracts.contractSpecifier.toString & vbCrLf & _
                        "which matched: " & vbCrLf & _
                        lTicker.Contracts.toString, _
                    MsgBoxInformation, _
                    "Attention"
Case ApiNotifyCodes.ApiNotifyContractSpecifierInvalid
    logMessage "Invalid contract details (" & ev.eventMessage & "): " & _
                        lTicker.Contracts.contractSpecifier.toString
    modelessMsgBox "Invalid contract details - you specified: " & vbCrLf & _
                        vbTab & lTicker.Contracts.contractSpecifier.toString & vbCrLf & _
                        ev.eventMessage, _
                    MsgBoxCritical, _
                    "Attention"
Case ApiNotifyCodes.ApiNotifyMarketDepthNotAvailable
    logMessage "No market depth for: " & _
                        lTicker.Contract.specifier.toString & _
                        ": " & ev.eventMessage
    modelessMsgBox "No market depth available for: " & vbCrLf & _
                                lTicker.Contract.specifier.localSymbol & vbCrLf & _
                                ev.eventMessage, _
                    MsgBoxInformation, _
                    "Attention"
Case ApiNotifyCodes.ApiNotifyRealtimeDataRequestFailed
    logMessage "Market data request failed for: " & _
                        lTicker.Contract.specifier.localSymbol & _
                        ": " & ev.eventMessage
    modelessMsgBox "Market data request failed for: " & vbCrLf & _
                        vbTab & lTicker.Contract.specifier.toString & vbCrLf & _
                        ev.eventMessage, _
                    MsgBoxCritical, _
                    "Attention"
Case Else
    logMessage "Notification " & ev.eventCode & ": " & ev.eventMessage
End Select

Exit Sub
Err:
handleFatalError Err.Number, Err.Description, "mTickers_tickerError"
End Sub

Private Sub mTickers_StateChange( _
                ev As StateChangeEvent)
Dim lTicker As Ticker

On Error GoTo Err

Set lTicker = ev.source

Select Case ev.State
Case TickerStateCreated
    ConfigureButton.Enabled = False
Case TickerStateStarting

Case TickerStateReady
    If lTicker Is mTicker Then setCurrentClock mTicker.Clock
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
        setCurrentClock getDefaultClock
    End If
    
    checkOkToLoadConfiguration
End Select

Exit Sub

Err:
handleFatalError Err.Number, Err.Description, "mTickers_TickerStateEvent"
End Sub

'================================================================================
' mTickfileManager Event Handlers
'================================================================================

Private Sub mTickfileManager_Notification( _
                ev As NotificationEvent)
On Error GoTo Err
logMessage "Notification " & ev.eventCode & ": " & ev.eventMessage

Exit Sub
Err:
handleFatalError Err.Number, Err.Description, "mTickfileManager_Notification"
End Sub

Private Sub mTickfileManager_QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As Contract, _
                continueMode As ReplayContinueModes)
On Error GoTo Err

If tickfileIndex <> 0 Then
    clearTickerFields
    Set mTicker = Nothing
    setCurrentClock getDefaultClock
End If

ReplayProgressBar.Min = 0
ReplayProgressBar.Max = 100
ReplayProgressBar.value = 0
TickfileList.ListIndex = tickfileIndex - 1
ReplayContractLabel.caption = Replace(pContract.specifier.toString, vbCrLf, "; ")

Exit Sub
Err:
handleFatalError Err.Number, Err.Description, "mTickfileManager_QueryReplayNextTickfile"
End Sub

Private Sub mTickfileManager_ReplayCompleted()
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

logMessage "Tickfile replay completed"

Exit Sub
Err:
handleFatalError Err.Number, Err.Description, "mTickfileManager_ReplayCompleted"
End Sub

Private Sub mTickfileManager_ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)
On Error GoTo Err
ReplayProgressBar.value = percentComplete
ReplayProgressLabel.caption = tickfileTimestamp & _
                                "  Processed " & _
                                eventsPlayed & _
                                " events"

Exit Sub
Err:
handleFatalError Err.Number, Err.Description, "mTickfileManager_ReplayProgress"
End Sub

Private Sub mTickfileManager_TickerAllocated(ByVal pTicker As Ticker)
On Error GoTo Err

Set mTicker = pTicker
mTicker.DOMEventsRequired = DOMProcessedEvents

Exit Sub
Err:
handleFatalError Err.Number, Err.Description, "mTickfileManager_TickerAllocated"
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

' fires when an unrecoverable error has occurred in TradeBuild
Private Sub mTradeBuildAPI_Error( _
                ByRef ev As ErrorEvent)
On Error GoTo Err

handleFatalError ev.errorCode, ev.errorMessage, "mTradeBuildAPI_Error"

Exit Sub
Err:

' an error has occurred in the fatal error handler!! All we can do is terminate abruptly
End

End Sub

Private Sub mTradeBuildAPI_Notification( _
                ByRef ev As NotificationEvent)
Dim spError As ServiceProviderError

On Error GoTo Err

Select Case ev.eventCode
Case ApiNotifyCodes.ApiNotifyServiceProviderError
    Set spError = mTradeBuildAPI.GetServiceProviderError
    logMessage "Error from " & _
                        spError.serviceProviderName & _
                        ": code " & spError.errorCode & _
                        ": " & spError.message

Case Else
    logMessage "Notification " & ev.eventCode & ": " & ev.eventMessage
End Select

Exit Sub
Err:
handleFatalError Err.Number, Err.Description, "mTradeBuildAPI_notification"
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Function configure() As Boolean
On Error GoTo Err

If getConfigToLoad() Is Nothing Then
    ' put the user on the configuration tab
    MainSSTAB.Tab = TabIndexConfiguration
Else
    loadAppConfig getConfigToLoad()
End If

configure = True

Exit Function

Err:
configure = False
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub checkOkToLoadConfiguration()
If mTickers.Count = 0 And _
    Not ConfigManager1.selectedAppConfig Is Nothing _
Then
    ConfigureButton.Enabled = True
Else
    ConfigureButton.Enabled = False
End If
End Sub

Private Sub checkOkToStartReplay()
If TickfileList.ListCount <> 0 Then
    PlayTickFileButton.Enabled = True
Else
    PlayTickFileButton.Enabled = False
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
Dim chartForm As fChart2
Dim tp As TimePeriod

If Not pTicker.State = TickerStateRunning Then Exit Sub

Set tp = TimeframeSelector1.timeframeDesignator
Set chartForm = New fChart2
chartForm.showChart pTicker, _
                    NumHistoryBarsText, _
                    IIf(SessionOnlyCheck = vbChecked, False, True), _
                    20, _
                    tp
chartForm.Show vbModeless
chartForm.Visible = True
End Sub

Private Function createTicker() As Ticker
Set createTicker = mTickers.Add(TickerOptions.TickerOptOrdersAreLive + _
                                TickerOptions.TickerOptUseExchangeTimeZone)
End Function

Private Sub displayTime()
Dim theTime As Date
theTime = mCurrentClock.TimeStamp
DateTimeText = FormatDateTime(theTime, vbShortDate) & vbCrLf & _
                Format(theTime, "hh:mm:ss")
End Sub

Private Function formatLogRecord(ByVal logrec As LogRecord) As String
Static formatter As LogFormatter
If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.formatRecord(logrec)
End Function

Private Function getConfigFilename() As String

getConfigFilename = gCommandLineParser.Arg(0)
If getConfigFilename = "" Then
    getConfigFilename = GetSpecialFolderPath(FolderIdLocalAppdata) & _
                        "\TradeWright\" & _
                        AppName & _
                        "\v" & _
                        App.Major & "." & App.Minor & _
                        "\settings.xml"
End If
End Function

Private Function getConfigToLoad() As ConfigItem
Static configToLoad As ConfigItem

If configToLoad Is Nothing Then
    If Not ConfigManager1.initialise(getConfigFilename, App.ProductName) Then
        MsgBox "The configuration file (" & _
                getConfigFilename & _
                ") is not the correct format for this program", _
                vbCritical, _
                "Error"
        Err.Raise ErrorCodes.ErrIllegalArgumentException
    End If
    
    On Error Resume Next
    Set configToLoad = getNamedConfig()
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    If configToLoad Is Nothing Then
        Set configToLoad = ConfigManager1.selectedAppConfig
    End If

    If configToLoad Is Nothing Then
        Set configToLoad = ConfigManager1.firstAppConfig
    End If

    If configToLoad Is Nothing Then
        If MsgBox("No existing configuration details can be found. Would you like to " & vbCrLf & _
                "proceed with a default configuration?" & vbCrLf & vbCrLf & _
                "The default configuration will connect to TWS running on the " & vbCrLf & _
                "same computer. It wll obtain contract data and historical data " & vbCrLf & _
                "from TWS, and will simulate any orders placed." & vbCrLf & vbCrLf & _
                "You may amend the default configuration by going to the " & vbCrLf & _
                "Configuration tab." & vbCrLf & vbCrLf & _
                "Click Yes to continue with the default configuration. Click No " & vbCrLf & _
                "to manually set up the configuration you want.", _
                vbYesNo Or vbQuestion, _
                "Attention!") = vbYes _
        Then
            logMessage ("Creating a new default configuration")
            ConfigManager1.createNewAppConfig DefaultConfigName, True, True
            Set configToLoad = ConfigManager1.appConfig(DefaultConfigName)
        End If
    End If
End If

Set getConfigToLoad = configToLoad

End Function

Private Function getDefaultClock() As Clock
Static lClock As Clock
If lClock Is Nothing Then Set lClock = GetClock("") ' create a clock running local time
Set getDefaultClock = lClock
End Function

Private Function getNamedConfig() As ConfigItem
Dim configName As String

If Not gCommandLineParser.Switch(SwitchConfig) Then Exit Function

configName = gCommandLineParser.SwitchValue(SwitchConfig)

If configName <> "" Then
    Set getNamedConfig = ConfigManager1.appConfig(configName)
    If getNamedConfig Is Nothing Then
        MsgBox "The required configuration does not exist: " & configName, _
                vbCritical, _
                "Error"
        Err.Raise ErrorCodes.ErrIllegalArgumentException
        Exit Function
    End If
End If
End Function

Private Function getOrderForm() As OrderForm
Static lOrderForm As OrderForm
If lOrderForm Is Nothing Then
    Set lOrderForm = New OrderForm
    lOrderForm.Show vbModeless
End If
Set getOrderForm = lOrderForm
End Function

Private Sub handleFatalError(ByVal errNum As Long, _
                            ByVal Description As String, _
                            ByVal source As String)
Set mTicker = Nothing

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

Private Sub loadAppConfig( _
                ByVal configToLoad As ConfigItem)
ExecutionsSummary1.Clear
removeServiceProviders

logMessage "Loading configuration: " & configToLoad.getAttribute(AttributeNameAppConfigName)

If Not setupServiceProviders(configToLoad) Or _
    Not setupStudyLibraries(configToLoad) _
Then
    logMessage "Failed loading configuration: " & configToLoad.getAttribute(AttributeNameAppConfigName)
    ' put the user on the configuration tab
    MainSSTAB.Tab = TabIndexConfiguration
    Exit Sub
End If

' now set up the timeframe selectors, which depends on what timeframes the historical data service
' provider supports (it obtains this info from TradeBuild)
TimeframeSelector1.initialise   ' use the default settings built-in to the control
TimeframeSelector1.selectTimeframe GetTimePeriod(5, TimePeriodMinute)
HistTimeframeSelector.initialise
HistTimeframeSelector.selectTimeframe GetTimePeriod(5, TimePeriodMinute)

setChartButtonTooltip
setHistChartButtonTooltip

logMessage "Loaded configuration: " & configToLoad.getAttribute(AttributeNameAppConfigName)
CurrentConfigNameText = configToLoad.getAttribute(AttributeNameAppConfigName)
End Sub

Private Sub logMessage(message As String)
If Not gLogger Is Nothing Then gLogger.Log LogLevelNormal, message
End Sub

Private Sub modelessMsgBox( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Dim lMsgBox As New fMsgBox

lMsgBox.initialise prompt, buttons, title

lMsgBox.Show vbModeless, Me
                
End Sub
Private Sub removeServiceProviders()
If Not mTradeBuildAPI Is Nothing Then mTradeBuildAPI.ServiceProviders.RemoveAll
RemoveAllStudyLibraries
End Sub

Private Sub setChartButtonTooltip()
Dim tp As TimePeriod

Set tp = TimeframeSelector1.timeframeDesignator

ChartButton.ToolTipText = "Show " & _
                        tp.toString & _
                        " chart"
GridChartButton.ToolTipText = ChartButton.ToolTipText
End Sub

Private Sub setCurrentClock( _
                ByVal pClock As Clock)
Set mCurrentClock = pClock
TimeZoneText = mCurrentClock.TimeZone.standardName
displayTime
End Sub

Private Sub setForeColor( _
                ByVal pControl As Control, _
                ByVal change As ValueChanges)
If change = ValueChangeUp Then
    pControl.ForeColor = IncreasedValueColor
ElseIf change = ValueChangeDown Then
    pControl.ForeColor = DecreasedValueColor
End If
End Sub

Private Sub setHistChartButtonTooltip()
Dim tp As TimePeriod

Set tp = HistTimeframeSelector.timeframeDesignator

HistChartButton.ToolTipText = "Show " & _
                        tp.toString & _
                        " chart"
End Sub

Private Sub setupLogging()
gLogger.addLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

gLogger.Log TWUtilities30.LogLevels.LogLevelNormal, "Log file: " & gLogFileName
gLogger.Log TWUtilities30.LogLevels.LogLevelNormal, "Log level: " & LogLevelToString(DefaultLogLevel)

End Sub

Private Sub setupReplaySpeedCombo()
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

End Sub

Private Function setupServiceProviders( _
                ByVal config As ConfigItem) As Boolean
Dim tradebuildEntry As ConfigItem

On Error Resume Next
Set tradebuildEntry = config.childItems.Item(ConfigNameTradeBuild)
On Error GoTo 0

If tradebuildEntry Is Nothing Then
    logMessage "No service providers defined in this configuration"
    setupServiceProviders = False
Else
    On Error GoTo Err
    mTradeBuildAPI.ServiceProviders.loadServiceProviderConfiguration tradebuildEntry
    setupServiceProviders = True
End If

Exit Function

Err:
logMessage "Service provider configuration failed: " & Err.Description
setupServiceProviders = False
End Function

Private Function setupStudyLibraries( _
                ByVal config As ConfigItem) As Boolean
Dim tradebuildEntry As ConfigItem

On Error Resume Next
Set tradebuildEntry = config.childItems.Item(ConfigNameTradeBuild)
On Error GoTo 0

If tradebuildEntry Is Nothing Then
    logMessage "No study libraries defined in this configuration"
    setupStudyLibraries = False
Else
    LoadStudyLibraryConfiguration tradebuildEntry
    setupStudyLibraries = True
End If

End Function

Private Sub showMarketDepthForm(ByVal pTicker As Ticker)
Dim mktDepthForm As fMarketDepth

If Not pTicker.State = TickerStateRunning Then Exit Sub

Set mktDepthForm = New fMarketDepth
mktDepthForm.numberOfRows = 100
mktDepthForm.Ticker = pTicker
mktDepthForm.Show vbModeless
End Sub




