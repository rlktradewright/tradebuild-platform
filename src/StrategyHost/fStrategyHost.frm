VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#294.0#0"; "TradingUI27.ocx"
Begin VB.Form fStrategyHost 
   Caption         =   "TradeBuild Strategy Host v2.7"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   47
      Top             =   3600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Price chart"
      TabPicture(0)   =   "fStrategyHost.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PriceChart"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Daily profit chart"
      TabPicture(1)   =   "fStrategyHost.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ProfitChart"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Trade chart"
      TabPicture(2)   =   "fStrategyHost.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TradeChart"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Bracket order details"
      TabPicture(3)   =   "fStrategyHost.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "BracketOrderList"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin TradingUI27.MultiChart PriceChart 
         Height          =   5175
         Left            =   0
         TabIndex        =   57
         Top             =   300
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   9128
      End
      Begin TradingUI27.MarketChart TradeChart 
         Height          =   3615
         Left            =   -75000
         TabIndex        =   56
         Top             =   300
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6376
      End
      Begin TradingUI27.MarketChart ProfitChart 
         Height          =   5295
         Left            =   -75000
         TabIndex        =   49
         Top             =   300
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9340
      End
      Begin MSComctlLib.ListView BracketOrderList 
         Height          =   5295
         Left            =   -75000
         TabIndex        =   48
         Top             =   300
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   9340
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3375
      Left            =   0
      TabIndex        =   21
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Controls"
      TabPicture(0)   =   "fStrategyHost.frx":0070
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "fStrategyHost.frx":008C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Parameters"
      TabPicture(2)   =   "fStrategyHost.frx":00A8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Listeners"
      TabPicture(3)   =   "fStrategyHost.frx":00C4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Log"
      TabPicture(4)   =   "fStrategyHost.frx":00E0
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LogPicture"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Results"
      TabPicture(5)   =   "fStrategyHost.frx":00FC
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "AskSizeText"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "TradeSizeText"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "BidSizeText"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "AskText"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "TradeText"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "BidText"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "MoreButton"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label7"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "MicrosecsPerEventLabel"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "EventsPerSecondLabel"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label3"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "PercentCompleteLabel"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Label2"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "EventsPlayedLabel"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "Label1"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Label8"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "Label10"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "Label9"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "Label4"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "Profit"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "Drawdown"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "Label12"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "Label5"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "MaxProfit"
      Tab(5).Control(23).Enabled=   0   'False
      Tab(5).Control(24)=   "Position"
      Tab(5).Control(24).Enabled=   0   'False
      Tab(5).Control(25)=   "Label14"
      Tab(5).Control(25).Enabled=   0   'False
      Tab(5).Control(26)=   "TheTime"
      Tab(5).Control(26).Enabled=   0   'False
      Tab(5).ControlCount=   27
      Begin VB.PictureBox LogPicture 
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   -74880
         ScaleHeight     =   2715
         ScaleWidth      =   10680
         TabIndex        =   54
         Top             =   480
         Width           =   10675
         Begin VB.TextBox LogText 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2700
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   55
            Top             =   0
            Width           =   10695
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2310
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   2310
         ScaleWidth      =   10695
         TabIndex        =   52
         Top             =   480
         Width           =   10695
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   10695
         TabIndex        =   50
         Top             =   480
         Width           =   10695
         Begin VB.ComboBox StopStrategyFactoryCombo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "fStrategyHost.frx":0118
            Left            =   6000
            List            =   "fStrategyHost.frx":011F
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox SymbolText 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            TabIndex        =   0
            Top             =   0
            Width           =   1815
         End
         Begin VB.ComboBox StrategyCombo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "fStrategyHost.frx":0146
            Left            =   6000
            List            =   "fStrategyHost.frx":014D
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Width           =   3495
         End
         Begin VB.CheckBox DummyProfitProfileCheck 
            Caption         =   "Dummy profit profile"
            Height          =   195
            Left            =   6000
            TabIndex        =   7
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox ProfitProfileCheck 
            Caption         =   "Profit profile"
            Height          =   195
            Left            =   6000
            TabIndex        =   6
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox NoMoneyManagement 
            Caption         =   "No money management"
            Height          =   195
            Left            =   6000
            TabIndex        =   8
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CheckBox SeparateSessionsCheck 
            Caption         =   "Separate session per tick file"
            Height          =   195
            Left            =   8040
            TabIndex        =   9
            Top             =   840
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.CommandButton StopButton 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9600
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Default         =   -1  'True
            Enabled         =   0   'False
            Height          =   375
            Left            =   9600
            TabIndex        =   4
            Top             =   0
            Width           =   1095
         End
         Begin VB.CheckBox LiveTradesCheck 
            Caption         =   "Live trades"
            Height          =   195
            Left            =   8040
            TabIndex        =   10
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox ResultsPathText 
            Height          =   255
            Left            =   6960
            TabIndex        =   11
            Top             =   1560
            Width           =   1995
         End
         Begin VB.CommandButton ResultsPathButton 
            Caption         =   "..."
            Height          =   255
            Left            =   9000
            TabIndex        =   12
            ToolTipText     =   "Select results path"
            Top             =   1560
            Width           =   375
         End
         Begin TradingUI27.TickfileOrganiser TickfileOrganiser1 
            Height          =   2325
            Left            =   0
            TabIndex        =   1
            Top             =   360
            Width           =   5610
            _ExtentX        =   9895
            _ExtentY        =   4101
         End
         Begin VB.Label Label 
            Caption         =   "Symbol"
            Height          =   375
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Results path"
            Height          =   255
            Left            =   6000
            TabIndex        =   51
            Top             =   1560
            Width           =   975
         End
      End
      Begin VB.TextBox AskSizeText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TradeSizeText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox BidSizeText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox AskText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TradeText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox BidText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton MoreButton 
         Caption         =   "Less <<<"
         Height          =   375
         Left            =   -68400
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1250
         Left            =   -74880
         ScaleHeight     =   1245
         ScaleWidth      =   7455
         TabIndex        =   25
         Top             =   360
         Width           =   7455
         Begin VB.TextBox ListenerFilenameText 
            Enabled         =   0   'False
            Height          =   255
            Left            =   1215
            TabIndex        =   19
            Top             =   870
            Width           =   4575
         End
         Begin VB.CommandButton SelectListenerFileButton 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   5880
            TabIndex        =   20
            Top             =   840
            Width           =   375
         End
         Begin VB.CheckBox RawDataCheck 
            Enabled         =   0   'False
            Height          =   195
            Left            =   1215
            TabIndex        =   18
            Top             =   660
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox ValueTypeText 
            Height          =   285
            Left            =   1215
            TabIndex        =   17
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox ListenerTypeCombo 
            Height          =   315
            ItemData        =   "fStrategyHost.frx":016E
            Left            =   1215
            List            =   "fStrategyHost.frx":0178
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   0
            Width           =   2175
         End
         Begin VB.Label Label19 
            Caption         =   "File name"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   870
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Log raw data?"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   645
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "Value type"
            Height          =   255
            Left            =   15
            TabIndex        =   27
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Listener type"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   2850
         Left            =   -74880
         ScaleHeight     =   2850
         ScaleWidth      =   10695
         TabIndex        =   24
         Top             =   360
         Width           =   10695
         Begin MSDataGridLib.DataGrid ParamGrid 
            Height          =   2775
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4895
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            Appearance      =   0
            BorderStyle     =   0
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   0
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
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1250
         Left            =   -74880
         ScaleHeight     =   1245
         ScaleWidth      =   7455
         TabIndex        =   23
         Top             =   360
         Width           =   7455
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1365
         Index           =   0
         Left            =   120
         ScaleHeight     =   1365
         ScaleWidth      =   7455
         TabIndex        =   22
         Top             =   360
         Width           =   7455
      End
      Begin VB.Label Label7 
         Caption         =   "Microsecs per event"
         Height          =   195
         Left            =   -71280
         TabIndex        =   65
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label MicrosecsPerEventLabel 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   64
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label EventsPerSecondLabel 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   63
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Events per second"
         Height          =   195
         Left            =   -71280
         TabIndex        =   62
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label PercentCompleteLabel 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   61
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Percent complete"
         Height          =   195
         Left            =   -71280
         TabIndex        =   60
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label EventsPlayedLabel 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   59
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Events played"
         Height          =   195
         Left            =   -71280
         TabIndex        =   58
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Bid"
         Height          =   195
         Left            =   -74880
         TabIndex        =   40
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Last"
         Height          =   195
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Ask"
         Height          =   195
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Profit/Loss"
         Height          =   195
         Left            =   -71280
         TabIndex        =   38
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Profit 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Drawdown 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Drawdown"
         Height          =   195
         Left            =   -71280
         TabIndex        =   35
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Max profit"
         Height          =   195
         Left            =   -71280
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.Label MaxProfit 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         Height          =   195
         Left            =   -69600
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Position"
         Height          =   195
         Left            =   -71280
         TabIndex        =   31
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label TheTime 
         Caption         =   " "
         Height          =   255
         Left            =   -74055
         TabIndex        =   30
         Top             =   1320
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog CommonDialogs 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "fStrategyHost"
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

Implements IGenericTickListener
Implements ILogListener
Implements IStrategyHost

'================================================================================
' Events
'================================================================================

'================================================================================
' Declares
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                        As String = "fStrategyHost"

Private Const LB_SETHORZEXTENT                  As Long = &H194&

'================================================================================
' Enums
'================================================================================

Private Enum BOListColumns
    ColumnKey = 1
    ColumnStartTime
    ColumnEndTime
    ColumnAction
    ColumnQuantity
    ColumnEntryPrice
    ColumnExitPrice
    ColumnProfit
    ColumnMaxProfit
    ColumnMaxLoss
    ColumnRisk
    ColumnQuantityOutstanding
    ColumnEntryReason
    ColumnTargetReason
    ColumnStopReason
    ColumnClosedOut
    ColumnDescription
End Enum

' Percentage widths of the bracket order list columns
Private Enum BOListColumnWidths
    WidthKey = 9
    WidthStartTime = 20
    WidthEndTime = 20
    WidthDescription = 50
    WidthAction = 8
    WidthQuantity = 5
    WidthQuantityOutstanding = 5
    WidthEntryPrice = 10
    WidthExitPrice = 10
    WidthProfit = 8
    WidthMaxProfit = 8
    WidthMaxLoss = 8
    WidthRisk = 8
    WidthEntryReason = 10
    WidthTargetReason = 10
    WidthStopReason = 10
    WidthClosedOut = 4
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTickfileIndex                                  As Long

Private mTicker                                         As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private mContract                                       As IContract
Private mSecType                                        As SecurityTypes
Private mTickSize                                       As Double

Private WithEvents mSession                             As Session
Attribute mSession.VB_VarHelpID = -1

Private mParams                                         As Parameters
Private mStrategyRunner                                 As StrategyRunner
Attribute mStrategyRunner.VB_VarHelpID = -1

Private mCurrTickfileIndex                              As Long

Private mPriceStudyBase                                 As StudyBaseForTickDataInput

Private mPriceChartTimePeriod                           As TimePeriod
Private mPriceChartIndex                                As Long

Private mProfitStudyBase                                As StudyBaseForDoubleInput

Private mTradeStudyBase                                 As StudyBaseForUserBarsInput
Private mTradeBar                                       As BarUtils27.Bar
Private mTradeBarNumber                                 As Long

Private mPosition As Long
Private mOverallProfit As Double
Private mSessionProfit As Double

Private mDetailsHidden As Boolean

Private mBracketOrderLineSeries                         As LineSeries

Private mPricePeriods                                   As Periods

Private mReplayStartTime                                As Date

Private mTotalElapsedSecs                               As Double
Private mElapsedSecsCurrTickfile                        As Double
Private mTotalEvents                                    As Long
Private mEventsCurrTickfile                             As Long

Private WithEvents mFutureWaiter                        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
Me.ScaleMode = vbTwips
InitialiseCommonControls
Set mFutureWaiter = New FutureWaiter
End Sub

Private Sub Form_Load()
setupLogging
setupBracketOrderList
If Not gTB.TickfileStoreInput Is Nothing Then
    TickfileOrganiser1.Initialise gTB.TickfileStoreInput, gTB.ContractStorePrimary
    TickfileOrganiser1.Enabled = True
End If
End Sub

Private Sub Form_Resize()
SSTab1.Width = ScaleWidth
SSTab2.Width = ScaleWidth

If ScaleHeight < minimumHeight Or mDetailsHidden Then
    Me.Height = minimumHeight + (Me.Height - Me.ScaleHeight)
    Exit Sub
End If

If Not mDetailsHidden Then
    If ScaleHeight - SSTab1.Top > 0 Then SSTab1.Height = ScaleHeight - SSTab1.Top
    PriceChart.Width = SSTab1.Width
    If SSTab1.Height - PriceChart.Top > 0 Then PriceChart.Height = SSTab1.Height - PriceChart.Top
    ProfitChart.Width = SSTab1.Width
    If SSTab1.Height - ProfitChart.Top > 0 Then ProfitChart.Height = SSTab1.Height - ProfitChart.Top
    TradeChart.Width = SSTab1.Width
    If SSTab1.Height - TradeChart.Top > 0 Then TradeChart.Height = SSTab1.Height - TradeChart.Top
    BracketOrderList.Width = SSTab1.Width
    If SSTab1.Height - BracketOrderList.Top > 0 Then BracketOrderList.Height = SSTab1.Height - BracketOrderList.Top
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

GetLogger("log").RemoveLogListener Me

LogMessage "Unloading main form"

If Not mStrategyRunner Is Nothing Then
    LogMessage "Stopping strategy host"
    mStrategyRunner.StopTesting
End If

LogMessage "Finishing charts"
PriceChart.Finish
ProfitChart.Finish
TradeChart.Finish

LogMessage "Closing other forms"
Dim f As Form
For Each f In Forms
    If Not TypeOf f Is fStrategyHost Then
        LogMessage "Closing form: caption=" & f.Caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' IGenericTickListener Interface Members
'================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypes.TickTypeAsk
    AskText.Text = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
    AskSizeText.Text = ev.Tick.Size
Case TickTypes.TickTypeBid
    BidText.Text = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
    BidSizeText.Text = ev.Tick.Size
Case TickTypes.TickTypeTrade
    TradeText.Text = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
    TradeSizeText.Text = ev.Tick.Size
    mPriceStudyBase.NotifyTick ev.Tick
Case TickTypes.TickTypeVolume
    mPriceStudyBase.NotifyTick ev.Tick
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IStrategyHost Interface Members
'================================================================================

Private Function IStrategyHost_AddStudy( _
                ByVal pName As String, _
                ByVal pUnderlyingStudy As IStudy, _
                ByRef pInputValueNames() As String, _
                Optional ByVal pParams As Parameters, _
                Optional ByVal pNumberOfValuesToCache As Long, _
                Optional ByVal pLibraryName As String) As IStudy
Const ProcName As String = "IStrategyHost_AddStudy"
On Error GoTo Err

Dim lChartManager As ChartManager
Set lChartManager = PriceChart.ChartManager

Dim lStudyConfig As StudyConfiguration
Set lStudyConfig = lChartManager.GetDefaultStudyConfiguration(pName, pLibraryName)

lStudyConfig.UnderlyingStudy = pUnderlyingStudy

Dim lInputValueNames() As String
lInputValueNames = pInputValueNames
lStudyConfig.InputValueNames = lInputValueNames
lStudyConfig.Parameters = pParams
lStudyConfig.StudyLibraryName = pLibraryName

Dim lStudy As IStudy
Set lStudy = lChartManager.AddStudyConfiguration(lStudyConfig)
lChartManager.StartStudy lStudy

Set IStrategyHost_AddStudy = lStudy

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStrategyHost_AddTimeframe( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pNumberOfBarsToFetch As Long, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean) As Timeframe
Const ProcName As String = "IStrategyHost_AddTimeframe"
On Error GoTo Err

Dim lIndex As Long
If TickfileOrganiser1.TickfileCount <> 0 Then
    lIndex = PriceChart.Add(pTimePeriod, mContract.Specifier.Symbol, False, pNumberOfBarsToFetch, pIncludeBarsOutsideSession)
Else
    lIndex = PriceChart.Add(pTimePeriod, mContract.Specifier.LocalSymbol, True, pNumberOfBarsToFetch, pIncludeBarsOutsideSession)
End If

If mPriceChartIndex = 0 Then
    mPriceChartIndex = lIndex
    Set mPriceChartTimePeriod = pTimePeriod
End If
If mPricePeriods Is Nothing Then Set mPricePeriods = PriceChart.BaseChartController.Periods
If mBracketOrderLineSeries Is Nothing Then Set mBracketOrderLineSeries = PriceChart.BaseChartController.Regions.Item(ChartRegionNamePrice).AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerHighestUser)

Set IStrategyHost_AddTimeframe = PriceChart.Timeframe

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IStrategyHost_ContractInvalid(ByVal pMessage As String)
MsgBox pMessage, vbCritical, "Invalid contract"
StartButton.Enabled = True
StopButton.Enabled = False
End Sub

Private Property Get IStrategyHost_ContractStorePrimary() As IContractStore
Set IStrategyHost_ContractStorePrimary = gTB.ContractStorePrimary
End Property

Private Property Get IStrategyHost_ContractStoreSecondary() As IContractStore
Set IStrategyHost_ContractStoreSecondary = gTB.ContractStoreSecondary
End Property

Private Property Get IStrategyHost_HistoricalDataStoreInput() As IHistoricalDataStore
Set IStrategyHost_HistoricalDataStoreInput = gTB.HistoricalDataStoreInput
End Property

Private Property Get IStrategyHost_LogDummyProfitProfile() As Boolean
IStrategyHost_LogDummyProfitProfile = IIf(DummyProfitProfileCheck = vbChecked, True, False)
End Property

Private Property Get IStrategyHost_LogParameters() As Boolean
IStrategyHost_LogParameters = True
End Property

Private Property Get IStrategyHost_LogProfitProfile() As Boolean
IStrategyHost_LogProfitProfile = IIf(ProfitProfileCheck = vbChecked, True, False)
End Property

Private Sub IStrategyHost_NotifyReplayCompleted()
Const ProcName As String = "IStrategyHost_NotifyReplayCompleted"
On Error GoTo Err

mTotalElapsedSecs = mTotalElapsedSecs + mElapsedSecsCurrTickfile
mElapsedSecsCurrTickfile = 0

mTotalEvents = mTotalEvents + mEventsCurrTickfile
mEventsCurrTickfile = 0

PriceChart.BaseChartController.EnableDrawing

If mTickfileIndex = TickfileOrganiser1.TickFileSpecifiers.Count - 1 Then
    Set mPriceStudyBase = Nothing
    StartButton.Enabled = True
    StopButton.Enabled = False
Else
    mTickfileIndex = mTickfileIndex + 1
    TickfileOrganiser1.ListIndex = mTickfileIndex
    If SeparateSessionsCheck = vbChecked Then
        clearFields
        mStrategyRunner.PrepareTickfile TickfileOrganiser1.TickFileSpecifiers(mTickfileIndex)
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_NotifyReplayEvent(ev As NotificationEventData)
Const ProcName As String = "IStrategyHost_NotifyReplayEvent"
On Error GoTo Err

Dim lMessage As String

Dim lEventCode As TickfileEventCodes
lEventCode = ev.EventCode
Select Case lEventCode
Case TickfileEventFileDoesNotExist
    lMessage = "Tickfile does not exist"
Case TickfileEventFileIsEmpty
    lMessage = "Tickfile is empty"
Case TickfileEventFileIsInvalid
    lMessage = "Tickfile is invalid"
Case TickfileEventFileFormatNotSupported
    lMessage = "Tickfile format is not supported"
Case TickfileEventNoContractDetails
    lMessage = "No contract details are available for this tickfile"
Case TickfileEventDataSourceNotAvailable
    lMessage = "Tickfile data source is not available"
Case TickfileEventAmbiguousContractDetails
    lMessage = "A unique contract for this tickfile cannot be determined"
Case Else
    lMessage = "An unspecified error has occurred"
End Select

If ev.EventMessage <> "" Then lMessage = lMessage & ev.EventMessage

MsgBox lMessage, vbCritical, "Tickfile problem"
StopButton.Enabled = False
StartButton.Enabled = True

mStrategyRunner.StopTesting

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_NotifyReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Single)
Const ProcName As String = "IStrategyHost_NotifyReplayProgress"
On Error GoTo Err

PercentCompleteLabel.Caption = Format(pPercentComplete, "0.0")

mElapsedSecsCurrTickfile = (GetTimestamp - mReplayStartTime) * 86400
Dim lTotalElapsedSecs As Double
lTotalElapsedSecs = mTotalElapsedSecs + mElapsedSecsCurrTickfile

mEventsCurrTickfile = pEventsPlayed
Dim lTotalEvents As Long
lTotalEvents = mTotalEvents + mEventsCurrTickfile

EventsPlayedLabel.Caption = lTotalEvents
EventsPerSecondLabel.Caption = Int(lTotalEvents / lTotalElapsedSecs)
MicrosecsPerEventLabel.Caption = Int(lTotalElapsedSecs * 1000000 / lTotalEvents)

PriceChart.BaseChartController.EnableDrawing
PriceChart.BaseChartController.DisableDrawing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Property Get IStrategyHost_OrderSubmitterFactoryLive() As IOrderSubmitterFactory
Set IStrategyHost_OrderSubmitterFactoryLive = gTB.OrderSubmitterFactoryLive
End Property

Private Property Get IStrategyHost_OrderSubmitterFactorySimulated() As IOrderSubmitterFactory
Set IStrategyHost_OrderSubmitterFactorySimulated = gTB.OrderSubmitterFactorySimulated
End Property

Private Property Get IStrategyHost_RealtimeTickers() As Tickers
Set IStrategyHost_RealtimeTickers = gTB.Tickers
End Property

Private Property Get IStrategyHost_ResultsPath() As String
IStrategyHost_ResultsPath = ResultsPathText
End Property

Private Property Get IStrategyHost_StudyLibraryManager() As StudyLibraryManager
Set IStrategyHost_StudyLibraryManager = gTB.StudyLibraryManager
End Property

Private Sub IStrategyHost_TickerCreated(ByVal pTicker As Ticker)
Const ProcName As String = "IStrategyHost_TickerCreated"
On Error GoTo Err

Set mTicker = pTicker
Set mContract = mTicker.ContractFuture.Value
mSecType = mContract.Specifier.SecType
mTickSize = mContract.TickSize
Set mSession = mTicker.SessionFuture.Value

If mPriceStudyBase Is Nothing Then
    Set mPriceStudyBase = New StudyBaseForTickDataInput
    mPriceStudyBase.Initialise gTB.StudyLibraryManager.CreateStudyManager( _
                                                        mContract.SessionStartTime, _
                                                        mContract.SessionEndTime, _
                                                        GetTimeZone(mContract.TimeZoneName)), _
                                mTicker.ContractFuture
    
    If mTicker.IsTickReplay Then
        mFutureWaiter.Add mTicker.ClockFuture
    Else
        initialisePriceChart
    End If
Else
    mTicker.AddGenericTickListener Me
    mStrategyRunner.StartReplay
    mReplayStartTime = GetTimestamp
    PriceChart.BaseChartController.DisableDrawing
End If
If mProfitStudyBase Is Nothing Then initialiseProfitChart
If mTradeStudyBase Is Nothing Then initialiseTradeChart

Me.Caption = "TradeBuild Strategy Trader - " & _
            StrategyCombo.Text & " - " & _
            mContract.Specifier.LocalSymbol

SSTab2.Tab = 5

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Property Get IStrategyHost_TickfileStoreInput() As ITickfileStore
Set IStrategyHost_TickfileStoreInput = gTB.TickfileStoreInput
End Property

Private Property Get IStrategyHost_UseMoneyManagement() As Boolean
IStrategyHost_UseMoneyManagement = IIf(NoMoneyManagement = vbChecked, False, True)
End Property

'================================================================================
' ILogListener Interface Members
'================================================================================

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal pLogrec As LogRecord)
Const ProcName As String = "ILogListener_Notify"
On Error GoTo Err

Select Case pLogrec.InfoType
Case "strategy.tradereason"
    writeLogText formatLogRecord(pLogrec, False)
Case "position.profit"
    Profit.Caption = Format(pLogrec.Data, "0.00")
    mSessionProfit = pLogrec.Data
    
    mProfitStudyBase.NotifyValue mOverallProfit + mSessionProfit, mTicker.TimeStamp
    
    If Not mTradeBar Is Nothing Then
        mTradeStudyBase.NotifyValue mOverallProfit + mSessionProfit, mTicker.TimeStamp
    End If
Case "position.drawdown"
    Drawdown.Caption = Format(pLogrec.Data, "0.00")
Case "position.maxprofit"
    MaxProfit.Caption = Format(pLogrec.Data, "0.00")
Case "position.bracketorderprofilestruct"
    Dim lListItem As ListItem
    Static sBracketOrderNumber As Long

    Dim lBracketOrderProfile As BracketOrderProfile
    lBracketOrderProfile = pLogrec.Data
    
    Dim lBracketOrderLine As ChartSkil27.Line
    Set lBracketOrderLine = mBracketOrderLineSeries.Add
    lBracketOrderLine.Point1 = NewPoint(mPricePeriods(BarStartTime(lBracketOrderProfile.StartTime, mPriceChartTimePeriod, mContract.SessionStartTime)).PeriodNumber, lBracketOrderProfile.EntryPrice)
    
    Dim lLineEndBarStartTime As Date
    lLineEndBarStartTime = BarStartTime(lBracketOrderProfile.EndTime, mPriceChartTimePeriod, mContract.SessionStartTime)
    
    On Error Resume Next
    Dim lPeriod As Period
    Set lPeriod = mPricePeriods(lLineEndBarStartTime)
    On Error GoTo 0
    If lPeriod Is Nothing Then
        ' this occurs when the execution that finished the order plex occurred
        ' at the start of a new bar but before the first price for the bar
        ' was reported. So add the bar now
        mPricePeriods.Add lLineEndBarStartTime
    End If
    lBracketOrderLine.Point2 = NewPoint(mPricePeriods(lLineEndBarStartTime).PeriodNumber, lBracketOrderProfile.ExitPrice)
    
    If lBracketOrderProfile.Action = OrderActionBuy Then
        lBracketOrderLine.Color = vbBlue
    Else
        lBracketOrderLine.Color = vbRed
    End If
    'If lBracketOrderProfile.QuantityOutstanding <> 0 Then
        lBracketOrderLine.ArrowEndStyle = ArrowClosed
        lBracketOrderLine.ArrowEndWidth = 8
        lBracketOrderLine.ArrowEndLength = 12
    'End If
    
    sBracketOrderNumber = sBracketOrderNumber + 1
    Set lListItem = BracketOrderList.ListItems.Add(, "K" & sBracketOrderNumber, lBracketOrderProfile.Key)
    lListItem.SubItems(BOListColumns.ColumnAction - 1) = IIf(lBracketOrderProfile.Action = OrderActionBuy, "BUY", "SELL")
    'lListItem.SubItems(BOListColumns.ColumnClosedOut - 1) = IIf(lBracketOrderProfile.closedOut, "Y", "")
    lListItem.SubItems(BOListColumns.ColumnDescription - 1) = lBracketOrderProfile.Description
    lListItem.SubItems(BOListColumns.ColumnEndTime - 1) = FormatDateTime(lBracketOrderProfile.EndTime, vbGeneralDate)
    lListItem.SubItems(BOListColumns.ColumnEntryPrice - 1) = FormatPrice(lBracketOrderProfile.EntryPrice, mSecType, mTickSize)
    lListItem.SubItems(BOListColumns.ColumnEntryReason - 1) = lBracketOrderProfile.EntryReason
    lListItem.SubItems(BOListColumns.ColumnExitPrice - 1) = FormatPrice(lBracketOrderProfile.ExitPrice, mSecType, mTickSize)
    lListItem.SubItems(BOListColumns.ColumnMaxLoss - 1) = lBracketOrderProfile.MaxLoss
    lListItem.SubItems(BOListColumns.ColumnMaxProfit - 1) = lBracketOrderProfile.MaxProfit
    lListItem.SubItems(BOListColumns.ColumnProfit - 1) = lBracketOrderProfile.Profit
    lListItem.SubItems(BOListColumns.ColumnQuantity - 1) = lBracketOrderProfile.Quantity
    'lListItem.SubItems(BOListColumns.ColumnQuantityOutstanding - 1) = IIf(lBracketOrderProfile.QuantityOutstanding <> 0, lBracketOrderProfile.QuantityOutstanding, "")
    lListItem.SubItems(BOListColumns.ColumnRisk - 1) = lBracketOrderProfile.Risk
    lListItem.SubItems(BOListColumns.ColumnStartTime - 1) = FormatDateTime(lBracketOrderProfile.StartTime, vbGeneralDate)
    lListItem.SubItems(BOListColumns.ColumnStopReason - 1) = lBracketOrderProfile.StopReason
    lListItem.SubItems(BOListColumns.ColumnTargetReason - 1) = lBracketOrderProfile.TargetReason
Case "position.position"
    Dim lNewPosition As Long
    lNewPosition = pLogrec.Data
    If (mPosition = 0 And lNewPosition <> 0) Or _
        (mPosition > 0 And lNewPosition <= 0) Or _
        (mPosition < 0 And lNewPosition >= 0) _
    Then
        mTradeBarNumber = mTradeBarNumber + 1
        mTradeStudyBase.NotifyBarNumber mTradeBarNumber, mTicker.TimeStamp
        mTradeStudyBase.NotifyValue mOverallProfit + mSessionProfit, mTicker.TimeStamp
    End If
    mPosition = lNewPosition
    Position.Caption = mPosition
Case "position.order", _
    "position.moneymanagement"
    LogMessage CStr(pLogrec.Data)
    writeLogText formatLogRecord(pLogrec, False)
Case "position.ordersimulated", _
    "position.moneymanagementsimulated"
    LogMessage CStr(pLogrec.Data)
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is Clock Then
    Dim lClock As Clock
    Set lClock = ev.Future.Value
    initialisePriceChart lClock.TimeStamp
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mSession Event Handlers
'================================================================================

Private Sub mSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mSession_SessionStarted"
On Error GoTo Err

mProfitStudyBase.NotifyValue mOverallProfit, mTicker.TimeStamp

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mTradeBarsBuilder Event Handlers
'================================================================================

Private Sub mTradeBarsBuilder_BarAdded(ByVal pBar As BarUtils27.Bar)
Const ProcName As String = "mTradeBarsBuilder_BarAdded"
On Error GoTo Err

Set mTradeBar = pBar

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub BracketOrderList_ColumnClick(ByVal ColumnHeader As ColumnHeader)
Const ProcName As String = "BracketOrderList_ColumnClick"
On Error GoTo Err

If BracketOrderList.SortKey = ColumnHeader.Index - 1 Then
    BracketOrderList.SortOrder = 1 - BracketOrderList.SortOrder
Else
    BracketOrderList.SortKey = ColumnHeader.Index - 1
    BracketOrderList.SortOrder = lvwAscending
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BracketOrderList_DblClick()
Const ProcName As String = "BracketOrderList_DblClick"
On Error GoTo Err

Dim ListItem As ListItem
Set ListItem = BracketOrderList.SelectedItem

Dim PeriodNumber As Long
PeriodNumber = mPricePeriods(BarStartTime(CDate(ListItem.SubItems(BOListColumns.ColumnStartTime - 1)), mPriceChartTimePeriod, mContract.SessionStartTime)).PeriodNumber
PriceChart.BaseChartController.LastVisiblePeriod = _
            PeriodNumber + _
            Int((PriceChart.BaseChartController.LastVisiblePeriod - _
            PriceChart.BaseChartController.FirstVisiblePeriod) / 2) - 1
SSTab1.Tab = 0

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'Private Sub ListenerFilenameText_Change()
'Dim fileListener As cFileListener
'Set fileListener = mStrategyListener
'fileListener.FileName = ListenerFilenameText
'End Sub

'Private Sub ListenerTypeCombo_Click()
'Select Case ListenerTypeCombo
'Case "File listener"
'    Dim fileListener As cFileListener
'    Set fileListener = New cFileListener
'    Set mStrategyListener = fileListener
'    fileListener.Overwrite = True
'    SelectListenerFileButton.Enabled = True
'    ListenerFilenameText.Enabled = True
'    RawDataCheck.Enabled = True
'    fileListener.FileName = ListenerFilenameText
'    fileListener.raw = IIf(RawDataCheck = vbChecked, True, False)
'Case "Swing writer"
'    Dim swingWriter As RLKSwingWriter
'    Set swingWriter = New RLKSwingWriter
'    Set mStrategyListener = swingWriter
'    SelectListenerFileButton.Enabled = False
'    ListenerFilenameText.Enabled = False
'    RawDataCheck.Enabled = False
'End Select
'End Sub

Private Sub MoreButton_Click()
Const ProcName As String = "MoreButton_Click"
On Error GoTo Err

If mDetailsHidden Then
    mDetailsHidden = False
    MoreButton.Caption = "Less <<<"
    Me.Height = 8985 + Me.Height - Me.ScaleHeight
Else
    mDetailsHidden = True
    MoreButton.Caption = "More >>>"
    Me.Height = minimumHeight + Me.Height - Me.ScaleHeight
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub PriceChart_ChartStateChanged(ByVal Index As Long, ev As TWUtilities40.StateChangeEventData)
Const ProcName As String = "PriceChart_ChartStateChanged"
On Error GoTo Err

If Index <> mPriceChartIndex Then Exit Sub

If ev.State <> ChartStates.ChartStateLoaded Then Exit Sub

mTicker.AddGenericTickListener Me

If mTicker.IsTickReplay Then
    mStrategyRunner.StartReplay
    mReplayStartTime = GetTimestamp
    PriceChart.BaseChartController.DisableDrawing
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'Private Sub RawDataCheck_Click()
'Dim fileListener As cFileListener
'Set fileListener = mStrategyListener
'fileListener.raw = IIf(RawDataCheck = vbChecked, True, False)
'End Sub

Private Sub ResultsPathButton_Click()
Const ProcName As String = "ResultsPathButton_Click"
On Error GoTo Err

ResultsPathText.Text = ChoosePath(ApplicationSettingsFolder & "Results")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'Private Sub SelectListenerFileButton_Click()
'CommonDialogs.CancelError = True
'On Error GoTo Err
'
'CommonDialogs.MaxFileSize = 32767
'CommonDialogs.DefaultExt = ".log"
'CommonDialogs.DialogTitle = "Save listener data to"
'CommonDialogs.Filter = "Text (*.log)|*.log|All files (*.*)|*.*"
'CommonDialogs.FilterIndex = 1
'CommonDialogs.Flags = cdlOFNLongNames + _
'                    cdlOFNPathMustExist + _
'                    cdlOFNExplorer
'CommonDialogs.ShowOpen
'
'ListenerFilenameText = CommonDialogs.FileName
'
'Exit Sub
'Err:
'
'End Sub

Private Sub StartButton_Click()
Const ProcName As String = "StartButton_Click"
On Error GoTo Err

StartButton.Enabled = False
StopButton.Enabled = True

PriceChart.Clear
ProfitChart.BaseChartController.ClearChart
TradeChart.BaseChartController.ClearChart
BracketOrderList.ListItems.Clear

mCurrTickfileIndex = -1

clearFields

mOverallProfit = 0#
mSessionProfit = 0#
Set mTradeBar = Nothing

If TickfileOrganiser1.TickfileCount <> 0 Then
    TickfileOrganiser1.ListIndex = 0
    mStrategyRunner.PrepareTickfile TickfileOrganiser1.TickFileSpecifiers(1)
Else
    mStrategyRunner.PrepareSymbol SymbolText.Text
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StopButton_Click()
Const ProcName As String = "StopButton_Click"
On Error GoTo Err

mStrategyRunner.StopTesting
StartButton.Enabled = True
StopButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StopStrategyFactoryCombo_Change()
Const ProcName As String = "StopStrategyFactoryCombo_Change"
On Error GoTo Err

SetStrategy

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StopStrategyFactoryCombo_Click()
Const ProcName As String = "StopStrategyFactoryCombo_Click"
On Error GoTo Err

SetStrategy

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StrategyCombo_Change()
Const ProcName As String = "StrategyCombo_Change"
On Error GoTo Err

SetStrategy

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StrategyCombo_Click()
Const ProcName As String = "StrategyCombo_Click"
On Error GoTo Err

SetStrategy

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
gHandleUnexpectedError ProcName, ModuleName
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

Private Sub clearFields()
BidText = ""
BidSizeText = ""
AskText = ""
AskSizeText = ""
TradeText = ""
TradeSizeText = ""
Profit.Caption = ""
Drawdown.Caption = ""
MaxProfit.Caption = ""
Position.Caption = ""
End Sub

Private Function formatLogRecord(ByVal pLogrec As LogRecord, ByVal pIncludeTimestamp As Boolean) As String
Const ProcName As String = "formatLogRecord"
On Error GoTo Err

Static formatter As ILogFormatter
Static formatterWithTimestamp As ILogFormatter

If pIncludeTimestamp Then
    If formatterWithTimestamp Is Nothing Then Set formatterWithTimestamp = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal, , True, False)
    formatLogRecord = formatterWithTimestamp.FormatRecord(pLogrec)
Else
    If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(, , False, False)
    formatLogRecord = formatter.FormatRecord(pLogrec)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initialisePriceChart(Optional ByVal pTimestamp As Date)
Const ProcName As String = "initialisePriceChart"
On Error GoTo Err

PriceChart.Initialise CreateTimeframes(mPriceStudyBase, _
                                    mTicker.ContractFuture, _
                                    gTB.HistoricalDataStoreInput, _
                                    mTicker.ClockFuture), _
                    gTB.HistoricalDataStoreInput.TimePeriodValidator, _
                    CreateChartSpecifier(200, , , pTimestamp), _
                    ChartStylesManager.DefaultStyle

mStrategyRunner.StartStrategy mParams

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseProfitChart()
Const ProcName As String = "initialiseProfitChart"
On Error GoTo Err

Set mProfitStudyBase = CreateStudyBaseForDoubleInput( _
                                    gTB.StudyLibraryManager.CreateStudyManager( _
                                                    mContract.SessionStartTime, _
                                                    mContract.SessionEndTime, _
                                                    GetTimeZone(mContract.TimeZoneName)))
ProfitChart.ShowChart CreateTimeframes(mProfitStudyBase), _
                        GetTimePeriod(1, TimePeriodDay), _
                        CreateChartSpecifier(0), _
                        ChartStylesManager.DefaultStyle, _
                        pTitle:="Profit by Session"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseTradeChart()
Const ProcName As String = "initialiseTradeChart"
On Error GoTo Err

Set mTradeStudyBase = CreateStudyBaseForUserBarsInput( _
                                    gTB.StudyLibraryManager.CreateStudyManager( _
                                                    mContract.SessionStartTime, _
                                                    mContract.SessionEndTime, _
                                                    GetTimeZone(mContract.TimeZoneName)))
TradeChart.ShowChart CreateTimeframes(mTradeStudyBase), _
                        GetTimePeriod(0, TimePeriodNone), _
                        CreateChartSpecifier(0), _
                        ChartStylesManager.DefaultStyle, _
                        pTitle:="Profit by Trade"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function minimumHeight() As Long
minimumHeight = SSTab2.Top + SSTab2.Height
End Function

Private Sub SetStrategy()
Const ProcName As String = "SetStrategy"
On Error GoTo Err

If StrategyCombo.Text = "" Then Exit Sub
If StopStrategyFactoryCombo.Text = "" Then Exit Sub

Set mStrategyRunner = CreateStrategyRunner(Me)
Set mParams = mStrategyRunner.SetStrategy(CreateObject(StrategyCombo.Text), CreateObject(StopStrategyFactoryCombo.Text))

Set ParamGrid.DataSource = mParams
ParamGrid.Columns(0).Width = ParamGrid.Width / 2
ParamGrid.Columns(1).Width = ParamGrid.Width / 2

StartButton.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupBracketOrderList()
Const ProcName As String = "setupBracketOrderList"
On Error GoTo Err

Dim pBOListWidth As Long

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnKey, , "Key"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnKey).Width = _
    BOListColumnWidths.WidthKey * BracketOrderList.Width / 100
pBOListWidth = BracketOrderList.ColumnHeaders(BOListColumns.ColumnKey).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnKey).Alignment = lvwColumnLeft

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnStartTime, , "Start time"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnStartTime).Width = _
    BOListColumnWidths.WidthStartTime * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnStartTime).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnStartTime).Alignment = lvwColumnLeft

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnEndTime, , "End time"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnEndTime).Width = _
    BOListColumnWidths.WidthEndTime * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnEndTime).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnEndTime).Alignment = lvwColumnLeft

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnAction, , "Action"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnAction).Width = _
    BOListColumnWidths.WidthAction * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnAction).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnAction).Alignment = lvwColumnLeft

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnQuantity, , "Qty"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnQuantity).Width = _
    BOListColumnWidths.WidthQuantity * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnQuantity).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnQuantity).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnEntryPrice, , "Entry"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnEntryPrice).Width = _
    BOListColumnWidths.WidthExitPrice * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnEntryPrice).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnEntryPrice).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnExitPrice, , "Exit"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnExitPrice).Width = _
    BOListColumnWidths.WidthExitPrice * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnExitPrice).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnExitPrice).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnProfit, , "Profit"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnProfit).Width = _
    BOListColumnWidths.WidthProfit * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnProfit).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnProfit).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnMaxProfit, , "Max profit"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnMaxProfit).Width = _
    BOListColumnWidths.WidthMaxProfit * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnMaxProfit).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnMaxProfit).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnMaxLoss, , "Max loss"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnMaxLoss).Width = _
    BOListColumnWidths.WidthMaxLoss * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnMaxLoss).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnMaxLoss).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnRisk, , "Risk"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnRisk).Width = _
    BOListColumnWidths.WidthRisk * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnRisk).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnRisk).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnQuantityOutstanding, , "OQty"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnQuantityOutstanding).Width = _
    BOListColumnWidths.WidthQuantityOutstanding * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnQuantityOutstanding).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnQuantityOutstanding).Alignment = lvwColumnRight

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnEntryReason, , "Entry reason"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnEntryReason).Width = _
    BOListColumnWidths.WidthEntryReason * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnEntryReason).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnEntryReason).Alignment = lvwColumnLeft

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnTargetReason, , "Target reason"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnTargetReason).Width = _
    BOListColumnWidths.WidthTargetReason * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnTargetReason).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnTargetReason).Alignment = lvwColumnLeft

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnStopReason, , "Stop reason"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnStopReason).Width = _
    BOListColumnWidths.WidthStopReason * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnStopReason).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnStopReason).Alignment = lvwColumnLeft

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnClosedOut, , "Closed out"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnClosedOut).Width = _
    BOListColumnWidths.WidthClosedOut * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnClosedOut).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnClosedOut).Alignment = lvwColumnCenter

BracketOrderList.ColumnHeaders.Add BOListColumns.ColumnDescription, , "Description"
BracketOrderList.ColumnHeaders(BOListColumns.ColumnDescription).Width = _
    BOListColumnWidths.WidthDescription * BracketOrderList.Width / 100
pBOListWidth = pBOListWidth + BracketOrderList.ColumnHeaders(BOListColumns.ColumnDescription).Width
BracketOrderList.ColumnHeaders(BOListColumns.ColumnDescription).Alignment = lvwColumnLeft

If Me.ScaleMode = vbTwips Then
    ' If using Twips then change to pixels
    pBOListWidth = pBOListWidth / Screen.TwipsPerPixelX
End If
SendMessage BracketOrderList.hWnd, LB_SETHORZEXTENT, pBOListWidth, 0

BracketOrderList.Sorted = True
BracketOrderList.SortKey = BOListColumns.ColumnEndTime - 1
BracketOrderList.SortOrder = lvwDescending

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupLogging()
Const ProcName As String = "setupLogging"
On Error GoTo Err

GetLogger("log").AddLogListener Me
GetLogger("position.profit").AddLogListener Me
GetLogger("position.drawdown").AddLogListener Me
GetLogger("position.maxprofit").AddLogListener Me
GetLogger("position.bracketorderprofilestruct").AddLogListener Me
GetLogger("position.position").AddLogListener Me
GetLogger("position.order").AddLogListener Me
GetLogger("position.ordersimulated").AddLogListener Me
GetLogger("position.moneymanagement").AddLogListener Me
GetLogger("position.moneymanagementsimulated").AddLogListener Me
GetLogger("strategy.tradereason").AddLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub writeLogText(ByVal pMessage As String)
Const ProcName As String = "writeLogText"
On Error GoTo Err

Dim lBytesNeeded As Long

lBytesNeeded = Len(LogText.Text) + Len(pMessage) - 32767
If lBytesNeeded > 0 Then
    ' clear some space at the start of the textbox
    LogText.SelStart = 0
    LogText.SelLength = 4 * lBytesNeeded
    LogText.SelText = ""
End If

LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = pMessage
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



