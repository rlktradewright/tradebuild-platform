VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#313.1#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.Form fStrategyHost 
   Caption         =   "TradeBuild Strategy Host v2.7"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab2 
      Height          =   3675
      Left            =   0
      TabIndex        =   14
      Top             =   -30
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   6482
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   15246432
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Controls"
      TabPicture(0)   =   "fStrategyHost.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Parameters"
      TabPicture(1)   =   "fStrategyHost.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Log"
      TabPicture(2)   =   "fStrategyHost.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LogPicture"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Results"
      TabPicture(3)   =   "fStrategyHost.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ResultsPicture"
      Tab(3).ControlCount=   1
      Begin VB.PictureBox ResultsPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -75000
         ScaleHeight     =   3375
         ScaleWidth      =   11070
         TabIndex        =   26
         Top             =   300
         Width           =   11070
         Begin TWControls40.TWButton MoreButton 
            Height          =   375
            Left            =   6480
            TabIndex        =   27
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Less <<<"
            DefaultBorderColor=   15793920
            DisabledBackColor=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseOverBackColor=   0
            PushedBackColor =   0
         End
         Begin VB.Label TheTime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1065
            TabIndex        =   55
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Position"
            Height          =   195
            Left            =   3600
            TabIndex        =   54
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Position 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   53
            Top             =   720
            Width           =   855
         End
         Begin VB.Label MaxProfit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   52
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Max profit"
            Height          =   195
            Left            =   3600
            TabIndex        =   51
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Drawdown"
            Height          =   195
            Left            =   3600
            TabIndex        =   50
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Drawdown 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   49
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Profit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   48
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Profit/Loss"
            Height          =   195
            Left            =   3600
            TabIndex        =   47
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Ask"
            Height          =   195
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Last"
            Height          =   195
            Left            =   0
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Bid"
            Height          =   195
            Left            =   0
            TabIndex        =   44
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Events played"
            Height          =   195
            Left            =   3600
            TabIndex        =   43
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label EventsPlayedLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   42
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Percent complete"
            Height          =   195
            Left            =   3600
            TabIndex        =   41
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label PercentCompleteLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   40
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Events per second"
            Height          =   195
            Left            =   3600
            TabIndex        =   39
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label EventsPerSecondLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   38
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label MicrosecsPerEventLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   37
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Microsecs per event"
            Height          =   195
            Left            =   3600
            TabIndex        =   36
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label AskLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   35
            Top             =   0
            Width           =   735
         End
         Begin VB.Label AskSizeLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   34
            Top             =   0
            Width           =   735
         End
         Begin VB.Label TradeLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.Label TradeSizeLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
         Begin VB.Label BidLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   31
            Top             =   480
            Width           =   735
         End
         Begin VB.Label BidSizeLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   30
            Top             =   480
            Width           =   735
         End
         Begin VB.Label VolumeLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   29
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Volume"
            Height          =   195
            Left            =   0
            TabIndex        =   28
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.PictureBox LogPicture 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -75000
         ScaleHeight     =   3375
         ScaleWidth      =   11070
         TabIndex        =   21
         Top             =   300
         Width           =   11070
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
            Height          =   3360
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   22
            Top             =   0
            Width           =   10935
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   0
         ScaleHeight     =   3375
         ScaleWidth      =   11070
         TabIndex        =   18
         Top             =   300
         Width           =   11070
         Begin TWControls40.TWImageCombo StopStrategyFactoryCombo 
            Height          =   330
            Left            =   6240
            TabIndex        =   3
            Top             =   600
            Width           =   3495
            _ExtentX        =   6165
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
            MouseIcon       =   "fStrategyHost.frx":0070
            Text            =   ""
         End
         Begin TWControls40.TWImageCombo StrategyCombo 
            Height          =   330
            Left            =   6240
            TabIndex        =   2
            Top             =   120
            Width           =   3495
            _ExtentX        =   6165
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
            MouseIcon       =   "fStrategyHost.frx":008C
            Text            =   ""
         End
         Begin TWControls40.TWButton ResultsPathButton 
            Height          =   285
            Left            =   10080
            TabIndex        =   10
            ToolTipText     =   "Click to select results path"
            Top             =   2040
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "..."
            DefaultBorderColor=   15793920
            DisabledBackColor=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseOverBackColor=   0
            PushedBackColor =   0
         End
         Begin TWControls40.TWButton StopButton 
            Height          =   375
            Left            =   9360
            TabIndex        =   12
            Top             =   2880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Stop"
            DefaultBorderColor=   15793920
            DisabledBackColor=   0
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
            MouseOverBackColor=   0
            PushedBackColor =   0
         End
         Begin TWControls40.TWButton StartButton 
            Default         =   -1  'True
            Height          =   375
            Left            =   9360
            TabIndex        =   11
            Top             =   2430
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Start"
            DefaultBorderColor=   15793920
            DisabledBackColor=   0
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
            MouseOverBackColor=   0
            PushedBackColor =   0
         End
         Begin TradingUI27.TickfileOrganiser TickfileOrganiser1 
            Height          =   2535
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4471
            Enabled         =   0   'False
         End
         Begin VB.CheckBox ShowChartCheck 
            Caption         =   "Show chart"
            Height          =   195
            Left            =   6240
            TabIndex        =   4
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox SymbolText 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            TabIndex        =   0
            Top             =   120
            Width           =   1815
         End
         Begin VB.CheckBox DummyProfitProfileCheck 
            Caption         =   "Dummy profit profile"
            Height          =   195
            Left            =   6240
            TabIndex        =   6
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CheckBox ProfitProfileCheck 
            Caption         =   "Profit profile"
            Height          =   195
            Left            =   6240
            TabIndex        =   5
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox NoMoneyManagementCheck 
            Caption         =   "No money management"
            Height          =   195
            Left            =   8280
            TabIndex        =   8
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox SeparateSessionsCheck 
            Caption         =   "Separate session per tick file"
            Height          =   195
            Left            =   8280
            TabIndex        =   7
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.TextBox ResultsPathText 
            Height          =   285
            Left            =   7200
            TabIndex        =   9
            Top             =   2040
            Width           =   2835
         End
         Begin VB.Label Label 
            Caption         =   "Symbol"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Results path"
            Height          =   255
            Left            =   6240
            TabIndex        =   19
            Top             =   2040
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -75000
         ScaleHeight     =   3375
         ScaleWidth      =   11070
         TabIndex        =   15
         Top             =   300
         Width           =   11070
         Begin MSDataGridLib.DataGrid ParamGrid 
            Height          =   2985
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   5265
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   16
      Top             =   3600
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   15246432
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Price chart"
      TabPicture(0)   =   "fStrategyHost.frx":00A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PriceChart"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Daily profit chart"
      TabPicture(1)   =   "fStrategyHost.frx":00C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ProfitChart"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Trade chart"
      TabPicture(2)   =   "fStrategyHost.frx":00E0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TradeChart"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Bracket order details"
      TabPicture(3)   =   "fStrategyHost.frx":00FC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "BracketOrderList"
      Tab(3).ControlCount=   1
      Begin TradingUI27.MarketChart ProfitChart 
         Height          =   5325
         Left            =   -75000
         TabIndex        =   24
         Top             =   300
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9393
      End
      Begin TradingUI27.MultiChart PriceChart 
         Height          =   5325
         Left            =   0
         TabIndex        =   23
         Top             =   300
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9393
      End
      Begin MSComctlLib.ListView BracketOrderList 
         Height          =   5325
         Left            =   -75000
         TabIndex        =   17
         Top             =   300
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   9393
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
      Begin TradingUI27.MarketChart TradeChart 
         Height          =   5325
         Left            =   -75000
         TabIndex        =   25
         Top             =   300
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9393
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
' Interfaces
'================================================================================

Implements IStrategyHostView

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

Private mModel                                          As IStrategyHostModel
Private mController                                     As IStrategyHostController

Private mContract                                       As IContract
Private mSecType                                        As SecurityTypes
Private mTickSize                                       As Double

Private WithEvents mSession                             As Session
Attribute mSession.VB_VarHelpID = -1
Private mTradingSessionInProgress                       As Boolean

Private mParams                                         As Parameters

Private mProfitStudyBase                                As StudyBaseForDoubleInput

Private mPriceChartTimePeriod                           As TimePeriod

Private mTradeStudyBase                                 As StudyBaseForIntegerInput

Private mPosition                                       As Long
Private mOverallProfit                                  As Double
Private mSessionProfit                                  As Double
Private mMaxProfit                                      As Double
Private mDrawdown                                       As Double

Private mDetailsHidden                                  As Boolean

Private mBracketOrderLineSeries                         As LineSeries

Private mPricePeriods                                   As Periods

Private mTheme                                          As ITheme

Private mChartStyle                                     As ChartStyle

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitialiseCommonControls
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

Me.ScaleMode = vbTwips
setupBracketOrderList
Set mChartStyle = gCreateChartStyle
LogMessage "Setting StrategyCombo"
StrategyCombo.ComboItems.Add , , "Strategies27.MACDStrategy21"
LogMessage "Setting StopStrategyFactoryCombo"
StopStrategyFactoryCombo.ComboItems.Add , , "Strategies27.StopStrategyFactory5"
LogMessage "Form loaded"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
On Error GoTo Err

SSTab1.Width = ScaleWidth
SSTab2.Width = ScaleWidth

If ScaleHeight < minimumHeight Or mDetailsHidden Then
    Me.Height = minimumHeight + (Me.Height - Me.ScaleHeight)
    Exit Sub
End If

Picture1.Width = SSTab2.Width
Picture4.Width = SSTab2.Width
LogPicture.Width = SSTab2.Width
ResultsPicture.Width = SSTab2.Width

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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

LogMessage "Unloading main form"

If mModel.ShowChart Then
    LogMessage "Finishing charts"
    PriceChart.Finish
    ProfitChart.Finish
    TradeChart.Finish
End If

LogMessage "Closing other forms"
Dim f As Form
For Each f In Forms
    If Not TypeOf f Is fStrategyHost Then
        LogMessage "Closing form: caption=" & f.Caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

gFinished = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' IStrategyHostView Interface Members
'================================================================================

Public Sub IStrategyHostView_AddStudyToChart( _
                ByVal pChartIndex As Long, _
                ByVal pStudy As IStudy, _
                ByVal pStudyValueNames As EnumerableCollection)
Const ProcName As String = "IStrategyHostView_AddStudyToChart"
On Error GoTo Err

Dim lChartManager As ChartManager
Set lChartManager = PriceChart.ChartManager(pChartIndex)

Dim lStudyConfig As StudyConfiguration
Set lStudyConfig = lChartManager.GetDefaultStudyConfiguration(pStudy.Name, pStudy.LibraryName)
lStudyConfig.Study = pStudy
lStudyConfig.UnderlyingStudy = pStudy.UnderlyingStudy

Assert Not lStudyConfig Is Nothing, "Can't get default study configuration"

Dim lSvc As StudyValueConfiguration
For Each lSvc In lStudyConfig.StudyValueConfigurations
    If pStudyValueNames.Contains(lSvc.ValueName) Then
        lSvc.IncludeInChart = True
    Else
        lSvc.IncludeInChart = False
    End If
Next

lChartManager.ApplyStudyConfiguration lStudyConfig, ReplayNumbers.ReplayAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IStrategyHostView_AddTimeframe( _
                ByVal pTimeframe As Timeframe) As Long
Const ProcName As String = "IStrategyHostView_AddTimeframe"
On Error GoTo Err

Dim lTitle As String
Dim lUpdatePerTick As Boolean

If mModel.IsTickReplay Then
    lTitle = ""
    lUpdatePerTick = False
Else
    lTitle = mModel.Contract.Specifier.LocalSymbol
    lUpdatePerTick = True
End If

Dim lIndex As Long
lIndex = PriceChart.AddRaw(pTimeframe, _
                        mModel.Ticker.StudyBase.StudyManager, _
                        mModel.Contract.Specifier.LocalSymbol, _
                        mModel.Contract.Specifier.SecType, _
                        mModel.Contract.Specifier.Exchange, _
                        mModel.Contract.TickSize, _
                        mModel.Contract.SessionStartTime, _
                        mModel.Contract.SessionEndTime, _
                        lTitle, _
                        lUpdatePerTick)

Dim lChartManager As ChartManager
Set lChartManager = PriceChart.ChartManager(lIndex)
lChartManager.SetBaseStudyConfiguration CreateBarsStudyConfig(pTimeframe, mModel.Contract.Specifier.SecType, mModel.Ticker.StudyBase.StudyManager.StudyLibraryManager), 0

If mPriceChartTimePeriod Is Nothing Then
    Set mPriceChartTimePeriod = pTimeframe.TimePeriod
    Set mPricePeriods = PriceChart.BaseChartController.Periods
    Set mBracketOrderLineSeries = PriceChart.BaseChartController.Regions.Item(ChartRegionNamePrice).AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerHighestUser)
    mBracketOrderLineSeries.Thickness = 2
    mBracketOrderLineSeries.ArrowEndStyle = ArrowClosed
    mBracketOrderLineSeries.ArrowEndWidth = 8
    mBracketOrderLineSeries.ArrowEndLength = 12
End If

IStrategyHostView_AddTimeframe = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IStrategyHostView_ClearPriceAndProfitFields()
ClearPriceAndProfitFields
End Sub

Private Sub IStrategyHostView_DisablePriceDrawing(Optional ByVal pTimeframeIndex As Long)
Const ProcName As String = "IStrategyHostView_DisablePriceDrawing"
On Error GoTo Err

If pTimeframeIndex = 0 Then
    Dim i As Long
    For i = 1 To PriceChart.Count
        PriceChart.BaseChartController(i).DisableDrawing
    Next
Else
    PriceChart.BaseChartController(pTimeframeIndex).DisableDrawing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_DisableProfitDrawing()
Const ProcName As String = "IStrategyHostView_DisableProfitDrawing"
On Error GoTo Err

ProfitChart.DisableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_DisableStart()
StartButton.Enabled = False
StopButton.Enabled = True
End Sub

Private Sub IStrategyHostView_DisableTradeDrawing()

Const ProcName As String = "IStrategyHostView_DisableTradeDrawing"
On Error GoTo Err

TradeChart.DisableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_EnablePriceDrawing(Optional ByVal pTimeframeIndex As Long)
Const ProcName As String = "IStrategyHostView_EnablePriceDrawing"
On Error GoTo Err

If pTimeframeIndex = 0 Then
    Dim i As Long
    For i = 1 To PriceChart.Count
        PriceChart.BaseChartController(i).EnableDrawing
    Next
Else
    PriceChart.BaseChartController(pTimeframeIndex).EnableDrawing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_EnableProfitDrawing()
Const ProcName As String = "IStrategyHostView_EnableProfitDrawing"
On Error GoTo Err

ProfitChart.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_EnableStart()
StartButton.Enabled = True
StopButton.Enabled = False
End Sub

Private Sub IStrategyHostView_EnableTradeDrawing()
Const ProcName As String = "IStrategyHostView_EnableTradeDrawing"
On Error GoTo Err

TradeChart.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifyBracketOrderProfile(ByRef Value As BracketOrderProfile)
Const ProcName As String = "IStrategyHostView_NotifyBracketOrderProfile"
On Error GoTo Err

Dim lListItem As ListItem
Static sBracketOrderNumber As Long

sBracketOrderNumber = sBracketOrderNumber + 1
Set lListItem = BracketOrderList.ListItems.Add(, "K" & sBracketOrderNumber, Value.Key)
lListItem.SubItems(BOListColumns.ColumnAction - 1) = IIf(Value.Action = OrderActionBuy, "BUY", "SELL")
'lListItem.SubItems(BOListColumns.ColumnClosedOut - 1) = IIf(Value.closedOut, "Y", "")
lListItem.SubItems(BOListColumns.ColumnDescription - 1) = Value.Description
lListItem.SubItems(BOListColumns.ColumnEndTime - 1) = FormatDateTime(Value.EndTime, vbGeneralDate)
lListItem.SubItems(BOListColumns.ColumnEntryPrice - 1) = FormatPrice(Value.EntryPrice, mSecType, mTickSize)
lListItem.SubItems(BOListColumns.ColumnEntryReason - 1) = Value.EntryReason
lListItem.SubItems(BOListColumns.ColumnExitPrice - 1) = FormatPrice(Value.ExitPrice, mSecType, mTickSize)
lListItem.SubItems(BOListColumns.ColumnMaxLoss - 1) = Value.MaxLoss
lListItem.SubItems(BOListColumns.ColumnMaxProfit - 1) = Value.MaxProfit
lListItem.SubItems(BOListColumns.ColumnProfit - 1) = Value.Profit
lListItem.SubItems(BOListColumns.ColumnQuantity - 1) = Value.Quantity
'lListItem.SubItems(BOListColumns.ColumnQuantityOutstanding - 1) = IIf(Value.QuantityOutstanding <> 0, Value.QuantityOutstanding, "")
lListItem.SubItems(BOListColumns.ColumnRisk - 1) = Value.Risk
lListItem.SubItems(BOListColumns.ColumnStartTime - 1) = FormatDateTime(Value.StartTime, vbGeneralDate)
lListItem.SubItems(BOListColumns.ColumnStopReason - 1) = Value.StopReason
lListItem.SubItems(BOListColumns.ColumnTargetReason - 1) = Value.TargetReason

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifyError(ByVal pTitle As String, ByVal pMessage As String, ByVal pSeverity As ErrorSeverities)
Select Case pSeverity
Case ErrorSeverityInformation
    ModelessMsgBox pMessage, MsgBoxInformation, pTitle, Me
Case ErrorSeverityWarning
    MsgBox pMessage, MsgBoxExclamation, pTitle
Case ErrorSeverityCritical
    MsgBox pMessage, MsgBoxCritical, pTitle
End Select
End Sub

Private Sub IStrategyHostView_NotifyEventsPerSecond(ByVal Value As Long)
EventsPerSecondLabel.Caption = Value
End Sub

Private Sub IStrategyHostView_NotifyEventsPlayed(ByVal Value As Long)
EventsPlayedLabel.Caption = Value
End Sub

Private Sub IStrategyHostView_NotifyMicrosecsPerEvent(ByVal Value As Long)
MicrosecsPerEventLabel.Caption = Value
End Sub

Private Sub IStrategyHostView_NotifyNewTradeBar(ByVal pBarNumber As Long, ByVal pTimestamp As Date)
Const ProcName As String = "IStrategyHostView_NotifyNewTradeBar"
On Error GoTo Err

If mModel.ShowChart Then
    mTradeStudyBase.NotifyBarNumber pBarNumber, pTimestamp
    mTradeStudyBase.NotifyValue mOverallProfit + mSessionProfit, pTimestamp
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifyPosition(ByVal Value As Long)
mPosition = Value
Position.Caption = mPosition
End Sub

Private Sub IStrategyHostView_NotifyReplayProgress(ByVal pTickfileTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Single)
Const ProcName As String = "IStrategyHostView_NotifyReplayProgress"
On Error GoTo Err

PercentCompleteLabel.Caption = Format(pPercentComplete, "0.0")
TheTime.Caption = FormatTimestamp(pTickfileTimestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)

processDrawdown
processMaxProfit
processProfit pTickfileTimestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifySessionDrawdown(ByVal Value As Double)
Const ProcName As String = "IStrategyHostView_NotifySessionDrawdown"
On Error GoTo Err

mDrawdown = Value
'If Not mIsTickReplay Then processDrawdown
processDrawdown

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifySessionMaxProfit(ByVal Value As Double)
Const ProcName As String = "IStrategyHostView_NotifySessionMaxProfit"
On Error GoTo Err

mMaxProfit = Value
'If Not mIsTickReplay Then processMaxProfit
processMaxProfit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifySessionProfit(ByVal Value As Double, ByVal pTimestamp As Date)
Const ProcName As String = "IStrategyHostView_NotifySessionProfit"
On Error GoTo Err

mSessionProfit = Value
'If Not mIsTickReplay Then processProfit pTimestamp
processProfit pTimestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IStrategyHostView_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypes.TickTypeAsk
    If Not mModel.IsTickReplay Then
        AskLabel.Caption = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
        AskSizeLabel.Caption = ev.Tick.Size
    End If
Case TickTypes.TickTypeBid
    If Not mModel.IsTickReplay Then
        BidLabel.Caption = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
        BidSizeLabel.Caption = ev.Tick.Size
    End If
Case TickTypes.TickTypeTrade
    If Not mModel.IsTickReplay Then
        TradeLabel.Caption = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
        TradeSizeLabel.Caption = ev.Tick.Size
    End If
Case TickTypes.TickTypeVolume
    If Not mModel.IsTickReplay Then VolumeLabel.Caption = ev.Tick.Size
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifyTickerCreated()
Const ProcName As String = "IStrategyHostView_NotifyTickerCreated"
On Error GoTo Err

Set mContract = mModel.Contract
mSecType = mContract.Specifier.SecType
mTickSize = mContract.TickSize
Set mSession = mModel.Ticker.SessionFuture.Value

If mSession.CurrentSessionEndTime > mModel.Ticker.TimeStamp Then mTradingSessionInProgress = True

Dim i As Long
For i = 1 To PriceChart.Count
    PriceChart.SetStudyManager mModel.Ticker.StudyBase.StudyManager, i
Next

If mProfitStudyBase Is Nothing Then initialiseProfitChart
If mTradeStudyBase Is Nothing Then initialiseTradeChart

Me.Caption = "TradeBuild Strategy Trader - " & _
            StrategyCombo.Text & " - " & _
            mContract.Specifier.LocalSymbol

SSTab2.Tab = 3


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_NotifyTickfileCompleted(ByVal pIndex As Long)
Const ProcName As String = "IStrategyHostView_NotifyTickfileCompleted"
On Error GoTo Err

If pIndex < TickfileOrganiser1.TickFileSpecifiers.Count - 1 Then
    TickfileOrganiser1.ListIndex = pIndex
End If
mOverallProfit = mOverallProfit + mSessionProfit
mSessionProfit = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStrategyHostView_Parameters() As Parameters
Set IStrategyHostView_Parameters = mParams
End Property

Private Sub IStrategyHostView_ResetProfitChart()
Set mProfitStudyBase = Nothing
End Sub

Private Sub IStrategyHostView_ResetTradeChart()
Set mTradeStudyBase = Nothing
End Sub

Private Sub IStrategyHostView_ShowTradeLine(ByVal pStartTime As Date, ByVal pEndTime As Date, ByVal pEntryPrice As Double, ByVal pExitPrice As Double, ByVal pProfit As Double)
Const ProcName As String = "IStrategyHostView_ShowTradeLine"
On Error GoTo Err

If Not mModel.ShowChart Then Exit Sub

On Error Resume Next
Dim lStartPeriod As Period
Set lStartPeriod = mPricePeriods(pStartTime)
On Error GoTo 0
If lStartPeriod Is Nothing Then
    ' occurs if market data was not being received at the time of the entry execution
    Exit Sub
End If

Dim lBracketOrderLine As ChartSkil27.Line
Set lBracketOrderLine = mBracketOrderLineSeries.Add
lBracketOrderLine.Point1 = NewPoint(lStartPeriod.PeriodNumber, pEntryPrice)

On Error Resume Next
Dim lEndPeriod As Period
Set lEndPeriod = mPricePeriods(pEndTime)
On Error GoTo 0
If lEndPeriod Is Nothing Then
    ' this occurs when the execution that finished the order plex occurred
    ' at the start of a new bar but before the first price for the bar
    ' was reported. So add the bar now
    mPricePeriods.Add pEndTime
    Set lEndPeriod = mPricePeriods(pEndTime)
End If
lBracketOrderLine.Point2 = NewPoint(lEndPeriod.PeriodNumber, pExitPrice)

If pProfit > 0 Then
    lBracketOrderLine.Color = vbBlue
    lBracketOrderLine.ArrowEndColor = vbBlue
    lBracketOrderLine.ArrowEndFillColor = vbBlue
ElseIf pProfit = 0 Then
    lBracketOrderLine.Color = vbBlack
    lBracketOrderLine.ArrowEndColor = vbBlack
    lBracketOrderLine.ArrowEndFillColor = vbBlack
Else
    lBracketOrderLine.Color = vbRed
    lBracketOrderLine.ArrowEndColor = vbRed
    lBracketOrderLine.ArrowEndFillColor = vbRed
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStrategyHostView_Strategy() As IStrategy
Const ProcName As String = "IStrategyHostView_Strategy"
On Error GoTo Err

Set IStrategyHostView_Strategy = CreateObject(StrategyCombo.Text)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStrategyHostView_UpdateLastChartBars()
Const ProcName As String = "IStrategyHostView_UpdateLastChartBars"
On Error GoTo Err

PriceChart.UpdateLastBar
ProfitChart.UpdateLastBar
TradeChart.UpdateLastBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostView_WriteLogText(ByVal pMessage As String)
Const ProcName As String = "IStrategyHostView_WriteLogText"
On Error GoTo Err

WriteLogText pMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mSession Event Handlers
'================================================================================

Private Sub mSession_SessionEnded(ev As SessionEventData)
If Not mModel.IsTickReplay And mTradingSessionInProgress Then Unload Me
End Sub

Private Sub mSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mSession_SessionStarted"
On Error GoTo Err

mTradingSessionInProgress = True
If mModel.ShowChart Then mProfitStudyBase.NotifyValue mOverallProfit, mSession.SessionCurrentTime

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

If Not mModel.ShowChart Then Exit Sub

Dim ListItem As ListItem
Set ListItem = BracketOrderList.SelectedItem

Dim lPeriodNumber As Long
lPeriodNumber = mPricePeriods(BarStartTime(CDate(ListItem.SubItems(BOListColumns.ColumnStartTime - 1)), mPriceChartTimePeriod, mContract.SessionStartTime)).PeriodNumber
PriceChart.BaseChartController(1).LastVisiblePeriod = _
            lPeriodNumber + _
            Int((PriceChart.BaseChartController.LastVisiblePeriod - _
            PriceChart.BaseChartController.FirstVisiblePeriod) / 2) - 1
SSTab1.Tab = 0

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DummyProfitProfileCheck_Click()
mModel.LogDummyProfitProfile = CBool(DummyProfitProfileCheck.Value)
End Sub

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

Private Sub NoMoneyManagementCheck_Click()
mModel.UseMoneyManagement = Not CBool(NoMoneyManagementCheck.Value)
End Sub

Private Sub ProfitProfileCheck_Click()
mModel.LogProfitProfile = CBool(ProfitProfileCheck.Value)
End Sub

Private Sub ResultsPathButton_Click()
Const ProcName As String = "ResultsPathButton_Click"
On Error GoTo Err

ResultsPathText.Text = ChoosePath(ApplicationSettingsFolder & "Results")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ResultsPathText_Change()
mModel.ResultsPath = ResultsPathText.Text
End Sub

Private Sub SeparateSessionsCheck_Click()
mModel.SeparateSessions = CBool(SeparateSessionsCheck.Value)
End Sub

Private Sub ShowChartCheck_Click()
mModel.ShowChart = CBool(ShowChartCheck.Value)
End Sub

Private Sub StartButton_Click()
Const ProcName As String = "StartButton_Click"
On Error GoTo Err

startprocessing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StopButton_Click()
Const ProcName As String = "StopButton_Click"
On Error GoTo Err

mController.Finish
StartButton.Enabled = True
StopButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StopStrategyFactoryCombo_Change()
Const ProcName As String = "StopStrategyFactoryCombo_Change"
On Error GoTo Err

getDefaultParams

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StopStrategyFactoryCombo_Click()
Const ProcName As String = "StopStrategyFactoryCombo_Click"
On Error GoTo Err

getDefaultParams

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StrategyCombo_Change()
Const ProcName As String = "StrategyCombo_Change"
On Error GoTo Err

getDefaultParams

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StrategyCombo_Click()
Const ProcName As String = "StrategyCombo_Click"
On Error GoTo Err

getDefaultParams

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal pModel As IStrategyHostModel, _
                ByVal pController As IStrategyHostController)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mModel = pModel
Set mController = pController

LogMessage "Initialising charts"
initialisePriceChart
ProfitChart.BaseChartController.Style = mChartStyle
TradeChart.BaseChartController.Style = mChartStyle

LogMessage "Applying theme"
applyTheme New BlackTheme

LogMessage "Setting controls from model"
ResultsPathText.Text = mModel.ResultsPath
NoMoneyManagementCheck.Value = IIf(mModel.UseMoneyManagement, vbUnchecked, vbChecked)
If mModel.StrategyClassName <> "" Then StrategyCombo.Text = mModel.StrategyClassName
If mModel.StopStrategyFactoryClassName <> "" Then StopStrategyFactoryCombo.Text = mModel.StopStrategyFactoryClassName
ShowChartCheck.Value = IIf(mModel.ShowChart, vbChecked, vbUnchecked)

If mModel.UseLiveBroker Then
    SymbolText.Enabled = True
    SymbolText.Text = mModel.Symbol.LocalSymbol
    SymbolText.SetFocus
Else
    LogMessage "Enabling TickfileOrganiser"
    TickfileOrganiser1.Enabled = True
    
    If Not mModel.TickfileStoreInput Is Nothing Then
        TickfileOrganiser1.Initialise mModel.TickfileStoreInput, mModel.ContractStorePrimary
    End If
    
    ' the following line moves focus to TickfileOrganiser1. Trying to do this
    ' with TickfileOrganiser1.SetFocus causes VB to go into a loop!
    SymbolText.Enabled = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Start()
Const ProcName As String = "Start"
On Error GoTo Err

getDefaultParams
startprocessing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyTheme(ByVal pTheme As ITheme)
Const ProcName As String = "applyTheme"
On Error GoTo Err

Set mTheme = pTheme
Me.BackColor = mTheme.BaseColor
gApplyTheme mTheme, Me.Controls

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearPerformanceFields()
EventsPlayedLabel = ""
PercentCompleteLabel = ""
EventsPerSecondLabel = ""
MicrosecsPerEventLabel = ""
End Sub

Private Sub ClearPriceAndProfitFields()
BidLabel = ""
BidSizeLabel = ""
AskLabel = ""
AskSizeLabel = ""
TradeLabel = ""
TradeSizeLabel = ""
Profit.Caption = ""
Drawdown.Caption = ""
MaxProfit.Caption = ""
Position.Caption = ""
End Sub

Private Sub getDefaultParams()
Const ProcName As String = "getDefaultParams"
On Error GoTo Err

If StrategyCombo.Text = "" Then Exit Sub
If StopStrategyFactoryCombo.Text = "" Then Exit Sub

Dim lPMFactories As New Collection
lPMFactories.Add CreateObject(StopStrategyFactoryCombo.Text)
Set mParams = mController.GetDefaultParameters(CreateObject(StrategyCombo.Text), lPMFactories)

Set ParamGrid.DataSource = mParams
ParamGrid.Columns(0).Width = ParamGrid.Width / 2
ParamGrid.Columns(1).Width = ParamGrid.Width / 2

StartButton.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialisePriceChart()
Const ProcName As String = "initialisePriceChart"
On Error GoTo Err

If Not mModel.ShowChart Then Exit Sub

PriceChart.InitialiseRaw mChartStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseProfitChart()
Const ProcName As String = "initialiseProfitChart"
On Error GoTo Err

If Not mModel.ShowChart Then Exit Sub

Set mProfitStudyBase = CreateStudyBaseForDoubleInput( _
                                    mModel.StudyLibraryManager.CreateStudyManager( _
                                                    mContract.SessionStartTime, _
                                                    mContract.SessionEndTime, _
                                                    GetTimeZone(mContract.TimeZoneName)))

ProfitChart.Initialise CreateTimeframes(mProfitStudyBase), False
ProfitChart.DisableDrawing
ProfitChart.ShowChart GetTimePeriod(1, TimePeriodDay), _
                        CreateChartSpecifier(0), _
                        mChartStyle, _
                        pTitle:="Profit by Session"
ProfitChart.PriceRegion.YScaleQuantum = 0.01

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseTradeChart()
Const ProcName As String = "initialiseTradeChart"
On Error GoTo Err

If Not mModel.ShowChart Then Exit Sub

Set mTradeStudyBase = CreateStudyBaseForIntegerInput( _
                                    mModel.StudyLibraryManager.CreateStudyManager( _
                                                    mContract.SessionStartTime, _
                                                    mContract.SessionEndTime, _
                                                    GetTimeZone(mContract.TimeZoneName)))

TradeChart.Initialise CreateTimeframes(mTradeStudyBase), False
TradeChart.DisableDrawing
TradeChart.ShowChart GetTimePeriod(0, TimePeriodNone), _
                    CreateChartSpecifier(0), _
                    mChartStyle, _
                    pTitle:="Profit by Trade"
TradeChart.PriceRegion.YScaleQuantum = 0.01

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function minimumHeight() As Long
minimumHeight = SSTab2.Top + SSTab2.Height
End Function

Private Sub processDrawdown()
Const ProcName As String = "processDrawdown"
On Error GoTo Err

Drawdown.Caption = Format(mDrawdown, "0.00")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processMaxProfit()
Const ProcName As String = "processMaxProfit"
On Error GoTo Err

MaxProfit.Caption = Format(mMaxProfit, "0.00")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processProfit(ByVal pTimestamp As Date)
Const ProcName As String = "processProfit"
On Error GoTo Err

Profit.Caption = Format(mSessionProfit, "0.00")

If mModel.ShowChart And mPosition <> 0 Then
    mProfitStudyBase.NotifyValue mOverallProfit + mSessionProfit, pTimestamp
    mTradeStudyBase.NotifyValue mOverallProfit + mSessionProfit, pTimestamp
End If

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

Private Sub startprocessing()
Const ProcName As String = "startprocessing"
On Error GoTo Err

StartButton.Enabled = False
StopButton.Enabled = True

PriceChart.Clear
ProfitChart.BaseChartController.ClearChart
TradeChart.BaseChartController.ClearChart
BracketOrderList.ListItems.Clear

ClearPriceAndProfitFields
clearPerformanceFields

mOverallProfit = 0#
mSessionProfit = 0#

Set mBracketOrderLineSeries = Nothing
Set mPricePeriods = Nothing

If TickfileOrganiser1.TickfileCount <> 0 Then
    mController.StartTickfileReplay TickfileOrganiser1.TickFileSpecifiers
Else
    mController.StartLiveProcessing mModel.Symbol
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteLogText(ByVal pMessage As String)
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



