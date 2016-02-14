VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#309.0#0"; "TradingUI27.ocx"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   32
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
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Trade chart"
      TabPicture(2)   =   "fStrategyHost.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TradeChart"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Bracket order details"
      TabPicture(3)   =   "fStrategyHost.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "BracketOrderList"
      Tab(3).ControlCount=   1
      Begin TradingUI27.MarketChart ProfitChart 
         Height          =   5295
         Left            =   -75000
         TabIndex        =   53
         Top             =   330
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9340
      End
      Begin TradingUI27.MultiChart PriceChart 
         Height          =   5295
         Left            =   0
         TabIndex        =   52
         Top             =   330
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9340
      End
      Begin MSComctlLib.ListView BracketOrderList 
         Height          =   5295
         Left            =   -75000
         TabIndex        =   33
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
      Begin TradingUI27.MarketChart TradeChart 
         Height          =   5295
         Left            =   -75000
         TabIndex        =   54
         Top             =   330
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9340
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3495
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Controls"
      TabPicture(0)   =   "fStrategyHost.frx":0070
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Parameters"
      TabPicture(1)   =   "fStrategyHost.frx":008C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Log"
      TabPicture(2)   =   "fStrategyHost.frx":00A8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LogPicture"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Results"
      TabPicture(3)   =   "fStrategyHost.frx":00C4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MoreButton"
      Tab(3).Control(1)=   "Label11"
      Tab(3).Control(2)=   "VolumeLabel"
      Tab(3).Control(3)=   "BidSizeLabel"
      Tab(3).Control(4)=   "BidLabel"
      Tab(3).Control(5)=   "TradeSizeLabel"
      Tab(3).Control(6)=   "TradeLabel"
      Tab(3).Control(7)=   "AskSizeLabel"
      Tab(3).Control(8)=   "AskLabel"
      Tab(3).Control(9)=   "Label7"
      Tab(3).Control(10)=   "MicrosecsPerEventLabel"
      Tab(3).Control(11)=   "EventsPerSecondLabel"
      Tab(3).Control(12)=   "Label3"
      Tab(3).Control(13)=   "PercentCompleteLabel"
      Tab(3).Control(14)=   "Label2"
      Tab(3).Control(15)=   "EventsPlayedLabel"
      Tab(3).Control(16)=   "Label1"
      Tab(3).Control(17)=   "Label8"
      Tab(3).Control(18)=   "Label10"
      Tab(3).Control(19)=   "Label9"
      Tab(3).Control(20)=   "Label4"
      Tab(3).Control(21)=   "Profit"
      Tab(3).Control(22)=   "Drawdown"
      Tab(3).Control(23)=   "Label12"
      Tab(3).Control(24)=   "Label5"
      Tab(3).Control(25)=   "MaxProfit"
      Tab(3).Control(26)=   "Position"
      Tab(3).Control(27)=   "Label14"
      Tab(3).Control(28)=   "TheTime"
      Tab(3).ControlCount=   29
      Begin VB.PictureBox LogPicture 
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   -74880
         ScaleHeight     =   3075
         ScaleWidth      =   10680
         TabIndex        =   38
         Top             =   360
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
            Height          =   3000
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   39
            Top             =   0
            Width           =   10695
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   10695
         TabIndex        =   35
         Top             =   480
         Width           =   10695
         Begin TradingUI27.TickfileOrganiser TickfileOrganiser1 
            Height          =   2535
            Left            =   0
            TabIndex        =   1
            Top             =   360
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4471
         End
         Begin VB.CheckBox ShowChartCheck 
            Caption         =   "Show chart"
            Height          =   195
            Left            =   6000
            TabIndex        =   5
            Top             =   840
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.ComboBox StopStrategyFactoryCombo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "fStrategyHost.frx":00E0
            Left            =   6000
            List            =   "fStrategyHost.frx":00E7
            Sorted          =   -1  'True
            TabIndex        =   4
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
            ItemData        =   "fStrategyHost.frx":010E
            Left            =   6000
            List            =   "fStrategyHost.frx":0115
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   0
            Width           =   3495
         End
         Begin VB.CheckBox DummyProfitProfileCheck 
            Caption         =   "Dummy profit profile"
            Height          =   195
            Left            =   6000
            TabIndex        =   7
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox ProfitProfileCheck 
            Caption         =   "Profit profile"
            Height          =   195
            Left            =   6000
            TabIndex        =   6
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox NoMoneyManagement 
            Caption         =   "No money management"
            Height          =   195
            Left            =   6000
            TabIndex        =   8
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox SeparateSessionsCheck 
            Caption         =   "Separate session per tick file"
            Height          =   195
            Left            =   8040
            TabIndex        =   9
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.CommandButton StopButton 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9600
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Default         =   -1  'True
            Enabled         =   0   'False
            Height          =   375
            Left            =   9600
            TabIndex        =   13
            Top             =   0
            Width           =   1095
         End
         Begin VB.CheckBox LiveTradesCheck 
            Caption         =   "Live trades"
            Height          =   195
            Left            =   8040
            TabIndex        =   10
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox ResultsPathText 
            Height          =   255
            Left            =   6960
            TabIndex        =   11
            Top             =   1800
            Width           =   1995
         End
         Begin VB.CommandButton ResultsPathButton 
            Caption         =   "..."
            Height          =   255
            Left            =   9000
            TabIndex        =   12
            ToolTipText     =   "Select results path"
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label 
            Caption         =   "Symbol"
            Height          =   375
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Results path"
            Height          =   255
            Left            =   6000
            TabIndex        =   36
            Top             =   1800
            Width           =   975
         End
      End
      Begin VB.CommandButton MoreButton 
         Caption         =   "Less <<<"
         Height          =   375
         Left            =   -68400
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   3090
         Left            =   -74880
         ScaleHeight     =   3090
         ScaleWidth      =   10695
         TabIndex        =   20
         Top             =   360
         Width           =   10695
         Begin MSDataGridLib.DataGrid ParamGrid 
            Height          =   2985
            Left            =   0
            TabIndex        =   15
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
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1365
         Index           =   0
         Left            =   120
         ScaleHeight     =   1365
         ScaleWidth      =   7455
         TabIndex        =   19
         Top             =   360
         Width           =   7455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Volume"
         Height          =   195
         Left            =   -74880
         TabIndex        =   56
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label VolumeLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   55
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label BidSizeLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.Label BidLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73800
         TabIndex        =   34
         Top             =   960
         Width           =   735
      End
      Begin VB.Label TradeSizeLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   40
         Top             =   720
         Width           =   735
      End
      Begin VB.Label TradeLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73800
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label AskSizeLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   51
         Top             =   480
         Width           =   735
      End
      Begin VB.Label AskLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73800
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Microsecs per event"
         Height          =   195
         Left            =   -71280
         TabIndex        =   49
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label MicrosecsPerEventLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   48
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label EventsPerSecondLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   47
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Events per second"
         Height          =   195
         Left            =   -71280
         TabIndex        =   46
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label PercentCompleteLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   45
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Percent complete"
         Height          =   195
         Left            =   -71280
         TabIndex        =   44
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label EventsPlayedLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   43
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Events played"
         Height          =   195
         Left            =   -71280
         TabIndex        =   42
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Bid"
         Height          =   195
         Left            =   -74880
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Last"
         Height          =   195
         Left            =   -74880
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Ask"
         Height          =   195
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Profit/Loss"
         Height          =   195
         Left            =   -71280
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Profit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Drawdown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   27
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Drawdown"
         Height          =   195
         Left            =   -71280
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Max profit"
         Height          =   195
         Left            =   -71280
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
      Begin VB.Label MaxProfit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69600
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Position"
         Height          =   195
         Left            =   -71280
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label TheTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73815
         TabIndex        =   21
         Top             =   1560
         Width           =   1815
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
Implements IStateChangeListener
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

Private Type StudyConfigToShow
    Timeframe           As Timeframe
    StudyConfig         As StudyConfiguration
End Type

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

Private mTimeframes                                     As EnumerableCollection

Private mPriceChartTimePeriod                           As TimePeriod
Private mPriceChartIndex                                As Long

Private mProfitStudyBase                                As StudyBaseForDoubleInput

Private mTradeStudyBase                                 As StudyBaseForIntegerInput
Private mTradeBarNumber                                 As Long

Private mPosition                                       As Long
Private mOverallProfit                                  As Double
Private mSessionProfit                                  As Double
Private mMaxProfit                                      As Double
Private mDrawdown                                       As Double

Private mDetailsHidden                                  As Boolean

Private mBracketOrderLineSeries                         As LineSeries

Private mPricePeriods                                   As Periods

Private mReplayStartTime                                As Date

Private mTotalElapsedSecs                               As Double
Private mElapsedSecsCurrTickfile                        As Double
Private mTotalEvents                                    As Long
Private mEventsCurrTickfile                             As Long

Private WithEvents mFutureWaiter                        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mShowChart                                      As Boolean

Private mNumberOfTimeframesLoading                      As Long

Private mIsTickReplay                                   As Boolean

Private mStudiesToShow()                                As StudyConfigToShow
Private mStudiesToShowIndex                             As Long

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

If mShowChart Then
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
    If Not mIsTickReplay Then
        AskLabel.Caption = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
        AskSizeLabel.Caption = ev.Tick.Size
    End If
Case TickTypes.TickTypeBid
    If Not mIsTickReplay Then
        BidLabel.Caption = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
        BidSizeLabel.Caption = ev.Tick.Size
    End If
Case TickTypes.TickTypeTrade
    If Not mIsTickReplay Then
        TradeLabel.Caption = FormatPrice(ev.Tick.Price, mSecType, mTickSize)
        TradeSizeLabel.Caption = ev.Tick.Size
    End If
Case TickTypes.TickTypeVolume
    If Not mIsTickReplay Then VolumeLabel.Caption = ev.Tick.Size
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IStateChangeListener Interface Members
'================================================================================

Private Sub IStateChangeListener_Change(ev As TWUtilities40.StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If ev.State <> TimeframeStates.TimeframeStateLoaded Then Exit Sub

Dim lTimeframe As Timeframe: Set lTimeframe = ev.Source
mNumberOfTimeframesLoading = mNumberOfTimeframesLoading - 1

Dim lChartManager As ChartManager
Set lChartManager = PriceChart.ChartManager(mTimeframes(lTimeframe.TimePeriod.ToString))

lChartManager.SetBaseStudyConfiguration CreateBarsStudyConfig(lTimeframe, mContract.Specifier.SecType, mTicker.StudyBase.StudyManager.StudyLibraryManager), 0

addStudiesForChart lTimeframe, lChartManager

If mNumberOfTimeframesLoading = 0 Then
    mTicker.AddGenericTickListener Me
    
    If mTicker.IsTickReplay Then
        mStrategyRunner.StartReplay
        mReplayStartTime = GetTimestamp
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IStrategyHost Interface Members
'================================================================================

Private Sub IStrategyHost_AddTimeframe( _
                ByVal pTimeframe As Timeframe)
Const ProcName As String = "IStrategyHost_AddTimeframe"
On Error GoTo Err

monitorTimeframe pTimeframe

If Not mShowChart Then Exit Sub
If mTimeframes.Contains(pTimeframe.TimePeriod.ToString) Then Exit Sub

Dim lTitle As String
Dim lUpdatePerTick As Boolean

If mIsTickReplay Then
    lTitle = ""
    lUpdatePerTick = False
Else
    lTitle = mContract.Specifier.LocalSymbol
    lUpdatePerTick = True
End If

Dim lIndex As Long
lIndex = PriceChart.AddRaw(pTimeframe, _
                        mTicker.StudyBase.StudyManager, _
                        mContract.Specifier.LocalSymbol, _
                        mContract.Specifier.SecType, _
                        mContract.Specifier.Exchange, _
                        mContract.TickSize, _
                        mContract.SessionStartTime, _
                        mContract.SessionEndTime, _
                        lTitle, _
                        lUpdatePerTick)

mTimeframes.Add lIndex, pTimeframe.TimePeriod.ToString

If mPriceChartIndex = 0 Then
    mPriceChartIndex = lIndex
    Set mPriceChartTimePeriod = pTimeframe.TimePeriod
End If

If mPricePeriods Is Nothing Then Set mPricePeriods = PriceChart.BaseChartController.Periods
If mBracketOrderLineSeries Is Nothing Then Set mBracketOrderLineSeries = PriceChart.BaseChartController.Regions.Item(ChartRegionNamePrice).AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerHighestUser)
mBracketOrderLineSeries.Thickness = 2
mBracketOrderLineSeries.ArrowEndStyle = ArrowClosed
mBracketOrderLineSeries.ArrowEndWidth = 8
mBracketOrderLineSeries.ArrowEndLength = 12

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_ChartStudyValue( _
                ByVal pStudy As IStudy, _
                ByVal pValueName As String, _
                ByVal pTimeframe As Timeframe)
Const ProcName As String = "IStrategyHost_ChartStudyValue"
On Error GoTo Err

Dim lStudyConfig As StudyConfiguration
Dim lSvc As StudyValueConfiguration

If Not findStudyConfig(pStudy, pTimeframe, lStudyConfig) Then
    Dim lChartManager As ChartManager
    Set lChartManager = PriceChart.ChartManager(PriceChart.GetIndexFromTimeframe(pTimeframe))
    Set lStudyConfig = lChartManager.GetDefaultStudyConfiguration(pStudy.Name, pStudy.LibraryName)
    lStudyConfig.Study = pStudy
    lStudyConfig.UnderlyingStudy = pStudy.UnderlyingStudy

    Assert Not lStudyConfig Is Nothing, "Can't get default study configuration"

    For Each lSvc In lStudyConfig.StudyValueConfigurations
        lSvc.IncludeInChart = False
    Next
    
    mStudiesToShowIndex = mStudiesToShowIndex + 1
    If mStudiesToShowIndex > UBound(mStudiesToShow) Then ReDim Preserve mStudiesToShow(2 * (UBound(mStudiesToShow) + 1) - 1) As StudyConfigToShow
    Set mStudiesToShow(mStudiesToShowIndex).StudyConfig = lStudyConfig
    Set mStudiesToShow(mStudiesToShowIndex).Timeframe = pTimeframe
End If

For Each lSvc In lStudyConfig.StudyValueConfigurations
    If lSvc.ValueName = pValueName Then lSvc.IncludeInChart = True
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

mTicker.RemoveGenericTickListener Me
mTicker.Finish

mTotalElapsedSecs = mTotalElapsedSecs + mElapsedSecsCurrTickfile
mElapsedSecsCurrTickfile = 0

mTotalEvents = mTotalEvents + mEventsCurrTickfile
mEventsCurrTickfile = 0

If mShowChart Then
    If mIsTickReplay Then
        ' ensure final bars in charts are displayed
        PriceChart.UpdateLastBar
        ProfitChart.UpdateLastBar
        TradeChart.UpdateLastBar
    End If
    PriceChart.BaseChartController.EnableDrawing
    ProfitChart.EnableDrawing
    TradeChart.EnableDrawing
End If

If mTickfileIndex = TickfileOrganiser1.TickFileSpecifiers.Count - 1 Then
    Set mTimeframes = New EnumerableCollection
    Set mProfitStudyBase = Nothing
    Set mTradeStudyBase = Nothing
    mTradeBarNumber = 0
    StartButton.Enabled = True
    StopButton.Enabled = False
Else
    mOverallProfit = mOverallProfit + mSessionProfit
    If SeparateSessionsCheck = vbChecked Then
        clearPriceAndProfitFields
    Else
    End If
    mSessionProfit = 0
    
    startNextTickfile
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_NotifyReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Single)
Const ProcName As String = "IStrategyHost_NotifyReplayProgress"
On Error GoTo Err

PercentCompleteLabel.Caption = Format(pPercentComplete, "0.0")
TheTime.Caption = FormatTimestamp(pTickfileTimestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)

processDrawdown
processMaxProfit
processProfit pTickfileTimestamp

mEventsCurrTickfile = pEventsPlayed
Dim lTotalEvents As Long
lTotalEvents = mTotalEvents + mEventsCurrTickfile

mElapsedSecsCurrTickfile = (GetTimestamp - mReplayStartTime) * 86400
Dim lTotalElapsedSecs As Double
lTotalElapsedSecs = mTotalElapsedSecs + mElapsedSecsCurrTickfile

EventsPlayedLabel.Caption = lTotalEvents
EventsPerSecondLabel.Caption = Int(lTotalEvents / lTotalElapsedSecs)
MicrosecsPerEventLabel.Caption = Int(lTotalElapsedSecs * 1000000 / lTotalEvents)

If mShowChart Then
    PriceChart.BaseChartController.EnableDrawing
    PriceChart.BaseChartController.DisableDrawing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_NotifyReplayStarted()
Const ProcName As String = "IStrategyHost_NotifyReplayStarted"
On Error GoTo Err

If mShowChart Then
    Dim i As Long
    For i = 1 To PriceChart.Count
        PriceChart.BaseChartController(i).DisableDrawing
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
mIsTickReplay = mTicker.IsTickReplay
Set mContract = mTicker.ContractFuture.Value
mSecType = mContract.Specifier.SecType
mTickSize = mContract.TickSize
Set mSession = mTicker.SessionFuture.Value

ReDim mStudiesToShow(3) As StudyConfigToShow
mStudiesToShowIndex = -1

If PriceChart.Count = 0 Then
    Set mTimeframes = New EnumerableCollection
    initialisePriceChart
Else
    Dim i As Long
    For i = 1 To PriceChart.Count
        PriceChart.SetStudyManager mTicker.StudyBase.StudyManager, i
    Next
End If
If mProfitStudyBase Is Nothing Then initialiseProfitChart
If mTradeStudyBase Is Nothing Then initialiseTradeChart

mStrategyRunner.StartStrategy CreateObject(StrategyCombo.Text), mParams

Me.Caption = "TradeBuild Strategy Trader - " & _
            StrategyCombo.Text & " - " & _
            mContract.Specifier.LocalSymbol

SSTab2.Tab = 3

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    mSessionProfit = pLogrec.Data
    processProfit mTicker.TimeStamp
Case "position.drawdown"
    mDrawdown = pLogrec.Data
    If Not mIsTickReplay Then processDrawdown
Case "position.maxprofit"
    mMaxProfit = pLogrec.Data
    If Not mIsTickReplay Then processMaxProfit
Case "position.bracketorderprofilestruct"
    Dim lListItem As ListItem
    Static sBracketOrderNumber As Long

    Dim lBracketOrderProfile As BracketOrderProfile
    lBracketOrderProfile = pLogrec.Data
    
    showBracketOrderLine lBracketOrderProfile
    
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
    Dim lPrevPosition As Long: lPrevPosition = mPosition
    
    mPosition = pLogrec.Data
    Position.Caption = mPosition
    
    If (mPosition <> 0 And lPrevPosition = 0) Or _
        (mPosition > 0 And lPrevPosition < 0) Or _
        (mPosition < 0 And lPrevPosition > 0) _
    Then
        If mIsTickReplay Then
            TradeChart.EnableDrawing
            TradeChart.DisableDrawing
        End If
        mTradeBarNumber = mTradeBarNumber + 1
        If mShowChart Then
            LogMessage "New trade bar: " & mTradeBarNumber & " at " & mTicker.TimeStamp
            mTradeStudyBase.NotifyBarNumber mTradeBarNumber, mTicker.TimeStamp
            mTradeStudyBase.NotifyValue mOverallProfit + mSessionProfit, mTicker.TimeStamp
        End If
    End If
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
    mStrategyRunner.StartStrategy CreateObject(StrategyCombo.Text), mParams
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

If mShowChart Then mProfitStudyBase.NotifyValue mOverallProfit, mTicker.TimeStamp

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

If Not mShowChart Then Exit Sub

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

Private Sub ResultsPathButton_Click()
Const ProcName As String = "ResultsPathButton_Click"
On Error GoTo Err

ResultsPathText.Text = ChoosePath(ApplicationSettingsFolder & "Results")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StartButton_Click()
Const ProcName As String = "StartButton_Click"
On Error GoTo Err

StartButton.Enabled = False
StopButton.Enabled = True

mShowChart = (ShowChartCheck = vbChecked)

PriceChart.Clear
ProfitChart.BaseChartController.ClearChart
TradeChart.BaseChartController.ClearChart
BracketOrderList.ListItems.Clear

clearPriceAndProfitFields
clearPerformanceFields

mOverallProfit = 0#
mSessionProfit = 0#

Set mBracketOrderLineSeries = Nothing
Set mPricePeriods = Nothing
mPriceChartIndex = 0

mTickfileIndex = -1

If TickfileOrganiser1.TickfileCount <> 0 Then
    startNextTickfile
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

Private Sub addStudiesForChart( _
                ByVal pTimeframe As Timeframe, _
                ByVal pChartManager As ChartManager)
Const ProcName As String = "addStudiesForChart"
On Error GoTo Err

Dim i As Long
For i = 0 To mStudiesToShowIndex
    If mStudiesToShow(i).Timeframe Is pTimeframe Then
        pChartManager.ApplyStudyConfiguration mStudiesToShow(i).StudyConfig, 0
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearPerformanceFields()
EventsPlayedLabel = ""
PercentCompleteLabel = ""
EventsPerSecondLabel = ""
MicrosecsPerEventLabel = ""
mTotalElapsedSecs = 0#
mElapsedSecsCurrTickfile = 0#
mTotalEvents = 0
mEventsCurrTickfile = 0
End Sub

Private Sub clearPriceAndProfitFields()
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

Private Function findStudyConfig( _
                ByVal pStudy As IStudy, _
                ByVal pTimeframe As Timeframe, _
                ByRef pStudyConfig As StudyConfiguration) As Boolean
Dim i As Long
For i = 0 To mStudiesToShowIndex
    If mStudiesToShow(i).StudyConfig.Study Is pStudy And _
        mStudiesToShow(i).Timeframe Is pTimeframe _
    Then
        Set pStudyConfig = mStudiesToShow(i).StudyConfig
        findStudyConfig = True
        Exit Function
    End If
Next
findStudyConfig = False
End Function

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

Private Sub getDefaultParams()
Const ProcName As String = "getDefaultParams"
On Error GoTo Err

If StrategyCombo.Text = "" Then Exit Sub
If StopStrategyFactoryCombo.Text = "" Then Exit Sub

Set mStrategyRunner = CreateStrategyRunner(Me)
Dim lPMFactories As New Collection
lPMFactories.Add CreateObject(StopStrategyFactoryCombo.Text)
Set mParams = mStrategyRunner.GetDefaultParameters(CreateObject(StrategyCombo.Text), lPMFactories)

Set ParamGrid.DataSource = mParams
ParamGrid.Columns(0).Width = ParamGrid.Width / 2
ParamGrid.Columns(1).Width = ParamGrid.Width / 2

StartButton.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialisePriceChart(Optional ByVal pTimestamp As Date)
Const ProcName As String = "initialisePriceChart"
On Error GoTo Err

If Not mShowChart Then Exit Sub

PriceChart.InitialiseRaw ChartStylesManager.DefaultStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseProfitChart()
Const ProcName As String = "initialiseProfitChart"
On Error GoTo Err

If Not mShowChart Then Exit Sub

Set mProfitStudyBase = CreateStudyBaseForDoubleInput( _
                                    gTB.StudyLibraryManager.CreateStudyManager( _
                                                    mContract.SessionStartTime, _
                                                    mContract.SessionEndTime, _
                                                    GetTimeZone(mContract.TimeZoneName)))

If mIsTickReplay Then ProfitChart.Initialise CreateTimeframes(mProfitStudyBase), False
ProfitChart.DisableDrawing
ProfitChart.ShowChart GetTimePeriod(1, TimePeriodDay), _
                        CreateChartSpecifier(0), _
                        ChartStylesManager.DefaultStyle, _
                        pTitle:="Profit by Session"
ProfitChart.PriceRegion.YScaleQuantum = 0.01

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseTradeChart()
Const ProcName As String = "initialiseTradeChart"
On Error GoTo Err

If Not mShowChart Then Exit Sub

Set mTradeStudyBase = CreateStudyBaseForIntegerInput( _
                                    gTB.StudyLibraryManager.CreateStudyManager( _
                                                    mContract.SessionStartTime, _
                                                    mContract.SessionEndTime, _
                                                    GetTimeZone(mContract.TimeZoneName)))

If mIsTickReplay Then TradeChart.Initialise CreateTimeframes(mTradeStudyBase), False
TradeChart.DisableDrawing
TradeChart.ShowChart GetTimePeriod(0, TimePeriodNone), _
                    CreateChartSpecifier(0), _
                    ChartStylesManager.DefaultStyle, _
                    pTitle:="Profit by Trade"
TradeChart.PriceRegion.YScaleQuantum = 0.01

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function minimumHeight() As Long
minimumHeight = SSTab2.Top + SSTab2.Height
End Function

Private Sub monitorTimeframe(ByVal pTimeframe As Timeframe)
Const ProcName As String = "monitorTimeframe"
On Error GoTo Err

If pTimeframe.State = TimeframeStateLoading Then
    mNumberOfTimeframesLoading = mNumberOfTimeframesLoading + 1
    pTimeframe.AddStateChangeListener Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

If mShowChart And mPosition <> 0 Then
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

Private Sub showBracketOrderLine(ByRef pBracketOrderProfile As BracketOrderProfile)
Const ProcName As String = "showBracketOrderLine"
On Error GoTo Err

If Not mShowChart Then Exit Sub

Dim lBracketOrderLine As ChartSkil27.Line
Set lBracketOrderLine = mBracketOrderLineSeries.Add
lBracketOrderLine.Point1 = NewPoint(mPricePeriods(BarStartTime(pBracketOrderProfile.StartTime, mPriceChartTimePeriod, mContract.SessionStartTime)).PeriodNumber, pBracketOrderProfile.EntryPrice)

Dim lLineEndBarStartTime As Date
lLineEndBarStartTime = BarStartTime(pBracketOrderProfile.EndTime, mPriceChartTimePeriod, mContract.SessionStartTime)

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
lBracketOrderLine.Point2 = NewPoint(mPricePeriods(lLineEndBarStartTime).PeriodNumber, pBracketOrderProfile.ExitPrice)

If pBracketOrderProfile.Profit > 0 Then
    lBracketOrderLine.Color = vbBlue
    lBracketOrderLine.ArrowEndColor = vbBlue
    lBracketOrderLine.ArrowEndFillColor = vbBlue
ElseIf pBracketOrderProfile.Profit = 0 Then
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

Private Sub startNextTickfile()
Const ProcName As String = "startNextTickfile"
On Error GoTo Err

mTickfileIndex = mTickfileIndex + 1
TickfileOrganiser1.ListIndex = mTickfileIndex
mStrategyRunner.PrepareTickFile TickfileOrganiser1.TickFileSpecifiers(mTickfileIndex + 1)

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



