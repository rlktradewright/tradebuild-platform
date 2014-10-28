VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#254.3#0"; "y.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#16.1#0"; "TWControls40.ocx"
Begin VB.Form fTradeSkilDemo 
   Caption         =   "TradeSkil Demo Edition Version 2.7"
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
      Height          =   240
      Left            =   16320
      MouseIcon       =   "fTradeSkilDemo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":0152
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      ToolTipText     =   "Show features"
      Top             =   9345
      Width           =   240
   End
   Begin VB.PictureBox HideFeaturesPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   16290
      MouseIcon       =   "fTradeSkilDemo.frx":06DC
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":082E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      ToolTipText     =   "Hide features"
      Top             =   5070
      Width           =   240
   End
   Begin VB.PictureBox ShowControlsPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      MouseIcon       =   "fTradeSkilDemo.frx":0DB8
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":0F0A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   56
      ToolTipText     =   "Show controls"
      Top             =   440
      Width           =   240
   End
   Begin VB.PictureBox HideControlsPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3900
      MouseIcon       =   "fTradeSkilDemo.frx":1494
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":15E6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   55
      ToolTipText     =   "Hide controls"
      Top             =   440
      Width           =   240
   End
   Begin TabDlg.SSTab FeaturesSSTAB 
      Height          =   4455
      Left            =   4320
      TabIndex        =   8
      Top             =   5040
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabsPerRow      =   6
      TabHeight       =   520
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
      TabPicture(0)   =   "fTradeSkilDemo.frx":1B70
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TickfileOrdersSummary"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SimulatedOrdersSummary"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LiveOrdersSummary"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OrdersSummaryTabStrip"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ModifyOrderPlexButton"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CancelOrderPlexButton"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ClosePositionsButton"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "OrderTicket1Button"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&2. Executions"
      TabPicture(1)   =   "fTradeSkilDemo.frx":1B8C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LiveExecutionsSummary"
      Tab(1).Control(1)=   "ExecutionsSummaryTabStrip"
      Tab(1).Control(2)=   "SimulatedExecutionsSummary"
      Tab(1).Control(3)=   "TickfileExecutionsSummary"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&3. Log"
      TabPicture(2)   =   "fTradeSkilDemo.frx":1BA8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LogText"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton OrderTicket1Button 
         Caption         =   "Order  Ticket"
         Height          =   495
         Left            =   11160
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox LogText 
         Height          =   3975
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
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
         TabIndex        =   12
         Top             =   3510
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderPlexButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   11160
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderPlexButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   11160
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin MSComctlLib.TabStrip OrdersSummaryTabStrip 
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         MultiRow        =   -1  'True
         Style           =   2
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
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
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tickfile"
               Object.ToolTipText     =   "Show tickfile orders"
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
      Begin TradingUI27.OrdersSummary LiveOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.OrdersSummary SimulatedOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.ExecutionsSummary LiveExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   52
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6376
      End
      Begin MSComctlLib.TabStrip ExecutionsSummaryTabStrip 
         Height          =   375
         Left            =   -74880
         TabIndex        =   53
         Top             =   3720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         MultiRow        =   -1  'True
         Style           =   2
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Live"
               Object.ToolTipText     =   "Show live executions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Simulated"
               Object.ToolTipText     =   "Show simulated executions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tickfile"
               Object.ToolTipText     =   "Show executions against tickfiles"
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
      Begin TradingUI27.ExecutionsSummary SimulatedExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   54
         Top             =   120
         Width           =   11995
         _ExtentX        =   21167
         _ExtentY        =   6376
      End
      Begin TradingUI27.OrdersSummary TickfileOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.ExecutionsSummary TickfileExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   68
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6376
      End
   End
   Begin MSComctlLib.TabStrip ControlsTabStrip 
      Height          =   580
      Left            =   120
      TabIndex        =   51
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
   Begin TabDlg.SSTab ControlsSSTab 
      Height          =   9015
      Left            =   120
      TabIndex        =   34
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   15901
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "fTradeSkilDemo.frx":1BC4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LiveContractSearch"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "fTradeSkilDemo.frx":1BE0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LiveChartStylesCombo"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "ChartButton"
      Tab(1).Control(3)=   "SessionOnlyCheck"
      Tab(1).Control(4)=   "NumHistoryBarsText"
      Tab(1).Control(5)=   "LiveChartTimeframeSelector"
      Tab(1).Control(6)=   "Label1"
      Tab(1).Control(7)=   "Label22"
      Tab(1).Control(8)=   "Label18"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "fTradeSkilDemo.frx":1BFC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(4)=   "Label8"
      Tab(2).Control(5)=   "HistChartStylesCombo"
      Tab(2).Control(6)=   "FromDatePicker"
      Tab(2).Control(7)=   "ToDatePicker"
      Tab(2).Control(8)=   "HistTimeframeSelector"
      Tab(2).Control(9)=   "NumHistBarsText"
      Tab(2).Control(10)=   "HistSessionOnlyCheck"
      Tab(2).Control(11)=   "HistContractSearch"
      Tab(2).Control(12)=   "Frame2"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "fTradeSkilDemo.frx":1C18
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TickfileOrganiser1"
      Tab(3).Control(1)=   "PlayTickFileButton"
      Tab(3).Control(2)=   "PauseReplayButton"
      Tab(3).Control(3)=   "StopReplayButton"
      Tab(3).Control(4)=   "ReplaySpeedCombo"
      Tab(3).Control(5)=   "ReplayProgressBar"
      Tab(3).Control(6)=   "ReplayProgressLabel"
      Tab(3).Control(7)=   "ReplayContractLabel"
      Tab(3).Control(8)=   "Label20"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "fTradeSkilDemo.frx":1C34
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label6"
      Tab(4).Control(1)=   "CurrentConfigNameText"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "ConfigEditorButton"
      Tab(4).ControlCount=   3
      Begin TradingUI27.TickfileOrganiser TickfileOrganiser1 
         Height          =   2520
         Left            =   -74880
         TabIndex        =   66
         Top             =   360
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   4445
      End
      Begin VB.Frame Frame2 
         Caption         =   "Change chart styles"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   62
         Top             =   7560
         Width           =   3855
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   60
            ScaleHeight     =   975
            ScaleWidth      =   3735
            TabIndex        =   63
            Top             =   240
            Width           =   3735
            Begin VB.CommandButton ChangeHistChartStylesButton 
               Caption         =   "Change ALL historical chart styles"
               Height          =   495
               Left            =   480
               TabIndex        =   64
               Top             =   480
               Width           =   2775
            End
            Begin VB.Label Label9 
               Caption         =   "Click this button to change the style of all existing historical charts to the style selected above."
               Height          =   495
               Left            =   120
               TabIndex        =   65
               Top             =   0
               Width           =   3495
            End
         End
      End
      Begin TWControls40.TWImageCombo LiveChartStylesCombo 
         Height          =   330
         Left            =   -73080
         TabIndex        =   16
         Top             =   1800
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
         MouseIcon       =   "fTradeSkilDemo.frx":1C50
         Text            =   ""
      End
      Begin VB.Frame Frame1 
         Caption         =   "Change chart styles"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   18
         Top             =   3600
         Width           =   3855
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   60
            ScaleHeight     =   975
            ScaleWidth      =   3735
            TabIndex        =   58
            Top             =   240
            Width           =   3735
            Begin VB.CommandButton ChangeLiveChartStylesButton 
               Caption         =   "Change ALL live chart styles"
               Height          =   495
               Left            =   480
               TabIndex        =   60
               Top             =   480
               Width           =   2775
            End
            Begin VB.Label Label7 
               Caption         =   "Click this button to change the style of all existing live charts to the style selected above."
               Height          =   495
               Left            =   120
               TabIndex        =   59
               Top             =   0
               Width           =   3495
            End
         End
      End
      Begin TradingUI27.ContractSearch LiveContractSearch 
         Height          =   5415
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9551
      End
      Begin TradingUI27.ContractSearch HistContractSearch 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   25
         Top             =   3000
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7858
      End
      Begin VB.CommandButton ConfigEditorButton 
         Caption         =   "Show config editor"
         Height          =   375
         Left            =   -72840
         TabIndex        =   32
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox CurrentConfigNameText 
         Height          =   285
         Left            =   -74640
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3375
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73080
         TabIndex        =   28
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -72360
         TabIndex        =   29
         ToolTipText     =   "Pause tickfile replay"
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -71640
         TabIndex        =   30
         ToolTipText     =   "Stop tickfile replay"
         Top             =   3480
         Width           =   615
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "fTradeSkilDemo.frx":1C6C
         Left            =   -73800
         List            =   "fTradeSkilDemo.frx":1C6E
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3000
         Width           =   2775
      End
      Begin VB.CheckBox HistSessionOnlyCheck 
         Caption         =   "Session only"
         Height          =   375
         Left            =   -72240
         TabIndex        =   21
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox NumHistBarsText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72000
         TabIndex        =   20
         Text            =   "500"
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Selected tickers"
         Height          =   1215
         Left            =   120
         TabIndex        =   37
         Top             =   6000
         Width           =   3615
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   3255
            TabIndex        =   38
            Top             =   240
            Width           =   3255
            Begin VB.CommandButton ChartButton1 
               Caption         =   "Chart"
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   2
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton StopTickerButton 
               Caption         =   "Sto&p"
               Enabled         =   0   'False
               Height          =   375
               Left            =   1080
               TabIndex        =   5
               Top             =   480
               Width           =   975
            End
            Begin VB.CommandButton OrderTicketButton 
               Caption         =   "&Order ticket"
               Enabled         =   0   'False
               Height          =   375
               Left            =   2160
               TabIndex        =   4
               Top             =   0
               Width           =   975
            End
            Begin VB.CommandButton MarketDepthButton 
               Caption         =   "&Mkt depth"
               Enabled         =   0   'False
               Height          =   375
               Left            =   1080
               TabIndex        =   3
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
         TabIndex        =   17
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox SessionOnlyCheck 
         Caption         =   "Session only"
         Height          =   375
         Left            =   -72240
         TabIndex        =   15
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox NumHistoryBarsText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72000
         TabIndex        =   14
         Text            =   "500"
         Top             =   840
         Width           =   975
      End
      Begin TradingUI27.TimeframeSelector LiveChartTimeframeSelector 
         Height          =   330
         Left            =   -73080
         TabIndex        =   13
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
      End
      Begin TradingUI27.TimeframeSelector HistTimeframeSelector 
         Height          =   330
         Left            =   -73080
         TabIndex        =   19
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker ToDatePicker 
         Height          =   375
         Left            =   -73080
         TabIndex        =   23
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd HH:mm"
         Format          =   20840451
         CurrentDate     =   39365
      End
      Begin MSComCtl2.DTPicker FromDatePicker 
         Height          =   375
         Left            =   -73080
         TabIndex        =   22
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd HH:mm"
         Format          =   20840451
         CurrentDate     =   39365
      End
      Begin MSComctlLib.ProgressBar ReplayProgressBar 
         Height          =   135
         Left            =   -74880
         TabIndex        =   44
         Top             =   4440
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin TWControls40.TWImageCombo HistChartStylesCombo 
         Height          =   330
         Left            =   -73080
         TabIndex        =   24
         Top             =   2520
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
         MouseIcon       =   "fTradeSkilDemo.frx":1C70
         Text            =   ""
      End
      Begin VB.Label Label8 
         Caption         =   "Style"
         Height          =   375
         Left            =   -74880
         TabIndex        =   61
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Style"
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Current configuration is:"
         Height          =   375
         Left            =   -74640
         TabIndex        =   47
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   -74880
         TabIndex        =   46
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   975
         Left            =   -74880
         TabIndex        =   45
         Top             =   4680
         Width           =   3855
      End
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74880
         TabIndex        =   43
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Number of history bars"
         Height          =   495
         Left            =   -74880
         TabIndex        =   42
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Timeframe"
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Number of history bars"
         Height          =   375
         Left            =   -74880
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Timeframe"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   33
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
   Begin TradingUI27.TickerGrid TickerGrid1 
      Height          =   4815
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8493
      AllowUserReordering=   3
      BackColorFixed  =   16053492
      Rows            =   100
      RowBackColorOdd =   16316664
      RowBackColorEven=   15658734
      GridLineWidth   =   0
      GridColorFixed  =   14737632
      ForeColorFixed  =   10526880
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
Implements StateChangeListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================
    
Private Const ModuleName                    As String = "fTradeSkilDemo"

Private Const ExecutionsTabCaptionLive      As String = "Live"
Private Const ExecutionsTabCaptionSimulated As String = "Simulated"
Private Const ExecutionsTabCaptionTickfile  As String = "Tickfile"

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

Private Enum FeaturesTabIndexNumbers
    FeaturesTabIndexOrders
    FeaturesTabIndexExecutions
    FeaturesTabIndexLog
End Enum

Private Enum OrdersTabIndexNumbers
    OrdersTabIndexLive = 1
    OrdersTabIndexSimulated
    OrderTabIndexTickfile
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI                   As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1

Private mTickers                                    As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private WithEvents mReplayController                As ReplayController
Attribute mReplayController.VB_VarHelpID = -1
Private WithEvents mTickfileReplayTC                As TaskController
Attribute mTickfileReplayTC.VB_VarHelpID = -1

Private mControlsHidden                             As Boolean
Private mFeaturesHidden                             As Boolean

Private mClockDisplay                               As ClockDisplay

Private mAppInstanceConfig                          As ConfigurationSection

Private WithEvents mOrderRecoveryFutureWaiter       As FutureWaiter
Attribute mOrderRecoveryFutureWaiter.VB_VarHelpID = -1
Private WithEvents mContractsFutureWaiter           As FutureWaiter
Attribute mContractsFutureWaiter.VB_VarHelpID = -1

Private mChartForms                                 As New ChartForms

Private mPreviousMainForm                           As fTradeSkilDemo

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
' ensure we get the Windows XP look and feel if running on XP
InitCommonControls
Set mOrderRecoveryFutureWaiter = New FutureWaiter
Set mContractsFutureWaiter = New FutureWaiter
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

setupLogging

Set mClockDisplay = New ClockDisplay
mClockDisplay.Initialise StatusBar1.Panels("datetime"), StatusBar1.Panels("timezone")
mClockDisplay.SetClock getDefaultClock

setupReplaySpeedCombo

FromDatePicker.Value = DateAdd("m", -1, Now)
FromDatePicker.Value = Empty    ' clear the checkbox
ToDatePicker.Value = Now

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_QueryUnload( _
                Cancel As Integer, _
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
On Error GoTo Err

Static prevHeight As Long
Static prevWidth As Long

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

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

LogMessage "Unloading main form"

LogMessage "Stopping tickfile replay"
' prevent event handler being fired on completion, which would
' reload the form again
Set mTickfileReplayTC = Nothing
stopTickfileReplay

LogMessage "Shutting down clock"
mClockDisplay.Finish

shutdown

LogMessage "Closing charts and market depth forms"
closeChartsAndMarketDepthForms

LogMessage "Closing config editor form"
gUnloadConfigEditor

LogMessage "Closing other forms"
Dim f As Form
For Each f In Forms
    If Not TypeOf f Is fTradeSkilDemo And Not TypeOf f Is fSplash Then
        LogMessage "Closing form: caption=" & f.caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

LogMessage "Stopping tickers"
If Not mTickers Is Nothing Then mTickers.Finish

killLoggingForThisForm

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

Private Sub LogListener_Notify(ByVal Logrec As LogRecord)
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

OrderTicketButton.Enabled = Not (getSelectedDataSource Is Nothing)
OrderTicket1Button.Enabled = OrderTicketButton.Enabled

Dim lDataSource As IMarketDataSource
Set lDataSource = ev.Source

Select Case ev.State
Case MarketDataSourceStates.MarketDataSourceStateCreated

Case MarketDataSourceStates.MarketDataSourceStateReady
    If lDataSource Is getSelectedDataSource Then mClockDisplay.SetClockFuture lDataSource.ClockFuture
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
    Else
        mClockDisplay.SetClockFuture getSelectedDataSource.ClockFuture
    End If
    
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

If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    Set op = LiveOrdersSummary.SelectedItem
ElseIf OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexSimulated Then
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
For Each lTicker In TickerGrid1.SelectedTickers
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

Private Sub ClosePositionsButton_Click()
Const ProcName As String = "ClosePositionsButton_Click"
On Error GoTo Err

If Not mTradeBuildAPI.ClosingPositions Then
    If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
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

Dim lNewAppInstanceConfig As ConfigurationSection
Set lNewAppInstanceConfig = gShowConfigEditor(mConfigStore, mAppInstanceConfig, Me)

If lNewAppInstanceConfig Is Nothing Then Exit Sub

shutdown
gLoadMainForm lNewAppInstanceConfig, Me

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
    If TickerGrid1.SelectedTickers.Count > 0 Then ChartButton1.Default = True
Case ControlsTabIndexNumbers.ControlsTabIndexTickfileReplay
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

Private Sub ControlsTabStrip_Click()
Const ProcName As String = "ControlsTabStrip_Click"
On Error GoTo Err

ControlsSSTab.SetFocus
ControlsSSTab.Tab = ControlsTabStrip.SelectedItem.Index - 1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ExecutionsSummaryTabStrip_Click()
Const ProcName As String = "ExecutionsSummaryTabStrip_Click"
On Error GoTo Err

Static currIndex As Long
If ExecutionsSummaryTabStrip.SelectedItem.Index = currIndex Then Exit Sub

Select Case ExecutionsSummaryTabStrip.SelectedItem.caption
Case ExecutionsTabCaptionLive
    LiveExecutionsSummary.Visible = True
    SimulatedExecutionsSummary.Visible = False
    TickfileExecutionsSummary.Visible = False
Case ExecutionsTabCaptionSimulated
    LiveExecutionsSummary.Visible = False
    SimulatedExecutionsSummary.Visible = True
    TickfileExecutionsSummary.Visible = False
Case ExecutionsTabCaptionTickfile
    LiveExecutionsSummary.Visible = False
    SimulatedExecutionsSummary.Visible = False
    TickfileExecutionsSummary.Visible = True
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Sub LiveContractSearch_Action()
Const ProcName As String = "LiveContractSearch_Action"
On Error GoTo Err

Dim lPreferredRow As Long
lPreferredRow = CLng(LiveContractSearch.Cookie)

Dim lContract As IContract
For Each lContract In LiveContractSearch.SelectedContracts
    TickerGrid1.StartTickerFromContract lContract, lPreferredRow
    If lPreferredRow <> 0 Then lPreferredRow = lPreferredRow + 1
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveContractSearch_NoContracts()
Const ProcName As String = "LiveContractSearch_NoContracts"
On Error GoTo Err

gModelessMsgBox "No contracts found", vbExclamation, "Attention"

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
On Error GoTo Err

Dim lTicker As Ticker
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

Dim os As OrdersSummary

If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    Set os = LiveOrdersSummary
ElseIf OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexSimulated Then
    Set os = SimulatedOrdersSummary
Else
    Set os = TickfileOrdersSummary
End If

If os.SelectedItem Is Nothing Then
    ModifyOrderPlexButton.Enabled = False
ElseIf os.IsSelectedItemModifiable Then
    getOrderTicket.Show vbModeless, Me
    getOrderTicket.ShowBracketOrder os.SelectedItem, os.SelectedOrderIndex
End If

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

Private Sub OrdersSummaryTabStrip_Click()
Const ProcName As String = "OrdersSummaryTabStrip_Click"
On Error GoTo Err

Static currIndex As Long
If OrdersSummaryTabStrip.SelectedItem.Index = currIndex Then Exit Sub

Select Case OrdersSummaryTabStrip.SelectedItem.Index
Case OrdersTabIndexNumbers.OrdersTabIndexLive
    LiveOrdersSummary.Visible = True
    SimulatedOrdersSummary.Visible = False
    TickfileOrdersSummary.Visible = False
    setOrdersSelection LiveOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrdersTabIndexLive
Case OrdersTabIndexNumbers.OrdersTabIndexSimulated
    LiveOrdersSummary.Visible = False
    SimulatedOrdersSummary.Visible = True
    TickfileOrdersSummary.Visible = False
    setOrdersSelection SimulatedOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrdersTabIndexSimulated
Case OrdersTabIndexNumbers.OrderTabIndexTickfile
    LiveOrdersSummary.Visible = False
    SimulatedOrdersSummary.Visible = False
    TickfileOrdersSummary.Visible = True
    setOrdersSelection TickfileOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrderTabIndexTickfile
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderTicket1Button_Click()
Const ProcName As String = "OrderTicket1Button_Click"
On Error GoTo Err

setupOrderTicket

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderTicketButton_Click()
Const ProcName As String = "OrderTicketButton_Click"
On Error GoTo Err

setupOrderTicket

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
    TickfileOrdersSummary.Initialise lTickfileDataManager
    
    Dim lOrderManager As New OrderManager
    TickfileOrdersSummary.MonitorPositions lOrderManager.PositionManagersSimulated
    TickfileExecutionsSummary.MonitorPositions lOrderManager.PositionManagersSimulated
    
    Set mReplayController = lTickfileDataManager.ReplayController
    
    Dim lTickers As Tickers
    Set lTickers = CreateTickers(lTickfileDataManager, mTradeBuildAPI.StudyLibraryManager, mTradeBuildAPI.HistoricalDataStoreInput, lOrderManager, , mTradeBuildAPI.OrderSubmitterFactorySimulated)
    
    Dim i As Long
    For i = 1 To TickfileOrganiser1.TickfileCount
        Dim lTicker As Ticker
        Set lTicker = lTickers.CreateTicker(mReplayController.TickStream(i - 1).ContractFuture, False)
        TickerGrid1.AddTickerFromDataSource lTicker
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

Private Sub SimulatedOrdersSummary_SelectionChanged()
Const ProcName As String = "SimulatedOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection SimulatedOrdersSummary

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

Private Sub TickerGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "TickerGrid1_KeyUp"
On Error GoTo Err

Select Case KeyCode
Case vbKeyDelete
    StopSelectedTickers
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickerGrid1_TickerSelectionChanged()
Const ProcName As String = "TickerGrid1_TickerSelectionChanged"
On Error GoTo Err

handleSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickerGrid1_TickerSymbolEntered(ByVal pSymbol As String, ByVal pPreferredRow As Long)
Const ProcName As String = "TickerGrid1_TickerSymbolEntered"
On Error GoTo Err

mContractsFutureWaiter.Add FetchContracts(CreateContractSpecifier(, pSymbol), mTradeBuildAPI.ContractStorePrimary, mTradeBuildAPI.ContractStoreSecondary), pPreferredRow

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickfileOrdersSummary_SelectionChanged()
Const ProcName As String = "SimulatedOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection TickfileOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickfileOrganiser1_TickfileCountChanged()
Const ProcName As String = "TickfileOrganiser1_TickfileCountChanged"
On Error GoTo Err

If TickfileOrganiser1.TickfileCount = 0 Then
    PlayTickFileButton.Enabled = False
    PauseReplayButton.Enabled = False
    StopReplayButton.Enabled = False
    ChartButton.Enabled = False
    ChartButton1.Enabled = False
Else
    PlayTickFileButton.Enabled = True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mContractsFutureWaiter Event Handlers
'================================================================================

Private Sub mContractsFutureWaiter_WaitCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mContractsFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lContracts As IContracts
Set lContracts = ev.Future.Value

If lContracts.Count = 1 Then
    TickerGrid1.StartTickerFromContract lContracts.ItemAtIndex(1), CLng(ev.ContinuationData)
Else
    LiveContractSearch.LoadContracts lContracts, CLng(ev.ContinuationData)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mOrderRecoveryFutureWaiter Event Handlers
'================================================================================

Private Sub mOrderRecoveryFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mOrderRecoveryFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    LogMessage "Order recovery failed"
ElseIf ev.Future.IsAvailable Then
    LogMessage "Order recovery completed    "
    loadAppInstanceConfig
    
    Me.Show vbModeless
    
    gUnloadSplashScreen
    
    If Not mPreviousMainForm Is Nothing Then
        Unload mPreviousMainForm
        Set mPreviousMainForm = Nothing
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

''================================================================================
'' mTickers Event Handlers
''================================================================================
'
'Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
'Const ProcName As String = "mTickers_CollectionChanged"
'On Error GoTo Err
'
'If ev.ChangeType <> CollItemAdded Then Exit Sub
'
'Dim lTicker As Ticker
'Set lTicker = ev.AffectedItem
'TickerGrid1.AddTickerFromDataSource lTicker
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

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
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_Notification( _
                ByRef ev As NotificationEventData)
Const ProcName As String = "mTradeBuildAPI_Notification"
On Error GoTo Err

Select Case ev.EventCode
Case ApiNotifyCodes.ApiNotifyServiceProviderError
    Dim spError As ServiceProviderError
    Set spError = mTradeBuildAPI.GetServiceProviderError
    LogMessage "Error from " & _
                        spError.ServiceProviderName & _
                        ": code " & spError.ErrorCode & _
                        ": " & spError.Message

Case Else
    LogMessage "Notification: code=" & ev.EventCode & "; source=" & TypeName(ev.Source) & ": " & _
                ev.EventMessage & vbCrLf
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

Friend Function Initialise( _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByVal pPreviousMainForm As fTradeSkilDemo, _
                ByRef pErrorMessage As String) As Boolean
Const ProcName As String = "initialise"
On Error GoTo Err

Set mTradeBuildAPI = pTradeBuildAPI
Set mConfigStore = pConfigStore
Set mAppInstanceConfig = pAppInstanceConfig
Set mPreviousMainForm = pPreviousMainForm

LogMessage "Loading configuration: " & mAppInstanceConfig.InstanceQualifier

mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionApplication

Set mTickers = mTradeBuildAPI.Tickers
If mTickers Is Nothing Then
    pErrorMessage = "No tickers object is available: one or more service providers may be missing or disabled"
    Initialise = False
    Exit Function
End If

LogMessage "Recovering orders from last session"
mOrderRecoveryFutureWaiter.Add CreateFutureFromTask(mTradeBuildAPI.RecoverOrders())

Initialise = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub MakeVisible()
Const ProcName As String = "MakeVisible"
On Error GoTo Err

Me.Show
If Not mControlsHidden Then ControlsTabStrip.Tabs(1).Selected = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyInstanceSettings()
Const ProcName As String = "applyInstanceSettings"
On Error GoTo Err

LogMessage "Loading configuration: positioning main form"
Select Case mAppInstanceConfig.GetSetting(ConfigSettingMainFormWindowstate, WindowStateNormal)
Case WindowStateMaximized
    Me.WindowState = FormWindowStateConstants.vbMaximized
Case WindowStateMinimized
    Me.WindowState = FormWindowStateConstants.vbMinimized
Case WindowStateNormal
    Me.left = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormLeft, 0)) * Screen.TwipsPerPixelX
    Me.Top = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormTop, 0)) * Screen.TwipsPerPixelY
    Me.Width = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
    Me.Height = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
End Select

mControlsHidden = CBool(mAppInstanceConfig.GetSetting(ConfigSettingMainFormControlsHidden, CStr(False)))
If mControlsHidden Then
    hideControls
Else
    showControls
End If

mFeaturesHidden = CBool(mAppInstanceConfig.GetSetting(ConfigSettingMainFormFeaturesHidden, CStr(False)))
If mFeaturesHidden Then
    hideFeatures
Else
    showFeatures
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearSelectedTickers()
Const ProcName As String = "clearSelectedTickers"
On Error GoTo Err

TickerGrid1.DeselectSelectedTickers
handleSelectedTickers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub closeChartsAndMarketDepthForms()
Const ProcName As String = "closeChartsAndMarketDepthForms"
On Error GoTo Err

mChartForms.Finish

Dim f As Form
For Each f In Forms
    If TypeOf f Is fMarketDepth Then
        LogMessage "Closing form: caption=" & f.caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
On Error GoTo Err

Static formatter As LogFormatter
If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDefaultClock() As Clock
Const ProcName As String = "getDefaultClock"
On Error GoTo Err

Static sClock As Clock
If sClock Is Nothing Then Set sClock = GetClock("") ' create a clock running local time
Set getDefaultClock = sClock

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getOrderTicket() As fOrderTicket
Const ProcName As String = "getOrderTicket"
On Error GoTo Err

Static sOrderTicket As fOrderTicket

If sOrderTicket Is Nothing Then
    Set sOrderTicket = New fOrderTicket
    sOrderTicket.Initialise mAppInstanceConfig
End If
Set getOrderTicket = sOrderTicket

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSelectedDataSource() As IMarketDataSource
Const ProcName As String = "getSelectedDataSource"
On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 1 Then Set getSelectedDataSource = TickerGrid1.SelectedTickers.Item(1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleSelectedTickers()
Const ProcName As String = "handleSelectedTickers"
On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 0 Then
    StopTickerButton.Enabled = False
    ChartButton.Enabled = False
    ChartButton1.Enabled = False
    MarketDepthButton.Enabled = False
    OrderTicketButton.Enabled = False
    OrderTicket1Button.Enabled = False
    mClockDisplay.SetClock getDefaultClock
Else
    StopTickerButton.Enabled = True
    
    ChartButton.Enabled = False
    ChartButton1.Enabled = False
    MarketDepthButton.Enabled = False
    OrderTicketButton.Enabled = False
    OrderTicket1Button.Enabled = False
    
    If ControlsSSTab.Tab = ControlsTabIndexNumbers.ControlsTabIndexLiveCharts Then
        ChartButton.Default = True
    ElseIf ControlsSSTab.Tab = ControlsTabIndexNumbers.ControlsTabIndexTickers Then
        ChartButton1.Default = True
    End If
    
    Dim lTicker As Ticker
    Set lTicker = getSelectedDataSource
    If lTicker Is Nothing Then
        mClockDisplay.SetClock getDefaultClock
    ElseIf lTicker.State = MarketDataSourceStateRunning Then
        mClockDisplay.SetClockFuture lTicker.ClockFuture
        ChartButton.Enabled = True
        ChartButton1.Enabled = True
        Dim lContract As IContract
        Set lContract = lTicker.ContractFuture.Value
        If (lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled) And lContract.Specifier.SecType <> SecTypeIndex Then
            OrderTicketButton.Enabled = True
            OrderTicket1Button.Enabled = True
            MarketDepthButton.Enabled = True
        End If
    Else
        mClockDisplay.SetClock getDefaultClock
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub killLoggingForThisForm()
Const ProcName As String = "killLoggingForThisForm"
On Error GoTo Err

GetLogger("log").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadAppInstanceConfig()
Const ProcName As String = "loadAppInstanceConfig"
On Error GoTo Err

LogMessage "Loading configuration: " & mAppInstanceConfig.InstanceQualifier

LogMessage "Loading configuration: Setting up tickfile organiser"
setupTickfileOrganiser

LogMessage "Loading configuration: Setting up contract search"
setupContractSearch

LogMessage "Loading configuration: Setting up ticker grid"
setupTickerGrid

LogMessage "Loading configuration: Setting up order summaries"
setupOrderSummaries

LogMessage "Loading configuration: Setting up execution summaries"
setupExecutionSummaries

LogMessage "Loading configuration: Setting up timeframeselectors"
setupTimeframeSelectors

applyInstanceSettings

LogMessage "Loading configuration: loading tickers into ticker grid"
TickerGrid1.LoadFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionTickerGrid)

LogMessage "Loading configuration: loading default study configurations"
LoadDefaultStudyConfigurationsFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionDefaultStudyConfigs)

LogMessage "Loading configuration: setting current chart style"
setCurrentChartStyles

LogMessage "Loading configuration: creating charts"
startCharts

LogMessage "Loading configuration: creating historical charts"
startHistoricalCharts

FeaturesSSTAB.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexOrders

LogMessage "Loaded configuration: " & mAppInstanceConfig.InstanceQualifier
CurrentConfigNameText = mAppInstanceConfig.InstanceQualifier
Me.caption = gAppTitle & _
            " - " & mAppInstanceConfig.InstanceQualifier

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

Private Sub Resize()
Const ProcName As String = "Resize"
On Error GoTo Err

Dim left As Long
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

HideFeaturesPicture.Move FeaturesSSTAB.left + FeaturesSSTAB.Width - HideFeaturesPicture.Width - 2 * Screen.TwipsPerPixelX, _
                        FeaturesSSTAB.Top + Screen.TwipsPerPixelY
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
    SimulatedOrdersSummary.Width = LiveOrdersSummary.Width
    TickfileOrdersSummary.Width = LiveOrdersSummary.Width
Else
    OrderTicket1Button.left = FeaturesSSTAB.Width - OrderTicket1Button.Width - 120 - SSTabInactiveControlAdjustment
    ModifyOrderPlexButton.left = FeaturesSSTAB.Width - ModifyOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    CancelOrderPlexButton.left = FeaturesSSTAB.Width - CancelOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    ClosePositionsButton.left = FeaturesSSTAB.Width - CancelOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    
    LiveOrdersSummary.Width = ModifyOrderPlexButton.left + SSTabInactiveControlAdjustment - 120 - 120
    SimulatedOrdersSummary.Width = LiveOrdersSummary.Width
    TickfileOrdersSummary.Width = LiveOrdersSummary.Width
End If

LogText.Width = FeaturesSSTAB.Width - 120 - 120
LiveExecutionsSummary.Width = FeaturesSSTAB.Width - 120 - 120
SimulatedExecutionsSummary.Width = FeaturesSSTAB.Width - 120 - 120

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
    If Not ChartStylesManager.Contains(ChartStyleNameAppDefault) Then setupChartStyles pComboItems
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupChartStyles(ByVal pComboItems As ComboItems)
Const ProcName As String = "setupChartStyles"
On Error GoTo Err

setupChartStyleAppDefault
pComboItems.Add , ChartStyleNameAppDefault, ChartStyleNameAppDefault

setupChartStyleBlack
pComboItems.Add , ChartStyleNameBlack, ChartStyleNameBlack

setupChartStyleDarkBlueFade
pComboItems.Add , ChartStyleNameDarkBlueFade, ChartStyleNameDarkBlueFade

setupChartStyleGoldFade
pComboItems.Add , ChartStyleNameGoldFade, ChartStyleNameGoldFade

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupChartStyleAppDefault()
Const ProcName As String = "setupChartStyleAppDefault"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameAppDefault) Then Exit Sub

Dim lCursorTextStyle As New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = &H80&
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0

Dim lFont As New StdFont
lFont.name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = RGB(192, 192, 192)
GradientFillColors(1) = RGB(248, 248, 248)
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors

lDefaultRegionStyle.SessionEndGridLineStyle.Color = &HD0D0D0
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &HD0D0D0
lDefaultRegionStyle.XGridLineStyle.Color = &HD0D0D0
lDefaultRegionStyle.YGridLineStyle.Color = &HD0D0D0
    
Dim lxAxisRegionStyle As ChartRegionStyle
Set lxAxisRegionStyle = GetDefaultChartXAxisRegionStyle.Clone
lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle
GradientFillColors(0) = RGB(230, 236, 207)
GradientFillColors(1) = RGB(222, 236, 215)
lxAxisRegionStyle.BackGradientFillColors = GradientFillColors
    
Dim lDefaultYAxisRegionStyle As ChartRegionStyle
Set lDefaultYAxisRegionStyle = GetDefaultChartYAxisRegionStyle.Clone
lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle
GradientFillColors(0) = RGB(234, 246, 254)
GradientFillColors(1) = RGB(226, 246, 255)
lDefaultYAxisRegionStyle.BackGradientFillColors = GradientFillColors
    
Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = &H7F

ChartStylesManager.Add ChartStyleNameAppDefault, _
                        ChartStylesManager.DefaultStyle, _
                        lDefaultRegionStyle, _
                        lxAxisRegionStyle, _
                        lDefaultYAxisRegionStyle, _
                        lCrosshairLineStyle


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupChartStyleBlack()
Const ProcName As String = "setupChartStyleBlack"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameBlack) Then Exit Sub

Dim lCursorTextStyle As New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = vbRed
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0

Dim lFont As StdFont
Set lFont = New StdFont
lFont.name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = &H202020
GradientFillColors(1) = &H202020
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors

lDefaultRegionStyle.XGridLineStyle.Color = &H303030
lDefaultRegionStyle.YGridLineStyle.Color = &H303030
lDefaultRegionStyle.SessionEndGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle.Color = &H303030
lDefaultRegionStyle.SessionStartGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &H303030
    
Dim lxAxisRegionStyle As ChartRegionStyle
Set lxAxisRegionStyle = GetDefaultChartXAxisRegionStyle.Clone
GradientFillColors(0) = RGB(0, 0, 0)
GradientFillColors(1) = RGB(0, 0, 0)
lxAxisRegionStyle.BackGradientFillColors = GradientFillColors
lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle

Dim lGridTextStyle As New TextStyle
lGridTextStyle.Box = True
lGridTextStyle.BoxFillWithBackgroundColor = True
lGridTextStyle.BoxStyle = LineInvisible
lGridTextStyle.Color = &HD0D0D0
lxAxisRegionStyle.XGridTextStyle = lGridTextStyle
    
Dim lDefaultYAxisRegionStyle As ChartRegionStyle
Set lDefaultYAxisRegionStyle = GetDefaultChartYAxisRegionStyle.Clone
GradientFillColors(0) = RGB(0, 0, 0)
GradientFillColors(1) = RGB(0, 0, 0)
lDefaultYAxisRegionStyle.BackGradientFillColors = GradientFillColors
lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle
lDefaultYAxisRegionStyle.YGridTextStyle = lGridTextStyle
    
Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = &H80&

ChartStylesManager.Add ChartStyleNameBlack, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        lxAxisRegionStyle, _
                        lDefaultYAxisRegionStyle, _
                        lCrosshairLineStyle


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupChartStyleDarkBlueFade()
Const ProcName As String = "setupChartStyleDarkBlueFade"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameDarkBlueFade) Then Exit Sub

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = &H643232
GradientFillColors(1) = &H804040
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors
    
lDefaultRegionStyle.XGridLineStyle.Color = &H505050
lDefaultRegionStyle.YGridLineStyle.Color = &H505050
    
lDefaultRegionStyle.SessionEndGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle.Color = &H505050
    
lDefaultRegionStyle.SessionStartGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &H505050

Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = vbRed

ChartStylesManager.Add ChartStyleNameDarkBlueFade, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        , _
                        , _
                        lCrosshairLineStyle


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupChartStyleGoldFade()
Const ProcName As String = "setupChartStyleGoldFade"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameGoldFade) Then Exit Sub

Dim lCursorTextStyle As New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = &H80&
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0

Dim lFont As New StdFont
lFont.name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = &H82DFE6
GradientFillColors(1) = &HEBFAFB
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors
    
lDefaultRegionStyle.XGridLineStyle.Color = &HE0E0E0
lDefaultRegionStyle.YGridLineStyle.Color = &HE0E0E0
    
lDefaultRegionStyle.SessionEndGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle.Color = &HE0E0E0
    
lDefaultRegionStyle.SessionStartGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &HE0E0E0

Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = 127

ChartStylesManager.Add ChartStyleNameGoldFade, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        , _
                        , _
                        lCrosshairLineStyle


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

Private Sub setupExecutionSummaries()
Const ProcName As String = "setupExecutionSummaries"
On Error GoTo Err

Do While ExecutionsSummaryTabStrip.Tabs.Count > 0
    ExecutionsSummaryTabStrip.Tabs.Remove 1
Loop

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    SimulatedExecutionsSummary.Visible = True
    ExecutionsSummaryTabStrip.Tabs.Add 1, , ExecutionsTabCaptionSimulated
Else
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated
    SimulatedExecutionsSummary.Visible = False
    LiveExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveExecutionsSummary.Visible = True
    ExecutionsSummaryTabStrip.Tabs.Add 1, , ExecutionsTabCaptionLive
    ExecutionsSummaryTabStrip.Tabs.Add 2, , ExecutionsTabCaptionSimulated
End If

If Not mTradeBuildAPI.TickfileStoreInput Is Nothing Then
    TickfileExecutionsSummary.Visible = False
    ExecutionsSummaryTabStrip.Tabs.Add ExecutionsSummaryTabStrip.Tabs.Count + 1, , ExecutionsTabCaptionTickfile
End If

If ExecutionsSummaryTabStrip.Tabs.Count = 1 Then
    ExecutionsSummaryTabStrip.Visible = False
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top + ExecutionsSummaryTabStrip.Height - SimulatedExecutionsSummary.Top
Else
    ExecutionsSummaryTabStrip.Visible = True
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    LiveExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    TickfileExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
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

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top + OrdersSummaryTabStrip.Height - SimulatedOrdersSummary.Top
    SimulatedOrdersSummary.Visible = True
    
    LiveOrdersSummary.Visible = False
    
    OrdersSummaryTabStrip.Visible = False
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexSimulated).Selected = True
Else
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top - SimulatedOrdersSummary.Top
    
    LiveOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
    LiveOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveOrdersSummary.Height = SimulatedOrdersSummary.Height
    
    OrdersSummaryTabStrip.Visible = True
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexLive).Selected = True
End If

SimulatedOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
SimulatedOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupOrderTicket()
Const ProcName As String = "setupOrderTicket"
On Error GoTo Err

If getSelectedDataSource Is Nothing Then
    gModelessMsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
Else
    getOrderTicket.Show vbModeless, Me
    getOrderTicket.Ticker = getSelectedDataSource
End If

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

Private Sub setupTickerGrid()
Const ProcName As String = "setupTickerGrid"
On Error GoTo Err

TickerGrid1.Initialise mTickers

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
gHandleUnexpectedError ProcName, ModuleName
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

Private Sub shutdown()
Const ProcName As String = "Shutdown"
On Error GoTo Err

Static sAlreadyShutdown As Boolean
If sAlreadyShutdown Then Exit Sub

sAlreadyShutdown = True

LogMessage "Finishing UI controls"
finishUIControls

LogMessage "Removing service providers"
mTradeBuildAPI.ServiceProviders.RemoveAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub startCharts()
Const ProcName As String = "startCharts"
On Error GoTo Err

mChartForms.LoadChartsFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionCharts), _
                                mTickers, _
                                mTradeBuildAPI.BarFormatterLibManager, _
                                mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator, _
                                gMainForm

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub startHistoricalCharts()
Const ProcName As String = "startHistoricalCharts"
On Error GoTo Err

mChartForms.LoadHistoricalChartsFromConfig _
                    mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionHistoricCharts), _
                    mTradeBuildAPI.StudyLibraryManager, _
                    mTradeBuildAPI.HistoricalDataStoreInput, _
                    mTradeBuildAPI.BarFormatterLibManager, _
                    gMainForm
                    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub StopSelectedTickers()
Const ProcName As String = "StopSelectedTickers"
On Error GoTo Err

Dim lTickers As SelectedTickers
Set lTickers = TickerGrid1.SelectedTickers

TickerGrid1.StopSelectedTickers

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

Private Sub updateInstanceSettings()
Const ProcName As String = "updateInstanceSettings"
On Error GoTo Err

If Not mAppInstanceConfig Is Nothing Then
    mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionMainForm
    Select Case Me.WindowState
    Case FormWindowStateConstants.vbMaximized
        mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMaximized
    Case FormWindowStateConstants.vbMinimized
        mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMinimized
    Case FormWindowStateConstants.vbNormal
        mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateNormal
        mAppInstanceConfig.SetSetting ConfigSettingMainFormLeft, Me.left / Screen.TwipsPerPixelX
        mAppInstanceConfig.SetSetting ConfigSettingMainFormTop, Me.Top / Screen.TwipsPerPixelY
        mAppInstanceConfig.SetSetting ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX
        mAppInstanceConfig.SetSetting ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY
    End Select
    
    mAppInstanceConfig.SetSetting ConfigSettingMainFormControlsHidden, CStr(mControlsHidden)
    mAppInstanceConfig.SetSetting ConfigSettingMainFormFeaturesHidden, CStr(mFeaturesHidden)
    
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


