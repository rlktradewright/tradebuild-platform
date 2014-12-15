VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#279.0#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#27.1#0"; "TWControls40.ocx"
Begin VB.UserControl InfoPanel 
   Appearance      =   0  'Flat
   BackColor       =   &H00CDF3FF&
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12285
   DefaultCancel   =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   12285
   Begin TabDlg.SSTab InfoSSTab 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   -2147483630
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
      TabPicture(0)   =   "InfoPanel.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TickfileOrdersSummary"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SimulatedOrdersSummary"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LiveOrdersSummary"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OrdersSummaryTabStrip"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OrderTicket1Button"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ModifyOrderPlexButton"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CancelOrderPlexButton"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ClosePositionsButton"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&2. Executions"
      TabPicture(1)   =   "InfoPanel.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LiveExecutionsSummary"
      Tab(1).Control(1)=   "ExecutionsSummaryTabStrip"
      Tab(1).Control(2)=   "SimulatedExecutionsSummary"
      Tab(1).Control(3)=   "TickfileExecutionsSummary"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&3. Log"
      TabPicture(2)   =   "InfoPanel.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LogText"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox LogText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Status messages"
         Top             =   120
         Width           =   11955
      End
      Begin TWControls40.TWButton ClosePositionsButton 
         Height          =   495
         Left            =   11160
         TabIndex        =   1
         Top             =   3480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Close all positions!"
      End
      Begin TWControls40.TWButton CancelOrderPlexButton 
         Height          =   495
         Left            =   11160
         TabIndex        =   2
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
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
         Caption         =   "&Cancel"
      End
      Begin TWControls40.TWButton ModifyOrderPlexButton 
         Height          =   495
         Left            =   11160
         TabIndex        =   3
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
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
         Caption         =   "&Modify"
      End
      Begin TWControls40.TWButton OrderTicket1Button 
         Height          =   495
         Left            =   11160
         TabIndex        =   4
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
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
         Caption         =   "Order Ticket"
      End
      Begin MSComctlLib.TabStrip OrdersSummaryTabStrip 
         Height          =   375
         Left            =   120
         TabIndex        =   6
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
         TabIndex        =   8
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.ExecutionsSummary LiveExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   9
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6376
      End
      Begin MSComctlLib.TabStrip ExecutionsSummaryTabStrip 
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   120
         Width           =   11995
         _ExtentX        =   21167
         _ExtentY        =   6376
      End
      Begin TradingUI27.OrdersSummary TickfileOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.ExecutionsSummary TickfileExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   13
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6376
      End
   End
End
Attribute VB_Name = "InfoPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

