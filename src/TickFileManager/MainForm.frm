VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   Caption         =   "TradeBuild Tickfile Manager"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox StatusText 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Status messages"
      Top             =   5160
      Width           =   10695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tickfile selection"
      TabPicture(0)   =   "MainForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ReplayContractLabel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ReplayProgressLabel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label15"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ReplayProgressBar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SelectTickfilesButton"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ClearTickfileListButton"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TickfileList"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "StopButton"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ConvertButton"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "QTServerText"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "QTPortText"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "OutputPathText"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "OutputPathButton"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Contract details"
      TabPicture(1)   =   "MainForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label17"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label21"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ServerText"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "ClientIDText"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "PortText"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "StrikePriceText"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "ExchangeText"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "ExpiryText"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "SymbolText"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "TypeCombo"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "RightCombo"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "GetContractButton"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "DisconnectButton"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "ContractDetailsText"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "Bar output"
      TabPicture(2)   =   "MainForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(3)=   "Frame7"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame7 
         Caption         =   "Period 4"
         Height          =   4095
         Left            =   -66960
         TabIndex        =   92
         Top             =   480
         Width           =   2535
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3735
            ScaleWidth      =   2295
            TabIndex        =   93
            Top             =   240
            Width           =   2295
            Begin VB.CheckBox IncludeBidAndAskCheck 
               Caption         =   "Include bid and ask"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   102
               Top             =   1920
               Width           =   2175
            End
            Begin VB.CheckBox EnableCheck 
               Caption         =   "Enable"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   96
               Top             =   120
               Width           =   2175
            End
            Begin VB.TextBox BarLengthText 
               Height          =   285
               Index           =   3
               Left            =   1320
               TabIndex        =   95
               Text            =   "60"
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox SaveIntervalText 
               Height          =   285
               Index           =   3
               Left            =   1320
               TabIndex        =   94
               Text            =   "60"
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label28 
               Caption         =   "Bar length (minutes)"
               Height          =   375
               Left            =   240
               TabIndex        =   98
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label27 
               Caption         =   "Save interval (seconds)"
               Height          =   615
               Left            =   240
               TabIndex        =   97
               Top             =   1080
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Period 3"
         Height          =   4095
         Left            =   -69600
         TabIndex        =   85
         Top             =   480
         Width           =   2535
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3735
            ScaleWidth      =   2295
            TabIndex        =   86
            Top             =   240
            Width           =   2295
            Begin VB.CheckBox IncludeBidAndAskCheck 
               Caption         =   "Include bid and ask"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   101
               Top             =   1920
               Width           =   2175
            End
            Begin VB.TextBox SaveIntervalText 
               Height          =   285
               Index           =   2
               Left            =   1320
               TabIndex        =   89
               Text            =   "60"
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox BarLengthText 
               Height          =   285
               Index           =   2
               Left            =   1320
               TabIndex        =   88
               Text            =   "15"
               Top             =   600
               Width           =   855
            End
            Begin VB.CheckBox EnableCheck 
               Caption         =   "Enable"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   87
               Top             =   120
               Width           =   2175
            End
            Begin VB.Label Label26 
               Caption         =   "Save interval (seconds)"
               Height          =   615
               Left            =   240
               TabIndex        =   91
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label25 
               Caption         =   "Bar length (minutes)"
               Height          =   375
               Left            =   240
               TabIndex        =   90
               Top             =   600
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Period 2"
         Height          =   4095
         Left            =   -72240
         TabIndex        =   78
         Top             =   480
         Width           =   2535
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3735
            ScaleWidth      =   2295
            TabIndex        =   79
            Top             =   240
            Width           =   2295
            Begin VB.CheckBox IncludeBidAndAskCheck 
               Caption         =   "Include bid and ask"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   100
               Top             =   1920
               Width           =   2175
            End
            Begin VB.CheckBox EnableCheck 
               Caption         =   "Enable"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   82
               Top             =   120
               Width           =   2175
            End
            Begin VB.TextBox BarLengthText 
               Height          =   285
               Index           =   1
               Left            =   1320
               TabIndex        =   81
               Text            =   "5"
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox SaveIntervalText 
               Height          =   285
               Index           =   1
               Left            =   1320
               TabIndex        =   80
               Text            =   "30"
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label24 
               Caption         =   "Bar length (minutes)"
               Height          =   375
               Left            =   240
               TabIndex        =   84
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "Save interval (seconds)"
               Height          =   615
               Left            =   240
               TabIndex        =   83
               Top             =   1080
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Period 1"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   71
         Top             =   480
         Width           =   2535
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3735
            ScaleWidth      =   2295
            TabIndex        =   72
            Top             =   240
            Width           =   2295
            Begin VB.CheckBox IncludeBidAndAskCheck 
               Caption         =   "Include bid and ask"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   99
               Top             =   1920
               Width           =   2175
            End
            Begin VB.TextBox SaveIntervalText 
               Height          =   285
               Index           =   0
               Left            =   1320
               TabIndex        =   77
               Text            =   "15"
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox BarLengthText 
               Height          =   285
               Index           =   0
               Left            =   1320
               TabIndex        =   75
               Text            =   "1"
               Top             =   600
               Width           =   855
            End
            Begin VB.CheckBox EnableCheck 
               Caption         =   "Enable"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   73
               Top             =   120
               Width           =   2175
            End
            Begin VB.Label Label22 
               Caption         =   "Save interval (seconds)"
               Height          =   615
               Left            =   240
               TabIndex        =   76
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label19 
               Caption         =   "Bar length (minutes)"
               Height          =   375
               Left            =   240
               TabIndex        =   74
               Top             =   600
               Width           =   735
            End
         End
      End
      Begin VB.CommandButton OutputPathButton 
         Caption         =   "..."
         Height          =   375
         Left            =   9840
         TabIndex        =   5
         ToolTipText     =   "Select output path"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox OutputPathText 
         Height          =   285
         Left            =   7800
         TabIndex        =   6
         ToolTipText     =   "Location of output tickfiles"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox QTPortText 
         Height          =   285
         Left            =   9240
         TabIndex        =   3
         ToolTipText     =   "Port for connecting to QuoteTracker"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox QTServerText 
         Height          =   285
         Left            =   9240
         TabIndex        =   2
         ToolTipText     =   "Name or address of computer hosting QuoteTracker"
         Top             =   600
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Timestamps"
         Height          =   1335
         Left            =   7800
         TabIndex        =   63
         Top             =   3360
         Width           =   1455
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   1215
            TabIndex        =   64
            Top             =   240
            Width           =   1215
            Begin VB.TextBox AdjustSecondsEndText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   9
               Text            =   "0"
               ToolTipText     =   "Timestamp adjustment (seconds) at end of file"
               Top             =   645
               Width           =   495
            End
            Begin VB.TextBox AdjustSecondsStartText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   8
               Text            =   "0"
               ToolTipText     =   "Timestamp adjustment (seconds) at start of file"
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox AdjustTimestampsCheck 
               Caption         =   "Adjust timestamps?"
               Height          =   375
               Left            =   0
               TabIndex        =   7
               ToolTipText     =   "Set if timestamps are to be adjusted"
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label Label12 
               Caption         =   "End"
               Height          =   255
               Left            =   720
               TabIndex        =   66
               Top             =   645
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Start"
               Height          =   255
               Left            =   720
               TabIndex        =   65
               Top             =   360
               Width           =   495
            End
         End
      End
      Begin VB.TextBox ContractDetailsText 
         Height          =   2535
         Left            =   -74640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3855
      End
      Begin VB.CommandButton DisconnectButton 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -67440
         TabIndex        =   24
         ToolTipText     =   "Disconnect from service provider"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton GetContractButton 
         Caption         =   "Get contract details"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -68880
         TabIndex        =   23
         ToolTipText     =   "Get contract details from specified source"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Output format"
         Height          =   1095
         Left            =   7800
         TabIndex        =   59
         Top             =   1440
         Width           =   2535
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   2295
            TabIndex        =   60
            Top             =   240
            Width           =   2295
            Begin VB.ListBox FormatList 
               Height          =   645
               ItemData        =   "MainForm.frx":0054
               Left            =   0
               List            =   "MainForm.frx":0056
               TabIndex        =   4
               ToolTipText     =   "Select output tickfile format"
               Top             =   0
               Width           =   2295
            End
         End
      End
      Begin VB.ComboBox RightCombo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "MainForm.frx":0058
         Left            =   -68880
         List            =   "MainForm.frx":005A
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3480
         Width           =   855
      End
      Begin VB.ComboBox TypeCombo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "MainForm.frx":005C
         Left            =   -68880
         List            =   "MainForm.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox SymbolText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox ExpiryText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox ExchangeText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   20
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox StrikePriceText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   21
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contract details source"
         Height          =   1095
         Left            =   -74640
         TabIndex        =   51
         Top             =   600
         Width           =   2535
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   2295
            TabIndex        =   52
            Top             =   240
            Width           =   2295
            Begin VB.OptionButton ContractFromServiceProviderOption 
               Caption         =   "Service provider"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               ToolTipText     =   "Get contract details from service provider"
               Top             =   480
               Width           =   1455
            End
            Begin VB.OptionButton ContractInTickfileOption 
               Caption         =   "In tickfile"
               Height          =   195
               Left            =   120
               TabIndex        =   12
               ToolTipText     =   "Tickfile contains contract details"
               Top             =   0
               Value           =   -1  'True
               Width           =   1455
            End
         End
      End
      Begin VB.TextBox PortText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   15
         Text            =   "7496"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox ClientIDText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox ServerText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton ConvertButton 
         Caption         =   "Convert"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   10
         ToolTipText     =   "Start tickfile conversion"
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton StopButton 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   11
         ToolTipText     =   "Stop tickfile conversion"
         Top             =   4320
         Width           =   735
      End
      Begin VB.ListBox TickfileList 
         Height          =   2400
         ItemData        =   "MainForm.frx":0060
         Left            =   120
         List            =   "MainForm.frx":0062
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   480
         Width           =   7575
      End
      Begin VB.CommandButton ClearTickfileListButton 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7800
         TabIndex        =   1
         ToolTipText     =   "Clear tickfile list"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton SelectTickfilesButton 
         Caption         =   "..."
         Height          =   375
         Left            =   7800
         TabIndex        =   0
         ToolTipText     =   "Select tickfile(s)"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton OrderButton 
         Caption         =   "&Order ticket"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -66720
         TabIndex        =   36
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -66720
         TabIndex        =   35
         Top             =   1620
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -66720
         TabIndex        =   34
         Top             =   1020
         Width           =   975
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69840
         TabIndex        =   33
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Left            =   -67680
         TabIndex        =   32
         ToolTipText     =   "Select tickfile(s)"
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67680
         TabIndex        =   31
         ToolTipText     =   "Clear tickfile list"
         Top             =   1500
         Width           =   495
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69120
         TabIndex        =   30
         ToolTipText     =   "Pause tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68400
         TabIndex        =   29
         ToolTipText     =   "Stop tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   -74640
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1020
         Width           =   6855
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "MainForm.frx":0064
         Left            =   -74040
         List            =   "MainForm.frx":0093
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Adjust tickfile replay speed"
         Top             =   2460
         Width           =   1575
      End
      Begin MSComctlLib.ListView OpenOrdersList 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   37
         ToolTipText     =   "Open orders"
         Top             =   420
         Width           =   8055
         _ExtentX        =   14208
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
         TabIndex        =   38
         ToolTipText     =   "Filled orders"
         Top             =   2580
         Width           =   8055
         _ExtentX        =   14208
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   -74640
         TabIndex        =   39
         Top             =   3180
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ProgressBar ReplayProgressBar 
         Height          =   135
         Left            =   120
         TabIndex        =   45
         Top             =   3360
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label18 
         Caption         =   "Output path"
         Height          =   255
         Left            =   7800
         TabIndex        =   70
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "QT Port"
         Height          =   255
         Left            =   8400
         TabIndex        =   68
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "QT Server"
         Height          =   255
         Left            =   8400
         TabIndex        =   67
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Current contract details"
         Height          =   255
         Left            =   -74640
         TabIndex        =   61
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "Right"
         Height          =   255
         Left            =   -70320
         TabIndex        =   58
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Strike price"
         Height          =   255
         Left            =   -70320
         TabIndex        =   57
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Symbol"
         Height          =   255
         Left            =   -70320
         TabIndex        =   56
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
         Height          =   255
         Left            =   -70320
         TabIndex        =   55
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Expiry"
         Height          =   255
         Left            =   -70320
         TabIndex        =   54
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Exchange"
         Height          =   255
         Left            =   -70320
         TabIndex        =   53
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Port"
         Height          =   255
         Left            =   -70320
         TabIndex        =   50
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Client id"
         Height          =   255
         Left            =   -70320
         TabIndex        =   49
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Server"
         Height          =   255
         Left            =   -70320
         TabIndex        =   48
         Top             =   600
         Width           =   615
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   3000
         Width           =   5655
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   855
         Left            =   120
         TabIndex        =   46
         Top             =   3720
         Width           =   5655
      End
      Begin VB.Label Label10 
         Caption         =   "Select tickfile(s)"
         Height          =   255
         Left            =   -74520
         TabIndex        =   43
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Output path"
         Height          =   855
         Left            =   -74640
         TabIndex        =   42
         Top             =   3420
         Width           =   5655
      End
      Begin VB.Label Label8 
         Caption         =   "qazqazqaz"
         Height          =   255
         Left            =   -74640
         TabIndex        =   41
         Top             =   2940
         Width           =   5655
      End
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74640
         TabIndex        =   40
         Top             =   2460
         Width           =   615
      End
   End
   Begin VB.Label Label16 
      Caption         =   "QT Port"
      Height          =   255
      Left            =   7920
      TabIndex        =   69
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "MainForm"
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

Implements TradeBuild.IListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1
Private WithEvents mTickfileManager As TradeBuild.TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1
Private WithEvents mContracts As TradeBuild.Contracts
Attribute mContracts.VB_VarHelpID = -1
Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private mOutputFormat As String
Private mOutputPath As String

Private mQuoteTrackerSP As QuoteTrackerSP.QTServiceProvider

Private mContract As Contract

Private mSupportedOutputFormats() As TradeBuild.TickfileFormatSpecifier

Private mArguments As cCommandLineArgs
Private mNoUI As Boolean
Private mRun As Boolean

Private mMonths(12) As String

Private mNoWriteBars As Boolean
Private mNoWriteTicks As Boolean

Private mTickfileSpecifiers() As TradeBuild.TickfileSpecifier

Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
On Error GoTo err
InitCommonControls

mMonths(1) = "Jan"
mMonths(2) = "Feb"
mMonths(3) = "Mar"
mMonths(4) = "Apr"
mMonths(5) = "May"
mMonths(6) = "Jun"
mMonths(7) = "Jul"
mMonths(8) = "Aug"
mMonths(9) = "Sep"
mMonths(10) = "Oct"
mMonths(11) = "Nov"
mMonths(12) = "Dec"

Exit Sub

err:
handleFatalError err.Number, _
                "Couldn't initialise common controls.", _
                "Form_Initialize"

End Sub

Private Sub Form_Load()
Dim TickfileSP As TickfileSP.TickfileServiceProvider
Dim SQLDBTickfileSP As TBInfoBase.TickfileServiceProvider
Dim contractInfoSP As TBInfoBase.ContractInfoServiceProvider
Dim histDataSP As TBInfoBase.HistDataServiceProvider
Dim i As Long

On Error Resume Next
Set mTradeBuildAPI = New TradeBuildAPI
On Error GoTo 0
If mTradeBuildAPI Is Nothing Then
    handleFatalError 999, _
                    "The required version of TradeBuild is not installed.", _
                    "Form_Load"
    Exit Sub
End If

mTradeBuildAPI.addListener Me, TradeBuild.StandardListenValueTypes.LogInfo

On Error Resume Next
Set SQLDBTickfileSP = New TBInfoBase.TickfileServiceProvider
On Error GoTo 0
If SQLDBTickfileSP Is Nothing Then
    handleFatalError 998, _
                    "The TradeBuild SQLDB Tickfile Service Provider is not installed.", _
                    "Form_Load"
    Exit Sub
End If
mTradeBuildAPI.ServiceProviders.Add SQLDBTickfileSP

On Error Resume Next
Set TickfileSP = New TickfileSP.TickfileServiceProvider
On Error GoTo 0
If TickfileSP Is Nothing Then
    handleFatalError 998, _
                    "The TradeBuild Tickfile Service Provider is not installed.", _
                    "Form_Load"
    Exit Sub
End If
mTradeBuildAPI.ServiceProviders.Add TickfileSP

On Error Resume Next
Set mQuoteTrackerSP = New QuoteTrackerSP.QTServiceProvider
On Error GoTo 0
If mQuoteTrackerSP Is Nothing Then
    handleFatalError 997, _
                    "The QuoteTracker Service Provider is not installed.", _
                    "Form_Load"
    Exit Sub
End If
mQuoteTrackerSP.ConnectionRetryIntervalSecs = 10
mQuoteTrackerSP.password = ""
mTradeBuildAPI.ServiceProviders.Add mQuoteTrackerSP

On Error Resume Next
Set contractInfoSP = New TBInfoBase.ContractInfoServiceProvider
On Error GoTo 0
If contractInfoSP Is Nothing Then
    handleFatalError 998, _
                    "The TradeBuild Contract Info Service Provider is not installed.", _
                    "Form_Load"
    Exit Sub
End If
mTradeBuildAPI.ServiceProviders.Add contractInfoSP

On Error Resume Next
Set histDataSP = New TBInfoBase.HistDataServiceProvider
On Error GoTo 0
If histDataSP Is Nothing Then
    handleFatalError 998, _
                    "The TradeBuild Historic Data Service Provider is not installed.", _
                    "Form_Load"
    Exit Sub
End If
mTradeBuildAPI.ServiceProviders.Add histDataSP

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

FormatList.AddItem "(None)"
For i = 0 To UBound(mSupportedOutputFormats)
    FormatList.AddItem mSupportedOutputFormats(i).Name
Next

FormatList.ListIndex = 0

TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.AddItem optionRightToString(OptionRights.OptCall)
RightCombo.AddItem optionRightToString(OptionRights.OptPut)

QTPortText.Text = "16240"

mOutputPath = App.Path

If Not ProcessCommandLineArgs Then
    Unload Me
End If
End Sub

'================================================================================
' IListener Interface Members
'================================================================================

Private Sub IListener_notify( _
                            ByVal valueType As Long, _
                            ByVal data As Variant, _
                            ByVal timestamp As Date)
Select Case valueType
Case TradeBuild.StandardListenValueTypes.LogInfo
    writeStatusMessage "Log: " & data
End Select
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub AdjustTimestampsCheck_Click()
If AdjustTimestampsCheck = vbChecked Then
    AdjustSecondsStartText.Enabled = True
    AdjustSecondsEndText.Enabled = True
Else
    AdjustSecondsStartText.Enabled = False
    AdjustSecondsEndText.Enabled = False
End If
End Sub

Private Sub ClearTickfileListButton_Click()
TickfileList.Clear
ClearTickfileListButton.Enabled = False
mTickfileManager.ClearTickfileSpecifiers
ConvertButton.Enabled = False
StopButton.Enabled = False
End Sub

Private Sub ContractFromServiceProviderOption_Click()
enableContractFields
SymbolText.SetFocus
End Sub

Private Sub ContractInTickfileOption_Click()
disableContractFields
End Sub

Private Sub ConvertButton_Click()
If ContractFromServiceProviderOption Then
    If mContract Is Nothing Then
        writeStatusMessage "Can't convert - no contract details are available"
        Exit Sub
    Else
        mTickfileManager.Contract = mContract
    End If
End If

SelectTickfilesButton.Enabled = False
ClearTickfileListButton.Enabled = False
ConvertButton.Enabled = False
StopButton.Enabled = True
ReplayProgressBar.Visible = True

mTickfileManager.replaySpeed = 0
If AdjustTimestampsCheck = vbChecked Then
    mTickfileManager.TimestampAdjustmentStart = AdjustSecondsStartText
    mTickfileManager.TimestampAdjustmentEnd = AdjustSecondsEndText
End If

mQuoteTrackerSP.QTServer = QTServerText.Text
mQuoteTrackerSP.QTPort = QTPortText.Text
QTServerText.Enabled = False
QTPortText.Enabled = False

writeStatusMessage "Tickfile conversion started"
mTickfileManager.StartReplay
End Sub

Private Sub DisconnectButton_Click()
mTradeBuildAPI.disconnect
GetContractButton.Enabled = True
enableContractFields
DisconnectButton.Enabled = False
End Sub

Private Sub ExchangeText_Change()
checkOKToGetContract
End Sub

Private Sub ExpiryText_Change()
checkOKToGetContract
End Sub

Private Sub FormatList_Click()
Dim i As Long
mOutputFormat = ""
For i = 0 To UBound(mSupportedOutputFormats)
    If FormatList.Text = mSupportedOutputFormats(i).Name Then
        mOutputFormat = mSupportedOutputFormats(i).FormalID
        Exit Sub
    End If
Next
End Sub

Private Sub GetContractButton_Click()
Dim lContractSpecifier As ContractSpecifier

On Error GoTo err

Set lContractSpecifier = mTradeBuildAPI.newContractSpecifier( _
                                    , _
                                    SymbolText, _
                                    ExchangeText, _
                                    secTypeFromString(TypeCombo), _
                                    , _
                                    ExpiryText, _
                                    IIf(StrikePriceText = "", 0, StrikePriceText), _
                                    optionRightFromString(RightCombo))

Set mContracts = mTradeBuildAPI.RequestContract(lContractSpecifier)
writeStatusMessage "Requesting contract details"
Exit Sub

err:
handleFatalError err.Number, err.description, "GetContractButton_Click"
End Sub

Private Sub OutputPathButton_Click()
Dim pathChooser As AppFramework.CPathChooser
Set pathChooser = New AppFramework.CPathChooser
pathChooser.Choose
OutputPathText.Text = pathChooser.Path
End Sub

Private Sub OutputPathText_Change()
mOutputPath = OutputPathText.Text
End Sub

Private Sub SelectTickfilesButton_Click()
Set mTickfileManager = mTradeBuildAPI.Tickers.createTickFileManager

mTickfileManager.ShowTickfileSelectionDialogue
QTServerText.Enabled = True
QTPortText.Enabled = True

End Sub

Private Sub StopButton_Click()
ConvertButton.Enabled = True
StopButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
mTicker.StopTicker
End Sub

Private Sub SymbolText_Change()
checkOKToGetContract
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
Case SecurityTypes.SecTypeBag
    writeStatusMessage "BAG type is not implemented"
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
End Select

checkOKToGetContract
End Sub

'================================================================================
' mContracts Event Handlers
'================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal ContractSpecifier As TradeBuild.ContractSpecifier)
writeStatusMessage "Invalid contract specifier: " & _
                    Replace(ContractSpecifier.ToString, vbCrLf, "; ")
End Sub

Private Sub mContracts_NoMoreContractDetails()
On Error GoTo err

If mContracts.Count > 1 Then
    writeStatusMessage "Unique contract not specified"
    Exit Sub
End If
Set mContract = mContracts(1)
ContractDetailsText = mContract.ToString

writeStatusMessage "Contract details received"
enableContractFields

Exit Sub
err:
handleFatalError err.Number, err.description, "mTradeBuildAPI_Contract"
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_errorMessage(ByVal timestamp As Date, _
                                ByVal id As Long, _
                                ByVal errorCode As TradeBuild.ApiErrorCodes, _
                                ByVal errorMsg As String)
On Error GoTo err
Select Case errorCode
Case ApiErrorCodes.NoContractDetails
    writeStatusMessage "Error " & errorCode & ": " & errorMsg
    
Case Else
    writeStatusMessage "Error " & errorCode & ": " & id & ": " & errorMsg
End Select

Exit Sub
err:
handleFatalError err.Number, err.description, "mTicker_errorMessage"
End Sub

Private Sub mTicker_outputTickfileCreated( _
                            ByVal timestamp As Date, _
                            ByVal filename As String)
writeStatusMessage "Created output tickfile: " & filename
End Sub

'================================================================================
' mTickfileManager Event Handlers
'================================================================================

Private Sub mTickfileManager_errorMessage(ByVal timestamp As Date, ByVal id As Long, ByVal errorCode As TradeBuild.ApiErrorCodes, ByVal errorMsg As String)
On Error GoTo err
Select Case errorCode
Case ApiErrorCodes.NoContractDetails
    writeStatusMessage "Error " & errorCode & ": " & errorMsg
    
Case Else
    writeStatusMessage "Error " & errorCode & ": " & id & ": " & errorMsg
End Select

Exit Sub
err:
handleFatalError err.Number, err.description, "mTickfileManager_errorMessage"
End Sub

Private Sub mTickfileManager_QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal tickfileSizeBytes As Long, _
                ByVal pContract As TradeBuild.Contract, _
                continueMode As TradeBuild.ReplayContinueModes)
On Error GoTo err

ReplayProgressBar.Min = 0
ReplayProgressBar.Max = 100
ReplayProgressBar.value = 0
TickfileList.ListIndex = tickfileIndex
writeStatusMessage "Converting " & TickfileList.List(TickfileList.ListIndex)
ReplayContractLabel.Caption = "Symbol:   " & pContract.specifier.Symbol & vbCrLf & _
                            "Type:     " & secTypeToString(pContract.specifier.SecType) & vbCrLf & _
                            IIf(pContract.specifier.SecType <> SecurityTypes.SecTypeStock, "Expiry:   " & pContract.specifier.Expiry & vbCrLf, "") & _
                            "Exchange: " & pContract.specifier.exchange

Exit Sub
err:
handleFatalError err.Number, err.description, "mTickfileManager_QueryReplayNextTickfile"
End Sub

Private Sub mTickfileManager_ReplayCompleted()
On Error GoTo err

ConvertButton.Enabled = True
StopButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
ReplayProgressBar.value = 0
ReplayProgressBar.Visible = False
ReplayContractLabel.Caption = ""
ReplayProgressLabel.Caption = ""

writeStatusMessage "Tickfile conversion completed"

If mRun Then Unload Me

Exit Sub
err:
handleFatalError err.Number, err.description, "mTickfileManager_ReplayCompleted"

End Sub

Private Sub mTickfileManager_ReplayProgress( _
                            ByVal tickfileTimestamp As Date, _
                            ByVal eventsPlayed As Long, _
                            ByVal percentComplete As Single)

On Error GoTo err
ReplayProgressBar.value = percentComplete
ReplayProgressBar.Refresh
ReplayProgressLabel.Caption = tickfileTimestamp & _
                                "  Processed " & _
                                eventsPlayed & _
                                " events"

Exit Sub
err:
handleFatalError err.Number, err.description, "mTickfileManager_ReplayProgress"
End Sub

Private Sub mTickfileManager_TickerAllocated(ByVal pTicker As TradeBuild.Ticker)
Dim i As Long
On Error GoTo err
Set mTicker = pTicker
mTicker.outputTickfilePath = mOutputPath
mTicker.outputTickfileFormat = mOutputFormat
If mOutputFormat <> "" Then
    mTicker.writeToTickfile = True
    mTicker.includeMarketDepthInTickfile = True
End If

For i = 0 To EnableCheck.UBound
    If EnableCheck(i).value = vbChecked Then
        mTicker.Timeframes.Add BarLengthText(i).Text, _
                                TradeBuild.TimePeriodUnits.Minute, _
                                BarLengthText(i).Text & "min", _
                                0, _
                                0, _
                                (IncludeBidAndAskCheck(i).value = vbChecked)
    End If
Next

Exit Sub
err:
handleFatalError err.Number, err.description, "mTickfileManager_TickerAllocated"
End Sub

Private Sub mTickfileManager_TickfilesSelected()
On Error GoTo err
Dim tickfiles() As TradeBuild.TickfileSpecifier
Dim i As Long
TickfileList.Clear
tickfiles = mTickfileManager.TickfileSpecifiers
For i = 0 To UBound(tickfiles)
    TickfileList.AddItem tickfiles(i).filename
Next
ConvertButton.Enabled = True
ClearTickfileListButton.Enabled = True

Exit Sub
err:
handleFatalError err.Number, err.description, "mTickfileManager_TickfilesSelected"
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Set mTickfileManager = mTradeBuildAPI.Tickers.createTickFileManager

mQuoteTrackerSP.QTServer = QTServerText.Text
mQuoteTrackerSP.QTPort = QTPortText.Text

QTServerText.Enabled = False
QTPortText.Enabled = False

mTickfileManager.TickfileSpecifiers = mTickfileSpecifiers
mTickfileManager.replaySpeed = 0

mTickfileManager.StartReplay
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_connecting(ByVal timestamp As Date)
On Error GoTo err
writeStatusMessage "Connecting"
DisconnectButton.Enabled = True

Exit Sub
err:
handleFatalError err.Number, err.description, "mTradeBuildAPI_connecting"
End Sub

Private Sub mTradeBuildAPI_connectionToTWSClosed( _
                ByVal timestamp As Date, _
                ByVal reconnecting As Boolean)
On Error GoTo err

DisconnectButton.Enabled = False
enableContractFields

writeStatusMessage "Connection closed"

Exit Sub
err:
handleFatalError err.Number, err.description, "mTradeBuildAPI_connectionClosed"
End Sub

Private Sub mTradeBuildAPI_errorMessage(ByVal timestamp As Date, _
                        ByVal id As Long, _
                        ByVal errorCode As ApiErrorCodes, _
                        ByVal errorMsg As String)
Dim spError As ServiceProviderError

On Error GoTo err

Select Case errorCode
Case ApiErrorCodes.ServiceProviderErrorNotification
    Set spError = mTradeBuildAPI.getServiceProviderError
    writeStatusMessage "Error from " & _
                        spError.serviceProviderName & _
                        ": code " & spError.errorCode & _
                        ": id " & id & ": " & _
                        spError.message

Case Else
    writeStatusMessage "Error " & errorCode & ": " & id & ": " & errorMsg
End Select


Exit Sub
err:
handleFatalError err.Number, err.description, "mTradeBuildAPI_errorMessage"
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



Private Sub checkOKToGetContract()
If ContractFromServiceProviderOption Then
    If SymbolText <> "" Then
'    If PortText <> "" And ClientIDText <> "" And _
'        SymbolText <> "" And _
'        TypeCombo.Text <> "" And _
'        IIf(TypeCombo.Text = StrSecTypeFuture Or _
'            TypeCombo.Text = StrSecTypeOption Or _
'            TypeCombo.Text = StrSecTypeOptionFuture, _
'            ExpiryText <> "", _
'            True) And _
'        IIf(TypeCombo.Text = StrSecTypeOption Or _
'            TypeCombo.Text = StrSecTypeOptionFuture, _
'            StrikePriceText <> "", _
'            True) And _
'        IIf(TypeCombo.Text = StrSecTypeOption Or _
'            TypeCombo.Text = StrSecTypeOptionFuture, _
'            RightCombo <> "", _
'            True) And _
'        ExchangeText <> "" _
'    Then
        GetContractButton.Enabled = True
    Else
        GetContractButton.Enabled = False
    End If
Else
    If SymbolText <> "" Then
        GetContractButton.Enabled = True
    Else
        GetContractButton.Enabled = False
    End If
End If
End Sub

Private Sub disableContractFields()
GetContractButton.Enabled = False
ConvertButton.Enabled = False
ServerText.Enabled = False
PortText.Enabled = False
ClientIDText.Enabled = False
SymbolText.Enabled = False
TypeCombo.Enabled = False
ExpiryText.Enabled = False
ExchangeText.Enabled = False
StrikePriceText.Enabled = False
RightCombo.Enabled = False
End Sub

Private Sub enableContractFields()
'If mTradeBuildAPI.connectionState = ConnNotConnected Then
'    ServerText.Enabled = True
'    PortText.Enabled = True
'    ClientIDText.Enabled = True
'End If
SymbolText.Enabled = True
TypeCombo.Enabled = True
ExchangeText.Enabled = True
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
Case SecurityTypes.SecTypeBag
    writeStatusMessage "BAG type is not implemented"
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case Else
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = True
    RightCombo.Enabled = True
End Select
checkOKToGetContract
End Sub

Private Sub handleFatalError(ByVal errNum As Long, _
                            ByVal description As String, _
                            ByVal source As String)

Dim i As Long

Set mTicker = Nothing
Set mTradeBuildAPI = Nothing

MsgBox "A fatal error has occurred. The program will close" & vbCrLf & _
        "Error number: " & errNum & vbCrLf & _
        "Description: " & description & vbCrLf & _
        "Source: TickFielManager.MainForm::" & source, _
        vbCritical, _
        "Fatal error"

For i = Forms.Count - 1 To 0 Step -1
   Unload Forms(i)
Next

End Sub

Private Function ProcessCommandLineArgs() As Boolean
Dim symbolValue As String
Dim localSymbolValue As String
Dim secTypeValue As String
Dim monthValue As String
Dim exchangeValue As String
Dim currencyValue As String
Dim strikevalue As String
Dim rightValue As String
Dim fromValue As String
Dim fromDate As Date
Dim fromTime As Date
Dim toValue As String
Dim toDate As Date
Dim toTime As Date
Dim sessionsValue As String
Dim numberOfSessions As Long
Dim startingSession As Long
Dim inFormatValue As String
Dim outFormatValue As String
Dim QTServerValue As String
Dim commaPosn As Long
Dim contractSpec As TradeBuild.ContractSpecifier
Dim i As Long
Dim j As Long
Dim lSupportedInputTickfileFormats() As TradeBuild.TickfileFormatSpecifier

Set mArguments = New cCommandLineArgs
mArguments.CommandLine = Command
mArguments.Separator = " "
mArguments.GetArgs

If mArguments.Switch("?") Then
    MsgBox vbCrLf & _
            "tickfilemanager [symbol  localSymbol|NOLOCALSYMBOL sectype " & vbCrLf & _
            "                month|NOMONTH exchange currency [strike] [right]]" & vbCrLf & _
            "                [/from:yyyymmdd[hhmmss]] " & vbCrLf & _
            "                [/to:yyyymmdd[hhmmss]] " & vbCrLf & _
            "                [/sessions:n[,m]]" & vbCrLf & _
            "                [/inFormat:inputTickfileFormat" & vbCrLf & _
            "                [/putFormat:outputTickfileFormat" & vbCrLf & _
            "                [/outpath:path]" & vbCrLf & _
            "                [/noWriteBars  |  /nwb]" & vbCrLf & _
            "                [/noUI]  [/run]" & vbCrLf & _
            "                [/QTserver:[server][,port]]" & vbCrLf & _
            vbCrLf & _
            "Notes:" & vbCrLf & _
            "   If /from is supplied, /sessions is ignored." & vbCrLf & _
            "   If /from is not supplied, /to is ignored." & vbCrLf & _
            "   In /sessions, n is the number of sessions to supply, and m" & vbCrLf & _
            "      is the number of sessions before current to start at." & vbCrLf & _
            "      m defaults to 1. If m is zero, the current session is" & vbCrLf & _
            "      supplied." & vbCrLf & _
            "   In /QTserver, port defaults to 16240.", _
            , _
            "Usage"
    ProcessCommandLineArgs = False
    Exit Function
End If

If mArguments.Switch("noui") Then
    mNoUI = True
End If

If mArguments.Switch("run") Then
    mRun = True
End If

symbolValue = mArguments.Arg(0)
If symbolValue = "" Then
    If mNoUI Then
        ProcessCommandLineArgs = False
        Exit Function
    ElseIf mRun Then
        MsgBox "Error - no symbol argument supplied"
        ProcessCommandLineArgs = False
        Exit Function
    End If
End If

localSymbolValue = mArguments.Arg(1)
If UCase(localSymbolValue) = "NOLOCALSYMBOL" Then localSymbolValue = ""

secTypeValue = mArguments.Arg(2)

monthValue = mArguments.Arg(3)
If UCase$(monthValue) = "NOMONTH" Then monthValue = ""

exchangeValue = mArguments.Arg(4)

currencyValue = mArguments.Arg(5)

strikevalue = mArguments.Arg(6)

rightValue = mArguments.Arg(7)

If mArguments.Switch("from") Then
    fromValue = mArguments.SwitchValue("from")
    If IsNumeric(fromValue) And _
        (Len(fromValue) = 8 _
            Or _
        Len(fromValue) = 14) _
    Then
        On Error Resume Next
        unpackDateTimeString fromValue, fromDate, fromTime
        If err.Number <> 0 Then
            MsgBox fromValue & " is not a valid date and time (format yyyymmdd[hhmmss])"
            ProcessCommandLineArgs = False
            Exit Function
        End If
        On Error GoTo 0
    Else
        If mNoUI Then
            ProcessCommandLineArgs = False
            Exit Function
        ElseIf mRun Then
            MsgBox "Error - from  " & fromValue & " not in format yyyymmdd[hhmmss]"
            ProcessCommandLineArgs = False
            Exit Function
        End If
    End If
End If

If mArguments.Switch("from") And mArguments.Switch("to") Then
    toValue = mArguments.SwitchValue("to")
    If IsNumeric(toValue) And _
        (Len(toValue) = 8 _
            Or _
        Len(toValue) = 14) _
    Then
        On Error Resume Next
        unpackDateTimeString toValue, toDate, toTime
        If err.Number <> 0 Then
            MsgBox toValue & " is not a valid date and time (format yyyymmdd[hhmmss])"
            ProcessCommandLineArgs = False
            Exit Function
        End If
        On Error GoTo 0
    Else
        If mNoUI Then
            ProcessCommandLineArgs = False
            Exit Function
        ElseIf mRun Then
            MsgBox "Error - to  " & toValue & " not in format yyyymmdd[hhmmss]"
            ProcessCommandLineArgs = False
            Exit Function
        End If
    End If
End If

startingSession = 1
If mArguments.Switch("sessions") Then
    sessionsValue = mArguments.SwitchValue("sessions")
    If Len(sessionsValue) = 0 Then
        MsgBox "Error - sessions should be /sessions:n[,m]"
        ProcessCommandLineArgs = False
        Exit Function
    End If
    
    On Error Resume Next
    If InStr(1, sessionsValue, ",") Then
        numberOfSessions = CLng(Left$(sessionsValue, InStr(1, sessionsValue, ",") - 1))
        If err.Number <> 0 Or numberOfSessions < 1 Then
            MsgBox "Error - sessions should be /sessions:n[,m] where n and m are integers, n>=1 and m>=0"
            ProcessCommandLineArgs = False
            Exit Function
        End If
        startingSession = CLng(Right$(sessionsValue, Len(sessionsValue) - InStr(1, sessionsValue, ",")))
        If err.Number <> 0 Or startingSession < 0 Then
            MsgBox "Error - sessions should be /sessions:n[,m] where n and m are integers, n>=1 and m>=0"
            ProcessCommandLineArgs = False
            Exit Function
        End If
    Else
        numberOfSessions = sessionsValue
        If err.Number <> 0 Or numberOfSessions < 1 Then
            MsgBox "Error - sessions should be /sessions:n[,m] where n and m are integers, n>=1 and m>=0"
            ProcessCommandLineArgs = False
            Exit Function
        End If
    End If
End If

If mArguments.Switch("outpath") Then
    If mArguments.SwitchValue("outpath") <> "" Then
        OutputPathText.Text = mArguments.SwitchValue("outpath")
    End If
End If

If mArguments.Switch("qtserver") Then
    QTServerValue = mArguments.SwitchValue("qtserver")
    commaPosn = InStr(1, QTServerValue, ",")
    Select Case commaPosn
    Case 0
        QTServerText.Text = QTServerValue
    Case 1
        If IsNumeric(QTServerValue) Then
            QTPortText.Text = QTServerValue
        Else
            MsgBox "Error - qtserver should be /qtserver:[server[,port] where server is a computer name or address, and port is an integer (port >0)"
            ProcessCommandLineArgs = False
            Exit Function
        End If
    Case Else
        QTServerText.Text = Left$(QTServerValue, commaPosn - 1)
        If IsNumeric(Right$(QTServerValue, Len(QTServerValue) - commaPosn)) Then
            QTPortText.Text = Right$(QTServerValue, Len(QTServerValue) - commaPosn)
        Else
            MsgBox "Error - qtserver should be /qtserver:[server[,port] where server is a computer name or address, and port is an integer (port >0)"
            ProcessCommandLineArgs = False
            Exit Function
        End If
    End Select
        
End If

If mArguments.Switch("informat") Then inFormatValue = mArguments.SwitchValue("informat")

If mArguments.Switch("outformat") Then
    outFormatValue = mArguments.SwitchValue("outformat")
    FormatList.Text = outFormatValue
End If

If mArguments.Switch("nwb") Or _
    mArguments.Switch("nowritebars") _
Then
    mNoWriteBars = True
End If

If mArguments.Switch("nwt") Or _
    mArguments.Switch("nowriteticks") _
Then
    mNoWriteTicks = True
End If

If symbolValue <> "" Then
    Set contractSpec = mTradeBuildAPI.newContractSpecifier( _
                                localSymbolValue, _
                                symbolValue, _
                                exchangeValue, _
                                secTypeFromString(secTypeValue), _
                                currencyValue, _
                                monthValue, _
                                IIf(StrikePriceText.Text = "", 0, StrikePriceText.Text), _
                                optionRightFromString(rightValue))


    If mArguments.Switch("from") Then numberOfSessions = 1
    If numberOfSessions > (startingSession + 1) Then numberOfSessions = startingSession + 1
    
    ReDim mTickfileSpecifiers(numberOfSessions - 1) As TradeBuild.TickfileSpecifier
    For i = 0 To UBound(mTickfileSpecifiers)
        With mTickfileSpecifiers(i)
            Set .ContractSpecifier = contractSpec
            If mArguments.Switch("from") Then
                .From = fromDate + fromTime
                If mArguments.Switch("to") Then
                    .To = toDate + toTime
                Else
                    .To = DateAdd("n", 1, Now)
                End If
            Else
                .EntireSession = True
                .From = DateAdd("d", -startingSession + i, Now)
            End If
                
            lSupportedInputTickfileFormats = mTradeBuildAPI.SupportedInputTickfileFormats
            For j = 0 To UBound(lSupportedInputTickfileFormats)
                If lSupportedInputTickfileFormats(j).Name = inFormatValue Then
                    .TickfileFormatID = lSupportedInputTickfileFormats(j).FormalID
                    Exit For
                End If
            Next
            
            If .EntireSession Then
                .filename = "Session (" & .From & ") " & _
                                Replace(contractSpec.ToString, vbCrLf, "; ")
            Else
                .filename = .From & "-" & .To & " " & _
                                Replace(contractSpec.ToString, vbCrLf, "; ")
            End If
        End With
    Next
    
    For i = 0 To UBound(mTickfileSpecifiers)
        TickfileList.AddItem mTickfileSpecifiers(i).filename
    Next
    
    Set mTimer = New TimerUtils.IntervalTimer
    mTimer.RepeatNotifications = False
    mTimer.TimerIntervalMillisecs = 10
    mTimer.StartTimer
End If

ProcessCommandLineArgs = True
End Function

Private Sub unpackDateTimeString( _
                            ByVal timestampString As String, _
                            ByRef dateOut As Date, _
                            ByRef timeOut As Date)
dateOut = CDate(mMonths(Mid$(timestampString, 5, 2)) & " " & _
                            Mid$(timestampString, 7, 2) & " " & _
                            Left$(timestampString, 4))
If Len(timestampString) = 14 Then
    timeOut = CDate(Mid$(timestampString, 9, 2) & ":" & _
                            Mid$(timestampString, 11, 2) & ":" & _
                            Mid(timestampString, 13, 2))
End If
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



