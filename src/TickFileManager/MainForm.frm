VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{948AEB4D-03C6-4FAB-ACD2-E61F7B7A0EB3}#1.0#0"; "TradeBuildUI27.ocx"
Begin VB.Form MainForm 
   Caption         =   "TradeBuild Tickfile Manager Version 2.7"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11925
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
      Top             =   6240
      Width           =   11655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Configuration"
      TabPicture(0)   =   "MainForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ConfigureButton"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tickfile selection"
      TabPicture(1)   =   "MainForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "SelectTickfilesButton"
      Tab(1).Control(2)=   "ClearTickfileListButton"
      Tab(1).Control(3)=   "TickfileList"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "StopButton"
      Tab(1).Control(5)=   "ConvertButton"
      Tab(1).Control(6)=   "ReplayProgressBar"
      Tab(1).Control(7)=   "ReplayContractLabel"
      Tab(1).Control(8)=   "ReplayProgressLabel"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Contract details"
      TabPicture(2)   =   "MainForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ContractSpecBuilder1"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "GetContractButton"
      Tab(2).Control(3)=   "ContractDetailsText"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label11"
      Tab(2).ControlCount=   5
      Begin TradeBuildUI27.ContractSpecBuilder ContractSpecBuilder1 
         Height          =   2895
         Left            =   -69840
         TabIndex        =   61
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   6509
      End
      Begin VB.CommandButton ConfigureButton 
         Caption         =   "&Configure"
         Height          =   495
         Left            =   10080
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.Frame Frame6 
         Caption         =   "Output configuration"
         Height          =   5295
         Left            =   3840
         TabIndex        =   11
         Top             =   480
         Width           =   3495
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   4935
            Left            =   120
            ScaleHeight     =   4935
            ScaleWidth      =   3255
            TabIndex        =   81
            Top             =   240
            Width           =   3255
            Begin VB.CheckBox AsyncWritesCheck 
               Caption         =   "Use asynchronous writes"
               Height          =   195
               Left            =   360
               TabIndex        =   97
               Top             =   4680
               Width           =   2655
            End
            Begin VB.TextBox ContractPasswordText 
               Enabled         =   0   'False
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1200
               PasswordChar    =   "*"
               TabIndex        =   17
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   2160
               Width           =   1815
            End
            Begin VB.TextBox ContractUsernameText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   16
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   1800
               Width           =   1815
            End
            Begin VB.ComboBox ContractDbTypeCombo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               TabIndex        =   14
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox ContractServerText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   13
               ToolTipText     =   "Name or address of computer hosting QuoteTracker"
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox ContractDatabaseText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   15
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   1440
               Width           =   1815
            End
            Begin VB.OptionButton FileOutputOption 
               Caption         =   "Output to file"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   0
               Value           =   -1  'True
               Width           =   2295
            End
            Begin VB.OptionButton DatabaseOutputOption 
               Caption         =   "Output to this database:"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   2520
               Width           =   2295
            End
            Begin VB.TextBox DatabaseOutText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   21
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   3600
               Width           =   1815
            End
            Begin VB.TextBox DbOutServerText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   19
               ToolTipText     =   "Name or address of computer hosting QuoteTracker"
               Top             =   2880
               Width           =   1815
            End
            Begin VB.ComboBox DbOutTypeCombo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               TabIndex        =   20
               Top             =   3240
               Width           =   1815
            End
            Begin VB.TextBox UsernameOutText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   22
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   3960
               Width           =   1815
            End
            Begin VB.TextBox PasswordOutText 
               Enabled         =   0   'False
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1200
               PasswordChar    =   "*"
               TabIndex        =   23
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   4320
               Width           =   1815
            End
            Begin VB.Label Label7 
               Caption         =   "Contract details from this database:"
               Height          =   255
               Left            =   360
               TabIndex        =   96
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label6 
               Caption         =   "Password"
               Height          =   255
               Left            =   360
               TabIndex        =   95
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "Username"
               Height          =   255
               Left            =   360
               TabIndex        =   94
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label4 
               Caption         =   "DB Type"
               Height          =   255
               Left            =   360
               TabIndex        =   93
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Server"
               Height          =   255
               Left            =   360
               TabIndex        =   92
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Database"
               Height          =   255
               Left            =   360
               TabIndex        =   91
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label29 
               Caption         =   "Database"
               Height          =   255
               Left            =   360
               TabIndex        =   86
               Top             =   3600
               Width           =   975
            End
            Begin VB.Label Label28 
               Caption         =   "Server"
               Height          =   255
               Left            =   360
               TabIndex        =   85
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label Label27 
               Caption         =   "DB Type"
               Height          =   255
               Left            =   360
               TabIndex        =   84
               Top             =   3240
               Width           =   975
            End
            Begin VB.Label Label26 
               Caption         =   "Username"
               Height          =   255
               Left            =   360
               TabIndex        =   83
               Top             =   3960
               Width           =   975
            End
            Begin VB.Label Label25 
               Caption         =   "Password"
               Height          =   255
               Left            =   360
               TabIndex        =   82
               Top             =   4320
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Input configuration"
         Height          =   5295
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3375
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   4935
            Left            =   120
            ScaleHeight     =   4935
            ScaleWidth      =   3135
            TabIndex        =   73
            Top             =   240
            Width           =   3135
            Begin VB.TextBox PasswordInText 
               Enabled         =   0   'False
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1200
               PasswordChar    =   "*"
               TabIndex        =   7
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   2160
               Width           =   1815
            End
            Begin VB.TextBox UsernameInText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   6
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   1800
               Width           =   1815
            End
            Begin VB.ComboBox DbInTypeCombo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               TabIndex        =   4
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox DbInServerText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   3
               ToolTipText     =   "Name or address of computer hosting QuoteTracker"
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox DatabaseInText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   5
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox QTServerText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   9
               ToolTipText     =   "Name or address of computer hosting QuoteTracker"
               Top             =   3000
               Width           =   1815
            End
            Begin VB.TextBox QTPortText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   10
               ToolTipText     =   "Port for connecting to QuoteTracker"
               Top             =   3360
               Width           =   1815
            End
            Begin VB.OptionButton QtInputOption 
               Caption         =   "Input from QuoteTracker:"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   2640
               Width           =   2295
            End
            Begin VB.OptionButton DatabaseInputOption 
               Caption         =   "Input from this database:"
               Height          =   255
               Left            =   120
               TabIndex        =   2
               Top             =   360
               Width           =   2295
            End
            Begin VB.OptionButton FileInputOption 
               Caption         =   "Input from file"
               Height          =   255
               Left            =   120
               TabIndex        =   1
               Top             =   0
               Value           =   -1  'True
               Width           =   2295
            End
            Begin VB.Label Label24 
               Caption         =   "Password"
               Height          =   255
               Left            =   360
               TabIndex        =   80
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label Label23 
               Caption         =   "Username"
               Height          =   255
               Left            =   360
               TabIndex        =   79
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label22 
               Caption         =   "DB Type"
               Height          =   255
               Left            =   360
               TabIndex        =   78
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label19 
               Caption         =   "Server"
               Height          =   255
               Left            =   360
               TabIndex        =   77
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label13 
               Caption         =   "Database"
               Height          =   255
               Left            =   360
               TabIndex        =   76
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label14 
               Caption         =   "QT Server"
               Height          =   255
               Left            =   360
               TabIndex        =   75
               Top             =   3000
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "QT Port"
               Height          =   255
               Left            =   360
               TabIndex        =   74
               Top             =   3360
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Conversion options"
         Height          =   4335
         Left            =   -66480
         TabIndex        =   67
         Top             =   480
         Width           =   2895
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   3975
            Left            =   120
            ScaleHeight     =   3975
            ScaleWidth      =   2715
            TabIndex        =   68
            Top             =   240
            Width           =   2715
            Begin VB.TextBox OutputPathText 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   33
               ToolTipText     =   "Location of output tickfiles"
               Top             =   2040
               Width           =   2535
            End
            Begin VB.CommandButton OutputPathButton 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   375
               Left            =   2160
               TabIndex        =   32
               ToolTipText     =   "Select output path"
               Top             =   1680
               Width           =   495
            End
            Begin VB.ListBox FormatList 
               Enabled         =   0   'False
               Height          =   645
               ItemData        =   "MainForm.frx":0054
               Left            =   120
               List            =   "MainForm.frx":0056
               TabIndex        =   31
               ToolTipText     =   "Select output tickfile format"
               Top             =   960
               Width           =   2535
            End
            Begin VB.Frame Frame3 
               Caption         =   "Timestamps"
               Height          =   1335
               Left            =   120
               TabIndex        =   69
               Top             =   2520
               Width           =   2535
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   975
                  Left            =   120
                  ScaleHeight     =   975
                  ScaleWidth      =   2295
                  TabIndex        =   70
                  Top             =   240
                  Width           =   2295
                  Begin VB.TextBox AdjustSecondsEndText 
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   1680
                     TabIndex        =   36
                     Text            =   "0"
                     ToolTipText     =   "Timestamp adjustment (seconds) at end of file"
                     Top             =   645
                     Width           =   495
                  End
                  Begin VB.TextBox AdjustSecondsStartText 
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   1680
                     TabIndex        =   35
                     Text            =   "0"
                     ToolTipText     =   "Timestamp adjustment (seconds) at start of file"
                     Top             =   360
                     Width           =   495
                  End
                  Begin VB.CheckBox AdjustTimestampsCheck 
                     Caption         =   "Adjust timestamps?"
                     Height          =   375
                     Left            =   120
                     TabIndex        =   34
                     ToolTipText     =   "Set if timestamps are to be adjusted"
                     Top             =   0
                     Width           =   1695
                  End
                  Begin VB.Label Label12 
                     Caption         =   "Seconds at end"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   72
                     Top             =   645
                     Width           =   1455
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Seconds at start"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   71
                     Top             =   360
                     Width           =   1455
                  End
               End
            End
            Begin VB.CheckBox WriteTickDataCheck 
               Caption         =   "Write tick data"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   0
               Width           =   1575
            End
            Begin VB.CheckBox WriteBarDataCheck 
               Caption         =   "Write bar data"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label30 
               Caption         =   "Output format"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label18 
               Caption         =   "Output path"
               Height          =   255
               Left            =   120
               TabIndex        =   87
               Top             =   1800
               Width           =   975
            End
         End
      End
      Begin VB.CommandButton SelectTickfilesButton 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67200
         TabIndex        =   27
         ToolTipText     =   "Select tickfile(s)"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton ClearTickfileListButton 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67200
         TabIndex        =   28
         ToolTipText     =   "Clear tickfile list"
         Top             =   1080
         Width           =   495
      End
      Begin VB.ListBox TickfileList 
         Height          =   4155
         ItemData        =   "MainForm.frx":0058
         Left            =   -74880
         List            =   "MainForm.frx":005A
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   480
         Width           =   7575
      End
      Begin VB.CommandButton StopButton 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -64800
         TabIndex        =   37
         ToolTipText     =   "Stop tickfile conversion"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton ConvertButton 
         Caption         =   "Convert"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -64800
         TabIndex        =   38
         ToolTipText     =   "Start tickfile conversion"
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contract details source"
         Height          =   1095
         Left            =   -74520
         TabIndex        =   58
         Top             =   840
         Width           =   2535
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   2295
            TabIndex        =   62
            Top             =   240
            Width           =   2295
            Begin VB.OptionButton ContractInTickfileOption 
               Caption         =   "In tickfile"
               Height          =   195
               Left            =   120
               TabIndex        =   59
               ToolTipText     =   "Tickfile contains contract details"
               Top             =   120
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton ContractFromServiceProviderOption 
               Caption         =   "Service provider"
               Height          =   195
               Left            =   120
               TabIndex        =   60
               ToolTipText     =   "Get contract details from service provider"
               Top             =   480
               Width           =   1455
            End
         End
      End
      Begin VB.CommandButton GetContractButton 
         Caption         =   "Get contract details"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -68760
         TabIndex        =   63
         ToolTipText     =   "Get contract details from specified source"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox ContractDetailsText 
         Height          =   2535
         Left            =   -74520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2400
         Width           =   3855
      End
      Begin VB.CommandButton OrderButton 
         Caption         =   "&Order ticket"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -66720
         TabIndex        =   48
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton CancelOrderButton 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -66720
         TabIndex        =   47
         Top             =   1620
         Width           =   975
      End
      Begin VB.CommandButton ModifyOrderButton 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -66720
         TabIndex        =   46
         Top             =   1020
         Width           =   975
      End
      Begin VB.CommandButton PlayTickFileButton 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69840
         TabIndex        =   45
         ToolTipText     =   "Start or resume tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Left            =   -67680
         TabIndex        =   44
         ToolTipText     =   "Select tickfile(s)"
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67680
         TabIndex        =   43
         ToolTipText     =   "Clear tickfile list"
         Top             =   1500
         Width           =   495
      End
      Begin VB.CommandButton PauseReplayButton 
         Caption         =   "P&ause"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69120
         TabIndex        =   42
         ToolTipText     =   "Pause tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton StopReplayButton 
         Caption         =   "St&op"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68400
         TabIndex        =   41
         ToolTipText     =   "Stop tickfile replay"
         Top             =   2340
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   -74640
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1020
         Width           =   6855
      End
      Begin VB.ComboBox ReplaySpeedCombo 
         Height          =   315
         ItemData        =   "MainForm.frx":005C
         Left            =   -74040
         List            =   "MainForm.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   39
         ToolTipText     =   "Adjust tickfile replay speed"
         Top             =   2460
         Width           =   1575
      End
      Begin MSComctlLib.ListView OpenOrdersList 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         Height          =   255
         Left            =   -74880
         TabIndex        =   66
         Top             =   5400
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label ReplayContractLabel 
         Height          =   375
         Left            =   -74880
         TabIndex        =   90
         Top             =   4800
         Width           =   7575
      End
      Begin VB.Label ReplayProgressLabel 
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   5160
         Width           =   7575
      End
      Begin VB.Label Label11 
         Caption         =   "Current contract details"
         Height          =   255
         Left            =   -74520
         TabIndex        =   64
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Select tickfile(s)"
         Height          =   255
         Left            =   -74520
         TabIndex        =   55
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Output path"
         Height          =   855
         Left            =   -74640
         TabIndex        =   54
         Top             =   3420
         Width           =   5655
      End
      Begin VB.Label Label8 
         Caption         =   "qazqazqaz"
         Height          =   255
         Left            =   -74640
         TabIndex        =   53
         Top             =   2940
         Width           =   5655
      End
      Begin VB.Label Label20 
         Caption         =   "Replay speed"
         Height          =   375
         Left            =   -74640
         TabIndex        =   52
         Top             =   2460
         Width           =   615
      End
   End
   Begin VB.Label Label16 
      Caption         =   "QT Port"
      Height          =   255
      Left            =   7920
      TabIndex        =   56
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

Implements LogListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "MainForm"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1

Private WithEvents mTradeBuildAPI As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1
Private WithEvents mTickfileManager As TradeBuild27.TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1
Private WithEvents mContracts As Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mTickers As Tickers
Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private mEt As ElapsedTimer

Private mRunningFromComandLine As Boolean

Private mOutputFormat As String
Private mOutputPath As String

Private mContract As Contract

Private mSupportedOutputFormats() As TradeBuild27.TickfileFormatSpecifier

Private mArguments As CommandLineParser
Private mNoUI As Boolean
Private mRun As Boolean

Private mMonths(12) As String

Private mNoWriteBars As Boolean
Private mNoWriteTicks As Boolean

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mNumberOfSessions As Long
Private mStartingSession As Long
Private mFromDate As Date
Private mFromTime As Date
Private mToDate As Date
Private mToTime As Date
Private mInFormatValue As String

Private mTickfiles As TickFileSpecifiers
Private mTickfilesSelected As Boolean

Private mLogFormatter As LogFormatter

Private mIsInDev                                    As Boolean

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

InitCommonControls

InitialiseTWUtilities

Set mUnhandledErrorHandler = UnhandledErrorHandler

ApplicationGroupName = "TradeWright"
ApplicationName = getAppTitle
SetupDefaultLogging Command

TaskQuantumMillisecs = 200
TaskConcurrency = 100
RunTasksAtLowerThreadPriority = False

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

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Load()

Const ProcName As String = "Form_Load"
On Error GoTo Err

On Error Resume Next
Set mTradeBuildAPI = TradeBuildAPI
On Error GoTo Err

If mTradeBuildAPI Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "The required version of TradeBuild is not installed"
End If

Set mTickers = mTradeBuildAPI.Tickers

AddStudyLibrary "CmnStudiesLib27.StudyLib", True, "Built-in"

setupDbTypeCombos

QTPortText.Text = "16240"

mOutputPath = App.path
OutputPathText = mOutputPath

disableInputDatabaseFields
disableQtFields
disableOutputDatabaseFields
enableOutputFileFields

WriteTickDataCheck.Value = vbChecked
WriteBarDataCheck.Value = vbChecked

If Not ProcessCommandLineArgs Then
    Unload Me
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

'================================================================================
' InfoListener Interface Members
'================================================================================

Private Sub LogListener_finish()

End Sub

Private Sub LogListener_Notify(ByVal logrec As TWUtilities30.LogRecord)
Const ProcName As String = "LogListener_Notify"
On Error GoTo Err

StatusText.SelStart = Len(StatusText.Text)
StatusText.SelLength = 0
If Len(StatusText.Text) <> 0 Then StatusText.SelText = vbCrLf
StatusText.SelText = mLogFormatter.FormatRecord(logrec)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub AdjustTimestampsCheck_Click()
Const ProcName As String = "AdjustTimestampsCheck_Click"
On Error GoTo Err

If AdjustTimestampsCheck = vbChecked Then
    AdjustSecondsStartText.Enabled = True
    AdjustSecondsEndText.Enabled = True
Else
    AdjustSecondsStartText.Enabled = False
    AdjustSecondsEndText.Enabled = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ClearTickfileListButton_Click()
Const ProcName As String = "ClearTickfileListButton_Click"
On Error GoTo Err

TickfileList.Clear
ClearTickfileListButton.Enabled = False
If Not mTickfileManager Is Nothing Then mTickfileManager.ClearTickfileSpecifiers
mTickfiles.Clear
ConvertButton.Enabled = False
StopButton.Enabled = False
mTickfilesSelected = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ConfigureButton_Click()
Const ProcName As String = "ConfigureButton_Click"
On Error GoTo Err

SelectTickfilesButton.Enabled = setupServiceProviders
TickfileList.Clear

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ContractFromServiceProviderOption_Click()
Const ProcName As String = "ContractFromServiceProviderOption_Click"
On Error GoTo Err

ContractSpecBuilder1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ContractInTickfileOption_Click()
Const ProcName As String = "ContractInTickfileOption_Click"
On Error GoTo Err

GetContractButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ContractSpecBuilder1_NotReady()
Const ProcName As String = "ContractSpecBuilder1_NotReady"
On Error GoTo Err

GetContractButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ContractSpecBuilder1_ready()
Const ProcName As String = "ContractSpecBuilder1_ready"
On Error GoTo Err

If ContractFromServiceProviderOption Then
    GetContractButton.Enabled = True
Else
    GetContractButton.Enabled = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ConvertButton_Click()
Const ProcName As String = "ConvertButton_Click"
On Error GoTo Err

If ContractFromServiceProviderOption And _
    mContract Is Nothing _
Then
    LogMessage "Can't convert - no contract details are available"
    Exit Sub
End If

SelectTickfilesButton.Enabled = False
ClearTickfileListButton.Enabled = False
ConvertButton.Enabled = False
StopButton.Enabled = True
ReplayProgressBar.Visible = True

Set mTickfileManager = mTickers.CreateTickFileManager(TickerOptions.TickerOptUseExchangeTimeZone Or _
                                        IIf(WriteTickDataCheck = vbChecked, TickerOptions.TickerOptWriteTickData Or TickerOptions.TickerOptIncludeMarketDepthInTickfile, 0) Or _
                                        IIf(WriteBarDataCheck = vbChecked, TickerOptions.TickerOptWriteTradeBarData, 0))
mTickfileManager.TickFileSpecifiers = mTickfiles
If ContractFromServiceProviderOption Then mTickfileManager.defaultContract = mContract
mTickfileManager.ReplaySpeed = 0
If AdjustTimestampsCheck = vbChecked Then
    mTickfileManager.TimestampAdjustmentStart = AdjustSecondsStartText
    mTickfileManager.TimestampAdjustmentEnd = AdjustSecondsEndText
End If

LogMessage "Tickfile conversion started"
mTickfileManager.StartReplay

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DatabaseInputOption_Click()
Const ProcName As String = "DatabaseInputOption_Click"
On Error GoTo Err

enableInputDatabaseFields
disableQtFields
disableContractDatabaseFields

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DatabaseOutputOption_Click()
Const ProcName As String = "DatabaseOutputOption_Click"
On Error GoTo Err

disableOutputFileFields
enableOutputDatabaseFields
disableContractDatabaseFields

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FileInputOption_Click()
Const ProcName As String = "FileInputOption_Click"
On Error GoTo Err

disableInputDatabaseFields
disableQtFields
If FileOutputOption Then
    enableContractDatabaseFields
Else
    disableContractDatabaseFields
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FileOutputOption_Click()
Const ProcName As String = "FileOutputOption_Click"
On Error GoTo Err

enableOutputFileFields
disableOutputDatabaseFields
If FileInputOption Or QtInputOption Then
    enableContractDatabaseFields
Else
    disableContractDatabaseFields
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FormatList_Click()
Dim i As Long
Const ProcName As String = "FormatList_Click"
On Error GoTo Err

mOutputFormat = ""
For i = 0 To UBound(mSupportedOutputFormats)
    If FormatList.Text = mSupportedOutputFormats(i).Name Then
        mOutputFormat = mSupportedOutputFormats(i).FormalID
        Exit Sub
    End If
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub GetContractButton_Click()
Const ProcName As String = "GetContractButton_Click"
On Error GoTo Err

Set mContracts = mTradeBuildAPI.LoadContracts(ContractSpecBuilder1.ContractSpecifier)
LogMessage "Requesting contract details"
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OutputPathButton_Click()
Dim PathChooser As PathChooser
Const ProcName As String = "OutputPathButton_Click"
On Error GoTo Err

Set PathChooser = New PathChooser
PathChooser.path = OutputPathText.Text
PathChooser.choose
If Not PathChooser.cancelled Then
    OutputPathText.Text = PathChooser.path
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OutputPathText_Change()
Const ProcName As String = "OutputPathText_Change"
On Error GoTo Err

mOutputPath = OutputPathText.Text

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub QtInputOption_Click()
Const ProcName As String = "QtInputOption_Click"
On Error GoTo Err

disableInputDatabaseFields
enableQtFields
If FileOutputOption Then
    enableContractDatabaseFields
Else
    disableContractDatabaseFields
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub SelectTickfilesButton_Click()
Dim tfs As TickfileSpecifier
Dim userCancelled As Boolean

Const ProcName As String = "SelectTickfilesButton_Click"
On Error GoTo Err

Set mTickfiles = SelectTickfiles(userCancelled)

If userCancelled Then Exit Sub

mTickfilesSelected = True

TickfileList.Clear
For Each tfs In mTickfiles
    TickfileList.AddItem tfs.FileName
Next
ClearTickfileListButton.Enabled = True

checkOkToConvert

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub StopButton_Click()
Const ProcName As String = "StopButton_Click"
On Error GoTo Err

ConvertButton.Enabled = True
StopButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
mTickfileManager.StopReplay

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub WriteBarDataCheck_Click()
Const ProcName As String = "WriteBarDataCheck_Click"
On Error GoTo Err

checkOkToConvert

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub WriteTickDataCheck_Click()
Const ProcName As String = "WriteTickDataCheck_Click"
On Error GoTo Err

checkOkToConvert

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mContracts Event Handlers
'================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
Const ProcName As String = "mContracts_ContractSpecifierInvalid"
On Error GoTo Err

LogMessage "Invalid contract specifier: " & _
                    Replace(mContracts.ContractSpecifier.ToString, vbCrLf, "; ")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mContracts_NoMoreContractDetails()
Dim tfs As TickfileSpecifier
Dim i As Long
Dim j As Long
Dim lSupportedInputTickfileFormats() As TradeBuild27.TickfileFormatSpecifier

Const ProcName As String = "mContracts_NoMoreContractDetails"
On Error GoTo Err

If mContracts.Count = 0 Then
    LogMessage "An invalid contract was specified"
    Exit Sub
End If

If mContracts.Count > 1 Then
    LogMessage "Unique contract not specified"
    Exit Sub
End If

Set mContract = mContracts(1)
ContractDetailsText = mContract.ToString

LogMessage "Contract details received"
If Not mRunningFromComandLine Then
    GetContractButton.Enabled = True
    Exit Sub
End If
    
If mArguments.Switch("from") Then mNumberOfSessions = 1
If mNumberOfSessions > (mStartingSession + 1) Then mNumberOfSessions = mStartingSession + 1

For i = 0 To mNumberOfSessions - 1
    Set tfs = New TickfileSpecifier
    mTickfiles.Add tfs
    With tfs
        Set .Contract = mContract
        If mArguments.Switch("from") Then
            .FromDate = mFromDate + mFromTime
            If mArguments.Switch("to") Then
                .ToDate = mToDate + mToTime
            Else
                .ToDate = DateAdd("n", 1, Now)
            End If
        Else
            .EntireSession = True
            .FromDate = DateAdd("d", -mStartingSession + i, Now)
        End If
            
        lSupportedInputTickfileFormats = mTradeBuildAPI.SupportedInputTickfileFormats
        For j = 0 To UBound(lSupportedInputTickfileFormats)
            If lSupportedInputTickfileFormats(j).Name = mInFormatValue Then
                .TickfileFormatID = lSupportedInputTickfileFormats(j).FormalID
                Exit For
            End If
        Next
        
        If .EntireSession Then
            .FileName = "Session (" & .FromDate & ") " & _
                            Replace(mContract.Specifier.ToString, vbCrLf, "; ")
        Else
            .FileName = .FromDate & "-" & .ToDate & " " & _
                            Replace(mContract.Specifier.ToString, vbCrLf, "; ")
        End If
        TickfileList.AddItem .FileName
    End With
Next

Set mTimer = CreateIntervalTimer(10)
mTimer.StartTimer

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_TickfileWriterNotification(ev As TradeBuild27.WriterEventData)
Const ProcName As String = "mTicker_TickfileWriterNotification"
On Error GoTo Err

If ev.Action = WriterFileCreated Then LogMessage "Created output tickfile: " & ev.FileName

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTickfileManager Event Handlers
'================================================================================

Private Sub mTickfileManager_QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal tickfileSizeBytes As Long, _
                ByVal pContract As Contract, _
                continueMode As ReplayContinueModes)
Const ProcName As String = "mTickfileManager_QueryReplayNextTickfile"
On Error GoTo Err

ReplayProgressBar.Min = 0
ReplayProgressBar.Max = 100
ReplayProgressBar.Value = 0
TickfileList.ListIndex = tickfileIndex - 1
LogMessage "Converting " & TickfileList.List(TickfileList.ListIndex)
ReplayContractLabel.Caption = pContract.Specifier.ToString
Set mEt = New ElapsedTimer
mEt.StartTiming

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mTickfileManager_ReplayCompleted()
Const ProcName As String = "mTickfileManager_ReplayCompleted"
On Error GoTo Err

ConvertButton.Enabled = True
StopButton.Enabled = False
SelectTickfilesButton.Enabled = True
ClearTickfileListButton.Enabled = True
ReplayProgressBar.Value = 0
ReplayProgressBar.Visible = False
ReplayContractLabel.Caption = ""
ReplayProgressLabel.Caption = ""

LogMessage "Tickfile conversion completed"

If mRun Then Unload Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub mTickfileManager_ReplayProgress( _
                            ByVal tickfileTimestamp As Date, _
                            ByVal eventsPlayed As Long, _
                            ByVal percentComplete As Single)

Const ProcName As String = "mTickfileManager_ReplayProgress"
On Error GoTo Err

ReplayProgressBar.Value = percentComplete
ReplayProgressBar.Refresh
ReplayProgressLabel.Caption = tickfileTimestamp & _
                                "  Processed " & _
                                eventsPlayed & _
                                " events" & _
                                IIf(percentComplete >= 1, Format(percentComplete, " \(0\%\)"), "")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mTickfileManager_TickerAllocated(ByVal pTicker As TradeBuild27.Ticker)
Const ProcName As String = "mTickfileManager_TickerAllocated"
On Error GoTo Err

Set mTicker = pTicker
mTicker.OutputTickfilePath = mOutputPath
mTicker.OutputTickfileFormat = mOutputFormat

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mTickfileManager_TickfileCompleted( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String)
Dim elapsed As Single
Const ProcName As String = "mTickfileManager_TickfileCompleted"
On Error GoTo Err

elapsed = mEt.ElapsedTimeMicroseconds

LogMessage "Processed " & mTicker.TickNumber & " ticks in " & Format(elapsed / 1000000, "0.0") & " seconds"
LogMessage "Ticks per second: " & CLng(mTicker.TickNumber / (elapsed / 1000000))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Const ProcName As String = "mTimer_TimerExpired"
On Error GoTo Err

Set mTickfileManager = mTickers.CreateTickFileManager(TickerOptions.TickerOptUseExchangeTimeZone Or _
                                        IIf(mNoWriteTicks, 0, TickerOptions.TickerOptWriteTickData Or TickerOptions.TickerOptIncludeMarketDepthInTickfile) Or _
                                        IIf(mNoWriteBars, 0, TickerOptions.TickerOptWriteTradeBarData))

mTickfileManager.TickFileSpecifiers = mTickfiles
mTickfileManager.ReplaySpeed = 0

mTickfileManager.StartReplay

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_Error( _
                ev As ErrorEventData)
Const ProcName As String = "mTradeBuildAPI_Error"
On Error Resume Next

handleFatalError
End Sub

Private Sub mTradeBuildAPI_Notification( _
                ev As NotificationEventData)
Dim spe As ServiceProviderError
Const ProcName As String = "mTradeBuildAPI_Notification"
On Error GoTo Err

If ev.EventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spe = mTradeBuildAPI.GetServiceProviderError
    LogMessage "Service provider error in " & spe.ServiceProviderName & ": error " & spe.ErrorCode & ": " & spe.Message
Else
    LogMessage "Notify " & ev.EventCode & ": " & ev.EventMessage
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As TWUtilities30.ErrorEventData)
handleFatalError
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

Private Sub checkOkToConvert()
Const ProcName As String = "checkOkToConvert"
On Error GoTo Err

If mTickfilesSelected Then
    If WriteTickDataCheck = vbChecked Or WriteBarDataCheck = vbChecked Then
        ConvertButton.Enabled = True
    Else
        ConvertButton.Enabled = False
    End If
Else
    ConvertButton.Enabled = False
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub disableContractDatabaseFields()
Const ProcName As String = "disableContractDatabaseFields"
On Error GoTo Err

disableControl ContractServerText
disableControl ContractDbTypeCombo
disableControl ContractDatabaseText
disableControl ContractUsernameText
disableControl ContractPasswordText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub disableControl( _
                ByVal ctrl As Control)
Const ProcName As String = "disableControl"
On Error GoTo Err

If TypeOf ctrl Is ComboBox Then
    Dim cb As ComboBox
    Set cb = ctrl
    cb.BackColor = vbButtonFace
    cb.Enabled = False
ElseIf TypeOf ctrl Is TextBox Then
    Dim tb As TextBox
    Set tb = ctrl
    tb.BackColor = vbButtonFace
    tb.Enabled = False
ElseIf TypeOf ctrl Is CommandButton Then
    Dim bt As CommandButton
    Set bt = ctrl
    bt.Enabled = False
ElseIf TypeOf ctrl Is ListBox Then
    Dim lb As ListBox
    Set lb = ctrl
    lb.Enabled = False
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub disableInputDatabaseFields()
Const ProcName As String = "disableInputDatabaseFields"
On Error GoTo Err

disableControl DbInServerText
disableControl DbInTypeCombo
disableControl DatabaseInText
disableControl UsernameInText
disableControl PasswordInText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub disableOutputDatabaseFields()
Const ProcName As String = "disableOutputDatabaseFields"
On Error GoTo Err

disableControl DbOutServerText
disableControl DbOutTypeCombo
disableControl DatabaseOutText
disableControl UsernameOutText
disableControl PasswordOutText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub disableOutputFileFields()
Const ProcName As String = "disableOutputFileFields"
On Error GoTo Err

disableControl FormatList
disableControl OutputPathText
disableControl OutputPathButton

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub disableQtFields()
Const ProcName As String = "disableQtFields"
On Error GoTo Err

disableControl QTServerText
disableControl QTPortText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub enableContractDatabaseFields()
Const ProcName As String = "enableContractDatabaseFields"
On Error GoTo Err

enableControl ContractServerText
enableControl ContractDbTypeCombo
enableControl ContractDatabaseText
enableControl ContractUsernameText
enableControl ContractPasswordText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub enableControl( _
                ByVal ctrl As Control)
Const ProcName As String = "enableControl"
On Error GoTo Err

If TypeOf ctrl Is ComboBox Then
    Dim cb As ComboBox
    Set cb = ctrl
    cb.BackColor = vbWindowBackground
    cb.Enabled = True
ElseIf TypeOf ctrl Is TextBox Then
    Dim tb As TextBox
    Set tb = ctrl
    tb.BackColor = vbWindowBackground
    tb.Enabled = True
ElseIf TypeOf ctrl Is CommandButton Then
    Dim bt As CommandButton
    Set bt = ctrl
    bt.Enabled = True
ElseIf TypeOf ctrl Is ListBox Then
    Dim lb As ListBox
    Set lb = ctrl
    lb.Enabled = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub enableInputDatabaseFields()
Const ProcName As String = "enableInputDatabaseFields"
On Error GoTo Err

enableControl DbInServerText
enableControl DbInTypeCombo
enableControl DatabaseInText
enableControl UsernameInText
enableControl PasswordInText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub enableOutputDatabaseFields()
Const ProcName As String = "enableOutputDatabaseFields"
On Error GoTo Err

enableControl DbOutServerText
enableControl DbOutTypeCombo
enableControl DatabaseOutText
enableControl UsernameOutText
enableControl PasswordOutText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub enableOutputFileFields()
Const ProcName As String = "enableOutputFileFields"
On Error GoTo Err

enableControl FormatList
enableControl OutputPathText
enableControl OutputPathButton

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub enableQtFields()
Const ProcName As String = "enableQtFields"
On Error GoTo Err

enableControl QTServerText
enableControl QTPortText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getAppTitle() As String
getAppTitle = AppName & _
                " v" & _
                App.Major & "." & App.Minor
End Function

Private Sub handleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly. Note that normally one would use the
' End statement to terminate a VB6 program abruptly. However the TWUtilities component interferes
' with the End statement's processing and prevents proper shutdown, so we use the
' TWUtilities component's EndProcess method instead. (However if we are running in the
' development environment, then we call End because the EndProcess method kills the
' entire development environment as well which can have undesirable side effects if other
' components are also loaded.)

If mIsInDev Then
    ' this tells TWUtilities that we've now handled this unhandled error. Not actually
    ' needed here because the End statement will prevent return to TWUtilities
    UnhandledErrorHandler.Handled = True
    End
Else
    EndProcess
End If

End Sub

Private Function inDev() As Boolean
Const ProcName As String = "inDev"

mIsInDev = True
inDev = True

End Function

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
Dim toValue As String
Dim sessionsValue As String
Dim outFormatValue As String
Dim QTServerValue As String
Dim commaPosn As Long
Dim contractSpec As ContractSpecifier


Const ProcName As String = "ProcessCommandLineArgs"
On Error GoTo Err

Set mArguments = CreateCommandLineParser(Command)

If mArguments.Switch("?") Then
    Dim s As String
    s = "tickfilemanager [symbol  localSymbol|NOLOCALSYMBOL sectype " & vbCrLf & _
            "                month|NOMONTH exchange currency [strike] [right]]" & vbCrLf & _
            "                [/from:yyyymmdd[hhmmss]] " & vbCrLf & _
            "                [/to:yyyymmdd[hhmmss]] " & vbCrLf & _
            "                [/sessions:n[,m]]" & vbCrLf & _
            "                [/inFormat:inputTickfileFormat" & vbCrLf & _
            "                [/outFormat:outputTickfileFormat" & vbCrLf & _
            "                [/outpath:path]" & vbCrLf & _
            "                [/noWriteTicks  |  /nwt]" & vbCrLf & _
            "                [/noWriteBars  |  /nwb]" & vbCrLf & _
            "                [/noUI]  [/run]" & vbCrLf & _
            "                [/QTserver:[server][,port]]" & vbCrLf & _
            "                [/log:filename] " & vbCrLf & _
            "                [/loglevel:levelName]" & vbCrLf
    s = s & "  where" & vbCrLf & _
            vbCrLf & _
            "    levelname is one of:" & vbCrLf & _
            "       None    or 0" & vbCrLf & _
            "       Severe  or S" & vbCrLf & _
            "       Warning or W" & vbCrLf & _
            "       Info    or I" & vbCrLf & _
            "       Normal  or N" & vbCrLf & _
            "       Detail  or D" & vbCrLf & _
            "       Medium  or M" & vbCrLf & _
            "       High    or H" & vbCrLf & _
            "       All     or A"
    s = s & vbCrLf & _
            "Notes:" & vbCrLf & _
            "   If /from is supplied, /sessions is ignored." & vbCrLf & _
            "   If /from is not supplied, /to is ignored." & vbCrLf & _
            "   In /sessions, n is the number of sessions to supply, and m" & vbCrLf & _
            "      is the number of sessions before current to start at." & vbCrLf & _
            "      m defaults to 1. If m is zero, the current session is" & vbCrLf & _
            "      supplied." & vbCrLf & _
            "   In /QTserver, port defaults to 16240."
    MsgBox s, _
            , _
            "Usage"
    ProcessCommandLineArgs = False
    Exit Function
End If

Set mLogFormatter = CreateBasicLogFormatter(TimestampTimeOnlyISO8601)
GetLogger("log").AddLogListener Me

LogMessage "Log file: " & DefaultLogFileName(Command)

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
        unpackDateTimeString fromValue, mFromDate, mFromTime
        If Err.Number <> 0 Then
            MsgBox fromValue & " is not a valid date and time (format yyyymmdd[hhmmss])"
            ProcessCommandLineArgs = False
            Exit Function
        End If
        On Error GoTo Err
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
        unpackDateTimeString toValue, mToDate, mToTime
        If Err.Number <> 0 Then
            MsgBox toValue & " is not a valid date and time (format yyyymmdd[hhmmss])"
            ProcessCommandLineArgs = False
            Exit Function
        End If
        On Error GoTo Err
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

mStartingSession = 1
If mArguments.Switch("sessions") Then
    sessionsValue = mArguments.SwitchValue("sessions")
    If Len(sessionsValue) = 0 Then
        MsgBox "Error - sessions should be /sessions:n[,m]"
        ProcessCommandLineArgs = False
        Exit Function
    End If
    
    On Error Resume Next
    If InStr(1, sessionsValue, ",") Then
        mNumberOfSessions = CLng(Left$(sessionsValue, InStr(1, sessionsValue, ",") - 1))
        If Err.Number <> 0 Or mNumberOfSessions < 1 Then
            MsgBox "Error - sessions should be /sessions:n[,m] where n and m are integers, n>=1 and m>=0"
            ProcessCommandLineArgs = False
            Exit Function
        End If
        mStartingSession = CLng(Right$(sessionsValue, Len(sessionsValue) - InStr(1, sessionsValue, ",")))
        If Err.Number <> 0 Or mStartingSession < 0 Then
            MsgBox "Error - sessions should be /sessions:n[,m] where n and m are integers, n>=1 and m>=0"
            ProcessCommandLineArgs = False
            Exit Function
        End If
    Else
        mNumberOfSessions = sessionsValue
        If Err.Number <> 0 Or mNumberOfSessions < 1 Then
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

If mArguments.Switch("informat") Then mInFormatValue = mArguments.SwitchValue("informat")

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
    mRunningFromComandLine = True
    
    Set contractSpec = CreateContractSpecifier( _
                                localSymbolValue, _
                                symbolValue, _
                                exchangeValue, _
                                SecTypeFromString(secTypeValue), _
                                currencyValue, _
                                monthValue, _
                                IIf(strikevalue = "", 0, strikevalue), _
                                OptionRightFromString(rightValue))
    Set mContracts = mTradeBuildAPI.LoadContracts(contractSpec)

End If

ProcessCommandLineArgs = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupContractDatabaseAsContractSP() As Boolean
Dim sp As Object
Const ProcName As String = "setupContractDatabaseAsContractSP"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TBInfoBase27.ContractInfoSrvcProvider", _
                            True, _
                            "Database Name=" & ContractDatabaseText & _
                            ";Database Type=" & ContractDbTypeCombo & _
                            ";Server=" & ContractServerText & _
                            ";User name=" & ContractUsernameText & _
                            ";Password=" & ContractPasswordText, _
                            "Contracts database", _
                            "Contract database")
On Error GoTo Err
If Not sp Is Nothing Then
    setupContractDatabaseAsContractSP = True
Else
    LogMessage "Can't configure Contract Info Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setupDbTypeCombos()
Const ProcName As String = "setupDbTypeCombos"
On Error GoTo Err

DbInTypeCombo.AddItem DatabaseTypeToString(DbMySQL5)
DbOutTypeCombo.AddItem DatabaseTypeToString(DbMySQL5)
ContractDbTypeCombo.AddItem DatabaseTypeToString(DbMySQL5)

DbInTypeCombo.AddItem DatabaseTypeToString(DbSQLServer)
DbOutTypeCombo.AddItem DatabaseTypeToString(DbSQLServer)
ContractDbTypeCombo.AddItem DatabaseTypeToString(DbSQLServer)

DbInTypeCombo.ListIndex = 0
DbOutTypeCombo.ListIndex = 0
ContractDbTypeCombo.ListIndex = 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function setupInDatabase() As Boolean
Dim sp As Object
Const ProcName As String = "setupInDatabase"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TBInfoBase27.TickfileServiceProvider", _
                            True, _
                            "Database Name=" & DatabaseInText & _
                            ";Database Type=" & DbInTypeCombo & _
                            ";Server=" & DbInServerText & _
                            ";User name=" & UsernameInText & _
                            ";Password=" & PasswordInText & _
                            ";Access mode=ReadOnly", _
                            "Input tick database", _
                            "Historical tick data input from database")
                                                        
On Error GoTo Err
If Not sp Is Nothing Then
    setupInDatabase = True
Else
    LogMessage "Can't configure Input Database Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupInDatabaseAsContractSP() As Boolean
Dim sp As Object
Const ProcName As String = "setupInDatabaseAsContractSP"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TBInfoBase27.ContractInfoSrvcProvider", _
                            True, _
                            "Database Name=" & DatabaseInText & _
                            ";Database Type=" & DbInTypeCombo & _
                            ";Server=" & DbInServerText & _
                            ";User name=" & UsernameInText & _
                            ";Password=" & PasswordInText, _
                            "Contract database", _
                            "Contract data from input database")
On Error GoTo Err
If Not sp Is Nothing Then
    setupInDatabaseAsContractSP = True
Else
    LogMessage "Can't configure Contract Info Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupInFileSP() As Boolean
Dim sp As Object
Const ProcName As String = "setupInFileSP"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TickfileSP27.TickfileServiceProvider", _
                            True, _
                            "Access mode=ReadOnly", _
                            "Input tickfiles", _
                            "Historical tick data input from files")
On Error GoTo Err
If Not sp Is Nothing Then
    setupInFileSP = True
Else
    LogMessage "Can't configure input Tickfile Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupOutBarDatabase() As Boolean
Dim sp As Object
Const ProcName As String = "setupOutBarDatabase"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TBInfoBase27.HistDataServiceProvider", _
                            True, _
                            "Database Name=" & DatabaseOutText & _
                            ";Database Type=" & DbOutTypeCombo & _
                            ";Server=" & DbOutServerText & _
                            ";User name=" & UsernameOutText & _
                            ";Password=" & PasswordOutText & _
                            ";Use Synchronous Writes=" & IIf(AsyncWritesCheck = vbChecked, "No", "Yes"), _
                            "Output bar database", _
                            "Historical bar data output to database")
On Error GoTo Err
If Not sp Is Nothing Then
    setupOutBarDatabase = True
Else
    LogMessage "Can't configure Historic Bar Data Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupOutTickDatabase() As Boolean
Dim sp As Object
Const ProcName As String = "setupOutTickDatabase"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TBInfoBase27.TickfileServiceProvider", _
                            True, _
                            "Database Name=" & DatabaseOutText & _
                            ";Database Type=" & DbOutTypeCombo & _
                            ";Server=" & DbOutServerText & _
                            ";User name=" & UsernameOutText & _
                            ";Password=" & PasswordOutText & _
                            ";Use Synchronous Writes=" & IIf(AsyncWritesCheck = vbChecked, "No", "Yes") & _
                            ";Access mode=WriteOnly", _
                            "Output tick database", _
                            "Historical tick data output to database")
On Error GoTo Err
If Not sp Is Nothing Then
    setupOutTickDatabase = True
Else
    LogMessage "Can't configure Historic Tick Data Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupOutTickDatabaseAsContractSP() As Boolean
Dim sp As Object
Const ProcName As String = "setupOutTickDatabaseAsContractSP"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TBInfoBase27.ContractInfoSrvcProvider", _
                            True, _
                            "Database Name=" & DatabaseOutText & _
                            ";Database Type=" & DbOutTypeCombo & _
                            ";Server=" & DbOutServerText & _
                            ";User name=" & UsernameOutText & _
                            ";Password=" & PasswordOutText, _
                            "Contract database", _
                            "Contract data from output database")
On Error GoTo Err
If Not sp Is Nothing Then
    setupOutTickDatabaseAsContractSP = True
Else
    LogMessage "Can't configure Contract Info Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupOutFileSP() As Boolean
Dim sp As Object
Const ProcName As String = "setupOutFileSP"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "TickfileSP27.TickfileServiceProvider", _
                            True, _
                            "Access mode=WriteOnly", _
                            "Output tickfiles", _
                            "Historical tick data output to files")
On Error GoTo Err
If Not sp Is Nothing Then
    setupOutFileSP = True
Else
    LogMessage "Can't configure output Tickfile Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupQtSP() As Boolean
Dim sp As Object
Const ProcName As String = "setupQtSP"
On Error GoTo Err

On Error Resume Next
Set sp = mTradeBuildAPI.ServiceProviders.Add( _
                            "QTSP27.QTTickfileServiceProvider", _
                            True, _
                            "Provider Key=QTIB" & _
                            ";Server=" & QTServerText & _
                            ";Port=" & QTPortText & _
                            ";Password=" & _
                            ";Connection Retry Interval Secs=10" & _
                            ";Keep connection=true", _
                            "QuoteTracker input tickdata", _
                            "Historical tick data input from QuoteTracker")
On Error GoTo Err
If Not sp Is Nothing Then
    setupQtSP = True
Else
    LogMessage "Can't configure QuoteTracker Service Provider"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupServiceProviders() As Boolean
Dim i As Long

Const ProcName As String = "setupServiceProviders"
On Error GoTo Err

setupServiceProviders = True

mTradeBuildAPI.ServiceProviders.RemoveAll
FormatList.Clear

WriteTickDataCheck.Enabled = False
WriteBarDataCheck.Enabled = False

If FileInputOption Then
    If Not setupInFileSP Then
        setupServiceProviders = False
    End If
End If

If FileOutputOption Then
    If Not setupOutFileSP Then
        setupServiceProviders = False
    End If
End If

If (FileInputOption Or QtInputOption) And FileOutputOption Then
    If Not setupContractDatabaseAsContractSP Then
        setupServiceProviders = False
    End If
End If

If DatabaseOutputOption Then
    If Not setupOutTickDatabaseAsContractSP Then
        setupServiceProviders = False
    End If
End If

If FileOutputOption Then
    WriteTickDataCheck.Enabled = True
    WriteBarDataCheck = vbUnchecked
End If

If Not (FileInputOption Or QtInputOption) And FileOutputOption Then
    If Not setupInDatabaseAsContractSP Then
        setupServiceProviders = False
    End If
End If

If DatabaseInputOption Then
    If Not setupInDatabase Then
        setupServiceProviders = False
    End If
    
End If

If DatabaseOutputOption Then
    If Not setupOutTickDatabase Then
        setupServiceProviders = False
    Else
        WriteTickDataCheck.Enabled = True
        WriteBarDataCheck.Enabled = True
    End If

    If Not setupOutBarDatabase Then
        setupServiceProviders = False
    End If
    
End If

If QtInputOption Then
    If Not setupQtSP Then
        setupServiceProviders = False
    End If
End If

If Not setupServiceProviders Then
    LogMessage "Service provider configuration failed"
    WriteTickDataCheck.Enabled = False
    WriteBarDataCheck.Enabled = False
    Exit Function
End If

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

For i = 0 To UBound(mSupportedOutputFormats)
    FormatList.AddItem mSupportedOutputFormats(i).Name
Next

FormatList.ListIndex = 0

LogMessage "Service provider configuration succeeded"

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub unpackDateTimeString( _
                            ByVal timestampString As String, _
                            ByRef dateOut As Date, _
                            ByRef timeOut As Date)
Const ProcName As String = "unpackDateTimeString"
On Error GoTo Err

dateOut = CDate(mMonths(Mid$(timestampString, 5, 2)) & " " & _
                            Mid$(timestampString, 7, 2) & " " & _
                            Left$(timestampString, 4))
If Len(timestampString) = 14 Then
    timeOut = CDate(Mid$(timestampString, 9, 2) & ":" & _
                            Mid$(timestampString, 11, 2) & ":" & _
                            Mid(timestampString, 13, 2))
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


