VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#321.2#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.UserControl FeaturesPanel 
   Appearance      =   0  'Flat
   BackColor       =   &H00CDF3FF&
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   DefaultCancel   =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   4065
   Begin TabDlg.SSTab FeaturesSSTab 
      Height          =   9030
      Left            =   -30
      TabIndex        =   1
      Top             =   645
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   15928
      _Version        =   393216
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
      TabPicture(0)   =   "FeaturesPanel.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TickersPicture"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FeaturesPanel.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LiveChartPicture"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FeaturesPanel.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "HistChartPicture"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "FeaturesPanel.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ReplayTickerPicture"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "FeaturesPanel.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ConfigPicture"
      Tab(4).ControlCount=   1
      Begin VB.PictureBox ConfigPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   42
         Top             =   0
         Width           =   4125
         Begin TWControls40.TWImageCombo CurrentConfigCombo 
            Height          =   270
            Left            =   240
            TabIndex        =   58
            Top             =   390
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   476
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
            Locked          =   -1  'True
            MouseIcon       =   "FeaturesPanel.ctx":008C
            Text            =   ""
         End
         Begin VB.Frame ChangeChartStylesFrame 
            Caption         =   "Change chart styles"
            Height          =   2895
            Left            =   240
            TabIndex        =   50
            Top             =   3720
            Width           =   2535
            Begin VB.PictureBox ChangeChartStylesPicture 
               BorderStyle     =   0  'None
               Height          =   2535
               Left            =   120
               ScaleHeight     =   2535
               ScaleWidth      =   2295
               TabIndex        =   51
               Top             =   240
               Width           =   2295
               Begin VB.CheckBox ApplyStyleHistCheck 
                  Appearance      =   0  'Flat
                  Caption         =   "Historical charts"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   480
                  TabIndex        =   57
                  Top             =   1440
                  Width           =   1935
               End
               Begin VB.CheckBox ApplyStyleLiveCheck 
                  Appearance      =   0  'Flat
                  Caption         =   "Live charts"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   480
                  TabIndex        =   56
                  Top             =   1200
                  Width           =   1455
               End
               Begin TWControls40.TWImageCombo ChartStylesCombo 
                  Height          =   270
                  Left            =   120
                  TabIndex        =   53
                  Top             =   480
                  Width           =   2160
                  _ExtentX        =   3810
                  _ExtentY        =   476
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
                  MouseIcon       =   "FeaturesPanel.ctx":00A8
                  Text            =   ""
               End
               Begin TWControls40.TWButton ApplyStyleButton 
                  Height          =   495
                  Left            =   120
                  TabIndex        =   52
                  Top             =   1920
                  Width           =   2160
                  _ExtentX        =   3810
                  _ExtentY        =   873
                  Appearance      =   0
                  Caption         =   "Apply style"
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
               End
               Begin VB.Label Label7 
                  Caption         =   "Apply to"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   55
                  Top             =   960
                  Width           =   1575
               End
               Begin VB.Label Label 
                  Caption         =   "Available styles"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   54
                  Top             =   120
                  Width           =   2055
               End
            End
         End
         Begin VB.Frame ThemeFrame 
            Caption         =   "Theme"
            Height          =   1815
            Left            =   240
            TabIndex        =   45
            Top             =   1680
            Width           =   2535
            Begin VB.PictureBox ThemePicture 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1455
               Left            =   120
               ScaleHeight     =   1455
               ScaleWidth      =   2295
               TabIndex        =   46
               Top             =   240
               Width           =   2295
               Begin VB.OptionButton BlueThemeOption 
                  Caption         =   "Blue"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   48
                  Top             =   480
                  Width           =   2295
               End
               Begin VB.OptionButton NativeThemeOption 
                  Caption         =   "Native"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   49
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.OptionButton BlackThemeOption 
                  Caption         =   "Black"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   47
                  Top             =   120
                  Width           =   1815
               End
            End
         End
         Begin TWControls40.TWButton ConfigEditorButton 
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Appearance      =   0
            Caption         =   "Show config editor"
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
         Begin VB.Label Label6 
            Caption         =   "Current configuration is:"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.PictureBox ReplayTickerPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   33
         Top             =   0
         Width           =   4125
         Begin TWControls40.TWImageCombo ReplaySpeedCombo 
            Height          =   270
            Left            =   1200
            TabIndex        =   38
            Top             =   4080
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   476
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
            MouseIcon       =   "FeaturesPanel.ctx":00C4
            Text            =   ""
         End
         Begin TWControls40.TWButton StopReplayButton 
            Height          =   495
            Left            =   3360
            TabIndex        =   37
            ToolTipText     =   "Stop tickfile replay"
            Top             =   4680
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            Appearance      =   0
            Caption         =   "St&op"
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
         End
         Begin TWControls40.TWButton PauseReplayButton 
            Height          =   495
            Left            =   2640
            TabIndex        =   36
            ToolTipText     =   "Pause tickfile replay"
            Top             =   4680
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            Appearance      =   0
            Caption         =   "P&ause"
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
         End
         Begin TWControls40.TWButton PlayTickFileButton 
            Height          =   495
            Left            =   1920
            TabIndex        =   35
            ToolTipText     =   "Start or resume tickfile replay"
            Top             =   4680
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            Appearance      =   0
            Caption         =   "&Play"
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
         End
         Begin TradingUI27.TickfileOrganiser TickfileOrganiser1 
            Height          =   3720
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   6562
         End
         Begin MSComctlLib.ProgressBar ReplayProgressBar 
            Height          =   135
            Left            =   120
            TabIndex        =   39
            Top             =   5640
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   238
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label20 
            Caption         =   "Replay speed"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label ReplayProgressLabel 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   5400
            Width           =   3855
         End
      End
      Begin VB.PictureBox HistChartPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   20
         Top             =   0
         Width           =   4125
         Begin MSComCtl2.DTPicker FromDatePicker 
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   128
            CalendarForeColor=   16711680
            CalendarTitleBackColor=   16777215
            CalendarTitleForeColor=   12632256
            CalendarTrailingForeColor=   65280
            CheckBox        =   -1  'True
            CustomFormat    =   "yyy-MM-dd HH:mm"
            Format          =   121831427
            CurrentDate     =   39365
         End
         Begin VB.TextBox NumHistHistoryBarsText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2760
            TabIndex        =   23
            Text            =   "500"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox HistSessionOnlyCheck 
            Appearance      =   0  'Flat
            Caption         =   "Session only"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin TradingUI27.ContractSearch HistContractSearch 
            Height          =   4455
            Left            =   120
            TabIndex        =   21
            Top             =   2760
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7858
         End
         Begin TradingUI27.TimeframeSelector HistChartTimeframeSelector 
            Height          =   270
            Left            =   1920
            TabIndex        =   24
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   476
         End
         Begin MSComCtl2.DTPicker ToDatePicker 
            Height          =   375
            Left            =   1920
            TabIndex        =   25
            Top             =   1800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyy-MM-dd HH:mm"
            Format          =   121831427
            CurrentDate     =   39365
         End
         Begin TWControls40.TWImageCombo HistChartStylesCombo 
            Height          =   270
            Left            =   1920
            TabIndex        =   27
            Top             =   2280
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   476
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
            MouseIcon       =   "FeaturesPanel.ctx":00E0
            Text            =   ""
         End
         Begin VB.Label Label5 
            Caption         =   "To"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "From"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Timeframe"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Number of history bars"
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Style"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   2280
            Width           =   1455
         End
      End
      Begin VB.PictureBox LiveChartPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75000
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   11
         Top             =   0
         Width           =   4125
         Begin VB.TextBox NumLiveHistoryBarsText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2760
            TabIndex        =   15
            Text            =   "500"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox LiveSessionOnlyCheck 
            Appearance      =   0  'Flat
            Caption         =   "Session only"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin TWControls40.TWButton LiveChartButton1 
            Height          =   375
            Left            =   3000
            TabIndex        =   13
            Top             =   2040
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "Show &Chart"
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
         End
         Begin TWControls40.TWImageCombo LiveChartStylesCombo 
            Height          =   270
            Left            =   1920
            TabIndex        =   12
            Top             =   1560
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   476
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
            MouseIcon       =   "FeaturesPanel.ctx":00FC
            Text            =   ""
         End
         Begin TradingUI27.TimeframeSelector LiveChartTimeframeSelector 
            Height          =   270
            Left            =   1920
            TabIndex        =   16
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   476
         End
         Begin VB.Label Label18 
            Caption         =   "Timeframe"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Number of history bars"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Style"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.PictureBox TickersPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   0
         ScaleHeight     =   9015
         ScaleWidth      =   4125
         TabIndex        =   7
         Top             =   0
         Width           =   4125
         Begin TradingUI27.ContractSearch LiveContractSearch 
            Height          =   5415
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   9551
         End
         Begin TWControls40.TWButton LiveChartButton 
            Height          =   375
            Left            =   2640
            TabIndex        =   2
            Top             =   6360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Appearance      =   0
            Caption         =   "Chart"
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
         End
         Begin TWControls40.TWButton StopTickerButton 
            Height          =   375
            Left            =   2640
            TabIndex        =   5
            Top             =   5880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Appearance      =   0
            Caption         =   "Sto&p"
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
         End
         Begin TWControls40.TWButton OrderTicketButton 
            Height          =   375
            Left            =   2640
            TabIndex        =   4
            Top             =   6840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Appearance      =   0
            Caption         =   "&Order ticket"
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
         End
         Begin TWControls40.TWButton MarketDepthButton 
            Height          =   375
            Left            =   2640
            TabIndex        =   3
            Top             =   7320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "&Mkt depth"
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
         End
      End
   End
   Begin VB.PictureBox HidePicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3720
      MouseIcon       =   "FeaturesPanel.ctx":0118
      MousePointer    =   99  'Custom
      Picture         =   "FeaturesPanel.ctx":026A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      ToolTipText     =   "Hide Features Panel"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox UnpinPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3360
      MouseIcon       =   "FeaturesPanel.ctx":07F4
      MousePointer    =   99  'Custom
      Picture         =   "FeaturesPanel.ctx":0946
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      ToolTipText     =   "Unpin Features Panel"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox PinPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00CDF3FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3360
      MouseIcon       =   "FeaturesPanel.ctx":0ED0
      MousePointer    =   99  'Custom
      Picture         =   "FeaturesPanel.ctx":1022
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      ToolTipText     =   "Pin Features Panel"
      Top             =   30
      Width           =   240
   End
   Begin MSComctlLib.TabStrip FeaturesTabStrip 
      Height          =   640
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   1138
      TabWidthStyle   =   1
      HotTracking     =   -1  'True
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
End
Attribute VB_Name = "FeaturesPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IStateChangeListener
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event ConfigsChanged()
Event Hide()
Event HistContractSearchCancelled()
Event HistContractSearchCleared()
Event HistContractsLoaded(ByVal pContracts As IContracts)
Event LiveContractSearchCancelled()
Event LiveContractSearchCleared()
Event LiveContractsLoaded(ByVal pContracts As IContracts)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Pin()
Event Unpin()

'@================================================================================
' Enums
'@================================================================================

Private Enum FeaturesTabIndexNumbers
    FeaturesTabIndexTickers
    FeaturesTabIndexLiveCharts
    FeaturesTabIndexHistoricalCharts
    FeaturesTabIndexTickfileReplay
    FeaturesTabIndexConfig
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "FeaturesPanel"

Private Const MinimumHeightTwips                    As Long = 8580
Private Const MinimumWidthTwips                     As Long = 4065

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildAPI                              As TradeBuildAPI
Private mConfigStore                                As ConfigurationStore
Private mAppInstanceConfig                          As ConfigurationSection

Private WithEvents mTickerGrid                      As TickerGrid
Attribute mTickerGrid.VB_VarHelpID = -1
Private WithEvents mTickers                         As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mInfoPanel                                  As InfoPanel
Private mInfoPanelFloating                          As InfoPanel

Private mChartForms                                 As ChartForms
Private mOrderTicket                                As fOrderTicket

Private WithEvents mReplayController                As ReplayController
Attribute mReplayController.VB_VarHelpID = -1
Private WithEvents mTickfileReplayTC                As TaskController
Attribute mTickfileReplayTC.VB_VarHelpID = -1

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent Mouseup(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()

If UserControl.Width < MinimumWidthTwips Then UserControl.Width = MinimumWidthTwips
If UserControl.Height < MinimumHeightTwips Then UserControl.Height = MinimumHeightTwips

UnpinPicture.Left = UserControl.Width - 47 * Screen.TwipsPerPixelX
PinPicture.Left = UserControl.Width - 47 * Screen.TwipsPerPixelX
HidePicture.Left = UserControl.Width - 23 * Screen.TwipsPerPixelX

FeaturesTabStrip.Width = UserControl.Width

FeaturesSSTab.Width = UserControl.Width + 4 * Screen.TwipsPerPixelX
FeaturesSSTab.Height = UserControl.Height + 2 * Screen.TwipsPerPixelX

TickersPicture.Width = UserControl.Width + 4 * Screen.TwipsPerPixelX
TickersPicture.Height = UserControl.Height - TickersPicture.Top

LiveContractSearch.Width = UserControl.Width - 16 * Screen.TwipsPerPixelX
StopTickerButton.Left = UserControl.Width - StopTickerButton.Width - 8 * Screen.TwipsPerPixelX
LiveChartButton.Left = StopTickerButton.Left
OrderTicketButton.Left = StopTickerButton.Left
MarketDepthButton.Left = StopTickerButton.Left

LiveChartPicture.Width = UserControl.Width + 4 * Screen.TwipsPerPixelX
LiveChartPicture.Height = UserControl.Height - LiveChartPicture.Top

LiveChartTimeframeSelector.Width = UserControl.Width - LiveChartTimeframeSelector.Left - 8 * Screen.TwipsPerPixelX
NumLiveHistoryBarsText.Width = UserControl.Width - NumLiveHistoryBarsText.Left - 8 * Screen.TwipsPerPixelX
LiveChartStylesCombo.Width = UserControl.Width - LiveChartStylesCombo.Left - 8 * Screen.TwipsPerPixelX
LiveChartButton1.Left = UserControl.Width - LiveChartButton1.Width - 8 * Screen.TwipsPerPixelX

HistChartPicture.Width = UserControl.Width + 4 * Screen.TwipsPerPixelX
HistChartPicture.Height = UserControl.Height - HistChartPicture.Top

HistChartTimeframeSelector.Width = UserControl.Width - HistChartTimeframeSelector.Left - 8 * Screen.TwipsPerPixelX
NumHistHistoryBarsText.Width = UserControl.Width - NumHistHistoryBarsText.Left - 8 * Screen.TwipsPerPixelX
HistChartStylesCombo.Width = UserControl.Width - HistChartStylesCombo.Left - 8 * Screen.TwipsPerPixelX
HistContractSearch.Width = UserControl.Width - 16 * Screen.TwipsPerPixelX
FromDatePicker.Width = UserControl.Width - FromDatePicker.Left - 8 * Screen.TwipsPerPixelX
ToDatePicker.Width = UserControl.Width - ToDatePicker.Left - 8 * Screen.TwipsPerPixelX

ReplayTickerPicture.Width = UserControl.Width + 4 * Screen.TwipsPerPixelX
ReplayTickerPicture.Height = UserControl.Height - ReplayTickerPicture.Top

TickfileOrganiser1.Width = UserControl.Width - 16 * Screen.TwipsPerPixelX
ReplaySpeedCombo.Width = UserControl.Width - ReplaySpeedCombo.Left - 8 * Screen.TwipsPerPixelX
StopReplayButton.Left = UserControl.Width - StopReplayButton.Width - 8 * Screen.TwipsPerPixelX
PauseReplayButton.Left = UserControl.Width - StopReplayButton.Width - PauseReplayButton.Width - 2 * 8 * Screen.TwipsPerPixelX
PlayTickFileButton.Left = UserControl.Width - StopReplayButton.Width - PauseReplayButton.Width - PlayTickFileButton.Width - 3 * 8 * Screen.TwipsPerPixelX
ReplayProgressLabel.Width = UserControl.Width - 16 * Screen.TwipsPerPixelX
ReplayProgressBar.Width = UserControl.Width - 16 * Screen.TwipsPerPixelX

ConfigPicture.Width = UserControl.Width + 4 * Screen.TwipsPerPixelX
ConfigPicture.Height = UserControl.Height - ConfigPicture.Top

CurrentConfigCombo.Width = UserControl.Width - CurrentConfigCombo.Left - 16 * Screen.TwipsPerPixelX

End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim lTicker As Ticker
Set lTicker = ev.Source

Select Case ev.State
Case MarketDataSourceStates.MarketDataSourceStateCreated

Case MarketDataSourceStates.MarketDataSourceStateReady
Case MarketDataSourceStates.MarketDataSourceStateRunning
    If lTicker Is getSelectedDataSource Then
        If lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled Then OrderTicketButton.Enabled = True
        MarketDepthButton.Enabled = True
        LiveChartButton1.Enabled = True
        LiveChartButton.Enabled = True
    End If
    
Case MarketDataSourceStates.MarketDataSourceStatePaused

Case MarketDataSourceStates.MarketDataSourceStateStopped
    If getSelectedDataSource Is Nothing Then
        StopTickerButton.Enabled = False
        MarketDepthButton.Enabled = False
        LiveChartButton1.Enabled = False
        LiveChartButton.Enabled = False
    End If
    
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ApplyStyleButton_Click()
Const ProcName As String = "ApplyStyleButton_Click"
On Error GoTo Err

If ApplyStyleHistCheck.Value = vbChecked Then setAllChartStyles ChartStylesCombo.Text, True
If ApplyStyleLiveCheck.Value = vbChecked Then setAllChartStyles ChartStylesCombo.Text, False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ApplyStyleHistCheck_Click()
Const ProcName As String = "ApplyStyleHistCheck_Click"
On Error GoTo Err

checkReadyToApplyStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ApplyStyleLiveCheck_Click()
Const ProcName As String = "ApplyStyleLiveCheck_Click"
On Error GoTo Err

checkReadyToApplyStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BlackThemeOption_Click()
Const ProcName As String = "BlackThemeOption_Click"
On Error GoTo Err

gMainForm.ApplyTheme "BLACK"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BlueThemeOption_Click()
Const ProcName As String = "BlueThemeOption_Click"
On Error GoTo Err

gMainForm.ApplyTheme "BLUE"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ChartStylesCombo_Change()
Const ProcName As String = "ChartStylesCombo_Change"
On Error GoTo Err

checkReadyToApplyStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ChartStylesCombo_Click()
Const ProcName As String = "ChartStylesCombo_Click"
On Error GoTo Err

checkReadyToApplyStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ConfigEditorButton_Click()
Const ProcName As String = "ConfigEditorButton_Click"
On Error GoTo Err

Dim lNewAppInstanceConfig As ConfigurationSection
Set lNewAppInstanceConfig = gShowConfigEditor(mConfigStore, mAppInstanceConfig, mTheme, gMainForm)

If lNewAppInstanceConfig Is Nothing Then
    SetupCurrentConfigCombo
    RaiseEvent ConfigsChanged
Else
    gMainForm.Shutdown
    gLoadMainForm lNewAppInstanceConfig
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub CurrentConfigCombo_Click()
Const ProcName As String = "CurrentConfigCombo_Click"
On Error GoTo Err

Dim lNewAppInstanceConfig As ConfigurationSection
Set lNewAppInstanceConfig = getAppInstanceConfig(mConfigStore, CurrentConfigCombo.SelectedItem.Key)

If lNewAppInstanceConfig Is mAppInstanceConfig Then Exit Sub

gMainForm.Shutdown
gLoadMainForm lNewAppInstanceConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesSSTab_Click(PreviousTab As Integer)
Const ProcName As String = "FeaturesSSTab_Click"
On Error GoTo Err

Select Case FeaturesSSTab.Tab
Case FeaturesTabIndexNumbers.FeaturesTabIndexConfig
    ConfigEditorButton.SetFocus
Case FeaturesTabIndexNumbers.FeaturesTabIndexHistoricalCharts
    HistContractSearch.SetFocus
Case FeaturesTabIndexNumbers.FeaturesTabIndexLiveCharts
    LiveChartTimeframeSelector.SetFocus
    If mTickerGrid.SelectedTickers.Count > 0 Then LiveChartButton1.Default = True
Case FeaturesTabIndexNumbers.FeaturesTabIndexTickers
    LiveContractSearch.SetFocus
    If mTickerGrid.SelectedTickers.Count > 0 Then LiveChartButton.Default = True
Case FeaturesTabIndexNumbers.FeaturesTabIndexTickfileReplay
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

Private Sub FeaturesTabStrip_Click()
Const ProcName As String = "FeaturesTabStrip_Click"
On Error GoTo Err

FeaturesSSTab.SetFocus
FeaturesSSTab.Tab = FeaturesTabStrip.SelectedItem.Index - 1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HidePicture_Click()
RaiseEvent Hide
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

Private Sub HistContractSearch_Cancelled()
RaiseEvent HistContractSearchCancelled
End Sub

Private Sub HistContractSearch_Cleared()
RaiseEvent HistContractSearchCleared
End Sub

Private Sub HistContractSearch_ContractsLoaded(ByVal pContracts As IContracts)
RaiseEvent HistContractsLoaded(pContracts)
End Sub

Private Sub HistContractSearch_NoContracts()
Const ProcName As String = "HistContractSearch_NoContracts"
On Error GoTo Err

gModelessMsgBox "No contracts found", vbExclamation, mTheme, "Attention"

Exit Sub

Err:
If Err.Number = 401 Then Exit Sub ' Can't show non-modal form when modal form is displayed
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveChartButton_Click()
Const ProcName As String = "LiveChartButton_Click"
On Error GoTo Err

LiveChartButton1_Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveChartButton1_Click()
Const ProcName As String = "LiveChartButton1_Click"
On Error GoTo Err

createCharts mTickerGrid.SelectedTickers

clearSelectedTickers

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

Private Sub LiveChartTimeframeSelector_Click()
Const ProcName As String = "LiveChartTimeframeSelector_Click"
On Error GoTo Err

setChartButtonTooltip

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveContractSearch_Action()
Const ProcName As String = "LiveContractSearch_Action"
On Error GoTo Err

Dim lPreferredRow As Long
lPreferredRow = CLng(LiveContractSearch.Cookie)

Dim lContract As IContract
For Each lContract In LiveContractSearch.SelectedContracts
    mTickerGrid.StartTickerFromContract lContract, lPreferredRow
    If lPreferredRow <> 0 Then lPreferredRow = lPreferredRow + 1
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveContractSearch_Cancelled()
RaiseEvent LiveContractSearchCancelled
End Sub

Private Sub LiveContractSearch_Cleared()
RaiseEvent LiveContractSearchCleared
End Sub

Private Sub LiveContractSearch_ContractsLoaded(ByVal pContracts As IContracts)
RaiseEvent LiveContractsLoaded(pContracts)
End Sub

Private Sub LiveContractSearch_NoContracts()
Const ProcName As String = "LiveContractSearch_NoContracts"
On Error GoTo Err

gModelessMsgBox "No contracts found", vbExclamation, mTheme, "Attention"

Exit Sub

Err:
If Err.Number = 401 Then Exit Sub ' Can't show non-modal form when modal form is displayed
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub MarketDepthButton_Click()
Const ProcName As String = "MarketDepthButton_Click"
On Error GoTo Err

Dim lTicker As Ticker
For Each lTicker In mTickerGrid.SelectedTickers
    showMarketDepthForm lTicker
Next

clearSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NativeThemeOption_Click()
Const ProcName As String = "NativeThemeOption_Click"
On Error GoTo Err

gMainForm.ApplyTheme "NATIVE"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub NumHistHistoryBarsText_Validate(Cancel As Boolean)
Const ProcName As String = "NumHistHistoryBarsText_Validate"
On Error GoTo Err

If Not IsInteger(NumHistHistoryBarsText.Text, 0, 2000) Then Cancel = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NumLiveHistoryBarsText_Validate(Cancel As Boolean)
Const ProcName As String = "NumLiveHistoryBarsText_Validate"
On Error GoTo Err

If Not IsInteger(NumLiveHistoryBarsText.Text, 0, 2000) Then Cancel = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderTicketButton_Click()
Const ProcName As String = "OrderTicketButton_Click"
On Error GoTo Err

showOrderTicket

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

Private Sub PinPicture_Click()
RaiseEvent Pin
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
                                                CInt(ReplaySpeedCombo.SelectedItem.Tag), _
                                                250)
    Dim lOrderManager As New OrderManager
    mInfoPanel.MonitorTickfilePositions lTickfileDataManager, lOrderManager.PositionManagersSimulated
    
    Set mReplayController = lTickfileDataManager.ReplayController
    
    Dim lTickers As Tickers
    Set lTickers = CreateTickers(lTickfileDataManager, mTradeBuildAPI.StudyLibraryManager, mTradeBuildAPI.HistoricalDataStoreInput, lOrderManager, , mTradeBuildAPI.OrderSubmitterFactorySimulated)
    
    Dim i As Long
    For i = 1 To TickfileOrganiser1.TickfileCount
        Dim lTicker As Ticker
        Set lTicker = lTickers.CreateTicker(mReplayController.TickStream(i - 1).ContractFuture, False)
        mTickerGrid.AddTickerFromDataSource lTicker
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
    mReplayController.ReplaySpeed = CInt(ReplaySpeedCombo.SelectedItem.Tag)
End If

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

Private Sub TickfileOrganiser1_TickfileCountChanged()
Const ProcName As String = "TickfileOrganiser1_TickfileCountChanged"
On Error GoTo Err

If TickfileOrganiser1.TickfileCount > 0 Then
    PlayTickFileButton.Enabled = True
Else
    PlayTickFileButton.Enabled = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UnpinPicture_Click()
RaiseEvent Unpin
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
ReplayProgressLabel.caption = ""

TickfileOrganiser1.Enabled = True

LogMessage "Tickfile replay completed"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mTickerGrid Event Handlers
'================================================================================

Private Sub mTickerGrid_TickerSelectionChanged()
Const ProcName As String = "mTickerGrid_TickerSelectionChanged"
On Error GoTo Err

handleSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

Dim lTicker As Ticker

Select Case ev.ChangeType
Case CollItemAdded
    Set lTicker = ev.AffectedItem
    lTicker.AddStateChangeListener Me
Case CollItemRemoved
    Set lTicker = ev.AffectedItem
    lTicker.RemoveStateChangeListener Me
Case CollItemChanged

Case CollOrderChanged

Case CollCollectionCleared

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get MinimumHeight() As Long
MinimumHeight = MinimumHeightTwips
End Property

Public Property Get MinimumWidth() As Long
MinimumWidth = MinimumWidthTwips
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property


If TypeOf mTheme Is BlackTheme Then
    BlackThemeOption.Value = True
ElseIf TypeOf mTheme Is BlueTheme Then
    BlueThemeOption.Value = True
ElseIf TypeOf mTheme Is NativeTheme Then
    NativeThemeOption.Value = True
End If

gApplyTheme mTheme, UserControl.Controls

PinPicture.BackColor = UserControl.BackColor
UnpinPicture.BackColor = UserControl.BackColor
HidePicture.BackColor = UserControl.BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub CancelHistContractSearch()
Const ProcName As String = "CancelHistContractSearch"
On Error GoTo Err

HistContractSearch.CancelSearch

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CancelLiveContractSearch()
Const ProcName As String = "CancelLiveContractSearch"
On Error GoTo Err

LiveContractSearch.CancelSearch

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearHistContractSearch()
Const ProcName As String = "ClearHistContractSearch"
On Error GoTo Err

HistContractSearch.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearLiveContractSearch()
Const ProcName As String = "ClearLiveContractSearch"
On Error GoTo Err

LiveContractSearch.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

LogMessage "Stopping tickfile replay"
' prevent event handler being fired on completion, which would
' reload the form again
Set mTickfileReplayTC = Nothing
stopTickfileReplay


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pPinned As Boolean, _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByVal pTickerGrid As TickerGrid, _
                ByVal pInfoPanel As InfoPanel, _
                ByVal pInfoPanelFloating As InfoPanel, _
                ByVal pChartForms As ChartForms, _
                ByVal pOrderTicket As fOrderTicket)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pPinned Then
    UnpinPicture.Visible = True
    PinPicture.Visible = False
Else
    UnpinPicture.Visible = False
    PinPicture.Visible = True
End If

Set mTradeBuildAPI = pTradeBuildAPI
Set mConfigStore = pConfigStore
Set mTickers = mTradeBuildAPI.Tickers
Set mAppInstanceConfig = pAppInstanceConfig
Set mTickerGrid = pTickerGrid
Set mInfoPanel = pInfoPanel
Set mInfoPanelFloating = pInfoPanelFloating
Set mChartForms = pChartForms
Set mOrderTicket = pOrderTicket

LogMessage "Initialising Features Panel: setting up contract search"
setupContractSearch

setupReplaySpeedCombo

LogMessage "Initialising Features Panel: setting up tickfile organiser"
setupTickfileOrganiser

LogMessage "Initialising Features Panel: setting up timeframeselectors"
setupTimeframeSelectors

LogMessage "Initialising Features Panel: setting current chart styles"
loadStyleComboItems ChartStylesCombo.ComboItems
setCurrentChartStyles

LogMessage "Initialising Features Panel: setting up date pickers"
FromDatePicker.Value = DateAdd("m", -1, Now)
FromDatePicker.Value = Empty    ' clear the checkbox
ToDatePicker.Value = Now

LogMessage "Initialising Features Panel: setting up current config combo"
SetupCurrentConfigCombo

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadHistContractsForUserChoice( _
                ByVal pContracts As IContracts)
Const ProcName As String = "LoadHistContractsForUserChoice"
On Error GoTo Err

HistContractSearch.LoadContracts pContracts, 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadLiveContractsForUserChoice( _
                ByVal pContracts As IContracts, _
                ByVal pPreferredTickerGridIndex)
Const ProcName As String = "LoadLiveContractsForUserChoice"
On Error GoTo Err

LiveContractSearch.LoadContracts pContracts, pPreferredTickerGridIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetupCurrentConfigCombo()
Const ProcName As String = "SetupCurrentConfigCombo"
On Error GoTo Err

CurrentConfigCombo.ComboItems.Clear

Dim lAppConfigs As ConfigurationSection
Set lAppConfigs = GetAppInstanceConfigs(mConfigStore)

Dim lAppConfig As ConfigurationSection
For Each lAppConfig In lAppConfigs
    If lAppConfig Is GetDefaultAppInstanceConfig(mConfigStore) Then
        CurrentConfigCombo.ComboItems.Add , lAppConfig.InstanceQualifier, "(Default) " & lAppConfig.InstanceQualifier
    Else
        CurrentConfigCombo.ComboItems.Add , lAppConfig.InstanceQualifier, lAppConfig.InstanceQualifier
    End If
Next

Set CurrentConfigCombo.SelectedItem = CurrentConfigCombo.ComboItems.Item(mAppInstanceConfig.InstanceQualifier)
CurrentConfigCombo.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowTickersPane()
Const ProcName As String = "ShowTickersPane"
On Error GoTo Err

If FeaturesSSTab.Tab <> FeaturesTabIndexNumbers.FeaturesTabIndexTickers Then FeaturesTabStrip.Tabs(FeaturesTabIndexNumbers.FeaturesTabIndexTickers + 1).Selected = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkReadyToApplyStyle()
Const ProcName As String = "checkReadyToApplyStyles"
On Error GoTo Err

ApplyStyleButton.Enabled = (ApplyStyleHistCheck.Value = vbChecked Or _
                            ApplyStyleLiveCheck.Value = vbChecked) And _
                            ChartStylesCombo.Text <> ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearSelectedTickers()
Const ProcName As String = "clearSelectedTickers"
On Error GoTo Err

mTickerGrid.DeselectSelectedTickers
handleSelectedTickers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createCharts(ByVal pTickers As Tickers)
Const ProcName As String = "createCharts"
On Error GoTo Err

mChartForms.AddAsync pTickers, _
                LiveChartTimeframeSelector.TimePeriod, _
                mTradeBuildAPI.BarFormatterLibManager, _
                mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator, _
                mAppInstanceConfig.AddConfigurationSection(ConfigSectionCharts), _
                CreateChartSpecifier(CLng(NumLiveHistoryBarsText.Text), Not (LiveSessionOnlyCheck = vbChecked)), _
                ChartStylesManager.Item(LiveChartStylesCombo.SelectedItem.Text), _
                gMainForm, _
                mTheme

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

mChartForms.AddHistoricAsync HistChartTimeframeSelector.TimePeriod, _
                    pContracts, _
                    mTradeBuildAPI.StudyLibraryManager, _
                    mTradeBuildAPI.HistoricalDataStoreInput, _
                    mTradeBuildAPI.BarFormatterLibManager, _
                    lConfig, _
                    CreateChartSpecifier(CLng(NumHistHistoryBarsText.Text), Not (HistSessionOnlyCheck = vbChecked), fromDate, toDate), _
                    ChartStylesManager.Item(HistChartStylesCombo.SelectedItem.Text), _
                    gMainForm, _
                    mTheme

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getSelectedDataSource() As IMarketDataSource
Const ProcName As String = "getSelectedDataSource"
On Error GoTo Err

If mTickerGrid.SelectedTickers.Count = 1 Then Set getSelectedDataSource = mTickerGrid.SelectedTickers.Item(1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleSelectedTickers()
Const ProcName As String = "handleSelectedTickers"
On Error GoTo Err

StopTickerButton.Enabled = False
LiveChartButton1.Enabled = False
LiveChartButton.Enabled = False
MarketDepthButton.Enabled = False
OrderTicketButton.Enabled = False

If mTickerGrid.SelectedTickers.Count = 0 Then Exit Sub
    
StopTickerButton.Enabled = True
LiveChartButton1.Enabled = True
LiveChartButton.Enabled = True
MarketDepthButton.Enabled = True

If FeaturesSSTab.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexLiveCharts Then
    LiveChartButton1.Default = True
ElseIf FeaturesSSTab.Tab = FeaturesTabIndexNumbers.FeaturesTabIndexTickers Then
    LiveChartButton.Default = True
End If

Dim lTicker As Ticker
Set lTicker = getSelectedDataSource

If lTicker Is Nothing Then Exit Sub

MarketDepthButton.Enabled = False
If lTicker.State <> MarketDataSourceStateRunning Then Exit Sub

Dim lContract As IContract
Set lContract = lTicker.ContractFuture.Value
If lContract.Specifier.SecType = SecTypeIndex Then Exit Sub

If lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled Then OrderTicketButton.Enabled = True
MarketDepthButton.Enabled = True

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

Private Sub setAllChartStyles(ByVal pStyleName As String, ByVal pHistorical As Boolean)
Const ProcName As String = "setAllChartStyles"
On Error GoTo Err

mChartForms.SetStyle ChartStylesManager.Item(pStyleName), pHistorical

If pHistorical Then
    HistChartStylesCombo.Text = pStyleName
Else
    LiveChartStylesCombo.Text = pStyleName
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setChartButtonTooltip()
Const ProcName As String = "setChartButtonTooltip"
On Error GoTo Err

Dim tp As TimePeriod
Set tp = LiveChartTimeframeSelector.TimePeriod

LiveChartButton1.ToolTipText = "Show " & tp.ToString & " chart"
LiveChartButton.ToolTipText = LiveChartButton1.ToolTipText

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

Private Sub setupReplaySpeedCombo()
Const ProcName As String = "setupReplaySpeedCombo"
On Error GoTo Err

ReplaySpeedCombo.ComboItems.Add , , "Continuous"
ReplaySpeedCombo.ComboItems(1).Tag = 0

ReplaySpeedCombo.ComboItems.Add , , "Actual speed"
ReplaySpeedCombo.ComboItems(2).Tag = 1
ReplaySpeedCombo.ComboItems(2).Selected = True

ReplaySpeedCombo.ComboItems.Add , , "2x Actual speed"
ReplaySpeedCombo.ComboItems(3).Tag = 2

ReplaySpeedCombo.ComboItems.Add , , "4x Actual speed"
ReplaySpeedCombo.ComboItems(4).Tag = 4

ReplaySpeedCombo.ComboItems.Add , , "8x Actual speed"
ReplaySpeedCombo.ComboItems(5).Tag = 8

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
HistChartTimeframeSelector.Initialise mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator
HistChartTimeframeSelector.SelectTimeframe GetTimePeriod(5, TimePeriodMinute)

setChartButtonTooltip

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showMarketDepthForm(ByVal pTicker As Ticker)
Const ProcName As String = "showMarketDepthForm"
On Error GoTo Err

If Not pTicker.State = MarketDataSourceStateRunning Then Exit Sub

Dim lContract As IContract
Set lContract = pTicker.ContractFuture.Value
If lContract.Specifier.SecType = SecTypeIndex Then Exit Sub

Dim mktDepthForm As New fMarketDepth
mktDepthForm.NumberOfRows = 100
mktDepthForm.Ticker = pTicker

If Not mTheme Is Nothing Then mktDepthForm.Theme = mTheme
mktDepthForm.Show vbModeless, gMainForm

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showOrderTicket()
Const ProcName As String = "showOrderTicket"
On Error GoTo Err

If getSelectedDataSource Is Nothing Then
    gModelessMsgBox "No ticker selected - please select a ticker", vbExclamation, mTheme, "Error"
Else
    mOrderTicket.Show vbModeless, gMainForm
    mOrderTicket.Ticker = getSelectedDataSource
End If

Exit Sub

Err:
If Err.Number = 401 Then Exit Sub ' Can't show non-modal form when modal form is displayed
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub StopSelectedTickers()
Const ProcName As String = "StopSelectedTickers"
On Error GoTo Err

Dim lTickers As SelectedTickers
Set lTickers = mTickerGrid.SelectedTickers

mTickerGrid.StopSelectedTickers

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
LiveChartButton1.Enabled = False
LiveChartButton.Enabled = False
If Not mReplayController Is Nothing Then
    mReplayController.StopReplay
    Set mReplayController = Nothing
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub



