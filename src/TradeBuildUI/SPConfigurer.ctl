VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#27.1#0"; "TWControls40.ocx"
Begin VB.UserControl SPConfigurer 
   BackStyle       =   0  'Transparent
   ClientHeight    =   12750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16740
   DefaultCancel   =   -1  'True
   ScaleHeight     =   12750
   ScaleWidth      =   16740
   Begin TWControls40.TWButton ApplyButton 
      Height          =   375
      Left            =   6360
      TabIndex        =   67
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Caption         =   "Apply"
   End
   Begin TWControls40.TWButton CancelButton 
      Height          =   375
      Left            =   5280
      TabIndex        =   66
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Caption         =   "Cancel"
   End
   Begin VB.PictureBox TfInputOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5040
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   59
      Top             =   9000
      Width           =   4815
      Begin TWControls40.TWButton InputPathChooserButton 
         Height          =   375
         Left            =   4320
         TabIndex        =   65
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin VB.TextBox InputTickfilePathText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   63
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox TfInputEnabledCheck 
         Appearance      =   0  'Flat
         Caption         =   "Enabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   "Input path"
         Height          =   375
         Left            =   0
         TabIndex        =   62
         Top             =   360
         Width           =   975
      End
   End
   Begin TWControls40.TWImageCombo OptionCombo 
      Height          =   270
      Left            =   4310
      TabIndex        =   0
      Top             =   710
      Width           =   3015
      _ExtentX        =   5318
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
      MouseIcon       =   "SPConfigurer.ctx":0000
      Text            =   ""
   End
   Begin VB.PictureBox CustomOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   10080
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   49
      Top             =   4080
      Width           =   4815
      Begin MSDataGridLib.DataGrid ParamsGrid 
         Height          =   1455
         Left            =   960
         TabIndex        =   52
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
         RowDividerStyle =   6
         AllowAddNew     =   -1  'True
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
      Begin VB.TextBox ProgIdText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   360
         Width           =   3855
      End
      Begin VB.CheckBox CustomEnabledCheck 
         Appearance      =   0  'Flat
         Caption         =   "Enabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Parameters"
         Height          =   255
         Left            =   0
         TabIndex        =   51
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Prog ID"
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox TfOutputOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5040
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   48
      Top             =   6480
      Width           =   4815
      Begin TWControls40.TWButton OutputPathChooserButton 
         Height          =   375
         Left            =   4320
         TabIndex        =   64
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin TWControls40.TWImageCombo TickfileGranularityCombo 
         Height          =   270
         Left            =   960
         TabIndex        =   26
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
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
         MouseIcon       =   "SPConfigurer.ctx":001C
         Text            =   ""
      End
      Begin VB.TextBox OutputTickfilePathText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox TfOutputEnabledCheck 
         Appearance      =   0  'Flat
         Caption         =   "Enabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label23 
         Caption         =   "Granularity"
         Height          =   255
         Left            =   0
         TabIndex        =   61
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Output path"
         Height          =   375
         Left            =   0
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.PictureBox BrOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   10080
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   47
      Top             =   6480
      Width           =   4815
      Begin VB.CheckBox BrEnabledCheck 
         Appearance      =   0  'Flat
         Caption         =   "Enabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox QtOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5040
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   42
      Top             =   4080
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox QtConnectRetryIntervalText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Text            =   "10"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox QtProviderKeyText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox QtKeepConnectionCheck 
         Appearance      =   0  'Flat
         Caption         =   "Keep connection"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox QtEnabledCheck 
         Appearance      =   0  'Flat
         Caption         =   "Enabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox QtPasswordText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox QtPortText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Text            =   "16240"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox QtServerText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Connection retry interval"
         Height          =   375
         Left            =   0
         TabIndex        =   57
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Provider key"
         Height          =   255
         Left            =   0
         TabIndex        =   56
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Password"
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Port"
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Server"
         Height          =   255
         Left            =   0
         TabIndex        =   43
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox DbOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   36
      Top             =   6480
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CheckBox DbUseAsyncWritesCheck 
         Appearance      =   0  'Flat
         Caption         =   "Use async writes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox DBUseAsyncReadsCheck 
         Appearance      =   0  'Flat
         Caption         =   "Use async reads"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin TWControls40.TWImageCombo DbTypeCombo 
         Height          =   270
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
         MouseIcon       =   "SPConfigurer.ctx":0038
         Text            =   ""
      End
      Begin VB.CheckBox DbEnabledCheck 
         Appearance      =   0  'Flat
         Caption         =   "Enabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox DbDatabaseText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox DbServerText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox DbUsernameText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox DbPasswordText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Database"
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Server"
         Height          =   255
         Left            =   0
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "DB Type"
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Username"
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Password"
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.PictureBox TwsOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   32
      Top             =   4080
      Visible         =   0   'False
      Width           =   4815
      Begin TWControls40.TWImageCombo TwsLogLevelCombo 
         Height          =   270
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         MouseIcon       =   "SPConfigurer.ctx":0054
         Text            =   ""
      End
      Begin VB.CheckBox TwsKeepConnectionCheck 
         Appearance      =   0  'Flat
         Caption         =   "Keep connection"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox TwsConnectRetryIntervalText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "10"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox TwsProviderKeyText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox TwsEnabledCheck 
         Appearance      =   0  'Flat
         Caption         =   "Enabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox TWSClientIdText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Text            =   "-1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TWSPortText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "7496"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TWSServerText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "TWS Log Level"
         Height          =   375
         Left            =   2520
         TabIndex        =   55
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Connection retry interval"
         Height          =   375
         Left            =   0
         TabIndex        =   54
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Provider key"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Client id"
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Port"
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Server"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ListBox CategoryList 
      Appearance      =   0  'Flat
      Height          =   3735
      ItemData        =   "SPConfigurer.ctx":0070
      Left            =   120
      List            =   "SPConfigurer.ctx":0072
      TabIndex        =   30
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label CategoryLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FE8100&
      Height          =   255
      Left            =   2400
      TabIndex        =   31
      Top             =   240
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E7D395&
      X1              =   2520
      X2              =   7320
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Shape OptionsBox 
      Height          =   2175
      Left            =   2520
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label OptionLabel 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2280
      TabIndex        =   46
      Top             =   720
      Width           =   1935
   End
   Begin VB.Shape OutlineBox 
      Height          =   4000
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E7D395&
      BorderWidth     =   2
      FillColor       =   &H80000005&
      Height          =   495
      Left            =   2280
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "SPConfigurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements CollectionChangeListener
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "SPConfigurer1"

Private Const RoleInput                     As String = "INPUT"
Private Const RoleOutput                    As String = "OUTPUT"

Private Const AttributeNameServiceProviderEnabled As String = "Enabled"
Private Const AttributeNameServiceProviderProgId As String = "ProgId"

Private Const ConfigNameProperty            As String = "Property"
Private Const ConfigNameProperties          As String = "Properties"
Private Const ConfigNameServiceProvider     As String = "ServiceProvider"
Private Const ConfigNameServiceProviders    As String = "ServiceProviders"

Private Const DbTypeMySql                   As String = "MySQL5"
Private Const DbTypeSqlServer7              As String = "SQL Server 7"
Private Const DbTypeSqlServer2000           As String = "SQL Server 2000"
Private Const DbTypeSqlServer2005           As String = "SQL Server 2005"

Private Const PropertyNameQtServer          As String = "Server"
Private Const PropertyNameQtPort            As String = "Port"
Private Const PropertyNameQtPassword        As String = "Password"
Private Const PropertyNameQtKeepConnection  As String = "Keep Connection"
Private Const PropertyNameQtConnectionRetryInterval As String = "Connection Retry Interval Secs"
Private Const PropertyNameQtProviderKey     As String = "Provider Key"

Private Const PropertyNameTbServer          As String = "Server"
Private Const PropertyNameTbDbType          As String = "Database Type"
Private Const PropertyNameTbDbName          As String = "Database Name"
Private Const PropertyNameTbUserName        As String = "User Name"
Private Const PropertyNameTbPassword        As String = "Password"
Private Const PropertyNameTbRole            As String = "Role"
Private Const PropertyNameTbUseSynchronousWrites As String = "Use Synchronous Writes"
Private Const PropertyNameTbUseSynchronousReads As String = "Use Synchronous Reads"

Private Const PropertyNameTfRole            As String = "Role"
Private Const PropertyNameTfTickfilePath    As String = "Tickfile Path"
Private Const PropertyNameTfTickfileGranularity As String = "Tickfile Granularity"

Private Const PropertyNameTwsServer         As String = "Server"
Private Const PropertyNameTwsPort           As String = "Port"
Private Const PropertyNameTwsClientId       As String = "Client Id"
Private Const PropertyNameTwsKeepConnection As String = "Keep Connection"
Private Const PropertyNameTwsConnectionRetryInterval    As String = "Connection Retry Interval Secs"
Private Const PropertyNameTwsProviderKey    As String = "Provider Key"
Private Const PropertyNameTwsLogLevel       As String = "TWS Log Level"

Private Const RolePrimary                   As String = "Primary"
Private Const RoleSecondary                 As String = "Secondary"

Private Const SpOptionCustomBarData         As String = "Custom"
Private Const SpOptionCustomContractData    As String = "Custom"
Private Const SpOptionCustomOrders          As String = "Custom"
Private Const SpOptionCustomRealtimeData    As String = "Custom"
Private Const SpOptionCustomTickData        As String = "Custom"

'Private Const SpOptionQtBarData             As String = "QuoteTracker"
'Private Const SpOptionQtRealtimeData        As String = "QuoteTracker"
'Private Const SpOptionQtTickData            As String = "QuoteTracker"

Private Const SpOptionTbBarData             As String = "TradeBuild Database"
Private Const SpOptionTbContractData        As String = "TradeBuild Database"
Private Const SpOptionTbOrders              As String = "TradeBuild Exchange Simulator"
Private Const SpOptionTbTickData            As String = "TradeBuild Database"

Private Const SpOptionFileTickData          As String = "Tickfiles"

Private Const SpOptionNotConfigured         As String = "(not configured or invalid)"

Private Const SpOptionTwsBarData            As String = "TWS"
Private Const SpOptionTwsContractData       As String = "TWS"
Private Const SpOptionTwsOrders             As String = "IB (via TWS)"
Private Const SpOptionTwsRealtimeData       As String = "TWS"

Private Const TickfileGranularityDay        As String = "File per day"
Private Const TickfileGranularityExecution  As String = "File per execution"
Private Const TickfileGranularitySession    As String = "File per trading session"
Private Const TickfileGranularityWeek       As String = "File per week"

Private Const TWSLogLevelDetail             As String = "Detail"
Private Const TWSLogLevelError              As String = "Error"
Private Const TWSLogLevelInformation        As String = "Information"
Private Const TWSLogLevelSystem             As String = "System"
Private Const TWSLogLevelWarning            As String = "Warning"

'@================================================================================
' Member variables
'@================================================================================

Private mCurrOptionsPic             As PictureBox

Private mConfig                     As ConfigurationSection

Private mCurrSPsList                As ConfigurationSection
Private mCurrSP                     As ConfigurationSection
Private mCurrProps                  As ConfigurationSection
Private mCurrCategory               As String
Private mCurrSpOption               As String

Private mCustomParams               As Parameters

Private mReadOnly                   As Boolean

Private mPermittedServiceProviderRoles As ServiceProviderRoles

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height

DbTypeCombo.ComboItems.Add , , ""
DbTypeCombo.ComboItems.Add , , DbTypeMySql
DbTypeCombo.ComboItems.Add , , DbTypeSqlServer7
DbTypeCombo.ComboItems.Add , , DbTypeSqlServer2000
DbTypeCombo.ComboItems.Add , , DbTypeSqlServer2005

TwsLogLevelCombo.ComboItems.Add , , TWSLogLevelDetail
TwsLogLevelCombo.ComboItems.Add , , TWSLogLevelError
TwsLogLevelCombo.ComboItems.Add , , TWSLogLevelInformation
TwsLogLevelCombo.ComboItems.Add , , TWSLogLevelSystem
TwsLogLevelCombo.ComboItems.Add , , TWSLogLevelWarning
TwsLogLevelCombo.Text = TWSLogLevelSystem

TickfileGranularityCombo.ComboItems.Add , , TickfileGranularitySession
TickfileGranularityCombo.ComboItems.Add , , TickfileGranularityDay
TickfileGranularityCombo.ComboItems.Add , , TickfileGranularityWeek
TickfileGranularityCombo.ComboItems.Add , , TickfileGranularityExecution

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_InitProperties()
UserControl.BackColor = UserControl.Ambient.BackColor
UserControl.ForeColor = UserControl.Ambient.ForeColor
End Sub

Private Sub UserControl_LostFocus()
Const ProcName As String = "UserControl_LostFocus"
On Error GoTo Err

checkForOutstandingUpdates

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.BackColor = UserControl.Ambient.BackColor
UserControl.ForeColor = UserControl.Ambient.ForeColor
End Sub

Private Sub UserControl_Resize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change( _
                ev As CollectionChangeEventData)
Const ProcName As String = "CollectionChangeListener_Change"
On Error GoTo Err

If ev.Source Is mCustomParams Then
    enableApplyButton True
    enableCancelButton True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ApplyButton_Click()
Const ProcName As String = "ApplyButton_Click"
On Error GoTo Err

applyProperties
enableApplyButton False
enableCancelButton False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BrEnabledCheck_Click()
Const ProcName As String = "BrEnabledCheck_Click"
On Error GoTo Err

enableApplyButton True
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"
On Error GoTo Err

enableApplyButton False
enableCancelButton False
Dim index As Long
index = CategoryList.ListIndex
CategoryList.ListIndex = -1
CategoryList.ListIndex = index

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CategoryList_Click()
Const ProcName As String = "CategoryList_Click"
On Error GoTo Err

If CategoryList.ListIndex = -1 Then
    Set mCurrSP = Nothing
    Set mCurrProps = Nothing
    mCurrCategory = ""
    mCurrSpOption = ""
    Exit Sub
End If

checkForOutstandingUpdates

hideSpOptions

mCurrCategory = CategoryList.Text

Select Case mCurrCategory
Case SPNameRealtimeData
    setupRealtimeDataSP
Case SPNameContractDataPrimary
    setupPrimaryContractDataSP
Case SPNameContractDataSecondary
    setupSecondaryContractDataSP
Case SPNameHistoricalDataInput
    setupHistoricalDataInputSP
Case SPNameHistoricalDataOutput
    setupHistoricalDataOutputSP
Case SPNameOrderSubmissionLive
    setupBrokerLiveSP
Case SPNameOrderSubmissionSimulated
    setupBrokerSimulatedSP
Case SPNameTickfileInput
    setupTickfileInputSP
Case SPNameTickfileOutput
    setupTickfileOutputSP
End Select

showSpOptions

enableApplyButton False
enableCancelButton False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DbDatabaseText_Change()
Const ProcName As String = "DbDatabaseText_Change"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DbEnabledCheck_Click()
Const ProcName As String = "DbEnabledCheck_Click"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DbPasswordText_Change()
Const ProcName As String = "DbPasswordText_Change"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DbServerText_Change()
Const ProcName As String = "DbServerText_Change"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DbTypeCombo_Click()
Const ProcName As String = "DbTypeCombo_Click"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DBUseAsyncReadsCheck_Click()
Const ProcName As String = "DBUseAsyncReadsCheck_Click"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DbUseAsyncWritesCheck_Click()
Const ProcName As String = "DbUseAsyncWritesCheck_Click"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub DbUsernameText_Change()
Const ProcName As String = "DbUsernameText_Change"
On Error GoTo Err

enableApplyButton isValidDbProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub InputPathChooserButton_Click()
Const ProcName As String = "InputPathChooserButton_Click"
On Error GoTo Err

InputTickfilePathText.Text = ChoosePath(InputTickfilePathText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub InputTickfilePathText_Change()
Const ProcName As String = "InputTickfilePathText_Change"
On Error GoTo Err

enableApplyButton True
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OptionCombo_Click()
Const ProcName As String = "OptionCombo_Click"
On Error GoTo Err

hideSpOptions
If OptionCombo.Text = SpOptionNotConfigured Then
    If Not mCurrSP Is Nothing Then enableApplyButton True
Else
    showSpOptions
    If mCurrSP Is Nothing Or OptionCombo.Text <> mCurrSpOption Then
        If mCurrOptionsPic Is DbOptionsPicture Then
            enableApplyButton isValidDbProperties
        ElseIf mCurrOptionsPic Is QtOptionsPicture Then
            enableApplyButton isValidQtProperties
        ElseIf mCurrOptionsPic Is TwsOptionsPicture Then
            enableApplyButton isValidTwsProperties
        ElseIf mCurrOptionsPic Is BrOptionsPicture Then
            enableApplyButton True
        ElseIf mCurrOptionsPic Is TfOutputOptionsPicture Then
            enableApplyButton True
        ElseIf mCurrOptionsPic Is TfInputOptionsPicture Then
            enableApplyButton True
        ElseIf mCurrOptionsPic Is CustomOptionsPicture Then
            enableApplyButton isValidCustomProperties
        End If
    End If
End If
mCurrSpOption = OptionCombo.Text
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OutputPathChooserButton_Click()
Const ProcName As String = "OutputPathChooserButton_Click"
On Error GoTo Err

OutputTickfilePathText.Text = ChoosePath(OutputTickfilePathText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OutputTickfilePathText_Change()
Const ProcName As String = "OutputTickfilePathText_Change"
On Error GoTo Err

enableApplyButton True
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ProgIdText_Change()
Const ProcName As String = "ProgIdText_Change"
On Error GoTo Err

enableApplyButton isValidCustomProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub QtEnabledCheck_Click()
Const ProcName As String = "QtEnabledCheck_Click"
On Error GoTo Err

enableApplyButton isValidQtProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub QtPasswordText_Change()
Const ProcName As String = "QtPasswordText_Change"
On Error GoTo Err

enableApplyButton isValidQtProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub QtPortText_Change()
Const ProcName As String = "QtPortText_Change"
On Error GoTo Err

enableApplyButton isValidQtProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub QtServerText_Change()
Const ProcName As String = "QtServerText_Change"
On Error GoTo Err

enableApplyButton isValidQtProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TfInputEnabledCheck_Click()
Const ProcName As String = "TfInputEnabledCheck_Click"
On Error GoTo Err

enableApplyButton True
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TfOutputEnabledCheck_Click()
Const ProcName As String = "TfOutputEnabledCheck_Click"
On Error GoTo Err

enableApplyButton True
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickfileGranularityCombo_Click()
Const ProcName As String = "TickfileGranularityCombo_Change"
On Error GoTo Err

enableApplyButton True
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWSClientIdText_Change()
Const ProcName As String = "TWSClientIdText_Change"
On Error GoTo Err

enableApplyButton isValidTwsProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TwsConnectRetryIntervalText_Change()
Const ProcName As String = "TwsConnectRetryIntervalText_Change"
On Error GoTo Err

enableApplyButton isValidTwsProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TwsEnabledCheck_Click()
Const ProcName As String = "TwsEnabledCheck_Click"
On Error GoTo Err

enableApplyButton isValidTwsProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWSPortText_Change()
Const ProcName As String = "TWSPortText_Change"
On Error GoTo Err

enableApplyButton isValidTwsProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TwsKeepConnectionCheck_Click()
Const ProcName As String = "TwsKeepConnectionCheck_Click"
On Error GoTo Err

enableApplyButton isValidTwsProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TwsProviderKeyText_Change()
Const ProcName As String = "TwsProviderKeyText_Change"
On Error GoTo Err

enableApplyButton isValidTwsProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWSServerText_Change()
Const ProcName As String = "TWSServerText_Change"
On Error GoTo Err

enableApplyButton isValidTwsProperties
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Dirty() As Boolean
Const ProcName As String = "Dirty"
On Error GoTo Err

Dirty = ApplyButton.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If Value Is Nothing Then Exit Property

Set mTheme = Value
UserControl.BackColor = mTheme.BackColor
gApplyTheme mTheme, UserControl.Controls

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

Public Sub ApplyChanges()
Const ProcName As String = "ApplyChanges"
On Error GoTo Err

applyProperties
enableApplyButton False
enableCancelButton False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mCustomParams Is Nothing Then
    Set ParamsGrid.DataSource = Nothing
    mCustomParams.RemoveCollectionChangeListener Me
End If

Set mCurrOptionsPic = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pPermittedServiceProviderRoles As ServiceProviderRoles, _
                ByVal pConfigdata As ConfigurationSection, _
                Optional ByVal pReadonly As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mPermittedServiceProviderRoles = pPermittedServiceProviderRoles

mReadOnly = pReadonly

checkForOutstandingUpdates

Set mCurrSPsList = Nothing
Set mCurrSP = Nothing
Set mCurrProps = Nothing
mCurrCategory = ""
mCurrSpOption = ""

loadConfig pConfigdata

If mReadOnly Then disableControls

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub applyBrProperties()
Const ProcName As String = "applyBrProperties"
On Error GoTo Err

If BrEnabledCheck = vbChecked Then
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "False"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyCustomProperties()
Const ProcName As String = "applyCustomProperties"
On Error GoTo Err

If CustomEnabledCheck = vbChecked Then
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "False"
End If

Dim param As Parameter
For Each param In mCustomParams
    setProperty mCurrProps, param.name, param.Value
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyDbProperties()
Const ProcName As String = "applyDbProperties"
On Error GoTo Err

If DbEnabledCheck = vbChecked Then
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "False"
End If
setProperty mCurrProps, PropertyNameTbServer, DbServerText
setProperty mCurrProps, PropertyNameTbDbType, DbTypeCombo
setProperty mCurrProps, PropertyNameTbDbName, DbDatabaseText
setProperty mCurrProps, PropertyNameTbUserName, DbUsernameText
setProperty mCurrProps, PropertyNameTbPassword, DbPasswordText
setProperty mCurrProps, PropertyNameTbUseSynchronousReads, CStr(DBUseAsyncReadsCheck.Value = vbUnchecked)
setProperty mCurrProps, PropertyNameTbUseSynchronousWrites, CStr(DbUseAsyncWritesCheck.Value = vbUnchecked)

If mCurrCategory = SPNameHistoricalDataInput Or _
    mCurrCategory = SPNameTickfileInput _
Then
    setProperty mCurrProps, PropertyNameTbRole, RoleInput
End If

If mCurrCategory = SPNameHistoricalDataOutput Or _
    mCurrCategory = SPNameTickfileOutput _
Then
    setProperty mCurrProps, PropertyNameTbRole, RoleOutput
End If

If mCurrCategory = SPNameContractDataPrimary Then
    setProperty mCurrProps, PropertyNameTbRole, RolePrimary
End If

If mCurrCategory = SPNameContractDataSecondary Then
    setProperty mCurrProps, PropertyNameTbRole, RoleSecondary
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyProperties()
Const ProcName As String = "applyProperties"
On Error GoTo Err

If mCurrSP Is Nothing Then
    createNewSp
End If

If OptionCombo.Text = SpOptionNotConfigured Then
    deleteSp
    hideSpOptions
    Exit Sub
End If

clearProperties

setProgId

If mCurrOptionsPic Is BrOptionsPicture Then
    applyBrProperties
ElseIf mCurrOptionsPic Is CustomOptionsPicture Then
    applyCustomProperties
ElseIf mCurrOptionsPic Is DbOptionsPicture Then
    applyDbProperties
ElseIf mCurrOptionsPic Is QtOptionsPicture Then
    applyQtProperties
ElseIf mCurrOptionsPic Is TfOutputOptionsPicture Then
    applyTfOutputProperties
ElseIf mCurrOptionsPic Is TfInputOptionsPicture Then
    applyTfInputProperties
ElseIf mCurrOptionsPic Is TwsOptionsPicture Then
    applyTwsProperties
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyQtProperties()
Const ProcName As String = "applyQtProperties"
On Error GoTo Err

If QtEnabledCheck = vbChecked Then
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "False"
End If
If QtKeepConnectionCheck = vbChecked Then
    setProperty mCurrProps, PropertyNameQtKeepConnection, "True"
Else
    setProperty mCurrProps, PropertyNameQtKeepConnection, "False"
End If
setProperty mCurrProps, PropertyNameQtServer, QtServerText
setProperty mCurrProps, PropertyNameQtPort, QtPortText
setProperty mCurrProps, PropertyNameQtPassword, QtPasswordText
setProperty mCurrProps, PropertyNameQtProviderKey, QtProviderKeyText
setProperty mCurrProps, PropertyNameQtConnectionRetryInterval, QtConnectRetryIntervalText

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyTfInputProperties()
Const ProcName As String = "applyTfInputProperties"
On Error GoTo Err

If TfInputEnabledCheck = vbChecked Then
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "False"
End If

setProperty mCurrProps, PropertyNameTfRole, RoleInput
setProperty mCurrProps, PropertyNameTfTickfilePath, InputTickfilePathText

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyTfOutputProperties()
Const ProcName As String = "applyTfOutputProperties"
On Error GoTo Err

If TfOutputEnabledCheck = vbChecked Then
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "False"
End If

setProperty mCurrProps, PropertyNameTfRole, RoleOutput
setProperty mCurrProps, PropertyNameTfTickfilePath, OutputTickfilePathText
setProperty mCurrProps, PropertyNameTfTickfileGranularity, TickfileGranularityCombo.Text

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyTwsProperties()
Const ProcName As String = "applyTwsProperties"
On Error GoTo Err

If TwsEnabledCheck = vbChecked Then
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.SetAttribute AttributeNameServiceProviderEnabled, "False"
End If
If TwsKeepConnectionCheck = vbChecked Then
    setProperty mCurrProps, PropertyNameTwsKeepConnection, "True"
Else
    setProperty mCurrProps, PropertyNameTwsKeepConnection, "False"
End If
setProperty mCurrProps, PropertyNameTwsServer, TWSServerText
setProperty mCurrProps, PropertyNameTwsPort, TWSPortText
setProperty mCurrProps, PropertyNameTwsClientId, TWSClientIdText
setProperty mCurrProps, PropertyNameTwsProviderKey, TwsProviderKeyText
setProperty mCurrProps, PropertyNameTwsConnectionRetryInterval, TwsConnectRetryIntervalText
setProperty mCurrProps, PropertyNameTwsLogLevel, TwsLogLevelCombo

If mCurrCategory = SPNameContractDataPrimary Then
    setProperty mCurrProps, PropertyNameTbRole, RolePrimary
End If

If mCurrCategory = SPNameContractDataSecondary Then
    setProperty mCurrProps, PropertyNameTbRole, RoleSecondary
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub checkForOutstandingUpdates()
Const ProcName As String = "checkForOutstandingUpdates"
On Error GoTo Err

If ApplyButton.Enabled Then
    If MsgBox("Do you want to apply the changes you have made?", _
            vbExclamation Or vbYesNo) = vbYes Then
        applyProperties
        enableApplyButton False
        enableCancelButton False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearProperties()
Const ProcName As String = "clearProperties"
On Error GoTo Err

mCurrProps.RemoveAllChildren

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createNewSp()
Const ProcName As String = "createNewSp"
On Error GoTo Err

If mCurrSPsList Is Nothing Then
    Set mCurrSPsList = mConfig.AddConfigurationSection(ConfigNameServiceProviders)
End If

Set mCurrSP = mCurrSPsList.AddConfigurationSection(ConfigNameServiceProvider & "(" & mCurrCategory & ")")
Set mCurrProps = mCurrSP.AddConfigurationSection(ConfigNameProperties)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteSp()
Const ProcName As String = "deleteSp"
On Error GoTo Err

mCurrSPsList.RemoveConfigurationSection ConfigNameServiceProvider & "(" & mCurrSP.InstanceQualifier & ")"
Set mCurrSP = Nothing
Set mCurrProps = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableControls()
Const ProcName As String = "disableControls"
On Error GoTo Err

CancelButton.Enabled = False
ApplyButton.Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableApplyButton( _
                ByVal enable As Boolean)
Const ProcName As String = "enableApplyButton"
On Error GoTo Err

If mReadOnly Then Exit Sub
If enable Then
    ApplyButton.Enabled = True
    ApplyButton.Default = True
Else
    ApplyButton.Enabled = False
    ApplyButton.Default = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableCancelButton( _
                ByVal enable As Boolean)
Const ProcName As String = "enableCancelButton"
On Error GoTo Err

If mReadOnly Then Exit Sub
If enable Then
    CancelButton.Enabled = True
    CancelButton.Cancel = True
Else
    CancelButton.Enabled = False
    CancelButton.Cancel = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findSp( _
                ByVal name As String) As Boolean
Const ProcName As String = "findSp"
On Error GoTo Err

Set mCurrSP = Nothing
Set mCurrProps = Nothing
mCurrSpOption = ""

If mCurrSPsList Is Nothing Then Exit Function

Set mCurrSP = mCurrSPsList.GetConfigurationSection(ConfigNameServiceProvider & "(" & name & ")")
If Not mCurrSP Is Nothing Then
    Set mCurrProps = mCurrSP.GetConfigurationSection(ConfigNameProperties)
    findSp = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getProperty( _
                ByVal name As String) As String
Const ProcName As String = "getProperty"
On Error GoTo Err

On Error Resume Next
getProperty = mCurrProps.GetSetting("." & ConfigNameProperty & "(" & name & ")")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub hideSpOptions()
Const ProcName As String = "hideSpOptions"
On Error GoTo Err

If Not mCurrOptionsPic Is Nothing Then
    mCurrOptionsPic.Visible = False
    Set mCurrOptionsPic = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isValidCustomProperties() As Boolean
Const ProcName As String = "isValidCustomProperties"
On Error GoTo Err

If ProgIdText = "" Then
ElseIf InStr(1, ProgIdText, ".") < 2 Then
ElseIf InStr(1, ProgIdText, ".") = Len(ProgIdText) Then
ElseIf Len(ProgIdText) > 39 Then
Else
    isValidCustomProperties = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidDbProperties() As Boolean
Const ProcName As String = "isValidDbProperties"
On Error GoTo Err

If DbDatabaseText = "" Then
ElseIf DbTypeCombo.Text = "" Then
ElseIf DbTypeCombo.Text = DbTypeMySql And DbUsernameText = "" Then
Else
    isValidDbProperties = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidQtProperties() As Boolean
Const ProcName As String = "isValidQtProperties"
On Error GoTo Err

If Not IsInteger(QtPortText, 1024) Then
Else
    isValidQtProperties = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidTwsProperties() As Boolean
Const ProcName As String = "isValidTwsProperties"
On Error GoTo Err

If Not IsInteger(TWSPortText, 1) Then
ElseIf Not IsInteger(TWSClientIdText) Then
ElseIf TwsConnectRetryIntervalText <> "" And Not IsInteger(TwsConnectRetryIntervalText, 0) Then
Else
    isValidTwsProperties = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadConfig( _
                ByVal configdata As ConfigurationSection)
Const ProcName As String = "loadConfig"
On Error GoTo Err

Set mConfig = configdata

On Error Resume Next
Set mCurrSPsList = mConfig.GetConfigurationSection(ConfigNameServiceProviders)
On Error GoTo Err

CategoryList.Clear

If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleRealtimeData Then
    CategoryList.AddItem SPNameRealtimeData
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleContractDataPrimary Then
    CategoryList.AddItem SPNameContractDataPrimary
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleContractDataSecondary Then
    CategoryList.AddItem SPNameContractDataSecondary
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleHistoricalDataInput Then
    CategoryList.AddItem SPNameHistoricalDataInput
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleHistoricalDataOutput Then
    CategoryList.AddItem SPNameHistoricalDataOutput
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderSubmissionLive Then
    CategoryList.AddItem SPNameOrderSubmissionLive
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderSubmissionSimulated Then
    CategoryList.AddItem SPNameOrderSubmissionSimulated
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleTickfileInput Then
    CategoryList.AddItem SPNameTickfileInput
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleTickfileOutput Then
    CategoryList.AddItem SPNameTickfileOutput
End If

If CategoryList.ListCount > 0 Then CategoryList.ListIndex = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setProgId()
Const ProcName As String = "setProgId"
On Error GoTo Err

If CategoryList.ListIndex = -1 Then Exit Sub

Dim progId As String

Select Case mCurrCategory
Case SPNameRealtimeData
    Select Case OptionCombo.Text
'    Case SpOptionQtRealtimeData
'        progId = SPProgIdQtRealtimeData
    Case SpOptionTwsRealtimeData
        progId = SPProgIdTwsRealtimeData
    Case SpOptionCustomRealtimeData
        progId = ProgIdText
    End Select
Case SPNameContractDataPrimary
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        progId = SPProgIdTbContractData
    Case SpOptionTwsContractData
        progId = SPProgIdTwsContractData
    Case SpOptionCustomContractData
        progId = ProgIdText
    End Select
Case SPNameContractDataSecondary
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        progId = SPProgIdTbContractData
    Case SpOptionTwsContractData
        progId = SPProgIdTwsContractData
    Case SpOptionCustomContractData
        progId = ProgIdText
    End Select
Case SPNameHistoricalDataInput
    Select Case OptionCombo.Text
'    Case SpOptionQtBarData
'        progId = SPProgIdQtBarData
    Case SpOptionTbBarData
        progId = SPProgIdTbBarData
    Case SpOptionTwsBarData
        progId = SPProgIdTwsBarData
    Case SpOptionCustomBarData
        progId = ProgIdText
    End Select
Case SPNameHistoricalDataOutput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        progId = SPProgIdTbBarData
    Case SpOptionCustomBarData
        progId = ProgIdText
    End Select
Case SPNameOrderSubmissionLive
    Select Case OptionCombo.Text
    Case SpOptionTwsOrders
        progId = SPProgIdTwsOrders
    Case SpOptionCustomOrders
        progId = ProgIdText
    End Select
Case SPNameOrderSubmissionSimulated
    Select Case OptionCombo.Text
    Case SpOptionTbOrders
        progId = SPProgIdTbOrders
    Case SpOptionCustomOrders
        progId = ProgIdText
    End Select
Case SPNameTickfileInput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        progId = SPProgIdTbTickData
'    Case SpOptionQtTickData
'        progId = SPProgIdQtTickData
    Case SpOptionFileTickData
        progId = SPProgIdFileTickData
    Case SpOptionCustomTickData
        progId = ProgIdText
    End Select
Case SPNameTickfileOutput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        progId = SPProgIdTbTickData
    Case SpOptionFileTickData
        progId = SPProgIdFileTickData
    Case SpOptionCustomTickData
        progId = ProgIdText
    End Select
End Select

mCurrSP.SetAttribute AttributeNameServiceProviderProgId, progId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setProperty( _
                ByVal props As ConfigurationSection, _
                ByVal name As String, _
                ByVal Value As String)
Const ProcName As String = "setProperty"
On Error GoTo Err

props.SetSetting "." & ConfigNameProperty & "(" & name & ")", Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupBrProperties()
Const ProcName As String = "setupBrProperties"
On Error GoTo Err

On Error Resume Next
BrEnabledCheck.Value = vbUnchecked
BrEnabledCheck.Value = IIf(mCurrSP.GetAttribute(AttributeNameServiceProviderEnabled) = "True", vbChecked, vbUnchecked)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupBrokerLiveSP()
Const ProcName As String = "setupBrokerLiveSP"
On Error GoTo Err

CategoryLabel = "Broker options (live orders)"
OptionLabel = "Select broker"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTwsOrders
OptionCombo.ComboItems.Add , , SpOptionCustomOrders

On Error Resume Next
findSp SPNameOrderSubmissionLive

Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTwsOrders
    OptionCombo.Text = SpOptionTwsOrders
    
    setupTwsProperties
Case Else
    OptionCombo.Text = SpOptionCustomOrders
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupBrokerSimulatedSP()
Const ProcName As String = "setupBrokerSimulatedSP"
On Error GoTo Err

CategoryLabel = "Broker options (simulated orders)"
OptionLabel = "Select broker"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTbOrders
OptionCombo.ComboItems.Add , , SpOptionCustomOrders

On Error Resume Next
findSp SPNameOrderSubmissionSimulated
Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTbOrders
    OptionCombo.Text = SpOptionTbOrders
    
    setupBrProperties
Case Else
    OptionCombo.Text = SpOptionCustomOrders
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupCustomProperties()
Const ProcName As String = "setupCustomProperties"
On Error GoTo Err

On Error Resume Next
CustomEnabledCheck.Value = IIf(mCurrSP.GetAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
ProgIdText = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")

mCustomParams.RemoveCollectionChangeListener Me

Set mCustomParams = New Parameters

Dim prop As ConfigurationSection
For Each prop In mCurrProps
    mCustomParams.SetParameterValue prop.InstanceQualifier, _
                                    prop.Value
Next

On Error GoTo Err

Set ParamsGrid.DataSource = mCustomParams

mCustomParams.AddCollectionChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupDbProperties()
Const ProcName As String = "setupDbProperties"
On Error GoTo Err

On Error Resume Next
DbEnabledCheck.Value = vbUnchecked
DbServerText = ""
DbTypeCombo = ""
DbDatabaseText = ""
DbUsernameText = ""
DbPasswordText = ""
DbEnabledCheck.Value = IIf(mCurrSP.GetAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
DbServerText = getProperty(PropertyNameTbServer)
DbTypeCombo = getProperty(PropertyNameTbDbType)
DbDatabaseText = getProperty(PropertyNameTbDbName)
DbUsernameText = getProperty(PropertyNameTbUserName)
DbPasswordText = getProperty(PropertyNameTbPassword)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupHistoricalDataInputSP()
Const ProcName As String = "setupHistoricalDataInputSP"
On Error GoTo Err

CategoryLabel = "Historical bar data retrieval source options"
OptionLabel = "Select historical bar data source"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTbBarData
'OptionCombo.ComboItems.Add , , SpOptionQtBarData
OptionCombo.ComboItems.Add , , SpOptionTwsBarData
OptionCombo.ComboItems.Add , , SpOptionCustomBarData

On Error Resume Next
findSp SPNameHistoricalDataInput
Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTwsBarData
    OptionCombo.Text = SpOptionTwsBarData
    
    setupTwsProperties
Case SPProgIdTbBarData
    OptionCombo.Text = SpOptionTbBarData
    
    setupDbProperties
'Case SPProgIdQtBarData
'    OptionCombo.Text = SpOptionQtBarData
'
'    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionCustomBarData
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupHistoricalDataOutputSP()
Const ProcName As String = "setupHistoricalDataOutputSP"
On Error GoTo Err

CategoryLabel = "Historical bar data storage options"
OptionLabel = "Select historical bar data source"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTbBarData
OptionCombo.ComboItems.Add , , SpOptionCustomBarData

On Error Resume Next
findSp SPNameHistoricalDataOutput
Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTbBarData
    OptionCombo.Text = SpOptionTbBarData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionCustomBarData
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupPrimaryContractDataSP()
Const ProcName As String = "setupPrimaryContractDataSP"
On Error GoTo Err

CategoryLabel = "Primary contract data source options"
OptionLabel = "Select primary contract data source"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTbContractData
OptionCombo.ComboItems.Add , , SpOptionTwsContractData
OptionCombo.ComboItems.Add , , SpOptionCustomContractData

On Error Resume Next
Dim progId As String
findSp SPNameContractDataPrimary
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTwsContractData
    OptionCombo.Text = SpOptionTwsContractData
    
    setupTwsProperties
Case SPProgIdTbContractData
    OptionCombo.Text = SpOptionTbContractData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionCustomContractData
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupRealtimeDataSP()
Const ProcName As String = "setupRealtimeDataSP"
On Error GoTo Err

CategoryLabel = "Realtime data source options"
OptionLabel = "Select realtime data source"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
'OptionCombo.ComboItems.Add , , SpOptionQtRealtimeData
OptionCombo.ComboItems.Add , , SpOptionTwsRealtimeData
OptionCombo.ComboItems.Add , , SpOptionCustomRealtimeData

On Error Resume Next
findSp SPNameRealtimeData
Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTwsRealtimeData
    OptionCombo.Text = SpOptionTwsRealtimeData

    setupTwsProperties
'Case SPProgIdQtRealtimeData
'    OptionCombo.Text = SpOptionQtRealtimeData
'
'    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionCustomRealtimeData

    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupQtProperties()
Const ProcName As String = "setupQtProperties"
On Error GoTo Err

On Error Resume Next
QtEnabledCheck.Value = IIf(mCurrSP.GetAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
QtKeepConnectionCheck.Value = IIf(getProperty(PropertyNameQtKeepConnection) = "True", vbChecked, vbUnchecked)
QtServerText = getProperty(PropertyNameQtServer)
QtPortText = getProperty(PropertyNameQtPort)
QtPasswordText = getProperty(PropertyNameQtPassword)
QtProviderKeyText = getProperty(PropertyNameQtProviderKey)
QtConnectRetryIntervalText = getProperty(PropertyNameQtConnectionRetryInterval)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupSecondaryContractDataSP()
Const ProcName As String = "setupSecondaryContractDataSP"
On Error GoTo Err

CategoryLabel = "Secondary contract data source options"
OptionLabel = "Select secondary contract data source"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTbContractData
OptionCombo.ComboItems.Add , , SpOptionTwsContractData
OptionCombo.ComboItems.Add , , SpOptionCustomContractData

On Error Resume Next
findSp SPNameContractDataSecondary
Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTwsContractData
    OptionCombo.Text = SpOptionTwsContractData
    
    setupTwsProperties
Case SPProgIdTbContractData
    OptionCombo.Text = SpOptionTbContractData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionCustomContractData
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupTfInputProperties()
Const ProcName As String = "setupTfInputProperties"
On Error GoTo Err

On Error Resume Next
TfInputEnabledCheck.Value = IIf(mCurrSP.GetAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTfOutputProperties()
Const ProcName As String = "setupTfOutputProperties"
On Error GoTo Err

On Error Resume Next
TfOutputEnabledCheck.Value = IIf(mCurrSP.GetAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
OutputTickfilePathText.Text = getProperty(PropertyNameTfTickfilePath)
TickfileGranularityCombo.Text = getProperty(PropertyNameTfTickfileGranularity)
If TickfileGranularityCombo.Text = "" Then TickfileGranularityCombo.Text = TickfileGranularityExecution
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTickfileInputSP()
Const ProcName As String = "setupTickfileInputSP"
On Error GoTo Err

CategoryLabel = "Tickfile replay data source options"
OptionLabel = "Select tickfile replay data source"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTbTickData
OptionCombo.ComboItems.Add , , SpOptionFileTickData
'OptionCombo.ComboItems.Add , , SpOptionQtTickData
OptionCombo.ComboItems.Add , , SpOptionCustomTickData

On Error Resume Next
findSp SPNameTickfileInput
Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTbTickData
    OptionCombo.Text = SpOptionTbTickData
    
    setupDbProperties
Case SPProgIdFileTickData
    OptionCombo.Text = SpOptionFileTickData
    
    setupTfInputProperties
'Case SPProgIdQtTickData
'    OptionCombo.Text = SpOptionQtTickData
'
'    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionCustomTickData
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTickfileOutputSP()
Const ProcName As String = "setupTickfileOutputSP"
On Error GoTo Err

CategoryLabel = "Tickfile storage options"
OptionLabel = "Select tickfile data store"
OptionCombo.ComboItems.Clear
OptionCombo.ComboItems.Add , , SpOptionNotConfigured
OptionCombo.ComboItems.Add , , SpOptionTbTickData
OptionCombo.ComboItems.Add , , SpOptionFileTickData
OptionCombo.ComboItems.Add , , SpOptionCustomTickData

On Error Resume Next
findSp SPNameTickfileOutput
Dim progId As String
progId = mCurrSP.GetAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo Err

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case SPProgIdTbTickData
    OptionCombo.Text = SpOptionTbTickData
    
    setupDbProperties
Case SPProgIdFileTickData
    OptionCombo.Text = SpOptionFileTickData
    
    setupTfOutputProperties
Case Else
    OptionCombo.Text = SpOptionCustomTickData
    
    setupCustomProperties
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTwsProperties()
Const ProcName As String = "setupTwsProperties"
On Error GoTo Err

On Error Resume Next
TwsEnabledCheck.Value = vbUnchecked
TwsKeepConnectionCheck.Value = vbUnchecked
TWSServerText = ""
TWSPortText = ""
TWSClientIdText = ""
TwsProviderKeyText = ""
TwsConnectRetryIntervalText = ""
TwsEnabledCheck.Value = IIf(mCurrSP.GetAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
TwsKeepConnectionCheck.Value = IIf(getProperty(PropertyNameTwsKeepConnection) = "True", vbChecked, vbUnchecked)
TWSServerText = getProperty(PropertyNameTwsServer)
TWSPortText = getProperty(PropertyNameTwsPort)
TWSClientIdText = getProperty(PropertyNameTwsClientId)
TwsProviderKeyText = getProperty(PropertyNameTwsProviderKey)
TwsConnectRetryIntervalText = getProperty(PropertyNameTwsConnectionRetryInterval)
Dim twsLogLevel As String
twsLogLevel = getProperty(PropertyNameTwsLogLevel)
If twsLogLevel = "" Then
    TwsLogLevelCombo.Text = TWSLogLevelSystem
Else
    TwsLogLevelCombo.Text = getProperty(PropertyNameTwsLogLevel)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showSpOptions()
Const ProcName As String = "showSpOptions"
On Error GoTo Err

Select Case mCurrCategory
Case SPNameRealtimeData
    Select Case OptionCombo.Text
'    Case SpOptionQtRealtimeData
'        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionTwsRealtimeData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomRealtimeData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameContractDataPrimary
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionTwsContractData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomContractData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameContractDataSecondary
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionTwsContractData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomContractData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameHistoricalDataInput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        Set mCurrOptionsPic = DbOptionsPicture
'    Case SpOptionQtBarData
'        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionTwsBarData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomBarData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameHistoricalDataOutput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionCustomBarData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameOrderSubmissionLive
    Select Case OptionCombo.Text
    Case SpOptionTwsOrders
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomOrders
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameOrderSubmissionSimulated
    Select Case OptionCombo.Text
    Case SpOptionTbOrders
        Set mCurrOptionsPic = BrOptionsPicture
    Case SpOptionCustomOrders
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameTickfileInput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        Set mCurrOptionsPic = DbOptionsPicture
'    Case SpOptionQtTickData
'        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionFileTickData
        Set mCurrOptionsPic = TfInputOptionsPicture
    Case SpOptionCustomTickData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case SPNameTickfileOutput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionFileTickData
        Set mCurrOptionsPic = TfOutputOptionsPicture
    Case SpOptionCustomTickData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
End Select

If Not mCurrOptionsPic Is Nothing Then
    mCurrOptionsPic.Left = OptionsBox.Left
    mCurrOptionsPic.Top = OptionsBox.Top
    mCurrOptionsPic.Visible = True
    mCurrOptionsPic.Refresh
End If

OptionCombo.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

