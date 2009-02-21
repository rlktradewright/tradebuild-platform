VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#34.0#0"; "TWControls10.ocx"
Begin VB.UserControl SPConfigurer 
   ClientHeight    =   12750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16740
   DefaultCancel   =   -1  'True
   ScaleHeight     =   12750
   ScaleWidth      =   16740
   Begin TWControls10.TWImageCombo OptionCombo 
      Height          =   330
      Left            =   4320
      TabIndex        =   0
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
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
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   360
         Width           =   3855
      End
      Begin VB.CheckBox CustomEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   26
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
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   28
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox TfOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5040
      ScaleHeight     =   2175
      ScaleWidth      =   4815
      TabIndex        =   48
      Top             =   6480
      Width           =   4815
      Begin VB.CommandButton PathChooserButton 
         Caption         =   "..."
         Height          =   375
         Left            =   4320
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TickfilePathText 
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox TfEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2535
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
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton ApplyButton 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   29
      Top             =   3480
      Width           =   975
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
         Height          =   285
         Left            =   960
         TabIndex        =   20
         Text            =   "10"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox QtProviderKeyText 
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox QtKeepConnectionCheck 
         Caption         =   "Keep connection"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox QtEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox QtPasswordText 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox QtPortText 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Text            =   "16240"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox QtServerText 
         Height          =   285
         Left            =   960
         TabIndex        =   16
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
      Begin TWControls10.TWImageCombo DbTypeCombo 
         Height          =   330
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
         MouseIcon       =   "SPConfigurer.ctx":001C
         Text            =   ""
      End
      Begin VB.CheckBox DbEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox DbDatabaseText 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox DbServerText 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox DbUsernameText 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox DbPasswordText 
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
      Begin TWControls10.TWImageCombo TwsLogLevelCombo 
         Height          =   330
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         MouseIcon       =   "SPConfigurer.ctx":0038
         Text            =   ""
      End
      Begin VB.CheckBox TwsKeepConnectionCheck 
         Caption         =   "Keep connection"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox TwsConnectRetryIntervalText 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "10"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox TwsProviderKeyText 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox TwsEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox TWSClientIdText 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Text            =   "-1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TWSPortText 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "7496"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TWSServerText 
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
      Height          =   3765
      ItemData        =   "SPConfigurer.ctx":0054
      Left            =   120
      List            =   "SPConfigurer.ctx":0056
      TabIndex        =   30
      Top             =   120
      Width           =   2055
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
   Begin VB.Label CategoryLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E7D395&
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

Private Const ProjectName                   As String = "TradeBuildUI26"
Private Const ModuleName                    As String = "SPConfigurer1"

Private Const AccessModeReadOnly            As String = "Read only"
Private Const AccessModeWriteOnly           As String = "Write only"
Private Const AccessModeReadWrite           As String = "Read write"

'Private Const AttributeNamePropertyName     As String = "Name"
'Private Const AttributeNamePropertyValue    As String = "Value"
Private Const AttributeNameServiceProviderEnabled As String = "Enabled"
Private Const AttributeNameServiceProviderName As String = "Name"
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
Private Const PropertyNameTbAccessMode      As String = "Access Mode"
Private Const PropertyNameTbRole            As String = "Role"

Private Const PropertyNameTfAccessMode      As String = "Access Mode"
Private Const PropertyNameTfTickfilePath    As String = "Tickfile Path"

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

Private Const SpOptionQtBarData             As String = "QuoteTracker"
Private Const SpOptionQtRealtimeData        As String = "QuoteTracker"
Private Const SpOptionQtTickData            As String = "QuoteTracker"

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

Private mSPs           As ServiceProviders

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height

DbTypeCombo.ComboItems.add , , ""
DbTypeCombo.ComboItems.add , , DbTypeMySql
DbTypeCombo.ComboItems.add , , DbTypeSqlServer7
DbTypeCombo.ComboItems.add , , DbTypeSqlServer2000
DbTypeCombo.ComboItems.add , , DbTypeSqlServer2005

TwsLogLevelCombo.ComboItems.add , , TWSLogLevelDetail
TwsLogLevelCombo.ComboItems.add , , TWSLogLevelError
TwsLogLevelCombo.ComboItems.add , , TWSLogLevelInformation
TwsLogLevelCombo.ComboItems.add , , TWSLogLevelSystem
TwsLogLevelCombo.ComboItems.add , , TWSLogLevelWarning
TwsLogLevelCombo.Text = TWSLogLevelSystem

End Sub

Private Sub UserControl_InitProperties()
UserControl.backColor = UserControl.Ambient.backColor
UserControl.foreColor = UserControl.Ambient.foreColor
End Sub

Private Sub UserControl_LostFocus()
checkForOutstandingUpdates
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.backColor = UserControl.Ambient.backColor
UserControl.foreColor = UserControl.Ambient.foreColor
End Sub

Private Sub UserControl_Resize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change( _
                ev As TWUtilities30.CollectionChangeEvent)
If ev.Source Is mCustomParams Then
    enableApplyButton True
    enableCancelButton True
End If
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ApplyButton_Click()
applyProperties
enableApplyButton False
enableCancelButton False
End Sub

Private Sub BrEnabledCheck_Click()
enableApplyButton True
enableCancelButton True
End Sub

Private Sub CancelButton_Click()
Dim index As Long
enableApplyButton False
enableCancelButton False
index = CategoryList.ListIndex
CategoryList.ListIndex = -1
CategoryList.ListIndex = index
End Sub

Private Sub CategoryList_Click()

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
Case mSPs.SPNameRealtimeData
    setupRealtimeDataSP
Case mSPs.SPNamePrimaryContractData
    setupPrimaryContractDataSP
Case mSPs.SPNameSecondryContractData
    setupSecondaryContractDataSP
Case mSPs.SPNameHistoricalDataInput
    setupHistoricalDataInputSP
Case mSPs.SPNameHistoricalDataOutput
    setupHistoricalDataOutputSP
Case mSPs.SPNameBrokerLive
    setupBrokerLiveSP
Case mSPs.SPNameBrokerSimulated
    setupBrokerSimulatedSP
Case mSPs.SPNameTickfileInput
    setupTickfileInputSP
Case mSPs.SPNameTickfileOutput
    setupTickfileOutputSP
End Select

showSpOptions

enableApplyButton False
enableCancelButton False
End Sub

Private Sub DbDatabaseText_Change()
enableApplyButton isValidDbProperties
enableCancelButton True
End Sub

Private Sub DbEnabledCheck_Click()
enableApplyButton isValidDbProperties
enableCancelButton True
End Sub

Private Sub DbPasswordText_Change()
enableApplyButton isValidDbProperties
enableCancelButton True
End Sub

Private Sub DbServerText_Change()
enableApplyButton isValidDbProperties
enableCancelButton True
End Sub

Private Sub DbTypeCombo_Click()
enableApplyButton isValidDbProperties
enableCancelButton True
End Sub

Private Sub DbUsernameText_Change()
enableApplyButton isValidDbProperties
enableCancelButton True
End Sub

Private Sub OptionCombo_Click()
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
        ElseIf mCurrOptionsPic Is TfOptionsPicture Then
            enableApplyButton True
        ElseIf mCurrOptionsPic Is CustomOptionsPicture Then
            enableApplyButton isValidCustomProperties
        End If
    End If
End If
mCurrSpOption = OptionCombo.Text
enableCancelButton True
End Sub

Private Sub PathChooserButton_Click()
Dim f As New fPathChooser
f.path = TickfilePathText.Text
f.Show vbModal
If Not f.cancelled Then TickfilePathText.Text = f.path
Unload f
End Sub

Private Sub ProgIdText_Change()
enableApplyButton isValidCustomProperties
enableCancelButton True
End Sub

Private Sub QtEnabledCheck_Click()
enableApplyButton isValidQtProperties
enableCancelButton True
End Sub

Private Sub QtPasswordText_Change()
enableApplyButton isValidQtProperties
enableCancelButton True
End Sub

Private Sub QtPortText_Change()
enableApplyButton isValidQtProperties
enableCancelButton True
End Sub

Private Sub QtServerText_Change()
enableApplyButton isValidQtProperties
enableCancelButton True
End Sub

Private Sub SpLogLevelCombo_Click()
If mCurrOptionsPic Is BrOptionsPicture Then
    enableApplyButton True
ElseIf mCurrOptionsPic Is CustomOptionsPicture Then
    enableApplyButton isValidCustomProperties
ElseIf mCurrOptionsPic Is DbOptionsPicture Then
    enableApplyButton isValidDbProperties
ElseIf mCurrOptionsPic Is QtOptionsPicture Then
    enableApplyButton isValidQtProperties
ElseIf mCurrOptionsPic Is TfOptionsPicture Then
    enableApplyButton True
ElseIf mCurrOptionsPic Is TwsOptionsPicture Then
    enableApplyButton isValidTwsProperties
End If
enableCancelButton True
End Sub

Private Sub TfEnabledCheck_Click()
enableApplyButton True
enableCancelButton True
End Sub

Private Sub TWSClientIdText_Change()
enableApplyButton isValidTwsProperties
enableCancelButton True
End Sub

Private Sub TwsConnectRetryIntervalText_Change()
enableApplyButton isValidTwsProperties
enableCancelButton True
End Sub

Private Sub TwsEnabledCheck_Click()
enableApplyButton isValidTwsProperties
enableCancelButton True
End Sub

Private Sub TWSPortText_Change()
enableApplyButton isValidTwsProperties
enableCancelButton True
End Sub

Private Sub TwsKeepConnectionCheck_Click()
enableApplyButton isValidTwsProperties
enableCancelButton True
End Sub

Private Sub TwsProviderKeyText_Change()
enableApplyButton isValidTwsProperties
enableCancelButton True
End Sub

Private Sub TWSServerText_Change()
enableApplyButton isValidTwsProperties
enableCancelButton True
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get dirty() As Boolean
dirty = ApplyButton.Enabled
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub applyChanges()
applyProperties
enableApplyButton False
enableCancelButton False
End Sub

Public Sub Initialise( _
                ByVal configdata As ConfigurationSection, _
                Optional ByVal readonly As Boolean)
Set mSPs = TradeBuildAPI.ServiceProviders

mReadOnly = readonly

checkForOutstandingUpdates

Set mCurrSPsList = Nothing
Set mCurrSP = Nothing
Set mCurrProps = Nothing
mCurrCategory = ""
mCurrSpOption = ""

Dim da As DataAdapter
If mCustomParams Is Nothing Then
    Set mCustomParams = New Parameters
    Set da = New DataAdapter
    Set da.Object = mCustomParams
    Set ParamsGrid.DataSource = da
    mCustomParams.addCollectionChangeListener Me
End If

loadConfig configdata

If mReadOnly Then disableControls
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub applyBrProperties()
If BrEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
End If
End Sub

Private Sub applyCustomProperties()
Dim param As Parameter

If CustomEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
End If

For Each param In mCustomParams
    setProperty mCurrProps, param.name, param.value
Next
End Sub

Private Sub applyDbProperties()
If DbEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
End If
setProperty mCurrProps, PropertyNameTbServer, DbServerText
setProperty mCurrProps, PropertyNameTbDbType, DbTypeCombo
setProperty mCurrProps, PropertyNameTbDbName, DbDatabaseText
setProperty mCurrProps, PropertyNameTbUserName, DbUsernameText
setProperty mCurrProps, PropertyNameTbPassword, DbPasswordText

If mCurrCategory = mSPs.SPNameHistoricalDataInput Or _
    mCurrCategory = mSPs.SPNameTickfileInput _
Then
    setProperty mCurrProps, PropertyNameTbAccessMode, AccessModeReadOnly
End If

If mCurrCategory = mSPs.SPNameHistoricalDataOutput Or _
    mCurrCategory = mSPs.SPNameTickfileOutput _
Then
    setProperty mCurrProps, PropertyNameTbAccessMode, AccessModeWriteOnly
End If

If mCurrCategory = mSPs.SPNamePrimaryContractData Then
    setProperty mCurrProps, PropertyNameTbRole, RolePrimary
End If

If mCurrCategory = mSPs.SPNameSecondryContractData Then
    setProperty mCurrProps, PropertyNameTbRole, RoleSecondary
End If
End Sub

Private Sub applyProperties()
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
ElseIf mCurrOptionsPic Is TfOptionsPicture Then
    applyTfProperties
ElseIf mCurrOptionsPic Is TwsOptionsPicture Then
    applyTwsProperties
End If
End Sub

Private Sub applyQtProperties()
If QtEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
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
End Sub

Private Sub applyTfProperties()
If TfEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
End If
If mCurrCategory = mSPs.SPNameTickfileInput Then
    setProperty mCurrProps, PropertyNameTfAccessMode, AccessModeReadOnly
End If
If mCurrCategory = mSPs.SPNameTickfileOutput Then
    setProperty mCurrProps, PropertyNameTfAccessMode, AccessModeWriteOnly
    setProperty mCurrProps, PropertyNameTfTickfilePath, TickfilePathText
End If
End Sub

Private Sub applyTwsProperties()
If TwsEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
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

If mCurrCategory = mSPs.SPNamePrimaryContractData Then
    setProperty mCurrProps, PropertyNameTbRole, RolePrimary
End If

If mCurrCategory = mSPs.SPNameSecondryContractData Then
    setProperty mCurrProps, PropertyNameTbRole, RoleSecondary
End If
End Sub

Private Sub checkForOutstandingUpdates()
If ApplyButton.Enabled Then
    If MsgBox("Do you want to apply the changes you have made?", _
            vbExclamation Or vbYesNoCancel) = vbYes Then
        applyProperties
        enableApplyButton False
        enableCancelButton False
    End If
End If
End Sub

Private Sub clearProperties()
mCurrProps.RemoveAllChildren
End Sub

Private Sub createNewSp()
If mCurrSPsList Is Nothing Then
    Set mCurrSPsList = mConfig.AddConfigurationSection(ConfigNameServiceProviders)
End If

Set mCurrSP = mCurrSPsList.AddConfigurationSection(ConfigNameServiceProvider & "(" & mCurrCategory & ")")
Set mCurrProps = mCurrSP.AddConfigurationSection(ConfigNameProperties)
End Sub

Private Sub deleteSp()
mCurrSPsList.RemoveConfigurationSection ConfigNameServiceProvider & "(" & mCurrSP.InstanceQualifier & ")"
Set mCurrSP = Nothing
Set mCurrProps = Nothing
End Sub

Private Sub disableControls()
CancelButton.Enabled = False
ApplyButton.Enabled = False
End Sub

Private Sub enableApplyButton( _
                ByVal enable As Boolean)
If mReadOnly Then Exit Sub
If enable Then
    ApplyButton.Enabled = True
    ApplyButton.Default = True
Else
    ApplyButton.Enabled = False
    ApplyButton.Default = False
End If
End Sub

Private Sub enableCancelButton( _
                ByVal enable As Boolean)
If mReadOnly Then Exit Sub
If enable Then
    CancelButton.Enabled = True
    CancelButton.Cancel = True
Else
    CancelButton.Enabled = False
    CancelButton.Cancel = False
End If
End Sub

'Private Function findProperty( _
'                ByVal name As String) As ConfigurationSection
'Dim prop As ConfigurationSection
'
'name = UCase$(name)
'For Each prop In mCurrProps.childItems
'    If UCase$(prop.getAttribute(AttributeNamePropertyName)) = name Then
'        Set findProperty = prop
'        Exit Function
'    End If
'Next
'End Function

Private Function findSp( _
                ByVal name As String) As Boolean
Dim sp As ConfigurationSection
Set mCurrSP = Nothing
Set mCurrProps = Nothing
mCurrSpOption = ""

If mCurrSPsList Is Nothing Then Exit Function

Set mCurrSP = mCurrSPsList.GetConfigurationSection(ConfigNameServiceProvider & "(" & name & ")")
If Not mCurrSP Is Nothing Then
    Set mCurrProps = mCurrSP.GetConfigurationSection(ConfigNameProperties)
    findSp = True
End If
End Function

Private Function getProperty( _
                ByVal name As String) As String
On Error Resume Next
getProperty = mCurrProps.GetSetting("." & ConfigNameProperty & "(" & name & ")")
End Function

Private Sub hideSpOptions()
If Not mCurrOptionsPic Is Nothing Then
    mCurrOptionsPic.Visible = False
    Set mCurrOptionsPic = Nothing
End If
End Sub

Private Function isValidCustomProperties() As Boolean
If ProgIdText = "" Then
ElseIf InStr(1, ProgIdText, ".") < 2 Then
ElseIf InStr(1, ProgIdText, ".") = Len(ProgIdText) Then
ElseIf Len(ProgIdText) > 39 Then
Else
    isValidCustomProperties = True
End If
End Function

Private Function isValidDbProperties() As Boolean
If DbDatabaseText = "" Then
ElseIf DbTypeCombo.Text = "" Then
ElseIf DbTypeCombo.Text = DbTypeMySql And DbUsernameText = "" Then
Else
    isValidDbProperties = True
End If
End Function

Private Function isValidQtProperties() As Boolean
If Not IsInteger(QtPortText, 1024) Then
Else
    isValidQtProperties = True
End If
End Function

Private Function isValidTwsProperties() As Boolean
If Not IsInteger(TWSPortText, 1) Then
ElseIf Not IsInteger(TWSClientIdText) Then
ElseIf TwsConnectRetryIntervalText <> "" And Not IsInteger(TwsConnectRetryIntervalText, 0) Then
Else
    isValidTwsProperties = True
End If
End Function

Private Sub loadConfig( _
                ByVal configdata As ConfigurationSection)
                
Set mConfig = configdata

On Error Resume Next
Set mCurrSPsList = mConfig.GetConfigurationSection(ConfigNameServiceProviders)
On Error GoTo 0

CategoryList.clear

If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPRealtimeData Then
    CategoryList.addItem mSPs.SPNameRealtimeData
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPPrimaryContractData Then
    CategoryList.addItem mSPs.SPNamePrimaryContractData
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPSecondaryContractData Then
    CategoryList.addItem mSPs.SPNameSecondryContractData
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPHistoricalDataInput Then
    CategoryList.addItem mSPs.SPNameHistoricalDataInput
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPHistoricalDataOutput Then
    CategoryList.addItem mSPs.SPNameHistoricalDataOutput
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPBrokerLive Then
    CategoryList.addItem mSPs.SPNameBrokerLive
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPBrokerSimulated Then
    CategoryList.addItem mSPs.SPNameBrokerSimulated
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPTickfileInput Then
    CategoryList.addItem mSPs.SPNameTickfileInput
End If
If mSPs.PermittedServiceProviderRoles And ServiceProviderRoles.SPTickfileOutput Then
    CategoryList.addItem mSPs.SPNameTickfileOutput
End If

If CategoryList.ListCount > 0 Then CategoryList.ListIndex = 0

End Sub

Private Sub setProgId()
Dim progId As String

If CategoryList.ListIndex = -1 Then Exit Sub

Select Case mCurrCategory
Case mSPs.SPNameRealtimeData
    Select Case OptionCombo.Text
    Case SpOptionQtRealtimeData
        progId = mSPs.SPProgIdQtRealtimeData
    Case SpOptionTwsRealtimeData
        progId = mSPs.SPProgIdTwsRealtimeData
    Case SpOptionCustomRealtimeData
        progId = ProgIdText
    End Select
Case mSPs.SPNamePrimaryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        progId = mSPs.SPProgIdTbContractData
    Case SpOptionTwsContractData
        progId = mSPs.SPProgIdTwsContractData
    Case SpOptionCustomContractData
        progId = ProgIdText
    End Select
Case mSPs.SPNameSecondryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        progId = mSPs.SPProgIdTbContractData
    Case SpOptionTwsContractData
        progId = mSPs.SPProgIdTwsContractData
    Case SpOptionCustomContractData
        progId = ProgIdText
    End Select
Case mSPs.SPNameHistoricalDataInput
    Select Case OptionCombo.Text
    Case SpOptionQtBarData
        progId = mSPs.SPProgIdQtBarData
    Case SpOptionTbBarData
        progId = mSPs.SPProgIdTbBarData
    Case SpOptionTwsBarData
        progId = mSPs.SPProgIdTwsBarData
    Case SpOptionCustomBarData
        progId = ProgIdText
    End Select
Case mSPs.SPNameHistoricalDataOutput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        progId = mSPs.SPProgIdTbBarData
    Case SpOptionCustomBarData
        progId = ProgIdText
    End Select
Case mSPs.SPNameBrokerLive
    Select Case OptionCombo.Text
    Case SpOptionTwsOrders
        progId = mSPs.SPProgIdTwsOrders
    Case SpOptionCustomOrders
        progId = ProgIdText
    End Select
Case mSPs.SPNameBrokerSimulated
    Select Case OptionCombo.Text
    Case SpOptionTbOrders
        progId = mSPs.SPProgIdTbOrders
    Case SpOptionCustomOrders
        progId = ProgIdText
    End Select
Case mSPs.SPNameTickfileInput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        progId = mSPs.SPProgIdTbTickData
    Case SpOptionQtTickData
        progId = mSPs.SPProgIdQtTickData
    Case SpOptionFileTickData
        progId = mSPs.SPProgIdFileTickData
    Case SpOptionCustomTickData
        progId = ProgIdText
    End Select
Case mSPs.SPNameTickfileOutput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        progId = mSPs.SPProgIdTbTickData
    Case SpOptionFileTickData
        progId = mSPs.SPProgIdFileTickData
    Case SpOptionCustomTickData
        progId = ProgIdText
    End Select
End Select

mCurrSP.setAttribute AttributeNameServiceProviderProgId, progId

End Sub

Private Sub setProperty( _
                ByVal props As ConfigurationSection, _
                ByVal name As String, _
                ByVal value As String)
props.SetSetting "." & ConfigNameProperty & "(" & name & ")", value
End Sub

Private Sub setupBrProperties()
On Error Resume Next
BrEnabledCheck.value = vbUnchecked
BrEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled) = "True", vbChecked, vbUnchecked)
End Sub

Private Sub setupBrokerLiveSP()
Dim progId As String
    
CategoryLabel = "Broker options (live orders)"
OptionLabel = "Select broker"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTwsOrders
OptionCombo.ComboItems.add , , SpOptionCustomOrders

On Error Resume Next
findSp mSPs.SPNameBrokerLive
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTwsOrders
    OptionCombo.Text = SpOptionTwsOrders
    
    setupTwsProperties
Case Else
    OptionCombo.Text = SpOptionCustomOrders
    
    setupCustomProperties
End Select

End Sub

Private Sub setupBrokerSimulatedSP()
Dim progId As String
    
CategoryLabel = "Broker options (simulated orders)"
OptionLabel = "Select broker"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTbOrders
OptionCombo.ComboItems.add , , SpOptionCustomOrders

On Error Resume Next
findSp mSPs.SPNameBrokerSimulated
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTbOrders
    OptionCombo.Text = SpOptionTbOrders
    
    setupBrProperties
Case Else
    OptionCombo.Text = SpOptionCustomOrders
    
    setupCustomProperties
End Select

End Sub

Private Sub setupCustomProperties()
Dim prop As ConfigurationSection
Dim da As DataAdapter

On Error Resume Next
CustomEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
ProgIdText = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")

mCustomParams.removeCollectionChangeListener Me

Set mCustomParams = New Parameters

For Each prop In mCurrProps
    mCustomParams.setParameterValue prop.InstanceQualifier, _
                                    prop.value
Next

On Error GoTo 0

Set da = New DataAdapter
Set da.Object = mCustomParams
Set ParamsGrid.DataSource = da

mCustomParams.addCollectionChangeListener Me

End Sub

Private Sub setupDbProperties()
On Error Resume Next
DbEnabledCheck.value = vbUnchecked
DbServerText = ""
DbTypeCombo = ""
DbDatabaseText = ""
DbUsernameText = ""
DbPasswordText = ""
DbEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
DbServerText = getProperty(PropertyNameTbServer)
DbTypeCombo = getProperty(PropertyNameTbDbType)
DbDatabaseText = getProperty(PropertyNameTbDbName)
DbUsernameText = getProperty(PropertyNameTbUserName)
DbPasswordText = getProperty(PropertyNameTbPassword)
End Sub

Private Sub setupHistoricalDataInputSP()
Dim progId As String
    
CategoryLabel = "Historical bar data retrieval source options"
OptionLabel = "Select historical bar data source"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTbBarData
OptionCombo.ComboItems.add , , SpOptionQtBarData
OptionCombo.ComboItems.add , , SpOptionTwsBarData
OptionCombo.ComboItems.add , , SpOptionCustomBarData

On Error Resume Next
findSp mSPs.SPNameHistoricalDataInput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTwsBarData
    OptionCombo.Text = SpOptionTwsBarData
    
    setupTwsProperties
Case mSPs.SPProgIdTbBarData
    OptionCombo.Text = SpOptionTbBarData
    
    setupDbProperties
Case mSPs.SPProgIdQtBarData
    OptionCombo.Text = SpOptionQtBarData
    
    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionCustomBarData
    
    setupCustomProperties
End Select

End Sub

Private Sub setupHistoricalDataOutputSP()
Dim progId As String
    
CategoryLabel = "Historical bar data storage options"
OptionLabel = "Select historical bar data source"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTbBarData
OptionCombo.ComboItems.add , , SpOptionTwsBarData
OptionCombo.ComboItems.add , , SpOptionCustomBarData

On Error Resume Next
findSp mSPs.SPNameHistoricalDataOutput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTbBarData
    OptionCombo.Text = SpOptionTbBarData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionCustomBarData
    
    setupCustomProperties
End Select

End Sub

Private Sub setupPrimaryContractDataSP()
Dim progId As String
    
CategoryLabel = "Primary contract data source options"
OptionLabel = "Select primary contract data source"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTbContractData
OptionCombo.ComboItems.add , , SpOptionTwsContractData
OptionCombo.ComboItems.add , , SpOptionCustomContractData

On Error Resume Next
findSp mSPs.SPNamePrimaryContractData
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTwsContractData
    OptionCombo.Text = SpOptionTwsContractData
    
    setupTwsProperties
Case mSPs.SPProgIdTbContractData
    OptionCombo.Text = SpOptionTbContractData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionCustomContractData
    
    setupCustomProperties
End Select

End Sub

Private Sub setupRealtimeDataSP()
Dim progId As String

CategoryLabel = "Realtime data source options"
OptionLabel = "Select realtime data source"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionQtRealtimeData
OptionCombo.ComboItems.add , , SpOptionTwsRealtimeData
OptionCombo.ComboItems.add , , SpOptionCustomRealtimeData

On Error Resume Next
findSp mSPs.SPNameRealtimeData
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTwsRealtimeData
    OptionCombo.Text = SpOptionTwsRealtimeData
    
    setupTwsProperties
Case mSPs.SPProgIdQtRealtimeData
    OptionCombo.Text = SpOptionQtRealtimeData

    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionCustomRealtimeData
    
    setupCustomProperties
End Select

End Sub

Private Sub setupQtProperties()
On Error Resume Next
QtEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
QtKeepConnectionCheck.value = IIf(getProperty(PropertyNameQtKeepConnection) = "True", vbChecked, vbUnchecked)
QtServerText = getProperty(PropertyNameQtServer)
QtPortText = getProperty(PropertyNameQtPort)
QtPasswordText = getProperty(PropertyNameQtPassword)
QtProviderKeyText = getProperty(PropertyNameQtProviderKey)
QtConnectRetryIntervalText = getProperty(PropertyNameQtConnectionRetryInterval)
End Sub

Private Sub setupSecondaryContractDataSP()
Dim progId As String
    
CategoryLabel = "Secondary contract data source options"
OptionLabel = "Select secondary contract data source"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTbContractData
OptionCombo.ComboItems.add , , SpOptionTwsContractData
OptionCombo.ComboItems.add , , SpOptionCustomContractData

On Error Resume Next
findSp mSPs.SPNameSecondryContractData
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTwsContractData
    OptionCombo.Text = SpOptionTwsContractData
    
    setupTwsProperties
Case mSPs.SPProgIdTbContractData
    OptionCombo.Text = SpOptionTbContractData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionCustomContractData
    
    setupCustomProperties
End Select

End Sub

Private Sub setupTfProperties()
On Error Resume Next
TfEnabledCheck.value = vbUnchecked
TickfilePathText.Text = ""
TfEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
TickfilePathText.Text = getProperty(PropertyNameTfTickfilePath)
End Sub

Private Sub setupTickfileInputSP()
Dim progId As String

CategoryLabel = "Tickfile replay data source options"
OptionLabel = "Select tickfile replay data source"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTbTickData
OptionCombo.ComboItems.add , , SpOptionFileTickData
OptionCombo.ComboItems.add , , SpOptionQtTickData
OptionCombo.ComboItems.add , , SpOptionCustomTickData

On Error Resume Next
findSp mSPs.SPNameTickfileInput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTbTickData
    OptionCombo.Text = SpOptionTbTickData
    
    setupDbProperties
Case mSPs.SPProgIdFileTickData
    OptionCombo.Text = SpOptionFileTickData
    
    setupTfProperties
Case mSPs.SPProgIdQtTickData
    OptionCombo.Text = SpOptionQtTickData

    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionCustomTickData
    
    setupCustomProperties
End Select

End Sub

Private Sub setupTickfileOutputSP()
Dim progId As String

CategoryLabel = "Tickfile storage options"
OptionLabel = "Select tickfile data store"
OptionCombo.ComboItems.clear
OptionCombo.ComboItems.add , , SpOptionNotConfigured
OptionCombo.ComboItems.add , , SpOptionTbTickData
OptionCombo.ComboItems.add , , SpOptionFileTickData
OptionCombo.ComboItems.add , , SpOptionCustomTickData

On Error Resume Next
findSp mSPs.SPNameTickfileOutput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId, "")
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case mSPs.SPProgIdTbTickData
    OptionCombo.Text = SpOptionTbTickData
    
    setupDbProperties
Case mSPs.SPProgIdFileTickData
    OptionCombo.Text = SpOptionFileTickData
    
    setupTfProperties
Case Else
    OptionCombo.Text = SpOptionCustomTickData
    
    setupCustomProperties
End Select

End Sub

Private Sub setupTwsProperties()
Dim twsLogLevel As String
On Error Resume Next
TwsEnabledCheck.value = vbUnchecked
TwsKeepConnectionCheck.value = vbUnchecked
TWSServerText = ""
TWSPortText = ""
TWSClientIdText = ""
TwsProviderKeyText = ""
TwsConnectRetryIntervalText = ""
TwsEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled, "False") = "True", vbChecked, vbUnchecked)
TwsKeepConnectionCheck.value = IIf(getProperty(PropertyNameTwsKeepConnection) = "True", vbChecked, vbUnchecked)
TWSServerText = getProperty(PropertyNameTwsServer)
TWSPortText = getProperty(PropertyNameTwsPort)
TWSClientIdText = getProperty(PropertyNameTwsClientId)
TwsProviderKeyText = getProperty(PropertyNameTwsProviderKey)
TwsConnectRetryIntervalText = getProperty(PropertyNameTwsConnectionRetryInterval)
twsLogLevel = getProperty(PropertyNameTwsLogLevel)
If twsLogLevel = "" Then
    TwsLogLevelCombo.Text = TWSLogLevelSystem
Else
    TwsLogLevelCombo.Text = getProperty(PropertyNameTwsLogLevel)
End If
End Sub

Private Sub showSpOptions()
Select Case mCurrCategory
Case mSPs.SPNameRealtimeData
    Select Case OptionCombo.Text
    Case SpOptionQtRealtimeData
        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionTwsRealtimeData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomRealtimeData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNamePrimaryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionTwsContractData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomContractData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNameSecondryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionTwsContractData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomContractData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNameHistoricalDataInput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionQtBarData
        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionTwsBarData
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomBarData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNameHistoricalDataOutput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionCustomBarData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNameBrokerLive
    Select Case OptionCombo.Text
    Case SpOptionTwsOrders
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionCustomOrders
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNameBrokerSimulated
    Select Case OptionCombo.Text
    Case SpOptionTbOrders
        Set mCurrOptionsPic = BrOptionsPicture
    Case SpOptionCustomOrders
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNameTickfileInput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionQtTickData
        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionFileTickData
        Set mCurrOptionsPic = TfOptionsPicture
        TickfilePathText.Enabled = False
        PathChooserButton.Enabled = False
    Case SpOptionCustomTickData
        Set mCurrOptionsPic = CustomOptionsPicture
    End Select
Case mSPs.SPNameTickfileOutput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionFileTickData
        Set mCurrOptionsPic = TfOptionsPicture
        TickfilePathText.Enabled = True
        PathChooserButton.Enabled = True
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
End Sub

