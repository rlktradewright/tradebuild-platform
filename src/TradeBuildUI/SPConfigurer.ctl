VERSION 5.00
Begin VB.UserControl SPConfigurer 
   BackStyle       =   0  'Transparent
   ClientHeight    =   12750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16740
   ScaleHeight     =   12750
   ScaleWidth      =   16740
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   37
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox TfOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3600
      ScaleHeight     =   735
      ScaleWidth      =   3495
      TabIndex        =   36
      Top             =   6480
      Width           =   3495
      Begin VB.CheckBox TfEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox BrOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3600
      ScaleHeight     =   735
      ScaleWidth      =   3495
      TabIndex        =   35
      Top             =   5640
      Width           =   3495
      Begin VB.CheckBox BrEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton ApplyButton 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox OptionCombo 
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.PictureBox QtOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3600
      ScaleHeight     =   1455
      ScaleWidth      =   3495
      TabIndex        =   30
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox QtEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox QtPasswordText 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox QtPortText 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Text            =   "16240"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox QtServerText 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Password"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Port"
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Server"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox DbOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   3495
      TabIndex        =   24
      Top             =   5760
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox DbEnabledCheck 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox DbDatabaseText 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         ToolTipText     =   "Port for connecting to QuoteTracker"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox DbServerText 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         ToolTipText     =   "Name or address of computer hosting QuoteTracker"
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox DbTypeCombo 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox DbUsernameText 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         ToolTipText     =   "Port for connecting to QuoteTracker"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox DbPasswordText 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   10
         ToolTipText     =   "Port for connecting to QuoteTracker"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "Database"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Server"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "DB Type"
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Username"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Password"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.PictureBox TwsOptionsPicture 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   3495
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
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
         Width           =   2535
      End
      Begin VB.TextBox TWSPortText 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "7496"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TWSServerText 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Client id"
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Port"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Server"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ListBox CategoryList 
      Height          =   3765
      ItemData        =   "SPConfigurer.ctx":0000
      Left            =   120
      List            =   "SPConfigurer.ctx":0002
      TabIndex        =   18
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape OptionsBox 
      Height          =   2175
      Left            =   2880
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label OptionLabel 
      Height          =   615
      Left            =   2280
      TabIndex        =   34
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
      TabIndex        =   19
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
Private Const ModuleName                    As String = "SPConfigurer"

Private Const AccessModeReadOnly            As String = "Read only"
Private Const AccessModeWriteOnly           As String = "Write only"
Private Const AccessModeReadWrite           As String = "Read write"

Private Const AttributeNameConfigName       As String = "Name"
Private Const AttributeNameDefault          As String = "Default"
Private Const AttributeNameEnabled          As String = "Enabled"
Private Const AttributeNameLogLevel         As String = "LogLevel"
Private Const AttributeNamePropertyName     As String = "Name"
Private Const AttributeNamePropertyValue    As String = "Value"
Private Const AttributeNameServiceProviderEnabled As String = "Enabled"
Private Const AttributeNameServiceProviderName As String = "Name"
Private Const AttributeNameServiceProviderProgId As String = "ProgId"

Private Const CategoryRealtimeData          As String = "Realtime data"
Private Const CategoryPrimaryContractData   As String = "Primary contract data"
Private Const CategorySecondaryContractData As String = "Secondary contract data"
Private Const CategoryHistoricalDataInput   As String = "Historical bar data retrieval"
Private Const CategoryHistoricalDataOutput  As String = "Historical bar data storage"
Private Const CategoryBroker                As String = "Broker"
Private Const CategoryTickfileInput         As String = "Tickfile replay"
Private Const CategoryTickfileOutput        As String = "Tickfile storage"

Private Const ConfigNameProperties          As String = "Properties"
Private Const ConfigNameProperty            As String = "Property"
Private Const ConfigNameServiceProvider     As String = "ServiceProvider"
Private Const ConfigNameServiceProviders    As String = "ServiceProviders"

Private Const DbTypeMySql                   As String = "MySQL5"
Private Const DbTypeSqlServer7              As String = "SQL Server 7"
Private Const DbTypeSqlServer2000           As String = "SQL Server 2000"
Private Const DbTypeSqlServer2005           As String = "SQL Server 2005"

Private Const PropertyNameQtServer          As String = "Server"
Private Const PropertyNameQtPort            As String = "Port"
Private Const PropertyNameQtPassword        As String = "Password"

Private Const PropertyNameTbServer          As String = "Server"
Private Const PropertyNameTbDbType          As String = "Database Type"
Private Const PropertyNameTbDbName          As String = "Database Name"
Private Const PropertyNameTbUserName        As String = "User Name"
Private Const PropertyNameTbPassword        As String = "Password"
Private Const PropertyNameTbAccessMode      As String = "Access Mode"
Private Const PropertyNameTbRole            As String = "Role"

Private Const PropertyNameTwsServer         As String = "Server"
Private Const PropertyNameTwsPort           As String = "Port"
Private Const PropertyNameTwsClientId       As String = "Client Id"

Private Const ProgIdQtBarData               As String = "QTSP26.QTHistDataServiceProvider"
Private Const ProgIdQtRealtimeData          As String = "QTSP26.QTRealtimeDataServiceProvider"
Private Const ProgIdQtTickData              As String = "QTSP26.QTTickfileServiceProvider"

Private Const ProgIdTbBarData               As String = "TBInfoBase26.HistDataServiceProvider"
Private Const ProgIdTbContractData          As String = "TBInfoBase26.ContractInfoSrvcProvider"
Private Const ProgIdTbOrders                As String = ""
Private Const ProgIdTbTickData              As String = "TBInfoBase26.TickfileServiceProvider"

Private Const ProgIdFileTickData            As String = "TickfileSP26.TickfileServiceProvider"

Private Const ProgIdTwsBarData              As String = "IBTWSSP26.HistDataServiceProvider"
Private Const ProgIdTwsContractData         As String = "IBTWSSP26.ContractInfoServiceProvider"
Private Const ProgIdTwsOrders               As String = "IBTWSSP26.OrderSubmissionSrvcProvider"
Private Const ProgIdTwsRealtimeData         As String = "IBTWSSP26.RealtimeDataServiceProvider"

Private Const RolePrimary                   As String = "Primary"
Private Const RoleSecondary                 As String = "Secondary"

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

'@================================================================================
' Member variables
'@================================================================================

Private mCurrOptionsPic             As PictureBox

Private mConfig                     As ConfigItem

Private mCurrSPsList                As ConfigItem
Private mCurrSP                     As ConfigItem
Private mCurrProps                  As ConfigItem
Private mCurrCategory               As String
Private mCurrSpOption               As String

Private mPermittedSPs               As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height

DbTypeCombo.addItem DbTypeMySql
DbTypeCombo.addItem DbTypeSqlServer7
DbTypeCombo.addItem DbTypeSqlServer2000
DbTypeCombo.addItem DbTypeSqlServer2005
End Sub

Private Sub UserControl_LostFocus()
checkForOutstandingUpdates
End Sub

Private Sub UserControl_Resize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ApplyButton_Click()
applyProperties
ApplyButton.Enabled = False
CancelButton.Enabled = False
End Sub

Private Sub BrEnabledCheck_Click()
ApplyButton.Enabled = True
CancelButton.Enabled = True
End Sub

Private Sub CancelButton_Click()
Dim index As Long
ApplyButton.Enabled = False
CancelButton.Enabled = False
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
Case CategoryRealtimeData
    setupRealtimeDataSP
Case CategoryPrimaryContractData
    setupPrimaryContractDataSP
Case CategorySecondaryContractData
    setupSecondaryContractDataSP
Case CategoryHistoricalDataInput
    setupHistoricalDataInputSP
Case CategoryHistoricalDataOutput
    setupHistoricalDataOutputSP
Case CategoryBroker
    setupBrokerSP
Case CategoryTickfileInput
    setupTickfileInputSP
Case CategoryTickfileOutput
    setupTickfileOutputSP
End Select

showSpOptions

ApplyButton.Enabled = False
CancelButton.Enabled = False
End Sub

Private Sub DbDatabaseText_Change()
ApplyButton.Enabled = isValidDbProperties
CancelButton.Enabled = True
End Sub

Private Sub DbEnabledCheck_Click()
ApplyButton.Enabled = isValidDbProperties
CancelButton.Enabled = True
End Sub

Private Sub DbPasswordText_Change()
ApplyButton.Enabled = isValidDbProperties
CancelButton.Enabled = True
End Sub

Private Sub DbServerText_Change()
ApplyButton.Enabled = isValidDbProperties
CancelButton.Enabled = True
End Sub

Private Sub DbTypeCombo_Click()
ApplyButton.Enabled = isValidDbProperties
CancelButton.Enabled = True
End Sub

Private Sub DbUsernameText_Change()
ApplyButton.Enabled = isValidDbProperties
CancelButton.Enabled = True
End Sub

Private Sub OptionCombo_Click()
hideSpOptions
If OptionCombo.Text = SpOptionNotConfigured Then
    If Not mCurrSP Is Nothing Then ApplyButton.Enabled = True
Else
    showSpOptions
    If mCurrSP Is Nothing Or OptionCombo.Text <> mCurrSpOption Then
        If mCurrOptionsPic Is DbOptionsPicture Then
            ApplyButton.Enabled = isValidDbProperties
        ElseIf mCurrOptionsPic Is QtOptionsPicture Then
            ApplyButton.Enabled = isValidQtProperties
        ElseIf mCurrOptionsPic Is TwsOptionsPicture Then
            ApplyButton.Enabled = isValidTwsProperties
        ElseIf mCurrOptionsPic Is BrOptionsPicture Then
            ApplyButton.Enabled = True
        ElseIf mCurrOptionsPic Is TfOptionsPicture Then
            ApplyButton.Enabled = True
        End If
    End If
End If
mCurrSpOption = OptionCombo.Text
CancelButton.Enabled = True
End Sub

Private Sub QtEnabledCheck_Click()
ApplyButton.Enabled = isValidQtProperties
CancelButton.Enabled = True
End Sub

Private Sub QtPasswordText_Change()
ApplyButton.Enabled = isValidQtProperties
CancelButton.Enabled = True
End Sub

Private Sub QtPortText_Change()
ApplyButton.Enabled = isValidQtProperties
CancelButton.Enabled = True
End Sub

Private Sub QtServerText_Change()
ApplyButton.Enabled = isValidQtProperties
CancelButton.Enabled = True
End Sub

Private Sub TfEnabledCheck_Click()
ApplyButton.Enabled = True
CancelButton.Enabled = True
End Sub

Private Sub TWSClientIdText_Change()
ApplyButton.Enabled = isValidTwsProperties
CancelButton.Enabled = True
End Sub

Private Sub TwsEnabledCheck_Click()
ApplyButton.Enabled = isValidTwsProperties
CancelButton.Enabled = True
End Sub

Private Sub TWSPortText_Change()
ApplyButton.Enabled = isValidTwsProperties
CancelButton.Enabled = True
End Sub

Private Sub TWSServerText_Change()
ApplyButton.Enabled = isValidTwsProperties
CancelButton.Enabled = True
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
ApplyButton.Enabled = False
CancelButton.Enabled = False
End Sub

Public Sub initialise( _
                ByVal configdata As ConfigItem, _
                ByVal permittedSPs As Long)
checkForOutstandingUpdates
mPermittedSPs = permittedSPs

Set mCurrSPsList = Nothing
Set mCurrSP = Nothing
Set mCurrProps = Nothing
mCurrCategory = ""
mCurrSpOption = ""

loadConfig configdata
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

Private Sub applyDbProperties()
If DbEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
End If
setProperty PropertyNameTbServer, DbServerText
setProperty PropertyNameTbDbType, DbTypeCombo
setProperty PropertyNameTbDbName, DbDatabaseText
setProperty PropertyNameTbUserName, DbUsernameText
setProperty PropertyNameTbPassword, DbPasswordText

If mCurrCategory = CategoryHistoricalDataInput Or _
    mCurrCategory = CategoryTickfileInput _
Then
    setProperty PropertyNameTbAccessMode, AccessModeReadOnly
End If

If mCurrCategory = CategoryHistoricalDataOutput Or _
    mCurrCategory = CategoryTickfileOutput _
Then
    setProperty PropertyNameTbAccessMode, AccessModeWriteOnly
End If

If mCurrCategory = CategoryPrimaryContractData Then
    setProperty PropertyNameTbRole, RolePrimary
End If

If mCurrCategory = CategorySecondaryContractData Then
    setProperty PropertyNameTbRole, RoleSecondary
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

If mCurrOptionsPic Is DbOptionsPicture Then
    applyDbProperties
ElseIf mCurrOptionsPic Is QtOptionsPicture Then
    applyQtProperties
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
setProperty PropertyNameQtServer, QtServerText
setProperty PropertyNameQtPort, QtPortText
setProperty PropertyNameQtPassword, QtPasswordText
End Sub

Private Sub applyTfProperties()
If TfEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
End If
If mCurrCategory = CategoryTickfileInput Then
    setProperty PropertyNameTbAccessMode, AccessModeReadOnly
End If
If mCurrCategory = CategoryTickfileOutput Then
    setProperty PropertyNameTbAccessMode, AccessModeWriteOnly
End If
End Sub

Private Sub applyTwsProperties()
If TwsEnabledCheck = vbChecked Then
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "True"
Else
    mCurrSP.setAttribute AttributeNameServiceProviderEnabled, "False"
End If
setProperty PropertyNameTwsServer, TWSServerText
setProperty PropertyNameTwsPort, TWSPortText
setProperty PropertyNameTwsClientId, TWSClientIdText
End Sub

Private Sub checkForOutstandingUpdates()
If ApplyButton.Enabled Then
    If MsgBox("Do you want to apply the changes you have made?", _
            vbExclamation Or vbYesNoCancel) = vbYes Then
        applyProperties
        ApplyButton.Enabled = False
        CancelButton.Enabled = False
    End If
End If
End Sub

Private Sub clearProperties()
mCurrProps.childItems.clear
End Sub

Private Sub createNewSp()
Set mCurrSP = mCurrSPsList.childItems.addItem(ConfigNameServiceProvider)
mCurrSP.setAttribute AttributeNameServiceProviderName, mCurrCategory
Set mCurrProps = mCurrSP.childItems.addItem(ConfigNameProperties)
End Sub

Private Sub deleteSp()
mCurrSPsList.childItems.remove mCurrSP
Set mCurrSP = Nothing
Set mCurrProps = Nothing
End Sub

Private Function findProperty( _
                ByVal name As String) As ConfigItem
Dim prop As ConfigItem
For Each prop In mCurrProps.childItems
    If prop.getAttribute(AttributeNameServiceProviderName) = name Then
        Set findProperty = prop
        Exit Function
    End If
Next
End Function

Private Function findSp( _
                ByVal name As String) As Boolean
Dim sp As ConfigItem
Set mCurrSP = Nothing
Set mCurrProps = Nothing
mCurrSpOption = ""
For Each sp In mCurrSPsList.childItems
    If sp.getAttribute(AttributeNameServiceProviderName) = name Then
        Set mCurrSP = sp
        Set mCurrProps = mCurrSP.childItems.item(ConfigNameProperties)
        findSp = True
        Exit Function
    End If
Next
End Function

Private Sub hideSpOptions()
If Not mCurrOptionsPic Is Nothing Then
    mCurrOptionsPic.Visible = False
    Set mCurrOptionsPic = Nothing
End If
End Sub

Private Function isValidDbProperties() As Boolean
If DbDatabaseText = "" Then
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
Else
    isValidTwsProperties = True
End If
End Function

Private Sub loadConfig( _
                ByVal configdata As ConfigItem)
                
Dim sp As ConfigItem

Set mConfig = configdata

On Error Resume Next
Set mCurrSPsList = mConfig.childItems.item(ConfigNameServiceProviders)
On Error GoTo 0

If mCurrSPsList Is Nothing Then
    Set mCurrSPsList = mConfig.childItems.addItem(ConfigNameServiceProviders)
End If

CategoryList.clear

If mPermittedSPs And PermittedServiceProviders.SPRealtimeData Then
    CategoryList.addItem CategoryRealtimeData
End If
If mPermittedSPs And PermittedServiceProviders.SPPrimaryContractData Then
    CategoryList.addItem CategoryPrimaryContractData
End If
If mPermittedSPs And PermittedServiceProviders.SPSecondaryContractData Then
    CategoryList.addItem CategorySecondaryContractData
End If
If mPermittedSPs And PermittedServiceProviders.SPHistoricalDataInput Then
    CategoryList.addItem CategoryHistoricalDataInput
End If
If mPermittedSPs And PermittedServiceProviders.SPHistoricalDataOutput Then
    CategoryList.addItem CategoryHistoricalDataOutput
End If
If mPermittedSPs And PermittedServiceProviders.SPBroker Then
    CategoryList.addItem CategoryBroker
End If
If mPermittedSPs And PermittedServiceProviders.SPTickfileInput Then
    CategoryList.addItem CategoryTickfileInput
End If
If mPermittedSPs And PermittedServiceProviders.SPTickfileOutput Then
    CategoryList.addItem CategoryTickfileOutput
End If

If CategoryList.ListCount > 0 Then CategoryList.ListIndex = 0

End Sub

Private Sub setProgId()
Dim progId As String

If CategoryList.ListIndex = -1 Then Exit Sub

Select Case mCurrCategory
Case CategoryRealtimeData
    Select Case OptionCombo.Text
    Case SpOptionQtRealtimeData
        progId = ProgIdQtRealtimeData
    Case SpOptionTwsRealtimeData
        progId = ProgIdTwsRealtimeData
    End Select
Case CategoryPrimaryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        progId = ProgIdTbContractData
    Case SpOptionTwsContractData
        progId = ProgIdTwsContractData
    End Select
Case CategorySecondaryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        progId = ProgIdTbContractData
    Case SpOptionTwsContractData
        progId = ProgIdTwsContractData
    End Select
Case CategoryHistoricalDataInput
    Select Case OptionCombo.Text
    Case SpOptionQtBarData
        progId = ProgIdQtBarData
    Case SpOptionTbBarData
        progId = ProgIdTbBarData
    Case SpOptionTwsBarData
        progId = ProgIdTwsBarData
    End Select
Case CategoryHistoricalDataOutput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        progId = ProgIdTbBarData
    End Select
Case CategoryBroker
    Select Case OptionCombo.Text
    Case SpOptionTbOrders
        progId = ProgIdTbOrders
    Case SpOptionTwsOrders
        progId = ProgIdTwsOrders
    End Select
Case CategoryTickfileInput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        progId = ProgIdTbTickData
    Case SpOptionQtTickData
        progId = ProgIdQtTickData
    Case SpOptionFileTickData
        progId = ProgIdFileTickData
    End Select
Case CategoryTickfileOutput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        progId = ProgIdTbTickData
    Case SpOptionFileTickData
        progId = ProgIdFileTickData
    End Select
End Select

mCurrSP.setAttribute AttributeNameServiceProviderProgId, progId

End Sub

Private Sub setProperty( _
                ByVal name As String, _
                ByVal value As String)
Dim prop As ConfigItem
Set prop = mCurrProps.childItems.addItem(ConfigNameProperty)
prop.setAttribute AttributeNamePropertyName, name
prop.setAttribute AttributeNamePropertyValue, value
End Sub

Private Sub setupBrProperties()
On Error Resume Next
BrEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled) = "True", vbChecked, vbUnchecked)
End Sub

Private Sub setupBrokerSP()
Dim progId As String
    
CategoryLabel = "Broker options"
OptionLabel = "Select broker"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionTbOrders
OptionCombo.addItem SpOptionTwsOrders

On Error Resume Next
findSp CategoryBroker
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTbOrders
    OptionCombo.Text = SpOptionTbOrders
    
    setupBrProperties
Case ProgIdTwsOrders
    OptionCombo.Text = SpOptionTwsOrders
    
    setupTwsProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupDbProperties()
On Error Resume Next
DbEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled) = "True", vbChecked, vbUnchecked)
DbServerText = findProperty(PropertyNameTbServer).getAttribute(AttributeNamePropertyValue)
DbTypeCombo = findProperty(PropertyNameTbDbType).getAttribute(AttributeNamePropertyValue)
DbDatabaseText = findProperty(PropertyNameTbDbName).getAttribute(AttributeNamePropertyValue)
DbUsernameText = findProperty(PropertyNameTbUserName).getAttribute(AttributeNamePropertyValue)
DbPasswordText = findProperty(PropertyNameTbPassword).getAttribute(AttributeNamePropertyValue)
End Sub

Private Sub setupHistoricalDataInputSP()
Dim progId As String
    
CategoryLabel = "Historical bar data retrieval source options"
OptionLabel = "Select historical bar data source"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionTbBarData
OptionCombo.addItem SpOptionQtBarData
OptionCombo.addItem SpOptionTwsBarData

On Error Resume Next
findSp CategoryHistoricalDataInput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTwsBarData
    OptionCombo.Text = SpOptionTwsBarData
    
    setupTwsProperties
Case ProgIdTbBarData
    OptionCombo.Text = SpOptionTbBarData
    
    setupDbProperties
Case ProgIdQtBarData
    OptionCombo.Text = SpOptionQtBarData
    
    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupHistoricalDataOutputSP()
Dim progId As String
    
CategoryLabel = "Historical bar data storage options"
OptionLabel = "Select historical bar data source"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionTbBarData

On Error Resume Next
findSp CategoryHistoricalDataOutput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTbBarData
    OptionCombo.Text = SpOptionTbBarData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupPrimaryContractDataSP()
Dim progId As String
    
CategoryLabel = "Primary contract data source options"
OptionLabel = "Select primary contract data source"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionTbContractData
OptionCombo.addItem SpOptionTwsContractData

On Error Resume Next
findSp CategoryPrimaryContractData
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTwsContractData
    OptionCombo.Text = SpOptionTwsContractData
    
    setupTwsProperties
Case ProgIdTbContractData
    OptionCombo.Text = SpOptionTbContractData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupRealtimeDataSP()
Dim progId As String

CategoryLabel = "Realtime data source options"
OptionLabel = "Select realtime data source"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionQtRealtimeData
OptionCombo.addItem SpOptionTwsRealtimeData

On Error Resume Next
findSp CategoryRealtimeData
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTwsRealtimeData
    OptionCombo.Text = SpOptionTwsRealtimeData
    
    setupTwsProperties
Case ProgIdQtRealtimeData
    OptionCombo.Text = SpOptionQtRealtimeData

    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupQtProperties()
On Error Resume Next
QtEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled) = "True", vbChecked, vbUnchecked)
QtServerText = findProperty(PropertyNameQtServer).getAttribute(AttributeNamePropertyValue)
QtPortText = findProperty(PropertyNameQtPort).getAttribute(AttributeNamePropertyValue)
QtPasswordText = findProperty(PropertyNameQtPassword).getAttribute(AttributeNamePropertyValue)
End Sub

Private Sub setupSecondaryContractDataSP()
Dim progId As String
    
CategoryLabel = "Secondary contract data source options"
OptionLabel = "Select secondary contract data source"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionTbContractData
OptionCombo.addItem SpOptionTwsContractData

On Error Resume Next
findSp CategorySecondaryContractData
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTwsContractData
    OptionCombo.Text = SpOptionTwsContractData
    
    setupTwsProperties
Case ProgIdTbContractData
    OptionCombo.Text = SpOptionTbContractData
    
    setupDbProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupTfProperties()
On Error Resume Next
TfEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled) = "True", vbChecked, vbUnchecked)
End Sub

Private Sub setupTickfileInputSP()
Dim progId As String

CategoryLabel = "Tickfile replay data source options"
OptionLabel = "Select tickfile replay data source"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionTbTickData
OptionCombo.addItem SpOptionFileTickData
OptionCombo.addItem SpOptionQtTickData

On Error Resume Next
findSp CategoryTickfileInput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTbTickData
    OptionCombo.Text = SpOptionTbTickData
    
    setupDbProperties
Case ProgIdFileTickData
    OptionCombo.Text = SpOptionFileTickData
    
    setupTfProperties
Case ProgIdQtTickData
    OptionCombo.Text = SpOptionQtTickData

    setupQtProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupTickfileOutputSP()
Dim progId As String

CategoryLabel = "Tickfile storage options"
OptionLabel = "Select tickfile data store"
OptionCombo.clear
OptionCombo.addItem SpOptionNotConfigured
OptionCombo.addItem SpOptionTbTickData
OptionCombo.addItem SpOptionFileTickData

On Error Resume Next
findSp CategoryTickfileOutput
progId = mCurrSP.getAttribute(AttributeNameServiceProviderProgId)
On Error GoTo 0

If mCurrSP Is Nothing Then
    OptionCombo.Text = SpOptionNotConfigured
    Exit Sub
End If

Select Case progId
Case ProgIdTbTickData
    OptionCombo.Text = SpOptionTbTickData
    
    setupDbProperties
Case ProgIdFileTickData
    OptionCombo.Text = SpOptionFileTickData
    
    setupTfProperties
Case Else
    OptionCombo.Text = SpOptionNotConfigured
End Select

End Sub

Private Sub setupTwsProperties()
On Error Resume Next
TwsEnabledCheck.value = IIf(mCurrSP.getAttribute(AttributeNameServiceProviderEnabled) = "True", vbChecked, vbUnchecked)
TWSServerText = findProperty(PropertyNameTwsServer).getAttribute(AttributeNamePropertyValue)
TWSPortText = findProperty(PropertyNameTwsPort).getAttribute(AttributeNamePropertyValue)
TWSClientIdText = findProperty(PropertyNameTwsClientId).getAttribute(AttributeNamePropertyValue)
End Sub

Private Sub showSpOptions()
Select Case mCurrCategory
Case CategoryRealtimeData
    Select Case OptionCombo.Text
    Case SpOptionQtRealtimeData
        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionTwsRealtimeData
        Set mCurrOptionsPic = TwsOptionsPicture
    End Select
Case CategoryPrimaryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionTwsContractData
        Set mCurrOptionsPic = TwsOptionsPicture
    End Select
Case CategorySecondaryContractData
    Select Case OptionCombo.Text
    Case SpOptionTbContractData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionTwsContractData
        Set mCurrOptionsPic = TwsOptionsPicture
    End Select
Case CategoryHistoricalDataInput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionQtBarData
        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionTwsBarData
        Set mCurrOptionsPic = TwsOptionsPicture
    End Select
Case CategoryHistoricalDataOutput
    Select Case OptionCombo.Text
    Case SpOptionTbBarData
        Set mCurrOptionsPic = DbOptionsPicture
    End Select
Case CategoryBroker
    Select Case OptionCombo.Text
    Case SpOptionTwsOrders
        Set mCurrOptionsPic = TwsOptionsPicture
    Case SpOptionTbOrders
        Set mCurrOptionsPic = BrOptionsPicture
    End Select
Case CategoryTickfileInput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionQtTickData
        Set mCurrOptionsPic = QtOptionsPicture
    Case SpOptionFileTickData
        Set mCurrOptionsPic = TfOptionsPicture
    End Select
Case CategoryTickfileOutput
    Select Case OptionCombo.Text
    Case SpOptionTbTickData
        Set mCurrOptionsPic = DbOptionsPicture
    Case SpOptionFileTickData
        Set mCurrOptionsPic = TfOptionsPicture
    End Select
Case CategoryTickfileOutput
End Select

If Not mCurrOptionsPic Is Nothing Then
    mCurrOptionsPic.Left = OptionsBox.Left
    mCurrOptionsPic.Top = OptionsBox.Top
    mCurrOptionsPic.Visible = True
End If
End Sub

