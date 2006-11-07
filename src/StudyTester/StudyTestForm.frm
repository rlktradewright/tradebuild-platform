VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DBED8E43-5960-49DE-B9A7-BBC22DB93A26}#12.1#0"; "ChartSkil.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D1E1CD3C-084A-4A4F-B2D9-56CE3669B04D}#10.0#0"; "TradeBuildUI.ocx"
Begin VB.Form StudyTestForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TradeBuild Study Test Harness"
   ClientHeight    =   10365
   ClientLeft      =   5070
   ClientTop       =   3540
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   14130
   Begin VB.CommandButton TestButton 
      Caption         =   "Test"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   29
      ToolTipText     =   "Test the study"
      Top             =   120
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   30
      Top             =   960
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   16325
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   2
      TabCaption(0)   =   "&Test data and results"
      TabPicture(0)   =   "StudyTestForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TestDataGrid"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TestDataFilenameText"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FindFileButton"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Study setup"
      TabPicture(1)   =   "StudyTestForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(3)=   "Label19"
      Tab(1).Control(4)=   "StudyConfigurer1"
      Tab(1).Control(5)=   "StudiesCombo"
      Tab(1).Control(6)=   "ServiceProviderClassNameText"
      Tab(1).Control(7)=   "BuiltInStudiesCheck"
      Tab(1).Control(8)=   "SPToAddText"
      Tab(1).Control(9)=   "AddSPButton"
      Tab(1).Control(10)=   "SpList"
      Tab(1).Control(11)=   "RemoveSPButton"
      Tab(1).Control(12)=   "SetSpButton"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "&Contract setup"
      TabPicture(2)   =   "StudyTestForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "SetContractButton"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "C&hart"
      TabPicture(3)   =   "StudyTestForm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Chart1"
      Tab(3).ControlCount=   1
      Begin ChartSkil.Chart Chart1 
         Height          =   8775
         Left            =   -74880
         TabIndex        =   57
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   15478
         autoscale       =   0   'False
      End
      Begin VB.CommandButton SetContractButton 
         Caption         =   "Set"
         Height          =   375
         Left            =   -66360
         TabIndex        =   28
         Top             =   4320
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Contract details"
         Height          =   3735
         Left            =   -71400
         TabIndex        =   46
         Top             =   480
         Width           =   5895
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   3375
            Left            =   120
            ScaleHeight     =   3375
            ScaleWidth      =   5655
            TabIndex        =   47
            Top             =   240
            Width           =   5655
            Begin VB.TextBox TradingClassText 
               Height          =   285
               Left            =   1440
               TabIndex        =   27
               Top             =   2880
               Width           =   1335
            End
            Begin VB.TextBox SessionStartTimeText 
               Height          =   285
               Left            =   1440
               TabIndex        =   26
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox SessionEndTimeText 
               Height          =   285
               Left            =   1440
               TabIndex        =   25
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox MultiplierText 
               Height          =   285
               Left            =   1440
               TabIndex        =   24
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox MinimumTickText 
               Height          =   285
               Left            =   1440
               TabIndex        =   23
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox MarketNameText 
               Height          =   285
               Left            =   1440
               TabIndex        =   22
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox ExpiryDateText 
               Height          =   285
               Left            =   1440
               TabIndex        =   21
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox DescriptionText 
               Height          =   285
               Left            =   1440
               TabIndex        =   20
               Top             =   360
               Width           =   4095
            End
            Begin VB.TextBox ContractIdText 
               Height          =   285
               Left            =   1440
               TabIndex        =   19
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label18 
               Caption         =   "Trading class"
               Height          =   255
               Left            =   0
               TabIndex        =   56
               Top             =   2880
               Width           =   1335
            End
            Begin VB.Label Label16 
               Caption         =   "Session start time"
               Height          =   255
               Left            =   0
               TabIndex        =   55
               Top             =   2520
               Width           =   1335
            End
            Begin VB.Label Label15 
               Caption         =   "Session end time"
               Height          =   255
               Left            =   0
               TabIndex        =   54
               Top             =   2160
               Width           =   1335
            End
            Begin VB.Label Label14 
               Caption         =   "Multiplier"
               Height          =   255
               Left            =   0
               TabIndex        =   53
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label13 
               Caption         =   "Minimum tick"
               Height          =   255
               Left            =   0
               TabIndex        =   52
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Market name"
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Expiry date"
               Height          =   255
               Left            =   0
               TabIndex        =   50
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "Description"
               Height          =   255
               Left            =   0
               TabIndex        =   49
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label9 
               Caption         =   "Contract id"
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contract specifier"
         Height          =   3735
         Left            =   -74640
         TabIndex        =   36
         Top             =   480
         Width           =   3015
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   120
            ScaleHeight     =   3255
            ScaleWidth      =   2775
            TabIndex        =   37
            Top             =   240
            Width           =   2775
            Begin VB.ComboBox RightCombo 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   2520
               Width           =   855
            End
            Begin VB.ComboBox TypeCombo 
               Height          =   315
               ItemData        =   "StudyTestForm.frx":0070
               Left            =   1440
               List            =   "StudyTestForm.frx":0072
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   705
               Width           =   1335
            End
            Begin VB.TextBox SymbolText 
               Height          =   285
               Left            =   1440
               TabIndex        =   12
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox ExpiryText 
               Height          =   285
               Left            =   1440
               TabIndex        =   14
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox StrikePriceText 
               Height          =   285
               Left            =   1440
               TabIndex        =   17
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox CurrencyText 
               Height          =   285
               Left            =   1440
               TabIndex        =   16
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox LocalSymbolText 
               Height          =   285
               Left            =   1440
               TabIndex        =   11
               Top             =   0
               Width           =   1335
            End
            Begin VB.TextBox ExchangeText 
               Height          =   285
               Left            =   1440
               TabIndex        =   15
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label21 
               Caption         =   "Right"
               Height          =   255
               Left            =   0
               TabIndex        =   45
               Top             =   2520
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "Strike price"
               Height          =   255
               Left            =   0
               TabIndex        =   44
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label Label8 
               Caption         =   "Symbol"
               Height          =   255
               Left            =   0
               TabIndex        =   43
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label7 
               Caption         =   "Type"
               Height          =   255
               Left            =   0
               TabIndex        =   42
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label5 
               Caption         =   "Expiry"
               Height          =   255
               Left            =   0
               TabIndex        =   41
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "Exchange"
               Height          =   255
               Left            =   0
               TabIndex        =   40
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label26 
               Caption         =   "Currency"
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label Label29 
               Caption         =   "Short name"
               Height          =   255
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Width           =   855
            End
         End
      End
      Begin VB.CommandButton SetSpButton 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   3
         ToolTipText     =   "Click to load your service provider"
         Top             =   540
         Width           =   855
      End
      Begin VB.CommandButton RemoveSPButton 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   9
         ToolTipText     =   "Remove the selected service provider from the list"
         Top             =   2340
         Width           =   855
      End
      Begin VB.ListBox SpList 
         Height          =   840
         ItemData        =   "StudyTestForm.frx":0074
         Left            =   -72600
         List            =   "StudyTestForm.frx":0076
         TabIndex        =   8
         ToolTipText     =   "Lists all studies service providers you need (except the built-in studies service provider)"
         Top             =   2340
         Width           =   3975
      End
      Begin VB.CommandButton AddSPButton 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   7
         ToolTipText     =   "Add this service provider to the list"
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox SPToAddText 
         Height          =   285
         Left            =   -72600
         TabIndex        =   6
         ToolTipText     =   "Enter the program id of any other studies service provider your service provider needs"
         Top             =   1860
         Width           =   3975
      End
      Begin VB.CheckBox BuiltInStudiesCheck 
         Caption         =   "Include Built-In Studies"
         Height          =   255
         Left            =   -72600
         TabIndex        =   5
         ToolTipText     =   "Set if your service provider uses any built-in studies"
         Top             =   1500
         Width           =   3015
      End
      Begin VB.TextBox ServiceProviderClassNameText 
         Height          =   285
         Left            =   -72600
         TabIndex        =   2
         ToolTipText     =   "Enter your service provider's program id in the form project.class"
         Top             =   540
         Width           =   3975
      End
      Begin VB.ComboBox StudiesCombo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72600
         TabIndex        =   4
         ToolTipText     =   "Select the study to test"
         Top             =   1020
         Width           =   3975
      End
      Begin VB.CommandButton FindFileButton 
         Caption         =   "..."
         Height          =   285
         Left            =   6720
         TabIndex        =   1
         ToolTipText     =   "Click to browse for the test data file"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox TestDataFilenameText 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "The file that contains the test data"
         Top             =   780
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid TestDataGrid 
         Height          =   7815
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1260
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   13785
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorBkg    =   -2147483636
         Appearance      =   0
      End
      Begin TradeBuildUI.StudyConfigurer StudyConfigurer1 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   10
         Top             =   3480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9975
      End
      Begin VB.Label Label19 
         Caption         =   "Configure the study - selected output values will appear both on the chart and in the grid"
         Height          =   375
         Left            =   -74760
         TabIndex        =   58
         Top             =   3240
         Width           =   11655
      End
      Begin VB.Label Label1 
         Caption         =   "Other study service providers to include"
         Height          =   615
         Left            =   -74760
         TabIndex        =   35
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Test data file"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Program id for Service Provider under test"
         Height          =   375
         Left            =   -74760
         TabIndex        =   33
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Study to test"
         Height          =   375
         Left            =   -74760
         TabIndex        =   32
         Top             =   1260
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "StudyTestForm"
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

' pretend to be a service provider, so that we can gain access to the common
' service consumer in TradeBuild
Implements TradeBuildSP.ICommonServiceProvider

' pretend to be a contract object, so we can intercept contract info
' requests from the study under test
Implements TradeBuildSP.IContract
Implements TradeBuildSP.IContractSpecifier

' masquerade as the underlying study for the study under test
Implements TradeBuildSP.IStudy
                                   
' handle requests from the study under test
Implements TradeBuildSP.IStudyServiceConsumer
                                    
'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const CellBackColorOdd As Long = &HF8F8F8
Private Const CellBackColorEven As Long = &HEEEEEE

Private Const PriceRegionName As String = "$price"
Private Const VolumeRegionName As String = "$volume"

Private Const TestDataGridRowsInitial As Long = 50
Private Const TestDataGridRowsIncrement As Long = 25

Private Const ParamNameBarLength As String = "Bar length"
Private Const ParamNameBarUnits As String = "Bar units"

Private Const BarsValueClose As String = "Close"
Private Const BarsValueVolume As String = "Volume"

'================================================================================
' Enums
'================================================================================

Private Enum TestDataFileColumns
    Timestamp
    openValue
    highValue
    lowValue
    closeValue
    volume
End Enum

Private Enum TestDataGridColumns
    Timestamp
    openValue
    highValue
    lowValue
    closeValue
    volume
    StudyValue1
End Enum

' Character widths of the TestDataGrid columns
Private Enum TestDataGridColumnWidths
    TimeStampWidth = 19
    openValueWidth = 9
    highValueWidth = 9
    lowValueWidth = 9
    closeValueWidth = 9
    volumeWidth = 9
    StudyValue1Width = 20
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTB As TradeBuildAPI

Private mLetterWidth As Single
Private mDigitWidth As Single

Private mPriceRegion As ChartSkil.ChartRegion
Private mVolumeRegion As ChartSkil.ChartRegion

Private mBarSeries As ChartSkil.BarSeries
Private mChartBar As ChartSkil.Bar

Private mVolumeSeries As ChartSkil.DataPointSeries
Private mVolumePoint As ChartSkil.DataPoint

Private mStudy As IStudy
Private mStudyParams As parameters

Private mDataLoaded As Boolean
Private mStudySet As Boolean

Private mStudiesServiceProvider As TradeBuildSP.IStudyServiceProvider

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer

Private mStudyDefinition As TradeBuild.StudyDefinition
Private mStudyConfiguration As TradeBuildUI.StudyConfiguration

Private mMyStudyId As String

Private mContract As TradeBuildSP.IContract
Private mContractSpecifier As TradeBuildSP.IContractSpecifier

Private mBaseStudyDefinition As TradeBuild.StudyDefinition
Private mBaseStudyConfiguration As TradeBuildUI.StudyConfiguration
Private mBaseStudyConfigurations As TradeBuildUI.StudyConfigurations

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Load()
Dim widthString As String
Dim paramDef As StudyParameterDefinition
Dim valueDef As StudyValueDefinition

mName = "TradeBuild Study Test Harness"

widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = Me.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = Me.TextWidth(widthString) / Len(widthString)

setupTestDataGrid

Set mTB = New TradeBuildAPI

mTB.ServiceProviders.Add Me

Set mBaseStudyDefinition = New TradeBuild.StudyDefinition

mBaseStudyDefinition.Description = "Formats the price stream into Open/High/Low/Close bars of an appropriate length."
mBaseStudyDefinition.name = "Bars"
mBaseStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice

Set paramDef = mBaseStudyDefinition.StudyParameterDefinitions.Add("Period Length")
paramDef.name = "Period Length"
paramDef.Description = "Length of one bar"
paramDef.parameterType = StudyParameterTypes.ParameterTypeInteger

Set paramDef = mBaseStudyDefinition.StudyParameterDefinitions.Add("Period Units")
paramDef.name = "Period Units"
paramDef.Description = "The units in which Period length is measured."
paramDef.parameterType = StudyParameterTypes.ParameterTypeString

Set valueDef = mBaseStudyDefinition.StudyValueDefinitions.Add(BarsValueClose)
valueDef.name = BarsValueClose
valueDef.Description = "The latest underlying value"
valueDef.isDefault = True
valueDef.valueType = StudyValueTypes.ValueTypeDouble

Set valueDef = mBaseStudyDefinition.StudyValueDefinitions.Add(BarsValueVolume)
valueDef.name = BarsValueVolume
valueDef.Description = "The cumulative size associated with the latest underlying value (where relevant)"
valueDef.valueType = StudyValueTypes.ValueTypeInteger

mMyStudyId = GenerateGUIDString

initialiseChart

TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.AddItem OptionRightToString(OptionRights.OptCall)
RightCombo.AddItem OptionRightToString(OptionRights.OptPut)

End Sub

'================================================================================
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim details As TradeBuildSP.ServiceProviderDetails
With details
    .Comments = App.Comments
    .EXEName = App.EXEName
    .FileDescription = App.FileDescription
    .LegalCopyright = App.LegalCopyright
    .LegalTrademarks = App.LegalTrademarks
    .Path = App.Path
    .ProductName = App.ProductName
    .Vendor = App.CompanyName
    .VersionMajor = App.Major
    .VersionMinor = App.Minor
    .VersionRevision = App.Revision
End With
ICommonServiceProvider_Details = details
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal handle As Long)
Set mCommonServiceConsumer = commonServiceConsumer

End Sub

Private Property Let ICommonServiceProvider_LogLevel(ByVal RHS As TradeBuildSP.LogLevels)

End Property

Private Property Get ICommonServiceProvider_name() As String
ICommonServiceProvider_name = name
End Property

Private Property Let ICommonServiceProvider_name(ByVal RHS As String)
mName = RHS
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'================================================================================
' IContract Interface Members
'================================================================================

Private Function IContract_BarStartTime(ByVal Timestamp As Date, ByVal BarLength As Long) As Date
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to request BarStartTime"
IContract_BarStartTime = mContract.BarStartTime(Timestamp, BarLength)
End Function

Private Function IContract_Clone() As TradeBuildSP.IContract
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to clone contract"
Set IContract_Clone = mContract.Clone
End Function

Private Property Let IContract_ContractID(ByVal RHS As Long)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify ContractId"
End Property

Private Property Get IContract_ContractID() As Long
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to request ContractId"
IContract_ContractID = mContract.contractID
End Property

Private Property Get IContract_CurrentSessionEndTime() As Date
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to request CurrentSessionEndTime"
IContract_CurrentSessionEndTime = mContract.currentSessionEndTime
End Property

Private Property Get IContract_CurrentSessionStartTime() As Date
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to request CurrentSessionStartTime"
IContract_CurrentSessionStartTime = mContract.currentSessionStartTime
End Property

Private Property Let IContract_DaysBeforeExpiryToSwitch(ByVal RHS As Long)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify DaysBeforeExpiryToSwitch"
End Property

Private Property Get IContract_DaysBeforeExpiryToSwitch() As Long
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to request DaysBeforeExpiryToSwitch"
IContract_DaysBeforeExpiryToSwitch = mContract.daysBeforeExpiryToSwitch
End Property

Private Property Let IContract_Description(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Description"
End Property

Private Property Get IContract_Description() As String
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to request Description"
IContract_Description = mContract.Description
End Property

Private Property Let IContract_ExpiryDate(ByVal RHS As Date)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify ExpiryDate"
End Property

Private Property Get IContract_ExpiryDate() As Date
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to request ExpiryDate"
IContract_ExpiryDate = mContract.ExpiryDate
End Property

Private Sub IContract_FromXML(ByVal contractXML As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to load contract from XML"
End Sub

Private Sub IContract_GetSessionTimes(ByVal Timestamp As Date, SessionStartTime As Date, SessionEndTime As Date)
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to call GetSessionTimes"
mContract.GetSessionTimes Timestamp, SessionStartTime, SessionEndTime
End Sub

Private Function IContract_IsTimeInSession(ByVal Timestamp As Date) As Boolean
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to call GetSessionTimes"
IContract_IsTimeInSession = mContract.isTimeInSession(Timestamp)
End Function

Private Property Get IContract_Key() As String
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Key"
IContract_Key = mContract.Key
End Property

Private Property Let IContract_MarketName(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify MarketName"
End Property

Private Property Get IContract_MarketName() As String
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get MarketName"
IContract_MarketName = mContract.marketName
End Property

Private Property Let IContract_MinimumTick(ByVal RHS As Double)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify MinimumTick"
End Property

Private Property Get IContract_MinimumTick() As Double
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get MinimumTick"
IContract_MinimumTick = mContract.MinimumTick
End Property

Private Property Let IContract_Multiplier(ByVal RHS As Long)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Multiplier"
End Property

Private Property Get IContract_Multiplier() As Long
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Multiplier"
IContract_Multiplier = mContract.multiplier
End Property

Private Property Get IContract_NumberOfDecimals() As Long
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get NumberOfDecimals"
IContract_NumberOfDecimals = mContract.NumberOfDecimals
End Property

Private Function IContract_OffsetBarStartTime(ByVal Timestamp As Date, ByVal BarLength As Long, ByVal Offset As Long) As Date
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get offsetBarStartTime"
IContract_OffsetBarStartTime = mContract.offsetBarStartTime(Timestamp, BarLength, Offset)
End Function

Private Property Get IContract_ProviderID(ByVal providerKey As String) As String
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get ProviderID"
IContract_ProviderID = mContract.ProviderID(providerKey)
End Property

Private Property Let IContract_ProviderIDs(RHS() As TradeBuildSP.DictionaryEntry)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify ProviderIDs"
End Property

Private Property Let IContract_SessionEndTime(ByVal RHS As Date)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify SessionEndTime"
End Property

Private Property Get IContract_SessionEndTime() As Date
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get SessionEndTime"
IContract_SessionEndTime = mContract.SessionEndTime
End Property

Private Property Let IContract_SessionStartTime(ByVal RHS As Date)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify SessionStartTime"
End Property

Private Property Get IContract_SessionStartTime() As Date
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get SessionEndTime"
IContract_SessionStartTime = mContract.SessionStartTime
End Property

Private Sub IContract_SetSession(ByVal Timestamp As Date)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to call SetSession"
End Sub

Private Property Let IContract_Specifier(ByVal RHS As TradeBuildSP.IContractSpecifier)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to set Specifier"
End Property

Private Property Get IContract_Specifier() As TradeBuildSP.IContractSpecifier
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get specifier"
Set IContract_Specifier = Me
End Property

Private Function IContract_ToString() As String
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to call ToString"
IContract_ToString = mContract.ToString
End Function

Private Function IContract_ToXML() As String
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to call ToString"
IContract_ToXML = mContract.ToXML
End Function

Private Property Let IContract_TradingClass(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify TradingClass"
End Property

Private Property Get IContract_TradingClass() As String
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get TradingClass"
IContract_TradingClass = mContract.tradingClass
End Property

Private Property Let IContract_ValidExchanges(RHS() As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify ValidExchanges"
End Property

Private Property Get IContract_ValidExchanges() As String()
If mContract Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get validExchanges"
IContract_ValidExchanges = mContract.validExchanges
End Property

'================================================================================
' IContractSpecifier Interface Members
'================================================================================

Private Function IContractSpecifier_Clone() As TradeBuildSP.IContractSpecifier
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to clone contractSpecifier"
Set IContractSpecifier_Clone = mContractSpecifier.Clone
End Function

Private Property Get IContractSpecifier_ComboLegs() As TradeBuildSP.IComboLegs
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get ComboLegs"
Set IContractSpecifier_ComboLegs = mContractSpecifier.ComboLegs
End Property

Private Property Let IContractSpecifier_ComboLegs(ByVal RHS As TradeBuildSP.IComboLegs)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify ComboLegs"
End Property

Private Property Let IContractSpecifier_CurrencyCode(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify CurrencyCode"
End Property

Private Property Get IContractSpecifier_CurrencyCode() As String
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get CurrencyCode"
IContractSpecifier_CurrencyCode = mContractSpecifier.currencyCode
End Property

Private Function IContractSpecifier_Equals(ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier) As Boolean
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to call Equals"
IContractSpecifier_Equals = mContractSpecifier.Equals(pContractSpecifier)
End Function

Private Property Let IContractSpecifier_Exchange(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Exchange"
End Property

Private Property Get IContractSpecifier_Exchange() As String
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Exchange"
IContractSpecifier_Exchange = mContractSpecifier.exchange
End Property

Private Property Let IContractSpecifier_Expiry(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Expiry"
End Property

Private Property Get IContractSpecifier_Expiry() As String
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Expiry"
IContractSpecifier_Expiry = mContractSpecifier.expiry
End Property

Private Function IContractSpecifier_FuzzyEquals(ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier) As Boolean
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to call FuzzyEquals"
IContractSpecifier_FuzzyEquals = mContractSpecifier.FuzzyEquals(pContractSpecifier)
End Function

Private Property Get IContractSpecifier_Key() As String
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Key"
IContractSpecifier_Key = mContractSpecifier.Key
End Property

Private Property Let IContractSpecifier_LocalSymbol(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify LocalSymbol"
End Property

Private Property Get IContractSpecifier_LocalSymbol() As String
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get LocalSymbol"
IContractSpecifier_LocalSymbol = mContractSpecifier.localSymbol
End Property

Private Property Let IContractSpecifier_Locked(ByVal RHS As Boolean)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Locked"
End Property

Private Property Get IContractSpecifier_Locked() As Boolean
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Locked"
IContractSpecifier_Locked = mContractSpecifier.Locked
End Property

Private Property Let IContractSpecifier_Right(ByVal RHS As TradeBuildSP.OptionRights)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Right"
End Property

Private Property Get IContractSpecifier_Right() As TradeBuildSP.OptionRights
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Right"
IContractSpecifier_Right = mContractSpecifier.Right
End Property

Private Property Let IContractSpecifier_SecType(ByVal RHS As TradeBuildSP.SecurityTypes)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify SecType"
End Property

Private Property Get IContractSpecifier_SecType() As TradeBuildSP.SecurityTypes
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get SecType"
IContractSpecifier_SecType = mContractSpecifier.sectype
End Property

Private Property Let IContractSpecifier_Strike(ByVal RHS As Double)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Strike"
End Property

Private Property Get IContractSpecifier_Strike() As Double
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get strike"
IContractSpecifier_Strike = mContractSpecifier.strike
End Property

Private Property Let IContractSpecifier_Symbol(ByVal RHS As String)
err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
            "Attempt to modify Symbol"
End Property

Private Property Get IContractSpecifier_Symbol() As String
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to get Symbol"
IContractSpecifier_Symbol = mContractSpecifier.symbol
End Property

Private Function IContractSpecifier_ToString() As String
If mContractSpecifier Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                            "Attempt to call ToString"
IContractSpecifier_ToString = mContractSpecifier.ToString
End Function

'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy(ByVal study As TradeBuildSP.IStudy, valueNames() As String, ByVal numUnderlyingValuesToUse As Long, Optional ByVal taskName As String, Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

End Function

Private Function IStudy_addStudyValueListener(ByVal listener As TradeBuildSP.IStudyValueListener, ByVal valueName As String, ByVal numberOfValuesToReplay As Long, Optional ByVal taskName As String, Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

End Function

Private Property Get IStudy_baseStudy() As TradeBuildSP.IStudy

End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)

End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters

End Property

Private Function IStudy_getStudyValue(ByVal valueName As String, ByVal ref As Long) As TradeBuildSP.StudyValue

End Function

Private Property Get IStudy_id() As String
IStudy_id = mMyStudyId
End Property

Private Sub IStudy_initialise(ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, ByVal id As String, ByVal parameters As TradeBuildSP.IParameters, ByVal numberOfValuesToCache As Long, valueNames() As String, ByVal underlyingStudy As TradeBuildSP.IStudy)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = "Test harness"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = "Test harness"
End Property

Private Sub IStudy_Notify(ev As TradeBuildSP.StudyValueEvent)

End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long

End Property

Private Function IStudy_numberOfCachedValues(Optional ByVal valueName As String = "") As Long

End Function

Private Property Get IStudy_parameters() As TradeBuildSP.IParameters

End Property

Private Sub IStudy_removeStudyValueListener(ByVal listener As TradeBuildSP.IStudyValueListener)

End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition

End Property

'================================================================================
' IStudyServiceConsumer Interface Members
'================================================================================

Private Property Get IStudyServiceConsumer_Contract() As TradeBuildSP.IContract
Set IStudyServiceConsumer_Contract = Me
End Property

Private Function IStudyServiceConsumer_replayStudyValues( _
                ByVal target As Object, _
                ByVal sourceStudy As TradeBuildSP.IStudy, _
                valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal discriminator As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

End Function

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub AddSPButton_Click()
SpList.AddItem SPToAddText
SpList.ListIndex = SpList.ListCount - 1
SPToAddText = ""
End Sub

Private Sub FindFileButton_Click()
Dim filePath As String
Dim fileExt As String

CommonDialog1.CancelError = True
On Error GoTo err

CommonDialog1.MaxFileSize = 32767
CommonDialog1.DialogTitle = "Open test data file"
CommonDialog1.Filter = "TradeBuild bar data files (*.tbd)|*.tbd"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = cdlOFNFileMustExist + _
                    cdlOFNLongNames + _
                    cdlOFNPathMustExist + _
                    cdlOFNExplorer + _
                    cdlOFNReadOnly
CommonDialog1.ShowOpen

TestDataFilenameText = CommonDialog1.FileName

LoadData

err:
End Sub

Private Sub RemoveSPButton_Click()
SpList.RemoveItem SpList.ListIndex
RemoveSPButton.Enabled = False
End Sub

Private Sub ServiceProviderClassNameText_Change()
If ServiceProviderClassNameText = "" Then
    SetSpButton.Enabled = False
Else
    SetSpButton.Enabled = True
End If
End Sub

Private Sub SetContractButton_Click()
setUpContract
End Sub

Private Sub SetSpButton_Click()
Dim availableStudies() As String
Dim i As Long

StudiesCombo.Clear
StudyConfigurer1.Clear
Set mStudyDefinition = Nothing
Set mStudiesServiceProvider = Nothing
mStudySet = False
TestButton.Enabled = False
TestDataGrid.Cols = TestDataGridColumns.StudyValue1

If ServiceProviderClassNameText = "" Then
    StudiesCombo.Enabled = False
    Exit Sub
End If

On Error Resume Next
Set mStudiesServiceProvider = CreateObject(ServiceProviderClassNameText)
On Error GoTo 0
If mStudiesServiceProvider Is Nothing Then
    StudiesCombo.Enabled = False
    MsgBox ServiceProviderClassNameText & " is not a valid Studies Service Provider"
    Exit Sub
End If

mTB.ServiceProviders.Add mStudiesServiceProvider

StudiesCombo.Enabled = True
availableStudies = mStudiesServiceProvider.implementedStudyNames
For i = 0 To UBound(availableStudies)
    StudiesCombo.AddItem availableStudies(i)
Next

End Sub

Private Sub SpList_Click()
If SpList.ListIndex = -1 Then
    RemoveSPButton.Enabled = False
Else
    RemoveSPButton.Enabled = True
End If
End Sub

Private Sub SPToAddText_Change()
If SPToAddText = "" Then
    AddSPButton.Enabled = False
Else
    AddSPButton.Enabled = True
End If
End Sub

Private Sub StudiesCombo_Click()
Dim regionNames(1) As String

Set mStudyParams = mTB.StudyDefaultParameters(StudiesCombo, "")

regionNames(0) = PriceRegionName
regionNames(1) = VolumeRegionName

Set mStudyDefinition = mTB.StudyDefinition(StudiesCombo, "")

StudyConfigurer1.initialise mStudyDefinition, _
                            "", _
                            regionNames, _
                            mBaseStudyConfigurations, _
                            Nothing, _
                            mStudyParams
mStudySet = True
If mDataLoaded Then TestButton.Enabled = True
End Sub

Private Sub TestButton_Click()
Dim i As Long
Dim j As Long
Dim when As String
Dim ev As TradeBuildSP.StudyValueEvent
Dim inputDefs As TradeBuildSP.IStudyInputDefinitions
Dim inputValueNames() As String
Dim valueName As String

On Error GoTo err
Screen.MousePointer = MousePointerConstants.vbArrowHourglass

mTB.ServiceProviders.RemoveAll

when = "adding additional service providers"
For i = 0 To SpList.ListCount - 1
    mTB.ServiceProviders.Add CreateObject(SpList.List(i))
Next

If BuiltInStudiesCheck.Value = vbChecked Then
    when = "adding Built-In Studies service provider"
    mTB.ServiceProviders.Add CreateObject("BuiltInStudiesSP.StudyServiceProvider")
End If

when = "adding your service provider"
mTB.ServiceProviders.Add mStudiesServiceProvider

mTB.ServiceProviders.Add Me

when = "creating the study to be tested"
Set mStudy = mStudiesServiceProvider.createStudy(StudiesCombo)

when = "setting up the Study Value grid"
Set mStudyConfiguration = StudyConfigurer1.StudyConfiguration
setupStudyValueGridColumns

when = "initialising the study to be tested"

inputValueNames = mStudyConfiguration.inputValueNames
mStudy.initialise mCommonServiceConsumer, _
                    Me, _
                    GenerateGUIDString, _
                    mStudyConfiguration.parameters, _
                    0, _
                    inputValueNames, _
                    Me

Set ev.Source = Me

Set inputDefs = mStudy.StudyDefinition.StudyInputDefinitions
For i = 1 To TestDataGrid.Rows
    TestDataGrid.row = i
    TestDataGrid.Col = TestDataGridColumns.Timestamp
    If TestDataGrid.Text = "" Then Exit For
    ev.Timestamp = CDate(TestDataGrid.Text)
    ev.barNumber = i

    For j = 0 To inputDefs.Count - 1
        valueName = inputValueNames(j)
        
        ev.valueName = inputDefs.Item(j + 1).name
        
        If valueName = BarsValueClose Then
            when = "notifying open value for bar " & i
            ev.Value = CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.openValue))
            mStudy.baseStudy.notify ev
            
            when = "notifying high value for bar " & i
            ev.Value = CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.highValue))
            mStudy.baseStudy.notify ev
            
            when = "notifying low value for bar " & i
            ev.Value = CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.lowValue))
            mStudy.baseStudy.notify ev
            
            when = "notifying close value for bar " & i
            ev.Value = CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.closeValue))
            mStudy.baseStudy.notify ev
        ElseIf valueName = BarsValueVolume Then
            If TestDataGrid.TextMatrix(i, TestDataGridColumns.volume) <> "" Then
                when = "notifying volume for bar " & i
                ev.Value = CLng(TestDataGrid.TextMatrix(i, TestDataGridColumns.volume))
                mStudy.baseStudy.notify ev
            End If
        End If
    Next
    
    processStudyValues i, when
Next

setTestDataGridRowBackColors 1

Screen.MousePointer = MousePointerConstants.vbDefault
Exit Sub

err:
setTestDataGridRowBackColors 1
MsgBox "Error " & err.Number & _
        " when " & when & _
        ": " & err.Description & _
        IIf(err.Source <> "", ": " & err.Source, "")
Screen.MousePointer = MousePointerConstants.vbDefault
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Public Function getBaseStudyDefinition() As StudyDefinition
Dim paramDef As StudyParameterDefinition
Dim valueDef As StudyValueDefinition

Set getBaseStudyDefinition = New StudyDefinition

getBaseStudyDefinition.Description = "Formats the price stream into Open/High/Low/Close bars of an appropriate length."
getBaseStudyDefinition.name = "Bars"
getBaseStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice

Set paramDef = getBaseStudyDefinition.StudyParameterDefinitions.Add(ParamNameBarLength)
paramDef.name = ParamNameBarLength
paramDef.Description = "Length of one bar"
paramDef.parameterType = StudyParameterTypes.ParameterTypeInteger
getBaseStudyDefinition.StudyParameterDefinitions.Add paramDef

Set paramDef = getBaseStudyDefinition.StudyParameterDefinitions.Add(ParamNameBarUnits)
paramDef.name = ParamNameBarUnits
paramDef.Description = "The units in which Period length is measured."
paramDef.parameterType = StudyParameterTypes.ParameterTypeString
getBaseStudyDefinition.StudyParameterDefinitions.Add paramDef

Set valueDef = getBaseStudyDefinition.StudyValueDefinitions.Add(BarsValueClose)
valueDef.name = BarsValueClose
valueDef.Description = "The latest underlying value"
valueDef.isDefault = True
valueDef.valueType = StudyValueTypes.ValueTypeDouble
getBaseStudyDefinition.StudyValueDefinitions.Add valueDef

Set valueDef = getBaseStudyDefinition.StudyValueDefinitions.Add(BarsValueVolume)
valueDef.name = BarsValueVolume
valueDef.Description = "The cumulative size associated with the latest underlying value (where relevant)"
valueDef.valueType = StudyValueTypes.ValueTypeInteger
getBaseStudyDefinition.StudyValueDefinitions.Add valueDef

End Function

Private Sub initialiseChart()

Chart1.suppressDrawing = True

Chart1.clearChart
Chart1.chartBackColor = vbWhite
Chart1.autoscale = True
Chart1.showCrosshairs = True
Chart1.twipsPerBar = 100
Chart1.showHorizontalScrollBar = True

Set mPriceRegion = Chart1.addChartRegion(100, 25, PriceRegionName)
mPriceRegion.gridlineSpacingY = 2
mPriceRegion.showGrid = True

Set mBarSeries = mPriceRegion.addBarSeries
mBarSeries.outlineThickness = 1
mBarSeries.tailThickness = 1
mBarSeries.barThickness = 1
mBarSeries.displayAsCandlestick = True
mBarSeries.solidUpBody = True

Set mVolumeRegion = Chart1.addChartRegion(20, , VolumeRegionName)
mVolumeRegion.gridlineSpacingY = 0.8
mVolumeRegion.minimumHeight = 10
mVolumeRegion.integerYScale = True
mVolumeRegion.showGrid = True
mVolumeRegion.setTitle "Volume", vbBlue, Nothing

Set mVolumeSeries = mVolumeRegion.addDataPointSeries
mVolumeSeries.displayMode = ChartSkil.DisplayModes.displayAsHistogram
mVolumeSeries.includeInAutoscale = True

Chart1.suppressDrawing = False

End Sub

Private Sub LoadData()
Dim fso As FileSystemObject
Dim ts As TextStream
Dim rec As String
Dim tokens() As String
Dim lPeriod As ChartSkil.Period
Dim lBar As ChartSkil.Bar
Dim lVolume As ChartSkil.DataPoint
Dim row As Long

On Error GoTo err

mDataLoaded = False
TestButton.Enabled = False

Screen.MousePointer = MousePointerConstants.vbArrowHourglass

TestDataGrid.Clear
setupTestDataGrid
If Not mStudyConfiguration Is Nothing Then setupStudyValueGridColumns
TestDataGrid.Refresh
initialiseChart
TestDataGrid.Redraw = False

Set mBaseStudyConfiguration = New TradeBuildUI.StudyConfiguration
mBaseStudyConfiguration.chartRegionName = PriceRegionName
mBaseStudyConfiguration.name = "Bars"
mBaseStudyConfiguration.StudyDefinition = mBaseStudyDefinition
mBaseStudyConfiguration.studyId = mMyStudyId
mBaseStudyConfiguration.instanceName = "Bars"
mBaseStudyConfiguration.instanceFullyQualifiedName = "Bars - " & TestDataFilenameText

Set mBaseStudyConfigurations = New TradeBuildUI.StudyConfigurations
mBaseStudyConfigurations.Add mBaseStudyConfiguration

Set fso = New FileSystemObject
Set ts = fso.OpenTextFile(TestDataFilenameText, ForReading)

Chart1.suppressDrawing = True
mPriceRegion.setTitle TestDataFilenameText, vbBlue, Nothing

Do While Not ts.AtEndOfStream
    rec = ts.ReadLine
    If rec <> "" And Left$(rec, 2) <> "//" Then
        tokens = Split(rec, ",")
        
        'update the chart
        Set lPeriod = Chart1.addperiod(CDate(tokens(TestDataFileColumns.Timestamp)))
        Set lBar = mBarSeries.addBar(lPeriod.periodNumber)
        lBar.Tick CDbl(tokens(TestDataFileColumns.openValue))
        lBar.Tick CDbl(tokens(TestDataFileColumns.highValue))
        lBar.Tick CDbl(tokens(TestDataFileColumns.lowValue))
        lBar.Tick CDbl(tokens(TestDataFileColumns.closeValue))
        If tokens(TestDataFileColumns.volume) <> "" Then
            Set lVolume = mVolumeSeries.addDataPoint(lPeriod.periodNumber)
            lVolume.dataValue = CLng(tokens(TestDataFileColumns.volume))
        End If
        Chart1.scrollX 1
        
        'update the grid
        row = row + 1
        If row > TestDataGrid.Rows - 1 Then TestDataGrid.Rows = TestDataGrid.Rows + TestDataGridRowsIncrement
        TestDataGrid.row = row
        TestDataGrid.Col = TestDataGridColumns.Timestamp
        TestDataGrid.Text = CDate(tokens(TestDataFileColumns.Timestamp))
        TestDataGrid.Col = TestDataGridColumns.openValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.openValue))
        TestDataGrid.Col = TestDataGridColumns.highValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.highValue))
        TestDataGrid.Col = TestDataGridColumns.lowValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.lowValue))
        TestDataGrid.Col = TestDataGridColumns.closeValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.closeValue))
        If tokens(TestDataFileColumns.volume) <> "" Then
            TestDataGrid.Col = TestDataGridColumns.volume
            TestDataGrid.Text = CLng(tokens(TestDataFileColumns.volume))
        End If
    End If
    
Loop

TestDataGrid.Redraw = True
setTestDataGridRowBackColors 1
Chart1.suppressDrawing = False

Screen.MousePointer = MousePointerConstants.vbDefault

mDataLoaded = True
If mStudySet Then TestButton.Enabled = True

Exit Sub

err:
TestDataGrid.Redraw = True
setTestDataGridRowBackColors 1
Chart1.suppressDrawing = False

Screen.MousePointer = MousePointerConstants.vbDefault

MsgBox "Can't load data file: " & TestDataFilenameText & vbCrLf & _
        "Error " & err.Number & ": " & err.Description
End Sub

Private Sub processStudyValues( _
                ByVal row As Long, _
                ByRef when As String)
Dim svd As StudyValueDefinition
Dim svc As StudyValueConfiguration
Dim lStudyValue As StudyValue
Dim i As Long
Dim j As Long

For i = 1 To mStudyConfiguration.StudyValueConfigurations.Count
    Set svc = mStudyConfiguration.StudyValueConfigurations.Item(i)
    If svc.includeInChart Then
        Set svd = mStudyDefinition.StudyValueDefinitions.Item(svc.valueName)
        when = "getting value for " & svc.valueName & " for bar " & row
        lStudyValue = mStudy.getStudyValue(svc.valueName, 0)
        
        TestDataGrid.TextMatrix(row, TestDataGridColumns.StudyValue1 + j) = lStudyValue.Value
        
        j = j + 1
    End If
Next

End Sub

Private Sub setTestDataGridRowBackColors( _
                ByVal startingIndex As Long)
Dim i As Long
Dim lTicker As Ticker

TestDataGrid.Redraw = False

For i = startingIndex To TestDataGrid.Rows - 1
    TestDataGrid.row = i
    TestDataGrid.Col = 0
    TestDataGrid.RowSel = i
    TestDataGrid.ColSel = TestDataGrid.Cols - 1
    TestDataGrid.CellBackColor = IIf(i Mod 2 = 0, CellBackColorEven, CellBackColorOdd)
    
Next

TestDataGrid.Redraw = True
End Sub

Private Sub setUpContract()
Set mContractSpecifier = mCommonServiceConsumer.newContractSpecifier
mContractSpecifier.currencyCode = CurrencyText
mContractSpecifier.exchange = ExchangeText
mContractSpecifier.expiry = ExpiryText
mContractSpecifier.localSymbol = LocalSymbolText
If RightCombo <> "" Then mContractSpecifier.Right = OptionRightFromString(RightCombo)
If TypeCombo <> "" Then mContractSpecifier.sectype = SecTypeFromString(TypeCombo)
mContractSpecifier.strike = IIf(StrikePriceText = "", 0, StrikePriceText)
mContractSpecifier.symbol = SymbolText

Set mContract = mCommonServiceConsumer.NewContract
If ContractIdText <> "" Then mContract.contractID = ContractIdText
mContract.Description = DescriptionText
If ExpiryDateText <> "" Then mContract.ExpiryDate = CDate(ExpiryDateText)
mContract.marketName = MarketNameText
mContract.MinimumTick = MinimumTickText
If MultiplierText <> "" Then mContract.multiplier = MultiplierText
If SessionEndTimeText <> "" Then mContract.SessionEndTime = CDate(SessionEndTimeText)
If SessionStartTimeText <> "" Then mContract.SessionStartTime = CDate(SessionStartTimeText)
mContract.tradingClass = TradingClassText

mContract.specifier = mContractSpecifier
End Sub

Private Sub setupStudyValueGridColumns()
Dim svd As StudyValueDefinition
Dim svc As StudyValueConfiguration
Dim i As Long
Dim j As Long

' remove any existing study value columns
TestDataGrid.Cols = TestDataGridColumns.StudyValue1

For i = 1 To mStudyConfiguration.StudyValueConfigurations.Count
    Set svc = mStudyConfiguration.StudyValueConfigurations.Item(i)
    If svc.includeInChart Then
        Set svd = mStudyDefinition.StudyValueDefinitions.Item(svc.valueName)
        setupTestDataGridColumn TestDataGridColumns.StudyValue1 + j, _
                                TestDataGridColumnWidths.StudyValue1Width, _
                                svd.name, _
                                IIf(svd.valueType = ValueTypeString, True, False), _
                                IIf(svd.valueType = ValueTypeString, AlignmentSettings.flexAlignLeftCenter, AlignmentSettings.flexAlignRightCenter)
        j = j + 1
    End If
Next

End Sub

Private Sub setupTestDataGrid()

With TestDataGrid
    .AllowBigSelection = True
    .AllowUserResizing = flexResizeBoth
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusNone
    .HighLight = flexHighlightNever
    
    .Cols = TestDataGridColumns.StudyValue1
    .Rows = TestDataGridRowsInitial
    .FixedRows = 1
    .FixedCols = 0
End With
    
setupTestDataGridColumn TestDataGridColumns.Timestamp, TestDataGridColumnWidths.TimeStampWidth, "Timestamp", False, AlignmentSettings.flexAlignLeftCenter
setupTestDataGridColumn TestDataGridColumns.openValue, TestDataGridColumnWidths.openValueWidth, "Open", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.highValue, TestDataGridColumnWidths.highValueWidth, "High", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.lowValue, TestDataGridColumnWidths.lowValueWidth, "Low", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.closeValue, TestDataGridColumnWidths.closeValueWidth, "Close", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.volume, TestDataGridColumnWidths.volumeWidth, "Volume", False, AlignmentSettings.flexAlignRightCenter

setTestDataGridRowBackColors 1
End Sub

Private Sub setupTestDataGridColumn( _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As AlignmentSettings)
    
Dim lColumnWidth As Long

With TestDataGrid
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .ColWidth(columnNumber) = 0
    End If
    
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
    .ColWidth(columnNumber) = lColumnWidth
    
    .ColAlignment(columnNumber) = align
    .TextMatrix(0, columnNumber) = columnHeader
End With
End Sub
                

