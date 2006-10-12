VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fStudyConfigurer 
   Caption         =   "Configure a study"
   ClientHeight    =   5745
   ClientLeft      =   1005
   ClientTop       =   1230
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   10725
   Begin VB.Frame LinesFrame 
      Caption         =   "Lines"
      Height          =   735
      Left            =   2640
      TabIndex        =   39
      Top             =   4320
      Width           =   6735
      Begin VB.PictureBox LinesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   6540
         TabIndex        =   40
         Top             =   240
         Width           =   6540
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   4
            Left            =   5280
            TabIndex        =   17
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   3
            Left            =   3960
            TabIndex        =   16
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   15
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   14
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   4
            Left            =   6000
            TabIndex        =   45
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   3
            Left            =   4680
            TabIndex        =   44
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   43
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   42
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   41
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton SetDefaultButton 
      Caption         =   "Set as &default"
      Height          =   615
      Left            =   9480
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame ValuesFrame 
      Caption         =   "Output values"
      Height          =   4095
      Left            =   2640
      TabIndex        =   25
      Top             =   120
      Width           =   6735
      Begin VB.PictureBox ValuesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   6495
         TabIndex        =   26
         Top             =   240
         Width           =   6495
         Begin VB.CommandButton AdvancedButton 
            Caption         =   "..."
            Height          =   375
            Index           =   0
            Left            =   5640
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin MSComctlLib.ImageCombo StyleCombo 
            Height          =   330
            Index           =   0
            Left            =   4440
            TabIndex        =   11
            ToolTipText     =   "Choose the line style (ignored if thickness is grater than 1)"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.CheckBox AutoscaleCheck 
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   6
            ToolTipText     =   "Set this to ensure that all values are visible when the chart is auto-scaling"
            Top             =   240
            Width           =   195
         End
         Begin MSComCtl2.UpDown ThicknessUpDown 
            Height          =   330
            Index           =   0
            Left            =   4080
            TabIndex        =   10
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "ThicknessText(0)"
            BuddyDispid     =   196618
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   240
            OrigRight       =   4335
            OrigBottom      =   570
            Max             =   5
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox ThicknessText 
            Alignment       =   2  'Center
            Height          =   330
            Index           =   0
            Left            =   3600
            TabIndex        =   9
            Text            =   "1"
            ToolTipText     =   "Choose the thickness of lines or points"
            Top             =   240
            Width           =   600
         End
         Begin MSComctlLib.ImageCombo DisplayAsCombo 
            Height          =   330
            Index           =   0
            Left            =   2520
            TabIndex        =   8
            ToolTipText     =   "Select how to display this value"
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.CheckBox IncludeCheck 
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Set to include this study value in the chart"
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label10 
            Caption         =   "Advanced"
            Height          =   255
            Left            =   5640
            TabIndex        =   38
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Style"
            Height          =   255
            Left            =   4440
            TabIndex        =   33
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Scale"
            Height          =   255
            Left            =   1560
            TabIndex        =   32
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Thickness"
            Height          =   255
            Left            =   3600
            TabIndex        =   31
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Display as"
            Height          =   255
            Left            =   2520
            TabIndex        =   30
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Color"
            Height          =   255
            Left            =   2160
            TabIndex        =   29
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Value name"
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label ValueNameLabel 
            Caption         =   "Label2"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label ColorLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   2160
            TabIndex        =   7
            ToolTipText     =   "Click to change the colour for this value"
            Top             =   240
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4935
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   2415
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   2175
         TabIndex        =   34
         Top             =   2760
         Width           =   2175
         Begin MSComctlLib.ImageCombo InputValueCombo 
            Height          =   330
            Left            =   0
            TabIndex        =   4
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin MSComctlLib.ImageCombo BaseStudiesCombo 
            Height          =   330
            Left            =   0
            TabIndex        =   3
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin MSComctlLib.ImageCombo ChartRegionCombo 
            Height          =   330
            Left            =   0
            TabIndex        =   2
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.Label Label9 
            Caption         =   "Input value"
            Height          =   255
            Left            =   0
            TabIndex        =   37
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Base study"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Chart region"
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   2175
         TabIndex        =   23
         Top             =   240
         Width           =   2175
         Begin VB.TextBox ParameterValueTemplateText 
            Height          =   285
            Left            =   1320
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSComCtl2.UpDown ParameterValueUpDown 
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   1
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "ParameterValueText(0)"
            BuddyDispid     =   196636
            BuddyIndex      =   0
            OrigLeft        =   1920
            OrigRight       =   2175
            OrigBottom      =   285
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox ParameterValueText 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   0
            Top             =   0
            Width           =   600
         End
         Begin VB.Label ParameterNameLabel 
            Caption         =   "Param name"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   9480
      TabIndex        =   19
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "&Add to chart"
      Height          =   615
      Left            =   9480
      TabIndex        =   18
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox StudyDescriptionText 
      BackColor       =   &H8000000F&
      Height          =   525
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5160
      Width           =   10455
   End
End
Attribute VB_Name = "fStudyConfigurer"
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

'================================================================================
' Events
'================================================================================

Event SetDefault(ByVal studyConfig As StudyConfiguration)
Event AddStudyConfiguration(ByVal studyConfig As StudyConfiguration)

'================================================================================
' Constants
'================================================================================

Private Const DisplayModeLine As String = "Line"
Private Const DisplayModePoint As String = "Point"
Private Const DisplayModeSteppedLine As String = "Stepped line"
Private Const DisplayModeHistogram As String = "Histogram"

Private Const LineStyleSolid As String = "Solid"
Private Const LineStyleDash As String = "Dash"
Private Const LineStyleDot As String = "Dot"
Private Const LineStyleDashDot As String = "Dash dot"
Private Const LineStyleDashDotDot As String = "Dash dot dot"
Private Const LineStyleInsideSolid As String = "Inside solid"
Private Const LineStyleInvisible As String = "Invisible"

Private Const RegionDefault As String = "Use default"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTicker As TradeBuild.ticker

Private mStudyname As String
Private mServiceProviderName As String

Private mStudyDefinition As TradeBuild.studyDefinition

Private mConfiguredStudies As studyConfigurations

Private mNextTabIndex As Long

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
mNextTabIndex = 2
End Sub

Private Sub Form_Load()
initialiseDisplayAsCombo DisplayAsCombo(0)
initialiseStyleCombo StyleCombo(0)
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub AddButton_Click()
RaiseEvent AddStudyConfiguration(createStudyConfig)
Unload Me
End Sub

Private Sub AdvancedButton_Click(index As Integer)
notImplemented
End Sub

Private Sub BaseStudiesCombo_Click()
initialiseInputValueCombo ""
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub ColorLabel_Click( _
                index As Integer)
Dim simpleColorPicker As New fSimpleColorPicker
Dim formFrameThickness As Long
Dim formTitleBarThickness As Long

formFrameThickness = (Me.Width - Me.ScaleWidth) / 2
formTitleBarThickness = Me.Height - Me.ScaleHeight - formFrameThickness

simpleColorPicker.Top = Me.Top + _
                        formTitleBarThickness + _
                        ValuesFrame.Top + _
                        ValuesPicture.Top + _
                        ColorLabel(index).Top + ColorLabel(index).Height / 2
simpleColorPicker.Left = Me.Left + _
                        formFrameThickness + _
                        ValuesFrame.Left + _
                        ValuesPicture.Left + _
                        ColorLabel(index).Left + - _
                        (simpleColorPicker.Width - ColorLabel(index).Width) / 2
simpleColorPicker.initialColor = ColorLabel(index).backColor
simpleColorPicker.Show vbModal, Me
ColorLabel(index).backColor = simpleColorPicker.selectedColor
Unload simpleColorPicker
End Sub

Private Sub DisplayAsCombo_Validate( _
                index As Integer, _
                Cancel As Boolean)
If DisplayAsCombo(index).selectedItem Is Nothing Then Cancel = True
End Sub

Private Sub LineColorLabel_Click(index As Integer)
Dim simpleColorPicker As New fSimpleColorPicker
Dim formFrameThickness As Long
Dim formTitleBarThickness As Long

formFrameThickness = (Me.Width - Me.ScaleWidth) / 2
formTitleBarThickness = Me.Height - Me.ScaleHeight - formFrameThickness

simpleColorPicker.Top = Me.Top + _
                        formTitleBarThickness + _
                        LinesFrame.Top + _
                        LinesPicture.Top + _
                        LineColorLabel(index).Top + LineColorLabel(index).Height / 2
simpleColorPicker.Left = Me.Left + _
                        formFrameThickness + _
                        LinesFrame.Left + _
                        LinesPicture.Left + _
                        LineColorLabel(index).Left + - _
                        (simpleColorPicker.Width - LineColorLabel(index).Width) / 2
simpleColorPicker.initialColor = LineColorLabel(index).backColor
simpleColorPicker.Show vbModal, Me
LineColorLabel(index).backColor = simpleColorPicker.selectedColor
Unload simpleColorPicker
End Sub

Private Sub SetDefaultButton_Click()
notImplemented
Exit Sub

RaiseEvent SetDefault(createStudyConfig)
End Sub

Private Sub StyleCombo_Validate( _
                index As Integer, _
                Cancel As Boolean)
If StyleCombo(index).selectedItem Is Nothing Then Cancel = True
End Sub

Private Sub ThicknessText_KeyPress(index As Integer, KeyAscii As Integer)
filterNonNumericKeyPress KeyAscii
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

Friend Sub initialise( _
                ByVal ticker As TradeBuild.ticker, _
                ByVal studyDef As TradeBuild.studyDefinition, _
                ByVal serviceProviderName As String, _
                ByRef regionNames() As String, _
                ByVal configuredStudies As studyConfigurations, _
                ByVal defaultConfiguration As StudyConfiguration)
                
Set mTicker = ticker
                
processStudyDefinition studyDef, defaultConfiguration

mServiceProviderName = serviceProviderName

If defaultConfiguration Is Nothing Then
    processRegionNames regionNames, ""
Else
    processRegionNames regionNames, defaultConfiguration.chartRegionName
End If

If defaultConfiguration Is Nothing Then
    processConfiguredStudies configuredStudies, ""
Else
    processConfiguredStudies configuredStudies, defaultConfiguration.underlyingStudyId
End If

If defaultConfiguration Is Nothing Then
    initialiseInputValueCombo ""
Else
    initialiseInputValueCombo defaultConfiguration.inputValueName
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function createStudyConfig() As StudyConfiguration
Dim studyConfig As StudyConfiguration
Dim params As TradeBuild.parameters
Dim studyValueDefs As TradeBuild.studyValueDefinitions
Dim studyValueDef As TradeBuild.StudyValueDefinition
Dim studyValueConfig As StudyValueConfiguration
Dim studyHorizRule As StudyHorizontalRule
Dim regionName As String
Dim i As Long

Set studyConfig = New StudyConfiguration
studyConfig.studyDefinition = mStudyDefinition
studyConfig.name = mStudyname
studyConfig.serviceProviderName = mServiceProviderName
studyConfig.underlyingStudyId = BaseStudiesCombo.selectedItem.Tag
studyConfig.inputValueName = InputValueCombo.selectedItem.text

If ChartRegionCombo.selectedItem.text = RegionDefault Then
    Select Case mStudyDefinition.defaultRegion
    Case DefaultRegionNone
        regionName = PriceRegionName
    Case DefaultRegionPrice
        regionName = PriceRegionName
    Case DefaultRegionVolume
        regionName = VolumeRegionName
    Case DefaultRegionCustom
        regionName = CustomRegionName
    End Select
Else
    regionName = ChartRegionCombo.selectedItem.text
End If
studyConfig.chartRegionName = regionName

Set params = New TradeBuild.parameters

For i = 0 To ParameterNameLabel.UBound
    params.setParameterValue ParameterNameLabel(i).caption, ParameterValueText(i).text
Next

studyConfig.parameters = params

Set studyValueDefs = mStudyDefinition.studyValueDefinitions

For i = 0 To ValueNameLabel.UBound
    Set studyValueConfig = studyConfig.studyValueConfigurations.add(ValueNameLabel(i).caption)
    studyValueConfig.includeInChart = (IncludeCheck(i).value = vbChecked)
    studyValueConfig.includeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
    studyValueConfig.color = ColorLabel(i).backColor
    
    Set studyValueDef = studyValueDefs.item(i + 1)
    
    studyValueConfig.multipleValuesPerBar = studyValueDef.multipleValuesPerBar
    
    Select Case studyValueDef.defaultRegion
    Case DefaultRegionNone
        studyValueConfig.chartRegionName = regionName
    Case DefaultRegionPrice
        studyValueConfig.chartRegionName = PriceRegionName
    Case DefaultRegionVolume
        studyValueConfig.chartRegionName = VolumeRegionName
    Case DefaultRegionCustom
        studyValueConfig.chartRegionName = CustomRegionName
    End Select
    
    Select Case DisplayAsCombo(i).selectedItem.text
    Case DisplayModeLine
        studyValueConfig.displayMode = DisplayAsLines
    Case DisplayModePoint
        studyValueConfig.displayMode = displayAsPoints
    Case DisplayModeSteppedLine
        studyValueConfig.displayMode = DisplayAsSteppedLines
    Case DisplayModeHistogram
        studyValueConfig.displayMode = displayAsHistogram
    End Select
    
    studyValueConfig.lineThickness = ThicknessText(i).text
    
    Select Case StyleCombo(i).selectedItem.text
    Case LineStyleSolid
        studyValueConfig.lineStyle = LineSolid
    Case LineStyleDash
        studyValueConfig.lineStyle = LineDash
    Case LineStyleDot
        studyValueConfig.lineStyle = LineDot
    Case LineStyleDashDot
        studyValueConfig.lineStyle = LineDashDot
    Case LineStyleDashDotDot
        studyValueConfig.lineStyle = LineDashDotDot
    End Select
Next

For i = 0 To 4
    If LineText(i).text <> "" Then
        Set studyHorizRule = studyConfig.studyHorizontalRules.add
        studyHorizRule.y = LineText(i).text
        studyHorizRule.color = LineColorLabel(i).backColor
    End If
Next

Set createStudyConfig = studyConfig
End Function
Private Sub initialiseDisplayAsCombo(ByVal combo As ImageCombo)
Dim item As ComboItem
Set item = combo.ComboItems.add(, , DisplayModeLine)
item.Selected = True
combo.ComboItems.add , , DisplayModePoint
combo.ComboItems.add , , DisplayModeSteppedLine
combo.ComboItems.add , , DisplayModeHistogram
End Sub

Private Sub initialiseInputValueCombo( _
                ByRef selectedValue As String)
Dim studyValueDefs As TradeBuild.studyValueDefinitions
Dim valueDef As TradeBuild.StudyValueDefinition
Dim item As ComboItem
Dim i As Long

Set item = BaseStudiesCombo.selectedItem
Set studyValueDefs = mConfiguredStudies.item(item.key).studyDefinition.studyValueDefinitions

InputValueCombo.ComboItems.clear

' There seems to be a bug in VB6: this doesn't work when the study definition has
' been created by a service provider
'For Each valueDef In studyValueDefs
'    InputValueCombo.ComboItems.add , , valueDef.name
'Next

For i = 1 To studyValueDefs.count
    Set valueDef = studyValueDefs.item(i)
    InputValueCombo.ComboItems.add , , valueDef.name
Next

If selectedValue = "" Then
    InputValueCombo.ComboItems(1).Selected = True
Else
    setComboSelection InputValueCombo, selectedValue
End If

InputValueCombo.Refresh
End Sub

Private Sub initialiseStyleCombo(ByVal combo As ImageCombo)
Dim item As ComboItem
Set item = combo.ComboItems.add(, , LineStyleSolid)
item.Selected = True
combo.ComboItems.add , , LineStyleDash
combo.ComboItems.add , , LineStyleDot
combo.ComboItems.add , , LineStyleDashDot
combo.ComboItems.add , , LineStyleDashDotDot
End Sub

Private Function nextTabIndex() As Long
nextTabIndex = mNextTabIndex
mNextTabIndex = mNextTabIndex + 1
End Function

Private Sub processConfiguredStudies( _
                ByVal studyConfigs As studyConfigurations, _
                ByRef selectedValue As String)
Dim studyConfig As StudyConfiguration
Dim item As ComboItem

Set mConfiguredStudies = studyConfigs

For Each studyConfig In mConfiguredStudies
    Set item = BaseStudiesCombo.ComboItems.add(, studyConfig.instanceFullyQualifiedName, studyConfig.instanceName)
    item.Tag = studyConfig.studyId
Next
If selectedValue = "" Then
    BaseStudiesCombo.ComboItems(1).Selected = True
Else
    setComboSelection BaseStudiesCombo, selectedValue
End If
BaseStudiesCombo.Refresh

End Sub

Private Sub processRegionNames( _
                ByRef regionNames() As String, _
                ByRef selectedValue As String)
Dim i As Long

ChartRegionCombo.ComboItems.add , , RegionDefault

For i = 0 To UBound(regionNames)
    ChartRegionCombo.ComboItems.add , , regionNames(i)
Next
If selectedValue = "" Then
    ChartRegionCombo.ComboItems.item(1).Selected = True
Else
    setComboSelection ChartRegionCombo, selectedValue
End If
End Sub

Private Sub processStudyDefinition( _
                ByVal value As TradeBuild.studyDefinition, _
                ByVal defaultConfig As StudyConfiguration)
Dim i As Long
Dim studyParameterDefinitions As TradeBuild.studyParameterDefinitions
Dim studyValueDefinitions As TradeBuild.studyValueDefinitions
Dim studyParam As TradeBuild.StudyParameterDefinition
Dim defaultParams As TradeBuild.parameters
Dim studyValue As TradeBuild.StudyValueDefinition
Dim studyValueConfigs As studyValueConfigurations
Dim studyValueConfig As StudyValueConfiguration
Dim firstParamIsInteger As Boolean

Set mStudyDefinition = value

mStudyname = mStudyDefinition.name

If Not defaultConfig Is Nothing Then
    Set defaultParams = defaultConfig.parameters
    Set studyValueConfigs = defaultConfig.studyValueConfigurations
Else
    Set defaultParams = mTicker.StudyDefaultParameters( _
                                            mStudyDefinition.name, _
                                            mServiceProviderName)
End If

Me.caption = mStudyDefinition.name
StudyDescriptionText.text = mStudyDefinition.Description

Set studyParameterDefinitions = mStudyDefinition.studyParameterDefinitions
Set studyValueDefinitions = mStudyDefinition.studyValueDefinitions

For i = 1 To studyParameterDefinitions.count
    Set studyParam = studyParameterDefinitions.item(i)
    If i <> 1 Then
        Load ParameterNameLabel(i - 1)
        ParameterNameLabel(i - 1).Top = ParameterNameLabel(i - 2).Top + 360
        ParameterNameLabel(i - 1).Left = ParameterNameLabel(i - 2).Left
        ParameterNameLabel(i - 1).Visible = True

        Load ParameterValueText(i - 1)
        ParameterValueText(i - 1).TabIndex = nextTabIndex
        ParameterValueText(i - 1).Width = ParameterValueTemplateText.Width
        ParameterValueText(i - 1).Top = ParameterValueText(i - 2).Top + 360
        ParameterValueText(i - 1).Left = ParameterValueText(i - 2).Left
        ParameterValueText(i - 1).Visible = True
    
        Load ParameterValueUpDown(i - 1)
        ParameterValueUpDown(i - 1).TabIndex = nextTabIndex
        ParameterValueUpDown(i - 1).Top = ParameterValueUpDown(i - 2).Top + 360
        'ParameterValueUpDown(i - 1).Left = ParameterValueUpDown(i - 2).Left
    End If
    
    ParameterNameLabel(i - 1).caption = studyParam.name
    ParameterValueText(i - 1).text = defaultParams.getParameterValue(studyParam.name)
    ParameterValueText(i - 1).ToolTipText = studyParam.Description
    
    If studyParam.parameterType = StudyParameterTypes.ParameterTypeInteger Then
        ParameterValueUpDown(i - 1).Min = 0
        ParameterValueUpDown(i - 1).Max = 255
        ParameterValueUpDown(i - 1).Visible = True
        If i <> 1 Then
            ParameterValueUpDown(i - 1).BuddyControl = ParameterValueText(i - 1)
            ' the following line is necessary because for some reason VB resizes
            ' the first parametervaluetext whenever BuddyControl is set to true
            ' on a later UpDown control  !!!
            If firstParamIsInteger Then
                ParameterValueText(0).Width = ParameterValueTemplateText.Width - ParameterValueUpDown(0).Width
            Else
                ParameterValueText(0).Width = ParameterValueTemplateText.Width
            End If
        Else
            firstParamIsInteger = True
        End If
    Else
        If i = 1 Then
            ParameterValueUpDown(0).Visible = False
            ParameterValueText(0).Width = ParameterValueTemplateText.Width
        End If
    End If
    
    If studyParam.parameterType = StudyParameterTypes.ParameterTypeInteger Or _
        studyParam.parameterType = StudyParameterTypes.ParameterTypeSingle Or _
        studyParam.parameterType = StudyParameterTypes.ParameterTypeDouble _
    Then
        ParameterValueText(i - 1).Alignment = AlignmentConstants.vbRightJustify
    Else
        ParameterValueText(i - 1).Alignment = AlignmentConstants.vbLeftJustify
    End If
Next

ChartRegionCombo.TabIndex = nextTabIndex
BaseStudiesCombo.TabIndex = nextTabIndex
InputValueCombo.TabIndex = nextTabIndex

IncludeCheck(0).TabIndex = nextTabIndex
AutoscaleCheck(0).TabIndex = nextTabIndex
ColorLabel(0).TabIndex = nextTabIndex
DisplayAsCombo(0).TabIndex = nextTabIndex
ThicknessText(0).TabIndex = nextTabIndex
ThicknessUpDown(0).TabIndex = nextTabIndex

For i = 1 To studyValueDefinitions.count
    Set studyValue = studyValueDefinitions.item(i)
    If Not studyValueConfigs Is Nothing Then
        Set studyValueConfig = studyValueConfigs.item(i)
    End If
    
    If i <> 1 Then
        Load IncludeCheck(i - 1)
        IncludeCheck(i - 1).Top = IncludeCheck(i - 2).Top + 360
        IncludeCheck(i - 1).Left = IncludeCheck(i - 2).Left
        IncludeCheck(i - 1).Visible = True
        IncludeCheck(i - 1).TabIndex = nextTabIndex
    
        Load ValueNameLabel(i - 1)
        ValueNameLabel(i - 1).Top = ValueNameLabel(i - 2).Top + 360
        ValueNameLabel(i - 1).Left = ValueNameLabel(i - 2).Left
        ValueNameLabel(i - 1).Visible = True
    
        Load AutoscaleCheck(i - 1)
        AutoscaleCheck(i - 1).Top = AutoscaleCheck(i - 2).Top + 360
        AutoscaleCheck(i - 1).Left = AutoscaleCheck(i - 2).Left
        AutoscaleCheck(i - 1).Visible = True
        AutoscaleCheck(i - 1).TabIndex = nextTabIndex
    
        Load ColorLabel(i - 1)
        ColorLabel(i - 1).Top = ColorLabel(i - 2).Top + 360
        ColorLabel(i - 1).Left = ColorLabel(i - 2).Left
        ColorLabel(i - 1).Visible = True
        ColorLabel(i - 1).TabIndex = nextTabIndex
    
        Load DisplayAsCombo(i - 1)
        DisplayAsCombo(i - 1).Top = DisplayAsCombo(i - 2).Top + 360
        DisplayAsCombo(i - 1).Left = DisplayAsCombo(i - 2).Left
        DisplayAsCombo(i - 1).Visible = True
        DisplayAsCombo(i - 1).TabIndex = nextTabIndex
        initialiseDisplayAsCombo DisplayAsCombo(i - 1)
    
        Load ThicknessText(i - 1)
        ThicknessText(i - 1).TabIndex = nextTabIndex
        ThicknessText(i - 1).Width = ThicknessUpDown(i - 2).Left + _
                                    ThicknessUpDown(i - 2).Width - _
                                    ThicknessText(i - 2).Left
        ThicknessText(i - 1).Top = ThicknessText(i - 2).Top + 360
        ThicknessText(i - 1).Left = ThicknessText(i - 2).Left
        ThicknessText(i - 1).Visible = True
    
        Load ThicknessUpDown(i - 1)
        ThicknessUpDown(i - 1).TabIndex = nextTabIndex
        ThicknessUpDown(i - 1).Top = ThicknessUpDown(i - 2).Top + 360
        ThicknessUpDown(i - 1).BuddyControl = ThicknessText(i - 1)
        ThicknessUpDown(i - 1).Visible = True
        ' need the following line otherwise VB increases the length
        ' of the first thicknessText !!!
        If i <> 1 Then ThicknessText(0).Width = ThicknessText(i - 1).Width

    
        Load StyleCombo(i - 1)
        StyleCombo(i - 1).Top = StyleCombo(i - 2).Top + 360
        StyleCombo(i - 1).Left = StyleCombo(i - 2).Left
        StyleCombo(i - 1).Visible = True
        StyleCombo(i - 1).TabIndex = nextTabIndex
        initialiseStyleCombo StyleCombo(i - 1)
    
        Load AdvancedButton(i - 1)
        AdvancedButton(i - 1).Top = AdvancedButton(i - 2).Top + 360
        AdvancedButton(i - 1).Left = AdvancedButton(i - 2).Left
        AdvancedButton(i - 1).Visible = True
        AdvancedButton(i - 1).TabIndex = nextTabIndex
        
    End If
    
    ValueNameLabel(i - 1).caption = studyValue.name
    ValueNameLabel(i - 1).ToolTipText = studyValue.Description

    If Not studyValueConfig Is Nothing Then
        IncludeCheck(i - 1) = vbChecked
        AutoscaleCheck(i - 1) = IIf(studyValueConfig.includeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).backColor = studyValueConfig.color
        
        Select Case studyValueConfig.displayMode
        Case DisplayAsLines
            setComboSelection DisplayAsCombo(i - 1), DisplayModeLine
        Case displayAsPoints
            setComboSelection DisplayAsCombo(i - 1), DisplayModePoint
        Case DisplayAsSteppedLines
            setComboSelection DisplayAsCombo(i - 1), DisplayModeSteppedLine
        Case displayAsHistogram
            setComboSelection DisplayAsCombo(i - 1), DisplayModeHistogram
        End Select
        
        ThicknessText(i - 1).text = studyValueConfig.lineThickness
        
        Select Case studyValueConfig.lineStyle
        Case LineSolid
            setComboSelection StyleCombo(i - 1), LineStyleSolid
        Case LineDash
            setComboSelection StyleCombo(i - 1), LineStyleDash
        Case LineDot
            setComboSelection StyleCombo(i - 1), LineStyleDot
        Case LineDashDot
            setComboSelection StyleCombo(i - 1), LineStyleDashDot
        Case LineDashDotDot
            setComboSelection StyleCombo(i - 1), LineStyleDashDotDot
        End Select
        
    End If

Next
End Sub

Private Sub setComboSelection( _
                ByVal combo As ImageCombo, _
                ByVal text As String)
Dim item As ComboItem
For Each item In combo.ComboItems
    If UCase$(item.text) = UCase$(text) Then
        item.Selected = True
        Exit For
    End If
Next
End Sub

