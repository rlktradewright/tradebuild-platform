VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl StudyConfigurer 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   ScaleHeight     =   5595
   ScaleWidth      =   11805
   Begin VB.Frame LinesFrame 
      Caption         =   "Horizontal lines"
      Height          =   735
      Left            =   5040
      TabIndex        =   33
      Top             =   4200
      Width           =   6735
      Begin VB.PictureBox LinesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   6540
         TabIndex        =   34
         Top             =   240
         Width           =   6540
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   4
            Left            =   5280
            TabIndex        =   39
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   3
            Left            =   3960
            TabIndex        =   38
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   37
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   36
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   4
            Left            =   6000
            TabIndex        =   44
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   3
            Left            =   4680
            TabIndex        =   43
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   42
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   41
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   40
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame ValuesFrame 
      Caption         =   "Output values"
      Height          =   4095
      Left            =   5040
      TabIndex        =   15
      Top             =   0
      Width           =   6735
      Begin VB.PictureBox ValuesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   6495
         TabIndex        =   16
         Top             =   240
         Width           =   6495
         Begin VB.CommandButton AdvancedButton 
            Caption         =   "..."
            Height          =   375
            Index           =   0
            Left            =   5640
            TabIndex        =   20
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox AutoscaleCheck 
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   19
            ToolTipText     =   "Set this to ensure that all values are visible when the chart is auto-scaling"
            Top             =   240
            Width           =   210
         End
         Begin VB.TextBox ThicknessText 
            Alignment       =   2  'Center
            Height          =   330
            Index           =   0
            Left            =   3600
            TabIndex        =   18
            Text            =   "1"
            ToolTipText     =   "Choose the thickness of lines or points"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox IncludeCheck 
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   17
            ToolTipText     =   "Set to include this study value in the chart"
            Top             =   240
            Width           =   195
         End
         Begin MSComctlLib.ImageCombo StyleCombo 
            Height          =   330
            Index           =   0
            Left            =   4440
            TabIndex        =   21
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
         Begin MSComctlLib.ImageCombo DisplayAsCombo 
            Height          =   330
            Index           =   0
            Left            =   2520
            TabIndex        =   22
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
         Begin MSComCtl2.UpDown ThicknessUpDown 
            Height          =   330
            Index           =   0
            Left            =   4080
            TabIndex        =   23
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "ThicknessText(0)"
            BuddyDispid     =   196623
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
         Begin VB.Label Label10 
            Caption         =   "Advanced"
            Height          =   255
            Left            =   5640
            TabIndex        =   32
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Style"
            Height          =   255
            Left            =   4440
            TabIndex        =   31
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Scale"
            Height          =   255
            Left            =   1560
            TabIndex        =   30
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Thickness"
            Height          =   255
            Left            =   3600
            TabIndex        =   29
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Display as"
            Height          =   255
            Left            =   2520
            TabIndex        =   28
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Color"
            Height          =   255
            Left            =   2160
            TabIndex        =   27
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Value name"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   0
            Width           =   975
         End
         Begin VB.Label ValueNameLabel 
            Caption         =   "Label2"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   25
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
            TabIndex        =   24
            ToolTipText     =   "Click to change the colour for this value"
            Top             =   240
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4935
      Left            =   2520
      TabIndex        =   9
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   2175
         TabIndex        =   10
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox ParameterValueCheck 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   46
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComctlLib.ImageCombo ParameterValueCombo 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   45
            Top             =   480
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.TextBox ParameterValueTemplateText 
            Height          =   330
            Left            =   1320
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox ParameterValueText 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
         Begin MSComCtl2.UpDown ParameterValueUpDown 
            Height          =   330
            Index           =   0
            Left            =   1920
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            BuddyControl    =   "ParameterValueText(0)"
            BuddyDispid     =   196631
            BuddyIndex      =   0
            OrigLeft        =   1920
            OrigRight       =   2175
            OrigBottom      =   285
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label ParameterNameLabel 
            Caption         =   "Param name"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin VB.TextBox StudyDescriptionText 
      BackColor       =   &H8000000F&
      Height          =   525
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5040
      Width           =   11775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Inputs"
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   2175
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         Begin MSComctlLib.ImageCombo InputValueCombo 
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   1440
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
            Top             =   840
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
            TabIndex        =   4
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Chart region"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Base study"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label InputValueNameLabel 
            Caption         =   "Input value"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   1200
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "StudyConfigurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private Const RegionCustom As String = "Use own region"

Private Const CC_RGBINIT         As Long = &H1
Private Const CC_FULLOPEN        As Long = &H2
Private Const CC_PREVENTFULLOPEN As Long = &H4
Private Const CC_SOLIDCOLOR      As Long = &H80
Private Const CC_ANYCOLOR        As Long = &H100

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type W32CHOOSECOLOR
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'================================================================================
' External procedure declarations
'================================================================================

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" ( _
                pChoosecolor As W32CHOOSECOLOR) As Long

'================================================================================
' member variables
'================================================================================

Private mStudyname As String
Private mStudyLibraryName As String

Private mStudyDefinition As StudyDefinition

Private mConfiguredStudies As StudyConfigurations

Private mNextTabIndex As Long

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
mNextTabIndex = 2
initialiseDisplayAsCombo DisplayAsCombo(0)
initialiseStyleCombo StyleCombo(0)
End Sub

'================================================================================
' XXXX Interface members
'================================================================================

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub AdvancedButton_Click(Index As Integer)
notImplemented
End Sub

Private Sub BaseStudiesCombo_Click()
Dim i As Long
For i = 0 To mStudyDefinition.studyInputDefinitions.Count - 1
    initialiseInputValueCombo i
Next
End Sub

Private Sub ColorLabel_Click( _
                Index As Integer)
'Dim simpleColorPicker As New fSimpleColorPicker
'Dim cursorpos As W32Point
'
'GetCursorPos cursorpos
'
'simpleColorPicker.Top = cursorpos.y * Screen.TwipsPerPixelY
'simpleColorPicker.Left = cursorpos.x * Screen.TwipsPerPixelX
'simpleColorPicker.initialColor = ColorLabel(Index).BackColor
'simpleColorPicker.Show vbModal, UserControl
'ColorLabel(Index).BackColor = simpleColorPicker.selectedColor
'Unload simpleColorPicker

Dim cc As W32CHOOSECOLOR


cc.flags = CC_FULLOPEN Or CC_RGBINIT Or CC_ANYCOLOR
cc.lStructSize = Len(cc)
cc.hwndOwner = Extender.Parent.hWnd
cc.lpCustColors = VarPtr(gCustColors(0))
cc.rgbResult = ColorLabel(Index).BackColor
ChooseColor cc
ColorLabel(Index).BackColor = cc.rgbResult
End Sub

Private Sub DisplayAsCombo_Validate( _
                Index As Integer, _
                Cancel As Boolean)
If DisplayAsCombo(Index).SelectedItem Is Nothing Then Cancel = True
End Sub

Private Sub LineColorLabel_Click(Index As Integer)
Dim simpleColorPicker As New fSimpleColorPicker
Dim formFrameThickness As Long
Dim formTitleBarThickness As Long

formFrameThickness = (UserControl.Width - UserControl.ScaleWidth) / 2
formTitleBarThickness = UserControl.Height - UserControl.ScaleHeight - formFrameThickness

simpleColorPicker.Top = Parent.Top + _
                        formTitleBarThickness + _
                        LinesFrame.Top + _
                        LinesPicture.Top + _
                        LineColorLabel(Index).Top + LineColorLabel(Index).Height / 2
simpleColorPicker.Left = Parent.Left + _
                        formFrameThickness + _
                        LinesFrame.Left + _
                        LinesPicture.Left + _
                        LineColorLabel(Index).Left + - _
                        (simpleColorPicker.Width - LineColorLabel(Index).Width) / 2
simpleColorPicker.initialColor = LineColorLabel(Index).BackColor
simpleColorPicker.Show vbModal, UserControl
LineColorLabel(Index).BackColor = simpleColorPicker.selectedColor
Unload simpleColorPicker
End Sub

Private Sub StyleCombo_Validate( _
                Index As Integer, _
                Cancel As Boolean)
If StyleCombo(Index).SelectedItem Is Nothing Then Cancel = True
End Sub

Private Sub ThicknessText_KeyPress(Index As Integer, KeyAscii As Integer)
filterNonNumericKeyPress KeyAscii
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get studyConfiguration() As studyConfiguration
Dim studyConfig As studyConfiguration
Dim params As Parameters2.Parameters
Dim studyParamDef As StudyParameterDefinition
Dim studyValueDefs As studyValueDefinitions
Dim studyValueDef As StudyValueDefinition
Dim studyValueConfig As StudyValueConfiguration
Dim studyHorizRule As StudyHorizontalRule
Dim regionName As String
Dim inputValueNames() As String
Dim i As Long
Dim scfg As studyConfiguration

Set studyConfig = New studyConfiguration
'studyConfig.studyDefinition = mStudyDefinition
studyConfig.name = mStudyname
studyConfig.StudyLibraryName = mStudyLibraryName
If Not BaseStudiesCombo.SelectedItem Is Nothing Then
    For Each scfg In mConfiguredStudies
        If scfg.Study.id = BaseStudiesCombo.SelectedItem.Tag Then
            studyConfig.underlyingStudy = scfg.Study
            Exit For
        End If
    Next
End If

ReDim inputValueNames(mStudyDefinition.studyInputDefinitions.Count - 1) As String
For i = 0 To UBound(inputValueNames)
    If Not InputValueCombo(i).SelectedItem Is Nothing Then
        inputValueNames(i) = InputValueCombo(i).SelectedItem.text
    End If
Next
studyConfig.inputValueNames = inputValueNames

If ChartRegionCombo.SelectedItem.text = RegionDefault Then
    Select Case mStudyDefinition.defaultRegion
    Case DefaultRegionNone
        regionName = RegionNameDefault
    Case DefaultRegionCustom
        regionName = RegionNameCustom
    End Select
ElseIf ChartRegionCombo.SelectedItem.text = RegionCustom Then
    regionName = RegionNameCustom
Else
    regionName = ChartRegionCombo.SelectedItem.text
End If
studyConfig.chartRegionName = regionName

Set params = New Parameters2.Parameters

For i = 0 To mStudyDefinition.studyParameterDefinitions.Count - 1
    Set studyParamDef = mStudyDefinition.studyParameterDefinitions.item(i + 1)
    If studyParamDef.parameterType = ParameterTypeBoolean Then
        params.setParameterValue ParameterNameLabel(i).Caption, _
                                IIf(ParameterValueCheck(i) = vbChecked, "True", "False")
    ElseIf ParameterValueText(i).Visible Then
        params.setParameterValue ParameterNameLabel(i).Caption, ParameterValueText(i).text
    Else
        params.setParameterValue ParameterNameLabel(i).Caption, ParameterValueCombo(i).text
    End If
Next

studyConfig.Parameters = params

Set studyValueDefs = mStudyDefinition.studyValueDefinitions

For i = 0 To ValueNameLabel.UBound
    Set studyValueConfig = studyConfig.StudyValueConfigurations.Add(ValueNameLabel(i).Caption)
    studyValueConfig.includeInChart = (IncludeCheck(i).value = vbChecked)
    studyValueConfig.includeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
    studyValueConfig.Color = ColorLabel(i).BackColor
    
    Set studyValueDef = studyValueDefs.item(i + 1)
    
    Select Case studyValueDef.defaultRegion
    Case DefaultRegionNone
        studyValueConfig.chartRegionName = RegionNameDefault
    Case DefaultRegionCustom
        studyValueConfig.chartRegionName = RegionNameCustom
    End Select
    
    Select Case DisplayAsCombo(i).SelectedItem.text
    Case DisplayModeLine
        studyValueConfig.dataPointDisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
    Case DisplayModePoint
        studyValueConfig.dataPointDisplayMode = DataPointDisplayModes.DataPointDisplayModePoint
    Case DisplayModeSteppedLine
        studyValueConfig.dataPointDisplayMode = DataPointDisplayModes.DataPointDisplayModeStep
    Case DisplayModeHistogram
        studyValueConfig.dataPointDisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram
    End Select
    
    studyValueConfig.lineThickness = ThicknessText(i).text
    
    Select Case StyleCombo(i).SelectedItem.text
    Case LineStyleSolid
        studyValueConfig.LineStyle = LineSolid
    Case LineStyleDash
        studyValueConfig.LineStyle = LineDash
    Case LineStyleDot
        studyValueConfig.LineStyle = LineDot
    Case LineStyleDashDot
        studyValueConfig.LineStyle = LineDashDot
    Case LineStyleDashDotDot
        studyValueConfig.LineStyle = LineDashDotDot
    End Select
    
    ' temporary fix until ChartSkil improves darwing op non-extended lines
    If studyValueDef.valueMode = ValueModeLine Then studyValueConfig.extended = True
Next

For i = 0 To 4
    If LineText(i).text <> "" Then
        Set studyHorizRule = studyConfig.StudyHorizontalRules.Add
        studyHorizRule.y = LineText(i).text
        studyHorizRule.Color = LineColorLabel(i).BackColor
    End If
Next

Set studyConfiguration = studyConfig
End Property

'================================================================================
' methods
'================================================================================

Public Sub clear()
initialiseControls
End Sub

Public Sub initialise( _
                ByVal studyDef As StudyDefinition, _
                ByVal StudyLibraryName As String, _
                ByRef regionNames() As String, _
                ByVal configuredStudies As StudyConfigurations, _
                ByVal defaultConfiguration As studyConfiguration, _
                ByVal defaultParameters As Parameters2.Parameters)
                
If Not defaultConfiguration Is Nothing And defaultParameters Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuildUI.StudyConfigurer::initialise", _
            "DefaultConfiguration and DefaultParameters cannot both be Nothing"
End If

initialiseControls

Set mStudyDefinition = studyDef
mStudyLibraryName = StudyLibraryName
Set mConfiguredStudies = configuredStudies

processRegionNames regionNames

setupBaseStudiesCombo

processStudyDefinition defaultConfiguration, defaultParameters
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub initialiseControls()
Dim i As Long

On Error Resume Next

For i = InputValueNameLabel.UBound To 1 Step -1
    Unload InputValueNameLabel(i)
Next
InputValueNameLabel(0).Caption = ""
InputValueNameLabel(0).Visible = False

For i = InputValueCombo.UBound To 1 Step -1
    Unload InputValueCombo(i)
Next

For i = ParameterNameLabel.UBound To 1 Step -1
    Unload ParameterNameLabel(i)
Next
ParameterNameLabel(0).Caption = ""
ParameterNameLabel(0).Visible = False

For i = ParameterValueText.UBound To 1 Step -1
    Unload ParameterValueText(i)
Next
ParameterValueText(0).text = ""
ParameterValueText(0).Visible = False

For i = ParameterValueCombo.UBound To 1 Step -1
    Unload ParameterValueCombo(i)
Next
ParameterValueCombo(0).text = ""
ParameterValueCombo(0).ComboItems.clear
ParameterValueCombo(0).Visible = False

For i = ParameterValueCheck.UBound To 1 Step -1
    Unload ParameterValueCheck(i)
Next
ParameterValueCombo(0).Visible = False

For i = ParameterValueUpDown.UBound To 1 Step -1
    Unload ParameterValueUpDown(i)
Next
ParameterValueUpDown(0).Visible = False

For i = IncludeCheck.UBound To 1 Step -1
    Unload IncludeCheck(i)
Next
IncludeCheck(0).value = vbUnchecked

For i = ValueNameLabel.UBound To 1 Step -1
    Unload ValueNameLabel(i)
Next
ValueNameLabel(0).Caption = ""

For i = AutoscaleCheck.UBound To 1 Step -1
    Unload AutoscaleCheck(i)
Next
AutoscaleCheck(0).value = vbUnchecked

For i = ColorLabel.UBound To 1 Step -1
    Unload ColorLabel(i)
Next
ColorLabel(0).BackColor = vbRed

For i = DisplayAsCombo.UBound To 1 Step -1
    Unload DisplayAsCombo(i)
Next
DisplayAsCombo(0).ComboItems(0).Selected = True

For i = ThicknessText.UBound To 1 Step -1
    Unload ThicknessText(i)
Next
ThicknessText(0).text = "1"

For i = ThicknessUpDown.UBound To 1 Step -1
    Unload ThicknessUpDown(i)
Next

For i = StyleCombo.UBound To 1 Step -1
    Unload StyleCombo(i)
Next
StyleCombo(0).ComboItems(0).Selected = True

For i = AdvancedButton.UBound To 1 Step -1
    Unload AdvancedButton(i)
Next

End Sub

Private Sub initialiseDisplayAsCombo(ByVal combo As ImageCombo)
Dim item As ComboItem
Set item = combo.ComboItems.Add(, , DisplayModeLine)
item.Selected = True
combo.ComboItems.Add , , DisplayModePoint
combo.ComboItems.Add , , DisplayModeSteppedLine
combo.ComboItems.Add , , DisplayModeHistogram
End Sub

Private Sub initialiseInputValueCombo( _
                ByVal Index As Long)
Dim studyValueDefs As studyValueDefinitions
Dim valueDef As StudyValueDefinition
Dim inputDef As StudyInputDefinition
Dim item As ComboItem
Dim i As Long
Dim selIndex As Long

If mConfiguredStudies Is Nothing Then Exit Sub

Set item = BaseStudiesCombo.SelectedItem
Set studyValueDefs = mConfiguredStudies.item(item.Key).Study.StudyDefinition.studyValueDefinitions
Set inputDef = mStudyDefinition.studyInputDefinitions.item(Index + 1)

InputValueCombo(Index).ComboItems.clear

selIndex = -1
InputValueCombo(Index).ComboItems.Add , , ""
For Each valueDef In studyValueDefs
    If typesCompatible(valueDef.valueType, inputDef.inputType) Then
        InputValueCombo(Index).ComboItems.Add , , valueDef.name
        If UCase$(inputDef.name) = UCase$(valueDef.name) Then selIndex = InputValueCombo(Index).ComboItems.Count
        If valueDef.isDefault And _
            selIndex = -1 Then selIndex = InputValueCombo(Index).ComboItems.Count
    End If
Next

If InputValueCombo(Index).ComboItems.Count <> 0 And selIndex <> -1 Then
    InputValueCombo(Index).ComboItems(IIf(selIndex <> 0, selIndex, 1)).Selected = True
End If

InputValueCombo(Index).Refresh
End Sub

Private Sub initialiseStyleCombo(ByVal combo As ImageCombo)
Dim item As ComboItem
Set item = combo.ComboItems.Add(, , LineStyleSolid)
item.Selected = True
combo.ComboItems.Add , , LineStyleDash
combo.ComboItems.Add , , LineStyleDot
combo.ComboItems.Add , , LineStyleDashDot
combo.ComboItems.Add , , LineStyleDashDotDot
End Sub

Private Function nextTabIndex() As Long
nextTabIndex = mNextTabIndex
mNextTabIndex = mNextTabIndex + 1
End Function

Private Sub processRegionNames( _
                ByRef regionNames() As String)
Dim i As Long

ChartRegionCombo.ComboItems.clear

ChartRegionCombo.ComboItems.Add , , RegionDefault
ChartRegionCombo.ComboItems.Add , , RegionCustom

For i = 0 To UBound(regionNames)
    ChartRegionCombo.ComboItems.Add , , regionNames(i)
Next
ChartRegionCombo.ComboItems.item(1).Selected = True
ChartRegionCombo.Refresh
End Sub

Private Sub processStudyDefinition( _
                ByVal defaultConfig As studyConfiguration, _
                ByVal defaultParams As Parameters2.Parameters)
Dim i As Long
Dim studyInputDefinitions As studyInputDefinitions
Dim studyParameterDefinitions As studyParameterDefinitions
Dim studyValueDefinitions As studyValueDefinitions
Dim studyinput As StudyInputDefinition
Dim studyParam As StudyParameterDefinition
Dim studyValue As StudyValueDefinition
Dim studyValueConfigs As StudyValueConfigurations
Dim studyValueConfig As StudyValueConfiguration
Dim studyHorizRules As StudyHorizontalRules
Dim studyHorizRule As StudyHorizontalRule
Dim firstParamIsInteger As Boolean
Dim permittedParamValues() As Variant
Dim permittedParamValue As Variant
Dim numPermittedParamValues As Long
Dim defaultParamValue As String

mNextTabIndex = 2

mStudyname = mStudyDefinition.name

If Not defaultConfig Is Nothing Then
    Set defaultParams = defaultConfig.Parameters
    Set studyValueConfigs = defaultConfig.StudyValueConfigurations
    Set studyHorizRules = defaultConfig.StudyHorizontalRules
End If

StudyDescriptionText.text = mStudyDefinition.Description

If Not defaultConfig Is Nothing Then
    If defaultConfig.chartRegionName = defaultConfig.instanceFullyQualifiedName Then
        '
        setComboSelection ChartRegionCombo, RegionCustom
    Else
        setComboSelection ChartRegionCombo, defaultConfig.chartRegionName
    End If
    
    For i = 1 To BaseStudiesCombo.ComboItems.Count
        If BaseStudiesCombo.ComboItems(i).Tag = defaultConfig.underlyingStudy.id Then
            BaseStudiesCombo.ComboItems(i).Selected = True
            Exit For
        End If
    Next
    
End If

Set studyInputDefinitions = mStudyDefinition.studyInputDefinitions
For i = 1 To studyInputDefinitions.Count
    Set studyinput = studyInputDefinitions.item(i)
    If i = 1 Then
        InputValueNameLabel(0).Visible = True
    Else
        Load InputValueNameLabel(i - 1)
        InputValueNameLabel(i - 1).Top = InputValueNameLabel(i - 2).Top + 600
        InputValueNameLabel(i - 1).Visible = True
        Load InputValueCombo(i - 1)
        InputValueCombo(i - 1).Top = InputValueCombo(i - 2).Top + 600
        InputValueCombo(i - 1).Visible = True
        InputValueCombo(i - 1).TabIndex = nextTabIndex
    End If
    InputValueNameLabel(i - 1).Caption = studyinput.name
    InputValueCombo(i - 1).ToolTipText = studyinput.Description

    initialiseInputValueCombo i - 1
    If Not defaultConfig Is Nothing Then
        Dim inputValueNames() As String
        inputValueNames = defaultConfig.inputValueNames
        setComboSelection InputValueCombo(i - 1), inputValueNames(i - 1)
    End If
    
Next

Set studyParameterDefinitions = mStudyDefinition.studyParameterDefinitions

For i = 1 To studyParameterDefinitions.Count
    Set studyParam = studyParameterDefinitions.item(i)
    If i = 1 Then
        ParameterNameLabel(0).Visible = True
        
        ParameterValueText(0).Visible = False
        ParameterValueText(0).TabIndex = nextTabIndex
        
        ParameterValueCombo(0).Top = ParameterValueText(0).Top
        ParameterValueCombo(0).Visible = False
        ParameterValueCombo(0).TabIndex = nextTabIndex
        
        ParameterValueCheck(0).Top = ParameterValueText(0).Top
        ParameterValueCheck(0).Visible = False
        ParameterValueCheck(0).TabIndex = nextTabIndex
    Else
        Load ParameterNameLabel(i - 1)
        ParameterNameLabel(i - 1).Top = ParameterNameLabel(i - 2).Top + 360
        ParameterNameLabel(i - 1).Left = ParameterNameLabel(i - 2).Left
        ParameterNameLabel(i - 1).Visible = True

        Load ParameterValueText(i - 1)
        ParameterValueText(i - 1).TabIndex = nextTabIndex
        ParameterValueText(i - 1).Width = ParameterValueTemplateText.Width
        ParameterValueText(i - 1).Top = ParameterValueText(i - 2).Top + 360
        ParameterValueText(i - 1).Left = ParameterValueText(i - 2).Left
        ParameterValueText(i - 1).Visible = False
    
        Load ParameterValueUpDown(i - 1)
        ParameterValueUpDown(i - 1).TabIndex = nextTabIndex
        ParameterValueUpDown(i - 1).Top = ParameterValueUpDown(i - 2).Top + 360
        
        Load ParameterValueCombo(i - 1)
        ParameterValueCombo(i - 1).TabIndex = nextTabIndex
        ParameterValueCombo(i - 1).Top = ParameterValueCombo(i - 2).Top + 360
    
        Load ParameterValueCheck(i - 1)
        ParameterValueCheck(i - 1).TabIndex = nextTabIndex
        ParameterValueCheck(i - 1).Top = ParameterValueCombo(i - 2).Top + 360
    End If
    
    permittedParamValues = studyParam.permittedValues
    
    numPermittedParamValues = -1
    On Error Resume Next
    numPermittedParamValues = UBound(permittedParamValues)
    On Error GoTo 0
    If numPermittedParamValues <> -1 Then
        For Each permittedParamValue In permittedParamValues
            ParameterValueCombo(i - 1).ComboItems.Add , , permittedParamValue
        Next
        ParameterValueCombo(i - 1).Visible = True
    ElseIf studyParam.parameterType = StudyParameterTypes.ParameterTypeInteger Then
        ParameterValueUpDown(i - 1).Min = 1
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
            ParameterValueText(0).Width = ParameterValueTemplateText.Width - ParameterValueUpDown(0).Width
            firstParamIsInteger = True
        End If
        ParameterValueText(i - 1).Visible = True
    ElseIf studyParam.parameterType = StudyParameterTypes.ParameterTypeBoolean Then
        ParameterValueCheck(i - 1).Visible = True
    Else
        If i = 1 Then
            ParameterValueUpDown(0).Visible = False
            ParameterValueText(0).Width = ParameterValueTemplateText.Width
        End If
        ParameterValueText(i - 1).Visible = True
    End If
    
    ParameterNameLabel(i - 1).Caption = studyParam.name
    defaultParamValue = defaultParams.getParameterValue(studyParam.name)
    If studyParam.parameterType = StudyParameterTypes.ParameterTypeBoolean Then
        Select Case UCase$(defaultParamValue)
        Case "Y", "YES", "T", "TRUE", "1"
            ParameterValueCheck(i - 1) = vbChecked
        Case "N", "NO", "F", "FALSE", "0"
            ParameterValueCheck(i - 1) = vbUnchecked
        End Select
    ElseIf numPermittedParamValues = -1 Then
        ParameterValueText(i - 1).text = defaultParamValue
        ParameterValueText(i - 1).ToolTipText = studyParam.Description
    Else
        ParameterValueCombo(i - 1).text = defaultParamValue
        ParameterValueCombo(i - 1).ToolTipText = studyParam.Description
    End If
    
    If studyParam.parameterType = StudyParameterTypes.ParameterTypeInteger Or _
        studyParam.parameterType = StudyParameterTypes.ParameterTypeReal _
    Then
        ParameterValueText(i - 1).Alignment = AlignmentConstants.vbRightJustify
    Else
        ParameterValueText(i - 1).Alignment = AlignmentConstants.vbLeftJustify
    End If
Next

IncludeCheck(0).TabIndex = nextTabIndex
AutoscaleCheck(0).TabIndex = nextTabIndex
ColorLabel(0).TabIndex = nextTabIndex
DisplayAsCombo(0).TabIndex = nextTabIndex
ThicknessText(0).TabIndex = nextTabIndex
ThicknessUpDown(0).TabIndex = nextTabIndex

Set studyValueDefinitions = mStudyDefinition.studyValueDefinitions
For i = 1 To studyValueDefinitions.Count
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
    
    AutoscaleCheck(i - 1) = vbChecked
    
    ValueNameLabel(i - 1).Caption = studyValue.name
    ValueNameLabel(i - 1).ToolTipText = studyValue.Description

    If Not studyValueConfig Is Nothing Then
        IncludeCheck(i - 1) = IIf(studyValueConfig.includeInChart, vbChecked, vbUnchecked)
        AutoscaleCheck(i - 1) = IIf(studyValueConfig.includeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = studyValueConfig.Color
        
        Select Case studyValueConfig.dataPointDisplayMode
        Case DataPointDisplayModes.DataPointDisplayModeLine
            setComboSelection DisplayAsCombo(i - 1), DisplayModeLine
        Case DataPointDisplayModes.DataPointDisplayModePoint
            setComboSelection DisplayAsCombo(i - 1), DisplayModePoint
        Case DataPointDisplayModes.DataPointDisplayModeStep
            setComboSelection DisplayAsCombo(i - 1), DisplayModeSteppedLine
        Case DataPointDisplayModes.DataPointDisplayModeHistogram
            setComboSelection DisplayAsCombo(i - 1), DisplayModeHistogram
        End Select
        
        ThicknessText(i - 1).text = studyValueConfig.lineThickness
        
        Select Case studyValueConfig.LineStyle
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

If Not studyHorizRules Is Nothing Then
    For i = 1 To studyHorizRules.Count
        Set studyHorizRule = studyHorizRules.item(i)
        LineText(i - 1) = studyHorizRule.y
        LineColorLabel(i - 1).BackColor = studyHorizRule.Color
    Next
End If
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

Private Sub setupBaseStudiesCombo()
Dim studyConfig As studyConfiguration
Dim item As ComboItem

BaseStudiesCombo.ComboItems.clear
If mConfiguredStudies Is Nothing Then Exit Sub
For Each studyConfig In mConfiguredStudies
    If Not TypeOf studyConfig.Study Is InputStudy Or _
        Not mStudyDefinition.needsBars _
    Then
        If studiesCompatible(studyConfig.Study.StudyDefinition, mStudyDefinition) Then
            Set item = BaseStudiesCombo.ComboItems.Add(, studyConfig.instanceFullyQualifiedName, studyConfig.Study.instanceName)
            item.Tag = studyConfig.Study.id
        End If
    End If
Next
BaseStudiesCombo.ComboItems(1).Selected = True
BaseStudiesCombo.Refresh

End Sub

Private Function studiesCompatible( _
                ByVal sourceStudyDefinition As StudyDefinition, _
                ByVal sinkStudyDefinition As StudyDefinition) As Boolean
Dim sourceValueDef As StudyValueDefinition
Dim sinkInputDef As StudyInputDefinition
Dim i As Long

For i = 1 To sinkStudyDefinition.studyInputDefinitions.Count
    Set sinkInputDef = sinkStudyDefinition.studyInputDefinitions.item(i)
    For Each sourceValueDef In sourceStudyDefinition.studyValueDefinitions
        If typesCompatible(sourceValueDef.valueType, sinkInputDef.inputType) Then
            studiesCompatible = True
            Exit For
        End If
    Next
    If Not studiesCompatible Then Exit Function
Next
End Function

Private Function typesCompatible( _
                ByVal sourceValueType As StudyValueTypes, _
                ByVal sinkInputType As StudyInputTypes) As Boolean
Select Case sourceValueType
Case ValueTypeInteger
    Select Case sinkInputType
    Case InputTypeInteger
        typesCompatible = True
    Case InputTypeReal
        typesCompatible = True
    End Select
Case ValueTypeReal
    Select Case sinkInputType
    Case InputTypeReal
        typesCompatible = True
    End Select
Case ValueTypeString
    Select Case sinkInputType
    Case InputTypeString
        typesCompatible = True
    End Select
Case ValueTypeDate
    Select Case sinkInputType
    Case InputTypeDate
        typesCompatible = True
    End Select
End Select
End Function



