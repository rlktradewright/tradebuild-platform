VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl StudyConfigurer 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12210
   ScaleHeight     =   5595
   ScaleWidth      =   12210
   Begin VB.Frame LinesFrame 
      Caption         =   "Horizontal lines"
      Height          =   735
      Left            =   5040
      TabIndex        =   37
      Top             =   4200
      Width           =   7095
      Begin VB.PictureBox LinesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   6900
         TabIndex        =   38
         Top             =   240
         Width           =   6900
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   4
            Left            =   5280
            TabIndex        =   43
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   3
            Left            =   3960
            TabIndex        =   42
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   41
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   40
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   4
            Left            =   6000
            TabIndex        =   48
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   3
            Left            =   4680
            TabIndex        =   47
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   46
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   45
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   44
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame ValuesFrame 
      Caption         =   "Output values"
      Height          =   4095
      Left            =   5040
      TabIndex        =   17
      Top             =   0
      Width           =   7095
      Begin VB.PictureBox ValuesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   6855
         TabIndex        =   18
         Top             =   240
         Width           =   6855
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3360
            Top             =   1920
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton FontButton 
            Caption         =   "Font..."
            Height          =   375
            Index           =   0
            Left            =   5400
            TabIndex        =   50
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton AdvancedButton 
            Caption         =   "..."
            Height          =   375
            Index           =   0
            Left            =   6360
            TabIndex        =   28
            ToolTipText     =   "Click for advanced features"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox AutoscaleCheck 
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   24
            ToolTipText     =   "Set this to ensure that all values are visible when the chart is auto-scaling"
            Top             =   240
            Width           =   210
         End
         Begin VB.TextBox ThicknessText 
            Alignment       =   2  'Center
            Height          =   330
            Index           =   0
            Left            =   4320
            TabIndex        =   25
            Text            =   "1"
            ToolTipText     =   "Choose the thickness of lines or points"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox IncludeCheck 
            Height          =   195
            Index           =   0
            Left            =   1560
            TabIndex        =   19
            ToolTipText     =   "Set to include this study value in the chart"
            Top             =   240
            Width           =   195
         End
         Begin MSComctlLib.ImageCombo StyleCombo 
            Height          =   330
            Index           =   0
            Left            =   5160
            TabIndex        =   27
            ToolTipText     =   "Choose the line style (ignored if thickness is greater than 1)"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin MSComctlLib.ImageCombo DisplayModeCombo 
            Height          =   330
            Index           =   0
            Left            =   3240
            TabIndex        =   23
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
            Left            =   4800
            TabIndex        =   26
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "ThicknessText(0)"
            BuddyDispid     =   196617
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   240
            OrigRight       =   4335
            OrigBottom      =   570
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Scale"
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Value name"
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Show"
            Height          =   255
            Left            =   1320
            TabIndex        =   49
            Top             =   0
            Width           =   375
         End
         Begin VB.Label DownColorLabel 
            BackColor       =   &H000000FF&
            Height          =   330
            Index           =   0
            Left            =   2865
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
         Begin VB.Label UpColorLabel 
            BackColor       =   &H0000FF00&
            Height          =   330
            Index           =   0
            Left            =   2520
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
         Begin VB.Label ColorLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   2175
            TabIndex        =   20
            ToolTipText     =   "Click to change the colour for this value"
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "Advanced"
            Height          =   255
            Left            =   6120
            TabIndex        =   36
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Style"
            Height          =   255
            Left            =   5040
            TabIndex        =   35
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Thickness"
            Height          =   255
            Left            =   4320
            TabIndex        =   33
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Display as"
            Height          =   255
            Left            =   3120
            TabIndex        =   32
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Colors"
            Height          =   255
            Left            =   2400
            TabIndex        =   31
            Top             =   0
            Width           =   495
         End
         Begin VB.Label ValueNameLabel 
            Caption         =   "Label2"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4935
      Left            =   2520
      TabIndex        =   13
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
         TabIndex        =   14
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox ParameterValueCheck 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   6
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComctlLib.ImageCombo ParameterValueCombo 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   4
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
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox ParameterValueText 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
         Begin MSComCtl2.UpDown ParameterValueUpDown 
            Height          =   330
            Index           =   0
            Left            =   1920
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            BuddyControl    =   "ParameterValueText(0)"
            BuddyDispid     =   196634
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
            TabIndex        =   16
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5040
      Width           =   12135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Inputs"
      Height          =   4935
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   2175
         TabIndex        =   8
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
            TabIndex        =   1
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
            TabIndex        =   0
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
            TabIndex        =   11
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Base study"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label InputValueNameLabel 
            Caption         =   "Input value"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   9
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

Private Const RegionDefault As String = "Use default"
Private Const RegionCustom As String = "Use own region"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mController As chartController
Private mStudyname As String
Private mStudyLibraryName As String

Private mStudyDefinition As StudyDefinition

Private mConfiguredStudies As StudyConfigurations

Private mNextTabIndex As Long

Private mDefaultConfiguration As studyConfiguration

Private mFonts() As StdFont

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
mNextTabIndex = 2
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
Dim studyValueDef As StudyValueDefinition
Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

ColorLabel(Index).BackColor = chooseAColor(ColorLabel(Index).BackColor, _
                                            IIf(studyValueDef.valueMode = ValueModeBar, True, False))

End Sub

Private Sub DisplayModeCombo_Click(Index As Integer)
Dim studyValueDef  As StudyValueDefinition
Dim studyValueconfig  As StudyValueConfiguration

Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

If Not mDefaultConfiguration Is Nothing Then
    Set studyValueconfig = mDefaultConfiguration.StudyValueConfigurations.item(Index + 1)
End If

Select Case studyValueDef.valueMode
Case ValueModeNone
    Dim dpStyle As dataPointStyle
    
    If Not studyValueconfig Is Nothing Then
        Set dpStyle = studyValueconfig.dataPointStyle
    Else
        Set dpStyle = mController.defaultDataPointStyle
    End If
        
    Select Case DisplayModeCombo(Index).SelectedItem.text
    Case PointDisplayModeLine
        initialiseLineStyleCombo StyleCombo(Index), dpStyle.lineStyle
    Case PointDisplayModePoint
        initialisePointStyleCombo StyleCombo(Index), dpStyle.pointStyle
    Case PointDisplayModeSteppedLine
        initialiseLineStyleCombo StyleCombo(Index), dpStyle.lineStyle
    Case PointDisplayModeHistogram
        initialiseHistogramStyleCombo StyleCombo(Index), dpStyle.histBarWidth
    End Select
Case ValueModeLine

Case ValueModeBar

Case ValueModeText

End Select

End Sub

Private Sub DisplayModeCombo_Validate( _
                Index As Integer, _
                Cancel As Boolean)
If DisplayModeCombo(Index).SelectedItem Is Nothing Then Cancel = True
End Sub

Private Sub DownColorLabel_Click(Index As Integer)
Dim studyValueDef As StudyValueDefinition
Dim allowNullColor As Boolean

Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

If studyValueDef.valueMode = ValueModeBar Or _
    studyValueDef.valueMode = ValueModeNone Then allowNullColor = True

DownColorLabel(Index).BackColor = chooseAColor(DownColorLabel(Index).BackColor, _
                                            allowNullColor)

End Sub

Private Sub FontButton_Click(Index As Integer)
Dim aFont As StdFont

CommonDialog1.flags = cdlCFBoth + cdlCFEffects
CommonDialog1.FontName = mFonts(Index).name
CommonDialog1.FontBold = mFonts(Index).Bold
CommonDialog1.FontItalic = mFonts(Index).Italic
CommonDialog1.FontSize = mFonts(Index).Size
CommonDialog1.FontStrikethru = mFonts(Index).Strikethrough
CommonDialog1.FontUnderline = mFonts(Index).Underline
CommonDialog1.Color = ColorLabel(Index).BackColor
CommonDialog1.ShowFont

Set aFont = New StdFont
aFont.Bold = CommonDialog1.FontBold
aFont.Italic = CommonDialog1.FontItalic
aFont.name = CommonDialog1.FontName
aFont.Size = CommonDialog1.FontSize
aFont.Strikethrough = CommonDialog1.FontStrikethru
aFont.Underline = CommonDialog1.FontUnderline

Set mFonts(Index) = aFont

ColorLabel(Index).BackColor = CommonDialog1.Color

End Sub

Private Sub LineColorLabel_Click(Index As Integer)
LineColorLabel(Index).BackColor = chooseAColor(LineColorLabel(Index).BackColor, False)
End Sub

Private Sub StyleCombo_Validate( _
                Index As Integer, _
                Cancel As Boolean)
If StyleCombo(Index).SelectedItem Is Nothing Then Cancel = True
End Sub

Private Sub ThicknessText_KeyPress(Index As Integer, KeyAscii As Integer)
filterNonNumericKeyPress KeyAscii
End Sub

Private Sub UpColorLabel_Click(Index As Integer)
Dim studyValueDef As StudyValueDefinition
Dim allowNullColor As Boolean

Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

If studyValueDef.valueMode = ValueModeBar Or _
    studyValueDef.valueMode = ValueModeNone Then allowNullColor = True

UpColorLabel(Index).BackColor = chooseAColor(UpColorLabel(Index).BackColor, _
                                            allowNullColor)

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
Dim studyValueconfig As StudyValueConfiguration
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
        If scfg.study.id = BaseStudiesCombo.SelectedItem.Tag Then
            studyConfig.underlyingStudy = scfg.study
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

For i = 0 To ValueNameLabel.ubound
    Set studyValueDef = studyValueDefs.item(i + 1)
    
    Set studyValueconfig = studyConfig.StudyValueConfigurations.Add(ValueNameLabel(i).Caption)
    studyValueconfig.includeInChart = (IncludeCheck(i).value = vbChecked)
    
    Select Case studyValueDef.defaultRegion
    Case DefaultRegionNone
        studyValueconfig.chartRegionName = RegionNameDefault
    Case DefaultRegionCustom
        studyValueconfig.chartRegionName = RegionNameCustom
    End Select
    
    Select Case studyValueDef.valueMode
    Case ValueModeNone
        Dim dpStyle As dataPointStyle
        
        Set dpStyle = mController.defaultDataPointStyle
        
        dpStyle.includeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
        dpStyle.Color = ColorLabel(i).BackColor
        dpStyle.downColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                            -1, _
                            DownColorLabel(i).BackColor)
        dpStyle.upColor = IIf(UpColorLabel(i).BackColor = NullColor, _
                            -1, _
                            UpColorLabel(i).BackColor)
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case PointDisplayModeLine
            dpStyle.displayMode = DataPointDisplayModes.DataPointDisplayModeLine
            Select Case StyleCombo(i).SelectedItem.text
            Case LineStyleSolid
                dpStyle.lineStyle = LineSolid
            Case LineStyleDash
                dpStyle.lineStyle = LineDash
            Case LineStyleDot
                dpStyle.lineStyle = LineDot
            Case LineStyleDashDot
                dpStyle.lineStyle = LineDashDot
            Case LineStyleDashDotDot
                dpStyle.lineStyle = LineDashDotDot
            End Select
        Case PointDisplayModePoint
            dpStyle.displayMode = DataPointDisplayModes.DataPointDisplayModePoint
            Select Case StyleCombo(0).SelectedItem.text
            Case PointStyleRound
                dpStyle.pointStyle = PointRound
            Case PointStyleSquare
                dpStyle.pointStyle = PointSquare
            End Select
        Case PointDisplayModeSteppedLine
            dpStyle.displayMode = DataPointDisplayModes.DataPointDisplayModeStep
            Select Case StyleCombo(i).SelectedItem.text
            Case LineStyleSolid
                dpStyle.lineStyle = LineSolid
            Case LineStyleDash
                dpStyle.lineStyle = LineDash
            Case LineStyleDot
                dpStyle.lineStyle = LineDot
            Case LineStyleDashDot
                dpStyle.lineStyle = LineDashDot
            Case LineStyleDashDotDot
                dpStyle.lineStyle = LineDashDotDot
            End Select
        Case PointDisplayModeHistogram
            dpStyle.displayMode = DataPointDisplayModes.DataPointDisplayModeHistogram
            Select Case StyleCombo(0).SelectedItem.text
            Case HistogramStyleNarrow
                dpStyle.histBarWidth = HistogramWidthNarrow
            Case HistogramStyleMedium
                dpStyle.histBarWidth = HistogramWidthMedium
            Case HistogramStyleWide
                dpStyle.histBarWidth = HistogramWidthWide
            Case CustomStyle
                dpStyle.histBarWidth = CSng(StyleCombo(0).SelectedItem.Tag)
            End Select
        End Select
        
        dpStyle.lineThickness = ThicknessText(i).text
        
        studyValueconfig.dataPointStyle = dpStyle
    Case ValueModeLine
        Dim lnStyle As lineStyle

        Set lnStyle = mController.defaultLineStyle
        
        lnStyle.includeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
        lnStyle.Color = ColorLabel(i).BackColor
        lnStyle.arrowStartColor = IIf(UpColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        UpColorLabel(i).BackColor = NullColor)
        lnStyle.arrowEndColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        DownColorLabel(i).BackColor = NullColor)
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case LineDisplayModePlain
            lnStyle.arrowEndStyle = ArrowNone
            lnStyle.arrowStartStyle = ArrowNone
        Case LineDisplayModeArrowEnd
            lnStyle.arrowEndStyle = ArrowClosed
            lnStyle.arrowStartStyle = ArrowNone
        Case LineDisplayModeArrowStart
            lnStyle.arrowEndStyle = ArrowNone
            lnStyle.arrowStartStyle = ArrowClosed
        Case LineDisplayModeArrowBoth
            lnStyle.arrowEndStyle = ArrowClosed
            lnStyle.arrowStartStyle = ArrowClosed
        End Select
            
        Select Case StyleCombo(i).SelectedItem.text
        Case LineStyleSolid
            lnStyle.lineStyle = LineSolid
        Case LineStyleDash
            lnStyle.lineStyle = LineDash
        Case LineStyleDot
            lnStyle.lineStyle = LineDot
        Case LineStyleDashDot
            lnStyle.lineStyle = LineDashDot
        Case LineStyleDashDotDot
            lnStyle.lineStyle = LineDashDotDot
        End Select
        
        lnStyle.thickness = ThicknessText(i).text
        ' temporary fix until ChartSkil improves drawing of non-extended lines
        lnStyle.extended = True
        
        studyValueconfig.lineStyle = lnStyle
    
    Case ValueModeBar
        Dim brStyle As barStyle
        
        Set brStyle = mController.defaultBarStyle
        
        brStyle.barColor = IIf(ColorLabel(i).BackColor = NullColor, _
                            -1, _
                            ColorLabel(i).BackColor)
        brStyle.downColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                            -1, _
                            DownColorLabel(i).BackColor)
        brStyle.upColor = UpColorLabel(i).BackColor
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case BarModeBar
            brStyle.displayMode = BarDisplayModes.BarDisplayModeBar
            brStyle.barThickness = ThicknessText(i).text
        Case BarModeCandle
            brStyle.displayMode = BarDisplayModes.BarDisplayModeCandlestick
            brStyle.solidUpBody = False
            brStyle.tailThickness = ThicknessText(i).text
        Case BarModeSolidCandle
            brStyle.displayMode = BarDisplayModes.BarDisplayModeCandlestick
            brStyle.solidUpBody = True
            brStyle.tailThickness = ThicknessText(i).text
        Case BarModeLine
            brStyle.displayMode = BarDisplayModes.BarDisplayModeLine
        End Select
        
        Select Case StyleCombo(0).SelectedItem.text
        Case BarStyleNarrow
            brStyle.barWidth = BarWidthNarrow
        Case BarStyleMedium
            brStyle.barWidth = BarWidthMedium
        Case BarStyleWide
            brStyle.barWidth = BarWidthWide
        Case CustomStyle
            brStyle.barWidth = CSng(StyleCombo(0).SelectedItem.Tag)
        End Select
        
        studyValueconfig.barStyle = brStyle
    
    Case ValueModeText
        Dim txStyle As textStyle

        Set txStyle = mController.defaultLineStyle
        
        txStyle.includeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
        txStyle.Color = ColorLabel(i).BackColor
        txStyle.boxFillColor = IIf(UpColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        UpColorLabel(i).BackColor = NullColor)
        txStyle.boxColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        DownColorLabel(i).BackColor = NullColor)
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case TextDisplayModePlain
            txStyle.box = False
        Case TextDisplayModeWIthBackground
            txStyle.box = True
            txStyle.boxStyle = LineInvisible
            txStyle.boxFillStyle = FillSolid
        Case TextDisplayModeWithBox
            txStyle.box = True
            txStyle.boxStyle = LineInsideSolid
            txStyle.boxFillStyle = FillTransparent
        Case TextDisplayModeWithFilledBox
            txStyle.box = True
            txStyle.boxStyle = LineInsideSolid
            txStyle.boxFillStyle = FillSolid
        End Select
            
        If TypeName(FontButton(i).Tag) <> "Nothing" Then
            txStyle.Font = mFonts(i)
        End If
        
        txStyle.boxThickness = ThicknessText(i).text
        ' temporary fix until ChartSkil improves drawing of non-extended texts
        txStyle.extended = True
        
        studyValueconfig.textStyle = txStyle
    

    End Select
    
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
                ByVal controller As chartController, _
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

Set mController = controller
Set mStudyDefinition = studyDef
mStudyLibraryName = StudyLibraryName
Set mConfiguredStudies = configuredStudies
Set mDefaultConfiguration = defaultConfiguration

processRegionNames regionNames

setupBaseStudiesCombo

processStudyDefinition defaultParameters
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function chooseAColor( _
                ByVal initialColor As Long, _
                ByVal allowNull As Boolean) As Long
Dim simpleColorPicker As New fSimpleColorPicker
Dim cursorpos As W32Point

GetCursorPos cursorpos

simpleColorPicker.Top = cursorpos.y * Screen.TwipsPerPixelY
simpleColorPicker.Left = cursorpos.x * Screen.TwipsPerPixelX
simpleColorPicker.initialColor = initialColor
If allowNull Then simpleColorPicker.NoColorButton.Enabled = True
simpleColorPicker.Show vbModal, UserControl
chooseAColor = simpleColorPicker.selectedColor
Unload simpleColorPicker
End Function

Private Sub initialiseControls()
Dim i As Long

On Error Resume Next

For i = InputValueNameLabel.ubound To 1 Step -1
    Unload InputValueNameLabel(i)
Next
InputValueNameLabel(0).Caption = ""
InputValueNameLabel(0).Visible = False

For i = InputValueCombo.ubound To 1 Step -1
    Unload InputValueCombo(i)
Next

For i = ParameterNameLabel.ubound To 1 Step -1
    Unload ParameterNameLabel(i)
Next
ParameterNameLabel(0).Caption = ""
ParameterNameLabel(0).Visible = False

For i = ParameterValueText.ubound To 1 Step -1
    Unload ParameterValueText(i)
Next
ParameterValueText(0).text = ""
ParameterValueText(0).Visible = False

For i = ParameterValueCombo.ubound To 1 Step -1
    Unload ParameterValueCombo(i)
Next
ParameterValueCombo(0).text = ""
ParameterValueCombo(0).ComboItems.clear
ParameterValueCombo(0).Visible = False

For i = ParameterValueCheck.ubound To 1 Step -1
    Unload ParameterValueCheck(i)
Next
ParameterValueCombo(0).Visible = False

For i = ParameterValueUpDown.ubound To 1 Step -1
    Unload ParameterValueUpDown(i)
Next
ParameterValueUpDown(0).Visible = False

For i = IncludeCheck.ubound To 1 Step -1
    Unload IncludeCheck(i)
Next
IncludeCheck(0).value = vbUnchecked

For i = ValueNameLabel.ubound To 1 Step -1
    Unload ValueNameLabel(i)
Next
ValueNameLabel(0).Caption = ""

For i = AutoscaleCheck.ubound To 1 Step -1
    Unload AutoscaleCheck(i)
Next
AutoscaleCheck(0).value = vbUnchecked

For i = ColorLabel.ubound To 1 Step -1
    Unload ColorLabel(i)
Next
ColorLabel(0).BackColor = vbBlue

For i = UpColorLabel.ubound To 1 Step -1
    Unload UpColorLabel(i)
Next
UpColorLabel(0).BackColor = vbGreen

For i = DownColorLabel.ubound To 1 Step -1
    Unload DownColorLabel(i)
Next
DownColorLabel(0).BackColor = vbRed

For i = DisplayModeCombo.ubound To 1 Step -1
    Unload DisplayModeCombo(i)
Next
DisplayModeCombo(0).ComboItems(0).selected = True

For i = ThicknessText.ubound To 1 Step -1
    Unload ThicknessText(i)
Next
ThicknessText(0).text = "1"

For i = ThicknessUpDown.ubound To 1 Step -1
    Unload ThicknessUpDown(i)
Next

For i = StyleCombo.ubound To 1 Step -1
    Unload StyleCombo(i)
Next
StyleCombo(0).ComboItems(0).selected = True

For i = FontButton.ubound To 1 Step -1
    Unload FontButton(i)
Next

ReDim mFonts(0) As StdFont

For i = AdvancedButton.ubound To 1 Step -1
    Unload AdvancedButton(i)
Next

End Sub

Private Sub initialiseBarDisplayModeCombo( _
                ByVal combo As ImageCombo, _
                ByVal pDisplayMode As BarDisplayModes, _
                ByVal pSolid As Boolean)
Dim item As ComboItem
combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , BarModeBar)
If pDisplayMode = BarDisplayModeBar Then item.selected = True

Set item = combo.ComboItems.Add(, , BarModeCandle)
If pDisplayMode = BarDisplayModeCandlestick And Not pSolid Then item.selected = True

Set item = combo.ComboItems.Add(, , BarModeSolidCandle)
If pDisplayMode = BarDisplayModeCandlestick And pSolid Then item.selected = True

Set item = combo.ComboItems.Add(, , BarModeLine)
If pDisplayMode = BarDisplayModeLine Then item.selected = True

combo.Refresh
End Sub

Private Sub initialiseBarStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal barWidth As Single)
Dim item As ComboItem
Dim selected As Boolean

combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , BarStyleMedium)
If barWidth = BarWidthMedium Then item.selected = True: selected = True

Set item = combo.ComboItems.Add(, , BarStyleNarrow)
If barWidth = BarWidthNarrow Then item.selected = True: selected = True

Set item = combo.ComboItems.Add(, , BarStyleWide)
If barWidth = BarWidthWide Then item.selected = True: selected = True

If Not selected Then
    Set item = combo.ComboItems.Add(1, , CustomStyle)
    item.selected = True
    item.Tag = barWidth
End If

combo.Refresh
End Sub

Private Sub initialiseHistogramStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal histBarWidth As Single)
Dim item As ComboItem
Dim selected As Boolean

combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , HistogramStyleMedium)
If histBarWidth = HistogramWidthMedium Then item.selected = True: selected = True

Set item = combo.ComboItems.Add(, , HistogramStyleNarrow)
If histBarWidth = HistogramWidthNarrow Then item.selected = True: selected = True

Set item = combo.ComboItems.Add(, , HistogramStyleWide)
If histBarWidth = HistogramWidthWide Then item.selected = True: selected = True

If Not selected Then
    Set item = combo.ComboItems.Add(1, , CustomStyle)
    item.selected = True
    item.Tag = histBarWidth
End If

combo.Refresh
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
Set studyValueDefs = mConfiguredStudies.item(item.Key).study.StudyDefinition.studyValueDefinitions
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
    InputValueCombo(Index).ComboItems(IIf(selIndex <> 0, selIndex, 1)).selected = True
End If

InputValueCombo(Index).Refresh
End Sub

Private Sub initialiseLineDisplayModeCombo( _
                ByVal combo As ImageCombo, _
                ByVal pArrowStart As Boolean, _
                ByVal pArrowEnd As Boolean)
Dim item As ComboItem
combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , LineDisplayModePlain)
If Not pArrowStart And Not pArrowEnd Then item.selected = True

Set item = combo.ComboItems.Add(, , LineDisplayModeArrowEnd)
If Not pArrowStart And pArrowEnd Then item.selected = True

Set item = combo.ComboItems.Add(, , LineDisplayModeArrowStart)
If pArrowStart And Not pArrowEnd Then item.selected = True

Set item = combo.ComboItems.Add(, , LineDisplayModeArrowBoth)
If pArrowStart And pArrowEnd Then item.selected = True

combo.Refresh
End Sub

Private Sub initialiseLineStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal pLineStyle As LineStyles)
Dim item As ComboItem

combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , LineStyleSolid)
If pLineStyle = LineSolid Then item.selected = True

Set item = combo.ComboItems.Add(, , LineStyleDash)
If pLineStyle = LineDash Then item.selected = True

Set item = combo.ComboItems.Add(, , LineStyleDot)
If pLineStyle = LineDot Then item.selected = True

Set item = combo.ComboItems.Add(, , LineStyleDashDot)
If pLineStyle = LineDashDot Then item.selected = True

Set item = combo.ComboItems.Add(, , LineStyleDashDotDot)
If pLineStyle = LineDashDotDot Then item.selected = True

combo.Refresh
End Sub

Private Sub initialisePointDisplayModeCombo( _
                ByVal combo As ImageCombo, _
                ByVal pDisplayMode As DataPointDisplayModes)
Dim item As ComboItem

combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , PointDisplayModeLine)
If pDisplayMode = DataPointDisplayModeLine Then item.selected = True

Set item = combo.ComboItems.Add(, , PointDisplayModePoint)
If pDisplayMode = DataPointDisplayModePoint Then item.selected = True

Set item = combo.ComboItems.Add(, , PointDisplayModeSteppedLine)
If pDisplayMode = DataPointDisplayModeStep Then item.selected = True

Set item = combo.ComboItems.Add(, , PointDisplayModeHistogram)
If pDisplayMode = DataPointDisplayModeHistogram Then item.selected = True

combo.Refresh
End Sub

Private Sub initialisePointStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal pPointStyle As PointStyles)
Dim item As ComboItem

combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , PointStyleRound)
If pPointStyle = PointRound Then item.selected = True

Set item = combo.ComboItems.Add(, , PointStyleSquare)
If pPointStyle = PointSquare Then item.selected = True

combo.Refresh
End Sub

Private Sub initialiseTextDisplayModeCombo( _
                ByVal combo As ImageCombo, _
                ByVal pBox As Boolean, _
                ByVal pBoxThickness As Long, _
                ByVal pBoxStyle As LineStyles, _
                ByVal pBoxColor As Long, _
                ByVal pBoxFillStyle As FillStyles, _
                ByVal pBoxFillColor As Long)
Dim item As ComboItem
Dim selected As Boolean

combo.ComboItems.clear

Set item = combo.ComboItems.Add(, , TextDisplayModePlain)
If Not pBox Then item.selected = True: selected = True

Set item = combo.ComboItems.Add(, , TextDisplayModeWIthBackground)
If pBox And (pBoxStyle = LineInvisible Or pBoxThickness = 0) And pBoxFillStyle = FillSolid Then item.selected = True: selected = True

Set item = combo.ComboItems.Add(, , TextDisplayModeWithBox)
If pBox And pBoxStyle <> LineInvisible And pBoxThickness = 0 And pBoxFillStyle = FillTransparent Then item.selected = True: selected = True

Set item = combo.ComboItems.Add(, , TextDisplayModeWithFilledBox)
If pBox And pBoxStyle <> LineInvisible And pBoxThickness = 0 And pBoxFillStyle = FillSolid Then item.selected = True: selected = True

If Not selected Then
    Set item = combo.ComboItems.Add(, , CustomDisplayMode)
End If

combo.Refresh
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
ChartRegionCombo.ComboItems.item(1).selected = True
ChartRegionCombo.Refresh
End Sub

Private Sub processStudyDefinition( _
                ByVal defaultParams As Parameters2.Parameters)
Dim i As Long
Dim studyInputDefinitions As studyInputDefinitions
Dim studyParameterDefinitions As studyParameterDefinitions
Dim studyValueDefinitions As studyValueDefinitions
Dim studyinput As StudyInputDefinition
Dim studyParam As StudyParameterDefinition
Dim studyValueDef As StudyValueDefinition
Dim studyValueconfigs As StudyValueConfigurations
Dim studyValueconfig As StudyValueConfiguration
Dim studyHorizRules As StudyHorizontalRules
Dim studyHorizRule As StudyHorizontalRule
Dim firstParamIsInteger As Boolean
Dim permittedParamValues() As Variant
Dim permittedParamValue As Variant
Dim numPermittedParamValues As Long
Dim defaultParamValue As String

mNextTabIndex = 2

mStudyname = mStudyDefinition.name

If Not mDefaultConfiguration Is Nothing Then
    Set defaultParams = mDefaultConfiguration.Parameters
    Set studyValueconfigs = mDefaultConfiguration.StudyValueConfigurations
    Set studyHorizRules = mDefaultConfiguration.StudyHorizontalRules
End If

StudyDescriptionText.text = mStudyDefinition.Description

If Not mDefaultConfiguration Is Nothing Then
    If mDefaultConfiguration.chartRegionName = mDefaultConfiguration.instanceFullyQualifiedName Then
        '
        setComboSelection ChartRegionCombo, RegionCustom
    Else
        setComboSelection ChartRegionCombo, mDefaultConfiguration.chartRegionName
    End If
    
    For i = 1 To BaseStudiesCombo.ComboItems.Count
        If BaseStudiesCombo.ComboItems(i).Tag = mDefaultConfiguration.underlyingStudy.id Then
            BaseStudiesCombo.ComboItems(i).selected = True
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
    If Not mDefaultConfiguration Is Nothing Then
        Dim inputValueNames() As String
        inputValueNames = mDefaultConfiguration.inputValueNames
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
DisplayModeCombo(0).TabIndex = nextTabIndex
ThicknessText(0).TabIndex = nextTabIndex
ThicknessUpDown(0).TabIndex = nextTabIndex
StyleCombo(0).TabIndex = nextTabIndex

Set studyValueDefinitions = mStudyDefinition.studyValueDefinitions
For i = 1 To studyValueDefinitions.Count
    Set studyValueDef = studyValueDefinitions.item(i)
    If Not studyValueconfigs Is Nothing Then
        Set studyValueconfig = studyValueconfigs.item(i)
    End If
    
    If i = 1 Then
        UpColorLabel(0).Visible = False
        DownColorLabel(0).Visible = False
        
        DisplayModeCombo(0).Visible = False
        
        StyleCombo(0).Visible = False
        
        FontButton(0).Visible = False
    Else
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
    
        Load UpColorLabel(i - 1)
        UpColorLabel(i - 1).Top = UpColorLabel(i - 2).Top + 360
        UpColorLabel(i - 1).Left = UpColorLabel(i - 2).Left
        UpColorLabel(i - 1).TabIndex = nextTabIndex
    
        Load DownColorLabel(i - 1)
        DownColorLabel(i - 1).Top = DownColorLabel(i - 2).Top + 360
        DownColorLabel(i - 1).Left = DownColorLabel(i - 2).Left
        DownColorLabel(i - 1).TabIndex = nextTabIndex
    
        Load DisplayModeCombo(i - 1)
        DisplayModeCombo(i - 1).Top = DisplayModeCombo(i - 2).Top + 360
        DisplayModeCombo(i - 1).Left = DisplayModeCombo(i - 2).Left
        DisplayModeCombo(i - 1).TabIndex = nextTabIndex
    
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
        StyleCombo(i - 1).TabIndex = nextTabIndex
        
        Load FontButton(i - 1)
        FontButton(i - 1).Top = StyleCombo(i - 2).Top + 360
        FontButton(i - 1).TabIndex = nextTabIndex
        
        ReDim mFonts(UBound(mFonts) + 1) As StdFont
    
        Load AdvancedButton(i - 1)
        AdvancedButton(i - 1).Top = AdvancedButton(i - 2).Top + 360
        AdvancedButton(i - 1).Left = AdvancedButton(i - 2).Left
        AdvancedButton(i - 1).Visible = True
        AdvancedButton(i - 1).TabIndex = nextTabIndex
        
    End If
    
    AutoscaleCheck(i - 1) = vbChecked
    
    ValueNameLabel(i - 1).Caption = studyValueDef.name
    ValueNameLabel(i - 1).ToolTipText = studyValueDef.Description

    Select Case studyValueDef.valueMode
    Case ValueModeNone
        Dim dpStyle As dataPointStyle
        
        UpColorLabel(i - 1).Visible = True
        DownColorLabel(i - 1).Visible = True
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            IncludeCheck(i - 1) = IIf(studyValueconfig.includeInChart, vbChecked, vbUnchecked)
            Set dpStyle = studyValueconfig.dataPointStyle
        Else
            Set dpStyle = mController.defaultDataPointStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(dpStyle.includeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = dpStyle.Color
        UpColorLabel(i - 1).BackColor = IIf(dpStyle.upColor = -1, NullColor, dpStyle.upColor)
        DownColorLabel(i - 1).BackColor = IIf(dpStyle.downColor = -1, NullColor, dpStyle.downColor)
        
        initialisePointDisplayModeCombo DisplayModeCombo(i - 1), dpStyle.displayMode
        Select Case dpStyle.displayMode
        Case DataPointDisplayModes.DataPointDisplayModeLine
            initialiseLineStyleCombo StyleCombo(i - 1), dpStyle.lineStyle
        Case DataPointDisplayModes.DataPointDisplayModePoint
            initialisePointStyleCombo StyleCombo(i - 1), dpStyle.pointStyle
        Case DataPointDisplayModes.DataPointDisplayModeStep
            initialiseLineStyleCombo StyleCombo(i - 1), dpStyle.lineStyle
        Case DataPointDisplayModes.DataPointDisplayModeHistogram
            initialiseHistogramStyleCombo StyleCombo(i - 1), dpStyle.histBarWidth
        End Select
        
        ThicknessText(i - 1).text = dpStyle.lineThickness
        
    Case ValueModeLine
        Dim lnStyle As lineStyle
        
        UpColorLabel(i - 1).Visible = True
        DownColorLabel(i - 1).Visible = True
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            IncludeCheck(i - 1) = IIf(studyValueconfig.includeInChart, vbChecked, vbUnchecked)
            Set lnStyle = studyValueconfig.dataPointStyle
        Else
            Set lnStyle = mController.defaultLineStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(lnStyle.includeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = lnStyle.Color
        UpColorLabel(i - 1).BackColor = IIf(lnStyle.arrowStartColor = -1, NullColor, lnStyle.arrowStartColor)
        DownColorLabel(i - 1).BackColor = IIf(lnStyle.arrowEndColor = -1, NullColor, lnStyle.arrowEndColor)
        
        initialiseLineDisplayModeCombo DisplayModeCombo(i - 1), _
                                        (lnStyle.arrowStartStyle <> ArrowNone), _
                                        (lnStyle.arrowEndStyle <> ArrowNone)
    
        initialiseLineStyleCombo StyleCombo(i - 1), lnStyle.lineStyle
        
        ThicknessText(i - 1).text = lnStyle.thickness
        
    Case ValueModeBar
        Dim brStyle As barStyle
        
        UpColorLabel(i - 1).Visible = True
        DownColorLabel(i - 1).Visible = True
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            IncludeCheck(i - 1) = IIf(studyValueconfig.includeInChart, vbChecked, vbUnchecked)
            Set brStyle = studyValueconfig.barStyle
        Else
            Set brStyle = mController.defaultBarStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(brStyle.includeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = IIf(brStyle.barColor = -1, NullColor, brStyle.barColor)
        UpColorLabel(i - 1).BackColor = IIf(brStyle.upColor = -1, NullColor, brStyle.upColor)
        DownColorLabel(i - 1).BackColor = IIf(brStyle.downColor = -1, NullColor, brStyle.downColor)
        
        initialiseBarDisplayModeCombo DisplayModeCombo(i - 1), _
                                        brStyle.displayMode, _
                                        brStyle.solidUpBody
        
        initialiseBarStyleCombo StyleCombo(i - 1), brStyle.barWidth
        
        Select Case DisplayModeCombo(i - 1).SelectedItem.text
        Case BarModeBar
            ThicknessText(i - 1).text = brStyle.barThickness
        Case BarModeCandle
            ThicknessText(i - 1).text = brStyle.tailThickness
        Case BarModeSolidCandle
            ThicknessText(i - 1).text = brStyle.tailThickness
        Case BarModeLine
            ThicknessText(i - 1).text = brStyle.barThickness
        End Select
        
    Case ValueModeText
        Dim txStyle As textStyle
        
        UpColorLabel(i - 1).Visible = True      ' box fill color
        DownColorLabel(i - 1).Visible = True    ' box outline color
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = False
        FontButton(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            IncludeCheck(i - 1) = IIf(studyValueconfig.includeInChart, vbChecked, vbUnchecked)
            Set txStyle = studyValueconfig.textStyle
        Else
            Set txStyle = mController.defaultTextStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(txStyle.includeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = txStyle.Color
        UpColorLabel(i - 1).BackColor = IIf(txStyle.boxFillColor = -1, NullColor, txStyle.boxFillColor)
        DownColorLabel(i - 1).BackColor = IIf(txStyle.boxColor = -1, NullColor, txStyle.boxColor)
        
        initialiseTextDisplayModeCombo DisplayModeCombo(i - 1), _
                                        txStyle.box, _
                                        txStyle.boxThickness, _
                                        txStyle.boxStyle, _
                                        txStyle.boxColor, _
                                        txStyle.boxFillStyle, _
                                        txStyle.boxFillColor
    
        ThicknessText(i - 1).text = txStyle.boxThickness
        
        Set mFonts(i - 1) = txStyle.Font
        
    End Select

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
        item.selected = True
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
    If Not TypeOf studyConfig.study Is InputStudy Or _
        Not mStudyDefinition.needsBars _
    Then
        If studiesCompatible(studyConfig.study.StudyDefinition, mStudyDefinition) Then
            Set item = BaseStudiesCombo.ComboItems.Add(, studyConfig.instanceFullyQualifiedName, studyConfig.study.instanceName)
            item.Tag = studyConfig.study.id
        End If
    End If
Next
BaseStudiesCombo.ComboItems(1).selected = True
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



