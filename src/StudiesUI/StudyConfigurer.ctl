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
      TabIndex        =   36
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
         TabIndex        =   37
         Top             =   240
         Width           =   6900
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   4
            Left            =   5280
            TabIndex        =   42
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   3
            Left            =   3960
            TabIndex        =   41
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   40
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   39
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   6000
            TabIndex        =   47
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   4680
            TabIndex        =   46
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   45
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   44
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   43
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame ValuesFrame 
      Caption         =   "Output values"
      Height          =   4095
      Left            =   5040
      TabIndex        =   16
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
         TabIndex        =   17
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
            TabIndex        =   49
            ToolTipText     =   "Click to select the font"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton AdvancedButton 
            Caption         =   "..."
            Height          =   375
            Index           =   0
            Left            =   6360
            TabIndex        =   27
            ToolTipText     =   "Click for advanced features"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox AutoscaleCheck 
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   23
            ToolTipText     =   "Set this to ensure that all values are visible when the chart is auto-scaling"
            Top             =   240
            Width           =   210
         End
         Begin VB.TextBox ThicknessText 
            Alignment       =   2  'Center
            Height          =   330
            Index           =   0
            Left            =   4320
            TabIndex        =   24
            Text            =   "1"
            ToolTipText     =   "Choose the thickness of lines or points"
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox IncludeCheck 
            Height          =   195
            Index           =   0
            Left            =   1560
            TabIndex        =   18
            ToolTipText     =   "Set to include this study value in the chart"
            Top             =   240
            Width           =   195
         End
         Begin MSComctlLib.ImageCombo StyleCombo 
            Height          =   330
            Index           =   0
            Left            =   5160
            TabIndex        =   26
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
            Left            =   4800
            TabIndex        =   25
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "ThicknessText(0)"
            BuddyDispid     =   196618
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
            TabIndex        =   33
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Value name"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Show"
            Height          =   255
            Left            =   1320
            TabIndex        =   48
            Top             =   0
            Width           =   495
         End
         Begin VB.Label DownColorLabel 
            BackColor       =   &H000000FF&
            Height          =   330
            Index           =   0
            Left            =   2865
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
         Begin VB.Label UpColorLabel 
            BackColor       =   &H0000FF00&
            Height          =   330
            Index           =   0
            Left            =   2520
            TabIndex        =   20
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
            TabIndex        =   19
            ToolTipText     =   "Click to change the colour for this value"
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "Advanced"
            Height          =   255
            Left            =   6120
            TabIndex        =   35
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Style"
            Height          =   255
            Left            =   5040
            TabIndex        =   34
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Thickness"
            Height          =   255
            Left            =   4320
            TabIndex        =   32
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Display as"
            Height          =   255
            Left            =   3120
            TabIndex        =   31
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Colors"
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            Top             =   0
            Width           =   495
         End
         Begin VB.Label ValueNameLabel 
            Caption         =   "Label2"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4935
      Left            =   2520
      TabIndex        =   12
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
         TabIndex        =   13
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox ParameterValueCheck 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComctlLib.ImageCombo ParameterValueCombo 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   3
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
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox ParameterValueText 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
         Begin MSComCtl2.UpDown ParameterValueUpDown 
            Height          =   330
            Index           =   0
            Left            =   1920
            TabIndex        =   14
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
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
         Begin VB.Label ParameterNameLabel 
            Caption         =   "Param name"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   15
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5040
      Width           =   12135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Inputs"
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.TreeView BaseStudiesTree 
         Height          =   1815
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3201
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         SingleSel       =   -1  'True
         Appearance      =   0
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   2175
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         Begin MSComctlLib.ImageCombo InputValueCombo 
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Top             =   3000
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
            TabIndex        =   10
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Base study"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label InputValueNameLabel 
            Caption         =   "Input value"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   2760
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
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyConfigurer"

Private Const CompatibleNode As String = "YES"

Private Const RegionDefault As String = "Use default"
Private Const RegionCustom As String = "Use own region"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mChart As ChartController
Private mStudyname As String
Private mStudyLibraryName As String

Private mStudyDefinition As StudyDefinition

Private mBaseStudyConfig As StudyConfiguration

Private mNextTabIndex As Long

Private mDefaultConfiguration As StudyConfiguration

Private mFonts() As StdFont

Private mFirstCompatibleNode As Node

Private mCompatibleStudies As Collection

Private mPrevSelectedBaseStudiesTreeNode As Node

Private mStudyValueRegionNames() As String

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
mNextTabIndex = 2
End Sub

'@================================================================================
' XXXX Interface members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AdvancedButton_Click(Index As Integer)
Const ProcName As String = "AdvancedButton_Click"
On Error GoTo Err

notImplemented

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub BaseStudiesTree_Click()
Dim i As Long
Const ProcName As String = "BaseStudiesTree_Click"
On Error GoTo Err

BaseStudiesTree.SelectedItem.Expanded = True
If Not BaseStudiesTree.SelectedItem.Tag = CompatibleNode Then
    Set BaseStudiesTree.SelectedItem = mPrevSelectedBaseStudiesTreeNode
Else
    Set mPrevSelectedBaseStudiesTreeNode = BaseStudiesTree.SelectedItem
    For i = 0 To mStudyDefinition.studyInputDefinitions.Count - 1
        initialiseInputValueCombo i
    Next
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub ColorLabel_Click( _
                Index As Integer)
Dim studyValueDef As StudyValueDefinition
Const ProcName As String = "ColorLabel_Click"
On Error GoTo Err

Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

ColorLabel(Index).BackColor = chooseAColor(ColorLabel(Index).BackColor, _
                                            IIf(studyValueDef.ValueMode = ValueModeBar, True, False))

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub DisplayModeCombo_Click(Index As Integer)
Dim studyValueDef  As StudyValueDefinition
Dim studyValueconfig  As StudyValueConfiguration

Const ProcName As String = "DisplayModeCombo_Click"
On Error GoTo Err

Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

If Not mDefaultConfiguration Is Nothing Then
    Set studyValueconfig = mDefaultConfiguration.StudyValueConfigurations.item(studyValueDef.name)
End If

Select Case studyValueDef.ValueMode
Case ValueModeNone
    Dim dpStyle As DataPointStyle
    
    If Not studyValueconfig Is Nothing Then
        Set dpStyle = studyValueconfig.DataPointStyle
    Else
        Set dpStyle = New DataPointStyle
    End If
        
    Select Case DisplayModeCombo(Index).SelectedItem.text
    Case PointDisplayModeLine
        initialiseLineStyleCombo StyleCombo(Index), dpStyle.LineStyle
    Case PointDisplayModePoint
        initialisePointStyleCombo StyleCombo(Index), dpStyle.PointStyle
    Case PointDisplayModeSteppedLine
        initialiseLineStyleCombo StyleCombo(Index), dpStyle.LineStyle
    Case PointDisplayModeHistogram
        initialiseHistogramStyleCombo StyleCombo(Index), dpStyle.HistogramBarWidth
    End Select
Case ValueModeLine

Case ValueModeBar

Case ValueModeText

End Select

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub DisplayModeCombo_Validate( _
                Index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "DisplayModeCombo_Validate"
On Error GoTo Err

If DisplayModeCombo(Index).SelectedItem Is Nothing Then Cancel = True

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub DownColorLabel_Click(Index As Integer)
Dim studyValueDef As StudyValueDefinition
Dim allowNullColor As Boolean

Const ProcName As String = "DownColorLabel_Click"
On Error GoTo Err

Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

If studyValueDef.ValueMode = ValueModeBar Or _
    studyValueDef.ValueMode = ValueModeNone Then allowNullColor = True

DownColorLabel(Index).BackColor = chooseAColor(DownColorLabel(Index).BackColor, _
                                            allowNullColor)

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub FontButton_Click(Index As Integer)
Dim aFont As StdFont

Const ProcName As String = "FontButton_Click"
On Error GoTo Err

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

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub LineColorLabel_Click(Index As Integer)
Const ProcName As String = "LineColorLabel_Click"
On Error GoTo Err

LineColorLabel(Index).BackColor = chooseAColor(LineColorLabel(Index).BackColor, False)

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub StyleCombo_Validate( _
                Index As Integer, _
                Cancel As Boolean)
Const ProcName As String = "StyleCombo_Validate"
On Error GoTo Err

If StyleCombo(Index).SelectedItem Is Nothing Then Cancel = True

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub ThicknessText_KeyPress(Index As Integer, KeyAscii As Integer)
Const ProcName As String = "ThicknessText_KeyPress"
On Error GoTo Err

filterNonNumericKeyPress KeyAscii

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub UpColorLabel_Click(Index As Integer)
Dim studyValueDef As StudyValueDefinition
Dim allowNullColor As Boolean

Const ProcName As String = "UpColorLabel_Click"
On Error GoTo Err

Set studyValueDef = mStudyDefinition.studyValueDefinitions.item(Index + 1)

If studyValueDef.ValueMode = ValueModeBar Or _
    studyValueDef.ValueMode = ValueModeNone Then allowNullColor = True

UpColorLabel(Index).BackColor = chooseAColor(UpColorLabel(Index).BackColor, _
                                            allowNullColor)

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get StudyConfiguration() As StudyConfiguration
Dim studyConfig As StudyConfiguration
Dim params As Parameters
Dim studyParamDef As StudyParameterDefinition
Dim studyValueDefs As studyValueDefinitions
Dim studyValueDef As StudyValueDefinition
Dim studyValueconfig As StudyValueConfiguration
Dim studyHorizRule As StudyHorizontalRule
Dim regionName As String
Dim inputValueNames() As String
Dim i As Long

Const ProcName As String = "StudyConfiguration"
On Error GoTo Err

Set studyConfig = New StudyConfiguration
'studyConfig.studyDefinition = mStudyDefinition
studyConfig.name = mStudyname
studyConfig.StudyLibraryName = mStudyLibraryName
If Not BaseStudiesTree.SelectedItem Is Nothing Then
    studyConfig.UnderlyingStudy = mCompatibleStudies(BaseStudiesTree.SelectedItem.Key)
End If

ReDim inputValueNames(mStudyDefinition.studyInputDefinitions.Count - 1) As String
For i = 0 To UBound(inputValueNames)
    If Not InputValueCombo(i).SelectedItem Is Nothing Then
        inputValueNames(i) = InputValueCombo(i).SelectedItem.text
    End If
Next
studyConfig.inputValueNames = inputValueNames

If ChartRegionCombo.SelectedItem.text = RegionDefault Then
    Select Case mStudyDefinition.DefaultRegion
    Case DefaultRegionNone
        regionName = ChartRegionNameDefault
    Case DefaultRegionCustom
        regionName = ChartRegionNameCustom
    End Select
ElseIf ChartRegionCombo.SelectedItem.text = RegionCustom Then
    regionName = ChartRegionNameCustom
Else
    regionName = ChartRegionCombo.SelectedItem.text
End If
studyConfig.ChartRegionName = regionName

Set params = New Parameters

For i = 0 To mStudyDefinition.studyParameterDefinitions.Count - 1
    Set studyParamDef = mStudyDefinition.studyParameterDefinitions.item(i + 1)
    If studyParamDef.ParameterType = ParameterTypeBoolean Then
        params.SetParameterValue ParameterNameLabel(i).Caption, _
                                IIf(ParameterValueCheck(i) = vbChecked, "True", "False")
    ElseIf ParameterValueText(i).Visible Then
        params.SetParameterValue ParameterNameLabel(i).Caption, ParameterValueText(i).text
    Else
        params.SetParameterValue ParameterNameLabel(i).Caption, ParameterValueCombo(i).text
    End If
Next

studyConfig.Parameters = params

Set studyValueDefs = mStudyDefinition.studyValueDefinitions

For i = 0 To ValueNameLabel.UBound
    Set studyValueDef = studyValueDefs.item(i + 1)
    
    Set studyValueconfig = studyConfig.StudyValueConfigurations.Add(ValueNameLabel(i).Caption)
    studyValueconfig.IncludeInChart = (IncludeCheck(i).value = vbChecked)
    
    If mStudyValueRegionNames(i) <> "" Then
        studyValueconfig.ChartRegionName = mStudyValueRegionNames(i)
    Else
        Select Case studyValueDef.DefaultRegion
        Case DefaultRegionNone
            studyValueconfig.ChartRegionName = ChartRegionNameDefault
        Case DefaultRegionCustom
            studyValueconfig.ChartRegionName = ChartRegionNameCustom
        End Select
    End If
    
    Select Case studyValueDef.ValueMode
    Case ValueModeNone
        Dim dpStyle As DataPointStyle
        
        Set dpStyle = New DataPointStyle
        
        dpStyle.IncludeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
        dpStyle.Color = ColorLabel(i).BackColor
        dpStyle.DownColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                            -1, _
                            DownColorLabel(i).BackColor)
        dpStyle.UpColor = IIf(UpColorLabel(i).BackColor = NullColor, _
                            -1, _
                            UpColorLabel(i).BackColor)
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case PointDisplayModeLine
            dpStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
            Select Case StyleCombo(i).SelectedItem.text
            Case LineStyleSolid
                dpStyle.LineStyle = LineSolid
            Case LineStyleDash
                dpStyle.LineStyle = LineDash
            Case LineStyleDot
                dpStyle.LineStyle = LineDot
            Case LineStyleDashDot
                dpStyle.LineStyle = LineDashDot
            Case LineStyleDashDotDot
                dpStyle.LineStyle = LineDashDotDot
            End Select
        Case PointDisplayModePoint
            dpStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint
            Select Case StyleCombo(i).SelectedItem.text
            Case PointStyleRound
                dpStyle.PointStyle = PointRound
            Case PointStyleSquare
                dpStyle.PointStyle = PointSquare
            End Select
        Case PointDisplayModeSteppedLine
            dpStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep
            Select Case StyleCombo(i).SelectedItem.text
            Case LineStyleSolid
                dpStyle.LineStyle = LineSolid
            Case LineStyleDash
                dpStyle.LineStyle = LineDash
            Case LineStyleDot
                dpStyle.LineStyle = LineDot
            Case LineStyleDashDot
                dpStyle.LineStyle = LineDashDot
            Case LineStyleDashDotDot
                dpStyle.LineStyle = LineDashDotDot
            End Select
        Case PointDisplayModeHistogram
            dpStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram
            Select Case StyleCombo(i).SelectedItem.text
            Case HistogramStyleNarrow
                dpStyle.HistogramBarWidth = HistogramWidthNarrow
            Case HistogramStyleMedium
                dpStyle.HistogramBarWidth = HistogramWidthMedium
            Case HistogramStyleWide
                dpStyle.HistogramBarWidth = HistogramWidthWide
            Case CustomStyle
                dpStyle.HistogramBarWidth = CSng(StyleCombo(i).SelectedItem.Tag)
            End Select
        End Select
        
        dpStyle.LineThickness = ThicknessText(i).text
        
        studyValueconfig.DataPointStyle = dpStyle
    Case ValueModeLine
        Dim lnStyle As LineStyle

        Set lnStyle = New LineStyle
        
        lnStyle.IncludeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
        lnStyle.Color = ColorLabel(i).BackColor
        lnStyle.ArrowStartColor = ColorLabel(i).BackColor
        lnStyle.ArrowEndColor = ColorLabel(i).BackColor
        lnStyle.ArrowStartFillColor = IIf(UpColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        UpColorLabel(i).BackColor)
        lnStyle.ArrowEndFillColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        DownColorLabel(i).BackColor)
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case LineDisplayModePlain
            lnStyle.ArrowEndStyle = ArrowNone
            lnStyle.ArrowStartStyle = ArrowNone
        Case LineDisplayModeArrowEnd
            lnStyle.ArrowEndStyle = ArrowClosed
            lnStyle.ArrowStartStyle = ArrowNone
        Case LineDisplayModeArrowStart
            lnStyle.ArrowEndStyle = ArrowNone
            lnStyle.ArrowStartStyle = ArrowClosed
        Case LineDisplayModeArrowBoth
            lnStyle.ArrowEndStyle = ArrowClosed
            lnStyle.ArrowStartStyle = ArrowClosed
        End Select
            
        Select Case StyleCombo(i).SelectedItem.text
        Case LineStyleSolid
            lnStyle.LineStyle = LineSolid
        Case LineStyleDash
            lnStyle.LineStyle = LineDash
        Case LineStyleDot
            lnStyle.LineStyle = LineDot
        Case LineStyleDashDot
            lnStyle.LineStyle = LineDashDot
        Case LineStyleDashDotDot
            lnStyle.LineStyle = LineDashDotDot
        End Select
        
        lnStyle.Thickness = ThicknessText(i).text
        ' temporary fix until ChartSkil improves drawing of non-extended lines
        lnStyle.Extended = True
        
        studyValueconfig.LineStyle = lnStyle
    
    Case ValueModeBar
        Dim brStyle As BarStyle
        
        Set brStyle = New BarStyle
        
        brStyle.Color = IIf(ColorLabel(i).BackColor = NullColor, _
                            -1, _
                            ColorLabel(i).BackColor)
        brStyle.DownColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                            -1, _
                            DownColorLabel(i).BackColor)
        brStyle.UpColor = UpColorLabel(i).BackColor
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case BarModeBar
            brStyle.DisplayMode = BarDisplayModes.BarDisplayModeBar
            brStyle.Thickness = ThicknessText(i).text
        Case BarModeCandle
            brStyle.DisplayMode = BarDisplayModes.BarDisplayModeCandlestick
            brStyle.SolidUpBody = False
            brStyle.TailThickness = ThicknessText(i).text
        Case BarModeSolidCandle
            brStyle.DisplayMode = BarDisplayModes.BarDisplayModeCandlestick
            brStyle.SolidUpBody = True
            brStyle.TailThickness = ThicknessText(i).text
        Case BarModeLine
            brStyle.DisplayMode = BarDisplayModes.BarDisplayModeLine
        End Select
        
        Select Case StyleCombo(i).SelectedItem.text
        Case BarStyleNarrow
            brStyle.Width = BarWidthNarrow
        Case BarStyleMedium
            brStyle.Width = BarWidthMedium
        Case BarStyleWide
            brStyle.Width = BarWidthWide
        Case CustomStyle
            brStyle.Width = CSng(StyleCombo(i).SelectedItem.Tag)
        End Select
        
        studyValueconfig.BarStyle = brStyle
    
    Case ValueModeText
        Dim txStyle As TextStyle

        Set txStyle = New LineStyle
        
        txStyle.IncludeInAutoscale = (AutoscaleCheck(i).value = vbChecked)
        txStyle.Color = ColorLabel(i).BackColor
        txStyle.BoxFillColor = IIf(UpColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        UpColorLabel(i).BackColor)
        txStyle.BoxColor = IIf(DownColorLabel(i).BackColor = NullColor, _
                                        -1, _
                                        DownColorLabel(i).BackColor)
        
        Select Case DisplayModeCombo(i).SelectedItem.text
        Case TextDisplayModePlain
            txStyle.Box = False
        Case TextDisplayModeWIthBackground
            txStyle.Box = True
            txStyle.BoxStyle = LineInvisible
            txStyle.BoxFillStyle = FillSolid
        Case TextDisplayModeWithBox
            txStyle.Box = True
            txStyle.BoxStyle = LineInsideSolid
            txStyle.BoxFillStyle = FillTransparent
        Case TextDisplayModeWithFilledBox
            txStyle.Box = True
            txStyle.BoxStyle = LineInsideSolid
            txStyle.BoxFillStyle = FillSolid
        End Select
            
        If TypeName(FontButton(i).Tag) <> "Nothing" Then
            txStyle.Font = mFonts(i)
        End If
        
        txStyle.BoxThickness = ThicknessText(i).text
        ' temporary fix until ChartSkil improves drawing of non-extended texts
        txStyle.Extended = True
        
        studyValueconfig.TextStyle = txStyle
    

    End Select
    
Next

For i = 0 To 4
    If LineText(i).text <> "" Then
        Set studyHorizRule = studyConfig.StudyHorizontalRules.Add
        studyHorizRule.Y = LineText(i).text
        studyHorizRule.Color = LineColorLabel(i).BackColor
    End If
Next

Set StudyConfiguration = studyConfig

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' methods
'@================================================================================

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Set mPrevSelectedBaseStudiesTreeNode = Nothing
initialiseControls

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Initialise( _
                ByVal pChart As ChartController, _
                ByVal studyDef As StudyDefinition, _
                ByVal StudyLibraryName As String, _
                ByRef regionNames() As String, _
                ByRef baseStudyConfig As StudyConfiguration, _
                ByVal defaultConfiguration As StudyConfiguration, _
                ByVal defaultParameters As Parameters, _
                ByVal noParameterModification As Boolean)
                
Const ProcName As String = "Initialise"
On Error GoTo Err

If Not defaultConfiguration Is Nothing And defaultParameters Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "DefaultConfiguration and DefaultParameters cannot both be Nothing"
End If

initialiseControls

Set mChart = pChart
Set mStudyDefinition = studyDef
mStudyLibraryName = StudyLibraryName
Set mBaseStudyConfig = baseStudyConfig
Set mDefaultConfiguration = defaultConfiguration

processRegionNames regionNames

setupBaseStudiesTree

processStudyDefinition defaultParameters, noParameterModification

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addBaseStudiesTreeEntry( _
                ByVal studyConfig As StudyConfiguration, _
                ByVal parentStudyConfig As StudyConfiguration)
Dim lNode As Node
Dim parentNode As Node
Dim childStudyConfig As StudyConfiguration

Const ProcName As String = "addBaseStudiesTreeEntry"
On Error GoTo Err

If studyConfig Is Nothing Then Exit Sub

If Not mDefaultConfiguration Is Nothing Then
    If mDefaultConfiguration.Study Is studyConfig.Study Then Exit Sub
End If

If parentStudyConfig Is Nothing Then
    Set lNode = BaseStudiesTree.Nodes.Add(, _
                                TreeRelationshipConstants.tvwChild, _
                                studyConfig.Study.Id, _
                                studyConfig.Study.InstanceName)
Else
    Set parentNode = BaseStudiesTree.Nodes.item(parentStudyConfig.Study.Id)
    Set lNode = BaseStudiesTree.Nodes.Add(parentNode, _
                                TreeRelationshipConstants.tvwChild, _
                                studyConfig.Study.Id, _
                                studyConfig.Study.InstanceName)
    parentNode.Expanded = True
End If

If (Not TypeOf studyConfig.Study Is InputStudy Or _
    Not mStudyDefinition.NeedsBars) _
Then
    If studiesCompatible(studyConfig.Study.StudyDefinition, mStudyDefinition) Then
        lNode.Tag = CompatibleNode
        If mPrevSelectedBaseStudiesTreeNode Is Nothing Then Set mPrevSelectedBaseStudiesTreeNode = lNode
        mCompatibleStudies.Add studyConfig.Study, lNode.Key
        If mFirstCompatibleNode Is Nothing Then
            lNode.selected = True
            Set mFirstCompatibleNode = lNode
        End If
    Else
        lNode.BackColor = &HC0C0C0
        lNode.ForeColor = &H808080
    End If
End If

For Each childStudyConfig In studyConfig.StudyConfigurations
    addBaseStudiesTreeEntry childStudyConfig, studyConfig
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function chooseAColor( _
                ByVal initialColor As Long, _
                ByVal allowNull As Boolean) As Long
Dim simpleColorPicker As New fSimpleColorPicker
Dim cursorpos As W32Point

Const ProcName As String = "chooseAColor"
On Error GoTo Err

GetCursorPos cursorpos

simpleColorPicker.Top = cursorpos.Y * Screen.TwipsPerPixelY
simpleColorPicker.Left = cursorpos.X * Screen.TwipsPerPixelX
simpleColorPicker.initialColor = initialColor
If allowNull Then simpleColorPicker.NoColorButton.Enabled = True
simpleColorPicker.Show vbModal, UserControl
chooseAColor = simpleColorPicker.selectedColor
Unload simpleColorPicker

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub initialiseControls()
Dim i As Long

Const ProcName As String = "initialiseControls"
On Error GoTo Err

On Error Resume Next

ReDim mFonts(0) As StdFont

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
ParameterValueCombo(0).ComboItems.Clear
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
ColorLabel(0).BackColor = vbBlue

For i = UpColorLabel.UBound To 1 Step -1
    Unload UpColorLabel(i)
Next
UpColorLabel(0).BackColor = vbGreen

For i = DownColorLabel.UBound To 1 Step -1
    Unload DownColorLabel(i)
Next
DownColorLabel(0).BackColor = vbRed

For i = DisplayModeCombo.UBound To 1 Step -1
    Unload DisplayModeCombo(i)
Next
DisplayModeCombo(0).ComboItems(0).selected = True

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
StyleCombo(0).ComboItems(0).selected = True

For i = FontButton.UBound To 1 Step -1
    Unload FontButton(i)
Next

For i = AdvancedButton.UBound To 1 Step -1
    Unload AdvancedButton(i)
Next

For i = 0 To LineText.UBound
    LineText(i).text = ""
    LineColorLabel(i).BackColor = vbBlack
Next

BaseStudiesTree.Enabled = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialiseBarDisplayModeCombo( _
                ByVal combo As ImageCombo, _
                ByVal pDisplayMode As BarDisplayModes, _
                ByVal pSolid As Boolean)
Dim item As ComboItem
Const ProcName As String = "initialiseBarDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

Set item = combo.ComboItems.Add(, , BarModeBar)
If pDisplayMode = BarDisplayModeBar Then item.selected = True

Set item = combo.ComboItems.Add(, , BarModeCandle)
If pDisplayMode = BarDisplayModeCandlestick And Not pSolid Then item.selected = True

Set item = combo.ComboItems.Add(, , BarModeSolidCandle)
If pDisplayMode = BarDisplayModeCandlestick And pSolid Then item.selected = True

Set item = combo.ComboItems.Add(, , BarModeLine)
If pDisplayMode = BarDisplayModeLine Then item.selected = True

combo.ToolTipText = "Select the type of bar"

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialiseBarStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal barWidth As Single)
Dim item As ComboItem
Dim selected As Boolean

Const ProcName As String = "initialiseBarStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

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

combo.ToolTipText = "Select the width of the bar"

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialiseHistogramStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal histBarWidth As Single)
Dim item As ComboItem
Dim selected As Boolean

Const ProcName As String = "initialiseHistogramStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

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

combo.ToolTipText = "Select the width of the histogram"

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialiseInputValueCombo( _
                ByVal Index As Long)
Dim lstudy As Study
Dim studyValueDefs As studyValueDefinitions
Dim valueDef As StudyValueDefinition
Dim inputDef As StudyInputDefinition
Dim selIndex As Long

Const ProcName As String = "initialiseInputValueCombo"
On Error GoTo Err

Set lstudy = mCompatibleStudies(BaseStudiesTree.SelectedItem.Key)
Set studyValueDefs = lstudy.StudyDefinition.studyValueDefinitions
Set inputDef = mStudyDefinition.studyInputDefinitions.item(Index + 1)

InputValueCombo(Index).ComboItems.Clear

selIndex = -1
'InputValueCombo(Index).ComboItems.Add , , ""
For Each valueDef In studyValueDefs
    If typesCompatible(valueDef.ValueType, inputDef.InputType) Then
        InputValueCombo(Index).ComboItems.Add , , valueDef.name
        If UCase$(inputDef.name) = UCase$(valueDef.name) Then selIndex = InputValueCombo(Index).ComboItems.Count
        If valueDef.IsDefault And _
            selIndex = -1 Then selIndex = InputValueCombo(Index).ComboItems.Count
    End If
Next

If selIndex <> -1 Then
    InputValueCombo(Index).ComboItems(selIndex).selected = True
ElseIf InputValueCombo(Index).ComboItems.Count <> 0 Then
    InputValueCombo(Index).ComboItems(1).selected = True
End If

InputValueCombo(Index).Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialiseLineDisplayModeCombo( _
                ByVal combo As ImageCombo, _
                ByVal pArrowStart As Boolean, _
                ByVal pArrowEnd As Boolean)
Dim item As ComboItem
Const ProcName As String = "initialiseLineDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

Set item = combo.ComboItems.Add(, , LineDisplayModePlain)
If Not pArrowStart And Not pArrowEnd Then item.selected = True

Set item = combo.ComboItems.Add(, , LineDisplayModeArrowEnd)
If Not pArrowStart And pArrowEnd Then item.selected = True

Set item = combo.ComboItems.Add(, , LineDisplayModeArrowStart)
If pArrowStart And Not pArrowEnd Then item.selected = True

Set item = combo.ComboItems.Add(, , LineDisplayModeArrowBoth)
If pArrowStart And pArrowEnd Then item.selected = True

combo.ToolTipText = "Select the type of line"

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialiseLineStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal pLineStyle As LineStyles)
Dim item As ComboItem

Const ProcName As String = "initialiseLineStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

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

combo.ToolTipText = "Select the style of the line"

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialisePointDisplayModeCombo( _
                ByVal combo As ImageCombo, _
                ByVal pDisplayMode As DataPointDisplayModes)
Dim item As ComboItem

Const ProcName As String = "initialisePointDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

Set item = combo.ComboItems.Add(, , PointDisplayModeLine)
If pDisplayMode = DataPointDisplayModeLine Then item.selected = True

Set item = combo.ComboItems.Add(, , PointDisplayModePoint)
If pDisplayMode = DataPointDisplayModePoint Then item.selected = True

Set item = combo.ComboItems.Add(, , PointDisplayModeSteppedLine)
If pDisplayMode = DataPointDisplayModeStep Then item.selected = True

Set item = combo.ComboItems.Add(, , PointDisplayModeHistogram)
If pDisplayMode = DataPointDisplayModeHistogram Then item.selected = True

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initialisePointStyleCombo( _
                ByVal combo As ImageCombo, _
                ByVal pPointStyle As PointStyles)
Dim item As ComboItem

Const ProcName As String = "initialisePointStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

Set item = combo.ComboItems.Add(, , PointStyleRound)
If pPointStyle = PointRound Then item.selected = True

Set item = combo.ComboItems.Add(, , PointStyleSquare)
If pPointStyle = PointSquare Then item.selected = True

combo.ToolTipText = "Select the shape of the point"

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Const ProcName As String = "initialiseTextDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

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

combo.ToolTipText = "Select the type of text"

combo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function nextTabIndex() As Long
Const ProcName As String = "nextTabIndex"
On Error GoTo Err

nextTabIndex = mNextTabIndex
mNextTabIndex = mNextTabIndex + 1

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub processRegionNames( _
                ByRef regionNames() As String)
Dim i As Long

Const ProcName As String = "processRegionNames"
On Error GoTo Err

ChartRegionCombo.ComboItems.Clear

ChartRegionCombo.ComboItems.Add , , RegionDefault
ChartRegionCombo.ComboItems.Add , , RegionCustom

For i = 0 To UBound(regionNames)
    ChartRegionCombo.ComboItems.Add , , regionNames(i)
Next
ChartRegionCombo.ComboItems.item(1).selected = True
ChartRegionCombo.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processStudyDefinition( _
                ByVal defaultParams As Parameters, _
                ByVal noParameterModification As Boolean)
Dim i As Long
Dim j As Long
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
Dim inputValueNames() As String
Dim noInputModification As Boolean

Const ProcName As String = "processStudyDefinition"
On Error GoTo Err

mNextTabIndex = 2

mStudyname = mStudyDefinition.name

If Not mDefaultConfiguration Is Nothing Then
    Set defaultParams = mDefaultConfiguration.Parameters
    Set studyValueconfigs = mDefaultConfiguration.StudyValueConfigurations
    Set studyHorizRules = mDefaultConfiguration.StudyHorizontalRules
End If

StudyDescriptionText.text = mStudyDefinition.Description

If Not mDefaultConfiguration Is Nothing Then
    If mDefaultConfiguration.ChartRegionName = mDefaultConfiguration.InstanceFullyQualifiedName Then
        '
        setComboSelection ChartRegionCombo, RegionCustom
    Else
        setComboSelection ChartRegionCombo, mDefaultConfiguration.ChartRegionName
    End If
    
    If Not mDefaultConfiguration.UnderlyingStudy Is Nothing Then
        If TypeOf mDefaultConfiguration.UnderlyingStudy Is InputStudy Then
            noInputModification = True
            mCompatibleStudies.Add mDefaultConfiguration.UnderlyingStudy, mDefaultConfiguration.UnderlyingStudy.Id
            BaseStudiesTree.Nodes.Clear
            BaseStudiesTree.Nodes.Add , _
                                    , _
                                    mDefaultConfiguration.UnderlyingStudy.Id, _
                                    mDefaultConfiguration.UnderlyingStudy.InstanceName
            BaseStudiesTree.Nodes(1).selected = True
            BaseStudiesTree.Enabled = False
        Else
            BaseStudiesTree.Nodes(mDefaultConfiguration.UnderlyingStudy.Id).selected = True
        End If
    End If
    
End If

Set studyInputDefinitions = mStudyDefinition.studyInputDefinitions
If Not mDefaultConfiguration Is Nothing Then inputValueNames = mDefaultConfiguration.inputValueNames

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
    If Not mDefaultConfiguration Is Nothing Then setComboSelection InputValueCombo(i - 1), _
                                                                    inputValueNames(i - 1)
    
    InputValueCombo(i - 1).Enabled = Not noInputModification
    
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
    
    permittedParamValues = studyParam.PermittedValues
    
    numPermittedParamValues = -1
    On Error Resume Next
    numPermittedParamValues = UBound(permittedParamValues)
    On Error GoTo Err
    If numPermittedParamValues <> -1 Then
        For Each permittedParamValue In permittedParamValues
            ParameterValueCombo(i - 1).ComboItems.Add , , permittedParamValue
        Next
        ParameterValueCombo(i - 1).Visible = True
        ParameterValueCombo(i - 1).Enabled = (Not noParameterModification)
    ElseIf studyParam.ParameterType = StudyParameterTypes.ParameterTypeInteger Then
        ParameterValueUpDown(i - 1).Min = 1
        ParameterValueUpDown(i - 1).Max = 255
        ParameterValueUpDown(i - 1).Visible = True
        ParameterValueUpDown(i - 1).Enabled = (Not noParameterModification)
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
        ParameterValueText(i - 1).Enabled = (Not noParameterModification)
    ElseIf studyParam.ParameterType = StudyParameterTypes.ParameterTypeBoolean Then
        ParameterValueCheck(i - 1).Visible = True
        ParameterValueCheck(i - 1).Enabled = (Not noParameterModification)
    Else
        If i = 1 Then
            ParameterValueUpDown(0).Visible = False
            ParameterValueText(0).Width = ParameterValueTemplateText.Width
        End If
        ParameterValueText(i - 1).Visible = True
        ParameterValueText(i - 1).Enabled = (Not noParameterModification)
    End If
    
    ParameterNameLabel(i - 1).Caption = studyParam.name
    defaultParamValue = defaultParams.GetParameterValue(studyParam.name)
    If studyParam.ParameterType = StudyParameterTypes.ParameterTypeBoolean Then
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
    
    If studyParam.ParameterType = StudyParameterTypes.ParameterTypeInteger Or _
        studyParam.ParameterType = StudyParameterTypes.ParameterTypeReal _
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
ReDim mStudyValueRegionNames(studyValueDefinitions.Count - 1) As String
j = 1
For i = 1 To studyValueDefinitions.Count
    Set studyValueDef = studyValueDefinitions.item(i)
    If Not studyValueconfigs Is Nothing Then
        Set studyValueconfig = Nothing
                
        On Error Resume Next
        Set studyValueconfig = studyValueconfigs.item(j)
        On Error GoTo Err
        If Not studyValueconfig Is Nothing Then
            If studyValueconfig.valueName = studyValueDef.name Then
                mStudyValueRegionNames(i - 1) = studyValueconfig.ChartRegionName
                j = j + 1
            Else
                Set studyValueconfig = Nothing
            End If
        End If
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

    If Not studyValueconfig Is Nothing Then
        IncludeCheck(i - 1) = IIf(studyValueconfig.IncludeInChart, vbChecked, vbUnchecked)
    Else
        IncludeCheck(i - 1) = IIf(studyValueDef.IncludeInChart, vbChecked, vbUnchecked)
    End If
        
    Select Case studyValueDef.ValueMode
    Case ValueModeNone
        Dim dpStyle As DataPointStyle
        
        ColorLabel(i - 1).ToolTipText = "Select the color for all values"
        
        UpColorLabel(i - 1).Visible = True
        UpColorLabel(i - 1).ToolTipText = "Optionally, select the color for higher values"
        
        DownColorLabel(i - 1).Visible = True
        DownColorLabel(i - 1).ToolTipText = "Optionally, select the color for lower values"
        
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            Set dpStyle = studyValueconfig.DataPointStyle
        ElseIf Not studyValueDef.ValueStyle Is Nothing Then
            Set dpStyle = studyValueDef.ValueStyle
        Else
            Set dpStyle = New DataPointStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(dpStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = dpStyle.Color
        UpColorLabel(i - 1).BackColor = IIf(dpStyle.UpColor = -1, NullColor, dpStyle.UpColor)
        DownColorLabel(i - 1).BackColor = IIf(dpStyle.DownColor = -1, NullColor, dpStyle.DownColor)
        
        initialisePointDisplayModeCombo DisplayModeCombo(i - 1), dpStyle.DisplayMode
        Select Case dpStyle.DisplayMode
        Case DataPointDisplayModes.DataPointDisplayModeLine
            initialiseLineStyleCombo StyleCombo(i - 1), dpStyle.LineStyle
        Case DataPointDisplayModes.DataPointDisplayModePoint
            initialisePointStyleCombo StyleCombo(i - 1), dpStyle.PointStyle
        Case DataPointDisplayModes.DataPointDisplayModeStep
            initialiseLineStyleCombo StyleCombo(i - 1), dpStyle.LineStyle
        Case DataPointDisplayModes.DataPointDisplayModeHistogram
            initialiseHistogramStyleCombo StyleCombo(i - 1), dpStyle.HistogramBarWidth
        End Select
        
        ThicknessText(i - 1).text = dpStyle.LineThickness
        
    Case ValueModeLine
        Dim lnStyle As LineStyle
        
        ColorLabel(i - 1).ToolTipText = "Select the color for the line"
        
        UpColorLabel(i - 1).Visible = True
        UpColorLabel(i - 1).ToolTipText = "Optionally, select the color for the start arrowhead"
        
        DownColorLabel(i - 1).Visible = True
        DownColorLabel(i - 1).ToolTipText = "Optionally, select the color for the end arrowhead"
        
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            Set lnStyle = studyValueconfig.LineStyle
        Else
            Set lnStyle = New LineStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(lnStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = lnStyle.Color
        UpColorLabel(i - 1).BackColor = IIf(lnStyle.ArrowStartFillColor = -1, NullColor, lnStyle.ArrowStartFillColor)
        DownColorLabel(i - 1).BackColor = IIf(lnStyle.ArrowEndFillColor = -1, NullColor, lnStyle.ArrowEndFillColor)
        
        initialiseLineDisplayModeCombo DisplayModeCombo(i - 1), _
                                        (lnStyle.ArrowStartStyle <> ArrowNone), _
                                        (lnStyle.ArrowEndStyle <> ArrowNone)
    
        initialiseLineStyleCombo StyleCombo(i - 1), lnStyle.LineStyle
        
        ThicknessText(i - 1).text = lnStyle.Thickness
        
    Case ValueModeBar
        Dim brStyle As BarStyle
        
        ColorLabel(i - 1).ToolTipText = "Optionally, select the color for the bar or the candlestick frame"
        
        UpColorLabel(i - 1).Visible = True
        UpColorLabel(i - 1).ToolTipText = "Select the color for up bars"
        
        DownColorLabel(i - 1).Visible = True
        DownColorLabel(i - 1).ToolTipText = "Optionally, select the color for down bars"
        
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            Set brStyle = studyValueconfig.BarStyle
        Else
            Set brStyle = New BarStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(brStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = IIf(brStyle.Color = -1, NullColor, brStyle.Color)
        UpColorLabel(i - 1).BackColor = IIf(brStyle.UpColor = -1, NullColor, brStyle.UpColor)
        DownColorLabel(i - 1).BackColor = IIf(brStyle.DownColor = -1, NullColor, brStyle.DownColor)
        
        initialiseBarDisplayModeCombo DisplayModeCombo(i - 1), _
                                        brStyle.DisplayMode, _
                                        brStyle.SolidUpBody
        
        initialiseBarStyleCombo StyleCombo(i - 1), brStyle.Width
        
        Select Case DisplayModeCombo(i - 1).SelectedItem.text
        Case BarModeBar
            ThicknessText(i - 1).text = brStyle.Thickness
        Case BarModeCandle
            ThicknessText(i - 1).text = brStyle.TailThickness
        Case BarModeSolidCandle
            ThicknessText(i - 1).text = brStyle.TailThickness
        Case BarModeLine
            ThicknessText(i - 1).text = brStyle.Thickness
        End Select
        
    Case ValueModeText
        Dim txStyle As TextStyle
        
        ColorLabel(i - 1).ToolTipText = "Select the color for the text"
        
        UpColorLabel(i - 1).Visible = True      ' box fill color
        UpColorLabel(i - 1).ToolTipText = "Optionally, select the color for the box fill"
        
        DownColorLabel(i - 1).Visible = True    ' box outline color
        UpColorLabel(i - 1).ToolTipText = "Optionally, select the color for the box outline"
        
        DisplayModeCombo(i - 1).Visible = True
        StyleCombo(i - 1).Visible = False
        FontButton(i - 1).Visible = True
        
        If Not studyValueconfig Is Nothing Then
            Set txStyle = studyValueconfig.TextStyle
        Else
            Set txStyle = New TextStyle
        End If
        
        AutoscaleCheck(i - 1) = IIf(txStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
        ColorLabel(i - 1).BackColor = txStyle.Color
        UpColorLabel(i - 1).BackColor = IIf(txStyle.BoxFillColor = -1, NullColor, txStyle.BoxFillColor)
        DownColorLabel(i - 1).BackColor = IIf(txStyle.BoxColor = -1, NullColor, txStyle.BoxColor)
        
        initialiseTextDisplayModeCombo DisplayModeCombo(i - 1), _
                                        txStyle.Box, _
                                        txStyle.BoxThickness, _
                                        txStyle.BoxStyle, _
                                        txStyle.BoxColor, _
                                        txStyle.BoxFillStyle, _
                                        txStyle.BoxFillColor
    
        ThicknessText(i - 1).text = txStyle.BoxThickness
        
        Set mFonts(i - 1) = txStyle.Font
        
    End Select

Next

If Not studyHorizRules Is Nothing Then
    For i = 1 To studyHorizRules.Count
        Set studyHorizRule = studyHorizRules.item(i)
        LineText(i - 1) = studyHorizRule.Y
        LineColorLabel(i - 1).BackColor = studyHorizRule.Color
    Next
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setComboSelection( _
                ByVal combo As ImageCombo, _
                ByVal text As String)
Dim item As ComboItem
Const ProcName As String = "setComboSelection"
On Error GoTo Err

For Each item In combo.ComboItems
    If UCase$(item.text) = UCase$(text) Then
        item.selected = True
        Exit For
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupBaseStudiesTree()
Dim studyConfig As StudyConfiguration

Const ProcName As String = "setupBaseStudiesTree"
On Error GoTo Err

BaseStudiesTree.Nodes.Clear
Set mCompatibleStudies = New Collection
Set mFirstCompatibleNode = Nothing

If mBaseStudyConfig Is Nothing Then Exit Sub

Set studyConfig = mBaseStudyConfig
addBaseStudiesTreeEntry studyConfig, Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function studiesCompatible( _
                ByVal sourceStudyDefinition As StudyDefinition, _
                ByVal sinkStudyDefinition As StudyDefinition) As Boolean
Dim sourceValueDef As StudyValueDefinition
Dim sinkInputDef As StudyInputDefinition
Dim i As Long

Const ProcName As String = "studiesCompatible"
On Error GoTo Err

For i = 1 To sinkStudyDefinition.studyInputDefinitions.Count
    Set sinkInputDef = sinkStudyDefinition.studyInputDefinitions.item(i)
    For Each sourceValueDef In sourceStudyDefinition.studyValueDefinitions
        If typesCompatible(sourceValueDef.ValueType, sinkInputDef.InputType) Then
            studiesCompatible = True
            Exit For
        End If
    Next
    If Not studiesCompatible Then Exit Function
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function typesCompatible( _
                ByVal sourceValueType As StudyValueTypes, _
                ByVal sinkInputType As StudyInputTypes) As Boolean
Const ProcName As String = "typesCompatible"
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function



