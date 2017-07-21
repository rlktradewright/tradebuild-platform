VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.UserControl StudyValueConfigurer 
   BackStyle       =   0  'Transparent
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ScaleHeight     =   375
   ScaleWidth      =   6885
   Begin TWControls40.TWButton AdvancedButton 
      Height          =   300
      Left            =   6300
      TabIndex        =   6
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      Caption         =   "..."
      DefaultBorderColor=   15793920
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
   Begin TWControls40.TWButton FontButton 
      Height          =   300
      Left            =   5250
      TabIndex        =   11
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      Caption         =   "Font..."
      DefaultBorderColor=   15793920
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
   Begin TWControls40.TWImageCombo StyleCombo 
      Height          =   270
      Left            =   5160
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      MouseIcon       =   "StudyValueConfigurer.ctx":0000
      Text            =   ""
   End
   Begin TWControls40.TWImageCombo DisplayModeCombo 
      Height          =   270
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
      MouseIcon       =   "StudyValueConfigurer.ctx":001C
      Text            =   ""
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox IncludeCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Set to include this study value in the chart"
      Top             =   0
      Width           =   195
   End
   Begin VB.TextBox ThicknessText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4320
      TabIndex        =   3
      Text            =   "1"
      ToolTipText     =   "Choose the thickness of lines or points"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox AutoscaleCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Set this to ensure that all values are visible when the chart is auto-scaling"
      Top             =   0
      Width           =   210
   End
   Begin MSComCtl2.UpDown ThicknessUpDown 
      Height          =   300
      Left            =   4800
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      OrigLeft        =   4080
      OrigTop         =   240
      OrigRight       =   4335
      OrigBottom      =   570
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.Label ValueNameLabel 
      Caption         =   "Label2"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2175
      TabIndex        =   9
      ToolTipText     =   "Click to change the colour for this value"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label UpColorLabel 
      BackColor       =   &H0000FF00&
      Height          =   300
      Left            =   2520
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.Label DownColorLabel 
      BackColor       =   &H000000FF&
      Height          =   300
      Left            =   2865
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "StudyValueConfigurer"
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

Private Const ModuleName                            As String = "StudyValueConfigurer"

'@================================================================================
' Member variables
'@================================================================================

Private mStudyValueDef As StudyValueDefinition
Private mStudyValueConfig As StudyValueConfiguration

Private mFont As StdFont

Private mTheme                              As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Height = FontButton.Height + 75
UserControl.Width = AdvancedButton.Left + AdvancedButton.Width
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AdvancedButton_Click()
Const ProcName As String = "AdvancedButton_Click"
On Error GoTo Err

gNotImplemented

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ColorLabel_Click()
Const ProcName As String = "ColorLabel_Click"
On Error GoTo Err

ColorLabel.BackColor = gChooseAColor(ColorLabel.BackColor, _
                                    IIf(mStudyValueDef.ValueMode = ValueModeBar, True, False), _
                                    gGetParentForm(Me))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DisplayModeCombo_Click()
Const ProcName As String = "DisplayModeCombo_Click"
On Error GoTo Err

Select Case mStudyValueDef.ValueMode
Case ValueModeNone
    Dim dpStyle As DataPointStyle
    
    If Not mStudyValueConfig Is Nothing Then
        Set dpStyle = mStudyValueConfig.DataPointStyle
    Else
        Set dpStyle = GetDefaultDataPointStyle.Clone
    End If
        
    Select Case DisplayModeCombo.SelectedItem.Text
    Case PointDisplayModeLine
        initialiseLineStyleCombo StyleCombo, dpStyle.LineStyle
    Case PointDisplayModePoint
        initialisePointStyleCombo StyleCombo, dpStyle.PointStyle
    Case PointDisplayModeSteppedLine
        initialiseLineStyleCombo StyleCombo, dpStyle.LineStyle
    Case PointDisplayModeHistogram
        initialiseHistogramStyleCombo StyleCombo, dpStyle.HistogramBarWidth
    End Select
Case ValueModeLine

Case ValueModeBar

Case ValueModeText

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DisplayModeCombo_Validate(Cancel As Boolean)
Const ProcName As String = "DisplayModeCombo_Validate"
On Error GoTo Err

If DisplayModeCombo.SelectedItem Is Nothing Then Cancel = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DownColorLabel_Click()
Const ProcName As String = "ColorLabel_Click"
On Error GoTo Err

DownColorLabel.BackColor = gChooseAColor(DownColorLabel.BackColor, _
                                        IIf(mStudyValueDef.ValueMode = ValueModeBar, True, False), _
                                        gGetParentForm(Me))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FontButton_Click()
Const ProcName As String = "FontButton_Click"
On Error GoTo Err

CommonDialog1.flags = cdlCFBoth + cdlCFEffects
CommonDialog1.FontName = mFont.name
CommonDialog1.FontBold = mFont.Bold
CommonDialog1.FontItalic = mFont.Italic
CommonDialog1.FontSize = mFont.Size
CommonDialog1.FontStrikethru = mFont.Strikethrough
CommonDialog1.FontUnderline = mFont.Underline
CommonDialog1.Color = ColorLabel.BackColor
CommonDialog1.ShowFont

Dim aFont As New StdFont
aFont.Bold = CommonDialog1.FontBold
aFont.Italic = CommonDialog1.FontItalic
aFont.name = CommonDialog1.FontName
aFont.Size = CommonDialog1.FontSize
aFont.Strikethrough = CommonDialog1.FontStrikethru
aFont.Underline = CommonDialog1.FontUnderline

Set mFont = aFont

ColorLabel.BackColor = CommonDialog1.Color

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub StyleCombo_Validate(Cancel As Boolean)
Const ProcName As String = "StyleCombo_Validate"
On Error GoTo Err

If StyleCombo.SelectedItem Is Nothing Then Cancel = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ThicknessText_KeyPress(KeyAscii As Integer)
Const ProcName As String = "ThicknessText_KeyPress"
On Error GoTo Err

gFilterNonNumericKeyPress KeyAscii

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ThicknessText_Validate(Cancel As Boolean)
If ThicknessText.Text = "" Or ThicknessText.Text = "0" Then
    ThicknessText.Text = "1"
End If
End Sub

Private Sub UpColorLabel_Click()
Const ProcName As String = "ColorLabel_Click"
On Error GoTo Err

UpColorLabel.BackColor = gChooseAColor(UpColorLabel.BackColor, _
                                        IIf(mStudyValueDef.ValueMode = ValueModeBar, True, False), _
                                        gGetParentForm(Me))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor

Dim lColor As Long: lColor = ColorLabel.BackColor
Dim lUpColor As Long: lUpColor = UpColorLabel.BackColor
Dim lDownColor As Long: lDownColor = DownColorLabel.BackColor

gApplyTheme mTheme, UserControl.Controls

ColorLabel.BackColor = lColor
UpColorLabel.BackColor = lUpColor
DownColorLabel.BackColor = lDownColor

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

Public Sub ApplyUpdates(ByVal pStudyValueConfig As StudyValueConfiguration)
Const ProcName As String = "ApplyUpdates"
On Error GoTo Err

If Not mStudyValueConfig Is Nothing Then
    pStudyValueConfig.BarFormatterFactoryName = mStudyValueConfig.BarFormatterFactoryName
    pStudyValueConfig.BarFormatterLibraryName = mStudyValueConfig.BarFormatterLibraryName
End If

pStudyValueConfig.IncludeInChart = (IncludeCheck.value = vbChecked)

pStudyValueConfig.ChartRegionName = getRegionName

Select Case mStudyValueDef.ValueMode
Case ValueModeNone
    Dim dpStyle As DataPointStyle
    
    Set dpStyle = GetDefaultDataPointStyle.Clone
    
    dpStyle.IncludeInAutoscale = (AutoscaleCheck.value = vbChecked)
    dpStyle.Color = ColorLabel.BackColor
    dpStyle.DownColor = IIf(DownColorLabel.BackColor = NullColor, _
                        -1, _
                        DownColorLabel.BackColor)
    dpStyle.UpColor = IIf(UpColorLabel.BackColor = NullColor, _
                        -1, _
                        UpColorLabel.BackColor)
    
    Select Case DisplayModeCombo.SelectedItem.Text
    Case PointDisplayModeLine
        dpStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
        Select Case StyleCombo.SelectedItem.Text
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
        Select Case StyleCombo.SelectedItem.Text
        Case PointStyleRound
            dpStyle.PointStyle = PointRound
        Case PointStyleSquare
            dpStyle.PointStyle = PointSquare
        End Select
    Case PointDisplayModeSteppedLine
        dpStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep
        Select Case StyleCombo.SelectedItem.Text
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
        Select Case StyleCombo.SelectedItem.Text
        Case HistogramStyleNarrow
            dpStyle.HistogramBarWidth = HistogramWidthNarrow
        Case HistogramStyleMedium
            dpStyle.HistogramBarWidth = HistogramWidthMedium
        Case HistogramStyleWide
            dpStyle.HistogramBarWidth = HistogramWidthWide
        Case CustomStyle
            dpStyle.HistogramBarWidth = CSng(StyleCombo.SelectedItem.Tag)
        End Select
    End Select
    
    dpStyle.LineThickness = ThicknessText.Text
    
    pStudyValueConfig.DataPointStyle = dpStyle
Case ValueModeLine
    Dim lnStyle As LineStyle

    Set lnStyle = GetDefaultLineStyle.Clone
    
    lnStyle.IncludeInAutoscale = (AutoscaleCheck.value = vbChecked)
    lnStyle.Color = ColorLabel.BackColor
    lnStyle.ArrowStartColor = ColorLabel.BackColor
    lnStyle.ArrowEndColor = ColorLabel.BackColor
    lnStyle.ArrowStartFillColor = IIf(UpColorLabel.BackColor = NullColor, _
                                    -1, _
                                    UpColorLabel.BackColor)
    lnStyle.ArrowEndFillColor = IIf(DownColorLabel.BackColor = NullColor, _
                                    -1, _
                                    DownColorLabel.BackColor)
    
    Select Case DisplayModeCombo.SelectedItem.Text
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
        
    Select Case StyleCombo.SelectedItem.Text
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
    
    lnStyle.Thickness = ThicknessText.Text
    ' temporary fix until ChartSkil improves drawing of non-extended lines
    lnStyle.Extended = True
    
    pStudyValueConfig.LineStyle = lnStyle

Case ValueModeBar
    Dim brStyle As BarStyle
    
    Set brStyle = GetDefaultBarStyle.Clone
    
    brStyle.IncludeInAutoscale = (AutoscaleCheck.value = vbChecked)
    brStyle.Color = IIf(ColorLabel.BackColor = NullColor, _
                        -1, _
                        ColorLabel.BackColor)
    brStyle.DownColor = IIf(DownColorLabel.BackColor = NullColor, _
                        -1, _
                        DownColorLabel.BackColor)
    brStyle.UpColor = UpColorLabel.BackColor
    
    Select Case DisplayModeCombo.SelectedItem.Text
    Case BarModeBar
        brStyle.DisplayMode = BarDisplayModes.BarDisplayModeBar
        brStyle.Thickness = ThicknessText.Text
    Case BarModeCandle
        brStyle.DisplayMode = BarDisplayModes.BarDisplayModeCandlestick
        brStyle.SolidUpBody = False
        brStyle.TailThickness = ThicknessText.Text
    Case BarModeSolidCandle
        brStyle.DisplayMode = BarDisplayModes.BarDisplayModeCandlestick
        brStyle.SolidUpBody = True
        brStyle.TailThickness = ThicknessText.Text
    Case BarModeLine
        brStyle.DisplayMode = BarDisplayModes.BarDisplayModeLine
    End Select
    
    Select Case StyleCombo.SelectedItem.Text
    Case BarStyleNarrow
        brStyle.Width = BarWidthNarrow
    Case BarStyleMedium
        brStyle.Width = BarWidthMedium
    Case BarStyleWide
        brStyle.Width = BarWidthWide
    Case CustomStyle
        brStyle.Width = CSng(StyleCombo.SelectedItem.Tag)
    End Select
    
    pStudyValueConfig.BarStyle = brStyle

Case ValueModeText
    Dim txStyle As TextStyle

    Set txStyle = GetDefaultTextStyle.Clone
    
    txStyle.IncludeInAutoscale = (AutoscaleCheck.value = vbChecked)
    txStyle.Color = ColorLabel.BackColor
    txStyle.BoxFillColor = IIf(UpColorLabel.BackColor = NullColor, _
                                    -1, _
                                    UpColorLabel.BackColor)
    txStyle.BoxColor = IIf(DownColorLabel.BackColor = NullColor, _
                                    -1, _
                                    DownColorLabel.BackColor)
    
    Select Case DisplayModeCombo.SelectedItem.Text
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
        
    If TypeName(FontButton.Tag) <> "Nothing" Then
        txStyle.Font = mFont
    End If
    
    txStyle.BoxThickness = ThicknessText.Text
    ' temporary fix until ChartSkil improves drawing of non-extended texts
    txStyle.Extended = True
    
    pStudyValueConfig.TextStyle = txStyle


End Select
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pStudyValueDef As StudyValueDefinition, _
                ByVal pStudyValueConfig As StudyValueConfiguration)
Const ProcName As String = "Initialise"
On Error GoTo Err

DownColorLabel.Visible = False
DisplayModeCombo.Visible = False
FontButton.Visible = False
StyleCombo.Visible = False
ThicknessUpDown.Visible = False
UpColorLabel.Visible = False

Set mStudyValueDef = pStudyValueDef
Set mStudyValueConfig = pStudyValueConfig

AutoscaleCheck = vbUnchecked

ValueNameLabel.Caption = mStudyValueDef.name
ValueNameLabel.ToolTipText = mStudyValueDef.Description

If Not mStudyValueConfig Is Nothing Then
    IncludeCheck = IIf(mStudyValueConfig.IncludeInChart, vbChecked, vbUnchecked)
Else
    IncludeCheck = IIf(mStudyValueDef.IncludeInChart, vbChecked, vbUnchecked)
End If
    
Select Case mStudyValueDef.ValueMode
Case ValueModeNone
    Dim dpStyle As DataPointStyle
    
    ColorLabel.ToolTipText = "Select the color for all values"
    
    UpColorLabel.Visible = True
    UpColorLabel.ToolTipText = "Optionally, select the color for higher values"
    
    DownColorLabel.Visible = True
    DownColorLabel.ToolTipText = "Optionally, select the color for lower values"
    
    DisplayModeCombo.Visible = True
    StyleCombo.Visible = True
    
    If Not mStudyValueConfig Is Nothing Then
        Set dpStyle = mStudyValueConfig.DataPointStyle
    ElseIf Not mStudyValueDef.ValueStyle Is Nothing Then
        Set dpStyle = mStudyValueDef.ValueStyle
    End If
    If dpStyle Is Nothing Then Set dpStyle = GetDefaultDataPointStyle.Clone
    
    AutoscaleCheck = IIf(dpStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
    ColorLabel.BackColor = dpStyle.Color
    UpColorLabel.BackColor = IIf(dpStyle.UpColor = -1, NullColor, dpStyle.UpColor)
    DownColorLabel.BackColor = IIf(dpStyle.DownColor = -1, NullColor, dpStyle.DownColor)
    
    initialisePointDisplayModeCombo DisplayModeCombo, dpStyle.DisplayMode
    Select Case dpStyle.DisplayMode
    Case DataPointDisplayModes.DataPointDisplayModeLine
        initialiseLineStyleCombo StyleCombo, dpStyle.LineStyle
    Case DataPointDisplayModes.DataPointDisplayModePoint
        initialisePointStyleCombo StyleCombo, dpStyle.PointStyle
    Case DataPointDisplayModes.DataPointDisplayModeStep
        initialiseLineStyleCombo StyleCombo, dpStyle.LineStyle
    Case DataPointDisplayModes.DataPointDisplayModeHistogram
        initialiseHistogramStyleCombo StyleCombo, dpStyle.HistogramBarWidth
    End Select
    
    ThicknessText.Text = dpStyle.LineThickness
    ThicknessText.Visible = True
Case ValueModeLine
    Dim lnStyle As LineStyle
    
    ColorLabel.ToolTipText = "Select the color for the line"
    
    UpColorLabel.Visible = True
    UpColorLabel.ToolTipText = "Optionally, select the color for the start arrowhead"
    
    DownColorLabel.Visible = True
    DownColorLabel.ToolTipText = "Optionally, select the color for the end arrowhead"
    
    DisplayModeCombo.Visible = True
    StyleCombo.Visible = True
    
    If Not mStudyValueConfig Is Nothing Then
        Set lnStyle = mStudyValueConfig.LineStyle
    ElseIf Not mStudyValueDef.ValueStyle Is Nothing Then
        Set lnStyle = mStudyValueDef.ValueStyle
    End If
    If lnStyle Is Nothing Then Set lnStyle = GetDefaultLineStyle.Clone
    
    AutoscaleCheck = IIf(lnStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
    ColorLabel.BackColor = lnStyle.Color
    UpColorLabel.BackColor = IIf(lnStyle.ArrowStartFillColor = -1, NullColor, lnStyle.ArrowStartFillColor)
    DownColorLabel.BackColor = IIf(lnStyle.ArrowEndFillColor = -1, NullColor, lnStyle.ArrowEndFillColor)
    
    initialiseLineDisplayModeCombo DisplayModeCombo, _
                                    (lnStyle.ArrowStartStyle <> ArrowNone), _
                                    (lnStyle.ArrowEndStyle <> ArrowNone)

    initialiseLineStyleCombo StyleCombo, lnStyle.LineStyle
    
    ThicknessText.Text = lnStyle.Thickness
    ThicknessText.Visible = True
Case ValueModeBar
    Dim brStyle As BarStyle
    
    ColorLabel.ToolTipText = "Optionally, select the color for the bar or the candlestick frame"
    
    UpColorLabel.Visible = True
    UpColorLabel.ToolTipText = "Select the color for up bars"
    
    DownColorLabel.Visible = True
    DownColorLabel.ToolTipText = "Optionally, select the color for down bars"
    
    DisplayModeCombo.Visible = True
    StyleCombo.Visible = True
    
    If Not mStudyValueConfig Is Nothing Then
        Set brStyle = mStudyValueConfig.BarStyle
    ElseIf Not mStudyValueDef.ValueStyle Is Nothing Then
        Set brStyle = mStudyValueDef.ValueStyle
    End If
    If brStyle Is Nothing Then Set brStyle = GetDefaultBarStyle.Clone
    
    AutoscaleCheck = IIf(brStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
    ColorLabel.BackColor = IIf(brStyle.Color = -1, NullColor, brStyle.Color)
    UpColorLabel.BackColor = IIf(brStyle.UpColor = -1, NullColor, brStyle.UpColor)
    DownColorLabel.BackColor = IIf(brStyle.DownColor = -1, NullColor, brStyle.DownColor)
    
    initialiseBarDisplayModeCombo DisplayModeCombo, _
                                    brStyle.DisplayMode, _
                                    brStyle.SolidUpBody
    
    initialiseBarStyleCombo StyleCombo, brStyle.Width
    
    Select Case DisplayModeCombo.SelectedItem.Text
    Case BarModeBar
        ThicknessText.Text = brStyle.Thickness
    Case BarModeCandle
        ThicknessText.Text = brStyle.TailThickness
    Case BarModeSolidCandle
        ThicknessText.Text = brStyle.TailThickness
    Case BarModeLine
        ThicknessText.Text = brStyle.Thickness
    End Select
    ThicknessText.Visible = True
Case ValueModeText
    Dim txStyle As TextStyle
    
    ColorLabel.ToolTipText = "Select the color for the text"
    
    UpColorLabel.Visible = True      ' box fill color
    UpColorLabel.ToolTipText = "Optionally, select the color for the box fill"
    
    DownColorLabel.Visible = True    ' box outline color
    UpColorLabel.ToolTipText = "Optionally, select the color for the box outline"
    
    DisplayModeCombo.Visible = True
    StyleCombo.Visible = False
    FontButton.Visible = True
    
    If Not mStudyValueConfig Is Nothing Then
        Set txStyle = mStudyValueConfig.TextStyle
    ElseIf Not mStudyValueDef.ValueStyle Is Nothing Then
        Set txStyle = mStudyValueDef.ValueStyle
    End If
    If txStyle Is Nothing Then Set txStyle = GetDefaultTextStyle.Clone
    
    AutoscaleCheck = IIf(txStyle.IncludeInAutoscale, vbChecked, vbUnchecked)
    ColorLabel.BackColor = txStyle.Color
    UpColorLabel.BackColor = IIf(txStyle.BoxFillColor = -1, NullColor, txStyle.BoxFillColor)
    DownColorLabel.BackColor = IIf(txStyle.BoxColor = -1, NullColor, txStyle.BoxColor)
    
    initialiseTextDisplayModeCombo DisplayModeCombo, _
                                    txStyle.Box, _
                                    txStyle.BoxThickness, _
                                    txStyle.BoxStyle, _
                                    txStyle.BoxColor, _
                                    txStyle.BoxFillStyle, _
                                    txStyle.BoxFillColor

    ThicknessText.Text = txStyle.BoxThickness
    ThicknessText.Visible = True
    
    Set mFont = txStyle.Font
    
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getDefaultRegionName() As String
Select Case mStudyValueDef.DefaultRegion
Case StudyValueDefaultRegionNone
    getDefaultRegionName = ChartRegionNameDefault
Case StudyValueDefaultRegionDefault
    getDefaultRegionName = ChartRegionNameDefault
Case StudyValueDefaultRegionCustom
    getDefaultRegionName = ChartRegionNameCustom
Case StudyValueDefaultRegionUnderlying
    getDefaultRegionName = ChartRegionNameUnderlying
End Select
End Function

Private Function getRegionName() As String
If useCurrentRegionName Then
    getRegionName = mStudyValueConfig.ChartRegionName
Else
    getRegionName = getDefaultRegionName
End If
End Function

Private Sub initialiseBarDisplayModeCombo( _
                ByVal combo As TWImageCombo, _
                ByVal pDisplayMode As BarDisplayModes, _
                ByVal pSolid As Boolean)
Const ProcName As String = "initialiseBarDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseBarStyleCombo( _
                ByVal combo As TWImageCombo, _
                ByVal barWidth As Single)
Const ProcName As String = "initialiseBarStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
Set item = combo.ComboItems.Add(, , BarStyleMedium)

Dim selected As Boolean
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseHistogramStyleCombo( _
                ByVal combo As TWImageCombo, _
                ByVal histBarWidth As Single)
Const ProcName As String = "initialiseHistogramStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
Set item = combo.ComboItems.Add(, , HistogramStyleMedium)

Dim selected As Boolean
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseLineDisplayModeCombo( _
                ByVal combo As TWImageCombo, _
                ByVal pArrowStart As Boolean, _
                ByVal pArrowEnd As Boolean)
Const ProcName As String = "initialiseLineDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseLineStyleCombo( _
                ByVal combo As TWImageCombo, _
                ByVal pLineStyle As LineStyles)
Const ProcName As String = "initialiseLineStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialisePointDisplayModeCombo( _
                ByVal combo As TWImageCombo, _
                ByVal pDisplayMode As DataPointDisplayModes)
Const ProcName As String = "initialisePointDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialisePointStyleCombo( _
                ByVal combo As TWImageCombo, _
                ByVal pPointStyle As PointStyles)
Const ProcName As String = "initialisePointStyleCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
Set item = combo.ComboItems.Add(, , PointStyleRound)
If pPointStyle = PointRound Then item.selected = True

Set item = combo.ComboItems.Add(, , PointStyleSquare)
If pPointStyle = PointSquare Then item.selected = True

combo.ToolTipText = "Select the shape of the point"

combo.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseTextDisplayModeCombo( _
                ByVal combo As TWImageCombo, _
                ByVal pBox As Boolean, _
                ByVal pBoxThickness As Long, _
                ByVal pBoxStyle As LineStyles, _
                ByVal pBoxColor As Long, _
                ByVal pBoxFillStyle As FillStyles, _
                ByVal pBoxFillColor As Long)
Const ProcName As String = "initialiseTextDisplayModeCombo"
On Error GoTo Err

combo.ComboItems.Clear

Dim item As ComboItem
Set item = combo.ComboItems.Add(, , TextDisplayModePlain)

Dim selected As Boolean
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function useCurrentRegionName() As Boolean
If mStudyValueConfig Is Nothing Then Exit Function
If mStudyValueConfig.ChartRegionName <> "" Then useCurrentRegionName = True
End Function


