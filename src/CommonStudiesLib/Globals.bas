Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                        As String = "CmnStudiesLib27"
Private Const ModuleName                        As String = "Globals"


Public Const MaxDouble                          As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble                          As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh                          As Double = MinDouble
Public Const DummyLow                           As Double = MaxDouble

Public Const DefaultStudyValueName              As String = "$default"

' study name constants

Public Const AccDistName                        As String = "Accumulation/Distribution"
Public Const AccDistShortName                   As String = "AccDist"

Public Const AtrName                            As String = "Average True Range"
Public Const AtrShortName                       As String = "ATR"

Public Const BbName                             As String = "Bollinger Bands"
Public Const BbShortName                        As String = "BB"

Public Const DoncName                           As String = "Donchian Channels"
Public Const DoncShortName                      As String = "Donc"

Public Const EmaName                            As String = "Exponential Moving Average"
Public Const EmaShortName                       As String = "EMA"

Public Const FIName                             As String = "Force Index"
Public Const FIShortName                        As String = "FI"

Public Const MacdName                           As String = "MACD"
Public Const MacdShortName                      As String = "MACD"

Public Const PsName                             As String = "Parabolic Stop"
Public Const PsShortName                        As String = "PS"

Public Const RsiName                            As String = "Relative Strength Index"
Public Const RsiShortName                       As String = "RSI"

Public Const SdName                             As String = "Standard Deviation"
Public Const SdShortName                        As String = "SD"

Public Const SStochName                         As String = "Slow Stochastic"
Public Const SStochShortName                    As String = "SStoch"

Public Const SmaName                            As String = "Simple Moving Average"
Public Const SmaShortName                       As String = "SMA"

Public Const StochName                          As String = "Stochastic"
Public Const StochShortName                     As String = "Stoch"

Public Const SwingName                          As String = "Swing"
Public Const SwingShortName                     As String = "Swing"

' generic study parameter names - these are parameter names that are common
' to many studies
Public Const ParamMovingAverageType             As String = "Mov avg type"
Public Const ParamPeriods                       As String = "Periods"

'@================================================================================
' Enums
'@================================================================================

Public Enum MyLayerNumbers
    LayerBars = LayerNumbers.LayerLowestUser + 25
    LayerDataPoints = LayerNumbers.LayerLowestUser + 50
    LayerLines = LayerNumbers.LayerLowestUser + 75
    LayerTexts = LayerNumbers.LayerLowestUser + 100
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateBarStudyDefinition( _
                ByVal pName As String, _
                ByVal pShortName As String, _
                ByVal pDescription As String, _
                ByVal pInputValueName As String, _
                Optional ByVal pInputTotalVolumeName As String, _
                Optional ByVal pInputTickVolumeName As String, _
                Optional ByVal pInputOpenInterestName As String, _
                Optional ByVal pInputBarNumberName As String) As StudyDefinition
Const ProcName As String = "gCreateBarStudyDefinition"
On Error GoTo Err

Dim lStudyDefinition As New StudyDefinition
lStudyDefinition.name = pName
lStudyDefinition.NeedsBars = False
lStudyDefinition.ShortName = pShortName
lStudyDefinition.Description = pDescription
lStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom

Dim inputDef As StudyInputDefinition
Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputValueName)
inputDef.InputType = InputTypeReal
inputDef.Description = "Value"

If pInputTotalVolumeName <> "" Then
    Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputTotalVolumeName)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
End If

If pInputTickVolumeName <> "" Then
    Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputTickVolumeName)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Tick volume"
End If
    
If pInputOpenInterestName <> "" Then
    Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputOpenInterestName)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Open interest"
End If

If pInputBarNumberName <> "" Then
    Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputBarNumberName)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Bar number"
End If

Dim valueDef As StudyValueDefinition
Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueBar)
valueDef.Description = "The user-defined bars"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.IncludeInChart = True
valueDef.ValueMode = ValueModeBar
valueDef.ValueStyle = gCreateBarStyle
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOpen)
valueDef.Description = "Bar open value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&H8000&)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHigh)
valueDef.Description = "Bar high value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(vbBlue, Layer:=LayerBars + 1)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueLow)
valueDef.Description = "Bar low value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(vbRed, Layer:=LayerBars + 1)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueClose)
valueDef.Description = "Bar close value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.IsDefault = True
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&H80&, Layer:=LayerBars + 1)
valueDef.ValueType = ValueTypeReal

If pInputTotalVolumeName <> "" Then
    Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueVolume)
    valueDef.Description = "Bar volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H80000001, DisplayMode:=DataPointDisplayModeHistogram, DownColor:=&H4040C0, Layer:=LayerDataPoints, UpColor:=&H40C040)
    valueDef.ValueType = ValueTypeInteger
End If

If pInputTickVolumeName <> "" Then
    Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueTickVolume)
    valueDef.Description = "Bar tick volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H800000, DisplayMode:=DataPointDisplayModeHistogram, Layer:=LayerDataPoints)
    valueDef.ValueType = ValueTypeInteger
End If

If pInputOpenInterestName <> "" Then
    Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOpenInterest)
    valueDef.Description = "Bar open interest"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H80&, DisplayMode:=DataPointDisplayModeHistogram, Layer:=LayerDataPoints)
    valueDef.ValueType = ValueTypeInteger
End If

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHL2)
valueDef.Description = "Bar H+L/2 value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&HFF&, Layer:=LayerBars + 2)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHLC3)
valueDef.Description = "Bar H+L+C/3 value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&HFF00&, Layer:=LayerBars + 2)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOHLC4)
valueDef.Description = "Bar O+H+L+C/4 value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&HFF0000, Layer:=LayerBars + 2)
valueDef.ValueType = ValueTypeReal

Set gCreateBarStudyDefinition = lStudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateBarStyle( _
                Optional ByVal Color As Long = -1, _
                Optional ByVal DisplayMode As BarDisplayModes = BarDisplayModes.BarDisplayModeCandlestick, _
                Optional ByVal DownColor As Long = &H7878D1, _
                Optional ByVal IncludeInAutoscale As Boolean = True, _
                Optional ByVal Layer As Long = LayerBars, _
                Optional ByVal OutlineThickness As Long = 1, _
                Optional ByVal SolidUpBody As Boolean = True, _
                Optional ByVal TailThickness As Long = 1, _
                Optional ByVal Thickness As Long = 2, _
                Optional ByVal UpColor As Long = &H9BDD9B, _
                Optional ByVal Width As Single = 0.6) As BarStyle
Dim lStyle As BarStyle
Set lStyle = New BarStyle
lStyle.Color = Color
lStyle.DisplayMode = DisplayMode
lStyle.DownColor = DownColor
lStyle.IncludeInAutoscale = IncludeInAutoscale
lStyle.Layer = Layer
lStyle.OutlineThickness = OutlineThickness
lStyle.SolidUpBody = SolidUpBody
lStyle.TailThickness = TailThickness
lStyle.Thickness = Thickness
lStyle.UpColor = UpColor
lStyle.Width = Width
Set gCreateBarStyle = lStyle
End Function

Public Function gCreateDataPointStyle( _
                Optional ByVal Color As Long = vbBlack, _
                Optional ByVal DisplayMode As DataPointDisplayModes = DataPointDisplayModeLine, _
                Optional ByVal DownColor As Long = -1, _
                Optional ByVal HistogramBarWidth As Single = 0.6, _
                Optional ByVal IncludeInAutoscale As Boolean = True, _
                Optional ByVal Layer As Long = LayerDataPoints, _
                Optional ByVal LineStyle As LineStyles = LineSolid, _
                Optional ByVal Linethickness As Long = 1, _
                Optional ByVal PointStyle As PointStyles = PointRound, _
                Optional ByVal UpColor As Long = -1) As DataPointStyle
Dim style As DataPointStyle
Const ProcName As String = "gCreateDataPointStyle"
On Error GoTo Err

Set style = New DataPointStyle
style.Color = Color
style.DisplayMode = DisplayMode
style.DownColor = DownColor
style.HistogramBarWidth = HistogramBarWidth
style.IncludeInAutoscale = IncludeInAutoscale
style.Layer = Layer
style.LineStyle = LineStyle
style.Linethickness = Linethickness
style.PointStyle = PointStyle
style.UpColor = UpColor
Set gCreateDataPointStyle = style

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Property Get gCreateLineStyle( _
                Optional ByVal ArrowEndColor As Long = vbBlack, _
                Optional ByVal ArrowEndFillColor As Long = vbBlack, _
                Optional ByVal ArrowEndFillStyle As FillStyles = FillStyles.FillSolid, _
                Optional ByVal ArrowEndLength As Long = 10, _
                Optional ByVal ArrowEndStyle As ArrowStyles = ArrowStyles.ArrowNone, _
                Optional ByVal ArrowEndWidth As Long = 10, _
                Optional ByVal ArrowStartColor As Long = vbBlack, _
                Optional ByVal ArrowStartFillColor As Long = vbBlack, _
                Optional ByVal ArrowStartFillStyle As FillStyles = FillStyles.FillSolid, _
                Optional ByVal ArrowStartLength As Long = 10, _
                Optional ByVal ArrowStartStyle As ArrowStyles = ArrowStyles.ArrowNone, _
                Optional ByVal ArrowStartWidth As Long = 10, _
                Optional ByVal Color As Long = vbBlack, _
                Optional ByVal ExtendAfter As Boolean = False, _
                Optional ByVal ExtendBefore As Boolean = False, _
                Optional ByVal Extended As Boolean = False, _
                Optional ByVal FixedX As Boolean = False, _
                Optional ByVal FixedY As Boolean = False, _
                Optional ByVal IncludeInAutoscale As Boolean = False, _
                Optional ByVal Layer As Long = LayerLines, _
                Optional ByVal LineStyle As LineStyles = LineStyles.LineSolid, _
                Optional ByVal Thickness As Long = 1) As LineStyle
Dim lStyle As LineStyle
Set lStyle = New LineStyle
lStyle.ArrowEndColor = ArrowEndColor
lStyle.ArrowEndFillColor = ArrowEndFillColor
lStyle.ArrowEndFillStyle = ArrowEndFillStyle
lStyle.ArrowEndLength = ArrowEndLength
lStyle.ArrowEndStyle = ArrowEndStyle
lStyle.ArrowEndWidth = ArrowEndWidth
lStyle.ArrowStartColor = ArrowStartColor
lStyle.ArrowStartFillColor = ArrowStartFillColor
lStyle.ArrowStartFillStyle = ArrowStartFillStyle
lStyle.ArrowStartLength = ArrowStartLength
lStyle.ArrowStartStyle = ArrowStartStyle
lStyle.ArrowStartWidth = ArrowStartWidth
lStyle.Color = Color
lStyle.ExtendAfter = ExtendAfter
lStyle.ExtendBefore = ExtendBefore
lStyle.Extended = Extended
lStyle.FixedX = FixedX
lStyle.FixedY = FixedY
lStyle.IncludeInAutoscale = IncludeInAutoscale
lStyle.Layer = Layer
lStyle.LineStyle = LineStyle
lStyle.Thickness = Thickness
Set gCreateLineStyle = lStyle
End Property

Public Function gCreateMA( _
                ByVal StudyManager As StudyManager, _
                ByVal maType As String, _
                ByVal periods As Long, _
                ByVal numberOfValuesToCache As Long) As IStudy
Const ProcName As String = "gCreateMA"
On Error GoTo Err

Dim valueNames(0) As String
valueNames(0) = "in"

Dim lparams As Parameters
Dim lStudy As IStudy
Dim lSf As New StudyFoundation

Select Case UCase$(maType)
Case UCase$(EmaShortName)
    Dim lEMA As EMA
    Set lEMA = New EMA
    Set lStudy = lEMA
    Set lparams = GEMA.defaultParameters
    lparams.SetParameterValue ParamPeriods, periods
Case Else
    Dim lSMA As SMA
    Set lSMA = New SMA
    Set lStudy = lSMA
    Set lparams = GSMA.defaultParameters
    lparams.SetParameterValue ParamPeriods, periods
End Select

lSf.Initialise "", _
            "", _
            StudyManager, _
            lStudy, _
            GenerateGUIDString, _
            lparams, _
            numberOfValuesToCache, _
            valueNames, _
            Nothing, _
            Nothing
lStudy.Initialise lSf
                
Set gCreateMA = lStudy

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateTextStyle(Optional ByVal Angle = 0, _
                Optional ByVal Align As TextAlignModes = TextAlignModes.AlignTopLeft, _
                Optional ByVal Box As Boolean = False, _
                Optional ByVal BoxColor As Long = vbBlack, _
                Optional ByVal BoxStyle As LineStyles = LineStyles.LineSolid, _
                Optional ByVal BoxThickness As Long = 1, _
                Optional ByVal BoxFillColor As Long = vbWhite, _
                Optional ByVal BoxFillStyle As FillStyles = FillStyles.FillSolid, _
                Optional ByVal BoxFillWithBackgroundColor As Boolean = False, _
                Optional ByVal Color As Long = vbBlack, _
                Optional ByVal Font As StdFont, _
                Optional ByVal Ellipsis As EllipsisModes = EllipsisModes.EllipsisNone, _
                Optional ByVal ExpandTabs As Boolean = True, _
                Optional ByVal Extended As Boolean = False, _
                Optional ByVal FixedX As Boolean = False, _
                Optional ByVal FixedY As Boolean = False, _
                Optional ByVal HideIfBlank As Boolean = True, _
                Optional ByVal IncludeInAutoscale As Boolean = False, _
                Optional ByVal Justification As TextJustifyModes = TextJustifyModes.JustifyLeft, _
                Optional ByVal Layer As Long = LayerTexts, _
                Optional ByVal MultiLine As Boolean = False, _
                Optional ByVal PaddingX As Long = 1, _
                Optional ByVal PaddingY As Long = 0, _
                Optional ByVal TabWidth As Long = 8, _
                Optional ByVal WordWrap As Boolean = True) As TextStyle
Dim lStyle As TextStyle
Set lStyle = New TextStyle

lStyle.Angle = 0
lStyle.Align = Align
lStyle.Box = Box
lStyle.BoxColor = BoxColor
lStyle.BoxStyle = BoxStyle
lStyle.BoxThickness = BoxThickness
lStyle.BoxFillColor = BoxFillColor
lStyle.BoxFillStyle = BoxFillStyle
lStyle.BoxFillWithBackgroundColor = BoxFillWithBackgroundColor
lStyle.Color = Color
If Not Font Is Nothing Then
    lStyle.Font = Font
Else
    Dim aFont As New StdFont
    aFont.Bold = False
    aFont.Italic = False
    aFont.name = "Arial"
    aFont.Size = 8
    aFont.Strikethrough = False
    aFont.Underline = False
    lStyle.Font = aFont
End If
lStyle.Ellipsis = Ellipsis
lStyle.ExpandTabs = ExpandTabs
lStyle.Extended = Extended
lStyle.FixedX = FixedX
lStyle.FixedY = FixedY
lStyle.HideIfBlank = HideIfBlank
lStyle.IncludeInAutoscale = IncludeInAutoscale
lStyle.Justification = Justification
lStyle.Layer = Layer
lStyle.MultiLine = MultiLine
lStyle.PaddingX = PaddingX
lStyle.PaddingY = PaddingY
lStyle.TabWidth = TabWidth
lStyle.WordWrap = WordWrap
Set gCreateTextStyle = lStyle
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gMaTypes() As Variant()
Dim ar(1) As Variant
Const ProcName As String = "gMaTypes"
On Error GoTo Err

ar(0) = EmaShortName
ar(1) = SmaShortName
gMaTypes = ar

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Function
'@================================================================================





