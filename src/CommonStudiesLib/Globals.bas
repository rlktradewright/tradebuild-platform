Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                    As String = "CmnStudiesLib26"
Private Const ModuleName                As String = "Globals"


Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh As Double = MinDouble
Public Const DummyLow As Double = MaxDouble

Public Const DefaultStudyValueName As String = "$default"

' study name constants

Public Const AccDistName As String = "Accumulation/Distribution"
Public Const AccDistShortName As String = "AccDist"

Public Const AtrName As String = "Average True Range"
Public Const AtrShortName As String = "ATR"

Public Const BbName As String = "Bollinger Bands"
Public Const BbShortName As String = "BB"

Public Const ConstMomentumBarsName As String = "Constant momentum bars"
Public Const ConstMomentumBarsShortName As String = "CM Bars"

Public Const ConstTimeBarsName As String = "Constant time bars"
Public Const ConstTimeBarsShortName As String = "Bars"

Public Const ConstVolBarsName As String = "Constant volume bars"
Public Const ConstVolBarsShortName As String = "CV Bars"

Public Const DoncName As String = "Donchian Channels"
Public Const DoncShortName As String = "Donc"

Public Const EmaName As String = "Exponential Moving Average"
Public Const EmaShortName As String = "EMA"

Public Const FIName As String = "Force Index"
Public Const FIShortName As String = "FI"

Public Const MacdName As String = "MACD"
Public Const MacdShortName As String = "MACD"

Public Const PsName As String = "Parabolic Stop"
Public Const PsShortName As String = "PS"

Public Const RsiName As String = "Relative Strength Index"
Public Const RsiShortName As String = "RSI"

Public Const SdName As String = "Standard Deviation"
Public Const SdShortName As String = "SD"

Public Const SStochName As String = "Slow Stochastic"
Public Const SStochShortName As String = "SStoch"

Public Const SmaName As String = "Simple Moving Average"
Public Const SmaShortName As String = "SMA"

Public Const StochName As String = "Stochastic"
Public Const StochShortName As String = "Stoch"

Public Const SwingName As String = "Swing"
Public Const SwingShortName As String = "Swing"

' generic study parameter names - these are parameter names that are common
' to many studies
Public Const ParamMovingAverageType As String = "Mov avg type"
Public Const ParamPeriods As String = "Periods"

' sub-value names for study values in bar mode
Public Const BarValueOpen As String = "Open"
Public Const BarValueHigh As String = "High"
Public Const BarValueLow As String = "Low"
Public Const BarValueClose As String = "Close"
Public Const BarValueVolume As String = "Volume"
Public Const BarValueTickVolume As String = "Tick Volume"
Public Const BarValueOpenInterest As String = "Open Interest"
Public Const BarValueHL2 As String = "(H+L)/2"
Public Const BarValueHLC3 As String = "(H+L+C)/3"
Public Const BarValueOHLC4 As String = "(O+H+L+C)/4"

'@================================================================================
' Enums
'@================================================================================

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

Public gLibraryManager As StudyLibraryManager

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

Public Function gCreateBarStyle( _
                Optional ByVal Color As Long = -1, _
                Optional ByVal DisplayMode As BarDisplayModes = BarDisplayModeCandlestick, _
                Optional ByVal DownColor As Long = vbBlack, _
                Optional ByVal IncludeInAutoscale As Boolean = True, _
                Optional ByVal Layer As Long = LayerLowestUser, _
                Optional ByVal OutlineThickness As Long = 1, _
                Optional ByVal SolidUpBody As Boolean = False, _
                Optional ByVal TailThickness As Long = 1, _
                Optional ByVal Thickness As Long = 2, _
                Optional ByVal UpColor As Long = vbBlack, _
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
style.LineStyle = LineStyle
style.Linethickness = Linethickness
style.PointStyle = PointStyle
style.UpColor = UpColor
Set gCreateDataPointStyle = style

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
                Optional ByVal Layer As Long = LayerHighestUser, _
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
                ByVal maType As String, _
                ByVal periods As Long, _
                ByVal numberOfValuesToCache As Long) As Study
Dim lparams As Parameters
Dim lStudy As Study
Dim valueNames(0) As String

Const ProcName As String = "gCreateMA"
On Error GoTo Err

valueNames(0) = "in"

Select Case UCase$(maType)
Case UCase$(EmaShortName)
    Dim lEMA As EMA
    Set lEMA = New EMA
    Set lStudy = lEMA
    Set lparams = GEMA.defaultParameters
    lparams.SetParameterValue ParamPeriods, periods
    lStudy.initialise GenerateGUIDString, _
                    lparams, _
                    numberOfValuesToCache, _
                    valueNames, _
                    Nothing, _
                    Nothing
                    
    Set gCreateMA = lEMA
Case Else
    Dim lSMA As SMA
    Set lSMA = New SMA
    Set lStudy = lSMA
    Set lparams = GSMA.defaultParameters
    lparams.SetParameterValue ParamPeriods, periods
    lStudy.initialise GenerateGUIDString, _
                    lparams, _
                    numberOfValuesToCache, _
                    valueNames, _
                    Nothing, _
                    Nothing
    Set gCreateMA = lSMA
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
                Optional ByVal Layer As Long = LayerHighestUser, _
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Function
'@================================================================================





