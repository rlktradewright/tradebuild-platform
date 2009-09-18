VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IGraphicObject

'@================================================================================
' Events
'@================================================================================

'Event Click()
'
'Event DblCLick()
'
'Event SelectionStateChanged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "DataPoint"

'================================================================================
' Member variables and constants
'================================================================================

Private mSeries As DataPointSeries

Private mViewport As Viewport
Private mCanvas As Canvas
Private mLayer As Long
Private mHandle As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mTimestamp As Date

Private WithEvents mLocalStyle As DataPointStyle
Attribute mLocalStyle.VB_VarHelpID = -1

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

Private mBoundingRect As TRectangle

Private mX As Double
Private mKey As String

Private mDataValue As Double
Private mBlank As Boolean   ' indicates that no value has been assigned

Private mPrevDataPointRef As WeakReference

Private mPropertyOverrideFlags As DataPointPropertyOverrideFlags

Private mSeriesUpdateNumber As Long

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True
mBoundingRect.isValid = False

End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "DataPoint terminated"
Debug.Print "DataPoint terminated"
End Sub

'================================================================================
' IGraphicObject Members
'================================================================================
Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Static prevGaugeX As Double
Static prevGaugeY As Double
Static prevRect As TRectangle
Dim viewportChanges As Long

Dim failpoint As Long
On Error GoTo Err

viewportChanges = mViewport.CompareTo(prevGaugeX, prevGaugeY, prevRect)
If (viewportChanges And ViewportComparisonCodes.GaugeChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.GaugeChangedY) Or _
    mSeriesUpdateNumber <> mSeries.SeriesUpdateNumber _
Then
    mSeriesUpdateNumber = mSeries.SeriesUpdateNumber
    calcBoundingRect
End If

prevGaugeX = mViewport.GaugeX
prevGaugeY = mViewport.GaugeY
prevRect = mViewport.Boundary

IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
IGraphicObject_capabilities = GraphicObjectCapabilities.PartialRedraw
End Property

Private Sub IGraphicObject_Click()
'RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
'RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
mCanvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areas(0))
Draw
mCanvas.ClearClippingRegion
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
Set mViewport = Nothing
End Sub

Private Property Get IGraphicObject_gaugeDependent() As Boolean
Select Case DisplayMode
Case DataPointDisplayModePoint
    IGraphicObject_gaugeDependent = True
Case DataPointDisplayModeLine
    IGraphicObject_gaugeDependent = True
Case DataPointDisplayModeStep
    IGraphicObject_gaugeDependent = True
Case DataPointDisplayModeHistogram
    IGraphicObject_gaugeDependent = False
End Select
End Property

Private Property Let IGraphicObject_Handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
IGraphicObject_HitTest = RectContainsPoint(mBoundingRect, X, Y)
End Function

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
IGraphicObject_IncludeInAutoscale = IncludeInAutoscale
End Property

Private Property Let IGraphicObject_inScope(ByVal value As Boolean)
mInScope = value
End Property

Private Property Get IGraphicObject_IsSelectable() As Boolean
IGraphicObject_IsSelectable = mIsSelectable
End Property

Private Property Get IGraphicObject_layer() As Long
IGraphicObject_layer = Layer
End Property

Private Property Get IGraphicObject_noDraw() As Boolean
IGraphicObject_noDraw = mBlank Or (DisplayMode = 0)
End Property

Private Property Get IGraphicObject_PeriodNumber() As Long
IGraphicObject_PeriodNumber = Int(mX)
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = False
End Property

Private Property Let IGraphicObject_Selected(ByVal RHS As Boolean)
If RHS = mIsSelected Then
Else
    mIsSelected = RHS
    'RaiseEvent SelectionStateChanged
End If
End Property

Private Property Get IGraphicObject_Selected() As Boolean
IGraphicObject_Selected = mIsSelected
End Property

Private Property Get IGraphicObject_SeriesID() As Long
IGraphicObject_SeriesID = mSeries.Id
End Property

Private Property Get IGraphicObject_Timestamp() As Date
IGraphicObject_Timestamp = mTimestamp
End Property

'Private Sub IGraphicObject_undraw( _
'                ByRef area As TRectangle, _
'                ByVal Hdc As Long)
'Dim failpoint As Long
'On Error GoTo Err
'
'undraw
'
'Exit Sub
'
'Err:
'Dim errNumber As Long: errNumber = Err.Number
'Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbcrlf, "") & ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint
'Dim errDescription As String: errDescription = Err.Description
'gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
'Err.Raise errNumber, errSource, errDescription
'
'
'End Sub

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property


'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "COLOR"
    initiateRedraw False, False
Case "DISPLAYMODE"
    initiateRedraw True, True
Case "DOWNCOLOR"
    initiateRedraw False, False
Case "HistogramBarWidth"
    initiateRedraw True, True
Case "INCLUDEINAUTOSCALE"
    initiateRedraw False, False
Case "LINESTYLE"
    initiateRedraw False, False
Case "LINETHICKNESS"
    initiateRedraw True, True
Case "POINTSTYLE"
    initiateRedraw False, False
Case "UPCOLOR"
    initiateRedraw False, False
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mLocalStyle_PropertyChanged", _
            "Unhandled property change"
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "color", _
            "Value is not a valid color"

If isPropertyOverrideFlagSet(DataPointIsSetColor) Then
    If getProperties.Color = value Then Exit Property
End If
setPropertyOverrideFlag DataPointIsSetColor
getProperties.Color = value
initiateRedraw False, False
End Property

Public Property Get Color() As Long
Color = mSeries.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If isPropertyOverrideFlagSet(DataPointIsSetColor) Then Color = getProperties.Color
End Property

Public Property Get DataValue() As Double
DataValue = mDataValue
End Property

Public Property Let DataValue(value As Double)
If value = mDataValue Then Exit Property
mBlank = False
mDataValue = value
initiateRedraw True, True
End Property

Public Property Let DisplayMode(ByVal value As DataPointDisplayModes)
If isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then
    If getProperties.DisplayMode = value Then Exit Property
End If
setPropertyOverrideFlag DataPointIsSetDisplayMode
getProperties.DisplayMode = value
initiateRedraw True, True
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
DisplayMode = mSeries.DisplayMode
If Not mLocalStyle Is Nothing Then DisplayMode = mLocalStyle.DisplayMode
If isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then DisplayMode = getProperties.DisplayMode
End Property

Public Property Let DownColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "downColor", _
            "Value is not a valid color or -1"

If isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then
    If getProperties.DownColor = value Then Exit Property
End If
setPropertyOverrideFlag DataPointIsSetDownColor
getProperties.DownColor = value
initiateRedraw False, False
End Property

Public Property Get DownColor() As Long
DownColor = mSeries.DownColor
If Not mLocalStyle Is Nothing Then DownColor = mLocalStyle.DownColor
If isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then DownColor = getProperties.DownColor
If DownColor = -1 Then DownColor = Color
End Property

Friend Sub Finish()
Set mPrevDataPointRef = Nothing
End Sub

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Let HistogramBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "HistogramBarWidth", _
            "HistogramBarWidth must be greater than zero but not greater than 1"
If isPropertyOverrideFlagSet(DataPointIsSetHistWidth) Then
    If getProperties.HistogramBarWidth = value Then Exit Property
End If
setPropertyOverrideFlag DataPointIsSetHistWidth
getProperties.HistogramBarWidth = value
initiateRedraw True, True
End Property

Public Property Get HistogramBarWidth() As Single
HistogramBarWidth = mSeries.HistogramBarWidth
If Not mLocalStyle Is Nothing Then HistogramBarWidth = mLocalStyle.HistogramBarWidth
If isPropertyOverrideFlagSet(DataPointIsSetHistWidth) Then HistogramBarWidth = getProperties.HistogramBarWidth
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mSeries.IncludeInAutoscale
If Not mLocalStyle Is Nothing Then IncludeInAutoscale = mLocalStyle.IncludeInAutoscale
If isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then IncludeInAutoscale = getProperties.IncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then
    If getProperties.IncludeInAutoscale = value Then Exit Property
End If
getProperties.IncludeInAutoscale = value
setPropertyOverrideFlag DataPointIsSetIncludeInAutoscale
initiateRedraw False, False
End Property

Public Property Get InScope() As Boolean
InScope = mInScope
End Property

Public Property Let IsSelectable(ByVal value As Boolean)
mIsSelectable = value
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let Layer(ByVal value As LayerNumbers)
If isPropertyOverrideFlagSet(DataPointIsSetLayer) Then
    If mLayer = value Then Exit Property
End If
mLayer = value
setPropertyOverrideFlag DataPointIsSetLayer
initiateRedraw False, False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mSeries.Layer
If isPropertyOverrideFlagSet(DataPointIsSetLayer) Then Layer = mLayer
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then
    If getProperties.LineStyle = value Then Exit Property
End If
getProperties.LineStyle = value
setPropertyOverrideFlag DataPointIsSetLineStyle
initiateRedraw False, False
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mSeries.LineStyle
If Not mLocalStyle Is Nothing Then LineStyle = mLocalStyle.LineStyle
If isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then LineStyle = getProperties.LineStyle
End Property

Public Property Let LineThickness(ByVal value As Long)
If isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then
    If getProperties.LineThickness = value Then Exit Property
End If
getProperties.LineThickness = value
setPropertyOverrideFlag DataPointIsSetLineThickness
initiateRedraw True, True
End Property

Public Property Get LineThickness() As Long
LineThickness = mSeries.LineThickness
If Not mLocalStyle Is Nothing Then LineThickness = mLocalStyle.LineThickness
If isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then LineThickness = getProperties.LineThickness
End Property

Public Property Get LocalStyle() As DataPointStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal value As DataPointStyle)

Set mLocalStyle = value
initiateRedraw True, True
End Property

Public Property Let PointStyle(ByVal value As PointStyles)
If isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then
    If getProperties.PointStyle = value Then Exit Property
End If
getProperties.PointStyle = value
setPropertyOverrideFlag DataPointIsSetPointStyle
initiateRedraw False, False
End Property

Public Property Get PointStyle() As PointStyles
PointStyle = mSeries.PointStyle
If Not mLocalStyle Is Nothing Then PointStyle = mLocalStyle.PointStyle
If isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then PointStyle = getProperties.PointStyle
End Property

Public Property Get PreviousDataPoint() As DataPoint
If mPrevDataPointRef Is Nothing Then Exit Property
Set PreviousDataPoint = prevDataPointObj
End Property

Friend Property Let PreviousDataPoint(ByVal value As DataPoint)
Set mPrevDataPointRef = CreateWeakReference(value)
End Property

Public Property Let UpColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "upColor", _
            "Value is not a valid color"
If isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then
    If getProperties.UpColor = value Then Exit Property
End If
getProperties.UpColor = value
setPropertyOverrideFlag DataPointIsSetUpColor
initiateRedraw False, False
End Property

Public Property Get UpColor() As Long
UpColor = mSeries.UpColor
If Not mLocalStyle Is Nothing Then UpColor = mLocalStyle.UpColor
If isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then UpColor = getProperties.UpColor
If UpColor = -1 Then UpColor = Color
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Public Property Let X(ByVal value As Double)
If value = mX Then Exit Property
mBlank = False
mX = value
initiateRedraw True, True
End Property

Public Property Get X() As Double
X = mX
End Property

'================================================================================
' Methods
'================================================================================

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
initiateRedraw True, True
End Sub

Friend Sub Initialise( _
                ByVal pSeries As DataPointSeries, _
                ByVal pTimestamp As Date, _
                ByVal pKey As String, _
                ByVal pGOSP As GraphObjServiceProvider, _
                ByVal pviewport As Viewport, _
                ByVal pLocalStyle As DataPointStyle)
                
Set mSeries = pSeries
mTimestamp = pTimestamp
mKey = pKey
Set mGOSP = pGOSP
Set mViewport = pviewport
Set mCanvas = mViewport.Canvas
Set mLocalStyle = pLocalStyle

calcBoundingRect
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcBoundingRect()

If DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    calcBoundingRectForHistogram
ElseIf drawAsPoint Then
    calcBoundingRectForPoint
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
    calcBoundingRectForLineOrSteppedLine
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
    calcBoundingRectForLineOrSteppedLine
End If
End Sub

Private Sub calcBoundingRectForHistogram()
Dim boundingRect As TRectangle
Dim pixelRect As RECT
    
With boundingRect
    .Left = mX - HistogramBarWidth / 2
    .Right = mX + HistogramBarWidth / 2
    If mDataValue >= 0 Then
        .Bottom = 0
        .Top = mDataValue
    Else
        .Bottom = mDataValue
        .Top = 0
    End If
End With
pixelRect = mViewport.ConvertTRectangleToRect(boundingRect)
mBoundingRect = mViewport.ConvertRectToTRectangle(pixelRect)
End Sub

Private Sub calcBoundingRectForLineOrSteppedLine()
Dim boundingRect As TRectangle
Dim pixelRect As RECT
Dim prevX As Double
Dim prevValue As Double

prevX = prevDataPointObj.X
prevValue = prevDataPointObj.DataValue

With boundingRect
    If prevX <= mX Then
        .Left = prevX
        .Right = mX
    Else
        .Left = mX
        .Right = prevX
    End If
    If mDataValue > prevValue Then
        .Bottom = prevValue
        .Top = mDataValue
    Else
        .Bottom = mDataValue
        .Top = prevValue
    End If
End With

pixelRect = mViewport.ConvertTRectangleToRect(boundingRect)

With pixelRect
    .Left = .Left - Int(LineThickness / 2)
    .Right = .Right + Int(LineThickness / 2)
    .Top = .Top - Int(LineThickness / 2)
    .Bottom = .Bottom + Int(LineThickness / 2)
End With

mBoundingRect = mViewport.ConvertRectToTRectangle(pixelRect)
End Sub

Private Sub calcBoundingRectForPoint()
Dim boundingRect As TRectangle
Dim pixelRect As RECT

With boundingRect
    .Left = mX
    .Right = mX
    .Bottom = mDataValue
    .Top = mDataValue
End With

pixelRect = mViewport.ConvertTRectangleToRect(boundingRect)

If PointStyle = PointRound Then
    If LineThickness <> 1 Then
        With pixelRect
            ' these offsets were discovered by measuring points drawn using PSet
            .Left = .Left - Int(LineThickness / 2)
            .Right = .Left + LineThickness - 1
            .Bottom = .Top + Int((LineThickness + 1) / 2)
            .Top = .Top - Int((LineThickness - 1) / 2)
        End With
    End If
Else
    With pixelRect
        .Left = .Left - Int(LineThickness / 2)
        .Right = .Left + LineThickness - 1
        .Top = .Top - Int(LineThickness / 2)
        .Bottom = .Top + LineThickness - 1
    End With
End If

mBoundingRect = mViewport.ConvertRectToTRectangle(pixelRect)
End Sub

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub Draw()
Dim lColor As Long

If Not mPrevDataPointRef Is Nothing Then
    If mDataValue >= prevDataPointObj.DataValue Then
        lColor = UpColor
    Else
        lColor = DownColor
    End If
Else
    lColor = Color
End If
DrawPoint lColor
End Sub

Private Function drawAsPoint() As Boolean
If mPrevDataPointRef Is Nothing Then
    drawAsPoint = True
ElseIf prevDataPointObj.Blank Then
    drawAsPoint = True
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Then
    drawAsPoint = True
Else
    drawAsPoint = False
End If
End Function

Private Sub DrawPoint(ByVal Color As Long)
Dim PreviousDataPoint As DataPoint

If Blank Then Exit Sub
If Not mVisible Then Exit Sub

If DisplayMode = 0 Then Exit Sub

With mViewport
    .Canvas.SetPenAttributes Color, _
                            LineThickness, _
                            LineStyle, _
                            DrawModes.DrawModeCopyPen
    
    If DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
        .Canvas.SetBrushAttributes Color, FillStyles.FillSolid
        .Canvas.DrawRectangle .NewPoint(mX - HistogramBarWidth / 2, 0), .NewPoint(mX + HistogramBarWidth / 2, mDataValue)
    ElseIf drawAsPoint Then
        If PointStyle = PointRound Then
            .Canvas.DrawPoint .NewPoint(mX, mDataValue)
        Else
            Dim ThicknessScaleX As Double
            Dim ThicknessScaleY As Double
            
            ThicknessScaleX = .ConvertPixelsToLogicalX(LineThickness)
            ThicknessScaleY = .ConvertPixelsToLogicalY(LineThickness)
            .Canvas.SetPenAttributes Color, _
                                    1, _
                                    LineInsideSolid, _
                                    DrawModes.DrawModeCopyPen
            .Canvas.SetBrushAttributes Color, FillSolid
            .Canvas.DrawRectangle .NewPoint(mX - ThicknessScaleX / 2, mDataValue - ThicknessScaleY / 2), _
                            .NewPoint(mX + ThicknessScaleX / 2, mDataValue + ThicknessScaleY / 2)
        End If
    ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
        Set PreviousDataPoint = prevDataPointObj
        .Canvas.DrawLine prevDataPointObj.X, prevDataPointObj.DataValue, mX, mDataValue
    ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
        Set PreviousDataPoint = prevDataPointObj
        .Canvas.DrawLine prevDataPointObj.X, prevDataPointObj.DataValue, mX - 0.5, prevDataPointObj.DataValue
        .Canvas.DrawLine mX - 0.5, prevDataPointObj.DataValue, mX - 0.5, mDataValue
        .Canvas.DrawLine mX - 0.5, mDataValue, mX, mDataValue
    End If
End With
End Sub

Private Function getProperties() As DataPointProperties
Static props As DataPointProperties
If props Is Nothing Then Set props = New DataPointProperties
Set getProperties = props
End Function

Private Sub initiateRedraw( _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
If undrawCurrentImage Then undraw
If recalcBoundaries Then
    calcBoundingRect
End If
mGOSP.ObjectChanged mHandle, mBoundingRect
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As DataPointPropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Function prevDataPointObj() As DataPoint
Set prevDataPointObj = mPrevDataPointRef.Target
End Function

Private Sub setPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub undraw()
If Not mVisible Then Exit Sub
mGOSP.ObjectUndrawn mHandle, mBoundingRect
End Sub


