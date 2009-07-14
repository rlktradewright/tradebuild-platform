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

Event Click()

Event DblCLick()

Event SelectionStateChanged()

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

Private mLineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mPointStyle As PointStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistBarWidth As Single
Private mIncludeInAutoscale As Boolean

Private mDataValue As Double
Private mBlank As Boolean   ' indicates that no value has been assigned

Private mPrevDataPoint As DataPoint

Private mPropertyOverrideFlags As DataPointPropertyOverrideFlags

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
Debug.Print "DataPoint terminated"
End Sub

'================================================================================
' IGraphicObject Members
'================================================================================
Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Static prevGaugeX As Double
Static prevGaugeY As Double
Static prevRect As TRectangle
Dim canvasChanges As Long

Dim failpoint As Long
On Error GoTo Err

canvasChanges = mCanvas.CompareTo(prevGaugeX, prevGaugeY, prevRect)
If (canvasChanges And CanvasComparisonCodes.GaugeChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.GaugeChangedY) _
Then
    calcBoundingRect
End If

prevGaugeX = mCanvas.GaugeX
prevGaugeY = mCanvas.GaugeY
prevRect = mCanvas.Boundary

IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
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
RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
Draw
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
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

Private Property Let IGraphicObject_handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_handle() As Long
IGraphicObject_handle = mHandle
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
    RaiseEvent SelectionStateChanged
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
'Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint & IIf(err.Source <> "", vbCrLf & err.Source, "")
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
Dim s As DataPointStyle
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("Color")
    initiateRedraw False, False
Case UCase$("displayMode")
    initiateRedraw True, True
Case UCase$("downColor")
    initiateRedraw False, False
Case UCase$("histBarWidth")
    initiateRedraw True, True
Case UCase$("LineStyle")
    initiateRedraw False, False
Case UCase$("lineThickness")
    initiateRedraw True, True
Case UCase$("pointStyle")
    initiateRedraw False, False
Case UCase$("upColor")
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

Public Property Let Color(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "color", _
            "Value is not a valid color"

If mColor = value And isPropertyOverrideFlagSet(DataPointIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag DataPointIsSetColor
initiateRedraw False, False
End Property

Public Property Get Color() As OLE_COLOR
Color = mSeries.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If isPropertyOverrideFlagSet(DataPointIsSetColor) Then Color = mColor
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
If mDisplayMode = value And isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then Exit Property
mDisplayMode = value
setPropertyOverrideFlag DataPointIsSetDisplayMode
initiateRedraw True, True
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
DisplayMode = mSeries.DisplayMode
If Not mLocalStyle Is Nothing Then DisplayMode = mLocalStyle.DisplayMode
If isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then DisplayMode = mDisplayMode
End Property

Public Property Let DownColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "downColor", _
            "Value is not a valid color or -1"

If mDownColor = value And isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then Exit Property
mDownColor = value
setPropertyOverrideFlag DataPointIsSetDownColor
initiateRedraw False, False
End Property

Public Property Get DownColor() As OLE_COLOR
DownColor = mSeries.DownColor
If Not mLocalStyle Is Nothing Then DownColor = mLocalStyle.DownColor
If isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then DownColor = mDownColor
If DownColor = -1 Then DownColor = Color
End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Let HistBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"
If mHistBarWidth = value And isPropertyOverrideFlagSet(DataPointIsSetHistBarWidth) Then Exit Property
mHistBarWidth = value
setPropertyOverrideFlag DataPointIsSetHistBarWidth
initiateRedraw True, True
End Property

Public Property Get HistBarWidth() As Single
HistBarWidth = mSeries.HistBarWidth
If Not mLocalStyle Is Nothing Then HistBarWidth = mLocalStyle.HistBarWidth
If isPropertyOverrideFlagSet(DataPointIsSetHistBarWidth) Then HistBarWidth = mHistBarWidth
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mSeries.IncludeInAutoscale
If isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
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
If mLayer = value And isPropertyOverrideFlagSet(DataPointIsSetLayer) Then Exit Property
mLayer = value
setPropertyOverrideFlag DataPointIsSetLayer
initiateRedraw False, False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mSeries.Layer
If isPropertyOverrideFlagSet(DataPointIsSetLayer) Then Layer = mLayer
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If mLineStyle = value And isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag DataPointIsSetLineStyle
initiateRedraw False, False
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mSeries.LineStyle
If Not mLocalStyle Is Nothing Then LineStyle = mLocalStyle.LineStyle
If isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then LineStyle = mLineStyle
End Property

Public Property Let LineThickness(ByVal value As Long)
If mLineThickness = value And isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then Exit Property
mLineThickness = value
setPropertyOverrideFlag DataPointIsSetLineThickness
initiateRedraw True, True
End Property

Public Property Get LineThickness() As Long
LineThickness = mSeries.LineThickness
If Not mLocalStyle Is Nothing Then LineThickness = mLocalStyle.LineThickness
If isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then LineThickness = mLineThickness
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
If mPointStyle = value And isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then Exit Property
mPointStyle = value
setPropertyOverrideFlag DataPointIsSetPointStyle
initiateRedraw False, False
End Property

Public Property Get PointStyle() As PointStyles
PointStyle = mSeries.PointStyle
If Not mLocalStyle Is Nothing Then PointStyle = mLocalStyle.PointStyle
If isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then PointStyle = mPointStyle
End Property

Public Property Get PrevDataPoint() As DataPoint
Set PrevDataPoint = mPrevDataPoint
End Property

Public Property Let PrevDataPoint(ByVal value As DataPoint)
Set mPrevDataPoint = value
End Property

Public Property Let UpColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "upColor", _
            "Value is not a valid color"

If mUpColor = value And isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then Exit Property
mUpColor = value
setPropertyOverrideFlag DataPointIsSetUpColor
initiateRedraw False, False
End Property

Public Property Get UpColor() As OLE_COLOR
UpColor = mSeries.UpColor
If Not mLocalStyle Is Nothing Then UpColor = mLocalStyle.UpColor
If isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then UpColor = mUpColor
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

Friend Sub Initialise( _
                ByVal series As DataPointSeries, _
                ByVal Timestamp As Date, _
                ByVal Key As String, _
                ByVal GOSP As GraphObjServiceProvider, _
                ByVal Canvas As Canvas, _
                ByVal LocalStyle As DataPointStyle)
                
Set mSeries = series
mTimestamp = Timestamp
mKey = Key
Set mGOSP = GOSP
Set mCanvas = Canvas
Set mLocalStyle = LocalStyle

calcBoundingRect
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcBoundingRect()
Dim boundingRect As TRectangle
Dim thicknessScaleX As Double
Dim thicknessScaleY As Double
Dim prevX As Long
Dim prevValue As Double

thicknessScaleX = mCanvas.ConvertPixelsToLogicalX(LineThickness + 2)
thicknessScaleY = mCanvas.ConvertPixelsToLogicalY(LineThickness + 2)

If DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    With boundingRect
        .Left = mX - HistBarWidth / 2
        .Right = mX + HistBarWidth / 2
        If mDataValue >= 0 Then
            .Bottom = 0
            .Top = mDataValue
        Else
            .Bottom = mDataValue
            .Top = 0
        End If
    End With
ElseIf drawAsPoint Then
    With boundingRect
        .Left = mX - thicknessScaleX / 2
        .Right = mX + thicknessScaleX / 2
        .Bottom = mDataValue - thicknessScaleY / 2
        .Top = mDataValue + thicknessScaleY / 2
    End With
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
    With mPrevDataPoint
        prevX = .X
        prevValue = .DataValue
    End With
    With boundingRect
        If prevX <= mX Then
            .Left = prevX - thicknessScaleX / 2
            .Right = mX + thicknessScaleX / 2
        Else
            .Left = mX - thicknessScaleX / 2
            .Right = prevX + thicknessScaleX / 2
        End If
        If mDataValue > prevValue Then
            .Bottom = prevValue - thicknessScaleY / 2
            .Top = mDataValue + thicknessScaleY / 2
        Else
            .Bottom = mDataValue - thicknessScaleY / 2
            .Top = prevValue + thicknessScaleY / 2
        End If
    End With
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
    With mPrevDataPoint
        prevX = .X
        prevValue = .DataValue
    End With
    With boundingRect
        If prevX <= mX Then
            .Left = prevX - thicknessScaleX / 2
            .Right = mX + thicknessScaleX / 2
        Else
            .Left = mX - thicknessScaleX / 2
            .Right = prevX + thicknessScaleX / 2
        End If
        If mDataValue > prevValue Then
            .Bottom = prevValue - thicknessScaleY / 2
            .Top = mDataValue + thicknessScaleY / 2
        Else
            .Bottom = mDataValue - thicknessScaleY / 2
            .Top = prevValue + thicknessScaleY / 2
        End If
    End With
End If
boundingRect.isValid = True
mBoundingRect = boundingRect
End Sub

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub Draw()
Dim lColor As Long

If Not mPrevDataPoint Is Nothing Then
    If mDataValue >= mPrevDataPoint.DataValue Then
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
If mPrevDataPoint Is Nothing Then
    drawAsPoint = True
ElseIf mPrevDataPoint.Blank Then
    drawAsPoint = True
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Then
'ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Or _
'    (DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine And Not mPrevDataPoint.inScope) Or _
'    (DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep And Not mPrevDataPoint.inScope) _
'Then
    drawAsPoint = True
Else
    drawAsPoint = False
End If
End Function

Private Sub DrawPoint(ByVal Color As Long)
If Blank Then Exit Sub
If Not mVisible Then Exit Sub

If DisplayMode = 0 Then Exit Sub

With mCanvas
    .SetPenAttributes Color, _
                            LineThickness, _
                            LineStyle, _
                            DrawModes.DrawModeCopyPen
    
    If DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
        .SetBrushAttributes Color, FillStyles.FillSolid
        .DrawRectangle .NewPoint(mX - HistBarWidth / 2, 0), .NewPoint(mX + HistBarWidth / 2, mDataValue)
    ElseIf drawAsPoint Then
        If PointStyle = PointRound Then
            .DrawPoint .NewPoint(mX, mDataValue)
        Else
            Dim thicknessScaleX As Double
            Dim thicknessScaleY As Double
            
            thicknessScaleX = .ConvertPixelsToLogicalX(LineThickness)
            thicknessScaleY = .ConvertPixelsToLogicalY(LineThickness)
            .SetPenAttributes Color, _
                                    1, _
                                    LineInsideSolid, _
                                    DrawModes.DrawModeCopyPen
            .SetBrushAttributes Color, FillSolid
            .DrawRectangle .NewPoint(mX - thicknessScaleX / 2, mDataValue - thicknessScaleY / 2), _
                            .NewPoint(mX + thicknessScaleX / 2, mDataValue + thicknessScaleY / 2)
        End If
    ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
        .DrawLine .NewPoint(mPrevDataPoint.X, mPrevDataPoint.DataValue), .NewPoint(mX, mDataValue)
    ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
        .DrawLine .NewPoint(mPrevDataPoint.X, mPrevDataPoint.DataValue), .NewPoint(mX - 0.5, mPrevDataPoint.DataValue)
        .DrawLine .NewPoint(mX - 0.5, mPrevDataPoint.DataValue), .NewPoint(mX - 0.5, mDataValue)
        .DrawLine .NewPoint(mX - 0.5, mDataValue), .NewPoint(mX, mDataValue)
    End If
End With
End Sub

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

Private Sub setPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub undraw()
If Not mVisible Then Exit Sub
mGOSP.ObjectUndrawn mHandle, mBoundingRect
End Sub


