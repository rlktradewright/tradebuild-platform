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

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "DataPoint"

'================================================================================
' Member variables and constants
'================================================================================

Private mCanvas As Canvas
Private mLayer As Long
Private mHandle As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean
Private mSeriesID As Long

Private mTimestamp As Date

Private mLocalStyle As DataPointStyle
Private mBaseStyle As DataPointStyle

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

Private mBoundingRect As TRectangle

Private mX As Double
Private mKey As String

Private mIsSetLineThickness As Boolean
Private mIsSetColor As Boolean
Private mIsSetUpColor As Boolean
Private mIsSetDownColor As Boolean
Private mIsSetLineStyle As Boolean
Private mIsSetPointStyle As Boolean
Private mIsSetDisplayMode As Boolean
Private mIsSetHistBarWidth As Boolean
Private mIsSetIncludeInAutoscale As Boolean

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

canvasChanges = mCanvas.compareTo(prevGaugeX, prevGaugeY, prevRect)
If (canvasChanges And CanvasComparisonCodes.GaugeChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.GaugeChangedY) _
Then
    mBoundingRect = calcBoundingRect
End If

prevGaugeX = mCanvas.gaugeX
prevGaugeY = mCanvas.gaugeY
prevRect = mCanvas.boundary

IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
Select Case displayMode
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

Private Function IGraphicObject_HitTest(ByVal x As Double, ByVal y As Double) As Boolean
IGraphicObject_HitTest = RectContains(mBoundingRect, x, y)
End Function

Private Property Get IGraphicObject_includeInAutoscale() As Boolean
IGraphicObject_includeInAutoscale = IncludeInAutoscale
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
IGraphicObject_noDraw = mBlank Or (displayMode = 0)
End Property

Private Property Get IGraphicObject_periodNumber() As Long
IGraphicObject_periodNumber = Int(mX)
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
IGraphicObject_SeriesID = mSeriesID
End Property

Private Property Get IGraphicObject_timestamp() As Date
IGraphicObject_timestamp = mTimestamp
End Property

'Private Sub IGraphicObject_undraw( _
'                ByRef area As TRectangle, _
'                ByVal hdc As Long)
'Dim failpoint As Long
'On Error GoTo Err
'
'undraw
'
'Exit Sub
'
'Err:
'Dim errNumber As Long: errNumber = Err.Number
'Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
'Dim errDescription As String: errDescription = Err.Description
'gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
'Err.Raise errNumber, errSource, errDescription
'
'
'End Sub

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property


'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get blank() As Boolean
blank = mBlank
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "color", _
            "Value is not a valid color"

If mColor = value And mIsSetColor Then Exit Property
mColor = value
mIsSetColor = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get Color() As Long
Color = mBaseStyle.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If mIsSetColor Then Color = mColor
End Property

Public Property Get dataValue() As Double
dataValue = mDataValue
End Property

Public Property Let dataValue(value As Double)
If value = mDataValue Then Exit Property
undraw  ' remove current point
mBlank = False
mDataValue = value
update
End Property

Public Property Let displayMode(ByVal value As DataPointDisplayModes)
If mDisplayMode = value And mIsSetDisplayMode Then Exit Property
undraw
mDisplayMode = value
mIsSetDisplayMode = True
calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get displayMode() As DataPointDisplayModes
displayMode = mBaseStyle.displayMode
If Not mLocalStyle Is Nothing Then displayMode = mLocalStyle.displayMode
If mIsSetDisplayMode Then displayMode = mDisplayMode
End Property

Public Property Let downColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "downColor", _
            "Value is not a valid color or -1"

If mDownColor = value And mIsSetDownColor Then Exit Property
mDownColor = value
mIsSetDownColor = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get downColor() As Long
downColor = mBaseStyle.downColor
If Not mLocalStyle Is Nothing Then downColor = mLocalStyle.downColor
If mIsSetDownColor Then downColor = mDownColor
If downColor = -1 Then downColor = Color
End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Let histBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"
If mHistBarWidth = value And mIsSetHistBarWidth Then Exit Property
undraw
mHistBarWidth = value
mIsSetHistBarWidth = True
calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get histBarWidth() As Single
histBarWidth = mBaseStyle.histBarWidth
If Not mLocalStyle Is Nothing Then histBarWidth = mLocalStyle.histBarWidth
If mIsSetHistBarWidth Then histBarWidth = mHistBarWidth
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And mIsSetIncludeInAutoscale Then Exit Property
If Not value Then undraw
mIncludeInAutoscale = value
mIsSetIncludeInAutoscale = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mBaseStyle.IncludeInAutoscale
If Not mLocalStyle Is Nothing Then IncludeInAutoscale = mLocalStyle.IncludeInAutoscale
If mIsSetIncludeInAutoscale Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Get InScope() As Boolean
InScope = mInScope
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Let Layer(ByVal value As Long)
If mLayer = value Then Exit Property
mLayer = value
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get Layer() As Long
Layer = mLayer
End Property

Public Property Let linestyle(ByVal value As LineStyles)
If mLineStyle = value And mIsSetLineStyle Then Exit Property
mLineStyle = value
mIsSetLineStyle = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get linestyle() As LineStyles
linestyle = mBaseStyle.linestyle
If Not mLocalStyle Is Nothing Then linestyle = mLocalStyle.linestyle
If mIsSetLineStyle Then linestyle = mLineStyle
End Property

Public Property Let lineThickness(ByVal value As Long)
If mLineThickness = value And mIsSetLineThickness Then Exit Property
undraw
mLineThickness = value
mIsSetLineThickness = True
calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get lineThickness() As Long
lineThickness = mBaseStyle.lineThickness
If Not mLocalStyle Is Nothing Then lineThickness = mLocalStyle.lineThickness
If mIsSetLineThickness Then lineThickness = mLineThickness
End Property

Public Property Get localStyle() As DataPointStyle
Set localStyle = mLocalStyle.clone
End Property

Public Property Let localStyle( _
                ByVal value As DataPointStyle)

undraw

Set mLocalStyle = value.clone
calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect

End Property

Public Property Let pointStyle(ByVal value As PointStyles)
If mPointStyle = value And mIsSetPointStyle Then Exit Property
mPointStyle = value
mIsSetPointStyle = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get pointStyle() As PointStyles
pointStyle = mBaseStyle.pointStyle
If Not mLocalStyle Is Nothing Then pointStyle = mLocalStyle.pointStyle
If mIsSetPointStyle Then pointStyle = mPointStyle
End Property

Public Property Get prevDataPoint() As DataPoint
Set prevDataPoint = mPrevDataPoint
End Property

Public Property Let prevDataPoint(ByVal value As DataPoint)
Set mPrevDataPoint = value
End Property

Public Property Let upColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPoint" & ":" & "upColor", _
            "Value is not a valid color"

If mUpColor = value And mIsSetUpColor Then Exit Property
mUpColor = value
mIsSetUpColor = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get upColor() As Long
upColor = mBaseStyle.upColor
If Not mLocalStyle Is Nothing Then upColor = mLocalStyle.upColor
If mIsSetUpColor Then upColor = mUpColor
If upColor = -1 Then upColor = Color
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Public Property Let x(ByVal value As Double)
If value = mX Then Exit Property
undraw  ' remove current point
mBlank = False
mX = value
update
End Property

Public Property Get x() As Double
x = mX
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal seriesId As Long, _
                ByVal timestamp As Date, _
                ByVal key As String, _
                ByVal GOSP As GraphObjServiceProvider, _
                ByVal Layer As Long, _
                ByVal Canvas As Canvas, _
                ByVal baseStyle As DataPointStyle, _
                ByVal localStyle As DataPointStyle, _
                Optional ByVal IsSelectable As Boolean)
                
mSeriesID = seriesId
mTimestamp = timestamp
mKey = key
Set mGOSP = GOSP
mLayer = Layer
Set mCanvas = Canvas
Set mBaseStyle = baseStyle
If Not localStyle Is Nothing Then Set mLocalStyle = localStyle.clone ' local style cannot be shared
mIsSelectable = IsSelectable

calcBoundingRect
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcBoundingRect() As TRectangle
Dim boundingRect As TRectangle
Dim thicknessScaleX As Double
Dim thicknessScaleY As Double
Dim prevX As Long
Dim prevValue As Double

thicknessScaleX = mCanvas.convertPixelsToLogicalX(lineThickness + 2)
thicknessScaleY = mCanvas.convertPixelsToLogicalY(lineThickness + 2)

If displayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    With boundingRect
        .left = mX - histBarWidth / 2
        .right = mX + histBarWidth / 2
        If mDataValue >= 0 Then
            .bottom = 0
            .top = mDataValue
        Else
            .bottom = mDataValue
            .top = 0
        End If
    End With
ElseIf drawAsPoint Then
    With boundingRect
        .left = mX - thicknessScaleX / 2
        .right = mX + thicknessScaleX / 2
        .bottom = mDataValue - thicknessScaleY / 2
        .top = mDataValue + thicknessScaleY / 2
    End With
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
    With mPrevDataPoint
        prevX = .x
        prevValue = .dataValue
    End With
    With boundingRect
        If prevX <= mX Then
            .left = prevX - thicknessScaleX / 2
            .right = mX + thicknessScaleX / 2
        Else
            .left = mX - thicknessScaleX / 2
            .right = prevX + thicknessScaleX / 2
        End If
        If mDataValue > prevValue Then
            .bottom = prevValue - thicknessScaleY / 2
            .top = mDataValue + thicknessScaleY / 2
        Else
            .bottom = mDataValue - thicknessScaleY / 2
            .top = prevValue + thicknessScaleY / 2
        End If
    End With
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
    With mPrevDataPoint
        prevX = .x
        prevValue = .dataValue
    End With
    With boundingRect
        If prevX <= mX Then
            .left = prevX - thicknessScaleX / 2
            .right = mX + thicknessScaleX / 2
        Else
            .left = mX - thicknessScaleX / 2
            .right = prevX + thicknessScaleX / 2
        End If
        If mDataValue > prevValue Then
            .bottom = prevValue - thicknessScaleY / 2
            .top = mDataValue + thicknessScaleY / 2
        Else
            .bottom = mDataValue - thicknessScaleY / 2
            .top = prevValue + thicknessScaleY / 2
        End If
    End With
End If
boundingRect.isValid = True
calcBoundingRect = boundingRect
End Function

Private Sub Draw()
Dim lColor As Long

If Not mPrevDataPoint Is Nothing Then
    If mDataValue >= mPrevDataPoint.dataValue Then
        lColor = upColor
    Else
        lColor = downColor
    End If
Else
    lColor = Color
End If
drawPoint lColor
End Sub

Private Function drawAsPoint() As Boolean
If mPrevDataPoint Is Nothing Then
    drawAsPoint = True
ElseIf mPrevDataPoint.blank Then
    drawAsPoint = True
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModePoint Then
'ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Or _
'    (DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine And Not mPrevDataPoint.inScope) Or _
'    (DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep And Not mPrevDataPoint.inScope) _
'Then
    drawAsPoint = True
Else
    drawAsPoint = False
End If
End Function

Private Sub drawPoint(ByVal Color As Long)
If blank Then Exit Sub
If Not mVisible Then Exit Sub

If displayMode = 0 Then Exit Sub

With mCanvas
    .setPenAttributes Color, _
                            lineThickness, _
                            linestyle, _
                            DrawModes.DrawModeCopyPen
    
    If displayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
        .setBrushAttributes Color, FillStyles.FillSolid
        .drawRectangle .newPoint(mX - histBarWidth / 2, 0), .newPoint(mX + histBarWidth / 2, mDataValue)
    ElseIf drawAsPoint Then
        If pointStyle = PointRound Then
            .drawPoint .newPoint(mX, mDataValue)
        Else
            Dim thicknessScaleX As Double
            Dim thicknessScaleY As Double
            
            thicknessScaleX = .convertPixelsToLogicalX(lineThickness)
            thicknessScaleY = .convertPixelsToLogicalY(lineThickness)
            .setPenAttributes Color, _
                                    1, _
                                    LineInsideSolid, _
                                    DrawModes.DrawModeCopyPen
            .setBrushAttributes Color, FillSolid
            .drawRectangle .newPoint(mX - thicknessScaleX / 2, mDataValue - thicknessScaleY / 2), _
                            .newPoint(mX + thicknessScaleX / 2, mDataValue + thicknessScaleY / 2)
        End If
    ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
        .drawLine .newPoint(mPrevDataPoint.x, mPrevDataPoint.dataValue), .newPoint(mX, mDataValue)
    ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
        .drawLine .newPoint(mPrevDataPoint.x, mPrevDataPoint.dataValue), .newPoint(mX - 0.5, mPrevDataPoint.dataValue)
        .drawLine .newPoint(mX - 0.5, mPrevDataPoint.dataValue), .newPoint(mX - 0.5, mDataValue)
        .drawLine .newPoint(mX - 0.5, mDataValue), .newPoint(mX, mDataValue)
    End If
End With
End Sub

Private Sub undraw()
If Not mVisible Then Exit Sub
'drawPoint mCanvas.backColor
mGOSP.ObjectUndrawn mHandle, mBoundingRect
End Sub

Private Sub update()
If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
mBoundingRect = calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect
End Sub

