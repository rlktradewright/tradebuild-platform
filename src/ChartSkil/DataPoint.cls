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

Private mViewport As ViewPort
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

Private mPropertyOverrideFlags As DataPointPropertyFlags

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
gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Proc", pModName:=ModuleName, pMsg:="DataPoint terminated"
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

Const ProcName As String = "IGraphicObject_boundingRectangle"

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
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IGraphicObject_BoundingRectangleInViewport() As TRectangle
Const ProcName As String = "IGraphicObject_BoundingRectangleInViewport"

On Error GoTo Err

IGraphicObject_BoundingRectangleInViewport = IGraphicObject_boundingRectangle
RectSetXInterval IGraphicObject_BoundingRectangleInViewport, _
                IntIntersection(RectGetXInterval(IGraphicObject_BoundingRectangleInViewport), _
                                RectGetXInterval(mViewport.Boundary))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
Const ProcName As String = "IGraphicObject_capabilities"

On Error GoTo Err

IGraphicObject_capabilities = GraphicObjectCapabilities.PartialRedraw

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IGraphicObject_Click()
'RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
'RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
Const ProcName As String = "IGraphicObject_draw"

On Error GoTo Err

'mCanvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areas(0))
Draw
'mCanvas.ClearClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Property Let IGraphicObject_Handle(ByVal Value As Long)
mHandle = Value
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
Const ProcName As String = "IGraphicObject_HitTest"

On Error GoTo Err

IGraphicObject_HitTest = RectContainsPoint(mBoundingRect, X, Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
Const ProcName As String = "IGraphicObject_IncludeInAutoscale"

On Error GoTo Err

IGraphicObject_IncludeInAutoscale = IncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let IGraphicObject_inScope(ByVal Value As Boolean)
mInScope = Value
End Property

Private Property Get IGraphicObject_IsSelectable() As Boolean
IGraphicObject_IsSelectable = mIsSelectable
End Property

Private Property Get IGraphicObject_layer() As Long
IGraphicObject_layer = layer
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
Const ProcName As String = "IGraphicObject_SeriesID"

On Error GoTo Err

IGraphicObject_SeriesID = mSeries.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IGraphicObject_Timestamp() As Date
IGraphicObject_Timestamp = mTimestamp
End Property

Private Property Get IGraphicObject_VerticalRangeInXinterval(ByRef pXInterval As TInterval) As TInterval
If Not pXInterval.isValid Then Exit Property
If mX < pXInterval.startValue Then Exit Property
If mX > pXInterval.endValue Then Exit Property

IGraphicObject_VerticalRangeInXinterval = RectGetYInterval(IGraphicObject_boundingRectangle)
End Property

Private Property Let IGraphicObject_visible(ByVal Value As Boolean)
mVisible = Value
End Property


'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mLocalStyle_PropertyChanged"

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
Case "LAYER"
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
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Unhandled property change"
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"

On Error GoTo Err

If Not gIsValidColor(Value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color"

If IsPropertySet(DataPointPropertyColor) Then
    If getProperties.Color = Value Then Exit Property
End If
setPropertyFlag DataPointPropertyColor
getProperties.Color = Value
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"

On Error GoTo Err

Color = mSeries.Color
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyColor) Then Color = mLocalStyle.Color
If IsPropertySet(DataPointPropertyColor) Then Color = getProperties.Color

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DataValue() As Double
DataValue = mDataValue
End Property

Public Property Let DataValue(Value As Double)
Const ProcName As String = "DataValue"

On Error GoTo Err

If Value = mDataValue Then Exit Property
mBlank = False
mDataValue = Value
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DisplayMode(ByVal Value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"

On Error GoTo Err

If IsPropertySet(DataPointPropertyDisplayMode) Then
    If getProperties.DisplayMode = Value Then Exit Property
End If
setPropertyFlag DataPointPropertyDisplayMode
getProperties.DisplayMode = Value
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
Const ProcName As String = "DisplayMode"

On Error GoTo Err

DisplayMode = mSeries.DisplayMode
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyDisplayMode) Then DisplayMode = mLocalStyle.DisplayMode
If IsPropertySet(DataPointPropertyDisplayMode) Then DisplayMode = getProperties.DisplayMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownColor(ByVal Value As Long)
Const ProcName As String = "DownColor"

On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color or -1"

If IsPropertySet(DataPointPropertyDownColor) Then
    If getProperties.DownColor = Value Then Exit Property
End If
setPropertyFlag DataPointPropertyDownColor
getProperties.DownColor = Value
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownColor() As Long
Const ProcName As String = "DownColor"

On Error GoTo Err

DownColor = mSeries.DownColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyDownColor) Then DownColor = mLocalStyle.DownColor
If IsPropertySet(DataPointPropertyDownColor) Then DownColor = getProperties.DownColor
If DownColor = -1 Then DownColor = Color

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Sub Finish()
Set mPrevDataPointRef = Nothing
End Sub

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Let HistogramBarWidth(ByVal Value As Single)
Const ProcName As String = "HistogramBarWidth"

On Error GoTo Err

If Value <= 0 Or Value > 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "HistogramBarWidth must be greater than zero but not greater than 1"
If IsPropertySet(DataPointPropertyHistWidth) Then
    If getProperties.HistogramBarWidth = Value Then Exit Property
End If
setPropertyFlag DataPointPropertyHistWidth
getProperties.HistogramBarWidth = Value
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HistogramBarWidth() As Single
Const ProcName As String = "HistogramBarWidth"

On Error GoTo Err

HistogramBarWidth = mSeries.HistogramBarWidth
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyHistWidth) Then HistogramBarWidth = mLocalStyle.HistogramBarWidth
If IsPropertySet(DataPointPropertyHistWidth) Then HistogramBarWidth = getProperties.HistogramBarWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

IncludeInAutoscale = mSeries.IncludeInAutoscale
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyIncludeInAutoscale) Then IncludeInAutoscale = mLocalStyle.IncludeInAutoscale
If IsPropertySet(DataPointPropertyIncludeInAutoscale) Then IncludeInAutoscale = getProperties.IncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

If IsPropertySet(DataPointPropertyIncludeInAutoscale) Then
    If getProperties.IncludeInAutoscale = Value Then Exit Property
End If
getProperties.IncludeInAutoscale = Value
setPropertyFlag DataPointPropertyIncludeInAutoscale
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InScope() As Boolean
InScope = mInScope
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
mIsSelectable = Value
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

Public Property Let layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"

On Error GoTo Err

If IsPropertySet(DataPointPropertyLayer) Then
    If mLayer = Value Then Exit Property
End If
mLayer = Value
setPropertyFlag DataPointPropertyLayer
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get layer() As LayerNumbers
Const ProcName As String = "Layer"

On Error GoTo Err

layer = mSeries.layer
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyLayer) Then layer = mLocalStyle.layer
If IsPropertySet(TextPropertyLayer) Then layer = mLayer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"

On Error GoTo Err

If IsPropertySet(DataPointPropertyLineStyle) Then
    If getProperties.LineStyle = Value Then Exit Property
End If
getProperties.LineStyle = Value
setPropertyFlag DataPointPropertyLineStyle
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LineStyle() As LineStyles
Const ProcName As String = "LineStyle"

On Error GoTo Err

LineStyle = mSeries.LineStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyLineStyle) Then LineStyle = mLocalStyle.LineStyle
If IsPropertySet(DataPointPropertyLineStyle) Then LineStyle = getProperties.LineStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LineThickness(ByVal Value As Long)
Const ProcName As String = "LineThickness"

On Error GoTo Err

If IsPropertySet(DataPointPropertyLineThickness) Then
    If getProperties.LineThickness = Value Then Exit Property
End If
getProperties.LineThickness = Value
setPropertyFlag DataPointPropertyLineThickness
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LineThickness() As Long
Const ProcName As String = "LineThickness"

On Error GoTo Err

LineThickness = mSeries.LineThickness
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyLineThickness) Then LineThickness = mLocalStyle.LineThickness
If IsPropertySet(DataPointPropertyLineThickness) Then LineThickness = getProperties.LineThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LocalStyle() As DataPointStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal Value As DataPointStyle)

Const ProcName As String = "LocalStyle"

On Error GoTo Err

Set mLocalStyle = Value
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointStyle(ByVal Value As PointStyles)
Const ProcName As String = "PointStyle"

On Error GoTo Err

If IsPropertySet(DataPointPropertyPointStyle) Then
    If getProperties.PointStyle = Value Then Exit Property
End If
getProperties.PointStyle = Value
setPropertyFlag DataPointPropertyPointStyle
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointStyle() As PointStyles
Const ProcName As String = "PointStyle"

On Error GoTo Err

PointStyle = mSeries.PointStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyPointStyle) Then PointStyle = mLocalStyle.PointStyle
If IsPropertySet(DataPointPropertyPointStyle) Then PointStyle = getProperties.PointStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PreviousDataPoint() As DataPoint
Const ProcName As String = "PreviousDataPoint"

On Error GoTo Err

If mPrevDataPointRef Is Nothing Then Exit Property
Set PreviousDataPoint = prevDataPointObj

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let PreviousDataPoint(ByVal Value As DataPoint)
Const ProcName As String = "PreviousDataPoint"

On Error GoTo Err

Set mPrevDataPointRef = CreateWeakReference(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpColor(ByVal Value As Long)
Const ProcName As String = "UpColor"

On Error GoTo Err

If Not gIsValidColor(Value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color"
If IsPropertySet(DataPointPropertyUpColor) Then
    If getProperties.UpColor = Value Then Exit Property
End If
getProperties.UpColor = Value
setPropertyFlag DataPointPropertyUpColor
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"

On Error GoTo Err

UpColor = mSeries.UpColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyUpColor) Then UpColor = mLocalStyle.UpColor
If IsPropertySet(DataPointPropertyUpColor) Then UpColor = getProperties.UpColor
If UpColor = -1 Then UpColor = Color

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Public Property Let X(ByVal Value As Double)
Const ProcName As String = "X"

On Error GoTo Err

If Value = mX Then Exit Property
mBlank = False
mX = Value
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get X() As Double
X = mX
End Property

'================================================================================
' Methods
'================================================================================

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"

On Error GoTo Err

mPropertyOverrideFlags = 0
initiateRedraw True, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pSeries As DataPointSeries, _
                ByVal pTimestamp As Date, _
                ByVal pKey As String, _
                ByVal pGOSP As GraphObjServiceProvider, _
                ByVal pViewport As ViewPort, _
                ByVal pLocalStyle As DataPointStyle)
                
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mSeries = pSeries
mTimestamp = pTimestamp
mKey = pKey
Set mGOSP = pGOSP
Set mViewport = pViewport
Set mCanvas = mViewport.Canvas
Set mLocalStyle = pLocalStyle

calcBoundingRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcBoundingRect()

Const ProcName As String = "calcBoundingRect"

On Error GoTo Err

If DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    calcBoundingRectForHistogram
ElseIf drawAsPoint Then
    calcBoundingRectForPoint
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
    calcBoundingRectForLineOrSteppedLine
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
    calcBoundingRectForLineOrSteppedLine
End If


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcBoundingRectForHistogram()
Dim boundingRect As TRectangle
Dim pixelRect As GDI_RECT
    
Const ProcName As String = "calcBoundingRectForHistogram"

On Error GoTo Err

With mBoundingRect
    .Left = mX - HistogramBarWidth / 2
    .Right = mX + HistogramBarWidth / 2
    If mDataValue >= 0 Then
        .Bottom = 0
        .Top = mDataValue
    Else
        .Bottom = mDataValue
        .Top = 0
    End If
    .isValid = True
End With
'pixelRect = mViewport.ConvertTRectangleToRect(boundingRect)
'pixelRect.Right = pixelRect.Right + 1
'pixelRect.Bottom = pixelRect.Bottom + 1
'mBoundingRect = mViewport.ConvertRectToTRectangle(pixelRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcBoundingRectForLineOrSteppedLine()
Dim boundingRect As TRectangle
Dim pixelRect As GDI_RECT
Dim prevX As Double
Dim prevValue As Double
Dim lt As Long

Const ProcName As String = "calcBoundingRectForLineOrSteppedLine"

On Error GoTo Err

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

lt = LineThickness
With pixelRect
    .Left = .Left - Int(lt / 2)
    .Right = .Right + Int(lt / 2) + 1
    .Top = .Top - Int(lt / 2)
    .Bottom = .Bottom + Int(lt / 2) + 1
End With

mBoundingRect = mViewport.ConvertRectToTRectangle(pixelRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcBoundingRectForPoint()
Dim pixelRect As GDI_RECT

Const ProcName As String = "calcBoundingRectForPoint"

On Error GoTo Err

pixelRect = getPaintingRect
pixelRect.Right = pixelRect.Right + 1
pixelRect.Bottom = pixelRect.Bottom + 1

mBoundingRect = mViewport.ConvertRectToTRectangle(pixelRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ClearProperty( _
                ByVal flag As DataPointPropertyFlags)
Const ProcName As String = "ClearProperty"

On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Draw()

Const ProcName As String = "Draw"

On Error GoTo Err

If Not mPrevDataPointRef Is Nothing Then
    If mDataValue >= prevDataPointObj.DataValue Then
        DrawPoint UpColor
    Else
        DrawPoint DownColor
    End If
Else
    DrawPoint Color
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function drawAsPoint() As Boolean
Const ProcName As String = "drawAsPoint"

On Error GoTo Err

If mPrevDataPointRef Is Nothing Then
    drawAsPoint = True
ElseIf prevDataPointObj.Blank Then
    drawAsPoint = True
ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Then
    drawAsPoint = True
Else
    drawAsPoint = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub DrawPoint(ByVal Color As Long)
Dim PreviousDataPoint As DataPoint

Const ProcName As String = "DrawPoint"

On Error GoTo Err

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
        .Canvas.DrawRectangle mX - HistogramBarWidth / 2, 0, mX + HistogramBarWidth / 2, mDataValue
    ElseIf drawAsPoint Then
        If PointStyle = PointRound Then
            .Canvas.DrawPoint gNewPoint(mX, mDataValue)
        Else
            .Canvas.SetPenAttributes Color, _
                                    1, _
                                    LineInsideSolid, _
                                    DrawModes.DrawModeCopyPen
            .Canvas.SetBrushAttributes Color, FillSolid
            .Canvas.DrawRectangleFromRect getPaintingRect
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getPaintingRect() As GDI_RECT
Dim pixelRect As GDI_RECT
Dim lt As Long

Const ProcName As String = "getPaintingRect"

On Error GoTo Err

lt = LineThickness

With pixelRect
    .Left = mViewport.ConvertCoordsLogicalToPixelsX(mX)
    .Right = .Left
    .Bottom = mViewport.ConvertCoordsLogicalToPixelsY(mDataValue)
    .Top = .Bottom

    If PointStyle = PointRound Then
        If lt <> 1 Then
            ' these offsets were discovered by measuring points drawn using PSet
            .Left = .Left - Int(lt / 2)
            .Right = .Left + lt - 1
            .Top = .Top - Int((lt - 1) / 2)
            .Bottom = .Top + 2 * Int((lt + 1) / 2) - 1
        End If
    Else
        .Left = .Left - Int(lt / 2)
        .Right = .Left + lt - 1
        .Top = .Top - Int(lt / 2)
        .Bottom = .Top + lt - 1
    End If
    
End With

getPaintingRect = pixelRect

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getProperties() As DataPointProperties
Static props As DataPointProperties
Const ProcName As String = "getProperties"

On Error GoTo Err

If props Is Nothing Then Set props = New DataPointProperties
Set getProperties = props

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initiateRedraw( _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "initiateRedraw"

On Error GoTo Err

If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
If undrawCurrentImage Then
    undraw
End If

If recalcBoundaries Then
    calcBoundingRect
End If

mGOSP.ObjectChanged mHandle, mBoundingRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IsPropertySet( _
                ByVal flag As DataPointPropertyFlags) As Boolean
Const ProcName As String = "IsPropertySet"

On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function prevDataPointObj() As DataPoint
Const ProcName As String = "prevDataPointObj"

On Error GoTo Err

Set prevDataPointObj = mPrevDataPointRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setPropertyFlag( _
                ByVal flag As DataPointPropertyFlags)
Const ProcName As String = "setPropertyFlag"

On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub undraw()
Const ProcName As String = "undraw"

On Error GoTo Err

If Not mVisible Then Exit Sub
mGOSP.ObjectUndrawn mHandle, mBoundingRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


