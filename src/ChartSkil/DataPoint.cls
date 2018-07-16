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

Private mLineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mPointStyle As PointStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistogramBarWidth As Single
Private mIncludeInAutoscale As Boolean

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

Private mBoundingRect As TRectangle

Private mX As Double
Private mKey As String

Private mDataValue As Double
Private mBlank As Boolean   ' indicates that no Value has been assigned

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

'Private Sub Class_Terminate()
'gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Class_Terminate", pModName:=ModuleName, pMsg:="DataPoint terminated"
'Debug.Print "DataPoint terminated"
'End Sub

'================================================================================
' IGraphicObject Members
'================================================================================

Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Const ProcName As String = "IGraphicObject_boundingRectangle"
On Error GoTo Err

Static sPrevGaugeX As Double
Static sPrevGaugeY As Double
Static sPrevRect As TRectangle

Dim viewportChanges As Long
viewportChanges = mViewport.CompareTo(sPrevGaugeX, sPrevGaugeY, sPrevRect)

If (viewportChanges And ViewportComparisonCodes.GaugeChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.GaugeChangedY) Or _
    mSeriesUpdateNumber <> mSeries.SeriesUpdateNumber _
Then
    mSeriesUpdateNumber = mSeries.SeriesUpdateNumber
    calcBoundingRect
End If

sPrevGaugeX = mViewport.GaugeX
sPrevGaugeY = mViewport.GaugeY
sPrevRect = mViewport.Boundary

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

Private Property Get IGraphicObject_Layer() As Long
IGraphicObject_Layer = Layer
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

If Not gIsValidColor(Value) And Value <> -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color or -1"

If mColor = Value And IsPropertySet(DataPointPropertyColor) Then Exit Property
mColor = Value
If mColor <> -1 Then
    setPropertyFlag DataPointPropertyColor
Else
    ClearProperty DataPointPropertyColor
End If
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
If IsPropertySet(DataPointPropertyColor) Then Color = mColor

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
    If mDisplayMode = Value Then Exit Property
End If
setPropertyFlag DataPointPropertyDisplayMode
mDisplayMode = Value
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
If IsPropertySet(DataPointPropertyDisplayMode) Then DisplayMode = mDisplayMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownColor(ByVal Value As Long)
Const ProcName As String = "DownColor"
On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color or -1"

If mDownColor = Value And IsPropertySet(DataPointPropertyDownColor) Then Exit Property
mDownColor = Value
If mDownColor <> -1 Then
    setPropertyFlag DataPointPropertyDownColor
Else
    ClearProperty DataPointPropertyDownColor
End If
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
If IsPropertySet(DataPointPropertyDownColor) Then DownColor = mDownColor
If DownColor = -1 Then DownColor = Color

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Sub Finish()
Set mPrevDataPointRef = Nothing
End Sub

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Let HistogramBarWidth(ByVal Value As Single)
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

If Value <= 0 Or Value > 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "HistogramBarWidth must be greater than zero but not greater than 1"
If IsPropertySet(DataPointPropertyHistWidth) Then
    If mHistogramBarWidth = Value Then Exit Property
End If
setPropertyFlag DataPointPropertyHistWidth
mHistogramBarWidth = Value
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
If IsPropertySet(DataPointPropertyHistWidth) Then HistogramBarWidth = mHistogramBarWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

IncludeInAutoscale = mSeries.IncludeInAutoscale
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(DataPointPropertyIncludeInAutoscale) Then IncludeInAutoscale = mLocalStyle.IncludeInAutoscale
If IsPropertySet(DataPointPropertyIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

If IsPropertySet(DataPointPropertyIncludeInAutoscale) Then
    If mIncludeInAutoscale = Value Then Exit Property
End If
mIncludeInAutoscale = Value
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

Public Property Let Layer(ByVal Value As LayerNumbers)
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

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mSeries.Layer
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyLayer) Then Layer = mLocalStyle.Layer
If IsPropertySet(TextPropertyLayer) Then Layer = mLayer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"
On Error GoTo Err

If IsPropertySet(DataPointPropertyLineStyle) Then
    If mLineStyle = Value Then Exit Property
End If
mLineStyle = Value
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
If IsPropertySet(DataPointPropertyLineStyle) Then LineStyle = mLineStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LineThickness(ByVal Value As Long)
Const ProcName As String = "LineThickness"
On Error GoTo Err

If IsPropertySet(DataPointPropertyLineThickness) Then
    If mLineThickness = Value Then Exit Property
End If
mLineThickness = Value
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
If IsPropertySet(DataPointPropertyLineThickness) Then LineThickness = mLineThickness

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
    If mPointStyle = Value Then Exit Property
End If
mPointStyle = Value
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
If IsPropertySet(DataPointPropertyPointStyle) Then PointStyle = mPointStyle

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

If Not gIsValidColor(Value) And Value <> -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color or -1"

If mUpColor = Value And IsPropertySet(DataPointPropertyUpColor) Then Exit Property
mUpColor = Value
If mUpColor <> -1 Then
    setPropertyFlag DataPointPropertyUpColor
Else
    ClearProperty DataPointPropertyUpColor
End If
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
If IsPropertySet(DataPointPropertyUpColor) Then UpColor = mUpColor
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
Const ProcName As String = "calcBoundingRectForHistogram"
On Error GoTo Err

Dim boundingRect As TRectangle
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

'Dim pixelRect As GDI_RECT
'pixelRect = mViewport.ConvertTRectangleToRect(boundingRect)
'pixelRect.Right = pixelRect.Right + 1
'pixelRect.Bottom = pixelRect.Bottom + 1
'mBoundingRect = mViewport.ConvertRectToTRectangle(pixelRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcBoundingRectForLineOrSteppedLine()
Const ProcName As String = "calcBoundingRectForLineOrSteppedLine"
On Error GoTo Err

Dim prevX As Double
prevX = prevDataPointObj.X

Dim prevValue As Double
prevValue = prevDataPointObj.DataValue

Dim boundingRect As TRectangle
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

Dim pixelRect As GDI_RECT
pixelRect = mViewport.ConvertTRectangleToRect(boundingRect)

Dim lt As Long
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
Const ProcName As String = "calcBoundingRectForPoint"
On Error GoTo Err

Dim pixelRect As GDI_RECT
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
    If mDataValue > prevDataPointObj.DataValue Then
        DrawPoint UpColor
    ElseIf mDataValue < prevDataPointObj.DataValue Then
        DrawPoint DownColor
    Else
        DrawPoint Color
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
    
    Dim PreviousDataPoint As DataPoint
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
Const ProcName As String = "getPaintingRect"
On Error GoTo Err

Dim lt As Long
lt = LineThickness

Dim pixelRect As GDI_RECT
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


