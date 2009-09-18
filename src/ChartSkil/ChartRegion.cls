VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements DeferredAction

'================================================================================
' Events
'================================================================================

Event AutoscalingChanged()

Event Click()
Attribute Click.VB_UserMemId = -600

Event DblCLick()
Attribute DblCLick.VB_UserMemId = -601

Event MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseDown.VB_UserMemId = -605
                
Event MouseMove(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseMove.VB_UserMemId = -606
                
Event MouseUp(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseUp.VB_UserMemId = -607

'================================================================================
' Enums
'================================================================================

Private Enum DeferredPaintActions
    DeferredPaintRepaintNone
    DeferredPaintRepaintChangedAreas
    DeferredPaintRepaintRegion
    DeferredPaintChangeGaugeandRepaint
    DeferredPaintScaleAndRepaint
    DeferredPaintCalculateAndRepaint
End Enum

Private Enum PropertyOverrideFlags
    AutoscalingIsSet = 1
    GridTextStyleIsSet = 2
    HasGridIsSet = 4
    HasGridTextIsSet = 8
    IntegerYScaleIsSet = &H10&
    YScaleQuantumIsSet = &H20&
    GridlineSpacingYIsSet = &H40&
    MinimumHeightIsSet = &H80&
    CursorSnapsToTickBoundariesIsSet = &H100&
    BackGradientFillColorsIsSet = &H200&
    GridLineStyleIsSet = &H400&
    YCursorTextStyleIsSet = &H800&
    YAxisTextStyleIsSet = &H1000&
End Enum

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "ChartRegion"

'================================================================================
' Types
'================================================================================

Private Type RedrawObjectsTableEntry
    graphObj            As IGraphicObject
    areaToRedraw        As TRectangle
End Type

'================================================================================
' Member variables
'================================================================================

Private mChartRef As WeakReference
Private WithEvents mData As ChartRegionData
Attribute mData.VB_VarHelpID = -1
Private mGraphObjSeriesCollection As GraphObjSeriesCollection

Private mHandle As Long

Private mActualHeight As Long

Private mCleared As Boolean

Private mName As String

Private mRegionType As RegionTypes

Private WithEvents mStyle As ChartRegionStyle
Attribute mStyle.VB_VarHelpID = -1

Private mTitle As Text

Private WithEvents mViewport As Viewport
Attribute mViewport.VB_VarHelpID = -1
Private mIsDrawingEnabled As Boolean

Private mMinInViewPeriod As Long
Private mMaxInViewPeriod As Long

Private mDrawn As Boolean

Private mGraphObjServiceProvider As GraphObjServiceProvider
Attribute mGraphObjServiceProvider.VB_VarHelpID = -1

Private mHorizontalGrid As HorizontalGrid

Private mCrosshairLineHoriz As Line
Private mCrosshairLineVert As Line

Private mYAxisRegion As ChartRegion
Private mYCursorText As Text      ' used to indicate the Position of the pointer
                                    ' in the Y axis region
Private mXGridTextsCoarse As TextSeries
Private mXGridTexts As TextSeries
Private mXGridLines As LineSeries

Private mPointerMode As PointerModes
Private mPointerStyle As PointerStyles
Private mPointerIcon As IPictureDisp
Private mPointerToolStyle As PointerStyles
Private mPointerToolIcon As IPictureDisp

Private mSessionOffsetMinutes As Date
Private mVerticalGridTimePeriod As TimePeriod

Private mPercentHeight As Double
Private mMinimumPercentHeight As Double
Private mYAxisPosition As Long

Private mPerformanceTextVisible As Boolean

Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mCurrX As Single
Private mCurrY As Single

Private mUseDeferredPainting As Boolean
Private mDeferredPaintPending As Boolean

Private mDeferredPaintAction As DeferredPaintActions

Private mUndrawnAreas() As TRectangle
Private mUndrawnAreasIndex As Long

Private mRedrawObjects() As RedrawObjectsTableEntry
Private mRedrawObjectsIndex As Long

Private mSelectedObject As IGraphicObject

' these fields override corresponding items in the Style
Private mAutoscaling As Boolean
Private mHasGrid As Boolean
Private mHasGridText As Boolean
Private mIntegerYScale As Boolean
Private mYScaleQuantum As Double
Private mGridlineSpacingY As Double
Private mMinimumHeight As Double
Private mCursorSnapsToTickBoundaries As Boolean
Private mBackGradientFillColors() As Long
Private WithEvents mGridLineStyle As LineStyle
Attribute mGridLineStyle.VB_VarHelpID = -1
Private WithEvents mGridTextStyle As TextStyle
Attribute mGridTextStyle.VB_VarHelpID = -1
Private WithEvents mYAxisTextStyle As TextStyle
Attribute mYAxisTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

Private mPropertyOverrideFlags As PropertyOverrideFlags

Private mPaintingInProgress As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Dim i As Long

Set mData = New ChartRegionData
Set mGraphObjSeriesCollection = New GraphObjSeriesCollection
mGraphObjSeriesCollection.Initialise Me

Set mStyle = New ChartRegionStyle

mUseDeferredPainting = True
deferredPaint DeferredPaintRepaintNone

Set mGraphObjServiceProvider = New GraphObjServiceProvider
mGraphObjServiceProvider.Region = Me

End Sub

Private Sub Class_Terminate()
If Not mViewport Is Nothing Then mViewport.Finish
Set mViewport = Nothing

gLogger.Log LogLevelHighDetail, "ChartRegion terminated (" & gRegionTypeToString(mRegionType) & ")"
Debug.Print "ChartRegion terminated (" & gRegionTypeToString(mRegionType) & ")"
End Sub

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
Dim et As ElapsedTimer
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If mCleared Then Exit Sub

'If gIsInDev Or gLogger.isLoggable(LogLevelHighDetail) Then
If gIsInDev Then
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

Select Case mDeferredPaintAction
Case DeferredPaintRepaintNone

Case DeferredPaintRepaintChangedAreas
    

    failpoint = 100

    For i = 0 To mUndrawnAreasIndex
        failpoint = failpoint + i
        undrawAreaEx mUndrawnAreas(i)
    Next
    
    
    failpoint = 200
    
    For i = 0 To mRedrawObjectsIndex
        failpoint = failpoint + i
        redrawObject mRedrawObjects(i).graphObj, mRedrawObjects(i).areaToRedraw
    Next
Case DeferredPaintRepaintRegion
    
    failpoint = 300
    
    Paint
Case DeferredPaintChangeGaugeandRepaint
    
    failpoint = 400
    
    mData.CalculateInscopeGaugeDependentObjects
    
    failpoint = 500
    
    
    failpoint = 600
    
    Paint
    
    
    failpoint = 700
    
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.ResizedY
Case DeferredPaintScaleAndRepaint
    
    failpoint = 800
    
    calculateVerticalBounds
    
    failpoint = 900
    
    mData.DetermineVisibleObjects
    
    failpoint = 1000
    
    Paint
    
    failpoint = 1100
    
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.SetVerticalScaleInternal mViewport.Bottom, mViewport.Top
    End If
Case DeferredPaintCalculateAndRepaint
    
    failpoint = 1200
    
    mData.DetermineInScopeObjects
    
    failpoint = 1300
    
    If Autoscaling Then
        calculateVerticalBounds
    End If
    
    failpoint = 1400
    
    mData.DetermineVisibleObjects
    
    failpoint = 1500
    
    Paint
    
    failpoint = 1600
    
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.SetVerticalScaleInternal mViewport.Bottom, mViewport.Top
    End If
End Select


'If gIsInDev Or gLogger.isLoggable(LogLevelHighDetail) Then
If gIsInDev Then
    Dim s As String
    s = "Deferred paint: " & Format(et.ElapsedTimeMicroseconds, "0")
    Debug.Print s
    'gLogger.Log LogLevelHighDetail, s
End If

deferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "DeferredAction_run" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

deferredPaint DeferredPaintRepaintNone

End Sub

'================================================================================
' GraphObjServiceProvider Methods
'================================================================================

Friend Sub GOSP_AddGraphicObject( _
                ByVal graphObj As IGraphicObject)
Dim failpoint As Long
On Error GoTo Err

mData.AddGraphicObject graphObj

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GOSP_addGraphicObject" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_ObjectChanged( _
                ByVal handle As Long, _
                ByRef areaToRedraw As TRectangle)
Dim graphObj As IGraphicObject

Dim failpoint As Long
On Error GoTo Err

Set graphObj = mData.GetGraphicObject(handle)
If Not mData.UpdateGraphicObject(graphObj) Then Exit Sub

If Not mIsDrawingEnabled Then Exit Sub

If Not graphObj.BoundingRectangle.isValid Then Exit Sub
If graphObj.NoDraw Then Exit Sub

If mPaintingInProgress Then Exit Sub

areaToRedraw = RectIntersection(areaToRedraw, mViewport.Boundary)

' check if the vertical bounds have changed, and redraw if so
If Not Autoscaling Then
    redrawChangedObject graphObj, areaToRedraw
ElseIf Not graphObj.IncludeInAutoscale Then
    redrawChangedObject graphObj, areaToRedraw
ElseIf graphObj.BoundingRectangle.Top < mViewport.Boundary.Top And _
        graphObj.BoundingRectangle.Bottom > (mViewport.Boundary.Bottom) Then
    redrawChangedObject graphObj, areaToRedraw
ElseIf mUseDeferredPainting Then
    deferredPaint DeferredPaintScaleAndRepaint
Else
    calculateVerticalBounds
    mData.DetermineVisibleObjects
    Paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GOSP_objectChanged" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Friend Sub GOSP_ObjectUndrawn(ByVal handle As Long, undrawnArea As TRectangle)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

UndrawArea undrawnArea, mData.GetGraphicObject(handle)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GOSP_objectUndrawn" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_PaintRegion( _
                ByVal recalculate As Boolean)

Dim failpoint As Long
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

If mUseDeferredPainting Then
    If recalculate Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        deferredPaint DeferredPaintRepaintRegion
    End If
Else
    If recalculate Then
        mData.DetermineInScopeObjects
        If Autoscaling Then
            calculateVerticalBounds
        End If
        mData.DetermineVisibleObjects
    End If
    Paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GOSP_paintRegion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_RecalcBoundary( _
                ByVal SeriesID As Long)
Dim failpoint As Long
On Error GoTo Err

mData.RecalcBoundaryForObjectsInSeries SeriesID

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GOSP_recalcBoundary" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Friend Sub GOSP_RemoveObject( _
                ByVal handle As Long)
RemoveGraphicObjectEx mData.GetGraphicObject(handle), True, False
End Sub

'================================================================================
' mViewport Event Handlers
'================================================================================

Private Sub mViewport_BoundaryChanged()
mData.SetViewport mViewport.Boundary
End Sub

'================================================================================
' mData Event Handlers
'================================================================================

Private Sub mData_PeriodAdded( _
                ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date)
If CDbl(pTimestamp) <> 0 Then setVerticalGridLine pTimestamp, pPeriodNumber
End Sub

'================================================================================
' mGridLineStyle Event Handlers
'================================================================================

Private Sub mGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setGridLineStyle
End Sub

'================================================================================
' mGridTextStyle Event Handlers
'================================================================================

Private Sub mGridTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setGridTextStyle
End Sub

'================================================================================
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("Autoscaling")
    setAutoscaling
Case UCase$("BackColor")
    setBackColor
Case UCase$("BackGradientFillColors")
    setBackGradientFillColors
Case UCase$("GridlineSpacingY")
    setGridlineSpacingY
Case UCase$("GridLineStyle")
    setGridLineStyle
Case UCase$("GridTextStyle")
    setGridTextStyle
Case UCase$("HasGrid")
    setHasGrid
Case UCase$("HasGridText")
    setHasGridText
Case UCase$("IntegerYScale")
    setIntegerYScale
Case UCase$("MinimumHeight")
    setMinimumHeight
Case UCase$("CursorSnapsToTickBoundaries")
    setCursorSnapsToTickBoundaries
Case UCase$("YAxisTextStyle")
    setYAxisTextStyle
Case UCase$("YCursorTextStyle")
    setYCursorTextStyle
Case UCase$("YScaleQuantum")
    setYScaleQuantum
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged", _
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
' mYAxisTextStyle Event Handlers
'================================================================================

Private Sub mYAxisTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setYAxisTextStyle
End Sub

'================================================================================
' mYCursorTextStyle Event Handlers
'================================================================================

Private Sub mYCursorTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setYCursorTextStyle
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let ActualHeight(ByVal value As Long)
mActualHeight = value
End Property

Friend Property Get ActualHeight() As Long
ActualHeight = mActualHeight
End Property

Public Property Get Autoscaling() As Boolean
Autoscaling = mStyle.Autoscaling
If isPropertyOverrideFlagSet(AutoscalingIsSet) Then Autoscaling = mAutoscaling
End Property

Public Property Let Autoscaling(ByVal value As Boolean)
If mAutoscaling = value And isPropertyOverrideFlagSet(AutoscalingIsSet) Then Exit Property
mAutoscaling = value
setPropertyOverrideFlag AutoscalingIsSet
setAutoscaling
End Property

Public Property Get BackColor() As Long
BackColor = mStyle.BackColor
If isPropertyOverrideFlagSet(BackGradientFillColorsIsSet) Then BackColor = mBackGradientFillColors(0)
End Property

Public Property Let BackColor(ByVal value As Long)
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = value
setPropertyOverrideFlag AutoscalingIsSet
setBackColor
End Property

Public Property Get BackGradientFillColors() As Long()
BackGradientFillColors = mStyle.BackGradientFillColors
If isPropertyOverrideFlagSet(BackGradientFillColorsIsSet) Then BackGradientFillColors = mBackGradientFillColors
End Property

Public Property Let BackGradientFillColors(ByRef value() As Long)
Dim ar() As Long
ar = value
mBackGradientFillColors = ar
setPropertyOverrideFlag BackGradientFillColorsIsSet
setBackGradientFillColors
End Property

Public Property Get Bottom() As Double
Bottom = mViewport.Boundary.Bottom
End Property

Friend Property Let Bottom(ByVal value As Double)
mViewport.Bottom = value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.Bottom = value
End If
End Property

Friend Property Get Canvas() As Canvas
Set Canvas = mViewport
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal value As Boolean)
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "snapCursorToTickBoundaries", _
            "Cannot set snapCursorToTickBoundaries for an axis region"
End If
If mCursorSnapsToTickBoundaries = value And isPropertyOverrideFlagSet(CursorSnapsToTickBoundariesIsSet) Then Exit Property
mCursorSnapsToTickBoundaries = value
setPropertyOverrideFlag CursorSnapsToTickBoundariesIsSet
setCursorSnapsToTickBoundaries
End Property

Public Property Get CursorSnapsToTickBoundaries() As Boolean
CursorSnapsToTickBoundaries = mStyle.CursorSnapsToTickBoundaries
If isPropertyOverrideFlagSet(CursorSnapsToTickBoundariesIsSet) Then CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
End Property

Public Property Get GridlineSpacingY() As Double
GridlineSpacingY = mStyle.GridlineSpacingY
If isPropertyOverrideFlagSet(GridlineSpacingYIsSet) Then GridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let GridlineSpacingY(ByVal value As Double)
If mGridlineSpacingY = value And isPropertyOverrideFlagSet(GridlineSpacingYIsSet) Then Exit Property
mGridlineSpacingY = value
setPropertyOverrideFlag GridlineSpacingYIsSet
setGridlineSpacingY
End Property

Public Property Get GridLineStyle() As LineStyle
Set GridLineStyle = mStyle.GridLineStyle
If isPropertyOverrideFlagSet(GridLineStyleIsSet) Then Set GridLineStyle = mGridLineStyle
End Property

Public Property Let GridLineStyle(ByVal value As LineStyle)
If mGridLineStyle Is value And isPropertyOverrideFlagSet(GridLineStyleIsSet) Then Exit Property
Set mGridLineStyle = value
If mGridLineStyle Is Nothing Then
    clearPropertyOverrideFlag GridLineStyleIsSet
Else
    setPropertyOverrideFlag GridLineStyleIsSet
End If
setGridLineStyle
End Property

Public Property Get GridTextStyle() As TextStyle
Set GridTextStyle = mStyle.GridTextStyle
If isPropertyOverrideFlagSet(GridTextStyleIsSet) Then Set GridTextStyle = mGridTextStyle
End Property

Public Property Let GridTextStyle(ByVal value As TextStyle)
If mGridTextStyle Is value And isPropertyOverrideFlagSet(GridTextStyleIsSet) Then Exit Property
Set mGridTextStyle = value
If mGridTextStyle Is Nothing Then
    clearPropertyOverrideFlag GridTextStyleIsSet
Else
    setPropertyOverrideFlag GridTextStyleIsSet
End If
setGridTextStyle
End Property

Friend Property Let handle(ByVal value As Long)
mHandle = value
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get HasGrid() As Boolean
HasGrid = mStyle.HasGrid
If isPropertyOverrideFlagSet(HasGridIsSet) Then HasGrid = mHasGrid
End Property

Public Property Let HasGrid(ByVal value As Boolean)
If mHasGrid = value And isPropertyOverrideFlagSet(HasGridIsSet) Then Exit Property
mHasGrid = value
setPropertyOverrideFlag HasGridIsSet
setHasGrid
End Property

Public Property Get HasGridText() As Boolean
HasGridText = mStyle.HasGridText
If isPropertyOverrideFlagSet(HasGridTextIsSet) Then HasGridText = mHasGridText
End Property

Public Property Let HasGridText(ByVal value As Boolean)
If mHasGridText = value And isPropertyOverrideFlagSet(HasGridTextIsSet) Then Exit Property
mHasGridText = value
setPropertyOverrideFlag HasGridTextIsSet
setHasGridText
End Property

Public Property Get Height() As Double
Height = mViewport.Height
End Property

Public Property Let IntegerYScale(ByVal value As Boolean)
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "integerYScale", _
            "Cannot set integerYScale for an axis region"
End If

If mIntegerYScale = value And isPropertyOverrideFlagSet(IntegerYScaleIsSet) Then Exit Property
mIntegerYScale = value
setPropertyOverrideFlag IntegerYScaleIsSet
setIntegerYScale
End Property

Public Property Get IntegerYScale() As Boolean
IntegerYScale = mStyle.IntegerYScale
If isPropertyOverrideFlagSet(IntegerYScaleIsSet) Then IntegerYScale = mIntegerYScale
End Property

Friend Property Let IsDrawingEnabled(ByVal value As Boolean)
If mIsDrawingEnabled = value Then Exit Property
mIsDrawingEnabled = value
If Not mIsDrawingEnabled Then
    mViewport.MousePointer = vbHourglass
Else
    If IsXAxisRegion Or IsYAxisRegion Then
        mViewport.MousePointer = vbDefault
    Else
        Select Case mPointerMode
        Case PointerModeDefault
            setMouseIcon mPointerStyle
        Case PointerModeTool
            setMouseIcon PointerStyles.PointerTool
        Case PointerModeSelection
            setMouseIcon mPointerToolStyle
        End Select
        showCursor mPrevCursorX, mPrevCursorY
    End If
    
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        mData.DetermineInScopeObjects
        If Autoscaling Then
            calculateVerticalBounds
        End If
        mData.DetermineVisibleObjects
        Paint
    End If
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.IsDrawingEnabled = mIsDrawingEnabled
End If
End Property

Public Property Get IsXAxisRegion() As Boolean
IsXAxisRegion = (mRegionType = RegionTypeXAxis)
End Property

Public Property Get IsYAxisRegion() As Boolean
IsYAxisRegion = (mRegionType = RegionTypeYAxis)
End Property

Public Property Get Left() As Double
Left = mViewport.Left
End Property

Friend Property Let Left(ByVal value As Double)
mViewport.Left = value
End Property

Public Property Get MinimumHeight() As Double
MinimumHeight = mStyle.MinimumHeight
If isPropertyOverrideFlagSet(MinimumHeightIsSet) Then MinimumHeight = mMinimumHeight
End Property

Public Property Let MinimumHeight(ByVal value As Double)
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "minimumHeight", _
            "Cannot set minimumHeight for an axis region"
End If

If mMinimumHeight = value And isPropertyOverrideFlagSet(MinimumHeightIsSet) Then Exit Property
mMinimumHeight = value
setPropertyOverrideFlag MinimumHeightIsSet
setMinimumHeight
End Property

Public Property Get MinimumPercentHeight() As Double
MinimumPercentHeight = mMinimumPercentHeight
End Property

Friend Property Let MinimumPercentHeight(ByVal value As Double)
mMinimumPercentHeight = value
End Property

Public Property Get MousePosition() As Point
Set MousePosition = NewPoint(mCurrX, mCurrY)
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get PercentHeight() As Double
PercentHeight = mPercentHeight
End Property

Friend Property Let PercentHeight(ByVal value As Double)
mPercentHeight = value
End Property

Public Property Get PerformanceTextVisible() As Boolean
PerformanceTextVisible = mPerformanceTextVisible
End Property

Public Property Let PerformanceTextVisible(ByVal value As Boolean)
mPerformanceTextVisible = value
End Property

Public Property Get PointerCrosshairsColor() As Long
PointerCrosshairsColor = mPointerCrosshairsColor
End Property

Friend Property Let PointerCrosshairsColor(ByVal value As Long)
Dim lUseDeferredPainting As Boolean

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PointerCrosshairsColor", _
            "Cannot set PointerCrossHairsColor for an axis region"
End If

mPointerCrosshairsColor = value

lUseDeferredPainting = mUseDeferredPainting
mUseDeferredPainting = False
mCrosshairLineHoriz.Color = mPointerCrosshairsColor
mCrosshairLineVert.Color = mPointerCrosshairsColor
mUseDeferredPainting = lUseDeferredPainting
End Property

Public Property Get PointerDiscColor() As Long
PointerDiscColor = mPointerDiscColor
End Property

Friend Property Let PointerDiscColor(ByVal value As Long)
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PointerDiscColor", _
            "Cannot set PointerDiscColor for an axis region"
End If

mPointerDiscColor = value
End Property

Public Property Get PointerIcon() As IPictureDisp
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal value As IPictureDisp)

If value Is Nothing Then Exit Property
If value Is mPointerIcon Then Exit Property

If mPointerMode <> PointerModeDefault Then
    Set mPointerIcon = value
    Exit Property
End If

If Not mDrawn Then
    Set mPointerIcon = value
    Exit Property
End If

If mIsDrawingEnabled Then hideCursor

Set mPointerIcon = value

If mIsDrawingEnabled Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If

End Property

Public Property Get PointerStyle() As PointerStyles
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Dim failpoint As Long
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "pointerStyle", _
            "Cannot set PointerStyle for an axis region"
End If


If mPointerStyle = value Then Exit Property

If mPointerMode <> PointerModeDefault Then
    mPointerStyle = value
    Exit Property
End If

If Not mDrawn Then
    mPointerStyle = value
    Exit Property
End If

If mIsDrawingEnabled Then hideCursor

mPointerStyle = value

If mIsDrawingEnabled Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "PointerStyle" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get ScaleGridSpacingY() As Single
ScaleGridSpacingY = mHorizontalGrid.ScaleGridSpacingY
End Property

Friend Property Let SessionStartTime(ByVal value As Date)
mSessionOffsetMinutes = calculateMinutesFromMidnight(value) Mod 60
End Property

Public Property Get Style() As ChartRegionStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal value As ChartRegionStyle)

Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Style", _
            "Value must not be Nothing"
End If

Set mStyle = value

If mViewport Is Nothing Then Exit Property

processStyle

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get Title() As Text
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "Title", _
            "An axis region does not have a title"
End If

If mTitle Is Nothing Then setupTitle
Set Title = mTitle
End Property

Public Property Get Top() As Double
Top = mViewport.Top
End Property

Friend Property Let Top(ByVal value As Double)
mViewport.Top = value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.Top = value
End If
End Property

Friend Property Let VerticalGridTimePeriod(ByVal value As TimePeriod)
Set mVerticalGridTimePeriod = value
End Property

Public Property Get Width() As Single
Width = mViewport.Width
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Friend Property Let YAxisRegion(ByVal value As ChartRegion)

Set mYAxisRegion = value

mYAxisRegion.HasGrid = False
mYAxisRegion.HasGridText = False

mYAxisRegion.IsDrawingEnabled = mIsDrawingEnabled
mYAxisRegion.Bottom = 0
mYAxisRegion.Top = 1
mYAxisRegion.SetPeriodsInView 0, 99

Dim aTextSeries As New TextSeries
mYAxisRegion.AddGraphicObjectSeries aTextSeries, LayerNumbers.LayerPointer
Set mYCursorText = aTextSeries.Add
mYCursorText.Align = AlignBoxCentreLeft
setYCursorTextStyle

mHorizontalGrid.YAxisRegion = mYAxisRegion
setYAxisTextStyle

End Property

Public Property Get YAxisRegion() As ChartRegion
Set YAxisRegion = mYAxisRegion
End Property

Public Property Get YAxisTextStyle() As TextStyle
Set YAxisTextStyle = mStyle.YAxisTextStyle
If isPropertyOverrideFlagSet(YAxisTextStyleIsSet) Then Set YAxisTextStyle = mYAxisTextStyle
End Property

Public Property Let YAxisTextStyle(ByVal value As TextStyle)
If mYAxisTextStyle Is value And isPropertyOverrideFlagSet(YAxisTextStyleIsSet) Then Exit Property
Set mYAxisTextStyle = value
If mYAxisTextStyle Is Nothing Then
    clearPropertyOverrideFlag YAxisTextStyleIsSet
Else
    setPropertyOverrideFlag YAxisTextStyleIsSet
End If
setYAxisTextStyle
End Property

Public Property Get YCursorTextStyle() As TextStyle
Set YCursorTextStyle = mStyle.YCursorTextStyle
If isPropertyOverrideFlagSet(YCursorTextStyleIsSet) Then Set YCursorTextStyle = mYCursorTextStyle
End Property

Public Property Let YCursorTextStyle(ByVal value As TextStyle)
If mYCursorTextStyle Is value And isPropertyOverrideFlagSet(YCursorTextStyleIsSet) Then Exit Property
Set mYCursorTextStyle = value
If mYCursorTextStyle Is Nothing Then
    clearPropertyOverrideFlag YCursorTextStyleIsSet
Else
    setPropertyOverrideFlag YCursorTextStyleIsSet
End If
setYCursorTextStyle
End Property

Public Property Let YScaleQuantum(ByVal value As Double)
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "YScaleQuantum", _
            "Cannot set YScaleQuantum for an axis region"
End If

If mYScaleQuantum = value And isPropertyOverrideFlagSet(YScaleQuantumIsSet) Then Exit Property
mYScaleQuantum = value
setPropertyOverrideFlag YScaleQuantumIsSet
setYScaleQuantum
End Property

Public Property Get YScaleQuantum() As Double
YScaleQuantum = mStyle.YScaleQuantum
If isPropertyOverrideFlagSet(YScaleQuantumIsSet) Then YScaleQuantum = mYScaleQuantum
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddGraphicObjectSeries( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal Name As String) As IGraphicObjectSeries
Dim failpoint As Long
On Error GoTo Err

Set AddGraphicObjectSeries = mGraphObjSeriesCollection.Add( _
                                            graphObjSeries, _
                                            mData.GetNextUserLayer(Layer), _
                                            Name, _
                                            mViewport, _
                                            chartObj, _
                                            mGraphObjServiceProvider)
Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "addGraphicObjectSeries" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function AddLine( _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As LineStyle) As Line
Static miscLines As LineSeries
If miscLines Is Nothing Then
    Set miscLines = AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerHighestUser - 1)
End If
Set AddLine = miscLines.Add
If pLayer <> -1 Then AddLine.Layer = pLayer
If Not Style Is Nothing Then AddLine.LocalStyle = Style
End Function

Friend Sub AddPeriod(ByVal PeriodNumber As Long, _
                ByVal pTimestamp As Date)
Dim failpoint As Long
On Error GoTo Err

mData.AddPeriod PeriodNumber, pTimestamp

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "AddPeriod" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Function AddText( _
                Optional ByVal value As String, _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As TextStyle) As Text
Static miscTexts As TextSeries
If miscTexts Is Nothing Then
    Set miscTexts = AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerHighestUser)
End If
Set AddText = miscTexts.Add
If value <> "" Then AddText.Text = value
If pLayer <> -1 Then AddText.Layer = pLayer
If Not Style Is Nothing Then AddText.LocalStyle = Style
End Function

Friend Sub ClearRegion()

Dim failpoint As Long
On Error GoTo Err

If mCleared Then Exit Sub

Set mTitle = Nothing

Set mCrosshairLineHoriz = Nothing
Set mCrosshairLineVert = Nothing

If Not mYAxisRegion Is Nothing Then mYAxisRegion.ClearRegion
Set mYAxisRegion = Nothing

Set mYCursorText = Nothing

Set mHorizontalGrid = Nothing

Set mXGridTextsCoarse = Nothing
Set mXGridTexts = Nothing
Set mXGridLines = Nothing

mCleared = True

mGraphObjSeriesCollection.Finish

mGraphObjServiceProvider.Finish
Set mGraphObjServiceProvider = Nothing

Set mChartRef = Nothing

mData.Clear
Set mData = Nothing
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "clearRegion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ClearSelection()
deSelectObjectOrSeries mSelectedObject
Set mSelectedObject = Nothing
End Sub

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
doRepaint
End Sub

Friend Sub Click()
clickSelectedObjectOrSeries
RaiseEvent Click
End Sub

Friend Sub DblCLick()
dblClickSelectedObjectOrSeries
RaiseEvent DblCLick
End Sub

Friend Sub DrawCursor(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Dim yVal As Single
Dim failpoint As Long
On Error GoTo Err

If Y = MinusInfinitySingle Then
    yVal = mViewport.Top + 2 * mViewport.Height  ' make sure the pointer isn't visible
Else
    yVal = Y
End If

If Round(X) >= YAxisPosition Then Exit Sub

hideCursor
showCursor Round(X), yVal

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "drawCursor" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Function FormatYValue(ByVal Y As Double) As String
FormatYValue = mHorizontalGrid.GenerateYAxisCursorPositionText(Y)
End Function

Public Sub HideGrid()
Dim failpoint As Long
On Error GoTo Err

If mHorizontalGrid.Hidden Then Exit Sub

mHorizontalGrid.Hidden = True
If mUseDeferredPainting Then
    deferredPaint DeferredPaintRepaintRegion
Else
    Paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "hideGrid" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub Initialise( _
                ByVal Name As String, _
                ByVal pChart As Chart, _
                ByVal pCanvas As Canvas, _
                ByVal pRegionType As RegionTypes)
mName = Name

Set mChartRef = CreateWeakReference(pChart)
Set mViewport = New Viewport
mViewport.Canvas = pCanvas
mViewport.MousePointer = vbDefault

mRegionType = pRegionType

mData.Initialise pChart, Me, mGraphObjSeriesCollection

setupCrosshairs
processStyle
End Sub

Public Function IsGridHidden() As Boolean
IsGridHidden = mHorizontalGrid.Hidden
End Function

Friend Sub MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mCurrX = X
mCurrY = Y
If mPointerMode = PointerModeSelection Then setSelectedObjectOrSeries
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
                
Friend Sub MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
mCurrX = X
mCurrY = Y
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Friend Sub MouseUp(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mCurrX = X
mCurrY = Y
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Function NewDimension(ByVal X As Double, _
                        ByVal Y As Double) As Dimension
Dim failpoint As Long
On Error GoTo Err

Set NewDimension = mViewport.NewDimension(X, Y)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "NewDimension" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function NewPoint(ByVal X As Double, _
                        ByVal Y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordinateSystems.CoordsLogical) As Point
Dim failpoint As Long
On Error GoTo Err

Set NewPoint = mViewport.NewPoint(X, Y, coordSystemX, coordSystemY)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "NewPoint" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Friend Sub PaintRegion()
doRepaint
End Sub

Public Sub RemoveGraphicObject( _
                ByVal graphObj As IGraphicObject)
RemoveGraphicObjectEx graphObj, True, True
End Sub

Friend Sub RemoveGraphicObjectEx( _
                ByVal graphObj As IGraphicObject, _
                ByVal undraw As Boolean, _
                ByVal removeFromSeries As Boolean)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

If Not mSelectedObject Is Nothing Then
    If graphObj Is mSelectedObject Then ClearSelection
End If

failpoint = 300

If removeFromSeries Then mGraphObjSeriesCollection.Item(graphObj.SeriesID).RemoveObject graphObj

failpoint = 800

If undraw Then UndrawArea graphObj.BoundingRectangle, graphObj

mData.RemoveGraphicObject graphObj

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "RemoveGraphicObjectEx" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub RemoveGraphicObjectSeries( _
                ByVal pSeries As IGraphicObjectSeries)
If pSeries Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveGraphicObjectSeries", _
            "The supplied argument is Nothing"
End If

mGraphObjSeriesCollection.Remove pSeries
End Sub

Friend Sub ResizedY()
' notifies the region that the physical Height of the canvas has changed.
' This means the Y gauge will have changed, so we need to recalculate
' the Position of any in-scope gauge-dependent objects.
Dim failpoint As Long
On Error GoTo Err

mViewport.ResizedY

If mUseDeferredPainting Then
    deferredPaint DeferredPaintChangeGaugeandRepaint
Else
    mData.CalculateInscopeGaugeDependentObjects
    Paint
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.ResizedY
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ResizedY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ScaleUp(ByVal factor As Single)
Dim failpoint As Long
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "scaleUp", _
            "Cannot scale an axis region"
End If

With mViewport
    SetVerticalScale .Bottom + factor * .Height / 2, _
                                .Top - factor * .Height / 2
End With

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "scaleUp" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ScrollVertical(ByVal amount As Single)
Dim failpoint As Long
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "scrollVertical", _
            "Cannot scroll an axis region"
End If

Autoscaling = False
SetVerticalScale mViewport.Bottom + amount, mViewport.Top + amount

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "scrollVertical" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ScrollVerticalProportion(ByVal proportion As Single)
Dim failpoint As Long
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "scrollVerticalProportion", _
            "Cannot scroll an axis region"
End If

Autoscaling = False
SetVerticalScale mViewport.Bottom + proportion * mViewport.Height, _
                mViewport.Top + proportion * mViewport.Height

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "scrollVerticalProportion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub SetPeriodBackgroundColor( _
                ByVal PeriodNumber As Long, _
                ByVal Color As Long)
Dim Layer As Long
Dim index As Long
Dim periodrect As TRectangle
Dim redrawAreas(0) As TRectangle

Dim failpoint As Long
On Error GoTo Err

mData.PeriodBackColor(PeriodNumber) = Color

If Not mIsDrawingEnabled Then Exit Sub

If Color = 0 Then Exit Sub

With mViewport
    .Canvas.SetPenAttributes Color, 1, LineInsideSolid, DrawModeCopyPen
    .Canvas.SetBrushAttributes Color, FillSolid
    .Canvas.DrawRectangle .NewPoint(PeriodNumber - 0.5, .Bottom), .NewPoint(PeriodNumber + 0.5, .Top)

    periodrect.Left = PeriodNumber - 0.5
    periodrect.Right = PeriodNumber + 0.5
    periodrect.Bottom = .Bottom
    periodrect.Top = .Top
    periodrect.isValid = True
End With
    
mData.RedrawArea periodrect, LayerNumbers.LayerMin

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "setPeriodBackgroundColor" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub SetPeriodsInView( _
                ByVal minValue As Single, _
                ByVal maxValue As Single)
Dim failpoint As Long
On Error GoTo Err

mMinInViewPeriod = minValue
mMaxInViewPeriod = maxValue
mYAxisPosition = maxValue + 1
mViewport.Left = minValue
mViewport.Right = maxValue + 0.99

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        mData.DetermineInScopeObjects
        If Autoscaling Then
            calculateVerticalBounds
        End If
        mData.DetermineVisibleObjects
        Paint
        If Not mYAxisRegion Is Nothing Then mYAxisRegion.PaintRegion
    End If
End If
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetPeriodsInView" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub SetPointerModeDefault()
If mPointerMode = PointerModeDefault Then Exit Sub

If mIsDrawingEnabled Then
    ClearSelection
    hideCursor
End If

mPointerMode = PointerModeDefault

If mIsDrawingEnabled Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If
End Sub

Friend Sub SetPointerModeSelection()
If mPointerMode = PointerModeSelection Then Exit Sub

If mIsDrawingEnabled Then hideCursor

mPointerMode = PointerModeSelection

If mIsDrawingEnabled Then
    setMouseIcon PointerStyles.PointerSelection
    showCursor mPrevCursorX, mPrevCursorY
End If
End Sub

Friend Sub SetPointerModeTool( _
                ByVal toolPointerStyle As PointerStyles, _
                ByVal icon As IPictureDisp)

If mIsDrawingEnabled Then hideCursor

mPointerMode = PointerModeTool
mPointerToolStyle = toolPointerStyle
Set mPointerToolIcon = icon

If mIsDrawingEnabled Then
    setMouseIcon mPointerToolStyle
    showCursor mPrevCursorX, mPrevCursorY
End If
End Sub

' Note that the arguments are ByRef because they may be adjusted
' if they are the same when expressed as Singles.
Public Sub SetVerticalScale(ByRef lowValue As Double, _
                            ByRef highValue As Double)

Dim failpoint As Long
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "setVerticalScale", _
            "Cannot set the scale for an axis region"
End If

Autoscaling = False

SetVerticalScaleInternal lowValue, highValue

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "setVerticalScale" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Friend Sub SetVerticalScaleInternal(ByRef lowValue As Double, _
                            ByRef highValue As Double)

Dim failpoint As Long
On Error GoTo Err

setViewportVerticalBounds lowValue, highValue

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        mData.DetermineVisibleObjects
        Paint
    End If
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.SetVerticalScaleInternal lowValue, highValue
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "setVerticalScaleInternal" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub ShowGrid()
Dim failpoint As Long
On Error GoTo Err

If Not mHorizontalGrid.Hidden Then Exit Sub

mHorizontalGrid.Hidden = False
If mUseDeferredPainting Then
    deferredPaint DeferredPaintRepaintRegion
Else
    Paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ShowGrid" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addVerticalGridLine(ByVal PeriodNumber As Long, _
                                ByVal Key As String, _
                                ByVal label As String, _
                                ByVal coarseKey As String, _
                                ByVal coarseLabel As String)
Dim vGridLine As ChartSkil26.Line
Dim vGridText As Text
Dim vGridTextCoarse As Text

If HasGrid Then
    On Error Resume Next
    Set vGridLine = mXGridLines.Item(Key)
    On Error GoTo 0
    If vGridLine Is Nothing Then
        Set vGridLine = mXGridLines.Add(Key)
        vGridLine.SetPosition NewPoint(PeriodNumber - 0.5, _
                                    0, _
                                    CoordsLogical, _
                                    CoordsRelative), _
                                NewPoint(PeriodNumber - 0.5, _
                                    100, _
                                    CoordsLogical, _
                                    CoordsRelative)
    Else
        If vGridLine.Point1.XLogical > PeriodNumber Then
            vGridLine.SetPosition NewPoint(PeriodNumber - 0.5, _
                                        0, _
                                        CoordsLogical, _
                                        CoordsRelative), _
                                    NewPoint(PeriodNumber - 0.5, _
                                        100, _
                                        CoordsLogical, _
                                        CoordsRelative)
        End If
    End If
End If

If HasGridText Then
    On Error Resume Next
    Set vGridText = mXGridTexts.Item(Key)
    On Error GoTo 0
    If vGridText Is Nothing Then
        Set vGridText = mXGridTexts.Add(Key)
        vGridText.FixedX = False
        vGridText.FixedY = True
        vGridText.position = NewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsDistance)
        vGridText.Text = label
    End If

    On Error Resume Next
    Set vGridTextCoarse = mXGridTextsCoarse.Item(coarseKey)
    On Error GoTo 0
    If vGridTextCoarse Is Nothing Then
        Set vGridTextCoarse = mXGridTextsCoarse.Add(coarseKey)
        vGridTextCoarse.FixedX = False
        vGridTextCoarse.FixedY = True
        vGridTextCoarse.position = NewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsCounterDistance)
        vGridTextCoarse.Text = coarseLabel
    End If
End If

End Sub

Private Function calculateMinutesFromMidnight( _
                            ByVal Timestamp As Date) As Long
calculateMinutesFromMidnight = Int((Timestamp - Int(Timestamp) + 1 / 86400000) * 1440)
End Function

Private Sub calculateVerticalBounds()
Dim scaleMax As Double
Dim scaleMin As Double

Dim failpoint As Long
On Error GoTo Err

If Not Autoscaling Then Exit Sub

scaleMax = MinusInfinityDouble
scaleMin = PlusInfinityDouble

If mData.CalcVerticalBoundsForNonGaugeDependentObjects( _
                                scaleMin, _
                                scaleMax) _
Then
    ' there are gauge-dependent objects that may affect the max and min scales.
    ' We need to recalculate their bounds with the new gauge determined from
    ' the other objects to get a better estimate of their sizes.
    setViewportVerticalBounds scaleMin, scaleMax
    mData.CalcVerticalBoundsForGaugeDependentObjects _
                                scaleMin, _
                                scaleMax
End If

If scaleMin = PlusInfinityDouble Or scaleMax = MinusInfinityDouble Then Exit Sub

setViewportVerticalBounds scaleMin - (scaleMax - scaleMin) * 0.025, _
                        scaleMax + (scaleMax - scaleMin) * 0.025

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "calculateVerticalBounds" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource & vbCrLf & ".Bottom=" & CStr(scaleMin - (scaleMax - scaleMin) * 0.025)
Err.Raise errNumber, errSource, errDescription
End Sub

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Function clickSelectedObjectOrSeries() As Boolean
If mSelectedObject Is Nothing Then
    clickSelectedObjectOrSeries = False
ElseIf mSelectedObject.IsSelectable Then
    mSelectedObject.Click
    clickSelectedObjectOrSeries = True
ElseIf mGraphObjSeriesCollection.Item(mSelectedObject.SeriesID).IsSelectable Then
    mGraphObjSeriesCollection.Item(mSelectedObject.SeriesID).Click
    clickSelectedObjectOrSeries = True
Else
    clickSelectedObjectOrSeries = False
End If
End Function

Private Function dblClickSelectedObjectOrSeries() As Boolean
If mSelectedObject Is Nothing Then
    dblClickSelectedObjectOrSeries = False
ElseIf mSelectedObject.IsSelectable Then
    mSelectedObject.DblCLick
    dblClickSelectedObjectOrSeries = True
ElseIf mGraphObjSeriesCollection.Item(mSelectedObject.SeriesID).IsSelectable Then
    mGraphObjSeriesCollection.Item(mSelectedObject.SeriesID).DblCLick
    dblClickSelectedObjectOrSeries = True
Else
    dblClickSelectedObjectOrSeries = False
End If
End Function

Private Function deSelectObjectOrSeries(ByVal graphObj As IGraphicObject) As Boolean
If graphObj Is Nothing Then
ElseIf graphObj.IsSelectable Then
    graphObj.Selected = False
    deSelectObjectOrSeries = True
ElseIf mGraphObjSeriesCollection.Item(graphObj.SeriesID).IsSelectable Then
    mGraphObjSeriesCollection.Item(graphObj.SeriesID).Selected = False
    deSelectObjectOrSeries = True
Else
    deSelectObjectOrSeries = False
End If
End Function

Private Sub deferredPaint(ByVal action As DeferredPaintActions)
If action = DeferredPaintRepaintNone Then
    mDeferredPaintAction = DeferredPaintRepaintNone
    mDeferredPaintPending = False
    ReDim mUndrawnAreas(3) As TRectangle
    mUndrawnAreasIndex = -1
    ReDim mRedrawObjects(3) As RedrawObjectsTableEntry
    mRedrawObjectsIndex = -1
Else
    Select Case mDeferredPaintAction
    Case DeferredPaintRepaintNone
        mDeferredPaintAction = action
    Case DeferredPaintRepaintChangedAreas
        If action = DeferredPaintRepaintRegion Or _
            action = DeferredPaintChangeGaugeandRepaint Or _
            action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintRepaintRegion
        If action = DeferredPaintChangeGaugeandRepaint Or _
            action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintChangeGaugeandRepaint
        If action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintScaleAndRepaint
        If action = DeferredPaintCalculateAndRepaint Then mDeferredPaintAction = action
    Case DeferredPaintCalculateAndRepaint
        
    End Select
    If Not mDeferredPaintPending Then
        mDeferredPaintPending = True
        DeferAction Me, Empty, 20, ExpiryTimeUnitMilliseconds
    End If
End If
End Sub

Private Sub doRepaint()
If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintRepaintRegion
    Else
        Paint
    End If
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.PaintRegion
End If
End Sub

Private Sub doScaleAndRepaint()
If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintScaleAndRepaint
    Else
        calculateVerticalBounds
        mData.DetermineVisibleObjects
        Paint
    End If
End If
If Not mYAxisRegion Is Nothing Then mYAxisRegion.PaintRegion
End Sub

Private Sub drawCrosshairsPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single

If Not mIsDrawingEnabled Then Exit Sub

'    .SetPenAttributes mPointerCrosshairsColor Xor .BackColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine .NewPoint(.Left, Y), .NewPoint(YAxisPosition, Y)
lineX = Round(X)
'    .DrawLine .NewPoint(lineX, .Bottom), .NewPoint(lineX, .Top)
'
mCrosshairLineHoriz.SetPosition NewPoint(0, Y, CoordsRelative), _
                                NewPoint(100, Y, CoordsRelative)
mCrosshairLineVert.SetPosition NewPoint(lineX, 0, , CoordsRelative), _
                                NewPoint(lineX, 100, , CoordsRelative)
    
mPrevCursorX = X
mPrevCursorY = Y

If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = mHorizontalGrid.GenerateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawCustomPointer(ByVal X As Single, ByVal Y As Single)

If Not mIsDrawingEnabled Then Exit Sub

mPrevCursorX = X
mPrevCursorY = Y
    
If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = mHorizontalGrid.GenerateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawDiscPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim size As Dimension

If Not mIsDrawingEnabled Then Exit Sub

With mViewport
    '.MousePointer = vbCrosshair
    lineX = Round(X)
    Set size = .NewDimension(0.4, 0.4)
    
    .Canvas.SetPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.DrawLine lineX - size.XLogical, Y, lineX + size.XLogical, Y
    .Canvas.DrawLine lineX, Y - size.YLogical, lineX, Y + size.YLogical
    
    .Canvas.SetPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.SetBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .Canvas.DrawCircle .NewPoint(lineX, Y), .NewDimension(0.8, 0.8).XLogical
    
    mPrevCursorX = X
    mPrevCursorY = Y
    
End With

If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = mHorizontalGrid.GenerateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Dim drawAreas(0) As TRectangle
drawAreas(0) = RectIntersection(mViewport.Boundary, areaToRedraw)

graphObj.Draw drawAreas
End Sub

Private Sub drawSelectionPointer(ByVal X As Single, ByVal Y As Single)

If Not mIsDrawingEnabled Then Exit Sub

' nothing to draw since the pointer is a cursor and Windows draws it
mPrevCursorX = X
mPrevCursorY = Y
    
If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = mHorizontalGrid.GenerateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawToolPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim size As Dimension

If Not mIsDrawingEnabled Then Exit Sub

With mViewport
    lineX = Round(X)
    Set size = .NewDimension(0.4, 0.4)
    
'    .SetPenAttributes vbBlack Xor .BackColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine .NewPoint(lineX - size.XLogical, Y), .NewPoint(lineX + size.XLogical, Y)
'    .DrawLine .NewPoint(lineX, Y - size.YLogical), .NewPoint(lineX, Y + size.YLogical)
    
mCrosshairLineHoriz.SetPosition .NewPoint(lineX - size.XLogical, Y), .NewPoint(lineX + size.XLogical, Y)
mCrosshairLineVert.SetPosition .NewPoint(lineX, Y - size.YLogical), .NewPoint(lineX, Y + size.YLogical)
End With

mPrevCursorX = X
mPrevCursorY = Y

If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = mHorizontalGrid.GenerateYAxisCursorPositionText(Y)
End If
End Sub

Private Function getPerformanceText() As Text
Static perfText As Text
Dim perffont As StdFont

If perfText Is Nothing Then
    Set perffont = New StdFont
    perffont.Name = "Lucida Console"
    perffont.size = 8
    perffont.Underline = False
    perffont.Bold = False
    
    Set perfText = AddText()
    perfText.Color = vbActiveTitleBarText
    perfText.Font = perffont
    perfText.Box = True
    perfText.BoxColor = vbBlack
    perfText.BoxStyle = LineStyles.LineInsideSolid
    perfText.BoxThickness = 1
    perfText.BoxFillColor = vbActiveTitleBar
    perfText.BoxFillStyle = FillStyles.FillSolid
    perfText.position = NewPoint(5, 2, CoordsRelative, CoordsRelative)
    perfText.FixedX = True
    perfText.FixedY = True
    perfText.Align = TextAlignModes.AlignBottomLeft
    perfText.IncludeInAutoscale = False
    perfText.PaddingX = 0.5
    perfText.Layer = LayerTitle
End If
Set getPerformanceText = perfText
End Function

Private Sub hideCursor()
Dim lPointerStyle As PointerStyles

mUseDeferredPainting = False

If mPointerMode = PointerModeTool Then
    lPointerStyle = mPointerToolStyle
ElseIf mPointerMode = PointerModeSelection Then
    lPointerStyle = PointerSelection
Else
    lPointerStyle = mPointerStyle
End If

Select Case lPointerStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairsPointer
Case PointerDisc
    undrawDiscPointer
Case PointerSelection
    undrawSelectionPointer
Case PointerTool
    undrawToolPointer
Case PointerCustom
    undrawCustomPointer
End Select

mUseDeferredPainting = True
End Sub

Private Sub Paint()
Static et As ElapsedTimer

Dim numVisibleObjects  As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

Static repaintCount As Long
repaintCount = repaintCount + 1
Debug.Print "Region " & mHandle & " repaint count = " & repaintCount

mPaintingInProgress = True

If mPerformanceTextVisible Then
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

mViewport.PaintBackground

paintPeriodBackgrounds

If Not (IsYAxisRegion Or IsXAxisRegion) Then mHorizontalGrid.Generate mViewport.Bottom, mViewport.Top, mViewport.HeightCm

numVisibleObjects = mData.PaintAllVisibleObjects

showCursor mPrevCursorX, mPrevCursorY

If mPerformanceTextVisible Then paintPerformanceText numVisibleObjects, et.ElapsedTimeMicroseconds / 1000

mViewport.ZOrder 0

mDrawn = True
mPaintingInProgress = False

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "paint" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription

End Sub

Private Sub paintPerformanceText( _
                ByVal numberOfObjectsRepainted As Long, _
                ByVal elapsedTimeMillisecs As Single)
Dim perfText As Text
If mPerformanceTextVisible Then
    Set perfText = getPerformanceText
    perfText.Text = "Repaint time: " & Format(elapsedTimeMillisecs, "0.0") & "ms" & vbCrLf & _
                    "Visible objects: " & numberOfObjectsRepainted & vbCrLf & _
                    "Total objects: " & mData.NumberOfObjects
    ' need to force the Text to draw now
    perfText.Draw
End If
End Sub

Private Sub paintPeriodBackgrounds()
Dim i As Long
With mViewport
    For i = IIf(mMinInViewPeriod < 1, 1, mMinInViewPeriod) To IIf(mData.MaxPeriodNumber < mMaxInViewPeriod, mData.MaxPeriodNumber, mMaxInViewPeriod)
        If mData.PeriodBackColor(i) <> 0 Then
            .Canvas.SetPenAttributes mData.PeriodBackColor(i), 1, LineInsideSolid, DrawModeCopyPen
            .Canvas.SetBrushAttributes mData.PeriodBackColor(i), FillSolid
            .Canvas.DrawRectangle .NewPoint(i - 0.5, .Bottom), .NewPoint(i + 0.5, .Top)
        End If
    Next
End With
End Sub

Private Sub processStyle()

setBackColor
setBackGradientFillColors
setHasGrid
setHasGridText
setGridTextStyle
If Not IsXAxisRegion And Not IsYAxisRegion Then
    setAutoscaling
    setGridlineSpacingY
    setIntegerYScale
    setMinimumHeight
    setCursorSnapsToTickBoundaries
    setYAxisTextStyle
    setYCursorTextStyle
    setYScaleQuantum
End If

End Sub

Private Sub RedrawArea( _
                ByRef areaToRedraw As TRectangle, _
                ByVal startLayer As Long, _
                Optional ByVal objToExclude As IGraphicObject)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

' redraw objects in the undrawn area, preserving layering
If mIsDrawingEnabled And areaToRedraw.isValid Then
    
    failpoint = 100

    If RectOverlaps(mViewport.Boundary, areaToRedraw) Then
        mData.RedrawArea RectIntersection(mViewport.Boundary, areaToRedraw), startLayer, objToExclude
    End If
    
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "redrawArea" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub redrawChangedObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Dim i As Long
Dim handled As Boolean

If areaToRedraw.isValid Then
    If mUseDeferredPainting Then
        For i = 0 To mRedrawObjectsIndex
            If mRedrawObjects(i).graphObj Is graphObj Then
                mRedrawObjects(i).areaToRedraw = RectUnion(areaToRedraw, mRedrawObjects(i).areaToRedraw)
                handled = True
            End If
        Next
            
        If Not handled Then
            mRedrawObjectsIndex = mRedrawObjectsIndex + 1
            If mRedrawObjectsIndex > UBound(mRedrawObjects) Then
                ReDim Preserve mRedrawObjects(2 * (UBound(mRedrawObjects) + 1) - 1) As RedrawObjectsTableEntry
            End If
            Set mRedrawObjects(mRedrawObjectsIndex).graphObj = graphObj
            mRedrawObjects(mRedrawObjectsIndex).areaToRedraw = areaToRedraw
        End If
        deferredPaint DeferredPaintRepaintChangedAreas
    Else
        redrawObject graphObj, areaToRedraw
    End If
End If

End Sub

Private Sub redrawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
'unDrawCursor areaToRedraw
drawObject graphObj, areaToRedraw
RedrawArea areaToRedraw, graphObj.Layer + 1
End Sub

Private Function selectObjectOrSeries(ByVal graphObj As IGraphicObject) As Boolean
If graphObj.IsSelectable Then
    graphObj.Selected = True
    selectObjectOrSeries = True
ElseIf mGraphObjSeriesCollection.Item(graphObj.SeriesID).IsSelectable Then
    mGraphObjSeriesCollection.Item(graphObj.SeriesID).Selected = True
    selectObjectOrSeries = True
Else
    selectObjectOrSeries = False
End If
End Function

Private Sub setAutoscaling()
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "autoScale", _
            "Cannot set autoscale for an axis region"
End If

If Autoscaling Then doScaleAndRepaint

RaiseEvent AutoscalingChanged
End Sub

Private Sub setBackColor()
mViewport.BackColor = BackColor
doRepaint
End Sub

Private Sub setBackGradientFillColors()
mViewport.GradientFillColors = BackGradientFillColors
doRepaint
End Sub

Private Sub setViewportVerticalBounds( _
                ByRef Bottom As Double, _
                ByRef Top As Double)
    
If Bottom > Top Then Exit Sub
    
With mViewport
    If MinimumHeight <> 0 Then
        If Top - Bottom < MinimumHeight Then
            Top = (Top + Bottom) / 2 + MinimumHeight / 2
            Bottom = Top - MinimumHeight
        End If
    ElseIf Bottom = Top Then
        Top = Top * 1.001
        Bottom = Top * 0.998
    End If
    .Top = Top
    .Bottom = Bottom
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.Top = .Top
        mYAxisRegion.Bottom = .Bottom
    End If
    ' return possibly adjusted values
    Top = .Top
    Bottom = .Bottom
End With
End Sub

Private Sub setGridlineSpacingY()
If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "gridlineSpacingY", _
            "Cannot set gridline spacing for an axis region"
End If

mHorizontalGrid.GridlineSpacingY = GridlineSpacingY
doRepaint
End Sub

Private Sub setGridLineStyle()
setHasGrid
End Sub

Private Sub setGridTextStyle()
setHasGridText
End Sub

Private Sub setHasGrid()
If HasGrid Then
    Set mXGridLines = Me.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
    mXGridLines.Style = GridLineStyle
    
    Set mHorizontalGrid = New HorizontalGrid
    mHorizontalGrid.Initialise Me, GridLineStyle
Else
    Set mXGridLines = Nothing
    Set mHorizontalGrid = Nothing
End If
doRepaint
End Sub

Private Sub setHasGridText()

If HasGridText Then
    Set mXGridTexts = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
    mXGridTexts.Style = GridTextStyle
    ' override the Style for the following, but don't change the Style itself
    
    mXGridTexts.Align = AlignBoxBottomCentre
    If Not IsXAxisRegion Then
        mXGridTexts.Box = True
        mXGridTexts.BoxFillWithBackgroundColor = True
        mXGridTexts.BoxStyle = LineInvisible
    End If
    mXGridTexts.Extended = True

    Set mXGridTextsCoarse = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
    mXGridTextsCoarse.Style = GridTextStyle
    
    ' override the Style for the following, but don't change the Style itself
    mXGridTextsCoarse.Align = AlignBoxTopCentre
    If Not IsXAxisRegion Then
        mXGridTextsCoarse.Box = True
        mXGridTexts.BoxFillWithBackgroundColor = True
        mXGridTextsCoarse.BoxStyle = LineInvisible
    End If
    mXGridTextsCoarse.Extended = True
Else
    Set mXGridTexts = Nothing
    Set mXGridTextsCoarse = Nothing
End If
doRepaint
End Sub

Private Sub setIntegerYScale()
mHorizontalGrid.IntegerYScale = IntegerYScale
doRepaint
End Sub

Private Sub setMinimumHeight()
doScaleAndRepaint
End Sub

Private Sub setMouseIcon( _
                ByVal PointerStyle As PointerStyles)
mViewport.MousePointer = vbCustom
Select Case PointerStyle
Case PointerNone
    mViewport.MouseIcon = gBlankCursor
Case PointerCrosshairs
    mViewport.MouseIcon = gBlankCursor
Case PointerDisc
    mViewport.MouseIcon = gBlankCursor
Case PointerSelection
    mViewport.MouseIcon = gSelectorCursor
Case PointerTool
    mViewport.MouseIcon = gBlankCursor
Case PointerCustom
    mViewport.MouseIcon = mPointerIcon
Case PointerWindowsDefault
    mViewport.MousePointer = vbDefault
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26.ChartRegion::(let)pointerStyle", _
            "Value must be a member of the PointerStyles enum"
End Select
End Sub

Private Sub setSelectedObjectOrSeries()
Dim newSelectedObject As IGraphicObject

Set newSelectedObject = mData.GetNextSelectableObjectOrSeriesAtPoint(mSelectedObject, mCurrX, mCurrY)

If newSelectedObject Is Nothing Then
    Set mSelectedObject = Nothing
ElseIf newSelectedObject = mSelectedObject Then
    
Else
    deSelectObjectOrSeries mSelectedObject
    Set mSelectedObject = newSelectedObject
    selectObjectOrSeries newSelectedObject
End If
End Sub

Private Sub setCursorSnapsToTickBoundaries()
' nothing to do
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub setupCrosshairs()
Set mCrosshairLineHoriz = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineHoriz.Color = vbRed
mCrosshairLineHoriz.Extended = True
'mCrosshairLineHoriz.FixedX = True
'mCrosshairLineHoriz.FixedY = True
mCrosshairLineHoriz.IncludeInAutoscale = False
mCrosshairLineHoriz.LineStyle = LineSolid
mCrosshairLineHoriz.Thickness = 1

Set mCrosshairLineVert = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineVert.Color = vbRed
mCrosshairLineVert.Extended = True
'mCrosshairLineVert.FixedX = True
'mCrosshairLineVert.FixedY = True
mCrosshairLineVert.IncludeInAutoscale = False
mCrosshairLineVert.LineStyle = LineSolid
mCrosshairLineVert.Thickness = 1
End Sub

Private Sub setupTitle()
Dim failpoint As Long
On Error GoTo Err

Set mTitle = AddText(, LayerTitle)
mTitle.Box = True
mTitle.BoxColor = vbBlack
mTitle.BoxStyle = LineStyles.LineInvisible
mTitle.BoxThickness = 1
mTitle.BoxFillColor = vbWhite
mTitle.BoxFillStyle = FillStyles.FillSolid
mTitle.position = NewPoint(0.2, 0.2, CoordsDistance, CoordsCounterDistance)
mTitle.FixedX = True
mTitle.FixedY = True
mTitle.Align = TextAlignModes.AlignBoxTopLeft
mTitle.IncludeInAutoscale = False
mTitle.PaddingX = 0.5

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "setupTitle" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource


End Sub

Private Sub setVerticalGridLine(ByVal Timestamp As Date, _
                            ByVal PeriodNumber As Long)
Dim gridLineTime As Date

If CDbl(Timestamp) = 0 Then Exit Sub

If mVerticalGridTimePeriod Is Nothing Then
    ' just draw a vertical gridline every 10 bars
    If PeriodNumber Mod 10 = 1 Then
        addVerticalGridLine PeriodNumber, _
                    Format(Timestamp, "yyyymmddhhnnss"), _
                    FormatDateTime(Timestamp, vbShortTime), _
                    Format(Timestamp, "yyyymmdd"), _
                    Format(Timestamp, "d Mmm yy")
    End If
    Exit Sub
End If

gridLineTime = BarStartTime(Timestamp, mVerticalGridTimePeriod)
    
Select Case mVerticalGridTimePeriod.Units
Case TimePeriodSecond
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmddhhnnss"), _
                    FormatDateTime(gridLineTime, vbLongTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodMinute
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmddhhnn"), _
                    FormatDateTime(gridLineTime, vbShortTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodHour
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmddhh"), _
                    FormatDateTime(gridLineTime, vbShortTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodDay
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d"), _
                    Format(gridLineTime, "yyyymm"), _
                    Format(gridLineTime, "Mmm yy")
Case TimePeriodWeek
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm"), _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "yyyy")
Case TimePeriodMonth
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymm"), _
                    Format(gridLineTime, "Mmm"), _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "yyyy")
Case TimePeriodYear
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "YYYY"), _
                    "$", _
                    ""
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If PeriodNumber Mod 10 = 1 Then
        addVerticalGridLine PeriodNumber, _
                        Format(gridLineTime, "yyyymmddhhnnss"), _
                        FormatDateTime(gridLineTime, vbLongTime), _
                        Format(gridLineTime, "yyyymmdd"), _
                        Format(gridLineTime, "d Mmm yy")
    End If
End Select
End Sub

Private Sub setYAxisTextStyle()
mHorizontalGrid.TextStyle = YAxisTextStyle
End Sub

Private Sub setYCursorTextStyle()
If Not mYCursorText Is Nothing Then mYCursorText.LocalStyle = YCursorTextStyle
End Sub

Private Sub setYScaleQuantum()
mHorizontalGrid.YScaleQuantum = YScaleQuantum
doRepaint
End Sub

Private Sub showCursor( _
                ByVal X As Single, _
                ByVal Y As Single)
Dim PointerStyle As PointerStyles

mUseDeferredPainting = False

If mPointerMode = PointerModeTool Then
    PointerStyle = mPointerToolStyle
ElseIf mPointerMode = PointerModeSelection Then
    PointerStyle = PointerSelection
Else
    PointerStyle = mPointerStyle
End If

Select Case PointerStyle
Case PointerNone

Case PointerCrosshairs
    drawCrosshairsPointer X, Y
Case PointerDisc
    drawDiscPointer X, Y
Case PointerSelection
    drawSelectionPointer X, Y
Case PointerTool
    drawToolPointer X, Y
Case PointerCustom
    drawCustomPointer X, Y
End Select

mUseDeferredPainting = True
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As PropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub UndrawArea( _
                ByRef undrawnArea As TRectangle, _
                ByVal objToExclude As IGraphicObject)
If mPaintingInProgress Then Exit Sub

If mIsDrawingEnabled And undrawnArea.isValid Then
    If mUseDeferredPainting Then
        Dim i As Long
        Dim handled As Boolean
        For i = 0 To mUndrawnAreasIndex
            If RectOverlaps(undrawnArea, mUndrawnAreas(i)) Then
                 mUndrawnAreas(i) = RectUnion(undrawnArea, mUndrawnAreas(i))
                 handled = True
            End If
        Next
        If Not handled Then
            mUndrawnAreasIndex = mUndrawnAreasIndex + 1
            If mUndrawnAreasIndex > UBound(mUndrawnAreas) Then
                ReDim Preserve mUndrawnAreas(2 * (UBound(mUndrawnAreas) + 1) - 1) As TRectangle
            End If
            mUndrawnAreas(mUndrawnAreasIndex) = undrawnArea
        End If
        deferredPaint DeferredPaintRepaintChangedAreas
    Else
        undrawAreaEx undrawnArea, objToExclude
    End If
End If

End Sub

Private Sub undrawAreaEx( _
                ByRef undrawnArea As TRectangle, _
                Optional ByVal objToExclude As IGraphicObject)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

' redraw objects in the undrawn area, preserving layering
If mIsDrawingEnabled And undrawnArea.isValid Then
    
    failpoint = 200

    mViewport.UndrawArea undrawnArea
    
    failpoint = 400
    
    RedrawArea undrawnArea, MinLayer, objToExclude
    
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "undrawAreaEx" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub undrawCrosshairsPointer()

If Not mIsDrawingEnabled Then Exit Sub

mCrosshairLineHoriz.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)
mCrosshairLineVert.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)

End Sub

Private Sub undrawCustomPointer()

' nothing to do

End Sub

Private Sub undrawDiscPointer()
Dim size As Dimension
Dim prevLineX As Single

If Not mIsDrawingEnabled Then Exit Sub

With mViewport
    prevLineX = Round(mPrevCursorX)
    Set size = .NewDimension(0.4, 0.4)
    
    .Canvas.SetPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.DrawLine prevLineX - size.XLogical, mPrevCursorY, prevLineX + size.XLogical, mPrevCursorY
    .Canvas.DrawLine prevLineX, mPrevCursorY - size.YLogical, prevLineX, mPrevCursorY + size.YLogical
    
    .Canvas.SetPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.SetBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .Canvas.DrawCircle .NewPoint(prevLineX, mPrevCursorY), .NewDimension(0.8, 0.8).XLogical
End With

End Sub

Private Sub undrawSelectionPointer()

If Not mIsDrawingEnabled Then Exit Sub

' nothing to do since the pointer is actually a proper cursor
End Sub

Private Sub undrawToolPointer()

If Not mIsDrawingEnabled Then Exit Sub

mCrosshairLineHoriz.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)
mCrosshairLineVert.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)
End Sub


