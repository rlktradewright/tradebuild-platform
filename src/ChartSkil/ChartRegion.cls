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

Private Enum PropertyOverrideFlags
    AutoscalingIsSet = 1
    GridTextStyleIsSet = 2
    HasGridIsSet = 4
    HasGridTextProperty = 8
    IntegerYScaleIsSet = &H10&
    YScaleQuantumIsSet = &H20&
    GridlineSpacingYIsSet = &H40&
    MinimumHeightIsSet = &H80&
    CursorSnapsToTickBoundariesIsSet = &H100&
    BackGradientFillColorsIsSet = &H200&
    GridLineStyleIsSet = &H400&
    YCursorTextStyleIsSet = &H800&
    YAxisTextStyleIsSet = &H1000&
    SessionEndGridLineStyleIsSet = &H2000&
    SessionStartGridLineStyleIsSet = &H4000&
End Enum

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "ChartRegion"

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mChartRef As WeakReference
Private mData As ChartRegionData
Attribute mData.VB_VarHelpID = -1
Private mGraphObjSeriesCollection As GraphObjSeriesCollection

Private mHandle As Long

Private mActualHeight As Long

Private mIsCleared As Boolean

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

Private mVerticalGridTimePeriod As TimePeriod

Private mSessionStartTime As Date

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
Private mDeferredPainter As DeferredRegionPainter

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
Private mBackGradientFillColors()                   As Long

Private WithEvents mGridLineStyle                   As LineStyle
Attribute mGridLineStyle.VB_VarHelpID = -1
Private WithEvents mGridTextStyle                   As TextStyle
Attribute mGridTextStyle.VB_VarHelpID = -1
Private WithEvents mSessionEndGridLineStyle         As LineStyle
Attribute mSessionEndGridLineStyle.VB_VarHelpID = -1
Private WithEvents mSessionStartGridLineStyle       As LineStyle
Attribute mSessionStartGridLineStyle.VB_VarHelpID = -1
Private WithEvents mYAxisTextStyle                  As TextStyle
Attribute mYAxisTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle                As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

Private mPropertyOverrideFlags                      As PropertyOverrideFlags

Private mPaintingInProgress                         As Boolean

Private mPriceFormatter                             As IPriceFormatter

Private mGridHidden                                 As Boolean

Private mHorizontalGridSpacer                       As HorizontalGridSpacer

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
Set mDeferredPainter = New DeferredRegionPainter
mDeferredPainter.Initialise Me

Set mGraphObjServiceProvider = New GraphObjServiceProvider
mGraphObjServiceProvider.Region = Me

Set mPriceFormatter = New PriceFormatter

Set mHorizontalGridSpacer = New HorizontalGridSpacer

End Sub

Private Sub Class_Terminate()
If Not mViewport Is Nothing Then mViewport.Finish
Set mViewport = Nothing

'gLogger.Log LogLevelHighDetail, "ChartRegion terminated (" & gRegionTypeToString(mRegionType) & ")"
Debug.Print "ChartRegion terminated (" & gRegionTypeToString(mRegionType) & ")"
End Sub

'================================================================================
' GraphObjServiceProvider Methods
'================================================================================

Friend Sub GOSP_AddGraphicObject( _
                ByVal graphObj As IGraphicObject)
Const ProcName As String = "GOSP_AddGraphicObject"
Dim failpoint As String
On Error GoTo Err

mData.AddGraphicObject graphObj

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub GOSP_ObjectChanged( _
                ByVal handle As Long, _
                ByRef areaToRedraw As TRectangle)
Dim graphObj As IGraphicObject
Dim RedrawArea As TRectangle

Const ProcName As String = "GOSP_ObjectChanged"
Dim failpoint As String
On Error GoTo Err

Set graphObj = mData.GetGraphicObject(handle)
If Not mData.UpdateGraphicObject(graphObj) Then Exit Sub

If Not mIsDrawingEnabled Then Exit Sub

If Not graphObj.BoundingRectangle.isValid Then Exit Sub
If graphObj.NoDraw Then Exit Sub

If mPaintingInProgress Then Exit Sub

RedrawArea = RectIntersection(areaToRedraw, mViewport.Boundary)

' check if the vertical bounds have changed, and redraw if so
If Autoscaling And graphObj.IncludeInAutoscale Then
    Dim verticalRangeInViewport As TInterval
    verticalRangeInViewport = graphObj.VerticalRangeInXinterval(RectGetXInterval(mViewport.Boundary))
    
    If verticalRangeInViewport.isValid And _
        Autoscaling And _
        (verticalRangeInViewport.endValue > mViewport.Boundary.Top Or _
            verticalRangeInViewport.startValue < (mViewport.Boundary.Bottom)) Then
        doScaleAndRepaint
        Exit Sub
    End If
End If

redrawChangedObject graphObj, RedrawArea

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub GOSP_ObjectUndrawn(ByVal handle As Long, undrawnArea As TRectangle)
Const ProcName As String = "GOSP_ObjectUndrawn"
Dim failpoint As String
On Error GoTo Err

UndrawArea undrawnArea, mData.GetGraphicObject(handle)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub GOSP_PaintRegion( _
                ByVal recalculate As Boolean)
Const ProcName As String = "GOSP_PaintRegion"
Dim failpoint As String
On Error GoTo Err

If recalculate Then
    doCalculateAndRepaint
Else
    doRepaint
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub GOSP_RecalcBoundary( _
                ByVal SeriesID As Long)
Const ProcName As String = "GOSP_RecalcBoundary"
Dim failpoint As String
On Error GoTo Err

mData.RecalcBoundaryForObjectsInSeries SeriesID

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub GOSP_RemoveObject( _
                ByVal handle As Long)
Const ProcName As String = "GOSP_RemoveObject"
Dim failpoint As String
On Error GoTo Err

RemoveGraphicObjectEx mData.GetGraphicObject(handle), True, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' mViewport Event Handlers
'================================================================================

Private Sub mViewport_BoundaryChanged()
Const ProcName As String = "mViewport_BoundaryChanged"
Dim failpoint As String
On Error GoTo Err

mData.SetViewport mViewport.Boundary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mGridLineStyle Event Handlers
'================================================================================

Private Sub mGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mGridLineStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

setGridLineStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mGridTextStyle Event Handlers
'================================================================================

Private Sub mGridTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mGridTextStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

setGridTextStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mSessionEndGridLineStyle Event Handlers
'================================================================================

Private Sub mSessionEndGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mSessionEndGridLineStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

setSessionEndGridLineStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mSessionStartGridLineStyle Event Handlers
'================================================================================

Private Sub mSessionStartGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mSessionStartGridLineStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

setSessionStartGridLineStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mStyle_PropertyChanged"
Dim failpoint As String
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
Case UCase$("SessionEndGridLineStyle")
    setSessionEndGridLineStyle
Case UCase$("SessionStartGridLineStyle")
    setSessionStartGridLineStyle
Case UCase$("YAxisTextStyle")
    setYAxisTextStyle
Case UCase$("YCursorTextStyle")
    setYCursorTextStyle
Case UCase$("YScaleQuantum")
    setYScaleQuantum
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unhandled property change"
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mYAxisTextStyle Event Handlers
'================================================================================

Private Sub mYAxisTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mYAxisTextStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

setYAxisTextStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mYCursorTextStyle Event Handlers
'================================================================================

Private Sub mYCursorTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mYCursorTextStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

setYCursorTextStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let ActualHeight(ByVal Value As Long)
mActualHeight = Value
End Property

Friend Property Get ActualHeight() As Long
ActualHeight = mActualHeight
End Property

Public Property Get Autoscaling() As Boolean
Const ProcName As String = "Autoscaling"
Dim failpoint As String
On Error GoTo Err

Autoscaling = mStyle.Autoscaling
If IsPropertySet(AutoscalingIsSet) Then Autoscaling = mAutoscaling

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Autoscaling(ByVal Value As Boolean)
Const ProcName As String = "Autoscaling"
Dim failpoint As String
On Error GoTo Err

If mAutoscaling = Value And IsPropertySet(AutoscalingIsSet) Then Exit Property
mAutoscaling = Value
setPropertyFlag AutoscalingIsSet
setAutoscaling

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BackColor() As Long
Const ProcName As String = "BackColor"
Dim failpoint As String
On Error GoTo Err

BackColor = mStyle.BackColor
If IsPropertySet(BackGradientFillColorsIsSet) Then BackColor = mBackGradientFillColors(0)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BackColor(ByVal Value As Long)
Const ProcName As String = "BackColor"
Dim failpoint As String
On Error GoTo Err

ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = Value
setPropertyFlag AutoscalingIsSet
setBackColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BackGradientFillColors() As Long()
Const ProcName As String = "BackGradientFillColors"
Dim failpoint As String
On Error GoTo Err

BackGradientFillColors = mStyle.BackGradientFillColors
If IsPropertySet(BackGradientFillColorsIsSet) Then BackGradientFillColors = mBackGradientFillColors

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BackGradientFillColors(ByRef Value() As Long)
Dim ar() As Long
Const ProcName As String = "BackGradientFillColors"
Dim failpoint As String
On Error GoTo Err

ar = Value
mBackGradientFillColors = ar
setPropertyFlag BackGradientFillColorsIsSet
setBackGradientFillColors

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Bottom() As Double
Const ProcName As String = "Bottom"
Dim failpoint As String
On Error GoTo Err

Bottom = mViewport.Boundary.Bottom

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Bottom(ByVal Value As Double)
Const ProcName As String = "Bottom"
Dim failpoint As String
On Error GoTo Err

mViewport.Bottom = Value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.Bottom = Value
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get Canvas() As Canvas
Const ProcName As String = "Canvas"
Dim failpoint As String
On Error GoTo Err

Set Canvas = mViewport.Canvas

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal Value As Boolean)
Const ProcName As String = "CursorSnapsToTickBoundaries"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set snapCursorToTickBoundaries for an axis region"
End If
If mCursorSnapsToTickBoundaries = Value And IsPropertySet(CursorSnapsToTickBoundariesIsSet) Then Exit Property
mCursorSnapsToTickBoundaries = Value
setPropertyFlag CursorSnapsToTickBoundariesIsSet
setCursorSnapsToTickBoundaries

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CursorSnapsToTickBoundaries() As Boolean
Const ProcName As String = "CursorSnapsToTickBoundaries"
Dim failpoint As String
On Error GoTo Err

CursorSnapsToTickBoundaries = mStyle.CursorSnapsToTickBoundaries
If IsPropertySet(CursorSnapsToTickBoundariesIsSet) Then CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GridlineSpacingY() As Double
Const ProcName As String = "GridlineSpacingY"
Dim failpoint As String
On Error GoTo Err

GridlineSpacingY = mStyle.GridlineSpacingY
If IsPropertySet(GridlineSpacingYIsSet) Then GridlineSpacingY = mGridlineSpacingY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let GridlineSpacingY(ByVal Value As Double)
Const ProcName As String = "GridlineSpacingY"
Dim failpoint As String
On Error GoTo Err

If mGridlineSpacingY = Value And IsPropertySet(GridlineSpacingYIsSet) Then Exit Property
mGridlineSpacingY = Value
setPropertyFlag GridlineSpacingYIsSet
setGridlineSpacingY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GridLineStyle() As LineStyle
Const ProcName As String = "GridLineStyle"
Dim failpoint As String
On Error GoTo Err

Set GridLineStyle = mStyle.GridLineStyle
If IsPropertySet(GridLineStyleIsSet) Then Set GridLineStyle = mGridLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let GridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "GridLineStyle"
Dim failpoint As String
On Error GoTo Err

If mGridLineStyle Is Value And IsPropertySet(GridLineStyleIsSet) Then Exit Property
Set mGridLineStyle = Value
If mGridLineStyle Is Nothing Then
    ClearProperty GridLineStyleIsSet
Else
    setPropertyFlag GridLineStyleIsSet
End If
setGridLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GridTextStyle() As TextStyle
Const ProcName As String = "GridTextStyle"
Dim failpoint As String
On Error GoTo Err

Set GridTextStyle = mStyle.GridTextStyle
If IsPropertySet(GridTextStyleIsSet) Then Set GridTextStyle = mGridTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let GridTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "GridTextStyle"
Dim failpoint As String
On Error GoTo Err

If mGridTextStyle Is Value And IsPropertySet(GridTextStyleIsSet) Then Exit Property
Set mGridTextStyle = Value
If mGridTextStyle Is Nothing Then
    ClearProperty GridTextStyleIsSet
Else
    setPropertyFlag GridTextStyleIsSet
End If
setGridTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let handle(ByVal Value As Long)
mHandle = Value
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get HasGrid() As Boolean
Const ProcName As String = "HasGrid"
Dim failpoint As String
On Error GoTo Err

HasGrid = mStyle.HasGrid
If IsPropertySet(HasGridIsSet) Then HasGrid = mHasGrid

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HasGrid(ByVal Value As Boolean)
Const ProcName As String = "HasGrid"
Dim failpoint As String
On Error GoTo Err

If mHasGrid = Value And IsPropertySet(HasGridIsSet) Then Exit Property
mHasGrid = Value
setPropertyFlag HasGridIsSet
setHasGrid

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HasGridText() As Boolean
Const ProcName As String = "HasGridText"
Dim failpoint As String
On Error GoTo Err

HasGridText = mStyle.HasGridText
If IsPropertySet(HasGridTextProperty) Then HasGridText = mHasGridText

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HasGridText(ByVal Value As Boolean)
Const ProcName As String = "HasGridText"
Dim failpoint As String
On Error GoTo Err

If mHasGridText = Value And IsPropertySet(HasGridTextProperty) Then Exit Property
mHasGridText = Value
setPropertyFlag HasGridTextProperty
setHasGridText

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Height() As Double
Const ProcName As String = "Height"
Dim failpoint As String
On Error GoTo Err

Height = mViewport.Height

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IntegerYScale(ByVal Value As Boolean)
Const ProcName As String = "IntegerYScale"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set integerYScale for an axis region"
End If

If mIntegerYScale = Value And IsPropertySet(IntegerYScaleIsSet) Then Exit Property
mIntegerYScale = Value
setPropertyFlag IntegerYScaleIsSet
setIntegerYScale

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IntegerYScale() As Boolean
Const ProcName As String = "IntegerYScale"
Dim failpoint As String
On Error GoTo Err

IntegerYScale = mStyle.IntegerYScale
If IsPropertySet(IntegerYScaleIsSet) Then IntegerYScale = mIntegerYScale

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get IsCleared() As Boolean
IsCleared = mIsCleared
End Property

Friend Property Let IsDrawingEnabled(ByVal Value As Boolean)
Const ProcName As String = "IsDrawingEnabled"
Dim failpoint As String
On Error GoTo Err

If mIsDrawingEnabled = Value Then Exit Property
mIsDrawingEnabled = Value
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
    
    doCalculateAndRepaint
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.IsDrawingEnabled = mIsDrawingEnabled
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsXAxisRegion() As Boolean
IsXAxisRegion = (mRegionType = RegionTypeXAxis)
End Property

Public Property Get IsYAxisRegion() As Boolean
IsYAxisRegion = (mRegionType = RegionTypeYAxis)
End Property

Public Property Get Left() As Double
Const ProcName As String = "Left"
Dim failpoint As String
On Error GoTo Err

Left = mViewport.Left

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Left(ByVal Value As Double)
Const ProcName As String = "Left"
Dim failpoint As String
On Error GoTo Err

mViewport.Left = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MinimumHeight() As Double
Const ProcName As String = "MinimumHeight"
Dim failpoint As String
On Error GoTo Err

MinimumHeight = mStyle.MinimumHeight
If IsPropertySet(MinimumHeightIsSet) Then MinimumHeight = mMinimumHeight

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let MinimumHeight(ByVal Value As Double)
Const ProcName As String = "MinimumHeight"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set minimumHeight for an axis region"
End If

If mMinimumHeight = Value And IsPropertySet(MinimumHeightIsSet) Then Exit Property
mMinimumHeight = Value
setPropertyFlag MinimumHeightIsSet
setMinimumHeight

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MinimumPercentHeight() As Double
MinimumPercentHeight = mMinimumPercentHeight
End Property

Friend Property Let MinimumPercentHeight(ByVal Value As Double)
mMinimumPercentHeight = Value
End Property

Public Property Get MousePosition() As Point
Const ProcName As String = "MousePosition"
Dim failpoint As String
On Error GoTo Err

Set MousePosition = gNewPoint(mCurrX, mCurrY)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get PercentHeight() As Double
PercentHeight = mPercentHeight
End Property

Friend Property Let PercentHeight(ByVal Value As Double)
mPercentHeight = Value
End Property

Public Property Get PerformanceTextVisible() As Boolean
PerformanceTextVisible = mPerformanceTextVisible
End Property

Public Property Let PerformanceTextVisible(ByVal Value As Boolean)
mPerformanceTextVisible = Value
End Property

Public Property Get PointerCrosshairsColor() As Long
PointerCrosshairsColor = mPointerCrosshairsColor
End Property

Friend Property Let PointerCrosshairsColor(ByVal Value As Long)
Dim lUseDeferredPainting As Boolean

Const ProcName As String = "PointerCrosshairsColor"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set PointerCrossHairsColor for an axis region"
End If

mPointerCrosshairsColor = Value

lUseDeferredPainting = mUseDeferredPainting
mUseDeferredPainting = False
mCrosshairLineHoriz.Color = mPointerCrosshairsColor
mCrosshairLineVert.Color = mPointerCrosshairsColor
mUseDeferredPainting = lUseDeferredPainting

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerDiscColor() As Long
PointerDiscColor = mPointerDiscColor
End Property

Friend Property Let PointerDiscColor(ByVal Value As Long)
Const ProcName As String = "PointerDiscColor"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set PointerDiscColor for an axis region"
End If

mPointerDiscColor = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerIcon() As IPictureDisp
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal Value As IPictureDisp)

Const ProcName As String = "PointerIcon"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then Exit Property
If Value Is mPointerIcon Then Exit Property

If mPointerMode <> PointerModeDefault Then
    Set mPointerIcon = Value
    Exit Property
End If

If Not mDrawn Then
    Set mPointerIcon = Value
    Exit Property
End If

If mIsDrawingEnabled Then hideCursor

Set mPointerIcon = Value

If mIsDrawingEnabled Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get PointerStyle() As PointerStyles
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal Value As PointerStyles)
Dim failpoint As Long
Const ProcName As String = "PointerStyle"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set PointerStyle for an axis region"
End If


If mPointerStyle = Value Then Exit Property

If mPointerMode <> PointerModeDefault Then
    mPointerStyle = Value
    Exit Property
End If

If Not mDrawn Then
    mPointerStyle = Value
    Exit Property
End If

If mIsDrawingEnabled Then hideCursor

mPointerStyle = Value

If mIsDrawingEnabled Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If
Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PriceFormatter(ByVal Value As IPriceFormatter)
If Not Value Is Nothing Then
    Set mPriceFormatter = Value
Else
    Set mPriceFormatter = New PriceFormatter
End If

mPriceFormatter.IntegerYScale = IntegerYScale
mPriceFormatter.YScaleQuantum = YScaleQuantum

If Not mHorizontalGrid Is Nothing Then
    mHorizontalGrid.PriceFormatter = Value
    doRepaint
End If
End Property

Public Property Get PriceFormatter() As IPriceFormatter
Set PriceFormatter = mHorizontalGrid.PriceFormatter
End Property

Public Property Get SessionEndGridLineStyle() As LineStyle
Const ProcName As String = "SessionEndGridLineStyle"
Dim failpoint As String
On Error GoTo Err

Set SessionEndGridLineStyle = mStyle.SessionEndGridLineStyle
If IsPropertySet(SessionEndGridLineStyleIsSet) Then Set SessionEndGridLineStyle = mSessionEndGridLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let SessionEndGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionEndGridLineStyle"
Dim failpoint As String
On Error GoTo Err

If mSessionEndGridLineStyle Is Value And IsPropertySet(SessionEndGridLineStyleIsSet) Then Exit Property
Set mSessionEndGridLineStyle = Value
If mSessionEndGridLineStyle Is Nothing Then
    ClearProperty SessionEndGridLineStyleIsSet
Else
    setPropertyFlag SessionEndGridLineStyleIsSet
End If
setSessionEndGridLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionStartGridLineStyle() As LineStyle
Const ProcName As String = "SessionStartGridLineStyle"
Dim failpoint As String
On Error GoTo Err

Set SessionStartGridLineStyle = mStyle.SessionStartGridLineStyle
If IsPropertySet(SessionStartGridLineStyleIsSet) Then Set SessionStartGridLineStyle = mSessionStartGridLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let SessionStartGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionStartGridLineStyle"
Dim failpoint As String
On Error GoTo Err

If mSessionStartGridLineStyle Is Value And IsPropertySet(SessionStartGridLineStyleIsSet) Then Exit Property
Set mSessionStartGridLineStyle = Value
If mSessionStartGridLineStyle Is Nothing Then
    ClearProperty SessionStartGridLineStyleIsSet
Else
    setPropertyFlag SessionStartGridLineStyleIsSet
End If
setSessionStartGridLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let SessionStartTime(ByVal Value As Date)
mSessionStartTime = Value
End Property

Public Property Get Style() As ChartRegionStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal Value As ChartRegionStyle)
Const ProcName As String = "Style"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must not be Nothing"
End If

Set mStyle = Value

If mViewport Is Nothing Then Exit Property

processStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Title() As Text
Const ProcName As String = "Title"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "An axis region does not have a title"
End If

If mTitle Is Nothing Then setupTitle
Set Title = mTitle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Top() As Double
Const ProcName As String = "Top"
Dim failpoint As String
On Error GoTo Err

Top = mViewport.Top

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Top(ByVal Value As Double)
Const ProcName As String = "Top"
Dim failpoint As String
On Error GoTo Err

mViewport.Top = Value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.Top = Value
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let VerticalGridTimePeriod(ByVal Value As TimePeriod)
Set mVerticalGridTimePeriod = Value
End Property

Friend Property Get Viewport() As Viewport
Set Viewport = mViewport
End Property

Public Property Get Width() As Single
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

Width = mViewport.Width

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Friend Property Let YAxisRegion(ByVal Value As ChartRegion)

Const ProcName As String = "YAxisRegion"
Dim failpoint As String
On Error GoTo Err

Set mYAxisRegion = Value

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
mYCursorText.Size = gNewSize(90, 0, CoordsRelative)
setYCursorTextStyle

If Not mHorizontalGrid Is Nothing Then mHorizontalGrid.YAxisRegion = mYAxisRegion
setYAxisTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get YAxisRegion() As ChartRegion
Set YAxisRegion = mYAxisRegion
End Property

Public Property Get YAxisTextStyle() As TextStyle
Const ProcName As String = "YAxisTextStyle"
Dim failpoint As String
On Error GoTo Err

Set YAxisTextStyle = mStyle.YAxisTextStyle
If IsPropertySet(YAxisTextStyleIsSet) Then Set YAxisTextStyle = mYAxisTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YAxisTextStyle"
Dim failpoint As String
On Error GoTo Err

If mYAxisTextStyle Is Value And IsPropertySet(YAxisTextStyleIsSet) Then Exit Property
Set mYAxisTextStyle = Value
If mYAxisTextStyle Is Nothing Then
    ClearProperty YAxisTextStyleIsSet
Else
    setPropertyFlag YAxisTextStyleIsSet
End If
setYAxisTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YCursorTextStyle() As TextStyle
Const ProcName As String = "YCursorTextStyle"
Dim failpoint As String
On Error GoTo Err

Set YCursorTextStyle = mStyle.YCursorTextStyle
If IsPropertySet(YCursorTextStyleIsSet) Then Set YCursorTextStyle = mYCursorTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YCursorTextStyle"
Dim failpoint As String
On Error GoTo Err

If mYCursorTextStyle Is Value And IsPropertySet(YCursorTextStyleIsSet) Then Exit Property
Set mYCursorTextStyle = Value
If mYCursorTextStyle Is Nothing Then
    ClearProperty YCursorTextStyleIsSet
Else
    setPropertyFlag YCursorTextStyleIsSet
End If
setYCursorTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YScaleGridSpacing() As Single
Const ProcName As String = "YScaleGridSpacing"
Dim failpoint As String
On Error GoTo Err

YScaleGridSpacing = mHorizontalGridSpacer.YScaleGridSpacing

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YScaleQuantum(ByVal Value As Double)
Const ProcName As String = "YScaleQuantum"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set YScaleQuantum for an axis region"
End If

If Value <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "YScaleQuantum must be greater than zero"
End If

If mYScaleQuantum = Value And IsPropertySet(YScaleQuantumIsSet) Then Exit Property
mYScaleQuantum = Value
setPropertyFlag YScaleQuantumIsSet
setYScaleQuantum

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YScaleQuantum() As Double
Const ProcName As String = "YScaleQuantum"
Dim failpoint As String
On Error GoTo Err

YScaleQuantum = mStyle.YScaleQuantum
If IsPropertySet(YScaleQuantumIsSet) Then YScaleQuantum = mYScaleQuantum

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddGraphicObjectSeries( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal Name As String) As IGraphicObjectSeries
Const ProcName As String = "AddGraphicObjectSeries"
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function AddLine( _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As LineStyle) As Line
Static miscLines As LineSeries
Dim lLine As Line

Const ProcName As String = "AddLine"
Dim failpoint As String
On Error GoTo Err

If miscLines Is Nothing Then
    Set miscLines = AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerHighestUser - 1)
End If
Set lLine = miscLines.Add
If pLayer <> -1 Then lLine.Layer = pLayer
If Not Style Is Nothing Then lLine.LocalStyle = Style

Set AddLine = lLine

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub AddPeriod(ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date)
Const ProcName As String = "AddPeriod"
Dim failpoint As String
On Error GoTo Err

mData.AddPeriod pPeriodNumber, pTimestamp
If CDbl(pTimestamp) <> 0 Then setVerticalGridLine pTimestamp, pPeriodNumber, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub AddSessionEndGridline( _
                ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date)
                
Const ProcName As String = "AddSessionEndGridline"
Dim failpoint As String
On Error GoTo Err

setVerticalGridLine(pTimestamp, pPeriodNumber, True).LocalStyle = SessionEndGridLineStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub AddSessionStartGridline( _
                ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date)
                
Const ProcName As String = "AddSessionStartGridline"
Dim failpoint As String
On Error GoTo Err

setVerticalGridLine(pTimestamp, pPeriodNumber, True).LocalStyle = SessionStartGridLineStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function AddText( _
                Optional ByVal Value As String, _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As TextStyle) As Text
Static miscTexts As TextSeries
Dim lText As Text

Const ProcName As String = "AddText"
Dim failpoint As String
On Error GoTo Err

If miscTexts Is Nothing Then
    Set miscTexts = AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerHighestUser)
End If
Set lText = miscTexts.Add
If Value <> "" Then lText.Text = Value
If pLayer <> -1 Then lText.Layer = pLayer
If Not Style Is Nothing Then lText.LocalStyle = Style

Set AddText = lText

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub CalculateAndRepaint()
Const ProcName As String = "CalculateAndRepaint"
Dim failpoint As String
On Error GoTo Err

mData.DetermineInScopeObjects
calculateVerticalBounds
mData.DetermineVisibleObjects
paint

If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.SetVerticalScaleInternal mViewport.Bottom, mViewport.Top
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub ChangeGaugeAndRepaint()
Const ProcName As String = "changeGaugeAndRepaint"
Dim failpoint As String
On Error GoTo Err

mData.CalculateInscopeGaugeDependentObjects
paint
If Not mYAxisRegion Is Nothing Then mYAxisRegion.NotifyResizedY

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub ClearRegion()
Const ProcName As String = "ClearRegion"
Dim failpoint As String
On Error GoTo Err

If mIsCleared Then Exit Sub

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

mIsCleared = True

mGraphObjSeriesCollection.Finish

mGraphObjServiceProvider.Finish
Set mGraphObjServiceProvider = Nothing

Set mChartRef = Nothing

mData.Clear
Set mData = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearSelection()
Const ProcName As String = "ClearSelection"
Dim failpoint As String
On Error GoTo Err

deSelectObjectOrSeries mSelectedObject
Set mSelectedObject = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = 0
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Click()
Const ProcName As String = "Click"
Dim failpoint As String
On Error GoTo Err

clickSelectedObjectOrSeries
RaiseEvent Click

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub DblCLick()
Const ProcName As String = "DblCLick"
Dim failpoint As String
On Error GoTo Err

dblClickSelectedObjectOrSeries
RaiseEvent DblCLick

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub DrawCursor(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "DrawCursor"
Dim failpoint As String
On Error GoTo Err

If Y = MinusInfinitySingle Then
    Y = mViewport.Top + 2 * mViewport.Height  ' make sure the pointer isn't visible
End If

If Round(X) >= YAxisPosition Then Exit Sub

hideCursor
showCursor Round(X), Y

mPrevCursorX = X
mPrevCursorY = Y

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function FormatYValue(ByVal Y As Double) As String
Const ProcName As String = "FormatYValue"
Dim failpoint As String
On Error GoTo Err

FormatYValue = mPriceFormatter.FormatPrice(Y)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub HideGrid()
Const ProcName As String = "HideGrid"
Dim failpoint As String
On Error GoTo Err

If mGridHidden Then Exit Sub

mHorizontalGrid.Hidden = True
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal Name As String, _
                ByVal pChart As Chart, _
                ByVal pCanvas As Canvas, _
                ByVal pRegionType As RegionTypes)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

mName = Name

Set mChartRef = CreateWeakReference(pChart)
Set mViewport = New Viewport
mViewport.Region = Me
mViewport.Canvas = pCanvas
mViewport.MousePointer = vbDefault

mRegionType = pRegionType

mData.Initialise pChart, Me, mGraphObjSeriesCollection

setupCrosshairs
processStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function IsGridHidden() As Boolean
Const ProcName As String = "IsGridHidden"
Dim failpoint As String
On Error GoTo Err

IsGridHidden = mHorizontalGrid.Hidden

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Const ProcName As String = "MouseDown"
Dim failpoint As String
On Error GoTo Err

mCurrX = X
mCurrY = Y
If mPointerMode = PointerModeSelection Then setSelectedObjectOrSeries
RaiseEvent MouseDown(Button, Shift, X, Y)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Friend Sub NotifyResizedY()
' notifies the region that the physical Height of the canvas has changed.
' This means the Y gauge will have changed, so we need to recalculate
' the Position of any in-scope gauge-dependent objects.

Const ProcName As String = "NotifyResizedY"
Dim failpoint As String
On Error GoTo Err

mViewport.NotifyResizedY

If mUseDeferredPainting Then
    mDeferredPainter.ChangeGaugeAndRepaint
Else
    ChangeGaugeAndRepaint
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RepaintAreas( _
                ByRef pAreas() As TRectangle, _
                ByVal pCount As Long)
                
Const ProcName As String = "RepaintAreas"
Dim i As Long

Dim failpoint As String
On Error GoTo Err

mData.MarkObjectsForRedraw pAreas, pCount

For i = 0 To pCount - 1
    If i = 0 Then
        mViewport.Canvas.SetClippingRegion mViewport.ConvertTRectangleToRect(pAreas(i))
    Else
        mViewport.Canvas.AddToClippingRegion mViewport.ConvertTRectangleToRect(pAreas(i))
    End If
Next

mViewport.PaintBackground

Dim numMarkedObjects As Long
numMarkedObjects = mData.PaintMarkedObjects
Debug.Print "Region " & mHandle & ": " & numMarkedObjects & " marked objects painted"

mViewport.Canvas.ClearClippingRegion

mViewport.Canvas.Surface.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RepaintRegion()
Const ProcName As String = "RepaintRegion"
Dim failpoint As String
On Error GoTo Err

paint
If Not mYAxisRegion Is Nothing Then mYAxisRegion.RepaintRegion

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RedrawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
'unDrawCursor areaToRedraw
Const ProcName As String = "RedrawObject"
Dim failpoint As String
On Error GoTo Err

drawObject graphObj, areaToRedraw
RedrawArea areaToRedraw, graphObj.Layer + 1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveGraphicObject( _
                ByVal graphObj As IGraphicObject)
Const ProcName As String = "RemoveGraphicObject"
Dim failpoint As String
On Error GoTo Err

RemoveGraphicObjectEx graphObj, True, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveGraphicObjectEx( _
                ByVal graphObj As IGraphicObject, _
                ByVal undraw As Boolean, _
                ByVal removeFromSeries As Boolean)
Const ProcName As String = "RemoveGraphicObjectEx"
Dim failpoint As String
On Error GoTo Err

failpoint = 100

If Not mSelectedObject Is Nothing Then
    If graphObj Is mSelectedObject Then ClearSelection
End If

failpoint = 300

If removeFromSeries Then mGraphObjSeriesCollection.Item(graphObj.SeriesID).RemoveObject graphObj

failpoint = 800

If undraw Then UndrawArea graphObj.BoundingRectangleInViewport, graphObj

mData.RemoveGraphicObject graphObj

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveGraphicObjectSeries( _
                ByVal pSeries As IGraphicObjectSeries)
Const ProcName As String = "RemoveGraphicObjectSeries"
Dim failpoint As String
On Error GoTo Err

If pSeries Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "The supplied argument is Nothing"
End If

mGraphObjSeriesCollection.Remove pSeries

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub ScaleAndRepaint()
Const ProcName As String = "ScaleAndRepaint"
Dim failpoint As String
On Error GoTo Err

calculateVerticalBounds
mData.DetermineVisibleObjects
paint

If Not mYAxisRegion Is Nothing Then mYAxisRegion.SetVerticalScaleInternal mViewport.Bottom, mViewport.Top

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ScaleUp(ByVal factor As Single)
Const ProcName As String = "ScaleUp"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot scale an axis region"
End If

With mViewport
    SetVerticalScale .Bottom + factor * .Height / 2, _
                                .Top - factor * .Height / 2
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ScrollVertical(ByVal amount As Single)
Const ProcName As String = "ScrollVertical"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot scroll an axis region"
End If

Autoscaling = False
SetVerticalScale mViewport.Bottom + amount, mViewport.Top + amount

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ScrollVerticalProportion(ByVal proportion As Single)
Const ProcName As String = "ScrollVerticalProportion"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot scroll an axis region"
End If

Autoscaling = False
SetVerticalScale mViewport.Bottom + proportion * mViewport.Height, _
                mViewport.Top + proportion * mViewport.Height

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetPeriodBackgroundColor( _
                ByVal PeriodNumber As Long, _
                ByVal Color As Long)
Dim Layer As Long
Dim index As Long
Dim periodrect As TRectangle
Dim redrawAreas(0) As TRectangle

Const ProcName As String = "SetPeriodBackgroundColor"
Dim failpoint As String
On Error GoTo Err

mData.PeriodBackColor(PeriodNumber) = Color

If Not mIsDrawingEnabled Then Exit Sub

If Color = 0 Then Exit Sub

With mViewport
    .Canvas.SetPenAttributes Color, 1, LineInsideSolid, DrawModeCopyPen
    .Canvas.SetBrushAttributes Color, FillSolid
    .Canvas.DrawRectangle PeriodNumber - 0.5, .Bottom, PeriodNumber + 0.5, .Top

    periodrect.Left = PeriodNumber - 0.5
    periodrect.Right = PeriodNumber + 0.5
    periodrect.Bottom = .Bottom
    periodrect.Top = .Top
    periodrect.isValid = True
End With
    
mData.RedrawArea periodrect, LayerNumbers.LayerMin

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SetPeriodsInView( _
                ByVal minValue As Single, _
                ByVal maxValue As Single)
Const ProcName As String = "SetPeriodsInView"
Dim failpoint As String
On Error GoTo Err

mMinInViewPeriod = minValue
mMaxInViewPeriod = maxValue
mYAxisPosition = maxValue + 1
mViewport.Left = minValue
mViewport.Right = maxValue + 0.99

doCalculateAndRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SetPointerModeDefault()
Const ProcName As String = "SetPointerModeDefault"
Dim failpoint As String
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SetPointerModeSelection()
Const ProcName As String = "SetPointerModeSelection"
Dim failpoint As String
On Error GoTo Err

If mPointerMode = PointerModeSelection Then Exit Sub

If mIsDrawingEnabled Then hideCursor

mPointerMode = PointerModeSelection

If mIsDrawingEnabled Then
    setMouseIcon PointerStyles.PointerSelection
    showCursor mPrevCursorX, mPrevCursorY
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SetPointerModeTool( _
                ByVal toolPointerStyle As PointerStyles, _
                ByVal icon As IPictureDisp)

Const ProcName As String = "SetPointerModeTool"
Dim failpoint As String
On Error GoTo Err

If mIsDrawingEnabled Then hideCursor

mPointerMode = PointerModeTool
mPointerToolStyle = toolPointerStyle
Set mPointerToolIcon = icon

If mIsDrawingEnabled Then
    setMouseIcon mPointerToolStyle
    showCursor mPrevCursorX, mPrevCursorY
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

' Note that the arguments are ByRef because they may be adjusted
' if they are the same when expressed as Singles.
Public Sub SetVerticalScale(ByRef lowValue As Double, _
                            ByRef highValue As Double)
Const ProcName As String = "SetVerticalScale"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set the scale for an axis region"
End If

Autoscaling = False

SetVerticalScaleInternal lowValue, highValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SetVerticalScaleInternal(ByRef lowValue As Double, _
                            ByRef highValue As Double)
Const ProcName As String = "SetVerticalScaleInternal"
Dim failpoint As String
On Error GoTo Err

setViewportVerticalBounds lowValue, highValue

doCalculateAndRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ShowGrid()
Const ProcName As String = "ShowGrid"
Dim failpoint As String
On Error GoTo Err

If Not mGridHidden Then Exit Sub

mHorizontalGrid.Hidden = False
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub UndrawAreaEx( _
                ByRef undrawnArea As TRectangle, _
                Optional ByVal objToExclude As IGraphicObject)
    Const ProcName As String = "UndrawAreaEx"
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function addVerticalGridLine(ByVal PeriodNumber As Long, _
                                ByVal Key As String, _
                                ByVal label As String, _
                                ByVal coarseKey As String, _
                                ByVal coarseLabel As String) As ChartSkil26.Line
Dim vGridLine As ChartSkil26.Line
Dim vGridText As Text
Dim vGridTextCoarse As Text

Const ProcName As String = "addVerticalGridLine"
Dim failpoint As String
On Error GoTo Err

If HasGrid Then
    On Error Resume Next
    Set vGridLine = mXGridLines.Item(Key)
    On Error GoTo Err
    If vGridLine Is Nothing Then
        Set vGridLine = mXGridLines.Add(Key)
        vGridLine.FixedX = False
        vGridLine.FixedY = True
        vGridLine.SetPosition gNewPoint(PeriodNumber - 0.5, _
                                    0, _
                                    CoordsLogical, _
                                    CoordsRelative), _
                                gNewPoint(PeriodNumber - 0.5, _
                                    100, _
                                    CoordsLogical, _
                                    CoordsRelative)
    Else
        If vGridLine.Point1.XLogical(mViewport.Region) > PeriodNumber Then
            vGridLine.SetPosition gNewPoint(PeriodNumber - 0.5, _
                                        0, _
                                        CoordsLogical, _
                                        CoordsRelative), _
                                    gNewPoint(PeriodNumber - 0.5, _
                                        100, _
                                        CoordsLogical, _
                                        CoordsRelative)
        End If
    End If
    Set addVerticalGridLine = vGridLine
End If

If HasGridText Then
    On Error Resume Next
    Set vGridText = mXGridTexts.Item(Key)
    On Error GoTo Err
    If vGridText Is Nothing Then
        Set vGridText = mXGridTexts.Add(Key)
        vGridText.FixedX = False
        vGridText.FixedY = True
        vGridText.Position = gNewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsDistance)
        vGridText.Text = label
    End If

    On Error Resume Next
    Set vGridTextCoarse = mXGridTextsCoarse.Item(coarseKey)
    On Error GoTo Err
    If vGridTextCoarse Is Nothing Then
        Set vGridTextCoarse = mXGridTextsCoarse.Add(coarseKey)
        vGridTextCoarse.FixedX = False
        vGridTextCoarse.FixedY = True
        vGridTextCoarse.Position = gNewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsCounterDistance)
        vGridTextCoarse.Text = coarseLabel
    End If
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Private Sub calculateVerticalBounds()
Dim scaleMax As Double
Dim scaleMin As Double

Const ProcName As String = "calculateVerticalBounds"
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function chartObj() As Chart
Const ProcName As String = "chartObj"
Dim failpoint As String
On Error GoTo Err

Set chartObj = mChartRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub ClearProperty( _
                ByVal flag As PropertyOverrideFlags)
Const ProcName As String = "ClearProperty"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function clickSelectedObjectOrSeries() As Boolean
Const ProcName As String = "clickSelectedObjectOrSeries"
Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function dblClickSelectedObjectOrSeries() As Boolean
Const ProcName As String = "dblClickSelectedObjectOrSeries"
Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function deSelectObjectOrSeries(ByVal graphObj As IGraphicObject) As Boolean
Const ProcName As String = "deSelectObjectOrSeries"
Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub doRepaint()
Const ProcName As String = "doRepaint"
Dim failpoint As String
On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.RepaintRegion
    Else
        paint
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub doScaleAndRepaint()
Const ProcName As String = "doScaleAndRepaint"
Dim failpoint As String
On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.ScaleAndRepaint
    Else
        ScaleAndRepaint
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub doCalculateAndRepaint()
Const ProcName As String = "doCalculateAndRepaint"
Dim failpoint As String
On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.CalculateAndRepaint
    Else
        CalculateAndRepaint
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawCrosshairsPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single

Const ProcName As String = "drawCrosshairsPointer"
Dim failpoint As String
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

lineX = Round(X)

mCrosshairLineHoriz.SetPosition gNewPoint(0, Y, CoordsRelative), _
                                gNewPoint(100, Y, CoordsRelative)
mCrosshairLineVert.SetPosition gNewPoint(lineX, 0, , CoordsRelative), _
                                gNewPoint(lineX, 100, , CoordsRelative)
    
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawCustomPointer(ByVal X As Single, ByVal Y As Single)

Const ProcName As String = "drawCustomPointer"
Dim failpoint As String
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawDiscPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim Size As Size

Const ProcName As String = "drawDiscPointer"
Dim failpoint As String
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

With mViewport
    lineX = Round(X)
    Set Size = gNewSize(0.4, 0.4)
    
    .Canvas.SetPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.DrawLine lineX - Size.WidthLogical(mViewport.Region), Y, lineX + Size.WidthLogical(mViewport.Region), Y
    .Canvas.DrawLine lineX, Y - Size.HeightLogical(mViewport.Region), lineX, Y + Size.HeightLogical(mViewport.Region)
    
    .Canvas.SetPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.SetBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .Canvas.DrawCircle gNewPoint(lineX, Y), gNewDimension(0.8).LengthLogicalX(mViewport.Region)
    
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Dim drawAreas(0) As TRectangle
Const ProcName As String = "drawObject"
Dim failpoint As String
On Error GoTo Err

drawAreas(0) = RectIntersection(mViewport.Boundary, areaToRedraw)

graphObj.Draw drawAreas

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawSelectionPointer(ByVal X As Single, ByVal Y As Single)

Const ProcName As String = "drawSelectionPointer"
Dim failpoint As String
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

' nothing to draw since the pointer is a cursor and Windows draws it

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawToolPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim lSize As Size

Const ProcName As String = "drawToolPointer"
Dim failpoint As String
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

With mViewport
    lineX = Round(X)
    Set lSize = gNewSize(0.4, 0.4)
    
    mCrosshairLineHoriz.SetPosition gNewPoint(lineX - lSize.WidthLogical(mViewport.Region), Y), gNewPoint(lineX + lSize.WidthLogical(mViewport.Region), Y)
    mCrosshairLineVert.SetPosition gNewPoint(lineX, Y - lSize.HeightLogical(mViewport.Region)), gNewPoint(lineX, Y + lSize.HeightLogical(mViewport.Region))
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawYCursorText(ByVal Y As Single)
Const ProcName As String = "drawYCursorText"
Dim failpoint As String
On Error GoTo Err

If Not mYAxisRegion Is Nothing Then
    mYCursorText.Position = gNewPoint(5, Y, CoordsRelative, CoordsLogical)
    mYCursorText.Text = mPriceFormatter.FormatPrice(Y)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getPerformanceText() As Text
Static perfText As Text
Dim perffont As StdFont

Const ProcName As String = "getPerformanceText"
Dim failpoint As String
On Error GoTo Err

If perfText Is Nothing Then
    Set perffont = New StdFont
    perffont.Name = "Lucida Console"
    perffont.Size = 8
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
    perfText.Position = gNewPoint(5, 2, CoordsRelative, CoordsRelative)
    perfText.FixedX = True
    perfText.FixedY = True
    perfText.Align = TextAlignModes.AlignBottomLeft
    perfText.IncludeInAutoscale = False
    perfText.MultiLine = True
    perfText.PaddingX = 0.5
    perfText.Layer = LayerTitle
End If
Set getPerformanceText = perfText

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub hideCursor()
Dim lPointerStyle As PointerStyles

Const ProcName As String = "hideCursor"
Dim failpoint As String
On Error GoTo Err

'mUseDeferredPainting = False
mUseDeferredPainting = True

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paint()
Static et As ElapsedTimer

Dim numVisibleObjects  As Long
Dim index As Long

Static repaintCount As Long
Const ProcName As String = "paint"
Dim failpoint As String
On Error GoTo Err

repaintCount = repaintCount + 1
Debug.Print "Region " & mHandle & " repaint count = " & repaintCount

mPaintingInProgress = True

If mPerformanceTextVisible Then
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

mViewport.PaintBackground

paintPeriodBackgrounds

paintHorizontalGrid

numVisibleObjects = mData.PaintAllVisibleObjects
Debug.Print "Region " & mHandle & ": " & numVisibleObjects & " objects painted"

showCursor mPrevCursorX, mPrevCursorY

If mPerformanceTextVisible Then paintPerformanceText numVisibleObjects, et.ElapsedTimeMicroseconds / 1000

mViewport.Canvas.Surface.Refresh

mViewport.ZOrder 0

mDrawn = True
mPaintingInProgress = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintHorizontalGrid()
Const ProcName As String = "paintHorizontalGrid"
On Error GoTo Err

If Not (IsYAxisRegion Or IsXAxisRegion) And HasGrid Then
    mPriceFormatter.YScaleGridSpacing = mHorizontalGridSpacer.Calculate(mViewport.Height, mViewport.HeightCm)
    
    mHorizontalGrid.Generate mViewport.Bottom, mViewport.Top, mHorizontalGridSpacer.YScaleGridSpacing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintPerformanceText( _
                ByVal numberOfObjectsRepainted As Long, _
                ByVal elapsedTimeMillisecs As Single)
Dim perfText As Text
Const ProcName As String = "paintPerformanceText"
Dim failpoint As String
On Error GoTo Err

If mPerformanceTextVisible Then
    Set perfText = getPerformanceText
    perfText.Text = "Repaint time: " & Format(elapsedTimeMillisecs, "0.0") & "ms" & vbCrLf & _
                    "Visible objects: " & numberOfObjectsRepainted & vbCrLf & _
                    "Total objects: " & mData.NumberOfObjects
    ' need to force the Text to draw now
    perfText.Draw
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintPeriodBackgrounds()
Dim i As Long
Const ProcName As String = "paintPeriodBackgrounds"
Dim failpoint As String
On Error GoTo Err

With mViewport
    For i = IIf(mMinInViewPeriod < mData.MinPeriodNumber, mData.MinPeriodNumber, mMinInViewPeriod) To IIf(mData.MaxPeriodNumber < mMaxInViewPeriod, mData.MaxPeriodNumber, mMaxInViewPeriod)
        If mData.PeriodBackColor(i) <> 0 Then
            .Canvas.SetPenAttributes mData.PeriodBackColor(i), 1, LineInsideSolid, DrawModeCopyPen
            .Canvas.SetBrushAttributes mData.PeriodBackColor(i), FillSolid
            .Canvas.DrawRectangle i - 0.5, .Bottom, i + 0.5, .Top
        End If
    Next
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processStyle()

Const ProcName As String = "processStyle"
Dim failpoint As String
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub RedrawArea( _
                ByRef areaToRedraw As TRectangle, _
                ByVal startLayer As Long, _
                Optional ByVal objToExclude As IGraphicObject)
Const ProcName As String = "redrawArea"
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub redrawChangedObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Dim i As Long
Dim handled As Boolean

Const ProcName As String = "redrawChangedObject"
Dim failpoint As String
On Error GoTo Err

If areaToRedraw.isValid Then
    If mUseDeferredPainting Then
        'mDeferredPainter.AddChangedObject graphObj, areaToRedraw
        mDeferredPainter.AddAreaToRedraw areaToRedraw
    Else
        RedrawObject graphObj, areaToRedraw
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function selectObjectOrSeries(ByVal graphObj As IGraphicObject) As Boolean
Const ProcName As String = "selectObjectOrSeries"
Dim failpoint As String
On Error GoTo Err

If graphObj.IsSelectable Then
    graphObj.Selected = True
    selectObjectOrSeries = True
ElseIf mGraphObjSeriesCollection.Item(graphObj.SeriesID).IsSelectable Then
    mGraphObjSeriesCollection.Item(graphObj.SeriesID).Selected = True
    selectObjectOrSeries = True
Else
    selectObjectOrSeries = False
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setAutoscaling()
Const ProcName As String = "setAutoscaling"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set autoscale for an axis region"
End If

If Autoscaling Then doScaleAndRepaint

RaiseEvent AutoscalingChanged

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setBackColor()
Const ProcName As String = "setBackColor"
Dim failpoint As String
On Error GoTo Err

mViewport.BackColor = BackColor
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setBackGradientFillColors()
Const ProcName As String = "setBackGradientFillColors"
Dim failpoint As String
On Error GoTo Err

mViewport.GradientFillColors = BackGradientFillColors
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setViewportVerticalBounds( _
                ByRef Bottom As Double, _
                ByRef Top As Double)
    
Const ProcName As String = "setViewportVerticalBounds"
Dim failpoint As String
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setGridlineSpacingY()
Const ProcName As String = "setGridlineSpacingY"
Dim failpoint As String
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set gridline spacing for an axis region"
End If

mHorizontalGridSpacer.GridlineSpacingY = GridlineSpacingY
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setGridLineStyle()
Const ProcName As String = "setGridLineStyle"
Dim failpoint As String
On Error GoTo Err

If Not mHorizontalGrid Is Nothing Then
    mHorizontalGrid.GridLineStyle = GridLineStyle
    doRepaint
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setGridTextStyle()
Const ProcName As String = "setGridTextStyle"
Dim failpoint As String
On Error GoTo Err

setHasGridText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setHasGrid()
Const ProcName As String = "setHasGrid"
Dim failpoint As String
On Error GoTo Err

If HasGrid Then
    Set mXGridLines = Me.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
    mXGridLines.Style = GridLineStyle
    
    If mHorizontalGrid Is Nothing Then Set mHorizontalGrid = New HorizontalGrid
    mHorizontalGrid.Initialise Me
    mHorizontalGrid.GridLineStyle = GridLineStyle
    mHorizontalGrid.Hidden = mGridHidden
    mHorizontalGrid.PriceFormatter = mPriceFormatter
    mHorizontalGrid.TextStyle = YAxisTextStyle
    mHorizontalGrid.YAxisRegion = mYAxisRegion
Else
    Set mXGridLines = Nothing
    If Not mHorizontalGrid Is Nothing Then
        mHorizontalGrid.Finish
        Set mHorizontalGrid = Nothing
    End If
End If
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setHasGridText()

Const ProcName As String = "setHasGridText"
Dim failpoint As String
On Error GoTo Err

If HasGridText Then
    Set mXGridTexts = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText + 1)
    mXGridTexts.Style = GridTextStyle
    ' override the Style for the following, but don't change the Style itself
    
    mXGridTexts.Align = AlignBoxBottomCentre
    'If Not IsXAxisRegion Then
        mXGridTexts.Box = True
        mXGridTexts.BoxThickness = 0
        mXGridTexts.BoxFillWithBackgroundColor = True
        mXGridTexts.BoxStyle = LineInvisible
        mXGridTexts.PaddingY = 0
        mXGridTexts.PaddingX = 4
    'End If
    mXGridTexts.Extended = True

    Set mXGridTextsCoarse = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
    mXGridTextsCoarse.Style = GridTextStyle
    
    ' override the Style for the following, but don't change the Style itself
    mXGridTextsCoarse.Align = AlignBoxTopCentre
    'If Not IsXAxisRegion Then
        mXGridTextsCoarse.Box = True
        mXGridTextsCoarse.BoxThickness = 0
        mXGridTexts.BoxFillWithBackgroundColor = True
        mXGridTextsCoarse.BoxStyle = LineInvisible
        mXGridTextsCoarse.PaddingY = 0
    'End If
    mXGridTextsCoarse.Extended = True
Else
    Set mXGridTexts = Nothing
    Set mXGridTextsCoarse = Nothing
End If
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setIntegerYScale()
Const ProcName As String = "setIntegerYScale"
Dim failpoint As String
On Error GoTo Err

If Not mHorizontalGrid Is Nothing Then
    mPriceFormatter.IntegerYScale = IntegerYScale
    doRepaint
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setMinimumHeight()
Const ProcName As String = "setMinimumHeight"
Dim failpoint As String
On Error GoTo Err

doScaleAndRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setMouseIcon( _
                ByVal PointerStyle As PointerStyles)
Const ProcName As String = "setMouseIcon"
Dim failpoint As String
On Error GoTo Err

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
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be a member of the PointerStyles enum"
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSelectedObjectOrSeries()
Dim newSelectedObject As IGraphicObject

Const ProcName As String = "setSelectedObjectOrSeries"
Dim failpoint As String
On Error GoTo Err

Set newSelectedObject = mData.GetNextSelectableObjectOrSeriesAtPoint(mSelectedObject, mCurrX, mCurrY)

If newSelectedObject Is Nothing Then
    deSelectObjectOrSeries mSelectedObject
    Set mSelectedObject = Nothing
ElseIf newSelectedObject Is mSelectedObject Then
    
Else
    deSelectObjectOrSeries mSelectedObject
    Set mSelectedObject = newSelectedObject
    selectObjectOrSeries newSelectedObject
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setCursorSnapsToTickBoundaries()
' nothing to do
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As PropertyOverrideFlags)
Const ProcName As String = "setPropertyFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSessionEndGridLineStyle()
Const ProcName As String = "setSessionEndGridLineStyle"
Dim failpoint As String
On Error GoTo Err

setHasGrid

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSessionStartGridLineStyle()
Const ProcName As String = "setSessionStartGridLineStyle"
Dim failpoint As String
On Error GoTo Err

setHasGrid

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupCrosshairs()
Const ProcName As String = "setupCrosshairs"
Dim failpoint As String
On Error GoTo Err

Set mCrosshairLineHoriz = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineHoriz.Color = vbRed
mCrosshairLineHoriz.Extended = True
mCrosshairLineHoriz.IncludeInAutoscale = False
mCrosshairLineHoriz.LineStyle = LineSolid
mCrosshairLineHoriz.Thickness = 1

Set mCrosshairLineVert = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineVert.Color = vbRed
mCrosshairLineVert.Extended = True
mCrosshairLineVert.IncludeInAutoscale = False
mCrosshairLineVert.LineStyle = LineSolid
mCrosshairLineVert.Thickness = 1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupTitle()
Const ProcName As String = "setupTitle"
Dim failpoint As String
On Error GoTo Err

Set mTitle = AddText(, LayerTitle)
mTitle.Box = True
mTitle.BoxColor = vbBlack
mTitle.BoxStyle = LineStyles.LineInvisible
mTitle.BoxThickness = 1
mTitle.BoxFillColor = vbWhite
mTitle.BoxFillStyle = FillStyles.FillSolid
mTitle.Position = gNewPoint(0.2, 0.2, CoordsDistance, CoordsCounterDistance)
mTitle.FixedX = True
mTitle.FixedY = True
mTitle.Align = TextAlignModes.AlignBoxTopLeft
mTitle.IncludeInAutoscale = False
mTitle.PaddingX = 0.5

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function setVerticalGridLine(ByVal Timestamp As Date, _
                            ByVal PeriodNumber As Long, _
                            ByVal pIgnoreVerticalGridTimePeriod As Boolean) As ChartSkil26.Line
Dim gridLineTime As Date

Const ProcName As String = "setVerticalGridLine"
Dim failpoint As String
On Error GoTo Err

If CDbl(Timestamp) = 0 Then Exit Function

If pIgnoreVerticalGridTimePeriod Then
    ' just draw a vertical gridline at the specified period
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                Format(Timestamp, "yyyymmddhhnnss"), _
                                                FormatDateTime(Timestamp, vbShortTime), _
                                                Format(Timestamp, "yyyymmdd"), _
                                                Format(Timestamp, "d Mmm yy"))
    Exit Function
End If

If mVerticalGridTimePeriod Is Nothing Then
    ' just draw a vertical gridline every 10 bars
    If PeriodNumber Mod 10 = 1 Then
        Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                    Format(Timestamp, "yyyymmddhhnnss"), _
                                                    FormatDateTime(Timestamp, vbShortTime), _
                                                    Format(Timestamp, "yyyymmdd"), _
                                                    Format(Timestamp, "d Mmm yy"))
    End If
    Exit Function
End If

gridLineTime = BarStartTime(Timestamp, mVerticalGridTimePeriod, mSessionStartTime)
    
Select Case mVerticalGridTimePeriod.Units
Case TimePeriodSecond
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                Format(gridLineTime, "yyyymmddhhnnss"), _
                                                FormatDateTime(gridLineTime, vbLongTime), _
                                                Format(gridLineTime, "yyyymmdd"), _
                                                Format(gridLineTime, "d Mmm yy"))
Case TimePeriodMinute
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                Format(gridLineTime, "yyyymmddhhnn"), _
                                                FormatDateTime(gridLineTime, vbShortTime), _
                                                Format(gridLineTime, "yyyymmdd"), _
                                                Format(gridLineTime, "d Mmm yy"))
Case TimePeriodHour
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                Format(gridLineTime, "yyyymmddhh"), _
                                                FormatDateTime(gridLineTime, vbShortTime), _
                                                Format(gridLineTime, "yyyymmdd"), _
                                                Format(gridLineTime, "d Mmm yy"))
Case TimePeriodDay
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                Format(gridLineTime, "yyyymmdd"), _
                                                Format(gridLineTime, "d"), _
                                                Format(gridLineTime, "yyyymm"), _
                                                Format(gridLineTime, "Mmm yy"))
Case TimePeriodWeek
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                Format(gridLineTime, "yyyymmdd"), _
                                                Format(gridLineTime, "d Mmm"), _
                                                Format(gridLineTime, "yyyy"), _
                                                Format(gridLineTime, "yyyy"))
Case TimePeriodMonth
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                Format(gridLineTime, "yyyymm"), _
                                                Format(gridLineTime, "Mmm"), _
                                                Format(gridLineTime, "yyyy"), _
                                                Format(gridLineTime, "yyyy"))
Case TimePeriodYear
    Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                        Format(gridLineTime, "yyyy"), _
                                        Format(gridLineTime, "YYYY"), _
                                        "$", _
                                        "")
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If PeriodNumber Mod 10 = 1 Then
        Set setVerticalGridLine = addVerticalGridLine(PeriodNumber, _
                                                    Format(gridLineTime, "yyyymmddhhnnss"), _
                                                    FormatDateTime(gridLineTime, vbLongTime), _
                                                    Format(gridLineTime, "yyyymmdd"), _
                                                    Format(gridLineTime, "d Mmm yy"))
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setYAxisTextStyle()
Const ProcName As String = "setYAxisTextStyle"
Dim failpoint As String
On Error GoTo Err

If Not mHorizontalGrid Is Nothing Then
    mHorizontalGrid.TextStyle = YAxisTextStyle
    doRepaint
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setYCursorTextStyle()
Const ProcName As String = "setYCursorTextStyle"
Dim failpoint As String
On Error GoTo Err

If Not mYCursorText Is Nothing Then mYCursorText.LocalStyle = YCursorTextStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setYScaleQuantum()
Const ProcName As String = "setYScaleQuantum"
Dim failpoint As String
On Error GoTo Err

mPriceFormatter.YScaleQuantum = YScaleQuantum

If Not mHorizontalGrid Is Nothing Then
    mPriceFormatter.YScaleQuantum = YScaleQuantum
    mHorizontalGridSpacer.YScaleQuantum = YScaleQuantum
    doRepaint
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub showCursor( _
                ByVal X As Single, _
                ByVal Y As Single)
Dim PointerStyle As PointerStyles

Const ProcName As String = "showCursor"
Dim failpoint As String
On Error GoTo Err

'mUseDeferredPainting = False
mUseDeferredPainting = True

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

drawYCursorText Y

mUseDeferredPainting = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function IsPropertySet( _
                ByVal flag As PropertyOverrideFlags) As Boolean
Const ProcName As String = "IsPropertySet"
Dim failpoint As String
On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub UndrawArea( _
                ByRef undrawnArea As TRectangle, _
                ByVal objToExclude As IGraphicObject)
Const ProcName As String = "UndrawArea"
Dim failpoint As String
On Error GoTo Err

If mPaintingInProgress Then Exit Sub

If mIsDrawingEnabled And undrawnArea.isValid Then
    If mUseDeferredPainting Then
        mDeferredPainter.AddAreaToRedraw undrawnArea
    Else
        UndrawAreaEx undrawnArea, objToExclude
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub undrawCrosshairsPointer()


' nothing to do

End Sub

Private Sub undrawCustomPointer()

' nothing to do

End Sub

Private Sub undrawDiscPointer()
Dim lSize As Size
Dim prevLineX As Single

Const ProcName As String = "undrawDiscPointer"
Dim failpoint As String
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

With mViewport
    prevLineX = Round(mPrevCursorX)
    Set lSize = gNewSize(0.4, 0.4)
    
    .Canvas.SetPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.DrawLine prevLineX - lSize.WidthLogical(mViewport.Region), mPrevCursorY, prevLineX + lSize.WidthLogical(mViewport.Region), mPrevCursorY
    .Canvas.DrawLine prevLineX, mPrevCursorY - lSize.HeightLogical(mViewport.Region), prevLineX, mPrevCursorY + lSize.HeightLogical(mViewport.Region)
    
    .Canvas.SetPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .Canvas.SetBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .Canvas.DrawCircle gNewPoint(prevLineX, mPrevCursorY), gNewDimension(0.8).LengthLogicalX(mViewport.Region)
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub undrawSelectionPointer()

If Not mIsDrawingEnabled Then Exit Sub

' nothing to do since the pointer is actually a proper cursor
End Sub

Private Sub undrawToolPointer()

Const ProcName As String = "undrawToolPointer"
Dim failpoint As String
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub

mCrosshairLineHoriz.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)
mCrosshairLineVert.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


