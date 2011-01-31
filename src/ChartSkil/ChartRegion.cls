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

Private Type PeriodLabels
    PeriodNumber As Long
    Key As String
    Label As String
    CoarseKey As String
    CoarseLabel As String
End Type

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

Private WithEvents mBaseStyle As ChartRegionStyle
Attribute mBaseStyle.VB_VarHelpID = -1
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
Private WithEvents mSessionEndGridlineStyle         As LineStyle
Attribute mSessionEndGridlineStyle.VB_VarHelpID = -1
Private WithEvents mSessionStartGridlineStyle       As LineStyle
Attribute mSessionStartGridlineStyle.VB_VarHelpID = -1
Private WithEvents mYAxisTextStyle                  As TextStyle
Attribute mYAxisTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle                As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

Private mPropertyOverrideFlags                      As PropertyOverrideFlags

Private mPaintingInProgress                         As Boolean

Private mPriceFormatter                             As IPriceFormatter

Private mHorizontalGridSpacer                       As HorizontalGridSpacer

Private mSessStartGridlineStyle                     As LineStyle
Private mSessEndGridlineStyle                       As LineStyle

Private mDivider                                    As PictureBox
Private mDividerGradientFill                        As GradientFill

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Dim i As Long

Set mData = New ChartRegionData
Set mGraphObjSeriesCollection = New GraphObjSeriesCollection
mGraphObjSeriesCollection.Initialise Me

Set mBaseStyle = gDefaultChartRegionStyle

mUseDeferredPainting = True
Set mDeferredPainter = New DeferredRegionPainter
mDeferredPainter.Initialise Me

Set mGraphObjServiceProvider = New GraphObjServiceProvider
mGraphObjServiceProvider.Region = Me

Set mDividerGradientFill = New GradientFill
mDividerGradientFill.Color = &HD0D0D0

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

On Error GoTo Err

UndrawArea undrawnArea, mData.GetGraphicObject(handle)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub GOSP_PaintRegion( _
                ByVal recalculate As Boolean)
Const ProcName As String = "GOSP_PaintRegion"

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

On Error GoTo Err

mData.RecalcBoundaryForObjectsInSeries SeriesID

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub GOSP_RemoveObject( _
                ByVal handle As Long)
Const ProcName As String = "GOSP_RemoveObject"

On Error GoTo Err

RemoveGraphicObjectEx mData.GetGraphicObject(handle), True, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' mBaseStyle Event Handlers
'================================================================================

Private Sub mBaseStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mBaseStyle_PropertyChanged"

On Error GoTo Err

handleStylePropertyChanged ev

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mViewport Event Handlers
'================================================================================

Private Sub mViewport_BoundaryChanged()
Const ProcName As String = "mViewport_BoundaryChanged"

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

On Error GoTo Err

handleStylePropertyChanged ev

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mYAxisTextStyle Event Handlers
'================================================================================

Private Sub mYAxisTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mYAxisTextStyle_PropertyChanged"

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

On Error GoTo Err

If IsPropertySet(AutoscalingIsSet) Then
    Autoscaling = mAutoscaling
ElseIf Not mStyle Is Nothing Then
    Autoscaling = mStyle.Autoscaling
Else
    Autoscaling = mBaseStyle.Autoscaling
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Autoscaling(ByVal Value As Boolean)
Const ProcName As String = "Autoscaling"

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

On Error GoTo Err

If IsPropertySet(BackGradientFillColorsIsSet) Then
    BackColor = mBackGradientFillColors(0)
ElseIf Not mStyle Is Nothing Then
    BackColor = mStyle.BackColor
Else
    BackColor = mBaseStyle.BackColor
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BackColor(ByVal Value As Long)
Const ProcName As String = "BackColor"

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

On Error GoTo Err

If IsPropertySet(BackGradientFillColorsIsSet) Then
    BackGradientFillColors = mBackGradientFillColors
ElseIf Not mStyle Is Nothing Then
    BackGradientFillColors = mStyle.BackGradientFillColors
Else
    BackGradientFillColors = mBaseStyle.BackGradientFillColors
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BackGradientFillColors(ByRef Value() As Long)
Dim ar() As Long
Const ProcName As String = "BackGradientFillColors"

On Error GoTo Err

ar = Value
mBackGradientFillColors = ar
setPropertyFlag BackGradientFillColorsIsSet
setBackGradientFillColors

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let BaseStyle( _
                ByVal Value As ChartRegionStyle)
Const ProcName As String = "BaseStyle"

On Error GoTo Err

If Value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must not be Nothing"
End If

Set mBaseStyle = Value

If mViewport Is Nothing Then Exit Property

processStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Bottom() As Double
Const ProcName As String = "Bottom"

On Error GoTo Err

Bottom = mViewport.Boundary.Bottom

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Bottom(ByVal Value As Double)
Const ProcName As String = "Bottom"

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

On Error GoTo Err

Set Canvas = mViewport.Canvas

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let CrosshairLineStyle(ByVal Value As LineStyle)
Dim lUseDeferredPainting As Boolean

Const ProcName As String = "CrosshairLineStyle"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot set CrosshairLineStyle for an axis region"
End If

lUseDeferredPainting = mUseDeferredPainting
mUseDeferredPainting = False
If Not Value Is Nothing Then
    mCrosshairLineHoriz.LocalStyle = Value
    mCrosshairLineVert.LocalStyle = Value
Else
    mCrosshairLineHoriz.LocalStyle = gDefaultLineStyle
    mCrosshairLineVert.LocalStyle = gDefaultLineStyle
End If
mUseDeferredPainting = lUseDeferredPainting

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal Value As Boolean)
Const ProcName As String = "CursorSnapsToTickBoundaries"

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

On Error GoTo Err

If IsPropertySet(CursorSnapsToTickBoundariesIsSet) Then
    CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
ElseIf Not mStyle Is Nothing Then
    CursorSnapsToTickBoundaries = mStyle.CursorSnapsToTickBoundaries
Else
    CursorSnapsToTickBoundaries = mBaseStyle.CursorSnapsToTickBoundaries
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Divider(ByVal Value As PictureBox)
Const ProcName As String = "Divider"
On Error GoTo Err

Set mDivider = Value

setDividerColors
PaintDivider

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property


Public Property Get GridlineSpacingY() As Double
Const ProcName As String = "GridlineSpacingY"

On Error GoTo Err

If IsPropertySet(GridlineSpacingYIsSet) Then
    GridlineSpacingY = mGridlineSpacingY
ElseIf Not mStyle Is Nothing Then
    GridlineSpacingY = mStyle.GridlineSpacingY
Else
    GridlineSpacingY = mBaseStyle.GridlineSpacingY
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let GridlineSpacingY(ByVal Value As Double)
Const ProcName As String = "GridlineSpacingY"

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

On Error GoTo Err

If IsPropertySet(GridLineStyleIsSet) Then
    Set GridLineStyle = mGridLineStyle
ElseIf Not mStyle Is Nothing Then
    Set GridLineStyle = mStyle.GridLineStyle
Else
    Set GridLineStyle = mBaseStyle.GridLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let GridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "GridLineStyle"

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

On Error GoTo Err

If IsPropertySet(GridTextStyleIsSet) Then
    Set GridTextStyle = mGridTextStyle
ElseIf Not mStyle Is Nothing Then
    Set GridTextStyle = mStyle.GridTextStyle
Else
    Set GridTextStyle = mBaseStyle.GridTextStyle
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let GridTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "GridTextStyle"

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

On Error GoTo Err

If IsPropertySet(HasGridIsSet) Then
    HasGrid = mHasGrid
ElseIf Not mStyle Is Nothing Then
    HasGrid = mStyle.HasGrid
Else
    HasGrid = mBaseStyle.HasGrid
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HasGrid(ByVal Value As Boolean)
Const ProcName As String = "HasGrid"

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

On Error GoTo Err

If IsPropertySet(HasGridTextProperty) Then
    HasGridText = mHasGridText
ElseIf Not mStyle Is Nothing Then
    HasGridText = mStyle.HasGridText
Else
    HasGridText = mBaseStyle.HasGridText
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HasGridText(ByVal Value As Boolean)
Const ProcName As String = "HasGridText"

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

On Error GoTo Err

Height = mViewport.Height

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IntegerYScale(ByVal Value As Boolean)
Const ProcName As String = "IntegerYScale"

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

On Error GoTo Err

If IsPropertySet(IntegerYScaleIsSet) Then
    IntegerYScale = mIntegerYScale
ElseIf Not mStyle Is Nothing Then
    IntegerYScale = mStyle.IntegerYScale
Else
    IntegerYScale = mBaseStyle.IntegerYScale
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get IsCleared() As Boolean
IsCleared = mIsCleared
End Property

Friend Property Let IsDrawingEnabled(ByVal Value As Boolean)
Const ProcName As String = "IsDrawingEnabled"

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

On Error GoTo Err

Left = mViewport.Left

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Left(ByVal Value As Double)
Const ProcName As String = "Left"

On Error GoTo Err

mViewport.Left = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MinimumHeight() As Double
Const ProcName As String = "MinimumHeight"

On Error GoTo Err

If IsPropertySet(MinimumHeightIsSet) Then
    MinimumHeight = mMinimumHeight
ElseIf Not mStyle Is Nothing Then
    MinimumHeight = mStyle.MinimumHeight
Else
    MinimumHeight = mBaseStyle.MinimumHeight
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let MinimumHeight(ByVal Value As Double)
Const ProcName As String = "MinimumHeight"

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

Public Property Get PointerDiscColor() As Long
PointerDiscColor = mPointerDiscColor
End Property

Friend Property Let PointerDiscColor(ByVal Value As Long)
Const ProcName As String = "PointerDiscColor"

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

mHorizontalGrid.PriceFormatter = Value
doRepaint

End Property

Public Property Get PriceFormatter() As IPriceFormatter
Set PriceFormatter = mHorizontalGrid.PriceFormatter
End Property

Public Property Get SessionEndGridLineStyle() As LineStyle
Const ProcName As String = "SessionEndGridLineStyle"

On Error GoTo Err

If IsPropertySet(SessionEndGridLineStyleIsSet) Then
    Set SessionEndGridLineStyle = mSessionEndGridlineStyle
ElseIf Not mStyle Is Nothing Then
    Set SessionEndGridLineStyle = mStyle.SessionEndGridLineStyle
Else
    Set SessionEndGridLineStyle = mBaseStyle.SessionEndGridLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let SessionEndGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionEndGridLineStyle"

On Error GoTo Err

If mSessionEndGridlineStyle Is Value And IsPropertySet(SessionEndGridLineStyleIsSet) Then Exit Property
Set mSessionEndGridlineStyle = Value
If mSessionEndGridlineStyle Is Nothing Then
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

On Error GoTo Err

If IsPropertySet(SessionStartGridLineStyleIsSet) Then
    Set SessionStartGridLineStyle = mSessionStartGridlineStyle
ElseIf Not mStyle Is Nothing Then
    Set SessionStartGridLineStyle = mStyle.SessionStartGridLineStyle
Else
    Set SessionStartGridLineStyle = mBaseStyle.SessionStartGridLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let SessionStartGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionStartGridLineStyle"

On Error GoTo Err

If mSessionStartGridlineStyle Is Value And IsPropertySet(SessionStartGridLineStyleIsSet) Then Exit Property
Set mSessionStartGridlineStyle = Value
If mSessionStartGridlineStyle Is Nothing Then
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

On Error GoTo Err

Set mStyle = Value

If mViewport Is Nothing Then Exit Property

processStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Title() As Text
Const ProcName As String = "Title"

On Error GoTo Err

'If IsXAxisRegion Or IsYAxisRegion Then
'    Err.Raise ErrorCodes.ErrIllegalStateException, _
'            ProjectName & "." & ModuleName & ":" & ProcName, _
'            "An axis region does not have a title"
'End If

If mTitle Is Nothing Then setupTitle
Set Title = mTitle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Top() As Double
Const ProcName As String = "Top"

On Error GoTo Err

Top = mViewport.Top

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Top(ByVal Value As Double)
Const ProcName As String = "Top"

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

On Error GoTo Err

Set mYAxisRegion = Value

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

mHorizontalGrid.YAxisRegion = mYAxisRegion
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

On Error GoTo Err

If IsPropertySet(YAxisTextStyleIsSet) Then
    Set YAxisTextStyle = mYAxisTextStyle
ElseIf Not mStyle Is Nothing Then
    Set YAxisTextStyle = mStyle.YAxisTextStyle
Else
    Set YAxisTextStyle = mBaseStyle.YAxisTextStyle
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YAxisTextStyle"

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

On Error GoTo Err

If IsPropertySet(YCursorTextStyleIsSet) Then
    Set YCursorTextStyle = mYCursorTextStyle
ElseIf Not mStyle Is Nothing Then
    Set YCursorTextStyle = mStyle.YCursorTextStyle
Else
    Set YCursorTextStyle = mBaseStyle.YCursorTextStyle
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YCursorTextStyle"

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

On Error GoTo Err

YScaleGridSpacing = mHorizontalGridSpacer.YScaleGridSpacing

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YScaleQuantum(ByVal Value As Double)
Const ProcName As String = "YScaleQuantum"

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

On Error GoTo Err

If IsPropertySet(YScaleQuantumIsSet) Then
    YScaleQuantum = mYScaleQuantum
ElseIf Not mStyle Is Nothing Then
    YScaleQuantum = mStyle.YScaleQuantum
Else
    YScaleQuantum = mBaseStyle.YScaleQuantum
End If

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

On Error GoTo Err

mData.AddPeriod pPeriodNumber, pTimestamp
If CDbl(pTimestamp) <> 0 Then setVerticalGridLineAndText pTimestamp, pPeriodNumber, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub AddSessionEndGridline( _
                ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date)
               
Const ProcName As String = "AddSessionEndGridline"

On Error GoTo Err

addStyledVerticalGridline pPeriodNumber, pTimestamp, mSessEndGridlineStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub AddSessionStartGridline( _
                ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date)
Const ProcName As String = "AddSessionStartGridline"

On Error GoTo Err

addStyledVerticalGridline pPeriodNumber, pTimestamp, mSessStartGridlineStyle

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

On Error GoTo Err

deSelectObjectOrSeries mSelectedObject
Set mSelectedObject = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"

On Error GoTo Err

mPropertyOverrideFlags = 0
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Click()
Const ProcName As String = "Click"

On Error GoTo Err

clickSelectedObjectOrSeries
RaiseEvent Click

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub DblCLick()
Const ProcName As String = "DblCLick"

On Error GoTo Err

dblClickSelectedObjectOrSeries
RaiseEvent DblCLick

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub DrawCursor(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "DrawCursor"

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

On Error GoTo Err

FormatYValue = mPriceFormatter.FormatPrice(Y)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub HideGrid()
Const ProcName As String = "HideGrid"

On Error GoTo Err

If mHorizontalGrid.Hidden Then Exit Sub

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

On Error GoTo Err

mName = Name

Set mChartRef = CreateWeakReference(pChart)
Set mViewport = New Viewport
mViewport.Region = Me
mViewport.Canvas = pCanvas
mViewport.MousePointer = vbDefault

mRegionType = pRegionType

mData.Initialise pChart, Me, mGraphObjSeriesCollection

setupGrid
setupCrosshairs
processStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function IsGridHidden() As Boolean
Const ProcName As String = "IsGridHidden"

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

Friend Sub PaintDivider()
Const ProcName As String = "PaintDivider"
On Error GoTo Err

If mDivider Is Nothing Then Exit Sub

setDividerGradientFillArea
mDividerGradientFill.Fill mDivider.hDC
mDivider.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RepaintAreas( _
                ByRef pAreas() As TRectangle, _
                ByVal pCount As Long)
                
Const ProcName As String = "RepaintAreas"
Dim i As Long


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
On Error GoTo Err

If Not mSelectedObject Is Nothing Then
    If graphObj Is mSelectedObject Then ClearSelection
End If

If removeFromSeries Then mGraphObjSeriesCollection.Item(graphObj.SeriesID).RemoveObject graphObj

If undraw Then UndrawArea graphObj.BoundingRectangleInViewport, graphObj

mData.RemoveGraphicObject graphObj

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveGraphicObjectSeries( _
                ByVal pSeries As IGraphicObjectSeries)
Const ProcName As String = "RemoveGraphicObjectSeries"

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
    
mData.RedrawArea periodrect, LayerNumbers.LayerMinVisible

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SetPeriodsInView( _
                ByVal minValue As Single, _
                ByVal maxValue As Single)
Const ProcName As String = "SetPeriodsInView"

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

On Error GoTo Err

setViewportVerticalBounds lowValue, highValue

doCalculateAndRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setViewportVerticalBounds( _
                ByRef Bottom As Double, _
                ByRef Top As Double)
    
Const ProcName As String = "setViewportVerticalBounds"

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

Public Sub ShowGrid()
Const ProcName As String = "ShowGrid"

On Error GoTo Err

If Not mHorizontalGrid.Hidden Then Exit Sub

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

On Error GoTo Err

' redraw objects in the undrawn area, preserving layering
If mIsDrawingEnabled And undrawnArea.isValid Then
    mViewport.UndrawArea undrawnArea
    RedrawArea undrawnArea, LayerMinVisible, objToExclude
    
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Friend Sub addStyledVerticalGridline( _
                ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date, _
                ByVal pStyle As LineStyle)
               
Const ProcName As String = "addStyledVerticalGridline"

On Error GoTo Err

Dim lLine As ChartSkil26.Line

Set lLine = setVerticalGridLineAndText(pTimestamp, pPeriodNumber, True)
If Not lLine Is Nothing Then lLine.LocalStyle = pStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addVerticalGridText(ByRef pLabels As PeriodLabels)
Const ProcName As String = "addVerticalGridText"

On Error GoTo Err

Dim vGridText As Text
Dim vGridTextCoarse As Text

On Error Resume Next
Set vGridText = mXGridTexts.Item(pLabels.Key)
On Error GoTo Err
If vGridText Is Nothing Then
    Set vGridText = mXGridTexts.Add(pLabels.Key)
    vGridText.Position = gNewPoint(pLabels.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsDistance)
    vGridText.Text = pLabels.Label
End If

On Error Resume Next
Set vGridTextCoarse = mXGridTextsCoarse.Item(pLabels.CoarseKey)
On Error GoTo Err
If vGridTextCoarse Is Nothing Then
    Set vGridTextCoarse = mXGridTextsCoarse.Add(pLabels.CoarseKey)
    vGridTextCoarse.Position = gNewPoint(pLabels.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsCounterDistance)
    vGridTextCoarse.Text = pLabels.CoarseLabel
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function addVerticalGridLine(ByRef pLabels As PeriodLabels) As ChartSkil26.Line
Dim vGridLine As ChartSkil26.Line

Const ProcName As String = "addVerticalGridLine"

On Error GoTo Err

On Error Resume Next
Set vGridLine = mXGridLines.Item(pLabels.Key)
On Error GoTo Err
If vGridLine Is Nothing Then
    Set vGridLine = mXGridLines.Add(pLabels.Key)
    vGridLine.FixedX = False
    vGridLine.FixedY = True
    vGridLine.SetPosition gNewPoint(pLabels.PeriodNumber - 0.5, _
                                0, _
                                CoordsLogical, _
                                CoordsRelative), _
                            gNewPoint(pLabels.PeriodNumber - 0.5, _
                                100, _
                                CoordsLogical, _
                                CoordsRelative)
Else
    If vGridLine.Point1.XLogical(mViewport.Region) > pLabels.PeriodNumber Then
        vGridLine.SetPosition gNewPoint(pLabels.PeriodNumber - 0.5, _
                                    0, _
                                    CoordsLogical, _
                                    CoordsRelative), _
                                gNewPoint(pLabels.PeriodNumber - 0.5, _
                                    100, _
                                    CoordsLogical, _
                                    CoordsRelative)
    End If
End If
Set addVerticalGridLine = vGridLine

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub calculateVerticalBounds()
Dim scaleMax As Double
Dim scaleMin As Double

Const ProcName As String = "calculateVerticalBounds"

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

On Error GoTo Err

Set chartObj = mChartRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub ClearProperty( _
                ByVal flag As PropertyOverrideFlags)
Const ProcName As String = "ClearProperty"

On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function clickSelectedObjectOrSeries() As Boolean
Const ProcName As String = "clickSelectedObjectOrSeries"

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

On Error GoTo Err

drawAreas(0) = RectIntersection(mViewport.Boundary, areaToRedraw)

If graphObj.Layer <> LayerInvisible Then graphObj.Draw drawAreas

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawSelectionPointer(ByVal X As Single, ByVal Y As Single)

Const ProcName As String = "drawSelectionPointer"

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

Private Function getPeriodLabels(ByVal pTimestamp As Date, _
                            ByVal pPeriodNumber As Long, _
                            ByVal pIgnoreVerticalGridTimePeriod As Boolean) As PeriodLabels
Const ProcName As String = "getPeriodLabels"

Dim lGridLineTime As Date

On Error GoTo Err

With getPeriodLabels
    .PeriodNumber = pPeriodNumber
    If pIgnoreVerticalGridTimePeriod Then
        ' just draw a vertical gridline at the specified period
        .Key = Format(pTimestamp, "yyyymmddhhnnss")
        .Label = FormatDateTime(pTimestamp, vbShortTime)
        .CoarseKey = Format(pTimestamp, "yyyymmdd")
        .CoarseLabel = Format(pTimestamp, "d Mmm yy")
        Exit Function
    End If
    
    If mVerticalGridTimePeriod Is Nothing Then
        ' just draw a vertical gridline every 10 bars
        If pPeriodNumber Mod 10 = 1 Then
            .Key = Format(pTimestamp, "yyyymmddhhnnss")
            .Label = FormatDateTime(pTimestamp, vbShortTime)
            .CoarseKey = Format(pTimestamp, "yyyymmdd")
            .CoarseLabel = Format(pTimestamp, "d Mmm yy")
        End If
        Exit Function
    End If
    
    lGridLineTime = BarStartTime(pTimestamp, mVerticalGridTimePeriod, mSessionStartTime)
        
    Select Case mVerticalGridTimePeriod.Units
    Case TimePeriodSecond
        .Key = Format(lGridLineTime, "yyyymmddhhnnss")
        .Label = FormatDateTime(lGridLineTime, vbLongTime)
        .CoarseKey = Format(lGridLineTime, "yyyymmdd")
        .CoarseLabel = Format(lGridLineTime, "d Mmm yy")
    Case TimePeriodMinute
        .Key = Format(lGridLineTime, "yyyymmddhhnn")
        .Label = FormatDateTime(lGridLineTime, vbShortTime)
        .CoarseKey = Format(lGridLineTime, "yyyymmdd")
        .CoarseLabel = Format(lGridLineTime, "d Mmm yy")
    Case TimePeriodHour
        .Key = Format(lGridLineTime, "yyyymmddhh")
        .Label = FormatDateTime(lGridLineTime, vbShortTime)
        .CoarseKey = Format(lGridLineTime, "yyyymmdd")
        .CoarseLabel = Format(lGridLineTime, "d Mmm yy")
    Case TimePeriodDay
        .Key = Format(lGridLineTime, "yyyymmdd")
        .Label = Format(lGridLineTime, "d")
        .CoarseKey = Format(lGridLineTime, "yyyymm")
        .CoarseLabel = Format(lGridLineTime, "Mmm yy")
    Case TimePeriodWeek
        .Key = Format(lGridLineTime, "yyyymmdd")
        .Label = Format(lGridLineTime, "d Mmm")
        .CoarseKey = Format(lGridLineTime, "yyyy")
        .CoarseLabel = Format(lGridLineTime, "yyyy")
    Case TimePeriodMonth
        .Key = Format(lGridLineTime, "yyyymm")
        .Label = Format(lGridLineTime, "Mmm")
        .CoarseKey = Format(lGridLineTime, "yyyy")
        .CoarseLabel = Format(lGridLineTime, "yyyy")
    Case TimePeriodYear
        .Key = Format(lGridLineTime, "yyyy")
        .Label = Format(lGridLineTime, "YYYY")
        .CoarseKey = "$"
        .CoarseLabel = ""
    Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
        If pPeriodNumber Mod 10 = 1 Then
            .Key = Format(lGridLineTime, "yyyymmddhhnnss")
            .Label = FormatDateTime(lGridLineTime, vbLongTime)
            .CoarseKey = Format(lGridLineTime, "yyyymmdd")
            .CoarseLabel = Format(lGridLineTime, "d Mmm yy")
        End If
    End Select
End With

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub handleStylePropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "handleStylePropertyChanged"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub hideCursor()
Dim lPointerStyle As PointerStyles

Const ProcName As String = "hideCursor"

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

mPriceFormatter.YScaleGridSpacing = mHorizontalGridSpacer.Calculate(mViewport.Height, mViewport.HeightCm)

mHorizontalGrid.Generate mViewport.Bottom, mViewport.Top, mHorizontalGridSpacer.YScaleGridSpacing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintPerformanceText( _
                ByVal numberOfObjectsRepainted As Long, _
                ByVal elapsedTimeMillisecs As Single)
Dim perfText As Text
Const ProcName As String = "paintPerformanceText"

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

On Error GoTo Err

setBackColor
setBackGradientFillColors
setHasGrid
setHasGridText
setGridTextStyle
setGridlineSpacingY
setSessionEndGridLineStyle
setSessionStartGridLineStyle
If Not IsXAxisRegion And Not IsYAxisRegion Then
    setAutoscaling
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

On Error GoTo Err

' redraw objects in the undrawn area, preserving layering
If mIsDrawingEnabled And areaToRedraw.isValid Then
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

On Error GoTo Err

mViewport.BackColor = BackColor
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setBackGradientFillColors()
Const ProcName As String = "setBackGradientFillColors"

On Error GoTo Err

mViewport.GradientFillColors = BackGradientFillColors
doRepaint

setDividerColors
PaintDivider

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setDividerColors()
If Not mDivider Is Nothing Then mDividerGradientFill.Colors = BackGradientFillColors
End Sub

Private Sub setDividerGradientFillArea()
Dim lRect As GDI_RECT

Const ProcName As String = "setDividerGradientFillArea"
On Error GoTo Err

lRect.Right = mDivider.Width / Screen.TwipsPerPixelX
lRect.Bottom = mDivider.Height / Screen.TwipsPerPixelY
mDividerGradientFill.Area = lRect

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setGridlineSpacingY()
Const ProcName As String = "setGridlineSpacingY"

On Error GoTo Err

mHorizontalGridSpacer.GridlineSpacingY = GridlineSpacingY
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setGridLineStyle()
Const ProcName As String = "setGridLineStyle"

On Error GoTo Err

mHorizontalGrid.GridLineStyle = GridLineStyle
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setGridTextStyle()
Const ProcName As String = "setGridTextStyle"

On Error GoTo Err

setHasGridText

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setHasGrid()
Const ProcName As String = "setHasGrid"

On Error GoTo Err

If HasGrid Then
    mHorizontalGrid.GridLineStyle = GridLineStyle
    mXGridLines.Style = GridLineStyle
    mXGridLines.Layer = LayerGrid
Else
    mXGridLines.Layer = LayerInvisible
End If
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setHasGridText()

Const ProcName As String = "setHasGridText"

On Error GoTo Err

If HasGridText Then
    mXGridTexts.Layer = LayerGridText + 1
    mXGridTextsCoarse.Layer = LayerGridText
    mXGridTexts.Style = GridTextStyle
    mXGridTextsCoarse.Style = GridTextStyle
Else
    mXGridTexts.Layer = LayerInvisible
    mXGridTextsCoarse.Layer = LayerInvisible
End If
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setIntegerYScale()
Const ProcName As String = "setIntegerYScale"

On Error GoTo Err

mPriceFormatter.IntegerYScale = IntegerYScale
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setMinimumHeight()
Const ProcName As String = "setMinimumHeight"

On Error GoTo Err

doScaleAndRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setMouseIcon( _
                ByVal PointerStyle As PointerStyles)
Const ProcName As String = "setMouseIcon"

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

On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSessionEndGridLineStyle()
Const ProcName As String = "setSessionEndGridLineStyle"

On Error GoTo Err

SessionEndGridLineStyle.SyncTo mSessEndGridlineStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSessionStartGridLineStyle()
Const ProcName As String = "setSessionStartGridLineStyle"

On Error GoTo Err

SessionStartGridLineStyle.SyncTo mSessStartGridlineStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupCrosshairs()
Const ProcName As String = "setupCrosshairs"

On Error GoTo Err

Set mCrosshairLineHoriz = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineHoriz.LocalStyle = gDefaultLineStyle
mCrosshairLineHoriz.Extended = True
mCrosshairLineHoriz.IncludeInAutoscale = False
mCrosshairLineHoriz.Layer = LayerPointer

Set mCrosshairLineVert = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineVert.LocalStyle = gDefaultLineStyle
mCrosshairLineVert.Extended = True
mCrosshairLineVert.IncludeInAutoscale = False
mCrosshairLineVert.Layer = LayerPointer

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupTitle()
Const ProcName As String = "setupTitle"

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

Private Sub setupGrid()
Const ProcName As String = "setupGrid"
On Error GoTo Err

Set mHorizontalGrid = New HorizontalGrid
mHorizontalGrid.Initialise Me

Set mXGridLines = Me.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)

Set mPriceFormatter = New PriceFormatter
mHorizontalGrid.PriceFormatter = mPriceFormatter

Set mHorizontalGridSpacer = New HorizontalGridSpacer

Set mXGridTexts = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText + 1)
mXGridTexts.Align = AlignBoxBottomCentre
mXGridTexts.FixedX = False
mXGridTexts.FixedY = True
mXGridTexts.Extended = True

Set mXGridTextsCoarse = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
mXGridTextsCoarse.Align = AlignBoxTopCentre
mXGridTextsCoarse.FixedX = False
mXGridTextsCoarse.FixedY = True
mXGridTextsCoarse.Extended = True

Set mSessEndGridlineStyle = gDefaultLineStyle.clone
Set mSessStartGridlineStyle = gDefaultLineStyle.clone

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function setVerticalGridLineAndText(ByVal pTimestamp As Date, _
                            ByVal pPeriodNumber As Long, _
                            ByVal pIgnoreVerticalGridTimePeriod As Boolean) As ChartSkil26.Line
Const ProcName As String = "setVerticalGridLineAndText"

On Error GoTo Err

Dim lPeriodLabels As PeriodLabels

If CDbl(pTimestamp) = 0 Then Exit Function

lPeriodLabels = getPeriodLabels(pTimestamp, pPeriodNumber, pIgnoreVerticalGridTimePeriod)
Set setVerticalGridLineAndText = addVerticalGridLine(lPeriodLabels)
addVerticalGridText lPeriodLabels

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setYAxisTextStyle()
Const ProcName As String = "setYAxisTextStyle"

On Error GoTo Err

mHorizontalGrid.TextStyle = YAxisTextStyle
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setYCursorTextStyle()
Const ProcName As String = "setYCursorTextStyle"

On Error GoTo Err

If Not mYCursorText Is Nothing Then mYCursorText.LocalStyle = YCursorTextStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setYScaleQuantum()
Const ProcName As String = "setYScaleQuantum"

On Error GoTo Err

mPriceFormatter.YScaleQuantum = YScaleQuantum

mPriceFormatter.YScaleQuantum = YScaleQuantum
mHorizontalGridSpacer.YScaleQuantum = YScaleQuantum
doRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub showCursor( _
                ByVal X As Single, _
                ByVal Y As Single)
Dim PointerStyle As PointerStyles

Const ProcName As String = "showCursor"

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


