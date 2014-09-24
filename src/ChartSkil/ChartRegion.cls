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
    AutoscalingIsSet = &H1&
    XGridTextStyleIsSet = &H2&
    HasXGridIsSet = &H4&
    HasXGridTextIsSet = &H8&
    IntegerYScaleIsSet = &H10&
    YScaleQuantumIsSet = &H20&
    YGridlineSpacingIsSet = &H40&
    MinimumHeightIsSet = &H80&
    CursorSnapsToTickBoundariesIsSet = &H100&
    BackGradientFillColorsIsSet = &H200&
    XGridLineStyleIsSet = &H400&
    YCursorTextStyleIsSet = &H800&
    YAxisTextStyleIsSet = &H1000&
    SessionEndGridLineStyleIsSet = &H2000&
    SessionStartGridLineStyleIsSet = &H4000&
    HasYGridIsSet = &H8000&
    HasYGridTextIsSet = &H10000
    CursorTextModeIsSet = &H20000
    CursorTextStyleIsSet = &H40000
    YGridLineStyleIsSet = &H80000
    YGridTextStyleIsSet = &H100000
    XGridTextPositionIsSet = &H200000
    YGridTextPositionIsSet = &H400000
    XCursorTextPositionIsSet = &H800000
    YCursorTextPositionIsSet = &H1000000
    XCursorTextStyleIsSet = &H2000000
    CursorTextPositionIsSet = &H4000000
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

Private mPeriods                                    As Periods
Private mData                                       As ChartRegionData
Attribute mData.VB_VarHelpID = -1
Private mGraphObjSeriesCollection                   As GraphObjSeriesCollection

Private mHandle                                     As Long

Private mActualHeight                               As Long

Private mIsCleared                                  As Boolean

Private mName                                       As String

Private mRegionType                                 As RegionTypes

Private WithEvents mBaseStyle                       As ChartRegionStyle
Attribute mBaseStyle.VB_VarHelpID = -1
Private WithEvents mStyle                           As ChartRegionStyle
Attribute mStyle.VB_VarHelpID = -1

Private mTitle                                      As Text

Private mViewport                                   As ViewPort
Attribute mViewport.VB_VarHelpID = -1
Private mIsDrawingEnabled                           As Boolean

Private mMinInViewPeriod                            As Long
Private mMaxInViewPeriod                            As Long

Private mDrawn                                      As Boolean

Private mGraphObjServiceProvider                    As GraphObjServiceProvider
Attribute mGraphObjServiceProvider.VB_VarHelpID = -1

Private mHorizontalGrid                             As HorizontalGrid
Private mVerticalGrid                               As VerticalGrid

Private mPointerController                          As PointerController

Private mYAxisRegion                                As ChartRegion

Private mCursorText                                 As Text
Private mXCursorText                                As Text
Private mYCursorText                                As Text

Private mPercentHeight                              As Double
Private mMinimumPercentHeight                       As Double
Private mYAxisPosition                              As Long

Private mPerformanceTextVisible                     As Boolean

Private mPointerDiscColor                           As Long

Private mCurrX                                      As Single
Private mCurrY                                      As Single

Private mUseDeferredPainting                        As Boolean
Private mDeferredPainter                            As DeferredRegionPainter

Private mSelectedObject                             As IGraphicObject

' these fields override corresponding items in the Style
Private mAutoscaling                                As Boolean
Private mHasXGrid                                   As Boolean
Private mHasYGrid                                   As Boolean
Private mHasXGridText                               As Boolean
Private mHasYGridText                               As Boolean

Private mXGridTextPosition                          As XGridTextPositions
Private mYGridTextPosition                          As YGridTextPositions

Private mCursorTextMode                             As CursorTextModes
Private mCursorTextPosition                         As CursorTextPositions

Private mXCursorTextPosition                        As CursorTextPositions

Private mYCursorTextPosition                        As CursorTextPositions

Private mIntegerYScale                              As Boolean
Private mYScaleQuantum                              As Double
Private mYGridlineSpacing                           As Double
Private mMinimumHeight                              As Double
Private mCursorSnapsToTickBoundaries                As Boolean
Private mBackGradientFillColors()                   As Long

Private WithEvents mXGridLineStyle                  As LineStyle
Attribute mXGridLineStyle.VB_VarHelpID = -1
Private WithEvents mXGridTextStyle                  As TextStyle
Attribute mXGridTextStyle.VB_VarHelpID = -1
Private WithEvents mYGridLineStyle                  As LineStyle
Attribute mYGridLineStyle.VB_VarHelpID = -1
Private WithEvents mYGridTextStyle                  As TextStyle
Attribute mYGridTextStyle.VB_VarHelpID = -1
Private WithEvents mSessionEndGridLineStyle         As LineStyle
Attribute mSessionEndGridLineStyle.VB_VarHelpID = -1
Private WithEvents mSessionStartGridLineStyle       As LineStyle
Attribute mSessionStartGridLineStyle.VB_VarHelpID = -1
Private WithEvents mCursorTextStyle                 As TextStyle
Attribute mCursorTextStyle.VB_VarHelpID = -1
Private WithEvents mXCursorTextStyle                As TextStyle
Attribute mXCursorTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle                As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

Private mPropertyOverrideFlags                      As PropertyOverrideFlags

Private mPaintingInProgress                         As Boolean

Private mPriceFormatter                             As IPriceFormatter

Private mHorizontalGridSpacer                       As HorizontalGridSpacer

Private mDivider                                    As PictureBox
Private mDividerGradientFill                        As GradientFill

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Dim i As Long

Set mData = New ChartRegionData
Set mGraphObjSeriesCollection = New GraphObjSeriesCollection
mGraphObjSeriesCollection.Initialise Me

Set mBaseStyle = gDefaultChartDataRegionStyle

mUseDeferredPainting = True
Set mDeferredPainter = New DeferredRegionPainter
mDeferredPainter.Initialise Me

Set mGraphObjServiceProvider = New GraphObjServiceProvider
mGraphObjServiceProvider.Initialise Me, mData

Set mDividerGradientFill = New GradientFill
mDividerGradientFill.Color = &HD0D0D0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
If Not mViewport Is Nothing Then mViewport.Finish
Set mViewport = Nothing

gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Proc", pModName:=ModuleName, pMsg:="ChartRegion terminated (" & gRegionTypeToString(mRegionType) & ")"
Debug.Print "ChartRegion terminated (" & gRegionTypeToString(mRegionType) & ")"
End Sub

'================================================================================
' mBaseStyle Event Handlers
'================================================================================

Private Sub mBaseStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mBaseStyle_PropertyChanged"
On Error GoTo Err

handleStylePropertyChanged ev

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mCursorTextStyle Event Handlers
'================================================================================

Private Sub mCursorTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mCursorTextStyle_PropertyChanged"
On Error GoTo Err

setCursorTextStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mSessionEndGridLineStyle Event Handlers
'================================================================================

Private Sub mSessionEndGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
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

Private Sub mSessionStartGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
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

Private Sub mStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mStyle_PropertyChanged"
On Error GoTo Err

handleStylePropertyChanged ev

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mXCursorTextStyle Event Handlers
'================================================================================

Private Sub mXCursorTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mXCursorTextStyle_PropertyChanged"
On Error GoTo Err

setXCursorTextStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mXGridLineStyle Event Handlers
'================================================================================

Private Sub mXGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mXGridLineStyle_PropertyChanged"
On Error GoTo Err

setXGridLineStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mXGridTextStyle Event Handlers
'================================================================================

Private Sub mXGridTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mXGridTextStyle_PropertyChanged"
On Error GoTo Err

setXGridTextStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mYCursorTextStyle Event Handlers
'================================================================================

Private Sub mYCursorTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mYCursorTextStyle_PropertyChanged"
On Error GoTo Err

setYCursorTextStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mYGridLineStyle Event Handlers
'================================================================================

Private Sub mYGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mYGridLineStyle_PropertyChanged"
On Error GoTo Err

setYGridLineStyle

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mYGridTextStyle Event Handlers
'================================================================================

Private Sub mYGridTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mYGridTextStyle_PropertyChanged"
On Error GoTo Err

setYGridTextStyle

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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BackGradientFillColors(ByRef Value() As Long)
Const ProcName As String = "BackGradientFillColors"
On Error GoTo Err

Dim ar() As Long
ar = Value
mBackGradientFillColors = ar
setPropertyFlag BackGradientFillColorsIsSet
setBackGradientFillColors

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let BaseStyle( _
                ByVal Value As ChartRegionStyle)
Const ProcName As String = "BaseStyle"
On Error GoTo Err

If Value Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must not be Nothing"

Set mBaseStyle = Value

If mViewport Is Nothing Then Exit Property

processStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Bottom() As Double
Const ProcName As String = "Bottom"
On Error GoTo Err

Bottom = mViewport.Boundary.Bottom

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Bottom(ByVal Value As Double)
Const ProcName As String = "Bottom"
On Error GoTo Err

mViewport.SetVerticalBounds Value, mViewport.Top
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.SetViewportVerticalBounds Value, mYAxisRegion.Top
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Canvas() As Canvas
Const ProcName As String = "Canvas"
On Error GoTo Err

Set Canvas = mViewport.Canvas

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let CrosshairLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

Dim lUseDeferredPainting As Boolean

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set CrosshairLineStyle for an axis region"

lUseDeferredPainting = mUseDeferredPainting
mUseDeferredPainting = False
mPointerController.CrosshairLineStyle = Value
mUseDeferredPainting = lUseDeferredPainting

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal Value As Boolean)
Const ProcName As String = "CursorSnapsToTickBoundaries"
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set snapCursorToTickBoundaries for an axis region"

If mCursorSnapsToTickBoundaries = Value And IsPropertySet(CursorSnapsToTickBoundariesIsSet) Then Exit Property
mCursorSnapsToTickBoundaries = Value
setPropertyFlag CursorSnapsToTickBoundariesIsSet
setCursorSnapsToTickBoundaries

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CursorTextMode() As CursorTextModes
Const ProcName As String = "CursorTextMode"
On Error GoTo Err

If IsPropertySet(CursorTextModeIsSet) Then
    CursorTextMode = mCursorTextMode
ElseIf Not mStyle Is Nothing Then
    CursorTextMode = mStyle.CursorTextMode
Else
    CursorTextMode = mBaseStyle.CursorTextMode
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CursorTextMode(ByVal Value As CursorTextModes)
Const ProcName As String = "CursorTextMode"
On Error GoTo Err

If mCursorTextMode = Value And IsPropertySet(CursorTextModeIsSet) Then Exit Property
mCursorTextMode = Value
setPropertyFlag CursorTextModeIsSet
setCursorTextMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CursorTextPosition() As CursorTextPositions
Const ProcName As String = "CursorTextPosition"
On Error GoTo Err

If IsPropertySet(CursorTextPositionIsSet) Then
    CursorTextPosition = mCursorTextPosition
ElseIf Not mStyle Is Nothing Then
    CursorTextPosition = mStyle.CursorTextPosition
Else
    CursorTextPosition = mBaseStyle.CursorTextPosition
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CursorTextPosition(ByVal Value As CursorTextPositions)
Const ProcName As String = "CursorTextPosition"
On Error GoTo Err

If mCursorTextPosition = Value And IsPropertySet(CursorTextPositionIsSet) Then Exit Property
mCursorTextPosition = Value
setPropertyFlag CursorTextPositionIsSet
setCursorTextPosition mCursorText, CursorTextPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CursorTextStyle() As TextStyle
Const ProcName As String = "CursorTextStyle"
On Error GoTo Err

If IsPropertySet(CursorTextStyleIsSet) Then
    Set CursorTextStyle = mCursorTextStyle
ElseIf Not mStyle Is Nothing Then
    Set CursorTextStyle = mStyle.CursorTextStyle
Else
    Set CursorTextStyle = mBaseStyle.CursorTextStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "CursorTextStyle"
On Error GoTo Err

If mCursorTextStyle Is Value And IsPropertySet(CursorTextStyleIsSet) Then Exit Property
Set mCursorTextStyle = Value
If mCursorTextStyle Is Nothing Then
    ClearProperty CursorTextStyleIsSet
Else
    setPropertyFlag CursorTextStyleIsSet
End If
setCursorTextStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Divider(ByVal Value As PictureBox)
Const ProcName As String = "Divider"
On Error GoTo Err

Set mDivider = Value

setDividerColors
PaintDivider

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property


Public Property Get YGridlineSpacing() As Double
Const ProcName As String = "YGridlineSpacing"
On Error GoTo Err

If IsPropertySet(YGridlineSpacingIsSet) Then
    YGridlineSpacing = mYGridlineSpacing
ElseIf Not mStyle Is Nothing Then
    YGridlineSpacing = mStyle.YGridlineSpacing
Else
    YGridlineSpacing = mBaseStyle.YGridlineSpacing
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YGridlineSpacing(ByVal Value As Double)
Const ProcName As String = "YGridlineSpacing"
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set YGridlineSpacing for an axis region"
YGridlineSpacingEx = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let YGridlineSpacingEx(ByVal Value As Double)
Const ProcName As String = "YGridlineSpacingEx"
On Error GoTo Err

If mYGridlineSpacing = Value And IsPropertySet(YGridlineSpacingIsSet) Then Exit Property
mYGridlineSpacing = Value
setPropertyFlag YGridlineSpacingIsSet
setYGridlineSpacing

If Not mYAxisRegion Is Nothing Then mYAxisRegion.YGridlineSpacingEx = YGridlineSpacing

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let handle(ByVal Value As Long)
mHandle = Value
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get HasXGrid() As Boolean
Const ProcName As String = "HasXGrid"
On Error GoTo Err

If IsPropertySet(HasXGridIsSet) Then
    HasXGrid = mHasXGrid
ElseIf Not mStyle Is Nothing Then
    HasXGrid = mStyle.HasXGrid
Else
    HasXGrid = mBaseStyle.HasXGrid
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HasXGrid(ByVal Value As Boolean)
Const ProcName As String = "HasXGrid"
On Error GoTo Err

If mHasXGrid = Value And IsPropertySet(HasXGridIsSet) Then Exit Property
mHasXGrid = Value
setPropertyFlag HasXGridIsSet
setHasXGrid

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasXGridText() As Boolean
Const ProcName As String = "HasXGridText"
On Error GoTo Err

If IsPropertySet(HasXGridTextIsSet) Then
    HasXGridText = mHasXGridText
ElseIf Not mStyle Is Nothing Then
    HasXGridText = mStyle.HasXGridText
Else
    HasXGridText = mBaseStyle.HasXGridText
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HasXGridText(ByVal Value As Boolean)
Const ProcName As String = "HasXGridText"
On Error GoTo Err

If mHasXGridText = Value And IsPropertySet(HasXGridTextIsSet) Then Exit Property
mHasXGridText = Value
setPropertyFlag HasXGridTextIsSet
setHasXGridText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasYGrid() As Boolean
Const ProcName As String = "HasYGrid"
On Error GoTo Err

If IsPropertySet(HasYGridIsSet) Then
    HasYGrid = mHasYGrid
ElseIf Not mStyle Is Nothing Then
    HasYGrid = mStyle.HasYGrid
Else
    HasYGrid = mBaseStyle.HasYGrid
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HasYGrid(ByVal Value As Boolean)
Const ProcName As String = "HasYGrid"
On Error GoTo Err

If mHasYGrid = Value And IsPropertySet(HasYGridIsSet) Then Exit Property
mHasYGrid = Value
setPropertyFlag HasYGridIsSet
setHasYGrid

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasYGridText() As Boolean
Const ProcName As String = "HasYGridText"
On Error GoTo Err

If IsPropertySet(HasYGridTextIsSet) Then
    HasYGridText = mHasYGridText
ElseIf Not mStyle Is Nothing Then
    HasYGridText = mStyle.HasYGridText
Else
    HasYGridText = mBaseStyle.HasYGridText
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HasYGridText(ByVal Value As Boolean)
Const ProcName As String = "HasYGridText"
On Error GoTo Err

If mHasYGridText = Value And IsPropertySet(HasYGridTextIsSet) Then Exit Property
mHasYGridText = Value
setPropertyFlag HasYGridTextIsSet
setHasYGridText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Height() As Double
Const ProcName As String = "Height"
On Error GoTo Err

Height = mViewport.Height

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IntegerYScale(ByVal Value As Boolean)
Const ProcName As String = "IntegerYScale"
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set integerYScale for an axis region"

IntegerYScaleEx = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let IntegerYScaleEx(ByVal Value As Boolean)
Const ProcName As String = "IntegerYScaleEx"
On Error GoTo Err

If mIntegerYScale = Value And IsPropertySet(IntegerYScaleIsSet) Then Exit Property
mIntegerYScale = Value
setPropertyFlag IntegerYScaleIsSet

setIntegerYScale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get IsCleared() As Boolean
IsCleared = mIsCleared
End Property

Friend Property Get IsDrawingEnabled() As Boolean
IsDrawingEnabled = mIsDrawingEnabled
End Property

Friend Property Let IsDrawingEnabled(ByVal Value As Boolean)
Const ProcName As String = "IsDrawingEnabled"
On Error GoTo Err

If mIsDrawingEnabled = Value Then Exit Property
mIsDrawingEnabled = Value
If Not mIsDrawingEnabled Then
    mPointerController.UseWaitCursor = True
Else
    mPointerController.UseWaitCursor = False
    CalculateAndPaint
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.IsDrawingEnabled = mIsDrawingEnabled
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Left(ByVal Value As Double)
Const ProcName As String = "Left"

On Error GoTo Err

mViewport.Left = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

If MinimumHeight < YScaleQuantum Then MinimumHeight = YScaleQuantum

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let MinimumHeight(ByVal Value As Double)
Const ProcName As String = "MinimumHeight"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set minimumHeight for an axis region"

If mMinimumHeight = Value And IsPropertySet(MinimumHeightIsSet) Then Exit Property
mMinimumHeight = Value
setPropertyFlag MinimumHeightIsSet
setMinimumHeight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Friend Property Get PaintingInProgress() As Boolean
PaintingInProgress = mPaintingInProgress
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

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set PointerDiscColor for an axis region"

mPointerDiscColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerIcon() As IPictureDisp
Set PointerIcon = mPointerController.PointerIcon
End Property

Public Property Let PointerIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "PointerIcon"
On Error GoTo Err

mPointerController.PointerIcon = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerStyle() As PointerStyles
PointerStyle = mPointerController.PointerStyle
End Property

Public Property Let PointerStyle(ByVal Value As PointerStyles)
Const ProcName As String = "PointerStyle"
On Error GoTo Err

mPointerController.PointerStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PriceFormatter(ByVal Value As IPriceFormatter)
Const ProcName As String = "PriceFormatter"
On Error GoTo Err

If Not Value Is Nothing Then
    Set mPriceFormatter = Value
Else
    Set mPriceFormatter = New PriceFormatter
End If

mPriceFormatter.IntegerYScale = IntegerYScale
mPriceFormatter.YScaleQuantum = YScaleQuantum

mHorizontalGrid.PriceFormatter = Value
Paint

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PriceFormatter() As IPriceFormatter
Set PriceFormatter = mHorizontalGrid.PriceFormatter
End Property

Public Property Get SessionEndGridLineStyle() As LineStyle
Const ProcName As String = "SessionEndGridLineStyle"

On Error GoTo Err

If IsPropertySet(SessionEndGridLineStyleIsSet) Then
    Set SessionEndGridLineStyle = mSessionEndGridLineStyle
ElseIf Not mStyle Is Nothing Then
    Set SessionEndGridLineStyle = mStyle.SessionEndGridLineStyle
Else
    Set SessionEndGridLineStyle = mBaseStyle.SessionEndGridLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SessionEndGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionEndGridLineStyle"

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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionStartGridLineStyle() As LineStyle
Const ProcName As String = "SessionStartGridLineStyle"

On Error GoTo Err

If IsPropertySet(SessionStartGridLineStyleIsSet) Then
    Set SessionStartGridLineStyle = mSessionStartGridLineStyle
ElseIf Not mStyle Is Nothing Then
    Set SessionStartGridLineStyle = mStyle.SessionStartGridLineStyle
Else
    Set SessionStartGridLineStyle = mBaseStyle.SessionStartGridLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SessionStartGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionStartGridLineStyle"

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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Title() As Text
Const ProcName As String = "Title"

On Error GoTo Err

If mTitle Is Nothing Then setupTitle
Set Title = mTitle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Top() As Double
Const ProcName As String = "Top"
On Error GoTo Err

Top = mViewport.Top

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Top(ByVal Value As Double)
Const ProcName As String = "Top"
On Error GoTo Err

mViewport.SetVerticalBounds mViewport.Bottom, Value
If Not mYAxisRegion Is Nothing Then mYAxisRegion.SetViewportVerticalBounds mYAxisRegion.Bottom, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ViewPort() As ViewPort
Set ViewPort = mViewport
End Property

Public Property Get Width() As Single
Const ProcName As String = "Width"

On Error GoTo Err

Width = mViewport.Width

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XCursorTextPosition() As CursorTextPositions
Const ProcName As String = "XCursorTextPosition"
On Error GoTo Err

If IsPropertySet(XCursorTextPositionIsSet) Then
    XCursorTextPosition = mXCursorTextPosition
ElseIf Not mStyle Is Nothing Then
    XCursorTextPosition = mStyle.XCursorTextPosition
Else
    XCursorTextPosition = mBaseStyle.XCursorTextPosition
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XCursorTextPosition(ByVal Value As CursorTextPositions)
Const ProcName As String = "XCursorTextPosition"
On Error GoTo Err

If mXCursorTextPosition = Value And IsPropertySet(XCursorTextPositionIsSet) Then Exit Property
mXCursorTextPosition = Value
setPropertyFlag XCursorTextPositionIsSet
setXCursorTextPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XCursorTextStyle() As TextStyle
Const ProcName As String = "XCursorTextStyle"

On Error GoTo Err

If IsPropertySet(XCursorTextStyleIsSet) Then
    Set XCursorTextStyle = mXCursorTextStyle
ElseIf Not mStyle Is Nothing Then
    Set XCursorTextStyle = mStyle.XCursorTextStyle
Else
    Set XCursorTextStyle = mBaseStyle.XCursorTextStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "XCursorTextStyle"

On Error GoTo Err

If mXCursorTextStyle Is Value And IsPropertySet(XCursorTextStyleIsSet) Then Exit Property
Set mXCursorTextStyle = Value
If mXCursorTextStyle Is Nothing Then
    ClearProperty XCursorTextStyleIsSet
Else
    setPropertyFlag XCursorTextStyleIsSet
End If
setXCursorTextStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XGridLineStyle() As LineStyle
Const ProcName As String = "XGridLineStyle"

On Error GoTo Err

If IsPropertySet(XGridLineStyleIsSet) Then
    Set XGridLineStyle = mXGridLineStyle
ElseIf Not mStyle Is Nothing Then
    Set XGridLineStyle = mStyle.XGridLineStyle
Else
    Set XGridLineStyle = mBaseStyle.XGridLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "XGridLineStyle"

On Error GoTo Err

If mXGridLineStyle Is Value And IsPropertySet(XGridLineStyleIsSet) Then Exit Property
Set mXGridLineStyle = Value
If mXGridLineStyle Is Nothing Then
    ClearProperty XGridLineStyleIsSet
Else
    setPropertyFlag XGridLineStyleIsSet
End If
setXGridLineStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XGridTextPosition() As XGridTextPositions
Const ProcName As String = "XGridTextPosition"
On Error GoTo Err

If IsPropertySet(XGridTextPositionIsSet) Then
    XGridTextPosition = mXGridTextPosition
ElseIf Not mStyle Is Nothing Then
    XGridTextPosition = mStyle.XGridTextPosition
Else
    XGridTextPosition = mBaseStyle.XGridTextPosition
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XGridTextPosition(ByVal Value As XGridTextPositions)
Const ProcName As String = "XGridTextPosition"
On Error GoTo Err

If mXGridTextPosition = Value And IsPropertySet(XGridTextPositionIsSet) Then Exit Property
mXGridTextPosition = Value
setPropertyFlag XGridTextPositionIsSet
setXGridTextPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XGridTextStyle() As TextStyle
Const ProcName As String = "XGridTextStyle"

On Error GoTo Err

If IsPropertySet(XGridTextStyleIsSet) Then
    Set XGridTextStyle = mXGridTextStyle
ElseIf Not mStyle Is Nothing Then
    Set XGridTextStyle = mStyle.XGridTextStyle
Else
    Set XGridTextStyle = mBaseStyle.XGridTextStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XGridTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "XGridTextStyle"

On Error GoTo Err

If mXGridTextStyle Is Value And IsPropertySet(XGridTextStyleIsSet) Then Exit Property
Set mXGridTextStyle = Value
If mXGridTextStyle Is Nothing Then
    ClearProperty XGridTextStyleIsSet
Else
    setPropertyFlag XGridTextStyleIsSet
End If
setXGridTextStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Friend Property Let YAxisRegion(ByVal Value As ChartRegion)
Const ProcName As String = "YAxisRegion"
On Error GoTo Err

Set mYAxisRegion = Value

mYAxisRegion.IsDrawingEnabled = mIsDrawingEnabled
mYAxisRegion.YScaleQuantumEx = YScaleQuantum
mYAxisRegion.IntegerYScaleEx = IntegerYScale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get YAxisRegion() As ChartRegion
Set YAxisRegion = mYAxisRegion
End Property

Public Property Get YCursorTextPosition() As CursorTextPositions
Const ProcName As String = "YCursorTextPosition"
On Error GoTo Err

If IsPropertySet(YCursorTextPositionIsSet) Then
    YCursorTextPosition = mYCursorTextPosition
ElseIf Not mStyle Is Nothing Then
    YCursorTextPosition = mStyle.YCursorTextPosition
Else
    YCursorTextPosition = mBaseStyle.YCursorTextPosition
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YCursorTextPosition(ByVal Value As CursorTextPositions)
Const ProcName As String = "YCursorTextPosition"
On Error GoTo Err

If mYCursorTextPosition = Value And IsPropertySet(YCursorTextPositionIsSet) Then Exit Property
mYCursorTextPosition = Value
setPropertyFlag YCursorTextPositionIsSet
setYCursorTextPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YGridLineStyle() As LineStyle
Const ProcName As String = "YGridLineStyle"

On Error GoTo Err

If IsPropertySet(YGridLineStyleIsSet) Then
    Set YGridLineStyle = mYGridLineStyle
ElseIf Not mStyle Is Nothing Then
    Set YGridLineStyle = mStyle.YGridLineStyle
Else
    Set YGridLineStyle = mBaseStyle.YGridLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "YGridLineStyle"

On Error GoTo Err

If mYGridLineStyle Is Value And IsPropertySet(YGridLineStyleIsSet) Then Exit Property
Set mYGridLineStyle = Value
If mYGridLineStyle Is Nothing Then
    ClearProperty YGridLineStyleIsSet
Else
    setPropertyFlag YGridLineStyleIsSet
End If
setYGridLineStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YGridTextPosition() As YGridTextPositions
Const ProcName As String = "YGridTextPosition"
On Error GoTo Err

If IsPropertySet(YGridTextPositionIsSet) Then
    YGridTextPosition = mYGridTextPosition
ElseIf Not mStyle Is Nothing Then
    YGridTextPosition = mStyle.YGridTextPosition
Else
    YGridTextPosition = mBaseStyle.YGridTextPosition
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YGridTextPosition(ByVal Value As YGridTextPositions)
Const ProcName As String = "YGridTextPosition"
On Error GoTo Err

If mYGridTextPosition = Value And IsPropertySet(YGridTextPositionIsSet) Then Exit Property
mYGridTextPosition = Value
setPropertyFlag YGridTextPositionIsSet
setYGridTextPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YGridTextStyle() As TextStyle
Const ProcName As String = "YGridTextStyle"

On Error GoTo Err

If IsPropertySet(YGridTextStyleIsSet) Then
    Set YGridTextStyle = mYGridTextStyle
ElseIf Not mStyle Is Nothing Then
    Set YGridTextStyle = mStyle.YGridTextStyle
Else
    Set YGridTextStyle = mBaseStyle.YGridTextStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YGridTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YGridTextStyle"

On Error GoTo Err

If mYGridTextStyle Is Value And IsPropertySet(YGridTextStyleIsSet) Then Exit Property
Set mYGridTextStyle = Value
If mYGridTextStyle Is Nothing Then
    ClearProperty YGridTextStyleIsSet
Else
    setPropertyFlag YGridTextStyleIsSet
End If
setYGridTextStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YScaleGridSpacing() As Single
Const ProcName As String = "YScaleGridSpacing"

On Error GoTo Err

YScaleGridSpacing = mHorizontalGridSpacer.YScaleGridSpacing

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YScaleQuantum(ByVal Value As Double)
Const ProcName As String = "YScaleQuantum"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set YScaleQuantum for an axis region"
YScaleQuantumEx = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let YScaleQuantumEx(ByVal Value As Double)
Const ProcName As String = "YScaleQuantumEx"

On Error GoTo Err

If Value <= 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, , "YScaleQuantum must be greater than zero"

If mYScaleQuantum = Value And IsPropertySet(YScaleQuantumIsSet) Then Exit Property
mYScaleQuantum = Value
setPropertyFlag YScaleQuantumIsSet
setYScaleQuantum

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddGraphicObjectSeries( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                Optional ByVal layer As Long = -1, _
                Optional ByVal Name As String) As IGraphicObjectSeries
Const ProcName As String = "AddGraphicObjectSeries"

On Error GoTo Err

Set AddGraphicObjectSeries = mGraphObjSeriesCollection.Add( _
                                            graphObjSeries, _
                                            mData.GetNextUserLayer(layer), _
                                            Name, _
                                            mViewport, _
                                            mPeriods, _
                                            mGraphObjServiceProvider)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddLine( _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As LineStyle) As Line
Const ProcName As String = "AddLine"
On Error GoTo Err

Static miscLines As LineSeries
Dim lLine As Line

If miscLines Is Nothing Then
    Set miscLines = AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerHighestUser - 1)
End If
Set lLine = miscLines.Add
If pLayer <> -1 Then lLine.layer = pLayer
If Not Style Is Nothing Then lLine.LocalStyle = Style

Set AddLine = lLine

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub AddPeriod(ByVal pPeriod As Period)
Const ProcName As String = "AddPeriod"
On Error GoTo Err

mData.AddPeriod pPeriod.PeriodNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddText( _
                Optional ByVal Value As String, _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As TextStyle) As Text
Const ProcName As String = "AddText"
On Error GoTo Err

Static miscTexts As TextSeries
Dim lText As Text

If miscTexts Is Nothing Then
    Set miscTexts = AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerHighestUser)
End If
Set lText = miscTexts.Add
If Value <> "" Then lText.Text = Value
If pLayer <> -1 Then lText.layer = pLayer
If Not Style Is Nothing Then lText.LocalStyle = Style

Set AddText = lText

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub CalculateAndPaint()
Const ProcName As String = "CalculateAndPaint"

On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.CalculateAndPaint
    Else
        CalculateAndPaintSync
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub CalculateAndPaintSync()
Const ProcName As String = "CalculateAndPaintSync"

On Error GoTo Err

mData.DetermineInScopeObjects
calculateVerticalBounds
mData.DetermineVisibleObjects
PaintSync

If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.SetVerticalScaleInternal mViewport.Bottom, mViewport.Top
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ChangeGaugeAndPaint()
Const ProcName As String = "ChangeGaugeAndPaint"

On Error GoTo Err

mData.CalculateInscopeGaugeDependentObjects
PaintSync
If Not mYAxisRegion Is Nothing Then mYAxisRegion.NotifyResizedY

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ClearRegion()
Const ProcName As String = "ClearRegion"
On Error GoTo Err

If mIsCleared Then Exit Sub

mDeferredPainter.Finish
Set mDeferredPainter = Nothing

Set mTitle = Nothing

Set mPointerController = Nothing

If Not mYAxisRegion Is Nothing Then mYAxisRegion.ClearRegion
Set mYAxisRegion = Nothing

Set mYCursorText = Nothing
Set mXCursorText = Nothing

Set mHorizontalGrid = Nothing

mIsCleared = True

mGraphObjSeriesCollection.Finish

mGraphObjServiceProvider.Finish
Set mGraphObjServiceProvider = Nothing

Set mPeriods = Nothing

mData.Clear
Set mData = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearSelection()
Const ProcName As String = "ClearSelection"

On Error GoTo Err

deSelectObjectOrSeries mSelectedObject
Set mSelectedObject = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"

On Error GoTo Err

mPropertyOverrideFlags = 0
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Click()
Const ProcName As String = "Click"

On Error GoTo Err

clickSelectedObjectOrSeries
RaiseEvent Click

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DblCLick()
Const ProcName As String = "DblCLick"

On Error GoTo Err

dblClickSelectedObjectOrSeries
RaiseEvent DblCLick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DrawCursor( _
                ByVal Button As Integer, _
                ByVal Shift As Integer, _
                ByVal X As Single, _
                ByVal Y As Single, _
                ByVal pRegionContainsPointer As Boolean)
Const ProcName As String = "DrawCursor"
On Error GoTo Err

mPointerController.DrawPointer Button, Shift, X, Y, pRegionContainsPointer

drawCursorText X, Y, pRegionContainsPointer
drawXCursorText X, Y, pRegionContainsPointer
drawYCursorText X, Y, pRegionContainsPointer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function FormatYValue(ByVal Y As Double) As String
Const ProcName As String = "FormatYValue"

On Error GoTo Err

FormatYValue = mPriceFormatter.FormatPrice(Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub HideGrid()
Const ProcName As String = "HideGrid"

On Error GoTo Err

If mHorizontalGrid.Hidden Then Exit Sub

mHorizontalGrid.Hidden = True
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal Name As String, _
                ByVal pPeriods As Periods, _
                ByVal pViewport As ViewPort, _
                ByVal pRegionType As RegionTypes)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = Name

Set mPeriods = pPeriods
Set mViewport = pViewport

mRegionType = pRegionType

mData.Initialise mViewport, mGraphObjSeriesCollection

Dim aTextSeries As New TextSeries
AddGraphicObjectSeries aTextSeries

Set mCursorText = aTextSeries.Add
mCursorText.layer = LayerNumbers.LayerPointer + 1
mCursorText.HideIfBlank = True
mCursorText.Extended = True ' need this so it displays when cursor is outside the period range

Set mXCursorText = aTextSeries.Add
mXCursorText.layer = LayerNumbers.LayerPointer + 1
mXCursorText.HideIfBlank = True
mXCursorText.Extended = True ' need this so it displays when cursor is outside the period range

Set mYCursorText = aTextSeries.Add
mYCursorText.layer = LayerNumbers.LayerPointer + 1
mYCursorText.HideIfBlank = True
mYCursorText.Extended = True ' need this so it displays when cursor is outside the period range

Set mPointerController = New PointerController
mPointerController.Initialise mViewport, mRegionType, AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerPointer)

setupGrid
processStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsGridHidden() As Boolean
Const ProcName As String = "IsGridHidden"

On Error GoTo Err

IsGridHidden = mHorizontalGrid.Hidden

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Const ProcName As String = "MouseDown"
On Error GoTo Err

mCurrX = X
mCurrY = Y
If mPointerController.PointerMode = PointerModeSelection Then setSelectedObjectOrSeries
RaiseEvent MouseDown(Button, Shift, X, Y)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    mDeferredPainter.ChangeGaugeAndPaint
Else
    ChangeGaugeAndPaint
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Paint()
Const ProcName As String = "Paint"
On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.Paint
    Else
        PaintSync
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PaintSync()
Const ProcName As String = "PaintSync"
On Error GoTo Err

Static et As ElapsedTimer

Dim numVisibleObjects  As Long
Dim index As Long

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

generateGrid

numVisibleObjects = mData.PaintAllVisibleObjects
Debug.Print "Region " & mHandle & ": " & numVisibleObjects & " objects painted"

mPaintingInProgress = False

mUseDeferredPainting = False
mPointerController.RedrawPointer
mUseDeferredPainting = True

If mPerformanceTextVisible Then paintPerformanceText numVisibleObjects, et.ElapsedTimeMicroseconds / 1000

mViewport.Canvas.Surface.Refresh

mViewport.ZOrder 0

mDrawn = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PaintAreas( _
                ByRef pAreas() As TRectangle, _
                ByVal pCount As Long)
                
Const ProcName As String = "PaintAreas"
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RedrawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Const ProcName As String = "RedrawObject"
On Error GoTo Err

Dim i As Long
Dim handled As Boolean

If areaToRedraw.isValid Then
    If mUseDeferredPainting Then
        'mDeferredPainter.AddChangedObject graphObj, areaToRedraw
        mDeferredPainter.AddAreaToRedraw areaToRedraw
    Else
        redrawObjectSync graphObj, areaToRedraw
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGraphicObject( _
                ByVal graphObj As IGraphicObject)
Const ProcName As String = "RemoveGraphicObject"
On Error GoTo Err

RemoveGraphicObjectEx graphObj, True, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGraphicObjectSeries( _
                ByVal pSeries As IGraphicObjectSeries)
Const ProcName As String = "RemoveGraphicObjectSeries"

On Error GoTo Err

If pSeries Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The supplied argument is Nothing"

mGraphObjSeriesCollection.Remove pSeries

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ScaleAndPaint()
Const ProcName As String = "ScaleAndPaint"

On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.ScaleAndPaint
    Else
        ScaleAndPaintSync
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ScaleAndPaintSync()
Const ProcName As String = "ScaleAndPaintSync"

On Error GoTo Err

calculateVerticalBounds
mData.DetermineVisibleObjects
PaintSync

If Not mYAxisRegion Is Nothing Then mYAxisRegion.SetVerticalScaleInternal mViewport.Bottom, mViewport.Top

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScaleUp(ByVal factor As Single)
Const ProcName As String = "ScaleUp"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot scale an axis region"

With mViewport
    SetVerticalScale .Bottom + factor * .Height / 2, _
                                .Top - factor * .Height / 2
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollVertical(ByVal amount As Single)
Const ProcName As String = "ScrollVertical"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot scroll an axis region"

Autoscaling = False
SetVerticalScale mViewport.Bottom + amount, mViewport.Top + amount

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollVerticalProportion(ByVal proportion As Single)
Const ProcName As String = "ScrollVerticalProportion"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot scroll an axis region"

Autoscaling = False
SetVerticalScale mViewport.Bottom + proportion * mViewport.Height, _
                mViewport.Top + proportion * mViewport.Height

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPeriodBackgroundColor( _
                ByVal PeriodNumber As Long, _
                ByVal Color As Long)
Dim layer As Long
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

mViewport.Canvas.SetClippingRegion mViewport.ConvertTRectangleToRect(periodrect)
mData.RedrawArea periodrect, LayerNumbers.LayerMinVisible
mViewport.Canvas.ClearClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

CalculateAndPaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetPointerModeDefault()
Const ProcName As String = "SetPointerModeDefault"
On Error GoTo Err

If mIsDrawingEnabled Then ClearSelection

mPointerController.PointerMode = PointerModeDefault

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetPointerModeSelection()
Const ProcName As String = "SetPointerModeSelection"
On Error GoTo Err

mPointerController.PointerMode = PointerModeSelection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetPointerModeTool( _
                ByVal toolPointerStyle As PointerStyles, _
                ByVal icon As IPictureDisp)
Const ProcName As String = "SetPointerModeTool"
On Error GoTo Err

mPointerController.SetToolPointer toolPointerStyle, icon
mPointerController.PointerMode = PointerModeTool

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' Note that the arguments are ByRef because they may be adjusted
' if they are the same when expressed as Singles.
Public Sub SetVerticalScale(ByRef lowValue As Double, _
                            ByRef highValue As Double)
Const ProcName As String = "SetVerticalScale"

On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set the scale for an axis region"

Autoscaling = False

SetVerticalScaleInternal lowValue, highValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetVerticalScaleInternal(ByRef lowValue As Double, _
                            ByRef highValue As Double)
Const ProcName As String = "SetVerticalScaleInternal"

On Error GoTo Err

SetViewportVerticalBounds lowValue, highValue

CalculateAndPaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetViewportVerticalBounds( _
                ByRef Bottom As Double, _
                ByRef Top As Double)
Const ProcName As String = "SetViewportVerticalBounds"
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
    .SetVerticalBounds Bottom, Top
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.SetViewportVerticalBounds Bottom, Top
    ' return possibly adjusted values
    Top = .Top
    Bottom = .Bottom
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowGrid()
Const ProcName As String = "ShowGrid"

On Error GoTo Err

If Not mHorizontalGrid.Hidden Then Exit Sub

mHorizontalGrid.Hidden = False
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub UndrawArea( _
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub UndrawAreaEx( _
                ByRef undrawnArea As TRectangle, _
                Optional ByVal objToExclude As IGraphicObject)
Const ProcName As String = "UndrawAreaEx"
On Error GoTo Err

' redraw objects in the undrawn area, preserving layering
If mIsDrawingEnabled And undrawnArea.isValid Then
    mViewport.UndrawArea undrawnArea
    mViewport.Canvas.SetClippingRegion mViewport.ConvertTRectangleToRect(undrawnArea)
    RedrawArea undrawnArea, LayerMinVisible, objToExclude
    mViewport.Canvas.ClearClippingRegion
    
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calculateVerticalBounds()
Const ProcName As String = "calculateVerticalBounds"
On Error GoTo Err

Dim scaleMax As Double
Dim scaleMin As Double

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
    SetViewportVerticalBounds scaleMin, scaleMax
    mData.CalcVerticalBoundsForGaugeDependentObjects _
                                scaleMin, _
                                scaleMax
End If

If scaleMin = PlusInfinityDouble Or scaleMax = MinusInfinityDouble Then Exit Sub

'If scaleMin < -200 Then Stop

SetViewportVerticalBounds scaleMin - (scaleMax - scaleMin) * 0.025, _
                        scaleMax + (scaleMax - scaleMin) * 0.025

gLogger.Log "Min=" & mViewport.Bottom & "; Max=" & mViewport.Top, ProcName, ModuleName, LogLevelMediumDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ClearProperty( _
                ByVal flag As PropertyOverrideFlags)
Const ProcName As String = "ClearProperty"
On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub drawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Dim drawAreas(0) As TRectangle
Const ProcName As String = "drawObject"

On Error GoTo Err

drawAreas(0) = RectIntersection(mViewport.Boundary, areaToRedraw)

If graphObj.layer <> LayerInvisible Then graphObj.Draw drawAreas

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub drawCursorText(ByVal X As Single, ByVal Y As Single, ByVal pRegionContainsPointer As Boolean)
Const ProcName As String = "drawCursorText"
On Error GoTo Err

If Not pRegionContainsPointer Then
    mCursorText.Text = ""
    Exit Sub
End If

If CursorTextMode = CursorTextModeCombined Then
    Select Case CursorTextPosition
    Case CursorTextPositionLeft
        mCursorText.Position = gNewPoint(0#, Y, CoordsDistance, CoordsLogical)
    Case CursorTextPositionTop
        mCursorText.Position = gNewPoint(X, 0#, CoordsLogical, CoordsCounterDistance)
    Case CursorTextPositionBottom
        mCursorText.Position = gNewPoint(X, 0#, CoordsLogical, CoordsDistance)
    Case CursorTextPositionRight
        mCursorText.Position = gNewPoint(0#, Y, CoordsCounterDistance, CoordsLogical)
    Case CursorTextPositionCentre
        mCursorText.Position = gNewPoint(50#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBelowCursor, _
            CursorTextPositionAboveCursor, _
            CursorTextPositionLeftCursor, _
            CursorTextPositionRightCursor, _
            CursorTextPositionAboveRightCursor, _
            CursorTextPositionBelowRightCursor, _
            CursorTextPositionBelowLeftCursor, _
            CursorTextPositionAboveLeftCursor
        mCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionTopLeftFixed
        mCursorText.Position = gNewPoint(0#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionTopCentreFixed
        mCursorText.Position = gNewPoint(50#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionTopRightFixed
        mCursorText.Position = gNewPoint(100#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreRightFixed
        mCursorText.Position = gNewPoint(100#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomRightFixed
        mCursorText.Position = gNewPoint(100#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomCentreFixed
        mCursorText.Position = gNewPoint(50#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomLeftFixed
        mCursorText.Position = gNewPoint(0#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreLeftFixed
        mCursorText.Position = gNewPoint(0#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreFixed
        mCursorText.Position = gNewPoint(50#, 50#, CoordsRelative, CoordsRelative)
    End Select

    mCursorText.Text = formatXCursorText(X) & ", " & mPriceFormatter.FormatPrice(Y)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub drawXCursorText(ByVal X As Single, ByVal Y As Single, ByVal pRegionContainsPointer As Boolean)
Const ProcName As String = "drawXCursorText"
On Error GoTo Err

If Not pRegionContainsPointer Then
    mXCursorText.Text = ""
    Exit Sub
End If

Select Case CursorTextMode
Case CursorTextModeXOnly, CursorTextModeBoth
    Select Case XCursorTextPosition
    Case CursorTextPositionLeft
        mXCursorText.Position = gNewPoint(0#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionTop
        mXCursorText.Position = gNewPoint(X, 100#, CoordsLogical, CoordsRelative)
    Case CursorTextPositionBottom
        mXCursorText.Position = gNewPoint(X, 0#, CoordsLogical, CoordsRelative)
    Case CursorTextPositionRight
        mXCursorText.Position = gNewPoint(100#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentre
        mXCursorText.Position = gNewPoint(X, 50#, CoordsLogical, CoordsRelative)
    Case CursorTextPositionBelowCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionAboveCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionLeftCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionRightCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionAboveRightCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionBelowRightCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionBelowLeftCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionAboveLeftCursor
        mXCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionTopLeftFixed
        mXCursorText.Position = gNewPoint(0#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionTopCentreFixed
        mXCursorText.Position = gNewPoint(50#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionTopRightFixed
        mXCursorText.Position = gNewPoint(100#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreRightFixed
        mXCursorText.Position = gNewPoint(100#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomRightFixed
        mXCursorText.Position = gNewPoint(100#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomCentreFixed
        mXCursorText.Position = gNewPoint(50#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomLeftFixed
        mXCursorText.Position = gNewPoint(0#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreLeftFixed
        mXCursorText.Position = gNewPoint(0#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreFixed
        mXCursorText.Position = gNewPoint(50#, 50#, CoordsRelative, CoordsRelative)
    End Select
    
    mXCursorText.Text = formatXCursorText(X)
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub drawYCursorText(ByVal X As Single, ByVal Y As Single, ByVal pRegionContainsPointer As Boolean)
Const ProcName As String = "drawYCursorText"
On Error GoTo Err

If Not pRegionContainsPointer Then
    mYCursorText.Text = ""
    Exit Sub
End If

Select Case CursorTextMode
Case CursorTextModeYOnly, CursorTextModeBoth
    Select Case YCursorTextPosition
    Case CursorTextPositionLeft
        mYCursorText.Position = gNewPoint(0#, Y, CoordsRelative, CoordsLogical)
    Case CursorTextPositionTop
        mYCursorText.Position = gNewPoint(50#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottom
        mYCursorText.Position = gNewPoint(50#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionRight
        mYCursorText.Position = gNewPoint(100#, Y, CoordsRelative, CoordsLogical)
    Case CursorTextPositionCentre
        mYCursorText.Position = gNewPoint(50#, Y, CoordsRelative, CoordsLogical)
    Case CursorTextPositionBelowCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionAboveCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionLeftCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionRightCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionAboveRightCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionBelowRightCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionBelowLeftCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionAboveLeftCursor
        mYCursorText.Position = gNewPoint(X, Y, CoordsLogical, CoordsLogical)
    Case CursorTextPositionTopLeftFixed
        mYCursorText.Position = gNewPoint(0#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionTopCentreFixed
        mYCursorText.Position = gNewPoint(50#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionTopRightFixed
        mYCursorText.Position = gNewPoint(100#, 100#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreRightFixed
        mYCursorText.Position = gNewPoint(100#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomRightFixed
        mYCursorText.Position = gNewPoint(100#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomCentreFixed
        mYCursorText.Position = gNewPoint(50#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionBottomLeftFixed
        mYCursorText.Position = gNewPoint(0#, 0#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreLeftFixed
        mYCursorText.Position = gNewPoint(0#, 50#, CoordsRelative, CoordsRelative)
    Case CursorTextPositionCentreFixed
        mYCursorText.Position = gNewPoint(50#, 50#, CoordsRelative, CoordsRelative)
    End Select
    
    mYCursorText.Text = mPriceFormatter.FormatPrice(Y)
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function formatXCursorText(ByVal X As Single) As String
Const ProcName As String = "formatXCursorText"
On Error GoTo Err

Dim lThisPeriod As Period
Dim lPeriodNumber As Long

If mPeriods.Count = 0 Then Exit Function

lPeriodNumber = Round(X)
If lPeriodNumber >= mYAxisPosition Then Exit Function

On Error Resume Next
Set lThisPeriod = mPeriods(lPeriodNumber)
On Error GoTo Err

If lThisPeriod Is Nothing Then
    formatXCursorText = ""
    Exit Function
End If

Select Case mPeriods.TimePeriod.Units
Case TimePeriodNone, TimePeriodMinute, TimePeriodHour
    formatXCursorText = FormatDateTime(lThisPeriod.Timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(lThisPeriod.Timestamp, vbShortTime)
Case TimePeriodSecond, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    formatXCursorText = FormatDateTime(lThisPeriod.Timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(lThisPeriod.Timestamp, vbLongTime)
Case Else
    formatXCursorText = FormatDateTime(lThisPeriod.Timestamp, vbShortDate)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub generateGrid()
Const ProcName As String = "generateGrid"
On Error GoTo Err

mPriceFormatter.YScaleGridSpacing = mHorizontalGridSpacer.Calculate(mViewport.Height, mViewport.HeightCm)

mHorizontalGrid.Generate mHorizontalGridSpacer.YScaleGridSpacing
mVerticalGrid.Generate

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    perfText.layer = LayerTitle
End If
Set getPerformanceText = perfText

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleStylePropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "handleStylePropertyChanged"

On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("Autoscaling")
    setAutoscaling
Case UCase$("BackGradientFillColors")
    setBackGradientFillColors
Case UCase$("YGridlineSpacing")
    setYGridlineSpacing
Case UCase$("CursorTextMode")
    setCursorTextMode
Case UCase$("CursorTextPosition")
    setCursorTextPosition mCursorText, CursorTextPosition
Case UCase$("CursorTextStyle")
    setCursorTextStyle
Case UCase$("HasXGrid")
    setHasXGrid
Case UCase$("HasXGridText")
    setHasXGridText
Case UCase$("HasYGrid")
    setHasYGrid
Case UCase$("HasYGridText")
    setHasYGridText
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
Case UCase$("XCursorTextPosition")
    setXCursorTextPosition
Case UCase$("XCursorTextStyle")
    setXCursorTextStyle
Case UCase$("XGridLineStyle")
    setXGridLineStyle
Case UCase$("XGridTextPosition")
    setXGridTextPosition
Case UCase$("XGridTextStyle")
    setXGridTextStyle
Case UCase$("YCursorTextPosition")
    setYCursorTextPosition
Case UCase$("YCursorTextStyle")
    setYCursorTextStyle
Case UCase$("YGridLineStyle")
    setYGridLineStyle
Case UCase$("YGridTextPosition")
    setYGridTextPosition
Case UCase$("YGridTextStyle")
    setYGridTextStyle
Case UCase$("YScaleQuantum")
    setYScaleQuantum
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Unhandled property change"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IsPropertySet( _
                ByVal flag As PropertyOverrideFlags) As Boolean
Const ProcName As String = "IsPropertySet"
On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStyle()
Const ProcName As String = "processStyle"

On Error GoTo Err

setBackGradientFillColors
setHasXGrid
setHasXGridText
setHasYGrid
setHasYGridText
setCursorTextMode
setCursorTextPosition mCursorText, CursorTextPosition
setCursorTextStyle
setXCursorTextPosition
setXCursorTextStyle
setXGridTextStyle
setXGridTextPosition
setYCursorTextPosition
setYCursorTextStyle
setYGridTextStyle
setYGridTextPosition
setYGridlineSpacing
setSessionEndGridLineStyle
setSessionStartGridLineStyle
If Not IsXAxisRegion And Not IsYAxisRegion Then
    setAutoscaling
    setIntegerYScale
    setMinimumHeight
    setCursorSnapsToTickBoundaries
    setYScaleQuantum
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub redrawObjectSync( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Const ProcName As String = "redrawObjectSync"
On Error GoTo Err

mViewport.Canvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areaToRedraw)
drawObject graphObj, areaToRedraw
RedrawArea areaToRedraw, graphObj.layer + 1
mViewport.Canvas.ClearClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setAutoscaling()
Const ProcName As String = "setAutoscaling"
On Error GoTo Err

If IsXAxisRegion Or IsYAxisRegion Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set autoscale for an axis region"

If Autoscaling Then ScaleAndPaint

RaiseEvent AutoscalingChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBackGradientFillColors()
Const ProcName As String = "setBackGradientFillColors"

On Error GoTo Err

mViewport.GradientFillColors = BackGradientFillColors
Paint

setDividerColors
PaintDivider

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCursorTextMode()
Const ProcName As String = "setCursorTextMode"
On Error GoTo Err

Select Case CursorTextMode
Case CursorTextModeNone
    mCursorText.Text = ""
    mXCursorText.Text = ""
    mYCursorText.Text = ""
Case CursorTextModeXOnly
    mCursorText.Text = ""
    mYCursorText.Text = ""
Case CursorTextModeYOnly
    mCursorText.Text = ""
    mXCursorText.Text = ""
Case CursorTextModeBoth
    mXCursorText.Text = ""
    mYCursorText.Text = ""
Case CursorTextModeCombined
    mCursorText.Text = ""
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCursorTextPosition(ByVal pCursorText As Text, ByVal pPosition As CursorTextPositions)
Const ProcName As String = "setCursorTextPosition"
On Error GoTo Err

Select Case pPosition
Case CursorTextPositionLeft
    pCursorText.Align = AlignBoxCentreLeft
    pCursorText.Justification = JustifyLeft
Case CursorTextPositionTop
    pCursorText.Align = AlignBoxTopCentre
    pCursorText.Justification = JustifyCentre
Case CursorTextPositionBottom
    pCursorText.Align = AlignBoxBottomCentre
    pCursorText.Justification = JustifyCentre
Case CursorTextPositionRight
    pCursorText.Align = AlignBoxCentreRight
    pCursorText.Justification = JustifyRight
Case CursorTextPositionCentre
    pCursorText.Align = AlignBoxCentreCentre
    pCursorText.Justification = JustifyCentre
Case CursorTextPositionBelowCursor
    pCursorText.Align = AlignBoxTopCentre
    pCursorText.Justification = JustifyCentre
Case CursorTextPositionAboveCursor
    pCursorText.Align = AlignBoxBottomCentre
    pCursorText.Justification = JustifyCentre
Case CursorTextPositionLeftCursor
    pCursorText.Align = AlignBoxCentreRight
    pCursorText.Justification = JustifyRight
Case CursorTextPositionRightCursor
    pCursorText.Align = AlignBoxCentreLeft
    pCursorText.Justification = JustifyLeft
Case CursorTextPositionAboveRightCursor
    pCursorText.Align = AlignBoxBottomLeft
    pCursorText.Justification = JustifyLeft
Case CursorTextPositionBelowRightCursor
    pCursorText.Align = AlignBoxTopLeft
    pCursorText.Justification = JustifyLeft
Case CursorTextPositionBelowLeftCursor
    pCursorText.Align = AlignBoxTopRight
    pCursorText.Justification = JustifyRight
Case CursorTextPositionAboveLeftCursor
    pCursorText.Align = AlignBoxBottomRight
    pCursorText.Justification = JustifyRight
Case CursorTextPositionTopLeftFixed
    pCursorText.Align = AlignBoxTopLeft
    pCursorText.Justification = JustifyLeft
Case CursorTextPositionTopCentreFixed
    pCursorText.Align = AlignBoxTopCentre
    pCursorText.Justification = JustifyCentre
Case CursorTextPositionTopRightFixed
    pCursorText.Align = AlignBoxTopRight
    pCursorText.Justification = JustifyRight
Case CursorTextPositionCentreRightFixed
    pCursorText.Align = AlignBoxCentreRight
    pCursorText.Justification = JustifyRight
Case CursorTextPositionBottomRightFixed
    pCursorText.Align = AlignBoxBottomRight
    pCursorText.Justification = JustifyRight
Case CursorTextPositionBottomCentreFixed
    pCursorText.Align = AlignBoxBottomCentre
    pCursorText.Justification = JustifyCentre
Case CursorTextPositionBottomLeftFixed
    pCursorText.Align = AlignBoxBottomLeft
    pCursorText.Justification = JustifyLeft
Case CursorTextPositionCentreLeftFixed
    pCursorText.Align = AlignBoxCentreLeft
    pCursorText.Justification = JustifyLeft
Case CursorTextPositionCentreFixed
    pCursorText.Align = AlignBoxCentreCentre
    pCursorText.Justification = JustifyCentre
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCursorTextStyle()
Const ProcName As String = "setCursorTextStyle"

On Error GoTo Err

If Not mCursorText Is Nothing Then mCursorText.LocalStyle = CursorTextStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setYGridlineSpacing()
Const ProcName As String = "setYGridlineSpacing"

On Error GoTo Err

mHorizontalGridSpacer.YGridlineSpacing = YGridlineSpacing
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setHasXGrid()
Const ProcName As String = "setHasXGrid"

On Error GoTo Err

If HasXGrid Then mVerticalGrid.GridLineStyle = XGridLineStyle

mVerticalGrid.HasGrid = HasXGrid
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setHasXGridText()
Const ProcName As String = "setHasXGridText"
On Error GoTo Err

mVerticalGrid.HasGridText = HasXGridText
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setHasYGrid()
Const ProcName As String = "setHasYGrid"
On Error GoTo Err

If HasYGrid Then mHorizontalGrid.GridLineStyle = YGridLineStyle

mHorizontalGrid.HasGrid = HasYGrid
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setHasYGridText()
Const ProcName As String = "setHasYGridText"
On Error GoTo Err

mHorizontalGrid.HasGridText = HasYGridText
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setIntegerYScale()
Const ProcName As String = "setIntegerYScale"
On Error GoTo Err

mPriceFormatter.IntegerYScale = IntegerYScale
Paint
If Not mYAxisRegion Is Nothing Then mYAxisRegion.IntegerYScaleEx = IntegerYScale

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setMinimumHeight()
Const ProcName As String = "setMinimumHeight"
On Error GoTo Err

If YScaleQuantum > MinimumHeight Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Minimum height cannot be less than YScaleQuantum"

ScaleAndPaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSelectedObjectOrSeries()
Const ProcName As String = "setSelectedObjectOrSeries"
On Error GoTo Err

Dim newSelectedObject As IGraphicObject

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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSessionEndGridLineStyle()
Const ProcName As String = "setSessionEndGridLineStyle"
On Error GoTo Err

If SessionEndGridLineStyle Is Nothing Then Exit Sub
'SessionEndGridLineStyle.SyncTo mSessEndGridlineStyle
mVerticalGrid.SessionEndGridLineStyle = SessionEndGridLineStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSessionStartGridLineStyle()
Const ProcName As String = "setSessionStartGridLineStyle"
On Error GoTo Err

If SessionStartGridLineStyle Is Nothing Then Exit Sub
'SessionStartGridLineStyle.SyncTo mSessStartGridlineStyle
mVerticalGrid.SessionStartGridLineStyle = SessionStartGridLineStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupGrid()
Const ProcName As String = "setupGrid"
On Error GoTo Err

Set mHorizontalGrid = New HorizontalGrid

mHorizontalGrid.Initialise AddGraphicObjectSeries(New LineSeries), AddGraphicObjectSeries(New TextSeries), mViewport

Set mPriceFormatter = New PriceFormatter
mHorizontalGrid.PriceFormatter = mPriceFormatter

Set mHorizontalGridSpacer = New HorizontalGridSpacer

Set mVerticalGrid = New VerticalGrid
mVerticalGrid.Initialise AddGraphicObjectSeries(New LineSeries), _
                        AddGraphicObjectSeries(New TextSeries), _
                        AddGraphicObjectSeries(New TextSeries), _
                        mPeriods, _
                        mViewport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setXCursorTextPosition()
Const ProcName As String = "setXCursorTextPosition"
On Error GoTo Err

setCursorTextPosition mXCursorText, XCursorTextPosition

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setXCursorTextStyle()
Const ProcName As String = "setXCursorTextStyle"
On Error GoTo Err

If Not mXCursorText Is Nothing Then mXCursorText.LocalStyle = XCursorTextStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setXGridLineStyle()
Const ProcName As String = "setXGridLineStyle"
On Error GoTo Err

mVerticalGrid.GridLineStyle = XGridLineStyle
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setXGridTextPosition()
Const ProcName As String = "setXGridTextPosition"
On Error GoTo Err

mVerticalGrid.GridTextPosition = XGridTextPosition

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setXGridTextStyle()
Const ProcName As String = "setXGridTextStyle"
On Error GoTo Err

mVerticalGrid.GridTextStyle = XGridTextStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setYCursorTextPosition()
Const ProcName As String = "setYCursorTextPosition"
On Error GoTo Err

setCursorTextPosition mYCursorText, YCursorTextPosition

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setYCursorTextStyle()
Const ProcName As String = "setYCursorTextStyle"
On Error GoTo Err

If Not mYCursorText Is Nothing Then mYCursorText.LocalStyle = YCursorTextStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setYGridLineStyle()
Const ProcName As String = "setYGridLineStyle"
On Error GoTo Err

mHorizontalGrid.GridLineStyle = YGridLineStyle
Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setYGridTextPosition()
Const ProcName As String = "setYGridTextPosition"
On Error GoTo Err

mHorizontalGrid.GridTextPosition = YGridTextPosition

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setYGridTextStyle()
Const ProcName As String = "setYGridTextStyle"
On Error GoTo Err

mHorizontalGrid.GridTextStyle = YGridTextStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setYScaleQuantum()
Const ProcName As String = "setYScaleQuantum"
On Error GoTo Err

mPriceFormatter.YScaleQuantum = YScaleQuantum
mHorizontalGridSpacer.YScaleQuantum = YScaleQuantum

Paint

If Not mYAxisRegion Is Nothing Then mYAxisRegion.YScaleQuantumEx = YScaleQuantum

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


