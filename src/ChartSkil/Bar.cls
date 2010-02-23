VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bar"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                As String = "Bar"

'================================================================================
' Member variables
'================================================================================

Private mSeries As BarSeries

Private mViewport As Viewport
Private mCanvas As Canvas
Private mLayer As Long
Private mHandle As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mTimestamp As Date

Private WithEvents mLocalStyle As BarStyle
Attribute mLocalStyle.VB_VarHelpID = -1

Private mTailThickness As Long
Private mOutlineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayMode As BarDisplayModes
Private mIncludeInAutoscale As Boolean
Private mSolidUpBody As Boolean
Private mThickness As Long
Private mWidth As Single

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

Private mBoundingRect As TRectangle

Private mX As Double

Private mKey As String

Private mOpen As Double
Private mHigh As Double
Private mLow As Double
Private mClose As Double
Private mBlank As Boolean   ' indicates that there were no ticks during the bar

Private mBarSemiWidth As Single ' half the Width of the bar - this actually depends
                                ' on the bar display mode, rather than being simply
                                ' half of Width

Private mPropertyOverrideFlags As BarPropertyOverrideFlags

Private mPrevBarRef As WeakReference

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mOpen = MinusInfinityDouble
mLow = PlusInfinityDouble
mHigh = MinusInfinityDouble
mClose = MinusInfinityDouble
mBlank = True

mBoundingRect.isValid = False

End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "Bar terminated"
Debug.Print "Bar terminated"
End Sub

'================================================================================
' IGraphicObject Members
'================================================================================

Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Const ProcName As String = "IGraphicObject_boundingRectangle"
Dim failpoint As String
On Error GoTo Err

setBarSemiWidth
calcBoundingRect
IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_BoundingRectangleInViewport() As TRectangle
Const ProcName As String = "IGraphicObject_BoundingRectangleInViewport"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_BoundingRectangleInViewport = IGraphicObject_boundingRectangle
RectSetXInterval IGraphicObject_BoundingRectangleInViewport, _
                IntIntersection(RectGetXInterval(IGraphicObject_BoundingRectangleInViewport), _
                                RectGetXInterval(mViewport.Boundary))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_BoundingRectanglesAt( _
                ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
Const ProcName As String = "IGraphicObject_capabilities"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_capabilities = GraphicObjectCapabilities.PartialRedraw

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub IGraphicObject_Click()
'RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
'RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
Dim i As Long

Const ProcName As String = "IGraphicObject_draw"
Dim failpoint As String
On Error GoTo Err

'For i = 0 To UBound(areas)
'    mCanvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areas(i))
    Draw
'    mCanvas.ClearClippingRegion
'Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
Set mViewport = Nothing
End Sub

Private Property Get IGraphicObject_gaugeDependent() As Boolean
IGraphicObject_gaugeDependent = False
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
Const ProcName As String = "IGraphicObject_HitTest"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_HitTest = RectContainsPoint(mBoundingRect, X, Y)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Let IGraphicObject_Handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
IGraphicObject_IncludeInAutoscale = True
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
IGraphicObject_noDraw = mBlank
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
Dim failpoint As String
On Error GoTo Err

IGraphicObject_SeriesID = mSeries.Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_Timestamp() As Date
IGraphicObject_Timestamp = mTimestamp
End Property

Private Property Get IGraphicObject_VerticalRangeInXinterval(ByRef pXInterval As TInterval) As TInterval
If Not pXInterval.isValid Then Exit Property
If mX < pXInterval.startValue Then Exit Property
If mX > pXInterval.endValue Then Exit Property

IGraphicObject_VerticalRangeInXinterval.startValue = mLow
IGraphicObject_VerticalRangeInXinterval.endValue = mHigh
IGraphicObject_VerticalRangeInXinterval.isValid = True
                    
End Property

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property

'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Const ProcName As String = "mLocalStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "COLOR"
    initiateRedraw False, False
Case "THICKNESS"
    initiateRedraw True, False
Case "WIDTH"
    initiateRedraw True, True
Case "DISPLAYMODE"
    initiateRedraw True, True
Case "DOWNCOLOR"
    initiateRedraw False, False
Case "INCLUDEINAUTOSCALE"
    initiateRedraw False, False
Case "OUTLINETHICKNESS"
    initiateRedraw False, False
Case "SOLIDUPBODY"
    initiateRedraw False, False
Case "TAILTHICKNESS"
    initiateRedraw True, False
Case "UPCOLOR"
    initiateRedraw False, False
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unhandled property change"
End Select

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"

If mColor = value And isPropertyOverrideFlagSet(BarIsSetColor) Then Exit Property
mColor = value
If mColor <> -1 Then
    setPropertyOverrideFlag BarIsSetColor
Else
    clearPropertyOverrideFlag BarIsSetColor
End If
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

Color = mSeries.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If isPropertyOverrideFlagSet(BarIsSetColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Thickness(ByVal value As Long)
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

If mThickness = value And isPropertyOverrideFlagSet(BarIsSetThickness) Then Exit Property
mThickness = value
setPropertyOverrideFlag BarIsSetThickness
initiateRedraw True, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Thickness() As Long
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

Thickness = mSeries.Thickness
If Not mLocalStyle Is Nothing Then Thickness = mLocalStyle.Thickness
If isPropertyOverrideFlagSet(BarIsSetThickness) Then Thickness = mThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Width(ByVal value As Single)
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

If mWidth = value And isPropertyOverrideFlagSet(BarIsSetWidth) Then Exit Property
mWidth = value
setPropertyOverrideFlag BarIsSetWidth
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Width() As Single
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

Width = mSeries.Width
If Not mLocalStyle Is Nothing Then Width = mLocalStyle.Width
If isPropertyOverrideFlagSet(BarIsSetWidth) Then Width = mWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Public Property Get ClosePrice() As Double
ClosePrice = mClose
End Property

Public Property Let DisplayMode(ByVal value As BarDisplayModes)
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

If mDisplayMode = value And isPropertyOverrideFlagSet(BarIsSetDisplayMode) Then Exit Property
mDisplayMode = value
setPropertyOverrideFlag BarIsSetDisplayMode
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DisplayMode() As BarDisplayModes
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

DisplayMode = mSeries.DisplayMode
If Not mLocalStyle Is Nothing Then DisplayMode = mLocalStyle.DisplayMode
If isPropertyOverrideFlagSet(BarIsSetDisplayMode) Then DisplayMode = mDisplayMode

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DownBar() As Boolean
DownBar = (mClose <= mOpen)
End Property

Public Property Let DownColor(ByVal value As Long)
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"

If mDownColor = value And isPropertyOverrideFlagSet(BarIsSetDownColor) Then Exit Property
mDownColor = value
If mDownColor <> -1 Then
    setPropertyOverrideFlag BarIsSetDownColor
Else
    clearPropertyOverrideFlag BarIsSetDownColor
End If
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DownColor() As Long
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

DownColor = mSeries.DownColor
If Not mLocalStyle Is Nothing Then DownColor = mLocalStyle.DownColor
If isPropertyOverrideFlagSet(BarIsSetDownColor) Then DownColor = mDownColor
If DownColor = -1 Then DownColor = UpColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get HighPrice() As Double
HighPrice = mHigh
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(BarIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag BarIsSetIncludeInAutoscale
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

IncludeInAutoscale = mSeries.IncludeInAutoscale
If Not mLocalStyle Is Nothing Then IncludeInAutoscale = mLocalStyle.IncludeInAutoscale
If isPropertyOverrideFlagSet(BarIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale
IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

If mLayer = value And isPropertyOverrideFlagSet(BarIsSetLayer) Then Exit Property
mLayer = value
setPropertyOverrideFlag BarIsSetLayer
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

Layer = mSeries.Layer
If isPropertyOverrideFlagSet(BarIsSetLayer) Then Layer = mLayer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LocalStyle() As BarStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal value As BarStyle)
Const ProcName As String = "LocalStyle"
Dim failpoint As String
On Error GoTo Err

Set mLocalStyle = value
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LowPrice() As Double
LowPrice = mLow
End Property

Public Property Get OpenPrice() As Double
OpenPrice = mOpen
End Property

Public Property Let OutlineThickness(ByVal value As Long)
Const ProcName As String = "OutlineThickness"
Dim failpoint As String
On Error GoTo Err

If mOutlineThickness = value And isPropertyOverrideFlagSet(BarIsSetOutlineThickness) Then Exit Property
mOutlineThickness = value
setPropertyOverrideFlag BarIsSetOutlineThickness
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OutlineThickness() As Long
Const ProcName As String = "OutlineThickness"
Dim failpoint As String
On Error GoTo Err

OutlineThickness = mSeries.OutlineThickness
If Not mLocalStyle Is Nothing Then OutlineThickness = mLocalStyle.OutlineThickness
If isPropertyOverrideFlagSet(BarIsSetOutlineThickness) Then OutlineThickness = mOutlineThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PreviousBar() As Bar
Const ProcName As String = "PreviousBar"
Dim failpoint As String
On Error GoTo Err

If mPrevBarRef Is Nothing Then Exit Property
Set PreviousBar = prevBarObj

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let PreviousBar(ByVal value As Bar)
Const ProcName As String = "PreviousBar"
Dim failpoint As String
On Error GoTo Err

Set mPrevBarRef = CreateWeakReference(value)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let SolidUpBody(ByVal value As Boolean)
Const ProcName As String = "SolidUpBody"
Dim failpoint As String
On Error GoTo Err

If mSolidUpBody = value And isPropertyOverrideFlagSet(BarIsSetSolidUpBody) Then Exit Property
mSolidUpBody = value
setPropertyOverrideFlag BarIsSetSolidUpBody
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SolidUpBody() As Boolean
Const ProcName As String = "SolidUpBody"
Dim failpoint As String
On Error GoTo Err

SolidUpBody = mSeries.SolidUpBody
If Not mLocalStyle Is Nothing Then SolidUpBody = mLocalStyle.SolidUpBody
If isPropertyOverrideFlagSet(BarIsSetSolidUpBody) Then SolidUpBody = mSolidUpBody

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TailThickness(ByVal value As Long)
Const ProcName As String = "TailThickness"
Dim failpoint As String
On Error GoTo Err

If mTailThickness = value And isPropertyOverrideFlagSet(BarIsSetTailThickness) Then Exit Property
mTailThickness = value
setPropertyOverrideFlag BarIsSetTailThickness
initiateRedraw True, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TailThickness() As Long
Const ProcName As String = "TailThickness"
Dim failpoint As String
On Error GoTo Err

TailThickness = mSeries.TailThickness
If Not mLocalStyle Is Nothing Then TailThickness = mLocalStyle.TailThickness
If isPropertyOverrideFlagSet(BarIsSetTailThickness) Then TailThickness = mTailThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get UpBar() As Boolean
UpBar = (mClose >= mOpen)
End Property

Public Property Let UpColor(ByVal value As Long)
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"

If mUpColor = value And isPropertyOverrideFlagSet(BarIsSetUpColor) Then Exit Property
mUpColor = value
If mUpColor <> -1 Then
    setPropertyOverrideFlag BarIsSetUpColor
Else
    clearPropertyOverrideFlag BarIsSetUpColor
End If
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

UpColor = mSeries.UpColor
If Not mLocalStyle Is Nothing Then UpColor = mLocalStyle.UpColor
If isPropertyOverrideFlagSet(BarIsSetUpColor) Then UpColor = mUpColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Friend Property Let X(ByVal value As Double)
Const ProcName As String = "X"
Dim failpoint As String
On Error GoTo Err

mX = value
calcBoundingRect

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get X() As Double
X = mX
End Property

'================================================================================
' Methods
'================================================================================

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = 0
initiateRedraw True, True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Initialise( _
                ByVal pSeries As BarSeries, _
                ByVal pTimestamp As Date, _
                ByVal pKey As String, _
                ByVal pGOSP As GraphObjServiceProvider, _
                ByVal pViewport As Viewport, _
                ByVal pLocalStyle As BarStyle)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mSeries = pSeries
mTimestamp = pTimestamp
mKey = pKey
Set mGOSP = pGOSP
Set mViewport = pViewport
Set mCanvas = mViewport.Canvas
Set mLocalStyle = LocalStyle

setBarSemiWidth
calcBoundingRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub ResetStyle()
' reset the Style to the initial Style (ie the Style for the bar series) with no
' local modifications
Const ProcName As String = "ResetStyle"
Dim failpoint As String
On Error GoTo Err

Set mLocalStyle = Nothing

setPropertyOverrideFlag BarIsSetTailThickness = False
setPropertyOverrideFlag BarIsSetOutlineThickness = False
setPropertyOverrideFlag BarIsSetColor = False
setPropertyOverrideFlag BarIsSetUpColor = False
setPropertyOverrideFlag BarIsSetDownColor = False
setPropertyOverrideFlag BarIsSetDisplayMode = False
setPropertyOverrideFlag BarIsSetSolidUpBody = False
setPropertyOverrideFlag BarIsSetThickness = False
setPropertyOverrideFlag BarIsSetWidth = False

initiateRedraw True, True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Tick(ByVal Price As Double)
Dim clearRect As TRectangle
Dim boundsChanged As Boolean

Const ProcName As String = "Tick"
Dim failpoint As String
On Error GoTo Err

If Price = mClose Then Exit Sub

clearRect = mBoundingRect

If mBlank Then
    mBlank = False
    mOpen = Price
Else
    With clearRect
        .isValid = False
        If UpBar Then
            If Price < mOpen Then
                ' is now a downbar so redraw the whole thing
                .isValid = True
            ElseIf Price < mClose Then
                .Bottom = Price
                .Top = mClose
                .isValid = True
            ElseIf Price > mClose Then
                .Bottom = mClose
                .Top = Price
                .isValid = True
            End If
        Else
            If Price >= mOpen Then
                ' now an upbar so redraw the whole thing
                .isValid = True
            ElseIf Price > mClose Then
                .Bottom = mClose
                .Top = Price
                .isValid = True
            ElseIf Price < mClose Then
                .Bottom = Price
                .Top = mClose
                .isValid = True
            End If
        End If
    End With
End If

If Price > mHigh Then
    mHigh = Price
    boundsChanged = True
End If
If Price < mLow Then
    mLow = Price
    boundsChanged = True
End If
mClose = Price

initiateRedraw True, boundsChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcBoundingRect()
Dim lRect As RECT
Dim prevBar As Bar

Const ProcName As String = "calcBoundingRect"
Dim failpoint As String
On Error GoTo Err

If mBlank Then
    mBoundingRect.isValid = False
    Exit Sub
End If

If DisplayMode = BarDisplayModeLine And Not mPrevBarRef Is Nothing Then
    Set prevBar = prevBarObj
    mBoundingRect.Left = prevBarObj.X
    mBoundingRect.Right = mX
    If prevBar.ClosePrice < mClose Then
        mBoundingRect.Bottom = prevBar.ClosePrice
        mBoundingRect.Top = mClose
    Else
        mBoundingRect.Bottom = mClose
        mBoundingRect.Top = prevBar.ClosePrice
    End If
Else
    mBoundingRect.Left = mX - mBarSemiWidth
    mBoundingRect.Right = mX + mBarSemiWidth
    mBoundingRect.Bottom = mLow
    mBoundingRect.Top = mHigh
End If

' align to pixel boundaries
lRect = mViewport.ConvertTRectangleToRect(mBoundingRect)

If DisplayMode = BarDisplayModeLine Then InflateRect lRect, Int(Thickness / 2), Int(Thickness / 2)
lRect.Right = lRect.Right + 1
lRect.Bottom = lRect.Bottom + 1

mBoundingRect = mViewport.ConvertRectToTRectangle(lRect)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As BarPropertyOverrideFlags)
Const ProcName As String = "clearPropertyOverrideFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Draw()
Const ProcName As String = "Draw"
Dim failpoint As String
On Error GoTo Err

drawBar IIf(UpBar, UpColor, DownColor), _
        IIf(Color <> -1, Color, IIf(UpBar, UpColor, DownColor))

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


Private Sub drawBar( _
                ByVal Color As Long, _
                ByVal outlineColor As Long)
Dim ThicknessScaleX As Double
Dim ThicknessScaleY As Double

Const ProcName As String = "drawBar"
Dim failpoint As String
On Error GoTo Err

If Not mVisible Then Exit Sub
If Blank Then Exit Sub

If DisplayMode = BarDisplayModeCandlestick Then
    If UpBar Then
        drawBody mOpen, mClose, Color, outlineColor
    Else
        drawBody mClose, mOpen, Color, outlineColor
    End If
    
    If UpBar Then
        If mLow < mOpen Then
            drawTail mLow, mOpen, outlineColor
        End If
        If mHigh > mClose Then
            drawTail mClose, mHigh, outlineColor
        End If
    Else
        If mHigh > mOpen Then
            drawTail mOpen, mHigh, outlineColor
        End If
        If mLow < mClose Then
            drawTail mLow, mClose, outlineColor
        End If
    End If
ElseIf DisplayMode = BarDisplayModeLine Then
    mCanvas.SetPenAttributes outlineColor, _
                            Thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    If Not mPrevBarRef Is Nothing Then
        Dim prevBar As Bar
        Set prevBar = prevBarObj
        mCanvas.DrawLine prevBar.X, prevBar.ClosePrice, mX, mClose
    Else
        mCanvas.DrawLine mX, mClose, mX, mClose
    End If
Else
    mCanvas.SetBrushAttributes outlineColor, FillStyles.FillSolid
    If Thickness = 1 Then
        Dim onePixel As Double
        onePixel = mViewport.ConvertPixelsToLogicalX(1)
        mCanvas.SetPenAttributes outlineColor, _
                                1, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen
        mCanvas.DrawLine mX - mBarSemiWidth, mOpen, _
                        mX, mOpen
        mCanvas.DrawLine mX, mLow, _
                        mX, mHigh
        mCanvas.DrawLine mX, mClose, _
                        mX + mBarSemiWidth + onePixel, mClose
    Else
        mCanvas.SetPenAttributes outlineColor, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        ThicknessScaleX = mViewport.ConvertPixelsToLogicalX(Thickness - 1)
        ThicknessScaleY = mViewport.ConvertPixelsToLogicalY(Thickness - 1)
        
        If (mOpen - ThicknessScaleY / 2) < mLow Then
            mCanvas.DrawRectangle mX - mBarSemiWidth, _
                                    mLow, _
                                    mX + ThicknessScaleX / 2, _
                                    mLow + ThicknessScaleY
        ElseIf (mOpen + ThicknessScaleY / 2) > mHigh Then
            mCanvas.DrawRectangle mX - mBarSemiWidth, mHigh - ThicknessScaleY, _
                                    mX + ThicknessScaleX / 2, mHigh
        Else
            mCanvas.DrawRectangle mX - mBarSemiWidth, mOpen - ThicknessScaleY / 2, _
                                    mX + ThicknessScaleX / 2, mOpen + ThicknessScaleY / 2
        End If
        
        mCanvas.DrawRectangle mX - ThicknessScaleX / 2, mLow, _
                                mX + ThicknessScaleX / 2, mHigh
        
        If (mClose - ThicknessScaleY / 2) < mLow Then
            mCanvas.DrawRectangle mX - ThicknessScaleX / 2, mLow, _
                                    mX + mBarSemiWidth, mLow + ThicknessScaleY
        ElseIf (mClose + ThicknessScaleY / 2) > mHigh Then
            mCanvas.DrawRectangle mX - ThicknessScaleX / 2, mHigh - ThicknessScaleY, _
                                    mX + mBarSemiWidth, mHigh
        Else
            mCanvas.DrawRectangle mX - ThicknessScaleX / 2, mClose - ThicknessScaleY / 2, _
                                    mX + mBarSemiWidth, mClose + ThicknessScaleY / 2
        End If
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub drawBody( _
                ByVal Bottom As Double, _
                ByVal Top As Double, _
                ByVal Color As Long, _
                ByVal outlineColor As Long)
Dim barOutlineScaleX As Double
Dim barOutlineScaleY As Double
Dim lOutlineThickness As Long
Dim wholeBody As Boolean

Const ProcName As String = "drawBody"
Dim failpoint As String
On Error GoTo Err

If UpBar Then
    If Bottom = mOpen And Top = mClose Then
        wholeBody = True
    End If
    If Not SolidUpBody Then Color = vbWhite
Else
    If Bottom = mClose And Top = mOpen Then
        wholeBody = True
    End If
End If

lOutlineThickness = IIf(OutlineThickness = 0, 1, OutlineThickness)

mCanvas.SetBrushAttributes Color, FillStyles.FillSolid
barOutlineScaleX = mViewport.ConvertPixelsToLogicalX(lOutlineThickness - 1)
barOutlineScaleY = mViewport.ConvertPixelsToLogicalY(lOutlineThickness - 1)

If wholeBody Then
    ' draw the whole of the body
    mCanvas.SetPenAttributes outlineColor, _
                            lOutlineThickness, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    mCanvas.DrawRectangle mX - mBarSemiWidth, Bottom, _
                            mX + mBarSemiWidth, Top
Else

    mCanvas.SetPenAttributes Color, _
                            1, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen

    mCanvas.DrawRectangle mX - mBarSemiWidth, Bottom, _
                    mX + mBarSemiWidth, Top

    mCanvas.SetPenAttributes outlineColor, _
                            1, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    ' we'll draw the outlines of the bar as filled boxes
    ' so set the fill Color to the outline Color
    mCanvas.SetBrushAttributes Color, FillStyles.FillSolid

    If Bottom = mOpen Then
        ' draw the base of the body
        If lOutlineThickness = 1 Then
            mCanvas.DrawLine mX - mBarSemiWidth, Bottom, _
                            mX + mBarSemiWidth, Bottom
        Else
            mCanvas.DrawRectangle mX - mBarSemiWidth, Bottom, _
                            mX + mBarSemiWidth, Bottom + barOutlineScaleY
        End If
    End If
    If Top = mClose Then
        ' draw the Top of the body
        If lOutlineThickness = 1 Then
            mCanvas.DrawLine mX - mBarSemiWidth, Top, _
                            mX + mBarSemiWidth, Top
        Else
            mCanvas.DrawRectangle mX - mBarSemiWidth, Top - barOutlineScaleY, _
                            mX + mBarSemiWidth, Top
        End If
    End If
    If lOutlineThickness = 1 Then
        mCanvas.DrawLine mX - mBarSemiWidth, Bottom, _
                        mX - mBarSemiWidth, Top
        mCanvas.DrawLine mX + mBarSemiWidth, Bottom, _
                        mX + mBarSemiWidth, Top
    Else
        mCanvas.DrawRectangle mX - mBarSemiWidth, Bottom, _
                            mX - mBarSemiWidth + barOutlineScaleX, Top
        mCanvas.DrawRectangle mX + mBarSemiWidth - barOutlineScaleX, Bottom, _
                            mX + mBarSemiWidth, Top
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'Private Sub drawPartialBar( _
'                ByRef Area As TRectangle, _
'                ByVal Color As Long, _
'                ByVal outlineColor As Long, _
'                ByVal tailColor As Long)
'Dim ThicknessScaleX As Double
'Dim ThicknessScaleY As Double
'If Not mVisible Then Exit Sub
'If Blank Then Exit Sub
'
'If DisplayMode = BarDisplayModeCandlestick Then
'    If DownBar Then
'        If Area.Top < mClose Then
'            drawTail IIf(mLow < Area.Bottom, Area.Bottom, mLow), Area.Top, tailColor
'        ElseIf Area.Top >= mClose And Area.Top <= mOpen And Area.Bottom < mClose Then
'            drawTail IIf(mLow < Area.Bottom, Area.Bottom, mLow), mClose, tailColor
'            drawBody mClose, Area.Top, Color, outlineColor
'        ElseIf Area.Top > mOpen And Area.Bottom < mClose Then
'            drawTail IIf(mLow < Area.Bottom, Area.Bottom, mLow), mClose, tailColor
'            drawBody mClose, mOpen, Color, outlineColor
'            drawTail mOpen, IIf(mHigh > Area.Top, Area.Top, mHigh), tailColor
'        ElseIf Area.Top <= mOpen And Area.Bottom >= mClose Then
'            drawBody Area.Bottom, Area.Top, Color, outlineColor
'        ElseIf Area.Top > mOpen And Area.Bottom <= mOpen And Area.Bottom >= mClose Then
'            drawTail mOpen, IIf(mHigh > Area.Top, Area.Top, mHigh), tailColor
'            drawBody Area.Bottom, mOpen, Color, outlineColor
'        Else
'            drawTail Area.Bottom, IIf(mHigh > Area.Top, Area.Top, mHigh), tailColor
'        End If
'    Else
'        If Area.Top < mOpen Then
'            drawTail IIf(mLow < Area.Bottom, Area.Bottom, mLow), Area.Top, tailColor
'        ElseIf Area.Top >= mOpen And Area.Top <= mClose And Area.Bottom < mOpen Then
'            drawTail IIf(mLow < Area.Bottom, Area.Bottom, mLow), mOpen, tailColor
'            drawBody mOpen, Area.Top, Color, outlineColor
'        ElseIf Area.Top > mClose And Area.Bottom < mOpen Then
'            drawTail IIf(mLow < Area.Bottom, Area.Bottom, mLow), mOpen, tailColor
'            drawBody mOpen, mClose, Color, outlineColor
'            drawTail mClose, IIf(mHigh > Area.Top, Area.Top, mHigh), tailColor
'        ElseIf Area.Top <= mClose And Area.Bottom >= mOpen Then
'            drawBody Area.Bottom, Area.Top, Color, outlineColor
'        ElseIf Area.Top > mClose And Area.Bottom <= mClose And Area.Bottom >= mOpen Then
'            drawTail mClose, IIf(mHigh > Area.Top, Area.Top, mHigh), tailColor
'            drawBody Area.Bottom, mClose, Color, outlineColor
'        Else
'            drawTail Area.Bottom, IIf(mHigh > Area.Top, Area.Top, mHigh), tailColor
'        End If
'    End If
'Else
'    mCanvas.SetPenAttributes outlineColor, _
'                            1, _
'                            LineStyles.LineInsideSolid, _
'                            DrawModes.DrawModeCopyPen
'    If Thickness = 1 Then
'        Dim onePixel As Double
'        onePixel = mViewport.ConvertPixelsToLogicalX(1)
'        If mOpen >= Area.Bottom And mOpen <= Area.Top Then
'            mCanvas.DrawLine mViewport.NewPoint(mX - mBarSemiWidth, mOpen), _
'                            mViewport.NewPoint(mX, mOpen)
'        End If
'        If mClose >= Area.Bottom And mClose <= Area.Top Then
'            mCanvas.DrawLine mViewport.NewPoint(mX, mClose), _
'                            mViewport.NewPoint(mX + mBarSemiWidth + onePixel, mClose)
'        End If
'        mCanvas.DrawLine mViewport.NewPoint(mX, IIf(Area.Bottom > mLow, Area.Bottom, mLow)), _
'                        mViewport.NewPoint(mX, IIf(Area.Top < mHigh, Area.Top, mHigh))
'    Else
'        mCanvas.SetBrushAttributes outlineColor, FillStyles.FillSolid
'        ThicknessScaleX = mViewport.ConvertPixelsToLogicalX(Thickness - 1)
'        ThicknessScaleY = mViewport.ConvertPixelsToLogicalY(Thickness - 1)
'
'        If mOpen >= Area.Bottom And mOpen <= Area.Top Then
'            If (mOpen - ThicknessScaleY / 2) < mLow Then
'                mCanvas.DrawRectangleFromPoints mViewport.NewPoint(mX - mBarSemiWidth, mLow), _
'                                mViewport.NewPoint(mX + ThicknessScaleX / 2, mLow + ThicknessScaleY)
'            ElseIf (mOpen + ThicknessScaleY / 2) > mHigh Then
'                mCanvas.DrawRectangleFromPoints mViewport.NewPoint(mX - mBarSemiWidth, mHigh - ThicknessScaleY), _
'                                mViewport.NewPoint(mX + ThicknessScaleX / 2, mHigh)
'            Else
'                mCanvas.DrawRectangleFromPoints mViewport.NewPoint(mX - mBarSemiWidth, mOpen - ThicknessScaleY / 2), _
'                                mViewport.NewPoint(mX + ThicknessScaleX / 2, mOpen + ThicknessScaleY / 2)
'            End If
'        End If
'
'        If mClose >= Area.Bottom And mClose <= Area.Top Then
'            If (mClose - ThicknessScaleY / 2) < mLow Then
'                mCanvas.DrawRectangleFromPoints mViewport.NewPoint(mX - ThicknessScaleX / 2, mLow), _
'                                mViewport.NewPoint(mX + mBarSemiWidth, mLow + ThicknessScaleY)
'            ElseIf (mClose + ThicknessScaleY / 2) > mHigh Then
'                mCanvas.DrawRectangleFromPoints mViewport.NewPoint(mX - ThicknessScaleX / 2, mHigh - ThicknessScaleY), _
'                                mViewport.NewPoint(mX + mBarSemiWidth, mHigh)
'            Else
'                mCanvas.DrawRectangleFromPoints mViewport.NewPoint(mX - ThicknessScaleX / 2, mClose - ThicknessScaleY / 2), _
'                                mViewport.NewPoint(mX + mBarSemiWidth, mClose + ThicknessScaleY / 2)
'            End If
'        End If
'
'        mCanvas.DrawRectangleFromPoints mViewport.NewPoint(mX - ThicknessScaleX / 2, IIf(Area.Bottom > mLow, Area.Bottom, mLow)), _
'                        mViewport.NewPoint(mX + ThicknessScaleX / 2, IIf(Area.Top < mHigh, Area.Top, mHigh))
'    End If
'End If
'
'End Sub

Private Sub drawTail(ByVal Bottom As Double, ByVal Top As Double, ByVal Color As Long)
Dim tailThicknessScaleX As Double
Const ProcName As String = "drawTail"
Dim failpoint As String
On Error GoTo Err

mCanvas.SetPenAttributes Color, _
                        1, _
                        LineStyles.LineInsideSolid, _
                        DrawModes.DrawModeCopyPen
If TailThickness > 1 Then
    mCanvas.SetBrushAttributes Color, FillStyles.FillSolid
    tailThicknessScaleX = mViewport.ConvertPixelsToLogicalX(TailThickness - 1)
    mCanvas.DrawRectangle mX - tailThicknessScaleX / 2, Bottom, _
                    mX + tailThicknessScaleX / 2, Top
Else
    mCanvas.DrawLine mX, Bottom, mX, Top
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initiatePartialRedraw( _
                ByRef areaToUndraw As TRectangle, _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "initiatePartialRedraw"
Dim failpoint As String
On Error GoTo Err

If undrawCurrentImage Then undraw areaToUndraw
If recalcBoundaries Then
    setBarSemiWidth
    calcBoundingRect
End If
mGOSP.ObjectChanged mHandle, areaToUndraw

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub initiateRedraw( _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "initiateRedraw"
Dim failpoint As String
On Error GoTo Err

initiatePartialRedraw mBoundingRect, undrawCurrentImage, recalcBoundaries

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As BarPropertyOverrideFlags) As Boolean
Const ProcName As String = "isPropertyOverrideFlagSet"
Dim failpoint As String
On Error GoTo Err

isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function prevBarObj() As Bar
Const ProcName As String = "prevBarObj"
Dim failpoint As String
On Error GoTo Err

Set prevBarObj = mPrevBarRef.Target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub setBarSemiWidth()
Select Case DisplayMode
Case BarDisplayModeBar
    mBarSemiWidth = Width / 2
Case BarDisplayModeCandlestick
    mBarSemiWidth = Width / 2
Case BarDisplayModeLine
    mBarSemiWidth = Width / 2
End Select
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As BarPropertyOverrideFlags)
Const ProcName As String = "setPropertyOverrideFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub undraw(ByRef Area As TRectangle)
Const ProcName As String = "undraw"
Dim failpoint As String
On Error GoTo Err

If Not mVisible Then Exit Sub
If Not Area.isValid Then
    mGOSP.ObjectUndrawn mHandle, mBoundingRect
Else
    mGOSP.ObjectUndrawn mHandle, Area
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



