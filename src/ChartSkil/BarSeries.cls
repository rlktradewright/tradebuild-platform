VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements IGraphicObjectSeries

'================================================================================
' Events
'================================================================================

Event Click()

Event DblCLick()

Event PropertyChanged(ev As PropertyChangedEvent)

Event SelectionStateChanged()

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "BarSeries"

'================================================================================
' Member variables
'================================================================================

Private mChartRef As WeakReference
Private mSeriesID As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mBars As Collection

Private mViewport As Viewport
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As BarStyle
Attribute mStyle.VB_VarHelpID = -1

Private mIncludeInAutoscale As Boolean

Private mName As String

Private mPrevBarRef As WeakReference

' these fields override corresponding items in the Style
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayMode As BarDisplayModes
Private mSolidUpBody As Boolean
Private mThickness As Long
Private mWidth As Single
Private mTailThickness As Long
Private mOutlineThickness As Long

Private mPropertyOverrideFlags As BarPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
Set mStyle = New BarStyle
End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "BarSeries terminated"
Debug.Print "BarSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Sub IGraphicObjectSeries_Click()
RaiseEvent Click
End Sub

Private Sub IGraphicObjectSeries_DblCLick()
RaiseEvent DblCLick
End Sub

Private Property Get IGraphicObjectSeries_count() As Long
IGraphicObjectSeries_count = mBars.Count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
Set mChartRef = Nothing
Set mPrevBarRef = Nothing
Set mGOSP = Nothing
Set mViewport = Nothing
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal pName As String, _
                ByVal pviewport As Viewport, _
                ByVal pLayer As LayerNumbers, _
                ByVal pChart As Chart, _
                ByVal pSeriesID As Long, _
                ByVal pGOSP As GraphObjServiceProvider)
mName = pName
mSeriesID = pSeriesID
Set mGOSP = pGOSP
mLayer = pLayer
Set mViewport = pviewport
Set mChartRef = CreateWeakReference(pChart)
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mBars.Item(index)
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Public Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mBars.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lBar As ChartSkil26.Bar

On Error Resume Next
Set lBar = value

mBars.Remove lBar.Key
End Sub

Private Property Let IGraphicObjectSeries_Selected(ByVal RHS As Boolean)
If RHS <> mIsSelected Then
    mIsSelected = RHS
    firePropertyChanged "IsSelected"
    RaiseEvent SelectionStateChanged
End If
End Property

Private Property Get IGraphicObjectSeries_Selected() As Boolean
IGraphicObjectSeries_Selected = mIsSelected
End Property

Private Property Get IGraphicObjectSeries_SeriesID() As Long
IGraphicObjectSeries_SeriesID = mSeriesID
End Property

'================================================================================
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "COLOR"
    repaintRegion False
Case "THICKNESS"
    repaintRegion False
Case "WIDTH"
    repaintRegion True
Case "DISPLAYMODE"
    repaintRegion True
Case "DOWNCOLOR"
    repaintRegion False
Case "INCLUDEINAUTOSCALE"
    repaintRegion True
Case "OUTLINETHICKNESS"
    repaintRegion False
Case "SOLIDUPBODY"
    repaintRegion False
Case "TAILTHICKNESS"
    repaintRegion False
Case "UPCOLOR"
    repaintRegion False
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged", _
            "Unhandled property change"
End Select

firePropertyChanged ev.PropertyName

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

Public Property Get Color() As Long
Color = mStyle.Color
If isPropertyOverrideFlagSet(BarIsSetColor) Then Color = mColor
End Property

Public Property Let Color(ByVal value As Long)
If mColor = value And isPropertyOverrideFlagSet(BarIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag BarIsSetColor
repaintRegion False
firePropertyChanged "Color"
End Property

Public Property Get Count() As Long
Count = mBars.Count
End Property

Public Property Get DisplayMode() As BarDisplayModes
DisplayMode = mStyle.DisplayMode
If isPropertyOverrideFlagSet(BarIsSetDisplayMode) Then DisplayMode = mDisplayMode
End Property

Public Property Let DisplayMode(ByVal value As BarDisplayModes)
If mDisplayMode = value And isPropertyOverrideFlagSet(BarIsSetDisplayMode) Then Exit Property
mDisplayMode = value
setPropertyOverrideFlag BarIsSetDisplayMode
repaintRegion True
firePropertyChanged "DisplayMode"
End Property

Public Property Get DownColor() As Long
DownColor = mStyle.DownColor
If isPropertyOverrideFlagSet(BarIsSetDownColor) Then DownColor = mDownColor
End Property

Public Property Let DownColor(ByVal value As Long)
If mDownColor = value And isPropertyOverrideFlagSet(BarIsSetDownColor) Then Exit Property
mDownColor = value
setPropertyOverrideFlag BarIsSetDownColor
repaintRegion False
firePropertyChanged "DownColor"
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mStyle.IncludeInAutoscale
If isPropertyOverrideFlagSet(BarIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(BarIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag BarIsSetIncludeInAutoscale
repaintRegion False
firePropertyChanged "IncludeInAutoscale"
End Property

Public Property Let IsSelectable(ByVal value As Boolean)
mIsSelectable = value
firePropertyChanged "IsSelectable"
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Let Layer(ByVal value As LayerNumbers)
If mLayer = value Then Exit Property
mLayer = value
repaintRegion False
firePropertyChanged "Layer"
End Property

Public Property Get Layer() As LayerNumbers
Layer = mLayer
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal value As String)
mName = value
firePropertyChanged "Name"
End Property

Public Property Get OutlineThickness() As Long
OutlineThickness = mStyle.OutlineThickness
If isPropertyOverrideFlagSet(BarIsSetOutlineThickness) Then OutlineThickness = mOutlineThickness
End Property

Public Property Let OutlineThickness(ByVal value As Long)
If mOutlineThickness = value And isPropertyOverrideFlagSet(BarIsSetOutlineThickness) Then Exit Property
mOutlineThickness = value
setPropertyOverrideFlag BarIsSetOutlineThickness
repaintRegion False
firePropertyChanged "OutlineThickness"
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get SolidUpBody() As Boolean
SolidUpBody = mStyle.SolidUpBody
If isPropertyOverrideFlagSet(BarIsSetSolidUpBody) Then SolidUpBody = mSolidUpBody
End Property

Public Property Let SolidUpBody(ByVal value As Boolean)
If mSolidUpBody = value And isPropertyOverrideFlagSet(BarIsSetSolidUpBody) Then Exit Property
mSolidUpBody = value
setPropertyOverrideFlag BarIsSetSolidUpBody
repaintRegion False
firePropertyChanged "SolidUpBody"
End Property

Public Property Get Style() As BarStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal value As BarStyle)
Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Style", _
            "Style is null reference"
End If

Set mStyle = value
repaintRegion True

firePropertyChanged "Style"
Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Property

Public Property Get TailThickness() As Long
TailThickness = mStyle.TailThickness
If isPropertyOverrideFlagSet(BarIsSetTailThickness) Then TailThickness = mTailThickness
End Property

Public Property Let TailThickness(ByVal value As Long)
If mTailThickness = value And isPropertyOverrideFlagSet(BarIsSetTailThickness) Then Exit Property
mTailThickness = value
setPropertyOverrideFlag BarIsSetTailThickness
repaintRegion False
firePropertyChanged "TailThickness"
End Property

Public Property Get Thickness() As Long
Thickness = mStyle.Thickness
If isPropertyOverrideFlagSet(BarIsSetThickness) Then Thickness = mThickness
End Property

Public Property Let Thickness(ByVal value As Long)
If mThickness = value And isPropertyOverrideFlagSet(BarIsSetThickness) Then Exit Property
mThickness = value
setPropertyOverrideFlag BarIsSetThickness
repaintRegion False
firePropertyChanged "Thickness"
End Property

Public Property Get UpColor() As Long
UpColor = mStyle.UpColor
If isPropertyOverrideFlagSet(BarIsSetUpColor) Then UpColor = mUpColor
End Property

Public Property Let UpColor(ByVal value As Long)
If mUpColor = value And isPropertyOverrideFlagSet(BarIsSetUpColor) Then Exit Property
mUpColor = value
setPropertyOverrideFlag BarIsSetUpColor
repaintRegion False
firePropertyChanged "UpColor"
End Property

Public Property Get Width() As Single
Width = mStyle.Width
If isPropertyOverrideFlagSet(BarIsSetWidth) Then Width = mWidth
End Property

Public Property Let Width(ByVal value As Single)
If mWidth = value And isPropertyOverrideFlagSet(BarIsSetWidth) Then Exit Property
mWidth = value
setPropertyOverrideFlag BarIsSetWidth
repaintRegion True
firePropertyChanged "Width"
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                ByVal Timestamp As Date, _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As BarStyle) As Bar

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.BarSeries::Add", _
                            "Bar series is finished"

Set Add = New Bar

If Key = "" Then
    Key = GenerateGUIDString
End If
Add.Initialise Me, Timestamp, Key, mGOSP, mViewport, LocalStyle

Add.X = chartObj.GetXFromTimestamp(Timestamp, True)

If Not mPrevBarRef Is Nothing Then Add.PreviousBar = prevBarObj
mBars.Add Add, Key
Set mPrevBarRef = CreateWeakReference(Add)
mGOSP.AddGraphicObject Add
End Function

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
repaintRegion True
End Sub

Friend Sub Finish()
mFinished = True
Set mViewport = Nothing
End Sub

Public Function Item(ByVal index As Variant) As Bar
Set Item = mBars(index)
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.Bar)
On Error GoTo Err
mBars.Remove value.Key
mGOSP.RemoveObject value.handle

Exit Sub

Err:
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mBars.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As BarPropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEvent

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As BarPropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Function prevBarObj() As Bar
Set prevBarObj = mPrevBarRef.Target
End Function

Private Sub repaintRegion( _
                ByVal recalcBoundaries As Boolean)
If mBars.Count <> 0 Then
    If recalcBoundaries Then
        mSeriesUpdateNumber = mSeriesUpdateNumber + 1
        mGOSP.RecalcBoundary mSeriesID
        mGOSP.PaintRegion True
    Else
        mGOSP.PaintRegion False
    End If
End If
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As BarPropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub



