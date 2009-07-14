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

Event SelectionStateChanged()

'================================================================================
' Enums
'================================================================================

Private Enum PropertyOverrideFlags
    BarColorIsSet = 1
    UpColorIsSet = 2
    DownColorIsSet = 4
    DisplayModeIsSet = 8
    SolidUpBodyIsSet = &H10&
    BarThicknessIsSet = &H20&
    BarWidthIsSet = &H40&
    TailThicknessIsSet = &H80&
    OutlineThicknessIsSet = &H100&
End Enum

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

Private mChart As Chart
Private mSeriesID As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mBars As Collection

Private mCanvas As Canvas
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As BarStyle
Attribute mStyle.VB_VarHelpID = -1

Private mIncludeInAutoscale As Boolean

Private mName As String

' these fields override corresponding items in the Style
Private mBarColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayMode As BarDisplayModes
Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mBarWidth As Single
Private mTailThickness As Long
Private mOutlineThickness As Long

Private mPropertyOverrideFlags As PropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
Set mStyle = New BarStyle
End Sub

Private Sub Class_Terminate()
Debug.Print "BarSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Let IGraphicObjectSeries_Canvas(ByVal value As Canvas)
Set mCanvas = value
End Property

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
Set mChart = Nothing
Set mGOSP = Nothing
Set mCanvas = Nothing
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal Name As String, _
                ByVal Canvas As Canvas, _
                ByVal Layer As LayerNumbers, _
                ByVal pChart As Chart, _
                ByVal SeriesID As Long, _
                ByVal GOSP As GraphObjServiceProvider)
mName = Name
mSeriesID = SeriesID
Set mGOSP = GOSP
mLayer = Layer
Set mCanvas = Canvas
Set mChart = pChart
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
Case UCase$("BarColor")
    repaintRegion False
Case UCase$("BarThickness")
    repaintRegion False
Case UCase$("BarWidth")
    repaintRegion True
Case UCase$("displayMode")
    repaintRegion True
Case UCase$("downColor")
    repaintRegion False
Case UCase$("OutlineThickness")
    repaintRegion False
Case UCase$("SolidUpBody")
    repaintRegion False
Case UCase$("TailThickness")
    repaintRegion False
Case UCase$("upColor")
    repaintRegion False
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
' Properties
'================================================================================

Public Property Get BarColor() As OLE_COLOR
BarColor = mStyle.BarColor
If isPropertyOverrideFlagSet(BarColorIsSet) Then BarColor = mBarColor
End Property

Public Property Let BarColor(ByVal value As OLE_COLOR)
If mBarColor = value And isPropertyOverrideFlagSet(BarColorIsSet) Then Exit Property
mBarColor = value
setPropertyOverrideFlag BarColorIsSet
repaintRegion False
End Property

Public Property Get BarThickness() As Long
BarThickness = mStyle.BarThickness
If isPropertyOverrideFlagSet(BarThicknessIsSet) Then BarThickness = mBarThickness
End Property

Public Property Let BarThickness(ByVal value As Long)
If mBarThickness = value And isPropertyOverrideFlagSet(BarThicknessIsSet) Then Exit Property
mBarThickness = value
setPropertyOverrideFlag BarThicknessIsSet
repaintRegion False
End Property

Public Property Get BarWidth() As Single
BarWidth = mStyle.BarWidth
If isPropertyOverrideFlagSet(BarWidthIsSet) Then BarWidth = mBarWidth
End Property

Public Property Let BarWidth(ByVal value As Single)
If mBarWidth = value And isPropertyOverrideFlagSet(BarWidthIsSet) Then Exit Property
mBarWidth = value
setPropertyOverrideFlag BarWidthIsSet
repaintRegion True
End Property

Public Property Get Count() As Long
Count = mBars.Count
End Property

Public Property Get DisplayMode() As BarDisplayModes
DisplayMode = mStyle.DisplayMode
If isPropertyOverrideFlagSet(DisplayModeIsSet) Then DisplayMode = mDisplayMode
End Property

Public Property Let DisplayMode(ByVal value As BarDisplayModes)
If mDisplayMode = value And isPropertyOverrideFlagSet(DisplayModeIsSet) Then Exit Property
mDisplayMode = value
setPropertyOverrideFlag DisplayModeIsSet
repaintRegion True
End Property

Public Property Get DownColor() As OLE_COLOR
DownColor = mStyle.DownColor
If isPropertyOverrideFlagSet(DownColorIsSet) Then DownColor = mDownColor
End Property

Public Property Let DownColor(ByVal value As OLE_COLOR)
If mDownColor = value And isPropertyOverrideFlagSet(DownColorIsSet) Then Exit Property
mDownColor = value
setPropertyOverrideFlag DownColorIsSet
repaintRegion False
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value Then Exit Property
mIncludeInAutoscale = value
repaintRegion False
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

Public Property Let Layer(ByVal value As LayerNumbers)
If mLayer = value Then Exit Property
mLayer = value
repaintRegion False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mLayer
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get OutlineThickness() As Long
OutlineThickness = mStyle.OutlineThickness
If isPropertyOverrideFlagSet(OutlineThicknessIsSet) Then OutlineThickness = mOutlineThickness
End Property

Public Property Let OutlineThickness(ByVal value As Long)
If mOutlineThickness = value And isPropertyOverrideFlagSet(OutlineThicknessIsSet) Then Exit Property
mOutlineThickness = value
setPropertyOverrideFlag OutlineThicknessIsSet
repaintRegion False
End Property

Public Property Get SolidUpBody() As Boolean
SolidUpBody = mStyle.SolidUpBody
If isPropertyOverrideFlagSet(SolidUpBodyIsSet) Then SolidUpBody = mSolidUpBody
End Property

Public Property Let SolidUpBody(ByVal value As Boolean)
If mSolidUpBody = value And isPropertyOverrideFlagSet(SolidUpBodyIsSet) Then Exit Property
mSolidUpBody = value
setPropertyOverrideFlag SolidUpBodyIsSet
repaintRegion False
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

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Property

Public Property Get TailThickness() As Long
TailThickness = mStyle.TailThickness
If isPropertyOverrideFlagSet(TailThicknessIsSet) Then TailThickness = mTailThickness
End Property

Public Property Let TailThickness(ByVal value As Long)
If mTailThickness = value And isPropertyOverrideFlagSet(TailThicknessIsSet) Then Exit Property
mTailThickness = value
setPropertyOverrideFlag TailThicknessIsSet
repaintRegion False
End Property

Public Property Get UpColor() As OLE_COLOR
UpColor = mStyle.UpColor
If isPropertyOverrideFlagSet(UpColorIsSet) Then UpColor = mUpColor
End Property

Public Property Let UpColor(ByVal value As OLE_COLOR)
If mUpColor = value And isPropertyOverrideFlagSet(UpColorIsSet) Then Exit Property
mUpColor = value
setPropertyOverrideFlag UpColorIsSet
repaintRegion False
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

Add.X = mChart.GetXFromTimestamp(Timestamp, True)
If Key = "" Then
    Key = GenerateGUIDString
End If
Add.Initialise Me, Timestamp, Key, mGOSP, mCanvas, LocalStyle

mBars.Add Add, Key
mGOSP.AddGraphicObject Add, mSeriesID
End Function

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
repaintRegion True
End Sub

Friend Sub Finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Public Function Item(ByVal index As Variant) As Bar
Set Item = mBars(index)
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.Bar)
On Error GoTo Err
mBars.Remove value.Key
mGOSP.RemoveObject value.Handle

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

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As PropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub repaintRegion( _
                ByVal recalcBoundaries As Boolean)
If mBars.Count <> 0 Then
    If recalcBoundaries Then
        mGOSP.RecalcBoundary mSeriesID
        mGOSP.PaintRegion True
    Else
        mGOSP.PaintRegion False
    End If
End If
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub



