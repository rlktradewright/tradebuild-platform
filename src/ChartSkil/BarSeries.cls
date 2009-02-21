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

Event StyleChanged()

Event SelectionStateChanged()

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "BarSeries"

Private mController As ChartController
Private mSeriesID As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mBars As Collection

Private mCanvas As Canvas
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private mStyle As BarStyle
Private mInhibitChangeEvent As Boolean

Private mName As String

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
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
Set mController = Nothing
Set mGOSP = Nothing
Set mCanvas = Nothing
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal name As String, _
                ByVal Canvas As Canvas, _
                ByVal Layer As LayerNumbers, _
                ByVal controller As ChartController, _
                ByVal seriesId As Long, _
                ByVal GOSP As GraphObjServiceProvider)
mName = name
mSeriesID = seriesId
Set mGOSP = GOSP
mLayer = Layer
Set mCanvas = Canvas
Set mController = controller
If mStyle Is Nothing Then Set mStyle = controller.DefaultBarStyle.clone
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mBars.Item(index)
End Function

Private Property Get IGraphicObjectSeries_name() As String
IGraphicObjectSeries_name = mName
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
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get barColor() As Long
barColor = mStyle.barColor
End Property

Public Property Let barColor(ByVal value As Long)
If mStyle.barColor = value Then Exit Property
mStyle.barColor = value
If mBars.Count <> 0 Then mGOSP.PaintRegion False
fireStyleChanged
End Property

Public Property Get barThickness() As Long
barThickness = mStyle.barThickness
End Property

Public Property Let barThickness(ByVal value As Long)
If mStyle.barThickness = value Then Exit Property
mStyle.barThickness = value
If mBars.Count <> 0 Then mGOSP.PaintRegion False
fireStyleChanged
End Property

Public Property Get barWidth() As Single
barWidth = mStyle.barWidth
End Property

Public Property Let barWidth(ByVal value As Single)
If mStyle.barWidth = value Then Exit Property
mStyle.barWidth = value
If mBars.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get Count() As Long
Count = mBars.Count
End Property

Public Property Get displayMode() As BarDisplayModes
displayMode = mStyle.displayMode
End Property

Public Property Let displayMode(ByVal value As BarDisplayModes)
If mStyle.displayMode = value Then Exit Property
mStyle.displayMode = value
If mBars.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get downColor() As Long
downColor = mStyle.downColor
End Property

Public Property Let downColor(ByVal value As Long)
If mStyle.downColor = value Then Exit Property
mStyle.downColor = value
If mBars.Count <> 0 Then mGOSP.PaintRegion False
fireStyleChanged
End Property

Friend Property Get includeInAutoscale() As Boolean
includeInAutoscale = mStyle.includeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mStyle.includeInAutoscale = value Then Exit Property
mStyle.includeInAutoscale = value
If mBars.Count <> 0 Then
    mGOSP.PaintRegion True
End If
fireStyleChanged
End Property

Public Property Let isSelectable(ByVal value As Boolean)
mIsSelectable = value
End Property

Public Property Get isSelectable() As Boolean
isSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Get Layer() As Long
Layer = mLayer
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get outlineThickness() As Long
outlineThickness = mStyle.outlineThickness
End Property

Public Property Let outlineThickness(ByVal value As Long)
If mStyle.outlineThickness = value Then Exit Property
mStyle.outlineThickness = value
If mBars.Count <> 0 Then mGOSP.PaintRegion False
fireStyleChanged
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mStyle.solidUpBody
End Property

Public Property Let solidUpBody(ByVal value As Boolean)
If mStyle.solidUpBody = value Then Exit Property
mStyle.solidUpBody = value
If mBars.Count <> 0 Then mGOSP.PaintRegion False
fireStyleChanged
End Property

Public Property Get style() As BarStyle
Set style = mStyle.clone
End Property

Public Property Let style( _
                ByVal value As BarStyle)
Dim somethingChanged As Boolean

Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "style", _
            "Style is null reference"
End If

mInhibitChangeEvent = True

If value.barColor <> mStyle.barColor Then
    somethingChanged = True
    mStyle.barColor = value.barColor
End If
If value.barThickness <> mStyle.barThickness Then
    somethingChanged = True
    mStyle.barThickness = value.barThickness
End If
If value.barWidth <> mStyle.barWidth Then
    somethingChanged = True
    mStyle.barWidth = value.barWidth
    If mBars.Count <> 0 Then mGOSP.RecalcBoundary mSeriesID
End If
If value.displayMode <> mStyle.displayMode Then
    somethingChanged = True
    mStyle.displayMode = value.displayMode
End If
If value.downColor <> mStyle.downColor Then
    somethingChanged = True
    mStyle.downColor = value.downColor
End If
If value.outlineThickness <> mStyle.outlineThickness Then
    somethingChanged = True
    mStyle.outlineThickness = value.outlineThickness
End If
If value.solidUpBody <> mStyle.solidUpBody Then
    somethingChanged = True
    mStyle.solidUpBody = value.solidUpBody
End If
If value.tailThickness <> mStyle.tailThickness Then
    somethingChanged = True
    mStyle.tailThickness = value.tailThickness
End If
If value.upColor <> mStyle.upColor Then
    somethingChanged = True
    mStyle.upColor = value.upColor
End If

mInhibitChangeEvent = False

If somethingChanged Then
    If mBars.Count <> 0 Then mGOSP.PaintRegion False
    fireStyleChanged
End If

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "style" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Property

Public Property Get tailThickness() As Long
tailThickness = mStyle.tailThickness
End Property

Public Property Let tailThickness(ByVal value As Long)
If mStyle.tailThickness = value Then Exit Property
mStyle.tailThickness = value
If mBars.Count <> 0 Then mGOSP.PaintRegion False
fireStyleChanged
End Property

Public Property Get upColor() As Long
upColor = mStyle.upColor
End Property

Public Property Let upColor(ByVal value As Long)
If mStyle.upColor = value Then Exit Property
mStyle.upColor = value
If mBars.Count <> 0 Then mGOSP.PaintRegion False
fireStyleChanged
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal timestamp As Date, _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As BarStyle, _
                Optional ByVal isSelectable As Boolean) As Bar

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.BarSeries::add", _
                            "Bar series is finished"

Set add = New Bar

add.x = gCalculateX(timestamp, mController, True)
If Key = "" Then
    Key = GenerateGUIDString
End If
add.initialise mSeriesID, timestamp, Key, mGOSP, mLayer, mCanvas, mStyle, LocalStyle, isSelectable

mBars.add add, Key
mGOSP.AddGraphicObject add, mSeriesID
End Function

Friend Sub Finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Friend Sub initialise( _
                ByVal name As String, _
                ByVal style As BarStyle)
mName = name
Set mStyle = style.clone    ' style cannot be shared between BarSeries
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

Private Sub fireStyleChanged()
If Not mInhibitChangeEvent Then RaiseEvent StyleChanged
End Sub

