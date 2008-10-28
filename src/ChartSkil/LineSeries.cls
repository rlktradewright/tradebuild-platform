VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineSeries"
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
Private Const ModuleName                As String = "LineSeries"

Private mName As String
Private mSeriesID As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mLines As Collection

'Private mSurface As PictureBox
Private mCanvas As Canvas
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private mStyle As linestyle
Private mInhibitChangeEvent As Boolean

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mLines = New Collection
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
IGraphicObjectSeries_count = mLines.Count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
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
If mStyle Is Nothing Then Set mStyle = controller.DefaultLineStyle.clone
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mLines.Item(index)
End Function

Private Property Get IGraphicObjectSeries_name() As String
IGraphicObjectSeries_name = mName
End Property

Public Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mLines.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lLine As ChartSkil26.Line

On Error Resume Next
Set lLine = value

mLines.Remove lLine.Key
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

Public Property Get arrowEndColor() As Long
arrowEndColor = mStyle.arrowEndColor
End Property

Public Property Let arrowEndColor(ByVal value As Long)
If mStyle.arrowEndColor = value Then Exit Property
mStyle.arrowEndColor = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowEndFillColor() As Long
arrowEndFillColor = mStyle.arrowEndFillColor
End Property

Public Property Let arrowEndFillColor(ByVal value As Long)
If mStyle.arrowEndFillColor = value Then Exit Property
mStyle.arrowEndFillColor = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowEndFillStyle() As FillStyles
arrowEndFillStyle = mStyle.arrowEndFillStyle
End Property

Public Property Let arrowEndFillStyle(ByVal value As FillStyles)
If mStyle.arrowEndFillStyle = value Then Exit Property
mStyle.arrowEndFillStyle = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowEndLength() As Long
arrowEndLength = mStyle.arrowEndLength
End Property

Public Property Let arrowEndLength(ByVal value As Long)
If mStyle.arrowEndLength = value Then Exit Property
mStyle.arrowEndLength = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowEndStyle() As ArrowStyles
arrowEndStyle = mStyle.arrowEndStyle
End Property

Public Property Let arrowEndStyle(ByVal value As ArrowStyles)
If mStyle.arrowEndStyle = value Then Exit Property
mStyle.arrowEndStyle = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowEndWidth() As Long
arrowEndWidth = mStyle.arrowEndWidth
End Property

Public Property Let arrowEndWidth(ByVal value As Long)
If mStyle.arrowEndWidth = value Then Exit Property
mStyle.arrowEndWidth = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowStartColor() As Long
arrowStartColor = mStyle.arrowStartColor
End Property

Public Property Let arrowStartColor(ByVal value As Long)
If mStyle.arrowStartColor = value Then Exit Property
mStyle.arrowStartColor = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowStartFillColor() As Long
arrowStartFillColor = mStyle.arrowStartFillColor
End Property

Public Property Let arrowStartFillColor(ByVal value As Long)
If mStyle.arrowStartFillColor = value Then Exit Property
mStyle.arrowStartFillColor = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowStartfillstyle() As FillStyles
arrowStartfillstyle = arrowStartfillstyle
End Property

Public Property Let arrowStartfillstyle(ByVal value As FillStyles)
If mStyle.arrowStartfillstyle = value Then Exit Property
mStyle.arrowStartfillstyle = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowStartLength() As Long
arrowStartLength = mStyle.arrowStartLength
End Property

Public Property Let arrowStartLength(ByVal value As Long)
If mStyle.arrowStartLength = value Then Exit Property
mStyle.arrowStartLength = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowStartStyle() As ArrowStyles
arrowStartStyle = mStyle.arrowStartStyle
End Property

Public Property Let arrowStartStyle(ByVal value As ArrowStyles)
If mStyle.arrowStartStyle = value Then Exit Property
mStyle.arrowStartStyle = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get arrowStartWidth() As Long
arrowStartWidth = mStyle.arrowStartWidth
End Property

Public Property Let arrowStartWidth(ByVal value As Long)
If mStyle.arrowStartWidth = value Then Exit Property
mStyle.arrowStartWidth = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Friend Property Let Canvas(ByVal value As Canvas)
Set mCanvas = value
End Property

Public Property Get Color() As Long
Color = mStyle.Color
End Property

Public Property Let Color(ByVal value As Long)
If mStyle.Color = value Then Exit Property
mStyle.Color = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get Count() As Long
Count = mLines.Count
End Property

Friend Property Get extended() As Boolean
extended = mStyle.extended
End Property

Public Property Let extended(ByVal value As Boolean)
If mStyle.extended = value Then Exit Property
mStyle.extended = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get extendAfter() As Boolean
extendAfter = mStyle.extendAfter
End Property

Public Property Let extendAfter(ByVal value As Boolean)
If mStyle.extendAfter = value Then Exit Property
mStyle.extendAfter = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get extendBefore() As Boolean
extendBefore = mStyle.extendBefore
End Property

Public Property Let extendBefore(ByVal value As Boolean)
If mStyle.extendBefore = value Then Exit Property
mStyle.extendBefore = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Friend Property Get fixedX() As Boolean
fixedX = mStyle.fixedX
End Property

Public Property Let fixedX(ByVal value As Boolean)
If mStyle.fixedX = value Then Exit Property
mStyle.fixedX = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Friend Property Get fixedY() As Boolean
fixedY = mStyle.fixedY
End Property

Public Property Let fixedY(ByVal value As Boolean)
If mStyle.fixedY = value Then Exit Property
mStyle.fixedY = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mStyle.includeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mStyle.includeInAutoscale = value Then Exit Property
mStyle.includeInAutoscale = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion True
End If
fireStyleChanged
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

'Public Property Get keepInView() As Boolean
'keepInView = mStyle.KeepInView
'End Property

'Public Property Let keepInView(ByVal value As Boolean)
'If mStyle.KeepInView = value Then Exit Property
'undraw
'mStyle.KeepInView = value
'update
'End Property

Public Property Let Layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get Layer() As Long
Layer = mLayer
End Property

Public Property Get linestyle() As LineStyles
linestyle = mStyle.linestyle
End Property

Public Property Let linestyle(ByVal value As LineStyles)
If mStyle.linestyle = value Then Exit Property
mStyle.linestyle = value
If mLines.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get style() As linestyle
Set style = mStyle.clone
End Property

Public Property Let style(ByVal value As linestyle)
Dim somethingChanged As Boolean
Dim recalcBoundaries As Boolean
Dim recalcRegion As Boolean

Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "style", _
            "Style is null reference"
End If

mInhibitChangeEvent = True

If value.arrowEndColor <> mStyle.arrowEndColor Then
    mStyle.arrowEndColor = value.arrowEndColor
    somethingChanged = True
End If
If value.arrowEndFillColor <> mStyle.arrowEndFillColor Then
    mStyle.arrowEndFillColor = value.arrowEndFillColor
    somethingChanged = True
End If
If value.arrowEndFillStyle <> mStyle.arrowEndFillStyle Then
    mStyle.arrowEndFillStyle = value.arrowEndFillStyle
    somethingChanged = True
End If
If value.arrowEndLength <> mStyle.arrowEndLength Then
    mStyle.arrowEndLength = value.arrowEndLength
    somethingChanged = True
    recalcBoundaries = True
End If
If value.arrowEndStyle <> mStyle.arrowEndStyle Then
    mStyle.arrowEndStyle = value.arrowEndStyle
    somethingChanged = True
End If
If value.arrowEndWidth <> mStyle.arrowEndWidth Then
    mStyle.arrowEndWidth = value.arrowEndWidth
    somethingChanged = True
    recalcBoundaries = True
End If
If value.arrowStartColor <> mStyle.arrowStartColor Then
    mStyle.arrowStartColor = value.arrowStartColor
    somethingChanged = True
End If
If value.arrowStartFillColor <> mStyle.arrowStartFillColor Then
    mStyle.arrowStartFillColor = value.arrowStartFillColor
    somethingChanged = True
End If
If value.arrowStartfillstyle <> mStyle.arrowStartfillstyle Then
    mStyle.arrowStartfillstyle = value.arrowStartfillstyle
    somethingChanged = True
End If
If value.arrowStartLength <> mStyle.arrowStartLength Then
    mStyle.arrowStartLength = value.arrowStartLength
    somethingChanged = True
    recalcBoundaries = True
End If
If value.arrowStartStyle <> mStyle.arrowStartStyle Then
    mStyle.arrowStartStyle = value.arrowStartStyle
    somethingChanged = True
End If
If value.arrowStartWidth <> mStyle.arrowStartWidth Then
    mStyle.arrowStartWidth = value.arrowStartWidth
    somethingChanged = True
    recalcBoundaries = True
End If
If value.Color <> mStyle.Color Then
    mStyle.Color = value.Color
    somethingChanged = True
End If
If value.extendAfter <> mStyle.extendAfter Then
    mStyle.extendAfter = value.extendAfter
    somethingChanged = True
    recalcBoundaries = True
End If
If value.extendBefore <> mStyle.extendBefore Then
    mStyle.extendBefore = value.extendBefore
    somethingChanged = True
    recalcBoundaries = True
End If
If value.extended <> mStyle.extended Then
    mStyle.extended = value.extended
    somethingChanged = True
    recalcBoundaries = True
End If
If value.fixedX <> mStyle.fixedX Then
    mStyle.fixedX = value.fixedX
    somethingChanged = True
    recalcBoundaries = True
End If
If value.fixedY <> mStyle.fixedY Then
    mStyle.fixedY = value.fixedY
    somethingChanged = True
    recalcBoundaries = True
End If
If value.includeInAutoscale <> mStyle.includeInAutoscale Then
    mStyle.includeInAutoscale = value.includeInAutoscale
    somethingChanged = True
    recalcBoundaries = True
    recalcRegion = True
End If
If value.linestyle <> mStyle.linestyle Then
    mStyle.linestyle = value.linestyle
    somethingChanged = True
End If
If value.thickness <> mStyle.thickness Then
    mStyle.thickness = value.thickness
    somethingChanged = True
    recalcBoundaries = True
End If

mInhibitChangeEvent = False

If somethingChanged Then
    If recalcBoundaries Then If mLines.Count <> 0 Then mGOSP.RecalcBoundary mSeriesID
    If mLines.Count <> 0 Then mGOSP.PaintRegion recalcRegion
    fireStyleChanged
End If

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "style" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get thickness() As Long
thickness = mStyle.thickness
End Property

Public Property Let thickness(ByVal value As Long)
If mStyle.thickness = value Then Exit Property
mStyle.thickness = value
If mLines.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As linestyle, _
                Optional ByVal IsSelectable As Boolean) As ChartSkil26.Line
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.LineSeries::add", _
                            "Line series is finished"
Set Add = New ChartSkil26.Line
If Key = "" Then
    Key = GenerateGUIDString
End If
Add.initialise mSeriesID, Key, mGOSP, mLayer, mCanvas, mStyle, LocalStyle, IsSelectable

mLines.Add Add, Key
mGOSP.AddGraphicObject Add, mSeriesID
End Function

Friend Sub Finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Friend Sub initialise( _
                ByVal name As String, _
                ByVal style As linestyle)
mName = name
Set mStyle = style.clone    ' style cannot be shared between LineSeries
End Sub

Public Function Item(ByVal index As Variant) As ChartSkil26.Line
Attribute Item.VB_UserMemId = 0
Set Item = mLines(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = mLines.[_NewEnum]
End Function

Friend Sub RemoveAll()
StartTask Me, PriorityLow
End Sub

Public Sub RemoveObject(ByVal value As ChartSkil26.Line)
On Error GoTo Err
mLines.Remove value.Key
mGOSP.RemoveObject value.Handle

Exit Sub

Err:
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub fireStyleChanged()
If Not mInhibitChangeEvent Then RaiseEvent StyleChanged
End Sub


