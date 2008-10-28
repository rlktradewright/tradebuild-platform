VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextSeries"
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
Private Const ModuleName                As String = "TextSeries"

Private mSeriesID As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mTexts As Collection

Private mCanvas As Canvas
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private mStyle As TextStyle
Private mInhibitChangeEvent As Boolean

Private mName As String

Private mTaskContext As TaskContext

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTexts = New Collection
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
IGraphicObjectSeries_count = mTexts.Count
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
If mStyle Is Nothing Then Set mStyle = controller.DefaultTextStyle.clone
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mTexts.Item(index)
End Function

Private Property Get IGraphicObjectSeries_name() As String
IGraphicObjectSeries_name = mName
End Property

Public Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mTexts.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lText As ChartSkil26.text

On Error Resume Next
Set lText = value

mTexts.Remove lText.Key
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

Public Property Get align() As TextAlignModes
align = mStyle.align
End Property

Public Property Let align(ByVal value As TextAlignModes)
If mStyle.align = value Then Exit Property
mStyle.align = value
If mTexts.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Friend Property Get box() As Boolean
box = mStyle.box
End Property

Public Property Let box(ByVal value As Boolean)
If mStyle.box = value Then Exit Property
mStyle.box = value
If mTexts.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get boxColor() As Long
boxColor = mStyle.boxColor
End Property

Public Property Let boxColor(ByVal value As Long)
If mStyle.boxColor = value Then Exit Property
mStyle.boxColor = value
If mTexts.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Friend Property Get boxFillColor() As Long
boxFillColor = mStyle.boxFillColor
End Property

Public Property Let boxFillColor(ByVal value As Long)
If mStyle.boxFillColor = value Then Exit Property
mStyle.boxFillColor = value
If mTexts.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get boxFillStyle() As FillStyles
boxFillStyle = mStyle.boxFillStyle
End Property

Public Property Let boxFillStyle(ByVal value As FillStyles)
If mStyle.boxFillStyle = value Then Exit Property
mStyle.boxFillStyle = value
If mTexts.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get boxStyle() As LineStyles
boxStyle = mStyle.boxStyle
End Property

Public Property Let boxStyle(ByVal value As LineStyles)
If mStyle.boxStyle = value Then Exit Property
mStyle.boxStyle = value
If mTexts.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get boxThickness() As Long
boxThickness = mStyle.boxThickness
End Property

Public Property Let boxThickness(ByVal value As Long)
If mStyle.boxThickness = value Then Exit Property
mStyle.boxThickness = value
If mTexts.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get Color() As Long
Color = mStyle.Color
End Property

Public Property Let Color(ByVal value As Long)
If mStyle.Color = value Then Exit Property
mStyle.Color = value
If mTexts.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get Count() As Long
Count = mTexts.Count
End Property

Friend Property Get extended() As Boolean
extended = mStyle.extended
End Property

Public Property Let extended(ByVal value As Boolean)
If mStyle.extended = value Then Exit Property
mStyle.extended = value
If mTexts.Count <> 0 Then
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
If mTexts.Count <> 0 Then
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
If mTexts.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get font() As StdFont
Set font = mStyle.font
End Property

Public Property Let font(ByVal value As StdFont)
'If mStyle.font is value Then Exit Property
mStyle.font = value
If mTexts.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Friend Property Get includeInAutoscale() As Boolean
includeInAutoscale = mStyle.includeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mStyle.includeInAutoscale = value Then Exit Property
mStyle.includeInAutoscale = value
If mTexts.Count <> 0 Then
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

Friend Property Let Layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get Layer() As Long
Layer = mLayer
End Property

Public Property Get paddingX() As Double
paddingX = mStyle.paddingX
End Property

Public Property Let paddingX(ByVal value As Double)
If mStyle.paddingX = value Then Exit Property
mStyle.paddingX = value
If mTexts.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get paddingY() As Double
paddingY = mStyle.paddingY
End Property

Public Property Let paddingY(ByVal value As Double)
If mStyle.paddingY = value Then Exit Property
mStyle.paddingY = value
If mTexts.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get style() As TextStyle
Set style = mStyle.clone
End Property

Public Property Let style(ByVal value As TextStyle)
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

If value.align <> mStyle.align Then
    mStyle.align = value.align
    somethingChanged = True
    recalcBoundaries = True
End If
If value.box <> mStyle.box Then
    mStyle.box = value.box
    somethingChanged = True
    recalcBoundaries = True
End If
If value.boxColor <> mStyle.boxColor Then
    mStyle.boxColor = value.boxColor
    somethingChanged = True
End If
If value.boxFillColor <> mStyle.boxFillColor Then
    mStyle.boxFillColor = value.boxFillColor
    somethingChanged = True
End If
If value.boxFillStyle <> mStyle.boxFillStyle Then
    mStyle.boxFillStyle = value.boxFillStyle
    somethingChanged = True
End If
If value.boxStyle <> mStyle.boxStyle Then
    mStyle.boxStyle = value.boxStyle
    somethingChanged = True
    recalcBoundaries = True
End If
If value.boxThickness <> mStyle.boxThickness Then
    mStyle.boxThickness = value.boxThickness
    somethingChanged = True
    recalcBoundaries = True
End If
If value.Color <> mStyle.Color Then
    mStyle.Color = value.Color
    somethingChanged = True
End If
If value.extended <> mStyle.extended Then
    mStyle.extended = value.extended
    somethingChanged = True
End If
If value.fixedX <> mStyle.fixedX Then
    mStyle.fixedX = value.fixedX
    somethingChanged = True
End If
If value.fixedY <> mStyle.fixedY Then
    mStyle.fixedY = value.fixedY
    somethingChanged = True
End If
If Not value.font Is mStyle.font Then
    mStyle.font = value.font
    somethingChanged = True
    recalcBoundaries = True
End If
If value.includeInAutoscale <> mStyle.includeInAutoscale Then
    mStyle.includeInAutoscale = value.includeInAutoscale
    somethingChanged = True
    recalcBoundaries = True
    recalcRegion = True
End If
If value.paddingX <> mStyle.paddingX Then
    mStyle.paddingX = value.paddingX
    somethingChanged = True
    recalcBoundaries = True
End If
If value.paddingY <> mStyle.paddingY Then
    mStyle.paddingY = value.paddingY
    somethingChanged = True
    recalcBoundaries = True
End If

mInhibitChangeEvent = False

If somethingChanged Then
    If recalcBoundaries Then If mTexts.Count <> 0 Then mGOSP.RecalcBoundary mSeriesID
    If mTexts.Count <> 0 Then mGOSP.PaintRegion recalcRegion
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

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As TextStyle, _
                Optional ByVal IsSelectable As Boolean) As text
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.TextSeries::add", _
                            "Text series is finished"
Set Add = New text
If Key = "" Then
    Key = GenerateGUIDString
End If
Add.initialise mSeriesID, Key, mGOSP, mLayer, mCanvas, mStyle, LocalStyle, IsSelectable

mTexts.Add Add, Key
mGOSP.AddGraphicObject Add, mSeriesID
End Function

Friend Sub Finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Friend Sub initialise( _
                ByVal name As String, _
                ByVal style As TextStyle)
mName = name
Set mStyle = style.clone    ' style cannot be shared between TextSeries
End Sub

Public Function Item(ByVal index As Variant) As text
Attribute Item.VB_UserMemId = 0
Set Item = mTexts(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mTexts.[_NewEnum]
End Function

Friend Sub RemoveAll()
StartTask Me, PriorityLow
End Sub

Public Sub RemoveObject(ByVal value As ChartSkil26.text)
On Error GoTo Err
mTexts.Remove value.Key
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


