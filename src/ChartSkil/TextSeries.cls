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


Private Const ModuleName                As String = "TextSeries"

'================================================================================
' Member variables
'================================================================================

Private mSeriesID As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mTexts As Collection

Private mViewport As Viewport
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As TextStyle
Attribute mStyle.VB_VarHelpID = -1

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mName As String

' these fields override corresponding items in the Style
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillWithBackgroundColor As Boolean
Private mBoxFillStyle As FillStyles
Private mAlign As TextAlignModes
Private mPaddingX As Double
Private mPaddingY As Double
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Private mPropertyOverrideFlags As DataPointPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTexts = New Collection
Set mStyle = New TextStyle

End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "TextSeries terminated"
Debug.Print "TextSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Let IGraphicObjectSeries_Canvas(ByVal value As Canvas)
Set mViewport = value
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
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mTexts.Item(index)
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Public Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mTexts.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lText As ChartSkil26.Text

On Error Resume Next
Set lText = value

mTexts.Remove lText.Key
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
' mFont Event Handlers
'================================================================================

Private Sub mFont_FontChanged(ByVal PropertyName As String)
firePropertyChanged "Font"
End Sub

'================================================================================
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "ALIGN"
    repaintRegion True
Case "BOX"
    repaintRegion True
Case "BOXCOLOR"
    repaintRegion False
Case "BOXFILLCOLOR"
    repaintRegion False
Case "BOXFILLSTYLE"
    repaintRegion False
Case "BOXSTYLE"
    repaintRegion True
Case "BOXTHICKNESS"
    repaintRegion True
Case "COLOR"
    repaintRegion False
Case "FONT"
    repaintRegion True
Case "PADDINGX"
    repaintRegion True
Case "PADDINGY"
    repaintRegion True
Case "EXTENDED"
    repaintRegion True
Case "FIXEDX"
    repaintRegion True
Case "FIXEDY"
    repaintRegion True
Case "INCLUDEINAUTOSCALE"
    repaintRegion True
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

Public Property Get Align() As TextAlignModes
Align = mStyle.Align
If isPropertyOverrideFlagSet(TextIsSetAlign) Then Align = mAlign
End Property

Public Property Let Align(ByVal value As TextAlignModes)
If mAlign = value And isPropertyOverrideFlagSet(TextIsSetAlign) Then Exit Property
mAlign = value
setPropertyOverrideFlag TextIsSetAlign
repaintRegion True
firePropertyChanged "Align"
End Property

Public Property Get Box() As Boolean
Box = mStyle.Box
If isPropertyOverrideFlagSet(TextIsSetBox) Then Box = mBox
End Property

Public Property Let Box(ByVal value As Boolean)
If mBox = value And isPropertyOverrideFlagSet(TextIsSetBox) Then Exit Property
mBox = value
setPropertyOverrideFlag TextIsSetBox
repaintRegion True
firePropertyChanged "Box"
End Property

Public Property Get BoxColor() As Long
BoxColor = mStyle.BoxColor
If isPropertyOverrideFlagSet(TextIsSetBoxColor) Then BoxColor = mBoxColor
End Property

Public Property Let BoxColor(ByVal value As Long)
If mBoxColor = value And isPropertyOverrideFlagSet(TextIsSetBoxColor) Then Exit Property
mBoxColor = value
setPropertyOverrideFlag TextIsSetBoxColor
repaintRegion False
firePropertyChanged "BoxColor"
End Property

Public Property Get BoxFillColor() As Long
BoxFillColor = mStyle.BoxFillColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then BoxFillColor = mBoxFillColor
End Property

Public Property Let BoxFillColor(ByVal value As Long)
If mBoxFillColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then Exit Property
mBoxFillColor = value
setPropertyOverrideFlag TextIsSetBoxFillColor
repaintRegion False
firePropertyChanged "BoxFillColor"
End Property

Public Property Get BoxFillStyle() As FillStyles
BoxFillStyle = mStyle.BoxFillStyle
If isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then BoxFillStyle = mBoxFillStyle
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
If mBoxFillStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then Exit Property
mBoxFillStyle = value
setPropertyOverrideFlag TextIsSetBoxFillStyle
repaintRegion False
firePropertyChanged "BoxFillStyle"
End Property

Public Property Get BoxStyle() As LineStyles
BoxStyle = mStyle.BoxStyle
If isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then BoxStyle = mBoxStyle
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
If mBoxStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then Exit Property
mBoxStyle = value
setPropertyOverrideFlag TextIsSetBoxStyle
repaintRegion True
firePropertyChanged "BoxStyle"
End Property

Public Property Get BoxThickness() As Long
BoxThickness = mStyle.BoxThickness
If isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then BoxThickness = mBoxThickness
End Property

Public Property Let BoxThickness(ByVal value As Long)
If mBoxThickness = value And isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then Exit Property
mBoxThickness = value
setPropertyOverrideFlag TextIsSetBoxThickness
repaintRegion True
firePropertyChanged "BoxThickness"
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
BoxFillWithBackgroundColor = mStyle.BoxFillWithBackgroundColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
If mBoxFillWithBackgroundColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then Exit Property
mBoxFillWithBackgroundColor = value
setPropertyOverrideFlag TextIsSetBoxFillWithBackGroundColor
repaintRegion False
firePropertyChanged "BoxFillWithBackgroundColor"
End Property

Public Property Get Color() As Long
Color = mStyle.Color
If isPropertyOverrideFlagSet(TextIsSetColor) Then Color = mColor
End Property

Public Property Let Color(ByVal value As Long)
If mColor = value And isPropertyOverrideFlagSet(TextIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag TextIsSetColor
repaintRegion False
firePropertyChanged "Color"
End Property

Public Property Get Count() As Long
Count = mTexts.Count
End Property

Public Property Get Extended() As Boolean
Extended = mStyle.Extended
If isPropertyOverrideFlagSet(TextIsSetExtended) Then Extended = mExtended
End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value And isPropertyOverrideFlagSet(TextIsSetExtended) Then Exit Property
mExtended = value
setPropertyOverrideFlag TextIsSetExtended
repaintRegion True
firePropertyChanged "Extended"
End Property

Public Property Get FixedX() As Boolean
FixedX = mStyle.FixedX
If isPropertyOverrideFlagSet(TextIsSetFixedX) Then FixedX = mFixedX
End Property

Public Property Let FixedX(ByVal value As Boolean)
If mFixedX = value And isPropertyOverrideFlagSet(TextIsSetFixedX) Then Exit Property
mFixedX = value
setPropertyOverrideFlag TextIsSetFixedX
repaintRegion True
firePropertyChanged "FixedX"
End Property

Public Property Get FixedY() As Boolean
FixedY = mStyle.FixedY
If isPropertyOverrideFlagSet(TextIsSetFixedY) Then FixedY = mFixedY
End Property

Public Property Let FixedY(ByVal value As Boolean)
If mFixedY = value And isPropertyOverrideFlagSet(TextIsSetFixedY) Then Exit Property
mFixedY = value
setPropertyOverrideFlag TextIsSetFixedY
repaintRegion True
firePropertyChanged "FixedY"
End Property

Public Property Get Font() As StdFont
Set Font = mStyle.Font
If isPropertyOverrideFlagSet(TextIsSetFont) Then Set Font = mFont
End Property

Public Property Let Font(ByVal value As StdFont)
If mFont Is value And isPropertyOverrideFlagSet(TextIsSetFont) Then Exit Property
Set mFont = value
If mFont Is Nothing Then
    clearPropertyOverrideFlag TextIsSetFont
Else
    setPropertyOverrideFlag TextIsSetFont
End If
repaintRegion True
firePropertyChanged "Font"
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
repaintRegion True
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
mLayer = value
repaintRegion False
firePropertyChanged "Layer"
End Property

Public Property Get Layer() As LayerNumbers
Layer = mLayer
End Property

Public Property Get PaddingX() As Double
PaddingX = mStyle.PaddingX
If isPropertyOverrideFlagSet(TextIsSetPaddingX) Then PaddingX = mPaddingX
End Property

Public Property Let PaddingX(ByVal value As Double)
If mPaddingX = value And isPropertyOverrideFlagSet(TextIsSetPaddingX) Then Exit Property
mPaddingX = value
setPropertyOverrideFlag TextIsSetPaddingX
repaintRegion True
firePropertyChanged "PaddingX"
End Property

Public Property Get PaddingY() As Double
PaddingY = mStyle.PaddingY
If isPropertyOverrideFlagSet(TextIsSetPaddingY) Then PaddingY = mPaddingY
End Property

Public Property Let PaddingY(ByVal value As Double)
If mPaddingY = value And isPropertyOverrideFlagSet(TextIsSetPaddingY) Then Exit Property
mPaddingY = value
setPropertyOverrideFlag TextIsSetPaddingY
repaintRegion True
firePropertyChanged "PaddingY"
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get Style() As TextStyle
Set Style = mStyle
End Property

Public Property Let Style(ByVal value As TextStyle)
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

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal pKey As String, _
                Optional ByVal pLocalStyle As TextStyle) As Text
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.TextSeries::Add", _
                            "Text series is finished"

If pKey = "" Then pKey = GenerateGUIDString

Set Add = New Text
Add.Initialise Me, pKey, mGOSP, mViewport, pLocalStyle

mTexts.Add Add, pKey
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

Public Function Item(ByVal index As Variant) As Text
Attribute Item.VB_UserMemId = 0
Set Item = mTexts(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mTexts.[_NewEnum]
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.Text)
On Error GoTo Err
mTexts.Remove value.Key
mGOSP.RemoveObject value.handle

Exit Sub

Err:
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As TextPropertyOverrideFlags)
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
                ByVal flag As TextPropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub repaintRegion( _
                ByVal recalcBoundaries As Boolean)
If mTexts.Count <> 0 Then
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
                ByVal flag As TextPropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub



