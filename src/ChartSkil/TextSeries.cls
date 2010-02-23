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
Private mAngle As Double
Private mJustification As TextJustifyModes
Private mMultiLine As Boolean
Private mEllipsis As EllipsisModes
Private mExpandTabs As Boolean
Private mTabWidth As Long
Private mWordWrap As Boolean
Private mLeftMargin As DimensionX
Private mRightMargin As DimensionX
Private mHideIfBlank As Boolean

Private mOffset As Dimension
Private mSize As Dimension

Private mPropertyOverrideFlags As TextPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTexts = New Collection
Set mStyle = New TextStyle

End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "TextSeries terminated"
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
                ByVal pViewport As Viewport, _
                ByVal pLayer As LayerNumbers, _
                ByVal pChart As Chart, _
                ByVal pSeriesID As Long, _
                ByVal pGOSP As GraphObjServiceProvider)
mName = pName
mSeriesID = pSeriesID
Set mGOSP = pGOSP
mLayer = pLayer
Set mViewport = pViewport
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Const ProcName As String = "IGraphicObjectSeries_item"
Dim failpoint As String
On Error GoTo Err

Set IGraphicObjectSeries_item = mTexts.Item(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Private Sub IGraphicObjectSeries_Remove(ByVal index As Long)
Const ProcName As String = "IGraphicObjectSeries_Remove"
Dim failpoint As String
On Error GoTo Err

mTexts.Remove index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lText As ChartSkil26.Text

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set lText = value

mTexts.Remove lText.Key

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Let IGraphicObjectSeries_Selected(ByVal RHS As Boolean)
Const ProcName As String = "IGraphicObjectSeries_Selected"
Dim failpoint As String
On Error GoTo Err

If RHS <> mIsSelected Then
    mIsSelected = RHS
    firePropertyChanged "IsSelected"
    RaiseEvent SelectionStateChanged
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "mStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "ALIGN"
    RepaintRegion True
Case "ANGLE"
    RepaintRegion True
Case "BOX"
    RepaintRegion True
Case "BOXCOLOR"
    RepaintRegion False
Case "BOXFILLCOLOR"
    RepaintRegion False
Case "BOXFILLSTYLE"
    RepaintRegion False
Case "BOXSTYLE"
    RepaintRegion True
Case "BOXTHICKNESS"
    RepaintRegion True
Case "COLOR"
    RepaintRegion False
Case "ELLIPSIS"
    RepaintRegion False
Case "EXPANDTABS"
    RepaintRegion True
Case "EXTENDED"
    RepaintRegion True
Case "FIXEDX"
    RepaintRegion True
Case "FIXEDY"
    RepaintRegion True
Case "FONT"
    RepaintRegion True
Case "HIDEIFBLANK"
    RepaintRegion True
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
Case "JUSTIFICATION"
    RepaintRegion False
Case "LEFTMARGIN"
    RepaintRegion True
Case "MULTILINE"
    RepaintRegion True
Case "PADDINGX"
    RepaintRegion True
Case "PADDINGY"
    RepaintRegion True
Case "RIGHTMARGIN"
    RepaintRegion True
Case "TABWIDTH"
    RepaintRegion True
Case "WORDWRAP"
    RepaintRegion True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unhandled property change"
End Select

firePropertyChanged ev.PropertyName

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Align() As TextAlignModes
Const ProcName As String = "Align"
Dim failpoint As String
On Error GoTo Err

Align = mStyle.Align
If isPropertyOverrideFlagSet(TextIsSetAlign) Then Align = mAlign

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Align(ByVal value As TextAlignModes)
Const ProcName As String = "Align"
Dim failpoint As String
On Error GoTo Err

If mAlign = value And isPropertyOverrideFlagSet(TextIsSetAlign) Then Exit Property
mAlign = value
setPropertyOverrideFlag TextIsSetAlign
RepaintRegion True
firePropertyChanged "Align"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Angle() As Double
Const ProcName As String = "Angle"
Dim failpoint As String
On Error GoTo Err

Angle = mStyle.Angle
If isPropertyOverrideFlagSet(TextIsSetAngle) Then Angle = mAngle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Angle(ByVal value As Double)
Const ProcName As String = "Angle"
Dim failpoint As String
On Error GoTo Err

If mAngle = value And isPropertyOverrideFlagSet(TextIsSetAngle) Then Exit Property
mAngle = value
setPropertyOverrideFlag TextIsSetAngle
RepaintRegion True
firePropertyChanged "Angle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Box() As Boolean
Const ProcName As String = "Box"
Dim failpoint As String
On Error GoTo Err

Box = mStyle.Box
If isPropertyOverrideFlagSet(TextIsSetBox) Then Box = mBox

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Box(ByVal value As Boolean)
Const ProcName As String = "Box"
Dim failpoint As String
On Error GoTo Err

If mBox = value And isPropertyOverrideFlagSet(TextIsSetBox) Then Exit Property
mBox = value
setPropertyOverrideFlag TextIsSetBox
RepaintRegion True
firePropertyChanged "Box"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxColor() As Long
Const ProcName As String = "BoxColor"
Dim failpoint As String
On Error GoTo Err

BoxColor = mStyle.BoxColor
If isPropertyOverrideFlagSet(TextIsSetBoxColor) Then BoxColor = mBoxColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxColor(ByVal value As Long)
Const ProcName As String = "BoxColor"
Dim failpoint As String
On Error GoTo Err

If mBoxColor = value And isPropertyOverrideFlagSet(TextIsSetBoxColor) Then Exit Property
mBoxColor = value
setPropertyOverrideFlag TextIsSetBoxColor
RepaintRegion False
firePropertyChanged "BoxColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillColor() As Long
Const ProcName As String = "BoxFillColor"
Dim failpoint As String
On Error GoTo Err

BoxFillColor = mStyle.BoxFillColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then BoxFillColor = mBoxFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxFillColor(ByVal value As Long)
Const ProcName As String = "BoxFillColor"
Dim failpoint As String
On Error GoTo Err

If mBoxFillColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then Exit Property
mBoxFillColor = value
setPropertyOverrideFlag TextIsSetBoxFillColor
RepaintRegion False
firePropertyChanged "BoxFillColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillStyle() As FillStyles
Const ProcName As String = "BoxFillStyle"
Dim failpoint As String
On Error GoTo Err

BoxFillStyle = mStyle.BoxFillStyle
If isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then BoxFillStyle = mBoxFillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
Const ProcName As String = "BoxFillStyle"
Dim failpoint As String
On Error GoTo Err

If mBoxFillStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then Exit Property
mBoxFillStyle = value
setPropertyOverrideFlag TextIsSetBoxFillStyle
RepaintRegion False
firePropertyChanged "BoxFillStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxStyle() As LineStyles
Const ProcName As String = "BoxStyle"
Dim failpoint As String
On Error GoTo Err

BoxStyle = mStyle.BoxStyle
If isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then BoxStyle = mBoxStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
Const ProcName As String = "BoxStyle"
Dim failpoint As String
On Error GoTo Err

If mBoxStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then Exit Property
mBoxStyle = value
setPropertyOverrideFlag TextIsSetBoxStyle
RepaintRegion True
firePropertyChanged "BoxStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxThickness() As Long
Const ProcName As String = "BoxThickness"
Dim failpoint As String
On Error GoTo Err

BoxThickness = mStyle.BoxThickness
If isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then BoxThickness = mBoxThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxThickness(ByVal value As Long)
Const ProcName As String = "BoxThickness"
Dim failpoint As String
On Error GoTo Err

If mBoxThickness = value And isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then Exit Property
mBoxThickness = value
setPropertyOverrideFlag TextIsSetBoxThickness
RepaintRegion True
firePropertyChanged "BoxThickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
Const ProcName As String = "BoxFillWithBackgroundColor"
Dim failpoint As String
On Error GoTo Err

BoxFillWithBackgroundColor = mStyle.BoxFillWithBackgroundColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
Const ProcName As String = "BoxFillWithBackgroundColor"
Dim failpoint As String
On Error GoTo Err

If mBoxFillWithBackgroundColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then Exit Property
mBoxFillWithBackgroundColor = value
setPropertyOverrideFlag TextIsSetBoxFillWithBackGroundColor
RepaintRegion False
firePropertyChanged "BoxFillWithBackgroundColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

Color = mStyle.Color
If isPropertyOverrideFlagSet(TextIsSetColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If mColor = value And isPropertyOverrideFlagSet(TextIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag TextIsSetColor
RepaintRegion False
firePropertyChanged "Color"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

Count = mTexts.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Ellipsis(ByVal value As EllipsisModes)
Const ProcName As String = "Ellipsis"
Dim failpoint As String
On Error GoTo Err

If mEllipsis = value And isPropertyOverrideFlagSet(TextIsSetEllipsis) Then Exit Property
mEllipsis = value
setPropertyOverrideFlag TextIsSetEllipsis
RepaintRegion False
firePropertyChanged "Ellipsis"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Ellipsis() As EllipsisModes
Const ProcName As String = "Ellipsis"
Dim failpoint As String
On Error GoTo Err

Ellipsis = mStyle.Ellipsis
If isPropertyOverrideFlagSet(TextIsSetEllipsis) Then Ellipsis = mEllipsis

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExpandTabs(ByVal value As Boolean)
Const ProcName As String = "ExpandTabs"
Dim failpoint As String
On Error GoTo Err

If mExpandTabs = value And isPropertyOverrideFlagSet(TextIsSetExpandTabs) Then Exit Property
mExpandTabs = value
setPropertyOverrideFlag TextIsSetExpandTabs
RepaintRegion True
firePropertyChanged "ExpandTabs"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExpandTabs() As Boolean
Const ProcName As String = "ExpandTabs"
Dim failpoint As String
On Error GoTo Err

ExpandTabs = mStyle.ExpandTabs
If isPropertyOverrideFlagSet(TextIsSetExpandTabs) Then ExpandTabs = mExpandTabs

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

Extended = mStyle.Extended
If isPropertyOverrideFlagSet(TextIsSetExtended) Then Extended = mExtended

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Extended(ByVal value As Boolean)
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

If mExtended = value And isPropertyOverrideFlagSet(TextIsSetExtended) Then Exit Property
mExtended = value
setPropertyOverrideFlag TextIsSetExtended
RepaintRegion True
firePropertyChanged "Extended"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedX() As Boolean
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

FixedX = mStyle.FixedX
If isPropertyOverrideFlagSet(TextIsSetFixedX) Then FixedX = mFixedX

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedX(ByVal value As Boolean)
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

If mFixedX = value And isPropertyOverrideFlagSet(TextIsSetFixedX) Then Exit Property
mFixedX = value
setPropertyOverrideFlag TextIsSetFixedX
RepaintRegion True
firePropertyChanged "FixedX"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedY() As Boolean
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

FixedY = mStyle.FixedY
If isPropertyOverrideFlagSet(TextIsSetFixedY) Then FixedY = mFixedY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedY(ByVal value As Boolean)
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

If mFixedY = value And isPropertyOverrideFlagSet(TextIsSetFixedY) Then Exit Property
mFixedY = value
setPropertyOverrideFlag TextIsSetFixedY
RepaintRegion True
firePropertyChanged "FixedY"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Font() As StdFont
Const ProcName As String = "Font"
Dim failpoint As String
On Error GoTo Err

Set Font = mStyle.Font
If isPropertyOverrideFlagSet(TextIsSetFont) Then Set Font = mFont

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Font(ByVal value As StdFont)
Const ProcName As String = "Font"
Dim failpoint As String
On Error GoTo Err

If mFont Is value And isPropertyOverrideFlagSet(TextIsSetFont) Then Exit Property
Set mFont = value
If mFont Is Nothing Then
    clearPropertyOverrideFlag TextIsSetFont
Else
    setPropertyOverrideFlag TextIsSetFont
End If
RepaintRegion True
firePropertyChanged "Font"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get HideIfBlank() As Boolean
Const ProcName As String = "HideIfBlank"
Dim failpoint As String
On Error GoTo Err

HideIfBlank = mStyle.HideIfBlank
If isPropertyOverrideFlagSet(TextIsSetHideIfBlank) Then HideIfBlank = mHideIfBlank

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let HideIfBlank(ByVal value As Boolean)
Const ProcName As String = "HideIfBlank"
Dim failpoint As String
On Error GoTo Err

If mHideIfBlank = value Then Exit Property
mHideIfBlank = value
RepaintRegion True
firePropertyChanged "HideIfBlank"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

IncludeInAutoscale = mStyle.IncludeInAutoscale
If isPropertyOverrideFlagSet(TextIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = value Then Exit Property
mIncludeInAutoscale = value
RepaintRegion True
firePropertyChanged "IncludeInAutoscale"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IsSelectable(ByVal value As Boolean)
Const ProcName As String = "IsSelectable"
Dim failpoint As String
On Error GoTo Err

mIsSelectable = value
firePropertyChanged "IsSelectable"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Let Justification(ByVal value As TextJustifyModes)
Const ProcName As String = "Justification"
Dim failpoint As String
On Error GoTo Err

If mJustification = value And isPropertyOverrideFlagSet(TextIsSetJustification) Then Exit Property
mJustification = value
setPropertyOverrideFlag TextIsSetJustification
RepaintRegion False
firePropertyChanged "Justification"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Justification() As TextJustifyModes
Justification = mJustification
End Property

Public Property Let Layer(ByVal value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

mLayer = value
RepaintRegion False
firePropertyChanged "Layer"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Layer() As LayerNumbers
Layer = mLayer
End Property

Public Property Let LeftMargin(ByVal value As DimensionX)
Const ProcName As String = "LeftMargin"
Dim failpoint As String
On Error GoTo Err

Set mLeftMargin = value
RepaintRegion True
firePropertyChanged "LeftMargin"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LeftMargin() As DimensionX
Set LeftMargin = mLeftMargin
End Property

Public Property Let MultiLine(ByVal value As Boolean)
Const ProcName As String = "MultiLine"
Dim failpoint As String
On Error GoTo Err

If mMultiLine = value And isPropertyOverrideFlagSet(TextIsSetMultiLine) Then Exit Property
mMultiLine = value
setPropertyOverrideFlag TextIsSetMultiLine
RepaintRegion True
firePropertyChanged "MultiLine"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MultiLine() As Boolean
Const ProcName As String = "MultiLine"
Dim failpoint As String

On Error GoTo Err

MultiLine = mStyle.MultiLine
If isPropertyOverrideFlagSet(TextIsSetMultiLine) Then MultiLine = mMultiLine

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Offset() As Dimension
Const ProcName As String = "Offset"
Dim failpoint As String
On Error GoTo Err

Set Offset = mOffset

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Offset(ByVal value As Dimension)
Const ProcName As String = "Offset"
Dim failpoint As String
On Error GoTo Err

Set mOffset = value
RepaintRegion True
firePropertyChanged "Offset"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PaddingX() As Double
Const ProcName As String = "PaddingX"
Dim failpoint As String
On Error GoTo Err

PaddingX = mStyle.PaddingX
If isPropertyOverrideFlagSet(TextIsSetPaddingX) Then PaddingX = mPaddingX

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let PaddingX(ByVal value As Double)
Const ProcName As String = "PaddingX"
Dim failpoint As String
On Error GoTo Err

If mPaddingX = value And isPropertyOverrideFlagSet(TextIsSetPaddingX) Then Exit Property
mPaddingX = value
setPropertyOverrideFlag TextIsSetPaddingX
RepaintRegion True
firePropertyChanged "PaddingX"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PaddingY() As Double
Const ProcName As String = "PaddingY"
Dim failpoint As String
On Error GoTo Err

PaddingY = mStyle.PaddingY
If isPropertyOverrideFlagSet(TextIsSetPaddingY) Then PaddingY = mPaddingY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let PaddingY(ByVal value As Double)
Const ProcName As String = "PaddingY"
Dim failpoint As String
On Error GoTo Err

If mPaddingY = value And isPropertyOverrideFlagSet(TextIsSetPaddingY) Then Exit Property
mPaddingY = value
setPropertyOverrideFlag TextIsSetPaddingY
RepaintRegion True
firePropertyChanged "PaddingY"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let RightMargin(ByVal value As DimensionX)
Const ProcName As String = "RightMargin"
Dim failpoint As String
On Error GoTo Err

Set mRightMargin = value
RepaintRegion True
firePropertyChanged "RightMargin"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get RightMargin() As DimensionX
Set RightMargin = mRightMargin
End Property

Public Property Get Size() As Dimension
Const ProcName As String = "Size"
Dim failpoint As String
On Error GoTo Err

Set Size = mSize

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Size(ByVal value As Dimension)
Const ProcName As String = "Size"
Dim failpoint As String
On Error GoTo Err

Set mSize = value
RepaintRegion True
firePropertyChanged "Size"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get Style() As TextStyle
Set Style = mStyle
End Property

Public Property Let Style(ByVal value As TextStyle)
Const ProcName As String = "Style"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Style is null reference"
End If

Set mStyle = value
RepaintRegion True

firePropertyChanged "Style"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TabWidth(ByVal value As Long)
Const ProcName As String = "TabWidth"
Dim failpoint As String
On Error GoTo Err

If mTabWidth = value And isPropertyOverrideFlagSet(TextIsSetTabWidth) Then Exit Property
mTabWidth = value
setPropertyOverrideFlag TextIsSetTabWidth
RepaintRegion True
firePropertyChanged "TabWidth"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TabWidth() As Long
Const ProcName As String = "TabWidth"
Dim failpoint As String
On Error GoTo Err

TabWidth = mStyle.TabWidth
If isPropertyOverrideFlagSet(TextIsSetTabWidth) Then TabWidth = mTabWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let WordWrap(ByVal value As Boolean)
Const ProcName As String = "WordWrap"
Dim failpoint As String
On Error GoTo Err

If mWordWrap = value And isPropertyOverrideFlagSet(TextIsSetWordWrap) Then Exit Property
mWordWrap = value
setPropertyOverrideFlag TextIsSetWordWrap
RepaintRegion True
firePropertyChanged "WordWrap"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get WordWrap() As Boolean
Const ProcName As String = "WordWrap"
Dim failpoint As String
On Error GoTo Err

WordWrap = mStyle.WordWrap
If isPropertyOverrideFlagSet(TextIsSetWordWrap) Then WordWrap = mWordWrap

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal pKey As String, _
                Optional ByVal pLocalStyle As TextStyle) As Text
Const ProcName As String = "Add"
Dim lText As Text

Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Text series is finished"

If pKey = "" Then pKey = GenerateGUIDString

Set lText = New Text
lText.Initialise Me, pKey, mGOSP, mViewport, pLocalStyle

mTexts.Add lText, pKey
mGOSP.AddGraphicObject lText

Set Add = lText

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = 0
RepaintRegion True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Finish()
mFinished = True
Set mViewport = Nothing
End Sub

Public Function Item(ByVal index As Variant) As Text
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mTexts(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "clearPropertyOverrideFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEvent

Const ProcName As String = "firePropertyChanged"
Dim failpoint As String
On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As TextPropertyOverrideFlags) As Boolean
Const ProcName As String = "isPropertyOverrideFlagSet"
Dim failpoint As String
On Error GoTo Err

isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub RepaintRegion( _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "RepaintRegion"
Dim failpoint As String
On Error GoTo Err

If mTexts.Count <> 0 Then
    If recalcBoundaries Then
        mSeriesUpdateNumber = mSeriesUpdateNumber + 1
        mGOSP.RecalcBoundary mSeriesID
        mGOSP.PaintRegion True
    Else
        mGOSP.PaintRegion False
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As TextPropertyOverrideFlags)
Const ProcName As String = "setPropertyOverrideFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



