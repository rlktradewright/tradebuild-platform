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


Private Const ModuleName                As String = "LineSeries"

'================================================================================
' Member variables
'================================================================================

Private mName As String
Private mSeriesID As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mLines As Collection

'Private mSurface As PictureBox
Private mViewport As Viewport
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As LineStyle
Attribute mStyle.VB_VarHelpID = -1

' these fields override corresponding items in the Style
Private mColor As Long
Private mThickness As Long
Private mLineStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private mArrowStartfillStyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mPropertyOverrideFlags As LinePropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mLines = New Collection
Set mStyle = New LineStyle

End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "LineSeries terminated"
Debug.Print "LineSeries terminated"
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
Const ProcName As String = "IGraphicObjectSeries_count"
Dim failpoint As String
On Error GoTo Err

IGraphicObjectSeries_count = mLines.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Set IGraphicObjectSeries_item = mLines.Item(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Private Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mLines.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lLine As ChartSkil26.Line

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set lLine = value

mLines.Remove lLine.Key

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "mStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "ARROWENDCOLOR"
    RepaintRegion False
Case "ARROWENDFILLCOLOR"
    RepaintRegion False
Case "ARROWENDFILLSTYLE"
    RepaintRegion False
Case "ARROWENDLENGTH"
    RepaintRegion True
Case "ARROWENDSTYLE"
    RepaintRegion True
Case "ARROWENDWIDTH"
    RepaintRegion True
Case "ARROWSTARTCOLOR"
    RepaintRegion False
Case "ARROWSTARTFILLCOLOR"
    RepaintRegion False
Case "ARROWSTARTFILLSTYLE"
    RepaintRegion False
Case "ARROWSTARTLENGTH"
    RepaintRegion True
Case "ARROWSTARTSTYLE"
    RepaintRegion True
Case "ARROWSTARTWIDTH"
    RepaintRegion True
Case "COLOR"
    RepaintRegion False
Case "EXTENDAFTER"
    RepaintRegion True
Case "EXTENDBEFORE"
    RepaintRegion True
Case "LINESTYLE"
    RepaintRegion False
Case "THICKNESS"
    RepaintRegion True
Case "EXTENDED"
    RepaintRegion True
Case "FIXEDX"
    RepaintRegion True
Case "FIXEDY"
    RepaintRegion True
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unhandled property change"
End Select

firePropertyChanged ev.PropertyName

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get ArrowEndColor() As Long
Const ProcName As String = "ArrowEndColor"
Dim failpoint As String
On Error GoTo Err

ArrowEndColor = mStyle.ArrowEndColor
If isPropertyOverrideFlagSet(LineIsSetArrowEndColor) Then ArrowEndColor = mArrowEndColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndColor(ByVal value As Long)
Const ProcName As String = "ArrowEndColor"
Dim failpoint As String
On Error GoTo Err

If mArrowEndColor = value And isPropertyOverrideFlagSet(LineIsSetArrowEndColor) Then Exit Property
mArrowEndColor = value
setPropertyOverrideFlag LineIsSetArrowEndColor
RepaintRegion False
firePropertyChanged "ArrowEndColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillColor() As Long
Const ProcName As String = "ArrowEndFillColor"
Dim failpoint As String
On Error GoTo Err

ArrowEndFillColor = mStyle.ArrowEndFillColor
If isPropertyOverrideFlagSet(LineIsSetArrowEndFillColor) Then ArrowEndFillColor = mArrowEndFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndFillColor(ByVal value As Long)
Const ProcName As String = "ArrowEndFillColor"
Dim failpoint As String
On Error GoTo Err

If mArrowEndFillColor = value And isPropertyOverrideFlagSet(LineIsSetArrowEndFillColor) Then Exit Property
mArrowEndFillColor = value
setPropertyOverrideFlag LineIsSetArrowEndFillColor
RepaintRegion False
firePropertyChanged "ArrowEndFillColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
Const ProcName As String = "ArrowEndFillStyle"
Dim failpoint As String
On Error GoTo Err

ArrowEndFillStyle = mStyle.ArrowEndFillStyle
If isPropertyOverrideFlagSet(LineIsSetArrowEndFillStyle) Then ArrowEndFillStyle = mArrowEndFillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndFillStyle(ByVal value As FillStyles)
Const ProcName As String = "ArrowEndFillStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowEndFillStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowEndFillStyle) Then Exit Property
mArrowEndFillStyle = value
setPropertyOverrideFlag LineIsSetArrowEndFillStyle
RepaintRegion False
firePropertyChanged "ArrowEndFillStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndLength() As Long
Const ProcName As String = "ArrowEndLength"
Dim failpoint As String
On Error GoTo Err

ArrowEndLength = mStyle.ArrowEndLength
If isPropertyOverrideFlagSet(LineIsSetArrowEndLength) Then ArrowEndLength = mArrowEndLength

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndLength(ByVal value As Long)
Const ProcName As String = "ArrowEndLength"
Dim failpoint As String
On Error GoTo Err

If mArrowEndLength = value And isPropertyOverrideFlagSet(LineIsSetArrowEndLength) Then Exit Property
mArrowEndLength = value
setPropertyOverrideFlag LineIsSetArrowEndLength
RepaintRegion True
firePropertyChanged "ArrowEndLength"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
Const ProcName As String = "ArrowEndStyle"
Dim failpoint As String
On Error GoTo Err

ArrowEndStyle = mStyle.ArrowEndStyle
If isPropertyOverrideFlagSet(LineIsSetArrowEndStyle) Then ArrowEndStyle = mArrowEndStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndStyle(ByVal value As ArrowStyles)
Const ProcName As String = "ArrowEndStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowEndStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowEndStyle) Then Exit Property
mArrowEndStyle = value
setPropertyOverrideFlag LineIsSetArrowEndStyle
RepaintRegion True
firePropertyChanged "ArrowEndStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndWidth() As Long
Const ProcName As String = "ArrowEndWidth"
Dim failpoint As String
On Error GoTo Err

ArrowEndWidth = mStyle.ArrowEndWidth
If isPropertyOverrideFlagSet(LineIsSetArrowEndWidth) Then ArrowEndWidth = mArrowEndWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndWidth(ByVal value As Long)
Const ProcName As String = "ArrowEndWidth"
Dim failpoint As String
On Error GoTo Err

If mArrowEndWidth = value And isPropertyOverrideFlagSet(LineIsSetArrowEndWidth) Then Exit Property
mArrowEndWidth = value
setPropertyOverrideFlag LineIsSetArrowEndWidth
RepaintRegion True
firePropertyChanged "ArrowEndWidth"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartColor() As Long
Const ProcName As String = "ArrowStartColor"
Dim failpoint As String
On Error GoTo Err

ArrowStartColor = mStyle.ArrowStartColor
If isPropertyOverrideFlagSet(LineIsSetArrowStartColor) Then ArrowStartColor = mArrowStartColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartColor(ByVal value As Long)
Const ProcName As String = "ArrowStartColor"
Dim failpoint As String
On Error GoTo Err

If mArrowStartColor = value And isPropertyOverrideFlagSet(LineIsSetArrowStartColor) Then Exit Property
mArrowStartColor = value
setPropertyOverrideFlag LineIsSetArrowStartColor
RepaintRegion False
firePropertyChanged "ArrowStartColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillColor() As Long
Const ProcName As String = "ArrowStartFillColor"
Dim failpoint As String
On Error GoTo Err

ArrowStartFillColor = mStyle.ArrowStartFillColor
If isPropertyOverrideFlagSet(LineIsSetArrowStartFillColor) Then ArrowStartFillColor = mArrowStartFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartFillColor(ByVal value As Long)
Const ProcName As String = "ArrowStartFillColor"
Dim failpoint As String
On Error GoTo Err

If mArrowStartFillColor = value And isPropertyOverrideFlagSet(LineIsSetArrowStartFillColor) Then Exit Property
mArrowStartFillColor = value
setPropertyOverrideFlag LineIsSetArrowStartFillColor
RepaintRegion False
firePropertyChanged "ArrowStartFillColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
Const ProcName As String = "ArrowStartFillStyle"
Dim failpoint As String
On Error GoTo Err

ArrowStartFillStyle = mStyle.ArrowStartFillStyle
If isPropertyOverrideFlagSet(LineIsSetArrowStartFillStyle) Then ArrowStartFillStyle = mArrowStartfillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartFillStyle(ByVal value As FillStyles)
Const ProcName As String = "ArrowStartFillStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowStartfillStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowStartFillStyle) Then Exit Property
mArrowStartfillStyle = value
setPropertyOverrideFlag LineIsSetArrowStartFillStyle
RepaintRegion False
firePropertyChanged "ArrowStartFillStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartLength() As Long
Const ProcName As String = "ArrowStartLength"
Dim failpoint As String
On Error GoTo Err

ArrowStartLength = mStyle.ArrowStartLength
If isPropertyOverrideFlagSet(LineIsSetArrowStartLength) Then ArrowStartLength = mArrowStartLength

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartLength(ByVal value As Long)
Const ProcName As String = "ArrowStartLength"
Dim failpoint As String
On Error GoTo Err

If mArrowStartLength = value And isPropertyOverrideFlagSet(LineIsSetArrowStartLength) Then Exit Property
mArrowStartLength = value
setPropertyOverrideFlag LineIsSetArrowStartLength
RepaintRegion True
firePropertyChanged "ArrowStartLength"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
Const ProcName As String = "ArrowStartStyle"
Dim failpoint As String
On Error GoTo Err

ArrowStartStyle = mStyle.ArrowStartStyle
If isPropertyOverrideFlagSet(LineIsSetArrowStartStyle) Then ArrowStartStyle = mArrowStartStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartStyle(ByVal value As ArrowStyles)
Const ProcName As String = "ArrowStartStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowStartStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowStartStyle) Then Exit Property
mArrowStartStyle = value
setPropertyOverrideFlag LineIsSetArrowStartStyle
RepaintRegion True
firePropertyChanged "ArrowStartStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartWidth() As Long
Const ProcName As String = "ArrowStartWidth"
Dim failpoint As String
On Error GoTo Err

ArrowStartWidth = mStyle.ArrowStartWidth
If isPropertyOverrideFlagSet(LineIsSetArrowStartWidth) Then ArrowStartWidth = mArrowStartWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartWidth(ByVal value As Long)
Const ProcName As String = "ArrowStartWidth"
Dim failpoint As String
On Error GoTo Err

If mArrowStartWidth = value And isPropertyOverrideFlagSet(LineIsSetArrowStartWidth) Then Exit Property
mArrowStartWidth = value
setPropertyOverrideFlag LineIsSetArrowStartWidth
RepaintRegion True
firePropertyChanged "ArrowStartWidth"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Canvas(ByVal value As Canvas)
Set mViewport = value
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

Color = mStyle.Color
If isPropertyOverrideFlagSet(LineIsSetColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If mColor = value And isPropertyOverrideFlagSet(LineIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag LineIsSetColor
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

Count = mLines.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExtendAfter() As Boolean
Const ProcName As String = "ExtendAfter"
Dim failpoint As String
On Error GoTo Err

ExtendAfter = mStyle.ExtendAfter
If isPropertyOverrideFlagSet(LineIsSetExtendAfter) Then ExtendAfter = mExtendAfter

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExtendAfter(ByVal value As Boolean)
Const ProcName As String = "ExtendAfter"
Dim failpoint As String
On Error GoTo Err

If mExtendAfter = value And isPropertyOverrideFlagSet(LineIsSetExtendAfter) Then Exit Property
mExtendAfter = value
setPropertyOverrideFlag LineIsSetExtendAfter
RepaintRegion True
firePropertyChanged "ExtendAfter"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExtendBefore() As Boolean
Const ProcName As String = "ExtendBefore"
Dim failpoint As String
On Error GoTo Err

ExtendBefore = mStyle.ExtendBefore
If isPropertyOverrideFlagSet(LineIsSetExtendBefore) Then ExtendBefore = mExtendBefore

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExtendBefore(ByVal value As Boolean)
Const ProcName As String = "ExtendBefore"
Dim failpoint As String
On Error GoTo Err

If mExtendBefore = value And isPropertyOverrideFlagSet(LineIsSetExtendBefore) Then Exit Property
mExtendBefore = value
setPropertyOverrideFlag LineIsSetExtendBefore
RepaintRegion True
firePropertyChanged "ExtendBefore"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

Extended = mStyle.Extended
If isPropertyOverrideFlagSet(LineIsSetExtended) Then Extended = mExtended

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Extended(ByVal value As Boolean)
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

If mExtended = value And isPropertyOverrideFlagSet(LineIsSetExtended) Then Exit Property
mExtended = value
setPropertyOverrideFlag LineIsSetExtended
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
If isPropertyOverrideFlagSet(LineIsSetFixedX) Then FixedX = mFixedX

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedX(ByVal value As Boolean)
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

If mFixedX = value And isPropertyOverrideFlagSet(LineIsSetFixedX) Then Exit Property
mFixedX = value
setPropertyOverrideFlag LineIsSetFixedX
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
If isPropertyOverrideFlagSet(LineIsSetFixedY) Then FixedY = mFixedY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedY(ByVal value As Boolean)
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

If mFixedY = value And isPropertyOverrideFlagSet(LineIsSetFixedY) Then Exit Property
mFixedY = value
setPropertyOverrideFlag LineIsSetFixedY
RepaintRegion True
firePropertyChanged "FixedY"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

IncludeInAutoscale = mStyle.IncludeInAutoscale
If isPropertyOverrideFlagSet(LineIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(LineIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag LineIsSetIncludeInAutoscale
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

Public Property Get LineStyle() As LineStyles
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

LineStyle = mStyle.LineStyle
If isPropertyOverrideFlagSet(LineIsSetLineStyle) Then LineStyle = mLineStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

If mLineStyle = value And isPropertyOverrideFlagSet(LineIsSetLineStyle) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag LineIsSetLineStyle
RepaintRegion False
firePropertyChanged "LineStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get Style() As LineStyle
Set Style = mStyle
End Property

Public Property Let Style(ByVal value As LineStyle)
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

Public Property Get Thickness() As Long
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

Thickness = mStyle.Thickness
If isPropertyOverrideFlagSet(LineIsSetThickness) Then Thickness = mThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Thickness(ByVal value As Long)
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

If mThickness = value And isPropertyOverrideFlagSet(LineIsSetThickness) Then Exit Property
mThickness = value
setPropertyOverrideFlag LineIsSetThickness
RepaintRegion True
firePropertyChanged "Thickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As LineStyle) As ChartSkil26.Line
Const ProcName As String = "Add"

Dim lLine As ChartSkil26.Line

Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Line series is finished"
If Key = "" Then
    Key = GenerateGUIDString
End If
Set lLine = New ChartSkil26.Line
lLine.Initialise Me, Key, mGOSP, mViewport, LocalStyle

mGOSP.AddGraphicObject lLine
mLines.Add lLine, Key

Set Add = lLine
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

Public Function Item(ByVal index As Variant) As ChartSkil26.Line
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mLines(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Const ProcName As String = "NewEnum"
Dim failpoint As String
On Error GoTo Err

   Set NewEnum = mLines.[_NewEnum]

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.Line)
On Error GoTo Err
mLines.Remove value.Key
mGOSP.RemoveObject value.handle

Exit Sub

Err:
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As LinePropertyOverrideFlags)
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
                ByVal flag As LinePropertyOverrideFlags) As Boolean
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

If mLines.Count <> 0 Then
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
                ByVal flag As LinePropertyOverrideFlags)
Const ProcName As String = "setPropertyOverrideFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



