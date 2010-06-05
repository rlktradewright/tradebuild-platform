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
Private mBaseStyle As LineStyle

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
Private mOffset1 As Size
Private mOffset2 As Size

Private mPropertyOverrideFlags As LinePropertyFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mLines = New Collection
Set mBaseStyle = gDefaultLineStyle
Set mStyle = New LineStyle
End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "LineSeries terminated"
Debug.Print "LineSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Let IGraphicObjectSeries_Canvas(ByVal Value As Canvas)
Set mViewport = Value
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

Private Sub IGraphicObjectSeries_RemoveObject(ByVal Value As IGraphicObject)
Dim lLine As ChartSkil26.Line

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set lLine = Value

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
Case "EXTENDED"
    RepaintRegion True
Case "FIXEDX"
    RepaintRegion True
Case "FIXEDY"
    RepaintRegion True
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
Case "LAYER"
    RepaintRegion False
Case "LINESTYLE"
    RepaintRegion False
Case "OFFSET1"
    RepaintRegion True
Case "OFFSET2"
    RepaintRegion True
Case "THICKNESS"
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

ArrowEndColor = mBaseStyle.ArrowEndColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowEndColor) Then ArrowEndColor = mStyle.ArrowEndColor
If IsPropertySet(LinePropertyArrowEndColor) Then ArrowEndColor = mArrowEndColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndColor(ByVal Value As Long)
Const ProcName As String = "ArrowEndColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mArrowEndColor, LinePropertyArrowEndColor, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillColor() As Long
Const ProcName As String = "ArrowEndFillColor"
Dim failpoint As String
On Error GoTo Err

ArrowEndFillColor = mBaseStyle.ArrowEndFillColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowEndFillColor) Then ArrowEndFillColor = mStyle.ArrowEndFillColor
If IsPropertySet(LinePropertyArrowEndFillColor) Then ArrowEndFillColor = mArrowEndFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndFillColor(ByVal Value As Long)
Const ProcName As String = "ArrowEndFillColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mArrowEndFillColor, LinePropertyArrowEndFillColor, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
Const ProcName As String = "ArrowEndFillStyle"
Dim failpoint As String
On Error GoTo Err

ArrowEndFillStyle = mBaseStyle.ArrowEndFillStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowEndFillStyle) Then ArrowEndFillStyle = mStyle.ArrowEndFillStyle
If IsPropertySet(LinePropertyArrowEndFillStyle) Then ArrowEndFillStyle = mArrowEndFillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "ArrowEndFillStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowEndFillStyle, LinePropertyArrowEndFillStyle, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndLength() As Long
Const ProcName As String = "ArrowEndLength"
Dim failpoint As String
On Error GoTo Err

ArrowEndLength = mBaseStyle.ArrowEndLength
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowEndLength) Then ArrowEndLength = mStyle.ArrowEndLength
If IsPropertySet(LinePropertyArrowEndLength) Then ArrowEndLength = mArrowEndLength

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndLength(ByVal Value As Long)
Const ProcName As String = "ArrowEndLength"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowEndLength, LinePropertyArrowEndLength, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
Const ProcName As String = "ArrowEndStyle"
Dim failpoint As String
On Error GoTo Err

ArrowEndStyle = mBaseStyle.ArrowEndStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowEndStyle) Then ArrowEndStyle = mStyle.ArrowEndStyle
If IsPropertySet(LinePropertyArrowEndStyle) Then ArrowEndStyle = mArrowEndStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndStyle(ByVal Value As ArrowStyles)
Const ProcName As String = "ArrowEndStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowEndStyle, LinePropertyArrowEndStyle, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndWidth() As Long
Const ProcName As String = "ArrowEndWidth"
Dim failpoint As String
On Error GoTo Err

ArrowEndWidth = mBaseStyle.ArrowEndWidth
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowEndWidth) Then ArrowEndWidth = mStyle.ArrowEndWidth
If IsPropertySet(LinePropertyArrowEndWidth) Then ArrowEndWidth = mArrowEndWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndWidth(ByVal Value As Long)
Const ProcName As String = "ArrowEndWidth"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowEndWidth, LinePropertyArrowEndWidth, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartColor() As Long
Const ProcName As String = "ArrowStartColor"
Dim failpoint As String
On Error GoTo Err

ArrowStartColor = mBaseStyle.ArrowStartColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowStartColor) Then ArrowStartColor = mStyle.ArrowStartColor
If IsPropertySet(LinePropertyArrowStartColor) Then ArrowStartColor = mArrowStartColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartColor(ByVal Value As Long)
Const ProcName As String = "ArrowStartColor"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowStartColor, LinePropertyArrowStartColor, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillColor() As Long
Const ProcName As String = "ArrowStartFillColor"
Dim failpoint As String
On Error GoTo Err

ArrowStartFillColor = mBaseStyle.ArrowStartFillColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowStartFillColor) Then ArrowStartFillColor = mStyle.ArrowStartFillColor
If IsPropertySet(LinePropertyArrowStartFillColor) Then ArrowStartFillColor = mArrowStartFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartFillColor(ByVal Value As Long)
Const ProcName As String = "ArrowStartFillColor"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowStartFillColor, LinePropertyArrowStartFillColor, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
Const ProcName As String = "ArrowStartFillStyle"
Dim failpoint As String
On Error GoTo Err

ArrowStartFillStyle = mBaseStyle.ArrowStartFillStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowStartFillStyle) Then ArrowStartFillStyle = mStyle.ArrowStartFillStyle
If IsPropertySet(LinePropertyArrowStartFillStyle) Then ArrowStartFillStyle = mArrowStartfillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "ArrowStartFillStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowStartfillStyle, LinePropertyArrowStartFillStyle, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartLength() As Long
Const ProcName As String = "ArrowStartLength"
Dim failpoint As String
On Error GoTo Err

ArrowStartLength = mBaseStyle.ArrowStartLength
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowStartLength) Then ArrowStartLength = mStyle.ArrowStartLength
If IsPropertySet(LinePropertyArrowStartLength) Then ArrowStartLength = mArrowStartLength

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartLength(ByVal Value As Long)
Const ProcName As String = "ArrowStartLength"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowStartLength, LinePropertyArrowStartLength, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
Const ProcName As String = "ArrowStartStyle"
Dim failpoint As String
On Error GoTo Err

ArrowStartStyle = mBaseStyle.ArrowStartStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowStartStyle) Then ArrowStartStyle = mStyle.ArrowStartStyle
If IsPropertySet(LinePropertyArrowStartStyle) Then ArrowStartStyle = mArrowStartStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartStyle(ByVal Value As ArrowStyles)
Const ProcName As String = "ArrowStartStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowStartStyle, LinePropertyArrowStartStyle, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartWidth() As Long
Const ProcName As String = "ArrowStartWidth"
Dim failpoint As String
On Error GoTo Err

ArrowStartWidth = mBaseStyle.ArrowStartWidth
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyArrowStartWidth) Then ArrowStartWidth = mStyle.ArrowStartWidth
If IsPropertySet(LinePropertyArrowStartWidth) Then ArrowStartWidth = mArrowStartWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartWidth(ByVal Value As Long)
Const ProcName As String = "ArrowStartWidth"
Dim failpoint As String
On Error GoTo Err

setLongProperty mArrowStartWidth, LinePropertyArrowStartWidth, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Canvas(ByVal Value As Canvas)
Set mViewport = Value
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

Color = mBaseStyle.Color
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyColor) Then Color = mStyle.Color
If IsPropertySet(LinePropertyColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

setLongProperty mColor, LinePropertyColor, Value
RepaintRegion False

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

ExtendAfter = mBaseStyle.ExtendAfter
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyExtendAfter) Then ExtendAfter = mStyle.ExtendAfter
If IsPropertySet(LinePropertyExtendAfter) Then ExtendAfter = mExtendAfter

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExtendAfter(ByVal Value As Boolean)
Const ProcName As String = "ExtendAfter"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mExtendAfter, LinePropertyExtendAfter, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExtendBefore() As Boolean
Const ProcName As String = "ExtendBefore"
Dim failpoint As String
On Error GoTo Err

ExtendBefore = mBaseStyle.ExtendBefore
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyExtendBefore) Then ExtendBefore = mStyle.ExtendBefore
If IsPropertySet(LinePropertyExtendBefore) Then ExtendBefore = mExtendBefore

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExtendBefore(ByVal Value As Boolean)
Const ProcName As String = "ExtendBefore"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mExtendBefore, LinePropertyExtendBefore, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

Extended = mBaseStyle.Extended
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyExtended) Then Extended = mStyle.Extended
If IsPropertySet(LinePropertyExtended) Then Extended = mExtended

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Extended(ByVal Value As Boolean)
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mExtended, LinePropertyExtended, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedX() As Boolean
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

FixedX = mBaseStyle.FixedX
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyFixedX) Then FixedX = mStyle.FixedX
If IsPropertySet(LinePropertyFixedX) Then FixedX = mFixedX

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedX(ByVal Value As Boolean)
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mFixedX, LinePropertyFixedX, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedY() As Boolean
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

FixedY = mBaseStyle.FixedY
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyFixedY) Then FixedY = mStyle.FixedY
If IsPropertySet(LinePropertyFixedY) Then FixedY = mFixedY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedY(ByVal Value As Boolean)
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mFixedY, LinePropertyFixedY, Value
RepaintRegion True

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

IncludeInAutoscale = mBaseStyle.IncludeInAutoscale
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyIncludeInAutoscale) Then IncludeInAutoscale = mStyle.IncludeInAutoscale
If IsPropertySet(LinePropertyIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, LinePropertyIncludeInAutoscale, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"
Dim failpoint As String
On Error GoTo Err

mIsSelectable = Value
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

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

mLayer = Value
RepaintRegion False
firePropertyChanged "Layer"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mBaseStyle.Layer
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyLayer) Then Layer = mStyle.Layer
If IsPropertySet(BarPropertyLayer) Then Layer = mLayer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineStyle() As LineStyles
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

LineStyle = mBaseStyle.LineStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyLineStyle) Then LineStyle = mStyle.LineStyle
If IsPropertySet(LinePropertyLineStyle) Then LineStyle = mLineStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mLineStyle, LinePropertyLineStyle, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Offset1(ByVal Value As Size)

Const ProcName As String = "Offset1"
On Error GoTo Err

setSizeProperty mOffset1, LinePropertyOffset1, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Get Offset1() As Size
Const ProcName As String = "Offset1"
On Error GoTo Err

Set Offset1 = mBaseStyle.Offset1
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyOffset1) Then Set Offset1 = mStyle.Offset1
If IsPropertySet(LinePropertyOffset1) Then Set Offset1 = mOffset1

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Offset2(ByVal Value As Size)

Const ProcName As String = "Offset2"
On Error GoTo Err

setSizeProperty mOffset2, LinePropertyOffset2, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Get Offset2() As Size
Const ProcName As String = "Offset2"
On Error GoTo Err

Set Offset2 = mBaseStyle.Offset2
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyOffset2) Then Set Offset2 = mStyle.Offset2
If IsPropertySet(LinePropertyOffset2) Then Set Offset2 = mOffset2

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

Public Property Let Style(ByVal Value As LineStyle)
Const ProcName As String = "Style"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Style is null reference"
End If

Set mStyle = Value
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

Thickness = mBaseStyle.Thickness
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(LinePropertyThickness) Then Thickness = mStyle.Thickness
If IsPropertySet(LinePropertyThickness) Then Thickness = mThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Thickness(ByVal Value As Long)
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

setLongProperty mThickness, LinePropertyThickness, Value
RepaintRegion True

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

Public Sub RemoveObject(ByVal Value As ChartSkil26.Line)
On Error GoTo Err
mLines.Remove Value.Key
mGOSP.RemoveObject Value.handle

Exit Sub

Err:
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub ClearProperty( _
                ByVal flag As LinePropertyFlags)
Const ProcName As String = "ClearProperty"
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

Private Function IsPropertySet( _
                ByVal flag As LinePropertyFlags) As Boolean
Const ProcName As String = "isPropertySet"
Dim failpoint As String
On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

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

Private Sub setBooleanProperty( _
                ByRef pProp As Boolean, _
                ByVal pPropFlag As LinePropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setbooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As LinePropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As LinePropertyFlags)
Const ProcName As String = "setPropertyFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setSizeProperty( _
                ByRef pProp As Size, _
                ByVal pPropFlag As LinePropertyFlags, _
                ByVal Value As Size)
Const ProcName As String = "setSizeProperty"
On Error GoTo Err

If pProp Is Value And IsPropertySet(pPropFlag) Then Exit Sub
Set pProp = Value
If pProp Is Nothing Then
    ClearProperty pPropFlag
Else
    setPropertyFlag pPropFlag
End If
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



