VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointSeries"
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

'================================================================================
' Member variables
'================================================================================


Private Const ModuleName                As String = "DataPointSeries"

Private mChartRef           As WeakReference
Private mSeriesID           As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished           As Boolean

Private mDataPoints         As Collection

Private mViewport           As Viewport
Private mLayer              As Long

Private mGOSP               As GraphObjServiceProvider

Private WithEvents mStyle   As DataPointStyle
Attribute mStyle.VB_VarHelpID = -1
Private mBaseStyle          As DataPointStyle

Private mName               As String

Private mPrevDataPointRef   As WeakReference

' these fields override corresponding items in the Style
Private mLineThickness      As Long
Private mColor              As Long
Private mUpColor            As Long
Private mDownColor          As Long
Private mLineStyle          As LineStyles
Private mPointStyle         As PointStyles
Private mDisplayMode        As DataPointDisplayModes
Private mHistWidth          As Single
Private mIncludeInAutoscale As Boolean

Private mPropertyOverrideFlags As DataPointPropertyFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mDataPoints = New Collection
Set mBaseStyle = gDefaultDataPointStyle
Set mStyle = New DataPointStyle
End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "DataPointSeries terminated"
Debug.Print "DataPointSeries terminated"
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

IGraphicObjectSeries_count = mDataPoints.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
Set mChartRef = Nothing
Set mPrevDataPointRef = Nothing
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
Const ProcName As String = "IGraphicObjectSeries_Initialise"
Dim failpoint As String
On Error GoTo Err

mName = pName
mSeriesID = pSeriesID
Set mGOSP = pGOSP
mLayer = pLayer
Set mViewport = pViewport
Set mChartRef = CreateWeakReference(pChart)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Const ProcName As String = "IGraphicObjectSeries_item"
Dim failpoint As String
On Error GoTo Err

Set IGraphicObjectSeries_item = mDataPoints.Item(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Private Sub IGraphicObjectSeries_Remove(ByVal index As Long)
Dim dp As DataPoint
Const ProcName As String = "IGraphicObjectSeries_Remove"
Dim failpoint As String
On Error GoTo Err

Set dp = mDataPoints.Item(index)
dp.Finish
mDataPoints.Remove index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal Value As IGraphicObject)
Dim dp As ChartSkil26.DataPoint

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set dp = Value

mDataPoints.Remove dp.Key

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
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Const ProcName As String = "mStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "COLOR"
    RepaintRegion False
Case "DISPLAYMODE"
    RepaintRegion True
Case "DOWNCOLOR"
    RepaintRegion False
Case "HISTWIDTH"
    RepaintRegion True
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
Case "LAYER"
    RepaintRegion False
Case "LINESTYLE"
    RepaintRegion False
Case "LINETHICKNESS"
    RepaintRegion True
Case "POINTSTYLE"
    RepaintRegion False
Case "UPCOLOR"
    RepaintRegion False
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

Public Property Get Color() As Long
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

Color = mBaseStyle.Color
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyColor) Then Color = mStyle.Color
If IsPropertySet(DataPointPropertyColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mColor, DataPointPropertyColor, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

Count = mDataPoints.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

DisplayMode = mBaseStyle.DisplayMode
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyDisplayMode) Then DisplayMode = mStyle.DisplayMode
If IsPropertySet(DataPointPropertyDisplayMode) Then DisplayMode = mDisplayMode

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DisplayMode(ByVal Value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

setLongProperty mDisplayMode, DataPointPropertyDisplayMode, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DownColor() As Long
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

DownColor = mBaseStyle.DownColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyDownColor) Then DownColor = mStyle.DownColor
If IsPropertySet(DataPointPropertyDownColor) Then DownColor = mDownColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DownColor(ByVal Value As Long)
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"

setLongProperty mDownColor, DataPointPropertyDownColor, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HistogramBarWidth() As Single
Const ProcName As String = "HistogramBarWidth"
Dim failpoint As String
On Error GoTo Err

HistogramBarWidth = mBaseStyle.HistogramBarWidth
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyHistWidth) Then HistogramBarWidth = mStyle.HistogramBarWidth
If IsPropertySet(DataPointPropertyHistWidth) Then HistogramBarWidth = mHistWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let HistogramBarWidth(ByVal Value As Single)
Const ProcName As String = "HistogramBarWidth"
Dim failpoint As String
On Error GoTo Err

If Value <= 0 Or Value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "HistogramBarWidth must be greater than zero but not greater than 1"

setSingleProperty mHistWidth, DataPointPropertyHistWidth, Value
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
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyIncludeInAutoscale) Then IncludeInAutoscale = mStyle.IncludeInAutoscale
If IsPropertySet(DataPointPropertyIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, DataPointPropertyIncludeInAutoscale, Value
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
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyLineStyle) Then LineStyle = mStyle.LineStyle
If IsPropertySet(DataPointPropertyLineStyle) Then LineStyle = mLineStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mLineStyle, DataPointPropertyLineStyle, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineThickness() As Long
Const ProcName As String = "LineThickness"
Dim failpoint As String
On Error GoTo Err

LineThickness = mBaseStyle.LineThickness
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyLineThickness) Then LineThickness = mStyle.LineThickness
If IsPropertySet(DataPointPropertyLineThickness) Then LineThickness = mLineThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineThickness(ByVal Value As Long)
Const ProcName As String = "LineThickness"
Dim failpoint As String
On Error GoTo Err

setLongProperty mLineThickness, DataPointPropertyLineThickness, Value
RepaintRegion True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"
Dim failpoint As String
On Error GoTo Err

mName = Value
firePropertyChanged "Name"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PointStyle() As PointStyles
Const ProcName As String = "PointStyle"
Dim failpoint As String
On Error GoTo Err

PointStyle = mBaseStyle.PointStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyPointStyle) Then PointStyle = mStyle.PointStyle
If IsPropertySet(DataPointPropertyPointStyle) Then PointStyle = mPointStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let PointStyle(ByVal Value As PointStyles)
Const ProcName As String = "PointStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mPointStyle, DataPointPropertyPointStyle, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get Style() As DataPointStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal Value As DataPointStyle)
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

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

UpColor = mBaseStyle.UpColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(DataPointPropertyUpColor) Then UpColor = mStyle.UpColor
If IsPropertySet(DataPointPropertyUpColor) Then UpColor = mUpColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let UpColor(ByVal Value As Long)
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mUpColor, DataPointPropertyUpColor, Value
RepaintRegion False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                ByVal Timestamp As Date, _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As DataPointStyle) As DataPoint
Static duplicateCount As Long       ' counts the number of successive adds with
                                    ' the same Timestamp
Static prevTimestamp As Date
Dim dp As DataPoint

Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "DataPoint series is finished"

If Not mPrevDataPointRef Is Nothing Then
    If Timestamp = prevTimestamp Then
        duplicateCount = duplicateCount + 1
    Else
        duplicateCount = 0
    End If
End If

Set dp = New DataPoint
If CDbl(Timestamp) <> 0 Then dp.X = chartObj.GetXFromTimestamp(Timestamp, False, duplicateCount)
If Key = "" Then
    Key = GenerateGUIDString
End If
dp.Initialise Me, Timestamp, Key, mGOSP, mViewport, LocalStyle

If Not mPrevDataPointRef Is Nothing Then dp.PreviousDataPoint = prevDataPointObj

mDataPoints.Add dp, Key
Set mPrevDataPointRef = CreateWeakReference(dp)
prevTimestamp = Timestamp
mGOSP.AddGraphicObject dp

Set Add = dp

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

Public Function Item(ByVal index As Variant) As DataPoint
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mDataPoints(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub RemoveObject(ByVal Value As ChartSkil26.DataPoint)
On Error GoTo Err
mDataPoints.Remove Value.Key
mGOSP.RemoveObject Value.handle

Exit Sub

Err:
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mDataPoints.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function chartObj() As Chart
Const ProcName As String = "chartObj"
Dim failpoint As String
On Error GoTo Err

Set chartObj = mChartRef.Target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub ClearProperty( _
                ByVal flag As DataPointPropertyFlags)
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
                ByVal flag As DataPointPropertyFlags) As Boolean
Const ProcName As String = "IsPropertySet"
Dim failpoint As String
On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function prevDataPointObj() As DataPoint
Const ProcName As String = "prevDataPointObj"
Dim failpoint As String
On Error GoTo Err

Set prevDataPointObj = mPrevDataPointRef.Target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub RepaintRegion( _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "RepaintRegion"
Dim failpoint As String
On Error GoTo Err

If mDataPoints.Count <> 0 Then
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
                ByVal pPropFlag As DataPointPropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setBooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gDataPointPropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As DataPointPropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gDataPointPropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setSingleProperty( _
                ByRef pProp As Single, _
                ByVal pPropFlag As DataPointPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setSingleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gDataPointPropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As DataPointPropertyFlags)
Const ProcName As String = "setPropertyFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



