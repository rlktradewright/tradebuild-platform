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

Private mPropertyOverrideFlags As DataPointPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mDataPoints = New Collection
Set mStyle = New DataPointStyle

End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "DataPointSeries terminated"
Debug.Print "DataPointSeries terminated"
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

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim dp As ChartSkil26.DataPoint

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set dp = value

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
Case "HistogramBarWidth"
    RepaintRegion True
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
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

Color = mStyle.Color
If isPropertyOverrideFlagSet(DataPointIsSetColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
If mColor = value And isPropertyOverrideFlagSet(DataPointIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag DataPointIsSetColor
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

Count = mDataPoints.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

DisplayMode = mStyle.DisplayMode
If isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then DisplayMode = mDisplayMode

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DisplayMode(ByVal value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

If mDisplayMode = value And isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then Exit Property
mDisplayMode = value
setPropertyOverrideFlag DataPointIsSetDisplayMode
RepaintRegion True
firePropertyChanged "DisplayMode"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DownColor() As Long
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

DownColor = mStyle.DownColor
If isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then DownColor = mDownColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DownColor(ByVal value As Long)
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"
If mDownColor = value And isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then Exit Property
mDownColor = value
setPropertyOverrideFlag DataPointIsSetDownColor
RepaintRegion False
firePropertyChanged "DownColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HistogramBarWidth() As Single
Const ProcName As String = "HistogramBarWidth"
Dim failpoint As String
On Error GoTo Err

HistogramBarWidth = mStyle.HistogramBarWidth
If isPropertyOverrideFlagSet(DataPointIsSetHistWidth) Then HistogramBarWidth = mHistWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let HistogramBarWidth(ByVal value As Single)
Const ProcName As String = "HistogramBarWidth"
Dim failpoint As String
On Error GoTo Err

If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "HistogramBarWidth must be greater than zero but not greater than 1"
If mHistWidth = value And isPropertyOverrideFlagSet(DataPointIsSetHistWidth) Then Exit Property
mHistWidth = value
setPropertyOverrideFlag DataPointIsSetHistWidth
RepaintRegion True
firePropertyChanged "HistogramBarWidth"

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
If isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag DataPointIsSetIncludeInAutoscale
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
If isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then LineStyle = mLineStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

If mLineStyle = value And isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag DataPointIsSetLineStyle
RepaintRegion False
firePropertyChanged "LineStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineThickness() As Long
Const ProcName As String = "LineThickness"
Dim failpoint As String
On Error GoTo Err

LineThickness = mStyle.LineThickness
If isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then LineThickness = mLineThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineThickness(ByVal value As Long)
Const ProcName As String = "LineThickness"
Dim failpoint As String
On Error GoTo Err

If mLineThickness = value And isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then Exit Property
mLineThickness = value
setPropertyOverrideFlag DataPointIsSetLineThickness
RepaintRegion True
firePropertyChanged "LineThickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal value As String)
Const ProcName As String = "Name"
Dim failpoint As String
On Error GoTo Err

mName = value
firePropertyChanged "Name"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PointStyle() As PointStyles
Const ProcName As String = "PointStyle"
Dim failpoint As String
On Error GoTo Err

PointStyle = mStyle.PointStyle
If isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then PointStyle = mPointStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let PointStyle(ByVal value As PointStyles)
Const ProcName As String = "PointStyle"
Dim failpoint As String
On Error GoTo Err

If mPointStyle = value And isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then Exit Property
mPointStyle = value
setPropertyOverrideFlag DataPointIsSetPointStyle
RepaintRegion False
firePropertyChanged "PointStyle"

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
                ByVal value As DataPointStyle)
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

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

UpColor = mStyle.UpColor
If isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then UpColor = mUpColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let UpColor(ByVal value As Long)
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
If mUpColor = value And isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then Exit Property
mUpColor = value
setPropertyOverrideFlag DataPointIsSetUpColor
RepaintRegion False
firePropertyChanged "UpColor"

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

Public Sub RemoveObject(ByVal value As ChartSkil26.DataPoint)
On Error GoTo Err
mDataPoints.Remove value.Key
mGOSP.RemoveObject value.handle

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

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
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
                ByVal flag As DataPointPropertyOverrideFlags) As Boolean
Const ProcName As String = "isPropertyOverrideFlagSet"
Dim failpoint As String
On Error GoTo Err

isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)

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

Private Sub setPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
Const ProcName As String = "setPropertyOverrideFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



