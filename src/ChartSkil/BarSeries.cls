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

Event PropertyChanged(ev As PropertyChangedEventData)

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


Private Const ModuleName                As String = "BarSeries"

'================================================================================
' Member variables
'================================================================================

Private mPeriods As Periods
Private mSeriesID As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mBars As Collection

Private mViewport As ViewPort
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As BarStyle
Attribute mStyle.VB_VarHelpID = -1
Private mBaseStyle As BarStyle

Private mIncludeInAutoscale As Boolean

Private mName As String

Private mPrevBarRef As WeakReference

' these fields override corresponding items in the Style
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayMode As BarDisplayModes
Private mSolidUpBody As Boolean
Private mThickness As Long
Private mWidth As Single
Private mTailThickness As Long
Private mOutlineThickness As Long

Private mPropertyOverrideFlags As BarPropertyFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
Set mBaseStyle = gDefaultBarStyle
Set mStyle = New BarStyle
End Sub

Private Sub Class_Terminate()
gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Class_Terminate", pModName:=ModuleName, pMsg:="BarSeries terminated"
Debug.Print "BarSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Function IGraphicObjectSeries_Add() As IGraphicObject
Set IGraphicObjectSeries_Add = Add
End Function

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
Set mPeriods = Nothing
Set mPrevBarRef = Nothing
Set mGOSP = Nothing
Set mViewport = Nothing
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal pName As String, _
                ByVal pViewport As ViewPort, _
                ByVal pLayer As LayerNumbers, _
                ByVal pPeriods As Periods, _
                ByVal pSeriesID As Long, _
                ByVal pGOSP As GraphObjServiceProvider)
Const ProcName As String = "IGraphicObjectSeries_Initialise"

On Error GoTo Err

mName = pName
mSeriesID = pSeriesID
Set mGOSP = pGOSP
Layer = pLayer
Set mViewport = pViewport
Set mPeriods = pPeriods

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Const ProcName As String = "IGraphicObjectSeries_item"

On Error GoTo Err

Set IGraphicObjectSeries_item = mBars.Item(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Private Sub IGraphicObjectSeries_Remove(ByVal index As Long)
Const ProcName As String = "IGraphicObjectSeries_Remove"

On Error GoTo Err

mBars.Remove index

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal Value As IGraphicObject)
Dim lBar As ChartSkil27.Bar

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"

On Error GoTo Err

On Error Resume Next
Set lBar = Value

mBars.Remove lBar.Key

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IGraphicObjectSeries_Selected(ByVal RHS As Boolean)
Const ProcName As String = "IGraphicObjectSeries_Selected"

On Error GoTo Err

If RHS <> mIsSelected Then
    mIsSelected = RHS
    firePropertyChanged "IsSelected"
    RaiseEvent SelectionStateChanged
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Sub mStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mStyle_PropertyChanged"

On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "COLOR"
    RepaintRegion False
Case "WIDTH"
    RepaintRegion True
Case "DISPLAYMODE"
    RepaintRegion True
Case "DOWNCOLOR"
    RepaintRegion False
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
Case "LAYER"
    RepaintRegion True
Case "OUTLINETHICKNESS"
    RepaintRegion False
Case "SOLIDUPBODY"
    RepaintRegion False
Case "TAILTHICKNESS"
    RepaintRegion False
Case "THICKNESS"
    RepaintRegion False
Case "UPCOLOR"
    RepaintRegion False
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Unhandled property change"
End Select

firePropertyChanged ev.PropertyName

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Color() As Long
Const ProcName As String = "Color"

On Error GoTo Err

Color = mBaseStyle.Color
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyColor) Then Color = mStyle.Color
If IsPropertySet(BarPropertyColor) Then Color = mColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"

On Error GoTo Err


If Not gIsValidColor(Value) And Value <> -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color or -1"

setLongProperty mColor, BarPropertyColor, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"

On Error GoTo Err

Count = mBars.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisplayMode() As BarDisplayModes
Const ProcName As String = "DisplayMode"

On Error GoTo Err

DisplayMode = mBaseStyle.DisplayMode
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyDisplayMode) Then DisplayMode = mStyle.DisplayMode
If IsPropertySet(BarPropertyDisplayMode) Then DisplayMode = mDisplayMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DisplayMode(ByVal Value As BarDisplayModes)
Const ProcName As String = "DisplayMode"

On Error GoTo Err

setLongProperty mDisplayMode, BarPropertyDisplayMode, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownColor() As Long
Const ProcName As String = "DownColor"

On Error GoTo Err

DownColor = mBaseStyle.DownColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyDownColor) Then DownColor = mStyle.DownColor
If IsPropertySet(BarPropertyDownColor) Then DownColor = mDownColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownColor(ByVal Value As Long)
Const ProcName As String = "DownColor"

On Error GoTo Err

setLongProperty mDownColor, BarPropertyDownColor, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

IncludeInAutoscale = mBaseStyle.IncludeInAutoscale
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyIncludeInAutoscale) Then IncludeInAutoscale = mStyle.IncludeInAutoscale
If IsPropertySet(BarPropertyIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, BarPropertyIncludeInAutoscale, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"

On Error GoTo Err

mIsSelectable = Value
firePropertyChanged "IsSelectable"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"

On Error GoTo Err

setLongProperty mLayer, BarPropertyLayer, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mBaseStyle.Layer
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyLayer) Then Layer = mStyle.Layer
If IsPropertySet(BarPropertyLayer) Then Layer = mLayer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"

On Error GoTo Err

mName = Value
firePropertyChanged "Name"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OutlineThickness() As Long
Const ProcName As String = "OutlineThickness"

On Error GoTo Err

OutlineThickness = mBaseStyle.OutlineThickness
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyOutlineThickness) Then OutlineThickness = mStyle.OutlineThickness
If IsPropertySet(BarPropertyOutlineThickness) Then OutlineThickness = mOutlineThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let OutlineThickness(ByVal Value As Long)
Const ProcName As String = "OutlineThickness"

On Error GoTo Err

setLongProperty mOutlineThickness, BarPropertyOutlineThickness, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get SolidUpBody() As Boolean
Const ProcName As String = "SolidUpBody"

On Error GoTo Err

SolidUpBody = mBaseStyle.SolidUpBody
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertySolidUpBody) Then SolidUpBody = mStyle.SolidUpBody
If IsPropertySet(BarPropertySolidUpBody) Then SolidUpBody = mSolidUpBody

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SolidUpBody(ByVal Value As Boolean)
Const ProcName As String = "SolidUpBody"

On Error GoTo Err

setBooleanProperty mSolidUpBody, BarPropertySolidUpBody, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Style() As BarStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal Value As BarStyle)
Const ProcName As String = "Style"

On Error GoTo Err

Set mStyle = Value
RepaintRegion True

firePropertyChanged "Style"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TailThickness() As Long
Const ProcName As String = "TailThickness"

On Error GoTo Err

TailThickness = mBaseStyle.TailThickness
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyTailThickness) Then TailThickness = mStyle.TailThickness
If IsPropertySet(BarPropertyTailThickness) Then TailThickness = mTailThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TailThickness(ByVal Value As Long)
Const ProcName As String = "TailThickness"

On Error GoTo Err

setLongProperty mTailThickness, BarPropertyTailThickness, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Thickness() As Long
Const ProcName As String = "Thickness"

On Error GoTo Err

Thickness = mBaseStyle.Thickness
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyThickness) Then Thickness = mStyle.Thickness
If IsPropertySet(BarPropertyThickness) Then Thickness = mThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Thickness(ByVal Value As Long)
Const ProcName As String = "Thickness"

On Error GoTo Err

setLongProperty mThickness, BarPropertyThickness, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"

On Error GoTo Err

UpColor = mBaseStyle.UpColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyUpColor) Then UpColor = mStyle.UpColor
If IsPropertySet(BarPropertyUpColor) Then UpColor = mUpColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpColor(ByVal Value As Long)
Const ProcName As String = "UpColor"

On Error GoTo Err

setLongProperty mUpColor, BarPropertyUpColor, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Width() As Single
Const ProcName As String = "Width"

On Error GoTo Err

Width = mBaseStyle.Width
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(BarPropertyWidth) Then Width = mStyle.Width
If IsPropertySet(BarPropertyWidth) Then Width = mWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Width(ByVal Value As Single)
Const ProcName As String = "Width"

On Error GoTo Err

setSingleProperty mWidth, BarPropertyWidth, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal Timestamp As Date, _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As BarStyle) As Bar

Const ProcName As String = "Add"
Dim lBar As Bar


On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Bar series is finished"

Set lBar = New Bar

If Key = "" Then
    Key = GenerateGUIDString
End If
lBar.Initialise Me, Timestamp, Key, mGOSP, mViewport, LocalStyle

If CDbl(Timestamp) <> 0# Then lBar.X = mPeriods.GetXFromTimestamp(Timestamp, False)

If Not mPrevBarRef Is Nothing Then lBar.PreviousBar = prevBarObj
mBars.Add lBar, Key
Set mPrevBarRef = CreateWeakReference(lBar)
mGOSP.AddGraphicObject lBar

Set Add = lBar

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"

On Error GoTo Err

mPropertyOverrideFlags = 0
RepaintRegion True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
mFinished = True
Set mViewport = Nothing
End Sub

Public Function Item(ByVal index As Variant) As Bar
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mBars(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveObject(ByVal Value As ChartSkil27.Bar)
On Error GoTo Err
mBars.Remove Value.Key
mGOSP.RemoveObject Value.handle

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

Private Sub ClearProperty( _
                ByVal flag As BarPropertyFlags)
Const ProcName As String = "ClearProperty"

On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEventData

Const ProcName As String = "firePropertyChanged"

On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IsPropertySet( _
                ByVal flag As BarPropertyFlags) As Boolean
Const ProcName As String = "IsPropertySet"

On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function prevBarObj() As Bar
Const ProcName As String = "prevBarObj"

On Error GoTo Err

Set prevBarObj = mPrevBarRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub RepaintRegion( _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "RepaintRegion"

On Error GoTo Err

If mBars.Count <> 0 Then
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBooleanProperty( _
                ByRef pProp As Boolean, _
                ByVal pPropFlag As BarPropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setBooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gBarPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As BarPropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gBarPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSingleProperty( _
                ByRef pProp As Single, _
                ByVal pPropFlag As BarPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setSingleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gBarPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As BarPropertyFlags)
Const ProcName As String = "setPropertyFlag"

On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



