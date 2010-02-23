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


Private Const ModuleName                As String = "BarSeries"

'================================================================================
' Member variables
'================================================================================

Private mChartRef As WeakReference
Private mSeriesID As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mBars As Collection

Private mViewport As Viewport
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As BarStyle
Attribute mStyle.VB_VarHelpID = -1

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

Private mPropertyOverrideFlags As BarPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
Set mStyle = New BarStyle
End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "BarSeries terminated"
Debug.Print "BarSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

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
Set mChartRef = Nothing
Set mPrevBarRef = Nothing
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

Set IGraphicObjectSeries_item = mBars.Item(index)

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

mBars.Remove index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lBar As ChartSkil26.Bar

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set lBar = value

mBars.Remove lBar.Key

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
Case "THICKNESS"
    RepaintRegion False
Case "WIDTH"
    RepaintRegion True
Case "DISPLAYMODE"
    RepaintRegion True
Case "DOWNCOLOR"
    RepaintRegion False
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
Case "OUTLINETHICKNESS"
    RepaintRegion False
Case "SOLIDUPBODY"
    RepaintRegion False
Case "TAILTHICKNESS"
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
If isPropertyOverrideFlagSet(BarIsSetColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If mColor = value And isPropertyOverrideFlagSet(BarIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag BarIsSetColor
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

Count = mBars.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DisplayMode() As BarDisplayModes
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

DisplayMode = mStyle.DisplayMode
If isPropertyOverrideFlagSet(BarIsSetDisplayMode) Then DisplayMode = mDisplayMode

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DisplayMode(ByVal value As BarDisplayModes)
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

If mDisplayMode = value And isPropertyOverrideFlagSet(BarIsSetDisplayMode) Then Exit Property
mDisplayMode = value
setPropertyOverrideFlag BarIsSetDisplayMode
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
If isPropertyOverrideFlagSet(BarIsSetDownColor) Then DownColor = mDownColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DownColor(ByVal value As Long)
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

If mDownColor = value And isPropertyOverrideFlagSet(BarIsSetDownColor) Then Exit Property
mDownColor = value
setPropertyOverrideFlag BarIsSetDownColor
RepaintRegion False
firePropertyChanged "DownColor"

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
If isPropertyOverrideFlagSet(BarIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(BarIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag BarIsSetIncludeInAutoscale
RepaintRegion False
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

If mLayer = value Then Exit Property
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

Public Property Get OutlineThickness() As Long
Const ProcName As String = "OutlineThickness"
Dim failpoint As String
On Error GoTo Err

OutlineThickness = mStyle.OutlineThickness
If isPropertyOverrideFlagSet(BarIsSetOutlineThickness) Then OutlineThickness = mOutlineThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let OutlineThickness(ByVal value As Long)
Const ProcName As String = "OutlineThickness"
Dim failpoint As String
On Error GoTo Err

If mOutlineThickness = value And isPropertyOverrideFlagSet(BarIsSetOutlineThickness) Then Exit Property
mOutlineThickness = value
setPropertyOverrideFlag BarIsSetOutlineThickness
RepaintRegion False
firePropertyChanged "OutlineThickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get SolidUpBody() As Boolean
Const ProcName As String = "SolidUpBody"
Dim failpoint As String
On Error GoTo Err

SolidUpBody = mStyle.SolidUpBody
If isPropertyOverrideFlagSet(BarIsSetSolidUpBody) Then SolidUpBody = mSolidUpBody

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let SolidUpBody(ByVal value As Boolean)
Const ProcName As String = "SolidUpBody"
Dim failpoint As String
On Error GoTo Err

If mSolidUpBody = value And isPropertyOverrideFlagSet(BarIsSetSolidUpBody) Then Exit Property
mSolidUpBody = value
setPropertyOverrideFlag BarIsSetSolidUpBody
RepaintRegion False
firePropertyChanged "SolidUpBody"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Style() As BarStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal value As BarStyle)
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

Public Property Get TailThickness() As Long
Const ProcName As String = "TailThickness"
Dim failpoint As String
On Error GoTo Err

TailThickness = mStyle.TailThickness
If isPropertyOverrideFlagSet(BarIsSetTailThickness) Then TailThickness = mTailThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TailThickness(ByVal value As Long)
Const ProcName As String = "TailThickness"
Dim failpoint As String
On Error GoTo Err

If mTailThickness = value And isPropertyOverrideFlagSet(BarIsSetTailThickness) Then Exit Property
mTailThickness = value
setPropertyOverrideFlag BarIsSetTailThickness
RepaintRegion False
firePropertyChanged "TailThickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Thickness() As Long
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

Thickness = mStyle.Thickness
If isPropertyOverrideFlagSet(BarIsSetThickness) Then Thickness = mThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Thickness(ByVal value As Long)
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

If mThickness = value And isPropertyOverrideFlagSet(BarIsSetThickness) Then Exit Property
mThickness = value
setPropertyOverrideFlag BarIsSetThickness
RepaintRegion False
firePropertyChanged "Thickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

UpColor = mStyle.UpColor
If isPropertyOverrideFlagSet(BarIsSetUpColor) Then UpColor = mUpColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let UpColor(ByVal value As Long)
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

If mUpColor = value And isPropertyOverrideFlagSet(BarIsSetUpColor) Then Exit Property
mUpColor = value
setPropertyOverrideFlag BarIsSetUpColor
RepaintRegion False
firePropertyChanged "UpColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Width() As Single
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

Width = mStyle.Width
If isPropertyOverrideFlagSet(BarIsSetWidth) Then Width = mWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Width(ByVal value As Single)
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

If mWidth = value And isPropertyOverrideFlagSet(BarIsSetWidth) Then Exit Property
mWidth = value
setPropertyOverrideFlag BarIsSetWidth
RepaintRegion True
firePropertyChanged "Width"

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
                Optional ByVal LocalStyle As BarStyle) As Bar

Const ProcName As String = "Add"
Dim lBar As Bar

Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Bar series is finished"

Set lBar = New Bar

If Key = "" Then
    Key = GenerateGUIDString
End If
lBar.Initialise Me, Timestamp, Key, mGOSP, mViewport, LocalStyle

lBar.X = chartObj.GetXFromTimestamp(Timestamp, True)

If Not mPrevBarRef Is Nothing Then lBar.PreviousBar = prevBarObj
mBars.Add lBar, Key
Set mPrevBarRef = CreateWeakReference(lBar)
mGOSP.AddGraphicObject lBar

Set Add = lBar

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

Public Function Item(ByVal index As Variant) As Bar
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mBars(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.Bar)
On Error GoTo Err
mBars.Remove value.Key
mGOSP.RemoveObject value.handle

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
                ByVal flag As BarPropertyOverrideFlags)
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
                ByVal flag As BarPropertyOverrideFlags) As Boolean
Const ProcName As String = "isPropertyOverrideFlagSet"
Dim failpoint As String
On Error GoTo Err

isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function prevBarObj() As Bar
Const ProcName As String = "prevBarObj"
Dim failpoint As String
On Error GoTo Err

Set prevBarObj = mPrevBarRef.Target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub RepaintRegion( _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "RepaintRegion"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As BarPropertyOverrideFlags)
Const ProcName As String = "setPropertyOverrideFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



