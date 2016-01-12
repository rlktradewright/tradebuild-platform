VERSION 5.00
Begin VB.UserControl Chart 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10665
   KeyPreview      =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10665
   Begin VB.PictureBox YRegionDividerPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Index           =   0
      Left            =   4680
      ScaleHeight     =   45
      ScaleWidth      =   3855
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox SelectorPicture 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2640
      Picture         =   "ChartArea.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox BlankPicture 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "ChartArea.ctx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   7455
   End
   Begin VB.PictureBox RegionDividerPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Index           =   0
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9375
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.PictureBox YAxisPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   8400
      ScaleHeight     =   615
      ScaleWidth      =   975
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox XAxisPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   9390
      TabIndex        =   1
      Top             =   6960
      Width           =   9390
   End
   Begin VB.PictureBox ChartRegionPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00602008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   0
      MouseIcon       =   "ChartArea.ctx":0D0C
      MousePointer    =   99  'Custom
      ScaleHeight     =   615
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   8415
   End
End
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IConfigurable

'================================================================================
' Events
'================================================================================

Event ChartCleared()
Event Click()
Event DblCLick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseDown.VB_UserMemId = -605
                
Event MouseMove(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseMove.VB_UserMemId = -606
                
Event MouseUp(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Event PointerModeChanged()
Event PeriodsChanged(ev As CollectionChangeEventData)
Event RegionSelected(ByVal Region As ChartRegion)
Event StyleChanged(ByVal pNewStyle As ChartStyle)

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                                As String = "Chart"

Private Const PropNameHorizontalMouseScrollingAllowed   As String = "HorizontalMouseScrollingAllowed"
Private Const PropNameVerticalMouseScrollingAllowed     As String = "VerticalMouseScrollingAllowed"
Private Const PropNameAutoscrolling                     As String = "Autoscrolling"
Private Const PropNameChartBackColor                    As String = "ChartBackColor"
'Private Const PropNamePeriodLength                      As String = "PeriodLength"
'Private Const PropNamePeriodUnits                       As String = "PeriodUnits"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNamePointerStyle                      As String = "PointerStyle"
Private Const PropNameHorizontalScrollBarVisible        As String = "HorizontalScrollBarVisible"
Private Const PropNameTwipsPerPeriod                    As String = "TwipsPerPeriod"
Private Const PropNameVerticalGridSpacing               As String = "VerticalGridSpacing"
Private Const PropNameVerticalGridUnits                 As String = "VerticalGridUnits"
Private Const PropNameXAxisVisible                      As String = "XAxisVisible"
Private Const PropNameYAxisVisible                      As String = "YAxisVisible"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltHorizontalMouseScrollingAllowed   As Boolean = True
Private Const PropDfltVerticalMouseScrollingAllowed     As Boolean = True
Private Const PropDfltAutoscrolling                     As Boolean = True
Private Const PropDfltChartBackColor                    As Long = &H643232
'Private Const PropDfltPeriodLength                      As Long = 5
'Private Const PropDfltPeriodUnits                       As Long = TimePeriodMinute
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltHorizontalScrollBarVisible        As Boolean = True
Private Const PropDfltPeriodWidth                       As Long = DefaultPeriodWidth
Private Const PropDfltVerticalGridSpacing               As Long = 1
Private Const PropDfltVerticalGridUnits                 As Long = TimePeriodHour
Private Const PropDfltXAxisVisible                      As Boolean = True
Private Const PropDfltYAxisVisible                      As Boolean = True
Private Const PropDfltYAxisWidthCm                      As Single = DefaultYAxisWidthCm

'================================================================================
' Member variables
'================================================================================

Private mStyle                                          As ChartStyle

Private mConfig                                         As ConfigurationSection

Private WithEvents mRegions                             As ChartRegions
Attribute mRegions.VB_VarHelpID = -1
Private mRegionMap                                      As ChartRegionMap

Private mPeriods                                        As Periods
Attribute mPeriods.VB_VarHelpID = -1

Private mController                                     As ChartController

Private mInitialised                                    As Boolean

Private WithEvents mEPhost                              As ExtendedPropertyHost
Attribute mEPhost.VB_VarHelpID = -1

Private mScaleWidth As Single
Private mScaleHeight As Single
Private mScaleLeft As Single
Private mScaleTop As Single

Private mPrevHeight As Single
Private mPrevWidth As Single

Private mXAxisRegion As ChartRegion

Private mYAxisPosition As Long

' indicates whether grids in regions are currently
' hidden. Note that a region's hasGrid property
' indicates whether it has a grid, not whether it
' is currently visible
Private mHideGrid As Boolean

Private mPointerMode As PointerModes
Private mPointerStyle As PointerStyles
Private mPointerIcon As IPictureDisp
Private mToolPointerStyle As PointerStyles
Private mToolIcon As IPictureDisp
Private mPointerDiscColor As Long

Private mSuppressDrawingCount As Long

Private mLeftDragStartPosnX As Long
Private mLeftDragStartPosnY As Single

Private mUserResizingRegions As Boolean

Private mMouseScrollingInProgress As Boolean

Private mReferenceTime As Date

Private mBackGroundViewport As ViewPort

'================================================================================
' User Control Event Handlers
'================================================================================

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblCLick
End Sub

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error Resume Next

gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Proc", pModName:=ModuleName, pMsg:="ChartSkil chart created"

On Error GoTo Err

GChart.gRegisterProperties

Set mRegionMap = New ChartRegionMap

Set gBlankCursor = BlankPicture.Picture
Set gSelectorCursor = SelectorPicture.Picture

ReDim mChartBackGradientFillColors(0) As Long
mChartBackGradientFillColors(0) = PropDfltChartBackColor

Set mController = New ChartController
mController.Chart = Me

Set mEPhost = New ExtendedPropertyHost

Style = gChartStylesManager.DefaultStyle
Initialise

PointerStyle = PointerCrosshairs

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "UserControl_KeyDown"

On Error GoTo Err

RaiseEvent KeyDown(KeyCode, Shift)
mController.fireKeyDown KeyCode, Shift

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
Const ProcName As String = "UserControl_KeyPress"

On Error GoTo Err

RaiseEvent KeyPress(KeyAscii)
mController.fireKeyPress KeyAscii

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "UserControl_KeyUp"

On Error GoTo Err

RaiseEvent KeyUp(KeyCode, Shift)
mController.fireKeyUp KeyCode, Shift

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "UserControl_MouseDown"
On Error GoTo Err

RaiseEvent MouseDown(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "UserControl_MouseMove"
On Error GoTo Err

RaiseEvent MouseMove(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "UserControl_MouseUp"
On Error GoTo Err

RaiseEvent MouseUp(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"

On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="width=" & UserControl.Width & "; height=" & UserControl.Height, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

If UserControl.Width <> 0 And UserControl.Height <> 0 Then Resize (UserControl.Width <> mPrevWidth), (UserControl.Height <> mPrevHeight)
mPrevHeight = UserControl.Height
mPrevWidth = UserControl.Width

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Terminate()
gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Proc", pModName:=ModuleName, pMsg:="ChartSkil chart terminated"
Debug.Print "ChartSkil chart terminated"
End Sub

'@================================================================================
' IConfigurable Interface Members
'@================================================================================

Private Property Let IConfigurable_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "IConfigurable_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IConfigurable_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "IConfigurable_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IConfigurable_RemoveFromConfig()
Const ProcName As String = "IConfigurable_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_Click(index As Integer)
Const ProcName As String = "ChartRegionPicture_Click"

On Error GoTo Err

If index = 0 Then Exit Sub

getDataRegionFromPictureIndex(index).Click

RaiseEvent Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartRegionPicture_DblClick(index As Integer)
Const ProcName As String = "ChartRegionPicture_DblClick"

On Error GoTo Err

If index = 0 Then Exit Sub

getDataRegionFromPictureIndex(index).DblCLick

RaiseEvent DblCLick

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartRegionPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim Region As ChartRegion

Const ProcName As String = "ChartRegionPicture_MouseDown"

On Error GoTo Err

If index = 0 Then Exit Sub

Set Region = getDataRegionFromPictureIndex(index)


If CBool(Button And MouseButtonConstants.vbLeftButton) Then mMouseScrollingInProgress = True

' we notify the region selection first so that the application has a chance to
' turn off scrolling and snapping before getting the MouseDown event
RaiseEvent RegionSelected(Region)
mController.fireRegionSelected Region

If (mPointerMode = PointerModeDefault And _
        ((Region.CursorSnapsToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
        (Not Region.CursorSnapsToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
    (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = Region.YScaleQuantum
    If YScaleQuantum <> 0 Then Y = YScaleQuantum * Int((Y + YScaleQuantum / 10000) / YScaleQuantum)
End If

If mPointerMode = PointerModeDefault And _
    (HorizontalMouseScrollingAllowed Or VerticalMouseScrollingAllowed) _
Then
    mLeftDragStartPosnX = Int(X)
    mLeftDragStartPosnY = Y
End If

Region.MouseDown Button, Shift, Round(X), Y
RaiseEvent MouseDown(Button, _
                    Shift, _
                    convertChartRegionPictureMouseXtoContainerCoords(index, X), _
                    convertChartRegionPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartRegionPicture_MouseMove(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
Const ProcName As String = "ChartRegionPicture_MouseMove"
On Error GoTo Err

If index = 0 Then Exit Sub

Static sPrevX As Single
Static sPrevY As Single
If X = sPrevX And Y = sPrevY Then Exit Sub
sPrevX = X
sPrevY = Y

Dim lregion As ChartRegion
Set lregion = getDataRegionFromPictureIndex(index)

If CBool(Button And MouseButtonConstants.vbLeftButton) Then
    If mPointerMode = PointerModeDefault And _
        (HorizontalMouseScrollingAllowed Or VerticalMouseScrollingAllowed) And _
        mMouseScrollingInProgress _
    Then
        mouseScroll lregion, X, Y
    Else
        mMouseScrollingInProgress = False
        MouseMove lregion, Button, Shift, X, Y
    End If
Else
    MouseMove lregion, Button, Shift, X, Y
End If

lregion.MouseMove Button, Shift, Round(X), Y

RaiseEvent MouseMove(Button, _
                    Shift, _
                    convertChartRegionPictureMouseXtoContainerCoords(index, X), _
                    convertChartRegionPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartRegionPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim Region As ChartRegion

Const ProcName As String = "ChartRegionPicture_MouseUp"

On Error GoTo Err

If index = 0 Then Exit Sub

mMouseScrollingInProgress = False

Set Region = getDataRegionFromPictureIndex(index)

If (mPointerMode = PointerModeDefault And _
        ((Region.CursorSnapsToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
        (Not Region.CursorSnapsToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
    (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = Region.YScaleQuantum
    If YScaleQuantum <> 0 Then Y = YScaleQuantum * Int(Y / YScaleQuantum)
End If
2
Region.MouseUp Button, Shift, Round(X), Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertChartRegionPictureMouseXtoContainerCoords(index, X), _
                    convertChartRegionPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' HScroll Event Handlers
'================================================================================

Private Sub HScroll_Change()
Const ProcName As String = "HScroll_Change"
On Error GoTo Err

LastVisiblePeriod = Round((CLng(HScroll.Value) - CLng(HScroll.Min)) / (CLng(HScroll.Max) - CLng(HScroll.Min)) * (mPeriods.CurrentPeriodNumber + ChartWidth - 1))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' RegionDividerPicture Event Handlers
'================================================================================

Private Sub RegionDividerPicture_Click(index As Integer)
RaiseEvent Click
End Sub

Private Sub RegionDividerPicture_DblClick(index As Integer)
RaiseEvent DblCLick
End Sub

Private Sub RegionDividerPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Const ProcName As String = "RegionDividerPicture_MouseDown"
On Error GoTo Err

If CBool(Button And MouseButtonConstants.vbLeftButton) Then
    mLeftDragStartPosnX = Int(X)
    mLeftDragStartPosnY = Y
    mUserResizingRegions = True
End If
RaiseEvent MouseDown(Button, _
                    Shift, _
                    convertRegionDividerPictureMouseXtoContainerCoords(index, X), _
                    convertRegionDividerPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RegionDividerPicture_MouseMove( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Const ProcName As String = "RegionDividerPicture_MouseMove"
On Error GoTo Err

Static sPrevX As Single
Static sPrevY As Single
If X = sPrevX And Y = sPrevY Then Exit Sub
sPrevX = X
sPrevY = Y

If Not CBool(Button And MouseButtonConstants.vbLeftButton) Then Exit Sub
If Y = mLeftDragStartPosnY Then Exit Sub

If mRegions.ResizeRegion(getDataRegionFromPictureIndex(index), _
                            mLeftDragStartPosnY - Y) _
Then
    setRegionViewSizes
End If
                            
RaiseEvent MouseMove(Button, _
                    Shift, _
                    convertRegionDividerPictureMouseXtoContainerCoords(index, X), _
                    convertRegionDividerPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RegionDividerPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Const ProcName As String = "RegionDividerPicture_MouseUp"
On Error GoTo Err

mUserResizingRegions = False

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertRegionDividerPictureMouseXtoContainerCoords(index, X), _
                    convertRegionDividerPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' XAxisPicture Event Handlers
'================================================================================

Private Sub XAxisPicture_Click()
Const ProcName As String = "XAxisPicture_Click"
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.Click

RaiseEvent Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub XAxisPicture_DblClick()
Const ProcName As String = "XAxisPicture_DblClick"
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.DblCLick

RaiseEvent DblCLick

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub XAxisPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "XAxisPicture_MouseDown"
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.MouseDown Button, Shift, X, Y

RaiseEvent MouseDown(Button, _
                    Shift, _
                    convertXAxisPictureMouseXtoContainerCoords(X), _
                    convertXAxisPictureMouseYtoContainerCoords(Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub XAxisPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "XAxisPicture_MouseMove"
On Error GoTo Err

Static sPrevX As Single
Static sPrevY As Single
If X = sPrevX And Y = sPrevY Then Exit Sub
sPrevX = X
sPrevY = Y

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.MouseMove Button, Shift, X, Y

RaiseEvent MouseMove(Button, _
                    Shift, _
                    convertXAxisPictureMouseXtoContainerCoords(X), _
                    convertXAxisPictureMouseYtoContainerCoords(Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub XAxisPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "XAxisPicture_MouseUp"
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.MouseUp Button, Shift, X, Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertXAxisPictureMouseXtoContainerCoords(X), _
                    convertXAxisPictureMouseYtoContainerCoords(Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' YAxisPicture Event Handlers
'================================================================================

Private Sub YAxisPicture_Click(index As Integer)
Const ProcName As String = "YAxisPicture_Click"
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).Click

RaiseEvent Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub YAxisPicture_DblClick(index As Integer)
Const ProcName As String = "YAxisPicture_DblClick"
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).DblCLick

RaiseEvent DblCLick

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub YAxisPicture_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "YAxisPicture_MouseDown"
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).MouseDown Button, Shift, X, Y

RaiseEvent MouseDown(Button, _
                    Shift, _
                    convertYAxisPictureMouseXtoContainerCoords(index, X), _
                    convertYAxisPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub YAxisPicture_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "YAxisPicture_MouseMove"
On Error GoTo Err

Static sPrevX As Single
Static sPrevY As Single
If X = sPrevX And Y = sPrevY Then Exit Sub
sPrevX = X
sPrevY = Y

getYAxisRegionFromPictureIndex(index).MouseMove Button, Shift, X, Y

RaiseEvent MouseMove(Button, _
                    Shift, _
                    convertYAxisPictureMouseXtoContainerCoords(index, X), _
                    convertYAxisPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub YAxisPicture_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "YAxisPicture_MouseUp"
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).MouseUp Button, Shift, X, Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertYAxisPictureMouseXtoContainerCoords(index, X), _
                    convertYAxisPictureMouseYtoContainerCoords(index, Y))

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mEPhost Event Handlers
'================================================================================

Private Sub mEPhost_Change(pEv As ChangeEventData)
Const ProcName As String = "mEPhost_Change"
On Error GoTo Err

If Not mBackGroundViewport Is Nothing Then mBackGroundViewport.GradientFillColors = gCreateColorArray(ChartBackColor, ChartBackColor)

HScroll.Visible = HorizontalScrollBarVisible

XAxisPicture.Visible = XAxisVisible

Dim lregion As ChartRegion
If Not mRegions Is Nothing Then
    For Each lregion In mRegions
        lregion.CrosshairLineStyle = CrosshairLineStyle
    Next
    
    mRegions.DefaultDataRegionStyle = DefaultRegionStyle
    mRegions.DefaultYAxisRegionStyle = DefaultYAxisRegionStyle
End If

If Not mXAxisRegion Is Nothing Then mXAxisRegion.BaseStyle = XAxisRegionStyle

If Not mRegions Is Nothing Then Resize True, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mEPhost_ExtendedPropertyChanged(pEv As ExtendedPropertyChangedEventData)
Const ProcName As String = "mEPhost_ExtendedPropertyChanged"
On Error GoTo Err

Dim lregion As ChartRegion

If pEv.ExtendedProperty Is gChartBackColorProperty Then
    mBackGroundViewport.GradientFillColors = gCreateColorArray(ChartBackColor, ChartBackColor)
    mBackGroundViewport.PaintBackground
ElseIf pEv.ExtendedProperty Is gHorizontalScrollBarVisibleProperty Then
    HScroll.Visible = HorizontalScrollBarVisible
    Resize False, True
ElseIf pEv.ExtendedProperty Is gPeriodWidthProperty Then
    resizeX
ElseIf pEv.ExtendedProperty Is gXAxisVisibleProperty Then
    mRegions.ResizeY mUserResizingRegions
    XAxisPicture.Visible = XAxisVisible
    setRegionViewSizes
ElseIf pEv.ExtendedProperty Is gYAxisVisibleProperty Then
    resizeX
ElseIf pEv.ExtendedProperty Is gYAxisWidthCmProperty Then
    resizeX
ElseIf pEv.ExtendedProperty Is gCrosshairLineStyleProperty Then
    For Each lregion In mRegions
        lregion.CrosshairLineStyle = CrosshairLineStyle
    Next
ElseIf pEv.ExtendedProperty Is gDefaultRegionStyleProperty Then
    mRegions.DefaultDataRegionStyle = DefaultRegionStyle
ElseIf pEv.ExtendedProperty Is gDefaultYAxisRegionStyleProperty Then
    mRegions.DefaultYAxisRegionStyle = DefaultYAxisRegionStyle
ElseIf pEv.ExtendedProperty Is gXAxisRegionStyleProperty Then
    mXAxisRegion.BaseStyle = XAxisRegionStyle
End If

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' mRegions Event Handlers
'================================================================================

Private Sub mRegions_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mRegions_CollectionChanged"
On Error GoTo Err

Dim rgn As ChartRegion

Select Case ev.changeType
Case CollItemAdded
    Set rgn = ev.affectedItem
    mapRegion rgn
    setRegionViewSizes
Case CollItemRemoved
    Set rgn = ev.affectedItem
    unmapRegion rgn
    setRegionViewSizes
Case CollCollectionCleared

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Autoscrolling() As Boolean
Attribute Autoscrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Autoscrolling.VB_MemberFlags = "400"
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

Autoscrolling = mEPhost.GetValue(GChart.gAutoscrollingProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Autoscrolling(ByVal Value As Boolean)
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

setProperty GChart.gAutoscrollingProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAutoscrolling, Value
PropertyChanged PropNameAutoscrolling

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Availableheight() As Long
Const ProcName As String = "Availableheight"
On Error GoTo Err

Availableheight = IIf(XAxisVisible, XAxisPicture.Top, IIf(HorizontalScrollBarVisible, HScroll.Top, UserControl.ScaleHeight)) - _
                    (mRegions.Count - 1) * RegionDividerPicture(0).Height
If Availableheight < 1 Then Availableheight = 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Attribute ChartBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ChartBackColor.VB_MemberFlags = "400"
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

ChartBackColor = mEPhost.GetValue(GChart.gChartBackColorProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ChartBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

setProperty GChart.gChartBackColorProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingChartBackColor, gLongToHexString(Value)
resizeBackground
PropertyChanged PropNameChartBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartLeft() As Double
Attribute ChartLeft.VB_MemberFlags = "400"
ChartLeft = mScaleLeft
End Property

Public Property Get ChartWidth() As Double
Attribute ChartWidth.VB_MemberFlags = "400"
Const ProcName As String = "ChartWidth"
On Error GoTo Err

ChartWidth = YAxisPosition - mScaleLeft

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Attribute ConfigurationSection.VB_MemberFlags = "400"
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

If Value Is mConfig Then Exit Property
Set mConfig = Value

mConfig.SetSetting ConfigSettingStyle, mStyle.Name

If isLocalValueSet(GChart.gPeriodWidthProperty) Then mConfig.SetSetting ConfigSettingPeriodWidth, mEPhost.getLocalValue(GChart.gPeriodWidthProperty)

If isLocalValueSet(GChart.gAutoscrollingProperty) Then mConfig.SetSetting ConfigSettingAutoscrolling, mEPhost.getLocalValue(GChart.gAutoscrollingProperty)
If isLocalValueSet(GChart.gChartBackColorProperty) Then mConfig.SetSetting ConfigSettingChartBackColor, mEPhost.getLocalValue(GChart.gChartBackColorProperty)
If isLocalValueSet(GChart.gHorizontalMouseScrollingAllowedProperty) Then mConfig.SetSetting ConfigSettingHorizontalMouseScrollingAllowed, mEPhost.getLocalValue(GChart.gHorizontalMouseScrollingAllowedProperty)
If isLocalValueSet(GChart.gHorizontalScrollBarVisibleProperty) Then mConfig.SetSetting ConfigSettingHorizontalScrollBarVisible, mEPhost.getLocalValue(GChart.gHorizontalScrollBarVisibleProperty)
If isLocalValueSet(GChart.gVerticalMouseScrollingAllowedProperty) Then mConfig.SetSetting ConfigSettingVerticalMouseScrollingAllowed, mEPhost.getLocalValue(GChart.gVerticalMouseScrollingAllowedProperty)
If isLocalValueSet(GChart.gXAxisVisibleProperty) Then mConfig.SetSetting ConfigSettingXAxisVisible, mEPhost.getLocalValue(GChart.gXAxisVisibleProperty)
If isLocalValueSet(GChart.gYAxisVisibleProperty) Then mConfig.SetSetting ConfigSettingYAxisVisible, mEPhost.getLocalValue(GChart.gYAxisVisibleProperty)
If isLocalValueSet(GChart.gYAxisWidthCmProperty) Then mConfig.SetSetting ConfigSettingYAxisWidthCm, mEPhost.getLocalValue(GChart.gYAxisWidthCmProperty)

If isLocalValueSet(GChart.gCrosshairLineStyleProperty) Then mEPhost.getLocalValue(GChart.gCrosshairLineStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionCrosshairLineStyle)
If isLocalValueSet(GChart.gDefaultRegionStyleProperty) Then mEPhost.getLocalValue(GChart.gDefaultRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultRegionStyle)
If isLocalValueSet(GChart.gDefaultYAxisRegionStyleProperty) Then mEPhost.getLocalValue(GChart.gDefaultYAxisRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
If isLocalValueSet(GChart.gXAxisRegionStyleProperty) Then mEPhost.getLocalValue(GChart.gXAxisRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXAxisRegionStyle)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Controller() As ChartController
Attribute Controller.VB_MemberFlags = "400"
Set Controller = mController
End Property

Public Property Let CrosshairLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

Dim prevValue As LineStyle
If setProperty(GChart.gCrosshairLineStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionCrosshairLineStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CrosshairLineStyle() As LineStyle
Attribute CrosshairLineStyle.VB_MemberFlags = "400"
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

Set CrosshairLineStyle = mEPhost.GetValue(GChart.gCrosshairLineStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentPeriodNumber() As Long
Attribute CurrentPeriodNumber.VB_MemberFlags = "400"
Const ProcName As String = "CurrentPeriodNumber"
On Error GoTo Err

CurrentPeriodNumber = mPeriods.CurrentPeriodNumber

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentSessionEndTime() As Date
Attribute CurrentSessionEndTime.VB_MemberFlags = "400"
CurrentSessionEndTime = mPeriods.CurrentSessionEndTime
End Property

Public Property Get CurrentSessionStartTime() As Date
Attribute CurrentSessionStartTime.VB_MemberFlags = "400"
CurrentSessionStartTime = mPeriods.CurrentSessionStartTime
End Property

Public Property Let DefaultRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultRegionStyle"
On Error GoTo Err

Dim prevValue As ChartRegionStyle
If setProperty(GChart.gDefaultRegionStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultRegionStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Attribute DefaultRegionStyle.VB_MemberFlags = "400"
Const ProcName As String = "DefaultRegionStyle"
On Error GoTo Err

Set DefaultRegionStyle = mEPhost.GetValue(GChart.gDefaultRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DefaultYAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

Dim prevValue As ChartRegionStyle
If setProperty(GChart.gDefaultYAxisRegionStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Attribute DefaultYAxisRegionStyle.VB_MemberFlags = "400"
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

Set DefaultYAxisRegionStyle = mEPhost.GetValue(GChart.gDefaultYAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirstVisiblePeriod() As Long
Attribute FirstVisiblePeriod.VB_MemberFlags = "400"
FirstVisiblePeriod = mScaleLeft
End Property

Public Property Let FirstVisiblePeriod(ByVal Value As Long)
Const ProcName As String = "FirstVisiblePeriod"
On Error GoTo Err

ScrollX Value - Int(mScaleLeft + 1)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Attribute HorizontalMouseScrollingAllowed.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute HorizontalMouseScrollingAllowed.VB_MemberFlags = "400"
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

HorizontalMouseScrollingAllowed = mEPhost.GetValue(GChart.gHorizontalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

setProperty GChart.gHorizontalMouseScrollingAllowedProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHorizontalMouseScrollingAllowed, Value
PropertyChanged PropNameHorizontalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Attribute HorizontalScrollBarVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute HorizontalScrollBarVisible.VB_MemberFlags = "400"
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

HorizontalScrollBarVisible = mEPhost.GetValue(GChart.gHorizontalScrollBarVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal Value As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

gLogger.Log "HorizontalScrollBarVisible = " & Value, ProcName, ModuleName
setProperty GChart.gHorizontalScrollBarVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHorizontalScrollBarVisible, Value
PropertyChanged PropNameHorizontalScrollBarVisible

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDrawingEnabled() As Boolean
Attribute IsDrawingEnabled.VB_MemberFlags = "400"
IsDrawingEnabled = (mSuppressDrawingCount = 0)
End Property

Public Property Get IsGridHidden() As Boolean
Attribute IsGridHidden.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute IsGridHidden.VB_MemberFlags = "400"
IsGridHidden = mHideGrid
End Property

Public Property Get LastVisiblePeriod() As Long
Attribute LastVisiblePeriod.VB_MemberFlags = "400"
LastVisiblePeriod = mYAxisPosition - 1
End Property

Public Property Let LastVisiblePeriod(ByVal Value As Long)
Const ProcName As String = "LastVisiblePeriod"
On Error GoTo Err

ScrollX Value - mYAxisPosition + 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Periods() As Periods
Attribute Periods.VB_MemberFlags = "400"
Set Periods = mPeriods
End Property

Public Property Let PeriodWidth(ByVal Value As Long)
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

setProperty GChart.gPeriodWidthProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPeriodWidth, Value
PropertyChanged PropNameTwipsPerPeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PeriodWidth() As Long
Attribute PeriodWidth.VB_MemberFlags = "400"
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

PeriodWidth = mEPhost.GetValue(GChart.gPeriodWidthProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
PointerCrosshairsColor = CrosshairLineStyle.Color
End Property

Public Property Let PointerCrosshairsColor(ByVal Value As OLE_COLOR)
Attribute PointerCrosshairsColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
Attribute PointerCrosshairsColor.VB_MemberFlags = "400"
Const ProcName As String = "PointerCrosshairsColor"
On Error GoTo Err

CrosshairLineStyle.Color = Value
PropertyChanged PropNamePointerCrosshairsColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Attribute PointerDiscColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute PointerDiscColor.VB_MemberFlags = "400"
PointerDiscColor = mPointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "PointerDiscColor"
On Error GoTo Err

mPointerDiscColor = Value
Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.PointerDiscColor = Value
Next
PropertyChanged PropNamePointerDiscColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerIcon() As IPictureDisp
Attribute PointerIcon.VB_MemberFlags = "400"
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "PointerIcon"
On Error GoTo Err

If Value Is Nothing Then Exit Property
If Value Is mPointerIcon Then Exit Property

Set mPointerIcon = Value

If mPointerStyle = PointerCustom Then
    Dim lregion As ChartRegion
    For Each lregion In mRegions
        lregion.PointerStyle = PointerCustom
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerMode() As PointerModes
Attribute PointerMode.VB_MemberFlags = "400"
PointerMode = mPointerMode
End Property

Public Property Get PointerStyle() As PointerStyles
Attribute PointerStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute PointerStyle.VB_MemberFlags = "400"
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal Value As PointerStyles)
Const ProcName As String = "PointerStyle"
On Error GoTo Err

If Value = mPointerStyle Then Exit Property

mPointerStyle = Value

If mPointerStyle = PointerCustom And mPointerIcon Is Nothing Then
    ' we'll notify the region when an icon is supplied
    Exit Property
End If

Dim Region As ChartRegion
For Each Region In mRegions
    If mPointerStyle = PointerCustom Then Region.PointerIcon = mPointerIcon
    Region.PointerStyle = mPointerStyle
Next
PropertyChanged PropNamePointerStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Regions() As ChartRegions
Attribute Regions.VB_MemberFlags = "400"
Set Regions = mRegions
End Property

Public Property Get SessionEndTime() As Date
Attribute SessionEndTime.VB_MemberFlags = "400"
SessionEndTime = mPeriods.SessionEndTime
End Property

Public Property Let SessionEndTime(ByVal val As Date)
Const ProcName As String = "SessionEndTime"
On Error GoTo Err

mPeriods.SessionEndTime = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionStartTime() As Date
Attribute SessionStartTime.VB_MemberFlags = "400"
SessionStartTime = mPeriods.SessionStartTime
End Property

Public Property Let SessionStartTime(ByVal Value As Date)
Const ProcName As String = "SessionStartTime"
On Error GoTo Err

mPeriods.SessionStartTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Style(ByVal Value As ChartStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Set mStyle = Value
If mStyle Is Nothing Then Set mStyle = gChartStylesManager.DefaultStyle
gLogger.Log "Using chart style", ProcName, ModuleName, , mStyle.Name
mEPhost.Style = mStyle.ExtendedPropertyHost
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingStyle, mStyle.Name

RaiseEvent StyleChanged(mStyle)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Style() As ChartStyle
Attribute Style.VB_MemberFlags = "400"
Set Style = mStyle
End Property

Public Property Let TimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "TimePeriod"
On Error GoTo Err

mPeriods.TimePeriod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimePeriod() As TimePeriod
Set TimePeriod = mPeriods.TimePeriod
End Property

Friend Property Get TwipsPerPeriod() As Long
Const ProcName As String = "TwipsPerPeriod"
On Error GoTo Err

TwipsPerPeriod = PeriodWidth * Screen.TwipsPerPixelX

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let VerticalGridTimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "VerticalGridTimePeriod"
On Error GoTo Err

mPeriods.VerticalGridTimePeriod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VerticalGridTimePeriod() As TimePeriod
Attribute VerticalGridTimePeriod.VB_MemberFlags = "400"
Set VerticalGridTimePeriod = mPeriods.VerticalGridTimePeriod
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Attribute VerticalMouseScrollingAllowed.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute VerticalMouseScrollingAllowed.VB_MemberFlags = "400"
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

VerticalMouseScrollingAllowed = mEPhost.GetValue(GChart.gVerticalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

setProperty GChart.gVerticalMouseScrollingAllowedProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingVerticalMouseScrollingAllowed, Value
PropertyChanged PropNameVerticalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisRegion() As ChartRegion
Attribute XAxisRegion.VB_MemberFlags = "400"
Set XAxisRegion = mXAxisRegion
End Property

Public Property Let XAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "XAxisRegionStyle"
On Error GoTo Err

Dim prevValue As ChartRegionStyle
If setProperty(GChart.gXAxisRegionStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXAxisRegionStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Attribute XAxisRegionStyle.VB_MemberFlags = "400"
Const ProcName As String = "XAxisRegionStyle"
On Error GoTo Err

Set XAxisRegionStyle = mEPhost.GetValue(GChart.gXAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisVisible() As Boolean
Attribute XAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute XAxisVisible.VB_MemberFlags = "400"
Const ProcName As String = "XAxisVisible"
On Error GoTo Err

XAxisVisible = mEPhost.GetValue(GChart.gXAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "XAxisVisible"
On Error GoTo Err

setProperty GChart.gXAxisVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingXAxisVisible, Value
PropertyChanged PropNameXAxisVisible

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisPosition() As Long
Attribute YAxisPosition.VB_MemberFlags = "400"
YAxisPosition = mYAxisPosition
End Property

Public Property Get YAxisVisible() As Boolean
Attribute YAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute YAxisVisible.VB_MemberFlags = "400"
Const ProcName As String = "YAxisVisible"
On Error GoTo Err

YAxisVisible = mEPhost.GetValue(GChart.gYAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "YAxisVisible"
On Error GoTo Err

setProperty GChart.gYAxisVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYAxisVisible, Value
PropertyChanged PropNameYAxisVisible

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute YAxisWidthCm.VB_MemberFlags = "400"
YAxisWidthCm = mEPhost.GetValue(GChart.gYAxisWidthCmProperty)
End Property

Public Property Let YAxisWidthCm(ByVal Value As Single)
Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

setProperty GChart.gYAxisWidthCmProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYAxisWidthCm, Value
PropertyChanged PropNameYAxisWidthCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub AddPeriod( _
                ByVal pPeriod As Period)
Const ProcName As String = "AddPeriod"
On Error GoTo Err

Dim Region As ChartRegion

For Each Region In mRegions
    Region.AddPeriod pPeriod
Next

mXAxisRegion.AddPeriod pPeriod
If IsDrawingEnabled Then setHorizontalScrollBar
If Autoscrolling Then ScrollX 1
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ClearChart()
Const ProcName As String = "ClearChart"
On Error GoTo Err

gLogger.Log "Clearing chart", ProcName, ModuleName
DisableDrawing

Clear

Initialise
mYAxisPosition = 1

EnableDrawing

RaiseEvent ChartCleared
mController.fireChartCleared
Debug.Print "Chart cleared"

gLogger.Log "Chart cleared", ProcName, ModuleName
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateViewport(ByVal pRegion As ChartRegion, ByVal pRegionType As RegionTypes) As ViewPort
Const ProcName As String = "CreateViewport"
On Error GoTo Err

Dim lCanvas As Canvas

Select Case pRegionType
Case RegionTypeData
    Set lCanvas = createDataRegionCanvas(pRegion.handle)
Case RegionTypeXAxis
    Set lCanvas = createXAxisRegionCanvas
Case RegionTypeYAxis
    Set lCanvas = createYAxisRegionCanvas(pRegion.handle)
Case RegionTypeBackground
    Set lCanvas = createBackgroundRegionCanvas
End Select

Set CreateViewport = New ViewPort
CreateViewport.Initialise lCanvas, pRegion, pRegionType

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub DisableDrawing()
Const ProcName As String = "DisableDrawing"
On Error GoTo Err

SuppressDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub EnableDrawing()
Const ProcName As String = "EnableDrawing"
On Error GoTo Err

SuppressDrawing False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

DisableDrawing
Clear
mController.Finished
mRegions.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetXFromTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal forceNewPeriod As Boolean, _
                Optional ByVal duplicateNumber As Long) As Double
Const ProcName As String = "GetXFromTimestamp"
On Error GoTo Err

GetXFromTimestamp = mPeriods.GetXFromTimestamp(Timestamp, forceNewPeriod, duplicateNumber)


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub HideGrid()
Const ProcName As String = "HideGrid"
On Error GoTo Err

If mHideGrid Then Exit Sub

mHideGrid = True
Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.HideGrid
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
Const ProcName As String = "IsTimeInSession"
On Error GoTo Err

IsTimeInSession = mPeriods.IsTimeInSession(Timestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

If mConfig.GetSetting(ConfigSettingStyle, "") = "" Then
    Style = gChartStylesManager.DefaultStyle
Else
    Style = gChartStylesManager(mConfig.GetSetting(ConfigSettingStyle))
End If

If mConfig.GetSetting(ConfigSettingAutoscrolling) <> "" Then Autoscrolling = mConfig.GetSetting(ConfigSettingAutoscrolling, "True")
If mConfig.GetSetting(ConfigSettingChartBackColor) <> "" Then ChartBackColor = mConfig.GetSetting(ConfigSettingChartBackColor, CStr(vbWhite))
If mConfig.GetSetting(ConfigSettingHorizontalMouseScrollingAllowed) <> "" Then HorizontalMouseScrollingAllowed = mConfig.GetSetting(ConfigSettingHorizontalMouseScrollingAllowed, "true")
If mConfig.GetSetting(ConfigSettingHorizontalScrollBarVisible) <> "" Then HorizontalScrollBarVisible = mConfig.GetSetting(ConfigSettingHorizontalScrollBarVisible, "true")
If mConfig.GetSetting(ConfigSettingPeriodWidth) <> "" Then PeriodWidth = mConfig.GetSetting(ConfigSettingPeriodWidth, DefaultPeriodWidth)
If mConfig.GetSetting(ConfigSettingVerticalMouseScrollingAllowed) <> "" Then VerticalMouseScrollingAllowed = mConfig.GetSetting(ConfigSettingVerticalMouseScrollingAllowed, "true")
If mConfig.GetSetting(ConfigSettingXAxisVisible) <> "" Then XAxisVisible = mConfig.GetSetting(ConfigSettingXAxisVisible, "true")
If mConfig.GetSetting(ConfigSettingYAxisVisible) <> "" Then YAxisVisible = mConfig.GetSetting(ConfigSettingYAxisVisible, "true")
If mConfig.GetSetting(ConfigSettingYAxisWidthCm) <> "" Then YAxisWidthCm = mConfig.GetSetting(ConfigSettingYAxisWidthCm, DefaultYAxisWidthCm)

Dim ls As LineStyle
If Not mConfig.GetConfigurationSection(ConfigSectionCrosshairLineStyle) Is Nothing Then
    Set ls = New LineStyle
    ls.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionCrosshairLineStyle)
    CrosshairLineStyle = ls
End If

Dim crs As ChartRegionStyle
If Not mConfig.GetConfigurationSection(ConfigSectionDefaultRegionStyle) Is Nothing Then
    Set crs = New ChartRegionStyle
    crs.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionDefaultRegionStyle)
    DefaultRegionStyle = crs
End If

If Not mConfig.GetConfigurationSection(ConfigSectionDefaultYAxisRegionStyle) Is Nothing Then
    Set crs = New ChartRegionStyle
    crs.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
    DefaultYAxisRegionStyle = crs
End If

If Not mConfig.GetConfigurationSection(ConfigSectionXAxisRegionStyle) Is Nothing Then
    Set crs = New ChartRegionStyle
    crs.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionXAxisRegionStyle)
    XAxisRegionStyle = crs
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollX(ByVal Value As Long)
Const ProcName As String = "ScrollX"
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="value=" & CStr(Value), pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

If Value = 0 Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="ScrollX", pProjectName:=ProjectName, pModuleName:=ModuleName
    Exit Sub
End If

If (LastVisiblePeriod + Value) > _
        (mPeriods.CurrentPeriodNumber + ChartWidth - 1) Then
    Value = mPeriods.CurrentPeriodNumber + ChartWidth - 1 - LastVisiblePeriod
ElseIf (LastVisiblePeriod + Value) < 1 Then
    Value = 1 - LastVisiblePeriod
End If

mYAxisPosition = mYAxisPosition + Value
mScaleLeft = calcScaleLeft
XAxisPicture.ScaleLeft = mScaleLeft

If Not IsDrawingEnabled Then
    #If trace Then
        gTracer.ExitProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Sub
End If

Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.SetPeriodsInView mScaleLeft, mYAxisPosition - 1
Next

mXAxisRegion.SetPeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
setHorizontalScrollBar

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPointerModeDefault()
Const ProcName As String = "SetPointerModeDefault"
On Error GoTo Err

mPointerMode = PointerModeDefault
Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.SetPointerModeDefault
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPointerModeSelection()
Const ProcName As String = "SetPointerModeSelection"
On Error GoTo Err

mPointerMode = PointerModeSelection

Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.SetPointerModeSelection
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPointerModeTool( _
                Optional ByVal toolPointerStyle As PointerStyles = PointerTool, _
                Optional ByVal icon As IPictureDisp)
Const ProcName As String = "SetPointerModeTool"
On Error GoTo Err

mPointerMode = PointerModeTool
mToolPointerStyle = toolPointerStyle
Set mToolIcon = icon

Select Case toolPointerStyle
Case PointerNone
Case PointerCrosshairs
Case PointerDisc
Case PointerTool
Case PointerCustom
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "toolPointerStyle must be a member of the PointerStyles enum"
End Select

Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.SetPointerModeTool toolPointerStyle, icon
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowGrid()
Const ProcName As String = "ShowGrid"
On Error GoTo Err

If Not mHideGrid Then Exit Sub

mHideGrid = False

Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.ShowGrid
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcScaleLeft() As Single
Const ProcName As String = "calcScaleLeft"
On Error GoTo Err

calcScaleLeft = mYAxisPosition + _
            IIf(YAxisVisible, YAxisWidthCm * TwipsPerCm / XAxisPicture.Width * mScaleWidth, 0) - _
            mScaleWidth

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcScaleWidth() As Single
Const ProcName As String = "calcScaleWidth"
On Error GoTo Err

calcScaleWidth = CSng(XAxisPicture.Width) / CSng(TwipsPerPeriod)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Private Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Dim en As Enumerator
Set en = mRegions.Enumerator

Do While en.MoveNext
    en.Remove
Loop

If Not mXAxisRegion Is Nothing Then mXAxisRegion.ClearRegion
XAxisPicture.Cls
Set mXAxisRegion = Nothing
If Not mPeriods Is Nothing Then mPeriods.Finish
Set mPeriods = Nothing

finishBackgroundCanvas

mRegions.Finish

mSuppressDrawingCount = 0

mInitialised = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function convertChartRegionPictureMouseXtoContainerCoords( _
                ByVal index As Long, _
                ByVal X As Single) As Single
Const ProcName As String = "convertChartRegionPictureMouseXtoContainerCoords"
On Error GoTo Err

convertChartRegionPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(ChartRegionPicture(index), X)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertChartRegionPictureMouseYtoContainerCoords( _
                ByVal index As Long, _
                ByVal Y As Single) As Single
Const ProcName As String = "convertChartRegionPictureMouseYtoContainerCoords"
On Error GoTo Err

convertChartRegionPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(ChartRegionPicture(index), Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertPictureMouseXtoContainerCoords( _
                ByVal pPicture As PictureBox, _
                ByVal X As Single) As Single
Const ProcName As String = "convertPictureMouseXtoContainerCoords"
On Error GoTo Err

convertPictureMouseXtoContainerCoords = _
    ScaleX(pPicture.Left + _
            pPicture.ScaleX(X - pPicture.ScaleLeft, _
                            pPicture.ScaleMode, _
                            vbTwips), _
            vbTwips, _
            vbContainerPosition)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertPictureMouseYtoContainerCoords( _
                ByVal pPicture As PictureBox, _
                ByVal Y As Single) As Single
Const ProcName As String = "convertPictureMouseYtoContainerCoords"
On Error GoTo Err

convertPictureMouseYtoContainerCoords = _
    ScaleY(pPicture.Top + _
            pPicture.ScaleY(Y - pPicture.ScaleTop, _
                            pPicture.ScaleMode, _
                            vbTwips), _
            vbTwips, _
            vbContainerPosition)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertRegionDividerPictureMouseXtoContainerCoords( _
                ByVal index As Long, _
                ByVal X As Single) As Single
convertRegionDividerPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(RegionDividerPicture(index), X)
End Function

Private Function convertRegionDividerPictureMouseYtoContainerCoords( _
                ByVal index As Long, _
                ByVal Y As Single) As Single
Const ProcName As String = "convertRegionDividerPictureMouseYtoContainerCoords"
On Error GoTo Err

convertRegionDividerPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(RegionDividerPicture(index), Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertXAxisPictureMouseXtoContainerCoords( _
                ByVal X As Single) As Single
Const ProcName As String = "convertXAxisPictureMouseXtoContainerCoords"
On Error GoTo Err

convertXAxisPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(XAxisPicture, X)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertXAxisPictureMouseYtoContainerCoords( _
                ByVal Y As Single) As Single
Const ProcName As String = "convertXAxisPictureMouseYtoContainerCoords"
On Error GoTo Err

convertXAxisPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(XAxisPicture, Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertYAxisPictureMouseXtoContainerCoords( _
                ByVal index As Long, _
                ByVal X As Single) As Single
Const ProcName As String = "convertYAxisPictureMouseXtoContainerCoords"
On Error GoTo Err

convertYAxisPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(YAxisPicture(index), X)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertYAxisPictureMouseYtoContainerCoords( _
                ByVal index As Long, _
                ByVal Y As Single) As Single
Const ProcName As String = "convertYAxisPictureMouseYtoContainerCoords"
On Error GoTo Err

convertYAxisPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(YAxisPicture(index), Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createBackgroundRegionCanvas() As Canvas
Const ProcName As String = "createBackgroundRegionCanvas"
On Error GoTo Err

Set createBackgroundRegionCanvas = createCanvas(ChartRegionPicture(0), RegionTypeBackground)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createCanvas( _
                ByVal Surface As PictureBox, _
                ByVal pRegionType As RegionTypes) As Canvas
Const ProcName As String = "createCanvas"
On Error GoTo Err

Set createCanvas = New Canvas
createCanvas.Surface = Surface
createCanvas.RegionType = pRegionType

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createDataRegionCanvas(ByVal pIndex As Long) As Canvas
Const ProcName As String = "createDataRegionCanvas"
On Error GoTo Err

Load ChartRegionPicture(pIndex)
Set createDataRegionCanvas = createCanvas(ChartRegionPicture(pIndex), RegionTypeData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createXAxisRegionCanvas() As Canvas
Const ProcName As String = "createXAxisRegionCanvas"
On Error GoTo Err

Set createXAxisRegionCanvas = createCanvas(XAxisPicture, RegionTypeXAxis)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createYAxisRegionCanvas(ByVal pIndex As Long) As Canvas
Const ProcName As String = "createYAxisRegionCanvas"
On Error GoTo Err

Load YAxisPicture(pIndex)
Set createYAxisRegionCanvas = createCanvas(YAxisPicture(pIndex), RegionTypeYAxis)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub createXAxisRegion()
Const ProcName As String = "createXAxisRegion"
On Error GoTo Err

Dim afont As StdFont

Set mXAxisRegion = New ChartRegion

mXAxisRegion.Initialise "", mPeriods, CreateViewport(mXAxisRegion, RegionTypeXAxis), RegionTypeXAxis
mXAxisRegion.BaseStyle = XAxisRegionStyle

mXAxisRegion.Bottom = 0
mXAxisRegion.Top = 1

mXAxisRegion.IsDrawingEnabled = IsDrawingEnabled

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub finishBackgroundCanvas()
Const ProcName As String = "finishBackgroundCanvas"
On Error GoTo Err

gLogger.Log "Finish background canvas", ProcName, ModuleName, LogLevelHighDetail
If Not mBackGroundViewport Is Nothing Then mBackGroundViewport.Finish
Set mBackGroundViewport = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getDataRegionFromPictureIndex( _
                ByVal index As Long) As ChartRegion
Const ProcName As String = "getDataRegionFromPictureIndex"
On Error GoTo Err

Set getDataRegionFromPictureIndex = mRegionMap.Item(CLng(ChartRegionPicture(index).Tag))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getYAxisRegionFromPictureIndex( _
                ByVal index As Long) As ChartRegion
Const ProcName As String = "getYAxisRegionFromPictureIndex"
On Error GoTo Err

Set getYAxisRegionFromPictureIndex = mRegions.ItemFromHandle(CLng(YAxisPicture(index).Tag))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub Initialise()
Const ProcName As String = "Initialise"
On Error GoTo Err

gLogger.Log "Initialising chart", ProcName, ModuleName

Set mPeriods = New Periods
mPeriods.Chart = Me

Set mRegions = New ChartRegions
mRegions.Initialise Me, mPeriods

mRegions.DefaultDataRegionStyle = DefaultRegionStyle
mRegions.DefaultYAxisRegionStyle = DefaultYAxisRegionStyle

createXAxisRegion

mPrevHeight = UserControl.Height

Set mBackGroundViewport = CreateViewport(Nothing, RegionTypeBackground)

mPointerMode = PointerModes.PointerModeDefault

mYAxisPosition = 1
mScaleWidth = calcScaleWidth
mScaleLeft = calcScaleLeft
mScaleHeight = -100
mScaleTop = 100

HScroll.Value = 0

Resize True, True

mInitialised = True
gLogger.Log "Chart initialised", ProcName, ModuleName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function isLocalValueSet(ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "isLocalValueSet"
On Error GoTo Err

isLocalValueSet = mEPhost.IsPropertySet(pExtProp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub mapRegion(pRegion As ChartRegion)
Const ProcName As String = "mapRegion"
On Error GoTo Err

If pRegion Is Nothing Then Exit Sub

Dim index As Long
index = pRegion.handle

Dim mapHandle As Long
mapHandle = mRegionMap.Append(pRegion)

ChartRegionPicture(index).Tag = mapHandle
ChartRegionPicture(index).Visible = True
ChartRegionPicture(index).Align = vbAlignNone
ChartRegionPicture(index).Width = _
    IIf(YAxisVisible, UserControl.ScaleWidth - YAxisWidthCm * TwipsPerCm, UserControl.ScaleWidth)
ChartRegionPicture(index).ZOrder 1

pRegion.IsDrawingEnabled = IsDrawingEnabled
pRegion.PointerStyle = mPointerStyle
pRegion.PointerIcon = mPointerIcon
pRegion.CrosshairLineStyle = CrosshairLineStyle
pRegion.PointerDiscColor = mPointerDiscColor
Select Case mPointerMode
Case PointerModeDefault
    pRegion.SetPointerModeDefault
Case PointerModeTool
    pRegion.SetPointerModeTool mToolPointerStyle, mToolIcon
Case PointerModeSelection
    pRegion.SetPointerModeSelection
End Select
pRegion.Left = mScaleLeft
pRegion.Bottom = 0
pRegion.Top = 1
pRegion.SetPeriodsInView mScaleLeft, mYAxisPosition - 1

If mHideGrid Then pRegion.HideGrid

Load RegionDividerPicture(index)
RegionDividerPicture(index).Tag = mapHandle
RegionDividerPicture(index).ZOrder 0
RegionDividerPicture(index).Visible = (Not mRegionMap.IsFirst(mapHandle))
pRegion.Divider = RegionDividerPicture(index)

Dim yIndex As Long
yIndex = pRegion.YAxisRegion.handle

Load YRegionDividerPicture(yIndex)
YRegionDividerPicture(yIndex).Tag = mapHandle
YRegionDividerPicture(yIndex).ZOrder 0
YRegionDividerPicture(yIndex).Visible = (Not mRegionMap.IsFirst(mapHandle))
pRegion.YAxisRegion.Divider = YRegionDividerPicture(yIndex)

YAxisPicture(yIndex).Tag = yIndex
YAxisPicture(yIndex).Visible = True
YAxisPicture(yIndex).Align = vbAlignNone
YAxisPicture(yIndex).Left = ChartRegionPicture(index).Width
YAxisPicture(yIndex).Width = YAxisWidthCm * TwipsPerCm

XAxisPicture.Visible = XAxisVisible

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub MouseMove( _
                ByVal targetRegion As ChartRegion, _
                ByVal Button As Long, _
                ByVal Shift As Long, _
                ByRef X As Single, _
                ByRef Y As Single)
Const ProcName As String = "MouseMove"
On Error GoTo Err

Dim lregion As ChartRegion
For Each lregion In mRegions
    If lregion Is targetRegion Then
        'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & X & " y=" & Y
        If (mPointerMode = PointerModeDefault And _
                ((lregion.CursorSnapsToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
                (Not lregion.CursorSnapsToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
            (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
        Then
            Dim YScaleQuantum As Double
            YScaleQuantum = lregion.YScaleQuantum
            If YScaleQuantum <> 0 Then Y = YScaleQuantum * Int((Y + YScaleQuantum / 10000) / YScaleQuantum)
        End If
        lregion.DrawCursor Button, Shift, X, Y, True
        If Not lregion.YAxisRegion Is Nothing Then lregion.YAxisRegion.DrawCursor Button, Shift, 0!, Y, True
        
    Else
        'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & X & " y=" & MinusInfinitySingle
        lregion.DrawCursor Button, Shift, X, 0!, False
        If Not lregion.YAxisRegion Is Nothing Then lregion.YAxisRegion.DrawCursor Button, Shift, 0!, 0!, False
    End If
Next
XAxisRegion.DrawCursor Button, Shift, X, 0!, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mouseScroll( _
                ByVal targetRegion As ChartRegion, _
                ByRef X As Single, _
                ByRef Y As Single)
Const ProcName As String = "mouseScroll"
On Error GoTo Err

If HorizontalMouseScrollingAllowed Then
    ' the chart needs to be scrolled so that current mouse Position
    ' is the value contained in mLeftDragStartPosnX
    If mLeftDragStartPosnX <> Int(X) Then
        If (LastVisiblePeriod + mLeftDragStartPosnX - Int(X)) <= _
                (mPeriods.CurrentPeriodNumber + ChartWidth - 1) And _
            (LastVisiblePeriod + mLeftDragStartPosnX - Int(X)) >= 1 _
        Then
            ScrollX mLeftDragStartPosnX - Int(X)
        End If
    End If
End If
If VerticalMouseScrollingAllowed Then
    If mLeftDragStartPosnY <> Y Then
        With targetRegion
            If Not .Autoscaling Then
                .ScrollVertical mLeftDragStartPosnY - Y
            End If
        End With
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Resize( _
    ByVal resizeWidth As Boolean, _
    ByVal resizeHeight As Boolean)
Const ProcName As String = "Resize"
On Error GoTo Err

If Not mInitialised Then Exit Sub

resizeBackground

If resizeWidth Then
    HScroll.Width = UserControl.Width
    XAxisPicture.Width = UserControl.Width
    resizeX
End If

If resizeHeight Then
    HScroll.Top = UserControl.Height - IIf(HorizontalScrollBarVisible, HScroll.Height, 0)
    XAxisPicture.Top = HScroll.Top - IIf(XAxisVisible, XAxisPicture.Height, 0)
    mRegions.ResizeY mUserResizingRegions
    setRegionViewSizes
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resizeBackground()
Const ProcName As String = "resizeBackground"
On Error GoTo Err

If mRegions Is Nothing Then Exit Sub
If mRegions.Count > 0 Then Exit Sub
XAxisPicture.Visible = False
ChartRegionPicture(0).Visible = False
ChartRegionPicture(0).Move 0, 0, UserControl.Width, UserControl.Height
mBackGroundViewport.GradientFillColors = gCreateColorArray(ChartBackColor, ChartBackColor)
mBackGroundViewport.Left = 0
mBackGroundViewport.Right = 1
mBackGroundViewport.SetVerticalBounds 0#, 1#
mBackGroundViewport.PaintBackground
mBackGroundViewport.Canvas.ZOrder 1
ChartRegionPicture(0).Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resizeX()
Const ProcName As String = "resizeX"
On Error GoTo Err

If Not mInitialised Then Exit Sub

mScaleWidth = calcScaleWidth
mScaleLeft = calcScaleLeft

Dim lregion As ChartRegion
For Each lregion In mRegionMap
    Dim lYAxisPicture As PictureBox
    Set lYAxisPicture = YAxisPicture(lregion.YAxisRegion.handle)
    lYAxisPicture.Width = YAxisWidthCm * TwipsPerCm
    
    If YAxisVisible Then
        If (UserControl.Width - lYAxisPicture.Width) > 0 Then lYAxisPicture.Left = UserControl.Width - lYAxisPicture.Width
    Else
        lYAxisPicture.Left = UserControl.Width
    End If
    ChartRegionPicture(lregion.handle).Width = lYAxisPicture.Left
    
    RegionDividerPicture(lregion.handle).Width = lYAxisPicture.Left
    YRegionDividerPicture(lregion.YAxisRegion.handle).Width = lYAxisPicture.Width
    YRegionDividerPicture(lregion.YAxisRegion.handle).Left = lYAxisPicture.Left
Next

For Each lregion In mRegionMap
    lregion.SetPeriodsInView mScaleLeft, mYAxisPosition - 1
    lregion.PaintDivider
Next

If Not mXAxisRegion Is Nothing Then
    mXAxisRegion.SetPeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
End If

setHorizontalScrollBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setHorizontalScrollBar()
Const ProcName As String = "setHorizontalScrollBar"
On Error GoTo Err

If mPeriods.CurrentPeriodNumber + ChartWidth - 1 > 32767 Then
    HScroll.Max = 32767
ElseIf mPeriods.CurrentPeriodNumber + ChartWidth - 1 < 1 Then
    HScroll.Max = 1
Else
    HScroll.Max = mPeriods.CurrentPeriodNumber + ChartWidth - 1
End If
HScroll.Min = 0

' NB the following calculation has to be done using doubles as for very large charts it can cause an overflow using integers
Dim hscrollVal As Integer
hscrollVal = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.CurrentPeriodNumber + ChartWidth - 1)))
If hscrollVal > HScroll.Max Then
    HScroll.Value = HScroll.Max
ElseIf hscrollVal < HScroll.Min Then
    HScroll.Value = HScroll.Min
Else
    HScroll.Value = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.CurrentPeriodNumber + ChartWidth - 1)))
End If

HScroll.SmallChange = 1
If (ChartWidth - 1) < 1 Then
    HScroll.LargeChange = 1
Else
    HScroll.LargeChange = ChartWidth - 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setProperty( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant, _
                Optional ByRef pPrevValue As Variant) As Boolean
Const ProcName As String = "setProperty"
On Error GoTo Err

setProperty = gSetProperty(mEPhost, pExtProp, pNewValue, pPrevValue)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setRegionDividerLocation( _
                ByVal pRegion As ChartRegion, _
                ByVal currTop As Long) As Long
Const ProcName As String = "setRegionDividerLocation"
On Error GoTo Err

RegionDividerPicture(pRegion.handle).Top = currTop
YRegionDividerPicture(pRegion.YAxisRegion.handle).Top = currTop
If mRegionMap.IsFirst(CLng(RegionDividerPicture(pRegion.handle).Tag)) Then
    RegionDividerPicture(pRegion.handle).Visible = False
    YRegionDividerPicture(pRegion.YAxisRegion.handle).Visible = False
    setRegionDividerLocation = 0
Else
    RegionDividerPicture(pRegion.handle).Visible = True
    YRegionDividerPicture(pRegion.YAxisRegion.handle).Visible = True
    setRegionDividerLocation = RegionDividerPicture(pRegion.handle).Height
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setRegionViewSizeAndLocation( _
                ByVal pRegion As ChartRegion, _
                ByVal currTop As Long) As Long
Const ProcName As String = "setRegionViewSizeAndLocation"
On Error GoTo Err

ChartRegionPicture(pRegion.handle).Height = pRegion.ActualHeight
YAxisPicture(pRegion.YAxisRegion.handle).Height = pRegion.ActualHeight
ChartRegionPicture(pRegion.handle).Top = currTop
YAxisPicture(pRegion.YAxisRegion.handle).Top = currTop
pRegion.NotifyResizedY
setRegionViewSizeAndLocation = pRegion.ActualHeight

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setRegionViewSizes()
Const ProcName As String = "setRegionViewSizes"
On Error GoTo Err

' Now actually set the Heights and positions for the picture boxes

If Not IsDrawingEnabled Then Exit Sub

Dim currTop As Long
Dim lregion As ChartRegion
For Each lregion In mRegionMap
    currTop = currTop + setRegionDividerLocation(lregion, currTop)
    currTop = currTop + setRegionViewSizeAndLocation(lregion, currTop)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub SuppressDrawing(ByVal suppress As Boolean)
Const ProcName As String = "SuppressDrawing"
On Error GoTo Err

Dim lChange As Boolean

If suppress Then
    mSuppressDrawingCount = mSuppressDrawingCount + 1
    If mSuppressDrawingCount = 1 Then lChange = True
ElseIf mSuppressDrawingCount = 0 Then
    lChange = False
Else
    mSuppressDrawingCount = mSuppressDrawingCount - 1
    If mSuppressDrawingCount = 0 Then
        Resize True, True
        lChange = True
    Else
        lChange = False
    End If
End If

gLogger.Log "Suppress drawing count = " & mSuppressDrawingCount, ProcName, ModuleName, LogLevelHighDetail

If lChange Then
    Dim Region As ChartRegion
    For Each Region In mRegions
        Region.IsDrawingEnabled = IsDrawingEnabled
    Next
    If Not mXAxisRegion Is Nothing Then mXAxisRegion.IsDrawingEnabled = IsDrawingEnabled
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unmapRegion( _
                    ByVal Region As ChartRegion)
Const ProcName As String = "unmapRegion"
On Error GoTo Err

mRegionMap.Remove CLng(ChartRegionPicture(Region.handle).Tag)
Unload ChartRegionPicture(Region.handle)
Unload RegionDividerPicture(Region.handle)
If Not Region.YAxisRegion Is Nothing Then
    Unload YRegionDividerPicture(Region.YAxisRegion.handle)
    Unload YAxisPicture(Region.YAxisRegion.handle)
End If
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

