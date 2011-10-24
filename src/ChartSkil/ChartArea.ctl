VERSION 5.00
Begin VB.UserControl Chart 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
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

'================================================================================
' Events
'================================================================================

Event ChartCleared()
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
Private Const PropNamePeriodLength                      As String = "PeriodLength"
Private Const PropNamePeriodUnits                       As String = "PeriodUnits"
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
Private Const PropDfltPeriodLength                      As Long = 5
Private Const PropDfltPeriodUnits                       As Long = TimePeriodMinute
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltHorizontalScrollBarVisible        As Boolean = True
Private Const PropDfltTwipsPerPeriod                    As Long = 150
Private Const PropDfltVerticalGridSpacing               As Long = 1
Private Const PropDfltVerticalGridUnits                 As Long = TimePeriodHour
Private Const PropDfltXAxisVisible                      As Boolean = True
Private Const PropDfltYAxisVisible                      As Boolean = True
Private Const PropDfltYAxisWidthCm                      As Single = 1.5

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
Private mXCursorText As Text

Private mYAxisPosition As Long

Private mSessionBuilder As SessionBuilder
Private mSession As Session
Attribute mSession.VB_VarHelpID = -1

Private mPeriodLength As TimePeriod
Private mPeriodLengthSet As Boolean

Private mVerticalGridTimePeriod As TimePeriod
Private mVerticalGridTimePeriodSet As Boolean

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

Private mBackGroundViewport As Viewport

'================================================================================
' User Control Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"

On Error GoTo Err

Set mRegions = New ChartRegions
mRegions.Chart = Me

Set mRegionMap = New ChartRegionMap

GChart.gRegisterProperties
Set mEPhost = New ExtendedPropertyHost
mEPhost.Style = gChartStylesManager.DefaultStyle.ExtendedPropertyHost

' set a (possibly) temporary style for the saved properties to be applied to
Style = gDefaultChartStyle

Set gBlankCursor = BlankPicture.Picture
Set gSelectorCursor = SelectorPicture.Picture

ReDim mChartBackGradientFillColors(0) As Long
mChartBackGradientFillColors(0) = PropDfltChartBackColor

Set mController = New ChartController
mController.Chart = Me

createXAxisRegion

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_InitProperties()
Const ProcName As String = "UserControl_InitProperties"

On Error GoTo Err

Initialise

Style.HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
Style.VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
Style.Autoscrolling = PropDfltAutoscrolling
Set mPeriodLength = GetTimePeriod(PropDfltPeriodLength, PropDfltPeriodUnits)
PointerCrosshairsColor = PropDfltPointerCrosshairsColor
PointerDiscColor = PropDfltPointerDiscColor
PointerStyle = PropDfltPointerStyle
Style.HorizontalScrollBarVisible = PropDfltHorizontalScrollBarVisible
Style.TwipsPerPeriod = PropDfltTwipsPerPeriod
Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
Style.XAxisVisible = PropDfltXAxisVisible
Style.YAxisWidthCm = PropDfltYAxisWidthCm
Style.YAxisVisible = PropDfltYAxisVisible

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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_ReadProperties"

On Error GoTo Err

Initialise

Style.Autoscrolling = PropBag.ReadProperty(PropNameAutoscrolling, PropDfltAutoscrolling)

Set mPeriodLength = GetTimePeriod(PropBag.ReadProperty(PropNamePeriodLength, PropDfltPeriodLength), _
                                    PropBag.ReadProperty(PropNamePeriodUnits, PropDfltPeriodUnits))


Style.ChartBackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)

Style.HorizontalMouseScrollingAllowed = PropBag.ReadProperty(PropNameHorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed)

Style.HorizontalScrollBarVisible = PropBag.ReadProperty(PropNameHorizontalScrollBarVisible, PropDfltHorizontalScrollBarVisible)


PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)

PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)

PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)

Style.TwipsPerPeriod = PropBag.ReadProperty(PropNameTwipsPerPeriod, PropDfltTwipsPerPeriod)

Set mVerticalGridTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNameVerticalGridSpacing, PropDfltVerticalGridSpacing), _
                        PropBag.ReadProperty(PropNameVerticalGridUnits, PropDfltVerticalGridUnits))

Style.VerticalMouseScrollingAllowed = PropBag.ReadProperty(PropNameVerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed)

Style.XAxisVisible = PropBag.ReadProperty(PropNameXAxisVisible, PropDfltXAxisVisible)

Style.YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)

Style.YAxisVisible = PropBag.ReadProperty(PropNameYAxisVisible, PropDfltYAxisVisible)

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
'gLogger.Log LogLevelDetail, "ChartSkil chart terminated"
Debug.Print "ChartSkil chart terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

PropBag.WriteProperty PropNameHorizontalMouseScrollingAllowed, Style.HorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed
PropBag.WriteProperty PropNameVerticalMouseScrollingAllowed, Style.VerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed
PropBag.WriteProperty PropNameAutoscrolling, Style.Autoscrolling, PropDfltAutoscrolling
PropBag.WriteProperty PropNameChartBackColor, Style.ChartBackColor
PropBag.WriteProperty PropNamePeriodLength, mPeriodLength.Length, PropDfltPeriodLength
PropBag.WriteProperty PropNamePeriodUnits, mPeriodLength.Units, PropDfltPeriodUnits
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNamePointerStyle, mPointerStyle, PropDfltPointerStyle
PropBag.WriteProperty PropNameHorizontalScrollBarVisible, Style.HorizontalScrollBarVisible, PropDfltHorizontalScrollBarVisible
PropBag.WriteProperty PropNameTwipsPerPeriod, Style.TwipsPerPeriod, PropDfltTwipsPerPeriod
PropBag.WriteProperty PropNameVerticalGridSpacing, mVerticalGridTimePeriod.Length, PropDfltVerticalGridSpacing
PropBag.WriteProperty PropNameVerticalGridUnits, mVerticalGridTimePeriod.Units, PropDfltVerticalGridUnits
PropBag.WriteProperty PropNameXAxisVisible, Style.XAxisVisible, PropDfltXAxisVisible
PropBag.WriteProperty PropNameYAxisVisible, Style.YAxisVisible, PropDfltYAxisVisible
PropBag.WriteProperty PropNameYAxisWidthCm, Style.YAxisWidthCm, PropDfltYAxisWidthCm

End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_Click(index As Integer)
Const ProcName As String = "ChartRegionPicture_Click"

On Error GoTo Err

If index = 0 Then Exit Sub

getDataRegionFromPictureIndex(index).Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartRegionPicture_DblClick(index As Integer)
Const ProcName As String = "ChartRegionPicture_DblClick"

On Error GoTo Err

If index = 0 Then Exit Sub

getDataRegionFromPictureIndex(index).DblCLick

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
Dim lregion As ChartRegion

Const ProcName As String = "ChartRegionPicture_MouseMove"

On Error GoTo Err

If index = 0 Then Exit Sub

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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub XAxisPicture_DblClick()
Const ProcName As String = "XAxisPicture_DblClick"

On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.DblCLick

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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub YAxisPicture_DblClick(index As Integer)
Const ProcName As String = "YAxisPicture_DblClick"

On Error GoTo Err

getYAxisRegionFromPictureIndex(index).DblCLick

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

Private Sub mEPhost_Change(pEv As TWUtilities30.ChangeEventData)
Const ProcName As String = "mEPhost_Change"
On Error GoTo Err

If Not mBackGroundViewport Is Nothing Then mBackGroundViewport.BackColor = ChartBackColor

HScroll.Visible = HorizontalScrollBarVisible

XAxisPicture.Visible = XAxisVisible

Dim lregion As ChartRegion
For Each lregion In mRegions
    lregion.CrosshairLineStyle = CrosshairLineStyle
Next

mRegions.DefaultDataRegionStyle = DefaultRegionStyle
mRegions.DefaultYAxisRegionStyle = DefaultYAxisRegionStyle

If Not mXAxisRegion Is Nothing Then mXAxisRegion.BaseStyle = XAxisRegionStyle

If Not mXCursorText Is Nothing Then mXCursorText.LocalStyle = XCursorTextStyle

Resize True, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mEPhost_ExtendedPropertyChanged(pEv As ExtProps30.ExtendedPropertyChangedEventData)
Const ProcName As String = "mEPhost_ExtendedPropertyChanged"
On Error GoTo Err

Dim lregion As ChartRegion

If pEv.ExtendedProperty Is gChartBackColorProperty Then
    mBackGroundViewport.BackColor = ChartBackColor
ElseIf pEv.ExtendedProperty Is gHorizontalScrollBarVisibleProperty Then
    HScroll.Visible = HorizontalScrollBarVisible
    Resize False, True
ElseIf pEv.ExtendedProperty Is gTwipsPerPeriodProperty Then
    resizeX
ElseIf pEv.ExtendedProperty Is gXAxisVisibleProperty Then
    mRegions.ResizeY mUserResizingRegions
    XAxisPicture.Visible = XAxisVisible
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
ElseIf pEv.ExtendedProperty Is gXCursorTextStyleProperty Then
    mXCursorText.LocalStyle = XCursorTextStyle
End If

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' mRegions Event Handlers
'================================================================================

Private Sub mRegions_CollectionChanged(ev As TWUtilities30.CollectionChangeEventData)
Dim rgn As ChartRegion

Const ProcName As String = "mRegions_CollectionChanged"

On Error GoTo Err

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
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

Autoscrolling = mEPhost.GetValue(GChart.gAutoscrollingProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Autoscrolling(ByVal Value As Boolean)
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

setProperty GChart.gAutoscrollingProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAutoscrolling, Value
PropertyChanged PropNameAutoscrolling

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get Availableheight() As Long
Const ProcName As String = "Availableheight"

On Error GoTo Err

Availableheight = IIf(XAxisVisible, XAxisPicture.Top, UserControl.ScaleHeight) - _
                    (mRegions.Count - 1) * RegionDividerPicture(0).Height
If Availableheight < 1 Then Availableheight = 1

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Attribute ChartBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

ChartBackColor = mEPhost.GetValue(GChart.gChartBackColorProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

If Value Is mConfig Then Exit Property
Set mConfig = Value

If isLocalValueSet(GChart.gTwipsPerPeriodProperty) Then mConfig.SetSetting ConfigSettingTwipsPerPeriod, mEPhost.GetLocalValue(GChart.gTwipsPerPeriodProperty)

If isLocalValueSet(GChart.gAutoscrollingProperty) Then mConfig.SetSetting ConfigSettingAutoscrolling, mEPhost.GetLocalValue(GChart.gAutoscrollingProperty)
If isLocalValueSet(GChart.gChartBackColorProperty) Then mConfig.SetSetting ConfigSettingChartBackColor, mEPhost.GetLocalValue(GChart.gChartBackColorProperty)
If isLocalValueSet(GChart.gHorizontalMouseScrollingAllowedProperty) Then mConfig.SetSetting ConfigSettingHorizontalMouseScrollingAllowed, mEPhost.GetLocalValue(GChart.gHorizontalMouseScrollingAllowedProperty)
If isLocalValueSet(GChart.gHorizontalScrollBarVisibleProperty) Then mConfig.SetSetting ConfigSettingHorizontalScrollBarVisible, mEPhost.GetLocalValue(GChart.gHorizontalScrollBarVisibleProperty)
If isLocalValueSet(GChart.gVerticalMouseScrollingAllowedProperty) Then mConfig.SetSetting ConfigSettingVerticalMouseScrollingAllowed, mEPhost.GetLocalValue(GChart.gVerticalMouseScrollingAllowedProperty)
If isLocalValueSet(GChart.gXAxisVisibleProperty) Then mConfig.SetSetting ConfigSettingXAxisVisible, mEPhost.GetLocalValue(GChart.gXAxisVisibleProperty)
If isLocalValueSet(GChart.gYAxisVisibleProperty) Then mConfig.SetSetting ConfigSettingYAxisVisible, mEPhost.GetLocalValue(GChart.gYAxisVisibleProperty)
If isLocalValueSet(GChart.gYAxisWidthCmProperty) Then mConfig.SetSetting ConfigSettingYAxisWidthCm, mEPhost.GetLocalValue(GChart.gYAxisWidthCmProperty)

If isLocalValueSet(GChart.gCrosshairLineStyleProperty) Then mEPhost.GetLocalValue(GChart.gCrosshairLineStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionCrosshairLineStyle)
If isLocalValueSet(GChart.gDefaultRegionStyleProperty) Then mEPhost.GetLocalValue(GChart.gDefaultRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultRegionStyle)
If isLocalValueSet(GChart.gDefaultYAxisRegionStyleProperty) Then mEPhost.GetLocalValue(GChart.gDefaultYAxisRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
If isLocalValueSet(GChart.gXAxisRegionStyleProperty) Then mEPhost.GetLocalValue(GChart.gXAxisRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXAxisRegionStyle)
If isLocalValueSet(GChart.gXCursorTextStyleProperty) Then mEPhost.GetLocalValue(GChart.gXCursorTextStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXCursorTextStyle)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Controller() As ChartController
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CrosshairLineStyle() As LineStyle
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

Set CrosshairLineStyle = mEPhost.GetValue(GChart.gCrosshairLineStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentPeriodNumber() As Long
Attribute CurrentPeriodNumber.VB_MemberFlags = "400"
Const ProcName As String = "CurrentPeriodNumber"

On Error GoTo Err

CurrentPeriodNumber = mPeriods.CurrentPeriodNumber

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentSessionEndTime() As Date
Attribute CurrentSessionEndTime.VB_MemberFlags = "400"
CurrentSessionEndTime = mSession.CurrentSessionEndTime
End Property

Public Property Get CurrentSessionStartTime() As Date
Attribute CurrentSessionStartTime.VB_MemberFlags = "400"
CurrentSessionStartTime = mSession.CurrentSessionStartTime
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultRegionStyle"
On Error GoTo Err

Set DefaultRegionStyle = mEPhost.GetValue(GChart.gDefaultRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

Set DefaultYAxisRegionStyle = mEPhost.GetValue(GChart.gDefaultYAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FirstVisiblePeriod() As Long
Attribute FirstVisiblePeriod.VB_MemberFlags = "400"
FirstVisiblePeriod = mScaleLeft
End Property

Public Property Let FirstVisiblePeriod(ByVal Value As Long)
Const ProcName As String = "FirstVisiblePeriod"

On Error GoTo Err

ScrollX Value - mScaleLeft + 1

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

HorizontalMouseScrollingAllowed = mEPhost.GetValue(GChart.gHorizontalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"

On Error GoTo Err

setProperty GChart.gHorizontalMouseScrollingAllowedProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHorizontalMouseScrollingAllowed, Value
PropertyChanged PropNameHorizontalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Const ProcName As String = "HorizontalScrollBarVisible"

On Error GoTo Err

HorizontalScrollBarVisible = mEPhost.GetValue(GChart.gHorizontalScrollBarVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal Value As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"

On Error GoTo Err

setProperty GChart.gHorizontalScrollBarVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHorizontalScrollBarVisible, Value
PropertyChanged PropNameHorizontalScrollBarVisible

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsDrawingEnabled() As Boolean
IsDrawingEnabled = (mSuppressDrawingCount = 0)
End Property

Public Property Get IsGridHidden() As Boolean
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PeriodLength( _
                ByVal Value As TimePeriod)
Const ProcName As String = "PeriodLength"

On Error GoTo Err

If mPeriodLengthSet Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "PeriodLength has already been set"

Set mPeriodLength = Value

mPeriodLengthSet = True

If Not mVerticalGridTimePeriodSet Then calcVerticalGridParams
setRegionVerticalGridTimePeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get PeriodLength() As TimePeriod
Set PeriodLength = mPeriodLength
End Property

Public Property Get Periods() As Periods
Attribute Periods.VB_MemberFlags = "400"
Set Periods = mPeriods
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
PointerCrosshairsColor = CrosshairLineStyle.Color
End Property

Public Property Let PointerCrosshairsColor(ByVal Value As OLE_COLOR)
Attribute PointerCrosshairsColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
Const ProcName As String = "PointerCrosshairsColor"

On Error GoTo Err
CrosshairLineStyle.Color = Value
PropertyChanged PropNamePointerCrosshairsColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Attribute PointerDiscColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerDiscColor = mPointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal Value As OLE_COLOR)
Dim Region As ChartRegion
Const ProcName As String = "PointerDiscColor"

On Error GoTo Err

mPointerDiscColor = Value
For Each Region In mRegions
    Region.PointerDiscColor = Value
Next
PropertyChanged PropNamePointerDiscColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerIcon() As IPictureDisp
Attribute PointerIcon.VB_MemberFlags = "400"
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal Value As IPictureDisp)
Dim Region As ChartRegion

Const ProcName As String = "PointerIcon"

On Error GoTo Err

If Value Is Nothing Then Exit Property
If Value Is mPointerIcon Then Exit Property

Set mPointerIcon = Value

If mPointerStyle = PointerCustom Then
    For Each Region In mRegions
        Region.PointerStyle = PointerCustom
    Next
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerMode() As PointerModes
Attribute PointerMode.VB_MemberFlags = "400"
PointerMode = mPointerMode
End Property

Public Property Get PointerStyle() As PointerStyles
Attribute PointerStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal Value As PointerStyles)
Dim Region As ChartRegion

Const ProcName As String = "PointerStyle"

On Error GoTo Err

If Value = mPointerStyle Then Exit Property

mPointerStyle = Value

If mPointerStyle = PointerCustom And mPointerIcon Is Nothing Then
    ' we'll notify the region when an icon is supplied
    Exit Property
End If

For Each Region In mRegions
    If mPointerStyle = PointerCustom Then Region.PointerIcon = mPointerIcon
    Region.PointerStyle = mPointerStyle
Next
PropertyChanged PropNamePointerStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Regions() As ChartRegions
Set Regions = mRegions
End Property

Public Property Get SessionEndTime() As Date
Attribute SessionEndTime.VB_MemberFlags = "400"
SessionEndTime = mSession.SessionEndTime
End Property

Public Property Let SessionEndTime(ByVal val As Date)
Const ProcName As String = "SessionEndTime"

On Error GoTo Err

If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Value must be a time only"
mSessionBuilder.SessionEndTime = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionStartTime() As Date
Attribute SessionStartTime.VB_MemberFlags = "400"
SessionStartTime = mSession.SessionStartTime
End Property

Public Property Let SessionStartTime(ByVal val As Date)
Const ProcName As String = "SessionStartTime"

On Error GoTo Err

If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Value must be a time only"
mSessionBuilder.SessionStartTime = val

Dim Region As ChartRegion
For Each Region In mRegions
    Region.SessionStartTime = val
Next

mXAxisRegion.SessionStartTime = val
Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Style(ByVal Value As ChartStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Set mStyle = Value
If mStyle Is Nothing Then Set mStyle = gChartStylesManager.DefaultStyle
mEPhost.Style = mStyle.ExtendedPropertyHost
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingStyle, mStyle.Name

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Style() As ChartStyle
Set Style = mStyle
End Property

Public Property Get TwipsPerPeriod() As Long
Const ProcName As String = "TwipsPerPeriod"

On Error GoTo Err

TwipsPerPeriod = mEPhost.GetValue(GChart.gTwipsPerPeriodProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TwipsPerPeriod(ByVal Value As Long)
Attribute TwipsPerPeriod.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
Const ProcName As String = "TwipsPerPeriod"

On Error GoTo Err

setProperty GChart.gTwipsPerPeriodProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingTwipsPerPeriod, Value
PropertyChanged PropNameTwipsPerPeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Set VerticalGridTimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "VerticalGridTimePeriod"

On Error GoTo Err

If mVerticalGridTimePeriodSet Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "verticalGridTimePeriod has already been set"

If Value.Length <= 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "verticalGridTimePeriod length must be >0"
Select Case Value.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case TimePeriodWeek
Case TimePeriodMonth
Case TimePeriodYear
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "verticalGridTimePeriod Units must be a member of the TimePeriodUnits enum"
End Select

Set mVerticalGridTimePeriod = Value
mVerticalGridTimePeriodSet = True

setRegionVerticalGridTimePeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get VerticalGridTimePeriod() As TimePeriod
Attribute VerticalGridTimePeriod.VB_MemberFlags = "400"
Set VerticalGridTimePeriod = mVerticalGridTimePeriod
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

VerticalMouseScrollingAllowed = mEPhost.GetValue(GChart.gVerticalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

setProperty GChart.gVerticalMouseScrollingAllowedProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingVerticalMouseScrollingAllowed, Value
PropertyChanged PropNameVerticalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "XAxisRegionStyle"
On Error GoTo Err

Set XAxisRegionStyle = mEPhost.GetValue(GChart.gXAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XAxisVisible() As Boolean
Attribute XAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "XAxisVisible"
On Error GoTo Err

XAxisVisible = mEPhost.GetValue(GChart.gXAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let XAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "XAxisVisible"

On Error GoTo Err

setProperty GChart.gXAxisVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingXAxisVisible, Value
PropertyChanged PropNameXAxisVisible

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let XCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "XCursorTextStyle"

On Error GoTo Err

Dim prevValue As TextStyle
If setProperty(GChart.gXCursorTextStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXCursorTextStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XCursorTextStyle() As TextStyle
Const ProcName As String = "XCursorTextStyle"
On Error GoTo Err

Set XCursorTextStyle = mEPhost.GetValue(GChart.gXCursorTextStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisPosition() As Long
Attribute YAxisPosition.VB_MemberFlags = "400"
YAxisPosition = mYAxisPosition
End Property

Public Property Get YAxisVisible() As Boolean
Attribute YAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "YAxisVisible"
On Error GoTo Err

YAxisVisible = mEPhost.GetValue(GChart.gYAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "YAxisVisible"

On Error GoTo Err

setProperty GChart.gYAxisVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYAxisVisible, Value
PropertyChanged PropNameYAxisVisible

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_ProcData.VB_Invoke_Property = ";Appearance"
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub AddPeriod( _
                ByVal pPeriodNumber As Long, _
                ByVal pTimestamp As Date)
Dim Region As ChartRegion
Dim ev As SessionEventData

Const ProcName As String = "AddPeriod"

On Error GoTo Err

ev = mSessionBuilder.SetSessionCurrentTime(pTimestamp)

For Each Region In mRegions
    Region.AddPeriod pPeriodNumber, pTimestamp
    Select Case mPeriodLength.Units
    Case TimePeriodSecond, _
            TimePeriodMinute, _
            TimePeriodHour, _
            TimePeriodTickMovement, _
            TimePeriodTickVolume, _
            TimePeriodVolume
        If ev.changeType = SessionChangeEnd Then
            Region.AddSessionEndGridline pPeriodNumber, ev.Timestamp
        ElseIf ev.changeType = SessionChangeStart Then
            Region.AddSessionStartGridline pPeriodNumber, ev.Timestamp
        End If
    End Select
Next

mXAxisRegion.AddPeriod pPeriodNumber, pTimestamp
If IsDrawingEnabled Then setHorizontalScrollBar
If Autoscrolling Then ScrollX 1
    
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function ClearChart()

Const ProcName As String = "ClearChart"

On Error GoTo Err

DisableDrawing

Clear

Initialise
mYAxisPosition = 1
createXAxisRegion

EnableDrawing

RaiseEvent ChartCleared
mController.fireChartCleared
Debug.Print "Chart cleared"

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function CreateDataRegionCanvas(ByVal index As Long) As Canvas
Const ProcName As String = "CreateDataRegionCanvas"

On Error GoTo Err

Load ChartRegionPicture(index)
Set CreateDataRegionCanvas = CreateCanvas(ChartRegionPicture(index), RegionTypeData)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function CreateYAxisRegionCanvas(ByVal index As Long) As Canvas
Const ProcName As String = "CreateYAxisRegionCanvas"

On Error GoTo Err

Load YAxisPicture(index)
Set CreateYAxisRegionCanvas = CreateCanvas(YAxisPicture(index), RegionTypeYAxis)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub DisableDrawing()
Const ProcName As String = "DisableDrawing"

On Error GoTo Err

SuppressDrawing True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub EnableDrawing()
Const ProcName As String = "EnableDrawing"

On Error GoTo Err

SuppressDrawing False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function GetXFromTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal forceNewPeriod As Boolean, _
                Optional ByVal duplicateNumber As Long) As Double
Dim lPeriod As Period
Dim periodEndtime As Date

Const ProcName As String = "GetXFromTimestamp"

On Error GoTo Err

Select Case PeriodLength.Units
Case TimePeriodNone, _
        TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    
    On Error Resume Next
    Set lPeriod = mPeriods.Item(Timestamp)
    On Error GoTo Err
    
    If lPeriod Is Nothing Then
        If mPeriods.Count = 0 Then
            Set lPeriod = mPeriods.AddPeriod(Timestamp)
        ElseIf Timestamp < mPeriods.Item(1).Timestamp Then
            Set lPeriod = mPeriods.Item(1)
            Timestamp = lPeriod.Timestamp
        Else
            Set lPeriod = mPeriods.AddPeriod(Timestamp)
        End If
    End If
    
    periodEndtime = BarEndTime(lPeriod.Timestamp, _
                            PeriodLength, _
                            SessionStartTime)
    GetXFromTimestamp = lPeriod.PeriodNumber + (Timestamp - lPeriod.Timestamp) / (periodEndtime - lPeriod.Timestamp)
    
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If Not forceNewPeriod Then
        On Error Resume Next
        Set lPeriod = mPeriods.ItemDup(Timestamp, duplicateNumber)
        On Error GoTo Err
        
        If lPeriod Is Nothing Then
            Set lPeriod = mPeriods.AddPeriod(Timestamp, True)
        End If
        GetXFromTimestamp = lPeriod.PeriodNumber
    Else
        Set lPeriod = mPeriods.AddPeriod(Timestamp, True)
        GetXFromTimestamp = lPeriod.PeriodNumber
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Public Sub HideGrid()
Dim Region As ChartRegion

Const ProcName As String = "HideGrid"

On Error GoTo Err

If mHideGrid Then Exit Sub

mHideGrid = True
For Each Region In mRegions
    Region.HideGrid
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean

Const ProcName As String = "IsTimeInSession"

On Error GoTo Err

IsTimeInSession = mSession.IsTimeInSession(Timestamp)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
If mConfig.GetSetting(ConfigSettingTwipsPerPeriod) <> "" Then TwipsPerPeriod = mConfig.GetSetting(ConfigSettingTwipsPerPeriod, "120")
If mConfig.GetSetting(ConfigSettingVerticalMouseScrollingAllowed) <> "" Then VerticalMouseScrollingAllowed = mConfig.GetSetting(ConfigSettingVerticalMouseScrollingAllowed, "true")
If mConfig.GetSetting(ConfigSettingXAxisVisible) <> "" Then XAxisVisible = mConfig.GetSetting(ConfigSettingXAxisVisible, "true")
If mConfig.GetSetting(ConfigSettingYAxisVisible) <> "" Then YAxisVisible = mConfig.GetSetting(ConfigSettingYAxisVisible, "true")
If mConfig.GetSetting(ConfigSettingYAxisWidthCm) <> "" Then YAxisWidthCm = mConfig.GetSetting(ConfigSettingYAxisWidthCm, "1.8")

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

Dim ts As TextStyle
If Not mConfig.GetConfigurationSection(ConfigSectionXCursorTextStyle) Is Nothing Then
    Set ts = New TextStyle
    ts.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionXCursorTextStyle)
    XCursorTextStyle = ts
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"

On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ScrollX(ByVal Value As Long)
Dim Region As ChartRegion

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

For Each Region In mRegions
    Region.SetPeriodsInView mScaleLeft, mYAxisPosition - 1
Next

mXAxisRegion.SetPeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
setHorizontalScrollBar

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetPointerModeDefault()
Dim Region As ChartRegion

Const ProcName As String = "SetPointerModeDefault"

On Error GoTo Err

mPointerMode = PointerModeDefault
For Each Region In mRegions
    Region.SetPointerModeDefault
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetPointerModeSelection()
Dim Region As ChartRegion

Const ProcName As String = "SetPointerModeSelection"

On Error GoTo Err

mPointerMode = PointerModeSelection

For Each Region In mRegions
    Region.SetPointerModeSelection
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetPointerModeTool( _
                Optional ByVal toolPointerStyle As PointerStyles = PointerTool, _
                Optional ByVal icon As IPictureDisp)
Dim Region As ChartRegion

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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "toolPointerStyle must be a member of the PointerStyles enum"
End Select
For Each Region In mRegions
    Region.SetPointerModeTool toolPointerStyle, icon
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ShowGrid()
Dim Region As ChartRegion

Const ProcName As String = "ShowGrid"

On Error GoTo Err

If Not mHideGrid Then Exit Sub

mHideGrid = False
For Each Region In mRegions
    Region.ShowGrid
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub calcVerticalGridParams()

Const ProcName As String = "calcVerticalGridParams"

On Error GoTo Err

Select Case mPeriodLength.Units
Case TimePeriodNone
    Set mVerticalGridTimePeriod = Nothing
Case TimePeriodSecond
    Select Case mPeriodLength.Length
    Case 1
        Set mVerticalGridTimePeriod = GetTimePeriod(15, TimePeriodSecond)
    Case 2
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodSecond)
    Case 3
        Set mVerticalGridTimePeriod = GetTimePeriod(20, TimePeriodSecond)
    Case 4
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMinute)
    Case 5
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMinute)
    Case 6
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 10
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 12
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 15
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 20
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 30
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case Else
        Set mVerticalGridTimePeriod = Nothing
    End Select
Case TimePeriodMinute
    Select Case mPeriodLength.Length
    Case 1
        Set mVerticalGridTimePeriod = GetTimePeriod(15, TimePeriodMinute)
    Case 2
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodMinute)
    Case 3
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodMinute)
    Case 4
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 5
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 6
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 10
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 12
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 15
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 20
        Set mVerticalGridTimePeriod = GetTimePeriod(4, TimePeriodHour)
    Case 30
        Set mVerticalGridTimePeriod = GetTimePeriod(4, TimePeriodHour)
    Case 60
        Set mVerticalGridTimePeriod = GetTimePeriod(6, TimePeriodHour)
    Case Else
        Set mVerticalGridTimePeriod = Nothing
    End Select
Case TimePeriodHour
        Set mVerticalGridTimePeriod = GetTimePeriod(6, TimePeriodHour)
Case TimePeriodDay
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodWeek)
Case TimePeriodWeek
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodMonth
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodYear)
Case TimePeriodYear
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodYear)
Case TimePeriodVolume
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodVolume)
Case TimePeriodTickVolume
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodTickVolume)
Case TimePeriodTickMovement
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodTickMovement)
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
  
End Sub

Private Sub Clear()
Dim lregion As ChartRegion
Dim en As Enumerator

Const ProcName As String = "Clear"

On Error GoTo Err

Set en = mRegions.Enumerator

Do While en.MoveNext
    Set lregion = en.Current
    lregion.ClearRegion
    en.Remove
Loop

If Not mXAxisRegion Is Nothing Then mXAxisRegion.ClearRegion
XAxisPicture.Cls
Set mXAxisRegion = Nothing
Set mXCursorText = Nothing
If Not mPeriods Is Nothing Then mPeriods.Finish
Set mPeriods = Nothing

finishBackgroundCanvas

mInitialised = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function convertXAxisPictureMouseXtoContainerCoords( _
                ByVal X As Single) As Single
Const ProcName As String = "convertXAxisPictureMouseXtoContainerCoords"

On Error GoTo Err

convertXAxisPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(XAxisPicture, X)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function convertXAxisPictureMouseYtoContainerCoords( _
                ByVal Y As Single) As Single
Const ProcName As String = "convertXAxisPictureMouseYtoContainerCoords"

On Error GoTo Err

convertXAxisPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(XAxisPicture, Y)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function CreateCanvas( _
                ByVal Surface As PictureBox, _
                ByVal pRegionType As RegionTypes) As Canvas
Const ProcName As String = "CreateCanvas"

On Error GoTo Err

Set CreateCanvas = New Canvas
CreateCanvas.Surface = Surface
CreateCanvas.RegionType = pRegionType

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub createXAxisRegion()
Dim afont As StdFont

Const ProcName As String = "createXAxisRegion"

On Error GoTo Err

Set mXAxisRegion = New ChartRegion

mXAxisRegion.Initialise "", Me, CreateCanvas(XAxisPicture, RegionTypeXAxis), RegionTypeXAxis
                        
mXAxisRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
mXAxisRegion.Bottom = 0
mXAxisRegion.Top = 1

Set mXCursorText = mXAxisRegion.AddText(LayerNumbers.LayerPointer)
mXCursorText.LocalStyle = XCursorTextStyle


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub displayXAxisLabel(ByVal X As Single)
Dim thisPeriod As Period
Dim PeriodNumber As Long

Const ProcName As String = "displayXAxisLabel"

On Error GoTo Err

If Round(X) >= mYAxisPosition Then Exit Sub
If mPeriods.Count = 0 Then Exit Sub

On Error Resume Next
PeriodNumber = Round(X)
Set thisPeriod = mPeriods(PeriodNumber)
On Error GoTo Err
If thisPeriod Is Nothing Then
    mXCursorText.Text = ""
    Exit Sub
End If

mXCursorText.Position = gNewPoint( _
                            PeriodNumber, _
                            0, _
                            CoordsLogical, _
                            CoordsCounterDistance)

Select Case mPeriodLength.Units
Case TimePeriodNone, TimePeriodMinute, TimePeriodHour
    mXCursorText.Text = FormatDateTime(thisPeriod.Timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(thisPeriod.Timestamp, vbShortTime)
Case TimePeriodSecond, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    mXCursorText.Text = FormatDateTime(thisPeriod.Timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(thisPeriod.Timestamp, vbLongTime)
Case Else
    mXCursorText.Text = FormatDateTime(thisPeriod.Timestamp, vbShortDate)
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub finishBackgroundCanvas()
Const ProcName As String = "finishBackgroundCanvas"

On Error GoTo Err

gLogger.Log "Finish background canvas", ProcName, ModuleName, LogLevelHighDetail
If Not mBackGroundViewport Is Nothing Then mBackGroundViewport.Finish
Set mBackGroundViewport = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getDataRegionFromPictureIndex( _
                ByVal index As Long) As ChartRegion
Const ProcName As String = "getDataRegionFromPictureIndex"

On Error GoTo Err

Set getDataRegionFromPictureIndex = mRegionMap.Item(CLng(ChartRegionPicture(index).Tag))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getYAxisRegionFromPictureIndex( _
                ByVal index As Long) As ChartRegion
Const ProcName As String = "getYAxisRegionFromPictureIndex"

On Error GoTo Err

Set getYAxisRegionFromPictureIndex = mRegions.ItemFromHandle(CLng(YAxisPicture(index).Tag))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub Initialise()
Const ProcName As String = "Initialise"

On Error GoTo Err

Dim btn As Button

mPrevHeight = UserControl.Height

Set mPeriods = New Periods
mPeriods.Chart = Me

Set mSessionBuilder = New SessionBuilder
Set mSession = mSessionBuilder.Session

setupBackgroundViewport

mPeriodLengthSet = False
mVerticalGridTimePeriodSet = False

mPointerMode = PointerModes.PointerModeDefault

mYAxisPosition = 1
mScaleWidth = CSng(XAxisPicture.Width) / CSng(TwipsPerPeriod) - 0.5!
mScaleLeft = calcScaleLeft
mScaleHeight = -100
mScaleTop = 100

HScroll.Value = 0

mInitialised = True

Resize True, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function isLocalValueSet(ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "isLocalValueSet"
On Error GoTo Err

isLocalValueSet = mEPhost.IsPropertySet(pExtProp)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub mapRegion(pRegion As ChartRegion)
Dim index As Long
Dim mapHandle As Long
Dim btn As Button

Const ProcName As String = "mapRegion"

On Error GoTo Err

If pRegion Is Nothing Then Exit Sub

index = pRegion.handle
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
pRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
pRegion.SessionStartTime = mSession.SessionStartTime

If mHideGrid Then pRegion.HideGrid

Load RegionDividerPicture(index)
RegionDividerPicture(index).Tag = mapHandle
RegionDividerPicture(index).ZOrder 0
RegionDividerPicture(index).Visible = (Not mRegionMap.IsFirst(mapHandle))
pRegion.Divider = RegionDividerPicture(index)

Load YRegionDividerPicture(index)
YRegionDividerPicture(index).Tag = mapHandle
YRegionDividerPicture(index).ZOrder 0
YRegionDividerPicture(index).Visible = (Not mRegionMap.IsFirst(mapHandle))
pRegion.YAxisRegion.Divider = YRegionDividerPicture(index)

YAxisPicture(index).Tag = pRegion.YAxisRegion.handle
YAxisPicture(index).Align = vbAlignNone
YAxisPicture(index).Left = ChartRegionPicture(index).Width
YAxisPicture(index).Width = YAxisWidthCm * TwipsPerCm
YAxisPicture(index).Visible = YAxisVisible

XAxisPicture.Visible = XAxisVisible

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub MouseMove( _
                ByVal targetRegion As ChartRegion, _
                ByVal Button As Long, _
                ByVal Shift As Long, _
                ByRef X As Single, _
                ByRef Y As Single)
Dim Region As ChartRegion


Const ProcName As String = "MouseMove"

On Error GoTo Err

For Each Region In mRegions
    If Region Is targetRegion Then
        'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & X & " y=" & Y
        If (mPointerMode = PointerModeDefault And _
                ((Region.CursorSnapsToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
                (Not Region.CursorSnapsToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
            (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
        Then
            Dim YScaleQuantum As Double
            YScaleQuantum = Region.YScaleQuantum
            If YScaleQuantum <> 0 Then Y = YScaleQuantum * Int((Y + YScaleQuantum / 10000) / YScaleQuantum)
        End If
        Region.DrawCursor Button, Shift, X, Y
        
    Else
        'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & X & " y=" & MinusInfinitySingle
        Region.DrawCursor Button, Shift, X, MinusInfinitySingle
    End If
Next
displayXAxisLabel Round(X)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub resizeBackground()
Const ProcName As String = "resizeBackground"

On Error GoTo Err

If mRegions.Count > 0 Then Exit Sub
XAxisPicture.Visible = False
ChartRegionPicture(0).Visible = False
ChartRegionPicture(0).Move 0, 0, UserControl.Width, UserControl.Height
mBackGroundViewport.BackColor = ChartBackColor
mBackGroundViewport.Left = 0
mBackGroundViewport.Right = 1
mBackGroundViewport.Bottom = 0
mBackGroundViewport.Top = 1
mBackGroundViewport.PaintBackground
mBackGroundViewport.Canvas.ZOrder 1
ChartRegionPicture(0).Visible = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub resizeX()
Dim Region As ChartRegion

Const ProcName As String = "resizeX"

On Error GoTo Err

If Not mInitialised Then Exit Sub

mScaleWidth = CSng(XAxisPicture.Width) / CSng(TwipsPerPeriod)
mScaleLeft = calcScaleLeft

For Each Region In mRegionMap
    If (UserControl.Width - YAxisPicture(Region.handle).Width) > 0 Then
        YAxisPicture(Region.handle).Left = UserControl.Width - IIf(YAxisVisible, YAxisPicture(Region.handle).Width, 0)
        ChartRegionPicture(Region.handle).Width = YAxisPicture(Region.handle).Left
    End If
    RegionDividerPicture(Region.handle).Width = YAxisPicture(Region.handle).Left
    YRegionDividerPicture(Region.handle).Width = YAxisPicture(Region.handle).Width
    YRegionDividerPicture(Region.handle).Left = YAxisPicture(Region.handle).Left
Next

For Each Region In mRegionMap
    Region.SetPeriodsInView mScaleLeft, mYAxisPosition - 1
    Region.PaintDivider
Next

If Not mXAxisRegion Is Nothing Then
    mXAxisRegion.SetPeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
End If

setHorizontalScrollBar

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setHorizontalScrollBar()
Dim hscrollVal As Integer

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setRegionDividerLocation( _
                ByVal pRegion As ChartRegion, _
                ByVal currTop As Long) As Long
Const ProcName As String = "setRegionDividerLocation"

On Error GoTo Err

RegionDividerPicture(pRegion.handle).Top = currTop
YRegionDividerPicture(pRegion.handle).Top = currTop
If mRegionMap.IsFirst(CLng(RegionDividerPicture(pRegion.handle).Tag)) Then
    RegionDividerPicture(pRegion.handle).Visible = False
    YRegionDividerPicture(pRegion.handle).Visible = False
    setRegionDividerLocation = 0
Else
    RegionDividerPicture(pRegion.handle).Visible = True
    YRegionDividerPicture(pRegion.handle).Visible = True
    setRegionDividerLocation = RegionDividerPicture(pRegion.handle).Height
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setRegionVerticalGridTimePeriod()
Dim Region As ChartRegion
Const ProcName As String = "setRegionVerticalGridTimePeriod"

On Error GoTo Err

mXAxisRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
For Each Region In mRegions
    Region.VerticalGridTimePeriod = mVerticalGridTimePeriod
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function setRegionViewSizeAndLocation( _
                ByVal pRegion As ChartRegion, _
                ByVal currTop As Long) As Long
Const ProcName As String = "setRegionViewSizeAndLocation"

On Error GoTo Err

ChartRegionPicture(pRegion.handle).Height = pRegion.ActualHeight
YAxisPicture(pRegion.handle).Height = pRegion.ActualHeight
ChartRegionPicture(pRegion.handle).Top = currTop
YAxisPicture(pRegion.handle).Top = currTop
pRegion.NotifyResizedY
setRegionViewSizeAndLocation = pRegion.ActualHeight

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setRegionViewSizes()
Dim lregion As ChartRegion
Dim currTop As Long

' Now actually set the Heights and positions for the picture boxes

Const ProcName As String = "setRegionViewSizes"

On Error GoTo Err

If Not IsDrawingEnabled Then Exit Sub

For Each lregion In mRegionMap
    currTop = currTop + setRegionDividerLocation(lregion, currTop)
    currTop = currTop + setRegionViewSizeAndLocation(lregion, currTop)
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupBackgroundViewport()
Dim lcanvas As New Canvas
Const ProcName As String = "setupBackgroundViewport"

On Error GoTo Err

lcanvas.Surface = ChartRegionPicture(0)
lcanvas.RegionType = RegionTypeBackground
lcanvas.MousePointer = vbDefault
Set mBackGroundViewport = New Viewport
mBackGroundViewport.Canvas = lcanvas
mBackGroundViewport.RegionType = RegionTypeBackground

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub SuppressDrawing(ByVal suppress As Boolean)
Dim Region As ChartRegion
Const ProcName As String = "SuppressDrawing"

On Error GoTo Err

If suppress Then
    mSuppressDrawingCount = mSuppressDrawingCount + 1
Else
    If mSuppressDrawingCount > 0 Then
        mSuppressDrawingCount = mSuppressDrawingCount - 1
    End If
End If

If mSuppressDrawingCount = 0 Then
    Resize True, True
End If

For Each Region In mRegions
    Region.IsDrawingEnabled = IsDrawingEnabled
Next
If Not mXAxisRegion Is Nothing Then mXAxisRegion.IsDrawingEnabled = IsDrawingEnabled

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub unmapRegion( _
                    ByVal Region As ChartRegion)
Const ProcName As String = "unmapRegion"

On Error GoTo Err

mRegionMap.Remove CLng(ChartRegionPicture(Region.handle).Tag)
Unload ChartRegionPicture(Region.handle)
Unload RegionDividerPicture(Region.handle)
Unload YRegionDividerPicture(Region.handle)
Unload YAxisPicture(Region.handle)
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

