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
      Picture         =   "ChartArea.ctx":0152
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
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   0
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
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
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
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
      MouseIcon       =   "ChartArea.ctx":0594
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
Event PeriodsChanged(ev As CollectionChangeEvent)
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
Private Const PropNameTwipsPerBar                       As String = "TwipsPerBar"
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
Private Const PropDfltTwipsPerBar                       As Long = 150
Private Const PropDfltVerticalGridSpacing               As Long = 1
Private Const PropDfltVerticalGridUnits                 As Long = TimePeriodHour
Private Const PropDfltXAxisVisible                      As Boolean = True
Private Const PropDfltYAxisVisible                      As Boolean = True
Private Const PropDfltYAxisWidthCm                      As Single = 1.3

'================================================================================
' Member variables
'================================================================================

Private WithEvents mRegions                             As ChartRegions
Attribute mRegions.VB_VarHelpID = -1
Private mRegionMap                                      As ChartRegionMap

Private WithEvents mPeriods                             As Periods
Attribute mPeriods.VB_VarHelpID = -1

Private mController                                     As ChartController

Private mInitialised                                    As Boolean

Private mScaleWidth As Single
Private mScaleHeight As Single
Private mScaleLeft As Single
Private mScaleTop As Single

Private mPrevHeight As Single
Private mPrevWidth As Single

Private mTwipsPerBar As Long

Private mXAxisVisible  As Boolean
Private mXAxisRegion As ChartRegion
Private mXCursorText As Text

Private mYAxisPosition As Long
Private mYAxisWidthCm As Single
Private mYAxisVisible  As Boolean

Private mSessionStartTime As Date
Private mSessionEndTime As Date

Private mCurrentSessionStartTime As Date
Private mCurrentSessionEndTime As Date

Private mBarTimePeriod As TimePeriod
Private mBarTimePeriodSet As Boolean

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
Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mSuppressDrawingCount As Long

Private mLeftDragStartPosnX As Long
Private mLeftDragStartPosnY As Single

Private mUserResizingRegions As Boolean

Private mHorizontalMouseScrollingAllowed As Boolean
Private mVerticalMouseScrollingAllowed As Boolean

Private mMouseScrollingInProgress As Boolean

Private mHorizontalScrollBarVisible As Boolean

Private mReferenceTime As Date

Private mAutoscrolling As Boolean

Private mBackGroundViewport As Viewport
Private mChartBackColor As Long

'================================================================================
' User Control Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
Dim failpoint As Long
On Error GoTo Err

Set mRegions = New ChartRegions
mRegions.Chart = Me

Set mRegionMap = New ChartRegionMap

Set gBlankCursor = BlankPicture.Picture
Set gSelectorCursor = SelectorPicture.Picture

ReDim mChartBackGradientFillColors(0) As Long
mChartBackGradientFillColors(0) = PropDfltChartBackColor

Set mController = New ChartController
mController.Chart = Me

mTwipsPerBar = PropDfltTwipsPerBar

createXAxisRegion

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "UserControl_Initialize" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

Initialise

HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
Autoscrolling = PropDfltAutoscrolling
Set mBarTimePeriod = GetTimePeriod(PropDfltPeriodLength, PropDfltPeriodUnits)
PointerCrosshairsColor = PropDfltPointerCrosshairsColor
PointerDiscColor = PropDfltPointerDiscColor
PointerStyle = PropDfltPointerStyle
HorizontalScrollBarVisible = PropDfltHorizontalScrollBarVisible
TwipsPerBar = PropDfltTwipsPerBar
Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
XAxisVisible = PropDfltXAxisVisible
YAxisWidthCm = PropDfltYAxisWidthCm
YAxisVisible = PropDfltYAxisVisible

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
mController.fireKeyDown KeyCode, Shift
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
mController.fireKeyPress KeyAscii
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
mController.fireKeyUp KeyCode, Shift
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

Initialise

Autoscrolling = PropBag.ReadProperty(PropNameAutoscrolling, PropDfltAutoscrolling)

Set mBarTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNamePeriodLength, PropDfltPeriodLength), _
                                    PropBag.ReadProperty(PropNamePeriodUnits, PropDfltPeriodUnits))


ChartBackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)

HorizontalMouseScrollingAllowed = PropBag.ReadProperty(PropNameHorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed)

HorizontalScrollBarVisible = PropBag.ReadProperty(PropNameHorizontalScrollBarVisible, PropDfltHorizontalScrollBarVisible)


PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)

PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)

PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)

TwipsPerBar = PropBag.ReadProperty(PropNameTwipsPerBar, PropDfltTwipsPerBar)

Set mVerticalGridTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNameVerticalGridSpacing, PropDfltVerticalGridSpacing), _
                        PropBag.ReadProperty(PropNameVerticalGridUnits, PropDfltVerticalGridUnits))

VerticalMouseScrollingAllowed = PropBag.ReadProperty(PropNameVerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed)

XAxisVisible = PropBag.ReadProperty(PropNameXAxisVisible, PropDfltXAxisVisible)

YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)

YAxisVisible = PropBag.ReadProperty(PropNameYAxisVisible, PropDfltYAxisVisible)

End Sub

Private Sub UserControl_Resize()

Dim failpoint As Long
On Error GoTo Err

gTracer.EnterProcedure pInfo:="width=" & UserControl.Width & "; height=" & UserControl.Height, pProcedureName:="UserControl_Resize", pProjectName:=ProjectName, pModuleName:=ModuleName

If UserControl.Width <> 0 And UserControl.Height <> 0 Then Resize (UserControl.Width <> mPrevWidth), (UserControl.Height <> mPrevHeight)
mPrevHeight = UserControl.Height
mPrevWidth = UserControl.Width

gTracer.ExitProcedure pInfo:="", pProcedureName:="UserControl_Resize", pProjectName:=ProjectName, pModuleName:=ModuleName
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "UserControl_Resize" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

Private Sub UserControl_Terminate()
gLogger.Log LogLevelDetail, "ChartSkil chart terminated"
Debug.Print "ChartSkil chart terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameHorizontalMouseScrollingAllowed, HorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed
PropBag.WriteProperty PropNameVerticalMouseScrollingAllowed, VerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed
PropBag.WriteProperty PropNameAutoscrolling, Autoscrolling, PropDfltAutoscrolling
PropBag.WriteProperty PropNameChartBackColor, mChartBackColor
PropBag.WriteProperty PropNamePeriodLength, mBarTimePeriod.Length, PropDfltPeriodLength
PropBag.WriteProperty PropNamePeriodUnits, mBarTimePeriod.Units, PropDfltPeriodUnits
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNamePointerStyle, mPointerStyle, PropDfltPointerStyle
PropBag.WriteProperty PropNameHorizontalScrollBarVisible, HorizontalScrollBarVisible, PropDfltHorizontalScrollBarVisible
PropBag.WriteProperty PropNameTwipsPerBar, TwipsPerBar, PropDfltTwipsPerBar
PropBag.WriteProperty PropNameVerticalGridSpacing, mVerticalGridTimePeriod.Length, PropDfltVerticalGridSpacing
PropBag.WriteProperty PropNameVerticalGridUnits, mVerticalGridTimePeriod.Units, PropDfltVerticalGridUnits
PropBag.WriteProperty PropNameXAxisVisible, XAxisVisible, PropDfltXAxisVisible
PropBag.WriteProperty PropNameYAxisVisible, YAxisVisible, PropDfltYAxisVisible
PropBag.WriteProperty PropNameYAxisWidthCm, YAxisWidthCm, PropDfltYAxisWidthCm
End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_Click(index As Integer)
Dim failpoint As Long
On Error GoTo Err

If index = 0 Then Exit Sub

getDataRegionFromPictureIndex(index).Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_Click" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_DblClick(index As Integer)
Dim failpoint As Long
On Error GoTo Err

If index = 0 Then Exit Sub

getDataRegionFromPictureIndex(index).DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_DblClick" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim Region As ChartRegion

Dim failpoint As Long

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
    (mHorizontalMouseScrollingAllowed Or mVerticalMouseScrollingAllowed) _
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseDown" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseMove(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
Dim lRegion As ChartRegion

Dim failpoint As Long
On Error GoTo Err

If index = 0 Then Exit Sub

Set lRegion = getDataRegionFromPictureIndex(index)

If CBool(Button And MouseButtonConstants.vbLeftButton) Then
    If mPointerMode = PointerModeDefault And _
        (mHorizontalMouseScrollingAllowed Or mVerticalMouseScrollingAllowed) And _
        mMouseScrollingInProgress _
    Then
        mouseScroll lRegion, X, Y
    Else
        mMouseScrollingInProgress = False
        MouseMove lRegion, Button, Shift, X, Y
    End If
Else
    MouseMove lRegion, Button, Shift, X, Y
End If

lRegion.MouseMove Button, Shift, Round(X), Y

RaiseEvent MouseMove(Button, _
                    Shift, _
                    convertChartRegionPictureMouseXtoContainerCoords(index, X), _
                    convertChartRegionPictureMouseYtoContainerCoords(index, Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseMove" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim Region As ChartRegion

Dim failpoint As Long
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

Region.MouseUp Button, Shift, Round(X), Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertChartRegionPictureMouseXtoContainerCoords(index, X), _
                    convertChartRegionPictureMouseYtoContainerCoords(index, Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseUp" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' HScroll Event Handlers
'================================================================================

Private Sub HScroll_Change()
Dim failpoint As Long
On Error GoTo Err

LastVisiblePeriod = Round((CLng(HScroll.value) - CLng(HScroll.Min)) / (CLng(HScroll.Max) - CLng(HScroll.Min)) * (mPeriods.CurrentPeriodNumber + ChartWidth - 1))

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "HScroll_Change" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
Dim failpoint As Long
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseDown" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub RegionDividerPicture_MouseMove( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim failpoint As Long
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseMove" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub RegionDividerPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim failpoint As Long
On Error GoTo Err

mUserResizingRegions = False

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertRegionDividerPictureMouseXtoContainerCoords(index, X), _
                    convertRegionDividerPictureMouseYtoContainerCoords(index, Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseUp" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' XAxisPicture Event Handlers
'================================================================================

Private Sub XAxisPicture_Click()
Dim failpoint As Long
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "XAxisPicture_Click" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_DblClick()
Dim failpoint As Long
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "XAxisPicture_DblClick" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.MouseDown Button, Shift, X, Y

RaiseEvent MouseDown(Button, _
                    Shift, _
                    convertXAxisPictureMouseXtoContainerCoords(X), _
                    convertXAxisPictureMouseYtoContainerCoords(Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseDown" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.MouseMove Button, Shift, X, Y

RaiseEvent MouseMove(Button, _
                    Shift, _
                    convertXAxisPictureMouseXtoContainerCoords(X), _
                    convertXAxisPictureMouseYtoContainerCoords(Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseMove" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

If mXAxisRegion Is Nothing Then Exit Sub
mXAxisRegion.MouseUp Button, Shift, X, Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertXAxisPictureMouseXtoContainerCoords(X), _
                    convertXAxisPictureMouseYtoContainerCoords(Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseUp" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' YAxisPicture Event Handlers
'================================================================================

Private Sub YAxisPicture_Click(index As Integer)
Dim failpoint As Long
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "YAxisPicture_Click" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_DblClick(index As Integer)
Dim failpoint As Long
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "YAxisPicture_DblClick" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).MouseDown Button, Shift, X, Y

RaiseEvent MouseDown(Button, _
                    Shift, _
                    convertYAxisPictureMouseXtoContainerCoords(index, X), _
                    convertYAxisPictureMouseYtoContainerCoords(index, Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseDown" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).MouseMove Button, Shift, X, Y

RaiseEvent MouseMove(Button, _
                    Shift, _
                    convertYAxisPictureMouseXtoContainerCoords(index, X), _
                    convertYAxisPictureMouseYtoContainerCoords(index, Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseMove" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

getYAxisRegionFromPictureIndex(index).MouseUp Button, Shift, X, Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    convertYAxisPictureMouseXtoContainerCoords(index, X), _
                    convertYAxisPictureMouseYtoContainerCoords(index, Y))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseUp" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' mRegions Event Handlers
'================================================================================

Private Sub mRegions_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
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
End Sub

'================================================================================
' mPeriods Event Handlers
'================================================================================

Private Sub mPeriods_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
Dim Region As ChartRegion
Dim Period As Period

Dim failpoint As Long
On Error GoTo Err

Set Period = ev.affectedItem

Select Case ev.changeType
Case CollItemAdded
    For Each Region In mRegions
        Region.AddPeriod Period.PeriodNumber, Period.Timestamp
    Next
    
    mXAxisRegion.AddPeriod Period.PeriodNumber, Period.Timestamp
    If IsDrawingEnabled Then setHorizontalScrollBar
    setSession Period.Timestamp
    If mAutoscrolling Then ScrollX 1
    
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mPeriods_PeriodAdded" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Get availableheight() As Long
availableheight = IIf(mXAxisVisible, XAxisPicture.Top, UserControl.ScaleHeight) - _
                    (mRegions.Count - 1) * RegionDividerPicture(0).Height
If availableheight < 1 Then availableheight = 1
End Property

Public Property Let BarTimePeriod( _
                ByVal value As TimePeriod)
If mBarTimePeriodSet Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "barTimePeriod", _
                                    "BarTimePeriod has already been set"
If value.Length < 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "barTimePeriod", _
                                    "BarTimePeriod length cannot be negative"
                                    
Select Case value.Units
Case TimePeriodNone
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case TimePeriodWeek
Case TimePeriodMonth
Case TimePeriodYear
Case TimePeriodVolume
Case TimePeriodTickVolume
Case TimePeriodTickMovement
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil" & "." & "Chart" & ":" & "setPeriodParameters", _
            "Invalid period unit - must be a member of the TimePeriodUnits enum"
End Select

Set mBarTimePeriod = value

mBarTimePeriodSet = True

If Not mVerticalGridTimePeriodSet Then calcVerticalGridParams
setRegionPeriodAndVerticalGridParameters

End Property

Public Property Get BarTimePeriod() As TimePeriod
Attribute BarTimePeriod.VB_MemberFlags = "400"
Set BarTimePeriod = mBarTimePeriod
End Property

Public Property Get Autoscrolling() As Boolean
Autoscrolling = mAutoscrolling
End Property

Public Property Let Autoscrolling(ByVal value As Boolean)
mAutoscrolling = value
PropertyChanged PropNameAutoscrolling
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Attribute ChartBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
ChartBackColor = mChartBackColor
End Property

Public Property Let ChartBackColor(ByVal value As OLE_COLOR)
mChartBackColor = value
resizeBackground
PropertyChanged PropNameChartBackColor
End Property

Public Property Get ChartLeft() As Double
Attribute ChartLeft.VB_MemberFlags = "400"
ChartLeft = mScaleLeft
End Property

Public Property Get ChartWidth() As Double
Attribute ChartWidth.VB_MemberFlags = "400"
ChartWidth = YAxisPosition - mScaleLeft
End Property

Public Property Get Controller() As ChartController
Set Controller = mController
End Property

Public Property Get CurrentPeriodNumber() As Long
Attribute CurrentPeriodNumber.VB_MemberFlags = "400"
CurrentPeriodNumber = mPeriods.CurrentPeriodNumber
End Property

Public Property Get CurrentSessionEndTime() As Date
Attribute CurrentSessionEndTime.VB_MemberFlags = "400"
CurrentSessionEndTime = mCurrentSessionEndTime
End Property

Public Property Get CurrentSessionStartTime() As Date
Attribute CurrentSessionStartTime.VB_MemberFlags = "400"
CurrentSessionStartTime = mCurrentSessionStartTime
End Property

Public Property Get FirstVisiblePeriod() As Long
Attribute FirstVisiblePeriod.VB_MemberFlags = "400"
FirstVisiblePeriod = mScaleLeft
End Property

Public Property Let FirstVisiblePeriod(ByVal value As Long)
ScrollX value - mScaleLeft + 1
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
HorizontalMouseScrollingAllowed = mHorizontalMouseScrollingAllowed
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal value As Boolean)
mHorizontalMouseScrollingAllowed = value
PropertyChanged PropNameHorizontalMouseScrollingAllowed
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
HorizontalScrollBarVisible = mHorizontalScrollBarVisible
PropertyChanged PropNameHorizontalScrollBarVisible
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
mHorizontalScrollBarVisible = val
If mHorizontalScrollBarVisible Then
    HScroll.Visible = True
Else
    HScroll.Visible = False
End If
Resize False, True
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

Public Property Let LastVisiblePeriod(ByVal value As Long)
ScrollX value - mYAxisPosition + 1
End Property

Public Property Get Periods() As Periods
Attribute Periods.VB_MemberFlags = "400"
Set Periods = mPeriods
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
Attribute PointerCrosshairsColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerCrosshairsColor = mPointerCrosshairsColor
End Property

Public Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Dim Region As ChartRegion
mPointerCrosshairsColor = value
For Each Region In mRegions
    Region.PointerCrosshairsColor = value
Next
PropertyChanged PropNamePointerCrosshairsColor
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Attribute PointerDiscColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerDiscColor = mPointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Dim Region As ChartRegion
mPointerDiscColor = value
For Each Region In mRegions
    Region.PointerDiscColor = value
Next
PropertyChanged PropNamePointerDiscColor
End Property

Public Property Get PointerIcon() As IPictureDisp
Attribute PointerIcon.VB_MemberFlags = "400"
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal value As IPictureDisp)
Dim Region As ChartRegion

If value Is Nothing Then Exit Property
If value Is mPointerIcon Then Exit Property

Set mPointerIcon = value

If mPointerStyle = PointerCustom Then
    For Each Region In mRegions
        Region.PointerStyle = PointerCustom
    Next
End If
End Property

Public Property Get PointerMode() As PointerModes
Attribute PointerMode.VB_MemberFlags = "400"
PointerMode = mPointerMode
End Property

Public Property Get PointerStyle() As PointerStyles
Attribute PointerStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Dim Region As ChartRegion

If value = mPointerStyle Then Exit Property

mPointerStyle = value

If mPointerStyle = PointerCustom And mPointerIcon Is Nothing Then
    ' we'll notify the region when an icon is supplied
    Exit Property
End If

For Each Region In mRegions
    If mPointerStyle = PointerCustom Then Region.PointerIcon = mPointerIcon
    Region.PointerStyle = mPointerStyle
Next
PropertyChanged PropNamePointerStyle
End Property

Public Property Get Regions() As ChartRegions
Set Regions = mRegions
End Property

Public Property Get SessionEndTime() As Date
Attribute SessionEndTime.VB_MemberFlags = "400"
SessionEndTime = mSessionEndTime
End Property

Public Property Let SessionEndTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil26.Chart::(Let)sessionEndTime", _
                "Value must be a time only"
mSessionEndTime = val
End Property

Public Property Get SessionStartTime() As Date
Attribute SessionStartTime.VB_MemberFlags = "400"
SessionStartTime = mSessionStartTime
End Property

Public Property Let SessionStartTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil26.Chart::(Let)sessionStartTime", _
                "Value must be a time only"
mSessionStartTime = val
End Property

Public Property Let TwipsPerBar(ByVal val As Long)
Attribute TwipsPerBar.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
mTwipsPerBar = val
resizeX
PropertyChanged PropNameTwipsPerBar
End Property

Public Property Set VerticalGridTimePeriod( _
                ByVal value As TimePeriod)
If mVerticalGridTimePeriodSet Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "verticalGridTimePeriod", _
                                    "verticalGridTimePeriod has already been set"

If value.Length <= 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "verticalGridTimePeriod", _
                                    "verticalGridTimePeriod length must be >0"
Select Case value.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case TimePeriodWeek
Case TimePeriodMonth
Case TimePeriodYear
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil" & "." & "Chart" & ":" & "verticalGridTimePeriod", _
                "verticalGridTimePeriod Units must be a member of the TimePeriodUnits enum"
End Select

Set mVerticalGridTimePeriod = value
mVerticalGridTimePeriodSet = True

setRegionPeriodAndVerticalGridParameters

End Property

Public Property Get VerticalGridTimePeriod() As TimePeriod
Attribute VerticalGridTimePeriod.VB_MemberFlags = "400"
Set VerticalGridTimePeriod = mVerticalGridTimePeriod
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
VerticalMouseScrollingAllowed = mVerticalMouseScrollingAllowed
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal value As Boolean)
mVerticalMouseScrollingAllowed = value
PropertyChanged PropNameVerticalMouseScrollingAllowed
End Property

Public Property Get XAxisRegion() As ChartRegion
Attribute XAxisRegion.VB_MemberFlags = "400"
Set XAxisRegion = mXAxisRegion
End Property

Public Property Get XAxisVisible() As Boolean
Attribute XAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
XAxisVisible = mXAxisVisible
End Property

Public Property Let XAxisVisible(ByVal value As Boolean)
mXAxisVisible = value
mRegions.ResizeY mUserResizingRegions
XAxisPicture.Visible = mXAxisVisible
PropertyChanged PropNameXAxisVisible
End Property

Public Property Let XCursorTextStyle(ByVal value As TextStyle)
mXCursorText.LocalStyle = value
End Property

Public Property Get XCursorTextStyle() As TextStyle
Set XCursorTextStyle = mXCursorText.LocalStyle
End Property

Public Property Get YAxisPosition() As Long
Attribute YAxisPosition.VB_MemberFlags = "400"
YAxisPosition = mYAxisPosition
End Property

Public Property Get YAxisVisible() As Boolean
Attribute YAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
YAxisVisible = mYAxisVisible
End Property

Public Property Let YAxisVisible(ByVal value As Boolean)
mYAxisVisible = value
resizeX
PropertyChanged PropNameYAxisVisible
End Property

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_ProcData.VB_Invoke_Property = ";Appearance"
YAxisWidthCm = mYAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
If value <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "YAxisWidthCm", _
            "Y axis Width must be greater than 0"
End If

mYAxisWidthCm = value
resizeX
PropertyChanged PropNameYAxisWidthCm
End Property

'================================================================================
' Methods
'================================================================================

Public Function ClearChart()

DisableDrawing

Clear

Initialise
mYAxisPosition = 1
createXAxisRegion

EnableDrawing

RaiseEvent ChartCleared
mController.fireChartCleared
Debug.Print "Chart cleared"
End Function

Friend Function CreateDataRegionCanvas(ByVal index As Long) As Canvas
Load ChartRegionPicture(index)
Set CreateDataRegionCanvas = CreateCanvas(ChartRegionPicture(index), RegionTypeData)
End Function

Friend Function CreateYAxisRegionCanvas(ByVal index As Long) As Canvas
Load YAxisPicture(index)
Set CreateYAxisRegionCanvas = CreateCanvas(YAxisPicture(index), RegionTypeYAxis)
End Function

Public Sub DisableDrawing()
SuppressDrawing True
End Sub

Public Sub EnableDrawing()
SuppressDrawing False
End Sub

Public Sub Finish()
DisableDrawing
Clear
mController.Finished
mRegions.Finish
End Sub

Public Function GetXFromTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal forceNewPeriod As Boolean, _
                Optional ByVal duplicateNumber As Long) As Double
Dim lPeriod As Period
Dim periodEndtime As Date

Select Case BarTimePeriod.Units
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
    On Error GoTo 0
    
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
                            BarTimePeriod, _
                            SessionStartTime)
    GetXFromTimestamp = lPeriod.PeriodNumber + (Timestamp - lPeriod.Timestamp) / (periodEndtime - lPeriod.Timestamp)
    
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If Not forceNewPeriod Then
        On Error Resume Next
        Set lPeriod = mPeriods.ItemDup(Timestamp, duplicateNumber)
        On Error GoTo 0
        
        If lPeriod Is Nothing Then
            Set lPeriod = mPeriods.AddPeriod(Timestamp, True)
        End If
        GetXFromTimestamp = lPeriod.PeriodNumber
    Else
        Set lPeriod = mPeriods.AddPeriod(Timestamp, True)
        GetXFromTimestamp = lPeriod.PeriodNumber
    End If
End Select

End Function

Public Sub HideGrid()
Dim Region As ChartRegion

If mHideGrid Then Exit Sub

mHideGrid = True
For Each Region In mRegions
    Region.HideGrid
Next
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean

If Timestamp >= mCurrentSessionStartTime And _
    Timestamp < mCurrentSessionEndTime _
Then
    IsTimeInSession = True
End If
End Function

Public Sub ScrollX(ByVal value As Long)
Dim Region As ChartRegion
Dim failpoint As Long
On Error GoTo Err

gTracer.EnterProcedure pInfo:="value=" & CStr(value), pProcedureName:="ScrollX", pProjectName:=ProjectName, pModuleName:=ModuleName

If value = 0 Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="ScrollX", pProjectName:=ProjectName, pModuleName:=ModuleName
    Exit Sub
End If

If (LastVisiblePeriod + value) > _
        (mPeriods.CurrentPeriodNumber + ChartWidth - 1) Then
    value = mPeriods.CurrentPeriodNumber + ChartWidth - 1 - LastVisiblePeriod
ElseIf (LastVisiblePeriod + value) < 1 Then
    value = 1 - LastVisiblePeriod
End If

mYAxisPosition = mYAxisPosition + value
mScaleLeft = calcScaleLeft
XAxisPicture.ScaleLeft = mScaleLeft

If Not IsDrawingEnabled Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="ScrollX", pProjectName:=ProjectName, pModuleName:=ModuleName
    Exit Sub
End If

For Each Region In mRegions
    Region.SetPeriodsInView mScaleLeft, mYAxisPosition - 1
Next

mXAxisRegion.SetPeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
setHorizontalScrollBar

gTracer.ExitProcedure pInfo:="", pProcedureName:="ScrollX", pProjectName:=ProjectName, pModuleName:=ModuleName
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ScrollX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetPointerModeDefault()
Dim Region As ChartRegion
Dim failpoint As Long
On Error GoTo Err

mPointerMode = PointerModeDefault
For Each Region In mRegions
    Region.SetPointerModeDefault
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetPointerModeDefault" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetPointerModeSelection()
Dim Region As ChartRegion

Dim failpoint As Long
On Error GoTo Err

mPointerMode = PointerModeSelection

For Each Region In mRegions
    Region.SetPointerModeSelection
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetPointerModeSelection" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub SetPointerModeTool( _
                Optional ByVal toolPointerStyle As PointerStyles = PointerTool, _
                Optional ByVal icon As IPictureDisp)
Dim Region As ChartRegion
Dim failpoint As Long
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
            ProjectName & "." & ModuleName & ":" & "SetPointerModeTool", _
            "toolPointerStyle must be a member of the PointerStyles enum"
End Select
For Each Region In mRegions
    Region.SetPointerModeTool toolPointerStyle, icon
Next

RaiseEvent PointerModeChanged
mController.firePointerModeChanged
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetPointerModeTool" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub ShowGrid()
Dim Region As ChartRegion

If Not mHideGrid Then Exit Sub

mHideGrid = False
For Each Region In mRegions
    Region.ShowGrid
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcScaleLeft() As Single
calcScaleLeft = mYAxisPosition + _
            IIf(mYAxisVisible, mYAxisWidthCm * TwipsPerCm / XAxisPicture.Width * mScaleWidth, 0) - _
            mScaleWidth
End Function

Private Sub CalcSessionTimes(ByVal Timestamp As Date, _
                            ByRef SessionStartTime As Date, _
                            ByRef SessionEndTime As Date)
Dim i As Long

i = -1
Do
    i = i + 1
Loop Until calcSessionTimesHelper(Timestamp + i, SessionStartTime, SessionEndTime)
End Sub

Friend Function calcSessionTimesHelper(ByVal Timestamp As Date, _
                            ByRef SessionStartTime As Date, _
                            ByRef SessionEndTime As Date) As Boolean
Dim referenceDate As Date
Dim referenceTime As Date
Dim weekday As Long

referenceDate = DateValue(Timestamp)
referenceTime = TimeValue(Timestamp)

If mSessionStartTime < mSessionEndTime Then
    ' session doesn't span midnight
    If referenceTime < mSessionEndTime Then
        SessionStartTime = referenceDate + mSessionStartTime
        SessionEndTime = referenceDate + mSessionEndTime
    Else
        SessionStartTime = referenceDate + 1 + mSessionStartTime
        SessionEndTime = referenceDate + 1 + mSessionEndTime
    End If
ElseIf mSessionStartTime > mSessionEndTime Then
    ' session spans midnight
    If referenceTime >= mSessionEndTime Then
        SessionStartTime = referenceDate + mSessionStartTime
        SessionEndTime = referenceDate + 1 + mSessionEndTime
    Else
        SessionStartTime = referenceDate - 1 + mSessionStartTime
        SessionEndTime = referenceDate + mSessionEndTime
    End If
Else
    ' this instrument trades 24hrs, or the contract service provider doesn't know
    ' the session start and end times
    SessionStartTime = referenceDate
    SessionEndTime = referenceDate + 1
End If

weekday = DatePart("w", SessionStartTime)
If mSessionStartTime < mSessionEndTime Then
    ' session doesn't span midnight
    If weekday <> vbSaturday And weekday <> vbSunday Then calcSessionTimesHelper = True
ElseIf mSessionStartTime > mSessionEndTime Then
    ' session DOES span midnight
    If weekday <> vbFriday And weekday <> vbSaturday Then calcSessionTimesHelper = True
Else
    ' 24-hour session or no session times known
    If weekday <> vbSaturday And weekday <> vbSunday Then calcSessionTimesHelper = True
End If
End Function

Private Sub calcVerticalGridParams()

Select Case mBarTimePeriod.Units
Case TimePeriodNone
    Set mVerticalGridTimePeriod = Nothing
Case TimePeriodSecond
    Select Case mBarTimePeriod.Length
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
    Select Case mBarTimePeriod.Length
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
    Case Else
        Set mVerticalGridTimePeriod = Nothing
    End Select
Case TimePeriodHour
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodDay)
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
  
End Sub

Private Sub Clear()
Dim lRegion As ChartRegion
Dim en As Enumerator

Set en = mRegions.Enumerator

Do While en.moveNext
    Set lRegion = en.current
    lRegion.ClearRegion
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
End Sub

Private Function convertChartRegionPictureMouseXtoContainerCoords( _
                ByVal index As Long, _
                ByVal X As Single) As Single
convertChartRegionPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(ChartRegionPicture(index), X)
End Function

Private Function convertChartRegionPictureMouseYtoContainerCoords( _
                ByVal index As Long, _
                ByVal Y As Single) As Single
convertChartRegionPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(ChartRegionPicture(index), Y)
End Function

Private Function convertPictureMouseXtoContainerCoords( _
                ByVal pPicture As PictureBox, _
                ByVal X As Single) As Single
convertPictureMouseXtoContainerCoords = _
    ScaleX(pPicture.Left + _
            pPicture.ScaleX(X - pPicture.ScaleLeft, _
                            pPicture.ScaleMode, _
                            vbTwips), _
            vbTwips, _
            vbContainerPosition)
End Function

Private Function convertPictureMouseYtoContainerCoords( _
                ByVal pPicture As PictureBox, _
                ByVal Y As Single) As Single
convertPictureMouseYtoContainerCoords = _
    ScaleY(pPicture.Top + _
            pPicture.ScaleY(Y - pPicture.ScaleTop, _
                            pPicture.ScaleMode, _
                            vbTwips), _
            vbTwips, _
            vbContainerPosition)
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
convertRegionDividerPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(RegionDividerPicture(index), Y)
End Function

Private Function convertXAxisPictureMouseXtoContainerCoords( _
                ByVal X As Single) As Single
convertXAxisPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(XAxisPicture, X)
End Function

Private Function convertXAxisPictureMouseYtoContainerCoords( _
                ByVal Y As Single) As Single
convertXAxisPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(XAxisPicture, Y)
End Function

Private Function convertYAxisPictureMouseXtoContainerCoords( _
                ByVal index As Long, _
                ByVal X As Single) As Single
convertYAxisPictureMouseXtoContainerCoords = _
    convertPictureMouseXtoContainerCoords(YAxisPicture(index), X)
End Function

Private Function convertYAxisPictureMouseYtoContainerCoords( _
                ByVal index As Long, _
                ByVal Y As Single) As Single
convertYAxisPictureMouseYtoContainerCoords = _
    convertPictureMouseYtoContainerCoords(YAxisPicture(index), Y)
End Function

Private Function CreateCanvas( _
                ByVal Surface As PictureBox, _
                ByVal pRegionType As RegionTypes) As Canvas
Set CreateCanvas = New Canvas
CreateCanvas.Surface = Surface
CreateCanvas.RegionType = pRegionType
End Function

Private Sub createXAxisRegion()
Dim aFont As StdFont

Set mXAxisRegion = New ChartRegion

mXAxisRegion.Initialise "", Me, CreateCanvas(XAxisPicture, RegionTypeXAxis), RegionTypeXAxis
                        
mXAxisRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
mXAxisRegion.Bottom = 0
mXAxisRegion.Top = 1
mXAxisRegion.SessionStartTime = mSessionStartTime
mXAxisRegion.HasGrid = False
mXAxisRegion.HasGridText = True

Set mXCursorText = mXAxisRegion.AddText(LayerNumbers.LayerPointer)
mXCursorText.Align = AlignBoxTopCentre

Dim txtStyle As New TextStyle
txtStyle.Color = vbBlack
txtStyle.Box = True
txtStyle.BoxFillColor = vbWhite
txtStyle.BoxStyle = LineSolid
txtStyle.BoxColor = vbBlack
txtStyle.PaddingX = 1
Set aFont = New StdFont
aFont.Name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
txtStyle.Font = aFont
mXCursorText.LocalStyle = txtStyle
End Sub

Private Sub displayXAxisLabel(ByVal X As Single)
Dim thisPeriod As Period
Dim PeriodNumber As Long

If Round(X) >= mYAxisPosition Then Exit Sub
If mPeriods.Count = 0 Then Exit Sub

On Error Resume Next
PeriodNumber = Round(X)
Set thisPeriod = mPeriods(PeriodNumber)
On Error GoTo 0
If thisPeriod Is Nothing Then
    mXCursorText.Text = ""
    Exit Sub
End If

mXCursorText.position = mXAxisRegion.NewPoint( _
                            PeriodNumber, _
                            0, _
                            CoordsLogical, _
                            CoordsCounterDistance)

Select Case mBarTimePeriod.Units
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

End Sub

Private Sub finishBackgroundCanvas()
gLogger.Log LogLevelHighDetail, "Finish background canvas"
If Not mBackGroundViewport Is Nothing Then mBackGroundViewport.Finish
Set mBackGroundViewport = Nothing
End Sub

Private Function getDataRegionFromPictureIndex( _
                ByVal index As Long) As ChartRegion
Set getDataRegionFromPictureIndex = mRegionMap.Item(CLng(ChartRegionPicture(index).Tag))
End Function

Private Function getYAxisRegionFromPictureIndex( _
                ByVal index As Long) As ChartRegion
Set getYAxisRegionFromPictureIndex = mRegions.ItemFromHandle(CLng(YAxisPicture(index).Tag))
End Function

Private Sub Initialise()
Static firstInitialisationDone As Boolean
Dim btn As Button

mPrevHeight = UserControl.Height

Set mPeriods = New Periods
mPeriods.Chart = Me

setupBackgroundViewport

mBarTimePeriodSet = False
mVerticalGridTimePeriodSet = False

mPointerMode = PointerModes.PointerModeDefault

mYAxisPosition = 1
mScaleWidth = CSng(XAxisPicture.Width) / CSng(mTwipsPerBar) - 0.5!
mScaleLeft = calcScaleLeft
mScaleHeight = -100
mScaleTop = 100

HScroll.value = 0

mInitialised = True

Resize True, True

End Sub

Private Sub mapRegion(pRegion As ChartRegion)
Dim index As Long
Dim mapHandle As Long
Dim btn As Button

If pRegion Is Nothing Then Exit Sub

index = pRegion.handle
mapHandle = mRegionMap.Append(pRegion)

ChartRegionPicture(index).Tag = mapHandle
ChartRegionPicture(index).Visible = True
ChartRegionPicture(index).Align = vbAlignNone
ChartRegionPicture(index).Width = _
    IIf(mYAxisVisible, UserControl.ScaleWidth - mYAxisWidthCm * TwipsPerCm, UserControl.ScaleWidth)
ChartRegionPicture(index).ZOrder 1

pRegion.IsDrawingEnabled = IsDrawingEnabled
pRegion.PointerStyle = mPointerStyle
pRegion.PointerIcon = mPointerIcon
pRegion.PointerCrosshairsColor = mPointerCrosshairsColor
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
pRegion.SessionStartTime = mSessionStartTime

If mHideGrid Then pRegion.HideGrid

Load RegionDividerPicture(index)
RegionDividerPicture(index).Tag = mapHandle
RegionDividerPicture(index).ZOrder 0
RegionDividerPicture(index).Visible = (Not mRegionMap.IsFirst(mapHandle))

YAxisPicture(index).Tag = pRegion.YAxisRegion.handle
YAxisPicture(index).Align = vbAlignNone
YAxisPicture(index).Left = ChartRegionPicture(index).Width
YAxisPicture(index).Width = mYAxisWidthCm * TwipsPerCm
YAxisPicture(index).Visible = mYAxisVisible

XAxisPicture.Visible = mXAxisVisible

End Sub

Private Sub MouseMove( _
                ByVal targetRegion As ChartRegion, _
                ByVal Button As Long, _
                ByVal Shift As Long, _
                ByRef X As Single, _
                ByRef Y As Single)
Dim Region As ChartRegion


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
End Sub

Private Sub mouseScroll( _
                ByVal targetRegion As ChartRegion, _
                ByRef X As Single, _
                ByRef Y As Single)

If mHorizontalMouseScrollingAllowed Then
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
If mVerticalMouseScrollingAllowed Then
    If mLeftDragStartPosnY <> Y Then
        With targetRegion
            If Not .Autoscaling Then
                .ScrollVertical mLeftDragStartPosnY - Y
            End If
        End With
    End If
End If
End Sub

Private Sub Resize( _
    ByVal resizeWidth As Boolean, _
    ByVal resizeHeight As Boolean)
Dim failpoint As Long

On Error GoTo Err

failpoint = 100

'gLogger.Log LogLevelDetail, "ChartSkil: Resize: enter"

If Not mInitialised Then Exit Sub

resizeBackground

If resizeWidth Then
    HScroll.Width = UserControl.Width
    XAxisPicture.Width = UserControl.Width
    resizeX
End If

failpoint = 200

If resizeHeight Then
    HScroll.Top = UserControl.Height - IIf(mHorizontalScrollBarVisible, HScroll.Height, 0)
    XAxisPicture.Top = HScroll.Top - IIf(mXAxisVisible, XAxisPicture.Height, 0)
    mRegions.ResizeY mUserResizingRegions
    setRegionViewSizes
End If

'gLogger.Log LogLevelDetail, "ChartSkil: Resize: exit"

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Resize" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Sub resizeBackground()
If mRegions.Count > 0 Then Exit Sub
XAxisPicture.Visible = False
ChartRegionPicture(0).Visible = False
ChartRegionPicture(0).Move 0, 0, UserControl.Width, UserControl.Height
mBackGroundViewport.BackColor = mChartBackColor
mBackGroundViewport.Left = 0
mBackGroundViewport.Right = 1
mBackGroundViewport.Bottom = 0
mBackGroundViewport.Top = 1
mBackGroundViewport.PaintBackground
mBackGroundViewport.Canvas.ZOrder 1
ChartRegionPicture(0).Visible = True
End Sub

Private Sub resizeX()
Dim Region As ChartRegion

Dim failpoint As Long
On Error GoTo Err


failpoint = 100

If Not mInitialised Then Exit Sub

failpoint = 200

mScaleWidth = CSng(XAxisPicture.Width) / CSng(mTwipsPerBar)
mScaleLeft = calcScaleLeft


failpoint = 400

For Each Region In mRegionMap
    If (UserControl.Width - YAxisPicture(Region.handle).Width) > 0 Then
        YAxisPicture(Region.handle).Left = UserControl.Width - IIf(mYAxisVisible, YAxisPicture(Region.handle).Width, 0)
        ChartRegionPicture(Region.handle).Width = YAxisPicture(Region.handle).Left
    End If
    RegionDividerPicture(Region.handle).Width = UserControl.Width
Next


failpoint = 600

For Each Region In mRegionMap
    Region.SetPeriodsInView mScaleLeft, mYAxisPosition - 1
Next

failpoint = 700

If Not mXAxisRegion Is Nothing Then
    mXAxisRegion.SetPeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
End If


failpoint = 800

setHorizontalScrollBar

Exit Sub

Err:
gErrorLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "resizeX" & "." & failpoint & _
                            IIf(Err.Source <> "", Err.Source & vbCrLf, "") & vbCrLf & _
                            Err.Description
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "resizeX" & "." & failpoint & _
        IIf(Err.Source <> "", Err.Source & vbCrLf, ""), _
        Err.Description

End Sub

Private Sub setHorizontalScrollBar()
Dim failpoint As Long
Dim hscrollVal As Integer
On Error GoTo Err

If mPeriods.CurrentPeriodNumber + ChartWidth - 1 > 32767 Then

    failpoint = 100

    HScroll.Max = 32767
ElseIf mPeriods.CurrentPeriodNumber + ChartWidth - 1 < 1 Then

    failpoint = 200

    HScroll.Max = 1
Else

    failpoint = 300
    
    HScroll.Max = mPeriods.CurrentPeriodNumber + ChartWidth - 1
End If
HScroll.Min = 0


failpoint = 400

' NB the following calculation has to be done using doubles as for very large charts it can cause an overflow using integers
hscrollVal = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.CurrentPeriodNumber + ChartWidth - 1)))
If hscrollVal > HScroll.Max Then
    HScroll.value = HScroll.Max
ElseIf hscrollVal < HScroll.Min Then
    HScroll.value = HScroll.Min
Else
    HScroll.value = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.CurrentPeriodNumber + ChartWidth - 1)))
End If

failpoint = 500

HScroll.SmallChange = 1
If (ChartWidth - 1) < 1 Then
    HScroll.LargeChange = 1
Else
    HScroll.LargeChange = ChartWidth - 1
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "setHorizontalScrollBar" & "." & failpoint & _
                            IIf(Err.Source <> "", Err.Source & vbCrLf, "") & vbCrLf & _
                            errDescription
Err.Raise errNumber, _
        ProjectName & "." & ModuleName & ":" & "setHorizontalScrollBar" & "." & failpoint & _
        IIf(Err.Source <> "", Err.Source & vbCrLf, ""), _
        errDescription

End Sub

Private Function setRegionDividerLocation( _
                ByVal pRegion As ChartRegion, _
                ByVal currTop As Long) As Long
RegionDividerPicture(pRegion.handle).Top = currTop
If mRegionMap.IsFirst(CLng(RegionDividerPicture(pRegion.handle).Tag)) Then
    RegionDividerPicture(pRegion.handle).Visible = False
    setRegionDividerLocation = 0
Else
    RegionDividerPicture(pRegion.handle).Visible = True
    setRegionDividerLocation = RegionDividerPicture(pRegion.handle).Height
End If
End Function

Private Sub setRegionPeriodAndVerticalGridParameters()
Dim Region As ChartRegion
mXAxisRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
For Each Region In mRegions
    Region.VerticalGridTimePeriod = mVerticalGridTimePeriod
Next
End Sub

Private Function setRegionViewSizeAndLocation( _
                ByVal pRegion As ChartRegion, _
                ByVal currTop As Long) As Long
ChartRegionPicture(pRegion.handle).Height = pRegion.ActualHeight
YAxisPicture(pRegion.handle).Height = pRegion.ActualHeight
ChartRegionPicture(pRegion.handle).Top = currTop
YAxisPicture(pRegion.handle).Top = currTop
pRegion.ResizedY
setRegionViewSizeAndLocation = pRegion.ActualHeight
End Function

Private Sub setRegionViewSizes()
Dim lRegion As ChartRegion
Dim currTop As Long

' Now actually set the Heights and positions for the picture boxes

If Not IsDrawingEnabled Then Exit Sub

For Each lRegion In mRegionMap
    currTop = currTop + setRegionDividerLocation(lRegion, currTop)
    currTop = currTop + setRegionViewSizeAndLocation(lRegion, currTop)
Next
End Sub

Private Sub setupBackgroundViewport()
Dim lcanvas As New Canvas
lcanvas.Surface = ChartRegionPicture(0)
lcanvas.RegionType = RegionTypeBackground
lcanvas.MousePointer = vbDefault
Set mBackGroundViewport = New Viewport
mBackGroundViewport.Canvas = lcanvas
mBackGroundViewport.RegionType = RegionTypeBackground
End Sub

Private Sub setSession( _
                ByVal Timestamp As Date)
If Timestamp >= mCurrentSessionEndTime Or _
    Timestamp < mReferenceTime _
Then
    mReferenceTime = Timestamp
    CalcSessionTimes Timestamp, mCurrentSessionStartTime, mCurrentSessionEndTime
End If
End Sub

Private Sub SuppressDrawing(ByVal suppress As Boolean)
Dim Region As ChartRegion
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
End Sub

Public Property Get TwipsPerBar() As Long
TwipsPerBar = mTwipsPerBar
End Property

Private Sub unmapRegion( _
                    ByVal Region As ChartRegion)

Dim failpoint As Long
On Error GoTo Err

mRegionMap.Remove CLng(ChartRegionPicture(Region.handle).Tag)
Unload ChartRegionPicture(Region.handle)
Unload RegionDividerPicture(Region.handle)
Unload YAxisPicture(Region.handle)
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "RemoveChartRegion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

