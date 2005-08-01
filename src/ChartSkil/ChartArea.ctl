VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Chart 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ScaleHeight     =   7575
   ScaleWidth      =   9390
   Begin VB.PictureBox YAxisPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   615
      Index           =   0
      Left            =   8400
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox RegionDividerPicture 
      BorderStyle     =   0  'None
      Height          =   25
      Index           =   0
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   9375
      TabIndex        =   3
      Top             =   6240
      Width           =   9375
   End
   Begin MSComCtl2.FlatScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   2
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin VB.PictureBox XAxisPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9330
      TabIndex        =   1
      Top             =   6960
      Width           =   9390
   End
   Begin VB.PictureBox ChartRegionPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8355
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

'================================================================================
' Types
'================================================================================

Private Type RegionTableEntry
    region              As ChartRegion
    percentheight       As Double
    actualHeight        As Long
    useAvailableSpace   As Boolean
End Type

'================================================================================
' Member variables and constants
'================================================================================

Private Const DefaultTwipsPerBar As Long = 150

Private mRegions() As RegionTableEntry
Private mRegionsIndex As Long

Private mCurrentPeriodNumber As Long

Private mPeriods As Collection

Private mAutoscale As Boolean
Private mScaleWidth As Single
Private mScaleHeight As Single
Private mScaleLeft As Single
Private mScaleTop As Single

Private mPrevHeight As Single

Private mTwipsPerBar As Long

Private mYAxisPosition As Long
Private mYAxisWidthCm As Single

Private mBackColour As Long
Private mGridColour As Long
Private mShowGrid As Boolean
Private mShowCrosshairs As Boolean

Private mNotFirstMouseMove As Boolean
Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mSuppressDrawing As Boolean
Private mPainted As Boolean

Private mCurrentTool As ToolTypes

Private mLeftDragging As Boolean    ' set when the mouse is being dragged with
                                    ' the left button depressed
Private mLeftDragStartPosnX As Long
Private mLeftDragStartPosnY As Single

Private mUserResizingRegions As Boolean

Private mAllowHorizontalMouseScrolling As Boolean
Private mAllowVerticalMouseScrolling As Boolean

Private mShowHorizontalScrollBar As Boolean

'================================================================================
' Enums
'================================================================================

Enum ArrowStyles
    ArrowNone
    ArrowSingleOpen
    ArrowDoubleOpen
    ArrowClosed
    ArrowSingleBar
    ArrowDoubleBar
    ArrowLollipop
    ArrowDiamond
    ArrowBarb
End Enum

Enum FillStyles
    FillSolid = vbFSSolid ' 0 Solid
    FillTransparent = vbFSTransparent ' 1 (Default) Transparent
    FillHorizontalLine = vbHorizontalLine ' 2 Horizontal Line
    FillVerticalLine = vbVerticalLine ' 3 Vertical Line
    FillUpwardDiagonal = vbUpwardDiagonal ' 4 Upward Diagonal
    FillDownwardDiagonal = vbDownwardDiagonal ' 5 Downward Diagonal
    FillCross = vbCross ' 6 Cross
    FillDiagonalCross = vbDiagonalCross ' 7 Diagonal Cross
End Enum

Enum LineStyles
    LineSolid = vbSolid
    LineDash = vbDash
    LineDot = vbDot
    LineDashDot = vbDashDot
    LineDashDotDot = vbDashDotDot
    LineInvisible = vbInvisible
    LineInsideSolid = vbInsideSolid
End Enum

Enum TextAlignModes
    AlignTopLeft
    AlignCentreLeft
    AlignBottomLeft
    AlignTopCentre
    AlignCentreCentre
    AlignBottomCentre
    AlignTopRight
    AlignCentreRight
    AlignBottomRight
End Enum

Enum ToolTypes
    ToolPointer
    ToolLine
    ToolLineExtended
    ToolLineRay
    ToolLineHorizontal
    ToolLineVertical
    ToolFibonacciRetracement
    ToolFibonacciExtension
    ToolFibonacciCircle
    ToolFibonacciTime
    ToolRegressionChannel
    ToolRegressionEnvelope
    ToolText
    ToolPitchfork
    ToolCircle
    ToolRectangle
End Enum

Enum Verticals
    VerticalNot
    VerticalUp
    VerticalDown
End Enum

Enum Quadrants
    NE
    NW
    SW
    SE
End Enum

'================================================================================
' User Control Event Handlers
'================================================================================

Private Sub UserControl_Initialize()

mPrevHeight = UserControl.Height

ReDim mRegions(100) As RegionTableEntry
mRegionsIndex = -1

Set mPeriods = New Collection

mBackColour = vbWhite
mGridColour = &HC0C0C0
mShowGrid = True
mShowCrosshairs = True

mTwipsPerBar = DefaultTwipsPerBar
mScaleHeight = -100
mScaleTop = 100
mYAxisWidthCm = 1.5

resizeX

mAllowHorizontalMouseScrolling = True
mAllowVerticalMouseScrolling = True

HScroll.Height = 0

End Sub

Private Sub UserControl_Paint()
Static paintcount As Long
paintcount = paintcount + 1
Debug.Print "Control_paint" & paintcount
mPainted = True
paintAll
End Sub

Private Sub UserControl_Resize()
Static resizeCount As Long
resizeCount = resizeCount + 1
'debug.print "Control_resize: count = " & resizeCount
mNotFirstMouseMove = False
HScroll.top = UserControl.Height - HScroll.Height
HScroll.Width = UserControl.Width
XAxisPicture.top = HScroll.top - XAxisPicture.Height
XAxisPicture.Width = UserControl.Width
resizeX
resizeY
paintAll
'debug.print "Exit Control_resize"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "autoscale", mAutoscale, True
End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
If Button = vbLeftButton Then mLeftDragging = True
mLeftDragStartPosnX = Int(X)
mLeftDragStartPosnY = Y
End Sub

Private Sub ChartRegionPicture_MouseMove(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

Dim region As ChartRegion
Dim i As Long

If mLeftDragging = True Then
    If mAllowHorizontalMouseScrolling Then
        ' the chart needs to be scrolled so that current mouse position
        ' is the value contained in mLeftDragStartPosnX
        If mLeftDragStartPosnX <> Int(X) Then
            scrollX mLeftDragStartPosnX - Int(X)
        End If
    End If
    If mAllowVerticalMouseScrolling Then
        If mLeftDragStartPosnY <> Y Then
            With mRegions(index - 1).region
                If Not .autoscale Then
                    .scrollVertical mLeftDragStartPosnY - Y
                End If
            End With
        End If
    End If
Else
    For i = 0 To mRegionsIndex
        Set region = mRegions(i).region
        If i = index - 1 Then
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & y
            region.MouseMove Button, Shift, X, Y
        Else
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & MinusInfinitySingle
            region.MouseMove Button, Shift, X, MinusInfinitySingle
        End If
    Next
    displayXAxisLabel X, 100
End If
End Sub

Private Sub ChartRegionPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
If Button = vbLeftButton Then mLeftDragging = False
End Sub

'================================================================================
' HScroll Event Handlers
'================================================================================

Private Sub HScroll_Change()
scrollX HScroll.value - lastVisiblePeriod
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
If Button = vbLeftButton Then mLeftDragging = True
mLeftDragStartPosnX = Int(X)
mLeftDragStartPosnY = Y
mUserResizingRegions = True
End Sub

Private Sub RegionDividerPicture_MouseMove( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim vertChange As Long
Dim currRegion As Long
Dim newHeight As Long
Dim prevPercentHeight As Double

If Not mLeftDragging = True Then Exit Sub

currRegion = index + 1  ' we resize the region below the divider
vertChange = mLeftDragStartPosnY - Y
newHeight = mRegions(currRegion).actualHeight + vertChange

' the region table indicates the requested percentage used by each region
' and the actual height allocation. We need to work out the new percentage
' for the region to be resized.

prevPercentHeight = mRegions(currRegion).region.percentheight
If Not mRegions(currRegion).useAvailableSpace Then
    mRegions(currRegion).region.percentheight = mRegions(currRegion).percentheight * newHeight / mRegions(currRegion).actualHeight
Else
    ' this is a 'use available space' region that's being resized. Now change
    ' it to use a specific percentage
    mRegions(currRegion).region.percentheight = 100 * newHeight / calcAvailableHeight
End If

If sizeRegions Then
    paintAll
Else
    ' the regions couldn't be resized so reset the region's percent height
    mRegions(currRegion).region.percentheight = prevPercentHeight
End If
End Sub

Private Sub RegionDividerPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
If Button = vbLeftButton Then mLeftDragging = False
mUserResizingRegions = False
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get allowHorizontalMouseScrolling() As Boolean
allowHorizontalMouseScrolling = mAllowHorizontalMouseScrolling
End Property

Public Property Let allowHorizontalMouseScrolling(ByVal value As Boolean)
mAllowHorizontalMouseScrolling = value
PropertyChanged "allowHorizontalMouseScrolling"
End Property

Public Property Get allowVerticalMouseScrolling() As Boolean
allowVerticalMouseScrolling = mAllowVerticalMouseScrolling
End Property

Public Property Let allowVerticalMouseScrolling(ByVal value As Boolean)
mAllowVerticalMouseScrolling = value
PropertyChanged "allowVerticalMouseScrolling"
End Property

Public Property Get autoscale() As Boolean
autoscale = mAutoscale
End Property

Public Property Let autoscale(ByVal value As Boolean)
mAutoscale = value
PropertyChanged "autoscale"
End Property

'Public Property Get barSpacingPercent() As Single
'barSpacingPercent = mBarSpacingPercent
'End Property
'
'Public Property Let barSpacingPercent(ByVal value As Single)
'mBarSpacingPercent = value
'mCandleWidth = 100! / (100! + mBarSpacingPercent)
'PropertyChanged "barspacingpercent"
'End Property

Public Property Get chartBackColor() As Long
chartBackColor = mBackColour
End Property

Public Property Let chartBackColor(ByVal val As Long)
Dim i As Long

mBackColour = val
XAxisPicture.backColor = val

For i = 0 To mRegionsIndex
    mRegions(i).region.regionBackColor = val
Next
paintAll
End Property

Public Property Get chartLeft() As Single
chartLeft = mScaleLeft
End Property

Public Property Get chartWidth() As Single
chartWidth = YAxisPosition - mScaleLeft
End Property

Public Property Get currentPeriodNumber() As Long
currentPeriodNumber = mCurrentPeriodNumber
End Property

Public Property Get currentTool() As ToolTypes
currentTool = mCurrentTool
End Property

Public Property Let currentTool(ByVal value As ToolTypes)
Select Case value
Case ToolPointer
    mCurrentTool = value
Case ToolLine
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineExtended
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineRay
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineHorizontal
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineVertical
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciRetracement
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciExtension
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciCircle
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciTime
    mCurrentTool = ToolTypes.ToolPointer
Case ToolRegressionChannel
    mCurrentTool = ToolTypes.ToolPointer
Case ToolRegressionEnvelope
    mCurrentTool = ToolTypes.ToolPointer
Case ToolText
    mCurrentTool = ToolTypes.ToolPointer
Case ToolPitchfork
    mCurrentTool = ToolTypes.ToolPointer
End Select
End Property

Public Property Get gridColor() As Long
gridColor = mGridColour
End Property

Public Property Let gridColor(ByVal val As Long)
mGridColour = val
End Property

Public Property Get lastVisiblePeriod() As Long
lastVisiblePeriod = mYAxisPosition - 1
End Property

Public Property Let lastVisiblePeriod(ByVal value As Long)
scrollX value - mYAxisPosition + 1
End Property

Public Property Get showCrosshairs() As Boolean
showCrosshairs = mShowCrosshairs
End Property

Public Property Let showCrosshairs(ByVal val As Boolean)
Dim i As Long
Dim region As ChartRegion
mShowCrosshairs = val
For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.showCrosshairs = val
Next
End Property

Public Property Get showGrid() As Boolean
showGrid = mShowGrid
End Property

Public Property Let showGrid(ByVal val As Boolean)
mShowGrid = val
End Property

Public Property Get showHorizontalScrollBar() As Boolean
showHorizontalScrollBar = mShowHorizontalScrollBar
End Property

Public Property Let showHorizontalScrollBar(ByVal val As Boolean)
mShowHorizontalScrollBar = val
If mShowHorizontalScrollBar Then
    HScroll.Height = 255
Else
    HScroll.Height = 0
End If
resizeY
End Property

Public Property Get suppressDrawing() As Boolean
suppressDrawing = mSuppressDrawing
End Property

Public Property Let suppressDrawing(ByVal val As Boolean)
Dim i As Long
Dim region As ChartRegion
mSuppressDrawing = val
For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.suppressDrawing = val
Next
'If Not mSuppressDrawing Then paintAll
End Property

Public Property Get twipsPerBar() As Long
twipsPerBar = mTwipsPerBar
End Property

Public Property Let twipsPerBar(ByVal val As Long)
mTwipsPerBar = val
resizeX
setHorizontalScrollBar
paintAll
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Public Property Get YAxisWidthCm() As Single
YAxisWidthCm = mYAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
mYAxisWidthCm = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function addChartRegion(ByVal percentheight As Double, _
                    Optional ByVal minimumPercentHeight As Double) As ChartRegion
'
' NB: percentHeight=100 means the region will use whatever space
' is available
'

Dim YAxisRegion As ChartRegion

Set addChartRegion = New ChartRegion

Load ChartRegionPicture(ChartRegionPicture.UBound + 1)
ChartRegionPicture(ChartRegionPicture.UBound).align = vbAlignNone
ChartRegionPicture(ChartRegionPicture.UBound).Width = _
    UserControl.ScaleWidth * (mYAxisPosition - chartLeft) / XAxisPicture.ScaleWidth
ChartRegionPicture(ChartRegionPicture.UBound).visible = True

Load YAxisPicture(YAxisPicture.UBound + 1)
YAxisPicture(YAxisPicture.UBound).align = vbAlignNone
YAxisPicture(YAxisPicture.UBound).left = ChartRegionPicture(ChartRegionPicture.UBound).Width
YAxisPicture(YAxisPicture.UBound).Width = UserControl.ScaleWidth - YAxisPicture(YAxisPicture.UBound).left
YAxisPicture(YAxisPicture.UBound).visible = True

addChartRegion.surface = ChartRegionPicture(ChartRegionPicture.UBound)
addChartRegion.suppressDrawing = mSuppressDrawing
addChartRegion.currentTool = mCurrentTool
addChartRegion.gridColor = mGridColour
addChartRegion.minimumPercentHeight = minimumPercentHeight
addChartRegion.percentheight = percentheight
addChartRegion.regionBackColor = mBackColour
addChartRegion.regionLeft = mScaleLeft
addChartRegion.regionHeight = 1
addChartRegion.regionNumber = mRegionsIndex + 2
addChartRegion.regionTop = 1
addChartRegion.regionWidth = mYAxisPosition - mScaleLeft
addChartRegion.showCrosshairs = mShowCrosshairs
addChartRegion.showGrid = mShowGrid
addChartRegion.periodsInView mScaleLeft, mYAxisPosition - 1
addChartRegion.autoscale = mAutoscale

If mRegionsIndex = UBound(mRegions) Then
    ReDim Preserve mRegions(UBound(mRegions) + 100) As RegionTableEntry
End If

mRegionsIndex = mRegionsIndex + 1
Set mRegions(mRegionsIndex).region = addChartRegion
mRegions(mRegionsIndex).percentheight = percentheight
mRegions(mRegionsIndex).useAvailableSpace = (percentheight = 100#)

If mRegionsIndex <> 0 Then
    Load RegionDividerPicture(mRegionsIndex)
    RegionDividerPicture(mRegionsIndex).visible = True
End If

Set YAxisRegion = New ChartRegion
YAxisRegion.surface = YAxisPicture(YAxisPicture.UBound)
YAxisRegion.regionHeight = 1
YAxisRegion.regionTop = 1
addChartRegion.YAxisRegion = YAxisRegion

If Not sizeRegions Then
    ' can't fit this all in! So remove the added region,
    Set addChartRegion = Nothing
    Set mRegions(mRegionsIndex).region = Nothing
    mRegions(mRegionsIndex).percentheight = 0
    mRegions(mRegionsIndex).actualHeight = 0
    mRegions(mRegionsIndex).useAvailableSpace = False
    mRegionsIndex = mRegionsIndex - 1
    Unload ChartRegionPicture(ChartRegionPicture.UBound)
    Unload RegionDividerPicture(RegionDividerPicture.UBound)
    Unload YAxisPicture(YAxisPicture.UBound)
End If

End Function

Public Function addPeriod(ByVal timestamp As Date) As Period
Dim i As Long
Dim region As ChartRegion

Set addPeriod = New Period
mCurrentPeriodNumber = mCurrentPeriodNumber + 1
addPeriod.periodNumber = mCurrentPeriodNumber
addPeriod.timestamp = timestamp
addPeriod.backColor = mBackColour
mPeriods.Add addPeriod, CStr(addPeriod.periodNumber)

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.addPeriod mCurrentPeriodNumber
Next
setHorizontalScrollBar
End Function

Public Function refresh()
UserControl.refresh
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function calcAvailableHeight() As Long
calcAvailableHeight = XAxisPicture.top - _
                    mRegionsIndex * RegionDividerPicture(0).Height
End Function

Private Sub displayXAxisLabel(X As Single, Y As Single)
Dim thisPeriod As Period
Dim periodNumber As Long
Dim prevPeriodNumber As Long
Dim prevPeriod As Period

If Round(X) >= mYAxisPosition Then Exit Sub
If mPeriods.count = 0 Then Exit Sub

On Error Resume Next
periodNumber = Round(X)
Set thisPeriod = mPeriods(periodNumber)
On Error GoTo 0
If thisPeriod Is Nothing Then Exit Sub

With XAxisPicture
    If mNotFirstMouseMove Then
        prevPeriodNumber = Round(mPrevCursorX)
        If prevPeriodNumber > 0 Then
            .DrawMode = vbXorPen
            .ForeColor = .backColor
            .CurrentX = prevPeriodNumber
            .CurrentY = 100
            Set prevPeriod = mPeriods(prevPeriodNumber)
            XAxisPicture.Print Format(prevPeriod.timestamp, "dd/mm hh:nn")
        End If
    End If
    
    mPrevCursorX = X
    mPrevCursorY = Y
    mNotFirstMouseMove = True
    
    If Round(X) <= 0 Then Exit Sub
    
    .DrawMode = vbXorPen
    .CurrentX = periodNumber
    .CurrentY = Y
    .ForeColor = vbRed 'Xor BackColor 'vbWhite 'Xor BackColor
    XAxisPicture.Print Format(thisPeriod.timestamp, "dd/mm hh:nn")
    
End With

End Sub

Private Sub paintAll()
Dim region As ChartRegion
Dim i As Long

If mSuppressDrawing Then Exit Sub

mNotFirstMouseMove = False

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.paintRegion
Next

XAxisPicture.Cls
displayXAxisLabel mPrevCursorX, 100

End Sub

Private Sub resizeX()
Dim newScaleWidth As Single
Dim i As Long
Dim region As ChartRegion

newScaleWidth = CSng(XAxisPicture.Width) / CSng(mTwipsPerBar) - 0.5!
If newScaleWidth = mScaleWidth Then Exit Sub

mScaleWidth = newScaleWidth

mScaleLeft = mYAxisPosition + _
            (mYAxisWidthCm * TwipsPerCm / XAxisPicture.Width * mScaleWidth) - _
            mScaleWidth
XAxisPicture.ScaleWidth = mScaleWidth
XAxisPicture.ScaleLeft = mScaleLeft

For i = 0 To ChartRegionPicture.UBound
    YAxisPicture(i).left = UserControl.Width - YAxisPicture(i).Width
    ChartRegionPicture(i).Width = YAxisPicture(i).left
Next

For i = 0 To RegionDividerPicture.UBound
    RegionDividerPicture(i).Width = UserControl.Width
Next

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.regionWidth = mYAxisPosition - mScaleLeft
    region.periodsInView mScaleLeft, mYAxisPosition - 1
Next

End Sub

Private Sub resizeY()
If UserControl.Height = mPrevHeight Then Exit Sub

'debug.print "resizeY"

mPrevHeight = UserControl.Height
sizeRegions
End Sub

Public Sub scrollX(ByVal value As Long)
Dim region As ChartRegion
Dim i As Long
mYAxisPosition = mYAxisPosition + value
mScaleLeft = mYAxisPosition + _
            (mYAxisWidthCm * TwipsPerCm / XAxisPicture.Width * mScaleWidth) - _
            mScaleWidth
XAxisPicture.ScaleLeft = mScaleLeft
For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.periodsInView mScaleLeft, mYAxisPosition - 1
Next
setHorizontalScrollBar
paintAll
End Sub

Private Sub setHorizontalScrollBar()
HScroll.Max = IIf(lastVisiblePeriod > mCurrentPeriodNumber, lastVisiblePeriod, mCurrentPeriodNumber)
HScroll.Min = IIf(lastVisiblePeriod < (chartWidth - 1), lastVisiblePeriod, chartWidth - 1)
If HScroll.Max = HScroll.Min Then HScroll.Min = HScroll.Min - 1
HScroll.value = lastVisiblePeriod
HScroll.SmallChange = 1
If chartWidth > (HScroll.Max - HScroll.Min) Then
    HScroll.LargeChange = HScroll.Max - HScroll.Min
Else
    HScroll.LargeChange = chartWidth
End If
End Sub

Private Function sizeRegions() As Boolean
'
' NB: percentHeight=100 means the region will use whatever space
' is available
'
Dim i As Long
Dim top As Long
Dim aRegion As ChartRegion
Dim numAvailableSpaceRegions As Long
Dim heightReductionFactor As Double
Dim totalMinimumPercents As Double
Dim nonFixedAvailableSpacePercent As Double
Dim availableSpacePercent As Double
Dim availableHeight As Long     ' the space available for the region picture boxes
                                ' excluding the divider pictures

availableSpacePercent = 100
nonFixedAvailableSpacePercent = 100
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    mRegions(i).percentheight = aRegion.percentheight
    If Not mRegions(i).useAvailableSpace Then
        availableSpacePercent = availableSpacePercent - mRegions(i).percentheight
        nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).percentheight
    Else
        If aRegion.minimumPercentHeight <> 0 Then
            availableSpacePercent = availableSpacePercent - aRegion.minimumPercentHeight
        End If
        numAvailableSpaceRegions = numAvailableSpaceRegions + 1
    End If
Next

If availableSpacePercent < 0 And mUserResizingRegions Then
    sizeRegions = False
    Exit Function
End If

heightReductionFactor = 1
Do While availableSpacePercent < 0
    availableSpacePercent = 100
    nonFixedAvailableSpacePercent = 100
    heightReductionFactor = heightReductionFactor * 0.66666667
    For i = 0 To mRegionsIndex
        Set aRegion = mRegions(i).region
        If Not mRegions(i).useAvailableSpace Then
            If aRegion.minimumPercentHeight <> 0 Then
                If aRegion.percentheight * heightReductionFactor >= _
                    aRegion.minimumPercentHeight _
                Then
                    mRegions(i).percentheight = aRegion.percentheight * heightReductionFactor
                Else
                    mRegions(i).percentheight = aRegion.minimumPercentHeight
                    totalMinimumPercents = totalMinimumPercents + aRegion.minimumPercentHeight
                End If
            Else
                mRegions(i).percentheight = aRegion.percentheight * heightReductionFactor
            End If
            availableSpacePercent = availableSpacePercent - mRegions(i).percentheight
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).percentheight
        Else
            If aRegion.minimumPercentHeight <> 0 Then
                availableSpacePercent = availableSpacePercent - aRegion.minimumPercentHeight
                totalMinimumPercents = totalMinimumPercents + aRegion.minimumPercentHeight
            End If
        End If
    Next
    If totalMinimumPercents > 100 Then
        ' can't possibly fit this all in!
        sizeRegions = False
        Exit Function
    End If
Loop

If numAvailableSpaceRegions = 0 Then
    ' we must adjust the percentages on the other regions so they
    ' total 100.
    For i = 0 To mRegionsIndex
        mRegions(i).percentheight = 100 * mRegions(i).percentheight / (100 - nonFixedAvailableSpacePercent)
    Next
End If

' calculate the actual available height to put these regions in
availableHeight = calcAvailableHeight

' first set heights for fixed height regions
For i = 0 To mRegionsIndex
    If Not mRegions(i).useAvailableSpace Then
        mRegions(i).actualHeight = mRegions(i).percentheight * availableHeight / 100
    End If
Next

' now set heights for 'available space' regions with a minimum height
' that needs to be respected
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    If mRegions(i).useAvailableSpace And _
        aRegion.minimumPercentHeight <> 0 _
    Then
        mRegions(i).actualHeight = 0
        If (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) < aRegion.minimumPercentHeight Then
            mRegions(i).actualHeight = aRegion.minimumPercentHeight * availableHeight / 100
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.minimumPercentHeight
            numAvailableSpaceRegions = numAvailableSpaceRegions - 1
        End If
    End If
Next

' finally set heights for all other 'available space' regions
For i = 0 To mRegionsIndex
    If mRegions(i).useAvailableSpace And _
        mRegions(i).actualHeight = 0 _
    Then
        mRegions(i).actualHeight = (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) * availableHeight / 100
    End If
Next

' Now actually set the heights and positions for the picture boxes
top = 0
    
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    ChartRegionPicture(aRegion.regionNumber).Height = mRegions(i).actualHeight
    YAxisPicture(aRegion.regionNumber).Height = mRegions(i).actualHeight
    ChartRegionPicture(aRegion.regionNumber).top = top
    YAxisPicture(aRegion.regionNumber).top = top
    top = top + ChartRegionPicture(aRegion.regionNumber).Height
    If i <> mRegionsIndex Then
        RegionDividerPicture(i).top = top
        top = top + RegionDividerPicture(i).Height
    End If
Next

sizeRegions = True
End Function

Private Sub zoom(ByRef rect As TRectangle)

End Sub

