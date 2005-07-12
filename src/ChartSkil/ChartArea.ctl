VERSION 5.00
Begin VB.UserControl Chart 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ScaleHeight     =   7335
   ScaleWidth      =   9390
   Begin VB.PictureBox YBorderPicture 
      Height          =   6375
      Left            =   8400
      ScaleHeight     =   6375
      ScaleWidth      =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   30
   End
   Begin VB.PictureBox XBorderPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8415
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.PictureBox XAxisPicture 
      Align           =   2  'Align Bottom
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
   Begin VB.PictureBox ChartZonePicture 
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
    region As ChartRegion
    percentheight As Double
    actualHeight As Long
End Type

'================================================================================
' Member variables and constants
'================================================================================

Private Const DefaultTwipsPerBar As Long = 150

Private mRegions() As RegionTableEntry
Private mRegionsIndex As Long

Private mCurrentPeriodNumber As Long

Private mPeriods As Collection
Private mXScaleRegion As ChartRegion

Private mAutoscale As Boolean
Private mScaleWidth As Single
Private mScaleHeight As Single
Private mScaleLeft As Single
Private mScaleTop As Single

Private mPrevHeight As Single

Private mTwipsPerBar As Long

Private mYAxisPosition As Long

Private mVertGridSpacing As Double
Private mYScaleFormatStr As String
Private mYScaleSubFormatStr As String
Private mGridTextHeight As Double

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
Dim i As Long
Dim aBar As bar
Dim last As Double
Dim aBarSeries As BarSeries

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

resizeX

'mScaleWidth = UserControl.Width / mTwipsPerBar
'mScaleLeft = 7 - chartWidth
'
'XBorderPicture.Width = ((YAxisPosition - mScaleLeft) / mScaleWidth) * XAxisPicture.Width
'XBorderPicture.left = 0
'XBorderPicture.top = XAxisPicture.top
'YBorderPicture.Height = XAxisPicture.top + Screen.TwipsPerPixelY
'YBorderPicture.left = XBorderPicture.Width
'
'XAxisPicture.scaleWidth = mScaleWidth
'XAxisPicture.scaleLeft = mScaleLeft
'lastVisiblePeriod = 0

End Sub

Private Sub UserControl_Paint()
Static paintcount As Long
paintcount = paintcount + 1
'debug.print "Control_paint" & paintcount
mPainted = True
paintAll
End Sub

Private Sub UserControl_Resize()
Static resizeCount As Long
resizeCount = resizeCount + 1
'debug.print "Control_resize: count = " & resizeCount
mNotFirstMouseMove = False
resizeX
resizeY
paintAll
'debug.print "Exit Control_resize"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "autoscale", mAutoscale, True
End Sub

'================================================================================
' ChartZonePicture Event Handlers
'================================================================================

Private Sub ChartZonePicture_MouseMove(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

Dim region As ChartRegion
Dim i As Long

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    If i = index - 1 Then
        'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & y
        region.MouseMove Button, Shift, x, y
    Else
        'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & MinusInfinitySingle
        region.MouseMove Button, Shift, x, MinusInfinitySingle
    End If
Next
displayXAxisLabel x, 100
End Sub

'Private Sub ChartZonePicture_Paint(index As Integer)
'mRegions(index - 1).region.paintRegion
'End Sub

'Private Sub ChartZonePicture_Resize(Index As Integer)
'refresh
'End Sub

'================================================================================
' XAxisPicture Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

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
mBackColour = val
End Property

Public Property Get chartHeight() As Single
chartHeight = mScaleHeight
End Property

Public Property Let chartHeight(ByVal value As Single)
mScaleHeight = value
PropertyChanged "chartheight"
End Property

Public Property Get chartLeft() As Single
chartLeft = mScaleLeft
End Property

Public Property Let chartLeft(ByVal value As Single)
Dim shiftPeriods As Long
shiftPeriods = value - mScaleLeft
mScaleLeft = value
scrollX shiftPeriods
PropertyChanged "chartleft"
End Property

Public Property Get chartTop() As Single
chartTop = mScaleTop
End Property

Public Property Let chartTop(ByVal value As Single)
mScaleTop = value
PropertyChanged "charttop"
End Property

Public Property Get chartWidth() As Single
chartWidth = mScaleWidth
End Property

Public Property Let chartWidth(ByVal value As Single)
mScaleWidth = value
PropertyChanged "chartwidth"
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

Public Property Get showGrid() As Boolean
showGrid = mShowGrid
End Property

Public Property Let showGrid(ByVal val As Boolean)
mShowGrid = val
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
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
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
Dim i As Long
Dim top As Long
Dim aRegion As ChartRegion

Set addChartRegion = New ChartRegion
Load ChartZonePicture(ChartZonePicture.UBound + 1)
ChartZonePicture(ChartZonePicture.UBound).align = vbAlignNone
ChartZonePicture(ChartZonePicture.UBound).Width = UserControl.Width
ChartZonePicture(ChartZonePicture.UBound).visible = True
addChartRegion.surface = ChartZonePicture(ChartZonePicture.UBound)
addChartRegion.suppressDrawing = mSuppressDrawing
addChartRegion.autoscale = mAutoscale
addChartRegion.currentTool = mCurrentTool
addChartRegion.gridColor = mGridColour
addChartRegion.minimumPercentHeight = minimumPercentHeight
addChartRegion.percentheight = percentheight
addChartRegion.regionBackColor = mBackColour
addChartRegion.regionLeft = mScaleLeft
addChartRegion.regionHeight = -1
addChartRegion.regionNumber = mRegionsIndex + 2
addChartRegion.regionTop = 1
addChartRegion.regionWidth = mScaleWidth
addChartRegion.showCrosshairs = mShowCrosshairs
addChartRegion.showGrid = mShowGrid
addChartRegion.periodsInView mScaleLeft, YAxisPosition - 1

If mRegionsIndex = UBound(mRegions) Then
    ReDim Preserve mRegions(UBound(mRegions) + 100) As RegionTableEntry
End If

mRegionsIndex = mRegionsIndex + 1
Set mRegions(mRegionsIndex).region = addChartRegion
mRegions(mRegionsIndex).percentheight = percentheight

If Not sizeRegions Then
    ' can't fit this all in! So remove the added region,
    Set addChartRegion = Nothing
    Set mRegions(mRegionsIndex).region = Nothing
    mRegions(mRegionsIndex).percentheight = 0
    mRegions(mRegionsIndex).actualHeight = 0
    mRegionsIndex = mRegionsIndex - 1
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

'shift 1
End Function

Public Function newPoint(ByVal x As Double, _
                        ByVal y As Double, _
                        Optional ByVal relative As Boolean = False) As Point
Set newPoint = New Point
newPoint.x = x
newPoint.y = y
newPoint.relative = relative
End Function

Public Function refresh()
UserControl.refresh
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub displayXAxisLabel(x As Single, y As Single)
Dim thisPeriod As Period
Dim periodNumber As Long
Dim prevPeriodNumber As Long
Dim prevPeriod As Period

If Round(x) >= YAxisPosition Then Exit Sub
If mPeriods.count = 0 Then Exit Sub

With XAxisPicture
    .scaleWidth = chartWidth
    .scaleLeft = chartLeft
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
    
    mPrevCursorX = x
    mPrevCursorY = y
    mNotFirstMouseMove = True
    
    If Round(x) <= 0 Then Exit Sub
    
    .DrawMode = vbXorPen
    periodNumber = Round(x)
    .CurrentX = periodNumber
    .CurrentY = y
    .ForeColor = vbRed 'Xor BackColor 'vbWhite 'Xor BackColor
    Set thisPeriod = mPeriods(periodNumber)
    XAxisPicture.Print Format(thisPeriod.timestamp, "dd/mm hh:nn")
    
End With

End Sub

Private Sub paintAll()
Dim region As ChartRegion
Dim aPeriod As Period
Dim i As Long

If mSuppressDrawing Then Exit Sub
'If Not mPainted Then Exit Sub

mNotFirstMouseMove = False

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.paintRegion
Next

XAxisPicture.Cls
'For Each aPeriod In mPeriods
'    If aPeriod.periodNumber >= mScaleLeft And _
'        aPeriod.periodNumber < YAxisPosition _
'    Then
'        XAxisPicture.Line (aPeriod.periodNumber, 0)-(aPeriod.periodNumber + 1, XAxisPicture.Height), aPeriod.backColor, BF
'    End If
'Next
displayXAxisLabel mPrevCursorX, 100

End Sub

Private Sub resizeX()
Dim newScaleWidth As Single
Dim i As Long
Dim region As ChartRegion

newScaleWidth = CSng(XAxisPicture.Width) / CSng(mTwipsPerBar) - 0.5!
If newScaleWidth = mScaleWidth Then Exit Sub

mScaleWidth = newScaleWidth
mScaleLeft = YAxisPosition + 7 - mScaleWidth
XAxisPicture.scaleWidth = mScaleWidth
XAxisPicture.scaleLeft = mScaleLeft

XBorderPicture.Width = ((YAxisPosition - chartLeft) / chartWidth) * XAxisPicture.Width
XBorderPicture.left = 0
XBorderPicture.top = XAxisPicture.top
YBorderPicture.Height = XAxisPicture.top + Screen.TwipsPerPixelY
YBorderPicture.left = XBorderPicture.Width

For i = 0 To ChartZonePicture.UBound
    ChartZonePicture(i).Width = UserControl.Width
Next

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.regionWidth = mScaleWidth
    region.periodsInView mScaleLeft, YAxisPosition - 1
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
mScaleLeft = mYAxisPosition + 7 - mScaleWidth
XAxisPicture.scaleLeft = mScaleLeft
For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.periodsInView mScaleLeft, mYAxisPosition - 1
Next
'mPrevCursorX = mYAxisPosition - 1
paintAll
End Sub

Private Function sizeRegions() As Boolean
'
' NB: percentHeight=100 means the region will use whatever space
' is available
'
Dim i As Long
Dim top As Long
Dim aRegion As ChartRegion
Dim num100percentRegions As Long
Dim heightReductionFactor As Double
Dim totalMinimumPercents As Double
Dim nonFixedAvailableSpacePercent As Double
Dim availableSpacePercent As Double
Dim drawingWasAllowed As Boolean

availableSpacePercent = 100
nonFixedAvailableSpacePercent = 100
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    mRegions(i).actualHeight = 0
    mRegions(i).percentheight = aRegion.percentheight
    If aRegion.percentheight <> 100 Then
        availableSpacePercent = availableSpacePercent - aRegion.percentheight
        nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.percentheight
    Else
        If aRegion.minimumPercentHeight <> 0 Then
            availableSpacePercent = availableSpacePercent - aRegion.minimumPercentHeight
        End If
        num100percentRegions = num100percentRegions + 1
    End If
Next

heightReductionFactor = 1
Do While availableSpacePercent < 0
    availableSpacePercent = 100
    heightReductionFactor = heightReductionFactor * 0.66666667
    For i = 0 To mRegionsIndex
        Set aRegion = mRegions(i).region
        If aRegion.percentheight <> 100 Then
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
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.percentheight
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


' first set heights for fixed height regions
For i = 0 To mRegionsIndex
    If mRegions(i).percentheight <> 100 Then
        mRegions(i).actualHeight = mRegions(i).percentheight * (UserControl.Height - XAxisPicture.Height) / 100
    End If
Next

' now set heights for 'available space' regions with a minimum height
' that needs to be respected
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    If mRegions(i).percentheight = 100 And _
        aRegion.minimumPercentHeight <> 0 _
    Then
        If (nonFixedAvailableSpacePercent / num100percentRegions) < aRegion.minimumPercentHeight Then
            mRegions(i).actualHeight = aRegion.minimumPercentHeight * (UserControl.Height - XAxisPicture.Height) / 100
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.minimumPercentHeight
            num100percentRegions = num100percentRegions - 1
        End If
    End If
Next

' finally set heights for all other 'available space' regions
For i = 0 To mRegionsIndex
    If mRegions(i).percentheight = 100 And _
        mRegions(i).actualHeight = 0 _
    Then
        mRegions(i).actualHeight = (nonFixedAvailableSpacePercent / num100percentRegions) * (UserControl.Height - XAxisPicture.Height) / 100
    End If
Next

' Now actually set the heights and positions for the picture boxes
top = 0
'If Not suppressDrawing Then
'    drawingWasAllowed = True
'    suppressDrawing = True
'End If
    
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    ChartZonePicture(aRegion.regionNumber).Height = mRegions(i).actualHeight
    ChartZonePicture(aRegion.regionNumber).top = top
    top = top + ChartZonePicture(aRegion.regionNumber).Height
'    aRegion.regionHeight = aRegion.regionHeight ' reset the drawing surface ScaleHeight
'    aRegion.regionTop = aRegion.regionTop   ' reset the drawing surface ScaleTop
Next
'If drawingWasAllowed Then suppressDrawing = False
sizeRegions = True
End Function

Private Sub zoom(ByRef rect As TRectangle)

End Sub

