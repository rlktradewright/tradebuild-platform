VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PointerController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "PointerController"

'@================================================================================
' Member variables
'@================================================================================

Private mViewport                                   As ViewPort

Private mRegionType As RegionTypes
Private mLineSeries As LineSeries

Private mPointerMode                                As PointerModes
Private mPointerStyle                               As PointerStyles
Private mPointerIcon                                As IPictureDisp
Private mPointerToolStyle                           As PointerStyles
Private mPointerToolIcon                            As IPictureDisp

Private mCrosshairLineHoriz                         As Line
Private mCrosshairLineVert                          As Line
Private mCrosshairOffsetLeft                        As Size
Private mCrosshairOffsetRight                       As Size
Private mCrosshairOffsetBottom                      As Size
Private mCrosshairOffsetTop                         As Size

Private mUsingWaitCursor                            As Boolean

Private mPrevCursorX                                As Single
Private mPrevCursorY                                As Single
Private mPrevRegionContainsPointer                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let CrosshairLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

If Not Value Is Nothing Then
    mCrosshairLineHoriz.LocalStyle = Value
    mCrosshairLineVert.LocalStyle = Value
Else
    mCrosshairLineHoriz.LocalStyle = gDefaultLineStyle
    mCrosshairLineVert.LocalStyle = gDefaultLineStyle
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let PointerIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "PointerIcon"
On Error GoTo Err

If Value Is Nothing Then Exit Property
If Value Is mPointerIcon Then Exit Property

If mPointerMode <> PointerModeDefault Then
    Set mPointerIcon = Value
    Exit Property
End If

HidePointer

Set mPointerIcon = Value

If Not mUsingWaitCursor Then
    setMouseIcon mPointerStyle
    showPointer mPrevCursorX, mPrevCursorY, mPrevRegionContainsPointer
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Friend Property Get PointerIcon() As IPictureDisp
Set PointerIcon = mPointerIcon
End Property

Friend Property Let PointerMode(ByVal Value As PointerModes)
Const ProcName As String = "PointerMode"
On Error GoTo Err

If mPointerMode = Value Then Exit Property

HidePointer

mPointerMode = Value

If mUsingWaitCursor Then Exit Property

setMousePointer
setPointerStyle
RedrawPointer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get PointerMode() As PointerModes
PointerMode = mPointerMode
End Property

Friend Property Let PointerStyle(ByVal Value As PointerStyles)
Const ProcName As String = "PointerStyle"
On Error GoTo Err

If Not mRegionType = RegionTypeData Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot set PointerStyle for an axis region"

If mPointerStyle = Value Then Exit Property

If mPointerMode <> PointerModeDefault Then
    mPointerStyle = Value
    Exit Property
End If

HidePointer

mPointerStyle = Value

setPointerStyle

If Not mUsingWaitCursor Then
    setMouseIcon mPointerStyle
    showPointer mPrevCursorX, mPrevCursorY, mPrevRegionContainsPointer
End If
Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get PointerStyle() As PointerStyles
PointerStyle = mPointerStyle
End Property

Friend Property Let UseWaitCursor(ByVal Value As Boolean)
If Value Then
    HidePointer
    mViewport.MousePointer = vbHourglass
    mUsingWaitCursor = True
Else
    mUsingWaitCursor = False
    setMousePointer
    RedrawPointer
End If
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub DrawPointer( _
                ByVal Button As Integer, _
                ByVal Shift As Integer, _
                ByVal X As Single, _
                ByVal Y As Single, _
                ByVal pRegionContainsPointer As Boolean)
Const ProcName As String = "DrawPointer"
On Error GoTo Err

mPrevCursorX = X
mPrevCursorY = Y
mPrevRegionContainsPointer = pRegionContainsPointer

showPointer X, Y, pRegionContainsPointer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub HidePointer()
Const ProcName As String = "HidePointer"
On Error GoTo Err

Select Case getCurrentPointerStyle
Case PointerNone

Case PointerCrosshairs
    hideCrosshairsPointer
Case PointerDisc
    hideDiscPointer
Case PointerSelection
    hideSelectionPointer
Case PointerTool
    hideToolPointer
Case PointerCustom
    hideCustomPointer
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pViewport As ViewPort, _
                ByVal pRegionType As RegionTypes, _
                ByVal pLineSeries As LineSeries)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mViewport = pViewport
mRegionType = pRegionType
Set mLineSeries = pLineSeries

If mRegionType = RegionTypeData Then
    setupCrosshairs
    mPointerStyle = PointerCrosshairs
Else
    mPointerStyle = PointerWindowsDefault
End If

setMousePointer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RedrawPointer()
Const ProcName As String = "RedrawPointer"
On Error GoTo Err

showPointer mPrevCursorX, mPrevCursorY, mPrevRegionContainsPointer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetToolPointer( _
                ByVal toolPointerStyle As PointerStyles, _
                ByVal icon As IPictureDisp)
Const ProcName As String = "SetToolPointer"
On Error GoTo Err

mPointerToolStyle = toolPointerStyle
Set mPointerToolIcon = icon

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub drawCrosshairsPointer( _
                ByVal X As Single, _
                ByVal Y As Single)
Const ProcName As String = "drawCrosshairsPointer"
On Error GoTo Err

Dim lineX As Single

If mUsingWaitCursor Then Exit Sub

lineX = Round(X)

mCrosshairLineHoriz.SetPosition gNewPoint(X, Y, , , mCrosshairOffsetLeft), _
                                gNewPoint(X, Y, , , mCrosshairOffsetRight)
mCrosshairLineVert.SetPosition gNewPoint(lineX, Y, , , mCrosshairOffsetBottom), _
                                gNewPoint(lineX, Y, , , mCrosshairOffsetTop)
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub drawCustomPointer(ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "drawCustomPointer"
On Error GoTo Err

If mUsingWaitCursor Then Exit Sub

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub drawDiscPointer(ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "drawDiscPointer"
On Error GoTo Err

drawCrosshairsPointer X, Y

'Dim lineX As Single
'Dim Size As Size
'
'If mUsingWaitCursor Then Exit Sub
'
'With mViewport.Canvas
'    Dim lColor As Long
'    lColor = .GetColorAtPoint(X, Y)
'
'    lineX = Round(X)
'    Set Size = gNewSize(0.4, 0.4)
'
'    .SetPenAttributes vbBlack Xor lColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine lineX - Size.WidthLogical(mViewport.Region), Y, lineX + Size.WidthLogical(mViewport.Region), Y
'    .DrawLine lineX, Y - Size.HeightLogical(mViewport.Region), lineX, Y + Size.HeightLogical(mViewport.Region)
'
'    .SetPenAttributes mPointerDiscColor Xor lColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .SetBrushAttributes mPointerDiscColor Xor lColor, FillStyles.FillSolid
'    .DrawCircle gNewPoint(lineX, Y), gNewDimension(0.8).LengthLogicalX(mViewport.Region)
'
'End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub drawSelectionPointer(ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "drawSelectionPointer"
On Error GoTo Err

If mUsingWaitCursor Then Exit Sub

' nothing to draw since the pointer is a cursor and Windows draws it

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub drawToolPointer(ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "drawToolPointer"
On Error GoTo Err

drawCrosshairsPointer X, Y

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCurrentPointerStyle() As PointerStyles
Dim lPointerStyle As PointerStyles

If mPointerMode = PointerModeTool Then
    lPointerStyle = mPointerToolStyle
ElseIf mPointerMode = PointerModeSelection Then
    lPointerStyle = PointerSelection
Else
    lPointerStyle = mPointerStyle
End If
getCurrentPointerStyle = lPointerStyle
End Function

Private Sub hideCrosshairsPointer()
Const ProcName As String = "hideCrosshairsPointer"
On Error GoTo Err

mCrosshairLineHoriz.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)
mCrosshairLineVert.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub hideCustomPointer()

' nothing to do

End Sub

Private Sub hideDiscPointer()
Const ProcName As String = "undrawDiscPointer"
On Error GoTo Err

mCrosshairLineHoriz.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)
mCrosshairLineVert.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub hideSelectionPointer()

If mUsingWaitCursor Then Exit Sub

' nothing to do since the pointer is actually a proper cursor
End Sub

Private Sub hideToolPointer()
Const ProcName As String = "undrawToolPointer"
On Error GoTo Err

If mUsingWaitCursor Then Exit Sub

mCrosshairLineHoriz.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)
mCrosshairLineVert.SetPosition gNewPoint(-1000, 0), _
                                gNewPoint(-1000, 0)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setMouseIcon( _
                ByVal PointerStyle As PointerStyles)
Const ProcName As String = "setMouseIcon"
On Error GoTo Err

mViewport.MousePointer = vbCustom
Select Case PointerStyle
Case PointerNone
    mViewport.MouseIcon = gBlankCursor
Case PointerCrosshairs
    mViewport.MouseIcon = gBlankCursor
Case PointerDisc
    mViewport.MouseIcon = gBlankCursor
Case PointerSelection
    mViewport.MouseIcon = gSelectorCursor
Case PointerTool
    mViewport.MouseIcon = gBlankCursor
Case PointerCustom
    mViewport.MouseIcon = mPointerIcon
Case PointerWindowsDefault
    mViewport.MousePointer = vbDefault
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be a member of the PointerStyles enum"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setMousePointer()
Const ProcName As String = "setMousePointer"
On Error GoTo Err

Select Case mPointerMode
Case PointerModeDefault
    setMouseIcon mPointerStyle
Case PointerModeTool
    setMouseIcon mPointerToolStyle
Case PointerModeSelection
    setMouseIcon PointerStyles.PointerSelection
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPointerCrosshairs()
Const ProcName As String = "setPointerCrosshairs"
On Error GoTo Err

mCrosshairLineHoriz.ClearProperty LinePropertyThickness
mCrosshairLineVert.ClearProperty LinePropertyThickness
mCrosshairLineHoriz.ClearProperty LinePropertyLineStyle
mCrosshairLineVert.ClearProperty LinePropertyLineStyle

Set mCrosshairOffsetLeft = gNewSize(-100, 0, CoordsRelative)
Set mCrosshairOffsetRight = gNewSize(100, 0, CoordsRelative)
Set mCrosshairOffsetBottom = gNewSize(0, -100, CoordsRelative)
Set mCrosshairOffsetTop = gNewSize(0, 100, CoordsRelative)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPointerDisc()
Const ProcName As String = "setPointerDisc"
On Error GoTo Err

mCrosshairLineHoriz.ClearProperty LinePropertyThickness
mCrosshairLineVert.ClearProperty LinePropertyThickness
mCrosshairLineHoriz.ClearProperty LinePropertyLineStyle
mCrosshairLineVert.ClearProperty LinePropertyLineStyle

Set mCrosshairOffsetLeft = gNewSize(-0.4, 0)
Set mCrosshairOffsetRight = gNewSize(0.4, 0)
Set mCrosshairOffsetBottom = gNewSize(0, -0.4)
Set mCrosshairOffsetTop = gNewSize(0, 0.4)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPointerTool()
Const ProcName As String = "setPointerTool"
On Error GoTo Err

mCrosshairLineHoriz.Thickness = 1
mCrosshairLineVert.Thickness = 1
mCrosshairLineHoriz.LineStyle = LineSolid
mCrosshairLineVert.LineStyle = LineSolid

Set mCrosshairOffsetLeft = gNewSize(-0.6, 0)
Set mCrosshairOffsetRight = gNewSize(0.6, 0)
Set mCrosshairOffsetBottom = gNewSize(0, -0.6)
Set mCrosshairOffsetTop = gNewSize(0, 0.6)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPointerStyle()
Const ProcName As String = "setPointerStyle"
On Error GoTo Err

Select Case getCurrentPointerStyle
Case PointerNone

Case PointerCrosshairs
    setPointerCrosshairs
Case PointerDisc
    setPointerDisc
Case PointerTool
    setPointerTool
Case PointerCustom

Case PointerSelection

Case PointerWindowsDefault

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupCrosshairs()
Const ProcName As String = "setupCrosshairs"
On Error GoTo Err

Set mCrosshairLineHoriz = mLineSeries.Add
mCrosshairLineHoriz.LocalStyle = gDefaultLineStyle
mCrosshairLineHoriz.Extended = True
mCrosshairLineHoriz.IncludeInAutoscale = False
mCrosshairLineHoriz.Layer = LayerPointer
Set mCrosshairOffsetLeft = gNewSize(-100, 0, CoordsRelative)
Set mCrosshairOffsetRight = gNewSize(100, 0, CoordsRelative)

Set mCrosshairLineVert = mLineSeries.Add
mCrosshairLineVert.LocalStyle = gDefaultLineStyle
mCrosshairLineVert.Extended = True
mCrosshairLineVert.IncludeInAutoscale = False
mCrosshairLineVert.Layer = LayerPointer
Set mCrosshairOffsetBottom = gNewSize(0, -100, CoordsRelative)
Set mCrosshairOffsetTop = gNewSize(0, 100, CoordsRelative)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showPointer( _
                ByVal X As Single, _
                ByVal Y As Single, _
                ByVal pRegionContainsPointer As Boolean)
Const ProcName As String = "showPointer"
On Error GoTo Err

X = Round(X)
If X >= mViewport.Boundary.Right Then Exit Sub

If Not pRegionContainsPointer Then Y = mViewport.Boundary.Top + 2 * (mViewport.Boundary.Top - mViewport.Boundary.Bottom)

Select Case getCurrentPointerStyle
Case PointerNone

Case PointerCrosshairs
    drawCrosshairsPointer X, Y
Case PointerDisc
    drawDiscPointer X, Y
Case PointerSelection
    drawSelectionPointer X, Y
Case PointerTool
    drawToolPointer X, Y
Case PointerCustom
    drawCustomPointer X, Y
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Private Sub undrawDiscPointer()
'Const ProcName As String = "undrawDiscPointer"
'On Error GoTo Err
'
'Dim lSize As Size
'Dim prevLineX As Single
'
'If mUsingWaitCursor Then Exit Sub
'
'With mViewport.Canvas
'    Dim lColor As Long
'    lColor = .GetColorAtPoint(prevLineX, mPrevCursorY)
'
'    prevLineX = Round(mPrevCursorX)
'    Set lSize = gNewSize(0.4, 0.4)
'
'    .SetPenAttributes vbBlack Xor lColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine prevLineX - lSize.WidthLogical(mViewport.Region), mPrevCursorY, prevLineX + lSize.WidthLogical(mViewport.Region), mPrevCursorY
'    .DrawLine prevLineX, mPrevCursorY - lSize.HeightLogical(mViewport.Region), prevLineX, mPrevCursorY + lSize.HeightLogical(mViewport.Region)
'
'    .SetPenAttributes mPointerDiscColor Xor lColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .SetBrushAttributes mPointerDiscColor Xor lColor, FillStyles.FillSolid
'    .DrawCircle gNewPoint(prevLineX, mPrevCursorY), gNewDimension(0.8).LengthLogicalX(mViewport.Region)
'End With
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'
'End Sub




