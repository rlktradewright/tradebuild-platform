VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Canvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "Canvas"

' Used to indicate that there was no currently defined clip region when
' SaveClippingRegion was called
Private Const NoClipRegion              As Long = -1
'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External FUnction Declarations
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mSurface As PictureBox
Attribute mSurface.VB_VarHelpID = -1
Private mViewport As ViewPort

' pen attributes
Private mPenColor As Long
Private mPenWidth As Long
Private mPenStyle As LineStyles
Private mPenMode As DrawModes

' brush attributes
Private mBrushColor As Long
Private mBrushStyle As FillStyles

' Text attributes
Private mTextColor As Long
Private mTextFont As StdFont
Private mTextTransparent As Boolean

Private mRegionType As RegionTypes

Private mSavedClippingRegionHandle As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
End Sub

'Private Sub Class_Terminate()
'gLogger.Log LogLevels.LogLevelHighDetail, "Canvas terminated (" & gRegionTypeToString(mRegionType) & ")"
'Debug.Print "Canvas terminated (" & gRegionTypeToString(mRegionType) & ")"
'End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get hDC() As Long
Const ProcName As String = "hDC"

On Error GoTo Err

hDC = mSurface.hDC

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Height() As Double
Const ProcName As String = "Height"

On Error GoTo Err

Height = mSurface.Height

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let MouseIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "MouseIcon"

On Error GoTo Err

mSurface.MouseIcon = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let MousePointer(ByVal Value As VBRUN.MousePointerConstants)
Const ProcName As String = "MousePointer"

On Error GoTo Err

mSurface.MousePointer = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
Const ProcName As String = "MousePointer"

On Error GoTo Err

MousePointer = mSurface.MousePointer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let RegionType(ByVal Value As RegionTypes)
Const ProcName As String = "RegionType"

On Error GoTo Err

mRegionType = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Surface(ByVal Value As PictureBox)
Const ProcName As String = "Surface"

On Error GoTo Err

Set mSurface = Value
mSurface.MousePointer = MousePointerConstants.vbDefault
SetGraphicsMode mSurface.hDC, GM_ADVANCED

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Surface() As PictureBox
Const ProcName As String = "Surface"

On Error GoTo Err

Set Surface = mSurface

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let ViewPort(ByVal Value As ViewPort)
Set mViewport = Value
End Property

Friend Property Get Width() As Double
Const ProcName As String = "Width"

On Error GoTo Err

Width = mSurface.Width

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub AddToClippingRegion( _
                ByRef clipArea As GDI_RECT)
Dim lRegionHandle As Long
Dim res As Long

Const ProcName As String = "AddToClippingRegion"

On Error GoTo Err

lRegionHandle = CreateRectRgn(clipArea.Left, _
                        clipArea.Top, _
                        clipArea.Right, _
                        clipArea.Bottom)

If lRegionHandle = 0 Then Err.Raise ErrorCodes.ErrRuntimeException, , "Failed to create GDI region"

res = ExtSelectClipRgn(mSurface.hDC, lRegionHandle, RGN_OR)

DeleteObject lRegionHandle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Clear()
Const ProcName As String = "Clear"

On Error GoTo Err

mSurface.Cls

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearClippingRegion()
Const ProcName As String = "ClearClippingRegion"

On Error GoTo Err

SelectClipRgn mSurface.hDC, 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawCircle( _
                ByVal centre As Point, _
                ByVal radius As Double)
Const ProcName As String = "DrawCircle"

On Error GoTo Err

applyPenSettings
applyFillSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
mSurface.Circle (centre.XLogical(mViewport.Region), centre.yLogical(mViewport.Region)), radius

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawLine( _
                ByVal x1 As Single, _
                ByVal y1 As Single, _
                ByVal x2 As Single, _
                ByVal y2 As Single)
Const ProcName As String = "DrawLine"

On Error GoTo Err

applyPenSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
mSurface.Line (x1, y1)-(x2, y2)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, pErrorDesc:=Err.Description & ": Line: (" & x1 & ", " & y1 & "), (" & x2 & ", " & y2 & ")"
End Sub

Public Sub DrawPoint( _
                ByVal p As Point)
Const ProcName As String = "DrawPoint"

On Error GoTo Err

applyPenSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
mSurface.PSet (p.XLogical(mViewport.Region), p.yLogical(mViewport.Region))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawPolygon( _
                ByRef points() As GDI_POINT)
Const ProcName As String = "DrawPolygon"

On Error GoTo Err

applyPenSettings
applyFillSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
Polygon mSurface.hDC, VarPtr(points(0)), UBound(points) + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangle( _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "DrawRectangle"
Dim lTRect As TRectangle

On Error GoTo Err

lTRect.Left = x1Logical
lTRect.Bottom = IIf(y1Logical <= y2Logical, y1Logical, y2Logical)
lTRect.Right = x2Logical
lTRect.Top = IIf(y1Logical <= y2Logical, y2Logical, y1Logical)
RectValidate lTRect, True

DrawRectangleFromTRectangle lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangleFromPoints( _
                ByVal p1 As Point, _
                ByVal p2 As Point)
Const ProcName As String = "DrawRectangleFromPoints"
Dim lTRect As TRectangle

On Error GoTo Err

lTRect.Left = p1.XLogical(mViewport.Region)
lTRect.Bottom = p1.yLogical(mViewport.Region)
lTRect.Right = p2.XLogical(mViewport.Region)
lTRect.Top = p2.yLogical(mViewport.Region)
RectValidate lTRect

DrawRectangleFromTRectangle lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangleFromRect( _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "DrawRectangleFromRect"

On Error GoTo Err

applyPenSettings
applyFillSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
Rectangle mSurface.hDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangleFromTRectangle( _
                ByRef pRect As TRectangle)
Const ProcName As String = "DrawRectangleFromTRectangle"

On Error GoTo Err

If Not pRect.isValid Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Argument is not a valid TRectangle"

DrawRectangleFromRect mViewport.ConvertTRectangleToRect(pRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawText( _
                ByVal pText As String, _
                ByRef pClippingRect As TRectangle, _
                ByVal pJustification As TextJustifyModes, _
                ByVal pMultiLine As Boolean, _
                ByVal pEllipsis As EllipsisModes, _
                ByVal pExpandTabs As Boolean, _
                ByVal pTabWidth As Long, _
                ByVal pWordWrap As Boolean, _
                ByVal pLeftMargin As Dimension, _
                ByVal pRightMargin As Dimension)
Dim textLines() As String
Dim i As Long
Dim clipRect As GDI_RECT
Dim res As Long
Dim dtp As DRAWTEXTPARAMS

Const ProcName As String = "DrawText"

On Error GoTo Err

setfont
mSurface.FontTransparent = mTextTransparent
mSurface.ForeColor = mTextColor

SetGraphicsMode mSurface.hDC, GM_ADVANCED

clipRect = mViewport.ConvertTRectangleToRect(pClippingRect)
    
dtp.cbSize = Len(dtp)
dtp.iTabLength = pTabWidth
If Not pLeftMargin Is Nothing Then dtp.iLeftMargin = mViewport.ConvertLogicalToPixelsX(pLeftMargin.LengthLogicalX(mViewport.Region))
If Not pLeftMargin Is Nothing Then dtp.iRightMargin = mViewport.ConvertLogicalToPixelsX(pRightMargin.LengthLogicalX(mViewport.Region))

res = User32.DrawTextEx(mSurface.hDC, _
                        StrPtr(pText), _
                        Len(pText), _
                        clipRect, _
                        pJustification + _
                            pEllipsis + _
                            ((Not pMultiLine) And DT_SINGLELINE) + _
                            (pExpandTabs And DT_EXPANDTABS) + _
                            (pWordWrap And DT_WORDBREAK), _
                        dtp)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
' See this Microsoft bug: http://support.microsoft.com/kb/q190511/
' We have to make sure that all references to ControlArray members are released
' to ensure that this ActiveX control's memory is released
Const ProcName As String = "Finish"
On Error GoTo Err

gLogger.Log "Canvas finished", ProcName, ModuleName, LogLevelHighDetail
Set mSurface = Nothing

Set mViewport = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetColorAtPoint(ByVal X As Double, ByVal Y As Double) As Long
GetColorAtPoint = GetPixel(mSurface.hDC, _
                        mViewport.ConvertCoordsLogicalToPixelsX(X), _
                        mViewport.ConvertCoordsLogicalToPixelsY(Y))
End Function

Public Function GetTextSize( _
                ByVal pText As String, _
                ByRef pSize As Size, _
                ByVal pJustification As TextJustifyModes, _
                ByVal pMultiLine As Boolean, _
                ByVal pEllipsis As EllipsisModes, _
                ByVal pExpandTabs As Boolean, _
                ByVal pTabWidth As Long, _
                ByVal pWordWrap As Boolean, _
                ByVal pLeftMargin As Dimension, _
                ByVal pRightMargin As Dimension) As Size
Dim textLines() As String
Dim i As Long
Dim clipRect As GDI_RECT
Dim res As Long
Dim dtp As DRAWTEXTPARAMS
Dim sizeHeight As Long
Dim sizeWidth As Long

Const ProcName As String = "GetTextSize"

On Error GoTo Err

setfont
mSurface.FontTransparent = mTextTransparent
mSurface.ForeColor = mTextColor

SetGraphicsMode mSurface.hDC, GM_ADVANCED

If Not pSize Is Nothing Then
    sizeWidth = mViewport.ConvertLogicalToPixelsX(pSize.WidthLogical(mViewport.Region))
    sizeHeight = mViewport.ConvertLogicalToPixelsY(pSize.HeightLogical(mViewport.Region))
    clipRect.Right = sizeWidth
    clipRect.Bottom = sizeHeight
Else
    clipRect.Right = &H7FFFFFF
    clipRect.Bottom = 0
End If
    
dtp.cbSize = Len(dtp)
dtp.iTabLength = pTabWidth
If Not pLeftMargin Is Nothing Then dtp.iLeftMargin = mViewport.ConvertLogicalToPixelsX(pLeftMargin.LengthLogicalX(mViewport.Region))
If Not pLeftMargin Is Nothing Then dtp.iRightMargin = mViewport.ConvertLogicalToPixelsX(pRightMargin.LengthLogicalX(mViewport.Region))

res = User32.DrawTextEx(mSurface.hDC, _
                        StrPtr(pText), _
                        Len(pText), _
                        clipRect, _
                        DT_CALCRECT + _
                            pJustification + _
                            pEllipsis + _
                            ((Not pMultiLine) And DT_SINGLELINE) + _
                            (pExpandTabs And DT_EXPANDTABS) + _
                            (pWordWrap And DT_WORDBREAK), _
                        dtp)
                        
If Not pSize Is Nothing Then
    If sizeHeight <> 0 Then clipRect.Bottom = sizeHeight
    If sizeWidth <> 0 Then clipRect.Right = sizeWidth
End If

Set GetTextSize = gNewSize(mViewport.ConvertPixelsToLogicalX(clipRect.Right), _
                                            mViewport.ConvertPixelsToLogicalY(clipRect.Bottom), _
                                            CoordsLogical, _
                                            CoordsLogical)
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub IntersectClippingRegion( _
                ByRef clipArea As GDI_RECT)
Dim lRegionHandle As Long
Dim res As Long

Const ProcName As String = "IntersectClippingRegion"

On Error GoTo Err

lRegionHandle = CreateRectRgn(clipArea.Left, _
                        clipArea.Top, _
                        clipArea.Right, _
                        clipArea.Bottom)

If lRegionHandle = 0 Then Err.Raise ErrorCodes.ErrRuntimeException, , "Failed to create GDI region"

res = ExtSelectClipRgn(mSurface.hDC, lRegionHandle, RGN_AND)

DeleteObject lRegionHandle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RestoreClippingRegion()
Dim res As Long

Const ProcName As String = "RestoreClippingRegion"

On Error GoTo Err

If mSavedClippingRegionHandle = NoClipRegion Then
    ClearClippingRegion
Else
    res = SelectClipRgn(mSurface.hDC, mSavedClippingRegionHandle)
    DeleteObject mSavedClippingRegionHandle
End If

mSavedClippingRegionHandle = 0
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SaveClippingRegion()

Const ProcName As String = "SaveClippingRegion"

On Error GoTo Err

Dim res As Long

If mSavedClippingRegionHandle <> 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Clipping region has already been saved and not yet restored"

mSavedClippingRegionHandle = CreateRectRgn(0, 0, 2, 2)
res = GetClipRgn(mSurface.hDC, mSavedClippingRegionHandle)

If res = -1 Then Err.Raise ErrorCodes.ErrRuntimeException, , "Failed to get current GDI clipping region"

If res = 0 Then
    ' means there is no current clipping region - see documentation for GetClipRgn
    DeleteObject mSavedClippingRegionHandle
    mSavedClippingRegionHandle = NoClipRegion
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetBrushAttributes( _
                ByVal Color As Long, _
                ByVal Style As FillStyles)
Const ProcName As String = "SetBrushAttributes"

On Error GoTo Err

mBrushColor = Color
mBrushStyle = Style

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetClippingRegion( _
                ByRef clipArea As GDI_RECT)
Dim lRegionHandle As Long
Dim res As Long

Const ProcName As String = "SetClippingRegion"

On Error GoTo Err

lRegionHandle = CreateRectRgn(clipArea.Left, _
                        clipArea.Top, _
                        clipArea.Right, _
                        clipArea.Bottom)
If lRegionHandle = 0 Then Err.Raise ErrorCodes.ErrRuntimeException, , "Failed to create GDI region"

res = SelectClipRgn(mSurface.hDC, lRegionHandle)

DeleteObject lRegionHandle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetScale(ByRef pBoundary As TRectangle)
Const ProcName As String = "SetScale"

On Error GoTo Err

If mSurface Is Nothing Then Exit Sub
mSurface.ScaleHeight = pBoundary.Bottom - pBoundary.Top
mSurface.ScaleWidth = pBoundary.Right - pBoundary.Left
mSurface.ScaleTop = pBoundary.Top
mSurface.ScaleLeft = pBoundary.Left

'now adjust pBoundary to reflect the fact that ScaleHeight etc are singles not doubles
pBoundary.Top = mSurface.ScaleTop
pBoundary.Left = mSurface.ScaleLeft
pBoundary.Bottom = mSurface.ScaleHeight + pBoundary.Top
pBoundary.Right = mSurface.ScaleWidth + pBoundary.Left
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPenAttributes( _
                ByVal Color As Long, _
                ByVal Width As Long, _
                ByVal Style As LineStyles, _
                ByVal mode As DrawModes)
Const ProcName As String = "SetPenAttributes"

On Error GoTo Err

mPenColor = Color
mPenWidth = IIf(Width <= 0, 1, Width)
mPenStyle = Style
mPenMode = mode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetTextAttributes( _
                ByVal Color As Long, _
                ByVal Font As StdFont, _
                ByVal transparent As Boolean)
Const ProcName As String = "SetTextAttributes"

On Error GoTo Err

mTextColor = Color
Set mTextFont = Font
mTextTransparent = transparent

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function TextHeight(ByVal Value As String) As Double
Const ProcName As String = "TextHeight"

On Error GoTo Err

setfont
TextHeight = Abs(mSurface.TextHeight(Value))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TextWidth(ByVal Value As String) As Double
Const ProcName As String = "TextWidth"

On Error GoTo Err

setfont
TextWidth = mSurface.TextWidth(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ZOrder( _
                ByVal Position As Integer)
Const ProcName As String = "ZOrder"

On Error GoTo Err

mSurface.ZOrder Position

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyFillSettings()
Const ProcName As String = "applyFillSettings"

On Error GoTo Err

mSurface.FillColor = mBrushColor
mSurface.FillStyle = mBrushStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyPenSettings()
Const ProcName As String = "applyPenSettings"

On Error GoTo Err

mSurface.ForeColor = mPenColor

' must set DrawStyle before DrawWidth, otherwise DrawWidth may be forced to be 1
' if existing DrawStyle does not allow thick lines
mSurface.DrawStyle = mPenStyle
mSurface.DrawWidth = mPenWidth
mSurface.DrawMode = mPenMode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setfont()
Const ProcName As String = "setfont"

On Error GoTo Err

If Not mTextFont Is Nothing Then
    With mSurface.Font
        If mTextFont.Size < 8 Then .Size = mTextFont.Size
       .Name = mTextFont.Name
       .Bold = mTextFont.Bold
       .Italic = mTextFont.Italic
       .Size = mTextFont.Size
       .Charset = mTextFont.Charset
       .Strikethrough = mTextFont.Strikethrough
       .Underline = mTextFont.Underline
       .Weight = mTextFont.Weight
    End With
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

