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
Private mViewport As Viewport

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

Private Sub Class_Terminate()
'gLogger.Log LogLevels.LogLevelHighDetail, "Canvas terminated (" & gRegionTypeToString(mRegionType) & ")"
Debug.Print "Canvas terminated (" & gRegionTypeToString(mRegionType) & ")"
End Sub

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
Dim failpoint As String
On Error GoTo Err

hDC = mSurface.hDC

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get Height() As Double
Const ProcName As String = "Height"
Dim failpoint As String
On Error GoTo Err

Height = mSurface.Height

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let MouseIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "MouseIcon"
Dim failpoint As String
On Error GoTo Err

mSurface.MouseIcon = Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let MousePointer(ByVal Value As VBRUN.MousePointerConstants)
Const ProcName As String = "MousePointer"
Dim failpoint As String
On Error GoTo Err

mSurface.MousePointer = Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
Const ProcName As String = "MousePointer"
Dim failpoint As String
On Error GoTo Err

MousePointer = mSurface.MousePointer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let RegionType(ByVal Value As RegionTypes)
Const ProcName As String = "RegionType"
Dim failpoint As String
On Error GoTo Err

mRegionType = Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Surface(ByVal Value As PictureBox)
Const ProcName As String = "Surface"
Dim failpoint As String
On Error GoTo Err

Set mSurface = Value
SetGraphicsMode mSurface.hDC, GM_ADVANCED

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get Surface() As PictureBox
Const ProcName As String = "Surface"
Dim failpoint As String
On Error GoTo Err

Set Surface = mSurface

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Viewport(ByVal Value As Viewport)
Set mViewport = Value
End Property

Friend Property Get Width() As Double
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

Width = mSurface.Width

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub AddToClippingRegion( _
                ByRef clipArea As RECT)
Dim lRegionHandle As Long
Dim res As Long

Const ProcName As String = "AddToClippingRegion"
Dim failpoint As String
On Error GoTo Err

lRegionHandle = CreateRectRgn(clipArea.Left, _
                        clipArea.Top, _
                        clipArea.Right, _
                        clipArea.Bottom)
If lRegionHandle = 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Failed to create GDI region"
End If

res = ExtSelectClipRgn(mSurface.hDC, lRegionHandle, RGN_OR)

DeleteObject lRegionHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Clear()
Const ProcName As String = "Clear"
Dim failpoint As String
On Error GoTo Err

mSurface.Cls

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub ClearClippingRegion()
Const ProcName As String = "ClearClippingRegion"
Dim failpoint As String
On Error GoTo Err

SelectClipRgn mSurface.hDC, 0

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawCircle( _
                ByVal centre As Point, _
                ByVal radius As Double)
Const ProcName As String = "DrawCircle"
Dim failpoint As String
On Error GoTo Err

applyPenSettings
applyFillSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
mSurface.Circle (centre.XLogical(mViewport), centre.yLogical(mViewport)), radius

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawLine( _
                ByVal x1 As Single, _
                ByVal y1 As Single, _
                ByVal x2 As Single, _
                ByVal y2 As Single)
Const ProcName As String = "DrawLine"
Dim failpoint As String
On Error GoTo Err

applyPenSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
mSurface.Line (x1, y1)-(x2, y2)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawPoint( _
                ByVal p As Point)
Const ProcName As String = "DrawPoint"
Dim failpoint As String
On Error GoTo Err

applyPenSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
mSurface.PSet (p.XLogical(mViewport), p.yLogical(mViewport))

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawPolygon( _
                ByRef points() As W32Point)
Const ProcName As String = "DrawPolygon"
Dim failpoint As String
On Error GoTo Err

applyPenSettings
applyFillSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
Polygon mSurface.hDC, VarPtr(points(0)), UBound(points) + 1

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawRectangle( _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "DrawRectangle"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

lTRect.Left = x1Logical
lTRect.Bottom = IIf(y1Logical <= y2Logical, y1Logical, y2Logical)
lTRect.Right = x2Logical
lTRect.Top = IIf(y1Logical <= y2Logical, y2Logical, y1Logical)
RectValidate lTRect, True

DrawRectangleFromTRectangle lTRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawRectangleFromPoints( _
                ByVal p1 As Point, _
                ByVal p2 As Point)
Const ProcName As String = "DrawRectangleFromPoints"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

lTRect.Left = p1.XLogical(mViewport)
lTRect.Bottom = p1.yLogical(mViewport)
lTRect.Right = p2.XLogical(mViewport)
lTRect.Top = p2.yLogical(mViewport)
RectValidate lTRect

DrawRectangleFromTRectangle lTRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawRectangleFromRect( _
                ByRef pRect As RECT)
Const ProcName As String = "DrawRectangleFromRect"
Dim failpoint As String
On Error GoTo Err

applyPenSettings
applyFillSettings
SetGraphicsMode mSurface.hDC, GM_ADVANCED
Rectangle mSurface.hDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub DrawRectangleFromTRectangle( _
                ByRef pRect As TRectangle)
Const ProcName As String = "DrawRectangleFromTRectangle"
Dim failpoint As String
On Error GoTo Err

If Not pRect.isValid Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Argument is not a valid TRectangle"
End If

DrawRectangleFromRect mViewport.ConvertTRectangleToRect(pRect)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Dim clipRect As RECT
Dim res As Long
Dim dtp As DRAWTEXTPARAMS

Const ProcName As String = "DrawText"
Dim failpoint As String
On Error GoTo Err

setfont
mSurface.FontTransparent = mTextTransparent
mSurface.ForeColor = mTextColor

SetGraphicsMode mSurface.hDC, GM_ADVANCED

clipRect = mViewport.ConvertTRectangleToRect(pClippingRect)
    
dtp.cbSize = Len(dtp)
dtp.iTabLength = pTabWidth
If Not pLeftMargin Is Nothing Then dtp.iLeftMargin = mViewport.ConvertLogicalToPixelsX(pLeftMargin.LengthLogicalX(mViewport))
If Not pLeftMargin Is Nothing Then dtp.iRightMargin = mViewport.ConvertLogicalToPixelsX(pRightMargin.LengthLogicalX(mViewport))

res = User32.DrawTextEx(mSurface.hDC, _
                        StrPtr(pText), _
                        Len(pText), _
                        VarPtr(clipRect), _
                        pJustification + _
                            pEllipsis + _
                            ((Not pMultiLine) And DT_SINGLELINE) + _
                            (pExpandTabs And DT_EXPANDTABS) + _
                            (pWordWrap And DT_WORDBREAK), _
                        VarPtr(dtp))

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Finish()
' See this Microsoft bug: http://support.microsoft.com/kb/q190511/
' We have to make sure that all references to ControlArray members are released
' to ensure that this ActiveX control's memory is released
gLogger.Log LogLevelHighDetail, "Canvas finished"
Set mSurface = Nothing

Set mViewport = Nothing
End Sub

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
Dim clipRect As RECT
Dim res As Long
Dim dtp As DRAWTEXTPARAMS
Dim sizeHeight As Long
Dim sizeWidth As Long

Const ProcName As String = "GetTextSize"
Dim failpoint As String
On Error GoTo Err

setfont
mSurface.FontTransparent = mTextTransparent
mSurface.ForeColor = mTextColor

SetGraphicsMode mSurface.hDC, GM_ADVANCED

If Not pSize Is Nothing Then
    sizeWidth = mViewport.ConvertLogicalToPixelsX(pSize.WidthLogical(mViewport))
    sizeHeight = mViewport.ConvertLogicalToPixelsY(pSize.HeightLogical(mViewport))
    clipRect.Right = sizeWidth
    clipRect.Bottom = sizeHeight
Else
    clipRect.Right = &H7FFFFFF
    clipRect.Bottom = 0
End If
    
dtp.cbSize = Len(dtp)
dtp.iTabLength = pTabWidth
If Not pLeftMargin Is Nothing Then dtp.iLeftMargin = mViewport.ConvertLogicalToPixelsX(pLeftMargin.LengthLogicalX(mViewport))
If Not pLeftMargin Is Nothing Then dtp.iRightMargin = mViewport.ConvertLogicalToPixelsX(pRightMargin.LengthLogicalX(mViewport))

res = User32.DrawTextEx(mSurface.hDC, _
                        StrPtr(pText), _
                        Len(pText), _
                        VarPtr(clipRect), _
                        DT_CALCRECT + _
                            pJustification + _
                            pEllipsis + _
                            ((Not pMultiLine) And DT_SINGLELINE) + _
                            (pExpandTabs And DT_EXPANDTABS) + _
                            (pWordWrap And DT_WORDBREAK), _
                        VarPtr(dtp))
                        
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub IntersectClippingRegion( _
                ByRef clipArea As RECT)
Dim lRegionHandle As Long
Dim res As Long

Const ProcName As String = "IntersectClippingRegion"
Dim failpoint As String
On Error GoTo Err

lRegionHandle = CreateRectRgn(clipArea.Left, _
                        clipArea.Top, _
                        clipArea.Right, _
                        clipArea.Bottom)
If lRegionHandle = 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Failed to create GDI region"
End If

res = ExtSelectClipRgn(mSurface.hDC, lRegionHandle, RGN_AND)

DeleteObject lRegionHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RestoreClippingRegion()
Dim res As Long

Const ProcName As String = "RestoreClippingRegion"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SaveClippingRegion()

Const ProcName As String = "SaveClippingRegion"
Dim failpoint As String
On Error GoTo Err

Dim res As Long

If mSavedClippingRegionHandle <> 0 Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Clipping region has already been saved and not yet restored"
End If

mSavedClippingRegionHandle = CreateRectRgn(0, 0, 2, 2)
res = GetClipRgn(mSurface.hDC, mSavedClippingRegionHandle)

If res = -1 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Failed to get current GDI clipping region"
ElseIf res = 0 Then
    ' means there is no current clipping region - see documentation for GetClipRgn
    DeleteObject mSavedClippingRegionHandle
    mSavedClippingRegionHandle = NoClipRegion
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetBrushAttributes( _
                ByVal Color As Long, _
                ByVal Style As FillStyles)
Const ProcName As String = "SetBrushAttributes"
Dim failpoint As String
On Error GoTo Err

mBrushColor = Color
mBrushStyle = Style

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetClippingRegion( _
                ByRef clipArea As RECT)
Dim lRegionHandle As Long
Dim res As Long

Const ProcName As String = "SetClippingRegion"
Dim failpoint As String
On Error GoTo Err

lRegionHandle = CreateRectRgn(clipArea.Left, _
                        clipArea.Top, _
                        clipArea.Right, _
                        clipArea.Bottom)
If lRegionHandle = 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Failed to create GDI region"
End If

res = SelectClipRgn(mSurface.hDC, lRegionHandle)

DeleteObject lRegionHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub SetScale(ByRef pBoundary As TRectangle)
Const ProcName As String = "SetScale"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetPenAttributes( _
                ByVal Color As Long, _
                ByVal Width As Long, _
                ByVal Style As LineStyles, _
                ByVal mode As DrawModes)
Const ProcName As String = "SetPenAttributes"
Dim failpoint As String
On Error GoTo Err

mPenColor = Color
mPenWidth = IIf(Width <= 0, 1, Width)
mPenStyle = Style
mPenMode = mode

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetTextAttributes( _
                ByVal Color As Long, _
                ByVal Font As StdFont, _
                ByVal transparent As Boolean)
Const ProcName As String = "SetTextAttributes"
Dim failpoint As String
On Error GoTo Err

mTextColor = Color
Set mTextFont = Font
mTextTransparent = transparent

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function TextHeight(ByVal Value As String) As Double
Const ProcName As String = "TextHeight"
Dim failpoint As String
On Error GoTo Err

setfont
TextHeight = Abs(mSurface.TextHeight(Value))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function TextWidth(ByVal Value As String) As Double
Const ProcName As String = "TextWidth"
Dim failpoint As String
On Error GoTo Err

setfont
TextWidth = mSurface.TextWidth(Value)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub ZOrder( _
                ByVal Position As Integer)
Const ProcName As String = "ZOrder"
Dim failpoint As String
On Error GoTo Err

mSurface.ZOrder Position

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyFillSettings()
Const ProcName As String = "applyFillSettings"
Dim failpoint As String
On Error GoTo Err

mSurface.FillColor = mBrushColor
mSurface.FillStyle = mBrushStyle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub applyPenSettings()
Const ProcName As String = "applyPenSettings"
Dim failpoint As String
On Error GoTo Err

mSurface.ForeColor = mPenColor
mSurface.DrawWidth = mPenWidth
mSurface.DrawStyle = mPenStyle
mSurface.DrawMode = mPenMode

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setfont()
Const ProcName As String = "setfont"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

