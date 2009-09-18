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


'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External FUnction Declarations
'================================================================================

'Private Declare Function CreateRectRgn Lib "gdi32" ( _
'                ByVal x1 As Long, _
'                ByVal y1 As Long, _
'                ByVal x2 As Long, _
'                ByVal y2 As Long) As Long
'
'Private Declare Function CreateSolidBrush Lib "gdi32" ( _
'                ByVal crColor As Long) As Long
'
'Private Declare Function DeleteObject Lib "gdi32" ( _
'                ByVal hObject As Long) As Long
'
'Private Declare Function FillRect Lib "user32" ( _
'                ByVal hDC As Long, _
'                lpRect As RECT, _
'                ByVal hBrush As Long) As Long
'
'Private Declare Function GetBkColor Lib "gdi32" ( _
'                ByVal hDC As Long) As Long
'
'Private Declare Function GetClipRgn Lib "gdi32" ( _
'                ByVal hDC As Long, _
'                ByVal hRgn As Long) As Long
'
'Private Declare Function Polygon Lib "gdi32" ( _
'                ByVal hDC As Long, _
'                lpPoint As POINTAPI, _
'                ByVal nCount As Long) As Long
'
'Private Declare Function SelectClipRgn Lib "gdi32" ( _
'                ByVal hDC As Long, _
'                ByVal hRgn As Long) As Long
'
'================================================================================
' Member variables
'================================================================================

Private mSurface As PictureBox
Attribute mSurface.VB_VarHelpID = -1

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

Private mClipRegion As Long

Private mRegionType As RegionTypes

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
gLogger.Log LogLevels.LogLevelHighDetail, "Canvas terminated (" & gRegionTypeToString(mRegionType) & ")"
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
hDC = mSurface.hDC
End Property

Friend Property Get Height() As Double
Height = mSurface.Height
End Property

Friend Property Let MouseIcon(ByVal value As IPictureDisp)
mSurface.MouseIcon = value
End Property

Friend Property Let MousePointer(ByVal value As VBRUN.MousePointerConstants)
mSurface.MousePointer = value
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
MousePointer = mSurface.MousePointer
End Property

Friend Property Let RegionType(ByVal value As RegionTypes)
mRegionType = value
End Property

Friend Property Let Surface(ByVal value As PictureBox)
Set mSurface = value
End Property

Friend Property Get Surface() As PictureBox
Set Surface = mSurface
End Property

Friend Property Get Width() As Double
Width = mSurface.Width
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Clear()
Dim failpoint As Long
On Error GoTo Err

mSurface.Cls

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Clear" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub ClearClippingRegion()
Dim failpoint As Long
On Error GoTo Err

SelectClipRgn mSurface.hDC, 0
DeleteObject mClipRegion

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ClearClippingRegion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawCircle( _
                ByVal centre As Point, _
                ByVal radius As Double)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
applyFillSettings
mSurface.Circle (centre.XLogical, centre.YLogical), radius

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "DrawCircle" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawLine( _
                ByVal x1 As Single, _
                ByVal y1 As Single, _
                ByVal x2 As Single, _
                ByVal y2 As Single)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
mSurface.Line (x1, y1)-(x2, y2)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "DrawLine" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawPoint( _
                ByVal p As Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
mSurface.PSet (p.XLogical, p.YLogical)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "DrawPoint" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawPolygon( _
                ByRef points() As W32Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
applyFillSettings
Polygon mSurface.hDC, points(0), UBound(points) + 1

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "DrawPolygon" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawRectangle( _
                ByVal p1 As Point, _
                ByVal p2 As Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
applyFillSettings
mSurface.Line (p1.XLogical, p1.YLogical)-(p2.XLogical, p2.YLogical), , B

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "DrawRectangle" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawText( _
                ByVal Text As String, _
                ByVal location As Point)
Dim textLines() As String
Dim i As Long
                
Dim failpoint As Long
On Error GoTo Err

setfont
    
mSurface.FontTransparent = mTextTransparent
mSurface.CurrentX = location.XLogical
mSurface.CurrentY = location.YLogical
mSurface.ForeColor = mTextColor
textLines = Split(Text, vbCrLf)
For i = 0 To UBound(textLines)
    mSurface.Print textLines(i)
    mSurface.CurrentX = location.XLogical
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "DrawText" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Friend Sub Finish()
' See this Microsoft bug: http://support.microsoft.com/kb/q190511/
' We have to make sure that all references to ControlArray members are released
' to ensure that this ActiveX control's memory is released
gLogger.Log LogLevelHighDetail, "Canvas finished"
Set mSurface = Nothing
End Sub

Public Sub SetBrushAttributes( _
                ByVal Color As Long, _
                ByVal Style As FillStyles)
Dim failpoint As Long
On Error GoTo Err

mBrushColor = Color
mBrushStyle = Style

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetBrushAttributes" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub SetClippingRegion( _
                ByRef clipArea As RECT)
Dim res As Long
Dim failpoint As Long
On Error GoTo Err


mClipRegion = CreateRectRgn(clipArea.Left, _
                        clipArea.Top, _
                        clipArea.Right, _
                        clipArea.Bottom)
If mClipRegion = 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            ProjectName & "." & ModuleName & ":" & "SetClippingRegion", _
            "Failed to create GDI region"
End If

res = SelectClipRgn(mSurface.hDC, mClipRegion)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetClippingRegion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub SetScale(ByRef pBoundary As TRectangle)
If mSurface Is Nothing Then Exit Sub
mSurface.ScaleHeight = pBoundary.Bottom - pBoundary.Top
mSurface.ScaleWidth = pBoundary.Right - pBoundary.Left
mSurface.ScaleTop = pBoundary.Top
mSurface.ScaleLeft = pBoundary.Left
End Sub

Public Sub SetPenAttributes( _
                ByVal Color As Long, _
                ByVal Width As Long, _
                ByVal Style As LineStyles, _
                ByVal mode As DrawModes)
Dim failpoint As Long
On Error GoTo Err

mPenColor = Color
mPenWidth = Width
mPenStyle = Style
mPenMode = mode

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetPenAttributes" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub SetTextAttributes( _
                ByVal Color As Long, _
                ByVal Font As StdFont, _
                ByVal transparent As Boolean)
Dim failpoint As Long
On Error GoTo Err

mTextColor = Color
Set mTextFont = Font
mTextTransparent = transparent

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetTextAttributes" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Function TextHeight(ByVal value As String) As Double
Dim failpoint As Long
On Error GoTo Err

setfont
TextHeight = Abs(mSurface.TextHeight(value))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "TextHeight" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function TextWidth(ByVal value As String) As Double
Dim failpoint As Long
On Error GoTo Err

setfont
TextWidth = mSurface.TextWidth(value)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "TextWidth" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Friend Function ZOrder( _
                ByVal position As Integer)
mSurface.ZOrder position
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyFillSettings()
mSurface.FillColor = mBrushColor
mSurface.FillStyle = mBrushStyle
End Sub

Private Sub applyPenSettings()
mSurface.ForeColor = mPenColor
mSurface.DrawWidth = mPenWidth
mSurface.DrawStyle = mPenStyle
mSurface.DrawMode = mPenMode
End Sub

Private Sub setfont()
If Not mTextFont Is Nothing Then
    With mSurface.Font
        If mTextFont.size < 8 Then .size = mTextFont.size
       .Name = mTextFont.Name
       .Bold = mTextFont.Bold
       .Italic = mTextFont.Italic
       .size = mTextFont.size
       .Charset = mTextFont.Charset
       .Strikethrough = mTextFont.Strikethrough
       .Underline = mTextFont.Underline
       .Weight = mTextFont.Weight
    End With
End If
End Sub

