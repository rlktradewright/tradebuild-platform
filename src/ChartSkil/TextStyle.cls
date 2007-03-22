VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
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

'@================================================================================
' Member variables
'@================================================================================

Private mFont As StdFont
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillStyle As FillStyles
Private mAlign As TextAlignModes
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean
Private mPaddingX As Double
Private mPaddingY As Double

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

Public Property Let align(ByVal value As TextAlignModes)
mAlign = value
End Property

Public Property Get align() As TextAlignModes
align = mAlign
End Property

Public Property Let box(ByVal value As Boolean)
mBox = value
End Property

Public Property Get box() As Boolean
box = mBox
End Property

Public Property Let boxColor(ByVal value As Long)
mBoxColor = value
End Property

Public Property Get boxColor() As Long
boxColor = mBoxColor
End Property

Public Property Let boxFillColor(ByVal value As Long)
mBoxFillColor = value
End Property

Public Property Get boxFillColor() As Long
boxFillColor = mBoxFillColor
End Property

Public Property Let boxFillStyle(ByVal value As FillStyles)
mBoxFillStyle = value
End Property

Public Property Get boxFillStyle() As FillStyles
boxFillStyle = mBoxFillStyle
End Property

Public Property Let boxStyle(ByVal value As LineStyles)
mBoxStyle = value
End Property

Public Property Get boxStyle() As LineStyles
boxStyle = mBoxStyle
End Property

Public Property Let boxThickness(ByVal value As Long)
mBoxThickness = value
End Property

Public Property Get boxThickness() As Long
boxThickness = mBoxThickness
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "TextStyle" & ":" & "color", _
            "Value is not a valid color"
mColor = value
End Property

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let extended(ByVal value As Boolean)
mExtended = value
End Property

Public Property Get extended() As Boolean
extended = mExtended
End Property

Public Property Let fixedX(ByVal value As Boolean)
mFixedX = value
End Property

Public Property Get fixedX() As Boolean
fixedX = mFixedX
End Property

Public Property Let fixedY(ByVal value As Boolean)
mFixedY = value
End Property

Public Property Get fixedY() As Boolean
fixedY = mFixedY
End Property

Public Property Let font(ByVal value As StdFont)
Set mFont = value
End Property

Public Property Get font() As StdFont
Set font = mFont
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
mIncludeInAutoscale = value
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let paddingX(ByVal value As Double)
mPaddingX = value
End Property

Public Property Get paddingX() As Double
paddingX = mPaddingX
End Property

Public Property Let paddingY(ByVal value As Double)
mPaddingY = value
End Property

Public Property Get paddingY() As Double
paddingY = mPaddingY
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As TextStyle
Set clone = New TextStyle
clone.align = mAlign
clone.box = mBox
clone.boxColor = mBoxColor
clone.boxFillColor = mBoxFillColor
clone.boxFillStyle = mBoxFillStyle
clone.boxStyle = mBoxStyle
clone.boxThickness = mBoxThickness
clone.Color = mColor
clone.extended = mExtended
clone.fixedX = mFixedX
clone.fixedY = mFixedY
clone.font = mFont
clone.includeInAutoscale = mIncludeInAutoscale
clone.paddingX = mPaddingX
clone.paddingY = mPaddingY

End Function

'@================================================================================
' Helper Functions
'@================================================================================


