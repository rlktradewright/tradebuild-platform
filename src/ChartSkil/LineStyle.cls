VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineStyle"
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

Private mColor As Long
Private mThickness As Long
Private mLineStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private marrowStartfillstyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

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

Public Property Let arrowStartColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "LineStyle" & ":" & "arrowStartColor", _
            "Value is not a valid color"
mArrowStartColor = value
End Property

Public Property Get arrowStartColor() As Long
arrowStartColor = mArrowStartColor
End Property

Public Property Let arrowStartFillColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "LineStyle" & ":" & "arrowStartFillColor", _
            "Value is not a valid color"
mArrowStartFillColor = value
End Property

Public Property Get arrowStartFillColor() As Long
arrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let arrowStartfillstyle(ByVal value As FillStyles)
marrowStartfillstyle = value
End Property

Public Property Get arrowStartfillstyle() As FillStyles
arrowStartfillstyle = marrowStartfillstyle
End Property

Public Property Let arrowStartLength(ByVal value As Long)
mArrowStartLength = value
End Property

Public Property Get arrowStartLength() As Long
arrowStartLength = mArrowStartLength
End Property

Public Property Let arrowStartStyle(ByVal value As ArrowStyles)
mArrowStartStyle = value
End Property

Public Property Get arrowStartStyle() As ArrowStyles
arrowStartStyle = mArrowStartStyle
End Property

Public Property Let arrowStartWidth(ByVal value As Long)
mArrowStartWidth = value
End Property

Public Property Get arrowStartWidth() As Long
arrowStartWidth = mArrowStartWidth
End Property

Public Property Let arrowEndColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "LineStyle" & ":" & "arrowEndColor", _
            "Value is not a valid color"
mArrowEndColor = value
End Property

Public Property Get arrowEndColor() As Long
arrowEndColor = mArrowEndColor
End Property

Public Property Let arrowEndFillColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "LineStyle" & ":" & "arrowEndFillColor", _
            "Value is not a valid color"
mArrowEndFillColor = value
End Property

Public Property Get arrowEndFillColor() As Long
arrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let arrowEndFillStyle(ByVal value As FillStyles)
mArrowEndFillStyle = value
End Property

Public Property Get arrowEndFillStyle() As FillStyles
arrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let arrowEndLength(ByVal value As Long)
mArrowEndLength = value
End Property

Public Property Get arrowEndLength() As Long
arrowEndLength = mArrowEndLength
End Property

Public Property Let arrowEndStyle(ByVal value As ArrowStyles)
mArrowEndStyle = value
End Property

Public Property Get arrowEndStyle() As ArrowStyles
arrowEndStyle = mArrowEndStyle
End Property

Public Property Let arrowEndWidth(ByVal value As Long)
mArrowEndWidth = value
End Property

Public Property Get arrowEndWidth() As Long
arrowEndWidth = mArrowEndWidth
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "LineStyle" & ":" & "color", _
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

Public Property Let extendAfter(ByVal value As Boolean)
mExtendAfter = value
End Property

Public Property Get extendAfter() As Boolean
extendAfter = mExtendAfter
End Property

Public Property Let extendBefore(ByVal value As Boolean)
mExtendBefore = value
End Property

Public Property Get extendBefore() As Boolean
extendBefore = mExtendBefore
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

Public Property Let includeInAutoscale(ByVal value As Boolean)
mIncludeInAutoscale = value
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let linestyle(ByVal value As LineStyles)
mLineStyle = value
End Property

Public Property Get linestyle() As LineStyles
linestyle = mLineStyle
End Property

Public Property Let thickness(ByVal value As Long)
mThickness = value
End Property

Public Property Get thickness() As Long
thickness = mThickness
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As linestyle
Set clone = New linestyle
clone.arrowEndColor = mArrowEndColor
clone.arrowEndFillColor = mArrowEndFillColor
clone.arrowEndFillStyle = mArrowEndFillStyle
clone.arrowEndLength = mArrowEndLength
clone.arrowEndStyle = mArrowEndStyle
clone.arrowEndWidth = mArrowEndWidth
clone.arrowStartColor = mArrowStartColor
clone.arrowStartFillColor = mArrowStartFillColor
clone.arrowStartfillstyle = marrowStartfillstyle
clone.arrowStartLength = mArrowStartLength
clone.arrowStartStyle = mArrowStartStyle
clone.arrowStartWidth = mArrowStartWidth
clone.Color = mColor
clone.extendAfter = mExtendAfter
clone.extendBefore = mExtendBefore
clone.extended = mExtended
clone.fixedX = mFixedX
clone.fixedY = mFixedY
clone.includeInAutoscale = mIncludeInAutoscale
clone.linestyle = mLineStyle
clone.thickness = mThickness
End Function

'@================================================================================
' Helper Functions
'@================================================================================


