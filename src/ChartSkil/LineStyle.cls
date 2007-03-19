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
Private mArrowStartFillStyle As FillStyles
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

Public Property Let arrowEndStyle(ByVal value As ArrowStyles)
mArrowEndStyle = value
End Property

Public Property Get arrowEndStyle() As ArrowStyles
arrowEndStyle = mArrowEndStyle
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

Public Property Let lineStyle(ByVal value As LineStyles)
mLineStyle = value
End Property

Public Property Get lineStyle() As LineStyles
lineStyle = mLineStyle
End Property

Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private mArrowStartFillStyle As FillStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Public Property Let thickness(ByVal value As Long)
mThickness = value
End Property

Public Property Get thickness() As Long
thickness = mThickness
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


