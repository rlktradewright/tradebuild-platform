VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointStyle"
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
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mLineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mPointStyle As PointStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistBarWidth As Single
Private mIncludeInAutoscale As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mColor = vbBlack
mUpColor = -1
mDownColor = -1
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointStyle" & ":" & "color", _
            "Value is not a valid color"
mColor = value
End Property

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let displayMode(ByVal value As DataPointDisplayModes)
mDisplayMode = value
End Property

Public Property Get displayMode() As DataPointDisplayModes
displayMode = mDisplayMode
End Property

Public Property Let downColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointStyle" & ":" & "downColor", _
            "Value is not a valid color or -1"
mDownColor = value
End Property

Public Property Get downColor() As Long
downColor = mDownColor
End Property

Public Property Let histBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointStyle" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"

mHistBarWidth = value
End Property

Public Property Get histBarWidth() As Single
histBarWidth = mHistBarWidth
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
mIncludeInAutoscale = value
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let lineStyle(ByVal value As LineStyles)
mLineStyle = value
End Property

Public Property Get lineStyle() As LineStyles
lineStyle = mLineStyle
End Property

Public Property Let lineThickness(ByVal value As Long)
mLineThickness = value
End Property

Public Property Get lineThickness() As Long
lineThickness = mLineThickness
End Property

Public Property Let pointStyle(ByVal value As PointStyles)
mPointStyle = value
End Property

Public Property Get pointStyle() As PointStyles
pointStyle = mPointStyle
End Property

Public Property Let upColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointStyle" & ":" & "downColor", _
            "Value is not a valid color or -1"
mUpColor = value
End Property

Public Property Get upColor() As Long
upColor = mUpColor
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As DataPointStyle
Set clone = New DataPointStyle
clone.Color = mColor
clone.displayMode = mDisplayMode
clone.downColor = mDownColor
clone.histBarWidth = mHistBarWidth
clone.includeInAutoscale = mIncludeInAutoscale
clone.lineStyle = mLineStyle
clone.lineThickness = mLineThickness
clone.pointStyle = pointStyle
clone.upColor = mUpColor
End Function

'@================================================================================
' Helper Functions
'@================================================================================


