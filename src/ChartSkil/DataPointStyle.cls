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
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistBarWidth As Single
Private mIncludeInAutoscale As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mUpColor = vbBlack
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

Public Property Let displayMode(ByVal Value As DataPointDisplayModes)
mDisplayMode = Value
End Property

Public Property Get displayMode() As DataPointDisplayModes
displayMode = mDisplayMode
End Property

Public Property Let downColor(ByVal Value As Long)
If Not gIsValidColor(Value) And Value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointStyle" & ":" & "downColor", _
            "Value is not a valid color or -1"
mDownColor = Value
End Property

Public Property Get downColor() As Long
downColor = mDownColor
End Property

Public Property Let histBarWidth(ByVal Value As Single)
If Value <= 0 Or Value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointStyle" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"

mHistBarWidth = Value
End Property

Public Property Get histBarWidth() As Single
histBarWidth = mHistBarWidth
End Property

Public Property Let includeInAutoscale(ByVal Value As Boolean)
mIncludeInAutoscale = Value
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let lineStyle(ByVal Value As LineStyles)
mLineStyle = Value
End Property

Public Property Get lineStyle() As LineStyles
lineStyle = mLineStyle
End Property

Public Property Let lineThickness(ByVal Value As Long)
mLineThickness = Value
End Property

Public Property Get lineThickness() As Long
lineThickness = mLineThickness
End Property

Public Property Let upColor(ByVal Value As Long)
If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointStyle" & ":" & "downColor", _
            "Value is not a valid color"
mUpColor = Value
End Property

Public Property Get upColor() As Long
upColor = mUpColor
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As DataPointStyle
Set clone = New DataPointStyle
clone.displayMode = mDisplayMode
clone.downColor = mDownColor
clone.histBarWidth = mHistBarWidth
clone.includeInAutoscale = mIncludeInAutoscale
clone.lineStyle = mLineStyle
clone.lineThickness = mLineThickness
clone.upColor = mUpColor
End Function

'@================================================================================
' Helper Functions
'@================================================================================


