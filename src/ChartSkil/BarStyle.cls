VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarStyle"
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

Private mTailThickness As Long
Private mOutlineThickness As Long

Private mBarColor As Long
Private mUpColor As Long
Private mDownColor As Long

Private mDisplayMode As BarDisplayModes

Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mBarWidth As Single

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mBarColor = -1
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

Public Property Let barColor(ByVal value As Long)
mBarColor = value
End Property

Public Property Get barColor() As Long
barColor = mBarColor
End Property

Public Property Let barThickness(ByVal value As Long)
mBarThickness = value
End Property

Public Property Get barThickness() As Long
barThickness = mBarThickness
End Property

Public Property Let barWidth(ByVal value As Single)
mBarWidth = value
End Property

Public Property Get barWidth() As Single
barWidth = mBarWidth
End Property

Public Property Let displayMode(ByVal value As BarDisplayModes)
mDisplayMode = value
End Property

Public Property Get displayMode() As BarDisplayModes
displayMode = mDisplayMode
End Property

Public Property Let downColor(ByVal value As Long)
mDownColor = value
End Property

Public Property Get downColor() As Long
downColor = mDownColor
End Property

Public Property Let outlineThickness(ByVal value As Long)
mOutlineThickness = value
End Property

Public Property Get outlineThickness() As Long
outlineThickness = mOutlineThickness
End Property

Public Property Let solidUpBody(ByVal value As Boolean)
mSolidUpBody = value
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mSolidUpBody
End Property

Public Property Let tailThickness(ByVal value As Long)
mTailThickness = value
End Property

Public Property Get tailThickness() As Long
tailThickness = mTailThickness
End Property

Public Property Let upColor(ByVal value As Long)
mUpColor = value
End Property

Public Property Get upColor() As Long
upColor = mUpColor
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As BarStyle
Set clone = New BarStyle
clone.barColor = mBarColor
clone.barThickness = mBarThickness
clone.barWidth = mBarWidth
clone.displayMode = mDisplayMode
clone.downColor = mDownColor
clone.outlineThickness = mOutlineThickness
clone.solidUpBody = mSolidUpBody
clone.tailThickness = mTailThickness
clone.upColor = mUpColor
End Function

'@================================================================================
' Helper Functions
'@================================================================================




