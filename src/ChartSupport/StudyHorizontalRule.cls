VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyHorizontalRule"
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
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mY As Double
Private mColor As Long
Private mThickness As Long
Private mStyle As ChartSkil.LineStyles

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mColor = vbBlack
mThickness = 1
mStyle = LineSolid
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

Public Property Let color(ByVal value As Long)
mColor = value
End Property

Public Property Get color() As Long
color = mColor
End Property

Public Property Let style(ByVal value As ChartSkil.LineStyles)
mStyle = value
End Property

Public Property Get style() As ChartSkil.LineStyles
style = mStyle
End Property

Public Property Let thickness(ByVal value As Long)
mThickness = value
End Property

Public Property Get thickness() As Long
thickness = mThickness
End Property

Public Property Let y(ByVal value As Double)
mY = value
End Property

Public Property Get y() As Double
y = mY
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


