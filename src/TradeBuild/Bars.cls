VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bars"
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

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mBars As Collection
Private mBarType As TradeBuild.BarTypes
Private mBarLength As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
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

Public Property Get BarLength() As Long
BarLength = mBarLength
End Property

Friend Property Let BarLength(ByVal value As Long)
mBarLength = value
End Property

Friend Property Let BarType(ByVal value As TradeBuild.BarTypes)
mBarType = value
End Property

Public Property Get Count() As Long
Count = mBars.Count
End Property

'================================================================================
' Methods
'================================================================================

Public Function add(ByVal key As String) As Bar
Set add = New Bar
add.BarType = mBarType
add.BarLength = mBarLength
mBars.add add, key
End Function

Public Function item(index As Variant) As Bar
Attribute item.VB_UserMemId = 0
Set item = mBars(index)
End Function

Public Sub Remove(index As Variant)
mBars.Remove index
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Set NewEnum = mBars.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================





