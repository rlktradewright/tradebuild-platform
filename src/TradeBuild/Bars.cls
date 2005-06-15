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

Event BarAdded(ByVal theBar As TradeBuild.Bar)
Event BarUpdated(ByVal theBar As TradeBuild.Bar)
Event HistoricBarAdded(ByVal theBar As TradeBuild.Bar)

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

Public Property Get barLength() As Long
barLength = mBarLength
End Property

Friend Property Let barLength(ByVal value As Long)
mBarLength = value
End Property

Friend Property Let barType(ByVal value As TradeBuild.BarTypes)
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
add.barType = mBarType
add.barLength = mBarLength
mBars.add add, key
End Function

Friend Function addHistoric( _
                            ByVal key As String, _
                            ByVal timestamp As Date, _
                            ByVal barType As TradeBuild.BarTypes, _
                            ByVal barLength As Long, _
                            ByVal openPrice As Double, _
                            ByVal highPrice As Double, _
                            ByVal lowPrice As Double, _
                            ByVal closePrice As Double, _
                            ByVal volume As Long, _
                            ByVal tickVolume As Long, _
                            ByVal openInterest As Long) As Bar
Debug.Assert barType = mBarType
Debug.Assert barLength = mBarLength
Set addHistoric = New Bar
With addHistoric
    .barType = barType
    .barLength = barLength
    .DateTime = timestamp
    .openPrice = openPrice
    .highPrice = highPrice
    .lowPrice = lowPrice
    .closePrice = closePrice
    .volume = volume
    .tickVolume = tickVolume
    .openInterest = openInterest
End With
mBars.add addHistoric, key
RaiseEvent HistoricBarAdded(addHistoric)
End Function

Public Function item(index As Variant) As Bar
Attribute item.VB_UserMemId = 0
Set item = mBars(index)
End Function

Friend Function NotifyAdd(index As Variant)
RaiseEvent BarAdded(item(index))
End Function

Friend Function NotifyUpdate(index As Variant)
RaiseEvent BarUpdated(item(index))
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





