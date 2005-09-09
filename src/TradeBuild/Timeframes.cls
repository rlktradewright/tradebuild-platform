VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timeframes"
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

Private mTimeframes As Collection

Private mTicker As TradeBuild.Ticker

Private mContract As TradeBuild.Contract

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTimeframes = New Collection
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

Friend Property Let Contract(ByVal value As Contract)
Dim lTimeframe As Timeframe
Set mContract = value
For Each lTimeframe In mTimeframes
    lTimeframe.Contract = mContract
Next
End Property

Public Property Get Count() As Long
Count = mTimeframes.Count
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Set mTicker = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                            ByVal timeperiod As Long, _
                            ByVal timeperiodUnit As TradeBuild.TimePeriodUnits, _
                            ByVal key As String, _
                            ByVal NumberOfBarsToFetch As Long, _
                            Optional ByVal SaveIntervalSeconds As Long = -1&, _
                            Optional ByVal IncludeBidAndAsk As Boolean) As Timeframe
Set add = New Timeframe
mTimeframes.add add, key
add.key = key
add.Ticker = mTicker
add.timeperiod = timeperiod
add.timeperiodUnit = timeperiodUnit
add.SaveIntervalSeconds = SaveIntervalSeconds
add.IncludeBidAndAsk = IncludeBidAndAsk
add.NumberOfBarsToFetch = NumberOfBarsToFetch
If Not mContract Is Nothing Then add.Contract = mContract
End Function

Friend Sub Finish()
Dim lTimeframe As TradeBuild.Timeframe
For Each lTimeframe In mTimeframes
    lTimeframe.Finish
Next
Set mTicker = Nothing
End Sub

Public Function Item(index As Variant) As Timeframe
Set Item = mTimeframes(index)
End Function

Friend Sub Reinitialise()
Dim lTimeframe As TradeBuild.Timeframe
For Each lTimeframe In mTimeframes
    lTimeframe.Reinitialise
Next
End Sub

Public Sub Remove(index As Variant)
Dim lTimeframe As TradeBuild.Timeframe
On Error Resume Next
Set lTimeframe = mTimeframes(index)
On Error GoTo 0
If Not lTimeframe Is Nothing Then
    lTimeframe.Finish
    mTimeframes.Remove index
End If
End Sub

Public Sub RemoveAll()
Dim i As Long
Dim lTimeframe As TradeBuild.Timeframe
For i = mTimeframes.Count To 1 Step -1
    Set lTimeframe = mTimeframes(i)
    lTimeframe.Finish
    mTimeframes.Remove i
Next
End Sub

Public Function NewEnum() As IUnknown
Set NewEnum = mTimeframes.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================

