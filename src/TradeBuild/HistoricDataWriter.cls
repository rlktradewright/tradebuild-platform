VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoricDataWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private mWriter As TradeBuildSP.IBarDataWriter
Private mServiceProviderWriterReady As Boolean

Private mPendingWrites As Collection


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Collection
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

Friend Property Let ServiceProviderWriter(ByVal value As TradeBuildSP.IBarDataWriter)
Set mWriter = value
End Property

Friend Property Let ServiceProviderWriterReady(ByVal value As Boolean)
Dim pendingWrite As Bar

mServiceProviderWriterReady = value

If mPendingWrites.Count <> 0 Then
    For Each pendingWrite In mPendingWrites
        WriteBar pendingWrite
    Next
End If

Set mPendingWrites = Nothing
Set mPendingWrites = New Collection
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub ReleaseDataStore()
mWriter.ReleaseDataStore
End Sub

Friend Sub WriteBar(ByVal theBar As TradeBuild.Bar)
Dim spBar As TradeBuildSP.Bar

If Not mServiceProviderWriterReady Then
    mPendingWrites.add theBar
Else
    With spBar
        .BarType = theBar.BarType
        .ClosePrice = theBar.ClosePrice
        .HighPrice = theBar.HighPrice
        .LowPrice = theBar.LowPrice
        .OpenPrice = theBar.OpenPrice
        .periodMinutes = theBar.BarLength
        .TickVolume = theBar.TickVolume
        .timestamp = theBar.DateTime
        .Volume = theBar.Volume
    End With
    mWriter.WriteBar spBar
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

