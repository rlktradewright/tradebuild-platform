VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DunniganBarFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements BarFormatter

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

Private Const ModuleName                    As String = "DunniganBarFormatter"

'@================================================================================
' Member variables
'@================================================================================

Private mBars As TimeframeUtils26.bars
 
'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' BarFormatter Interface Members
'@================================================================================

Private Sub BarFormatter_FormatBar( _
                ByVal sourceBar As TimeframeUtils26.Bar, _
                ByVal chartBar As ChartSkil26.Bar)

If mBars.Count < 2 Then Exit Sub

If sourceBar.highValue > mBars.Bar(-1).highValue And _
    sourceBar.lowValue >= mBars.Bar(-1).lowValue _
Then
    chartBar.barColor = &H1D9311
    chartBar.upColor = &H1D9311
    chartBar.downColor = &H1D9311
ElseIf sourceBar.highValue <= mBars.Bar(-1).highValue And _
    sourceBar.lowValue < mBars.Bar(-1).lowValue _
Then
    chartBar.barColor = &H43FC2
    chartBar.upColor = &H43FC2
    chartBar.downColor = &H43FC2
ElseIf sourceBar.highValue <= mBars.Bar(-1).highValue And _
    sourceBar.lowValue >= mBars.Bar(-1).lowValue _
Then
    chartBar.barColor = vbMagenta
    chartBar.upColor = vbMagenta
    chartBar.downColor = vbMagenta
ElseIf sourceBar.highValue >= mBars.Bar(-1).highValue And _
    sourceBar.lowValue <= mBars.Bar(-1).lowValue _
Then
    chartBar.barColor = vbBlue
    chartBar.upColor = vbBlue
    chartBar.downColor = vbBlue
End If
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let bars(ByVal value As TimeframeUtils26.bars)
Set mBars = value
End Property
 
'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


