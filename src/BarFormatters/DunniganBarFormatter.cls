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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Const ProcName As String = "BarFormatter_FormatBar"
On Error GoTo Err

If sourceBar.BarNumber <= 2 Or mBars.Count < 2 Then Exit Sub

If sourceBar.HighValue > mBars.Bar(sourceBar.BarNumber - 1).HighValue And _
    sourceBar.LowValue >= mBars.Bar(sourceBar.BarNumber - 1).LowValue _
Then
    chartBar.Color = &H1D9311
    chartBar.UpColor = &H1D9311
    chartBar.DownColor = &H1D9311
ElseIf sourceBar.HighValue <= mBars.Bar(sourceBar.BarNumber - 1).HighValue And _
    sourceBar.LowValue < mBars.Bar(sourceBar.BarNumber - 1).LowValue _
Then
    chartBar.Color = &H43FC2
    chartBar.UpColor = &H43FC2
    chartBar.DownColor = &H43FC2
ElseIf sourceBar.HighValue <= mBars.Bar(sourceBar.BarNumber - 1).HighValue And _
    sourceBar.LowValue >= mBars.Bar(sourceBar.BarNumber - 1).LowValue _
Then
    chartBar.Color = vbMagenta
    chartBar.UpColor = vbMagenta
    chartBar.DownColor = vbMagenta
ElseIf sourceBar.HighValue >= mBars.Bar(sourceBar.BarNumber - 1).HighValue And _
    sourceBar.LowValue <= mBars.Bar(sourceBar.BarNumber - 1).LowValue _
Then
    chartBar.Color = vbBlue
    chartBar.UpColor = vbBlue
    chartBar.DownColor = vbBlue
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let bars(ByVal value As TimeframeUtils26.bars)
Const ProcName As String = "bars"
On Error GoTo Err

Set mBars = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property
 
'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


