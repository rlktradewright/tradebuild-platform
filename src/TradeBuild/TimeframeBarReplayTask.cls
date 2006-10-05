VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeframeBarReplayTask"
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

Implements Task

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

Private mTaskContext As TaskContext

Private mBars As TradeBuild.bars
Private mMinIndex As Long

Private mItemsProcessed As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim i As Long

If mBars Is Nothing Then
    mTaskContext.finish
    Set mTaskContext = Nothing
    Exit Sub
End If

For i = (mMinIndex + mItemsProcessed) To _
        (mMinIndex + mItemsProcessed + 19)
    
    mBars.replayBar i
    
    If i = mBars.Count Then
        mTaskContext.finish
        Set mTaskContext = Nothing
        Set mBars = Nothing
        Exit Sub
    End If
    
    mItemsProcessed = mItemsProcessed + 1
    
Next

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal pTimeframe As Timeframe, _
                ByVal barType As TradeBuild.BarTypes, _
                ByVal barsToReplay As Long)
Dim lTimeframe As Timeframe
Dim numBars As Long

Set lTimeframe = pTimeframe

Select Case barType
Case BarTypeTrade
    Set mBars = lTimeframe.TradeBars
Case BarTypeBid
    Set mBars = lTimeframe.BidBars
Case BarTypeAsk
    Set mBars = lTimeframe.AskBars
End Select
    
If Not mBars Is Nothing Then numBars = mBars.Count

If barsToReplay <= 0 Then
    mMinIndex = 1
ElseIf numBars >= barsToReplay Then
    mMinIndex = mBars.Count - barsToReplay + 1
Else
    mMinIndex = 1
End If

End Sub

'================================================================================
' Helper Functions
'================================================================================









