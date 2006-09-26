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

Private mTradeBuildChart As TradeBuildChart
Private mMinIndex As Long
Private mMaxIndex As Long

Private mItemsProcessed As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim i As Long

For i = (mMinIndex + mItemsProcessed) To _
        (mMinIndex + mItemsProcessed + 19)
    
    mTradeBuildChart.replaybar i
    
    If i = mMaxIndex Then
        mTaskContext.finish
        Set mTaskContext = Nothing
        Set mTradeBuildChart = Nothing
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

Friend Sub initialise( _
                ByVal pTradeBuildChart As TradeBuildChart, _
                ByVal minIndex As Long, _
                ByVal maxIndex As Long)
Set mTradeBuildChart = pTradeBuildChart
mMinIndex = minIndex
mMaxIndex = maxIndex
End Sub

'================================================================================
' Helper Functions
'================================================================================







