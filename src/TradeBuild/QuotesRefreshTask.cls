VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuotesRefreshTask"
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

Implements Task

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

Private Const ModuleName                            As String = "QuotesRefreshTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext
Private mTickers                                    As Tickers
Private mIndex                                      As Long
Private mCancelled                                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()
mCancelled = True
End Sub

Private Sub Task_run()

If mCancelled Then
    mTaskContext.finish Empty, True
    Exit Sub
End If

mIndex = mIndex + 1
If mIndex > mTickers.Count Then
    mTaskContext.finish Empty, False
    Exit Sub
End If

mTickers.item(i).RefreshQuotes
mTickers.item(i).RefreshPriceChange

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
Set mTickers = mTaskContext.data
End Property

Private Property Get Task_TaskName() As String

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


