VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CloseoutProcessor"
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

Private Const ModuleName                            As String = "CloseoutProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderManager                               As New OrderManager

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mNumberOfLivePositionsClosed                As Long
Private mNumberOfLivePositionsNotClosed             As Long

Private mNumberOfSimulatedPositionsClosed           As Long
Private mNumberOfSimulatedPositionsNotClosed        As Long


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

If mNumberOfLivePositionsClosed <> 0 Then gWriteLineToConsole "Live positions closed: " & mNumberOfLivePositionsClosed, True
If mNumberOfLivePositionsNotClosed <> 0 Then gWriteLineToConsole "Live positions not closed: " & mNumberOfLivePositionsNotClosed, True
If mNumberOfSimulatedPositionsClosed <> 0 Then gWriteLineToConsole "Simulated positions closed: " & mNumberOfSimulatedPositionsClosed, True
If mNumberOfSimulatedPositionsNotClosed <> 0 Then gWriteLineToConsole "Simulated positions not closed: " & mNumberOfSimulatedPositionsNotClosed, True

mNumberOfLivePositionsClosed = 0
mNumberOfLivePositionsNotClosed = 0
mNumberOfSimulatedPositionsClosed = 0
mNumberOfSimulatedPositionsNotClosed = 0

gInputPaused = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Dim lResult As ClosePositionsResult: Set lResult = ev.Future.Value

mNumberOfLivePositionsClosed = mNumberOfLivePositionsClosed + lResult.NumberOfLivePositionsClosed
mNumberOfLivePositionsNotClosed = mNumberOfLivePositionsNotClosed + lResult.NumberOfLivePositionsNotClosed
mNumberOfSimulatedPositionsClosed = mNumberOfSimulatedPositionsClosed + lResult.NumberOfSimulatedPositionsClosed
mNumberOfSimulatedPositionsNotClosed = mNumberOfSimulatedPositionsNotClosed + lResult.NumberOfSimulatedPositionsNotClosed

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub CloseoutAll()
Const ProcName As String = "CloseoutAll"
On Error GoTo Err

Dim lVar As Variant
For Each lVar In mOrderManager.GetGroupNames
    Dim lGroupName As String: lGroupName = lVar
    CloseoutGroup lGroupName
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub CloseoutGroup(ByVal pGroupName As String)
Const ProcName As String = "CloseoutGroup"
On Error GoTo Err

AssertArgument mOrderManager.GetGroupNames.Contains(pGroupName), "No such group"
mFutureWaiter.Add mOrderManager.CloseGroupPositions(pGroupName)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise(ByVal pOrderManager As OrderManager)
Set mOrderManager = pOrderManager
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




