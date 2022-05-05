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

Private mGroups                                     As Groups

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mNumberOfLivePositionsClosed                As Long
Private mNumberOfLivePositionsNotClosed             As Long

Private mNumberOfSimulatedPositionsClosed           As Long
Private mNumberOfSimulatedPositionsNotClosed        As Long

Private mCloseoutMode                               As CloseoutModes
Private mPriceSpec                                  As PriceSpecifier

Private mLiveErrorMessages                          As String
Private mSimulatedErrorMessages                     As String

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
'mFutureWaiter.DiagnosticID = "CloseoutProcessor"
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
If mLiveErrorMessages <> "" Then gWriteLineToConsole mLiveErrorMessages

If mNumberOfSimulatedPositionsClosed <> 0 Then gWriteLineToConsole "Simulated positions closed: " & mNumberOfSimulatedPositionsClosed, True
If mNumberOfSimulatedPositionsNotClosed <> 0 Then gWriteLineToConsole "Simulated positions not closed: " & mNumberOfSimulatedPositionsNotClosed, True
If mSimulatedErrorMessages <> "" Then gWriteLineToConsole mSimulatedErrorMessages

Set mSelfRef = Nothing

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
Dim lMessage As Variant
If lResult.LiveErrorMessages.Count <> 0 Then
    For Each lMessage In lResult.LiveErrorMessages
        mLiveErrorMessages = mLiveErrorMessages & lMessage & vbCrLf
    Next
End If

mNumberOfSimulatedPositionsClosed = mNumberOfSimulatedPositionsClosed + lResult.NumberOfSimulatedPositionsClosed
mNumberOfSimulatedPositionsNotClosed = mNumberOfSimulatedPositionsNotClosed + lResult.NumberOfSimulatedPositionsNotClosed
If lResult.SimulatedErrorMessages.Count <> 0 Then
    For Each lMessage In lResult.SimulatedErrorMessages
        mSimulatedErrorMessages = mSimulatedErrorMessages & lMessage & vbCrLf
    Next
End If

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

mGroups.DiscardPendingItemsForAllGroups

mFutureWaiter.Add mOrderManager.CloseAllPositions(PositionTypeLive)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub CloseoutGroup( _
                ByVal pGroupName As String)
Const ProcName As String = "CloseoutGroup"
On Error GoTo Err

mGroups.DiscardPendingItemsForGroup pGroupName

mFutureWaiter.Add mOrderManager.CloseGroupPositions( _
                                    pGroupName, _
                                    mCloseoutMode, _
                                    mPriceSpec, _
                                    , _
                                    gBracketOrderListener.Base)
                                    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pOrderManager As OrderManager, _
                ByVal pGroups As Groups, _
                ByVal pCloseoutMode As CloseoutModes, _
                ByVal pPriceSpec As PriceSpecifier)
Set mSelfRef = Me
Set mOrderManager = pOrderManager
Set mGroups = pGroups
mCloseoutMode = pCloseoutMode
Set mPriceSpec = pPriceSpec
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




