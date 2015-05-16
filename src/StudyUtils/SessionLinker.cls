VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SessionLinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "SessionLinker"

'@================================================================================
' Member variables
'@================================================================================

Private mStudySession                               As Session
Private mSessionToLinkTo                            As Session

Private mSelfRef                                    As Object

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

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

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If ev.ContinuationData = "studysession" Then
    Set mStudySession = ev.Future.Value
ElseIf ev.ContinuationData = "mainsession" Then
    Set mSessionToLinkTo = ev.Future.Value
End If

checkLink

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

Friend Sub Link(ByVal pStudyManager As StudyManager, ByVal pStudy As IStudy)
Const ProcName As String = "Link"
On Error GoTo Err

Dim lSessionFuture As IFuture
Set lSessionFuture = CreateSessionFuture(pStudyManager.GetStudyEntry(pStudy).SessionBuilderFuture)
If lSessionFuture.IsAvailable Then
    Set mStudySession = lSessionFuture.Value
Else
    mFutureWaiter.Add lSessionFuture, "studysession"
End If

If pStudyManager.SessionFuture.IsAvailable Then
    Set mSessionToLinkTo = pStudyManager.SessionFuture.Value
Else
    mFutureWaiter.Add pStudyManager.SessionFuture, "mainsession"
End If

checkLink

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkLink()
Const ProcName As String = "checkLink"
On Error GoTo Err

If mStudySession Is Nothing Or mSessionToLinkTo Is Nothing Then Exit Sub

mStudySession.LinkToSession mSessionToLinkTo
Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



