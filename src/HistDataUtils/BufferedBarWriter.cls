VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BufferedBarWriter"
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

Implements IBarOutputMonitor
Implements IBarWriter

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

Private Const ModuleName                            As String = "BufferedBarWriter"

'@================================================================================
' Member variables
'@================================================================================

Private mWriter                                     As IBarWriter
Private mReady                                      As Boolean

Private mPendingWrites                              As Collection

Private mPendingClose                               As Boolean

Private mSessionBuilder                             As SessionBuilder
Private WithEvents mSession                         As Session
Attribute mSession.VB_VarHelpID = -1

Private mContractFuture                             As IFuture
Private mContract                                   As IContract

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mOutputMonitor                              As IBarOutputMonitor

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Set mPendingWrites = New Collection
End Sub

'@================================================================================
' IBarOutputMonitor Interface Members
'@================================================================================

Private Sub IBarOutputMonitor_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "IBarOutputMonitor_NotifyEvent"
On Error GoTo Err

If mOutputMonitor Is Nothing Then Exit Sub

Set ev.Source = Me
mOutputMonitor.NotifyEvent ev

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyNotReady(ByVal pSource As Object)
Const ProcName As String = "IBarOutputMonitor_NotifyNotReady"
On Error GoTo Err

mReady = False

If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyNotReady Me

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyOutputFileClosed(ByVal pSource As Object)
Const ProcName As String = "IBarOutputMonitor_NotifyOutputFileClosed"
On Error GoTo Err

If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyOutputFileClosed Me

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyOutputFileCreated(ByVal pSource As Object, ByVal pFilename As String)
Const ProcName As String = "IBarOutputMonitor_NotifyOutputFileCreated"
On Error GoTo Err

If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyOutputFileCreated Me, pFilename

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyReady(ByVal pSource As Object)
Const ProcName As String = "IBarOutputMonitor_NotifyReady"
On Error GoTo Err

mReady = True

Dim lItem As Variant
For Each lItem In mPendingWrites
    If TypeName(lItem) = "BarWriterPendingComment" Then
        Dim lComment As BarWriterPendingComment
        lComment = lItem
        mWriter.WriteComment lComment.Comment, lComment.Timestamp
    Else
        Dim lPendingBar As BarWriterPendingBar
        lPendingBar = lItem
        mWriter.WriteBar lPendingBar.Bar, lPendingBar.BarType, lPendingBar.TimePeriod
    End If
Next

Set mPendingWrites = New Collection
If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyReady Me

If mPendingClose Then Finish

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IBarWriter Interface Members
'@================================================================================

Private Property Get IBarWriter_ContractFuture() As IFuture
Set IBarWriter_ContractFuture = ContractFuture
End Property

Private Sub IBarWriter_Finish()
Const ProcName As String = "IBarWriter_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBarWriter_IsReady() As Boolean
IBarWriter_IsReady = IsReady
End Property

Private Sub IBarWriter_WriteBar(ByVal pBar As Bar, ByVal pBarType As BarTypes, ByVal pTimePeriod As TimePeriod)
Const ProcName As String = "IBarWriter_WriteBar"
On Error GoTo Err

WriteBar pBar, pBarType, pTimePeriod

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarWriter_WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)
Const ProcName As String = "IBarWriter_WriteComment"
On Error GoTo Err

WriteComment pComment, pTimestamp

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    setupSession
End If

Exit Sub

Err:
GHistData.NotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mSession Event Handlers
'================================================================================

Private Sub mSession_SessionEnded(ev As SessionEventData)
Const ProcName As String = "mSession_SessionEnded"
On Error GoTo Err

WriteComment "Session ended at " & FormatTimestamp(mSession.CurrentSessionEndTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), mSession.CurrentSessionEndTime

Exit Sub

Err:
GHistData.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mSession_SessionStarted"
On Error GoTo Err

WriteComment "Session started at " & FormatTimestamp(mSession.CurrentSessionStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), mSession.CurrentSessionEndTime

Exit Sub

Err:
GHistData.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Public Property Get NumberOfPendingWrites() As Long
Const ProcName As String = "NumberOfPendingWrites"
On Error GoTo Err

NumberOfPendingWrites = mPendingWrites.Count

Exit Property

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get IsReady() As Boolean
IsReady = mReady
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mReady Then
    mPendingClose = True
Else
    WriteComment "Bar writing finished", mSession.CurrentSessionEndTime
    GHistData.Logger.Log "Closing output file", ProcName, ModuleName
    mWriter.Finish
    mPendingClose = False
    mReady = False
End If

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pWriter As IBarWriter, _
                ByVal pContractFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOutputMonitor = pOutputMonitor
Set mWriter = pWriter
Set mContractFuture = pContractFuture
mFutureWaiter.Add mContractFuture

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)
Const ProcName As String = "WriteComment"
On Error GoTo Err

If Not mReady Then
    Dim lComment As BarWriterPendingComment
    lComment.Comment = pComment
    lComment.Timestamp = pTimestamp
    mPendingWrites.Add lComment
Else
    mSessionBuilder.SetSessionCurrentTime pTimestamp
    mWriter.WriteComment pComment, pTimestamp
End If

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub WriteBar( _
                ByVal pBar As Bar, _
                ByVal pBarType As BarTypes, _
                ByVal pTimePeriod As TimePeriod)
Const ProcName As String = "WriteBar"
On Error GoTo Err

If Not mReady Then
    Dim lBar As BarWriterPendingBar
    Set lBar.Bar = pBar
    lBar.BarType = pBarType
    Set lBar.TimePeriod = pTimePeriod
    mPendingWrites.Add lBar
Else
    mSessionBuilder.SetSessionCurrentTime pBar.Timestamp
    mWriter.WriteBar pBar, pBarType, pTimePeriod
End If

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setupSession()
Const ProcName As String = "setupSession"
On Error GoTo Err

Set mSessionBuilder = CreateSessionBuilder( _
                                mContract.SessionStartTime, _
                                mContract.SessionEndTime, _
                                Nothing)
Set mSession = mSessionBuilder.Session

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub





