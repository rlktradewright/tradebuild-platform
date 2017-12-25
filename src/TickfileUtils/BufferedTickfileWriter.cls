VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BufferedTickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements ITickfileOutputMonitor
Implements ITickfileWriter

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "BufferedTickfileWriter"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mWriter                         As ITickfileWriter

Private mReady                          As Boolean

Private mPendingWrites                  As Collection

Private mPendingClose                   As Boolean

Private mSessionBuilder                 As SessionBuilder
Private WithEvents mSession             As Session
Attribute mSession.VB_VarHelpID = -1

Private mOutputMonitor                  As ITickfileOutputMonitor

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Collection
End Sub

'@================================================================================
' ITickfileOutputMonitor Interface Members
'@================================================================================

Private Sub ITickfileOutputMonitor_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "ITickfileOutputMonitor_NotifyEvent"
On Error GoTo Err

If mOutputMonitor Is Nothing Then Exit Sub

Set ev.Source = Me
mOutputMonitor.NotifyEvent ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyNotReady(ByVal pSource As Object)
Const ProcName As String = "ITickfileOutputMonitor_NotifyNotReady"
On Error GoTo Err

mReady = False

If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyNotReady Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyOutputFileClosed(ByVal pSource As Object)
Const ProcName As String = "ITickfileOutputMonitor_NotifyOutputFileClosed"
On Error GoTo Err

If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyOutputFileClosed Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyOutputFileCreated(ByVal pSource As Object, ByVal pFilename As String)
Const ProcName As String = "ITickfileOutputMonitor_NotifyOutputFileCreated"
On Error GoTo Err

If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyOutputFileCreated Me, pFilename

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyReady(ByVal pSource As Object)
Const ProcName As String = "ITickfileOutputMonitor_NotifyReady"
On Error GoTo Err

mReady = True

setupSession

writePendingitems
If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyReady Me

If mPendingClose Then Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITickfileWriter Interface Members
'@================================================================================

Private Property Get ITickfileWriter_ContractFuture() As IFuture
Set ITickfileWriter_ContractFuture = ContractFuture
End Property

Private Sub ITickfileWriter_Finish()
Const ProcName As String = "ITickfileWriter_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickfileWriter_IsReady() As Boolean
ITickfileWriter_IsReady = IsReady
End Property

Private Sub ITickfileWriter_WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)
Const ProcName As String = "ITickfileWriter_WriteComment"
On Error GoTo Err

WriteComment pComment, pTimestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileWriter_WriteTick(pTick As GenericTick)
Const ProcName As String = "ITickfileWriter_WriteTick"
On Error GoTo Err

WriteTick pTick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mSession Event Handlers
'================================================================================

Private Sub mSession_SessionEnded(ev As SessionEventData)
Const ProcName As String = "mSession_SessionEnded"
On Error GoTo Err

WriteComment "Session ended at " & FormatTimestamp(mSession.CurrentSessionEndTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), mSession.SessionCurrentTime

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mSession_SessionStarted"
On Error GoTo Err

WriteComment "Session started at " & FormatTimestamp(mSession.CurrentSessionStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), mSession.SessionCurrentTime

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractFuture() As IFuture
Set ContractFuture = mWriter.ContractFuture
End Property

Public Property Get NumberOfPendingWrites() As Long
Const ProcName As String = "NumberOfPendingWrites"
On Error GoTo Err

NumberOfPendingWrites = mPendingWrites.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    WriteComment "Tick recording finished at " & FormatTimestamp(Now, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), mSession.SessionCurrentTime
    gLogger.Log "Closing output tickfile", ProcName, ModuleName
    mWriter.Finish
    mPendingClose = False
    mReady = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pWriter As ITickfileWriter)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOutputMonitor = pOutputMonitor
Set mWriter = pWriter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)
Const ProcName As String = "WriteComment"
On Error GoTo Err

If Not mReady Then
    Dim lComment As TickWriterPendingComment
    lComment.Comment = pComment
    lComment.Timestamp = pTimestamp
    mPendingWrites.Add lComment
Else
    mWriter.WriteComment pComment, pTimestamp
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub WriteTick( _
                pTick As GenericTick)
Const ProcName As String = "WriteTick"
On Error GoTo Err

Static sFirstTickReceived As Boolean

If Not mReady Then
    mPendingWrites.Add pTick
Else
    If Not sFirstTickReceived Then
        sFirstTickReceived = True
        WriteComment "Tick recording started at " & FormatTimestamp(Now, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), pTick.Timestamp
    End If
    mSessionBuilder.SetSessionCurrentTime pTick.Timestamp
    mWriter.WriteTick pTick
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setupSession()
Const ProcName As String = "setupSession"
On Error GoTo Err

Dim lContract As IContract
Set lContract = mWriter.ContractFuture.Value
Set mSessionBuilder = CreateSessionBuilder(lContract.sessionStartTime, lContract.sessionEndTime, Nothing)
Set mSession = mSessionBuilder.Session

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub writePendingitems()
Const ProcName As String = "writePendingitems"
On Error GoTo Err

If mPendingWrites.Count = 0 Then Exit Sub

Dim lItem As Variant
For Each lItem In mPendingWrites
    If TypeName(lItem) = "TickWriterPendingComment" Then
        Dim lComment As TickWriterPendingComment
        lComment = lItem
        WriteComment lComment.Comment, lComment.Timestamp
    Else
        Dim lTick As GenericTick
        lTick = lItem
        WriteTick lTick
    End If
Next

Set mPendingWrites = New Collection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
