VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SessionUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum SessionChangeTypes
    SessionChangeNone
    SessionChangeEnd
    SessionChangeStart
    SessionChangeDateChange
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type SessionEventData
    Source              As Object
    Timestamp           As Date
    changeType          As SessionChangeTypes
End Type

Public Type SessionTimes
    StartTime       As Date
    EndTime         As Date
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "SessionUtils"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function CreateSessionBuilder( _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                Optional ByVal pTimeZone As TimeZone, _
                Optional ByVal pInitialSessionTime As Date) As SessionBuilder
Const ProcName As String = "CreateSessionBuilder"
On Error GoTo Err

Set CreateSessionBuilder = GSessionUtils.CreateSessionBuilder(pSessionStartTime, pSessionEndTime, pTimeZone, pInitialSessionTime)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionBuilderFuture(ByVal pSessionFuture As IFuture) As IFuture
Const ProcName As String = "CreateSessionBuilderFuture"
On Error GoTo Err

Set CreateSessionBuilderFuture = GSessionUtils.CreateSessionBuilderFuture(pSessionFuture)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionFuture(ByVal pSessionBuilderFuture As IFuture) As IFuture
Const ProcName As String = "CreateSessionFuture"
On Error GoTo Err

Set CreateSessionFuture = GSessionUtils.CreateSessionFuture(pSessionBuilderFuture)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetOffsetSessionTimes"
On Error GoTo Err

GetOffsetSessionTimes = GSessionUtils.GetOffsetSessionTimes(Timestamp, _
                            offset, _
                            StartTime, _
                            EndTime)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetOffsetSessionTimesIgnoringWeekend( _
                ByVal Timestamp As Date, _
                ByVal offset As Long, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetOffsetSessionTimesIgnoringWeekend"
On Error GoTo Err

GetOffsetSessionTimesIgnoringWeekend = GSessionUtils.GetOffsetSessionTimesIgnoringWeekend(Timestamp, _
                            offset, _
                            StartTime, _
                            EndTime)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSessionTimes( _
                ByVal Timestamp As Date, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetSessionTimes"
On Error GoTo Err

GetSessionTimes = GSessionUtils.GetSessionTimes(Timestamp, _
                            StartTime, _
                            EndTime)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSessionTimesIgnoringWeekend( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As SessionTimes
Const ProcName As String = "GetSessionTimesIgnoringWeekend"
On Error GoTo Err

GetSessionTimesIgnoringWeekend = GSessionUtils.GetSessionTimesIgnoringWeekend(pTimestamp, pSessionStartTime, pSessionEndTime)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function


'@================================================================================
' Helper Functions
'@================================================================================


