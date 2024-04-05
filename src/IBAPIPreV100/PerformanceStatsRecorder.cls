VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerformanceStatsRecorder"
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

Private Type StatsEntry
    LastSecondCount         As Long
    LastSecondTime          As Single
    LastPeriodCount         As Long
    LastPeriodTime          As Single
    TotalCount              As Long
    TotalTime               As Single
    MaxSecondCount          As Long
    LongestTime             As Single
    ShortestTime            As Single
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "PerformanceStatsRecorder"

'@================================================================================
' Member variables
'@================================================================================

Private mPerformanceStats(MAX_SOCKET_INMSG)         As StatsEntry

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

Friend Sub AccumulateStats()
Const ProcName As String = "AccumulateStats"
Dim i As Long

On Error GoTo Err

For i = 0 To TwsSocketInMsgTypes.MAX_SOCKET_INMSG
    With mPerformanceStats(i)
        .LastPeriodCount = .LastPeriodCount + .LastSecondCount
        .LastPeriodTime = .LastPeriodTime + .LastSecondTime
        If .LastSecondTime > .MaxSecondCount Then .MaxSecondCount = .LastSecondCount
        .TotalCount = .TotalCount + .LastSecondCount
        .TotalTime = .TotalTime + .LastSecondTime
        .LastSecondCount = 0
        .LastSecondTime = 0
    End With
Next
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Function GenerateStats() As String
Const ProcName As String = "GenerateStats"
On Error GoTo Err

Dim i As Long
Dim lSb As StringBuilder

Set lSb = CreateStringBuilder("Message type          Last  Last avg    Total  Total avg Max/sec   Longest  Shortest" & vbCrLf)

For i = 0 To TwsSocketInMsgTypes.MAX_SOCKET_INMSG
    With mPerformanceStats(i)
        .LastPeriodCount = .LastPeriodCount + .LastSecondCount
        .LastPeriodTime = .LastPeriodTime + .LastSecondTime
        .TotalCount = .TotalCount + .LastSecondCount
        .TotalTime = .TotalTime + .LastSecondTime
        
        If .TotalCount <> 0 Then
            lSb.Append gSocketInMsgTypeToString(i)
            lSb.Append Format(.LastPeriodCount, " 00000")
            If .LastPeriodCount <> 0 Then
                lSb.Append Format(.LastPeriodTime / .LastPeriodCount, " 0000000.0")
            Else
                lSb.Append "       0.0"
            End If
            lSb.Append Format(.TotalCount, " 00000000")
            If .TotalCount <> 0 Then
                lSb.Append Format(.TotalTime / .TotalCount, "  0000000.0")
            Else
                lSb.Append "        0.0"
            End If
            lSb.Append Format(.MaxSecondCount, "    0000")
            lSb.Append Format(.LongestTime, " 0000000.0")
            lSb.AppendLine Format(.ShortestTime, " 0000000.0")
        End If
        
        .LastPeriodCount = 0
        .LastPeriodTime = 0
        .LastSecondCount = 0
        .LastSecondTime = 0
        .MaxSecondCount = 0
    End With
Next

GenerateStats = lSb.ToString

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Friend Sub UpdateMessageTypeStats( _
                ByVal pMessageId As TwsSocketInMsgTypes, _
                ByVal pEt As Single)
Const ProcName As String = "UpdateMessageTypeStats"
On Error GoTo Err

With mPerformanceStats(pMessageId)
    .LastSecondCount = .LastSecondCount + 1
    .LastSecondTime = .LastSecondTime + pEt
    If pEt > .LongestTime Then .LongestTime = pEt
    If pEt < .ShortestTime Or .ShortestTime = 0 Then .ShortestTime = pEt
End With

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




