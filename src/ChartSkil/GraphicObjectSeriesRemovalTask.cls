VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphObjSeriesRemovalTask"
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

Implements ITask

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

Private Const ModuleName                    As String = "GraphObjSeriesRemovalTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext

Private mGraphObjSeries                     As IGraphicObjectSeries
Private mRegion                             As ChartRegion

'@================================================================================
' Class Event Handlers
'@================================================================================

'Private Sub Class_Terminate()
'Debug.Print "GraphObjSeriesRemovalTask terminated"
'End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

Static sTotalNumberReleased As Long
Static sYieldCount As Long
Static sClearStartTime As Date
Static sEN As Enumerator
If sEN Is Nothing Then Set sEN = mGraphObjSeries.Enumerator

Dim results As Parameters

If sClearStartTime = 0 Then sClearStartTime = GetTimestamp

If mTaskContext.CancelPending Then
    mGraphObjSeries.Finish
    gLogger.Log "Released (cancelled) " & sTotalNumberReleased & " entries in " & sYieldCount & " yields", ProcName, ModuleName, LogLevelHighDetail
    Set results = generateResults(False, sTotalNumberReleased, sYieldCount, sClearStartTime)
    mTaskContext.Finish results, True
    Exit Sub
End If

Static sNumberToRelease As Long

Static sLoopCount As Long

Const HeuristicNumber As Long = 50
Const HeuristicTime As Long = 5

If sLoopCount Mod 50 = 0 Then
    sNumberToRelease = calibrate(sEN, HeuristicNumber, HeuristicTime, sTotalNumberReleased, sYieldCount)
ElseIf Not releaseEntries(sEN, sNumberToRelease, sTotalNumberReleased, sYieldCount) Then
    sNumberToRelease = 0
End If

If sNumberToRelease = 0 Then
    mGraphObjSeries.Finish
    gLogger.Log "Released " & sTotalNumberReleased & " entries in " & sYieldCount & " yields", ProcName, ModuleName, LogLevelHighDetail
    Set results = generateResults(False, sTotalNumberReleased, sYieldCount, sClearStartTime)
    mTaskContext.Finish results, False
End If

sLoopCount = sLoopCount + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
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

Friend Sub Initialise( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                ByVal Region As ChartRegion)
Set mGraphObjSeries = graphObjSeries
AssertArgument Not Region Is Nothing, "Region is Nothing"
Set mRegion = Region
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calibrate( _
                ByVal pEn As Enumerator, _
                ByVal pNumberToRelease As Long, _
                ByVal pHeuristicTime As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Long
Const ProcName As String = "calibrate"
On Error GoTo Err

Dim lPrevCpuTime As Double: lPrevCpuTime = GetCurrentProcessCpuTime
Dim lAligned As Boolean

Dim lNumberReleased As Long
Do While releaseEntries(pEn, pNumberToRelease, pTotalNumberReleased, pYieldCount)
    lNumberReleased = lNumberReleased + pNumberToRelease
    Dim lCurrCpuTime As Double: lCurrCpuTime = GetCurrentProcessCpuTime
    Dim lCpuTimeMillisecs As Double: lCpuTimeMillisecs = 1000 * (lCurrCpuTime - lPrevCpuTime)
    If lCpuTimeMillisecs > 0 Then
        If Not lAligned Then
            lNumberReleased = 0
            lPrevCpuTime = lCurrCpuTime
            lAligned = True
        Else
            gLogger.Log "Time to release " & lNumberReleased & " entries: " & lCpuTimeMillisecs, ProcName, ModuleName, LogLevelHighDetail
            calibrate = CLng(pHeuristicTime / lCpuTimeMillisecs * lNumberReleased)
            gLogger.Log "Number of entries to release per yield: " & calibrate, ProcName, ModuleName, LogLevelHighDetail
            Exit Function
        End If
    End If
Loop

calibrate = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateResults( _
                ByVal pCancelled As Boolean, _
                ByVal pNumberReleased As Long, _
                ByVal pNumberOfYields As Long, _
                ByVal pClearStartTime As Date) As Parameters
Const ProcName As String = "generateResults"
On Error GoTo Err

Dim results As New Parameters
results.SetParameterValue "Cancelled", pCancelled
results.SetParameterValue "Number of entries released", pNumberReleased
results.SetParameterValue "Number of timeslots", pNumberOfYields

Dim lDuration As Double: lDuration = 86400# * (GetTimestamp - pClearStartTime)
If lDuration < 0.001 Then
    results.SetParameterValue "Duration (microsecs)", Format(lDuration * 1000000#, "0")
ElseIf lDuration < 1# Then
    results.SetParameterValue "Duration (millisecs)", Format(lDuration * 1000#, "0.000")
Else
    results.SetParameterValue "Duration (secs)", Format(lDuration, "0.000")
End If

Set generateResults = results

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function releaseEntries( _
                ByVal pEn As Enumerator, _
                ByVal pNumberToRelease As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Boolean
Const ProcName As String = "releaseEntries"
On Error GoTo Err

Dim lNumberReleased As Long

Do While pEn.MoveNext
    Dim lGraphObj As IGraphicObject
    Set lGraphObj = pEn.Current
    mRegion.RemoveGraphicObjectEx lGraphObj, False, False
    lNumberReleased = lNumberReleased + 1
    If lNumberReleased >= pNumberToRelease Then
        releaseEntries = True
        pTotalNumberReleased = pTotalNumberReleased + lNumberReleased
        pYieldCount = pYieldCount + 1
        Exit Function
    End If
Loop

pTotalNumberReleased = pTotalNumberReleased + lNumberReleased
pYieldCount = pYieldCount + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


