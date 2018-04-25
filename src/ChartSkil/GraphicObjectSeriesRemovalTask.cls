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

Private mRemoveObjectsFromRegion             As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "GraphObjSeriesRemovalTask terminated"
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

Static sEN As Enumerator
If sEN Is Nothing Then Set sEN = mGraphObjSeries.Enumerator

Static sETimer As New ElapsedTimer

sETimer.StartTiming

Do While sEN.MoveNext
    Dim lGraphObj As IGraphicObject
    Set lGraphObj = sEN.Current
    If mRemoveObjectsFromRegion Then mRegion.RemoveGraphicObjectEx lGraphObj, False, False
    sEN.Remove
    If sETimer.ElapsedTimeMicroseconds > 20000 Then Exit Sub
Loop

mGraphObjSeries.Finish
Set mGraphObjSeries = Nothing
Set mRegion = Nothing

mTaskContext.Finish Empty, False

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
                ByVal Region As ChartRegion, _
                ByVal removeObjectsFromRegion As Boolean)
Set mGraphObjSeries = graphObjSeries
If removeObjectsFromRegion Then
    AssertArgument Not Region Is Nothing, "Region is Nothing"
    Set mRegion = Region
    mRemoveObjectsFromRegion = True
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


