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

Private Const ProjectName                   As String = "ChartSkil26"
Private Const ModuleName                    As String = "GraphicObjectSeriesRemovalTask"

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
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Dim graphObj As IGraphicObject
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Do While mGraphObjSeries.Count <> 0

    failpoint = 200

    Set graphObj = mGraphObjSeries.Item(1)
    
    failpoint = 300
    
    If mRemoveObjectsFromRegion Then mRegion.RemoveGraphicObjectEx graphObj, False, False
    
    failpoint = 400
    
    mGraphObjSeries.Remove 1
    i = i + 1
    If i = 100 Then
        Exit Sub
    End If
Loop

mGraphObjSeries.Finish
Set mGraphObjSeries = Nothing
Set mRegion = Nothing

failpoint = 500

mTaskContext.Finish Empty, False

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Task_run" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
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

Friend Sub initialise( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                ByVal Region As ChartRegion, _
                ByVal removeObjectsFromRegion As Boolean)
Set mGraphObjSeries = graphObjSeries
Set mRegion = Region
mRemoveObjectsFromRegion = removeObjectsFromRegion
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


