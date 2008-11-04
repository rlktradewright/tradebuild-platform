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

Private mTaskContext As TaskContext

Private mGraphObjSeries As IGraphicObjectSeries
Private mRegion As ChartRegion

'@================================================================================
' Class Event Handlers
'@================================================================================

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

    Set graphObj = mGraphObjSeries.Item(mGraphObjSeries.Count)
    
    failpoint = 300
    
    mRegion.RemoveGraphicObjectEx graphObj, False, False
    
    failpoint = 400
    
    mGraphObjSeries.Remove mGraphObjSeries.Count
    i = i + 1
    If i = 100 Then Exit Sub
Loop

mGraphObjSeries.Finish

failpoint = 500

mTaskContext.Finish Empty, False

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Task_run" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
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
                ByVal Region As ChartRegion)
Set mGraphObjSeries = graphObjSeries
Set mRegion = Region
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


