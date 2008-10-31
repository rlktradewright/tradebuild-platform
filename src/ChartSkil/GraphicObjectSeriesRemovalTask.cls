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

Do While mGraphObjSeries.Count <> 0
    Set graphObj = mGraphObjSeries.Item(mGraphObjSeries.Count)
    mRegion.RemoveGraphicObjectEx graphObj, False, False
    mGraphObjSeries.Remove mGraphObjSeries.Count
    i = i + 1
    If i = 100 Then Exit Sub
Loop

mGraphObjSeries.Finish

mTaskContext.Finish Empty, False
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


