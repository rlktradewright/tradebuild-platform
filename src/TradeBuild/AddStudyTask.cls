VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddStudyTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements Task

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTaskContext As TaskContext

Private mBars As bars

Private mStudy As TradeBuildSP.IStudy
Private mStudyToNotify As TradeBuildSP.IStudy
Private mValueNames() As String
Private mNotifyNames() As String

Private mNextBarNumber As Long
Private mBarsNotified As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim lBar As Bar
Dim i As Long
Dim ev As TradeBuildSP.StudyValueEvent
Dim valueName As String

If mBars.Count = 0 Then
    mBars.studies.add mStudy, mValueNames
    mTaskContext.finish
    Set mTaskContext = Nothing
    Exit Sub
End If

Set ev.source = mBars

Set lBar = mBars(mNextBarNumber)

If Not lBar.Blank Then
    For i = 0 To UBound(mValueNames)
        valueName = mValueNames(i)
        ev.barNumber = mNextBarNumber
        ev.valueName = mNotifyNames(i)
        
        If valueName = BarsValueClose Then
            ev.timestamp = lBar.datetime
            ev.value = lBar.openValue
            mStudyToNotify.notify ev
        
            If lBar.IsDownBar Then
                ev.timestamp = lBar.datetime + (lBar.barLength / 3) / 1440
                ev.value = lBar.highValue
                mStudyToNotify.notify ev
                
                ev.timestamp = lBar.datetime + (2 * lBar.barLength / 3) / 1440
                ev.value = lBar.lowValue
                mStudyToNotify.notify ev
            Else
                ev.timestamp = lBar.datetime + (lBar.barLength / 3) / 1440
                ev.value = lBar.lowValue
                mStudyToNotify.notify ev

                ev.timestamp = lBar.datetime + (2 * lBar.barLength / 3) / 1440
                ev.value = lBar.highValue
                mStudyToNotify.notify ev
            End If
            
            ev.timestamp = lBar.datetime + lBar.barLength / 1440 - OneSecond
            ev.value = lBar.closeValue
            mStudyToNotify.notify ev
        ElseIf valueName = BarsValueVolume Then
            ev.timestamp = lBar.datetime
            ev.value = lBar.volume
            mStudyToNotify.notify ev
        ElseIf valueName = BarsValueTickVolume Then
            ev.timestamp = lBar.datetime
            ev.value = lBar.tickVolume
            mStudyToNotify.notify ev
        End If
    Next
End If

mNextBarNumber = mNextBarNumber + 1
mBarsNotified = mBarsNotified + 1

If mNextBarNumber = mBars.Count Then
    mBars.studies.add mStudyToNotify, mValueNames
    mTaskContext.finish
    Set mTaskContext = Nothing
End If

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal bars As bars, _
                ByVal study As TradeBuildSP.IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long)
Dim numBarsToNotify As Long
Dim inputDefs As TradeBuildSP.IStudyInputDefinitions
Dim i As Long

Set mBars = bars
Set mStudy = study
Set mStudyToNotify = mStudy.baseStudy
mValueNames = valueNames

Set inputDefs = study.studyDefinition.studyInputDefinitions
ReDim mNotifyNames(inputDefs.Count - 1) As String
For i = 0 To inputDefs.Count - 1
    mNotifyNames(i) = inputDefs.Item(i + 1).name
Next

If numUnderlyingValuesToUse < 0 Then
    ' just replay the number of values needed by the study
    numBarsToNotify = study.numberOfBarsRequired
Else
    numBarsToNotify = numUnderlyingValuesToUse
End If
If numBarsToNotify > mBars.Count Then numBarsToNotify = mBars.Count

mNextBarNumber = mBars.Count - numBarsToNotify + 1

End Sub

'================================================================================
' Helper Functions
'================================================================================

