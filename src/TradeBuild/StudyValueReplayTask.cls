VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueReplayTask"
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
Implements TradeBuildSP.IValueReplayTask

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

Private mTarget As Object
Private mTargetIsStudy As Boolean

Private mTargetStudy As TradeBuildSP.IStudy
Private mStudyToNotify As TradeBuildSP.IStudy
Private mTargetListener As TradeBuildSP.IStudyValueListener
Private mSourceStudy As TradeBuildSP.IStudy
Private mDiscriminator As Long

Private mNextValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String
Private mValueNames() As String
Private mNotifyNames() As String

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IValueReplayTask Interface Members
'================================================================================

Private Property Get IValueReplayTask_discriminator() As Long
IValueReplayTask_discriminator = mDiscriminator
End Property

Private Property Get IValueReplayTask_sourceStudy() As TradeBuildSP.IStudy
Set IValueReplayTask_sourceStudy = mSourceStudy
End Property

Private Property Get IValueReplayTask_target() As Object
Set IValueReplayTask_target = mTarget
End Property

Private Property Get IValueReplayTask_ValueName() As String
IValueReplayTask_ValueName = mValueName
End Property

Private Property Get IValueReplayTask_valueNames() As String()
IValueReplayTask_valueNames = mValueNames
End Property

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim i As Long
Dim ev As TradeBuildSP.StudyValueEvent
Dim valueEntry As TradeBuildSP.StudyValue

If mSourceStudy.numberOfCachedValues = 0 Then
    mTaskContext.finish
    Set mTaskContext = Nothing
    Exit Sub
End If

For i = 0 To UBound(mValueNames)
    valueEntry = mSourceStudy.getStudyValue(mValueNames(i), mNextValueNumber)
     
    If Not IsEmpty(valueEntry.value) Then
        Set ev.source = mSourceStudy
        ev.timestamp = valueEntry.timestamp
        ev.barNumber = valueEntry.barNumber
        ev.value = valueEntry.value
        
        If mTargetIsStudy Then
            ev.valueName = mNotifyNames(i)
            mStudyToNotify.notify ev
        Else
            ev.valueName = mValueNames(i)
            mTargetListener.notify ev
        End If
    End If
Next

mNextValueNumber = mNextValueNumber + 1
mValuesNotified = mValuesNotified + 1
If mNextValueNumber = mSourceStudy.numberOfCachedValues Then
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

Public Property Get discriminator() As Long
discriminator = mDiscriminator
End Property

Public Property Get sourceStudy() As TradeBuildSP.IStudy
Set sourceStudy = mSourceStudy
End Property

Public Property Get target() As Object
Set target = mTarget
End Property

Public Property Get valueName() As String
valueName = mValueName
End Property

Public Property Get valueNames() As String()
valueNames = mValueNames
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal target As Object, _
                ByVal sourceStudy As TradeBuildSP.IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                ByVal discriminator As Long)
Dim numValuesToNotify As Long
Dim inputDefs As TradeBuildSP.IStudyInputDefinitions
Dim i As Long

Set mTarget = target
Set mSourceStudy = sourceStudy
mValueNames = valueNames
mValueName = mValueNames(0)

mDiscriminator = discriminator

If TypeOf target Is TradeBuildSP.IStudy Then
    mTargetIsStudy = True
    Set mTargetStudy = target
    Set mStudyToNotify = mTargetStudy.baseStudy
    
    Set inputDefs = mTargetStudy.studyDefinition.studyInputDefinitions
    ReDim mNotifyNames(inputDefs.Count - 1) As String
    For i = 0 To inputDefs.Count - 1
        mNotifyNames(i) = inputDefs.Item(i + 1).name
    Next
    
    If numUnderlyingValuesToUse < 0 Then
        ' just replay the number of values needed by the study
        numValuesToNotify = mTargetStudy.numberOfBarsRequired
    Else
        numValuesToNotify = numUnderlyingValuesToUse
    End If
Else
    Set mTargetListener = target
    If numUnderlyingValuesToUse > 0 Then
        numValuesToNotify = numUnderlyingValuesToUse
    End If
End If
If numValuesToNotify > mSourceStudy.numberOfCachedValues Then numValuesToNotify = mSourceStudy.numberOfCachedValues

mNextValueNumber = mSourceStudy.numberOfCachedValues - numValuesToNotify + 1

End Sub

'================================================================================
' Helper Functions
'================================================================================





