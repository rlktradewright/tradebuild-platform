VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CacheReplayTask"
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

Private mTargetStudy As TradeBuildSP.IStudy
Private mSourceStudy As TradeBuildSP.IStudy

Private mValueCache As TradeBuild.StudyValueCache

Private mResolver As TradeBuildSP.IStudyValueResolver

Private mStartValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IValueReplayTask Interface Members
'================================================================================

Private Property Get IValueReplayTask_sourceStudy() As TradeBuildSP.IStudy
Set IValueReplayTask_sourceStudy = mSourceStudy
End Property

Private Property Get IValueReplayTask_targetStudy() As TradeBuildSP.IStudy
Set IValueReplayTask_targetStudy = mTargetStudy
End Property

Private Property Get IValueReplayTask_ValueName() As String
IValueReplayTask_ValueName = mValueName
End Property

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim i As Long
Dim ev As TradeBuildSP.StudyValueEvent
Dim valueEntry As StudyValueStruct

For i = (mStartValueNumber + mValuesNotified) To _
        (mStartValueNumber + mValuesNotified + 19)
    
    valueEntry = mValueCache.getStudyValue(i)
    If Not IsEmpty(valueEntry.Value) Then
        Set ev.source = mSourceStudy
        ev.Timestamp = valueEntry.Timestamp
        ev.barNumber = valueEntry.barNumber
        ev.ValueName = mValueName
        If mResolver Is Nothing Then
            ev.Value = valueEntry.Value
        Else
            Dim sv As TradeBuildSP.StudyValue
            sv = valueEntry.Value
            ev.Value = mResolver.resolveValue(sv, mValueName)
        End If
        mTargetStudy.notify ev
        
    End If
    
    mValuesNotified = mValuesNotified + 1
    If i = mValueCache.Count Then
        mTaskContext.finish
        Set mTaskContext = Nothing
        Exit For
    End If
Next

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

Public Property Get sourceStudy() As TradeBuildSP.IStudy
Set sourceStudy = mSourceStudy
End Property

Public Property Get targetStudy() As TradeBuildSP.IStudy
Set targetStudy = mTargetStudy
End Property

Public Property Get ValueName() As String
ValueName = mValueName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal valueCache As StudyValueCache, _
                ByVal targetStudy As TradeBuildSP.IStudy, _
                ByVal sourceStudy As TradeBuildSP.IStudy, _
                ByVal ValueName As String, _
                ByVal resolver As TradeBuildSP.IStudyValueResolver, _
                ByVal numUnderlyingValuesToUse As Long)
Dim numValuesToNotify As Long

Set mValueCache = valueCache
Set mTargetStudy = targetStudy
Set mSourceStudy = sourceStudy
Set mResolver = resolver
mValueName = ValueName

numValuesToNotify = targetStudy.numberOfBarsRequired
If numUnderlyingValuesToUse = 0 Then numUnderlyingValuesToUse = mValueCache.Count
If numUnderlyingValuesToUse > numValuesToNotify Then numValuesToNotify = numUnderlyingValuesToUse
If numValuesToNotify > mValueCache.Count Then numValuesToNotify = mValueCache.Count

mStartValueNumber = mValueCache.Count - numValuesToNotify + 1

End Sub

'================================================================================
' Helper Functions
'================================================================================



