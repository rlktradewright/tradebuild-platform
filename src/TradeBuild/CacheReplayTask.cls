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

Private mValuesToNotify As Long
Private mValuesNotified As Long

Private mValueName As String

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim i As Long
Dim ev As TradeBuildSP.StudyValueEvent
Dim valueEntry As StudyValueStruct

For i = (mValueCache.Count - mValuesToNotify + mValuesNotified + 1) To _
        (mValueCache.Count - mValuesToNotify + mValuesNotified + 20)
    
    If mValuesNotified = mValueCache.Count Then Exit For
    
    mValuesNotified = mValuesNotified + 1
    
    valueEntry = mValueCache.getStudyValue(i)
    If Not IsEmpty(valueEntry.value) Then
        Set ev.source = mSourceStudy
        ev.Timestamp = valueEntry.Timestamp
        ev.barNumber = valueEntry.barNumber
        ev.ValueName = mValueName
        If mResolver Is Nothing Then
            ev.value = valueEntry.value
        Else
            Dim sv As TradeBuildSP.StudyValue
            sv = valueEntry.value
            ev.value = mResolver.resolveValue(sv, mValueName)
        End If
        mTargetStudy.notify ev
        
    End If
Next

If mValuesNotified = mValueCache.Count Then
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

Public Property Get targetStudy() As TradeBuildSP.IStudy
Set targetStudy = mTargetStudy
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
                ByVal useAllUnderlyingValues As Boolean)
Set mValueCache = valueCache
Set mTargetStudy = targetStudy
Set mSourceStudy = sourceStudy
Set mResolver = resolver
mValueName = ValueName

mValuesToNotify = targetStudy.numberOfBarsRequired
If mValuesToNotify = 0 Or useAllUnderlyingValues Then mValuesToNotify = mValueCache.Count
If mValuesToNotify > mValueCache.Count Then mValuesToNotify = mValueCache.Count

End Sub

'================================================================================
' Helper Functions
'================================================================================



