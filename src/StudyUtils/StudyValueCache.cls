VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueCache"
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

'Implements StudyValueCache

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ProjectName                   As String = "StudyUtils26"
Private Const ModuleName                    As String = "StudyValueCache"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mValueName As String
Private mValues() As StudyValue
Private mNewest As Long
Private mNumValues As Long
Private mCyclicSize As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IStudyValueCache Interface Members
'================================================================================

'Private Sub IStudyValueCache_addValue( _
'                ByVal value As Variant, _
'                ByVal barNumber As Long, _
'                ByVal timestamp As Date, _
'                ByVal barStartTime As Date)
'addValue value, barNumber, timestamp, barStartTime
'End Sub
'
'Private Property Get IStudyValueCache_Count() As Long
'IStudyValueCache_Count = mNumValues
'End Property
'
'Private Property Get IStudyValueCache_cyclicSize() As Long
'IStudyValueCache_cyclicSize = mCyclicSize
'End Property
'
'Private Function IStudyValueCache_getStudyValue( _
'                Optional ByVal ref As Long) As TradeBuildSP.StudyValue
'IStudyValueCache_getStudyValue = mValues(refToIndex(ref))
'End Function
'
'Private Function IStudyValueCache_getValue( _
'                Optional ByVal ref As Long) As Variant
'If TypeOf getValue(ref) Is Object  Then
'    Set IStudyValueCache_getValue = getValue(ref)
'Else
'    IStudyValueCache_getValue = getValue(ref)
'End If
'End Function
'
'Private Function IStudyValueCache_replayValues( _
'                ByVal target As Object, _
'                ByVal sourceStudy As Study, _
'                ByVal numberOfValuesToReplay As Long, _
'                Optional ByVal discriminator As Long, _
'                Optional ByVal taskName As String, _
'                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
'Set IStudyValueCache_replayValues = replayValues(target, _
'                                                sourceStudy, _
'                                                numberOfValuesToReplay, _
'                                                discriminator, _
'                                                taskName, _
'                                                taskData)
'End Function
'
'Private Sub IStudyValueCache_updateValue( _
'                ByVal value As Variant, _
'                Optional ByVal barNumber As Long, _
'                Optional ByVal timestamp As Date, _
'                Optional ByVal barStartTime As Date, _
'                Optional ByVal ref As Long)
'updateValue value, barNumber, timestamp, barStartTime, ref
'End Sub
'
'Private Property Get IStudyValueCache_ValueName() As String
'IStudyValueCache_ValueName = mValueName
'End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get count() As Long
count = mNumValues
End Property

Public Property Get cyclicSize() As Long
cyclicSize = mCyclicSize
End Property

Public Property Get valueName() As String
valueName = mValueName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addValue( _
                ByVal value As Variant, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date, _
                ByVal barStartTime As Date)

If mCyclicSize > 0 Then
    If mNumValues < mCyclicSize Then
        mNewest = mNewest + 1
        mNumValues = mNumValues + 1
    Else
        mNewest = (mNewest + 1) Mod mCyclicSize
    End If
Else
    If mNewest = UBound(mValues) Then
        ReDim Preserve mValues(2 * UBound(mValues)) As StudyValue
    End If
    mNewest = mNewest + 1
    mNumValues = mNumValues + 1
End If

If TypeName(value) = "Nothing" Then
    Set mValues(mNewest).value = Nothing
ElseIf TypeOf value Is Object  Then
    Set mValues(mNewest).value = value
Else
    mValues(mNewest).value = value
End If
mValues(mNewest).barNumber = barNumber
mValues(mNewest).timestamp = timestamp
mValues(mNewest).barStartTime = barStartTime
End Sub

Public Function getStudyValue(Optional ByVal ref As Long) As StudyValue
Dim valueEntry As StudyValue
Dim failpoint As Long
On Error GoTo Err

valueEntry = mValues(refToIndex(ref))
getStudyValue.barNumber = valueEntry.barNumber
getStudyValue.timestamp = valueEntry.timestamp
getStudyValue.barStartTime = valueEntry.barStartTime
If TypeName(valueEntry.value) = "Nothing" Then
    Set getStudyValue.value = Nothing
ElseIf TypeOf valueEntry.value Is Object  Then
    Set getStudyValue.value = valueEntry.value
Else
    getStudyValue.value = valueEntry.value
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "getStudyValue" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.description

End Function

Public Function getValue(Optional ByVal ref As Long) As Variant
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

index = refToIndex(ref)
If TypeOf mValues(index).value Is Object  Then
    Set getValue = mValues(index).value
Else
    getValue = mValues(index).value
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "getValue" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.description

End Function

Friend Sub initialise( _
                ByVal cyclicSize As Long, _
                ByVal valueName As String)
If cyclicSize > 0 Then
    mCyclicSize = cyclicSize
    ReDim mValues(mCyclicSize - 1) As StudyValue
Else
    ReDim mValues(1000) As StudyValue
    mCyclicSize = 0
End If
mNewest = -1

mValueName = valueName
End Sub



Public Function replayValues( _
                ByVal target As Object, _
                ByVal sourceStudy As study, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal discriminator As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
Dim replayTask As New CacheReplayTask

If numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll _
Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "StudiesSupp" & "." & "StudyValueCache" & ":" & "replayValues", _
            "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"
    
End If
replayTask.initialise Me, _
                    target, _
                    sourceStudy, _
                    numberOfValuesToReplay, _
                    discriminator

Set replayValues = StartTask( _
                                replayTask, _
                                PriorityNormal, _
                                IIf(taskName <> "", taskName, GenerateGUIDString), _
                                taskData)
End Function

Public Sub updateValue( _
                ByVal value As Variant, _
                Optional ByVal barNumber As Long, _
                Optional ByVal timestamp As Date, _
                Optional ByVal barStartTime As Date, _
                Optional ByVal ref As Long)
Dim i As Long
i = refToIndex(ref)
If TypeOf value Is Object  Then
    Set mValues(i).value = value
Else
    mValues(i).value = value
End If
If barNumber > 0 Then mValues(i).barNumber = barNumber
If CDbl(timestamp) > 0 Then mValues(i).timestamp = timestamp
If CDbl(barStartTime) > 0 Then mValues(i).barStartTime = barStartTime
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function refToIndex(ByVal ref As Long) As Long
If ref > mNumValues Or _
    ref < (1 - mNumValues) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "StudyUtils25.StudyValueCache::refToIndex", _
            "<ref> out of bounds"
End If
    
If mCyclicSize <> 0 Then
    If ref > 0 Then
        If mNumValues < mCyclicSize Then
            refToIndex = ref - 1
        Else
            refToIndex = (mNewest + ref) Mod mCyclicSize
        End If
    Else
        refToIndex = (mCyclicSize + mNewest + ref) Mod mCyclicSize
    End If
Else
    If ref > 0 Then
        refToIndex = ref - 1
    Else
        refToIndex = mNewest + ref
    End If
End If

End Function




