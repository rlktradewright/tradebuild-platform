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

Implements TradeBuildSP.IStudyValueCache

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

Private mTradeBuildAPIRef           As WeakReference

Private mValueName As String
Private mValues() As TradeBuildSP.StudyValue
Private mNewest As Long
Private mNumValues As Long
Private mCyclicSize As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IStudyValueCache Interface Members
'================================================================================

Private Sub IStudyValueCache_addValue( _
                ByVal value As Variant, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date)
addValue value, barNumber, timestamp
End Sub

Private Property Get IStudyValueCache_Count() As Long
IStudyValueCache_Count = Count
End Property

Private Property Get IStudyValueCache_cyclicSize() As Long
IStudyValueCache_cyclicSize = mCyclicSize
End Property

Private Function IStudyValueCache_getStudyValue( _
                Optional ByVal ref As Long) As TradeBuildSP.StudyValue
IStudyValueCache_getStudyValue = mValues(refToIndex(ref))
End Function

Private Function IStudyValueCache_getValue( _
                Optional ByVal ref As Long) As Variant
If TypeOf getValue(ref) Is Object  Then
    Set IStudyValueCache_getValue = getValue(ref)
Else
    IStudyValueCache_getValue = getValue(ref)
End If
End Function

Private Function IStudyValueCache_replayValues( _
                ByVal target As Object, _
                ByVal sourceStudy As TradeBuildSP.IStudy, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal discriminator As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Set IStudyValueCache_replayValues = replayValues(target, _
                                                sourceStudy, _
                                                numUnderlyingValuesToUse, _
                                                discriminator, _
                                                taskName, _
                                                taskData)
End Function

Private Sub IStudyValueCache_updateValue( _
                ByVal value As Variant, _
                Optional ByVal barNumber As Long, _
                Optional ByVal timestamp As Date, _
                Optional ByVal ref As Long)
updateValue value, barNumber, timestamp, ref
End Sub

Private Property Get IStudyValueCache_ValueName() As String
IStudyValueCache_ValueName = mValueName
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let cyclicSize(value As Long)
If value > 0 Then
    If value > mCyclicSize Then
        mCyclicSize = value
        ReDim mValues(mCyclicSize - 1) As TradeBuildSP.StudyValue
    End If
Else
    ReDim mValues(1000) As TradeBuildSP.StudyValue
    mCyclicSize = 0
End If
mNewest = -1
End Property

Friend Property Get cyclicSize() As Long
cyclicSize = mCyclicSize
End Property

Friend Property Get Count() As Long
Count = mNumValues
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = New WeakReference
mTradeBuildAPIRef.Initialise value
End Property

Friend Property Let valueName(value As String)
mValueName = value
End Property

Friend Property Get valueName() As String
valueName = mValueName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub addValue( _
                ByVal value As Variant, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date)

If mNumValues > 0 Then
'    If Timestamp < getStudyValue(0).Timestamp Then Stop
End If

If mCyclicSize > 0 Then
    If mNumValues < mCyclicSize Then
        mNewest = mNewest + 1
        mNumValues = mNumValues + 1
    Else
        mNewest = (mNewest + 1) Mod mCyclicSize
    End If
Else
    If mNewest = UBound(mValues) Then
        ReDim Preserve mValues(2 * UBound(mValues)) As TradeBuildSP.StudyValue
    End If
    mNewest = mNewest + 1
    mNumValues = mNumValues + 1
End If

If TypeOf value Is Object  Then
    Set mValues(mNewest).value = value
Else
    mValues(mNewest).value = value
End If
mValues(mNewest).barNumber = barNumber
mValues(mNewest).timestamp = timestamp
End Sub

Friend Function getValue(Optional ByVal ref As Long) As Variant
If TypeOf mValues(refToIndex(ref)).value Is Object  Then
    Set getValue = mValues(refToIndex(ref)).value
Else
    getValue = mValues(refToIndex(ref)).value
End If
End Function

Friend Function getStudyValue(Optional ByVal ref As Long) As StudyValueStruct
Dim valueEntry As TradeBuildSP.StudyValue
valueEntry = mValues(refToIndex(ref))
getStudyValue.barNumber = valueEntry.barNumber
getStudyValue.timestamp = valueEntry.timestamp
If TypeOf valueEntry.value Is Object  Then
    Set getStudyValue.value = valueEntry.value
Else
    getStudyValue.value = valueEntry.value
End If
End Function

Friend Function replayValues( _
                ByVal target As Object, _
                ByVal sourceStudy As TradeBuildSP.IStudy, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal discriminator As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
Dim replayTask As New CacheReplayTask

replayTask.Initialise Me, _
                    target, _
                    sourceStudy, _
                    numUnderlyingValuesToUse, _
                    discriminator

Set replayValues = tb.taskManager.startTask(replayTask, IIf(taskName <> "", taskName, gGenerateGUIDString), taskData)
End Function

Friend Sub updateValue( _
                ByVal value As Variant, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date, _
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
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function refToIndex(ByVal ref As Long) As Long
If ref > mNumValues Or _
    ref < (1 - mNumValues) _
Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild.StudyValueCache::refToIndex", _
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

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.target
End Function


