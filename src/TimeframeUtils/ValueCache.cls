VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValueCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "ValueCache"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mValueName As String
Private mValues() As SValue
Private mNewest As Long
Private mNumValues As Long
Private mCyclicSize As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStudyValueCache Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get count() As Long
count = mNumValues
End Property

Public Property Get cyclicSize() As Long
cyclicSize = mCyclicSize
End Property

Public Property Get valueName() As String
valueName = mValueName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addValue( _
                ByVal value As Variant, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date, _
                ByVal BarStartTime As Date)

If mCyclicSize > 0 Then
    If mNumValues < mCyclicSize Then
        mNewest = mNewest + 1
        mNumValues = mNumValues + 1
    Else
        mNewest = (mNewest + 1) Mod mCyclicSize
    End If
Else
    If mNewest = UBound(mValues) Then
        ReDim Preserve mValues(2 * (UBound(mValues) + 1) - 1) As SValue
    End If
    mNewest = mNewest + 1
    mNumValues = mNumValues + 1
End If

If IsObject(value) Then
    Set mValues(mNewest).value = value
Else
    mValues(mNewest).value = value
End If
mValues(mNewest).barNumber = barNumber
mValues(mNewest).timestamp = timestamp
mValues(mNewest).BarStartTime = BarStartTime
End Sub

Public Sub appendValues( _
                ByVal pValueCache As ValueCache)
Dim i As Long
Dim val As SValue

For i = 1 To pValueCache.count
    val = pValueCache.getSValue(i)
    addValue val.value, val.barNumber, val.timestamp, val.BarStartTime
Next
End Sub

Public Function appendValuesAsync( _
                ByVal pValueCache As ValueCache) As TaskController
Dim avt As New AppendValuesTask

avt.initialise Me, pValueCache

Set appendValuesAsync = StartTask(avt, PriorityNormal)
End Function

Public Function getSValue(Optional ByVal ref As Long) As SValue
Dim valueEntry As SValue
Dim failpoint As Long
On Error GoTo Err

valueEntry = mValues(refToIndex(ref))
getSValue.barNumber = valueEntry.barNumber
getSValue.timestamp = valueEntry.timestamp
getSValue.BarStartTime = valueEntry.BarStartTime
If IsObject(valueEntry.value) Then
    Set getSValue.value = valueEntry.value
Else
    getSValue.value = valueEntry.value
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "getSValue" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Function

Public Function getValue(Optional ByVal ref As Long) As Variant
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

index = refToIndex(ref)
If IsObject(mValues(index).value) Then
    Set getValue = mValues(index).value
Else
    getValue = mValues(index).value
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "getValue" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Function

Friend Sub initialise( _
                ByVal cyclicSize As Long, _
                ByVal valueName As String)
If cyclicSize > 0 Then
    mCyclicSize = cyclicSize
    ReDim mValues(mCyclicSize - 1) As SValue
Else
    ReDim mValues(15) As SValue
    mCyclicSize = 0
End If
mNewest = -1

mValueName = valueName
End Sub

Public Sub updateValue( _
                ByVal value As Variant, _
                Optional ByVal barNumber As Long, _
                Optional ByVal timestamp As Date, _
                Optional ByVal BarStartTime As Date, _
                Optional ByVal ref As Long)
Dim i As Long
i = refToIndex(ref)
If IsObject(value) Then
    Set mValues(i).value = value
Else
    mValues(i).value = value
End If
If barNumber > 0 Then mValues(i).barNumber = barNumber
If CDbl(timestamp) > 0 Then mValues(i).timestamp = timestamp
If CDbl(BarStartTime) > 0 Then mValues(i).BarStartTime = BarStartTime
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function refToIndex(ByVal ref As Long) As Long
If ref > mNumValues Or _
    ref < (1 - mNumValues) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "refToIndex", _
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






