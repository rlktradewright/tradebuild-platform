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

Public Property Get Count() As Long
Count = mNumValues
End Property

Public Property Get CyclicSize() As Long
CyclicSize = mCyclicSize
End Property

Public Property Get ValueName() As String
ValueName = mValueName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddValue( _
                ByVal value As Variant, _
                ByVal BarNumber As Long, _
                ByVal Timestamp As Date, _
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
mValues(mNewest).BarNumber = BarNumber
mValues(mNewest).Timestamp = Timestamp
mValues(mNewest).BarStartTime = BarStartTime
End Sub

Public Sub AppendValues( _
                ByVal pValueCache As ValueCache)
Dim i As Long
Dim val As SValue

For i = 1 To pValueCache.Count
    val = pValueCache.GetSValue(i)
    AddValue val.value, val.BarNumber, val.Timestamp, val.BarStartTime
Next
End Sub

Public Function AppendValuesAsync( _
                ByVal pValueCache As ValueCache) As TaskController
Dim avt As New AppendValuesTask

avt.Initialise Me, pValueCache

Set AppendValuesAsync = StartTask(avt, PriorityNormal)
End Function

Public Function GetSValue(Optional ByVal ref As Long) As SValue
Dim valueEntry As SValue
Dim failpoint As Long

On Error GoTo Err

valueEntry = mValues(refToIndex(ref))
GetSValue.BarNumber = valueEntry.BarNumber
GetSValue.Timestamp = valueEntry.Timestamp
GetSValue.BarStartTime = valueEntry.BarStartTime
If IsObject(valueEntry.value) Then
    Set GetSValue.value = valueEntry.value
Else
    GetSValue.value = valueEntry.value
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GetSValue" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function GetValue(Optional ByVal ref As Long) As Variant
Dim index As Long

Dim failpoint As Long

On Error GoTo Err

index = refToIndex(ref)
If IsObject(mValues(index).value) Then
    Set GetValue = mValues(index).value
Else
    GetValue = mValues(index).value
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GetValue" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Friend Sub Initialise( _
                ByVal CyclicSize As Long, _
                ByVal ValueName As String)
If CyclicSize > 0 Then
    mCyclicSize = CyclicSize
    ReDim mValues(mCyclicSize - 1) As SValue
Else
    ReDim mValues(15) As SValue
    mCyclicSize = 0
End If
mNewest = -1

mValueName = ValueName
End Sub

Public Sub UpdateValue( _
                ByVal value As Variant, _
                Optional ByVal BarNumber As Long, _
                Optional ByVal Timestamp As Date, _
                Optional ByVal BarStartTime As Date, _
                Optional ByVal ref As Long)
Dim i As Long
i = refToIndex(ref)
If IsObject(value) Then
    Set mValues(i).value = value
Else
    mValues(i).value = value
End If
If BarNumber > 0 Then mValues(i).BarNumber = BarNumber
If CDbl(Timestamp) > 0 Then mValues(i).Timestamp = Timestamp
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






