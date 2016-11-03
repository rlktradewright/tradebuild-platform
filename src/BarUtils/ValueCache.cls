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
                ByVal Value As Variant, _
                ByVal BarNumber As Long, _
                ByVal Timestamp As Date, _
                ByVal BarStartTime As Date)
Const ProcName As String = "AddValue"
On Error GoTo Err

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

gSetVariant mValues(mNewest).Value, Value
mValues(mNewest).BarNumber = BarNumber
mValues(mNewest).Timestamp = Timestamp
mValues(mNewest).BarStartTime = BarStartTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AppendValues( _
                ByVal pValueCache As ValueCache)
Const ProcName As String = "AppendValues"
On Error GoTo Err

Dim i As Long
For i = 1 To pValueCache.Count
    Dim val As SValue
    val = pValueCache.GetSValue(i)
    AddValue val.Value, val.BarNumber, val.Timestamp, val.BarStartTime
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AppendValuesAsync( _
                ByVal pValueCache As ValueCache) As TaskController
Const ProcName As String = "AppendValuesAsync"
On Error GoTo Err

Dim avt As New AppendValuesTask
avt.Initialise Me, pValueCache

Set AppendValuesAsync = StartTask(avt, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSValue(Optional ByVal ref As Long) As SValue
Const ProcName As String = "GetSValue"
On Error GoTo Err

GetSValue = mValues(refToIndex(ref))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetValue(Optional ByVal ref As Long) As Variant
Const ProcName As String = "GetValue"
On Error GoTo Err

gSetVariant GetValue, mValues(refToIndex(ref)).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal CyclicSize As Long, _
                ByVal ValueName As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

If CyclicSize > 0 Then
    mCyclicSize = CyclicSize
    ReDim mValues(mCyclicSize - 1) As SValue
Else
    ReDim mValues(15) As SValue
    mCyclicSize = 0
End If
mNewest = -1

mValueName = ValueName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UpdateValue( _
                ByVal Value As Variant, _
                Optional ByVal BarNumber As Long, _
                Optional ByVal Timestamp As Date, _
                Optional ByVal BarStartTime As Date, _
                Optional ByVal ref As Long)
Const ProcName As String = "UpdateValue"
On Error GoTo Err

Dim i As Long
i = refToIndex(ref)
gSetVariant mValues(i).Value, Value
If BarNumber > 0 Then mValues(i).BarNumber = BarNumber
If CDbl(Timestamp) > 0 Then mValues(i).Timestamp = Timestamp
If CDbl(BarStartTime) > 0 Then mValues(i).BarStartTime = BarStartTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function refToIndex(ByVal ref As Long) As Long
Const ProcName As String = "refToIndex"
On Error GoTo Err

AssertArgument ref <= mNumValues And ref >= (1 - mNumValues), "<ref> out of bounds"
    
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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function






