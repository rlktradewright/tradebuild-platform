VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Studies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Studies"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type LinkedStudyTableEntry
    TheStudy            As IStudy
    StudyToNotify       As IStudy
    ValueName           As String
    NotifyName          As String
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mDefaultValuename As String

Private mStudyTable() As LinkedStudyTableEntry
Private mNextStudyTableIndex As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mStudyTable(3) As LinkedStudyTableEntry
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let DefaultValueName(ByVal Value As String)
mDefaultValuename = UCase$(Value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add( _
                ByVal Study As IStudy, _
                ByRef ValueNames() As String)
Const ProcName As String = "Add"
On Error GoTo Err

Dim inputDefs As StudyInputDefinitions
Set inputDefs = Study.StudyDefinition.StudyInputDefinitions

Dim i As Long
For i = 0 To UBound(ValueNames)
    If ValueNames(i) <> "" Then
        If mNextStudyTableIndex > UBound(mStudyTable) Then
            ReDim Preserve mStudyTable(2 * (UBound(mStudyTable) + 1) - 1) As LinkedStudyTableEntry
        End If
        Dim tableEntry As LinkedStudyTableEntry
        Set tableEntry.TheStudy = Study
        Set tableEntry.StudyToNotify = Study.BaseStudy
        tableEntry.ValueName = UCase$(ValueNames(i))
        tableEntry.NotifyName = UCase$(inputDefs.Item(i + 1).Name)
        mStudyTable(mNextStudyTableIndex) = tableEntry
        mNextStudyTableIndex = mNextStudyTableIndex + 1
    End If
Next

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddEx( _
                ByVal Study As IStudy, _
                ByRef ValueNames() As String, _
                ByRef NotifyNames() As String)
Const ProcName As String = "AddEx"
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(ValueNames)
    If ValueNames(i) <> "" Then
        If mNextStudyTableIndex > UBound(mStudyTable) Then
            ReDim Preserve mStudyTable(2 * (UBound(mStudyTable) + 1) - 1) As LinkedStudyTableEntry
        End If
        Dim tableEntry As LinkedStudyTableEntry
        Set tableEntry.TheStudy = Study
        Set tableEntry.StudyToNotify = Study.BaseStudy
        tableEntry.ValueName = UCase$(ValueNames(i))
        tableEntry.NotifyName = UCase$(NotifyNames(i))
        mStudyTable(mNextStudyTableIndex) = tableEntry
        mNextStudyTableIndex = mNextStudyTableIndex + 1
    End If
Next

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Notify(ev As StudyValueEventData)
Const ProcName As String = "Notify"
On Error GoTo Err

ev.ValueName = UCase$(ev.ValueName)

Dim evOut As StudyValueEventData
evOut = ev

Dim i As Long
For i = 0 To mNextStudyTableIndex - 1
    If ev.ValueName = mDefaultValuename Then
        If mStudyTable(i).ValueName = DefaultStudyValueNameStr Then
            evOut.ValueName = mStudyTable(i).NotifyName
            mStudyTable(i).StudyToNotify.Notify evOut
        End If
    End If
    If mStudyTable(i).ValueName = ev.ValueName Then
        evOut.ValueName = mStudyTable(i).NotifyName
        mStudyTable(i).StudyToNotify.Notify evOut
    End If
Next

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NumberOfBarsRequired() As Long
Const ProcName As String = "NumberOfBarsRequired"
On Error GoTo Err

Dim num As Long

Dim i As Long
For i = 0 To mNextStudyTableIndex - 1
    Dim Study As IStudy
    Set Study = mStudyTable(i).TheStudy
    
    Dim numStudy As Long
    numStudy = Study.NumberOfBarsRequired
    If numStudy > num Then num = numStudy
Next
NumberOfBarsRequired = num

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub NotifyValueTicksize( _
                ByVal ValueName As String, _
                ByVal TickSize As Double)
Const ProcName As String = "NotifyValueTicksize"
On Error GoTo Err

ValueName = UCase$(ValueName)

Dim i As Long
For i = 0 To mNextStudyTableIndex - 1
    If mStudyTable(i).ValueName = ValueName Then
        mStudyTable(i).TheStudy.SetInputTicksize mStudyTable(i).NotifyName, TickSize
    End If
Next

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




