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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type StudyTableEntry
    theStudy            As Study
    studyToNotify       As Study
    ValueName           As String
    NotifyName          As String
End Type

'@================================================================================
' Member variables
'@================================================================================

'Private mUnderlyingStudy As Study
Private mDefaultValuename As String

Private mStudyTable() As StudyTableEntry
Private mNextStudyTableIndex As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mStudyTable(3) As StudyTableEntry
End Sub

Private Sub Class_Terminate()
Debug.Print "Studies terminated"
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let DefaultValueName(ByVal value As String)
mDefaultValuename = UCase$(value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add( _
                ByVal Study As Study, _
                ByRef ValueNames() As String)
Dim tableEntry As StudyTableEntry
Dim inputDefs As StudyInputDefinitions
Dim i As Long

Set inputDefs = Study.StudyDefinition.StudyInputDefinitions

For i = 0 To UBound(ValueNames)
    If ValueNames(i) <> "" Then
        If mNextStudyTableIndex > UBound(mStudyTable) Then
            ReDim Preserve mStudyTable(2 * (UBound(mStudyTable) + 1) - 1) As StudyTableEntry
        End If
        Set tableEntry.theStudy = Study
        Set tableEntry.studyToNotify = Study.BaseStudy
        tableEntry.ValueName = UCase$(ValueNames(i))
        tableEntry.NotifyName = UCase$(inputDefs.Item(i + 1).Name)
        mStudyTable(mNextStudyTableIndex) = tableEntry
        mNextStudyTableIndex = mNextStudyTableIndex + 1
    End If
Next
End Sub

Public Sub AddEx( _
                ByVal Study As Study, _
                ByRef ValueNames() As String, _
                ByRef NotifyNames() As String)
Dim tableEntry As StudyTableEntry
Dim i As Long

For i = 0 To UBound(ValueNames)
    If ValueNames(i) <> "" Then
        If mNextStudyTableIndex > UBound(mStudyTable) Then
            ReDim Preserve mStudyTable(2 * (UBound(mStudyTable) + 1) - 1) As StudyTableEntry
        End If
        Set tableEntry.theStudy = Study
        Set tableEntry.studyToNotify = Study.BaseStudy
        tableEntry.ValueName = UCase$(ValueNames(i))
        tableEntry.NotifyName = UCase$(NotifyNames(i))
        mStudyTable(mNextStudyTableIndex) = tableEntry
        mNextStudyTableIndex = mNextStudyTableIndex + 1
    End If
Next
End Sub

Public Sub Notify(ev As StudyValueEvent)
Dim i As Long
Dim evOut As StudyValueEvent

ev.ValueName = UCase$(ev.ValueName)
evOut = ev
For i = 0 To mNextStudyTableIndex - 1
    If ev.ValueName = mDefaultValuename Then
        If mStudyTable(i).ValueName = DefaultStudyValueNameStr Then
            evOut.ValueName = mStudyTable(i).NotifyName
            mStudyTable(i).studyToNotify.Notify evOut
        End If
    End If
    If mStudyTable(i).ValueName = ev.ValueName Then
        evOut.ValueName = mStudyTable(i).NotifyName
        mStudyTable(i).studyToNotify.Notify evOut
    End If
Next
End Sub

Public Function NumberOfBarsRequired() As Long
Dim num As Long
Dim numStudy As Long
Dim Study As Study
Dim i As Long

For i = 0 To mNextStudyTableIndex - 1
    Set Study = mStudyTable(i).theStudy
    numStudy = Study.NumberOfBarsRequired
    If numStudy > num Then num = numStudy
Next
NumberOfBarsRequired = num
End Function

Public Sub NotifyValueTicksize( _
                ByVal ValueName As String, _
                ByVal tickSize As Double)
Dim i As Long

ValueName = UCase$(ValueName)
For i = 0 To mNextStudyTableIndex - 1
    If mStudyTable(i).ValueName = ValueName Then
        mStudyTable(i).theStudy.SetInputTicksize mStudyTable(i).NotifyName, tickSize
    End If
Next
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




