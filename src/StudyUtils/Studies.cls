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

Private Type StudyTableEntry
    theStudy            As Study
    studyToNotify       As Study
    valueName           As String
    notifyName          As String
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
                ByRef valueNames() As String)
Dim tableEntry As StudyTableEntry
Dim inputDefs As StudyInputDefinitions
Dim i As Long

Const ProcName As String = "Add"
On Error GoTo Err

Set inputDefs = Study.StudyDefinition.StudyInputDefinitions

For i = 0 To UBound(valueNames)
    If valueNames(i) <> "" Then
        If mNextStudyTableIndex > UBound(mStudyTable) Then
            ReDim Preserve mStudyTable(2 * (UBound(mStudyTable) + 1) - 1) As StudyTableEntry
        End If
        Set tableEntry.theStudy = Study
        Set tableEntry.studyToNotify = Study.BaseStudy
        tableEntry.valueName = UCase$(valueNames(i))
        tableEntry.notifyName = UCase$(inputDefs.Item(i + 1).Name)
        mStudyTable(mNextStudyTableIndex) = tableEntry
        mNextStudyTableIndex = mNextStudyTableIndex + 1
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub AddEx( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByRef notifyNames() As String)
Dim tableEntry As StudyTableEntry
Dim i As Long

Const ProcName As String = "AddEx"
On Error GoTo Err

For i = 0 To UBound(valueNames)
    If valueNames(i) <> "" Then
        If mNextStudyTableIndex > UBound(mStudyTable) Then
            ReDim Preserve mStudyTable(2 * (UBound(mStudyTable) + 1) - 1) As StudyTableEntry
        End If
        Set tableEntry.theStudy = Study
        Set tableEntry.studyToNotify = Study.BaseStudy
        tableEntry.valueName = UCase$(valueNames(i))
        tableEntry.notifyName = UCase$(notifyNames(i))
        mStudyTable(mNextStudyTableIndex) = tableEntry
        mNextStudyTableIndex = mNextStudyTableIndex + 1
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Notify(ev As StudyValueEventData)
Dim i As Long
Dim evOut As StudyValueEventData

Const ProcName As String = "Notify"
On Error GoTo Err

ev.valueName = UCase$(ev.valueName)
evOut = ev
For i = 0 To mNextStudyTableIndex - 1
    If ev.valueName = mDefaultValuename Then
        If mStudyTable(i).valueName = DefaultStudyValueNameStr Then
            evOut.valueName = mStudyTable(i).notifyName
            mStudyTable(i).studyToNotify.Notify evOut
        End If
    End If
    If mStudyTable(i).valueName = ev.valueName Then
        evOut.valueName = mStudyTable(i).notifyName
        mStudyTable(i).studyToNotify.Notify evOut
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function NumberOfBarsRequired() As Long
Dim num As Long
Dim numStudy As Long
Dim Study As Study
Dim i As Long

Const ProcName As String = "NumberOfBarsRequired"
On Error GoTo Err

For i = 0 To mNextStudyTableIndex - 1
    Set Study = mStudyTable(i).theStudy
    numStudy = Study.NumberOfBarsRequired
    If numStudy > num Then num = numStudy
Next
NumberOfBarsRequired = num

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub NotifyValueTicksize( _
                ByVal valueName As String, _
                ByVal tickSize As Double)
Dim i As Long

Const ProcName As String = "NotifyValueTicksize"
On Error GoTo Err

valueName = UCase$(valueName)
For i = 0 To mNextStudyTableIndex - 1
    If mStudyTable(i).valueName = valueName Then
        mStudyTable(i).theStudy.SetInputTicksize mStudyTable(i).notifyName, tickSize
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




