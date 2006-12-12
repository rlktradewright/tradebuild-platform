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

Implements TradeBuildSP.IStudies

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

Private Type StudyTableEntry
    theStudy            As TradeBuildSP.IStudy
    studyToNotify       As TradeBuildSP.IStudy
    valueName           As String
    notifyName          As String
End Type

'================================================================================
' Member variables
'================================================================================

'Private mUnderlyingStudy As TradeBuildSP.IStudy
Private mDefaultValuename As String

Private mStudyTable() As StudyTableEntry
Private mNextStudyTableIndex As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mStudyTable(10) As StudyTableEntry
End Sub

Private Sub Class_Terminate()
Debug.Print "Studies terminated"
End Sub

'================================================================================
' IStudies Interface Members
'================================================================================

Private Sub IStudies_Add( _
                ByVal study As TradeBuildSP.IStudy, _
                ByRef valueNames() As String)
AddStudy study, valueNames
End Sub

Private Property Let iStudies_defaultValueName( _
                ByVal RHS As String)
defaultValueName = RHS
End Property

Private Sub IStudies_Notify(ev As TradeBuildSP.StudyValueEvent)
Dim i As Long
Dim evOut As StudyValueEvent

evOut = ev
For i = 0 To mNextStudyTableIndex - 1
    If ev.valueName = mDefaultValuename Then
        If mStudyTable(i).valueName = "$default" Then
            evOut.valueName = mStudyTable(i).notifyName
            mStudyTable(i).studyToNotify.notify evOut
        End If
    End If
    If mStudyTable(i).valueName = ev.valueName Then
        evOut.valueName = mStudyTable(i).notifyName
        mStudyTable(i).studyToNotify.notify evOut
    End If
Next
End Sub

Private Function IStudies_numberOfBarsRequired() As Long
IStudies_numberOfBarsRequired = numberOfBarsRequired
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let defaultValueName(ByVal value As String)
mDefaultValuename = value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub AddStudy( _
                ByVal study As TradeBuildSP.IStudy, _
                ByRef valueNames() As String)
Dim tableEntry As StudyTableEntry
Dim inputDefs As studyInputDefinitions
Dim i As Long

Set inputDefs = study.studyDefinition.studyInputDefinitions

For i = 0 To UBound(valueNames)
    If mNextStudyTableIndex > UBound(mStudyTable) Then
        ReDim Preserve mStudyTable(UBound(mStudyTable) + 10) As StudyTableEntry
    End If
    Set tableEntry.theStudy = study
    Set tableEntry.studyToNotify = study.baseStudy
    tableEntry.valueName = valueNames(i)
    tableEntry.notifyName = inputDefs.Item(i + 1).name
    mStudyTable(mNextStudyTableIndex) = tableEntry
    mNextStudyTableIndex = mNextStudyTableIndex + 1
Next
End Sub

Friend Function numberOfBarsRequired() As Long
Dim num As Long
Dim numStudy As Long
Dim study As TradeBuildSP.IStudy
Dim i As Long

For i = 0 To mNextStudyTableIndex - 1
    Set study = mStudyTable(i).theStudy
    numStudy = study.numberOfBarsRequired
    If numStudy > num Then num = numStudy
Next
numberOfBarsRequired = num
End Function



