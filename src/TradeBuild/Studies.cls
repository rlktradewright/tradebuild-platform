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

Implements TradeBuildSP.iStudies

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
    ValueName           As String
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
                ByVal ValueName As String)
AddStudy study, ValueName
End Sub

Private Property Let iStudies_defaultValueName( _
                ByVal RHS As String)
defaultValueName = RHS
End Property

Private Sub IStudies_Notify(ev As TradeBuildSP.StudyValueEvent)
notifyValue ev
End Sub

Private Function IStudies_numberOfBarsRequired() As Long
IStudies_numberOfBarsRequired = numberOfBarsRequired
End Function

'Private Property Let IStudies_underlyingStudy( _
'                ByVal RHS As TradeBuildSP.IStudy)
'underlyingStudy = RHS
'End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let defaultValueName(ByVal value As String)
mDefaultValuename = value
End Property

'Public Property Let underlyingStudy(value As TradeBuildSP.IStudy)
''   don't do the following as it causes circular references
''Set mUnderlyingStudy = value
'End Property

'================================================================================
' Methods
'================================================================================

Public Sub add( _
                study As TradeBuild.study, _
                ValueName As String)
AddStudy study.study, ValueName
End Sub

Friend Sub AddStudy( _
                study As TradeBuildSP.IStudy, _
                ValueName As String)
Dim tableEntry As StudyTableEntry

If mNextStudyTableIndex > UBound(mStudyTable) Then
    ReDim Preserve mStudyTable(UBound(mStudyTable) + 10) As StudyTableEntry
End If
Set tableEntry.theStudy = study
tableEntry.ValueName = ValueName
mStudyTable(mNextStudyTableIndex) = tableEntry
mNextStudyTableIndex = mNextStudyTableIndex + 1

End Sub

Public Sub notify( _
                ev As TradeBuild.StudyValueEvent)
Dim spEv As TradeBuildSP.StudyValueEvent

spEv.barNumber = ev.barNumber
Set spEv.source = ev.source
spEv.timestamp = ev.timestamp
spEv.value = ev.value
spEv.ValueName = ev.ValueName

notifyValue spEv
End Sub

Public Function numberOfBarsRequired() As Long
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

'================================================================================
' Helper Functions
'================================================================================

Private Sub notifyValue( _
                ev As TradeBuildSP.StudyValueEvent)
Dim i As Long
Dim study As TradeBuildSP.IStudy

For i = 0 To mNextStudyTableIndex - 1
    Set study = mStudyTable(i).theStudy
    If ev.ValueName = mDefaultValuename Then
        If mStudyTable(i).ValueName = "$default" Then
            study.notify ev
        End If
    End If
    If mStudyTable(i).ValueName = ev.ValueName Then
        study.notify ev
    End If
Next
End Sub

