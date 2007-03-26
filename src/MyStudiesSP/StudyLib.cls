VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
' A Study Library which serves up Study objects that are integrated
' into the TradeWright study handling and charting mechanisms.

'================================================================================
' Interfaces
'================================================================================

Implements StudyLibrary

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

Private mName As String
Private mHandle As Long

Private mStudyLongNames() As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
'
' TODO: set the name of your service provider to something unique
'
mName = "Richard's Sample Study Library"

ReDim mStudyLongNames(0) As String

'
' TODO: add a similar line for each of your studies
'
addStudyName RsiName
'
'----------------------------------------------------

End Sub

'================================================================================
' StudyServiceProvider Interface Members
'================================================================================

Private Function StudyLibrary_createStudy( _
                ByVal name As String) As study
'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    Set StudyLibrary_createStudy = New rsi
End Select
'
'----------------------------------------------------

End Function

Private Function StudyLibrary_getImplementedStudyNames() As String()
StudyLibrary_getImplementedStudyNames = mStudyLongNames
End Function

Private Function StudyLibrary_getStudyDefaultParameters( _
                ByVal name As String) As Parameters
'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    Set StudyLibrary_getStudyDefaultParameters = GRsi.defaultParameters
End Select
'
'----------------------------------------------------

End Function

Private Function StudyLibrary_getStudyDefinition( _
                ByVal name As String) As studyDefinition
'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    Set StudyLibrary_getStudyDefinition = GRsi.studyDefinition
End Select

End Function

Private Sub StudyLibrary_initialise( _
                ByVal libManager As StudyUtils25.StudyLibraryManager, _
                ByVal handle As Long)
Set gLibraryManager = libManager
mHandle = handle
End Sub

Private Property Let StudyLibrary_name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get StudyLibrary_name() As String
StudyLibrary_name = mName
End Property

Private Function StudyLibrary_setStudyDefaultParameters( _
                ByVal name As String, _
                ByVal Parameters As Parameters) As Boolean

StudyLibrary_setStudyDefaultParameters = True

'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    GRsi.defaultParameters = Parameters
Case Else
    StudyLibrary_setStudyDefaultParameters = False
End Select

End Function

Private Sub StudyLibrary_terminate()
' nothing to do
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub addStudyName(ByVal name As String)
If mStudyLongNames(0) <> "" Then
    ReDim Preserve mStudyLongNames(UBound(mStudyLongNames) + 1) As String
End If
mStudyLongNames(UBound(mStudyLongNames)) = name
End Sub



