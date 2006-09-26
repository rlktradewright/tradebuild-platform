VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private mStudy As IStudy                ' reference to the study itself
Private mStudyName As String            ' name of the study
Private mSpName As String               ' sevice provider name
Private mStudyID As String              ' id of this study
Private mUnderlyingStudyID As String    ' id of entry for study this study is to be
                                        ' added to
Private mInputValueName As String       ' name of value from the underlying study used
                                        ' as input to this study
Private mNumberOfCachedValues As Long   ' the number of study output values that should be
                                        ' for access by higher level studies
Private mStudyStarted As Boolean        ' set when this study has been added to the
                                        ' underlying study

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get inputValueName() As String
inputValueName = mInputValueName
End Property

Public Property Let inputValueName(ByRef newValue As String)
mInputValueName = newValue
End Property

Public Property Get numberOfCachedValues() As Long
numberOfCachedValues = mNumberOfCachedValues
End Property

Public Property Let numberOfCachedValues(ByVal newValue As Long)
mNumberOfCachedValues = newValue
End Property

Public Property Get spName() As String
spName = mSpName
End Property

Public Property Let spName(ByVal newValue As String)
mSpName = newValue
End Property

Public Property Get study() As IStudy
Set study = mStudy
End Property

Public Property Let study(ByVal newValue As IStudy)
Set mStudy = newValue
End Property

Public Property Get studyID() As String
studyID = mStudyID
End Property

Public Property Let studyID(ByVal newValue As String)
mStudyID = newValue
End Property

Public Property Get studyName() As String
studyName = mStudyName
End Property

Public Property Let studyName(ByVal newValue As String)
mStudyName = newValue
End Property

Public Property Get studyStarted() As Boolean
studyStarted = mStudyStarted
End Property

Public Property Let studyStarted(ByVal newValue As Boolean)
mStudyStarted = newValue
End Property

Public Property Get underlyingStudyID() As String
underlyingStudyID = mUnderlyingStudyID
End Property

Public Property Let underlyingStudyID(ByVal newValue As String)
mUnderlyingStudyID = newValue
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


