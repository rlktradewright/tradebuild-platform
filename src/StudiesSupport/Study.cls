VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Study"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
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

Public Enum ReplayNumbers
    ReplayAll = -1
    ReplayAsRequired = -2
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

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

Public Property Get baseStudy() As study

End Property

Public Property Get id() As String

End Property

Public Property Get instanceName() As String

End Property

Public Property Get instancePath() As String

End Property

Public Property Get numberOfBarsRequired() As Long

End Property

Public Property Get parameters() As Parameters2.parameters

End Property

Public Property Get sessionEndTime() As Date

End Property

Public Property Get sessionStartTime() As Date

End Property

Public Property Get studyDefinition() As studyDefinition

End Property

'================================================================================
' Methods
'================================================================================

Public Function addStudy( _
                ByVal pStudy As study, _
                ByRef valueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
                
End Function
                
Public Function addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
                
End Function

Public Function getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue

End Function

Public Function getValueTicksize( _
                ByVal valueName As String) As Double
                
End Function
     
Public Sub initialise( _
                ByVal id As String, _
                ByVal params As Parameters2.parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef valueNames() As String, _
                ByVal underlyingStudy As study)

End Sub

Public Sub notify( _
                ByRef ev As StudyValueEvent)

End Sub

Public Function numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
                
End Function

Public Sub removeStudyValueListener( _
                ByVal listener As StudyValueListener)
                
End Sub

Public Sub setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
                
End Sub
     
'================================================================================
' Helper Functions
'================================================================================


