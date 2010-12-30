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

'@================================================================================
' Description
'@================================================================================
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

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get BaseStudy() As Study

End Property

Public Property Get Id() As String

End Property

Public Property Get InstanceName() As String

End Property

Public Property Get InstancePath() As String

End Property

Public Property Get NumberOfBarsRequired() As Long

End Property

Public Property Get Parameters() As Parameters

End Property

Public Property Get StudyDefinition() As StudyDefinition

End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddStudy( _
                ByVal pStudy As Study, _
                ByRef valueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
                
End Function
                
Public Function AddStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
                
End Function

Public Function GetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue

End Function

Public Function GetValueTicksize( _
                ByVal valueName As String) As Double
                
End Function
     
Public Sub initialise( _
                ByVal Id As String, _
                ByVal params As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef valueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

End Sub

Public Sub Notify( _
                ByRef ev As StudyValueEventData)

End Sub

Public Function NumberOfCachedValues( _
                Optional ByVal valueName As String) As Long
                
End Function

Public Sub RemoveStudyValueListener( _
                ByVal listener As StudyValueListener)
                
End Sub

Public Sub SetInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
                
End Sub
     
'@================================================================================
' Helper Functions
'@================================================================================


