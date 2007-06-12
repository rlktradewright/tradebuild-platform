VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyLibrary"
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

Public Property Get name() As String

End Property

Public Property Let name( _
                ByVal value As String)

End Property

'================================================================================
' Methods
'================================================================================

Public Function createStudy( _
                ByVal name As String) As Study
                
End Function

Public Function getImplementedStudyNames() As String()

End Function

Public Function GetStudyDefaultParameters( _
                ByVal name As String) As parameters
                
End Function

Public Function GetStudyDefinition( _
                ByVal name As String) As studyDefinition
    
End Function
    
Public Sub initialise( _
                ByVal libManager As StudyLibraryManager, _
                ByVal handle As Long)

End Sub

Public Function SetStudyDefaultParameters( _
                ByVal name As String, _
                ByVal params As parameters) As Boolean

End Function

Public Sub terminate()

End Sub

'================================================================================
' Helper Functions
'================================================================================


