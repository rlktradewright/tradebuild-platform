VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStudyLibrary"
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

Public Property Get Name() As String

End Property

Public Property Let Name( _
                ByVal Value As String)

End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateStudy( _
                ByVal Name As String) As IStudy
                
End Function

Public Function GetImplementedStudyNames() As String()

End Function

Public Function GetStudyDefaultParameters( _
                ByVal Name As String) As Parameters
                
End Function

Public Function GetStudyDefinition( _
                ByVal Name As String) As StudyDefinition
    
End Function
    
Public Sub Initialise( _
                ByVal libManager As StudyLibraryManager, _
                ByVal handle As Long)

End Sub

Public Function SetStudyDefaultParameters( _
                ByVal Name As String, _
                ByVal params As Parameters) As Boolean

End Function

Public Sub Terminate()

End Sub

'@================================================================================
' Helper Functions
'@================================================================================


