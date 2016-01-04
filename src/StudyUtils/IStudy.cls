VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStudy"
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

Public Property Get BaseStudy() As IStudy

End Property

Public Property Get Id() As String

End Property

Public Property Get InstanceName() As String

End Property

Public Property Get InstancePath() As String

End Property

Public Property Get LibraryName() As String

End Property

Public Property Get Name() As String

End Property

Public Property Get NumberOfBarsRequired() As Long

End Property

Public Property Get Parameters() As Parameters

End Property

Public Property Get StudyDefinition() As StudyDefinition

End Property

Public Property Get UnderlyingStudy() As IStudy

End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddStudy( _
                ByVal pStudy As IStudy, _
                ByRef pValueNames() As String, _
                ByVal pNumberOfValuesToReplay As Long, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pTaskData As Variant) As TaskController
                
End Function
                
Public Function AddStudyValueListener( _
                ByVal pListener As IStudyValueListener, _
                ByVal pValueName As String, _
                ByVal pNumberOfValuesToReplay As Long, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pTaskData As Variant) As TaskController
                
End Function

Public Function GetStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long) As SValue

End Function

Public Function GetValueTicksize( _
                ByVal ValueName As String) As Double
                
End Function
     
Public Sub Initialise( _
                ByVal pStudyFoundation As StudyFoundation)

End Sub

Public Sub Notify( _
                ByRef ev As StudyValueEventData)

End Sub

Public Function NumberOfCachedValues( _
                Optional ByVal ValueName As String) As Long
                
End Function

Public Sub RemoveStudyValueListener( _
                ByVal pListener As IStudyValueListener)
                
End Sub

Public Sub SetInputTicksize( _
                ByVal pInputName As String, _
                ByVal pTickSize As Double)
                
End Sub
     
'@================================================================================
' Helper Functions
'@================================================================================


