Attribute VB_Name = "Globals"
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

Public Const ProjectName                        As String = "StudyUtils26"

Public Const AttributeNameEnabled               As String = "Enabled"
Public Const AttributeNameStudyLibraryBuiltIn   As String = "BuiltIn"
Public Const AttributeNameStudyLibraryProgId    As String = "ProgId"

Public Const BuiltInStudyLibProgId              As String = "CmnStudiesLib26.StudyLib"
Public Const BuiltInStudyLibName                As String = "BuiltIn"

Public Const ConfigNameStudyLibraries           As String = "StudyLibraries"
Public Const ConfigNameStudyLibrary             As String = "StudyLibrary"

Public Const MaxDouble                          As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble                          As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh                          As Double = MinDouble
Public Const DummyLow                           As Double = MaxDouble

Public Const DefaultStudyValueNameStr           As String = "$DEFAULT"
Public Const MovingAverageStudyValueNameStr     As String = "MA"

Public Const OneMicroSecond                     As Double = 1.15740740740741E-11
Public Const OneSecond                          As Double = 1 / 86400
Public Const OneMinute                          As Double = 1 / 1440
Public Const OneHour                            As Double = 1 / 24

Public Const StudyLibrariesRenderer             As String = "StudiesUI26.StudyLibConfigurer"


'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyLibraryManager    As New StudyLibraryManager

Private mStudies                As New Collection

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

Public Property Get gLogger() As Logger
Dim lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
End Property

Public Property Get StudyLibraryManager() As StudyLibraryManager
Set StudyLibraryManager = mStudyLibraryManager
End Property

Public Property Get StudiesCollection() As Collection
Set StudiesCollection = mStudies
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================



