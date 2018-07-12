VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyTableEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "StudyTableEntry"

'@================================================================================
' Member variables
'@================================================================================

' the study manager responsible for this study
Private mStudyManager                               As StudyManager
    
' the SessionBuilder used when replaying historicalinput
' values to the study
Private mSessionBuilder                             As SessionBuilder
    
' the study itself
Private mStudy                                      As IStudy
    
' name of the study
Private mStudyName                                  As String
    
' study library Name
Private mLibraryName                                As String
    
' Id of this study
Private mStudyID                                    As String
    
' Id of entry for study this study is to be added to
Private mUnderlyingStudyID                          As String
    
' names of values from the underlying study
' used as input to this study
Private mInputValueNames()                          As String
    
' the number of study output values that should be available
' for access by higher level studies
Private mNumberOfCachedValues                       As Long
    
Private mIncludeDataOutsideSession                  As Boolean
    
' set when this study has been added to the underlying study
Private mStudyStarted                               As Boolean

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

Friend Property Let StudyManager(ByVal Value As StudyManager)
Set mStudyManager = Value
End Property
    
Friend Property Get StudyManager() As StudyManager
Set StudyManager = mStudyManager
End Property
    
Friend Property Let SessionBuilder(ByVal Value As SessionBuilder)
Set mSessionBuilder = Value
End Property
    
Friend Property Get SessionBuilder() As SessionBuilder
Set SessionBuilder = mSessionBuilder
End Property
    
Friend Property Let Study(ByVal Value As IStudy)
Set mStudy = Value
End Property
    
Friend Property Get Study() As IStudy
Set Study = mStudy
End Property
    
Friend Property Let StudyName(ByVal Value As String)
mStudyName = Value
End Property
    
Friend Property Get StudyName() As String
StudyName = mStudyName
End Property
    
Friend Property Let LibraryName(ByVal Value As String)
mLibraryName = Value
End Property
    
Friend Property Get LibraryName() As String
LibraryName = mLibraryName
End Property
    
Friend Property Let StudyID(ByVal Value As String)
mStudyID = Value
End Property
    
Friend Property Get StudyID() As String
StudyID = mStudyID
End Property
    
Friend Property Let UnderlyingStudyID(ByVal Value As String)
mUnderlyingStudyID = Value
End Property
    
Friend Property Get UnderlyingStudyID() As String
UnderlyingStudyID = mUnderlyingStudyID
End Property
    
Friend Property Let InputValueNames(ByRef Value() As String)
mInputValueNames = Value
End Property
    
Friend Property Get InputValueNames() As String()
InputValueNames = mInputValueNames
End Property
    
Friend Property Let NumberOfCachedValues(ByVal Value As Long)
mNumberOfCachedValues = Value
End Property
    
Friend Property Get NumberOfCachedValues() As Long
NumberOfCachedValues = mNumberOfCachedValues
End Property
    
Friend Property Let IncludeDataOutsideSession(ByVal Value As Boolean)
mIncludeDataOutsideSession = Value
End Property
    
Friend Property Get IncludeDataOutsideSession() As Boolean
IncludeDataOutsideSession = mIncludeDataOutsideSession
End Property
    
Friend Property Let StudyStarted(ByVal Value As Boolean)
mStudyStarted = Value
End Property

Friend Property Get StudyStarted() As Boolean
StudyStarted = mStudyStarted
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




