VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyBaseForNullInput"
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

Implements IStudyBase

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

Private Const ModuleName                            As String = "StudyBaseForNullInput"

'@================================================================================
' Member variables
'@================================================================================

Private mStudyManager                               As StudyManager
Private mStudyInputHandler                          As StudyInputHandler

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStudyBase Interface Members
'@================================================================================

Private Property Get IStudyBase_BaseStudy() As IStudy
Const ProcName As String = "IStudyBase_BaseStudy"
On Error GoTo Err

Set IStudyBase_BaseStudy = mStudyInputHandler

Exit Property

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudyBase_StudyManager() As StudyManager
Const ProcName As String = "IStudyBase_StudyManager"
On Error GoTo Err

Set IStudyBase_StudyManager = mStudyManager

Exit Property

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Sub Initialise( _
                ByVal pStudyManager As StudyManager, _
                ByVal pName As String)
Set mStudyManager = pStudyManager

If pName = "" Then pName = GenerateGUIDString

Set mStudyInputHandler = mStudyManager.CreateStudyInputHandler(pName)
End Sub

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




