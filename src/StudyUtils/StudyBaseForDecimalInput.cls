VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyBaseForDecimalInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "StudyBaseForDecimalInput"

'@================================================================================
' Member variables
'@================================================================================

Private mStudyManager                               As StudyManager
Private mStudyInputHandler                          As StudyInputHandler

Private mInputHandleValue                           As Long
Private mInputHandleBarNumber                       As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStudyBase Interface Members
'@================================================================================

Private Property Get IStudyBase_BaseStudy() As IStudy
Set IStudyBase_BaseStudy = BaseStudy
End Property

Private Property Get IStudyBase_StudyManager() As StudyManager
Set IStudyBase_StudyManager = StudyManager
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get BaseStudy() As IStudy
Set BaseStudy = mStudyInputHandler
End Property

Public Property Get StudyManager() As StudyManager
Set StudyManager = mStudyManager
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pStudyManager As StudyManager, _
                ByVal pQuantum As BoxedDecimal, _
                ByVal pName As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStudyManager = pStudyManager

If pName = "" Then pName = GenerateGUIDString

Set mStudyInputHandler = mStudyManager.CreateStudyInputHandler(pName)
mInputHandleValue = mStudyManager.AddInput(mStudyInputHandler, _
                        ValueInputName, _
                        "Value", _
                        StudyInputTypes.InputTypeReal, _
                        True, _
                        pQuantum)
mInputHandleBarNumber = mStudyManager.AddInput(mStudyInputHandler, _
                        BarNumberInputName, _
                        "Bar number", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        DecimalOne)

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyBarNumber(pBarNumber As Long, ByVal pTimestamp As Date)
Const ProcName As String = "NotifyBarNumber"
On Error GoTo Err

mStudyManager.NotifyInput mInputHandleBarNumber, pBarNumber, pTimestamp

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyValue(pValue As BoxedDecimal, ByVal pTimestamp As Date)
Const ProcName As String = "NotifyValue"
On Error GoTo Err

mStudyManager.NotifyInput mInputHandleValue, pValue, pTimestamp

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




