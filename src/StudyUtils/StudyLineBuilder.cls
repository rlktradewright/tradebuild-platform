VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyLineBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                As String = "StudyLineBuilder"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyLine As StudyLine

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyLine = New StudyLine
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Point1(ByVal Value As StudyPoint)
Const ProcName As String = "Point1"
On Error GoTo Err

mStudyLine.Point1 = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Point2(ByVal Value As StudyPoint)
Const ProcName As String = "Point2"
On Error GoTo Err

mStudyLine.Point2 = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyLine() As StudyLine
Const ProcName As String = "StudyLine"
On Error GoTo Err

Set StudyLine = mStudyLine

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================



