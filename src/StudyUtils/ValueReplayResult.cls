VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValueReplayResult"
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


Private Const ModuleName                    As String = "ValueReplayResult"

'@================================================================================
' Member variables
'@================================================================================

Private mTarget As Object
Private mSourceStudy As Study
Private mValueName As String
Private mValueNames() As String
Private mNotifyName As String
Private mNotifyNames() As String

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

Public Property Get notifyName() As String
notifyName = mNotifyName
End Property

Public Property Get notifyNames() As String()
notifyNames = mNotifyNames
End Property

Public Property Get sourceStudy() As Study
Set sourceStudy = mSourceStudy
End Property

Public Property Get target() As Object
Set target = mTarget
End Property

' the same as ValueNames(0) for user convenience
Public Property Get valueName() As String
valueName = mValueName
End Property

Public Property Get valueNames() As String()
valueNames = mValueNames
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal target As Object, _
                ByVal sourceStudy As Study, _
                ByRef valueNames() As String, _
                ByRef notifyNames() As String)

Const ProcName As String = "initialise"
On Error GoTo Err

Set mTarget = target
Set mSourceStudy = sourceStudy
mValueNames = valueNames
mValueName = mValueNames(0)
mNotifyNames = notifyNames
mNotifyName = mNotifyNames(0)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


