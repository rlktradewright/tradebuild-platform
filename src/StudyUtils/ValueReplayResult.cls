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
Private mSourceStudy As IStudy
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

Public Property Get NotifyName() As String
NotifyName = mNotifyName
End Property

Public Property Get NotifyNames() As String()
NotifyNames = mNotifyNames
End Property

Public Property Get SourceStudy() As IStudy
Set SourceStudy = mSourceStudy
End Property

Public Property Get Target() As Object
Set Target = mTarget
End Property

' the same as ValueNames(0) for user convenience
Public Property Get ValueName() As String
ValueName = mValueName
End Property

Public Property Get ValueNames() As String()
ValueNames = mValueNames
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal Target As Object, _
                ByVal SourceStudy As IStudy, _
                ByRef ValueNames() As String, _
                ByRef NotifyNames() As String)

Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTarget = Target
Set mSourceStudy = SourceStudy
mValueNames = ValueNames
mValueName = mValueNames(0)
mNotifyNames = NotifyNames
mNotifyName = mNotifyNames(0)

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


