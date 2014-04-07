Attribute VB_Name = "GPositionManager"
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

Private Const ModuleName                            As String = "GPositionManager"

'@================================================================================
' Member variables
'@================================================================================

Private mPositionManagers               As PositionManagers
Private mPositionManagersSimulated      As PositionManagers

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

'@================================================================================
' Methods
'@================================================================================

Public Function gCreatePositionManager( _
                ByVal pKey As String, _
                ByVal pWorkspace As WorkSpace) As PositionManager
Const ProcName As String = "gCreatePositionManager"

On Error GoTo Err

Set gCreatePositionManager = New PositionManager
gCreatePositionManager.Initialise pKey, pWorkspace
mPositionManagers.Add gCreatePositionManager

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Function gCreatePositionManagerSimulated( _
                ByVal pKey As String, _
                ByVal pWorkspace As WorkSpace) As PositionManager
Const ProcName As String = "gCreatePositionManagerSimulated"

On Error GoTo Err

Set gCreatePositionManagerSimulated = New PositionManager
gCreatePositionManagerSimulated.Initialise pKey, pWorkspace
mPositionManagersSimulated.Add gCreatePositionManagerSimulated

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Function gGetPositionManager( _
                ByVal pKey As String) As PositionManager
Const ProcName As String = "gGetPositionManager"

On Error GoTo Err

On Error Resume Next
Set gGetPositionManager = mPositionManagers.Item(pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Function gGetPositionManagersEnumerator() As Enumerator
Const ProcName As String = "gGetPositionManagersEnumerator"
On Error GoTo Err

Set gGetPositionManagersEnumerator = mPositionManagers.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Function gGetPositionManagerSimulated( _
                ByVal pKey As String) As PositionManager
Const ProcName As String = "gGetPositionManagerSimulated"

On Error GoTo Err

On Error Resume Next
Set gGetPositionManagerSimulated = mPositionManagersSimulated.Item(pKey)
On Error GoTo Err

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Function gGetPositionManagersSimulatedEnumerator() As Enumerator
Const ProcName As String = "gGetPositionManagersSimulatedEnumerator"
On Error GoTo Err

Set gGetPositionManagersSimulatedEnumerator = mPositionManagersSimulated.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Function gInitialise()

Const ProcName As String = "gInitialise"
On Error GoTo Err

Set mPositionManagers = New PositionManagers
Set mPositionManagersSimulated = New PositionManagers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Function gNextApplicationIndex() As Long
Static lNextApplicationIndex As Long

Const ProcName As String = "gNextApplicationIndex"

On Error GoTo Err

gNextApplicationIndex = lNextApplicationIndex
lNextApplicationIndex = lNextApplicationIndex + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function

Public Sub gRemovePositionManager( _
                ByVal pPositionManager As PositionManager)
Const ProcName As String = "gRemovePositionManager"

On Error GoTo Err

mPositionManagers.Remove pPositionManager

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





