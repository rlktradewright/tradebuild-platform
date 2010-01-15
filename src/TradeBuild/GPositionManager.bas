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

Private mPositionManagers               As New PositionManagers
Private mPositionManagersSimulated      As New PositionManagers

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
                ByVal pWorkspace As Workspace) As PositionManager
Const ProcName As String = "gCreatePositionManager"
Dim failpoint As String
On Error GoTo Err

Set gCreatePositionManager = New PositionManager
gCreatePositionManager.Initialise pKey, pWorkspace
mPositionManagers.Add gCreatePositionManager

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gCreatePositionManagerSimulated( _
                ByVal pKey As String, _
                ByVal pWorkspace As Workspace) As PositionManager
Const ProcName As String = "gCreatePositionManagerSimulated"
Dim failpoint As String
On Error GoTo Err

Set gCreatePositionManagerSimulated = New PositionManager
gCreatePositionManagerSimulated.Initialise pKey, pWorkspace
mPositionManagersSimulated.Add gCreatePositionManagerSimulated

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gGetPositionManager( _
                ByVal pKey As String) As PositionManager
Const ProcName As String = "gGetPositionManager"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set gGetPositionManager = mPositionManagers.Item(pKey)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gGetPositionManagerSimulated( _
                ByVal pKey As String) As PositionManager
Const ProcName As String = "gGetPositionManagerSimulated"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set gGetPositionManagerSimulated = mPositionManagersSimulated.Item(pKey)
On Error GoTo Err

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function gNextApplicationIndex() As Long
Static lNextApplicationIndex As Long

Const ProcName As String = "gNextApplicationIndex"
Dim failpoint As String
On Error GoTo Err

gNextApplicationIndex = lNextApplicationIndex
lNextApplicationIndex = lNextApplicationIndex + 1

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================





