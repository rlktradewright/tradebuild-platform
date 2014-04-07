VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderIdMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "OrderIdMap"

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildIds                              As Collection
Private mBrokerIds                                  As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTradeBuildIds = New Collection
Set mBrokerIds = New Collection
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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add( _
                ByRef pTradeBuildId As String, _
                ByRef pBrokerId As String)
Const ProcName As String = "Add"
On Error GoTo Err

On Error Resume Next
mTradeBuildIds.Add pTradeBuildId, pBrokerId
mBrokerIds.Add pBrokerId, pTradeBuildId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Sub

Friend Function GetBrokerId( _
                ByRef pTradeBuildId As String) As String
Const ProcName As String = "GetBrokerId"

On Error GoTo Err

On Error Resume Next
GetBrokerId = mBrokerIds.Item(pTradeBuildId)
On Error GoTo Err

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function
                
Friend Function GetTradeBuildId( _
                ByRef pBrokerId As String) As String
Const ProcName As String = "GetTradeBuildId"

On Error GoTo Err

On Error Resume Next
GetTradeBuildId = mTradeBuildIds.Item(pBrokerId)
On Error GoTo Err
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, pProjectName:=ProjectName
End Function
                
'@================================================================================
' Helper Functions
'@================================================================================


