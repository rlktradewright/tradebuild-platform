VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IServiceProviderManager"
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

Private Const ModuleName                            As String = "IServiceProviderManager"

'@================================================================================
' Member variables
'@================================================================================

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

Public Sub NotifyEvent( _
                ByVal pEventCode As Long, _
                ByVal pEventMessage As String, _
                ByVal pServiceProviderHandle As Long)
End Sub

'Public Sub NotifyFatalError( _
'                ByVal pErrorCode As Long, _
'                ByVal pSource As String, _
'                ByVal pErrorMessage As String, _
'                ByVal pServiceProviderHandle As Long)
'End Sub

Public Sub NotifyRecoverableError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMessage As String, _
                ByVal pServiceProviderHandle As Long)
End Sub

Public Sub RegisterServiceProvider( _
                ByVal pHandle As Long, _
                ByVal pRole As ServiceProviderRoles)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




