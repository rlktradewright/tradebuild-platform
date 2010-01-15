VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServiceProviderError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables and constants
'@================================================================================

Private mTimestamp As Date
Private mErrorCode As Long
Private mServiceProviderName As String
Private mMessage As String


'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ErrorCode(ByVal value As Long)
mErrorCode = value
End Property

Public Property Get ErrorCode() As Long
ErrorCode = mErrorCode
End Property

Friend Property Let Message(ByVal value As String)
mMessage = value
End Property

Public Property Get Message() As String
Message = mMessage
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

Public Property Get ServiceProviderName() As String
ServiceProviderName = mServiceProviderName
End Property

Friend Property Let Timestamp(ByVal value As Date)
mTimestamp = value
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property



'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


